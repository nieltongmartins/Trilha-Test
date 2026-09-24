"""Catálogo SQLite de metadados de versões, separado do checkpoint da auditoria."""

from __future__ import annotations

from dataclasses import dataclass, replace
import logging
import sqlite3
import time
from typing import Callable

from app.database import Database
from app.sources.base import SpreadsheetInfo, VersionInfo, VersionSource


logger = logging.getLogger("auditoria_excel.version_catalog")


class CatalogInconsistencyError(RuntimeError):
    """A fronteira local não pôde ser comprovada contra o SharePoint."""


@dataclass(frozen=True, slots=True)
class CatalogSyncResult:
    versions: tuple[VersionInfo, ...]
    source: str
    new_versions: int


def workbook_identity(spreadsheet: SpreadsheetInfo) -> str:
    return "|".join((spreadsheet.site_id, spreadsheet.drive_id, spreadsheet.drive_item_id))


class VersionCatalog:
    """Carrega e sincroniza somente metadados; nunca persiste arquivos XLSX."""

    def __init__(self, database: Database, source: VersionSource) -> None:
        self.database = database
        self.source = source

    def load(self, spreadsheet: SpreadsheetInfo) -> tuple[tuple[VersionInfo, ...], sqlite3.Row | None]:
        started = time.perf_counter()
        identity = workbook_identity(spreadsheet)
        state = self.database.connection.execute(
            "SELECT * FROM version_catalog_state WHERE workbook_identity=?",
            (identity,),
        ).fetchone()
        rows = self.database.connection.execute(
            """SELECT technical_version_id, version_label, created_at_sharepoint,
                      is_current_snapshot
                 FROM version_catalog WHERE workbook_identity=?
                ORDER BY CAST(technical_version_id AS INTEGER)""",
            (identity,),
        ).fetchall()
        versions = tuple(
            VersionInfo(str(row[0]), row[1], row[2], is_current=bool(row[3]))
            for row in rows
        )
        logger.info(
            "VERSION_CATALOG_LOAD workbook=%s count=%d duration=%.3f max_known_id=%s current_id=%s",
            spreadsheet.name, len(versions), time.perf_counter() - started,
            versions[-1].id if versions else None,
            state["last_known_current_id"] if state else None,
        )
        return versions, state

    def synchronize(
        self,
        spreadsheet: SpreadsheetInfo,
        progress_callback: Callable[[int], None] | None = None,
    ) -> CatalogSyncResult:
        total_started = time.perf_counter()
        local, state = self.load(spreadsheet)
        if not local or state is None or state["catalog_status"] != "VALID":
            return self._full_rebuild(spreadsheet, progress_callback, total_started)

        checked = time.perf_counter()
        current = self.source.get_current_version(spreadsheet)  # type: ignore[attr-defined]
        changed = (current.id, current.number) != (
            state["last_known_current_id"], state["last_known_current_label"]
        )
        logger.info(
            "VERSION_CATALOG_REMOTE_CHECK cached_current_id=%s remote_current_id=%s changed=%s duration=%.3f",
            state["last_known_current_id"], current.id, str(changed).lower(),
            time.perf_counter() - checked,
        )
        if not changed:
            try:
                self._verify_local(local, state)
            except CatalogInconsistencyError as error:
                logger.warning(
                    "Catálogo local corrompido workbook=%s motivo=%s",
                    spreadsheet.name, error,
                )
                return self._full_rebuild(
                    spreadsheet, progress_callback, total_started
                )
            self.database.connection.execute(
                """UPDATE version_catalog_state SET last_sync_at=CURRENT_TIMESTAMP
                    WHERE workbook_identity=?""", (workbook_identity(spreadsheet),)
            )
            self.database.connection.commit()
            self._log_complete(spreadsheet, "local_only", len(local), total_started)
            return CatalogSyncResult(local, "local_only", 0)

        try:
            if int(current.id) < int(state["last_known_current_id"]):
                raise CatalogInconsistencyError("current_id remoto voltou para trás")
            delta_started = time.perf_counter()
            delta = tuple(self.source.list_version_delta(  # type: ignore[attr-defined]
                spreadsheet,
                state["last_known_current_id"],
                state["last_known_current_label"],
                progress_callback,
            ))
            merged, new_count = self._validate_delta(local, state, current, delta)
            logger.info(
                "VERSION_CATALOG_DELTA anchor_id=%s new_historical_count=%d new_current_id=%s duration=%.3f",
                state["last_known_current_id"], max(new_count - 1, 0), current.id,
                time.perf_counter() - delta_started,
            )
            self._replace(spreadsheet, merged)
            self._log_complete(spreadsheet, "delta", len(merged), total_started)
            return CatalogSyncResult(merged, "delta", new_count)
        except (CatalogInconsistencyError, ValueError) as error:
            logger.warning("Catálogo requer reconciliação workbook=%s motivo=%s", spreadsheet.name, error)
            self.database.connection.execute(
                "UPDATE version_catalog_state SET catalog_status='NEEDS_RECONCILIATION' WHERE workbook_identity=?",
                (workbook_identity(spreadsheet),),
            )
            self.database.connection.commit()
            return self._full_rebuild(spreadsheet, progress_callback, total_started)

    def _full_rebuild(self, spreadsheet, progress_callback, total_started):
        versions = tuple(self.source.list_versions(spreadsheet, progress_callback=progress_callback))
        self._validate_versions(versions)
        self._replace(spreadsheet, versions)
        self._log_complete(spreadsheet, "full_rebuild", len(versions), total_started)
        return CatalogSyncResult(versions, "full_rebuild", len(versions))

    @staticmethod
    def _verify_local(versions, state) -> None:
        VersionCatalog._validate_versions(versions)
        current = versions[-1]
        if (current.id, current.number, len(versions)) != (
            state["last_known_current_id"], state["last_known_current_label"], state["catalog_count"]
        ):
            raise CatalogInconsistencyError("estado do catálogo local divergente")

    @staticmethod
    def _validate_versions(versions) -> None:
        if not versions:
            raise CatalogInconsistencyError("catálogo vazio")
        try:
            ids = [int(item.id) for item in versions]
        except ValueError as error:
            raise CatalogInconsistencyError("ID técnico não numérico") from error
        if ids != sorted(ids) or len(ids) != len(set(ids)):
            raise CatalogInconsistencyError("IDs técnicos não são únicos e monotônicos")
        by_id = {item.id: item.number for item in versions}
        if len(by_id) != len(versions) or len({item.number for item in versions}) != len(versions):
            raise CatalogInconsistencyError("duplicidade estrutural no catálogo")
        if sum(item.is_current for item in versions) != 1 or not versions[-1].is_current:
            raise CatalogInconsistencyError("current metadata incoerente")

    @classmethod
    def _validate_delta(cls, local, state, remote_current, delta):
        if not delta:
            raise CatalogInconsistencyError("delta vazio apesar de watermark alterado")
        cls._validate_versions(delta)
        anchor = delta[0]
        if (anchor.id, anchor.number) != (
            state["last_known_current_id"], state["last_known_current_label"]
        ):
            raise CatalogInconsistencyError("anchor técnico esperado não encontrado")
        if (delta[-1].id, delta[-1].number) != (remote_current.id, remote_current.number):
            raise CatalogInconsistencyError("watermark do delta diverge do current remoto")
        known = {item.id: item.number for item in local}
        for item in delta:
            if item.id in known and known[item.id] != item.number:
                raise CatalogInconsistencyError("mesmo ID técnico possui label conflitante")
        prefix = tuple(replace(item, is_current=False) for item in local[:-1])
        merged = prefix + tuple(delta)
        cls._validate_versions(merged)
        return merged, len(merged) - len(local)

    def _replace(self, spreadsheet: SpreadsheetInfo, versions: tuple[VersionInfo, ...]) -> None:
        identity = workbook_identity(spreadsheet)
        historical = versions[-2] if len(versions) > 1 else None
        connection = self.database.connection
        with connection:
            connection.execute("DELETE FROM version_catalog WHERE workbook_identity=?", (identity,))
            connection.executemany(
                """INSERT INTO version_catalog
                   (workbook_identity, technical_version_id, version_label,
                    created_at_sharepoint, is_current_snapshot)
                   VALUES (?, ?, ?, ?, ?)""",
                ((identity, v.id, v.number, v.modified_at, int(v.is_current)) for v in versions),
            )
            connection.execute(
                """INSERT INTO version_catalog_state
                   (workbook_identity, last_historical_id, last_historical_label,
                    last_known_current_id, last_known_current_label, catalog_count, catalog_status)
                   VALUES (?, ?, ?, ?, ?, ?, 'VALID')
                   ON CONFLICT(workbook_identity) DO UPDATE SET
                    last_historical_id=excluded.last_historical_id,
                    last_historical_label=excluded.last_historical_label,
                    last_known_current_id=excluded.last_known_current_id,
                    last_known_current_label=excluded.last_known_current_label,
                    catalog_count=excluded.catalog_count, last_sync_at=CURRENT_TIMESTAMP,
                    catalog_status='VALID'""",
                (identity, historical.id if historical else None,
                 historical.number if historical else None, versions[-1].id,
                 versions[-1].number, len(versions)),
            )

    @staticmethod
    def _log_complete(spreadsheet, source, count, started):
        logger.info(
            "VERSION_CATALOG_SYNC_COMPLETE workbook=%s source=%s count=%d duration=%.3f",
            spreadsheet.name, source, count, time.perf_counter() - started,
        )
