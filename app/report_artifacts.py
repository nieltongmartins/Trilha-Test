"""Identifica e gerencia somente relatórios produzidos pela aplicação."""

from __future__ import annotations

from dataclasses import dataclass
import hashlib
import logging
from pathlib import Path
import re
import sqlite3
from collections.abc import Callable, Iterable


logger = logging.getLogger("auditoria_excel.reports")


class ReportArtifactError(RuntimeError):
    """Indica falha segura ao localizar, abrir ou remover um relatório."""


@dataclass(frozen=True)
class ReportIdentity:
    spreadsheet_id: int
    name: str
    site_id: str
    drive_id: str
    unique_id: str


class ReportArtifactManager:
    """Regra única da associação entre identidade canônica e XLSX derivado."""

    def __init__(self, connection: sqlite3.Connection, directory: str | Path) -> None:
        self.connection = connection
        self.directory = Path(directory)

    @staticmethod
    def safe_stem(name: str) -> str:
        return re.sub(r"[^\w.-]+", "_", Path(name).stem).strip("_") or "Planilha"

    @staticmethod
    def identity_token(identity: ReportIdentity) -> str:
        canonical = "\x1f".join((identity.site_id, identity.drive_id, identity.unique_id))
        return hashlib.sha256(canonical.encode("utf-8")).hexdigest()[:24]

    def canonical_path(self, identity: ReportIdentity) -> Path:
        return self.directory / (
            f"{self.safe_stem(identity.name)}__{self.identity_token(identity)}"
            "_Trilha_Auditoria.xlsx"
        )

    def legacy_path(self, identity: ReportIdentity) -> Path:
        return self.directory / f"{self.safe_stem(identity.name)}_Trilha_Auditoria.xlsx"

    def identity(self, spreadsheet_id: int) -> ReportIdentity:
        row = self.connection.execute(
            "SELECT id, nome_atual, site_id, drive_id, drive_item_id FROM planilha WHERE id=?",
            (spreadsheet_id,),
        ).fetchone()
        if row is None:
            raise ValueError("Auditoria selecionada não existe.")
        return ReportIdentity(
            row["id"], row["nome_atual"], row["site_id"], row["drive_id"],
            row["drive_item_id"],
        )

    def _legacy_is_unambiguous(self, identity: ReportIdentity) -> bool:
        stems = [
            self.safe_stem(row["nome_atual"])
            for row in self.connection.execute("SELECT nome_atual FROM planilha")
        ]
        return stems.count(self.safe_stem(identity.name)) == 1

    def locate(self, spreadsheet_id: int) -> Path | None:
        identity = self.identity(spreadsheet_id)
        canonical = self.canonical_path(identity)
        if canonical.is_file():
            return canonical
        legacy = self.legacy_path(identity)
        if legacy.is_file() and self._legacy_is_unambiguous(identity):
            return legacy
        if legacy.exists():
            logger.warning("Relatório legado ambíguo ignorado arquivo=%s", legacy)
        return None

    def controlled_paths(self, identities: Iterable[ReportIdentity]) -> list[Path]:
        result: list[Path] = []
        for identity in identities:
            canonical = self.canonical_path(identity)
            if canonical.is_file():
                result.append(canonical)
            legacy = self.legacy_path(identity)
            if legacy.is_file() and self._legacy_is_unambiguous(identity):
                result.append(legacy)
        return list(dict.fromkeys(result))

    def delete_with_database(self, paths: Iterable[Path], delete: Callable[[], None]) -> None:
        """Move antes da transação e restaura se o banco não puder ser alterado."""
        staged: list[tuple[Path, Path]] = []
        try:
            for path in paths:
                temporary = path.with_name(f".{path.name}.deleting")
                path.replace(temporary)
                staged.append((path, temporary))
            delete()
        except Exception as error:
            for original, temporary in reversed(staged):
                if temporary.exists():
                    temporary.replace(original)
            logger.exception("Exclusão segura de relatório/auditoria falhou")
            raise ReportArtifactError(
                "Não foi possível excluir o relatório associado. Feche-o no Excel "
                "e tente novamente; a auditoria local foi preservada."
            ) from error
        for _original, temporary in staged:
            try:
                temporary.unlink()
            except OSError as error:
                logger.error("Artefato temporário não pôde ser removido: %s", temporary)
                raise ReportArtifactError(
                    f"A auditoria foi excluída, mas a limpeza final falhou: {temporary}"
                ) from error
