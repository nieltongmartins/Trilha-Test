"""Orquestra a auditoria incremental independentemente da fonte de versões."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
import logging
import sqlite3
from collections.abc import Callable
from uuid import uuid4

from app.database import Database
from app.excel.comparator import CellChange, compare_snapshots
from app.excel.reader import CellValue, Snapshot, read_workbook
from app.integrity import sha256_file
from app.models import AuditExecutionStatus, ProcessedVersionStatus
from app.sources.base import SpreadsheetInfo, VersionInfo, VersionSource


logger = logging.getLogger("auditoria_excel.audit")


@dataclass(frozen=True, slots=True)
class AuditResult:
    execution_code: str
    status: AuditExecutionStatus
    processed_versions: int
    changes: int
    initial_checkpoint: str | None
    final_version: str | None


class AuditService:
    """Executa comparações consecutivas e confirma cada uma atomicamente."""

    def __init__(
        self,
        database: Database,
        source: VersionSource,
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> None:
        self.database = database
        self.source = source
        self.progress_callback = progress_callback

    def audit(self, spreadsheet: SpreadsheetInfo) -> AuditResult:
        connection = self.database.connection
        spreadsheet_id = self._upsert_spreadsheet(connection, spreadsheet)
        checkpoint = connection.execute(
            "SELECT versao_id, versao_numero FROM checkpoint WHERE planilha_id = ?",
            (spreadsheet_id,),
        ).fetchone()
        initial_checkpoint = checkpoint["versao_numero"] if checkpoint else None
        execution_code = f"AUD-{datetime.now(timezone.utc):%Y%m%d%H%M%S}-{uuid4().hex[:8]}"
        execution_id = connection.execute(
            """
            INSERT INTO execucao_auditoria
                (codigo_execucao, planilha_id, checkpoint_inicial, status)
            VALUES (?, ?, ?, ?)
            """,
            (
                execution_code,
                spreadsheet_id,
                initial_checkpoint,
                AuditExecutionStatus.RUNNING.value,
            ),
        ).lastrowid
        connection.commit()
        assert execution_id is not None
        logger.info(
            "Auditoria iniciada execucao=%s planilha=%s identidade=%s/%s/%s checkpoint=%s",
            execution_code,
            spreadsheet.name,
            spreadsheet.site_id,
            spreadsheet.drive_id,
            spreadsheet.drive_item_id,
            initial_checkpoint or "nenhum",
        )

        try:
            versions = list(self.source.list_versions(spreadsheet))
            pairs = self._pending_pairs(versions, checkpoint["versao_id"] if checkpoint else None)
            logger.info(
                "Versões descobertas execucao=%s planilha=%s total=%d comparacoes_pendentes=%d",
                execution_code,
                spreadsheet.name,
                len(versions),
                len(pairs),
            )
            self._report_progress(0, len(pairs))
        except Exception as error:
            return self._record_failure(
                connection, execution_id, spreadsheet_id, execution_code,
                initial_checkpoint, 0, 0, None, None, error,
            )

        if not pairs:
            final = initial_checkpoint
            self._finish_execution(
                connection, execution_id, AuditExecutionStatus.COMPLETED_WITHOUT_UPDATES,
                final, 0, 0, None,
            )
            logger.info(
                "Auditoria sem novidades execucao=%s planilha=%s checkpoint=%s",
                execution_code,
                spreadsheet.name,
                final or "nenhum",
            )
            return AuditResult(
                execution_code, AuditExecutionStatus.COMPLETED_WITHOUT_UPDATES,
                0, 0, initial_checkpoint, final,
            )

        processed = 0
        total_changes = 0
        final = initial_checkpoint
        previous_snapshot = None
        for previous, current in pairs:
            try:
                logger.debug(
                    "Comparação iniciada execucao=%s planilha=%s versao_anterior=%s versao_atual=%s",
                    execution_code,
                    spreadsheet.name,
                    previous.number,
                    current.number,
                )
                if previous_snapshot is None:
                    previous_snapshot, _ = self._read_temporary_version(spreadsheet, previous)
                current_snapshot, current_hash = self._read_temporary_version(
                    spreadsheet, current
                )
                changes = compare_snapshots(previous_snapshot, current_snapshot)
                self._persist_comparison(
                    connection, spreadsheet_id, execution_id, previous, current,
                    current_hash, changes,
                )
                logger.debug(
                    "Comparação concluída execucao=%s planilha=%s versao_atual=%s alteracoes=%d",
                    execution_code,
                    spreadsheet.name,
                    current.number,
                    len(changes),
                )
            except Exception as error:
                connection.rollback()
                return self._record_failure(
                    connection, execution_id, spreadsheet_id, execution_code,
                    initial_checkpoint, processed, total_changes, previous, current, error,
                )
            processed += 1
            self._report_progress(processed, len(pairs))
            total_changes += len(changes)
            final = current.number
            previous_snapshot = current_snapshot

        self._finish_execution(
            connection, execution_id, AuditExecutionStatus.COMPLETED,
            final, processed, total_changes, None,
        )
        logger.info(
            "Auditoria concluída execucao=%s planilha=%s versoes_processadas=%d alteracoes=%d checkpoint=%s",
            execution_code,
            spreadsheet.name,
            processed,
            total_changes,
            final or "nenhum",
        )
        return AuditResult(
            execution_code, AuditExecutionStatus.COMPLETED, processed,
            total_changes, initial_checkpoint, final,
        )

    def _report_progress(self, completed: int, total: int) -> None:
        if self.progress_callback is not None:
            try:
                self.progress_callback(completed, total)
            except Exception:
                logger.warning("Falha ao publicar progresso da auditoria", exc_info=True)

    def _read_temporary_version(
        self, spreadsheet: SpreadsheetInfo, version: VersionInfo
    ) -> tuple[Snapshot, str]:
        path = self.source.get_version(spreadsheet, version)
        try:
            # O digest representa exatamente o binário adquirido nesta execução,
            # antes que o XLSX temporário seja descartado pela fonte.
            digest = sha256_file(path)
            return read_workbook(path), digest
        finally:
            try:
                self.source.release_version(path)
            except Exception as error:
                logger.warning(
                    "Falha ao remover XLSX temporário planilha=%s versao=%s erro=%s",
                    spreadsheet.name,
                    version.number,
                    error,
                )

    @staticmethod
    def _pending_pairs(
        versions: list[VersionInfo], checkpoint_id: str | None
    ) -> list[tuple[VersionInfo, VersionInfo]]:
        ids = [version.id for version in versions]
        if len(ids) != len(set(ids)):
            raise ValueError("A fonte retornou identificadores de versão duplicados")
        if checkpoint_id is None:
            start = 0
        else:
            try:
                start = ids.index(checkpoint_id)
            except ValueError as error:
                raise ValueError("A versão do checkpoint não está disponível na fonte") from error
        return list(zip(versions[start:], versions[start + 1 :]))

    @staticmethod
    def _upsert_spreadsheet(
        connection: sqlite3.Connection, spreadsheet: SpreadsheetInfo
    ) -> int:
        connection.execute(
            """
            INSERT INTO planilha
                (site_id, drive_id, drive_item_id, nome_atual, caminho_sharepoint)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(site_id, drive_id, drive_item_id) DO UPDATE SET
                nome_atual = excluded.nome_atual,
                caminho_sharepoint = excluded.caminho_sharepoint,
                data_atualizacao = CURRENT_TIMESTAMP
            """,
            (spreadsheet.site_id, spreadsheet.drive_id, spreadsheet.drive_item_id,
             spreadsheet.name, spreadsheet.path),
        )
        row = connection.execute(
            """SELECT id FROM planilha
               WHERE site_id = ? AND drive_id = ? AND drive_item_id = ?""",
            (spreadsheet.site_id, spreadsheet.drive_id, spreadsheet.drive_item_id),
        ).fetchone()
        connection.commit()
        return int(row["id"])

    @staticmethod
    def _serialize(value: CellValue) -> str | None:
        return None if value is None else str(value)

    def _persist_comparison(
        self,
        connection: sqlite3.Connection,
        spreadsheet_id: int,
        execution_id: int,
        previous: VersionInfo,
        current: VersionInfo,
        current_hash: str,
        changes: list[CellChange],
    ) -> None:
        status = (
            ProcessedVersionStatus.PROCESSED
            if changes
            else ProcessedVersionStatus.WITHOUT_CHANGES
        )
        with connection:
            cursor = connection.execute(
                """
                INSERT INTO versao_processada (
                    planilha_id, versao_anterior_id, versao_anterior_numero,
                    versao_atual_id, versao_atual_numero, data_hora_versao,
                    autor, autor_email, autor_login, comentario, tamanho,
                    url_origem, versao_atual, quantidade_alteracoes, status,
                    hash_origem, execucao_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (spreadsheet_id, previous.id, previous.number, current.id,
                 current.number, current.modified_at, current.author,
                 current.author_email, current.author_login, current.comment,
                 current.size, current.source_url, int(current.is_current),
                 len(changes), status.value, current_hash, execution_id),
            )
            processed_id = cursor.lastrowid
            assert processed_id is not None
            connection.executemany(
                """
                INSERT INTO alteracao (
                    versao_processada_id, planilha_id, tipo, aba, endereco,
                    valor_anterior, valor_novo
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    (processed_id, spreadsheet_id, change.change_type.value,
                     change.sheet, change.address,
                     self._serialize(change.previous_value),
                     self._serialize(change.new_value))
                    for change in changes
                ],
            )
            connection.execute(
                """
                INSERT INTO checkpoint
                    (planilha_id, versao_id, versao_numero, data_hora_versao)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(planilha_id) DO UPDATE SET
                    versao_id = excluded.versao_id,
                    versao_numero = excluded.versao_numero,
                    data_hora_versao = excluded.data_hora_versao,
                    data_atualizacao = CURRENT_TIMESTAMP
                """,
                (spreadsheet_id, current.id, current.number, current.modified_at),
            )

    @staticmethod
    def _finish_execution(
        connection: sqlite3.Connection,
        execution_id: int,
        status: AuditExecutionStatus,
        final: str | None,
        processed: int,
        changes: int,
        message: str | None,
    ) -> None:
        connection.execute(
            """
            UPDATE execucao_auditoria SET
                fim = CURRENT_TIMESTAMP, versao_final = ?,
                versoes_processadas = ?, alteracoes_encontradas = ?,
                status = ?, mensagem = ?
            WHERE id = ?
            """,
            (final, processed, changes, status.value, message, execution_id),
        )
        connection.commit()

    def _record_failure(
        self, connection: sqlite3.Connection, execution_id: int,
        spreadsheet_id: int, execution_code: str,
        initial_checkpoint: str | None, processed: int, changes: int,
        previous: VersionInfo | None, current: VersionInfo | None,
        error: Exception,
    ) -> AuditResult:
        message = str(error)
        connection.execute(
            """
            INSERT INTO erro_processamento (
                execucao_id, planilha_id, versao_anterior, versao_atual,
                tipo_erro, mensagem
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (execution_id, spreadsheet_id,
             previous.number if previous else None,
             current.number if current else None,
             type(error).__name__, message),
        )
        checkpoint = connection.execute(
            "SELECT versao_numero FROM checkpoint WHERE planilha_id = ?",
            (spreadsheet_id,),
        ).fetchone()
        final = checkpoint["versao_numero"] if checkpoint else None
        self._finish_execution(
            connection, execution_id, AuditExecutionStatus.FAILED,
            final, processed, changes, message,
        )
        logger.error(
            "Auditoria falhou execucao=%s planilha_id=%d etapa=%s->%s tipo_erro=%s checkpoint_preservado=%s erro=%s",
            execution_code,
            spreadsheet_id,
            previous.number if previous else "listagem",
            current.number if current else "listagem",
            type(error).__name__,
            final or "nenhum",
            error,
        )
        return AuditResult(
            execution_code, AuditExecutionStatus.FAILED, processed, changes,
            initial_checkpoint, final,
        )
