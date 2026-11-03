"""Orquestra a auditoria incremental independentemente da fonte de versões."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
import logging
import sqlite3
import threading
import time
from collections.abc import Callable
from uuid import uuid4

from app.audit_interval import AuditCoverage, get_audit_coverage
from app.database import Database
from app.excel.comparator import CellChange, compare_snapshots
from app.excel.formula_values import normalize_formula_value
from app.excel.reader import CellValue, ConsecutiveWorkbookReader, Snapshot, read_workbook
from app.integrity import sha256_file
from app.local_identity import local_executor_name
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
    error_message: str | None = None


@dataclass(frozen=True, slots=True)
class VersionProgress:
    """Evento imutável de uma etapa real do processamento de uma versão."""

    version: str
    percent: int
    stage: str
    occurred_at: float = field(default_factory=time.monotonic)


class AuditService:
    """Executa comparações consecutivas e confirma cada uma atomicamente."""

    def __init__(
        self,
        database: Database,
        source: VersionSource,
        progress_callback: Callable[[int, int], None] | None = None,
        checkpoint_callback: Callable[[str, int, int], None] | None = None,
        version_progress_callback: Callable[[VersionProgress], None] | None = None,
        pause_event: threading.Event | None = None,
        stop_event: threading.Event | None = None,
        control_callback: Callable[[str, str | None], None] | None = None,
        executor_name: str | None = None,
    ) -> None:
        self.database = database
        self.source = source
        self.progress_callback = progress_callback
        self.checkpoint_callback = checkpoint_callback
        self.version_progress_callback = version_progress_callback
        self.pause_event = pause_event or threading.Event()
        self.stop_event = stop_event or threading.Event()
        self.control_callback = control_callback
        self.executor_name = executor_name

    def audit(
        self,
        spreadsheet: SpreadsheetInfo,
        versions: list[VersionInfo] | tuple[VersionInfo, ...] | None = None,
        *,
        start_version_id: str | None = None,
        end_version_id: str | None = None,
    ) -> AuditResult:
        interval_mode = start_version_id is not None or end_version_id is not None
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
                (codigo_execucao, planilha_id, checkpoint_inicial, status, executor_local)
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                execution_code,
                spreadsheet_id,
                initial_checkpoint,
                AuditExecutionStatus.RUNNING.value,
                self.executor_name if self.executor_name is not None else local_executor_name(),
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
            if versions is None:
                list_versions = getattr(self.source, "list_versions")
                try:
                    versions = list(
                        list_versions(
                            spreadsheet,
                            checkpoint_id=checkpoint["versao_id"] if checkpoint else None,
                            checkpoint_label=(
                                checkpoint["versao_numero"] if checkpoint else None
                            ),
                        )
                    )
                except TypeError:
                    # Fontes locais/alternativas conservam o contrato mínimo.
                    versions = list(list_versions(spreadsheet))
                logger.info(
                    "Lista de versões obtida na auditoria execucao=%s total=%d",
                    execution_code,
                    len(versions),
                )
            else:
                versions = list(versions)
                logger.info(
                    "Lista de versões reutilizada do cache da interface execucao=%s total=%d",
                    execution_code,
                    len(versions),
                )
            coverage = get_audit_coverage(
                connection, spreadsheet_id, versions,
                start_version_id, end_version_id,
            )
            pairs = list(coverage.missing_pairs)
            self._configure_coverage_plan(coverage, interval_mode)
            if interval_mode:
                self._record_interval_request(connection, execution_id, coverage)
            else:
                self._record_full_coverage(connection, execution_id, coverage)
                self._advance_covered_prefix(connection, spreadsheet_id, coverage)
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
            current_checkpoint = connection.execute(
                "SELECT versao_numero FROM checkpoint WHERE planilha_id=?",
                (spreadsheet_id,),
            ).fetchone()
            final = current_checkpoint[0] if current_checkpoint else initial_checkpoint
            self._finish_execution(
                connection, execution_id, AuditExecutionStatus.COMPLETED_WITHOUT_UPDATES,
                final, 0, 0, None,
            )
            self._finish_coverage(connection, execution_id, 0, 0.0)
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
        workbook_reader = ConsecutiveWorkbookReader()
        self._workbook_reader = workbook_reader
        audit_perf_started = time.perf_counter()
        if self._wait_at_safe_point(initial_checkpoint):
            return self._stop_execution(
                connection, execution_id, execution_code, initial_checkpoint,
                0, 0, initial_checkpoint,
            )
        for pair_index, (previous, current) in enumerate(pairs):
            # Lacunas independentes não compartilham a borda. Reuso de snapshot
            # só é correto dentro do mesmo bloco consecutivo.
            if pair_index and pairs[pair_index - 1][1].id != previous.id:
                previous_snapshot = None
            pair_started = time.perf_counter()
            future_versions = tuple(pair[1] for pair in pairs[pair_index:])
            try:
                self._report_version_progress(
                    current.number, 0, f"Obtendo versão {current.number}..."
                )
                self._report_version_progress(current.number, 5, "Baixando dados...")
                logger.debug(
                    "Comparação iniciada execucao=%s planilha=%s versao_anterior=%s versao_atual=%s",
                    execution_code,
                    spreadsheet.name,
                    previous.number,
                    current.number,
                )
                previous_read_seconds = 0.0
                if previous_snapshot is None:
                    previous_started = time.perf_counter()
                    previous_snapshot, _ = self._read_temporary_version(
                        spreadsheet, previous,
                        prefetch_versions=self._planned_prefetch_versions(
                            previous, future_versions
                        ),
                    )
                    previous_read_seconds = time.perf_counter() - previous_started

                current_started = time.perf_counter()
                current_snapshot, current_hash = self._read_temporary_version(
                    spreadsheet, current,
                    prefetch_versions=self._planned_prefetch_versions(
                        current, future_versions[1:]
                    ),
                    report_stages=True,
                )
                current_read_seconds = time.perf_counter() - current_started

                compare_started = time.perf_counter()
                changes = compare_snapshots(previous_snapshot, current_snapshot)
                self._report_version_progress(current.number, 97, "Comparação concluída.")
                compare_seconds = time.perf_counter() - compare_started

                persist_started = time.perf_counter()
                self._report_version_progress(current.number, 97, "Salvando alterações...")
                if interval_mode:
                    self._persist_comparison(
                        connection, spreadsheet_id, execution_id, previous, current,
                        current_hash, changes, update_checkpoint=False,
                    )
                else:
                    self._persist_comparison(
                        connection, spreadsheet_id, execution_id, previous, current,
                        current_hash, changes,
                    )
                persist_seconds = time.perf_counter() - persist_started
                pair_seconds = time.perf_counter() - pair_started

                logger.info(
                    "PERF comparacao planilha=%s anterior=%s atual=%s "
                    "leitura_anterior=%.3fs leitura_atual=%.3fs comparar=%.3fs "
                    "banco=%.3fs total=%.3fs alteracoes=%d",
                    spreadsheet.name,
                    previous.number,
                    current.number,
                    previous_read_seconds,
                    current_read_seconds,
                    compare_seconds,
                    persist_seconds,
                    pair_seconds,
                    len(changes),
                )
                logger.debug(
                    "Comparação concluída execucao=%s planilha=%s versao_atual=%s alteracoes=%d",
                    execution_code,
                    spreadsheet.name,
                    current.number,
                    len(changes),
                )
            except Exception as error:
                self._cancel_prefetch()
                connection.rollback()
                return self._record_failure(
                    connection, execution_id, spreadsheet_id, execution_code,
                    initial_checkpoint, processed, total_changes, previous, current, error,
                )
            self._report_version_progress(current.number, 100, "Checkpoint confirmado.")
            processed += 1
            self._report_progress(processed, len(pairs))
            self._report_checkpoint(current.number, processed, len(pairs) - processed)
            total_changes += len(changes)
            final = current.number
            previous_snapshot = current_snapshot

            # Este é o único ponto de controle dentro do loop: a persistência e
            # o checkpoint da versão atual já foram confirmados atomicamente.
            if self._wait_at_safe_point(final):
                return self._stop_execution(
                    connection, execution_id, execution_code, initial_checkpoint,
                    processed, total_changes, final,
                )

        self._cancel_prefetch()
        if not interval_mode:
            checkpoint_after = connection.execute(
                "SELECT versao_numero FROM checkpoint WHERE planilha_id=?",
                (spreadsheet_id,),
            ).fetchone()
            final = checkpoint_after[0] if checkpoint_after else final
        self._finish_execution(
            connection, execution_id, AuditExecutionStatus.COMPLETED,
            final, processed, total_changes, None,
        )
        audit_perf_seconds = time.perf_counter() - audit_perf_started
        self._finish_coverage(connection, execution_id, processed, audit_perf_seconds)
        logger.info(
            "PERF auditoria_resumo planilha=%s comparacoes=%d total=%.3fs media=%.3fs_por_comparacao",
            spreadsheet.name,
            processed,
            audit_perf_seconds,
            (audit_perf_seconds / processed) if processed else 0.0,
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

    def _wait_at_safe_point(self, checkpoint: str | None) -> bool:
        """Obedece pausa/parada somente entre comparações confirmadas."""
        if self.stop_event.is_set():
            self._cancel_prefetch()
            return True
        if not self.pause_event.is_set():
            return False

        # Downloads especulativos não devem continuar ocupando o buffer durante
        # uma pausa. O WebDriver e a sessão autenticada não são encerrados.
        self._cancel_prefetch()
        self._report_control("paused", checkpoint)
        while self.pause_event.is_set():
            if self.stop_event.wait(0.1):
                return True
        self._report_control("resumed", checkpoint)
        return self.stop_event.is_set()

    def _report_control(self, state: str, checkpoint: str | None) -> None:
        if self.control_callback is not None:
            try:
                self.control_callback(state, checkpoint)
            except Exception:
                logger.warning("Falha ao publicar estado de controle", exc_info=True)

    def _stop_execution(
        self, connection: sqlite3.Connection, execution_id: int,
        execution_code: str, initial_checkpoint: str | None,
        processed: int, changes: int, final: str | None,
    ) -> AuditResult:
        self._cancel_prefetch()
        self._finish_execution(
            connection, execution_id, AuditExecutionStatus.STOPPED,
            final, processed, changes, "Interrompida pelo usuário.",
        )
        self._report_control("stopped", final)
        logger.info(
            "Auditoria interrompida pelo usuário execucao=%s checkpoint=%s",
            execution_code, final or "nenhum",
        )
        return AuditResult(
            execution_code, AuditExecutionStatus.STOPPED, processed, changes,
            initial_checkpoint, final,
        )

    def _report_version_progress(
        self, version: str, percent: int, stage: str
    ) -> None:
        if self.version_progress_callback is not None:
            try:
                self.version_progress_callback(VersionProgress(version, percent, stage))
            except Exception:
                logger.warning("Falha ao publicar progresso da versão", exc_info=True)

    def _report_progress(self, completed: int, total: int) -> None:
        if self.progress_callback is not None:
            try:
                self.progress_callback(completed, total)
            except Exception:
                logger.warning("Falha ao publicar progresso da auditoria", exc_info=True)

    def _report_checkpoint(
        self, version: str, processed: int, pending: int
    ) -> None:
        """Publica somente checkpoints cujo bloco transacional já foi confirmado."""
        if self.checkpoint_callback is not None:
            try:
                self.checkpoint_callback(version, processed, pending)
            except Exception:
                logger.warning(
                    "Falha ao publicar checkpoint confirmado da auditoria",
                    exc_info=True,
                )

    def _start_prefetch(
        self, spreadsheet: SpreadsheetInfo, version: VersionInfo
    ) -> None:
        prefetch = getattr(self.source, "prefetch_version", None)
        if not callable(prefetch):
            return
        try:
            prefetch(spreadsheet, version)
        except Exception as error:
            # Prefetch é apenas otimização. Uma falha aqui não invalida a
            # auditoria: get_version fará o download normal quando necessário.
            logger.warning(
                "Prefetch opcional falhou planilha=%s versao=%s erro=%s",
                spreadsheet.name,
                version.number,
                error,
            )

    def _planned_prefetch_versions(
        self,
        current: VersionInfo,
        candidates: tuple[VersionInfo, ...],
    ) -> tuple[VersionInfo, ...]:
        """Seleciona, em ordem, as próximas versões que cabem no buffer."""
        configured_size = getattr(self.source, "prefetch_buffer_size", 2)
        buffer_size = configured_size if isinstance(configured_size, int) else 2
        if buffer_size <= 0:
            return ()
        selected: list[VersionInfo] = []
        seen = {current.id}
        for candidate in candidates:
            if candidate.id in seen:
                continue
            seen.add(candidate.id)
            selected.append(candidate)
            if len(selected) >= buffer_size:
                break
        return tuple(selected)

    def _cancel_prefetch(self) -> None:
        cancel = getattr(self.source, "cancel_prefetch", None)
        if callable(cancel):
            try:
                cancel()
            except Exception:
                logger.debug("Falha ao cancelar prefetch opcional", exc_info=True)

    def _read_temporary_version(
        self,
        spreadsheet: SpreadsheetInfo,
        version: VersionInfo,
        *,
        prefetch_versions: tuple[VersionInfo, ...] = (),
        report_stages: bool = False,
    ) -> tuple[Snapshot, str]:
        total_started = time.perf_counter()
        download_started = time.perf_counter()
        path = self.source.get_version(spreadsheet, version)
        download_seconds = time.perf_counter() - download_started
        if report_stages:
            self._report_version_progress(version.number, 55, "Validando arquivo...")

        # Assim que o binário atual chegou ao disco, o Edge começa a buscar a
        # próxima versão em background. Enquanto isso, o Python calcula SHA,
        # lê o XLSX e compara a versão atual. Não há duas comparações paralelas
        # e o checkpoint continua sendo confirmado estritamente em ordem.
        logger.info(
            "PERF prefetch_planejado atual=%s futuras=[%s] quantidade=%d",
            version.number,
            ",".join(candidate.number for candidate in prefetch_versions),
            len(prefetch_versions),
        )
        for candidate in prefetch_versions:
            self._start_prefetch(spreadsheet, candidate)
        try:
            # O digest representa exatamente o binário adquirido nesta execução,
            # antes que o XLSX temporário seja descartado pela fonte.
            hash_started = time.perf_counter()
            digest = sha256_file(path)
            verify_digest = getattr(self.source, "verify_download_digest", None)
            if callable(verify_digest):
                verify_digest(path, digest)
            hash_seconds = time.perf_counter() - hash_started
            if report_stages:
                self._report_version_progress(version.number, 70, "Lendo XLSX...")

            read_started = time.perf_counter()
            workbook_reader = getattr(self, "_workbook_reader", None)
            snapshot = (
                workbook_reader.read(path)
                if workbook_reader is not None
                else read_workbook(path)
            )
            read_seconds = time.perf_counter() - read_started
            if report_stages:
                self._report_version_progress(version.number, 90, "Comparando...")
            total_seconds = time.perf_counter() - total_started
            try:
                file_size = path.stat().st_size
            except OSError:
                file_size = -1
            metrics = getattr(workbook_reader, "last_metrics", None)
            logger.info(
                "PERF versao planilha=%s versao=%s bytes=%d download=%.3fs sha256=%.3fs "
                "leitura_xlsx=%.3fs total=%.3fs abas=%s abas_reutilizadas=%s "
                "rows_total=%s rows_reutilizadas=%s rows_parseadas=%s "
                "celulas_reutilizadas=%s celulas_parseadas=%s "
                "tempo_hashes=%.3fs tempo_dependencias=%.3fs "
                "tempo_diff_estrutural=%.3fs tempo_parsing=%.3fs "
                "tempo_snapshot=%.3fs fallback=%s motivo_fallback=%s "
                "sharedstrings_total_anterior=%s sharedstrings_total_atual=%s "
                "sharedstrings_indices_iguais=%s sharedstrings_indices_alterados=%s "
                "sharedstrings_indices_novos=%s sharedstrings_indices_removidos=%s "
                "sharedstrings_shadow_rows_candidate=%s sharedstrings_shadow_rows_safe=%s "
                "sharedstrings_shadow_rows_invalidated=%s "
                "sharedstrings_shadow_indices_checked=%s "
                "sharedstrings_shadow_indices_changed=%s sharedstrings_shadow_indices_new=%s "
                "rows_dependentes_sharedstrings=%s rows_reutilizadas_sharedstrings=%s "
                "rows_invalidada_indice_alterado=%s "
                "rows_invalidada_dependencia_global=%s "
                "rows_invalidada_estrutura_nao_suportada=%s "
                "tempo_sharedstrings_diff=%.6fs tempo_sharedstrings_hash=%.6fs",
                spreadsheet.name,
                version.number,
                file_size,
                download_seconds,
                hash_seconds,
                read_seconds,
                total_seconds,
                metrics.worksheets if metrics else "n/a",
                metrics.worksheets_reused if metrics else "n/a",
                metrics.rows_total if metrics else "n/a",
                metrics.rows_reused if metrics else "n/a",
                metrics.rows_parsed if metrics else "n/a",
                metrics.cells_reused if metrics else "n/a",
                metrics.cells_parsed if metrics else "n/a",
                metrics.hash_seconds if metrics else 0.0,
                metrics.dependency_seconds if metrics else 0.0,
                metrics.structural_diff_seconds if metrics else 0.0,
                metrics.parsing_seconds if metrics else 0.0,
                metrics.snapshot_seconds if metrics else 0.0,
                metrics.fallback_used if metrics else "n/a",
                metrics.fallback_reason if metrics and metrics.fallback_reason else "nenhum",
                metrics.sharedstrings_total_previous if metrics else "n/a",
                metrics.sharedstrings_total_current if metrics else "n/a",
                metrics.sharedstrings_indices_equal if metrics else "n/a",
                metrics.sharedstrings_indices_changed if metrics else "n/a",
                metrics.sharedstrings_indices_new if metrics else "n/a",
                metrics.sharedstrings_indices_removed if metrics else "n/a",
                metrics.sharedstrings_shadow_rows_candidate if metrics else "n/a",
                metrics.sharedstrings_shadow_rows_safe if metrics else "n/a",
                metrics.sharedstrings_shadow_rows_invalidated if metrics else "n/a",
                metrics.sharedstrings_shadow_indices_checked if metrics else "n/a",
                metrics.sharedstrings_shadow_indices_changed if metrics else "n/a",
                metrics.sharedstrings_shadow_indices_new if metrics else "n/a",
                metrics.rows_dependent_sharedstrings if metrics else "n/a",
                metrics.rows_reused_sharedstrings if metrics else "n/a",
                metrics.rows_invalidated_changed_index if metrics else "n/a",
                metrics.rows_invalidated_global_dependency if metrics else "n/a",
                metrics.rows_invalidated_unsupported_structure if metrics else "n/a",
                metrics.sharedstrings_diff_seconds if metrics else 0.0,
                metrics.sharedstrings_hash_seconds if metrics else 0.0,
            )
            return snapshot, digest
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
        normalized = normalize_formula_value(value)
        return None if normalized is None else str(normalized)

    def _persist_comparison(
        self,
        connection: sqlite3.Connection,
        spreadsheet_id: int,
        execution_id: int,
        previous: VersionInfo,
        current: VersionInfo,
        current_hash: str,
        changes: list[CellChange],
        *,
        update_checkpoint: bool = True,
    ) -> None:
        status = (
            ProcessedVersionStatus.PROCESSED
            if changes
            else ProcessedVersionStatus.WITHOUT_CHANGES
        )
        with connection:
            # Segunda linha de defesa para duas execuções que tenham planejado o
            # mesmo par simultaneamente. Um registro final e íntegro é
            # autoritativo e nunca deve ser sobrescrito.
            existing = connection.execute(
                """SELECT id,status,hash_origem FROM versao_processada
                    WHERE planilha_id=? AND versao_anterior_id=? AND versao_atual_id=?""",
                (spreadsheet_id, previous.id, current.id),
            ).fetchone()
            if (existing is not None
                    and existing["status"] in ("PROCESSADA", "SEM_ALTERACOES")
                    and existing["hash_origem"] is not None
                    and len(existing["hash_origem"]) == 64):
                logger.info("PAIR_REUSED previous_id=%s current_id=%s", previous.id, current.id)
                if update_checkpoint:
                    self._advance_planned_checkpoint(connection, spreadsheet_id, previous, current)
                return
            if existing is not None:
                # Registro parcial/falhado não é cobertura. A recomputação o
                # substitui dentro da mesma transação, preservando a UNIQUE.
                connection.execute(
                    "DELETE FROM alteracao WHERE versao_processada_id=?", (existing["id"],)
                )
                connection.execute(
                    "DELETE FROM versao_processada WHERE id=?", (existing["id"],)
                )
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
            if update_checkpoint:
                self._advance_planned_checkpoint(connection, spreadsheet_id, previous, current)

    def _advance_planned_checkpoint(self, connection, spreadsheet_id, previous, current) -> None:
        target = getattr(self, "_checkpoint_targets", {}).get(
            (previous.id, current.id), current
        )
        connection.execute(
            """
            INSERT INTO checkpoint
                (planilha_id, versao_id, versao_numero, data_hora_versao)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(planilha_id) DO UPDATE SET
                versao_id=excluded.versao_id, versao_numero=excluded.versao_numero,
                data_hora_versao=excluded.data_hora_versao,
                data_atualizacao=CURRENT_TIMESTAMP
            """,
            (spreadsheet_id, target.id, target.number, target.modified_at),
        )

    def _configure_coverage_plan(self, coverage: AuditCoverage, interval_mode: bool) -> None:
        """Mapeia cada lacuna ao maior prefixo contínuo confirmado após seu commit."""
        self._checkpoint_targets = {}
        if interval_mode:
            return
        missing = {
            (left.id, right.id) for left, right in coverage.missing_pairs
        }
        requested = coverage.requested_pairs
        for index, pair in enumerate(requested):
            key = (pair[0].id, pair[1].id)
            if key not in missing:
                continue
            target = pair[1]
            for following in requested[index + 1:]:
                following_key = (following[0].id, following[1].id)
                if following_key in missing:
                    break
                target = following[1]
            self._checkpoint_targets[key] = target

    @staticmethod
    def _advance_covered_prefix(connection, spreadsheet_id, coverage: AuditCoverage) -> None:
        """Avança somente sobre o prefixo já coberto, nunca sobre uma lacuna."""
        covered = {(a.id, b.id) for a, b in coverage.covered_pairs}
        target = None
        for previous, current in coverage.requested_pairs:
            if (previous.id, current.id) not in covered:
                break
            target = current
        if target is None:
            return
        connection.execute(
            """INSERT INTO checkpoint
                   (planilha_id,versao_id,versao_numero,data_hora_versao)
               VALUES (?,?,?,?)
               ON CONFLICT(planilha_id) DO UPDATE SET
                   versao_id=excluded.versao_id,
                   versao_numero=excluded.versao_numero,
                   data_hora_versao=excluded.data_hora_versao,
                   data_atualizacao=CURRENT_TIMESTAMP""",
            (spreadsheet_id, target.id, target.number, target.modified_at),
        )
        connection.commit()

    @staticmethod
    def _record_full_coverage(connection, execution_id, coverage: AuditCoverage) -> None:
        total = len(coverage.requested_pairs)
        covered = len(coverage.covered_pairs)
        logger.info(
            "FULL_AUDIT_COVERAGE total_pairs=%d covered_pairs=%d missing_pairs=%d coverage_percent=%.2f",
            total, covered, len(coverage.missing_pairs), coverage.coverage_percent,
        )
        for left, right in coverage.covered_pairs:
            logger.info("PAIR_REUSED previous_id=%s current_id=%s", left.id, right.id)
        for block in coverage.missing_blocks:
            logger.info("MISSING_BLOCK start_id=%s end_id=%s", block.start.id, block.end.id)
        connection.execute(
            """UPDATE execucao_auditoria SET modo='COMPLETA',
               coverage_pairs_total=?, coverage_pairs_existing=?,
               coverage_pairs_missing=?, reused_percent=?,
               pairs_total=?, pairs_reused=?, pairs_missing_before=? WHERE id=?""",
            (total, covered, len(coverage.missing_pairs),
             100.0 * covered / total if total else 100.0,
             total, covered, len(coverage.missing_pairs), execution_id),
        )
        connection.commit()

    @classmethod
    def _finish_coverage(cls, connection, execution_id, processed, duration) -> None:
        cls._finish_interval(connection, execution_id, processed, duration)
        connection.execute(
            """UPDATE execucao_auditoria SET pairs_processed_now=?,
                      coverage_percent_final=coverage_percent WHERE id=?""",
            (processed, execution_id),
        )
        connection.commit()
        row = connection.execute(
            """SELECT coverage_pairs_total,coverage_pairs_existing,
                      coverage_pairs_processed_now,coverage_percent
                 FROM execucao_auditoria WHERE id=?""", (execution_id,),
        ).fetchone()
        logger.info(
            "AUDIT_COVERAGE_COMPLETE total_pairs=%d reused_pairs=%d processed_pairs=%d coverage_percent=%.2f",
            row[0], row[1], row[2], row[3],
        )

    @staticmethod
    def _record_interval_request(connection, execution_id, coverage) -> None:
        requested = coverage.requested_pairs
        start, end = requested[0][0], requested[-1][1]
        total, existing = len(requested), len(coverage.covered_pairs)
        logger.info("INTERVAL_REQUEST start_id=%s end_id=%s", start.id, end.id)
        logger.info("INTERVAL_RESOLVED start=%s end=%s pairs=%d", start.number, end.number, total)
        logger.info("INTERVAL_COVERAGE requested_pairs=%d existing_pairs=%d missing_pairs=%d", total, existing, len(coverage.missing_pairs))
        for left, right in coverage.covered_pairs:
            logger.info("INTERVAL_REUSED_PAIR previous_id=%s current_id=%s", left.id, right.id)
        for left, right in coverage.missing_pairs:
            logger.info("INTERVAL_MISSING_PAIR previous_id=%s current_id=%s", left.id, right.id)
        for block in coverage.missing_blocks:
            logger.info("INTERVAL_MISSING_BLOCK start_id=%s end_id=%s", block.start.id, block.end.id)
        connection.execute(
            """UPDATE execucao_auditoria SET modo='INTERVALO',
               requested_start_version_id=?, requested_start_label=?,
               requested_end_version_id=?, requested_end_label=?,
               coverage_pairs_total=?, coverage_pairs_existing=?,
               coverage_pairs_missing=?, reused_percent=?, pairs_total=?,
               pairs_reused=?, pairs_missing_before=? WHERE id=?""",
            (start.id, start.number, end.id, end.number, total, existing,
             len(coverage.missing_pairs), 100.0 * existing / total,
             total, existing, len(coverage.missing_pairs), execution_id),
        )
        connection.commit()
        logger.info("INTERVAL_PROCESSING_START missing_pairs=%d", len(coverage.missing_pairs))

    @staticmethod
    def _finish_interval(connection, execution_id, processed, duration) -> None:
        row = connection.execute(
            "SELECT coverage_pairs_total,coverage_pairs_existing FROM execucao_auditoria WHERE id=?",
            (execution_id,),
        ).fetchone()
        total, existing = row[0], row[1]
        remaining = max(total - existing - processed, 0)
        result = "REUSED_EXISTING_HISTORY" if processed == 0 and remaining == 0 else "PROCESSED_MISSING_PAIRS"
        connection.execute(
            """UPDATE execucao_auditoria SET coverage_pairs_processed_now=?,
               coverage_pairs_missing=?, coverage_percent=?, interval_duration=?,
               interval_result=? WHERE id=?""",
            (processed, remaining, 100.0 * (total - remaining) / total,
             duration, result, execution_id),
        )
        connection.commit()
        logger.info("INTERVAL_PROCESSING_COMPLETE processed_pairs=%d coverage_percent=%.2f interval_duration=%.3f", processed, 100.0 * (total - remaining) / total, duration)

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
        # O schema histórico restringe os valores persistidos. Uma interrupção
        # graciosa é uma execução concluída (não uma falha); a mensagem preserva
        # a causa sem exigir migração destrutiva da tabela existente.
        persisted_status = (
            AuditExecutionStatus.COMPLETED
            if status is AuditExecutionStatus.STOPPED
            else status
        )
        connection.execute(
            """
            UPDATE execucao_auditoria SET
                fim = CURRENT_TIMESTAMP, versao_final = ?,
                versoes_processadas = ?, alteracoes_encontradas = ?,
                status = ?, mensagem = ?
            WHERE id = ?
            """,
            (final, processed, changes, persisted_status.value, message, execution_id),
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
            initial_checkpoint, final, message,
        )
