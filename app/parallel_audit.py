"""Protótipo Fase 2: duas comparações físicas e promoção ordenada."""

from __future__ import annotations

from concurrent.futures import Future, ProcessPoolExecutor, ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import StrEnum
import json
import logging
import os
from pathlib import Path
import pickle
import sqlite3
import tempfile
import threading
import time
from typing import Callable, Literal
from uuid import uuid4

from app.audit_service import AuditResult, AuditService
from app.excel.comparator import CellChange, compare_snapshots
from app.excel.reader import read_workbook
from app.integrity import sha256_file
from app.models import AuditExecutionStatus
from app.sources.base import SpreadsheetInfo, VersionInfo, VersionSource


logger = logging.getLogger("auditoria_excel.parallel")


class TaskState(StrEnum):
    WAITING = "AGUARDANDO"
    RESERVED = "RESERVADO"
    DOWNLOAD = "DOWNLOAD"
    SHA = "VALIDACAO_SHA"
    PARSE = "LEITURA_XLSX"
    DEPENDENCY = "AGUARDANDO_DEPENDENCIA"
    COMPARE = "COMPARACAO"
    STAGED = "STAGED"
    COMPLETED = "CONCLUIDO"
    PAUSED = "PAUSADO"
    ERROR = "ERRO"


@dataclass(slots=True)
class VersionTask:
    technical_version_id: str
    version_label: str
    predecessor_technical_id: str
    sequence: int
    execution_id: int
    state: TaskState = TaskState.WAITING
    slot_id: int | None = None
    reservation_token: str | None = None
    retry_count: int = 0


@dataclass(frozen=True, slots=True)
class ComparisonTask:
    sequence: int
    previous: VersionInfo
    current: VersionInfo
    previous_path: str
    current_path: str
    current_hash: str
    reservation_token: str


@dataclass(frozen=True, slots=True)
class PromotionTask:
    sequence: int
    previous: VersionInfo
    current: VersionInfo


@dataclass(frozen=True, slots=True)
class SlotProgress:
    slot_id: int
    task_id: str | None
    technical_version_id: str | None
    version: str | None
    sequence: int | None
    state: TaskState
    percent: int
    stage: str
    duration: float = 0.0
    occurred_at: float = field(default_factory=time.monotonic)


@dataclass(frozen=True, slots=True)
class ParallelMetrics:
    elapsed: float
    committed: int
    throughput_recent: float
    staged_count: int
    active_slots: int
    checkpoint_sequence: int
    window_size: int
    rss_python: int
    rss_workers: int = 0
    rss_edge: int = 0


def _compare_pair(task: ComparisonTask) -> tuple[list[CellChange], dict[str, float]]:
    """Função isolada e serializável executada pelo worker."""
    started = time.perf_counter()
    previous = read_workbook(Path(task.previous_path))
    current = read_workbook(Path(task.current_path))
    parsed = time.perf_counter()
    changes = compare_snapshots(previous, current)
    finished = time.perf_counter()
    return changes, {
        "parse": parsed - started,
        "compare": finished - parsed,
        "duration": finished - started,
    }


class StagingStore:
    """SQLite descartável; nunca é fonte de verdade oficial."""

    def __init__(self, directory: Path, execution_code: str) -> None:
        directory.mkdir(parents=True, exist_ok=True)
        for orphan in directory.glob("parallel-*.sqlite3"):
            orphan.unlink(missing_ok=True)
        self.path = directory / f"parallel-{execution_code}.sqlite3"
        self.connection = sqlite3.connect(self.path)
        self.connection.execute(
            """CREATE TABLE exp_result (
                sequence INTEGER PRIMARY KEY, predecessor_id TEXT NOT NULL,
                technical_id TEXT NOT NULL, reservation_token TEXT NOT NULL,
                current_hash TEXT NOT NULL, changes BLOB NOT NULL,
                metrics TEXT NOT NULL, staged_at REAL NOT NULL, complete INTEGER NOT NULL
            )"""
        )
        self.connection.commit()

    def put(self, task: ComparisonTask, changes: list[CellChange], metrics: dict[str, float]) -> None:
        with self.connection:
            self.connection.execute(
                "INSERT INTO exp_result VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)",
                (task.sequence, task.previous.id, task.current.id,
                 task.reservation_token, task.current_hash,
                 sqlite3.Binary(pickle.dumps(changes, protocol=5)),
                 json.dumps(metrics), time.monotonic()),
            )

    def get(self, sequence: int) -> tuple[list[CellChange], str, float] | None:
        row = self.connection.execute(
            "SELECT changes, current_hash, staged_at FROM exp_result WHERE sequence=? AND complete=1",
            (sequence,),
        ).fetchone()
        return None if row is None else (pickle.loads(row[0]), row[1], row[2])

    def count(self) -> int:
        return int(self.connection.execute("SELECT count(*) FROM exp_result").fetchone()[0])

    def remove(self, sequence: int) -> None:
        with self.connection:
            self.connection.execute("DELETE FROM exp_result WHERE sequence=?", (sequence,))

    def close(self) -> None:
        self.connection.close()
        self.path.unlink(missing_ok=True)


class OrderedCommitCoordinator:
    """Único componente autorizado a persistir e avançar o checkpoint."""

    def __init__(self, service: "ParallelAuditService", staging: StagingStore,
                 connection: sqlite3.Connection, spreadsheet_id: int,
                 execution_id: int, pairs: list[tuple[VersionInfo, VersionInfo]]) -> None:
        self.service = service
        self.staging = staging
        self.connection = connection
        self.spreadsheet_id = spreadsheet_id
        self.execution_id = execution_id
        self.pairs = pairs
        self.next_sequence = 1
        self.committed = 0
        self.changes = 0
        self.final: str | None = None
        self.commit_times: list[float] = []

    def promote_available(self) -> list[int]:
        promoted: list[int] = []
        while self.next_sequence <= len(self.pairs):
            staged = self.staging.get(self.next_sequence)
            if staged is None:
                break
            changes, digest, staged_at = staged
            previous, current = self.pairs[self.next_sequence - 1]
            started = time.perf_counter()
            # _persist_comparison encapsula comparação, alterações e checkpoint
            # no mesmo `with connection`, com rollback automático em falha.
            self.service._persist_comparison(
                self.connection, self.spreadsheet_id, self.execution_id,
                previous, current, digest, changes,
            )
            now = time.monotonic()
            self.commit_times.append(now)
            self.commit_times = [stamp for stamp in self.commit_times if now - stamp <= 60]
            logger.info(
                "TELEMETRIA_PROMOCAO sequence=%d technical_id=%s staging=%.3fs commit_duration=%.3fs",
                self.next_sequence, current.id, started - staged_at,
                time.perf_counter() - started,
            )
            self.staging.remove(self.next_sequence)
            self.committed += 1
            self.changes += len(changes)
            self.final = current.number
            promoted.append(self.next_sequence)
            self.next_sequence += 1
        return promoted


class ParallelAuditService(AuditService):
    """Scheduler central limitado deliberadamente a um ou dois slots."""

    def __init__(self, database, source: VersionSource, *, slots: Literal[1, 2] = 2,
                 backend: Literal["sync", "thread", "process"] = "process",
                 slot_callback: Callable[[SlotProgress], None] | None = None,
                 metrics_callback: Callable[[ParallelMetrics], None] | None = None,
                 staging_directory: str | Path | None = None, max_retries: int = 1,
                 **kwargs) -> None:
        if slots not in (1, 2):
            raise ValueError("A Fase 2 aceita exclusivamente 1 ou 2 slots")
        super().__init__(database, source, **kwargs)
        self.slots = slots
        self.backend = backend
        self.slot_callback = slot_callback
        self.metrics_callback = metrics_callback
        self.max_retries = max_retries
        self.staging_directory = (
            Path(staging_directory)
            if staging_directory is not None
            else Path(tempfile.gettempdir()) / "trilha-staging"
        )
        self.telemetry: list[dict[str, float | int | str]] = []

    def _slot(self, task: VersionTask, state: TaskState, percent: int, stage: str,
              started: float = 0.0) -> None:
        task.state = state
        if self.slot_callback:
            self.slot_callback(SlotProgress(
                task.slot_id or 0, task.reservation_token, task.technical_version_id,
                task.version_label, task.sequence, state, percent, stage,
                time.perf_counter() - started if started else 0.0,
            ))

    def audit(self, spreadsheet: SpreadsheetInfo,
              versions: list[VersionInfo] | tuple[VersionInfo, ...] | None = None) -> AuditResult:
        connection = self.database.connection
        spreadsheet_id = self._upsert_spreadsheet(connection, spreadsheet)
        checkpoint = connection.execute(
            "SELECT versao_id, versao_numero FROM checkpoint WHERE planilha_id=?",
            (spreadsheet_id,),
        ).fetchone()
        initial = checkpoint["versao_numero"] if checkpoint else None
        code = f"AUD-P2-{datetime.now(timezone.utc):%Y%m%d%H%M%S}-{uuid4().hex[:8]}"
        execution_id = connection.execute(
            "INSERT INTO execucao_auditoria (codigo_execucao,planilha_id,checkpoint_inicial,status) VALUES (?,?,?,?)",
            (code, spreadsheet_id, initial, AuditExecutionStatus.RUNNING.value),
        ).lastrowid
        connection.commit()
        staging: StagingStore | None = None
        acquired: dict[str, Path] = {}
        started = time.perf_counter()
        try:
            if versions is None:
                try:
                    versions = list(self.source.list_versions(
                        spreadsheet, checkpoint_id=checkpoint["versao_id"] if checkpoint else None,
                        checkpoint_label=initial,
                    ))
                except TypeError:
                    versions = list(self.source.list_versions(spreadsheet))
            pairs = self._pending_pairs(list(versions), checkpoint["versao_id"] if checkpoint else None)
            self._report_progress(0, len(pairs))
            if not pairs:
                self._finish_execution(connection, execution_id,
                    AuditExecutionStatus.COMPLETED_WITHOUT_UPDATES, initial, 0, 0, None)
                return AuditResult(code, AuditExecutionStatus.COMPLETED_WITHOUT_UPDATES, 0, 0, initial, initial)

            staging = StagingStore(self.staging_directory, code)
            coordinator = OrderedCommitCoordinator(
                self, staging, connection, spreadsheet_id, execution_id, pairs
            )
            tasks = [VersionTask(cur.id, cur.number, prev.id, index, execution_id)
                     for index, (prev, cur) in enumerate(pairs, 1)]
            executor_type = ProcessPoolExecutor if self.backend == "process" else ThreadPoolExecutor
            workers = self.slots if self.backend != "sync" else 1
            with executor_type(max_workers=workers) as executor:
                running: dict[Future, tuple[VersionTask, ComparisonTask, float]] = {}
                while coordinator.committed < len(pairs):
                    if self.stop_event.is_set():
                        break
                    if self.pause_event.is_set():
                        if not running:
                            for slot in range(1, self.slots + 1):
                                paused = VersionTask("", "", "", 0, execution_id, slot_id=slot)
                                self._slot(paused, TaskState.PAUSED, 0, "Pausado em boundary seguro")
                            self._report_control("paused", coordinator.final or initial)
                            while self.pause_event.is_set() and not self.stop_event.wait(.05):
                                pass
                            self._report_control("resumed", coordinator.final or initial)
                        else:
                            self._collect_done(running, staging, block=True)
                            self._promote(coordinator, tasks, len(pairs))
                        continue

                    limit = min(len(tasks), coordinator.committed + 2 * self.slots)
                    reserved = {item[0].sequence for item in running.values()}
                    available_slots = [slot for slot in range(1, self.slots + 1)
                                       if slot not in {item[0].slot_id for item in running.values()}]
                    for task in tasks[coordinator.committed:limit]:
                        if not available_slots or task.sequence in reserved or task.state in (TaskState.STAGED, TaskState.COMPLETED):
                            continue
                        slot = available_slots.pop(0)
                        task.slot_id = slot
                        task.reservation_token = uuid4().hex
                        pair = pairs[task.sequence - 1]
                        task_started = time.perf_counter()
                        self._slot(task, TaskState.RESERVED, 2, "Reserva exclusiva", task_started)
                        for version in pair:
                            if version.id not in acquired:
                                self._slot(task, TaskState.DOWNLOAD, 15, "Download pelo ator WebDriver", task_started)
                                download_started = time.perf_counter()
                                acquired[version.id] = self.source.get_version(spreadsheet, version)
                                download_seconds = time.perf_counter() - download_started
                                self._slot(task, TaskState.SHA, 35, "Validando SHA-256", task_started)
                                digest = sha256_file(acquired[version.id])
                                verify = getattr(self.source, "verify_download_digest", None)
                                if callable(verify):
                                    verify(acquired[version.id], digest)
                                self.telemetry.append({
                                    "slot_id": slot, "task_id": task.reservation_token,
                                    "technical_version_id": version.id,
                                    "sequence": task.sequence, "stage": "DOWNLOAD",
                                    "download": download_seconds,
                                })
                        current_hash = sha256_file(acquired[pair[1].id])
                        comparison = ComparisonTask(
                            task.sequence, pair[0], pair[1], str(acquired[pair[0].id]),
                            str(acquired[pair[1].id]), current_hash, task.reservation_token,
                        )
                        self._slot(task, TaskState.PARSE, 50, "Leitura XLSX independente", task_started)
                        future = executor.submit(_compare_pair, comparison)
                        running[future] = (task, comparison, task_started)

                    self._collect_done(running, staging, block=not any(f.done() for f in running))
                    self._promote(coordinator, tasks, len(pairs))
                    self._metrics(started, coordinator, staging, len(running))

            if self.stop_event.is_set():
                return self._stop_execution(connection, execution_id, code, initial,
                                            coordinator.committed, coordinator.changes,
                                            coordinator.final or initial)
            self._finish_execution(connection, execution_id, AuditExecutionStatus.COMPLETED,
                                   coordinator.final, coordinator.committed,
                                   coordinator.changes, None)
            return AuditResult(code, AuditExecutionStatus.COMPLETED, coordinator.committed,
                               coordinator.changes, initial, coordinator.final)
        except Exception as error:
            connection.rollback()
            return self._record_failure(connection, execution_id, spreadsheet_id, code,
                                        initial, 0, 0, None, None, error)
        finally:
            if staging is not None:
                staging.close()
            for path in acquired.values():
                try:
                    self.source.release_version(path)
                except Exception:
                    logger.warning("Falha ao liberar download %s", path, exc_info=True)

    def _collect_done(self, running, staging: StagingStore, *, block: bool) -> None:
        if block and running:
            while not any(future.done() for future in running):
                time.sleep(.01)
        for future in [item for item in running if item.done()]:
            task, comparison, started = running.pop(future)
            try:
                changes, metrics = future.result()
                self._slot(task, TaskState.COMPARE, 90, "Comparando snapshots", started)
                staging.put(comparison, changes, metrics)
                metrics.update({"staging": 0.0, "queue_wait": 0.0})
                self.telemetry.append({
                    "slot_id": task.slot_id or 0,
                    "task_id": task.reservation_token or "",
                    "technical_version_id": task.technical_version_id,
                    "sequence": task.sequence, "stage": "STAGED",
                    **metrics,
                })
                self._slot(task, TaskState.STAGED, 100, "Concluída — aguardando promoção", started)
                logger.info(
                    "TELEMETRIA_SLOT slot_id=%d task_id=%s technical_version_id=%s sequence=%d "
                    "stage=STAGED duration=%.3f parse=%.3f compare=%.3f retry=%d",
                    task.slot_id, task.reservation_token, task.technical_version_id,
                    task.sequence, metrics["duration"], metrics["parse"],
                    metrics["compare"], task.retry_count,
                )
            except Exception as error:
                task.retry_count += 1
                task.state = TaskState.WAITING if task.retry_count <= self.max_retries else TaskState.ERROR
                self._slot(task, task.state, 0, f"Worker falhou: {error}", started)
                if task.state is TaskState.ERROR:
                    raise RuntimeError(
                        f"slot={task.slot_id} technical_id={task.technical_version_id} "
                        f"VersionLabel={task.version_label} etapa={task.state} "
                        f"retry={task.retry_count}: {error}"
                    ) from error

    def _promote(self, coordinator: OrderedCommitCoordinator,
                 tasks: list[VersionTask], total: int) -> None:
        for sequence in coordinator.promote_available():
            task = tasks[sequence - 1]
            self._slot(task, TaskState.COMPLETED, 100, "Checkpoint confirmado")
            self._report_progress(coordinator.committed, total)
            self._report_checkpoint(task.version_label, coordinator.committed,
                                    total - coordinator.committed)

    def _metrics(self, started: float, coordinator: OrderedCommitCoordinator,
                 staging: StagingStore, active: int) -> None:
        if not self.metrics_callback:
            return
        rss = 0
        try:
            import resource
            rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss * 1024
        except (ImportError, OSError):
            pass
        elapsed = time.perf_counter() - started
        recent_window = min(60.0, elapsed)
        recent = (
            len(coordinator.commit_times) * 60.0 / recent_window
            if recent_window > 0
            else 0.0
        )
        worker_rss = 0
        try:
            worker_rss = resource.getrusage(resource.RUSAGE_CHILDREN).ru_maxrss * 1024
        except (NameError, OSError):
            pass
        self.metrics_callback(ParallelMetrics(
            elapsed, coordinator.committed, float(recent), staging.count(), active,
            coordinator.committed, 2 * self.slots, rss, worker_rss, 0,
        ))
