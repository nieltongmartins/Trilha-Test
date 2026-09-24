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
import statistics
import tempfile
import threading
import time
from typing import Callable, Literal
from uuid import uuid4

from app.audit_service import AuditResult, AuditService
from app.excel.comparator import CellChange, compare_snapshots
from app.excel.reader import read_workbook
from app.integrity import sha256_file
from app.execution_timing import SharedExecutionTimingModel, TimedStage
from app.models import AuditExecutionStatus
from app.runtime_profile import RuntimeProfileStore, RuntimeSample, workbook_identity
from app.sources.base import SpreadsheetInfo, VersionInfo, VersionSource


logger = logging.getLogger("auditoria_excel.parallel")
MIN_WORKERS = 1
MAX_WORKERS = 8
MEMORY_PRESSURE_PERCENT = 90.0
MEMORY_MIN_AVAILABLE_BYTES = 512 * 1024 * 1024


class TaskState(StrEnum):
    WAITING = "AGUARDANDO"
    RESERVED = "RESERVANDO"
    DOWNLOAD = "BAIXANDO"
    SHA = "VALIDANDO"
    PARSE = "LENDO"
    DEPENDENCY = "AGUARDANDO_DEPENDENCIA"
    COMPARE = "COMPARANDO"
    STAGED = "STAGING"
    COMPLETED = "CONCLUIDO"
    PAUSED = "PAUSADO"
    ERROR = "ERRO"


class AcquisitionState(StrEnum):
    NOT_REQUESTED = "NOT_REQUESTED"
    PREFETCH_PENDING = "PREFETCH_PENDING"
    PREFETCH_READY = "PREFETCH_READY"
    CONSUMING = "CONSUMING"
    CONSUMED = "CONSUMED"
    FAILED = "FAILED"


MIB = 1024 * 1024
DEFAULT_MAX_PREFETCH_BYTES = 64 * MIB
PREFETCH_MIN_TARGET = 2
PREFETCH_MAX_TARGET = 16
# Public compatibility table for UI/benchmark callers.  The scheduler no longer
# uses it as its target; ``prefetch_base_target`` owns that decision.
PREFETCH_TARGET_BY_SLOTS = {slots: 2 * slots for slots in range(MIN_WORKERS, MAX_WORKERS + 1)}


def scheduler_window_for_slots(slots: int) -> int:
    """Reserve room for active comparisons and their complete look-ahead buffer."""
    return slots + PREFETCH_TARGET_BY_SLOTS[slots]


def prefetch_base_target(slots: int, avg_file_bytes: float) -> int:
    """Return a size-aware target without starving the selected worker pool."""
    operational_floor = 2 if slots <= 2 else 3 if slots <= 4 else 4
    if avg_file_bytes < MIB:
        target = min(2 * slots, PREFETCH_MAX_TARGET)
    elif avg_file_bytes < 4 * MIB:
        target = min(slots, 8)
    else:
        target = max(operational_floor, (slots + 1) // 2)
    return max(operational_floor, min(PREFETCH_MAX_TARGET, target))


@dataclass(slots=True)
class AcquisitionRecord:
    version: VersionInfo
    state: AcquisitionState = AcquisitionState.NOT_REQUESTED
    planned_at: float | None = None
    fetch_started_at: float | None = None
    fetch_finished_at: float | None = None
    acquisition_duration: float | None = None
    residual_wait: float = 0.0
    estimated_bytes: int = 0


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
    percent: float
    stage: str
    duration: float = 0.0
    stage_average: float | None = None
    task_average: float | None = None
    estimated_remaining: float | None = None
    learned: bool = False
    occurred_at: float = field(default_factory=time.monotonic)
    timed_stage: TimedStage | None = None
    completed_stages: tuple[TimedStage, ...] = ()
    stage_started_active: float | None = None
    task_started_active: float | None = None


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
    estimated_remaining: float | None = None
    window_occupancy: int = 0
    rss_workers_by_pid: tuple[tuple[int, int], ...] = ()
    rss_total: int = 0
    slot_utilization: tuple[tuple[int, float], ...] = ()
    download_starvation_time: float = 0.0
    cpu_coordinator_percent: float = 0.0
    cpu_workers_percent: float = 0.0
    mean_task: float | None = None
    timing_samples: int = 0
    prefetch_target: int = 0
    prefetch_ready: int = 0
    prefetch_pending: int = 0
    prefetch_hit_rate: float = 0.0
    worker_starvation_count: int = 0
    configured_workers: int = 0
    distinct_worker_count: int = 0
    worker_busy_time: float = 0.0
    worker_idle_time: float = 0.0
    worker_utilization: float = 0.0
    slot_busy_time: float = 0.0
    slot_idle_time: float = 0.0
    prefetch_bytes: int = 0
    prefetch_memory_files: int = 0
    prefetch_temporary_files: int = 0
    prefetch_ready_hit_rate: float = 0.0
    prefetch_pending_hit_rate: float = 0.0
    prefetch_miss_rate: float = 0.0
    ready_ratio: float = 0.0
    pending_ratio: float = 0.0
    prefetch_bytes_pending: int = 0
    prefetch_bytes_ready: int = 0
    slots_processing: int = 0
    slots_waiting: int = 0
    ready_zero_seconds: float = 0.0
    pending_only_seconds: float = 0.0


def _compare_pair(task: ComparisonTask) -> tuple[list[CellChange], dict[str, float]]:
    """Função isolada e serializável executada pelo worker."""
    started = time.perf_counter()
    previous = read_workbook(Path(task.previous_path))
    current = read_workbook(Path(task.current_path))
    parsed = time.perf_counter()
    changes = compare_snapshots(previous, current)
    finished = time.perf_counter()
    return changes, {
        "read_xlsx": parsed - started,
        "compare": finished - parsed,
        "duration": finished - started,
        "worker_pid": os.getpid(),
        "worker_thread_id": threading.get_ident(),
    }


def _initialize_worker() -> None:
    """Emit lifecycle evidence from each executor worker itself."""
    pid = os.getpid()
    logger.info("WORKER_CREATED pid=%d", pid)
    logger.info("WORKER_READY pid=%d", pid)


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
            commit_duration = time.perf_counter() - started
            self.service.timing_model.observe(TimedStage.WAIT_PROMOTION, started - staged_at)
            self.service.timing_model.observe(TimedStage.COMMIT, commit_duration)
            self.service.timing_model.record_commit(now)
            self.commit_times.append(now)
            self.commit_times = [stamp for stamp in self.commit_times if now - stamp <= 60]
            logger.info(
                "TELEMETRIA_PROMOCAO sequence=%d technical_id=%s staging=%.3fs commit_duration=%.3fs",
                self.next_sequence, current.id, started - staged_at,
                commit_duration,
            )
            self.staging.remove(self.next_sequence)
            self.committed += 1
            self.changes += len(changes)
            self.final = current.number
            promoted.append(self.next_sequence)
            self.next_sequence += 1
        return promoted


class ParallelAuditService(AuditService):
    """Scheduler com fila dinâmica e de um a oito workers independentes."""

    def __init__(self, database, source: VersionSource, *, slots: int = 2,
                 backend: Literal["sync", "thread", "process"] = "process",
                 slot_callback: Callable[[SlotProgress], None] | None = None,
                 metrics_callback: Callable[[ParallelMetrics], None] | None = None,
                 staging_directory: str | Path | None = None, max_retries: int = 1,
                 max_prefetch_bytes: int = DEFAULT_MAX_PREFETCH_BYTES,
                 initial_avg_file_bytes: float = 0.0,
                 recommended_prefetch_target: int | None = None,
                 timing_model: SharedExecutionTimingModel | None = None,
                 **kwargs) -> None:
        if not MIN_WORKERS <= slots <= MAX_WORKERS:
            raise ValueError("slots deve estar entre 1 e 8")
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
        self.timing_model = timing_model or SharedExecutionTimingModel()
        self._task_stages: dict[int, tuple[TimedStage, ...]] = {}
        self._last_slot_log: dict[int, tuple[float, TaskState]] = {}
        self._worker_pids: set[int] = set()
        self._created_worker_pids: set[int] = set()
        self._initialized_worker_pids: set[int] = set()
        self._completed_worker_pids: set[int] = set()
        self._worker_busy: dict[int, float] = {}
        self._slot_busy: dict[int, float] = {slot: 0.0 for slot in range(1, slots + 1)}
        self._download_starvation_time = 0.0
        self.max_prefetch_bytes = max_prefetch_bytes
        self._avg_file_bytes = max(0.0, initial_avg_file_bytes)
        calculated_target = prefetch_base_target(slots, self._avg_file_bytes)
        self.prefetch_target = max(
            PREFETCH_MIN_TARGET,
            min(PREFETCH_MAX_TARGET, recommended_prefetch_target or calculated_target),
        )
        self.prefetch_target_initial = self.prefetch_target
        self.prefetch_target_min = self.prefetch_target
        self.prefetch_target_max = self.prefetch_target
        self._target_calibrated = False
        self._last_target_review = time.monotonic()
        self._target_observations = 0
        self._pending_pressure_samples = 0
        self.scheduler_window = scheduler_window_for_slots(slots)
        self._acquisitions: dict[str, AcquisitionRecord] = {}
        self._prefetch_ready_hits = 0
        self._prefetch_pending_hits = 0
        self._prefetch_misses = 0
        self._queue_ages: list[float] = []
        self._active_fetch_times: list[float] = []
        self._residual_waits: list[float] = []
        self._download_normal: list[float] = []
        self._download_prefetch: list[float] = []
        self._worker_starvation_count = 0
        self._slot_states: dict[int, TaskState] = {
            slot: TaskState.WAITING for slot in range(1, slots + 1)
        }
        self._buffer_state_sampled_at = time.monotonic()
        self._ready_zero_seconds = 0.0
        self._pending_only_seconds = 0.0
        self._slot_wait_started: dict[int, tuple[float, int]] = {}
        self._slot_wait_durations: list[float] = []
        configure = getattr(source, "configure_prefetch_buffer", None)
        if callable(configure):
            configure(self.prefetch_target)

    def _slot(self, task: VersionTask, state: TaskState, percent: float, stage: str,
              started: float = 0.0, timed_stage: TimedStage | None = None,
              stage_started: float | None = None) -> None:
        task.state = state
        if task.slot_id:
            self._slot_states[task.slot_id] = state
        duration = time.perf_counter() - started if started else 0.0
        estimate = None
        if timed_stage is not None:
            estimate = self.timing_model.estimate_task(
                timed_stage, max(0.0, time.perf_counter() - (stage_started or started)),
                self._task_stages.get(task.sequence, ()), state is TaskState.COMPLETED,
            )
        now = time.monotonic()
        previous = self._last_slot_log.get(task.slot_id or 0)
        if previous is None or previous[1] is not state or now - previous[0] >= 3.0:
            self._last_slot_log[task.slot_id or 0] = (now, state)
            logger.info(
                "SLOT_PROGRESS slot_id=%d sequence=%s stage=%s elapsed=%.3f "
                "estimated_stage=%.3f task_progress=%.1f eta_task=%.3f",
                task.slot_id or 0, task.sequence, timed_stage.value if timed_stage else state.value,
                duration, estimate.stage_average if estimate else 0.0,
                estimate.progress if estimate else percent,
                estimate.remaining if estimate else 0.0,
            )
        if self.slot_callback:
            self.slot_callback(SlotProgress(
                task.slot_id or 0, task.reservation_token, task.technical_version_id,
                task.version_label, task.sequence, state,
                estimate.progress if estimate else percent, stage, duration,
                estimate.stage_average if estimate else None,
                estimate.task_average if estimate else None,
                estimate.remaining if estimate else None,
                estimate.learned if estimate else False,
                timed_stage=timed_stage,
                completed_stages=self._task_stages.get(task.sequence, ()),
                stage_started_active=(self.timing_model.active_now() - max(0.0, time.perf_counter() - (stage_started or started)))
                if timed_stage is not None else None,
                task_started_active=(self.timing_model.active_now() - duration) if started else None,
            ))

    @property
    def _prefetch_hits(self) -> int:
        """Compatibility aggregate; telemetry always exposes qualified hits."""
        return self._prefetch_ready_hits + self._prefetch_pending_hits

    def _prefetch_counts(self) -> tuple[int, int]:
        ready = sum(r.state is AcquisitionState.PREFETCH_READY for r in self._acquisitions.values())
        pending = sum(r.state is AcquisitionState.PREFETCH_PENDING for r in self._acquisitions.values())
        return ready, pending

    def _buffer_bytes(self) -> tuple[int, int]:
        pending = sum(
            record.estimated_bytes for record in self._acquisitions.values()
            if record.state is AcquisitionState.PREFETCH_PENDING
        )
        ready = sum(
            record.estimated_bytes for record in self._acquisitions.values()
            if record.state is AcquisitionState.PREFETCH_READY
        )
        return pending, ready

    def _set_prefetch_target(self, target: int, reason: str) -> None:
        floor = 2 if self.slots <= 2 else 3 if self.slots <= 4 else 4
        target = max(floor, min(PREFETCH_MAX_TARGET, target))
        if target == self.prefetch_target:
            return
        old = self.prefetch_target
        self.prefetch_target = target
        self.prefetch_target_min = min(self.prefetch_target_min, target)
        self.prefetch_target_max = max(self.prefetch_target_max, target)
        configure = getattr(self.source, "configure_prefetch_buffer", None)
        if callable(configure):
            configure(target)
        logger.info("PREFETCH_TARGET_CHANGED old=%d new=%d reason=%s", old, target, reason)

    def _review_prefetch_target(self) -> None:
        """Apply a slow controller, rather than reacting to every acquisition."""
        self._target_observations += 1
        now = time.monotonic()
        if self._target_observations < 10 and now - self._last_target_review < 30.0:
            return
        self._target_observations = 0
        self._last_target_review = now
        ready, pending = self._prefetch_counts()
        sampled = time.monotonic()
        interval = max(0.0, sampled - self._buffer_state_sampled_at)
        self._buffer_state_sampled_at = sampled
        if ready == 0:
            self._ready_zero_seconds += interval
            if pending:
                self._pending_only_seconds += interval
        avg_queue = statistics.fmean(self._queue_ages[-20:]) if self._queue_ages else 0.0
        avg_residual = statistics.fmean(self._residual_waits[-20:]) if self._residual_waits else 0.0
        avg_fetch = statistics.fmean(self._active_fetch_times[-20:]) if self._active_fetch_times else 0.0
        pressured = pending > ready * 2 or avg_queue > 30.0 or avg_residual > 10.0 or avg_fetch > 30.0
        self._pending_pressure_samples = self._pending_pressure_samples + 1 if pressured else 0
        if self._pending_pressure_samples >= 2:
            self._set_prefetch_target(self.prefetch_target - 1, "pending_or_latency_pressure")
            self._pending_pressure_samples = 0
        elif (ready == 0 and self._worker_starvation_count and avg_fetch and avg_fetch < 30.0
              and sum(self._buffer_bytes()) < self.max_prefetch_bytes // 2):
            self._set_prefetch_target(self.prefetch_target + 1, "ready_empty_fast_download")

    def _refill_prefetch(self, spreadsheet: SpreadsheetInfo,
                         candidates: list[VersionInfo]) -> None:
        """Reabastece somente dentro da janela entregue pelo scheduler.

        Este método roda exclusivamente no coordenador/proprietário do WebDriver.
        O mapa por ID técnico torna o agendamento idempotente.
        """
        prefetch = getattr(self.source, "prefetch_version", None)
        if not callable(prefetch):
            return
        if self._memory_pressure():
            logger.warning(
                "MEMORY_BACKPRESSURE slots=%d action=hold_prefetch buffer_occupied=%d",
                self.slots, sum(self._prefetch_counts()),
            )
            return
        ready, pending = self._prefetch_counts()
        known_sizes = [int(item.size) for item in candidates if isinstance(item.size, int) and item.size > 0]
        if known_sizes:
            self._avg_file_bytes = statistics.fmean(known_sizes)
            base = prefetch_base_target(self.slots, self._avg_file_bytes)
            if not self._target_calibrated:
                self._set_prefetch_target(base, "average_file_size")
                self.prefetch_target_initial = base
                self.prefetch_target_min = base
                self.prefetch_target_max = base
                self._target_calibrated = True
        pending_bytes, ready_bytes = self._buffer_bytes()
        for version in candidates:
            if ready + pending >= self.prefetch_target:
                break
            estimate = int(version.size or self._avg_file_bytes or 0)
            if pending_bytes + ready_bytes + estimate > self.max_prefetch_bytes:
                logger.info(
                    "PREFETCH_BYTES_LIMIT bytes_total=%d candidate_bytes=%d max_prefetch_bytes=%d",
                    pending_bytes + ready_bytes, estimate, self.max_prefetch_bytes,
                )
                break
            record = self._acquisitions.setdefault(version.id, AcquisitionRecord(version))
            if record.state is not AcquisitionState.NOT_REQUESTED:
                continue
            logger.info("PREFETCH_PLANEJADO technical_version_id=%s VersionLabel=%s", version.id, version.number)
            try:
                if not prefetch(spreadsheet, version):
                    continue
            except Exception:
                record.state = AcquisitionState.FAILED
                logger.warning("PREFETCH_FAILED technical_version_id=%s VersionLabel=%s", version.id, version.number, exc_info=True)
                continue
            record.state = AcquisitionState.PREFETCH_PENDING
            record.planned_at = time.perf_counter()
            record.fetch_started_at = record.planned_at
            record.estimated_bytes = estimate
            pending += 1
            pending_bytes += estimate
            logger.info(
                "PREFETCH_STARTED technical_version_id=%s VersionLabel=%s buffer_occupied=%d buffer_ready=%d buffer_pending=%d",
                version.id, version.number, ready + pending, ready, pending,
            )

    def _acquire(self, spreadsheet: SpreadsheetInfo, version: VersionInfo,
                 slot: int) -> tuple[Path, float, bool]:
        record = self._acquisitions.setdefault(version.id, AcquisitionRecord(version))
        prefetched = record.state in (AcquisitionState.PREFETCH_PENDING, AcquisitionState.PREFETCH_READY)
        if record.state is AcquisitionState.CONSUMING:
            raise RuntimeError(f"aquisição duplicada: technical_version_id={version.id}")
        record.state = AcquisitionState.CONSUMING
        started = time.perf_counter()
        try:
            path = self.source.get_version(spreadsheet, version)
        except Exception:
            record.state = AcquisitionState.FAILED
            raise
        duration = time.perf_counter() - started
        consume_metrics = getattr(self.source, "consume_download_metrics", None)
        source_metrics = consume_metrics(version.id) if callable(consume_metrics) else {}
        record.state = AcquisitionState.CONSUMED
        record.fetch_started_at = float(source_metrics.get("fetch_started_at", record.fetch_started_at or started))
        record.fetch_finished_at = float(source_metrics.get("fetch_finished_at", time.perf_counter()))
        record.acquisition_duration = duration
        residual = float(source_metrics.get("residual_wait", duration if prefetched else 0.0))
        record.residual_wait = residual
        if prefetched:
            age = float(source_metrics.get(
                "prefetch_age", time.perf_counter() - (record.planned_at or started)
            ))
            was_ready = bool(source_metrics.get("was_ready", duration < .1))
            queue_age = float(source_metrics.get("queue_age", age))
            active_fetch = float(source_metrics.get("active_fetch_elapsed", duration))
            self._queue_ages.append(queue_age)
            self._active_fetch_times.append(active_fetch)
            self._residual_waits.append(residual)
            self._download_prefetch.append(duration)
            logger.info(
                "PREFETCH_READY technical_version_id=%s VersionLabel=%s age=%.3f bytes=%s buffer_occupied=%d",
                version.id, version.number, age, source_metrics.get("bytes", 0), sum(self._prefetch_counts()),
            )
            logger.info(
                "PREFETCH_CONSUMIDO technical_version_id=%s VersionLabel=%s age=%.3f wait_residual=%.3f prefetch_pronto=%s",
                version.id, version.number, age, residual, was_ready,
            )
            if was_ready:
                self._prefetch_ready_hits += 1
                logger.info("PREFETCH_READY_HIT technical_version_id=%s VersionLabel=%s", version.id, version.number)
            else:
                self._prefetch_pending_hits += 1
                logger.info("PREFETCH_PENDING_HIT technical_version_id=%s VersionLabel=%s", version.id, version.number)
        else:
            self._prefetch_misses += 1
            self._download_normal.append(duration)
            logger.info("PREFETCH_MISS technical_version_id=%s VersionLabel=%s wait_residual=%.3f", version.id, version.number, duration)
        self._review_prefetch_target()
        return path, duration, prefetched

    def audit(self, spreadsheet: SpreadsheetInfo,
              versions: list[VersionInfo] | tuple[VersionInfo, ...] | None = None) -> AuditResult:
        connection = self.database.connection
        identity = workbook_identity(spreadsheet)
        # Fail safe: banco corrompido ou perfil de versão antiga jamais impede a auditoria.
        profile = RuntimeProfileStore(connection).load(identity)
        if profile is not None:
            self._avg_file_bytes = profile.avg_file_bytes
            target = profile.recommended_prefetch_target
            if target is None:
                target = prefetch_base_target(self.slots, self._avg_file_bytes)
            self.prefetch_target = max(PREFETCH_MIN_TARGET, min(PREFETCH_MAX_TARGET, target))
            self.prefetch_target_initial = self.prefetch_target
            self.prefetch_target_min = self.prefetch_target
            self.prefetch_target_max = self.prefetch_target
            self._target_calibrated = True
            configure = getattr(self.source, "configure_prefetch_buffer", None)
            if callable(configure):
                configure(self.prefetch_target)
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
        coordinator: OrderedCommitCoordinator | None = None
        learnable = False
        completed_normally = False
        acquired: dict[str, Path] = {}
        started = time.perf_counter()
        configured_workers = self.slots if self.backend != "sync" else 1
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
            with executor_type(max_workers=configured_workers, initializer=_initialize_worker) as executor:
                running: dict[Future, tuple[VersionTask, ComparisonTask, float, float]] = {}
                while coordinator.committed < len(pairs):
                    if self.stop_event.is_set():
                        break
                    if self.pause_event.is_set():
                        self.timing_model.pause()
                        if not running:
                            for slot in range(1, self.slots + 1):
                                paused = VersionTask("", "", "", 0, execution_id, slot_id=slot)
                                self._slot(paused, TaskState.PAUSED, 0, "Pausado em boundary seguro")
                            self._report_control("paused", coordinator.final or initial)
                            while self.pause_event.is_set() and not self.stop_event.wait(.05):
                                pass
                            self.timing_model.resume()
                            self._report_control("resumed", coordinator.final or initial)
                        else:
                            self._collect_done(running, staging, block=True)
                            self._promote(coordinator, tasks, len(pairs))
                        continue

                    # The bounded window contains the active slots plus a full
                    # 2x look-ahead buffer.  Refill runs on every scheduler turn,
                    # rather than waiting for the buffer to drain.
                    limit = min(len(tasks), coordinator.committed + self.scheduler_window)
                    window_versions: list[VersionInfo] = []
                    seen_window: set[str] = set()
                    for previous, current in pairs[coordinator.committed:limit]:
                        for candidate in (previous, current):
                            if candidate.id not in acquired and candidate.id not in seen_window:
                                seen_window.add(candidate.id)
                                window_versions.append(candidate)
                    self._refill_prefetch(spreadsheet, window_versions)
                    reserved = {item[0].sequence for item in running.values()}
                    available_slots = [slot for slot in range(1, self.slots + 1)
                                       if slot not in {item[0].slot_id for item in running.values()}]
                    for task in tasks[coordinator.committed:limit]:
                        if not available_slots or task.sequence in reserved or task.state in (TaskState.STAGED, TaskState.COMPLETED):
                            continue
                        slot = available_slots.pop(0)
                        while self._memory_pressure() and not self.stop_event.is_set():
                            logger.warning("MEMORY_BACKPRESSURE slots=%d action=hold_reservations", self.slots)
                            self._metrics(started, coordinator, staging, len(running))
                            if running:
                                break
                            time.sleep(.25)
                        if self._memory_pressure() and running:
                            break
                        task.slot_id = slot
                        wait = self._slot_wait_started.pop(slot, None)
                        if wait is not None:
                            wait_seconds = max(0.0, time.monotonic() - wait[0])
                            self._slot_wait_durations.append(wait_seconds)
                            logger.info(
                                "SLOT_WAIT_END slot_id=%d previous_sequence=%d "
                                "next_sequence=%d wait_reason=WAIT_SCHEDULER wait_seconds=%.3f",
                                slot, wait[1], task.sequence, wait_seconds,
                            )
                        task.reservation_token = uuid4().hex
                        pair = pairs[task.sequence - 1]
                        task_started = time.perf_counter()
                        self._task_stages[task.sequence] = ()
                        self._slot(task, TaskState.RESERVED, 2, "Reserva exclusiva", task_started)
                        for version in pair:
                            if version.id not in acquired:
                                download_started = time.perf_counter()
                                self._slot(task, TaskState.DOWNLOAD, 15, "Download pelo ator WebDriver",
                                           task_started, TimedStage.DOWNLOAD_TRANSFER, download_started)
                                acquired[version.id], download_seconds, was_prefetched = self._acquire(
                                    spreadsheet, version, slot
                                )
                                # Downloader serial: outros slots livres ficaram sem arquivo nesse intervalo.
                                idle_slots = max(0, self.slots - len(running) - 1)
                                if idle_slots and download_seconds > 0:
                                    self._worker_starvation_count += 1
                                self._download_starvation_time += download_seconds * idle_slots
                                self.timing_model.observe(TimedStage.DOWNLOAD_TRANSFER, download_seconds, slot)
                                self._task_stages[task.sequence] = (*self._task_stages[task.sequence], TimedStage.DOWNLOAD_TRANSFER)
                                sha_started = time.perf_counter()
                                self._slot(task, TaskState.SHA, 35, "Validando SHA-256", task_started,
                                           TimedStage.SHA, sha_started)
                                digest = sha256_file(acquired[version.id])
                                sha_seconds = time.perf_counter() - sha_started
                                self.timing_model.observe(TimedStage.SHA, sha_seconds, slot)
                                self._task_stages[task.sequence] = (*self._task_stages[task.sequence], TimedStage.SHA)
                                verify = getattr(self.source, "verify_download_digest", None)
                                if callable(verify):
                                    verify(acquired[version.id], digest)
                                self.telemetry.append({
                                    "slot_id": slot, "task_id": task.reservation_token,
                                    "technical_version_id": version.id,
                                    "sequence": task.sequence, "stage": "DOWNLOAD",
                                    "download_transfer": download_seconds, "sha": sha_seconds,
                                    "download_mode": "prefetch" if was_prefetched else "normal",
                                    # Acquisition cost remains end-to-end even when
                                    # Edge prepared the blob in advance; perceived
                                    # slot starvation is only the residual fetch wait.
                                    "slot_download_wait": (
                                        self._acquisitions[version.id].residual_wait
                                        if was_prefetched else download_seconds
                                    ),
                                    "worker_pid": os.getpid(),
                                    "worker_thread_id": threading.get_ident(),
                                    "coordinator_pid": os.getpid(),
                                })
                        current_hash = sha256_file(acquired[pair[1].id])
                        comparison = ComparisonTask(
                            task.sequence, pair[0], pair[1], str(acquired[pair[0].id]),
                            str(acquired[pair[1].id]), current_hash, task.reservation_token,
                        )
                        stage_started = time.perf_counter()
                        self._slot(task, TaskState.PARSE, 50, "Leitura XLSX independente", task_started,
                                   TimedStage.READ_XLSX, stage_started)
                        future = executor.submit(_compare_pair, comparison)
                        running[future] = (task, comparison, task_started, stage_started)

                    self._collect_done(running, staging, block=not any(f.done() for f in running))
                    self._promote(coordinator, tasks, len(pairs))
                    self._metrics(started, coordinator, staging, len(running))

            if self.stop_event.is_set():
                self.timing_model.stop()
                learnable = True
                return self._stop_execution(connection, execution_id, code, initial,
                                            coordinator.committed, coordinator.changes,
                                            coordinator.final or initial)
            self._finish_execution(connection, execution_id, AuditExecutionStatus.COMPLETED,
                                   coordinator.final, coordinator.committed,
                                   coordinator.changes, None)
            learnable = completed_normally = True
            return AuditResult(code, AuditExecutionStatus.COMPLETED, coordinator.committed,
                               coordinator.changes, initial, coordinator.final)
        except Exception as error:
            connection.rollback()
            return self._record_failure(connection, execution_id, spreadsheet_id, code,
                                        initial, 0, 0, None, None, error)
        finally:
            total_requests = self._prefetch_ready_hits + self._prefetch_pending_hits + self._prefetch_misses
            ready, pending = self._prefetch_counts()
            logger.info(
                "PREFETCH_SUMMARY slots_selected=%d target_initial=%d target_min=%d target_max=%d "
                "target_final=%d avg_file_bytes=%.0f max_prefetch_bytes=%d prefetch_ready=%d "
                "prefetch_pending=%d ready_hit_rate=%.4f pending_hit_rate=%.4f miss_rate=%.4f "
                "ready_ratio=%.4f pending_ratio=%.4f avg_queue_age=%.3f max_queue_age=%.3f "
                "avg_active_fetch=%.3f max_active_fetch=%.3f average_residual_wait=%.3f "
                "download_normal_mean=%.3f download_prefetch_mean=%.3f "
                "worker_starvation_count=%d worker_starvation_time=%.3f",
                self.slots, self.prefetch_target_initial, self.prefetch_target_min,
                self.prefetch_target_max, self.prefetch_target, self._avg_file_bytes,
                self.max_prefetch_bytes, ready, pending,
                self._prefetch_ready_hits / total_requests if total_requests else 0.0,
                self._prefetch_pending_hits / total_requests if total_requests else 0.0,
                self._prefetch_misses / total_requests if total_requests else 0.0,
                ready / max(1, ready + pending), pending / max(1, ready + pending),
                statistics.fmean(self._queue_ages) if self._queue_ages else 0.0,
                max(self._queue_ages, default=0.0),
                statistics.fmean(self._active_fetch_times) if self._active_fetch_times else 0.0,
                max(self._active_fetch_times, default=0.0),
                statistics.fmean(self._residual_waits) if self._residual_waits else 0.0,
                statistics.fmean(self._download_normal) if self._download_normal else 0.0,
                statistics.fmean(self._download_prefetch) if self._download_prefetch else 0.0,
                self._worker_starvation_count, self._download_starvation_time,
            )
            elapsed = max(time.perf_counter() - started, .001)
            worker_busy = sum(self._worker_busy.values())
            worker_capacity = elapsed * configured_workers
            slot_busy = sum(self._slot_busy.values())
            logger.info(
                "WORKER_POOL_SUMMARY slots_selected=%d configured_workers=%d "
                "created_worker_pids=%s initialized_worker_pids=%s workers_that_received_tasks=%s "
                "workers_that_completed_tasks=%s distinct_active_worker_count=%d "
                "worker_busy_time=%.3f worker_idle_time=%.3f worker_utilization=%.4f "
                "slot_busy_time=%.3f slot_idle_time=%.3f scheduler_window=%d",
                self.slots, configured_workers, sorted(self._created_worker_pids),
                sorted(self._initialized_worker_pids), sorted(self._worker_pids),
                sorted(self._completed_worker_pids), len(self._worker_pids),
                worker_busy, max(0.0, worker_capacity - worker_busy),
                worker_busy / worker_capacity, slot_busy,
                max(0.0, elapsed * self.slots - slot_busy), self.scheduler_window,
            )
            if learnable and coordinator is not None:
                staged = [item for item in self.telemetry if item.get("stage") == "STAGED"]
                downloads = [float(item["download_transfer"]) for item in self.telemetry
                             if item.get("stage") == "DOWNLOAD"]
                def average(key: str) -> float:
                    values = [float(item[key]) for item in staged if key in item]
                    return statistics.fmean(values) if values else 0.0
                utilization = worker_busy / worker_capacity if worker_capacity else 0.0
                sample = RuntimeSample(
                    slots=self.slots, prefetch_target=self.prefetch_target,
                    versions_processed=coordinator.committed, elapsed_seconds=elapsed,
                    avg_download=statistics.fmean(downloads) if downloads else 0.0,
                    avg_read_xlsx=average("read_xlsx"), avg_compare=average("compare"),
                    avg_worker_task=average("worker_task_duration"),
                    worker_utilization=utilization, avg_file_bytes=self._avg_file_bytes,
                    completed_normally=completed_normally,
                )
                logger.info(
                    "RUNTIME_BENCHMARK_SUMMARY workbook=%s slots=%d versions_processed=%d "
                    "elapsed=%.3f throughput=%.3f avg_read=%.3f avg_compare=%.3f "
                    "worker_utilization=%.4f prefetch_target=%d", identity, self.slots,
                    coordinator.committed, elapsed, sample.throughput, sample.avg_read_xlsx,
                    sample.avg_compare, utilization, self.prefetch_target,
                )
                try:
                    RuntimeProfileStore(connection).record(identity, sample)
                except sqlite3.Error:
                    logger.warning("Não foi possível persistir runtime profile", exc_info=True)
            # Freeze only the execution clock.  Samples and recent official
            # commits remain available if this spreadsheet is continued.
            self.timing_model.stop()
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
                for task, _comparison, task_started, stage_started in running.values():
                    self._slot(task, TaskState.PARSE, 50, "Leitura XLSX independente",
                               task_started, TimedStage.READ_XLSX, stage_started)
                time.sleep(.05)
        for future in [item for item in running if item.done()]:
            task, comparison, started, _stage_started = running.pop(future)
            try:
                changes, metrics = future.result()
                compare_started = time.perf_counter() - float(metrics["compare"])
                self._task_stages[task.sequence] = tuple(dict.fromkeys((
                    *self._task_stages.get(task.sequence, ()), TimedStage.READ_XLSX,
                )))
                self._slot(task, TaskState.COMPARE, 0, "Comparando snapshots", started,
                           TimedStage.COMPARE, compare_started)
                staging_started = time.perf_counter()
                self._task_stages[task.sequence] = tuple(dict.fromkeys((
                    *self._task_stages[task.sequence], TimedStage.COMPARE,
                )))
                self._slot(task, TaskState.STAGED, 0, "Preparando staging", started,
                           TimedStage.STAGING, staging_started)
                staging.put(comparison, changes, metrics)
                metrics.update({"staging": time.perf_counter() - staging_started, "queue_wait": 0.0})
                self.timing_model.observe(TimedStage.READ_XLSX, metrics["read_xlsx"], task.slot_id)
                self.timing_model.observe(TimedStage.COMPARE, metrics["compare"], task.slot_id)
                self.timing_model.observe(TimedStage.STAGING, metrics["staging"], task.slot_id)
                worker_duration = float(metrics["duration"]) + float(metrics["staging"])
                pipeline_latency = time.perf_counter() - started
                # TOTAL_TASK remains a compatibility series, but is deliberately
                # worker-only so old ETA consumers are no longer polluted by acquisition wait.
                self.timing_model.observe(TimedStage.WORKER_TASK_DURATION, worker_duration, task.slot_id)
                self.timing_model.observe(TimedStage.PIPELINE_LATENCY, pipeline_latency, task.slot_id)
                self.timing_model.observe(TimedStage.TOTAL_TASK, worker_duration, task.slot_id)
                self._slot_busy[task.slot_id or 0] = self._slot_busy.get(task.slot_id or 0, 0.0) + (
                    time.perf_counter() - started
                )
                self._worker_pids.add(int(metrics["worker_pid"]))
                worker_pid = int(metrics["worker_pid"])
                self._created_worker_pids.add(worker_pid)
                self._initialized_worker_pids.add(worker_pid)
                self._completed_worker_pids.add(worker_pid)
                self._worker_busy[worker_pid] = self._worker_busy.get(worker_pid, 0.0) + float(
                    metrics["duration"]
                )
                self.telemetry.append({
                    "slot_id": task.slot_id or 0,
                    "task_id": task.reservation_token or "",
                    "technical_version_id": task.technical_version_id,
                    "sequence": task.sequence, "stage": "STAGED",
                    "coordinator_pid": os.getpid(),
                    "worker_task_duration": worker_duration,
                    "pipeline_latency": pipeline_latency,
                    **metrics,
                })
                self._slot(task, TaskState.STAGED, 0,
                           "Staging concluído — aguardando promoção", started,
                           TimedStage.STAGING, staging_started)
                self._task_stages[task.sequence] = tuple(dict.fromkeys((
                    *self._task_stages[task.sequence], TimedStage.STAGING,
                )))
                logger.info(
                    "TELEMETRIA_SLOT slot_id=%d task_id=%s technical_version_id=%s sequence=%d "
                    "stage=STAGED duration=%.3f read_xlsx=%.3f compare=%.3f worker_pid=%d "
                    "worker_thread_id=%d coordinator_pid=%d retry=%d",
                    task.slot_id, task.reservation_token, task.technical_version_id,
                    task.sequence, metrics["duration"], metrics["read_xlsx"],
                    metrics["compare"], metrics["worker_pid"], metrics["worker_thread_id"],
                    os.getpid(), task.retry_count,
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
            logger.info(
                "SLOT_RELEASED slot_id=%d previous_sequence=%d next_sequence=%s",
                task.slot_id or 0, sequence, sequence + 1 if sequence < total else "none",
            )
            # Confirmation is feedback, never an explanation for idle time.
            self._slot(task, TaskState.WAITING, 0, "Aguardando próxima versão")
            self._slot_wait_started[task.slot_id or 0] = (time.monotonic(), sequence)
            logger.info(
                "SLOT_WAIT_START slot_id=%d previous_sequence=%d next_sequence=%s "
                "wait_reason=WAIT_SCHEDULER wait_seconds=0.000",
                task.slot_id or 0, sequence, sequence + 1 if sequence < total else "none",
            )
            self._report_progress(coordinator.committed, total)
            self._report_checkpoint(task.version_label, coordinator.committed,
                                    total - coordinator.committed)

    def _metrics(self, started: float, coordinator: OrderedCommitCoordinator,
                 staging: StagingStore, active: int) -> None:
        if not self.metrics_callback:
            return
        rss = self._process_rss(os.getpid())
        elapsed = time.perf_counter() - started
        recent_window = min(60.0, elapsed)
        recent = (
            len(coordinator.commit_times) * 60.0 / recent_window
            if recent_window > 0
            else 0.0
        )
        workers_by_pid = tuple((pid, self._process_rss(pid)) for pid in sorted(self._worker_pids))
        worker_rss = sum(value for _pid, value in workers_by_pid)
        edge_rss = self._edge_rss()
        cpu_coordinator, cpu_workers = self._cpu_percentages()
        utilization = tuple(
            (slot, min(100.0, busy / max(elapsed, .001) * 100.0))
            for slot, busy in sorted(self._slot_busy.items())
        )
        prefetch_metrics = self._prefetch_resource_metrics()
        worker_busy = sum(self._worker_busy.values())
        worker_capacity = elapsed * (self.slots if self.backend != "sync" else 1)
        slot_busy = sum(self._slot_busy.values())
        processing_states = {
            TaskState.DOWNLOAD, TaskState.SHA, TaskState.PARSE, TaskState.COMPARE,
            TaskState.STAGED,
        }
        slots_processing = sum(state in processing_states for state in self._slot_states.values())
        slots_waiting = sum(
            state in {TaskState.WAITING, TaskState.RESERVED, TaskState.COMPLETED}
            for state in self._slot_states.values()
        )
        logger.info(
            "TELEMETRIA_MEMORIA slots=%d rss_coordinator=%d rss_workers_total=%d "
            "rss_edge=%d rss_total=%d staged_count=%d window_occupancy=%d "
            "prefetch_bytes=%d prefetch_memory_files=%d prefetch_temporary_files=%d",
            self.slots, rss, worker_rss, edge_rss, rss + worker_rss + edge_rss,
            staging.count(), staging.count() + active,
            prefetch_metrics["bytes"], prefetch_metrics["memory_files"],
            prefetch_metrics["temporary_files"],
        )
        self.metrics_callback(ParallelMetrics(
            elapsed, coordinator.committed, float(recent), staging.count(), active,
            coordinator.committed, self.scheduler_window, rss, worker_rss, edge_rss,
            self.timing_model.global_eta(max(len(coordinator.pairs) - coordinator.committed, 0),
                                         max(active, self.slots)),
            staging.count() + active, workers_by_pid, rss + worker_rss + edge_rss,
            utilization, self._download_starvation_time, cpu_coordinator, cpu_workers,
            (self.timing_model.task_average()
             if self.timing_model.sample_count(TimedStage.TOTAL_TASK) else None),
            self.timing_model.sample_count(TimedStage.TOTAL_TASK),
            self.prefetch_target, *self._prefetch_counts(),
            ((self._prefetch_ready_hits + self._prefetch_pending_hits) /
             max(1, self._prefetch_ready_hits + self._prefetch_pending_hits + self._prefetch_misses)),
            self._worker_starvation_count,
            self.slots if self.backend != "sync" else 1, len(self._worker_pids),
            worker_busy, max(0.0, worker_capacity - worker_busy),
            worker_busy / max(worker_capacity, .001), slot_busy,
            max(0.0, elapsed * self.slots - slot_busy),
            prefetch_metrics["bytes"], prefetch_metrics["memory_files"],
            prefetch_metrics["temporary_files"],
            self._prefetch_ready_hits / max(1, self._prefetch_ready_hits + self._prefetch_pending_hits + self._prefetch_misses),
            self._prefetch_pending_hits / max(1, self._prefetch_ready_hits + self._prefetch_pending_hits + self._prefetch_misses),
            self._prefetch_misses / max(1, self._prefetch_ready_hits + self._prefetch_pending_hits + self._prefetch_misses),
            self._prefetch_counts()[0] / max(1, sum(self._prefetch_counts())),
            self._prefetch_counts()[1] / max(1, sum(self._prefetch_counts())),
            self._buffer_bytes()[0], self._buffer_bytes()[1],
            slots_processing, slots_waiting, self._ready_zero_seconds,
            self._pending_only_seconds,
        ))

    def _prefetch_resource_metrics(self) -> dict[str, int]:
        report = getattr(self.source, "prefetch_resource_metrics", None)
        if callable(report):
            values = report()
            return {
                "bytes": int(values.get("bytes", 0)),
                "memory_files": int(values.get("memory_files", 0)),
                "temporary_files": int(values.get("temporary_files", 0)),
            }
        return {"bytes": 0, "memory_files": 0, "temporary_files": 0}

    @staticmethod
    def _process_rss(pid: int) -> int:
        try:
            import psutil
            return int(psutil.Process(pid).memory_info().rss)
        except (ImportError, OSError):
            try:
                fields = Path(f"/proc/{pid}/statm").read_text(encoding="ascii").split()
                return int(fields[1]) * int(os.sysconf("SC_PAGE_SIZE"))
            except (OSError, IndexError, ValueError):
                return 0

    @staticmethod
    def _edge_rss() -> int:
        try:
            import psutil
            return sum(
                int(process.info["memory_info"].rss)
                for process in psutil.process_iter(("name", "memory_info"))
                if "edge" in (process.info["name"] or "").lower()
            )
        except (ImportError, OSError):
            return 0

    def _cpu_percentages(self) -> tuple[float, float]:
        try:
            import psutil
            coordinator = psutil.Process(os.getpid()).cpu_percent(None)
            workers = sum(psutil.Process(pid).cpu_percent(None) for pid in self._worker_pids)
            return float(coordinator), float(workers)
        except (ImportError, OSError):
            return 0.0, 0.0

    @staticmethod
    def _memory_pressure() -> bool:
        try:
            import psutil
            memory = psutil.virtual_memory()
            return memory.percent >= MEMORY_PRESSURE_PERCENT or memory.available < MEMORY_MIN_AVAILABLE_BYTES
        except ImportError:
            return False
