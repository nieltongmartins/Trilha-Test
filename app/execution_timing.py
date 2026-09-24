"""Modelo observacional compartilhado de tempo, progresso e ETA da auditoria."""

from __future__ import annotations

from collections import defaultdict, deque
from dataclasses import dataclass
from enum import StrEnum
import math
import statistics
import threading
import time


class TimedStage(StrEnum):
    DOWNLOAD_FETCH = "DOWNLOAD_FETCH"
    DOWNLOAD_TRANSFER = "DOWNLOAD_TRANSFER"
    SHA = "SHA"
    READ_XLSX = "READ_XLSX"
    COMPARE = "COMPARE"
    STAGING = "STAGING"
    WAIT_PROMOTION = "WAIT_PROMOTION"
    COMMIT = "COMMIT"
    TOTAL_TASK = "TOTAL_TASK"


OPERATIONAL_STAGES = (
    TimedStage.DOWNLOAD_FETCH, TimedStage.DOWNLOAD_TRANSFER, TimedStage.SHA,
    TimedStage.READ_XLSX, TimedStage.COMPARE, TimedStage.STAGING,
)

# Baseline visual conservador usado somente até a primeira observação compartilhada.
DEFAULT_BASELINE = {
    TimedStage.DOWNLOAD_FETCH: 1.0,
    TimedStage.DOWNLOAD_TRANSFER: 7.0,
    TimedStage.SHA: 0.15,
    TimedStage.READ_XLSX: 9.0,
    TimedStage.COMPARE: 1.2,
    TimedStage.STAGING: 0.1,
    TimedStage.WAIT_PROMOTION: 0.2,
    TimedStage.COMMIT: 0.05,
    TimedStage.TOTAL_TASK: 18.45,
}


@dataclass(frozen=True, slots=True)
class TimingEstimate:
    progress: float
    stage_average: float
    task_average: float
    remaining: float
    learned: bool


class SharedExecutionTimingModel:
    """Thread-safe rolling model shared by every slot.

    ETA uses a median/MAD winsorized mean of the latest ``window_size`` values.
    Raw values are retained separately so outliers remain available to telemetry.
    """

    def __init__(self, window_size: int = 20, baseline: dict[TimedStage, float] | None = None,
                 commit_window: int = 20) -> None:
        if window_size < 1:
            raise ValueError("window_size deve ser positivo")
        self.window_size = window_size
        self.baseline = {**DEFAULT_BASELINE, **(baseline or {})}
        self._samples = {stage: deque(maxlen=window_size) for stage in TimedStage}
        self.raw_observations: list[tuple[TimedStage, float, int | None]] = []
        self._commit_times: deque[float] = deque(maxlen=commit_window)
        self._lock = threading.RLock()
        self._paused_at: float | None = None
        self._paused_total = 0.0
        self._stopped_at: float | None = None

    def active_now(self, now: float | None = None) -> float:
        now = time.monotonic() if now is None else now
        with self._lock:
            end = self._stopped_at if self._stopped_at is not None else (
                self._paused_at if self._paused_at is not None else now)
            return end - self._paused_total

    def pause(self, now: float | None = None) -> None:
        with self._lock:
            if self._paused_at is None and self._stopped_at is None:
                self._paused_at = time.monotonic() if now is None else now

    def resume(self, now: float | None = None) -> None:
        now = time.monotonic() if now is None else now
        with self._lock:
            if self._paused_at is not None:
                self._paused_total += max(0.0, now - self._paused_at)
                self._paused_at = None

    def stop(self, now: float | None = None) -> None:
        with self._lock:
            instant = time.monotonic() if now is None else now
            self._stopped_at = self._paused_at if self._paused_at is not None else instant

    def observe(self, stage: TimedStage | str, seconds: float, slot_id: int | None = None) -> None:
        stage = TimedStage(stage)
        value = float(seconds)
        if not math.isfinite(value) or value < 0:
            return
        with self._lock:
            self.raw_observations.append((stage, value, slot_id))
            self._samples[stage].append(value)

    def samples(self, stage: TimedStage | str) -> tuple[float, ...]:
        with self._lock:
            return tuple(self._samples[TimedStage(stage)])

    def has_samples(self, stage: TimedStage | str) -> bool:
        with self._lock:
            return bool(self._samples[TimedStage(stage)])

    def average(self, stage: TimedStage | str) -> float:
        stage = TimedStage(stage)
        with self._lock:
            values = list(self._samples[stage])
        if not values:
            return self.baseline[stage]
        if len(values) < 3:
            return statistics.fmean(values)
        median = statistics.median(values)
        deviations = [abs(value - median) for value in values]
        mad = statistics.median(deviations)
        if mad == 0:
            cap = max(median * 3.0, median + 1.0)
            bounded = [min(value, cap) for value in values]
        else:
            spread = 3.0 * 1.4826 * mad
            bounded = [min(max(value, median - spread), median + spread) for value in values]
        return statistics.fmean(bounded)

    def task_average(self) -> float:
        if self.has_samples(TimedStage.TOTAL_TASK):
            return self.average(TimedStage.TOTAL_TASK)
        if not any(self.has_samples(stage) for stage in OPERATIONAL_STAGES):
            return self.baseline[TimedStage.TOTAL_TASK]
        return sum(self.average(stage) for stage in OPERATIONAL_STAGES)

    def estimate_task(self, current_stage: TimedStage | str, elapsed_in_stage: float,
                      completed: tuple[TimedStage, ...] = (), finished: bool = False) -> TimingEstimate:
        stage = TimedStage(current_stage)
        weights = {item: self.average(item) for item in OPERATIONAL_STAGES}
        total = max(sum(weights.values()), .001)
        completed_set = set(completed)
        completed_weight = sum(weights[item] for item in OPERATIONAL_STAGES if item in completed_set)
        expected = weights.get(stage, self.average(stage))
        ratio = max(0.0, elapsed_in_stage) / max(expected, .001)
        # 90% da fatia no tempo esperado; cauda assintótica nunca confirma a etapa.
        fraction = .9 * ratio if ratio <= 1 else .9 + .1 * (1 - math.exp(-(ratio - 1)))
        fraction = min(fraction, .999)
        progress = 100.0 if finished else min(99.9, 100 * (completed_weight + expected * fraction) / total)
        remaining_current = expected * max(0.0, 1.0 - fraction)
        try:
            index = OPERATIONAL_STAGES.index(stage)
            future = sum(weights[item] for item in OPERATIONAL_STAGES[index + 1:])
        except ValueError:
            future = 0.0
        return TimingEstimate(progress, expected, total, 0.0 if finished else remaining_current + future,
                              self.has_samples(stage))

    def record_commit(self, now: float | None = None) -> None:
        with self._lock:
            self._commit_times.append(time.monotonic() if now is None else now)

    def throughput_per_minute(self) -> float | None:
        with self._lock:
            times = tuple(self._commit_times)
        if len(times) < 2 or times[-1] <= times[0]:
            return None
        return (len(times) - 1) * 60.0 / (times[-1] - times[0])

    def global_eta(self, remaining_tasks: int, active_slots: int) -> float:
        if remaining_tasks <= 0:
            return 0.0
        throughput = self.throughput_per_minute()
        if throughput is not None and len(self._commit_times) >= 3:
            return remaining_tasks / throughput * 60.0
        return remaining_tasks * self.task_average() / max(1, active_slots)
