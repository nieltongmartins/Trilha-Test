"""Modelo observacional compartilhado de tempo, progresso e ETA da auditoria."""

from __future__ import annotations

from collections import defaultdict, deque
from dataclasses import dataclass
from enum import StrEnum
import math
import logging
import statistics
import threading
import time


logger = logging.getLogger("auditoria_excel.timing")


class TimedStage(StrEnum):
    DOWNLOAD_FETCH = "DOWNLOAD_FETCH"
    DOWNLOAD_TRANSFER = "DOWNLOAD_TRANSFER"
    SHA = "SHA"
    READ_XLSX = "READ_XLSX"
    COMPARE = "COMPARE"
    STAGING = "STAGING"
    WAIT_PROMOTION = "WAIT_PROMOTION"
    COMMIT = "COMMIT"
    DOWNLOAD_WAIT = "DOWNLOAD_WAIT"
    WORKER_TASK_DURATION = "WORKER_TASK_DURATION"
    PIPELINE_LATENCY = "PIPELINE_LATENCY"
    TOTAL_TASK = "TOTAL_TASK"


OPERATIONAL_STAGES = (
    # DOWNLOAD_TRANSFER is the slot-visible end-to-end acquisition.  The
    # lower-level browser fetch remains separately observable telemetry, but
    # must not create a second visual phase that the scheduler cannot close.
    TimedStage.DOWNLOAD_TRANSFER, TimedStage.SHA, TimedStage.READ_XLSX,
    TimedStage.COMPARE, TimedStage.STAGING,
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
    TimedStage.DOWNLOAD_WAIT: 0.0,
    TimedStage.WORKER_TASK_DURATION: 10.0,
    TimedStage.PIPELINE_LATENCY: 18.45,
    TimedStage.TOTAL_TASK: 18.45,
}


@dataclass(frozen=True, slots=True)
class TimingEstimate:
    progress: float
    stage_average: float
    task_average: float
    remaining: float
    learned: bool


class GlobalTimingStats:
    """Presentation-only statistics based on ordered, official commits.

    The last 20 commits form the moving throughput window.  Rate and ETA are
    displayed through EWMAs, while their raw values remain available for
    diagnostics.  All timestamps use the pause-aware clock supplied by the
    execution model.
    """

    def __init__(self, window_size: int = 20, rate_alpha: float = .25,
                 eta_alpha: float = .20, minimum_commits: int = 5) -> None:
        if window_size < 2 or not 0 < rate_alpha <= 1 or not 0 < eta_alpha <= 1:
            raise ValueError("configuração temporal inválida")
        self.recent_commit_timestamps: deque[float] = deque(maxlen=window_size)
        self.rate_alpha = rate_alpha
        self.eta_alpha = eta_alpha
        self.minimum_commits = minimum_commits
        self.recent_throughput: float | None = None
        self.recent_average: float | None = None
        self.last_valid_eta: float | None = None
        self.smoothed_eta: float | None = None

    def record_commit(self, timestamp: float) -> None:
        self.recent_commit_timestamps.append(timestamp)
        values = self.recent_commit_timestamps
        if len(values) < 2 or values[-1] <= values[0]:
            return
        raw_rate = (len(values) - 1) * 60.0 / (values[-1] - values[0])
        self.recent_throughput = (raw_rate if self.recent_throughput is None else
                                  self.rate_alpha * raw_rate +
                                  (1 - self.rate_alpha) * self.recent_throughput)
        self.recent_average = 60.0 / self.recent_throughput

    def update_eta(self, remaining_versions: int) -> float | None:
        if remaining_versions <= 0:
            self.last_valid_eta = self.smoothed_eta = 0.0
        elif (len(self.recent_commit_timestamps) >= self.minimum_commits and
              self.recent_throughput is not None and self.recent_throughput > 0):
            raw_eta = remaining_versions * 60.0 / self.recent_throughput
            self.smoothed_eta = (raw_eta if self.smoothed_eta is None else
                                 self.eta_alpha * raw_eta +
                                 (1 - self.eta_alpha) * self.smoothed_eta)
            self.last_valid_eta = self.smoothed_eta
        return self.last_valid_eta


class SlotTimingStats:
    """Robust moving average of the latest completed tasks for one slot."""

    def __init__(self, window_size: int = 15) -> None:
        self.completed_tasks = 0
        self.recent_durations: deque[float] = deque(maxlen=window_size)
        self.average_duration: float | None = None
        self.last_duration: float | None = None

    def complete(self, duration: float) -> None:
        if not math.isfinite(duration) or duration < 0:
            return
        self.completed_tasks += 1
        self.last_duration = duration
        self.recent_durations.append(duration)
        values = tuple(self.recent_durations)
        if len(values) < 3:
            self.average_duration = statistics.fmean(values)
            return
        median = statistics.median(values)
        deviations = tuple(abs(value - median) for value in values)
        mad = statistics.median(deviations)
        spread = max(1.0, 3.0 * 1.4826 * mad)
        self.average_duration = statistics.fmean(
            min(max(value, median - spread), median + spread) for value in values
        )


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
        self.global_stats = GlobalTimingStats(window_size=commit_window)
        self.slot_stats: defaultdict[int, SlotTimingStats] = defaultdict(SlotTimingStats)
        self._lock = threading.RLock()
        self._paused_at: float | None = None
        self._paused_total = 0.0
        self._stopped_at: float | None = None
        self._last_telemetry: dict[TimedStage, float] = {}

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
            if self._stopped_at is None:
                self._stopped_at = self._paused_at if self._paused_at is not None else instant

    def continue_execution(self, now: float | None = None) -> None:
        """Restart the pause-aware clock while retaining all learned samples."""
        instant = time.monotonic() if now is None else now
        with self._lock:
            if self._stopped_at is not None:
                self._paused_total += max(0.0, instant - self._stopped_at)
                self._stopped_at = None
            if self._paused_at is not None:
                self._paused_total += max(0.0, instant - self._paused_at)
                self._paused_at = None

    def observe(self, stage: TimedStage | str, seconds: float, slot_id: int | None = None) -> None:
        stage = TimedStage(stage)
        value = float(seconds)
        if not math.isfinite(value) or value < 0:
            return
        with self._lock:
            self.raw_observations.append((stage, value, slot_id))
            self._samples[stage].append(value)
            if stage is TimedStage.WORKER_TASK_DURATION and slot_id is not None:
                self.slot_stats[slot_id].complete(value)
            values = tuple(self._samples[stage])
            now = time.monotonic()
            should_log = stage not in self._last_telemetry or now - self._last_telemetry[stage] >= 2.0
            if should_log:
                self._last_telemetry[stage] = now
        if should_log:
            logger.info(
                "TIMING_MODEL stage=%s samples=%d mean=%.3f median=%.3f estimate=%.3f",
                stage.value, len(values), statistics.fmean(values), statistics.median(values),
                self.average(stage),
            )

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
        progress = 100.0 if finished else min(99.0, 100 * (completed_weight + expected * fraction) / total)
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
            # Use the model's pause-aware clock.  Commit intervals therefore do
            # not include time spent at a safe paused boundary.
            timestamp = self.active_now(now)
            self._commit_times.append(timestamp)
            self.global_stats.record_commit(timestamp)

    def sample_count(self, stage: TimedStage | str = TimedStage.TOTAL_TASK) -> int:
        """Return the rolling sample count without exposing mutable storage."""
        with self._lock:
            return len(self._samples[TimedStage(stage)])

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
