"""Estimador puramente visual das barras dos slots.

O estado deste módulo é deliberadamente derivado dos eventos funcionais: ele
não produz sinais de volta para o pipeline e não confirma fases ou tarefas.
"""

from __future__ import annotations

from dataclasses import dataclass
import math

from app.execution_timing import OPERATIONAL_STAGES, SharedExecutionTimingModel, TimedStage


OPERATIONAL_LIMIT = 99.0
PROMOTION_LIMIT = 99.8


@dataclass(frozen=True, slots=True)
class PhaseProgressRange:
    start: float
    end: float
    estimate: float


@dataclass(frozen=True, slots=True)
class TaskProgressPlan:
    """Snapshot imutável das médias e faixas de uma única task."""

    ranges: dict[TimedStage, PhaseProgressRange]
    promotion: PhaseProgressRange

    @classmethod
    def snapshot(cls, model: SharedExecutionTimingModel) -> "TaskProgressPlan":
        estimates = {stage: max(.001, model.average(stage)) for stage in OPERATIONAL_STAGES}
        total = sum(estimates.values())
        cursor = 0.0
        ranges: dict[TimedStage, PhaseProgressRange] = {}
        for stage in OPERATIONAL_STAGES:
            end = cursor + OPERATIONAL_LIMIT * estimates[stage] / total
            ranges[stage] = PhaseProgressRange(cursor, end, estimates[stage])
            cursor = end
        return cls(
            ranges,
            PhaseProgressRange(OPERATIONAL_LIMIT, PROMOTION_LIMIT,
                               max(.001, model.average(TimedStage.WAIT_PROMOTION))),
        )


@dataclass(slots=True)
class SlotVisualProgress:
    """Progresso monotônico de uma task, atualizado pelo ticker da UI."""

    task_id: str | None
    plan: TaskProgressPlan
    phase: TimedStage
    phase_started_monotonic: float
    last_displayed_progress: float = 0.0
    complete: bool = False
    waiting: bool = False

    @staticmethod
    def _local_fraction(elapsed: float, estimate: float) -> float:
        ratio = max(0.0, elapsed) / max(.001, estimate)
        if ratio <= 1.0:
            return .85 * ratio
        # Contínua em ratio=1 e assintótica a 99%: o fim da fase depende
        # exclusivamente do próximo evento funcional.
        return .85 + .14 * (1.0 - math.exp(-(ratio - 1.0)))

    def _phase_range(self) -> PhaseProgressRange:
        if self.phase is TimedStage.WAIT_PROMOTION:
            return self.plan.promotion
        return self.plan.ranges[self.phase]

    def enter_phase(self, phase: TimedStage, now: float) -> None:
        old_range = self._phase_range()
        self.last_displayed_progress = max(self.last_displayed_progress, old_range.end)
        self.phase = phase
        self.phase_started_monotonic = now
        self.waiting = False

    def tick(self, now: float) -> float:
        if self.complete:
            self.last_displayed_progress = 100.0
        elif not self.waiting:
            phase_range = self._phase_range()
            fraction = self._local_fraction(
                now - self.phase_started_monotonic, phase_range.estimate
            )
            calculated = phase_range.start + (phase_range.end - phase_range.start) * fraction
            self.last_displayed_progress = max(
                self.last_displayed_progress, min(calculated, phase_range.end)
            )
        return self.last_displayed_progress

