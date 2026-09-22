"""Estimativa exclusivamente visual do progresso de uma versão."""

from __future__ import annotations

from collections import deque
from dataclasses import dataclass, field
import math
import statistics


MILESTONES = (0.0, 5.0, 55.0, 70.0, 90.0, 97.0, 100.0)
DEFAULT_SECONDS = {
    0.0: 0.5,
    5.0: 10.0,
    55.0: 3.0,
    70.0: 4.0,
    90.0: 2.0,
    97.0: 2.0,
}


@dataclass(slots=True)
class SmoothVersionProgress:
    """Interpola entre marcos reais sem jamais antecipar o próximo marco.

    A mediana das vinte observações mais recentes limita naturalmente a
    influência de timeouts e recuperações excepcionais.
    """

    histories: dict[float, deque[float]] = field(
        default_factory=lambda: {
            milestone: deque(maxlen=20) for milestone in MILESTONES[:-1]
        }
    )
    total_history: deque[float] = field(default_factory=lambda: deque(maxlen=20))
    version: str | None = None
    value: float = 0.0
    milestone: float = 0.0
    phase_started_at: float | None = None
    version_started_at: float | None = None

    def observe(self, version: str, percent: float, now: float) -> float:
        """Registra um marco autoritativo e devolve progresso monotônico."""
        if percent == 0 and version != self.version:
            self.version = version
            self.value = self.milestone = 0.0
            self.phase_started_at = self.version_started_at = now
            return self.value
        if version != self.version:
            return self.value
        percent = float(percent)
        if percent < self.milestone:
            return self.value
        if percent > self.milestone and self.phase_started_at is not None:
            duration = max(0.0, now - self.phase_started_at)
            if duration > 0:
                self.histories[self.milestone].append(duration)
            self.milestone = percent
            self.phase_started_at = now
        self.value = max(self.value, percent)
        if percent == 100 and self.version_started_at is not None:
            self.total_history.append(max(0.0, now - self.version_started_at))
            self.version_started_at = None
            self.phase_started_at = None
        return self.value

    def estimate(self, now: float) -> float:
        """Avança suavemente, aproximando-se do marco seguinte sem alcançá-lo."""
        if self.phase_started_at is None or self.milestone in (0, 100):
            return self.value
        index = MILESTONES.index(self.milestone)
        upper = MILESTONES[index + 1]
        expected = self.phase_average(self.milestone)
        elapsed = max(0.0, now - self.phase_started_at)
        ratio = elapsed / max(expected, 0.001)
        # Usa 90% da faixa até o tempo esperado. Depois disso, uma cauda
        # exponencial mantém movimento perceptível sem confirmar o marco real.
        fraction = 0.9 * ratio if ratio <= 1 else 0.9 + 0.1 * (1 - math.exp(-(ratio - 1)))
        guard = 0.5
        calculated = min(self.milestone + (upper - self.milestone) * fraction, upper - guard)
        self.value = max(self.value, calculated)
        return self.value

    def phase_average(self, milestone: float) -> float:
        samples = self.histories[milestone]
        return statistics.median(samples) if samples else DEFAULT_SECONDS[milestone]

    def total_average(self) -> float | None:
        return statistics.median(self.total_history) if self.total_history else None
