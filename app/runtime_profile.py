"""Aprendizado conservador e persistente de desempenho por workbook."""

from __future__ import annotations

from dataclasses import dataclass
import logging
import math
import sqlite3
from typing import Iterable

from app.sources.base import SpreadsheetInfo


logger = logging.getLogger("auditoria_excel.runtime_profile")
PROFILE_VERSION = 1
MINIMUM_TASKS = 20
TECHNICAL_TIE = .05
DRIFT_RATIO = 1.75


def workbook_identity(sheet: SpreadsheetInfo) -> str:
    """Forma estável baseada exclusivamente na identidade técnica SharePoint."""
    return "|".join((sheet.site_id, sheet.drive_id, sheet.drive_item_id))


@dataclass(frozen=True, slots=True)
class RuntimeProfile:
    workbook_identity: str
    avg_file_bytes: float
    recommended_slots: int | None
    recommended_prefetch_target: int | None
    best_measured_throughput: float
    best_measured_slots: int | None
    recommendation_confidence: str
    profile_needs_revalidation: bool
    sample_count: int
    benchmark_count: int
    last_driver_count: int = 1


@dataclass(frozen=True, slots=True)
class RuntimeSample:
    slots: int
    prefetch_target: int
    versions_processed: int
    elapsed_seconds: float
    avg_download: float = 0.0
    avg_read_xlsx: float = 0.0
    avg_compare: float = 0.0
    avg_worker_task: float = 0.0
    worker_utilization: float = 0.0
    avg_file_bytes: float = 0.0
    completed_normally: bool = True
    run_kind: str = "normal_run"

    @property
    def throughput(self) -> float:
        return self.versions_processed * 60.0 / self.elapsed_seconds

    def valid(self) -> bool:
        values = (self.elapsed_seconds, self.avg_download, self.avg_read_xlsx,
                  self.avg_compare, self.avg_worker_task, self.avg_file_bytes)
        return (MINIMUM_TASKS <= self.versions_processed and 1 <= self.slots <= 8
                and self.prefetch_target >= 1 and self.run_kind in {"normal_run", "benchmark_run"}
                and all(math.isfinite(value) and value >= 0 for value in values))


class RuntimeProfileStore:
    """Repository SQLite. Chamadores executam seus métodos na worker thread."""

    def __init__(self, connection: sqlite3.Connection) -> None:
        self.connection = connection

    def load(self, identity: str) -> RuntimeProfile | None:
        try:
            row = self.connection.execute(
                "SELECT * FROM workbook_runtime_profile WHERE workbook_identity=?",
                (identity,),
            ).fetchone()
            if row is None or row["profile_version"] != PROFILE_VERSION:
                return None
            slots = row["recommended_slots"]
            prefetch = row["recommended_prefetch_target"]
            if slots is not None and not 1 <= slots <= 8:
                return None
            if prefetch is not None and prefetch < 1:
                return None
            result = RuntimeProfile(
                identity, float(row["avg_file_bytes"]), slots, prefetch,
                float(row["best_measured_throughput"]), row["best_measured_slots"],
                row["recommendation_confidence"], bool(row["profile_needs_revalidation"]),
                int(row["sample_count"]), int(row["benchmark_count"]),
                int(row["last_driver_count"]),
            )
            if (not math.isfinite(result.avg_file_bytes) or result.avg_file_bytes < 0
                    or result.recommendation_confidence not in {"LOW", "MEDIUM", "HIGH"}):
                return None
        except (sqlite3.Error, KeyError, TypeError, ValueError, OverflowError):
            logger.warning("Perfil inválido ignorado workbook=%s", identity, exc_info=True)
            return None
        logger.info(
            "RUNTIME_PROFILE_LOAD workbook=%s recommended_slots=%s recommended_prefetch=%s "
            "avg_file_bytes=%.0f confidence=%s", identity, result.recommended_slots,
            result.recommended_prefetch_target, result.avg_file_bytes,
            result.recommendation_confidence,
        )
        return result

    def save_driver_count(self, identity: str, driver_count: int) -> None:
        """Persiste somente a escolha explícita, sem criar recomendação automática."""
        if not 1 <= driver_count <= 4:
            raise ValueError("driver_count deve estar entre 1 e 4")
        with self.connection:
            self.connection.execute(
                """INSERT INTO workbook_runtime_profile (workbook_identity,last_driver_count)
                   VALUES (?,?) ON CONFLICT(workbook_identity) DO UPDATE SET
                   last_driver_count=excluded.last_driver_count,
                   last_updated_at=CURRENT_TIMESTAMP""",
                (identity, driver_count),
            )

    def record(self, identity: str, sample: RuntimeSample) -> RuntimeProfile | None:
        """Guarda amostra válida e recalcula recomendação por throughput comparável."""
        if not sample.valid():
            return self.load(identity)
        throughput = sample.throughput
        logger.info(
            "RUNTIME_PROFILE_SAMPLE slots=%d throughput=%.3f read_xlsx=%.3f "
            "compare=%.3f worker_utilization=%.4f", sample.slots, throughput,
            sample.avg_read_xlsx, sample.avg_compare, sample.worker_utilization,
        )
        old = self.load(identity)
        drift = bool(old and (
            self._drifted(old.avg_file_bytes, sample.avg_file_bytes)
            or self._read_drift(identity, sample.avg_read_xlsx)
        ))
        comparable = not drift
        with self.connection:
            self.connection.execute(
                """INSERT INTO workbook_runtime_benchmark
                   (workbook_identity,run_kind,slots,prefetch_target,versions_processed,
                    elapsed_seconds,throughput_per_minute,avg_download,avg_read_xlsx,
                    avg_compare,avg_worker_task,worker_utilization,avg_file_bytes,
                    completed_normally,comparable)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (identity, sample.run_kind, sample.slots, sample.prefetch_target,
                 sample.versions_processed, sample.elapsed_seconds, throughput,
                 sample.avg_download, sample.avg_read_xlsx, sample.avg_compare,
                 sample.avg_worker_task, sample.worker_utilization, sample.avg_file_bytes,
                 int(sample.completed_normally), int(comparable)),
            )
            self._upsert_averages(identity, sample, drift)
            self._recompute(identity)
        updated = self.load(identity)
        logger.info(
            "RUNTIME_PROFILE_UPDATED old_slots=%s new_slots=%s reason=%s",
            old.recommended_slots if old else None,
            updated.recommended_slots if updated else None,
            "drift_revalidation" if drift else "throughput_with_tie_tolerance",
        )
        return updated

    @staticmethod
    def _drifted(old: float, new: float) -> bool:
        return old > 0 and new > 0 and max(old, new) / min(old, new) >= DRIFT_RATIO

    def _read_drift(self, identity: str, new: float) -> bool:
        row = self.connection.execute(
            "SELECT avg_read_xlsx_seconds FROM workbook_runtime_profile WHERE workbook_identity=?",
            (identity,),
        ).fetchone()
        return bool(row and self._drifted(float(row[0]), new))

    def _upsert_averages(self, identity: str, sample: RuntimeSample, drift: bool) -> None:
        weight = 1.0 if sample.completed_normally else .5
        benchmark = int(sample.run_kind == "benchmark_run")
        # A média ponderada conserva o histórico; stop voluntário tem metade do peso.
        row = self.connection.execute(
            "SELECT * FROM workbook_runtime_profile WHERE workbook_identity=?", (identity,)
        ).fetchone()
        count = int(row["sample_count"]) if row else 0
        denominator = count + weight
        def mean(column: str, value: float) -> float:
            prior = float(row[column]) if row else 0.0
            return (prior * count + value * weight) / denominator
        values = (
            mean("avg_file_bytes", sample.avg_file_bytes),
            mean("avg_download_seconds", sample.avg_download),
            mean("avg_read_xlsx_seconds", sample.avg_read_xlsx),
            mean("avg_compare_seconds", sample.avg_compare),
            mean("avg_worker_task_seconds", sample.avg_worker_task),
            mean("avg_worker_utilization", sample.worker_utilization),
        )
        self.connection.execute(
            """INSERT INTO workbook_runtime_profile
               (workbook_identity,avg_file_bytes,last_slots_used,last_prefetch_target,
                avg_download_seconds,avg_read_xlsx_seconds,avg_compare_seconds,
                avg_worker_task_seconds,avg_worker_utilization,sample_count,benchmark_count,
                last_benchmark_at,profile_needs_revalidation,profile_version)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,CASE WHEN ? THEN CURRENT_TIMESTAMP END,?,?)
               ON CONFLICT(workbook_identity) DO UPDATE SET
                avg_file_bytes=excluded.avg_file_bytes,last_slots_used=excluded.last_slots_used,
                last_prefetch_target=excluded.last_prefetch_target,
                avg_download_seconds=excluded.avg_download_seconds,
                avg_read_xlsx_seconds=excluded.avg_read_xlsx_seconds,
                avg_compare_seconds=excluded.avg_compare_seconds,
                avg_worker_task_seconds=excluded.avg_worker_task_seconds,
                avg_worker_utilization=excluded.avg_worker_utilization,
                sample_count=workbook_runtime_profile.sample_count+1,
                benchmark_count=workbook_runtime_profile.benchmark_count+?,
                last_benchmark_at=CASE WHEN ? THEN CURRENT_TIMESTAMP ELSE last_benchmark_at END,
                profile_needs_revalidation=?,last_updated_at=CURRENT_TIMESTAMP""",
            (identity, *values[:1], sample.slots, sample.prefetch_target, *values[1:],
             1, benchmark, benchmark, int(drift), PROFILE_VERSION,
             benchmark, benchmark, int(drift)),
        )

    def _recompute(self, identity: str) -> None:
        rows = self.connection.execute(
            """SELECT slots,
                      SUM(throughput_per_minute *
                          CASE WHEN run_kind='benchmark_run' THEN 2.0 ELSE 1.0 END *
                          CASE WHEN completed_normally=1 THEN 1.0 ELSE 0.5 END) /
                      SUM(CASE WHEN run_kind='benchmark_run' THEN 2.0 ELSE 1.0 END *
                          CASE WHEN completed_normally=1 THEN 1.0 ELSE 0.5 END) throughput,
                      COUNT(*) samples,
                      CAST(ROUND(SUM(prefetch_target *
                          CASE WHEN run_kind='benchmark_run' THEN 2.0 ELSE 1.0 END) /
                          SUM(CASE WHEN run_kind='benchmark_run' THEN 2.0 ELSE 1.0 END)) AS INTEGER) prefetch
                 FROM workbook_runtime_benchmark
                WHERE workbook_identity=? AND comparable=1
                GROUP BY slots ORDER BY slots""", (identity,)
        ).fetchall()
        if not rows:
            return
        peak = max(float(row["throughput"]) for row in rows)
        # Configurações até 5% do pico são empate técnico: vence a de menos slots.
        candidates = [row for row in rows if float(row["throughput"]) >= peak * (1 - TECHNICAL_TIE)]
        best = min(candidates, key=lambda row: int(row["slots"]))
        consistent = int(best["samples"])
        confidence = "HIGH" if consistent >= 4 else "MEDIUM" if consistent >= 2 else "LOW"
        self.connection.execute(
            """UPDATE workbook_runtime_profile SET recommended_slots=?,
                      recommended_prefetch_target=?,best_measured_throughput=?,
                      best_measured_slots=?,recommendation_confidence=?,
                      last_updated_at=CURRENT_TIMESTAMP WHERE workbook_identity=?""",
            (best["slots"], best["prefetch"], best["throughput"], best["slots"],
             confidence, identity),
        )

    def import_benchmarks(self, identity: str, samples: Iterable[RuntimeSample]) -> RuntimeProfile | None:
        result = None
        for sample in samples:
            result = self.record(identity, sample)
        return result
