from pathlib import Path

import pytest

from app.database import Database
from app.interface import AuditApplication
from app.runtime_profile import (RuntimeProfileStore, RuntimeSample,
                                 RuntimeProfile, workbook_identity)
from app.sources.base import SpreadsheetInfo


def sample(slots: int, throughput: float, *, size: float = 8 * 1024**2,
           prefetch: int = 4, completed: bool = True) -> RuntimeSample:
    count = 20
    return RuntimeSample(
        slots=slots, prefetch_target=prefetch, versions_processed=count,
        elapsed_seconds=count * 60 / throughput, avg_download=2.0,
        avg_read_xlsx=10.0 + slots, avg_compare=1.0,
        avg_worker_task=11.0 + slots, worker_utilization=.8,
        avg_file_bytes=size, completed_normally=completed,
        run_kind="benchmark_run",
    )


@pytest.fixture
def store(tmp_path: Path):
    with Database(tmp_path / "profiles.db") as database:
        database.initialize()
        yield RuntimeProfileStore(database.connection)


def test_workbook_without_profile_uses_caller_default(store) -> None:
    assert store.load("unknown") is None
    default_slots = 3
    assert (store.load("unknown") or default_slots) == default_slots


def test_real_identity_is_stable_and_workbooks_are_independent(store) -> None:
    first = workbook_identity(SpreadsheetInfo("site", "drive", "one", "same.xlsx"))
    second = workbook_identity(SpreadsheetInfo("site", "drive", "two", "same.xlsx"))
    store.record(first, sample(7, 16.55))
    assert store.load(first).recommended_slots == 7
    assert store.load(second) is None


def test_cqlpa123_measured_results_choose_seven_and_survive_reopen(tmp_path: Path) -> None:
    path = tmp_path / "second-open.db"
    identity = "tenant-site|quality-drive|real-item-id"
    observed = [(2, 9.04, 10.1), (3, 11.19, 12.7), (4, 12.71, 14.3),
                (5, 14.04, 16.8), (6, 13.85, 19.8), (7, 16.55, 19.7),
                (8, 14.12, 25.0)]
    with Database(path) as database:
        database.initialize()
        store = RuntimeProfileStore(database.connection)
        for slots, throughput, read in observed:
            measured = sample(slots, throughput)
            measured = RuntimeSample(**{
                name: getattr(measured, name) for name in measured.__dataclass_fields__
                if name != "avg_read_xlsx"
            }, avg_read_xlsx=read)
            store.record(identity, measured)
        profile = store.load(identity)
        assert profile.recommended_slots == 7
        assert profile.recommended_prefetch_target == 4
        assert profile.avg_file_bytes == 8 * 1024**2
    with Database(path) as reopened:
        reopened.initialize()
        profile = RuntimeProfileStore(reopened.connection).load(identity)
        assert profile.recommended_slots == 7
        assert profile.avg_file_bytes > 0
        count = reopened.connection.execute(
            "SELECT COUNT(*) FROM workbook_runtime_benchmark WHERE workbook_identity=?",
            (identity,),
        ).fetchone()[0]
        assert count == 7


def test_significant_improvement_changes_recommendation(store) -> None:
    store.record("book", sample(7, 16.5))
    store.record("book", sample(6, 18.0))
    assert store.load("book").recommended_slots == 6


def test_technical_tie_prefers_fewer_slots(store) -> None:
    store.record("book", sample(6, 14.1))
    store.record("book", sample(5, 14.0))
    assert store.load("book").recommended_slots == 5


def test_short_or_invalid_run_does_not_change_profile(store) -> None:
    invalid = RuntimeSample(8, 16, 19, 1, avg_file_bytes=1)
    assert store.record("book", invalid) is None
    assert store.load("book") is None


def test_drift_preserves_history_and_marks_revalidation(store) -> None:
    store.record("book", sample(7, 16.5))
    store.record("book", sample(6, 20.0, size=18 * 1024**2))
    profile = store.load("book")
    assert profile.profile_needs_revalidation is True
    assert profile.recommended_slots == 7


def test_manual_stop_has_lower_average_weight_but_is_persisted(store) -> None:
    store.record("book", sample(7, 16.5, size=8_000_000))
    store.record("book", sample(7, 16.5, size=11_000_000, completed=False))
    assert store.load("book").avg_file_bytes == pytest.approx(9_000_000)


class _Value:
    def __init__(self, value=None):
        self.value = value

    def set(self, value):
        self.value = value

    def get(self):
        return self.value


def test_interface_applies_seven_but_preserves_manual_five() -> None:
    application = AuditApplication.__new__(AuditApplication)
    application.spreadsheets = [SpreadsheetInfo("s", "d", "i", "CQLPA123.xlsx")]
    application.selector = type("Selector", (), {"current": lambda self: 0})()
    application.worker_count = _Value(3)
    application.runtime_recommendation = _Value()
    application._audit_active = False
    application._manual_slot_identity = None
    application._rebuild_slot_frames = lambda: None
    profile = RuntimeProfile("s|d|i", 8 * 1024**2, 7, 4, 16.55, 7,
                             "HIGH", False, 7, 7)
    application._apply_runtime_profile(profile)
    assert application.worker_count.get() == 7
    assert "7 slots" in application.runtime_recommendation.value

    application.worker_count.set(5)
    application._manual_slot_identity = "s|d|i"
    application._apply_runtime_profile(profile)
    assert application.worker_count.get() == 5
