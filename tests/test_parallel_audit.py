from __future__ import annotations

from pathlib import Path
import sqlite3
import threading
import time

from openpyxl import Workbook
import pytest

from app.database import Database
from app.models import AuditExecutionStatus
from app.parallel_audit import (PREFETCH_TARGET_BY_SLOTS, ParallelAuditService,
                                prefetch_base_target,
                                StagingStore, TaskState, scheduler_window_for_slots)
from app.sources.base import SpreadsheetInfo, VersionInfo
from app.sources.local import LocalSource


SHEET = SpreadsheetInfo("site", "drive", "CQLPA123", "CQLPA123.xlsx")


def history(root: Path, count: int = 6):
    result = []
    for index in range(count):
        path = root / f"v{index}.xlsx"
        book = Workbook()
        sheet = book.active
        sheet["A1"] = index
        sheet["B1"] = f"=A1+{index}"
        book.save(path)
        book.close()
        result.append((VersionInfo(f"technical-{index}", f"5.{119 + index}"), path))
    return result


def source(items):
    return LocalSource([SHEET], {(SHEET.site_id, SHEET.drive_id, SHEET.drive_item_id): items})


def official_rows(connection: sqlite3.Connection):
    versions = [tuple(row) for row in connection.execute(
        "SELECT versao_anterior_id,versao_anterior_numero,versao_atual_id,"
        "versao_atual_numero,quantidade_alteracoes,status FROM versao_processada ORDER BY id"
    )]
    changes = [tuple(row) for row in connection.execute(
        "SELECT v.versao_atual_id,a.tipo,a.aba,a.endereco,a.valor_anterior,a.valor_novo "
        "FROM alteracao a JOIN versao_processada v ON v.id=a.versao_processada_id "
        "ORDER BY v.id,a.aba,a.endereco"
    )]
    checkpoint = tuple(connection.execute(
        "SELECT versao_id,versao_numero FROM checkpoint"
    ).fetchone())
    return versions, changes, checkpoint


def test_one_to_eight_slots_produce_identical_official_database(tmp_path: Path):
    items = history(tmp_path)
    outputs = []
    for slots in range(1, 9):
        with Database(tmp_path / f"audit-{slots}.db") as database:
            database.initialize()
            result = ParallelAuditService(
                database, source(items), slots=slots, backend="thread",
                staging_directory=tmp_path / f"stage-{slots}",
            ).audit(SHEET)
            assert result.status is AuditExecutionStatus.COMPLETED
            outputs.append(official_rows(database.connection))
    assert all(output == outputs[0] for output in outputs[1:])
    assert len(outputs[-1][0]) == len(items) - 1
    assert len({row[2] for row in outputs[-1][0]}) == len(items) - 1


@pytest.mark.parametrize("slots", [0, 9])
def test_worker_count_is_limited_to_one_through_eight(tmp_path: Path, slots: int):
    with Database(tmp_path / "invalid.db") as database:
        database.initialize()
        with pytest.raises(ValueError, match="entre 1 e 8"):
            ParallelAuditService(database, source([]), slots=slots)


def test_prefetch_target_is_bounded_benchmark_configuration():
    assert PREFETCH_TARGET_BY_SLOTS == {
        1: 2, 2: 4, 3: 6, 4: 8, 5: 10, 6: 12, 7: 14, 8: 16,
    }
    assert [scheduler_window_for_slots(slots) for slots in range(1, 9)] == [
        3, 6, 9, 12, 15, 18, 21, 24,
    ]


@pytest.mark.parametrize(
    ("slots", "average_bytes", "expected"),
    [(8, 512 * 1024, 16), (8, 2 * 1024**2, 8), (8, 8 * 1024**2, 4),
     (1, 8 * 1024**2, 2)],
)
def test_prefetch_base_target_adapts_to_file_size(slots, average_bytes, expected):
    assert prefetch_base_target(slots, average_bytes) == expected


def test_large_files_keep_operational_prefetch_floor_for_worker_count():
    assert prefetch_base_target(4, 64 * 1024**2) == 3
    assert prefetch_base_target(8, 64 * 1024**2) == 4


def test_parallel_scheduler_prefetches_once_by_technical_id(tmp_path: Path):
    items = history(tmp_path, 5)

    class PrefetchLocal(LocalSource):
        def __init__(self):
            super().__init__([SHEET], {(SHEET.site_id, SHEET.drive_id, SHEET.drive_item_id): items})
            self.planned = []
            self.limit = None

        def configure_prefetch_buffer(self, size):
            self.limit = size

        def prefetch_version(self, spreadsheet, version):
            self.planned.append(version.id)
            return True

    fake = PrefetchLocal()
    with Database(tmp_path / "prefetch-parallel.db") as database:
        database.initialize()
        service = ParallelAuditService(
            database, fake, slots=3, backend="thread",
            staging_directory=tmp_path / "stage-prefetch",
        )
        result = service.audit(SHEET)
    assert result.status is AuditExecutionStatus.COMPLETED
    assert fake.limit == 6
    assert len(fake.planned) == len(set(fake.planned))
    assert service._prefetch_hits > 0


def test_slot_phase_events_use_timing_model_instead_of_fixed_percentages(tmp_path: Path):
    events = []
    with Database(tmp_path / "phase-progress.db") as database:
        database.initialize()
        result = ParallelAuditService(
            database, source(history(tmp_path, 3)), slots=1, backend="thread",
            slot_callback=events.append, staging_directory=tmp_path / "stage-progress",
        ).audit(SHEET)

    assert result.status is AuditExecutionStatus.COMPLETED
    compare = [event for event in events if event.state is TaskState.COMPARE]
    staging = [event for event in events if event.state is TaskState.STAGED]
    assert compare and all(event.timed_stage is not None for event in compare)
    assert staging and all(event.percent <= 99 for event in staging)
    assert all(event.percent < 100 for event in events if event.state is not TaskState.COMPLETED)


def test_out_of_order_staging_never_advances_checkpoint(monkeypatch, tmp_path: Path):
    items = history(tmp_path, 5)
    staged = []
    checkpoints_during_staging = []
    import app.parallel_audit as parallel
    real_compare = parallel._compare_pair

    def delayed(task):
        # Primeira lacuna deliberadamente lenta; as três seguintes terminam antes.
        time.sleep({1: .30, 2: .02, 3: .03, 4: .01}[task.sequence])
        return real_compare(task)

    monkeypatch.setattr(parallel, "_compare_pair", delayed)
    with Database(tmp_path / "order.db") as database:
        database.initialize()

        def report(event):
            if event.state is TaskState.STAGED:
                staged.append(event.sequence)
                row = database.connection.execute("SELECT versao_numero FROM checkpoint").fetchone()
                checkpoints_during_staging.append(None if row is None else row[0])

        result = ParallelAuditService(
            database, source(items), slots=2, backend="thread", slot_callback=report,
            staging_directory=tmp_path / "stage-order",
        ).audit(SHEET)
        assert result.final_version == "5.123"
        assert staged[0] == 2
        assert checkpoints_during_staging[0] is None
        assert [row[0] for row in database.connection.execute(
            "SELECT versao_atual_numero FROM versao_processada ORDER BY id"
        )] == ["5.120", "5.121", "5.122", "5.123"]


def test_restart_discards_staging_and_resumes_from_official_checkpoint(tmp_path: Path):
    items = history(tmp_path, 4)
    with Database(tmp_path / "crash.db") as database:
        database.initialize()
        first = ParallelAuditService(
            database, source(items[:2]), slots=1, backend="thread",
            staging_directory=tmp_path / "staging",
        ).audit(SHEET)
        assert first.final_version == "5.120"

        orphan = StagingStore(tmp_path / "staging", "orphan")
        orphan.connection.close()  # simula encerramento abrupto, sem close/remoção
        assert orphan.path.exists()

        resumed = ParallelAuditService(
            database, source(items[1:]), slots=2, backend="thread",
            staging_directory=tmp_path / "staging",
        ).audit(SHEET)
        assert resumed.initial_checkpoint == "5.120"
        assert resumed.final_version == "5.122"
        assert not orphan.path.exists()
        rows = list(database.connection.execute("SELECT versao_atual_id FROM versao_processada"))
        assert len(rows) == 3
        assert len({row[0] for row in rows}) == 3


def test_pause_and_resume_at_safe_boundary_with_eight_slots(tmp_path: Path):
    items = history(tmp_path, 3)
    pause = threading.Event()
    pause.set()
    controls = []
    timer = threading.Timer(.1, pause.clear)
    timer.start()
    try:
        with Database(tmp_path / "pause.db") as database:
            database.initialize()
            result = ParallelAuditService(
                database, source(items), slots=8, backend="thread",
                pause_event=pause, control_callback=lambda state, checkpoint: controls.append(state),
                staging_directory=tmp_path / "stage-pause",
            ).audit(SHEET)
        assert result.status is AuditExecutionStatus.COMPLETED
        assert controls[:2] == ["paused", "resumed"]
    finally:
        timer.cancel()


def test_stop_with_eight_slots_preserves_empty_checkpoint(tmp_path: Path):
    stop = threading.Event()
    stop.set()
    with Database(tmp_path / "stop.db") as database:
        database.initialize()
        result = ParallelAuditService(
            database, source(history(tmp_path, 3)), slots=8, backend="thread",
            stop_event=stop, staging_directory=tmp_path / "stage-stop",
        ).audit(SHEET)
        assert result.status is AuditExecutionStatus.STOPPED
        assert database.connection.execute("SELECT count(*) FROM checkpoint").fetchone()[0] == 0


def test_completed_worker_is_promoted_while_webdriver_owner_is_blocked(
    monkeypatch, tmp_path: Path,
):
    """A Selenium wait must not hold staging or the ordered checkpoint."""
    items = history(tmp_path, 4)
    fetch_active = threading.Event()
    promoted_during_fetch = []

    class SlowDownloadSource(LocalSource):
        def get_version(self, spreadsheet, version):
            if version.id == "technical-2":
                fetch_active.set()
                time.sleep(.30)
                fetch_active.clear()
            return super().get_version(spreadsheet, version)

    import app.parallel_audit as parallel
    real_compare = parallel._compare_pair

    def short_compare(task):
        time.sleep(.03)
        return real_compare(task)

    monkeypatch.setattr(parallel, "_compare_pair", short_compare)
    slow_source = SlowDownloadSource(
        [SHEET], {(SHEET.site_id, SHEET.drive_id, SHEET.drive_item_id): items}
    )
    with Database(tmp_path / "promotion-during-fetch.db") as database:
        database.initialize()
        result = ParallelAuditService(
            database, slow_source, slots=4, backend="thread",
            checkpoint_callback=lambda *_args: promoted_during_fetch.append(fetch_active.is_set()),
            staging_directory=tmp_path / "stage-promotion-during-fetch",
        ).audit(SHEET)

    assert result.status is AuditExecutionStatus.COMPLETED
    assert promoted_during_fetch[0] is True


def test_four_workers_receive_ready_files_concurrently(monkeypatch, tmp_path: Path):
    items = history(tmp_path, 7)
    lock = threading.Lock()
    active = 0
    peak = 0
    import app.parallel_audit as parallel
    real_compare = parallel._compare_pair

    def measured_compare(task):
        nonlocal active, peak
        with lock:
            active += 1
            peak = max(peak, active)
        try:
            time.sleep(.12)
            return real_compare(task)
        finally:
            with lock:
                active -= 1

    monkeypatch.setattr(parallel, "_compare_pair", measured_compare)
    with Database(tmp_path / "four-workers.db") as database:
        database.initialize()
        service = ParallelAuditService(
            database, source(items), slots=4, backend="thread",
            staging_directory=tmp_path / "stage-four-workers",
        )
        result = service.audit(SHEET)

    assert result.status is AuditExecutionStatus.COMPLETED
    assert peak == 4
    staged = [item for item in service.telemetry if item.get("stage") == "STAGED"]
    assert all(float(item["worker_to_staging"]) < .5 for item in staged)
    assert all("task_reserved_at" in item and "promoted_at" in item for item in staged)
