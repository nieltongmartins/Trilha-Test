from datetime import date, datetime, time
from pathlib import Path

from openpyxl import Workbook
import pytest

from app.excel.read_ahead import READ_AHEAD_BUFFER_SIZE, ReadAheadExecutor
from app.excel.reader import read_workbook


def make_typed_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Tipos"
    values = {
        "A1": "texto",
        "A2": 42,
        "A3": 1.25,
        "A4": True,
        "A5": "=SUM(A2, 8)",
        "A6": date(2026, 9, 22),
        "A7": datetime(2026, 9, 22, 10, 30),
        "A8": time(10, 30),
        "A9": "#DIV/0!",
    }
    for coordinate, value in values.items():
        sheet[coordinate] = value
    sheet["A9"].data_type = "e"
    sheet["A10"] = None
    workbook.create_sheet("Vazia")
    workbook.save(path)
    workbook.close()


def test_process_snapshot_is_exactly_equivalent_to_official_reader(tmp_path: Path) -> None:
    path = tmp_path / "tipos.xlsx"
    make_typed_workbook(path)
    expected = read_workbook(path)
    executor = ReadAheadExecutor()
    try:
        token = executor.submit("id-12", "12.0", path)
        prepared, _ = executor.result("id-12", "12.0", path, token)
    finally:
        executor.close()

    assert READ_AHEAD_BUFFER_SIZE == 1
    assert prepared.snapshot == expected
    assert prepared.snapshot.keys() == expected.keys()
    for sheet, cells in expected.items():
        assert prepared.snapshot[sheet].keys() == cells.keys()
        for coordinate, value in cells.items():
            actual = prepared.snapshot[sheet][coordinate]
            assert actual == value
            assert type(actual) is type(value)
    assert "A10" not in prepared.snapshot["Tipos"]
    assert prepared.cell_count == sum(len(cells) for cells in expected.values())
    assert prepared.snapshot_bytes > 0
    assert prepared.version_id == "id-12"
    assert prepared.version_label == "12.0"
    assert prepared.origin == str(path.resolve())


def test_identity_mismatch_is_rejected_before_snapshot_use(tmp_path: Path) -> None:
    path = tmp_path / "tipos.xlsx"
    make_typed_workbook(path)
    executor = ReadAheadExecutor()
    token = executor.submit("id-correto", "3.4", path)
    try:
        with pytest.raises(RuntimeError, match="identidade"):
            executor.result("id-errado", "3.4", path, token)
    finally:
        executor.close()
