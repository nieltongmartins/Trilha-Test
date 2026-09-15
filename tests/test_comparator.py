from pathlib import Path

from app.excel.comparator import CellChange, compare_snapshots
from app.excel.reader import read_workbook
from app.models import ChangeType


def test_comparator_detects_required_changes_from_controlled_files(
    cql028_versions: Path,
) -> None:
    previous = read_workbook(cql028_versions / "0.84.xlsx")
    current = read_workbook(cql028_versions / "0.85.xlsx")

    changes = compare_snapshots(previous, current)

    assert changes == [
        CellChange("Detalhes", "B2", ChangeType.ADD, None, "Nova informação"),
        CellChange("Resumo", "A1", ChangeType.MOD, 10, 15),
        CellChange("Resumo", "B2", ChangeType.ADD, None, "OK"),
        CellChange("Resumo", "D4", ChangeType.ADD, None, 0),
        CellChange("Resumo", "E5", ChangeType.ADD, None, False),
        CellChange(
            "Resumo", "F6", ChangeType.MOD, "=SUM(A1:A10)", "=SUM(A1:A20)"
        ),
    ]


def test_comparator_detects_deletion_modification_and_added_sheet(
    cql028_versions: Path,
) -> None:
    changes = compare_snapshots(
        read_workbook(cql028_versions / "0.85.xlsx"),
        read_workbook(cql028_versions / "0.86.xlsx"),
    )

    assert changes == [
        CellChange(
            "Detalhes", "B2", ChangeType.MOD, "Nova informação", "Alterada"
        ),
        CellChange("Nova Aba", "A1", ChangeType.ADD, None, "Criada"),
        CellChange("Resumo", "C3", ChangeType.DEL, "Pendente", None),
    ]


def test_comparator_returns_no_changes_for_equal_snapshots(
    cql028_versions: Path,
) -> None:
    version = read_workbook(cql028_versions / "0.86.xlsx")

    assert compare_snapshots(version, version) == []
    assert compare_snapshots(
        version, read_workbook(cql028_versions / "0.87.xlsx")
    ) == []


def test_comparator_is_deterministic_and_uses_row_then_column_order() -> None:
    previous = {"Zeta": {}, "Alfa": {}}
    current = {
        "Zeta": {"A10": 1, "B2": 2, "A2": 3},
        "Alfa": {"AA1": 4, "Z1": 5},
    }

    first = compare_snapshots(previous, current)
    second = compare_snapshots(previous, current)

    assert first == second
    assert [(change.sheet, change.address) for change in first] == [
        ("Alfa", "Z1"),
        ("Alfa", "AA1"),
        ("Zeta", "A2"),
        ("Zeta", "B2"),
        ("Zeta", "A10"),
    ]


def test_comparator_distinguishes_false_zero_and_empty_string() -> None:
    previous = {"Aba": {"A1": False, "A2": 0}}
    current = {"Aba": {"A1": 0, "A2": False, "A3": ""}}

    assert compare_snapshots(previous, current) == [
        CellChange("Aba", "A1", ChangeType.MOD, False, 0),
        CellChange("Aba", "A2", ChangeType.MOD, 0, False),
        CellChange("Aba", "A3", ChangeType.ADD, None, ""),
    ]


def test_removed_sheet_is_represented_as_cell_deletions() -> None:
    changes = compare_snapshots({"Removida": {"B1": "valor"}}, {})

    assert changes == [
        CellChange("Removida", "B1", ChangeType.DEL, "valor", None)
    ]
