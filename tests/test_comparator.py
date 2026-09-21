from pathlib import Path
import random

from app.excel.comparator import CellChange, compare_snapshots
from app.excel.reader import read_workbook
from app.models import ChangeType


def _reference_compare(previous, current):
    """Implementação anterior mantida como oráculo de equivalência."""
    from app.excel.comparator import _address_key, _has_content, _values_equal

    changes = []
    for sheet in sorted(set(previous) | set(current)):
        previous_cells = previous.get(sheet, {})
        current_cells = current.get(sheet, {})
        for address in sorted(
            previous_cells.keys() | current_cells.keys(), key=_address_key
        ):
            existed = _has_content(previous_cells, address)
            exists = _has_content(current_cells, address)
            previous_value = previous_cells.get(address)
            new_value = current_cells.get(address)
            if not existed and exists:
                kind = ChangeType.ADD
            elif existed and not exists:
                kind = ChangeType.DEL
            elif existed and exists and not _values_equal(previous_value, new_value):
                kind = ChangeType.MOD
            else:
                continue
            changes.append(
                CellChange(sheet, address, kind, previous_value, new_value)
            )
    return changes


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


def test_optimized_comparator_is_exactly_equivalent_to_reference() -> None:
    randomizer = random.Random(20260921)
    values = [None, "", "texto", 0, 1, False, True, 1.5, "=SUM(A1:A2)"]

    for _ in range(100):
        previous = {}
        current = {}
        for sheet in ("Alfa", "Zeta", "Removida", "Nova"):
            addresses = [
                f"{column}{row}"
                for row in range(1, 25)
                for column in ("A", "B", "Z", "AA")
            ]
            randomizer.shuffle(addresses)
            previous_cells = {
                address: randomizer.choice(values)
                for address in addresses
                if randomizer.random() < 0.55
            }
            randomizer.shuffle(addresses)
            current_cells = {
                address: randomizer.choice(values)
                for address in addresses
                if randomizer.random() < 0.55
            }
            if sheet != "Nova":
                previous[sheet] = previous_cells
            if sheet != "Removida":
                current[sheet] = current_cells

        assert compare_snapshots(previous, current) == _reference_compare(
            previous, current
        )
