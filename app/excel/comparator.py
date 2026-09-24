"""Compara snapshots Excel sem acessar fonte externa ou persistência."""

from dataclasses import dataclass
from functools import lru_cache
import re

from app.excel.reader import CellValue, RowSheetSnapshot, Snapshot
from app.excel.formula_values import normalize_formula_value
from app.models import ChangeType


@dataclass(frozen=True, slots=True)
class CellChange:
    """Diferença imutável entre duas versões de uma célula."""

    sheet: str
    address: str
    change_type: ChangeType
    previous_value: CellValue
    new_value: CellValue


_COORDINATE_PATTERN = re.compile(r"^([A-Z]+)([1-9][0-9]*)$")


@lru_cache(maxsize=4096)
def _column_number(letters: str) -> int:
    number = 0
    for letter in letters:
        number = number * 26 + ord(letter) - ord("A") + 1
    return number


@lru_cache(maxsize=262144)
def _address_key(address: str) -> tuple[int, int, str]:
    """Ordena endereços em ordem Excel e reutiliza chaves já calculadas.

    Em históricos extensos, os mesmos endereços aparecem repetidamente entre
    versões consecutivas. O cache evita repetir regex e conversão da coluna em
    todas as comparações.
    """

    normalized = address.upper()
    match = _COORDINATE_PATTERN.fullmatch(normalized)
    if match is None:
        return (2**31 - 1, 2**31 - 1, address)
    column, row = match.groups()
    return (int(row), _column_number(column), address)


def _has_content(cells: dict[str, CellValue], address: str) -> bool:
    return address in cells and cells[address] is not None


def _values_equal(previous: CellValue, current: CellValue) -> bool:
    previous = normalize_formula_value(previous)
    current = normalize_formula_value(current)
    # Em Python, False == 0 e True == 1. No Excel são tipos distintos.
    if isinstance(previous, bool) or isinstance(current, bool):
        return type(previous) is type(current) and previous == current
    return previous == current


def _compare_cells(
    sheet: str,
    previous_cells: dict[str, CellValue],
    current_cells: dict[str, CellValue],
) -> list[CellChange]:
    """Compara uma unidade já limitada (aba comum ou row alterada)."""
    changes: list[CellChange] = []
    for address, previous_value in previous_cells.items():
        previous_value = normalize_formula_value(previous_value)
        existed = previous_value is not None
        exists = _has_content(current_cells, address)
        if existed and not exists:
            changes.append(CellChange(sheet, address, ChangeType.DEL, previous_value, None))
            continue
        if not existed or not exists:
            continue
        new_value = normalize_formula_value(current_cells[address])
        if not _values_equal(previous_value, new_value):
            changes.append(
                CellChange(sheet, address, ChangeType.MOD, previous_value, new_value)
            )
    for address, new_value in current_cells.items():
        new_value = normalize_formula_value(new_value)
        if new_value is not None and not _has_content(previous_cells, address):
            changes.append(CellChange(sheet, address, ChangeType.ADD, None, new_value))
    return changes


def compare_snapshots(previous: Snapshot, current: Snapshot) -> list[CellChange]:
    """Retorna ADD, DEL e MOD em ordem estável de aba, linha e coluna."""

    changes: list[CellChange] = []
    for sheet in sorted(set(previous) | set(current)):
        previous_cells = previous.get(sheet, {})
        current_cells = current.get(sheet, {})
        # Esta identidade só economiza trabalho quando o leitor incremental
        # compartilhou o objeto após prova SHA-256. Para snapshots comuns ela é
        # apenas um atalho correto (um objeto não pode diferir de si mesmo).
        if previous_cells is current_cells:
            continue
        sheet_changes: list[CellChange] = []
        if isinstance(previous_cells, RowSheetSnapshot) and isinstance(
            current_cells, RowSheetSnapshot
        ):
            previous_rows = {row.key: row.cells for row in previous_cells.rows}
            current_rows = {row.key: row.cells for row in current_cells.rows}
            for row_key in previous_rows.keys() | current_rows.keys():
                old = previous_rows.get(row_key, {})
                new = current_rows.get(row_key, {})
                if old is new:
                    continue
                sheet_changes.extend(_compare_cells(sheet, old, new))
        else:
            # Somente o conjunto normalmente pequeno de diferenças é ordenado.
            sheet_changes.extend(_compare_cells(sheet, previous_cells, current_cells))  # type: ignore[arg-type]

        sheet_changes.sort(key=lambda change: _address_key(change.address))
        changes.extend(sheet_changes)
    return changes
