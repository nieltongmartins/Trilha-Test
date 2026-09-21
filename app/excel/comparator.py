"""Compara snapshots Excel sem acessar fonte externa ou persistência."""

from dataclasses import dataclass
from functools import lru_cache
import re

from app.excel.reader import CellValue, Snapshot
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
    # Em Python, False == 0 e True == 1. No Excel são tipos distintos.
    if isinstance(previous, bool) or isinstance(current, bool):
        return type(previous) is type(current) and previous == current
    return previous == current


def compare_snapshots(previous: Snapshot, current: Snapshot) -> list[CellChange]:
    """Retorna ADD, DEL e MOD em ordem estável de aba, linha e coluna."""

    changes: list[CellChange] = []
    for sheet in sorted(set(previous) | set(current)):
        previous_cells = previous.get(sheet, {})
        current_cells = current.get(sheet, {})
        addresses = previous_cells.keys() | current_cells.keys()

        for address in sorted(addresses, key=_address_key):
            existed = _has_content(previous_cells, address)
            exists = _has_content(current_cells, address)
            previous_value = previous_cells.get(address)
            new_value = current_cells.get(address)

            if not existed and exists:
                change_type = ChangeType.ADD
            elif existed and not exists:
                change_type = ChangeType.DEL
            elif existed and exists and not _values_equal(previous_value, new_value):
                change_type = ChangeType.MOD
            else:
                continue

            changes.append(
                CellChange(
                    sheet=sheet,
                    address=address,
                    change_type=change_type,
                    previous_value=previous_value,
                    new_value=new_value,
                )
            )
    return changes
