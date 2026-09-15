"""Converte arquivos ``.xlsx`` em snapshots lógicos comparáveis."""

from pathlib import Path
from typing import TypeAlias

from openpyxl import load_workbook

CellValue: TypeAlias = str | int | float | bool | None
SheetSnapshot: TypeAlias = dict[str, CellValue]
Snapshot: TypeAlias = dict[str, SheetSnapshot]


def read_workbook(path: str | Path) -> Snapshot:
    """Lê um ``.xlsx`` sem modificá-lo e preserva valores de fórmula.

    Células sem conteúdo (valor ``None``) não integram o snapshot. Valores
    falsy significativos, como zero, ``False`` e string vazia, não são
    descartados por teste de veracidade.
    """

    workbook_path = Path(path)
    if workbook_path.suffix.lower() != ".xlsx":
        raise ValueError("O leitor aceita somente arquivos .xlsx")

    workbook = load_workbook(
        filename=workbook_path,
        read_only=True,
        data_only=False,
    )
    try:
        snapshot: Snapshot = {}
        for worksheet in workbook.worksheets:
            cells: SheetSnapshot = {}
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        cells[cell.coordinate] = cell.value
            snapshot[worksheet.title] = cells
        return snapshot
    finally:
        workbook.close()
