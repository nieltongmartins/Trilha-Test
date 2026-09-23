"""Leitura e comparação determinística de pastas de trabalho Excel."""

from app.excel.comparator import CellChange, compare_snapshots
from app.excel.reader import ConsecutiveWorkbookReader, Snapshot, read_workbook

__all__ = [
    "CellChange", "ConsecutiveWorkbookReader", "Snapshot",
    "compare_snapshots", "read_workbook",
]
