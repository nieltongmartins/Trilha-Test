"""Leitura e comparação determinística de pastas de trabalho Excel."""

from app.excel.comparator import CellChange, compare_snapshots
from app.excel.reader import Snapshot, read_workbook

__all__ = ["CellChange", "Snapshot", "compare_snapshots", "read_workbook"]
