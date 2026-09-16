from pathlib import Path

import pytest

from app.excel.reader import read_workbook


def test_reader_preserves_formulas_values_and_sheets(
    cql028_versions: Path,
) -> None:
    snapshot = read_workbook(cql028_versions / "0.85.xlsx")

    assert list(snapshot) == ["Resumo", "Detalhes"]
    assert snapshot["Resumo"] == {
        "A1": 15,
        "B2": "OK",
        "C3": "Pendente",
        "D4": 0,
        "E5": False,
        "F6": "=SUM(A1:A20)",
    }
    assert snapshot["Detalhes"] == {"B2": "Nova informação"}


def test_reader_ignores_explicitly_empty_cells(cql028_versions: Path) -> None:
    snapshot = read_workbook(cql028_versions / "0.84.xlsx")

    # B2 é gravada como None pela fixture programática; openpyxl a materializa
    # como uma célula vazia e o snapshot não deve preservá-la.
    assert "B2" not in snapshot["Resumo"]
    assert "Detalhes" in snapshot
    assert snapshot["Detalhes"] == {}


def test_reader_rejects_non_xlsx_file(tmp_path: Path) -> None:
    path = tmp_path / "arquivo.xls"
    path.touch()

    with pytest.raises(ValueError, match="somente arquivos .xlsx"):
        read_workbook(path)
