import shutil
from collections.abc import Iterator, Mapping
from pathlib import Path

import pytest


WorkbookState = Mapping[str, Mapping[str, object]]


@pytest.fixture(autouse=True)
def discard_temporary_xlsx(tmp_path: Path) -> Iterator[None]:
    """Garante que nenhuma planilha temporária sobreviva ao teste."""

    yield

    for workbook_path in tmp_path.rglob("*.xlsx"):
        workbook_path.unlink(missing_ok=True)


def _write_workbook(path: Path, sheets: WorkbookState) -> None:
    """Cria uma versão Excel controlada no diretório temporário do teste."""

    from openpyxl import Workbook

    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    for sheet_name, cells in sheets.items():
        worksheet = workbook.create_sheet(sheet_name)
        for address, value in cells.items():
            worksheet[address] = value

    workbook.save(path)
    workbook.close()


@pytest.fixture
def cql028_versions(tmp_path: Path) -> Iterator[Path]:
    """Gera e descarta quatro versões da CQL028 no diretório temporário."""

    versions_path = tmp_path / "CQL028"
    versions_path.mkdir()

    versions: dict[str, WorkbookState] = {
        "0.84.xlsx": {
            "Resumo": {
                "A1": 10,
                # Célula vazia explícita: deve ser descartada pelo reader sem
                # confundir esse caso com os valores válidos 0 e False.
                "B2": None,
                "C3": "Pendente",
                "F6": "=SUM(A1:A10)",
            },
            "Detalhes": {},
        },
        "0.85.xlsx": {
            "Resumo": {
                "A1": 15,
                "B2": "OK",
                "C3": "Pendente",
                "D4": 0,
                "E5": False,
                "F6": "=SUM(A1:A20)",
            },
            "Detalhes": {"B2": "Nova informação"},
        },
        "0.86.xlsx": {
            "Resumo": {
                "A1": 15,
                "B2": "OK",
                "D4": 0,
                "E5": False,
                "F6": "=SUM(A1:A20)",
            },
            "Detalhes": {"B2": "Alterada"},
            "Nova Aba": {"A1": "Criada"},
        },
        "0.87.xlsx": {
            "Resumo": {
                "A1": 15,
                "B2": "OK",
                "D4": 0,
                "E5": False,
                "F6": "=SUM(A1:A20)",
            },
            "Detalhes": {"B2": "Alterada"},
            "Nova Aba": {"A1": "Criada"},
        },
    }

    try:
        for filename, state in versions.items():
            _write_workbook(versions_path / filename, state)

        yield versions_path
    finally:
        shutil.rmtree(versions_path, ignore_errors=True)
