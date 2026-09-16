from pathlib import Path

from openpyxl import load_workbook
import pytest

from app.database import Database
from app.report_service import ReportService


@pytest.fixture
def populated_database(tmp_path: Path):
    with Database(tmp_path / "audit.db") as database:
        database.initialize()
        connection = database.connection
        spreadsheet_id = connection.execute(
            """INSERT INTO planilha
               (drive_item_id, nome_atual, site_id, drive_id, caminho_sharepoint)
               VALUES ('item-1', 'CQL028.xlsx', 'site-1', 'drive-1', '/docs/CQL028.xlsx')"""
        ).lastrowid
        execution_id = connection.execute(
            """INSERT INTO execucao_auditoria
               (codigo_execucao, planilha_id, status, fim, versoes_processadas,
                alteracoes_encontradas)
               VALUES ('AUD-1', ?, 'CONCLUIDA', CURRENT_TIMESTAMP, 1, 2)""",
            (spreadsheet_id,),
        ).lastrowid
        version_id = connection.execute(
            """INSERT INTO versao_processada
               (planilha_id, versao_anterior_id, versao_anterior_numero,
                versao_atual_id, versao_atual_numero, data_hora_versao, autor,
                comentario, quantidade_alteracoes, status, execucao_id)
               VALUES (?, 'v1', '0.1', 'v2', '0.2', '2026-09-16T10:00:00Z',
                       'Maria', 'Ajuste', 2, 'PROCESSADA', ?)""",
            (spreadsheet_id, execution_id),
        ).lastrowid
        connection.executemany(
            """INSERT INTO alteracao
               (versao_processada_id, planilha_id, tipo, aba, endereco,
                valor_anterior, valor_novo) VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (
                (version_id, spreadsheet_id, "ADD", "Dados", "A1", None, "novo"),
                (version_id, spreadsheet_id, "MOD", "Dados", "B2", "1", "2"),
            ),
        )
        connection.commit()
        yield database, spreadsheet_id


def test_report_contains_official_sheets_filters_and_database_data(
    populated_database, tmp_path: Path
) -> None:
    database, spreadsheet_id = populated_database

    report = ReportService(database.connection, tmp_path / "reports").generate(
        spreadsheet_id
    )
    workbook = load_workbook(report)

    assert report.name == "CQL028_Trilha_Auditoria.xlsx"
    assert workbook.sheetnames == ["RESUMO", "VERSOES", "TRILHA"]
    assert workbook["RESUMO"]["B2"].value == "CQL028.xlsx"
    assert workbook["RESUMO"]["B9"].value == 2
    assert workbook["VERSOES"]["A2"].value == "0.1"
    assert workbook["TRILHA"]["I2"].value == "ADD"
    assert all(sheet.auto_filter.ref for sheet in workbook.worksheets)


def test_report_can_be_regenerated_and_rejects_unknown_spreadsheet(
    populated_database, tmp_path: Path
) -> None:
    database, spreadsheet_id = populated_database
    service = ReportService(database.connection, tmp_path)

    first = service.generate(spreadsheet_id)
    first.unlink()
    second = service.generate(spreadsheet_id)

    assert second.is_file()
    with pytest.raises(ValueError, match="não encontrada"):
        service.generate(999)
