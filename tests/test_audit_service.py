from pathlib import Path
import sqlite3

from openpyxl import Workbook
import pytest

from app.audit_service import AuditService
from app.database import Database
from app.models import AuditExecutionStatus
from app.sources import LocalSource, SpreadsheetInfo, VersionInfo


SPREADSHEET = SpreadsheetInfo(
    site_id="local-site",
    drive_id="local-drive",
    drive_item_id="CQL028-id",
    name="CQL028.xlsx",
    path="dados_teste/CQL028",
)


def make_workbook(path: Path, sheets: dict[str, dict[str, object]]) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    for title, cells in sheets.items():
        worksheet = workbook.create_sheet(title)
        for address, value in cells.items():
            worksheet[address] = value
    workbook.save(path)
    workbook.close()


def version(number: str) -> VersionInfo:
    compact = number.replace(".", "")
    return VersionInfo(
        id=f"version-{compact}",
        number=number,
        modified_at=f"2026-09-15T12:{compact[-2:]}:00Z",
        author="Pessoa de Teste",
        comment=f"Versão {number}",
    )


@pytest.fixture
def database(tmp_path: Path) -> Database:
    with Database(tmp_path / "audit.db") as instance:
        instance.initialize()
        yield instance


@pytest.fixture
def local_history(tmp_path: Path) -> list[tuple[VersionInfo, Path]]:
    definitions = [
        ("0.84", {"Dados": {"A1": "Inicial", "B1": "Excluir"}}),
        ("0.85", {"Dados": {"A1": "Modificado", "C1": "Adicionado"}}),
        # Comparação sem alteração, que ainda precisa ser registrada.
        ("0.86", {"Dados": {"A1": "Modificado", "C1": "Adicionado"}}),
        ("0.99", {"Dados": {"A1": "Final", "C1": "Adicionado"}}),
    ]
    history = []
    for number, sheets in definitions:
        path = tmp_path / f"{number}.xlsx"
        make_workbook(path, sheets)
        history.append((version(number), path))
    return history


def source(history: list[tuple[VersionInfo, Path]]) -> LocalSource:
    identity = (SPREADSHEET.site_id, SPREADSHEET.drive_id, SPREADSHEET.drive_item_id)
    return LocalSource([SPREADSHEET], {identity: history})


def scalar(connection: sqlite3.Connection, query: str) -> int:
    return int(connection.execute(query).fetchone()[0])


def test_local_source_lists_identity_versions_and_historical_file(
    local_history: list[tuple[VersionInfo, Path]],
) -> None:
    local_source = source(local_history)

    assert local_source.list_spreadsheets() == (SPREADSHEET,)
    assert [item.number for item in local_source.list_versions(SPREADSHEET)] == [
        "0.84", "0.85", "0.86", "0.99"
    ]
    assert local_source.get_version(SPREADSHEET, version("0.84")) == local_history[0][1]


def test_complete_audit_records_changes_empty_version_checkpoint_and_execution(
    database: Database, local_history: list[tuple[VersionInfo, Path]],
) -> None:
    result = AuditService(database, source(local_history)).audit(SPREADSHEET)
    connection = database.connection

    assert result.status is AuditExecutionStatus.COMPLETED
    assert (result.processed_versions, result.changes, result.final_version) == (3, 4, "0.99")
    checkpoint = connection.execute("SELECT * FROM checkpoint").fetchone()
    assert (checkpoint["versao_id"], checkpoint["versao_numero"]) == (
        "version-099", "0.99"
    )
    processed = connection.execute(
        "SELECT versao_atual_numero, status, quantidade_alteracoes "
        "FROM versao_processada ORDER BY id"
    ).fetchall()
    assert [tuple(row) for row in processed] == [
        ("0.85", "PROCESSADA", 3),
        ("0.86", "SEM_ALTERACOES", 0),
        ("0.99", "PROCESSADA", 1),
    ]
    assert {
        tuple(row) for row in connection.execute(
            "SELECT tipo, aba, endereco FROM alteracao"
        )
    } >= {("MOD", "Dados", "A1"), ("DEL", "Dados", "B1"), ("ADD", "Dados", "C1")}
    execution = connection.execute("SELECT * FROM execucao_auditoria").fetchone()
    assert execution["status"] == "CONCLUIDA"
    assert execution["fim"] is not None


def test_reexecution_without_new_versions_is_idempotent(
    database: Database, local_history: list[tuple[VersionInfo, Path]],
) -> None:
    service = AuditService(database, source(local_history))
    service.audit(SPREADSHEET)
    result = service.audit(SPREADSHEET)

    assert result.status is AuditExecutionStatus.COMPLETED_WITHOUT_UPDATES
    assert (result.processed_versions, result.changes, result.final_version) == (0, 0, "0.99")
    assert scalar(database.connection, "SELECT COUNT(*) FROM versao_processada") == 3
    assert scalar(database.connection, "SELECT COUNT(*) FROM alteracao") == 4
    statuses = database.connection.execute(
        "SELECT status FROM execucao_auditoria ORDER BY id"
    ).fetchall()
    assert [row[0] for row in statuses] == ["CONCLUIDA", "CONCLUIDA_SEM_NOVIDADES"]


def test_incremental_audit_keeps_base_and_processes_only_new_pairs(
    database: Database, local_history: list[tuple[VersionInfo, Path]], tmp_path: Path,
) -> None:
    AuditService(database, source(local_history)).audit(SPREADSHEET)
    extended = list(local_history)
    for number, value in [("1.00", "Um"), ("1.01", "Dois"), ("1.02", "Três")]:
        path = tmp_path / f"{number}.xlsx"
        make_workbook(path, {"Dados": {"A1": value, "C1": "Adicionado"}})
        extended.append((version(number), path))

    result = AuditService(database, source(extended)).audit(SPREADSHEET)

    assert (result.processed_versions, result.final_version) == (3, "1.02")
    pairs = database.connection.execute(
        "SELECT versao_anterior_numero, versao_atual_numero "
        "FROM versao_processada ORDER BY id"
    ).fetchall()
    assert [tuple(row) for row in pairs[-3:]] == [
        ("0.99", "1.00"), ("1.00", "1.01"), ("1.01", "1.02")
    ]
    assert scalar(database.connection, "SELECT COUNT(*) FROM versao_processada") == 6


def test_failure_rolls_back_pair_keeps_last_checkpoint_and_can_resume(
    database: Database, local_history: list[tuple[VersionInfo, Path]], tmp_path: Path,
) -> None:
    AuditService(database, source(local_history)).audit(SPREADSHEET)
    extended = list(local_history)
    for number in ("1.00", "1.01"):
        path = tmp_path / f"{number}.xlsx"
        make_workbook(path, {"Dados": {"A1": number}})
        extended.append((version(number), path))
    missing_path = tmp_path / "1.02.xlsx"
    extended.append((version("1.02"), missing_path))

    failed = AuditService(database, source(extended)).audit(SPREADSHEET)

    assert failed.status is AuditExecutionStatus.FAILED
    assert (failed.processed_versions, failed.final_version) == (2, "1.01")
    assert database.connection.execute(
        "SELECT versao_numero FROM checkpoint"
    ).fetchone()[0] == "1.01"
    execution = database.connection.execute(
        "SELECT status, versoes_processadas FROM execucao_auditoria ORDER BY id DESC"
    ).fetchone()
    assert tuple(execution) == ("FALHA", 2)
    error = database.connection.execute(
        "SELECT versao_anterior, versao_atual, tipo_erro FROM erro_processamento"
    ).fetchone()
    assert tuple(error) == ("1.01", "1.02", "FileNotFoundError")
    assert not database.connection.execute(
        "SELECT 1 FROM versao_processada WHERE versao_atual_numero = '1.02'"
    ).fetchone()

    make_workbook(missing_path, {"Dados": {"A1": "recuperado"}})
    resumed = AuditService(database, source(extended)).audit(SPREADSHEET)

    assert resumed.status is AuditExecutionStatus.COMPLETED
    assert (resumed.processed_versions, resumed.final_version) == (1, "1.02")
    assert scalar(database.connection, "SELECT COUNT(*) FROM versao_processada") == 6
