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
    first = service.audit(SPREADSHEET)
    connection = database.connection
    checkpoint_before = tuple(connection.execute("SELECT * FROM checkpoint").fetchone())
    processed_before = [
        tuple(row)
        for row in connection.execute("SELECT * FROM versao_processada ORDER BY id")
    ]
    changes_before = [
        tuple(row) for row in connection.execute("SELECT * FROM alteracao ORDER BY id")
    ]

    # Simula o encerramento da aplicação. Os XLSX podem desaparecer depois da
    # primeira auditoria: sem novidades, a reexecução não deve tentar relê-los.
    database.close()
    for _, path in local_history:
        path.unlink()
    database.initialize()

    result = AuditService(database, source(local_history)).audit(SPREADSHEET)

    assert first.status is AuditExecutionStatus.COMPLETED
    assert result.status is AuditExecutionStatus.COMPLETED_WITHOUT_UPDATES
    assert (
        result.processed_versions,
        result.changes,
        result.initial_checkpoint,
        result.final_version,
    ) == (0, 0, "0.99", "0.99")

    connection = database.connection
    assert tuple(connection.execute("SELECT * FROM checkpoint").fetchone()) == (
        checkpoint_before
    )
    assert [
        tuple(row)
        for row in connection.execute("SELECT * FROM versao_processada ORDER BY id")
    ] == processed_before
    assert [
        tuple(row) for row in connection.execute("SELECT * FROM alteracao ORDER BY id")
    ] == changes_before
    assert scalar(connection, "SELECT COUNT(*) FROM planilha") == 1

    executions = connection.execute(
        """SELECT codigo_execucao, checkpoint_inicial, versao_final,
                  versoes_processadas, alteracoes_encontradas, status, fim, mensagem
           FROM execucao_auditoria ORDER BY id"""
    ).fetchall()
    assert len(executions) == 2
    assert executions[0][0] != executions[1][0]
    assert executions[0][5] == "CONCLUIDA"
    assert tuple(executions[1][1:6]) == (
        "0.99", "0.99", 0, 0, "CONCLUIDA_SEM_NOVIDADES",
    )
    assert executions[1][6] is not None and executions[1][7] is None
    assert connection.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
    assert connection.execute("PRAGMA foreign_key_check").fetchall() == []


def test_incremental_audit_keeps_base_and_processes_only_new_pairs(
    database: Database, local_history: list[tuple[VersionInfo, Path]], tmp_path: Path,
) -> None:
    AuditService(database, source(local_history)).audit(SPREADSHEET)
    processed_before = [
        tuple(row)
        for row in database.connection.execute(
            "SELECT * FROM versao_processada ORDER BY id"
        )
    ]
    changes_before = [
        tuple(row)
        for row in database.connection.execute("SELECT * FROM alteracao ORDER BY id")
    ]

    # Em uma continuação real, as versões anteriores ao checkpoint não precisam mais
    # estar disponíveis localmente. A versão 0.99 permanece como base de 0.99 -> 1.00.
    for _, path in local_history[:-1]:
        path.unlink()
    extended = list(local_history)
    for number, value in [("1.00", "Um"), ("1.01", "Dois"), ("1.02", "Três")]:
        path = tmp_path / f"{number}.xlsx"
        make_workbook(path, {"Dados": {"A1": value, "C1": "Adicionado"}})
        extended.append((version(number), path))

    incremental_source = source(extended)
    acquired: list[str] = []
    original_get_version = incremental_source.get_version

    def record_get_version(
        spreadsheet: SpreadsheetInfo, item: VersionInfo,
    ) -> Path:
        acquired.append(item.number)
        return original_get_version(spreadsheet, item)

    incremental_source.get_version = record_get_version  # type: ignore[method-assign]
    result = AuditService(database, incremental_source).audit(SPREADSHEET)

    assert result.status is AuditExecutionStatus.COMPLETED
    assert (
        result.processed_versions,
        result.changes,
        result.initial_checkpoint,
        result.final_version,
    ) == (3, 3, "0.99", "1.02")
    assert acquired == ["0.99", "1.00", "1.01", "1.02"]

    connection = database.connection
    processed = [
        tuple(row)
        for row in connection.execute("SELECT * FROM versao_processada ORDER BY id")
    ]
    changes = [
        tuple(row) for row in connection.execute("SELECT * FROM alteracao ORDER BY id")
    ]
    assert processed[:3] == processed_before
    assert changes[:4] == changes_before
    assert [(row[3], row[5]) for row in processed[-3:]] == [
        ("0.99", "1.00"), ("1.00", "1.01"), ("1.01", "1.02")
    ]
    assert len(processed) == 6
    assert len(changes) == 7
    checkpoint = connection.execute(
        "SELECT versao_id, versao_numero FROM checkpoint"
    ).fetchone()
    assert tuple(checkpoint) == ("version-102", "1.02")
    execution = connection.execute(
        """SELECT checkpoint_inicial, versao_final, versoes_processadas,
                  alteracoes_encontradas, status, fim, mensagem
           FROM execucao_auditoria ORDER BY id DESC"""
    ).fetchone()
    assert tuple(execution[:5]) == ("0.99", "1.02", 3, 3, "CONCLUIDA")
    assert execution[5] is not None and execution[6] is None
    assert connection.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
    assert connection.execute("PRAGMA foreign_key_check").fetchall() == []


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
