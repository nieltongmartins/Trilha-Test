from dataclasses import replace
from pathlib import Path
import sqlite3

from openpyxl import Workbook
import pytest

from app.audit_service import AuditService
from app.database import Database
from app.integrity import sha256_file
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
        "SELECT versao_atual_numero, status, quantidade_alteracoes, hash_origem "
        "FROM versao_processada ORDER BY id"
    ).fetchall()
    assert [tuple(row) for row in processed] == [
        ("0.85", "PROCESSADA", 3, sha256_file(local_history[1][1])),
        ("0.86", "SEM_ALTERACOES", 0, sha256_file(local_history[2][1])),
        ("0.99", "PROCESSADA", 1, sha256_file(local_history[3][1])),
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


def test_hashes_content_of_sharepoint_current_version(
    database: Database, local_history: list[tuple[VersionInfo, Path]],
) -> None:
    history = list(local_history)
    current_info, current_path = history[-1]
    history[-1] = (replace(current_info, is_current=True), current_path)

    AuditService(database, source(history)).audit(SPREADSHEET)

    row = database.connection.execute(
        """SELECT versao_atual_id, versao_atual, hash_origem
           FROM versao_processada ORDER BY id DESC LIMIT 1"""
    ).fetchone()
    assert tuple(row) == (current_info.id, 1, sha256_file(current_path))


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
    assert [row["hash_origem"] for row in connection.execute(
        "SELECT hash_origem FROM versao_processada ORDER BY id DESC LIMIT 3"
    )][::-1] == [sha256_file(path) for _, path in extended[-3:]]
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
    # Consolida o estado anterior até 1.00, que será o checkpoint válido da prova.
    initial_history = list(local_history)
    initial_path = tmp_path / "1.00.xlsx"
    make_workbook(initial_path, {"Dados": {"A1": "1.00"}})
    initial_history.append((version("1.00"), initial_path))
    AuditService(database, source(initial_history)).audit(SPREADSHEET)

    extended = list(initial_history)
    for number in ("1.01", "1.02", "1.03"):
        path = tmp_path / f"{number}.xlsx"
        make_workbook(path, {"Dados": {"A1": number}})
        extended.append((version(number), path))

    # A falha é injetada no último passo da unidade transacional oficial, depois
    # dos INSERTs de versão/alterações e antes da confirmação do checkpoint 1.02.
    # Isso comprova que dados já escritos pela comparação também sofrem rollback.
    connection = database.connection
    connection.execute(
        """
        CREATE TRIGGER fail_checkpoint_102
        BEFORE UPDATE ON checkpoint
        WHEN NEW.versao_numero = '1.02'
        BEGIN
            SELECT RAISE(ABORT, 'falha controlada ao atualizar checkpoint 1.02');
        END
        """
    )
    connection.commit()
    failing_source = source(extended)
    acquired_before_failure: list[str] = []
    original_get_version = failing_source.get_version

    def record_failed_attempt(
        spreadsheet: SpreadsheetInfo, item: VersionInfo,
    ) -> Path:
        acquired_before_failure.append(item.number)
        return original_get_version(spreadsheet, item)

    failing_source.get_version = record_failed_attempt  # type: ignore[method-assign]
    failed = AuditService(database, failing_source).audit(SPREADSHEET)

    assert failed.status is AuditExecutionStatus.FAILED
    assert (
        failed.processed_versions,
        failed.changes,
        failed.initial_checkpoint,
        failed.final_version,
    ) == (1, 1, "1.00", "1.01")
    assert acquired_before_failure == ["1.00", "1.01", "1.02"]
    assert connection.execute(
        "SELECT versao_id, versao_numero FROM checkpoint"
    ).fetchone()[:] == ("version-101", "1.01")
    failed_execution = connection.execute(
        """SELECT id, checkpoint_inicial, versao_final, versoes_processadas,
                  alteracoes_encontradas, status, fim, mensagem
           FROM execucao_auditoria ORDER BY id DESC"""
    ).fetchone()
    assert tuple(failed_execution[1:6]) == ("1.00", "1.01", 1, 1, "FALHA")
    assert failed_execution[6] is not None
    assert failed_execution[7] == "falha controlada ao atualizar checkpoint 1.02"
    error = connection.execute(
        """SELECT execucao_id, versao_anterior, versao_atual, tipo_erro, mensagem
           FROM erro_processamento"""
    ).fetchone()
    assert tuple(error) == (
        failed_execution[0], "1.01", "1.02", "IntegrityError",
        "falha controlada ao atualizar checkpoint 1.02",
    )
    assert [
        tuple(row)
        for row in connection.execute(
            """SELECT versao_anterior_numero, versao_atual_numero
               FROM versao_processada ORDER BY id"""
        )
    ][-2:] == [("0.99", "1.00"), ("1.00", "1.01")]
    assert not connection.execute(
        """SELECT 1 FROM versao_processada
           WHERE versao_atual_numero IN ('1.02', '1.03')"""
    ).fetchone()
    assert not connection.execute(
        """SELECT 1 FROM alteracao a JOIN versao_processada v
               ON v.id = a.versao_processada_id
           WHERE v.versao_atual_numero IN ('1.02', '1.03')"""
    ).fetchone()
    assert scalar(connection, "SELECT COUNT(*) FROM alteracao") == 7
    assert connection.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
    assert connection.execute("PRAGMA foreign_key_check").fetchall() == []

    connection.execute("DROP TRIGGER fail_checkpoint_102")
    connection.commit()
    resumed_source = source(extended)
    acquired_on_resume: list[str] = []
    original_resume_get_version = resumed_source.get_version

    def record_resume(
        spreadsheet: SpreadsheetInfo, item: VersionInfo,
    ) -> Path:
        acquired_on_resume.append(item.number)
        return original_resume_get_version(spreadsheet, item)

    resumed_source.get_version = record_resume  # type: ignore[method-assign]
    resumed = AuditService(database, resumed_source).audit(SPREADSHEET)

    assert resumed.status is AuditExecutionStatus.COMPLETED
    assert (
        resumed.processed_versions,
        resumed.changes,
        resumed.initial_checkpoint,
        resumed.final_version,
    ) == (2, 2, "1.01", "1.03")
    assert acquired_on_resume == ["1.01", "1.02", "1.03"]
    assert scalar(connection, "SELECT COUNT(*) FROM versao_processada") == 7
    assert scalar(connection, "SELECT COUNT(*) FROM alteracao") == 9
    assert connection.execute(
        "SELECT versao_id, versao_numero FROM checkpoint"
    ).fetchone()[:] == ("version-103", "1.03")
    assert [
        tuple(row)
        for row in connection.execute(
            """SELECT versao_anterior_numero, versao_atual_numero, COUNT(*)
               FROM versao_processada
               GROUP BY versao_anterior_numero, versao_atual_numero
               ORDER BY MIN(id)"""
        )
    ][-3:] == [
        ("1.00", "1.01", 1), ("1.01", "1.02", 1), ("1.02", "1.03", 1),
    ]
    resumed_execution = connection.execute(
        """SELECT checkpoint_inicial, versao_final, versoes_processadas,
                  alteracoes_encontradas, status, fim, mensagem
           FROM execucao_auditoria ORDER BY id DESC"""
    ).fetchone()
    assert tuple(resumed_execution[:5]) == ("1.01", "1.03", 2, 2, "CONCLUIDA")
    assert resumed_execution[5] is not None and resumed_execution[6] is None
    assert scalar(connection, "SELECT COUNT(*) FROM erro_processamento") == 1
    assert connection.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
    assert connection.execute("PRAGMA foreign_key_check").fetchall() == []
