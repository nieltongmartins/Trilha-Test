import sqlite3
from pathlib import Path

import pytest

from app.audit_storage import AuditStorageManager, BackupError, RestoreConflictError
from app.database import Database, SCHEMA_VERSION


@pytest.fixture
def populated(tmp_path: Path):
    database = Database(tmp_path / "database" / "audit.db")
    database.initialize()
    connection = database.connection
    spreadsheet_id = connection.execute(
        "INSERT INTO planilha (drive_item_id,nome_atual,site_id,drive_id,caminho_sharepoint) VALUES ('unique-1','CQL001.xlsx','site','drive','/docs/CQL001.xlsx')"
    ).lastrowid
    execution_id = connection.execute(
        "INSERT INTO execucao_auditoria (codigo_execucao,planilha_id,fim,versoes_processadas,alteracoes_encontradas,status) VALUES ('run-1',?,'2026-09-18T10:00:00Z',1,1,'CONCLUIDA')", (spreadsheet_id,)
    ).lastrowid
    version_id = connection.execute(
        "INSERT INTO versao_processada (planilha_id,versao_anterior_id,versao_anterior_numero,versao_atual_id,versao_atual_numero,quantidade_alteracoes,status,hash_origem,execucao_id) VALUES (?,'v1','1.0','v2','1.1',1,'PROCESSADA','abc123',?)", (spreadsheet_id, execution_id)
    ).lastrowid
    connection.execute("INSERT INTO alteracao (versao_processada_id,planilha_id,tipo,aba,endereco,valor_anterior,valor_novo) VALUES (?,?,'MOD','Dados','A1','a','b')", (version_id, spreadsheet_id))
    connection.execute("INSERT INTO checkpoint (planilha_id,versao_id,versao_numero) VALUES (?,'v2','1.1')", (spreadsheet_id,))
    connection.execute("INSERT INTO erro_processamento (execucao_id,planilha_id,tipo_erro,mensagem) VALUES (?,?,'Teste','preservado')", (execution_id, spreadsheet_id))
    connection.commit()
    manager = AuditStorageManager(database, tmp_path / "backups")
    yield database, manager, spreadsheet_id
    database.close()


def test_list_stored_audits_uses_database(populated):
    _, manager, spreadsheet_id = populated
    assert manager.list_audits() == [manager.list_audits()[0]]
    audit = manager.list_audits()[0]
    assert (audit.id, audit.name, audit.last_version, audit.processed_versions, audit.changes) == (spreadsheet_id, "CQL001.xlsx", "1.1", 1, 1)


def test_individual_backup_restore_preserves_checkpoint_hashes_and_foreign_keys(populated):
    database, manager, spreadsheet_id = populated
    backup = manager.backup_individual(spreadsheet_id)
    assert backup.is_file() and backup.with_suffix(backup.suffix + ".sha256").is_file()
    manager.delete_individual(spreadsheet_id)
    restored_id = manager.restore_individual(backup)
    assert database.connection.execute("SELECT versao_numero FROM checkpoint WHERE planilha_id=?", (restored_id,)).fetchone()[0] == "1.1"
    assert database.connection.execute("SELECT hash_origem FROM versao_processada WHERE planilha_id=?", (restored_id,)).fetchone()[0] == "abc123"
    assert database.connection.execute("PRAGMA foreign_key_check").fetchall() == []


def test_individual_restore_detects_conflict_and_can_replace(populated):
    database, manager, spreadsheet_id = populated
    backup = manager.backup_individual(spreadsheet_id)
    with pytest.raises(RestoreConflictError):
        manager.restore_individual(backup)
    manager.restore_individual(backup, replace=True)
    assert database.connection.execute("SELECT COUNT(*) FROM planilha").fetchone()[0] == 1


def test_individual_delete_rolls_back_on_failure(populated, monkeypatch):
    database, manager, spreadsheet_id = populated
    original = database.connection

    class FailingConnection:
        def __enter__(self): return self
        def __exit__(self, kind, value, traceback):
            original.rollback() if kind else original.commit()
        def execute(self, sql, parameters=()):
            if sql.startswith("DELETE FROM checkpoint"):
                raise sqlite3.OperationalError("falha injetada")
            return original.execute(sql, parameters)

    database._connection = FailingConnection()
    with pytest.raises(sqlite3.OperationalError):
        manager.delete_individual(spreadsheet_id)
    database._connection = original
    assert original.execute("SELECT COUNT(*) FROM alteracao").fetchone()[0] == 1


def test_complete_backup_restore_and_safety_backup(populated):
    database, manager, spreadsheet_id = populated
    backup = manager.backup_full()
    manager.delete_all()
    safety = manager.restore_full(backup)
    assert safety.is_file()
    assert database.connection.execute("SELECT id FROM planilha").fetchone()[0] == spreadsheet_id
    assert database.connection.execute("PRAGMA foreign_key_check").fetchall() == []


def test_delete_all_keeps_empty_schema(populated):
    database, manager, _ = populated
    manager.delete_all()
    assert manager.list_audits() == []
    tables = {row[0] for row in database.connection.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    assert {"planilha", "checkpoint", "versao_processada", "alteracao", "execucao_auditoria", "erro_processamento"} <= tables
    assert database.connection.execute("PRAGMA user_version").fetchone()[0] == SCHEMA_VERSION


def test_corrupted_and_incompatible_backups_are_rejected(populated):
    _, manager, spreadsheet_id = populated
    corrupt = manager.backup_individual(spreadsheet_id)
    corrupt.write_bytes(corrupt.read_bytes() + b"damage")
    with pytest.raises(BackupError, match="SHA-256"):
        manager.restore_individual(corrupt)

    incompatible = manager.backup_full()
    metadata = incompatible.with_suffix(incompatible.suffix + ".meta.json")
    metadata.write_text('{"format_version": 999, "backup_type": "complete", "schema_version": 2}', encoding="utf-8")
    with pytest.raises(BackupError, match="incompatível"):
        manager.restore_full(incompatible)
