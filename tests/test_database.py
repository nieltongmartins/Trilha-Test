import sqlite3
from pathlib import Path

import pytest

from app.database import Database
from app.exceptions import DatabaseNotConnectedError


EXPECTED_TABLES = {
    "planilha",
    "checkpoint",
    "versao_processada",
    "alteracao",
    "execucao_auditoria",
    "erro_processamento",
}


@pytest.fixture
def database(tmp_path: Path) -> Database:
    with Database(tmp_path / "database" / "audit.db") as instance:
        instance.initialize()
        yield instance


def insert_spreadsheet(connection: sqlite3.Connection) -> int:
    cursor = connection.execute(
        """
        INSERT INTO planilha (drive_item_id, nome_atual, site_id, drive_id)
        VALUES (?, ?, ?, ?)
        """,
        ("item-1", "CQL028.xlsx", "site-1", "drive-1"),
    )
    connection.commit()
    assert cursor.lastrowid is not None
    return cursor.lastrowid


def test_initialize_creates_database_and_all_official_tables(tmp_path: Path) -> None:
    path = tmp_path / "nested" / "audit.db"

    with Database(path) as database:
        database.initialize()
        tables = {
            row[0]
            for row in database.connection.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'"
            )
        }

    assert path.is_file()
    assert EXPECTED_TABLES <= tables


def test_repeated_initialization_preserves_existing_data(database: Database) -> None:
    spreadsheet_id = insert_spreadsheet(database.connection)

    database.initialize()

    row = database.connection.execute(
        "SELECT nome_atual FROM planilha WHERE id = ?", (spreadsheet_id,)
    ).fetchone()
    assert row["nome_atual"] == "CQL028.xlsx"


def test_identity_and_checkpoint_constraints_prevent_duplicates(
    database: Database,
) -> None:
    spreadsheet_id = insert_spreadsheet(database.connection)

    with pytest.raises(sqlite3.IntegrityError):
        database.connection.execute(
            """
            INSERT INTO planilha (drive_item_id, nome_atual, site_id, drive_id)
            VALUES ('item-1', 'Renomeada.xlsx', 'site-1', 'drive-1')
            """
        )
    database.connection.rollback()

    database.connection.execute(
        """
        INSERT INTO checkpoint (planilha_id, versao_id, versao_numero)
        VALUES (?, 'version-1', '0.99')
        """,
        (spreadsheet_id,),
    )
    database.connection.commit()

    with pytest.raises(sqlite3.IntegrityError):
        database.connection.execute(
            """
            INSERT INTO checkpoint (planilha_id, versao_id, versao_numero)
            VALUES (?, 'version-2', '1.00')
            """,
            (spreadsheet_id,),
        )


def test_foreign_keys_and_controlled_values_are_enforced(database: Database) -> None:
    with pytest.raises(sqlite3.IntegrityError):
        database.connection.execute(
            """
            INSERT INTO checkpoint (planilha_id, versao_id, versao_numero)
            VALUES (999, 'version-1', '1.0')
            """
        )
    database.connection.rollback()

    spreadsheet_id = insert_spreadsheet(database.connection)
    with pytest.raises(sqlite3.IntegrityError):
        database.connection.execute(
            """
            INSERT INTO execucao_auditoria
                (codigo_execucao, planilha_id, status)
            VALUES ('AUD-1', ?, 'STATUS_INVALIDO')
            """,
            (spreadsheet_id,),
        )


def test_processed_comparison_and_change_cannot_be_duplicated(
    database: Database,
) -> None:
    spreadsheet_id = insert_spreadsheet(database.connection)
    execution_id = database.connection.execute(
        """
        INSERT INTO execucao_auditoria (codigo_execucao, planilha_id, status)
        VALUES ('AUD-1', ?, 'EM_EXECUCAO')
        """,
        (spreadsheet_id,),
    ).lastrowid
    version_values = (
        spreadsheet_id,
        "v099",
        "0.99",
        "v100",
        "1.00",
        execution_id,
    )
    version_id = database.connection.execute(
        """
        INSERT INTO versao_processada (
            planilha_id, versao_anterior_id, versao_anterior_numero,
            versao_atual_id, versao_atual_numero, status, execucao_id
        ) VALUES (?, ?, ?, ?, ?, 'PROCESSADA', ?)
        """,
        version_values,
    ).lastrowid
    assert version_id is not None
    database.connection.commit()

    with pytest.raises(sqlite3.IntegrityError):
        database.connection.execute(
            """
            INSERT INTO versao_processada (
                planilha_id, versao_anterior_id, versao_anterior_numero,
                versao_atual_id, versao_atual_numero, status, execucao_id
            ) VALUES (?, ?, ?, ?, ?, 'PROCESSADA', ?)
            """,
            version_values,
        )
    database.connection.rollback()

    change_values = (version_id, spreadsheet_id, "Plan1", "A1")
    database.connection.execute(
        """
        INSERT INTO alteracao (
            versao_processada_id, planilha_id, tipo, aba, endereco, valor_novo
        ) VALUES (?, ?, 'ADD', ?, ?, 'novo')
        """,
        change_values,
    )
    database.connection.commit()
    with pytest.raises(sqlite3.IntegrityError):
        database.connection.execute(
            """
            INSERT INTO alteracao (
                versao_processada_id, planilha_id, tipo, aba, endereco, valor_novo
            ) VALUES (?, ?, 'ADD', ?, ?, 'duplicado')
            """,
            change_values,
        )


def test_context_manager_closes_connection(tmp_path: Path) -> None:
    database = Database(tmp_path / "audit.db")

    with database:
        database.initialize()
        assert database.is_connected

    assert not database.is_connected
    with pytest.raises(DatabaseNotConnectedError):
        _ = database.connection
