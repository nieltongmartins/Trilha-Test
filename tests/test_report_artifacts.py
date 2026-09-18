from pathlib import Path

import pytest

from app.report_artifacts import ReportArtifactError, ReportArtifactManager
from app.database import Database


@pytest.fixture
def database(tmp_path: Path):
    value = Database(tmp_path / "audit.sqlite3")
    value.initialize()
    try:
        yield value
    finally:
        value.close()


def _insert(database, name: str, unique_id: str) -> int:
    cursor = database.connection.execute(
        """INSERT INTO planilha
        (drive_item_id, nome_atual, site_id, drive_id, caminho_sharepoint)
        VALUES (?, ?, 'site', 'drive', ?)""",
        (unique_id, name, f"/docs/{name}"),
    )
    database.connection.commit()
    return cursor.lastrowid


def test_canonical_identity_does_not_confuse_similar_names(database, tmp_path: Path) -> None:
    first = _insert(database, "TEST001.xlsx", "uid-1")
    second = _insert(database, "TEST001_RENOMEADO.xlsx", "uid-2")
    manager = ReportArtifactManager(database.connection, tmp_path)
    first_path = manager.canonical_path(manager.identity(first))
    second_path = manager.canonical_path(manager.identity(second))
    first_path.write_bytes(b"first")
    second_path.write_bytes(b"second")

    manager.delete_with_database([manager.locate(first)], lambda: database.connection.execute("DELETE FROM planilha WHERE id=?", (first,)))

    assert not first_path.exists()
    assert second_path.read_bytes() == b"second"


def test_legacy_report_is_only_recognized_when_unambiguous(database, tmp_path: Path) -> None:
    first = _insert(database, "Relatório A.xlsx", "uid-1")
    manager = ReportArtifactManager(database.connection, tmp_path)
    legacy = manager.legacy_path(manager.identity(first))
    legacy.write_bytes(b"legacy")
    assert manager.locate(first) == legacy

    _insert(database, "Relatório A.xlsm", "uid-2")
    assert manager.locate(first) is None
    assert legacy.exists()


def test_blocked_report_preserves_database(database, tmp_path: Path, monkeypatch) -> None:
    spreadsheet_id = _insert(database, "Bloqueado.xlsx", "uid-block")
    manager = ReportArtifactManager(database.connection, tmp_path)
    report = manager.canonical_path(manager.identity(spreadsheet_id))
    report.write_bytes(b"xlsx")
    monkeypatch.setattr(Path, "replace", lambda self, target: (_ for _ in ()).throw(PermissionError("open")))

    with pytest.raises(ReportArtifactError, match="preservada"):
        manager.delete_with_database([report], lambda: database.connection.execute("DELETE FROM planilha WHERE id=?", (spreadsheet_id,)))

    assert database.connection.execute("SELECT COUNT(*) FROM planilha WHERE id=?", (spreadsheet_id,)).fetchone()[0] == 1


def test_complete_cleanup_keeps_unknown_file(database, tmp_path: Path) -> None:
    spreadsheet_id = _insert(database, "Controlada.xlsx", "uid-control")
    manager = ReportArtifactManager(database.connection, tmp_path)
    controlled = manager.canonical_path(manager.identity(spreadsheet_id))
    controlled.write_bytes(b"controlled")
    unknown = tmp_path / "manual.xlsx"
    unknown.write_bytes(b"manual")
    paths = manager.controlled_paths([manager.identity(spreadsheet_id)])

    manager.delete_with_database(paths, lambda: database.connection.execute("DELETE FROM planilha"))

    assert not controlled.exists()
    assert unknown.read_bytes() == b"manual"
