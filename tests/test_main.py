from pathlib import Path

import main


def test_application_starts_and_creates_database(tmp_path: Path, monkeypatch) -> None:
    database_path = tmp_path / "database" / "audit.db"
    log_path = tmp_path / "logs" / "audit.log"
    monkeypatch.setenv("AUDIT_DATABASE_PATH", str(database_path))
    monkeypatch.setenv("AUDIT_LOG_PATH", str(log_path))

    assert main.main(launch_ui=False) == 0
    assert database_path.is_file()
    assert log_path.is_file()
