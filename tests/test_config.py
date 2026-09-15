from pathlib import Path

from app.config import Settings


def test_settings_create_required_directories(tmp_path: Path) -> None:
    settings = Settings(
        database_path=tmp_path / "database" / "audit.db",
        log_path=tmp_path / "logs" / "audit.log",
        temp_directory=tmp_path / "temp",
        reports_directory=tmp_path / "reports",
    )

    settings.create_directories()

    assert settings.database_path.parent.is_dir()
    assert settings.log_path.parent.is_dir()
    assert settings.temp_directory.is_dir()
    assert settings.reports_directory.is_dir()
