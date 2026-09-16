import logging
from pathlib import Path

from app.audit_service import AuditService
from app.database import Database
from app.logging_config import configure_logging
from app.models import AuditExecutionStatus
from app.report_service import ReportService
from app.sources.base import SpreadsheetInfo, VersionInfo
from app.sources.local import LocalSource


def _spreadsheet(versions_path: Path) -> SpreadsheetInfo:
    return SpreadsheetInfo(
        site_id="site-test",
        drive_id="drive-test",
        drive_item_id="item-test",
        name="CQL028.xlsx",
        path=str(versions_path),
    )


def _source(spreadsheet: SpreadsheetInfo, versions_path: Path) -> LocalSource:
    history = [
        (VersionInfo(id=f"version-{path.stem}", number=path.stem), path)
        for path in sorted(versions_path.glob("*.xlsx"))
    ]
    identity = (
        spreadsheet.site_id,
        spreadsheet.drive_id,
        spreadsheet.drive_item_id,
    )
    return LocalSource([spreadsheet], {identity: history})


def test_log_file_tracks_audit_restart_no_updates_and_report(
    tmp_path: Path, cql028_versions: Path
) -> None:
    log_path = tmp_path / "logs" / "audit.log"
    logger = configure_logging(log_path)
    database_path = tmp_path / "database" / "audit.db"
    spreadsheet = _spreadsheet(cql028_versions)

    database = Database(database_path)
    database.initialize()
    first = AuditService(database, _source(spreadsheet, cql028_versions)).audit(spreadsheet)
    row = database.connection.execute("SELECT id FROM planilha").fetchone()
    report = ReportService(database.connection, tmp_path / "reports").generate(row["id"])
    database.close()

    # Simula uma reinicialização real: mesmo arquivo em append, mesmo SQLite.
    configure_logging(log_path)
    reopened = Database(database_path)
    reopened.initialize()
    second = AuditService(reopened, _source(spreadsheet, cql028_versions)).audit(spreadsheet)
    assert reopened.connection.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
    assert reopened.connection.execute("PRAGMA foreign_key_check").fetchall() == []
    reopened.close()

    assert first.status is AuditExecutionStatus.COMPLETED
    assert second.status is AuditExecutionStatus.COMPLETED_WITHOUT_UPDATES
    assert report.is_file()
    content = log_path.read_text(encoding="utf-8")
    assert f"execucao={first.execution_code}" in content
    assert f"execucao={second.execution_code}" in content
    assert "Auditoria concluída" in content
    assert "Auditoria sem novidades" in content
    assert "Relatório gerado" in content
    assert content.count("Auditoria iniciada") == 2

    for handler in logger.handlers:
        handler.flush()


def test_log_records_controlled_failure_and_redacts_secrets(
    tmp_path: Path, cql028_versions: Path, monkeypatch
) -> None:
    log_path = tmp_path / "logs" / "audit.log"
    logger = configure_logging(log_path)
    database = Database(tmp_path / "database" / "audit.db")
    database.initialize()
    spreadsheet = _spreadsheet(cql028_versions)
    source = _source(spreadsheet, cql028_versions)

    def fail(*_args, **_kwargs):
        raise RuntimeError(
            "password=hunter2 token=abc123 Authorization: Bearer xyz Cookie=session-value"
        )

    monkeypatch.setattr(source, "list_versions", fail)
    result = AuditService(database, source).audit(spreadsheet)
    logging.getLogger("auditoria_excel.test").error(
        "client_secret=very-secret bearer standalone-token"
    )
    for handler in logger.handlers:
        handler.flush()
    content = log_path.read_text(encoding="utf-8")

    assert result.status is AuditExecutionStatus.FAILED
    assert "Auditoria falhou" in content
    assert "tipo_erro=RuntimeError" in content
    assert content.count("[REDACTED]") >= 5
    for secret in ("hunter2", "abc123", "xyz", "session-value", "very-secret", "standalone-token"):
        assert secret not in content
    database.close()


def test_log_rotates_instead_of_growing_without_limit(tmp_path: Path) -> None:
    log_path = tmp_path / "nested" / "audit.log"
    logger = configure_logging(log_path, max_bytes=180, backup_count=2)

    for index in range(20):
        logger.info("evento operacional %02d %s", index, "x" * 40)
    for handler in logger.handlers:
        handler.flush()

    assert log_path.is_file()
    assert list(log_path.parent.glob("audit.log.*"))
    assert len(list(log_path.parent.glob("audit.log*"))) <= 3
