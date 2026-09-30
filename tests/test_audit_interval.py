from pathlib import Path

from openpyxl import Workbook

from app.audit_interval import get_audit_coverage, resolve_interval
from app.audit_service import AuditService
from app.database import Database
from app.sources.base import SpreadsheetInfo, VersionInfo
from app.sources.local import LocalSource


SHEET = SpreadsheetInfo("site", "drive", "interval-item", "Intervalo.xlsx")


def _history(tmp_path: Path, count: int = 7):
    result = []
    for number in range(1, count + 1):
        path = tmp_path / f"{number}.xlsx"
        workbook = Workbook()
        workbook.active["A1"] = number
        workbook.save(path)
        result.append((VersionInfo(str(number), f"9.{number}"), path))
    return result


def _source(history):
    return LocalSource([SHEET], {(SHEET.site_id, SHEET.drive_id, SHEET.drive_item_id): history})


def test_interval_reuses_overlap_and_does_not_move_official_checkpoint(tmp_path):
    history = _history(tmp_path)
    with Database(tmp_path / "audit.db") as database:
        database.initialize()
        service = AuditService(database, _source(history))
        first = service.audit(SHEET, start_version_id="1", end_version_id="4")
        second = service.audit(SHEET, start_version_id="3", end_version_id="7")

        assert first.processed_versions == 3
        assert second.processed_versions == 3  # 3->4 foi reutilizado
        assert database.connection.execute("SELECT COUNT(*) FROM versao_processada").fetchone()[0] == 6
        assert database.connection.execute("SELECT COUNT(*) FROM checkpoint").fetchone()[0] == 0
        execution = database.connection.execute(
            "SELECT * FROM execucao_auditoria ORDER BY id DESC"
        ).fetchone()
        assert execution["coverage_pairs_total"] == 4
        assert execution["coverage_pairs_existing"] == 1
        assert execution["coverage_pairs_processed_now"] == 3
        assert execution["coverage_pairs_missing"] == 0
        assert execution["coverage_percent"] == 100


def test_coverage_detects_multiple_holes_and_invalid_record(tmp_path):
    history = _history(tmp_path)
    with Database(tmp_path / "audit.db") as database:
        database.initialize()
        service = AuditService(database, _source(history))
        service.audit(SHEET, start_version_id="1", end_version_id="7")
        database.connection.execute(
            "UPDATE versao_processada SET hash_origem=NULL WHERE versao_anterior_id='3'"
        )
        doomed = database.connection.execute(
            "SELECT id FROM versao_processada WHERE versao_anterior_id='5'"
        ).fetchone()[0]
        database.connection.execute("DELETE FROM alteracao WHERE versao_processada_id=?", (doomed,))
        database.connection.execute("DELETE FROM versao_processada WHERE id=?", (doomed,))
        database.connection.commit()
        spreadsheet_id = database.connection.execute("SELECT id FROM planilha").fetchone()[0]

        coverage = get_audit_coverage(
            database.connection, spreadsheet_id,
            [item[0] for item in history], "1", "7",
        )
        assert [(a.id, b.id) for a, b in coverage.missing_pairs] == [("3", "4"), ("5", "6")]
        assert [(block.start.id, block.end.id) for block in coverage.missing_blocks] == [
            ("3", "4"), ("5", "6")
        ]


def test_interval_validation_and_open_bounds(tmp_path):
    versions = [item[0] for item in _history(tmp_path, 4)]
    assert [version.id for version in resolve_interval(versions, "2", None)] == ["2", "3", "4"]
    assert [version.id for version in resolve_interval(versions, None, "3")] == ["1", "2", "3"]
    for start, end in (("missing", "3"), ("4", "2"), ("2", "2")):
        try:
            resolve_interval(versions, start, end)
        except ValueError:
            pass
        else:
            raise AssertionError("intervalo inválido aceito")
