import base64
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook
import pytest

from app.audit_service import AuditService
from app.config import Settings
from app.database import Database
from app.models import AuditExecutionStatus
from app.sources import BrowserSharePointSource, SharePointReadError


SITE = "https://tenant.sharepoint.com/sites/qualidade"
ROOT = "/sites/qualidade/Documentos Compartilhados"


def workbook_bytes(value: str) -> bytes:
    output = BytesIO()
    workbook = Workbook()
    workbook.active["A1"] = value
    workbook.save(output)
    workbook.close()
    return output.getvalue()


class FakeBrowser:
    def __init__(self, responses: dict[str, object]) -> None:
        self.responses = responses
        self.calls: list[tuple[str, bool]] = []
        self.visited: list[str] = []
        self.quit_called = False

    def get(self, url: str) -> None:
        self.visited.append(url)

    def execute_async_script(self, script: str, url: str, binary: bool) -> object:
        assert "method: 'GET'" in script
        assert "document.cookie" not in script and "localStorage" not in script
        self.calls.append((url, binary))
        value = next(value for suffix, value in self.responses.items() if suffix in url)
        if isinstance(value, bytes):
            return {"base64": base64.b64encode(value).decode()}
        if isinstance(value, Exception):
            return {"error": str(value)}
        return {"json": value}

    def quit(self) -> None:
        self.quit_called = True


def discovery_responses() -> dict[str, object]:
    return {
        "Compartilhados')/Files": {"value": []},
        "Compartilhados')/Folders": {"value": [
            {"Name": "Setor A", "ServerRelativeUrl": f"{ROOT}/Setor A"},
            {"Name": "Setor B", "ServerRelativeUrl": f"{ROOT}/Setor B"},
        ]},
        "Setor%20A')/Files": {"value": [
            {"Name": "CQL028.xlsx", "ServerRelativeUrl": f"{ROOT}/Setor A/CQL028.xlsx", "UniqueId": "uuid-a", "Length": "10"}
        ]},
        "Setor%20A')/Folders": {"value": []},
        "Setor%20B')/Files": {"value": [
            {"Name": "CQL028.xlsx", "ServerRelativeUrl": f"{ROOT}/Setor B/CQL028.xlsx", "UniqueId": "uuid-b", "Length": "20"},
            {"Name": "ignorar.xls", "ServerRelativeUrl": f"{ROOT}/Setor B/ignorar.xls", "UniqueId": "uuid-x", "Length": "1"},
        ]},
        "Setor%20B')/Folders": {"value": []},
    }


def test_configuration_has_no_credentials_or_tokens(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("SHAREPOINT_SITE_URL", SITE)
    monkeypatch.setenv("SHAREPOINT_SCOPE_PATHS", f"{ROOT};{ROOT}/Outro")
    settings = Settings.from_environment()
    assert settings.require_browser_sharepoint() == (SITE, (ROOT, f"{ROOT}/Outro"))
    assert not any("password" in name or "token" in name or "cookie" in name for name in settings.__dict__)


def test_recursively_discovers_same_name_in_different_folders(tmp_path: Path) -> None:
    browser = FakeBrowser(discovery_responses())
    with BrowserSharePointSource(SITE, [ROOT], browser, temp_directory=tmp_path) as source:
        items = source.list_spreadsheets()
    assert [(item.drive_item_id, item.name, item.folder) for item in items] == [
        ("uuid-a", "CQL028.xlsx", f"{ROOT}/Setor A"),
        ("uuid-b", "CQL028.xlsx", f"{ROOT}/Setor B"),
    ]
    assert all(method is False for _, method in browser.calls)


def test_versions_keep_id_label_metadata_and_technical_order(tmp_path: Path) -> None:
    responses = discovery_responses()
    responses["/Versions?"] = {"value": [
        {"ID": 98, "VersionLabel": "0.98", "Created": "2026-09-15T13:00:00Z", "CreatedBy": {"Title": "Autora", "Email": "a@example.com", "LoginName": "i:0#.f|membership|a@example.com"}, "CheckInComment": "fim", "Size": "200", "Url": "hist-98", "IsCurrentVersion": False},
        {"ID": 2, "VersionLabel": "0.10", "Created": "2026-09-15T12:00:00Z", "CreatedBy": {"Title": "Autor"}, "Size": 100},
    ]}
    browser = FakeBrowser(responses)
    with BrowserSharePointSource(SITE, [ROOT], browser, temp_directory=tmp_path) as source:
        spreadsheet = source.list_spreadsheets()[0]
        versions = source.list_versions(spreadsheet)
    assert [(version.id, version.number) for version in versions] == [("2", "0.10"), ("98", "0.98")]
    assert versions[1].author == "Autora"
    assert versions[1].author_email == "a@example.com"
    assert versions[1].comment == "fim" and versions[1].source_url == "hist-98"
    assert versions[1].is_current is False


def test_path_encoding_download_and_xlsx_validation(tmp_path: Path) -> None:
    responses = discovery_responses()
    responses["/Versions?"] = {"value": [{"ID": 7, "VersionLabel": "0.7"}]}
    responses["Versions(7)/$value"] = workbook_bytes("válido")
    browser = FakeBrowser(responses)
    with BrowserSharePointSource(SITE, [ROOT], browser, temp_directory=tmp_path) as source:
        spreadsheet = source.list_spreadsheets()[0]
        version = source.list_versions(spreadsheet)[0]
        path = source.get_version(spreadsheet, version)
        assert path.is_file() and "_7_0.7.xlsx" in path.name
        download_url = browser.calls[-1][0]
        assert "Setor%20A/CQL028.xlsx" in download_url and "Versions(7)/$value" in download_url
    assert not path.parent.exists()


def test_invalid_or_missing_version_never_becomes_a_non_adjacent_comparison(tmp_path: Path) -> None:
    responses = discovery_responses()
    responses["/Versions?"] = {"value": [
        {"ID": 1, "VersionLabel": "0.1"}, {"ID": 2, "VersionLabel": "0.2"}, {"ID": 3, "VersionLabel": "0.3"}
    ]}
    responses["Versions(1)/$value"] = workbook_bytes("um")
    responses["Versions(2)/$value"] = "não é xlsx".encode()
    responses["Versions(3)/$value"] = workbook_bytes("três")
    browser = FakeBrowser(responses)
    with Database(tmp_path / "audit.db") as database:
        database.initialize()
        with BrowserSharePointSource(SITE, [ROOT], browser, temp_directory=tmp_path / "temp") as source:
            spreadsheet = source.list_spreadsheets()[0]
            result = AuditService(database, source).audit(spreadsheet)
        assert result.status is AuditExecutionStatus.FAILED
        assert result.processed_versions == 0
        assert database.connection.execute("SELECT COUNT(*) FROM versao_processada").fetchone()[0] == 0
        assert database.connection.execute("SELECT COUNT(*) FROM checkpoint").fetchone()[0] == 0


def test_rejects_write_or_cross_origin_endpoints(tmp_path: Path) -> None:
    source = BrowserSharePointSource(SITE, [ROOT], FakeBrowser({}), temp_directory=tmp_path)
    with pytest.raises(SharePointReadError, match="Somente endpoints GET"):
        source._fetch("https://evil.example/_api/web")
    source.close()
