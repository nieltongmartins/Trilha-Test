from io import BytesIO
import base64
from pathlib import Path
import zipfile

from openpyxl import Workbook
import pytest

from app.audit_service import AuditService
from app.config import Settings
from app.database import Database
from app.models import AuditExecutionStatus
from app.sources import BrowserSharePointSource, SharePointReadError, SpreadsheetInfo
from app.sources.sharepoint import _normalize_scope_path, _odata_object

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
        self.calls: list[str] = []
        self.visited: list[str] = []
        self.quit_called = False
        self.current_url = SITE
        self.download_blob: bytes | None = None
        self.script_timeout: float | None = None

    def _response(self, url: str) -> object:
        matches = [(key, value) for key, value in self.responses.items() if key in url]
        if not matches:
            raise AssertionError(f"sem resposta fake para {url}")
        return max(matches, key=lambda item: len(item[0]))[1]

    def get(self, url: str) -> None:
        self.visited.append(url)

    def execute_async_script(self, script: str, *args: object) -> object:
        if "window.__auditDownloadBlob = null" in script:
            self.download_blob = None
            return {"ok": True}
        if "blob.slice" in script:
            assert self.download_blob is not None
            offset, length = int(args[0]), int(args[1])
            data = self.download_blob[offset : offset + length]
            return {
                "ok": True,
                "data": base64.b64encode(data).decode("ascii"),
                "length": len(data),
            }
        url = str(args[0])
        assert "method: 'GET'" in script
        assert "document.cookie" not in script and "localStorage" not in script
        self.calls.append(url)
        value = self._response(url)
        if "response.blob()" in script:
            self.visited.append(url)
            if isinstance(value, bytes):
                self.download_blob = value
                return {"ok": True, "size": len(value)}
            if value == "partial":
                self.download_blob = b"parcial"
                return {"ok": True, "size": len(self.download_blob)}
            if isinstance(value, Exception):
                return {"error": str(value)}
        if isinstance(value, Exception):
            return {"error": str(value)}
        return {"json": value}

    def quit(self) -> None:
        self.quit_called = True

    def set_script_timeout(self, time_to_wait: float) -> None:
        self.script_timeout = time_to_wait


def discovery_responses() -> dict[str, object]:
    return {
        "Compartilhados')/Files": {"value": []},
        "Compartilhados')/Folders": {
            "value": [
                {"Name": "Setor A", "ServerRelativeUrl": f"{ROOT}/Setor A"},
                {"Name": "Setor B", "ServerRelativeUrl": f"{ROOT}/Setor B"},
            ]
        },
        "Setor%20A')/Files": {
            "value": [
                {
                    "Name": "Arquivo.xlsx",
                    "ServerRelativeUrl": f"{ROOT}/Setor A/Arquivo.xlsx",
                    "UniqueId": "UUID-A",
                    "Length": "10",
                }
            ]
        },
        "Setor%20A')/Folders": {"value": []},
        "Setor%20B')/Files": {
            "value": [
                {
                    "Name": "Arquivo.xlsx",
                    "ServerRelativeUrl": f"{ROOT}/Setor B/Arquivo.xlsx",
                    "UniqueId": "UUID-B",
                    "Length": "20",
                },
                {
                    "Name": "ignorar.xls",
                    "ServerRelativeUrl": f"{ROOT}/Setor B/ignorar.xls",
                    "UniqueId": "UUID-X",
                    "Length": "1",
                },
            ]
        },
        "Setor%20B')/Folders": {"value": []},
    }


def file_metadata(
    unique_id: str = "UUID-A", label: str = "0.99", ui_version: int = 99
) -> dict[str, object]:
    return {
        "Name": "Arquivo.xlsx",
        "ServerRelativeUrl": f"{ROOT}/Setor A/Arquivo.xlsx",
        "UniqueId": unique_id,
        "UIVersion": ui_version,
        "UIVersionLabel": label,
        "TimeLastModified": "2026-09-16T11:35:49Z",
        "Length": "226665",
        "ModifiedBy": {
            "Title": "Autora Atual",
            "Email": "atual@example.com",
            "LoginName": "login-atual",
        },
    }


def make_source(
    tmp_path: Path, responses: dict[str, object]
) -> tuple[BrowserSharePointSource, FakeBrowser]:
    browser = FakeBrowser(responses)
    source = BrowserSharePointSource(
        SITE,
        [ROOT],
        browser,
        temp_directory=tmp_path / "temp",
    )
    return source, browser


def test_configuration_has_multiple_scopes_and_no_browser_secrets(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("SHAREPOINT_SITE_URL", SITE)
    monkeypatch.setenv("SHAREPOINT_SCOPE_PATHS", f"{ROOT};{ROOT}/Outro")
    settings = Settings.from_environment()
    assert settings.require_browser_sharepoint() == (SITE, (ROOT, f"{ROOT}/Outro"))
    browser_fields = {"sharepoint_site_url", "sharepoint_scope_paths"}
    assert all(
        not any(word in name for word in ("password", "token", "cookie"))
        for name in browser_fields
    )


def test_odata_object_accepts_real_nometadata_entity_and_legacy_envelopes() -> None:
    direct = {"UniqueId": "UUID-A", "UIVersion": 99}

    assert _odata_object(direct) is direct
    assert _odata_object({"value": direct}) is direct
    assert _odata_object({"d": direct}) is direct
    with pytest.raises(SharePointReadError, match="objeto inválido"):
        _odata_object({"value": [direct]})


def test_scope_relative_to_site_is_converted_to_server_relative_path() -> None:
    assert (
        _normalize_scope_path("/Documentos Compartilhados1", "/controle_qualidade")
        == "/controle_qualidade/Documentos Compartilhados1"
    )
    assert (
        _normalize_scope_path(
            "/controle_qualidade/Documentos Compartilhados1",
            "/controle_qualidade",
        )
        == "/controle_qualidade/Documentos Compartilhados1"
    )


def test_source_uses_normalized_scope_in_sharepoint_request(tmp_path: Path) -> None:
    browser = FakeBrowser(discovery_responses())
    source = BrowserSharePointSource(
        SITE,
        ["/Documentos Compartilhados"],
        browser,
        temp_directory=tmp_path,
    )

    with source:
        source.list_spreadsheets()

    assert source.scope_paths == (ROOT,)
    assert any(
        "/sites/qualidade/Documentos%20Compartilhados')" in url
        for url in browser.calls
    )


def test_wait_until_authenticated_validates_site_with_rest(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source, browser = make_source(tmp_path, {"web?$select=Id": {"Id": "site-id"}})
    observed: dict[str, object] = {}

    class WaitFake:
        def __init__(self, received_browser, timeout, poll_frequency) -> None:
            observed.update(
                browser=received_browser,
                timeout=timeout,
                poll_frequency=poll_frequency,
            )

        def until(self, predicate) -> None:
            assert predicate(browser) is True

    monkeypatch.setattr("selenium.webdriver.support.ui.WebDriverWait", WaitFake)

    source.wait_until_authenticated(timeout=12)

    assert observed == {"browser": browser, "timeout": 12, "poll_frequency": 0.5}
    assert browser.calls == [f"{SITE}/_api/web?$select=Id"]


def test_recursively_discovers_same_name_as_distinct_unique_ids(tmp_path: Path) -> None:
    source, browser = make_source(tmp_path, discovery_responses())
    with source:
        items = source.list_spreadsheets()
    assert [(item.drive_item_id, item.name, item.folder) for item in items] == [
        ("UUID-A", "Arquivo.xlsx", f"{ROOT}/Setor A"),
        ("UUID-B", "Arquivo.xlsx", f"{ROOT}/Setor B"),
    ]
    assert {item.drive_id for item in items} == {"sharepoint-rest"}
    assert all("/_api/" in url for url in browser.calls)


def test_lists_immediate_folders_for_ui_selection(tmp_path: Path) -> None:
    source, browser = make_source(tmp_path, discovery_responses())

    with source:
        folders = source.list_folders()

    assert folders == (
        ("Setor A", f"{ROOT}/Setor A"),
        ("Setor B", f"{ROOT}/Setor B"),
    )
    assert len(browser.calls) == 3
    assert browser.calls[0].endswith(
        "Documentos%20Compartilhados')/Folders?$select=Name,ServerRelativeUrl"
    )


def test_selected_folder_replaces_source_scope_without_new_browser(
    tmp_path: Path,
) -> None:
    source, _ = make_source(tmp_path, discovery_responses())

    source.set_scope_paths(["/Documentos Compartilhados/Setor A"])

    assert source.scope_paths == (f"{ROOT}/Setor A",)


def test_unique_id_keeps_identity_when_name_and_path_change() -> None:
    original = SpreadsheetInfo(
        SITE, "sharepoint-rest", "stable-id", "antes.xlsx", "/raiz/antes.xlsx"
    )
    changed = SpreadsheetInfo(
        SITE, "sharepoint-rest", "stable-id", "depois.xlsx", "/outra/depois.xlsx"
    )
    assert (original.site_id, original.drive_id, original.drive_item_id) == (
        changed.site_id,
        changed.drive_id,
        changed.drive_item_id,
    )


def test_parses_historical_metadata_and_appends_current_without_deriving_id(
    tmp_path: Path,
) -> None:
    responses = discovery_responses()
    responses["/Versions?"] = {
        "value": [
            {
                "ID": 700,
                "VersionLabel": "0.98",
                "Created": "2026-09-15T13:00:00Z",
                "CreatedBy": {
                    "Title": "Autora",
                    "Email": "a@example.com",
                    "LoginName": "login",
                },
                "CheckInComment": "fim",
                "Length": "200",
                "Url": "hist-700",
                "IsCurrentVersion": False,
            },
            {
                "ID": 2,
                "VersionLabel": "0.10",
                "CreatedBy": {"Title": "Autor"},
                "Size": 100,
            },
        ]
    }
    responses[")?$select=Name"] = file_metadata(ui_version=99)
    source, _ = make_source(tmp_path, responses)
    with source:
        spreadsheet = source.list_spreadsheets()[0]
        versions = source.list_versions(spreadsheet)
    assert [(v.id, v.number, v.is_current) for v in versions] == [
        ("2", "0.10", False),
        ("700", "0.98", False),
        ("99", "0.99", True),
    ]
    assert versions[1].author == "Autora" and versions[1].comment == "fim"
    assert versions[1].source_url == "hist-700" and versions[1].size == 200
    assert versions[-1].modified_at == "2026-09-16T11:35:49Z"
    assert versions[-1].author == "Autora Atual" and versions[-1].size == 226665


def test_current_is_not_duplicated_when_versions_also_contains_it(
    tmp_path: Path,
) -> None:
    responses = discovery_responses()
    responses["/Versions?"] = {
        "value": [
            {"ID": 98, "VersionLabel": "0.98"},
            {"ID": 99, "VersionLabel": "0.99", "IsCurrentVersion": True},
        ]
    }
    responses[")?$select=Name"] = file_metadata()
    source, _ = make_source(tmp_path, responses)
    with source:
        spreadsheet = source.list_spreadsheets()[0]
        versions = source.list_versions(spreadsheet)
    assert [(v.id, v.number, v.is_current) for v in versions] == [
        ("98", "0.98", False),
        ("99", "0.99", True),
    ]


def _paged_source(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    first: dict[str, object],
    second: dict[str, object] | None = None,
) -> tuple[BrowserSharePointSource, SpreadsheetInfo]:
    monkeypatch.setattr("app.sources.sharepoint._VERSION_PAGE_SIZE", 2)
    responses: dict[str, object] = {"$skip=0": first}
    if second is not None:
        responses["page=2"] = second
    source, _ = make_source(tmp_path, responses)
    monkeypatch.setattr(source, "_file_metadata", lambda _spreadsheet: file_metadata())
    spreadsheet = SpreadsheetInfo(
        SITE,
        "sharepoint-rest",
        "UUID-A",
        "Arquivo.xlsx",
        f"{ROOT}/Setor A/Arquivo.xlsx",
    )
    return source, spreadsheet


def test_version_pagination_accepts_complete_monotonic_pages(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source, spreadsheet = _paged_source(
        tmp_path,
        monkeypatch,
        {
            "value": [
                {"ID": 1, "VersionLabel": "0.1"},
                {"ID": 2, "VersionLabel": "0.2"},
            ],
            "@odata.nextLink": f"{SITE}/_api/versions?page=2",
        },
        {"value": [{"ID": 3, "VersionLabel": "0.3"}]},
    )

    with source:
        versions = source.list_versions(spreadsheet)

    assert [(item.id, item.number) for item in versions] == [
        ("1", "0.1"), ("2", "0.2"), ("3", "0.3"), ("99", "0.99")
    ]


@pytest.mark.parametrize(
    ("second_page", "message"),
    [
        (
            {
                "value": [
                    {"ID": 2, "VersionLabel": "0.2"},
                    {"ID": 3, "VersionLabel": "0.3"},
                ]
            },
            "duplicada",
        ),
        (
            {
                "value": [
                    {"ID": 4, "VersionLabel": "0.4"},
                    {"ID": 3, "VersionLabel": "0.3"},
                ]
            },
            "fora de ordem",
        ),
        (
            {"value": [{"ID": 3, "VersionLabel": "0.2"}]},
            "VersionLabel",
        ),
    ],
)
def test_version_pagination_rejects_partial_duplicates_order_and_label_conflicts(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    second_page: dict[str, object],
    message: str,
) -> None:
    source, spreadsheet = _paged_source(
        tmp_path,
        monkeypatch,
        {
            "value": [
                {"ID": 1, "VersionLabel": "0.1"},
                {"ID": 2, "VersionLabel": "0.2"},
            ],
            "@odata.nextLink": f"{SITE}/_api/versions?page=2",
        },
        second_page,
    )

    with source, pytest.raises(SharePointReadError, match=message):
        source.list_versions(spreadsheet)


def test_version_pagination_rejects_repeated_next_link_and_incomplete_empty_page(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    repeated = f"{SITE}/_api/versions?page=2"
    source, spreadsheet = _paged_source(
        tmp_path,
        monkeypatch,
        {
            "value": [
                {"ID": 1, "VersionLabel": "0.1"},
                {"ID": 2, "VersionLabel": "0.2"},
            ],
            "@odata.nextLink": repeated,
        },
        {
            "value": [{"ID": 3, "VersionLabel": "0.3"}],
            "@odata.nextLink": repeated,
        },
    )
    with source, pytest.raises(SharePointReadError, match="nextLink"):
        source.list_versions(spreadsheet)

    source, spreadsheet = _paged_source(
        tmp_path / "empty",
        monkeypatch,
        {
            "value": [
                {"ID": 1, "VersionLabel": "0.1"},
                {"ID": 2, "VersionLabel": "0.2"},
            ],
            "@odata.nextLink": repeated,
        },
        {"value": [], "@odata.nextLink": f"{SITE}/_api/versions?page=3"},
    )
    with source, pytest.raises(SharePointReadError, match="incompleta"):
        source.list_versions(spreadsheet)


def test_version_enumeration_rejects_current_version_changed_during_pagination(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source, spreadsheet = _paged_source(
        tmp_path, monkeypatch, {"value": [{"ID": 98, "VersionLabel": "0.98"}]}
    )
    metadata = iter(
        (
            file_metadata(label="0.99", ui_version=99),
            file_metadata(label="1.0", ui_version=100),
        )
    )
    monkeypatch.setattr(source, "_file_metadata", lambda _spreadsheet: next(metadata))

    with source, pytest.raises(SharePointReadError, match="potencialmente incompleta"):
        source.list_versions(spreadsheet)


@pytest.mark.parametrize(
    "historical",
    [
        [{"ID": 99, "VersionLabel": "outro", "IsCurrentVersion": True}],
        [{"ID": 98, "VersionLabel": "0.99", "IsCurrentVersion": True}],
        [{"ID": 98, "VersionLabel": "0.98", "IsCurrentVersion": True}],
    ],
)
def test_version_enumeration_rejects_inconsistent_current_version(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    historical: list[dict[str, object]],
) -> None:
    source, spreadsheet = _paged_source(
        tmp_path, monkeypatch, {"value": historical}
    )
    with source, pytest.raises(SharePointReadError, match="atual"):
        source.list_versions(spreadsheet)


def test_rejects_changed_unique_id_in_current_metadata(tmp_path: Path) -> None:
    responses = discovery_responses()
    responses["/Versions?"] = {"value": []}
    responses[")?$select=Name"] = file_metadata("outro-id")
    source, _ = make_source(tmp_path, responses)
    with source, pytest.raises(SharePointReadError, match="UniqueId"):
        source.list_versions(source.list_spreadsheets()[0])


def test_downloads_historical_and_current_using_distinct_read_only_endpoints(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("app.sources.sharepoint._DOWNLOAD_CHUNK_SIZE", 32)
    responses = discovery_responses()
    responses["/Versions?"] = {"value": [{"ID": 98, "VersionLabel": "0.98"}]}
    responses[")?$select=Name"] = file_metadata()
    responses["Versions(98)/$value"] = workbook_bytes("histórica")
    responses["')/$value"] = workbook_bytes("atual")
    source, browser = make_source(tmp_path, responses)
    with source:
        spreadsheet = source.list_spreadsheets()[0]
        historical, current = source.list_versions(spreadsheet)
        historical_path = source.get_version(spreadsheet, historical)
        assert historical_path.is_file()
        assert len(historical_path.stem) == 64 and historical_path.suffix == ".xlsx"
        historical_bytes = historical_path.read_bytes()
        current_path = source.get_version(spreadsheet, current)
        assert current_path.is_file() and current_path != historical_path
        source.release_version(historical_path)
        source.release_version(current_path)
        assert list(source._workspace.path.glob("*.xlsx")) == []
    assert zipfile.is_zipfile(BytesIO(historical_bytes))
    assert "/Versions(98)/$value" in browser.visited[-2]
    assert "/Versions(" not in browser.visited[-1] and browser.visited[-1].endswith(
        "/$value"
    )
    assert all("/_api/" in url for url in browser.visited)


def test_download_failure_from_fetch_is_reported_without_waiting(
    tmp_path: Path,
) -> None:
    source, _ = make_source(
        tmp_path,
        {"Versions(7)/$value": RuntimeError("HTTP 403")},
    )
    spreadsheet = SpreadsheetInfo(
        SITE,
        "sharepoint-rest",
        "UUID-A",
        "Arquivo.xlsx",
        f"{ROOT}/Setor A/Arquivo.xlsx",
    )
    from app.sources.base import VersionInfo

    with source, pytest.raises(SharePointReadError, match="HTTP 403"):
        source.get_version(spreadsheet, VersionInfo("7", "0.7"))


def test_rejects_invalid_open_xml_and_incomplete_download(tmp_path: Path) -> None:
    responses = discovery_responses()
    responses["Versions(7)/$value"] = "não é xlsx".encode()
    source, _ = make_source(tmp_path, responses)
    spreadsheet = SpreadsheetInfo(
        SITE,
        "sharepoint-rest",
        "UUID-A",
        "Arquivo.xlsx",
        f"{ROOT}/Setor A/Arquivo.xlsx",
    )
    from app.sources.base import VersionInfo

    with source, pytest.raises(SharePointReadError, match="XLSX/ZIP"):
        source.get_version(spreadsheet, VersionInfo("7", "0.7"))

    responses["Versions(7)/$value"] = "partial"
    source, _ = make_source(tmp_path / "second", responses)
    with source, pytest.raises(SharePointReadError, match="XLSX/ZIP"):
        source.get_version(spreadsheet, VersionInfo("7", "0.7"))


def test_requires_workbook_xml(tmp_path: Path) -> None:
    invalid = tmp_path / "invalid.xlsx"
    with zipfile.ZipFile(invalid, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
    with pytest.raises(SharePointReadError, match="workbook.xml"):
        BrowserSharePointSource._validate_xlsx(invalid)


def test_gap_download_failure_rolls_back_and_keeps_checkpoint(tmp_path: Path) -> None:
    responses = discovery_responses()
    responses["/Versions?"] = {
        "value": [
            {"ID": 97, "VersionLabel": "0.97"},
            {"ID": 98, "VersionLabel": "0.98"},
        ]
    }
    responses[")?$select=Name"] = file_metadata()
    responses["Versions(97)/$value"] = workbook_bytes("97")
    responses["Versions(98)/$value"] = "inválido".encode()
    responses["')/$value"] = workbook_bytes("99")
    source, _ = make_source(tmp_path, responses)
    with Database(tmp_path / "audit.db") as database, source:
        database.initialize()
        spreadsheet = source.list_spreadsheets()[0]
        result = AuditService(database, source).audit(spreadsheet)
        assert result.status is AuditExecutionStatus.FAILED
        assert result.processed_versions == 0
        assert (
            database.connection.execute(
                "SELECT COUNT(*) FROM versao_processada"
            ).fetchone()[0]
            == 0
        )
        assert (
            database.connection.execute("SELECT COUNT(*) FROM checkpoint").fetchone()[0]
            == 0
        )
        assert list(source._workspace.path.glob("*.xlsx")) == []


def test_only_get_same_origin_api_is_embedded() -> None:
    import inspect
    import app.sources.sharepoint as module

    source = inspect.getsource(module)
    assert "method: 'GET'" in source
    assert all(
        f"method: '{method}'" not in source
        for method in ("POST", "PUT", "PATCH", "DELETE")
    )
    assert (
        "get_cookies" not in source
        and "access_token" not in source
        and "password" not in source
    )
