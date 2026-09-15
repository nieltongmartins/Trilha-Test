import json
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook
import pytest

from app.audit_service import AuditService
from app.config import Settings
from app.database import Database
from app.models import AuditExecutionStatus
from app.sources import GraphReadError, SharePointSource, SpreadsheetInfo, VersionInfo


def response(payload: object) -> tuple[bytes, dict[str, str]]:
    return json.dumps(payload).encode(), {"Content-Type": "application/json"}


def workbook_bytes(value: str) -> bytes:
    output = BytesIO()
    workbook = Workbook()
    workbook.active["A1"] = value
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def test_settings_require_complete_sharepoint_configuration(monkeypatch: pytest.MonkeyPatch) -> None:
    names = [
        "SHAREPOINT_TENANT_ID", "SHAREPOINT_CLIENT_ID", "SHAREPOINT_CLIENT_SECRET",
        "SHAREPOINT_SITE_ID", "SHAREPOINT_DRIVE_ID",
    ]
    for name in names:
        monkeypatch.delenv(name, raising=False)
    with pytest.raises(ValueError, match="SHAREPOINT_CLIENT_SECRET"):
        Settings.from_environment().require_sharepoint()

    for index, name in enumerate(names):
        monkeypatch.setenv(name, f"value-{index}")
    assert Settings.from_environment().require_sharepoint() == tuple(
        f"value-{index}" for index in range(5)
    )


def test_lists_xlsx_with_pagination_and_never_uses_write_method(tmp_path: Path) -> None:
    calls: list[tuple[str, str]] = []

    def get(url: str, token: str) -> tuple[bytes, dict[str, str]]:
        calls.append((url, token))
        if url == "next-page":
            return response({"value": [{"id": "item-2", "name": "B.XLSX", "file": {}}]})
        return response({
            "value": [
                {"id": "item-1", "name": "A.xlsx", "file": {},
                 "parentReference": {"path": "/drives/drive/root:"}},
                {"id": "folder", "name": "Pasta", "folder": {}},
            ],
            "@odata.nextLink": "next-page",
        })

    with SharePointSource("site", "drive", lambda: "token", temp_directory=tmp_path, get_request=get) as source:
        items = source.list_spreadsheets()

    assert [item.drive_item_id for item in items] == ["item-1", "item-2"]
    assert items[0].path == "/drives/drive/root:"
    assert all(token == "token" for _, token in calls)
    assert all("/drives/drive/root/children" in url or url == "next-page" for url, _ in calls)


def test_lists_versions_oldest_first_with_metadata(tmp_path: Path) -> None:
    spreadsheet = SpreadsheetInfo("site", "drive", "item/encoded", "A.xlsx")

    def get(url: str, _: str) -> tuple[bytes, dict[str, str]]:
        assert url.endswith("/items/item%2Fencoded/versions")
        return response({"value": [
            {"id": "1.0", "lastModifiedDateTime": "2026-09-15T12:00:00Z",
             "lastModifiedBy": {"user": {"displayName": "Atual"}}, "size": 20},
            {"id": "0.9", "lastModifiedDateTime": "2026-09-15T11:00:00Z",
             "lastModifiedBy": {"user": {"displayName": "Anterior"}}, "size": 10},
        ]})

    with SharePointSource("site", "drive", lambda: "token", temp_directory=tmp_path, get_request=get) as source:
        versions = source.list_versions(spreadsheet)

    assert [item.id for item in versions] == ["0.9", "1.0"]
    assert (versions[0].author, versions[0].size) == ("Anterior", 10)


def test_downloads_historical_content_to_disposable_directory(tmp_path: Path) -> None:
    spreadsheet = SpreadsheetInfo("site", "drive", "item", "A.xlsx")
    requested: list[str] = []

    def get(url: str, _: str) -> tuple[bytes, dict[str, str]]:
        requested.append(url)
        return b"xlsx-content", {}

    source = SharePointSource("site", "drive", lambda: "token", temp_directory=tmp_path, get_request=get)
    path = source.get_version(spreadsheet, VersionInfo("0.84", "0.84"))
    assert path.read_bytes() == b"xlsx-content"
    assert requested[0].endswith("/items/item/versions/0.84/content")
    temporary_parent = path.parent
    source.close()
    assert not temporary_parent.exists()


def test_invalid_graph_json_is_normalized(tmp_path: Path) -> None:
    with SharePointSource(
        "site", "drive", lambda: "token", temp_directory=tmp_path,
        get_request=lambda _url, _token: (b"not-json", {}),
    ) as source:
        with pytest.raises(GraphReadError, match="JSON inválido"):
            source.list_spreadsheets()


def test_audit_service_operates_with_sharepoint_source(tmp_path: Path) -> None:
    spreadsheet = SpreadsheetInfo("site", "drive", "item", "A.xlsx")
    contents = {"0.84": workbook_bytes("antes"), "0.85": workbook_bytes("depois")}

    def get(url: str, _: str) -> tuple[bytes, dict[str, str]]:
        if url.endswith("/versions"):
            return response({"value": [{"id": "0.85"}, {"id": "0.84"}]})
        version_id = url.removesuffix("/content").rsplit("/", 1)[-1]
        return contents[version_id], {}

    with Database(tmp_path / "audit.db") as database:
        database.initialize()
        with SharePointSource(
            "site", "drive", lambda: "token", temp_directory=tmp_path / "temp",
            get_request=get,
        ) as source:
            result = AuditService(database, source).audit(spreadsheet)

        assert result.status is AuditExecutionStatus.COMPLETED
        assert (result.processed_versions, result.changes, result.final_version) == (1, 1, "0.85")
