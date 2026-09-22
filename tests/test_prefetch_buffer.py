"""Contratos estruturais do buffer de prefetch executado no único Edge."""

from pathlib import Path

from app.sources.base import SpreadsheetInfo, VersionInfo
from app.sources.sharepoint import (
    BrowserSharePointSource,
    PREFETCH_BUFFER_SIZE,
    _CLEAR_PREFETCH_SCRIPT,
    _READ_PREFETCH_CHUNK_SCRIPT,
    _START_PREFETCH_SCRIPT,
    _WAIT_PREFETCH_SCRIPT,
)


SITE = "https://tenant.sharepoint.com/sites/auditoria"
ROOT = "/sites/auditoria/Documentos"


class SlotBrowser:
    """Simula apenas o despacho não bloqueante; fetches permanecem no JS real."""

    def __init__(self) -> None:
        self.started: list[tuple[object, ...]] = []
        self.cleared: list[str] = []

    def execute_async_script(self, script: str, *args: object) -> object:
        if script == _START_PREFETCH_SCRIPT:
            self.started.append(args)
            return {"ok": True, "state": "started"}
        if script == _CLEAR_PREFETCH_SCRIPT:
            self.cleared.append(str(args[0]))
            return {"ok": True}
        raise AssertionError("script inesperado")


def make_source(tmp_path: Path) -> tuple[BrowserSharePointSource, SlotBrowser, SpreadsheetInfo]:
    browser = SlotBrowser()
    source = BrowserSharePointSource(SITE, [ROOT], browser, temp_directory=tmp_path)
    spreadsheet = SpreadsheetInfo(
        SITE, "sharepoint-rest", "arquivo-1", "Auditoria.xlsx", f"{ROOT}/Auditoria.xlsx"
    )
    return source, browser, spreadsheet


def test_buffer_starts_empty_and_has_named_limit(tmp_path: Path) -> None:
    source, _, _ = make_source(tmp_path)
    assert PREFETCH_BUFFER_SIZE == 2
    assert source._prefetch_slots == {}


def test_buffer_accepts_one_then_two_independent_versions(tmp_path: Path) -> None:
    source, browser, spreadsheet = make_source(tmp_path)
    versions = [VersionInfo(str(number), f"1.{number}", size=100 + number) for number in (1, 2)]

    assert source.prefetch_version(spreadsheet, versions[0]) is True
    assert len(source._prefetch_slots) == 1
    assert source.prefetch_version(spreadsheet, versions[1]) is True
    assert len(source._prefetch_slots) == PREFETCH_BUFFER_SIZE

    tokens = list(source._prefetch_slots)
    urls = [slot["url"] for slot in source._prefetch_slots.values()]
    assert tokens[0] != tokens[1]
    assert urls[0] != urls[1]
    assert browser.started[0][2] != browser.started[1][2]


def test_third_version_waits_until_a_slot_is_released(tmp_path: Path) -> None:
    source, _, spreadsheet = make_source(tmp_path)
    versions = [VersionInfo(str(number), f"2.{number}", size=100) for number in (1, 2, 3)]
    assert all(source.prefetch_version(spreadsheet, version) for version in versions[:2])
    assert source.prefetch_version(spreadsheet, versions[2]) is False

    first_token = next(iter(source._prefetch_slots))
    source.cancel_prefetch(first_token, reason="consumido_no_teste")
    assert source.prefetch_version(spreadsheet, versions[2]) is True
    assert len(source._prefetch_slots) == PREFETCH_BUFFER_SIZE


def test_retry_replaces_only_failed_slot_token(tmp_path: Path) -> None:
    source, _, spreadsheet = make_source(tmp_path)
    first = VersionInfo("1", "3.1", size=101)
    second = VersionInfo("2", "3.2", size=102)
    source.prefetch_version(spreadsheet, first)
    source.prefetch_version(spreadsheet, second)
    old_tokens = list(source._prefetch_slots)

    source.cancel_prefetch(old_tokens[0], reason="retry")
    source.prefetch_version(spreadsheet, first)

    new_tokens = list(source._prefetch_slots)
    assert old_tokens[1] in new_tokens
    assert old_tokens[0] not in new_tokens


def test_refresh_cleanup_invalidates_every_slot_and_blob_reference(tmp_path: Path) -> None:
    source, browser, spreadsheet = make_source(tmp_path)
    source.prefetch_version(spreadsheet, VersionInfo("1", "4.1", size=101))
    source.prefetch_version(spreadsheet, VersionInfo("2", "4.2", size=102))

    source.cancel_prefetch(reason="refresh")

    assert source._prefetch_slots == {}
    assert browser.cleared[-1] == ""


def test_javascript_slots_validate_full_identity_and_release_references() -> None:
    assert "window.__auditPrefetchSlots" in _START_PREFETCH_SCRIPT
    assert "new AbortController()" in _START_PREFETCH_SCRIPT
    assert "startedAt" in _START_PREFETCH_SCRIPT
    for field in ("token", "version", "versionId", "url", "expectedSize"):
        assert field in _WAIT_PREFETCH_SCRIPT
        assert field in _READ_PREFETCH_CHUNK_SCRIPT
    assert "slot.state !== 'ready'" in _READ_PREFETCH_CHUNK_SCRIPT
    assert "slot.controller = null" in _CLEAR_PREFETCH_SCRIPT
    assert "delete slots[key]" in _CLEAR_PREFETCH_SCRIPT


def test_each_javascript_slot_owns_timeout_controller_blob_and_error() -> None:
    assert "const controller = new AbortController()" in _START_PREFETCH_SCRIPT
    assert "setTimeout(() => { slot.timedOut = true; controller.abort(); }" in _START_PREFETCH_SCRIPT
    assert "blob: null" in _START_PREFETCH_SCRIPT
    assert "error: null" in _START_PREFETCH_SCRIPT
    assert "slots[token] = slot" in _START_PREFETCH_SCRIPT
