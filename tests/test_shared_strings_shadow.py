from pathlib import Path
import shutil
import zipfile

import pytest
from openpyxl import Workbook

from app.excel.reader import (
    ConsecutiveWorkbookReader,
    _FastReaderUnsupported,
    _shared_strings_snapshot,
    read_workbook,
)

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _sst(*entries: str, extra: str = "") -> bytes:
    return (
        f'<sst xmlns="{NS}" count="{len(entries)}" uniqueCount="{len(entries)}">'
        + "".join(entries) + extra + "</sst>"
    ).encode()


def _si(text: str, attrs: str = "") -> str:
    return f"<si><t{attrs}>{text}</t></si>"


def _make_shared_workbook(path: Path, indices: tuple[int, ...], entries: tuple[str, ...]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    for row, index in enumerate(indices, 1):
        sheet.cell(row, 1, index)
    workbook.save(path)
    workbook.close()
    rewritten = path.with_suffix(".new.xlsx")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(rewritten, "w") as target:
        for item in source.infolist():
            payload = source.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                for index in indices:
                    payload = payload.replace(
                        f'<c r="A{indices.index(index) + 1}" t="n"><v>{index}</v></c>'.encode(),
                        f'<c r="A{indices.index(index) + 1}" t="s"><v>{index}</v></c>'.encode(),
                    )
            target.writestr(item, payload)
        target.writestr("xl/sharedStrings.xml", _sst(*entries))
    rewritten.replace(path)


def _replace_shared(path: Path, entries: tuple[str, ...]) -> None:
    rewritten = path.with_suffix(".new.xlsx")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(rewritten, "w") as target:
        for item in source.infolist():
            if item.filename != "xl/sharedStrings.xml":
                target.writestr(item, source.read(item.filename))
        target.writestr("xl/sharedStrings.xml", _sst(*entries))
    rewritten.replace(path)


def test_snapshot_distinguishes_append_change_remove_and_reorder() -> None:
    old = _shared_strings_snapshot(_sst(_si("a"), _si("b")))
    appended = _shared_strings_snapshot(_sst(_si("a"), _si("b"), _si("c")))
    changed = _shared_strings_snapshot(_sst(_si("a"), _si("B")))
    removed = _shared_strings_snapshot(_sst(_si("a")))
    reordered = _shared_strings_snapshot(_sst(_si("b"), _si("a")))

    assert old.signatures == appended.signatures[:2]
    assert old.signatures[1] != changed.signatures[1]
    assert len(removed.signatures) == 1
    assert old.signatures != reordered.signatures


def test_snapshot_hashes_rich_text_structure_not_only_visible_text() -> None:
    plain = _shared_strings_snapshot(_sst(_si("ab")))
    rich = _shared_strings_snapshot(_sst("<si><r><t>a</t></r><r><t>b</t></r></si>"))
    styled = _shared_strings_snapshot(
        _sst("<si><r><rPr><b/></rPr><t>a</t></r><r><t>b</t></r></si>")
    )

    assert plain.values == rich.values == styled.values == ("ab",)
    assert len({plain.signatures[0], rich.signatures[0], styled.signatures[0]}) == 3


@pytest.mark.parametrize(
    "entry",
    [
        _si(" a ", ' xml:space="preserve"'),
        _si("ação 漢字 😀"),
        "<si><r><t xml:space=\"preserve\"> a </t></r><rPh sb=\"0\" eb=\"1\"><t>á</t></rPh></si>",
    ],
)
def test_snapshot_preserves_whitespace_xml_space_unicode_and_phonetics(entry: str) -> None:
    snapshot = _shared_strings_snapshot(_sst(entry))
    assert snapshot.signatures[0] is not None


def test_unknown_shared_string_structure_never_proves_equality() -> None:
    snapshot = _shared_strings_snapshot(
        _sst(f'<si><t>a</t><ext xmlns="urn:unknown">x</ext></si>')
    )
    assert snapshot.values == ("a",)
    assert snapshot.signatures == (None,)


def test_absent_and_corrupt_shared_strings_are_conservative() -> None:
    assert _shared_strings_snapshot(None).present is False
    with pytest.raises(_FastReaderUnsupported, match="XML invalido"):
        _shared_strings_snapshot(b"<sst>")


def test_shadow_append_only_validates_rows_without_enabling_reuse(tmp_path: Path) -> None:
    previous_path = tmp_path / "previous.xlsx"
    current_path = tmp_path / "current.xlsx"
    _make_shared_workbook(previous_path, (0, 1), (_si("old-a"), _si("old-b")))
    shutil.copyfile(previous_path, current_path)
    _replace_shared(current_path, (_si("old-a"), _si("old-b"), _si("new")))

    incremental = ConsecutiveWorkbookReader()
    previous = incremental.read(previous_path)
    current = incremental.read(current_path)
    metrics = incremental.last_metrics

    assert previous == current == read_workbook(current_path)
    assert metrics.rows_reused == 0  # decisao oficial conservadora permanece ativa
    assert metrics.rows_parsed == 2
    assert metrics.sharedstrings_indices_equal == 2
    assert metrics.sharedstrings_indices_new == 1
    assert metrics.sharedstrings_shadow_rows_candidate == 2
    assert metrics.sharedstrings_shadow_rows_safe == 2
    assert metrics.sharedstrings_shadow_indices_checked == 2
    assert metrics.rows_reused_sharedstrings == 2


def test_shadow_invalidates_only_row_using_changed_index(tmp_path: Path) -> None:
    previous_path = tmp_path / "previous.xlsx"
    current_path = tmp_path / "current.xlsx"
    _make_shared_workbook(previous_path, (0, 1), (_si("stable"), _si("before")))
    shutil.copyfile(previous_path, current_path)
    _replace_shared(current_path, (_si("stable"), _si("after")))

    incremental = ConsecutiveWorkbookReader()
    incremental.read(previous_path)
    current = incremental.read(current_path)
    metrics = incremental.last_metrics

    assert current == read_workbook(current_path)
    assert metrics.sharedstrings_shadow_rows_safe == 1
    assert metrics.sharedstrings_shadow_rows_invalidated == 1
    assert metrics.rows_invalidated_changed_index == 1
    assert metrics.sharedstrings_shadow_indices_changed == 1


def test_missing_index_uses_safe_openpyxl_fallback(tmp_path: Path) -> None:
    path = tmp_path / "invalid.xlsx"
    _make_shared_workbook(path, (2,), (_si("only-zero"),))
    reader = ConsecutiveWorkbookReader()
    # openpyxl is also entitled to reject a corrupt workbook; importantly the
    # incremental proof itself raises its explicit unsupported barrier first.
    with pytest.raises((IndexError, ValueError)):
        reader.read(path)
