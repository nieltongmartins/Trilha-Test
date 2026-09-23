from pathlib import Path
import json
import shutil
import zipfile

from openpyxl import Workbook

from tools.diagnose_row_hashes import diagnose


def _rewrite_member(path: Path, member: str, replacements: list[tuple[bytes, bytes]]) -> None:
    rewritten = path.with_suffix(".rewritten.xlsx")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(rewritten, "w") as target:
        for item in source.infolist():
            payload = source.read(item.filename)
            if item.filename == member:
                for before, after in replacements:
                    payload = payload.replace(before, after, 1)
            target.writestr(item, payload)
    rewritten.replace(path)


def test_diagnostic_classifies_all_sample_groups_without_cell_values(tmp_path: Path) -> None:
    before = tmp_path / "1.0.xlsx"
    after = tmp_path / "1.1.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "confidential unchanged"
    sheet["A2"] = "confidential lexical"
    sheet["A3"] = 1234
    workbook.save(before)
    workbook.close()
    shutil.copyfile(before, after)
    _rewrite_member(
        after,
        "xl/worksheets/sheet1.xml",
        [
            (b'<row r="2">', b'<row customFormat="0" r="2">'),
            (b"<v>1234</v>", b"<v>987654321</v>"),
        ],
    )

    report = diagnose([before, after], sample_limit=5)

    aggregate = report["aggregate"]
    assert aggregate["rows_total"] == 3
    assert aggregate["sha_equal"] == 1
    assert aggregate["sha_different"] == 2
    assert aggregate["different_logically_equal"] == 1
    assert aggregate["different_logically_changed"] == 1
    assert aggregate["raw_hash_false_negative_percent"] == 50.0
    assert {name: len(report["pairs"][0]["samples"][name])
            for name in ("A", "B", "C")} == {"A": 1, "B": 1, "C": 1}
    assert "customFormat" in report["pairs"][0]["samples"]["A"][0][
        "changed_attribute_names"
    ]
    serialized = json.dumps(report)
    assert "confidential" not in serialized
    assert "987654321" not in serialized


def test_diagnostic_reports_global_dependency_changes(tmp_path: Path) -> None:
    before = tmp_path / "1.xlsx"
    after = tmp_path / "2.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = 1
    workbook.save(before)
    workbook.close()
    shutil.copyfile(before, after)
    _rewrite_member(
        after,
        "xl/workbook.xml",
        [(b"<workbookPr />", b'<workbookPr date1904="1" />')],
    )

    pair = diagnose([before, after])["pairs"][0]

    assert pair["global_dependencies"]["epoch"]["equal"] is False
    assert pair["global_dependencies"]["workbook_metadata"]["equal"] is False
    assert pair["global_dependencies"]["styles"]["equal"] is True
