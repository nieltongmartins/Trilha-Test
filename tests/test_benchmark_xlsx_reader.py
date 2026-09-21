from pathlib import Path

from tools.benchmark_xlsx_reader import main, profile_file


def test_profile_variants_preserve_typed_snapshot(cql028_versions: Path) -> None:
    result = profile_file(cql028_versions / "0.85.xlsx", repeat=1)

    assert result["zip_bytes"] > 0
    assert result["official_reader_median_ns"] > 0
    for variant in result["variants"].values():
        assert variant["snapshot_exactly_equal"] is True
        assert variant["worksheet_count"] == 2
        assert variant["cells_stored"] == 7
        assert variant["snapshot_entries"] == 7
        assert variant["xml_uncompressed_bytes"] > 0
        assert variant["phases_ns"]["zip_open"] > 0


def test_profile_command_writes_json(cql028_versions: Path, tmp_path: Path) -> None:
    output = tmp_path / "profile.json"

    assert main([str(cql028_versions), "--repeat", "1", "--output", str(output)]) == 0
    payload = output.read_text(encoding="utf-8")
    assert '"schema_version": 1' in payload
    assert payload.count('"snapshot_exactly_equal": true') == 12
