from pathlib import Path

from tools.benchmark_xlsx_reader import (
    ConsecutiveSheetPrototype,
    benchmark_official_before_after,
    benchmark_sequence,
    main,
    profile_file,
)
from app.excel.reader import read_workbook


def test_profile_variants_preserve_typed_snapshot(cql028_versions: Path) -> None:
    result = profile_file(cql028_versions / "0.85.xlsx", repeat=1)

    assert result["zip_bytes"] > 0
    assert result["official_reader_median_ns"] > 0
    comparison = result["official_before_after"]
    assert comparison["snapshot_exactly_equal"] is True
    assert comparison["repetitions_after_warmup"] == 1
    assert comparison["cells"] == 7
    for name in ("previous_reader", "optimized_reader"):
        assert comparison[name]["median_ns"] > 0
        assert comparison[name]["mean_ns"] > 0
        assert comparison[name]["min_ns"] > 0
        assert comparison[name]["max_ns"] > 0
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
    assert '"schema_version": 2' in payload
    assert payload.count('"snapshot_exactly_equal": true') == 21


def test_real_reader_benchmark_uses_five_post_warmup_runs(
    cql028_versions: Path,
) -> None:
    result = benchmark_official_before_after(cql028_versions / "0.85.xlsx")

    assert result["snapshot_exactly_equal"] is True
    assert result["repetitions_after_warmup"] == 5
    assert len(result["previous_reader"]["samples_ns"]) == 5
    assert len(result["optimized_reader"]["samples_ns"]) == 5


def test_consecutive_prototype_reuses_only_proven_identical_sheets(
    cql028_versions: Path,
) -> None:
    paths = [cql028_versions / name for name in ("0.86.xlsx", "0.87.xlsx")]

    result = benchmark_sequence(paths)

    assert result["snapshot_exactly_equal"] is True
    assert result["worksheets_reused"] == 3
    assert result["versions"][1]["cells_processed"] == 0
    assert result["versions"][1]["cells_effectively_changed"] == 0


def test_prototype_matches_official_reader_across_changed_sheet_sets(
    cql028_versions: Path,
) -> None:
    prototype = ConsecutiveSheetPrototype()
    paths = sorted(cql028_versions.glob("*.xlsx"))

    for path in paths:
        snapshot, metrics = prototype.read(path)
        assert snapshot == read_workbook(path)
        assert metrics["xlsx_bytes"] == path.stat().st_size
        assert all(len(sheet["sha256"]) == 64 for sheet in metrics["worksheets"])
