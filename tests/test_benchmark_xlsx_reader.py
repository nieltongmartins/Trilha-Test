from pathlib import Path

from tools.benchmark_xlsx_reader import (
    benchmark_official_before_after,
    main,
    profile_file,
)


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
    assert '"schema_version": 1' in payload
    assert payload.count('"snapshot_exactly_equal": true') == 16


def test_real_reader_benchmark_uses_five_post_warmup_runs(
    cql028_versions: Path,
) -> None:
    result = benchmark_official_before_after(cql028_versions / "0.85.xlsx")

    assert result["snapshot_exactly_equal"] is True
    assert result["repetitions_after_warmup"] == 5
    assert len(result["previous_reader"]["samples_ns"]) == 5
    assert len(result["optimized_reader"]["samples_ns"]) == 5
