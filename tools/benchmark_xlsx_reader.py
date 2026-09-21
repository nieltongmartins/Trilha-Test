"""Profile the official XLSX reader without changing audit state or checkpoints.

This is a diagnostic command, not an application entry point.  It deliberately
duplicates the XML hot loop so experimental parsing can be timed and its result
compared byte-for-byte (values *and* Python types) with ``read_workbook``.
"""

from __future__ import annotations

import argparse
from collections import Counter
from dataclasses import asdict, dataclass, field
from io import BytesIO
import gc
import json
from pathlib import Path
from statistics import mean, median
import sys
from time import perf_counter_ns
from typing import Any, Callable
import xml.etree.ElementTree as ET
import zipfile

# Allow the documented ``python tools/...`` invocation without installation.
REPOSITORY_ROOT = Path(__file__).resolve().parents[1]
if str(REPOSITORY_ROOT) not in sys.path:
    sys.path.insert(0, str(REPOSITORY_ROOT))

from app.excel import reader


NS = reader._NS
CELL_TAG = reader._tag("c")
FORMULA_TAG = reader._tag("f")
VALUE_TAG = reader._tag("v")
INLINE_TAG = reader._tag("is")
SI_TAG = reader._tag("si")


def _rss_bytes() -> int | None:
    """Return current RSS where the operating system exposes it cheaply."""
    try:
        fields = Path("/proc/self/statm").read_text(encoding="ascii").split()
        return int(fields[1]) * int(__import__("os").sysconf("SC_PAGE_SIZE"))
    except (OSError, ValueError, IndexError):
        return None


@dataclass
class TimerTotals:
    nanoseconds: Counter[str] = field(default_factory=Counter)

    def call(self, name: str, function: Callable[..., Any], *args: Any) -> Any:
        started = perf_counter_ns()
        try:
            return function(*args)
        finally:
            self.nanoseconds[name] += perf_counter_ns() - started


@dataclass
class SheetProfile:
    title: str
    target: str
    compressed_bytes: int
    xml_bytes: int
    cells_seen: int = 0
    cells_stored: int = 0
    formulas: int = 0
    shared_formula_translations: int = 0
    cell_types: Counter[str] = field(default_factory=Counter)
    nanoseconds: Counter[str] = field(default_factory=Counter)


def _children(cell: ET.Element) -> tuple[ET.Element | None, ET.Element | None, ET.Element | None]:
    formula = value = inline = None
    for child in cell:
        if child.tag == FORMULA_TAG:
            formula = child
        elif child.tag == VALUE_TAG:
            value = child
        elif child.tag == INLINE_TAG:
            inline = child
    return formula, value, inline


def _profiled_cell_value(
    cell: ET.Element,
    coordinate: str,
    shared: list[str],
    date_styles: set[int],
    epoch: object,
    shared_formulas: dict[str, tuple[str, str]],
    profile: SheetProfile,
    direct_children: bool,
) -> object:
    started = perf_counter_ns()
    formula, value_element, inline = _children(cell) if direct_children else (None, None, None)
    formula_started = perf_counter_ns()
    if direct_children:
        # _formula_value is reproduced only to benchmark removal of repeated find().
        if formula is None:
            value = None
        else:
            formula_type = formula.get("t")
            if formula_type in {"array", "dataTable"}:
                raise reader._FastReaderUnsupported(f"fórmula {formula_type} em {coordinate}")
            text = formula.text or ""
            if formula_type == "shared":
                shared_index = formula.get("si")
                if shared_index is None:
                    raise reader._FastReaderUnsupported("fórmula compartilhada sem si")
                if text:
                    value = "=" + text
                    shared_formulas[shared_index] = (coordinate, value)
                else:
                    master = shared_formulas.get(shared_index)
                    if master is None:
                        raise reader._FastReaderUnsupported("fórmula compartilhada sem mestre")
                    translate_started = perf_counter_ns()
                    value = reader.Translator(master[1], origin=master[0]).translate_formula(coordinate)
                    profile.nanoseconds["shared_formula_translation"] += perf_counter_ns() - translate_started
                    profile.shared_formula_translations += 1
            else:
                value = "=" + text
    else:
        value = reader._formula_value(cell, coordinate, shared_formulas)
    profile.nanoseconds["formula"] += perf_counter_ns() - formula_started
    if value is not None:
        profile.formulas += 1
        profile.nanoseconds["cell_value"] += perf_counter_ns() - started
        return value

    cell_type = cell.get("t", "n")
    profile.cell_types[cell_type] += 1
    if not direct_children:
        value_element = cell.find("m:v", NS)
    raw = value_element.text if value_element is not None else None
    conversion_started = perf_counter_ns()
    phase = "number"
    if cell_type == "inlineStr":
        if not direct_children:
            inline = cell.find("m:is", NS)
        result = reader._all_text(inline) or None
        phase = "inline_string"
    elif raw is None:
        result = None
        phase = "empty"
    elif cell_type == "s":
        result = shared[int(raw)]
        phase = "shared_string"
    elif cell_type == "b":
        result = raw == "1"
        phase = "boolean"
    elif cell_type in {"str", "e"}:
        result = raw
        phase = "error" if cell_type == "e" else "string"
    elif cell_type == "d":
        result = reader.from_ISO8601(raw)
        phase = "date"
    else:
        result = reader._cast_number(raw)
        style_text = cell.get("s")
        if style_text is not None:
            try:
                style_id = int(style_text)
            except ValueError:
                style_id = -1
            if style_id in date_styles:
                result = reader.from_excel(result, epoch)
                phase = "date"
    profile.nanoseconds[phase] += perf_counter_ns() - conversion_started
    profile.nanoseconds["cell_value"] += perf_counter_ns() - started
    return result


def _profile_variant(path: Path, *, direct_children: bool, disable_gc: bool) -> tuple[dict[str, Any], dict[str, Any]]:
    timers = TimerTotals()
    profiles: list[SheetProfile] = []
    rss_before = _rss_bytes()
    gc_before = gc.get_count()
    total_started = perf_counter_ns()
    archive_started = perf_counter_ns()
    archive = zipfile.ZipFile(path)
    timers.nanoseconds["zip_open"] += perf_counter_ns() - archive_started
    try:
        names = set(archive.namelist())
        info_by_name = {item.filename: item for item in archive.infolist()}

        shared: list[str] = []
        if "xl/sharedStrings.xml" in names:
            raw_shared = timers.call("shared_strings_decompression", archive.read, "xl/sharedStrings.xml")
            parse_started = perf_counter_ns()
            for _, element in ET.iterparse(BytesIO(raw_shared), events=("end",)):
                if element.tag == SI_TAG:
                    shared.append(reader._all_text(element))
                    element.clear()
            timers.nanoseconds["shared_strings_iterparse"] += perf_counter_ns() - parse_started

        styles_raw = timers.call("styles_decompression", archive.read, "xl/styles.xml") if "xl/styles.xml" in names else None
        if styles_raw is None:
            date_styles: set[int] = set()
        else:
            # Reuse the official function through a tiny in-memory archive would
            # distort timing; its exact style algorithm is timed against the ZIP.
            date_styles = timers.call("styles_parse", reader._date_styles, archive)

        workbook_raw = timers.call("workbook_decompression", archive.read, "xl/workbook.xml")
        timers.call("workbook_parse", ET.fromstring, workbook_raw)
        epoch = timers.call("workbook_epoch", reader._workbook_epoch, archive)
        rels_raw = timers.call("relationships_decompression", archive.read, "xl/_rels/workbook.xml.rels")
        timers.call("relationships_parse", ET.fromstring, rels_raw)
        sheets = timers.call("sheet_targets", reader._sheet_targets, archive)

        snapshot: dict[str, Any] = {}
        gc_was_enabled = gc.isenabled()
        if disable_gc and gc_was_enabled:
            gc.disable()
        try:
            for title, target in sheets:
                info = info_by_name[target]
                profile = SheetProfile(title, target, info.compress_size, info.file_size)
                profiles.append(profile)
                xml_started = perf_counter_ns()
                xml = archive.read(target)
                profile.nanoseconds["xml_decompression"] += perf_counter_ns() - xml_started
                cells: dict[str, Any] = {}
                shared_formulas: dict[str, tuple[str, str]] = {}
                parse_started = perf_counter_ns()
                iterator = ET.iterparse(BytesIO(xml), events=("end",))
                for _, element in iterator:
                    if element.tag != CELL_TAG:
                        continue
                    profile.cells_seen += 1
                    coordinate_started = perf_counter_ns()
                    coordinate = element.get("r")
                    profile.nanoseconds["coordinate"] += perf_counter_ns() - coordinate_started
                    if coordinate:
                        value = _profiled_cell_value(
                            element, coordinate, shared, date_styles, epoch,
                            shared_formulas, profile, direct_children,
                        )
                        if value is not None:
                            insert_started = perf_counter_ns()
                            cells[coordinate] = value
                            profile.nanoseconds["dict_insert"] += perf_counter_ns() - insert_started
                            profile.cells_stored += 1
                    element.clear()
                profile.nanoseconds["iterparse_hot_loop"] += perf_counter_ns() - parse_started
                snapshot_started = perf_counter_ns()
                snapshot[title] = cells
                profile.nanoseconds["snapshot_insert"] += perf_counter_ns() - snapshot_started
        finally:
            if disable_gc and gc_was_enabled:
                gc.enable()
    finally:
        archive.close()
    parse_total_ns = perf_counter_ns() - total_started
    gc_started = perf_counter_ns()
    collected = gc.collect()
    timers.nanoseconds["garbage_collection"] += perf_counter_ns() - gc_started
    total_with_gc_ns = perf_counter_ns() - total_started
    metadata = {
        "total_ns": parse_total_ns,
        "total_with_forced_gc_ns": total_with_gc_ns,
        "phases_ns": dict(timers.nanoseconds),
        "worksheets": [
            {**asdict(item), "cell_types": dict(item.cell_types), "nanoseconds": dict(item.nanoseconds)}
            for item in profiles
        ],
        "worksheet_count": len(profiles),
        "cells_stored": sum(item.cells_stored for item in profiles),
        "xml_uncompressed_bytes": sum(item.xml_bytes for item in profiles),
        "shared_strings": len(shared),
        "snapshot_entries": sum(len(values) for values in snapshot.values()),
        "rss_before": rss_before,
        "rss_after": _rss_bytes(),
        "gc_count_before": gc_before,
        "gc_collected": collected,
    }
    return snapshot, metadata


def _typed_equal(left: object, right: object) -> bool:
    if type(left) is not type(right):
        return False
    if isinstance(left, dict):
        return left.keys() == right.keys() and all(_typed_equal(left[key], right[key]) for key in left)  # type: ignore[index,union-attr]
    return left == right


def _summary(samples: list[int]) -> dict[str, int | float]:
    return {
        "median_ns": median(samples),
        "mean_ns": mean(samples),
        "min_ns": min(samples),
        "max_ns": max(samples),
    }


def benchmark_official_before_after(path: Path, repeat: int = 5) -> dict[str, Any]:
    """Measure the old and optimized real reader paths after one warm-up."""
    before_snapshot = reader._read_fast_repeated_find(path)
    optimized_snapshot = reader.read_workbook(path)
    rss_before = _rss_bytes()
    before_times: list[int] = []
    optimized_times: list[int] = []
    for _ in range(repeat):
        started = perf_counter_ns()
        before_snapshot = reader._read_fast_repeated_find(path)
        before_times.append(perf_counter_ns() - started)

        started = perf_counter_ns()
        optimized_snapshot = reader.read_workbook(path)
        optimized_times.append(perf_counter_ns() - started)
    rss_after = _rss_bytes()
    before_summary = _summary(before_times)
    optimized_summary = _summary(optimized_times)
    before_median = float(before_summary["median_ns"])
    return {
        "repetitions_after_warmup": repeat,
        "cells": sum(len(cells) for cells in optimized_snapshot.values()),
        "snapshot_exactly_equal": _typed_equal(before_snapshot, optimized_snapshot),
        "previous_reader": {"samples_ns": before_times, **before_summary},
        "optimized_reader": {"samples_ns": optimized_times, **optimized_summary},
        "median_gain_percent": (
            before_median - float(optimized_summary["median_ns"])
        ) * 100 / before_median,
        "rss_before": rss_before,
        "rss_after": rss_after,
        "rss_delta": None if rss_before is None or rss_after is None else rss_after - rss_before,
    }


def profile_file(path: Path, repeat: int = 1) -> dict[str, Any]:
    rss_before = _rss_bytes()
    official_times: list[int] = []
    official = None
    for _ in range(repeat):
        started = perf_counter_ns()
        official = reader.read_workbook(path)
        official_times.append(perf_counter_ns() - started)
    assert official is not None

    variants: dict[str, Any] = {}
    for name, direct, disable_gc in (
        ("instrumented_current", False, False),
        ("direct_children", True, False),
        ("direct_children_gc_disabled", True, True),
    ):
        try:
            snapshot, metrics = _profile_variant(path, direct_children=direct, disable_gc=disable_gc)
            metrics["fallback_openpyxl_ns"] = 0
            metrics["fallback_reason"] = None
        except reader._FastReaderUnsupported as error:
            # Mirror the public reader's integrity fallback and expose its cost.
            started = perf_counter_ns()
            snapshot = reader._read_openpyxl(path)
            elapsed = perf_counter_ns() - started
            metrics = {
                "total_ns": elapsed,
                "phases_ns": {},
                "worksheets": [],
                "worksheet_count": len(snapshot),
                "cells_stored": sum(len(cells) for cells in snapshot.values()),
                "xml_uncompressed_bytes": None,
                "shared_strings": None,
                "snapshot_entries": sum(len(cells) for cells in snapshot.values()),
                "rss_before": None,
                "rss_after": _rss_bytes(),
                "fallback_openpyxl_ns": elapsed,
                "fallback_reason": str(error),
            }
        metrics["snapshot_exactly_equal"] = _typed_equal(official, snapshot)
        baseline = metrics["total_ns"] if name == "instrumented_current" else variants["instrumented_current"]["total_ns"]
        metrics["gain_percent_vs_instrumented_current"] = (baseline - metrics["total_ns"]) * 100 / baseline
        variants[name] = metrics

    return {
        "file": str(path),
        "zip_bytes": path.stat().st_size,
        "official_reader_ns": official_times,
        "official_reader_median_ns": sorted(official_times)[len(official_times) // 2],
        "rss_before": rss_before,
        "rss_after": _rss_bytes(),
        "official_before_after": benchmark_official_before_after(path, repeat),
        "variants": variants,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Profile e compare o leitor XLSX oficial")
    parser.add_argument("paths", nargs="+", type=Path, help="arquivos ou diretórios com versões .xlsx")
    parser.add_argument(
        "--repeat", type=int, default=5,
        help="repetições após warm-up dos leitores oficiais anterior e otimizado",
    )
    parser.add_argument("--output", type=Path, help="JSON de saída (stdout por padrão)")
    args = parser.parse_args(argv)
    if args.repeat < 1:
        parser.error("--repeat deve ser positivo")
    files: list[Path] = []
    for path in args.paths:
        files.extend(sorted(path.rglob("*.xlsx")) if path.is_dir() else [path])
    files = list(dict.fromkeys(item.resolve() for item in files))
    if not files:
        parser.error("nenhum arquivo .xlsx encontrado")
    results = {"schema_version": 1, "files": [profile_file(path, args.repeat) for path in files]}
    payload = json.dumps(results, ensure_ascii=False, indent=2, default=str)
    if args.output:
        args.output.write_text(payload + "\n", encoding="utf-8")
    else:
        print(payload)
    return 1 if any(
        not variant["snapshot_exactly_equal"]
        for result in results["files"]
        for variant in result["variants"].values()
    ) else 0


if __name__ == "__main__":
    sys.exit(main())
