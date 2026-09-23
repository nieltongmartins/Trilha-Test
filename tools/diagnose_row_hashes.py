"""Diagnose false negatives in the byte-for-byte XLSX row cache.

This command is deliberately offline and read-only.  It compares consecutive
files, uses the production reader as the semantic oracle, and emits no cell
values.  It does *not* provide a new cache signature.
"""

from __future__ import annotations

import argparse
from collections import Counter
from contextlib import contextmanager
import hashlib
import json
from pathlib import Path
import re
import sys
from time import perf_counter_ns
import tracemalloc
from typing import Any, Iterator
import xml.etree.ElementTree as ET
import zipfile

REPOSITORY_ROOT = Path(__file__).resolve().parents[1]
if str(REPOSITORY_ROOT) not in sys.path:
    sys.path.insert(0, str(REPOSITORY_ROOT))

from app.excel import reader


_COORDINATE_ROW = re.compile(r"^[A-Za-z]+([0-9]+)$")
_DIGITS = re.compile(r"([0-9]+)")


def _natural_key(path: Path) -> tuple[object, ...]:
    return tuple(int(part) if part.isdigit() else part.casefold()
                 for part in _DIGITS.split(str(path)))


def _typed_equal(left: object, right: object) -> bool:
    """Keep ``False`` distinct from ``0``, as the production comparator does."""
    if type(left) is not type(right):
        return False
    if isinstance(left, dict):
        return left.keys() == right.keys() and all(
            _typed_equal(left[key], right[key]) for key in left  # type: ignore[index]
        )
    return left == right


def _snapshot_rows(snapshot: reader.Snapshot) -> dict[str, dict[str, dict[str, object]]]:
    result: dict[str, dict[str, dict[str, object]]] = {}
    for title, cells in snapshot.items():
        rows: dict[str, dict[str, object]] = {}
        for coordinate, value in cells.items():
            match = _COORDINATE_ROW.match(coordinate)
            if match:
                rows.setdefault(match.group(1), {})[coordinate] = value
        result[title] = rows
    return result


def _safe_name(name: str) -> str:
    """Return an expanded XML name; never return an attribute value."""
    return name


def _tree_signature(element: ET.Element, *, strip_whitespace: bool) -> tuple[Any, ...]:
    def text(value: str | None) -> str | None:
        if strip_whitespace and (value is None or not value.strip()):
            return None
        return value

    return (
        element.tag,
        tuple(sorted(element.attrib.items())),
        text(element.text),
        text(element.tail),
        tuple(_tree_signature(child, strip_whitespace=strip_whitespace)
              for child in element),
    )


def _shape(element: ET.Element) -> tuple[int, int, bool, bool]:
    descendants = list(element.iter())
    cells = sum(node.tag == reader._tag("c") for node in descendants)
    formulas = any(node.tag == reader._tag("f") for node in descendants)
    styles = any(
        node.tag == reader._tag("c") and "s" in node.attrib
        for node in descendants
    )
    return cells, len(descendants), formulas, styles


def _changed_attributes(before: ET.Element, after: ET.Element) -> list[str]:
    changed: set[str] = set()
    before_nodes = list(before.iter())
    after_nodes = list(after.iter())
    for left, right in zip(before_nodes, after_nodes):
        if left.tag != right.tag:
            continue
        for name in left.attrib.keys() | right.attrib.keys():
            if left.attrib.get(name) != right.attrib.get(name):
                changed.add(_safe_name(name))
    return sorted(changed)


def _difference_types(before: ET.Element, after: ET.Element) -> list[str]:
    """Classify without placing text or attribute values in the result."""
    categories: set[str] = set()
    if _tree_signature(before, strip_whitespace=False) == _tree_signature(
        after, strip_whitespace=False
    ):
        return ["serializacao_lexica"]
    if _tree_signature(before, strip_whitespace=True) == _tree_signature(
        after, strip_whitespace=True
    ):
        categories.add("whitespace")

    before_nodes, after_nodes = list(before.iter()), list(after.iter())
    if [node.tag for node in before_nodes] != [node.tag for node in after_nodes]:
        categories.add("elementos_ou_ordem")
    for left, right in zip(before_nodes, after_nodes):
        if left.tag != right.tag:
            continue
        changed = {
            name for name in left.attrib.keys() | right.attrib.keys()
            if left.attrib.get(name) != right.attrib.get(name)
        }
        if changed:
            if left is before:
                categories.add("atributos_row")
            elif left.tag == reader._tag("c"):
                categories.add("atributos_celula")
                if "s" in changed:
                    categories.add("estilo_id")
                if "t" in changed:
                    categories.add("tipo_celula")
            elif left.tag == reader._tag("f"):
                categories.add("atributos_formula")
            else:
                categories.add("atributos_outros")
        if (left.text or "") != (right.text or ""):
            if left.tag == reader._tag("v"):
                categories.add("valor_xml_ou_shared_string_id")
            elif left.tag == reader._tag("f"):
                categories.add("formula")
            elif left.tag in {reader._tag("t"), reader._tag("is")}:
                categories.add("texto_inline")
            elif (left.text or "").strip() or (right.text or "").strip():
                categories.add("texto_outro")
    if any(not str(node.tag).startswith("{" + reader._MAIN_NS + "}")
           for node in before_nodes + after_nodes):
        categories.add("extensao_ou_namespace_externo")
    return sorted(categories or {"estrutura_nao_classificada"})


def _rows(archive: zipfile.ZipFile, target: str) -> tuple[dict[str, tuple[str, bytes, ET.Element]], int]:
    xml = archive.read(target)
    wrapper = reader._namespace_wrapper(xml)
    result: dict[str, tuple[str, bytes, ET.Element]] = {}
    hashing_ns = 0
    for key, _digest_from_scanner, payload, _shared in reader._iter_row_parts(xml, wrapper):
        started = perf_counter_ns()
        digest = hashlib.sha256(payload).hexdigest()
        hashing_ns += perf_counter_ns() - started
        result[key] = (digest, payload, reader._row_element(payload, wrapper))
    return result, hashing_ns


def _member_hash(archive: zipfile.ZipFile, name: str) -> str:
    if name not in archive.namelist():
        return "missing"
    return hashlib.sha256(archive.read(name)).hexdigest()


def _xml_subtree_hash(archive: zipfile.ZipFile, member: str, tag: str) -> str:
    if member not in archive.namelist():
        return "missing"
    root = ET.fromstring(archive.read(member))
    child = root.find(tag, reader._NS)
    return "missing" if child is None else hashlib.sha256(
        ET.tostring(child, encoding="utf-8")
    ).hexdigest()


def _dependencies(before: zipfile.ZipFile, after: zipfile.ZipFile) -> dict[str, Any]:
    members = {
        "sharedStrings": "xl/sharedStrings.xml",
        "styles": "xl/styles.xml",
        "relationships": "xl/_rels/workbook.xml.rels",
        "workbook_metadata": "xl/workbook.xml",
    }
    result = {
        name: {"equal": _member_hash(before, member) == _member_hash(after, member)}
        for name, member in members.items()
    }
    result["numFmt"] = {
        "equal": _xml_subtree_hash(before, "xl/styles.xml", "m:numFmts")
        == _xml_subtree_hash(after, "xl/styles.xml", "m:numFmts")
    }
    result["cellXfs"] = {
        "equal": _xml_subtree_hash(before, "xl/styles.xml", "m:cellXfs")
        == _xml_subtree_hash(after, "xl/styles.xml", "m:cellXfs")
    }
    result["epoch"] = {
        "equal": reader._workbook_epoch(before) == reader._workbook_epoch(after),
        "before": "1904" if reader._workbook_epoch(before) == reader.CALENDAR_MAC_1904 else "1900",
        "after": "1904" if reader._workbook_epoch(after) == reader.CALENDAR_MAC_1904 else "1900",
    }
    return result


@contextmanager
def _archives(before: Path, after: Path) -> Iterator[tuple[zipfile.ZipFile, zipfile.ZipFile]]:
    with zipfile.ZipFile(before) as left, zipfile.ZipFile(after) as right:
        yield left, right


def diagnose_pair(before_path: Path, after_path: Path, sample_limit: int = 5) -> dict[str, Any]:
    """Return privacy-safe evidence for one pair of consecutive workbooks."""
    semantic_parse_started = perf_counter_ns()
    before_snapshot = _snapshot_rows(reader.read_workbook(before_path))
    after_snapshot = _snapshot_rows(reader.read_workbook(after_path))
    semantic_oracle_ns = perf_counter_ns() - semantic_parse_started
    samples: dict[str, list[dict[str, Any]]] = {"A": [], "B": [], "C": []}
    categories: Counter[str] = Counter()
    totals = Counter()
    hashing_ns = canonicalization_ns = semantic_comparison_ns = 0

    with _archives(before_path, after_path) as (before_zip, after_zip):
        dependencies = _dependencies(before_zip, after_zip)
        before_targets = dict(reader._sheet_targets(before_zip))
        after_targets = dict(reader._sheet_targets(after_zip))
        for title in before_targets.keys() & after_targets.keys():
            before_rows, elapsed = _rows(before_zip, before_targets[title])
            hashing_ns += elapsed
            after_rows, elapsed = _rows(after_zip, after_targets[title])
            hashing_ns += elapsed
            totals["rows_added_or_removed"] += len(
                before_rows.keys() ^ after_rows.keys()
            )
            for key in before_rows.keys() & after_rows.keys():
                totals["rows_total"] += 1
                before_digest, before_payload, before_element = before_rows[key]
                after_digest, after_payload, after_element = after_rows[key]
                same_sha = before_digest == after_digest
                row_number = key.partition("#")[0]
                compare_started = perf_counter_ns()
                logically_equal = _typed_equal(
                    before_snapshot.get(title, {}).get(row_number, {}),
                    after_snapshot.get(title, {}).get(row_number, {}),
                )
                semantic_comparison_ns += perf_counter_ns() - compare_started
                category = (
                    "C" if same_sha and logically_equal
                    else "A" if not same_sha and logically_equal
                    else "B" if not same_sha
                    else None
                )
                totals["sha_equal" if same_sha else "sha_different"] += 1
                if not same_sha:
                    totals["different_logically_equal" if logically_equal
                           else "different_logically_changed"] += 1
                    canonical_started = perf_counter_ns()
                    difference_types = _difference_types(before_element, after_element)
                    canonicalization_ns += perf_counter_ns() - canonical_started
                    categories.update(difference_types)
                else:
                    difference_types = []
                    if not logically_equal:
                        # A global input reinterpreted byte-identical XML. This
                        # is intentionally outside A/B/C and proves that a row
                        # digest alone can never authorize reuse.
                        totals["sha_equal_logically_changed_global_context"] += 1
                if category is not None and len(samples[category]) < sample_limit:
                    before_shape = _shape(before_element)
                    after_shape = _shape(after_element)
                    samples[category].append({
                        "sheet": title,
                        "row_number": row_number,
                        "cells_before": before_shape[0],
                        "cells_after": after_shape[0],
                        "elements_before": before_shape[1],
                        "elements_after": after_shape[1],
                        "has_formula_before": before_shape[2],
                        "has_formula_after": after_shape[2],
                        "has_style_before": before_shape[3],
                        "has_style_after": after_shape[3],
                        "changed_attribute_names": (
                            _changed_attributes(before_element, after_element)
                            if not same_sha else []
                        ),
                        "difference_types": difference_types,
                        "xml_bytes_before": len(before_payload),
                        "xml_bytes_after": len(after_payload),
                        "sha256_before": before_digest,
                        "sha256_after": after_digest,
                    })

        # Rows in an added/removed sheet have no same-sheet candidate. Counting
        # them requires scanning but never interpreting their values.
        for title in before_targets.keys() - after_targets.keys():
            missing_rows, elapsed = _rows(before_zip, before_targets[title])
            hashing_ns += elapsed
            totals["rows_added_or_removed"] += len(missing_rows)
        for title in after_targets.keys() - before_targets.keys():
            missing_rows, elapsed = _rows(after_zip, after_targets[title])
            hashing_ns += elapsed
            totals["rows_added_or_removed"] += len(missing_rows)

    different = totals["sha_different"]
    return {
        "before": before_path.name,
        "after": after_path.name,
        **totals,
        "raw_hash_false_negative_percent": (
            100.0 * totals["different_logically_equal"] / different
            if different else 0.0
        ),
        "difference_categories": dict(categories.most_common()),
        "global_dependencies": dependencies,
        "samples": samples,
        "timing_ns": {
            "raw_sha256": hashing_ns,
            "canonicalization_and_structural_classification": canonicalization_ns,
            "semantic_oracle_full_read": semantic_oracle_ns,
            "semantic_row_comparison": semantic_comparison_ns,
        },
    }


def diagnose(paths: list[Path], *, max_pairs: int = 50, sample_limit: int = 5) -> dict[str, Any]:
    files = sorted(dict.fromkeys(path.resolve() for path in paths), key=_natural_key)
    if len(files) < 2:
        raise ValueError("informe pelo menos dois XLSX consecutivos")
    selected = list(zip(files, files[1:]))[:max_pairs]
    tracemalloc.start()
    started = perf_counter_ns()
    pairs = [diagnose_pair(before, after, sample_limit) for before, after in selected]
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    aggregate = Counter()
    categories: Counter[str] = Counter()
    timing: Counter[str] = Counter()
    dependencies: Counter[str] = Counter()
    for pair in pairs:
        for name in ("rows_total", "sha_equal", "sha_different",
                     "different_logically_equal", "different_logically_changed",
                     "sha_equal_logically_changed_global_context",
                     "rows_added_or_removed"):
            aggregate[name] += pair.get(name, 0)
        categories.update(pair["difference_categories"])
        timing.update(pair["timing_ns"])
        for name, state in pair["global_dependencies"].items():
            if not state["equal"]:
                dependencies[name] += 1
    different = aggregate["sha_different"]
    return {
        "schema_version": 1,
        "diagnostic_only": True,
        "privacy": "cell values and formula text are never emitted",
        "pair_count": len(pairs),
        "aggregate": {
            **aggregate,
            "raw_hash_false_negative_percent": (
                100.0 * aggregate["different_logically_equal"] / different
                if different else 0.0
            ),
            "difference_categories": dict(categories.most_common()),
            "dependency_changes_by_pair": dict(dependencies.most_common()),
            "timing_ns": dict(timing),
            "diagnostic_wall_ns": perf_counter_ns() - started,
            "peak_traced_bytes": peak,
        },
        "pairs": pairs,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Diagnostica SHA diferente em rows logicamente iguais sem expor valores"
    )
    parser.add_argument("paths", nargs="+", type=Path, help="XLSX ou diretórios em ordem de versão")
    parser.add_argument("--max-pairs", type=int, default=50)
    parser.add_argument("--sample-limit", type=int, default=5)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args(argv)
    if args.max_pairs < 1 or args.sample_limit < 0:
        parser.error("--max-pairs deve ser positivo e --sample-limit não negativo")
    files: list[Path] = []
    for path in args.paths:
        files.extend(path.rglob("*.xlsx") if path.is_dir() else [path])
    try:
        result = diagnose(files, max_pairs=args.max_pairs, sample_limit=args.sample_limit)
    except (ValueError, OSError, zipfile.BadZipFile, ET.ParseError) as error:
        parser.error(str(error))
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.write_text(payload + "\n", encoding="utf-8")
    else:
        print(payload)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
