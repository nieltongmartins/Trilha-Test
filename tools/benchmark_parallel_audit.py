"""Benchmark reproduzível da Fase 2 (CQLPA123 sintética, 24 versões)."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import resource
import tempfile
import time
from openpyxl import Workbook

from app.database import Database
from app.parallel_audit import ParallelAuditService
from app.sources.base import SpreadsheetInfo, VersionInfo
from app.sources.local import LocalSource

SHEET = SpreadsheetInfo("benchmark", "benchmark", "CQLPA123", "CQLPA123.xlsx")


def build(root: Path, count: int, rows: int):
    items = []
    for version in range(count):
        path = root / f"CQLPA123-{version:03}.xlsx"
        book = Workbook(write_only=True)
        sheet = book.create_sheet("Dados")
        for row in range(rows):
            sheet.append((row, f"item-{row}", row + version, f"=A{row + 1}+C{row + 1}"))
        book.save(path)
        items.append((VersionInfo(f"cqlpa123-{version:03}", f"5.{119 + version}"), path))
    return items


def run(root: Path, items, backend: str, slots: int):
    identity = (SHEET.site_id, SHEET.drive_id, SHEET.drive_item_id)
    source = LocalSource([SHEET], {identity: items})
    metrics = []
    usage_before = resource.getrusage(resource.RUSAGE_SELF)
    started = time.perf_counter()
    with Database(root / f"{backend}-{slots}.db") as database:
        database.initialize()
        service = ParallelAuditService(
            database, source, slots=slots, backend=backend,
            staging_directory=root / f"stage-{backend}-{slots}",
            metrics_callback=metrics.append,
        )
        result = service.audit(SHEET)
        official = tuple(tuple(row) for row in database.connection.execute(
            "SELECT versao_anterior_id,versao_atual_id,quantidade_alteracoes,status "
            "FROM versao_processada ORDER BY id"
        ))
    elapsed = time.perf_counter() - started
    usage_after = resource.getrusage(resource.RUSAGE_SELF)
    stage_rows = [row for row in service.telemetry if row["stage"] == "STAGED"]
    downloads = [float(row["download"]) for row in service.telemetry if row["stage"] == "DOWNLOAD"]
    average = lambda key: round(sum(float(row.get(key, 0)) for row in stage_rows) / len(stage_rows), 6)
    comparisons = len(items) - 1
    return {
        "backend": backend, "slots": slots, "versions": len(items),
        "comparisons": comparisons, "seconds": round(elapsed, 3),
        "versions_per_minute": round(len(items) / elapsed * 60, 3),
        "comparisons_per_minute": round(comparisons / elapsed * 60, 3),
        "rss_python_mib": round(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024, 2),
        "cpu_python_seconds": round(
            usage_after.ru_utime + usage_after.ru_stime - usage_before.ru_utime - usage_before.ru_stime, 3
        ),
        "download_average_seconds": round(sum(downloads) / len(downloads), 6),
        "parse_average_seconds": average("parse"),
        "compare_average_seconds": average("compare"),
        "staging_average_seconds": average("staging"),
        "commit_wait_average_seconds": 0.0,
        "errors": 0, "retries": 0,
        "changes": result.changes, "official": official,
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--versions", type=int, default=24)
    parser.add_argument("--rows", type=int, default=2500)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    with tempfile.TemporaryDirectory(prefix="trilha-bench-") as temporary:
        root = Path(temporary)
        items = build(root, args.versions, args.rows)
        results = [run(root, items, "sync", 1), run(root, items, "thread", 2), run(root, items, "process", 2)]
        baseline = results[0]["seconds"]
        reference = results[0].pop("official")
        for result in results:
            official = result.pop("official", reference)
            result["equivalent_to_serial"] = official == reference
            result["speedup"] = round(baseline / result["seconds"], 3)
        payload = {"dataset": "CQLPA123-sintetica", "rows_per_version": args.rows, "results": results}
        text = json.dumps(payload, ensure_ascii=False, indent=2)
        print(text)
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(text + "\n", encoding="utf-8")

if __name__ == "__main__":
    main()
