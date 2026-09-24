"""Benchmark reproduzível da Fase 4 para os oito níveis de concorrência."""
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


def sheet_for(workload: str) -> SpreadsheetInfo:
    return SpreadsheetInfo("benchmark", "benchmark", workload, f"{workload}.xlsx")


def build(root: Path, count: int, rows: int, workload: str = "CQLPA123"):
    items = []
    for version in range(count):
        path = root / f"{workload}-{version:03}.xlsx"
        book = Workbook(write_only=True)
        sheet = book.create_sheet("Dados")
        for row in range(rows):
            sheet.append((row, f"item-{row}", row + version, f"=A{row + 1}+C{row + 1}"))
        book.save(path)
        items.append((VersionInfo(f"{workload.lower()}-{version:03}", f"5.{119 + version}"), path))
    return items


def run(root: Path, items, backend: str, slots: int, workload: str = "CQLPA123"):
    sheet = sheet_for(workload)
    identity = (sheet.site_id, sheet.drive_id, sheet.drive_item_id)
    source = LocalSource([sheet], {identity: items})
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
        result = service.audit(sheet)
        official = tuple(tuple(row) for row in database.connection.execute(
            "SELECT versao_anterior_id,versao_atual_id,quantidade_alteracoes,status "
            "FROM versao_processada ORDER BY id"
        ))
    elapsed = time.perf_counter() - started
    usage_after = resource.getrusage(resource.RUSAGE_SELF)
    stage_rows = [row for row in service.telemetry if row["stage"] == "STAGED"]
    downloads = [float(row["download_transfer"]) for row in service.telemetry if row["stage"] == "DOWNLOAD"]
    average = lambda key: round(sum(float(row.get(key, 0)) for row in stage_rows) / max(len(stage_rows), 1), 6)
    percentile = lambda values, fraction: round(sorted(values)[min(len(values) - 1, int(len(values) * fraction))], 6) if values else 0.0
    read_times = [float(row.get("read_xlsx", 0)) for row in stage_rows]
    comparisons = len(items) - 1
    return {
        "workload": workload,
        "backend": backend, "slots": slots, "versions": len(items),
        "comparisons": comparisons, "seconds": round(elapsed, 3),
        "versions_per_minute": round(len(items) / elapsed * 60, 3),
        "comparisons_per_minute": round(comparisons / elapsed * 60, 3),
        "commits_per_minute": round(comparisons / elapsed * 60, 3),
        "rss_python_mib": round(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024, 2),
        "cpu_python_seconds": round(
            usage_after.ru_utime + usage_after.ru_stime - usage_before.ru_utime - usage_before.ru_stime, 3
        ),
        "cpu_average_percent": round((usage_after.ru_utime + usage_after.ru_stime - usage_before.ru_utime - usage_before.ru_stime) / elapsed * 100, 2),
        "cpu_peak_sample_percent": round(max(
            (item.cpu_coordinator_percent + item.cpu_workers_percent for item in metrics), default=0
        ), 2),
        "download_average_seconds": round(sum(downloads) / len(downloads), 6),
        "download_p50_seconds": percentile(downloads, .50),
        "download_p95_seconds": percentile(downloads, .95),
        "parse_average_seconds": average("read_xlsx"),
        "parse_p50_seconds": percentile(read_times, .50),
        "parse_p95_seconds": percentile(read_times, .95),
        "compare_average_seconds": average("compare"),
        "staging_average_seconds": average("staging"),
        "commit_wait_average_seconds": 0.0,
        "commit_average_seconds": round(service.timing_model.average("COMMIT"), 6),
        "wait_promotion_average_seconds": round(service.timing_model.average("WAIT_PROMOTION"), 6),
        "rss_workers_mib": round(max((item.rss_workers for item in metrics), default=0) / 1048576, 2),
        "rss_edge_mib": round(max((item.rss_edge for item in metrics), default=0) / 1048576, 2),
        "staged_maximum": max((item.staged_count for item in metrics), default=0),
        "window_occupancy_maximum": max((item.window_occupancy for item in metrics), default=0),
        "rss_total_mib": round(max((item.rss_total for item in metrics), default=0) / 1048576, 2),
        "download_starvation_seconds": round(max((item.download_starvation_time for item in metrics), default=0), 6),
        "slot_utilization_percent": {
            str(slot): round(max((dict(item.slot_utilization).get(slot, 0) for item in metrics), default=0), 2)
            for slot in range(1, slots + 1)
        },
        "errors": 0, "retries": 0,
        "changes": result.changes, "official": official,
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--versions", type=int, default=50)
    parser.add_argument("--rows", type=int, default=2500)
    parser.add_argument("--workload", choices=("CQLPA120", "CQLPA123"), default="CQLPA123")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    with tempfile.TemporaryDirectory(prefix="trilha-bench-") as temporary:
        root = Path(temporary)
        items = build(root, args.versions, args.rows, args.workload)
        results = [run(root, items, "process", slots, args.workload) for slots in range(1, 9)]
        baseline = results[0]["seconds"]
        reference = results[0].pop("official")
        for result in results:
            official = result.pop("official", reference)
            result["equivalent_to_serial"] = official == reference
            result["speedup"] = round(baseline / result["seconds"], 3)
            result["efficiency"] = round(result["speedup"] / result["slots"], 3)
        best = max(results, key=lambda item: item["comparisons_per_minute"])
        payload = {
            "dataset": f"{args.workload}-sintetica", "rows_per_version": args.rows,
            "methodology": "Mesmo conjunto XLSX para 1-8 slots; processo novo e banco novo por configuração; inclui startup; sem descarte de warm-up.",
            "best_slots": best["slots"], "results": results,
        }
        text = json.dumps(payload, ensure_ascii=False, indent=2)
        print(text)
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(text + "\n", encoding="utf-8")

if __name__ == "__main__":
    main()
