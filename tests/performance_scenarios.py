"""Cenários manuais e sintéticos de performance da F6.

Este módulo não é coletado pelo pytest. Execute-o explicitamente para medir a
arquitetura sem acessar o SharePoint e sem deixar bancos ou XLSX de carga no
repositório::

    python tests/performance_scenarios.py

Os volumes padrão representam uma planilha com 3.001 versões, 100.000
alterações e um banco compartilhado por 2.000 planilhas. Um único XLSX local é
reutilizado como conteúdo controlado das versões para evitar milhares de
arquivos temporários; o fluxo completo de leitura, comparação e persistência
continua sendo exercitado para cada par.
"""

from __future__ import annotations

import json
from pathlib import Path
import sqlite3
import sys
from tempfile import TemporaryDirectory
from time import perf_counter
import tracemalloc

from openpyxl import Workbook


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.audit_service import AuditService  # noqa: E402
from app.database import Database  # noqa: E402
from app.report_service import ReportService  # noqa: E402
from app.sources.base import SpreadsheetInfo, VersionInfo  # noqa: E402


VERSION_COUNT = 3_001
CHANGE_COUNT = 100_000
SPREADSHEET_COUNT = 2_000


class SyntheticSource:
    """Fonte local instrumentada que reutiliza um XLSX controlado."""

    def __init__(self, workbook: Path, versions: list[VersionInfo]) -> None:
        self.workbook = workbook
        self.versions = versions
        self.downloads = 0

    def list_spreadsheets(self) -> tuple[SpreadsheetInfo, ...]:
        return (SPREADSHEET,)

    def list_versions(self, spreadsheet: SpreadsheetInfo) -> tuple[VersionInfo, ...]:
        return tuple(self.versions)

    def get_version(self, spreadsheet: SpreadsheetInfo, version: VersionInfo) -> Path:
        self.downloads += 1
        return self.workbook


SPREADSHEET = SpreadsheetInfo(
    site_id="site-performance",
    drive_id="drive-performance",
    drive_item_id="item-performance",
    name="Performance.xlsx",
    path="/synthetic/Performance.xlsx",
)


def measured(operation):
    """Retorna resultado, segundos e pico Python medido pelo tracemalloc."""
    tracemalloc.start()
    started = perf_counter()
    result = operation()
    elapsed = perf_counter() - started
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    return result, elapsed, peak


def versions(count: int) -> list[VersionInfo]:
    return [
        VersionInfo(id=f"version-{index}", number=str(index), size=4_800)
        for index in range(count)
    ]


def create_workbook(path: Path, cell_count: int = 100) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Dados"
    for index in range(1, cell_count + 1):
        sheet.cell(row=index, column=1, value=index)
    workbook.save(path)
    workbook.close()


def spreadsheet_id(connection: sqlite3.Connection) -> int:
    row = connection.execute(
        """SELECT id FROM planilha
           WHERE site_id = ? AND drive_id = ? AND drive_item_id = ?""",
        (SPREADSHEET.site_id, SPREADSHEET.drive_id, SPREADSHEET.drive_item_id),
    ).fetchone()
    assert row is not None
    return int(row["id"])


def seed_many_spreadsheets(connection: sqlite3.Connection) -> None:
    connection.executemany(
        """INSERT INTO planilha
           (drive_item_id, nome_atual, site_id, drive_id, caminho_sharepoint)
           VALUES (?, ?, ?, ?, ?)""",
        [
            (f"item-{index}", f"Planilha {index}.xlsx", "site-lote", "drive-lote", f"/lote/{index}")
            for index in range(SPREADSHEET_COUNT - 1)
        ],
    )
    connection.commit()


def seed_changes(connection: sqlite3.Connection, planilha_id: int) -> None:
    execution_id = connection.execute(
        """INSERT INTO execucao_auditoria
           (codigo_execucao, planilha_id, status)
           VALUES ('PERF-CHANGES', ?, 'CONCLUIDA')""",
        (planilha_id,),
    ).lastrowid
    processed_id = connection.execute(
        """INSERT INTO versao_processada
           (planilha_id, versao_anterior_id, versao_anterior_numero,
            versao_atual_id, versao_atual_numero, quantidade_alteracoes,
            status, execucao_id)
           VALUES (?, 'changes-before', 'C0', 'changes-after', 'C1', ?,
                   'PROCESSADA', ?)""",
        (planilha_id, CHANGE_COUNT, execution_id),
    ).lastrowid
    connection.executemany(
        """INSERT INTO alteracao
           (versao_processada_id, planilha_id, tipo, aba, endereco,
            valor_anterior, valor_novo)
           VALUES (?, ?, 'MOD', 'Dados', ?, 'antes', 'depois')""",
        (
            (processed_id, planilha_id, f"A{index}")
            for index in range(1, CHANGE_COUNT + 1)
        ),
    )
    connection.commit()


def query_plans(connection: sqlite3.Connection, planilha_id: int) -> dict[str, list[str]]:
    queries = {
        "checkpoint": "SELECT versao_id, versao_numero FROM checkpoint WHERE planilha_id = ?",
        "versions_for_report": "SELECT * FROM versao_processada WHERE planilha_id = ? ORDER BY id",
        "changes_for_report": (
            "SELECT a.id FROM alteracao a JOIN versao_processada v "
            "ON v.id = a.versao_processada_id WHERE a.planilha_id = ? ORDER BY a.id"
        ),
        "last_execution": (
            "SELECT inicio FROM execucao_auditoria "
            "WHERE planilha_id = ? ORDER BY id DESC LIMIT 1"
        ),
    }
    return {
        name: [row["detail"] for row in connection.execute(f"EXPLAIN QUERY PLAN {query}", (planilha_id,))]
        for name, query in queries.items()
    }


def run() -> dict[str, object]:
    with TemporaryDirectory(prefix="auditoria-performance-") as temporary:
        directory = Path(temporary)
        workbook_path = directory / "version.xlsx"
        database_path = directory / "audit.db"
        reports = directory / "reports"
        create_workbook(workbook_path)

        database = Database(database_path)
        database.initialize()
        source = SyntheticSource(workbook_path, versions(VERSION_COUNT))
        service = AuditService(database, source)

        initial_result, initial_seconds, initial_peak = measured(
            lambda: service.audit(SPREADSHEET)
        )
        initial_downloads = source.downloads
        planilha_id = spreadsheet_id(database.connection)
        size_after_initial = database_path.stat().st_size

        source.versions.extend(versions(VERSION_COUNT + 2)[VERSION_COUNT:])
        source.downloads = 0
        incremental_result, incremental_seconds, incremental_peak = measured(
            lambda: service.audit(SPREADSHEET)
        )

        _, spreadsheets_seconds, spreadsheets_peak = measured(
            lambda: seed_many_spreadsheets(database.connection)
        )
        _, changes_seconds, changes_peak = measured(
            lambda: seed_changes(database.connection, planilha_id)
        )
        size_before_report = database_path.stat().st_size

        report_path, report_seconds, report_peak = measured(
            lambda: ReportService(database.connection, reports).generate(planilha_id)
        )
        integrity = database.connection.execute("PRAGMA integrity_check").fetchone()[0]
        foreign_keys = database.connection.execute("PRAGMA foreign_key_check").fetchall()
        plans = query_plans(database.connection, planilha_id)
        counts = {
            table: database.connection.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
            for table in ("planilha", "versao_processada", "alteracao", "execucao_auditoria")
        }
        database.close()

        return {
            "volumes": {
                "cells_per_version": 100,
                "initial_versions": VERSION_COUNT,
                "incremental_new_versions": 2,
                "changes": CHANGE_COUNT,
                "spreadsheets": SPREADSHEET_COUNT,
            },
            "initial_audit": {
                "seconds": initial_seconds,
                "seconds_per_comparison": initial_seconds / initial_result.processed_versions,
                "peak_python_bytes": initial_peak,
                "downloads": initial_downloads,
                "processed_versions": initial_result.processed_versions,
                "database_bytes": size_after_initial,
            },
            "incremental_audit": {
                "seconds": incremental_seconds,
                "peak_python_bytes": incremental_peak,
                "downloads": source.downloads,
                "processed_versions": incremental_result.processed_versions,
            },
            "many_spreadsheets_insert": {
                "seconds": spreadsheets_seconds,
                "peak_python_bytes": spreadsheets_peak,
            },
            "many_changes_insert": {
                "seconds": changes_seconds,
                "peak_python_bytes": changes_peak,
                "database_bytes": size_before_report,
            },
            "report": {
                "seconds": report_seconds,
                "peak_python_bytes": report_peak,
                "xlsx_bytes": report_path.stat().st_size,
            },
            "row_counts": counts,
            "query_plans": plans,
            "integrity_check": integrity,
            "foreign_key_violations": len(foreign_keys),
            "temporary_directory_removed_on_exit": True,
        }


if __name__ == "__main__":
    print(json.dumps(run(), indent=2, sort_keys=True))
