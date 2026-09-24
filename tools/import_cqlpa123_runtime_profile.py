"""Importa medições reais fornecidas para a identidade técnica da CQLPA123.

O utilitário é deliberadamente explícito: não tenta reconhecer o workbook pelo
nome e, portanto, não pode contaminar outra planilha. O prefetch e o tamanho são
argumentos porque esses dois valores não constam da série de throughput recebida.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from app.database import Database
from app.runtime_profile import RuntimeProfileStore, RuntimeSample


OBSERVED = (
    (2, 9.04, 10.1),
    (3, 11.19, 12.7),
    (4, 12.71, 14.3),
    (5, 14.04, 16.8),
    (6, 13.85, 19.8),
    (7, 16.55, 19.7),
    (8, 14.12, 25.0),
)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("database", type=Path)
    parser.add_argument("workbook_identity", help="site_id|drive_id|drive_item_id")
    parser.add_argument("--avg-file-bytes", type=float, required=True)
    parser.add_argument("--prefetch-target", type=int, required=True)
    args = parser.parse_args()
    if len(args.workbook_identity.split("|")) != 3:
        parser.error("workbook_identity deve ser site_id|drive_id|drive_item_id")
    with Database(args.database) as database:
        database.initialize()
        store = RuntimeProfileStore(database.connection)
        for slots, throughput, read_xlsx in OBSERVED:
            versions = 20
            store.record(args.workbook_identity, RuntimeSample(
                slots=slots, prefetch_target=args.prefetch_target,
                versions_processed=versions,
                elapsed_seconds=versions * 60.0 / throughput,
                avg_read_xlsx=read_xlsx, avg_file_bytes=args.avg_file_bytes,
                run_kind="benchmark_run",
            ))
        profile = store.load(args.workbook_identity)
        print(profile)


if __name__ == "__main__":
    main()
