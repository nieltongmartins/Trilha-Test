"""Teste mínimo de Tk + worker daemon + BrowserSharePointSource."""

import argparse
import logging
from pathlib import Path
import threading
import tkinter as tk

from app.sources.sharepoint import BrowserSharePointSource


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(threadName)s %(message)s",
)
logger = logging.getLogger("sharepoint-source-tk")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("site_url", help="URL HTTPS do site SharePoint")
    parser.add_argument(
        "--scope",
        action="append",
        default=[],
        help="Escopo relativo; pode ser informado mais de uma vez",
    )
    parser.add_argument("--temp-directory", type=Path)
    args = parser.parse_args()
    if not args.site_url.lower().startswith("https://"):
        parser.error("site_url deve usar HTTPS")

    root = tk.Tk()
    root.title("Teste mínimo - SharePointSource + Tk")
    source_holder: list[BrowserSharePointSource] = []

    def worker() -> None:
        logger.info("Abrindo BrowserSharePointSource")
        source = BrowserSharePointSource.open_edge(
            args.site_url,
            tuple(args.scope),
            temp_directory=args.temp_directory,
        )
        source_holder.append(source)
        logger.info("Edge e site_url abertos; aguardando autenticação")
        source.wait_until_authenticated()
        logger.info("SharePoint autenticado; o mainloop deve permanecer aberto")

    threading.Thread(target=worker, daemon=True).start()
    try:
        root.mainloop()
    finally:
        if source_holder:
            source_holder[0].close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
