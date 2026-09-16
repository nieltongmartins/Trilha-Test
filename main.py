"""Ponto de entrada da aplicação desktop."""

import os
import sys
import tkinter as tk

from app.config import Settings
from app.database import Database
from app.logging_config import configure_logging
from app.interface import AuditApplication
from app.sources.sharepoint import BrowserSharePointSource


def main(*, launch_ui: bool | None = None) -> int:
    """Inicializa o banco e, quando há ambiente gráfico, abre a interface."""
    settings = Settings.from_environment()
    settings.create_directories()
    logger = configure_logging(settings.log_path, settings.log_level)

    database = Database(settings.database_path)
    database.initialize()

    logger.info(
        "Aplicação inicializada; banco disponível em %s", settings.database_path
    )
    if launch_ui is None:
        launch_ui = sys.platform == "win32" or bool(os.environ.get("DISPLAY"))
    if not launch_ui:
        database.close()
        return 0

    site_url, scopes = settings.require_browser_sharepoint()
    source = BrowserSharePointSource.open_edge(
        site_url, scopes, temp_directory=settings.temp_directory
    )
    try:
        root = tk.Tk()
        root.title("Auditor de Planilhas SharePoint")
        root.minsize(760, 220)
        AuditApplication(root, database, source, settings.reports_directory)
        root.mainloop()
    finally:
        source.close()
        database.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
