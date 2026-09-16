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
    logger.info("Inicialização da aplicação iniciada")
    database = Database(settings.database_path)
    source = None
    try:
        logger.info("Abertura/configuração do banco iniciada arquivo=%s", settings.database_path)
        database.initialize()
        logger.info("Banco aberto e configurado arquivo=%s", settings.database_path)
        if launch_ui is None:
            launch_ui = sys.platform == "win32" or bool(os.environ.get("DISPLAY"))
        if not launch_ui:
            logger.info("Aplicação inicializada sem interface gráfica")
            return 0

        site_url, scopes = settings.require_browser_sharepoint()
        source = BrowserSharePointSource.open_edge(
            site_url, scopes, temp_directory=settings.temp_directory
        )
        root = tk.Tk()
        root.title("Auditor de Planilhas SharePoint")
        root.minsize(760, 220)
        AuditApplication(root, database, source, settings.reports_directory)
        root.mainloop()
    except Exception:
        logger.critical("Erro inesperado encerrou a aplicação", exc_info=True)
        raise
    finally:
        if source is not None:
            source.close()
        database.close()
        logger.info("Aplicação encerrada")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
