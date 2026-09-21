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
    application = None
    try:
        if launch_ui is None:
            launch_ui = sys.platform == "win32" or bool(os.environ.get("DISPLAY"))
        if not launch_ui:
            logger.info(
                "Abertura/configuração do banco iniciada arquivo=%s",
                settings.database_path,
            )
            database.initialize()
            logger.info("Aplicação inicializada sem interface gráfica")
            return 0

        root = tk.Tk()
        root.title("Trilha de Auditoria")
        root.minsize(760, 390)
        logger.info("Janela Tk criada; nenhuma conexão SharePoint foi iniciada")
        database.initialize()

        def connect_source(site_url: str, scopes: tuple[str, ...]):
            source = BrowserSharePointSource.open_edge(
                site_url, scopes, temp_directory=settings.temp_directory
            )
            try:
                source.wait_until_authenticated()
            except Exception:
                source.close()
                raise
            return source

        application = AuditApplication(
            root,
            database,
            None,
            settings.reports_directory,
            backups_directory=settings.backups_directory,
            site_url=settings.sharepoint_site_url or "",
            scope_paths=settings.sharepoint_scope_paths,
            connect_source=connect_source,
            save_configuration=settings.save_browser_sharepoint,
        )
        root.protocol("WM_DELETE_WINDOW", application.request_close)
        logger.info("Interface pronta; aguardando ação do usuário para conectar")
        root.mainloop()
    except Exception:
        logger.critical("Erro inesperado encerrou a aplicação", exc_info=True)
        raise
    finally:
        if application is not None:
            application.close_source()
        database.close()
        logger.info("Aplicação encerrada")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
