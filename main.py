"""Ponto de entrada da aplicação desktop."""

import os
import multiprocessing
import sys
import tkinter as tk
import time

from app.config import Settings
from app.database import Database
from app.logging_config import configure_logging
from app.interface import AuditApplication
from app.sources.sharepoint import BrowserSharePointSource


def main(*, launch_ui: bool | None = None) -> int:
    """Inicializa o banco e, quando há ambiente gráfico, abre a interface."""
    startup_started = time.perf_counter()
    last_stage = startup_started

    def startup_log(logger, stage: str) -> None:
        nonlocal last_stage
        now = time.perf_counter()
        logger.info(
            "STARTUP %s etapa=%.3fs acumulado=%.3fs",
            stage, now - last_stage, now - startup_started,
        )
        last_stage = now

    settings = Settings.from_environment()
    settings.create_directories()
    logger = configure_logging(settings.log_path, settings.log_level)
    logger.info("Inicialização da aplicação iniciada")
    startup_log(logger, "settings")
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
            startup_log(logger, "database")
            logger.info("Aplicação inicializada sem interface gráfica")
            return 0

        root = tk.Tk()
        root.title("Trilha de Auditoria")
        root.minsize(760, 390)
        startup_log(logger, "janela_criada")
        logger.info("Janela Tk criada; nenhuma conexão SharePoint foi iniciada")
        database.initialize()
        startup_log(logger, "database")

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
        startup_log(logger, "interface_init")
        root.protocol("WM_DELETE_WINDOW", application.request_close)
        logger.info("Interface pronta; aguardando ação do usuário para conectar")
        update_idletasks = getattr(root, "update_idletasks", None)
        if callable(update_idletasks):
            update_idletasks()
        startup_log(logger, "janela_visivel")
        startup_log(logger, "mainloop")
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
    multiprocessing.freeze_support()
    raise SystemExit(main())
