"""Ponto de entrada da aplicação desktop."""

import os
import sys
import threading
import tkinter as tk
import traceback

from app.config import Settings
from app.database import Database
from app.logging_config import configure_logging
from app.interface import AuditApplication
from app.sources.sharepoint import BrowserSharePointSource


def _window_state(root: tk.Misc) -> str:
    """Return a best-effort snapshot without letting diagnostics break Tk."""
    values: dict[str, object] = {}
    for name in ("winfo_exists", "winfo_viewable", "state"):
        try:
            values[name.removeprefix("winfo_")] = getattr(root, name)()
        except (AttributeError, tk.TclError) as error:
            values[name.removeprefix("winfo_")] = f"indisponível:{type(error).__name__}"
    return " ".join(f"{key}={value}" for key, value in values.items())


class WindowLifecycle:
    """Diagnose and control the lifetime of the application's root window."""

    def __init__(self, root: tk.Tk, logger) -> None:
        self.root = root
        self.logger = logger
        self.close_requested = False
        self.destroy_observed = False

        root.protocol("WM_DELETE_WINDOW", self.user_requested_close)
        for sequence in ("<Destroy>", "<Unmap>", "<Map>", "<Visibility>"):
            root.bind(
                sequence,
                lambda event, event_name=sequence: self._window_event(
                    event, event_name
                ),
                add="+",
            )
        root.report_callback_exception = self.report_callback_exception

    def _log(self, event: str, *, origin: str, reason: str) -> None:
        self.logger.info(
            "Ciclo de vida Tk evento=%s thread=%s origem=%s motivo=%s %s",
            event,
            threading.current_thread().name,
            origin,
            reason,
            _window_state(self.root),
        )

    def _window_event(self, event, event_name: str) -> None:
        # Bindings on a toplevel also see child events. Only the root matters here.
        if getattr(event, "widget", None) is not self.root:
            return
        if event_name == "<Destroy>":
            self.destroy_observed = True
        self._log(event_name, origin="evento-tk", reason="evento da janela principal")

    def report_callback_exception(self, exc_type, exc_value, exc_traceback) -> None:
        formatted = "".join(
            traceback.format_exception(exc_type, exc_value, exc_traceback)
        )
        self.logger.error(
            "Exceção em callback Tkinter thread=%s origem=report_callback_exception "
            "motivo=callback lançou exceção %s\n%s",
            threading.current_thread().name,
            _window_state(self.root),
            formatted,
        )

    def user_requested_close(self) -> None:
        self.close_requested = True
        self._log(
            "WM_DELETE_WINDOW",
            origin="usuário",
            reason="WM_DELETE_WINDOW solicitado pelo usuário",
        )
        self.destroy(origin="handler-WM_DELETE_WINDOW", reason="fechamento explícito")

    def destroy(self, *, origin: str, reason: str) -> None:
        self._log("destroy solicitado", origin=origin, reason=reason)
        self.root.destroy()


def main(*, launch_ui: bool | None = None) -> int:
    """Inicializa o banco e, quando há ambiente gráfico, abre a interface."""
    settings = Settings.from_environment()
    settings.create_directories()
    logger = configure_logging(settings.log_path, settings.log_level)
    logger.info("Inicialização da aplicação iniciada")
    database = Database(settings.database_path)
    application = None
    lifecycle = None
    previous_threading_excepthook = threading.excepthook

    def diagnostic_threading_excepthook(args: threading.ExceptHookArgs) -> None:
        logger.error(
            "Exceção não capturada em worker nome=%s ident=%s daemon=%s "
            "tipo_erro=%s erro=%r",
            args.thread.name if args.thread is not None else "desconhecida",
            args.thread.ident if args.thread is not None else None,
            args.thread.daemon if args.thread is not None else None,
            args.exc_type.__name__,
            args.exc_value,
            exc_info=(args.exc_type, args.exc_value, args.exc_traceback),
        )
        previous_threading_excepthook(args)

    threading.excepthook = diagnostic_threading_excepthook
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
        lifecycle = WindowLifecycle(root, logger)
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
            site_url=settings.sharepoint_site_url or "",
            scope_paths=settings.sharepoint_scope_paths,
            connect_source=connect_source,
            save_configuration=settings.save_browser_sharepoint,
        )
        logger.info("Interface pronta; aguardando ação do usuário para conectar")
        logger.info("Entrando em root.mainloop")
        lifecycle._log(
            "antes de mainloop", origin="main.py", reason="início do loop Tk"
        )
        try:
            root.mainloop()
            if lifecycle.close_requested:
                logger.info("root.mainloop retornou após solicitação de fechamento")
            else:
                logger.warning(
                    "root.mainloop RETORNOU NORMALMENTE sem solicitação de fechamento"
                )
        except BaseException as error:
            logger.exception(
                "root.mainloop terminou por BaseException tipo_erro=%s erro=%r",
                type(error).__name__,
                error,
            )
            raise
        lifecycle._log(
            "depois de mainloop",
            origin="main.py",
            reason=(
                "fechamento explícito solicitado"
                if lifecycle.close_requested
                else "mainloop retornou inesperadamente"
            ),
        )
    except Exception:
        logger.critical("Erro inesperado encerrou a aplicação", exc_info=True)
        raise
    finally:
        threading.excepthook = previous_threading_excepthook
        if lifecycle is not None:
            lifecycle._log(
                "finally",
                origin="main.py",
                reason="cleanup provocado pelo retorno de mainloop ou exceção",
            )
        if application is not None:
            application.close_source()
        database.close()
        logger.info("Aplicação encerrada")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
