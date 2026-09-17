"""Interface desktop responsiva para configurar, conectar e auditar."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from pathlib import Path
import logging
import queue
import subprocess
import sys
import threading
from tkinter import messagebox, ttk
import tkinter as tk

from app.audit_service import AuditResult, AuditService
from app.database import Database
from app.report_service import ReportService
from app.sources.base import SpreadsheetInfo, VersionSource


logger = logging.getLogger("auditoria_excel.interface")
SourceFactory = Callable[[str, tuple[str, ...]], VersionSource]
ConfigurationSaver = Callable[[str, tuple[str, ...]], None]


class AuditApplication(ttk.Frame):
    """Tela operacional; todo trabalho demorado ocorre fora da thread Tk."""

    def __init__(
        self,
        master: tk.Misc,
        database: Database,
        source: VersionSource | None,
        reports_directory: str | Path,
        *,
        site_url: str = "",
        scope_paths: Sequence[str] = (),
        connect_source: SourceFactory | None = None,
        save_configuration: ConfigurationSaver | None = None,
    ) -> None:
        super().__init__(master, padding=12)
        self.database = database
        self.source = source
        self.connect_source = connect_source
        self.save_configuration = save_configuration
        self.reports_directory = Path(reports_directory)
        self.spreadsheets: list[SpreadsheetInfo] = []
        self.last_report: Path | None = None
        self._busy = False
        self._work_results: queue.SimpleQueue[tuple[bool, object]] = (
            queue.SimpleQueue()
        )
        self.site_url = tk.StringVar(value=site_url)
        self.scope_paths = tk.StringVar(value=";".join(scope_paths))
        # Preserve the public attribute used by installations upgraded from the
        # first folder-field implementation. It now backs the read-only selector.
        self.folder_names = tk.StringVar(value="")
        self.folder_paths: list[str] = []
        self.status = tk.StringVar(
            value=(
                "Conectado. Atualize a lista."
                if source is not None
                else "Desconectado. Confira a configuração e clique em Conectar."
            )
        )
        self.details = tk.StringVar(value="Selecione uma planilha.")
        self._build()

    def _build(self) -> None:
        self.grid(sticky="nsew")
        self.columnconfigure(0, weight=1)
        ttk.Label(
            self, text="Auditor de Planilhas", font=("TkDefaultFont", 14, "bold")
        ).grid(sticky="w")

        config = ttk.LabelFrame(self, text="Conexão SharePoint", padding=8)
        config.grid(row=1, column=0, sticky="ew", pady=(8, 4))
        config.columnconfigure(1, weight=1)
        ttk.Label(config, text="URL do site:").grid(row=0, column=0, sticky="w")
        ttk.Entry(config, textvariable=self.site_url).grid(
            row=0, column=1, sticky="ew", padx=(8, 0)
        )
        ttk.Label(config, text="Escopo(s):").grid(
            row=1, column=0, sticky="w", pady=(6, 0)
        )
        ttk.Entry(config, textvariable=self.scope_paths).grid(
            row=1, column=1, sticky="ew", padx=(8, 0), pady=(6, 0)
        )
        ttk.Label(config, text="Pasta(s):").grid(
            row=2, column=0, sticky="w", pady=(6, 0)
        )
        self.folder_selector = ttk.Combobox(
            config, textvariable=self.folder_names, state="readonly"
        )
        self.folder_selector.grid(
            row=2, column=1, sticky="ew", padx=(8, 0), pady=(6, 0)
        )
        self.copy_folder_button = ttk.Button(
            config, text="Copiar para escopo", command=self.copy_folder_to_scope
        )
        self.copy_folder_button.grid(row=2, column=2, padx=(6, 0), pady=(6, 0))
        ttk.Label(
            config,
            text=(
                "Após conectar, selecione uma pasta e copie seu caminho para o escopo."
            ),
        ).grid(
            row=3, column=1, sticky="w"
        )
        self.connect_button = ttk.Button(config, text="Conectar", command=self.connect)
        self.connect_button.grid(row=4, column=1, sticky="e", pady=(8, 0))

        self.selector = ttk.Combobox(self, state="readonly", width=70)
        self.selector.grid(row=2, column=0, sticky="ew", pady=8)
        self.selector.bind("<<ComboboxSelected>>", lambda _event: self.show_status())
        ttk.Label(self, textvariable=self.details).grid(row=3, column=0, sticky="w")
        buttons = ttk.Frame(self)
        buttons.grid(row=4, column=0, sticky="w", pady=10)
        self.refresh_button = ttk.Button(
            buttons, text="Atualizar lista", command=self.refresh
        )
        self.refresh_button.pack(side="left", padx=(0, 6))
        self.audit_button = ttk.Button(
            buttons, text="Auditar histórico", command=self.audit
        )
        self.audit_button.pack(side="left", padx=(0, 6))
        self.report_button = ttk.Button(
            buttons, text="Gerar relatório", command=self.generate_report
        )
        self.report_button.pack(side="left", padx=6)
        ttk.Button(buttons, text="Abrir relatório", command=self.open_report).pack(
            side="left", padx=6
        )
        ttk.Label(self, textvariable=self.status, wraplength=720).grid(
            row=5, column=0, sticky="w"
        )
        self._set_action_state()

    def _configured_values(self) -> tuple[str, tuple[str, ...]]:
        site_url = self.site_url.get().strip().rstrip("/")
        scopes = tuple(
            dict.fromkeys(
                part.strip()
                for part in self.scope_paths.get().split(";")
                if part.strip()
            )
        )
        folders = tuple(
            dict.fromkeys(
                part.strip().strip("/")
                for part in self.folder_names.get().split(";")
                if part.strip().strip("/")
            )
        )
        if folders:
            scopes = tuple(
                f"{scope.rstrip('/')}/{folder}" for scope in scopes for folder in folders
            )
        return site_url, scopes

    def connect(self) -> None:
        if self.connect_source is None:
            self.status.set("Conexão SharePoint não está disponível.")
            return
        connect_source = self.connect_source
        site_url, scopes = self._configured_values()
        try:
            if self.save_configuration is not None:
                self.save_configuration(site_url, scopes)
        except (OSError, ValueError) as error:
            self.status.set(f"Configuração inválida: {error}")
            return
        self._start_work(
            "Conectando ao SharePoint... Conclua o login/MFA no Edge.",
            lambda: connect_source(site_url, scopes),
            self._connected,
        )

    def _connected(self, source: VersionSource) -> None:
        previous = self.source
        self.source = source
        if previous is not None and previous is not source:
            try:
                previous.close()  # type: ignore[attr-defined]
            except (AttributeError, RuntimeError):
                logger.warning(
                    "Falha ao fechar sessão SharePoint anterior", exc_info=True
                )
        self.status.set("Conectado ao SharePoint. Clique em Atualizar lista.")
        self._set_action_state()
        list_folders = getattr(source, "list_folders", None)
        if callable(list_folders):
            self._start_work(
                "Carregando pastas do SharePoint...",
                lambda: list(list_folders()),
                self._folders_loaded,
            )

    def _folders_loaded(self, folders: list[tuple[str, str]]) -> None:
        self.folder_paths = [path for _name, path in folders]
        self.folder_selector["values"] = [name for name, _path in folders]
        if folders:
            self.folder_selector.current(0)
            self.status.set("Pastas carregadas. Selecione uma ou atualize a lista.")
        else:
            self.status.set(
                "Nenhuma subpasta encontrada; o escopo atual pode ser usado."
            )

    def copy_folder_to_scope(self) -> None:
        index = self.folder_selector.current()
        if index < 0 or index >= len(self.folder_paths):
            self.status.set("Selecione uma pasta para copiar para o escopo.")
            return
        path = self.folder_paths[index]
        self.scope_paths.set(path)
        if self.source is not None:
            set_scope_paths = getattr(self.source, "set_scope_paths", None)
            if callable(set_scope_paths):
                set_scope_paths((path,))
        if self.save_configuration is not None:
            self.save_configuration(self.site_url.get().strip().rstrip("/"), (path,))
        self.status.set("Escopo atualizado. Clique em Atualizar lista.")

    def refresh(self) -> None:
        if self.source is None:
            self.status.set("Conecte ao SharePoint antes de atualizar a lista.")
            return
        source = self.source
        self._start_work(
            "Consultando planilhas no SharePoint...",
            lambda: list(source.list_spreadsheets()),
            self._refresh_finished,
        )

    def _refresh_finished(self, spreadsheets: list[SpreadsheetInfo]) -> None:
        self.spreadsheets = spreadsheets
        self.selector["values"] = [
            f"{item.name} — {item.path or item.drive_item_id}"
            for item in self.spreadsheets
        ]
        if self.spreadsheets:
            self.selector.current(0)
            self.show_status()
        else:
            self.status.set(
                "Nenhuma planilha .xlsx encontrada nos escopos configurados."
            )

    def _selected(self) -> SpreadsheetInfo:
        index = self.selector.current()
        if index < 0:
            raise ValueError("Selecione uma planilha")
        return self.spreadsheets[index]

    def _database_row(self, spreadsheet: SpreadsheetInfo):
        return self.database.connection.execute(
            """SELECT p.id, c.versao_numero FROM planilha p
               LEFT JOIN checkpoint c ON c.planilha_id = p.id
               WHERE p.site_id=? AND p.drive_id=? AND p.drive_item_id=?""",
            (spreadsheet.site_id, spreadsheet.drive_id, spreadsheet.drive_item_id),
        ).fetchone()

    def show_status(self) -> None:
        try:
            spreadsheet = self._selected()
        except ValueError as error:
            self.status.set(str(error))
            return
        self._start_work(
            "Consultando versões...",
            lambda: self._spreadsheet_status(spreadsheet),
            self._show_status_finished,
        )

    def _spreadsheet_status(self, spreadsheet: SpreadsheetInfo):
        if self.source is None:
            raise RuntimeError("Conecte ao SharePoint antes de consultar versões.")
        row = self._database_row(spreadsheet)
        checkpoint = row["versao_numero"] if row else None
        versions = list(self.source.list_versions(spreadsheet))
        latest = versions[-1].number if versions else "—"
        ids = [version.id for version in versions]
        pending = max(len(versions) - 1, 0)
        if row and checkpoint:
            checkpoint_row = self.database.connection.execute(
                "SELECT versao_id FROM checkpoint WHERE planilha_id=?", (row["id"],)
            ).fetchone()
            if checkpoint_row and checkpoint_row["versao_id"] in ids:
                pending = max(
                    len(versions) - ids.index(checkpoint_row["versao_id"]) - 1, 0
                )
        return checkpoint, latest, pending

    def _show_status_finished(self, result: tuple[str | None, str, int]) -> None:
        checkpoint, latest, pending = result
        self.details.set(
            f"Última auditada: {checkpoint or '—'} | "
            f"Última disponível: {latest} | Pendentes: {pending}"
        )
        self.audit_button.configure(
            text="Continuar auditoria" if checkpoint else "Auditar histórico"
        )
        self.status.set("Pronto.")

    def audit(self) -> None:
        try:
            spreadsheet = self._selected()
        except ValueError as error:
            self.status.set(str(error))
            return
        if self.source is None:
            self.status.set("Conecte ao SharePoint antes de auditar.")
            return
        source = self.source
        self._start_work(
            "Auditoria em andamento...",
            lambda: AuditService(self.database, source).audit(spreadsheet),
            self._audit_finished,
        )

    def _audit_finished(self, result: AuditResult) -> None:
        self.status.set(
            f"{result.status.value}: {result.processed_versions} versões, "
            f"{result.changes} alterações."
        )
        self.show_status()

    def generate_report(self) -> None:
        try:
            spreadsheet = self._selected()
        except ValueError as error:
            self.status.set(str(error))
            return
        row = self._database_row(spreadsheet)
        if row is None:
            messagebox.showinfo(
                "Relatório", "Audite a planilha antes de gerar o relatório."
            )
            return
        self._start_work(
            "Gerando relatório...",
            lambda: ReportService(
                self.database.connection, self.reports_directory
            ).generate(row["id"]),
            self._report_finished,
        )

    def _report_finished(self, report: Path) -> None:
        self.last_report = report
        self.status.set(f"Relatório gerado: {self.last_report}")

    def _start_work(
        self,
        message: str,
        operation: Callable[[], object],
        finished: Callable,
    ) -> None:
        if self._busy:
            self.status.set("Aguarde a operação atual terminar.")
            return
        self._busy = True
        self.status.set(message)
        self._set_action_state()

        def worker() -> None:
            try:
                result = operation()
            except Exception as error:
                logger.warning(
                    "Operação da interface falhou tipo_erro=%s erro=%s",
                    type(error).__name__,
                    error,
                    exc_info=True,
                )
                self._work_results.put((False, error))
            else:
                self._work_results.put((True, result))

        # Tk, including ``after``, is only accessed by the main thread.  The
        # worker communicates exclusively through this queue.
        self.after(50, self._poll_work_result, finished)
        threading.Thread(target=worker, daemon=True).start()

    def _poll_work_result(self, finished: Callable) -> None:
        try:
            succeeded, result = self._work_results.get_nowait()
        except queue.Empty:
            self.after(50, self._poll_work_result, finished)
            return
        if succeeded:
            self._work_finished(finished, result)
        else:
            assert isinstance(result, Exception)
            self._work_failed(result)

    def _work_finished(self, finished: Callable, result: object) -> None:
        self._busy = False
        self._set_action_state()
        finished(result)

    def _work_failed(self, error: Exception) -> None:
        self._busy = False
        self._set_action_state()
        self.status.set(f"Falha na operação: {error}")

    def _set_action_state(self) -> None:
        state = "disabled" if self._busy else "normal"
        connected_state = state if self.source is not None else "disabled"
        for name in ("connect_button",):
            if hasattr(self, name):
                getattr(self, name).configure(state=state)
        if hasattr(self, "copy_folder_button"):
            self.copy_folder_button.configure(state=connected_state)
        for name in ("refresh_button", "audit_button", "report_button"):
            if hasattr(self, name):
                getattr(self, name).configure(state=connected_state)

    def close_source(self) -> None:
        if self.source is not None:
            close = getattr(self.source, "close", None)
            if close is not None:
                close()
            self.source = None

    def open_report(self) -> None:
        if self.last_report is None or not self.last_report.exists():
            messagebox.showinfo("Relatório", "Gere o relatório primeiro.")
            return
        if sys.platform == "win32":
            import os

            os.startfile(self.last_report)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(("open", str(self.last_report)))
        else:
            subprocess.Popen(("xdg-open", str(self.last_report)))
