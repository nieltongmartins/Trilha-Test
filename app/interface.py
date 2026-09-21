"""Interface desktop responsiva para configurar, conectar e auditar."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from pathlib import Path
import logging
import queue
import subprocess
import sys
import threading
import time
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

from app.audit_service import AuditResult, AuditService
from app.audit_storage import AuditStorageManager, RestoreConflictError
from app.database import Database
from app.models import AuditExecutionStatus
from app.report_service import ReportService
from app.report_artifacts import ReportArtifactManager
from app.sources.base import SpreadsheetInfo, VersionInfo, VersionSource
from app.spreadsheet_comparator import ComparatorFrame


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
        backups_directory: str | Path = "data/backups",
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
        self.storage = AuditStorageManager(database, backups_directory)
        self.report_artifacts = ReportArtifactManager(
            getattr(database, "connection", database), self.reports_directory
        )
        self.spreadsheets: list[SpreadsheetInfo] = []
        # Cache curto da enumeração de versões. A listagem pode levar minutos em
        # arquivos com dezenas de milhares de versões; reutilizá-la evita repetir
        # a mesma consulta ao clicar em "Continuar auditoria".
        self._version_cache: dict[str, tuple[float, tuple[VersionInfo, ...]]] = {}
        self._version_cache_ttl = 300.0
        self.last_report: Path | None = None
        self._busy = False
        self._work_results: queue.SimpleQueue[tuple[bool, object]] = (
            queue.SimpleQueue()
        )
        self._progress_updates: queue.SimpleQueue[tuple[int, int]] = queue.SimpleQueue()
        self._version_scan_updates: queue.SimpleQueue[int] = queue.SimpleQueue()
        self._version_scan_started_at: float | None = None
        self._version_scan_active = False
        self._audit_started_at: float | None = None
        self._progress_completed = 0
        self._progress_total = 0
        self._hidden_clicks: list[float] = []
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
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1)
        self.grid(sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=0, column=0, sticky="nsew")
        audit_tab = ttk.Frame(self.notebook, padding=4)
        stored_tab = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(audit_tab, text="Auditoria")
        self.notebook.add(stored_tab, text="Auditorias armazenadas")
        self._audit_tab = audit_tab
        self._comparator_tab = None
        audit_tab.columnconfigure(0, weight=1)
        title_label = ttk.Label(
            audit_tab, text="Auditor de Planilhas", font=("TkDefaultFont", 14, "bold")
        )
        title_label.grid(sticky="w")
        # Gesto deliberado e invisível: cinco cliques no título existente.
        title_label.bind("<Button-1>", self._hidden_comparator_gesture)

        config = ttk.LabelFrame(audit_tab, text="Conexão SharePoint", padding=8)
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
        ttk.Label(
            config,
            text="Use o caminho completo da biblioteca/pasta; separe vários por ';'.",
        ).grid(
            row=2, column=1, sticky="w"
        )
        self.connect_button = ttk.Button(config, text="Conectar", command=self.connect)
        self.connect_button.grid(row=4, column=1, sticky="e", pady=(8, 0))

        self.selector = ttk.Combobox(audit_tab, state="readonly", width=70)
        self.selector.grid(row=2, column=0, sticky="ew", pady=8)
        self.selector.bind("<<ComboboxSelected>>", lambda _event: self.show_status())
        ttk.Label(audit_tab, textvariable=self.details).grid(row=3, column=0, sticky="w")
        buttons = ttk.Frame(audit_tab)
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
        ttk.Label(audit_tab, textvariable=self.status, wraplength=720).grid(
            row=5, column=0, sticky="w"
        )
        self.progress_value = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            audit_tab, variable=self.progress_value, maximum=100, mode="determinate"
        )
        self.progress_bar.grid(row=6, column=0, sticky="ew", pady=(10, 2))
        self.progress_text = tk.StringVar(
            value="Progresso da auditoria: aguardando | Tempo total: 00:00"
        )
        ttk.Label(audit_tab, textvariable=self.progress_text).grid(
            row=7, column=0, sticky="w"
        )
        self._build_stored_tab(stored_tab)
        self._set_action_state()
        self.refresh_stored()

    def _build_stored_tab(self, tab: ttk.Frame) -> None:
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)
        ttk.Label(tab, text="Auditorias armazenadas", font=("TkDefaultFont", 14, "bold")).grid(row=0, column=0, sticky="w")
        columns = ("nome", "caminho", "unique_id", "versao", "processadas", "alteracoes", "ultima", "checkpoint")
        self.stored_tree = ttk.Treeview(tab, columns=columns, show="headings", selectmode="browse")
        labels = ("Planilha", "Caminho", "UniqueId", "Última versão", "Versões", "Alterações", "Última auditoria", "Checkpoint")
        widths = (150, 210, 140, 95, 65, 70, 130, 95)
        for column, label, width in zip(columns, labels, widths):
            self.stored_tree.heading(column, text=label)
            self.stored_tree.column(column, width=width, minwidth=55)
        self.stored_tree.grid(row=1, column=0, sticky="nsew", pady=8)
        scroll = ttk.Scrollbar(tab, orient="vertical", command=self.stored_tree.yview)
        scroll.grid(row=1, column=1, sticky="ns", pady=8)
        self.stored_tree.configure(yscrollcommand=scroll.set)
        self.stored_tree.bind("<<TreeviewSelect>>", self._stored_selection_changed)
        actions = ttk.Frame(tab)
        actions.grid(row=2, column=0, columnspan=2, sticky="ew")
        definitions = (
            ("Atualizar", self.refresh_stored),
            ("Backup selecionado", self.backup_selected),
            ("Restaurar selecionado", self.restore_individual),
            ("Abrir relatório", self.open_stored_report),
            ("Excluir auditoria selecionada", self.delete_selected),
            ("Backup completo", self.backup_complete),
            ("Restaurar backup completo", self.restore_complete),
            ("Excluir todas as auditorias", self.delete_all_audits),
        )
        self.storage_buttons = []
        for index, (text, command) in enumerate(definitions):
            button = ttk.Button(actions, text=text, command=command)
            button.grid(row=index // 4, column=index % 4, sticky="ew", padx=(0, 5), pady=3)
            self.storage_buttons.append(button)
            if text == "Abrir relatório":
                self.open_stored_report_button = button
        self._stored_selection_changed()

    def _stored_selection_changed(self, _event: object = None) -> None:
        if hasattr(self, "open_stored_report_button"):
            state = "normal" if self.stored_tree.selection() and not self._busy else "disabled"
            self.open_stored_report_button.configure(state=state)

    def _hidden_comparator_gesture(self, _event: object = None) -> None:
        now = time.monotonic()
        self._hidden_clicks = [click for click in self._hidden_clicks if now - click <= 2.0]
        self._hidden_clicks.append(now)
        if len(self._hidden_clicks) >= 5:
            self._hidden_clicks.clear()
            self._show_comparator()

    def _show_comparator(self) -> None:
        if self._comparator_tab is None:
            self._comparator_tab = ComparatorFrame(
                self.notebook, on_close=self._hide_comparator
            )
            self.notebook.add(self._comparator_tab, text="Comparador de Planilhas")
        self.notebook.select(self._comparator_tab)

    def _hide_comparator(self) -> None:
        if self._comparator_tab is None:
            return
        self.notebook.select(self._audit_tab)
        self.notebook.forget(self._comparator_tab)
        self._comparator_tab.destroy()
        self._comparator_tab = None

    def refresh_stored(self) -> None:
        if not hasattr(self, "stored_tree"):
            return
        for item in self.stored_tree.get_children():
            self.stored_tree.delete(item)
        for audit in self.storage.list_audits():
            self.stored_tree.insert("", "end", iid=str(audit.id), values=(
                audit.name, audit.path or "—", audit.unique_id,
                audit.last_version or "—", audit.processed_versions, audit.changes,
                audit.last_audit or "—", audit.checkpoint or "—"))

    def _selected_stored(self) -> tuple[int, str]:
        selection = self.stored_tree.selection()
        if not selection:
            raise ValueError("Selecione uma auditoria armazenada.")
        item = selection[0]
        return int(item), str(self.stored_tree.item(item, "values")[0])

    def backup_selected(self) -> None:
        try:
            spreadsheet_id, _ = self._selected_stored()
        except ValueError as error:
            messagebox.showinfo("Auditorias armazenadas", str(error))
            return
        self._start_work("Criando backup individual...", lambda: self.storage.backup_individual(spreadsheet_id), self._storage_finished)

    def backup_complete(self) -> None:
        self._start_work("Criando backup completo...", self.storage.backup_full, self._storage_finished)

    def restore_individual(self) -> None:
        path = filedialog.askopenfilename(title="Selecionar backup individual", filetypes=(("Backup SQLite", "*.sqlite3"),))
        if not path:
            return
        self._start_work(
            "Validando e restaurando backup individual...",
            lambda: self._try_restore_individual(path),
            self._individual_restore_finished,
        )

    def _try_restore_individual(self, path: str) -> tuple[str, str]:
        try:
            self.storage.restore_individual(path)
        except RestoreConflictError:
            return "conflict", path
        return "restored", path

    def _individual_restore_finished(self, result: tuple[str, str]) -> None:
        state, path = result
        if state == "conflict":
            if not messagebox.askyesno("Conflito de restauração", "Já existe auditoria para esta planilha. Substituir integralmente pelos dados do backup? Nenhum merge será realizado."):
                self.status.set("Restauração cancelada; dados locais preservados.")
                return
            self._start_work("Restaurando backup individual...", lambda: self.storage.restore_individual(path, replace=True), self._storage_finished)
            return
        self._storage_finished(Path(path))

    def delete_selected(self) -> None:
        try:
            spreadsheet_id, name = self._selected_stored()
        except ValueError as error:
            messagebox.showinfo("Auditorias armazenadas", str(error))
            return
        if not messagebox.askyesno("Excluir auditoria local", f"Excluir permanentemente a auditoria local de '{name}'?\n\nO arquivo no SharePoint não será alterado."):
            return
        def operation() -> None:
            report = self.report_artifacts.locate(spreadsheet_id)
            self.report_artifacts.delete_with_database(
                [report] if report else [],
                lambda: self.storage.delete_individual(spreadsheet_id),
            )
        self._start_work("Excluindo auditoria local...", operation, self._storage_finished)

    def open_stored_report(self) -> None:
        try:
            spreadsheet_id, _ = self._selected_stored()
            report = self.report_artifacts.locate(spreadsheet_id)
        except ValueError as error:
            messagebox.showinfo("Auditorias armazenadas", str(error))
            return
        if report is None:
            messagebox.showinfo(
                "Auditorias armazenadas",
                "Não existe relatório gerado para esta auditoria.",
            )
            return
        try:
            self._open_file(report)
        except OSError as error:
            logger.exception("Falha ao abrir relatório armazenado arquivo=%s", report)
            messagebox.showerror("Auditorias armazenadas", f"Não foi possível abrir o relatório: {error}")

    def restore_complete(self) -> None:
        path = filedialog.askopenfilename(title="Selecionar backup completo", filetypes=(("Backup SQLite", "*.sqlite3"),))
        if not path or not messagebox.askyesno("Restaurar banco completo", "Validar e substituir TODO o estado local pelo backup selecionado? Um backup de segurança será criado antes da substituição."):
            return
        self._start_work("Restaurando backup completo...", lambda: self.storage.restore_full(path), self._storage_finished)

    def delete_all_audits(self) -> None:
        if not messagebox.askyesno("Excluir todas as auditorias", "Esta operação excluirá permanentemente todas as auditorias armazenadas localmente. Os arquivos do SharePoint não serão alterados.\n\nDeseja continuar?"):
            return
        identities = [self.report_artifacts.identity(audit.id) for audit in self.storage.list_audits()]
        reports = self.report_artifacts.controlled_paths(identities)
        backup = messagebox.askyesno("Backup de proteção", "Fazer backup completo antes de excluir?")
        def operation() -> object:
            backup_path = self.storage.backup_full() if backup else None
            self.report_artifacts.delete_with_database(reports, self.storage.delete_all)
            return backup_path
        self._start_work("Excluindo todas as auditorias locais...", operation, self._storage_finished)

    def _storage_finished(self, result: object) -> None:
        self.refresh_stored()
        self.status.set(f"Operação local concluída{f': {result}' if isinstance(result, Path) else '.'}")

    def _configured_values(self) -> tuple[str, tuple[str, ...]]:
        site_url = self.site_url.get().strip().rstrip("/")
        scopes = tuple(
            dict.fromkeys(
                part.strip()
                for part in self.scope_paths.get().split(";")
                if part.strip()
            )
        )
        return site_url, scopes

    def _folders_loaded(self, folders: list[tuple[str, str]]) -> None:
        self.folder_paths = [path for _name, path in folders]
        self.folder_selector["values"] = [name for name, _path in folders]
        if folders:
            self.folder_selector.current(0)

    def copy_folder_to_scope(self) -> None:
        index = self.folder_selector.current()
        if index < 0 or index >= len(self.folder_paths):
            self.status.set("Selecione uma pasta.")
            return
        scope = self.folder_paths[index]
        self.scope_paths.set(scope)
        if self.source is not None:
            setter = getattr(self.source, "set_scope_paths", None)
            if setter is not None:
                setter((scope,))
        self.status.set("Escopo atualizado. Clique em Atualizar lista.")

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
        self._version_cache.clear()
        if previous is not None and previous is not source:
            try:
                previous.close()  # type: ignore[attr-defined]
            except (AttributeError, RuntimeError):
                logger.warning(
                    "Falha ao fechar sessão SharePoint anterior", exc_info=True
                )
        self.status.set("Conectado ao SharePoint. Clique em Atualizar lista.")
        self._set_action_state()

    def refresh(self) -> None:
        if self.source is None:
            self.status.set("Conecte ao SharePoint antes de atualizar a lista.")
            return
        source = self.source
        self._version_cache.clear()
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

    @staticmethod
    def _version_cache_key(spreadsheet: SpreadsheetInfo) -> str:
        return f"{spreadsheet.site_id}|{spreadsheet.drive_id}|{spreadsheet.drive_item_id}"

    def _cached_versions(
        self, spreadsheet: SpreadsheetInfo
    ) -> tuple[VersionInfo, ...] | None:
        key = self._version_cache_key(spreadsheet)
        cached = self._version_cache.get(key)
        if cached is None:
            return None
        created_at, versions = cached
        if time.monotonic() - created_at > self._version_cache_ttl:
            self._version_cache.pop(key, None)
            return None
        return versions

    def _store_versions(
        self, spreadsheet: SpreadsheetInfo, versions: tuple[VersionInfo, ...]
    ) -> None:
        self._version_cache[self._version_cache_key(spreadsheet)] = (
            time.monotonic(),
            versions,
        )

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

        if self._cached_versions(spreadsheet) is None:
            self._begin_version_scan()

        self._start_work(
            "Consultando histórico de versões no SharePoint...",
            lambda: self._spreadsheet_status(spreadsheet),
            self._show_status_finished,
        )

    def _spreadsheet_status(self, spreadsheet: SpreadsheetInfo):
        if self.source is None:
            raise RuntimeError("Conecte ao SharePoint antes de consultar versões.")
        row = self._database_row(spreadsheet)
        checkpoint = row["versao_numero"] if row else None
        cached = self._cached_versions(spreadsheet)
        if cached is None:
            list_versions = getattr(self.source, "list_versions")
            try:
                versions = tuple(
                    list_versions(
                        spreadsheet,
                        progress_callback=self._version_scan_updates.put,
                    )
                )
            except TypeError:
                # Compatibilidade com fontes alternativas que ainda implementem
                # a assinatura antiga do protocolo VersionSource.
                versions = tuple(list_versions(spreadsheet))
            self._store_versions(spreadsheet, versions)
        else:
            versions = cached
            logger.info(
                "Lista de versões reutilizada do cache para status planilha=%s total=%d",
                spreadsheet.name,
                len(versions),
            )
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
        return checkpoint, latest, pending, len(versions)

    def _show_status_finished(
        self, result: tuple[str | None, str, int, int]
    ) -> None:
        checkpoint, latest, pending, total_versions = result
        self._end_version_scan(total_versions=total_versions)
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
        cached_versions = self._cached_versions(spreadsheet)
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")
        self._audit_started_at = time.monotonic()
        self._progress_completed = 0
        self._progress_total = 0
        self.progress_value.set(0)
        self.progress_text.set(
            "Progresso da auditoria: preparando | Estimativa: calculando | "
            "Tempo total: 00:00"
        )

        def report_progress(completed: int, total: int) -> None:
            self._progress_updates.put((completed, total))

        self._start_work(
            "Auditoria em andamento...",
            lambda: AuditService(
                self.database, source, progress_callback=report_progress
            ).audit(spreadsheet, versions=cached_versions),
            self._audit_finished,
        )

    def _audit_finished(self, result: AuditResult) -> None:
        if result.status is AuditExecutionStatus.FAILED:
            self._finish_progress(failed=True)
        else:
            self._finish_progress()
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
        self._poll_progress_updates()
        self._poll_version_scan_updates()
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
        if self._version_scan_active:
            self._end_version_scan(failed=True)
        if getattr(self, "_audit_started_at", None) is not None:
            elapsed = time.monotonic() - self._audit_started_at
            self.progress_text.set(
                f"Auditoria interrompida | Tempo total: {self._format_duration(elapsed)}"
            )
            self._audit_started_at = None
        self.status.set(f"Falha na operação: {error}")

    def _begin_version_scan(self) -> None:
        """Mostra atividade contínua enquanto o total de versões ainda é desconhecido."""
        self._version_scan_started_at = time.monotonic()
        self._version_scan_active = True
        while True:
            try:
                self._version_scan_updates.get_nowait()
            except queue.Empty:
                break
        self.progress_bar.stop()
        self.progress_bar.configure(mode="indeterminate")
        self.progress_value.set(0)
        self.progress_bar.start(12)
        self.progress_text.set(
            "Carregando histórico do SharePoint... 0 versões encontradas | "
            "Tempo: 00:00"
        )

    def _poll_version_scan_updates(self) -> None:
        if not self._version_scan_active:
            return
        latest_count: int | None = None
        while True:
            try:
                latest_count = self._version_scan_updates.get_nowait()
            except queue.Empty:
                break
        if latest_count is None:
            return
        elapsed = (
            time.monotonic() - self._version_scan_started_at
            if self._version_scan_started_at is not None
            else 0.0
        )
        self.progress_text.set(
            f"Carregando histórico do SharePoint... "
            f"{latest_count:,} versões encontradas | "
            f"Tempo: {self._format_duration(elapsed)}"
        )

    def _end_version_scan(
        self, *, total_versions: int | None = None, failed: bool = False
    ) -> None:
        if not self._version_scan_active:
            return
        elapsed = (
            time.monotonic() - self._version_scan_started_at
            if self._version_scan_started_at is not None
            else 0.0
        )
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")
        self.progress_value.set(0 if failed else 100)
        if failed:
            self.progress_text.set(
                "Carregamento do histórico interrompido | "
                f"Tempo: {self._format_duration(elapsed)}"
            )
        else:
            total_text = (
                f"{total_versions:,} versões carregadas"
                if total_versions is not None
                else "histórico carregado"
            )
            self.progress_text.set(
                f"Histórico do SharePoint: {total_text} | "
                f"Tempo: {self._format_duration(elapsed)}"
            )
        self._version_scan_active = False
        self._version_scan_started_at = None

    def _poll_progress_updates(self) -> None:
        if not hasattr(self, "_progress_updates"):
            return
        while True:
            try:
                completed, total = self._progress_updates.get_nowait()
            except queue.Empty:
                break
            self._update_progress(completed, total)
        if getattr(self, "_audit_started_at", None) is not None:
            self._update_progress(self._progress_completed, self._progress_total)

    def _update_progress(self, completed: int, total: int) -> None:
        if self._audit_started_at is None:
            return
        self._progress_completed = completed
        self._progress_total = total
        elapsed = time.monotonic() - self._audit_started_at
        percent = 0 if total <= 0 else completed / total * 100
        self.progress_value.set(percent)
        if completed > 0 and completed < total:
            remaining = elapsed / completed * (total - completed)
            estimate = self._format_duration(remaining)
        elif total > 0 and completed >= total:
            estimate = "00:00"
        else:
            estimate = "calculando"
        self.progress_text.set(
            f"Progresso da auditoria: {completed}/{total} ({percent:.0f}%) | "
            f"Estimativa: {estimate} | Tempo total: {self._format_duration(elapsed)}"
        )

    def _finish_progress(self, *, failed: bool = False) -> None:
        if self._audit_started_at is None:
            return
        elapsed = time.monotonic() - self._audit_started_at
        if failed:
            self.progress_text.set(
                "Auditoria interrompida | "
                f"Tempo total: {self._format_duration(elapsed)}"
            )
        else:
            self.progress_value.set(100)
            self.progress_text.set(
                f"Progresso da auditoria: concluída (100%) | "
                f"Tempo total: {self._format_duration(elapsed)}"
            )
        self._audit_started_at = None

    @staticmethod
    def _format_duration(seconds: float) -> str:
        total_seconds = max(0, round(seconds))
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        if hours:
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        return f"{minutes:02d}:{seconds:02d}"

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
        for button in getattr(self, "storage_buttons", ()):
            button.configure(state=state)
        self._stored_selection_changed()

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
        self._open_file(self.last_report)

    @staticmethod
    def _open_file(path: Path) -> None:
        if sys.platform == "win32":
            import os

            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(("open", str(path)))
        else:
            subprocess.Popen(("xdg-open", str(path)))
