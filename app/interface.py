"""Interface desktop simples para operar a auditoria e gerar relatórios."""

from __future__ import annotations

from pathlib import Path
import subprocess
import sys
import threading
from tkinter import messagebox, ttk
import tkinter as tk

from app.audit_service import AuditService
from app.database import Database
from app.report_service import ReportService
from app.sources.base import SpreadsheetInfo, VersionSource


class AuditApplication(ttk.Frame):
    """Tela operacional deliberadamente pequena, sem estado fora do SQLite."""

    def __init__(
        self,
        master: tk.Misc,
        database: Database,
        source: VersionSource,
        reports_directory: str | Path,
    ) -> None:
        super().__init__(master, padding=12)
        self.database = database
        self.source = source
        self.reports_directory = Path(reports_directory)
        self.spreadsheets: list[SpreadsheetInfo] = []
        self.last_report: Path | None = None
        self.status = tk.StringVar(value="Carregando planilhas...")
        self.details = tk.StringVar(value="Selecione uma planilha.")
        self._build()
        self.refresh()

    def _build(self) -> None:
        self.grid(sticky="nsew")
        self.columnconfigure(0, weight=1)
        ttk.Label(
            self, text="Auditor de Planilhas", font=("TkDefaultFont", 14, "bold")
        ).grid(sticky="w")
        self.selector = ttk.Combobox(self, state="readonly", width=70)
        self.selector.grid(row=1, column=0, sticky="ew", pady=8)
        self.selector.bind("<<ComboboxSelected>>", lambda _event: self.show_status())
        ttk.Label(self, textvariable=self.details).grid(row=2, column=0, sticky="w")
        buttons = ttk.Frame(self)
        buttons.grid(row=3, column=0, sticky="w", pady=10)
        ttk.Button(buttons, text="Atualizar lista", command=self.refresh).pack(
            side="left", padx=(0, 6)
        )
        self.audit_button = ttk.Button(
            buttons, text="Auditar histórico", command=self.audit
        )
        self.audit_button.pack(side="left", padx=(0, 6))
        ttk.Button(buttons, text="Gerar relatório", command=self.generate_report).pack(
            side="left", padx=6
        )
        ttk.Button(buttons, text="Abrir relatório", command=self.open_report).pack(
            side="left", padx=6
        )
        ttk.Label(self, textvariable=self.status).grid(row=4, column=0, sticky="w")

    def refresh(self) -> None:
        try:
            self.spreadsheets = list(self.source.list_spreadsheets())
        except Exception as error:
            self.status.set(f"Falha ao listar planilhas: {error}")
            return
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
        spreadsheet = self._selected()
        row = self._database_row(spreadsheet)
        checkpoint = row["versao_numero"] if row else None
        try:
            versions = list(self.source.list_versions(spreadsheet))
            latest = versions[-1].number if versions else "—"
            ids = [version.id for version in versions]
            pending = (
                max(len(versions) - 1, 0)
                if not row
                else (
                    max(
                        len(versions)
                        - ids.index(
                            self.database.connection.execute(
                                "SELECT versao_id FROM checkpoint WHERE planilha_id=?",
                                (row["id"],),
                            ).fetchone()["versao_id"]
                        )
                        - 1,
                        0,
                    )
                    if checkpoint
                    else max(len(versions) - 1, 0)
                )
            )
        except (ValueError, TypeError, IndexError):
            latest, pending = "indisponível", "indisponível"
        self.details.set(
            f"Última auditada: {checkpoint or '—'} | Última disponível: {latest} | Pendentes: {pending}"
        )
        self.audit_button.configure(
            text="Continuar auditoria" if checkpoint else "Auditar histórico"
        )
        self.status.set("Pronto.")

    def audit(self) -> None:
        spreadsheet = self._selected()
        self.audit_button.configure(state="disabled")
        self.status.set("Auditoria em andamento...")
        threading.Thread(
            target=self._audit_worker, args=(spreadsheet,), daemon=True
        ).start()

    def _audit_worker(self, spreadsheet: SpreadsheetInfo) -> None:
        result = AuditService(self.database, self.source).audit(spreadsheet)
        self.after(0, self._audit_finished, result)

    def _audit_finished(self, result) -> None:
        self.audit_button.configure(state="normal")
        self.status.set(
            f"{result.status.value}: {result.processed_versions} versões, {result.changes} alterações."
        )
        self.show_status()

    def generate_report(self) -> None:
        spreadsheet = self._selected()
        row = self._database_row(spreadsheet)
        if row is None:
            messagebox.showinfo(
                "Relatório", "Audite a planilha antes de gerar o relatório."
            )
            return
        self.last_report = ReportService(
            self.database.connection, self.reports_directory
        ).generate(row["id"])
        self.status.set(f"Relatório gerado: {self.last_report}")

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
