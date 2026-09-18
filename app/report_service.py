"""Gera representações Excel da trilha persistida no banco oficial."""

from __future__ import annotations

from pathlib import Path
import logging
import sqlite3

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from app.report_artifacts import ReportArtifactManager


logger = logging.getLogger("auditoria_excel.report")


class ReportService:
    """Consulta somente o SQLite e produz um relatório regenerável."""

    def __init__(
        self, connection: sqlite3.Connection, output_directory: str | Path
    ) -> None:
        self.connection = connection
        self.output_directory = Path(output_directory)

    def generate(self, spreadsheet_id: int) -> Path:
        logger.info("Geração de relatório iniciada planilha_id=%d", spreadsheet_id)
        spreadsheet = self.connection.execute(
            "SELECT * FROM planilha WHERE id = ?", (spreadsheet_id,)
        ).fetchone()
        if spreadsheet is None:
            raise ValueError("Planilha não encontrada no banco de auditoria")

        versions = self.connection.execute(
            """
            SELECT * FROM versao_processada
            WHERE planilha_id = ? ORDER BY id
            """,
            (spreadsheet_id,),
        ).fetchall()
        changes = self.connection.execute(
            """
            SELECT a.id, v.versao_anterior_numero, v.versao_atual_numero,
                   v.data_hora_versao, v.autor, v.comentario,
                   a.aba, a.endereco, a.tipo, a.valor_anterior, a.valor_novo
            FROM alteracao a
            JOIN versao_processada v ON v.id = a.versao_processada_id
            WHERE a.planilha_id = ? ORDER BY a.id
            """,
            (spreadsheet_id,),
        ).fetchall()
        last_execution = self.connection.execute(
            """
            SELECT inicio FROM execucao_auditoria
            WHERE planilha_id = ? ORDER BY id DESC LIMIT 1
            """,
            (spreadsheet_id,),
        ).fetchone()

        workbook = Workbook()
        summary = workbook.active
        summary.title = "RESUMO"
        identity = "/".join(
            (
                spreadsheet["site_id"],
                spreadsheet["drive_id"],
                spreadsheet["drive_item_id"],
            )
        )
        counts = {
            kind: sum(row["tipo"] == kind for row in changes)
            for kind in ("ADD", "MOD", "DEL")
        }
        summary_rows = [
            ("Planilha", spreadsheet["nome_atual"]),
            ("Identidade técnica", identity),
            ("DriveItem ID", spreadsheet["drive_item_id"]),
            (
                "Primeira versão",
                versions[0]["versao_anterior_numero"] if versions else None,
            ),
            (
                "Última versão",
                versions[-1]["versao_atual_numero"] if versions else None,
            ),
            ("Última execução", last_execution["inicio"] if last_execution else None),
            ("Versões processadas", len(versions)),
            ("Total de alterações", len(changes)),
            ("ADD", counts["ADD"]),
            ("MOD", counts["MOD"]),
            ("DEL", counts["DEL"]),
        ]
        summary.append(("Campo", "Valor"))
        for row in summary_rows:
            summary.append(row)

        version_sheet = workbook.create_sheet("VERSOES")
        version_headers = (
            "Versão anterior",
            "Versão atual",
            "Data/hora",
            "Autor",
            "Comentário",
            "Status",
            "Alterações",
        )
        version_sheet.append(version_headers)
        for version in versions:
            version_sheet.append(
                tuple(
                    version[key]
                    for key in (
                        "versao_anterior_numero",
                        "versao_atual_numero",
                        "data_hora_versao",
                        "autor",
                        "comentario",
                        "status",
                        "quantidade_alteracoes",
                    )
                )
            )

        trail = workbook.create_sheet("TRILHA")
        trail_headers = (
            "ID",
            "Versão anterior",
            "Versão atual",
            "Data/hora",
            "Autor",
            "Comentário",
            "Aba",
            "Célula",
            "Tipo",
            "Valor anterior",
            "Valor novo",
        )
        trail.append(trail_headers)
        for change in changes:
            trail.append(tuple(change))

        for sheet in workbook.worksheets:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
            for column in sheet.columns:
                width = min(max(len(str(cell.value or "")) for cell in column) + 2, 50)
                sheet.column_dimensions[column[0].column_letter].width = width

        self.output_directory.mkdir(parents=True, exist_ok=True)
        artifacts = ReportArtifactManager(self.connection, self.output_directory)
        output = artifacts.canonical_path(artifacts.identity(spreadsheet_id))
        workbook.save(output)
        logger.info(
            "Relatório gerado planilha_id=%d planilha=%s versoes=%d alteracoes=%d arquivo=%s",
            spreadsheet_id,
            spreadsheet["nome_atual"],
            len(versions),
            len(changes),
            output,
        )
        return output
