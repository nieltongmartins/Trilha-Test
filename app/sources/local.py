"""Fonte local somente leitura para desenvolvimento e testes controlados."""

from __future__ import annotations

from pathlib import Path
from typing import Mapping, Sequence

from app.sources.base import SpreadsheetInfo, VersionInfo


class LocalSource:
    """Expõe arquivos históricos previamente preparados, sem alterá-los.

    A sequência recebida é a ordem confiável das versões. Isso evita ordenar
    números de exibição como texto e mantém o contrato compatível com uma
    futura fonte que use a ordem oficial da API.
    """

    def __init__(
        self,
        spreadsheets: Sequence[SpreadsheetInfo],
        versions: Mapping[
            tuple[str, str, str], Sequence[tuple[VersionInfo, str | Path]]
        ],
    ) -> None:
        self._spreadsheets = tuple(spreadsheets)
        self._versions = {
            identity: tuple((info, Path(path)) for info, path in items)
            for identity, items in versions.items()
        }

    @staticmethod
    def _identity(spreadsheet: SpreadsheetInfo) -> tuple[str, str, str]:
        return spreadsheet.site_id, spreadsheet.drive_id, spreadsheet.drive_item_id

    def list_spreadsheets(self) -> tuple[SpreadsheetInfo, ...]:
        return self._spreadsheets

    def list_versions(
        self, spreadsheet: SpreadsheetInfo
    ) -> tuple[VersionInfo, ...]:
        return tuple(info for info, _ in self._versions.get(self._identity(spreadsheet), ()))

    def get_version(
        self, spreadsheet: SpreadsheetInfo, version: VersionInfo
    ) -> Path:
        for info, path in self._versions.get(self._identity(spreadsheet), ()):
            if info.id == version.id:
                if not path.is_file():
                    raise FileNotFoundError(f"Arquivo da versão {version.number} não encontrado")
                return path
        raise KeyError(f"Versão desconhecida: {version.id}")
