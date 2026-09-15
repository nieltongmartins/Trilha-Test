"""Contrato pequeno e independente de tecnologia para fontes de versões."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Protocol, Sequence


@dataclass(frozen=True, slots=True)
class SpreadsheetInfo:
    """Identidade técnica e dados descritivos de uma planilha."""

    site_id: str
    drive_id: str
    drive_item_id: str
    name: str
    path: str | None = None


@dataclass(frozen=True, slots=True)
class VersionInfo:
    """Versão normalizada; a posição na fonte define sua ordem oficial."""

    id: str
    number: str
    modified_at: str | None = None
    author: str | None = None
    comment: str | None = None
    size: int | None = None


class VersionSource(Protocol):
    """Operações somente leitura necessárias ao núcleo da auditoria."""

    def list_spreadsheets(self) -> Sequence[SpreadsheetInfo]: ...

    def list_versions(self, spreadsheet: SpreadsheetInfo) -> Sequence[VersionInfo]: ...

    def get_version(
        self, spreadsheet: SpreadsheetInfo, version: VersionInfo
    ) -> Path: ...
