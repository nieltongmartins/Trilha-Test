"""Fontes de versões consumidas pelo serviço de auditoria."""

from app.sources.base import SpreadsheetInfo, VersionInfo, VersionSource
from app.sources.local import LocalSource
from app.sources.sharepoint import GraphReadError, SharePointSource

__all__ = [
    "GraphReadError", "LocalSource", "SharePointSource", "SpreadsheetInfo",
    "VersionInfo", "VersionSource",
]
