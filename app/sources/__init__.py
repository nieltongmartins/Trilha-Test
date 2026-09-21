"""Fontes de versões consumidas pelo serviço de auditoria."""

from app.sources.base import SpreadsheetInfo, VersionInfo, VersionSource


def __getattr__(name: str):
    """Carrega integrações somente quando a fonte correspondente é solicitada."""
    if name == "LocalSource":
        from app.sources.local import LocalSource

        return LocalSource
    if name in {"GraphReadError", "GraphSharePointSource"}:
        from app.sources.graph import GraphReadError, GraphSharePointSource

        return {
            "GraphReadError": GraphReadError,
            "GraphSharePointSource": GraphSharePointSource,
        }[name]
    if name in {"BrowserSharePointSource", "SharePointReadError", "SharePointSource"}:
        from app.sources.sharepoint import (
            BrowserSharePointSource,
            SharePointReadError,
            SharePointSource,
        )

        return {
            "BrowserSharePointSource": BrowserSharePointSource,
            "SharePointReadError": SharePointReadError,
            "SharePointSource": SharePointSource,
        }[name]
    raise AttributeError(name)

__all__ = [
    "BrowserSharePointSource", "GraphReadError", "GraphSharePointSource", "LocalSource",
    "SharePointReadError", "SharePointSource", "SpreadsheetInfo",
    "VersionInfo", "VersionSource",
]
