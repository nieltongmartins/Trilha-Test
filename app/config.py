"""Configuração da aplicação baseada em variáveis de ambiente."""

from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path


@dataclass(frozen=True)
class Settings:
    """Configurações operacionais sem credenciais embutidas."""

    database_path: Path = Path("data/database/auditoria.db")
    log_path: Path = Path("logs/auditoria.log")
    log_level: str = "INFO"
    temp_directory: Path = Path("data/temp")
    reports_directory: Path = Path("data/reports")
    sharepoint_tenant_id: str | None = None
    sharepoint_client_id: str | None = None
    sharepoint_client_secret: str | None = None
    sharepoint_site_id: str | None = None
    sharepoint_drive_id: str | None = None
    sharepoint_site_url: str | None = None
    sharepoint_scope_paths: tuple[str, ...] = ()

    @classmethod
    def from_environment(cls) -> "Settings":
        """Carrega apenas opções locais necessárias nesta fase."""
        return cls(
            database_path=Path(
                os.getenv("AUDIT_DATABASE_PATH", "data/database/auditoria.db")
            ),
            log_path=Path(os.getenv("AUDIT_LOG_PATH", "logs/auditoria.log")),
            log_level=os.getenv("AUDIT_LOG_LEVEL", "INFO").upper(),
            temp_directory=Path(os.getenv("AUDIT_TEMP_DIRECTORY", "data/temp")),
            reports_directory=Path(
                os.getenv("AUDIT_REPORTS_DIRECTORY", "data/reports")
            ),
            sharepoint_tenant_id=os.getenv("SHAREPOINT_TENANT_ID") or None,
            sharepoint_client_id=os.getenv("SHAREPOINT_CLIENT_ID") or None,
            sharepoint_client_secret=os.getenv("SHAREPOINT_CLIENT_SECRET") or None,
            sharepoint_site_id=os.getenv("SHAREPOINT_SITE_ID") or None,
            sharepoint_drive_id=os.getenv("SHAREPOINT_DRIVE_ID") or None,
            sharepoint_site_url=os.getenv("SHAREPOINT_SITE_URL") or None,
            sharepoint_scope_paths=tuple(
                path.strip()
                for path in os.getenv("SHAREPOINT_SCOPE_PATHS", "").split(";")
                if path.strip()
            ),
        )

    def require_sharepoint(self) -> tuple[str, str, str, str, str]:
        """Retorna a configuração Graph completa sem expor o segredo em erros."""
        values = {
            "SHAREPOINT_TENANT_ID": self.sharepoint_tenant_id,
            "SHAREPOINT_CLIENT_ID": self.sharepoint_client_id,
            "SHAREPOINT_CLIENT_SECRET": self.sharepoint_client_secret,
            "SHAREPOINT_SITE_ID": self.sharepoint_site_id,
            "SHAREPOINT_DRIVE_ID": self.sharepoint_drive_id,
        }
        missing = [name for name, value in values.items() if not value]
        if missing:
            raise ValueError(
                "Configuração SharePoint incompleta: " + ", ".join(missing)
            )
        return tuple(values.values())  # type: ignore[return-value]

    def require_browser_sharepoint(self) -> tuple[str, tuple[str, ...]]:
        """Valida somente localizações não sensíveis usadas pela sessão Edge."""
        missing = []
        if not self.sharepoint_site_url:
            missing.append("SHAREPOINT_SITE_URL")
        if not self.sharepoint_scope_paths:
            missing.append("SHAREPOINT_SCOPE_PATHS")
        if missing:
            raise ValueError("Configuração SharePoint incompleta: " + ", ".join(missing))
        return self.sharepoint_site_url, self.sharepoint_scope_paths

    def create_directories(self) -> None:
        """Cria os diretórios locais usados pela aplicação."""
        for directory in (
            self.database_path.parent,
            self.log_path.parent,
            self.temp_directory,
            self.reports_directory,
        ):
            directory.mkdir(parents=True, exist_ok=True)
