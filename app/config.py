"""Configuração operacional por ambiente e arquivo local não secreto."""

from __future__ import annotations

from dataclasses import dataclass
import json
import os
from pathlib import Path
from typing import Mapping
from urllib.parse import urlsplit


DEFAULT_SHAREPOINT_SITE_URL = "https://hypermarcas.sharepoint.com/controle_qualidade"
DEFAULT_SHAREPOINT_SCOPE_PATHS = (
    "/controle_qualidade/Documentos Compartilhados1",
)


def default_local_config_path() -> Path:
    """Retorna um local persistente por usuário, sem depender do diretório atual."""
    if os.name == "nt" and os.getenv("APPDATA"):
        return Path(os.environ["APPDATA"]) / "Trilha de Auditoria" / "config.json"
    config_home = Path(os.getenv("XDG_CONFIG_HOME", Path.home() / ".config"))
    return config_home / "trilha-de-auditoria" / "config.json"


def _read_local_config(path: Path) -> Mapping[str, object]:
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ValueError(f"Configuração local inválida em {path}: {error}") from error
    if not isinstance(payload, dict):
        raise ValueError(f"Configuração local inválida em {path}: objeto JSON esperado")
    return payload


def _parse_scopes(value: object) -> tuple[str, ...]:
    if isinstance(value, str):
        values = value.split(";")
    elif isinstance(value, list) and all(isinstance(item, str) for item in value):
        values = value
    elif value is None:
        values = []
    else:
        raise ValueError("Configuração local inválida: sharepoint_scope_paths")
    return tuple(path.strip() for path in values if path.strip())


@dataclass(frozen=True)
class Settings:
    """Configurações operacionais sem credenciais embutidas."""

    database_path: Path = Path("data/database/auditoria.db")
    log_path: Path = Path("logs/auditoria.log")
    log_level: str = "INFO"
    temp_directory: Path = Path("data/temp")
    reports_directory: Path = Path("data/reports")
    backups_directory: Path = Path("data/backups")
    sharepoint_tenant_id: str | None = None
    sharepoint_client_id: str | None = None
    sharepoint_client_secret: str | None = None
    sharepoint_site_id: str | None = None
    sharepoint_drive_id: str | None = None
    sharepoint_site_url: str | None = None
    sharepoint_scope_paths: tuple[str, ...] = ()
    local_config_path: Path | None = None

    @classmethod
    def from_environment(cls, *, config_path: str | Path | None = None) -> "Settings":
        """Carrega ambiente > arquivo local > padrão, sem persistir segredos."""
        local_path = Path(config_path) if config_path else default_local_config_path()
        local = _read_local_config(local_path)
        local_site = local.get("sharepoint_site_url")
        if local_site is not None and not isinstance(local_site, str):
            raise ValueError("Configuração local inválida: sharepoint_site_url")
        local_scopes = _parse_scopes(local.get("sharepoint_scope_paths"))
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
            backups_directory=Path(
                os.getenv("AUDIT_BACKUPS_DIRECTORY", "data/backups")
            ),
            sharepoint_tenant_id=os.getenv("SHAREPOINT_TENANT_ID") or None,
            sharepoint_client_id=os.getenv("SHAREPOINT_CLIENT_ID") or None,
            sharepoint_client_secret=os.getenv("SHAREPOINT_CLIENT_SECRET") or None,
            sharepoint_site_id=os.getenv("SHAREPOINT_SITE_ID") or None,
            sharepoint_drive_id=os.getenv("SHAREPOINT_DRIVE_ID") or None,
            sharepoint_site_url=(
                os.getenv("SHAREPOINT_SITE_URL")
                or local_site
                or DEFAULT_SHAREPOINT_SITE_URL
            ),
            sharepoint_scope_paths=(
                _parse_scopes(os.environ["SHAREPOINT_SCOPE_PATHS"])
                if "SHAREPOINT_SCOPE_PATHS" in os.environ
                else local_scopes or DEFAULT_SHAREPOINT_SCOPE_PATHS
            ),
            local_config_path=local_path,
        )

    def save_browser_sharepoint(
        self, site_url: str, scope_paths: tuple[str, ...]
    ) -> None:
        """Persiste somente localização operacional SharePoint, de forma atômica."""
        self.validate_browser_sharepoint(site_url, scope_paths)
        path = self.local_config_path or default_local_config_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_suffix(path.suffix + ".tmp")
        payload = {
            "sharepoint_site_url": site_url.strip().rstrip("/"),
            "sharepoint_scope_paths": list(scope_paths),
        }
        temporary.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        temporary.replace(path)

    @staticmethod
    def validate_browser_sharepoint(
        site_url: str, scope_paths: tuple[str, ...]
    ) -> None:
        """Valida os dois campos não secretos antes de salvar ou conectar."""
        parsed = urlsplit(site_url.strip())
        if parsed.scheme != "https" or not parsed.netloc:
            raise ValueError("Informe uma URL SharePoint HTTPS válida.")
        if not scope_paths:
            raise ValueError("Informe ao menos um caminho/escopo SharePoint.")
        if any(not path.startswith("/") for path in scope_paths):
            raise ValueError("Cada caminho/escopo SharePoint deve começar com '/'.")

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
            raise ValueError(
                "Configuração SharePoint incompleta: " + ", ".join(missing)
            )
        assert self.sharepoint_site_url is not None
        return self.sharepoint_site_url, self.sharepoint_scope_paths

    def create_directories(self) -> None:
        """Cria os diretórios locais usados pela aplicação."""
        for directory in (
            self.database_path.parent,
            self.log_path.parent,
            self.temp_directory,
            self.reports_directory,
            self.backups_directory,
        ):
            directory.mkdir(parents=True, exist_ok=True)
