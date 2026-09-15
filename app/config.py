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

    @classmethod
    def from_environment(cls) -> "Settings":
        """Carrega apenas opções locais necessárias nesta fase."""
        return cls(
            database_path=Path(
                os.getenv("AUDIT_DATABASE_PATH", "data/database/auditoria.db")
            ),
            log_path=Path(os.getenv("AUDIT_LOG_PATH", "logs/auditoria.log")),
            log_level=os.getenv("AUDIT_LOG_LEVEL", "INFO").upper(),
        )

    def create_directories(self) -> None:
        """Cria os diretórios locais usados pela aplicação."""
        for directory in (
            self.database_path.parent,
            self.log_path.parent,
            self.temp_directory,
            self.reports_directory,
        ):
            directory.mkdir(parents=True, exist_ok=True)
