"""Ponto de entrada da aplicação."""

from app.config import Settings
from app.database import Database
from app.logging_config import configure_logging


def main() -> int:
    """Inicializa configuração, logging e o banco de auditoria."""
    settings = Settings.from_environment()
    settings.create_directories()
    logger = configure_logging(settings.log_path, settings.log_level)

    with Database(settings.database_path) as database:
        database.initialize()

    logger.info("Aplicação inicializada; banco disponível em %s", settings.database_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
