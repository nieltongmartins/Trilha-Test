"""Configuração central de logging."""

import logging
from pathlib import Path


def configure_logging(log_path: Path, level: str = "INFO") -> logging.Logger:
    """Configura logger em arquivo sem incluir dados sensíveis."""
    log_path.parent.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("auditoria_excel")
    logger.setLevel(getattr(logging, level.upper(), logging.INFO))
    logger.propagate = False

    for handler in logger.handlers[:]:
        handler.close()
        logger.removeHandler(handler)

    handler = logging.FileHandler(log_path, encoding="utf-8")
    handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")
    )
    logger.addHandler(handler)
    return logger
