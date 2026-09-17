"""Configuração central de logging."""

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
import re


_SENSITIVE_VALUE = re.compile(
    r"(?i)\b(password|senha|token|client[_-]?secret|secret|cookie|authorization)"
    r"(\s*[:=]\s*)(?:bearer\s+)?[^\s,;]+"
)
_BEARER_VALUE = re.compile(r"(?i)\bbearer\s+[^\s,;]+")


def _redact(message: str) -> str:
    """Remove valores de campos de autenticação antes da escrita em disco."""
    message = _SENSITIVE_VALUE.sub(r"\1\2[REDACTED]", message)
    return _BEARER_VALUE.sub("Bearer [REDACTED]", message)


class _RedactingFormatter(logging.Formatter):
    """Formata primeiro para também proteger exceções e argumentos interpolados."""

    def format(self, record: logging.LogRecord) -> str:
        return _redact(super().format(record))


def configure_logging(
    log_path: Path,
    level: str = "INFO",
    *,
    max_bytes: int = 5 * 1024 * 1024,
    backup_count: int = 5,
) -> logging.Logger:
    """Configura log UTF-8 persistente, limitado e protegido contra segredos."""
    log_path.parent.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("auditoria_excel")
    logger.setLevel(getattr(logging, level.upper(), logging.INFO))
    logger.propagate = False

    for handler in logger.handlers[:]:
        handler.close()
        logger.removeHandler(handler)

    handler = RotatingFileHandler(
        log_path,
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding="utf-8",
    )
    handler.setFormatter(
        _RedactingFormatter(
            "%(asctime)s %(levelname)s %(name)s [thread=%(threadName)s]: %(message)s",
            datefmt="%Y-%m-%dT%H:%M:%S%z",
        )
    )
    logger.addHandler(handler)
    return logger
