"""Identidade mínima e confiável da sessão local que iniciou a auditoria."""

from __future__ import annotations

import getpass


def local_executor_name() -> str | None:
    """Retorna o usuário efetivo da sessão, ou vazio quando não identificável."""
    try:
        value = getpass.getuser()
    except (OSError, KeyError):
        return None
    normalized = value.strip() if isinstance(value, str) else ""
    return normalized or None
