"""Exceções específicas da aplicação."""


class AuditError(Exception):
    """Erro base do domínio da aplicação."""


class DatabaseNotConnectedError(AuditError):
    """Operação solicitada sem uma conexão aberta com o banco."""
