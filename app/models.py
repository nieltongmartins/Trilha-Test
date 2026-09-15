"""Valores controlados do modelo de persistência."""

from enum import Enum


class ChangeType(str, Enum):
    ADD = "ADD"
    DEL = "DEL"
    MOD = "MOD"


class ProcessedVersionStatus(str, Enum):
    PROCESSED = "PROCESSADA"
    WITHOUT_CHANGES = "SEM_ALTERACOES"
    ERROR = "ERRO"


class AuditExecutionStatus(str, Enum):
    RUNNING = "EM_EXECUCAO"
    COMPLETED = "CONCLUIDA"
    COMPLETED_WITHOUT_UPDATES = "CONCLUIDA_SEM_NOVIDADES"
    FAILED = "FALHA"
