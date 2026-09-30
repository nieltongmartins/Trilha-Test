"""Resolução e cobertura de auditorias por intervalo de IDs técnicos."""

from __future__ import annotations

from dataclasses import dataclass
import logging
import sqlite3

from app.sources.base import VersionInfo


logger = logging.getLogger("auditoria_excel.interval")
Pair = tuple[VersionInfo, VersionInfo]


@dataclass(frozen=True, slots=True)
class CoverageBlock:
    start: VersionInfo
    end: VersionInfo
    pairs: tuple[Pair, ...]


@dataclass(frozen=True, slots=True)
class AuditCoverage:
    requested_pairs: tuple[Pair, ...]
    covered_pairs: tuple[Pair, ...]
    missing_pairs: tuple[Pair, ...]
    coverage_blocks: tuple[CoverageBlock, ...]
    missing_blocks: tuple[CoverageBlock, ...]

    @property
    def coverage_percent(self) -> float:
        return (100.0 * len(self.covered_pairs) / len(self.requested_pairs)
                if self.requested_pairs else 100.0)


def resolve_interval(
    versions: list[VersionInfo] | tuple[VersionInfo, ...],
    start_version_id: str | None,
    end_version_id: str | None,
) -> tuple[VersionInfo, ...]:
    """Resolve limites exclusivamente por technical_version_id.

    Limite vazio significa a primeira/última versão do catálogo. Labels nunca
    participam da identidade ou da ordenação.
    """
    ordered = tuple(versions)
    if len(ordered) < 2:
        raise ValueError("O catálogo precisa conter ao menos duas versões.")
    ids = [version.id for version in ordered]
    if len(ids) != len(set(ids)):
        raise ValueError("O catálogo contém IDs técnicos duplicados.")
    start_id = start_version_id or ids[0]
    end_id = end_version_id or ids[-1]
    try:
        start = ids.index(start_id)
    except ValueError as error:
        raise ValueError(f"Versão inicial inexistente (ID técnico {start_id}).") from error
    try:
        end = ids.index(end_id)
    except ValueError as error:
        raise ValueError(f"Versão final inexistente (ID técnico {end_id}).") from error
    if start >= end:
        raise ValueError("O intervalo deve estar em ordem e conter ao menos duas versões.")
    return ordered[start : end + 1]


def _blocks(pairs: tuple[Pair, ...]) -> tuple[CoverageBlock, ...]:
    if not pairs:
        return ()
    groups: list[list[Pair]] = [[pairs[0]]]
    for pair in pairs[1:]:
        if groups[-1][-1][1].id == pair[0].id:
            groups[-1].append(pair)
        else:
            groups.append([pair])
    return tuple(CoverageBlock(group[0][0], group[-1][1], tuple(group)) for group in groups)


def get_audit_coverage(
    connection: sqlite3.Connection,
    workbook_id: int,
    versions: list[VersionInfo] | tuple[VersionInfo, ...],
    start_version_id: str | None,
    end_version_id: str | None,
) -> AuditCoverage:
    """Classifica pares usando uma consulta indexada e validação conservadora.

    Um registro é reutilizável somente quando a comparação foi persistida com
    status final e hash SHA-256, que são gravados no commit atômico do par. O
    status global da execução não invalida pares confirmados antes de uma falha.
    Registros legados incompletos
    são deliberadamente tratados como lacunas.
    """
    selected = resolve_interval(versions, start_version_id, end_version_id)
    requested = tuple(zip(selected, selected[1:]))
    wanted = {(left.id, right.id) for left, right in requested}
    start_id, end_id = selected[0].id, selected[-1].id
    rows = connection.execute(
        """SELECT v.versao_anterior_id, v.versao_atual_id
             FROM versao_processada v
            WHERE v.planilha_id=?
              AND CAST(v.versao_anterior_id AS INTEGER)>=CAST(? AS INTEGER)
              AND CAST(v.versao_atual_id AS INTEGER)<=CAST(? AS INTEGER)
              AND v.status IN ('PROCESSADA','SEM_ALTERACOES')
              AND length(v.hash_origem)=64""",
        (workbook_id, start_id, end_id),
    ).fetchall()
    persisted = {(str(row[0]), str(row[1])) for row in rows} & wanted
    covered = tuple(pair for pair in requested if (pair[0].id, pair[1].id) in persisted)
    missing = tuple(pair for pair in requested if (pair[0].id, pair[1].id) not in persisted)
    return AuditCoverage(requested, covered, missing, _blocks(covered), _blocks(missing))
