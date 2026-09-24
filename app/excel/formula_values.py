"""Representação canônica dos valores especiais de fórmula do openpyxl."""

from __future__ import annotations

import json
from typing import Any

from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula


def _canonical(kind: str, metadata: dict[str, Any]) -> str:
    """Codifica metadados sem depender de ``repr`` ou identidade do objeto."""
    payload = json.dumps(
        metadata,
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    )
    return f"{kind}|{payload}"


def normalize_formula_value(value: Any) -> Any:
    """Retorna um valor comparável e persistível para fórmulas do openpyxl.

    Strings de fórmula comuns e valores que não são fórmulas permanecem
    inalterados. Fórmulas estruturadas como objetos são convertidas usando
    somente seus atributos semânticos documentados pela versão instalada.
    """
    if isinstance(value, ArrayFormula):
        return _canonical(
            "ARRAYFORMULA",
            {"formula": value.text, "ref": value.ref},
        )
    if isinstance(value, DataTableFormula):
        return _canonical(
            "DATATABLEFORMULA",
            {
                "ca": value.ca,
                "del1": value.del1,
                "del2": value.del2,
                "dt2D": value.dt2D,
                "dtr": value.dtr,
                "r1": value.r1,
                "r2": value.r2,
                "ref": value.ref,
            },
        )
    return value
