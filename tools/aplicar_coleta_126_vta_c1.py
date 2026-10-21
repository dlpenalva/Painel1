"""Aplica cirurgicamente a Coleta 12.6 ao template oficial.

Somente ``xl/worksheets/sheet14.xml`` (MEMORIA_RESULTADOS) e
``xl/sharedStrings.xml`` (rotulo unico de cobertura_temporal!A6) podem mudar.
O XLSX resultante e validado e recalculado em Microsoft Excel pelos testes.
"""
from __future__ import annotations

import html
import os
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from _coleta_oficial import (  # noqa: E402
    _FORMULA_B15_126,
    _FORMULA_B44_126,
    _FORMULA_W50_126,
    _FORMULA_W67_126,
    _FORMULAS_RETRO_FINANCEIRO_126,
    _ROTULO_COBERTURA_B6_126,
)

TEMPLATE = ROOT / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"
MEMORIA_XML = "xl/worksheets/sheet14.xml"
STRINGS_XML = "xl/sharedStrings.xml"


def _trocar_formula(xml: bytes, celula: str, formula: str) -> bytes:
    padrao = re.compile(
        rb'(<c\b[^>]*\br="' + celula.encode("ascii") + rb'"[^>]*>.*?<f\b[^>]*>)'
        rb'(.*?)(</f>)',
        re.DOTALL,
    )
    nova = html.escape(formula.removeprefix("="), quote=False).encode("utf-8")
    saida, quantidade = padrao.subn(lambda m: m.group(1) + nova + m.group(3), xml)
    if quantidade != 1:
        raise RuntimeError(f"{celula}: esperava uma formula, encontrei {quantidade}")
    return saida


def main() -> None:
    with zipfile.ZipFile(TEMPLATE, "r") as origem:
        memoria = origem.read(MEMORIA_XML)
        for celula, formula in _FORMULAS_RETRO_FINANCEIRO_126.items():
            memoria = _trocar_formula(memoria, celula, formula)
        memoria = _trocar_formula(memoria, "W50", _FORMULA_W50_126)
        memoria = _trocar_formula(memoria, "W67", _FORMULA_W67_126)
        memoria = _trocar_formula(memoria, "B15", _FORMULA_B15_126)
        memoria = _trocar_formula(memoria, "B44", _FORMULA_B44_126)

        # Idempotente: o rotulo pode ja ter sido aplicado por execucao anterior.
        strings = origem.read(STRINGS_XML)
        antigo = "Inicio do ciclo atual".encode("utf-8")
        novo = html.escape(_ROTULO_COBERTURA_B6_126, quote=False).encode("utf-8")
        if strings.count(antigo) == 1:
            strings = strings.replace(antigo, novo, 1)
        elif strings.count(novo) != 1:
            raise RuntimeError("rotulo cobertura_temporal!A6 fora do formato esperado")

        fd, temporario = tempfile.mkstemp(
            suffix=".xlsx", prefix="coleta126_", dir=TEMPLATE.parent
        )
        os.close(fd)
        try:
            with zipfile.ZipFile(temporario, "w") as destino:
                for info in origem.infolist():
                    dados = origem.read(info.filename)
                    if info.filename == MEMORIA_XML:
                        dados = memoria
                    elif info.filename == STRINGS_XML:
                        dados = strings
                    destino.writestr(info, dados)
            shutil.copyfile(temporario, TEMPLATE)
        finally:
            if os.path.exists(temporario):
                os.unlink(temporario)


if __name__ == "__main__":
    main()
