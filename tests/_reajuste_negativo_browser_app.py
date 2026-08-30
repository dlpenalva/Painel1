"""Aplicativo auxiliar para prova manual em navegador do REAJ-NEG-1.

O fluxo de produção é executado integralmente; apenas a consulta externa ao IST
é substituída por respostas determinísticas para tornar o cenário reproduzível.
"""

from __future__ import annotations

import os
import runpy
import sys
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import _indice_utils  # noqa: E402


def _ist_deterministico(data: pd.Timestamp) -> dict:
    inicio = pd.Timestamp(data).replace(day=1)
    fim = inicio + pd.DateOffset(months=12)
    fluxo = os.environ.get("REAJ_NEG_BROWSER_FLUXO", "simples")
    variacao = -0.02 if fluxo == "simples" or inicio.year == 2023 else 0.05
    return {
        "variacao": variacao,
        "var": variacao,
        "i_ini": 100.0,
        "i_fim": 100.0 * (1.0 + variacao),
        "d_ini": inicio,
        "d_fim": fim,
        "p_ini": inicio,
        "p_fim": fim,
        "metodo": "IST determinístico de teste",
        "serie": "IST-TESTE",
        "dados": pd.DataFrame(
            {"data": [inicio, fim], "indice": [100.0, 100.0 * (1.0 + variacao)]}
        ),
    }


fluxo = os.environ.get("REAJ_NEG_BROWSER_FLUXO", "simples")
pagina = (
    ROOT / "pages" / "02_Calculo_Represados.py"
    if fluxo == "multiplos"
    else ROOT / "pages" / "01_Calculo_Simples.py"
)

original = _indice_utils.calcular_ist_numero_indice
_indice_utils.calcular_ist_numero_indice = _ist_deterministico
try:
    runpy.run_path(str(pagina), run_name="__main__")
finally:
    _indice_utils.calcular_ist_numero_indice = original
