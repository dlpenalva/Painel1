# -*- coding: utf-8 -*-
"""NIVEL B — fluxo REAL da Comunicacao a Contratada, fim a fim.

Prova a sequencia inteira SEM inserir memoria_calculo manualmente:

    indice (fonte real) -> res -> normalizar_memoria_calculo -> ciclo ->
    gerar_rascunho_email_contratada -> montar_txt_download (bytes finais).

Cenarios: ciclo unico + IPCA (SGS/BCB, rede; skip se indisponivel),
ciclo unico + IST (ist.csv local) e multiciclo com 2 ciclos IST.

Criterio de aceite: o TXT FINAL contem literalmente "MEMÓRIA DE CÁLCULO" e as
competencias/numeros-indice derivados do PROPRIO res calculado.
"""
from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import pandas as pd
import pytest
from dateutil.relativedelta import relativedelta

RAIZ = Path(__file__).resolve().parents[1]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from _email_contratada import (  # noqa: E402
    gerar_rascunho_email_contratada,
    montar_txt_download,
)
from _indice_utils import (  # noqa: E402
    calcular_ist_numero_indice,
    coletar_sgs_produtorio,
    serie_sgs_do_indice,
)
from _memoria_calculo import normalizar_memoria_calculo  # noqa: E402


def _ciclo_como_a_pagina(nome, res, financeiro_inicio):
    """Monta o dict do ciclo com o MESMO wiring das paginas 01/02:
    memoria_calculo = normalizar_memoria_calculo(res, fator, aplicado)."""
    percentual_indice = float(res["variacao"])
    percentual_aplicado = 0.0 if percentual_indice < 0 else percentual_indice
    fator = 1.0 if percentual_indice < 0 else 1 + percentual_indice
    return {
        "ciclo": nome,
        "situacao_aplicada": "Tempestivo",
        "percentual_indice": percentual_indice,
        "percentual_aplicado": percentual_aplicado,
        "variacao": percentual_aplicado,
        "variacao_formatada": f"{percentual_aplicado*100:,.2f}%".replace(".", ","),
        "financeiro_inicio": financeiro_inicio,
        "memoria_calculo": normalizar_memoria_calculo(
            res, fator, percentual_aplicado
        ),
    }


def _txt_final(ciclos, indice):
    assunto, corpo = gerar_rascunho_email_contratada(
        ciclos, numero_contrato="CT-1/2026", indice=indice
    )
    dados = montar_txt_download(assunto, corpo)
    assert dados.startswith("﻿".encode("utf-8") + b"ASSUNTO: ")
    return dados.decode("utf-8-sig")


def _competencia(ts) -> str:
    return pd.Timestamp(ts).strftime("%m/%Y")


def test_fluxo_real_ciclo_unico_ipca():
    """IPCA real via SGS/BCB (o caso em que a ausencia foi observada)."""
    dt_base = date(2025, 4, 1)
    try:
        res = coletar_sgs_produtorio(
            serie_sgs_do_indice("IPCA (433)"), dt_base,
            dt_base + relativedelta(months=11), timeout=20,
        )
    except Exception as exc:  # rede indisponivel nao pode mascarar regressao
        pytest.skip(f"SGS/BCB indisponivel: {exc}")
    if not res:
        pytest.skip("SGS/BCB sem dados no intervalo")

    txt = _txt_final([_ciclo_como_a_pagina("C2", res, "01/04/2026")], "IPCA (433)")
    assert "MEMÓRIA DE CÁLCULO" in txt
    secao = txt.split("MEMÓRIA DE CÁLCULO", 1)[1]
    assert "Ciclo 2" in secao
    # competencias derivadas do PROPRIO res (nada hardcoded)
    for ts in list(res["dados"]["data"])[:2]:
        assert f"{_competencia(ts)}: " in secao
    assert "Fator apurado: " in secao
    assert "Variação apurada: " in secao
    assert "Método/Fonte: Produtório de taxas mensais (SGS/BCB)" in secao


def _res_ist(dt_base: date):
    try:
        res = calcular_ist_numero_indice(dt_base, caminho=str(RAIZ / "ist.csv"))
    except Exception as exc:
        pytest.skip(f"serie IST indisponivel: {exc}")
    if not res:
        pytest.skip("IST sem competencias para o intervalo")
    return res


def test_fluxo_real_ciclo_unico_ist():
    """IST real (numero-indice; nunca vira memoria mensal ficticia)."""
    res = _res_ist(date(2024, 8, 2))
    txt = _txt_final([_ciclo_como_a_pagina("C1", res, "01/08/2025")], "IST (Anatel)")
    assert "MEMÓRIA DE CÁLCULO" in txt
    secao = txt.split("MEMÓRIA DE CÁLCULO", 1)[1]
    assert "Ciclo 1" in secao
    ini = f"{float(res['i_ini']):.4f}".replace(".", ",")
    fim = f"{float(res['i_fim']):.4f}".replace(".", ",")
    assert f"Número-índice inicial ({_competencia(res['d_ini'])}): {ini}" in secao
    assert f"Número-índice final ({_competencia(res['d_fim'])}): {fim}" in secao
    assert "Divisão de Número-Índice" in secao
    # a serie mensal completa da memoria NAO vira lista no TXT: so 1 inicial
    # e 1 final por ciclo (nem "final" repetido, nem meses ficticios)
    assert secao.count("Número-índice inicial") == 1
    assert secao.count("Número-índice final") == 1


def test_fluxo_real_multiciclo_dois_ciclos():
    """2 ciclos IST: uma secao de memoria por ciclo efetivamente calculado."""
    res1 = _res_ist(date(2023, 8, 2))
    res2 = _res_ist(date(2024, 8, 2))
    ciclos = [
        _ciclo_como_a_pagina("C1", res1, "01/08/2024"),
        _ciclo_como_a_pagina("C2", res2, "01/08/2025"),
    ]
    txt = _txt_final(ciclos, "IST (Anatel)")
    secao = txt.split("MEMÓRIA DE CÁLCULO", 1)[1]
    assert "Ciclo 1" in secao and "Ciclo 2" in secao
    for res in (res1, res2):
        ini = f"{float(res['i_ini']):.4f}".replace(".", ",")
        assert ini in secao
    assert secao.count("Número-índice inicial") == 2
    assert secao.count("Número-índice final") == 2
