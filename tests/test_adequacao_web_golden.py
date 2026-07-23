"""Homologacao funcional da Adequacao Orcamentaria pelo PERCURSO REAL da pagina.

Diferente dos testes estaticos (que apenas leem o texto da pagina), aqui a
pagina 12 e executada via AppTest com a sessao populada como no fluxo real
(Coleta -> Upload -> Adequacao). Isso pega defeitos de wiring/import que os
testes estaticos nao pegam (ex.: origem "Pedidos de compra" quebrando por
NameError) e prova que a WEB reproduz o golden normativo da planilha 04.

Tambem prova (Secao 21) que Financeiro e Pedidos de Compra diferem apenas na
formacao da media historica: dado o mesmo valor de media, o fluxo posterior
(referencia, projecao, diferenca, complemento, programacao) e identico.
"""
from __future__ import annotations

from datetime import date, datetime

import pandas as pd
import pytest

from _adequacao_orcamentaria import calcular_adequacao_orcamentaria as calc, Pedido
from tests.test_adequacao_orcamentaria import PEDIDOS_GOLDEN


def _itens_pc_golden():
    return [{"numero_pc": f"PC-{n}",
             "data_pc": datetime.strptime(d, "%Y-%m-%d").date(),
             "valor_pc": v} for n, (d, v) in enumerate(PEDIDOS_GOLDEN)]


def _sessao_golden():
    df_fin = pd.DataFrame({
        "Competência": ["01/2026", "02/2026", "03/2026", "04/2026", "05/2026", "06/2026"],
        "Valor": [10000.0] * 6,
    })
    return {
        "df_financeiro_mensal": df_fin,
        "itens_pc_v10": {"itens": _itens_pc_golden()},
        "valor_represado_a_pagar": 16888.59,
        "variacao_acumulada": 0.1201,
        "modo_apuracao": "Completo",
    }


def _run_pagina_pc():
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file("pages/12_Adequacao_Orcamentaria.py", default_timeout=120)
    at.session_state["resultado_valor_global"] = _sessao_golden()
    at.run()
    at.radio(key="adequacao_v2_origem").set_value("Pedidos de compra")
    at.text_input(key="adequacao_v2_data_final_vigencia").set_value("05/05/2027")
    at.run()
    return at


def test_pagina_pc_nao_quebra_e_reproduz_golden_na_web():  # T, Q, W
    at = _run_pagina_pc()
    assert not at.exception, f"pagina quebrou no ramo PC: {at.exception}"
    blob = "\n".join(str(m.value) for m in at.markdown)
    # cards principais (formato BR)
    for alvo in ("14.306,98", "16.025,24", "18.900,86", "16.888,59", "35.789,45"):
        assert alvo in blob, f"WEB nao exibiu {alvo}"


def test_pagina_pc_programacao_por_exercicio_na_web():  # O
    at = _run_pagina_pc()
    prog = None
    for df in at.dataframe:
        cols = [str(c) for c in df.value.columns]
        if "Exercício" in cols and "Valor" in cols:
            prog = df.value
    assert prog is not None, "tabela de programacao ausente"
    linhas = {str(r["Exercício"]): str(r["Valor"]) for _, r in prog.iterrows()}
    assert linhas.get("2026") == "R$ 27.198,15"
    assert linhas.get("2027") == "R$ 8.591,30"
    assert linhas.get("TOTAL") == "R$ 35.789,45"


def test_financeiro_e_pc_convergem_apos_a_media():  # R, Secao 21
    """Mesma media -> mesmo fluxo. Financeiro com k meses todos = M produz media
    M; PC com media M deve gerar identicos referencia/diferenca/complemento/prog."""
    M = 14306.976846153848
    comum = dict(percentual=0.1201, ultima_competencia=date(2026, 6, 1),
                 data_fim_vigencia=date(2027, 5, 5), retroativo=16888.59)
    r_fin = calc(origem="Financeiro mensal", financeiro_mensal=[M] * 6, **comum)
    # PC sintetico: um pedido no ultimo mes, janela de 1 mes (media = M)
    r_pc = calc(origem="Pedidos de compra", janela_meses=1,
                pedidos=[Pedido("p", date(2026, 6, 10), M)], **comum)
    assert r_fin["media_mensal"] == pytest.approx(M)
    assert r_pc["media_mensal"] == pytest.approx(M)
    for chave in ("referencia_reajustada", "diferenca_futura", "complemento_estimado",
                  "meses_projetados"):
        assert r_fin[chave] == pytest.approx(r_pc[chave]), f"divergencia em {chave}"
    prog_fin = {p["exercicio"]: round(p["valor"], 2) for p in r_fin["programacao_por_exercicio"]}
    prog_pc = {p["exercicio"]: round(p["valor"], 2) for p in r_pc["programacao_por_exercicio"]}
    assert prog_fin == prog_pc
