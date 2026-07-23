"""Adequacao Revisao 2 — base financeira: ZERO != VAZIO e janela de 6 competencias.

Prova (helper/motor) que:
  * valor_original_foi_informado distingue zero explicito de ausencia (Secao 6);
  * a janela financeira sao 6 competencias-CALENDARIO terminando na ultima
    INFORMADA, sem puxar competencia anterior, com zero na media e vazio fora do
    denominador (Secoes 4, 17, 21).
E (WEB/AppTest) que a pagina reflete isso e mostra os dados importados (Secao 19).
"""
from __future__ import annotations

import pandas as pd
import pytest

from _adequacao_orcamentaria import (
    valor_original_foi_informado as informado,
    janela_financeira_competencias as janela,
)


# ----------------------------------------------------------------- Secao 6 (helper)

def test_valor_informado_vazio_e_false():
    for v in (None, "", "   ", "—", "-", float("nan"), pd.NA, pd.NaT):
        assert informado(v) is False, v


def test_valor_informado_zero_e_true():
    for v in (0, 0.0, "0", "0,00", "0.00", "R$ 0,00"):
        assert informado(v) is True, v


def test_valor_informado_positivo_e_true():
    for v in ("1000", "1.000,00", "R$ 14.000,00", 1000, 1000.5):
        assert informado(v) is True, v


# ----------------------------------------------------------------- Secao 17 (janela)

def _pares(*comps):
    # comps: (ano, mes, valor|None)
    return list(comps)


def test_A_seis_positivos():
    r = janela(_pares((2026, 1, 10.0), (2026, 2, 12.0), (2026, 3, 11.0),
                      (2026, 4, 13.0), (2026, 5, 14.0), (2026, 6, 15.0)))
    assert r["total"] == 6 and r["competencias_informadas"] == 6
    assert r["media_mensal"] == pytest.approx(75.0 / 6)


def test_B_ultimo_zero_nao_puxa_anterior():
    # jun=0 e a ultima INFORMADA -> janela jan..jun; zero entra na media
    r = janela(_pares((2025, 12, 9000.0), (2026, 1, 10000.0), (2026, 2, 12000.0),
                      (2026, 4, 13000.0), (2026, 5, 14000.0), (2026, 6, 0.0)))
    comps = [(c["ano"], c["mes"]) for c in r["competencias"]]
    assert comps == [(2026, 1), (2026, 2), (2026, 3), (2026, 4), (2026, 5), (2026, 6)]
    assert (2025, 12) not in comps                     # nao puxou dezembro
    assert r["media_mensal"] == pytest.approx((10000 + 12000 + 13000 + 14000 + 0) / 5)


def test_C_zero_intermediario_entra_na_media():
    r = janela(_pares((2026, 1, 100.0), (2026, 2, 0.0), (2026, 3, 200.0),
                      (2026, 4, 300.0), (2026, 5, 400.0), (2026, 6, 500.0)))
    assert r["competencias_informadas"] == 6
    assert r["media_mensal"] == pytest.approx(1500.0 / 6)


def test_D_vazio_intermediario_fora_do_denominador():
    r = janela(_pares((2026, 1, 100.0), (2026, 2, 200.0), (2026, 3, None),
                      (2026, 4, 300.0), (2026, 5, 400.0), (2026, 6, 500.0)))
    assert r["competencias_informadas"] == 5 and r["competencias_sem_info"] == 1
    assert r["media_mensal"] == pytest.approx(1500.0 / 5)


def test_F_sexto_mais_recente_zero_nao_substitui():
    # dados de sobra (jul..ago 2025); jun/2026=0 e a ultima informada
    r = janela(_pares((2025, 7, 1.0), (2025, 8, 2.0), (2026, 1, 10.0),
                      (2026, 5, 14.0), (2026, 6, 0.0)))
    fim = r["competencias"][-1]
    assert (fim["ano"], fim["mes"]) == (2026, 6) and fim["valor"] == 0.0


def test_J_competencia_so_vazia_e_sem_info():
    r = janela(_pares((2026, 6, 100.0), (2026, 5, None)))
    # ultima informada jun -> janela jan..jun; maio None
    maio = [c for c in r["competencias"] if (c["ano"], c["mes"]) == (2026, 5)][0]
    assert maio["valor"] is None and maio["situacao"] == "Sem informação"


# ----------------------------------------------------------------- Secao 21 (golden fin)

def test_golden_financeiro_media_9800():
    r = janela(_pares((2025, 12, 9000.0), (2026, 1, 10000.0), (2026, 2, 12000.0),
                      (2026, 3, None), (2026, 4, 13000.0), (2026, 5, 14000.0),
                      (2026, 6, 0.0)))
    assert r["competencias_informadas"] == 5
    assert r["competencias_sem_info"] == 1
    assert r["total"] == 6
    assert r["media_mensal"] == pytest.approx(9800.0)
    ult = r["competencias"][-1]
    assert (ult["ano"], ult["mes"]) == (2026, 6)


# ----------------------------------------------------------------- Secao 19/21 (WEB)

def _run_pagina_financeiro():
    from streamlit.testing.v1 import AppTest
    df_fin = pd.DataFrame({
        "Competência": ["12/2025", "01/2026", "02/2026", "03/2026", "04/2026", "05/2026",
                        "06/2026", "06/2026"],
        # 03/2026 vazio; 06/2026 com duas linhas (0 informado, somando 0)
        "Valor": [9000.0, 10000.0, 12000.0, "", 13000.0, 14000.0, 0.0, 0.0],
    })
    resultado = {"df_financeiro_mensal": df_fin,
                 "valor_represado_a_pagar": 16888.59, "variacao_acumulada": 0.1201,
                 "modo_apuracao": "Completo"}
    at = AppTest.from_file("pages/12_Adequacao_Orcamentaria.py", default_timeout=120)
    at.session_state["resultado_valor_global"] = resultado
    at.run()
    return at


def test_web_financeiro_media_e_dados_importados():
    at = _run_pagina_financeiro()
    assert not at.exception, at.exception
    blob = "\n".join(str(m.value) for m in at.markdown)
    assert "9.800,00" in blob                    # media das competencias informadas
    assert "Importado da apuração" in blob       # retroativo/percentual importados
    assert "16.888,59" in blob                   # retroativo importado
    assert "12,01%" in blob                      # percentual importado
    assert "5 de 6" in blob                      # competencias informadas


def test_web_financeiro_situacao_zero_e_sem_info():
    at = _run_pagina_financeiro()
    situacoes = set()
    for df in at.dataframe:
        cols = [str(c) for c in df.value.columns]
        if "Situação" in cols and "Competência" in cols:
            situacoes |= set(df.value["Situação"].tolist())
    assert "Zero informado" in situacoes
    assert "Sem informação" in situacoes
    assert "Informado" in situacoes
