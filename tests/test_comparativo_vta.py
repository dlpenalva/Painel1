"""Aba comparativo_VTA — Bloco 1 (precos originais) e Bloco 2 (referencia).

Cobre Secao 12 / Secao 13 (COMPARATIVO): totais originais por ciclo, fator
historico INTEGRAL (nao so os ciclos da apuracao) e separacao explicita entre
o VTA oficial e o contrafactual reajustado.
"""
from __future__ import annotations

import _comparativo_vta as CV

_ITENS = [
    {"ITEM": 1, "VU_ORIGINAL": 100, "QTD_BASE_ORIGINAL": 10,
     "QTD_REM_AJUSTADA_C0": 10, "QTD_REM_AJUSTADA_C1": 8,
     "QTD_REM_AJUSTADA_C2": 5, "QTD_REM_AJUSTADA_C3": 3,
     "QTD_REM_AJUSTADA_C4": 2},
    {"ITEM": 2, "VU_ORIGINAL": 50, "QTD_BASE_ORIGINAL": 20,
     "QTD_REM_AJUSTADA_C0": 20, "QTD_REM_AJUSTADA_C1": 15,
     "QTD_REM_AJUSTADA_C2": 10, "QTD_REM_AJUSTADA_C3": 5,
     "QTD_REM_AJUSTADA_C4": 4},
]


def _leitura(fator=1.20, por_ciclo=True):
    pc = {"C4": {"fator_acumulado": fator}} if por_ciclo else {}
    return {
        "controle": {"modo": "principal", "ciclo_vigente": "C4"},
        "parametros_v10": {"por_ciclo": pc},
        "posicao_contratual": {"itens": _ITENS},
        "historico_vu": {"itens": [
            {"item": 1, "vu_ciclos": {"VU_C0": 100, "VU_C4": 120}},
            {"item": 2, "vu_ciclos": {"VU_C0": 50, "VU_C4": 60}},
        ]},
        "itens_pc_v10": {"itens": []},
    }


# ---- N. valores originais por ciclo (Bloco 1) ------------------------------
def test_bloco1_totais_por_ciclo_precos_originais():
    r = CV.montar_comparativo_vta(_leitura())
    t = r["bloco1_totais_por_ciclo"]
    assert t == {"C0": 2000.0, "C1": 1550.0, "C2": 1000.0,
                 "C3": 550.0, "C4": 400.0}
    # sem reajuste: item 1 C4 = 100 * 2
    linha1 = r["bloco1_itens"][0]
    assert linha1["ciclos"]["C4"]["valor"] == 200.0


# ---- O. fator historico INTEGRAL usado no contrafactual --------------------
def test_bloco2_fator_historico_integral():
    b2 = CV.montar_comparativo_vta(_leitura(fator=1.20))["bloco2"]
    assert b2["fator_historico_integral"] == 1.20
    assert b2["valor_original_contrato"] == 2000.0
    assert b2["contrato_cheio_reajustado"] == 2400.0  # 2000 * 1.20


def test_fator_integral_fallback_historico_vu():
    # sem por_ciclo, deriva VU_C4/VU_C0 = 120/100 = 1.20
    b2 = CV.montar_comparativo_vta(_leitura(por_ciclo=False))["bloco2"]
    assert b2["fator_historico_integral"] == 1.20


# ---- VTA oficial separado do contrafactual ---------------------------------
def test_vta_separado_do_contrafactual():
    b2 = CV.montar_comparativo_vta(_leitura())["bloco2"]
    assert b2["e_vta"] is False
    assert "NAO UTILIZAR COMO VTA" in b2["titulo"]
    assert "Nao utilizar este valor como VTA" in b2["alerta"]
    # campos distintos (contrafactual != campo do VTA)
    assert "vta_metodo_aplicado" in b2 and "contrato_cheio_reajustado" in b2
