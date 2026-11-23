"""Adequacao 12.23: amostra para PC irregular, vigencia valida e paridade WEB.

Garante que o automatico gera CENARIO DE PLANEJAMENTO, nao ocorrencias
mensais atestadas, somente quando existe evidencia historica suficiente.
"""
from datetime import date

import pandas as pd
import pytest

from _adequacao_orcamentaria import avaliar_evidencia_media_irregular_pc
from _adequacao_ui import validar_vigencia_para_adequacao
from test_adequacao_pc_janela_potencial import (
    _blob, _run, _sessao_pc, _valor_do_card,
)


def _evidencia(*, pedidos=2529, meses_com=10, janela=55, ciclos=None,
               saldo=39364302.31, horizonte=6, usa_c0=False):
    if ciclos is None:
        ciclos = ["C1", "C2", "C3"]
    base = {
        "pedidos_considerados": pedidos,
        "meses_com_pedido": meses_com,
        "meses_historico": janela,
        "media_mensal": 892877.92,
    }
    cadencia = {"ciclos_base": ciclos, "usa_c0": usa_c0}
    return avaliar_evidencia_media_irregular_pc(base, cadencia, saldo, horizonte)


def test_amostra_2529_pcs_55_meses_7_2_1_aceita_para_planejamento():
    assert _evidencia()["apta"]
    assert _evidencia()["quantidade_pcs"] == 2529
    assert _evidencia()["meses_historico"] == 55
    assert _evidencia()["meses_com_pc"] == 10


@pytest.mark.parametrize("alteracao", [
    {"pedidos": 3}, {"meses_com": 3}, {"janela": 5},
    {"ciclos": ["C1"]}, {"usa_c0": True},
    {"saldo": None}, {"horizonte": 0},
])
def test_evidencia_insuficiente_mantem_futuro_nao_determinado(alteracao):
    dados = _evidencia(**alteracao)
    assert not dados["apta"], alteracao
    assert dados["motivos"]


@pytest.mark.parametrize("valor,esperado", [
    ("16/04/2027", "16/04/2027"),
    ("29/02/2028", "29/02/2028"),
])
def test_vigencia_estrita_valida(valor, esperado):
    assert validar_vigencia_para_adequacao(valor, date(2026, 10, 19)) == (esperado, "")


@pytest.mark.parametrize("valor", [
    "", "16/4/2027", "16-04-2027", "2027-04-16",
    "31/02/2027", "29/02/2027", "16/10/2026", "01/09/2026",
])
def test_vigencia_invalida_falha_com_motivo(valor):
    final, motivo = validar_vigencia_para_adequacao(valor, date(2026, 10, 19))
    assert final == ""
    assert motivo


def _sessao_pc_suficiente():
    # 55 meses de abril/2022 a outubro/2026, com movimentacao em C1 (7),
    # C2 (2), C3 (1) e C0 (1). Muitos PCs no mesmo mes contam como UM mes.
    meses = [(2022, 4)]
    meses += [(2023, m) for m in (4, 5, 6, 7, 8, 9, 10)]
    meses += [(2024, 4), (2024, 5), (2025, 4)]
    meses += [(2023, 4)] * 7
    assert len(meses) == 18
    total = 892877.92 * 55
    unitario = round(total / len(meses), 2)
    itens = [{
        "numero_pc": f"PC-{i+1}", "data_pc": date(ano, mes, 12),
        "valor_pc": unitario if i < len(meses) - 1
        else round(total - (len(meses) - 1) * unitario, 2),
    } for i, (ano, mes) in enumerate(meses)]
    sessao = _sessao_pc(saldo=39364302.31, data_corte=date(2026, 10, 19))
    sessao["itens_pc_v10"]["itens"] = itens
    sessao["variacao_acumulada"] = 0.0381
    sessao["valor_represado_a_pagar"] = 36397.14
    sessao["resultado_consolidado"]["retroativo_reconhecido"] = 36397.14
    sessao["df_ciclos"] = pd.DataFrame({
        "Ciclo": ["C0", "C1", "C2", "C3", "C4"],
        "Data-base": [f"01/04/{ano}" for ano in (2022, 2023, 2024, 2025, 2026)],
        "Situação automática": ["Base", "Precluso", "Precluso", "Precluso", "Vigente"],
    })
    return sessao


def test_web_auto_irregular_apto_projeta_estimada_sem_mensalidade_comprovada():
    at = _run(_sessao_pc_suficiente(), vigencia="16/04/2027")
    assert not at.exception, at.exception
    blob = _blob(at)
    assert "C1: 7 meses; C2: 2 meses; C3: 1 mês" in " ".join(e.value for e in at.error)
    assert "não significa compras mensais previstas" in " ".join(e.value for e in at.error)
    assert "Automática — cenário pela média de PCs; baixa confiabilidade" in blob
    assert "COMPLEMENTAÇÃO ORÇAMENTÁRIA ESTIMADA" in blob
    assert "COMPLEMENTAÇÃO ORÇAMENTÁRIA PARCIAL" not in blob
    assert _valor_do_card(blob, "Retroativo reconhecido considerado") == pytest.approx(36397.14)
    assert _valor_do_card(blob, "Diferença futura projetada") == pytest.approx(204111.90)
    assert _valor_do_card(blob, "COMPLEMENTAÇÃO ORÇAMENTÁRIA ESTIMADA") == pytest.approx(240509.04)
    assert "Rateio ilustrativo" in blob
    assert "Vigência considerada no cálculo: 16/04/2027" in " ".join(x.value for x in at.success)


def test_web_vigencia_invalida_nao_gera_futuro_positivo():
    at = _run(_sessao_pc_suficiente(), vigencia="31/02/2027")
    assert not at.exception, at.exception
    erros = " ".join(e.value for e in at.error)
    assert "Essa data não existe" in erros
    blob = _blob(at)
    assert "COMPLEMENTAÇÃO ORÇAMENTÁRIA PARCIAL" in blob
    assert "Não determinada" in blob


def test_pagina_destaca_vigencia_sem_alterar_fonte_do_corte():
    pagina = open("pages/12_Adequacao_Orcamentaria.py", encoding="utf-8").read()
    assert ".st-key-adequacao_v3_data_final_vigencia" in pagina
    assert "Não é obtida automaticamente da apuração" in pagina
    assert "validar_vigencia_para_adequacao(" in pagina
    assert "data_corte_canonica" in pagina
