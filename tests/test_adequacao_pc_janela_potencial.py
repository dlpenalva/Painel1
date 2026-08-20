"""Adequacao Orcamentaria no metodo Pedidos de Compra.

Protege quatro decisoes desta entrega:
  * a janela historica e o historico REAL importado ate o corte (nao um bloco
    fixo de 39 meses que inventa competencias anteriores ao primeiro PC);
  * a competencia final do historico vem da DATA DE CORTE canonica — a data
    final da VIGENCIA e outro conceito e nao a substitui;
  * a projecao futura vai do mes seguinte ao corte ate o mes da vigencia;
  * retroativo reconhecido e retroativo potencial permanecem grandezas
    separadas: o potencial nunca integra a complementacao confirmada.
"""
from __future__ import annotations

import re
import sys
from datetime import date
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from _adequacao_orcamentaria import (  # noqa: E402
    janela_automatica_pcs,
    media_pedidos_compra,
    pedidos_de_itens_pc,
)
from _adequacao_ui import (  # noqa: E402
    gerar_periodos_projecao,
    parse_moeda_br,
    periodo_para_label,
)

CORTE = date(2026, 8, 19)
VIGENCIA = "29/02/2028"
RECONHECIDO = 57701.49
POTENCIAL = 12000.0

# Historico real do cenario que revelou o defeito: nada em 2024.
_ITENS_PC = [
    {"numero_pc": "PC-1", "data_pc": date(2025, 3, 14), "valor_pc": 12000.0},
    {"numero_pc": "PC-2", "data_pc": date(2025, 11, 5), "valor_pc": 18000.0},
    {"numero_pc": "PC-3", "data_pc": date(2026, 6, 20), "valor_pc": 9000.0},
]


def _sessao_pc(*, potencial=POTENCIAL, com_potencial=True, data_corte=CORTE):
    consolidado = {
        "medidas_pc_aplicaveis": True,
        "retroativo_reconhecido": RECONHECIDO,
        "valor_atualizado_em_analise": 41000.0,
        "fora_do_corte": {"aplicavel": True, "data_corte": data_corte},
    }
    if com_potencial:
        consolidado["retroativo_potencial"] = potencial
    return {
        "itens_pc_v10": {"itens": [dict(i) for i in _ITENS_PC]},
        "valor_represado_a_pagar": RECONHECIDO,
        "variacao_acumulada": 0.0289,
        "modo_apuracao": "Completo",
        "resultado_consolidado": consolidado,
    }


def _run(sessao, *, origem="Pedidos de Compra", vigencia=VIGENCIA):
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file("pages/12_Adequacao_Orcamentaria.py", default_timeout=180)
    at.session_state["resultado_valor_global"] = sessao
    at.run()
    at.radio(key="adequacao_v3_origem").set_value(origem)
    at.text_input(key="adequacao_v3_data_final_vigencia").set_value(vigencia)
    at.run()
    return at


def _blob(at):
    return "\n".join(str(m.value) for m in at.markdown)


def _valor_do_card(blob: str, rotulo: str) -> float:
    """Extrai o numero do card/leitura cujo rotulo e `rotulo`."""
    idx = blob.index(rotulo)
    achado = re.search(r"(-?[\d\.]+,\d{2})", blob[idx + len(rotulo):])
    assert achado, f"valor nao encontrado para {rotulo!r}"
    return parse_moeda_br(achado.group(1))


# ---------------------------------------------------------------------------
# Motor: janela automatica
# ---------------------------------------------------------------------------

def test_janela_comeca_no_primeiro_pc_e_termina_no_mes_do_corte():
    janela = janela_automatica_pcs(pedidos_de_itens_pc(_ITENS_PC), CORTE)
    assert janela == {
        "inicio_janela": date(2025, 3, 1),
        "fim_janela": date(2026, 8, 31),
        "meses": 18,
    }


def test_janela_nao_inventa_meses_anteriores_ao_primeiro_pc():
    janela = janela_automatica_pcs(pedidos_de_itens_pc(_ITENS_PC), CORTE)
    assert janela["inicio_janela"].year == 2025
    assert janela["meses"] != 39


def test_exclusao_manual_nao_desloca_o_inicio_da_janela():
    """Desmarcar um PC muda a media; jamais o inicio do historico."""
    todos = pedidos_de_itens_pc(_ITENS_PC)
    com_exclusao = pedidos_de_itens_pc(_ITENS_PC, exclusoes={"PC-1"})
    janela = janela_automatica_pcs(todos, CORTE)
    assert janela_automatica_pcs(com_exclusao, CORTE) == janela
    base_todos = media_pedidos_compra(todos, CORTE, janela["meses"])
    base_excl = media_pedidos_compra(com_exclusao, CORTE, janela["meses"])
    assert base_excl["media_mensal"] < base_todos["media_mensal"]
    assert base_excl["inicio_janela"] == base_todos["inicio_janela"]


def test_sem_corte_ou_sem_pc_nao_ha_janela_derivavel():
    assert janela_automatica_pcs(pedidos_de_itens_pc(_ITENS_PC), None) is None
    assert janela_automatica_pcs([], CORTE) is None
    # PC posterior ao corte nao abre janela sozinho.
    posterior = [{"numero_pc": "PC-9", "data_pc": date(2026, 12, 1), "valor_pc": 500.0}]
    assert janela_automatica_pcs(pedidos_de_itens_pc(posterior), CORTE) is None


def test_meses_sem_pc_dentro_da_janela_seguem_no_denominador():
    janela = janela_automatica_pcs(pedidos_de_itens_pc(_ITENS_PC), CORTE)
    base = media_pedidos_compra(pedidos_de_itens_pc(_ITENS_PC), CORTE, janela["meses"])
    assert base["pedidos_considerados"] == 3
    assert base["meses_com_pedido"] == 3
    assert base["meses_sem_pedido"] == 15
    assert base["media_mensal"] == pytest.approx(39000.0 / 18)


# ---------------------------------------------------------------------------
# Motor: projecao futura (corte != vigencia)
# ---------------------------------------------------------------------------

def test_projecao_vai_do_mes_seguinte_ao_corte_ate_a_vigencia():
    periodos = gerar_periodos_projecao(pd.Period("2026-08", freq="M"), VIGENCIA)
    assert len(periodos) == 18
    assert periodo_para_label(periodos[0]) == "set/26"
    assert periodo_para_label(periodos[-1]) == "fev/28"


def test_confundir_corte_com_vigencia_zerava_a_projecao():
    """Regressao: era isso que a competencia final digitada como 02/2028 fazia."""
    assert gerar_periodos_projecao(pd.Period("2028-02", freq="M"), VIGENCIA) == []


# ---------------------------------------------------------------------------
# CASO A — o bug real, ponta a ponta na pagina
# ---------------------------------------------------------------------------

def test_caso_a_pagina_usa_o_historico_real_e_projeta_ate_a_vigencia():
    at = _run(_sessao_pc())
    assert not at.exception, at.exception
    blob = _blob(at)
    assert "01/03/2025 a 31/08/2026" in blob
    assert "Automático — histórico disponível até o corte" in blob
    assert "01/12/2024" not in blob
    assert "ago/26" in blob                        # competencia final do historico
    assert "data de corte da apuração" in blob     # origem declarada
    assert "set/26 a fev/28" in "\n".join(str(e.label) for e in at.expander)


def test_caso_a_diferenca_futura_nao_fica_zerada():
    """Com premissa mensal explicita, as 18 competencias futuras produzem valor.

    A cadencia automatica continua fail-closed (Etapa 51B): historico esparso
    NAO vira mensalidade presumida. O que esta entrega corrige e a janela de
    projecao, que antes vinha vazia porque a competencia final do historico era
    a propria data final da vigencia.
    """
    at = _run(_sessao_pc())
    assert not at.exception, at.exception
    at.radio(key="adequacao_v3_premissa").set_value("Mensal (média)")
    at.run()
    assert not at.exception, at.exception
    blob = _blob(at)
    assert "18 meses" in blob
    assert _valor_do_card(blob, "Diferença futura projetada") > 0
    confirmada = _valor_do_card(blob, "COMPLEMENTAÇÃO CONFIRMADA")
    assert confirmada > RECONHECIDO


def test_caso_i_cadencia_esparsa_nao_vira_mensalidade_presumida():
    at = _run(_sessao_pc())
    assert not at.exception, at.exception
    blob = _blob(at)
    assert "Histórico sem periodicidade suficiente" in blob
    assert _valor_do_card(blob, "Diferença futura projetada") == 0


def test_caso_a_competencia_final_nao_e_redigitada():
    at = _run(_sessao_pc())
    assert not at.exception, at.exception
    assert "adequacao_v3_comp_ref_pc" not in at.session_state


# ---------------------------------------------------------------------------
# CASOS B / C / D — reconhecido x potencial
# ---------------------------------------------------------------------------

def test_caso_b_potencial_positivo_fica_separado_do_confirmado():
    at = _run(_sessao_pc(potencial=POTENCIAL))
    assert not at.exception, at.exception
    blob = _blob(at)
    assert "Retroativo reconhecido" in blob
    assert "Retroativo potencial" in blob
    assert "COMPLEMENTAÇÃO CONFIRMADA" in blob
    assert "Cenário com potencial" in blob
    assert "57.701,49" in blob
    assert "12.000,00" in blob


def test_caso_b_confirmada_e_cenario_diferem_exatamente_pelo_potencial():
    at = _run(_sessao_pc(potencial=POTENCIAL))
    assert not at.exception, at.exception
    blob = _blob(at)
    confirmada = _valor_do_card(blob, "COMPLEMENTAÇÃO CONFIRMADA")
    cenario = _valor_do_card(blob, "Cenário com potencial")
    reconhecido = _valor_do_card(blob, "Retroativo reconhecido considerado")
    assert reconhecido == pytest.approx(RECONHECIDO)
    assert cenario - confirmada == pytest.approx(POTENCIAL, abs=0.01)


def test_caso_c_potencial_zero_nao_gera_alarde():
    at = _run(_sessao_pc(potencial=0.0))
    assert not at.exception, at.exception
    blob = _blob(at)
    assert "Retroativo potencial" in blob
    confirmada = _valor_do_card(blob, "COMPLEMENTAÇÃO CONFIRMADA")
    cenario = _valor_do_card(blob, "Cenário com potencial")
    assert cenario == pytest.approx(confirmada)


def test_caso_d_potencial_ausente_nao_vira_zero():
    at = _run(_sessao_pc(com_potencial=False))
    assert not at.exception, at.exception
    blob = _blob(at)
    assert "Não localizado" in blob
    assert "Cenário com potencial" not in blob


# ---------------------------------------------------------------------------
# CASOS E / F — fallback manual e janela avancada
# ---------------------------------------------------------------------------

def test_caso_e_sem_data_de_corte_o_campo_manual_continua_disponivel():
    at = _run(_sessao_pc(data_corte=None))
    assert not at.exception, at.exception
    assert "adequacao_v3_comp_ref_pc" in [t.key for t in at.text_input]


def test_caso_f_janela_avancada_continua_disponivel_e_reversivel():
    at = _run(_sessao_pc())
    assert not at.exception, at.exception
    at.checkbox(key="adequacao_v3_janela_manual").set_value(True)
    at.run()
    at.slider(key="adequacao_v3_janela").set_value(6)
    at.run()
    assert not at.exception, at.exception
    blob = _blob(at)
    assert "Ajuste avançado — janela informada manualmente" in blob
    assert "01/03/2026 a 31/08/2026" in blob
    # Volta ao automatico: a janela derivada do historico e restaurada.
    at.checkbox(key="adequacao_v3_janela_manual").set_value(False)
    at.run()
    blob = _blob(at)
    assert "Automático — histórico disponível até o corte" in blob
    assert "01/03/2025 a 31/08/2026" in blob


# ---------------------------------------------------------------------------
# CASO H — metodo Financeiro sem regressao
# ---------------------------------------------------------------------------

def test_caso_h_financeiro_nao_expoe_medidas_de_pc():
    sessao = {
        "df_financeiro_mensal": pd.DataFrame({
            "Competência": ["01/2026", "02/2026", "03/2026",
                            "04/2026", "05/2026", "06/2026"],
            "Valor": [10000.0] * 6,
        }),
        "valor_represado_a_pagar": 16888.59,
        "variacao_acumulada": 0.1201,
        "modo_apuracao": "Completo",
    }
    at = _run(sessao, origem="Financeiro", vigencia="05/05/2027")
    assert not at.exception, at.exception
    blob = _blob(at)
    assert "Retroativo apurado" in blob
    assert "Retroativo potencial" not in blob
    assert "Cenário com potencial" not in blob
