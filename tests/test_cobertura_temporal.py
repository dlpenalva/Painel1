"""Etapa — Cobertura Temporal / Remanescente no ciclo em execucao.

Cenarios A-H (desenho original) + I-M (hotfix "ultima evidencia != cobertura
confirmada"), sobre o motor `_motor_cobertura_temporal` (puro Python, sem Excel).

Invariantes provados:
  * posicao FISICA nunca inventada pelo tempo transcorrido ("nao inventa junho");
  * MAX(data) e apenas ULTIMA EVIDENCIA — nunca "completo ate";
  * PC nao admite inferencia de completude (so confirmacao GCC);
  * Financeiro infere continuidade so com grade rigorosa (zero != vazio);
  * projecao fail-closed: a partir da cobertura ADOTADA, nunca da ultima evidencia;
  * fonte principal unica + demais conferencia => sem dupla contagem.
"""
from __future__ import annotations

from datetime import date

from _motor_cobertura_temporal import (
    montar_cobertura_temporal,
    MODO_POSICAO_DE_CORTE,
    MODO_POSICAO_ATUAL,
    MODO_FINANCEIRO_POSTERIOR,
    MODO_PC_POSTERIOR,
)

# Calendario base: C3 (ciclo em execucao) abre em 01/03/2026.
POR_CICLO = {
    "C0": {"data_inicio": date(2023, 1, 1), "data_fim": date(2023, 12, 31), "fator_acumulado": 1.0},
    "C1": {"data_inicio": date(2024, 1, 1), "data_fim": date(2024, 12, 31), "fator_acumulado": 1.10},
    "C2": {"data_inicio": date(2025, 1, 1), "data_fim": date(2025, 12, 31), "fator_acumulado": 1.21},
    "C3": {"data_inicio": date(2026, 3, 1), "data_fim": date(2027, 2, 28), "fator_acumulado": 1.33},
}
MARCO = date(2026, 3, 1)      # abertura do ciclo em execucao
MARCO_MAIS_1 = date(2026, 3, 2)
ANALISE = date(2026, 6, 30)   # data da analise


def _res(*, data_corte=None, itens_base=("ITEM-1",), remanescente=None,
         fotografias=("C1", "C2", "C3"), financeiro=None, itens_pc=None,
         confirmacao_gcc=None, ciclo_vigente="C3", data_analise=ANALISE):
    return {
        "controle": {
            "data_corte": data_corte,
            "ciclo_vigente": ciclo_vigente,
            "data_analise": data_analise,
        },
        "por_ciclo": POR_CICLO,
        "itens_base": list(itens_base),
        "remanescente_atual": dict(remanescente or {}),
        "fotografias_ciclo": list(fotografias),
        "financeiro": list(financeiro or []),
        "itens_pc": list(itens_pc or []),
        "confirmacao_gcc": dict(confirmacao_gcc or {}),
    }


# --------------------------------------------------------------------------- #
# CASO A — so fotografia fisica de marco; analise em junho.
# --------------------------------------------------------------------------- #
def test_caso_a_somente_fotografia_marco():
    r = montar_cobertura_temporal(_res())
    assert r.modo_temporal == MODO_POSICAO_DE_CORTE
    assert r.posicao_atual_completa is False
    assert r.posicao_fisica_conhecida_ate == MARCO
    assert r.posicao_observada["data"] == MARCO.isoformat()
    # Projecao autorizada a partir do dia seguinte a posicao fisica; nao junho.
    assert r.projecao_autorizada_a_partir_de == MARCO_MAIS_1
    assert r.posicao_observada["data"] != ANALISE.isoformat()


# --------------------------------------------------------------------------- #
# CASO B — foto marco + PC abril e maio.
# --------------------------------------------------------------------------- #
def test_caso_b_pc_posterior_nao_move_fisica():
    pcs = [
        {"numero_pc": "PC-1", "data_pc": date(2026, 4, 10), "valor_pc": 1000.0},
        {"numero_pc": "PC-2", "data_pc": date(2026, 5, 20), "valor_pc": 2000.0},
    ]
    r = montar_cobertura_temporal(_res(itens_pc=pcs))
    assert r.modo_temporal == MODO_PC_POSTERIOR
    assert r.pc_ultima_evidencia == date(2026, 5, 20)   # so ultima evidencia
    assert r.pc_cobertura_adotada_ate is None            # sem confirmacao GCC
    assert r.posicao_fisica_conhecida_ate == MARCO       # fisica NAO se move
    assert r.posicao_observada["data"] == MARCO.isoformat()


# --------------------------------------------------------------------------- #
# CASO C — foto marco + financeiro abril/maio.
# --------------------------------------------------------------------------- #
def test_caso_c_financeiro_posterior_nao_move_fisica():
    fin = [
        {"competencia": date(2026, 4, 1), "ciclo": "C3", "valor": 1500.0},
        {"competencia": date(2026, 5, 1), "ciclo": "C3", "valor": 1500.0},
    ]
    r = montar_cobertura_temporal(_res(financeiro=fin))
    assert r.modo_temporal == MODO_FINANCEIRO_POSTERIOR
    assert r.financeiro_ultima_evidencia == date(2026, 5, 1)
    assert r.posicao_fisica_conhecida_ate == MARCO
    assert r.fonte_principal == "financeiro"


# --------------------------------------------------------------------------- #
# CASO D — foto marco + nova fotografia fisica maio (completa).
# --------------------------------------------------------------------------- #
def test_caso_d_fotografia_recente_preserva_marco():
    r = montar_cobertura_temporal(_res(
        data_corte=date(2026, 5, 31), remanescente={"ITEM-1": 40.0}))
    assert r.modo_temporal == MODO_POSICAO_ATUAL
    assert r.posicao_atual_completa is True
    assert r.data_fotografia_recente == date(2026, 5, 31)
    assert r.posicao_fisica_conhecida_ate == date(2026, 5, 31)
    assert r.data_fotografia_corte == MARCO                  # abertura preservada
    assert r.ciclo_referencia == "C3"


# --------------------------------------------------------------------------- #
# CASO E — foto marco + PC abril/maio + foto maio (completa).
# --------------------------------------------------------------------------- #
def test_caso_e_hibrido_sem_dupla_contagem():
    pcs = [
        {"numero_pc": "PC-1", "data_pc": date(2026, 4, 10), "valor_pc": 1000.0},
        {"numero_pc": "PC-2", "data_pc": date(2026, 5, 20), "valor_pc": 2000.0},
    ]
    fin = [{"competencia": date(2026, 5, 1), "ciclo": "C3", "valor": 3000.0}]
    r = montar_cobertura_temporal(_res(
        data_corte=date(2026, 5, 31), remanescente={"ITEM-1": 40.0},
        itens_pc=pcs, financeiro=fin))
    assert r.posicao_fisica_conhecida_ate == date(2026, 5, 31)
    assert r.pc_ultima_evidencia == date(2026, 5, 20)
    assert r.financeiro_ultima_evidencia == date(2026, 5, 1)
    assert r.dupla_contagem_prevenida is True
    assert r.fontes_somadas == ("financeiro",)
    assert "pc" in r.fontes_conferencia


# --------------------------------------------------------------------------- #
# CASO F — foto marco + Financeiro e PC para a MESMA execucao.
# --------------------------------------------------------------------------- #
def test_caso_f_fonte_principal_unica():
    pcs = [{"numero_pc": "PC-1", "data_pc": date(2026, 4, 10), "valor_pc": 5000.0}]
    fin = [{"competencia": date(2026, 4, 1), "ciclo": "C3", "valor": 5000.0}]
    r = montar_cobertura_temporal(_res(itens_pc=pcs, financeiro=fin))
    assert r.fonte_principal == "financeiro"
    assert r.fontes_conferencia == ("pc",)
    assert r.fontes_somadas == ("financeiro",)   # PC nao soma
    assert r.dupla_contagem_prevenida is True


# --------------------------------------------------------------------------- #
# CASO G — foto marco apenas; analise junho.
# --------------------------------------------------------------------------- #
def test_caso_g_reconciliacao_pela_fotografia_disponivel():
    r = montar_cobertura_temporal(_res())
    rec = r.reconciliacao_fisica
    assert rec["reconciliado"] is True
    assert rec["data_posicao_utilizada"] == MARCO.isoformat()   # marco, nao junho
    assert rec["execucao_desde_corte"] == 0.0                    # fallback: exec 0
    assert r.posicao_observada["data"] == MARCO.isoformat()


# --------------------------------------------------------------------------- #
# CASO H — nova fotografia PARCIAL/incompleta.
# --------------------------------------------------------------------------- #
def test_caso_h_fotografia_parcial_fallback_global():
    r = montar_cobertura_temporal(_res(
        data_corte=date(2026, 5, 31),
        itens_base=("ITEM-1", "ITEM-2"),
        remanescente={"ITEM-1": 40.0},   # ITEM-2 sem quantidade -> incompleta
    ))
    assert r.posicao_atual_completa is False
    assert r.modo_temporal == MODO_POSICAO_DE_CORTE
    assert r.ciclo_referencia == "C3"
    assert r.posicao_fisica_conhecida_ate == MARCO   # abertura do ciclo, nao maio
    assert "INCOMPLETA" in r.origem_posicao


# --------------------------------------------------------------------------- #
# CASO I — PC sem confirmacao GCC.
# ultima evidencia PC = maio; cobertura completa NAO confirmada; projecao NAO
# comeca automaticamente em junho por causa desse PC.
# --------------------------------------------------------------------------- #
def test_caso_i_pc_sem_confirmacao_gcc():
    pcs = [
        {"numero_pc": "PC-1", "data_pc": date(2026, 4, 10), "valor_pc": 1000.0},
        {"numero_pc": "PC-2", "data_pc": date(2026, 5, 20), "valor_pc": 2000.0},
    ]
    r = montar_cobertura_temporal(_res(itens_pc=pcs))
    assert r.pc_ultima_evidencia == date(2026, 5, 20)
    assert r.pc_cobertura_confirmada_ate is None
    assert r.pc_cobertura_adotada_ate is None
    # projecao ancora na fisica (marco), NAO em junho por causa do PC de maio.
    assert r.projecao_autorizada_a_partir_de == MARCO_MAIS_1
    assert r.projecao_autorizada_a_partir_de.month == 3


# --------------------------------------------------------------------------- #
# CASO J — PC com confirmacao GCC ate 31/05.
# ultima evidencia = maio; confirmada = 31/05; projecao pode comecar 01/06.
# --------------------------------------------------------------------------- #
def test_caso_j_pc_com_confirmacao_gcc():
    pcs = [
        {"numero_pc": "PC-1", "data_pc": date(2026, 4, 10), "valor_pc": 1000.0},
        {"numero_pc": "PC-2", "data_pc": date(2026, 5, 20), "valor_pc": 2000.0},
    ]
    r = montar_cobertura_temporal(_res(
        itens_pc=pcs, confirmacao_gcc={"pc_ate": date(2026, 5, 31)}))
    assert r.pc_ultima_evidencia == date(2026, 5, 20)
    assert r.pc_cobertura_confirmada_ate == date(2026, 5, 31)
    assert r.pc_cobertura_adotada_ate == date(2026, 5, 31)
    assert r.projecao_autorizada_a_partir_de == date(2026, 6, 1)  # dia seguinte


# --------------------------------------------------------------------------- #
# CASO K — Financeiro continuo com zero.
# marco=10000, abril=0, maio=12000: abril conta como informado; inferida ate maio.
# --------------------------------------------------------------------------- #
def test_caso_k_financeiro_continuo_com_zero():
    fin = [
        {"competencia": date(2026, 3, 1), "ciclo": "C3", "valor": 10000.0},
        {"competencia": date(2026, 4, 1), "ciclo": "C3", "valor": 0.0},   # zero informado
        {"competencia": date(2026, 5, 1), "ciclo": "C3", "valor": 12000.0},
    ]
    r = montar_cobertura_temporal(_res(financeiro=fin))
    assert r.financeiro_ultima_evidencia == date(2026, 5, 1)
    assert r.financeiro_cobertura_inferida_ate == date(2026, 5, 1)   # continuidade ate maio
    assert r.financeiro_cobertura_adotada_ate == date(2026, 5, 1)


# --------------------------------------------------------------------------- #
# CASO L — Financeiro com lacuna.
# marco=10000, abril=VAZIO, maio=12000: nao infere ate maio; abril e lacuna;
# nao inicia projecao em junho apenas por MAX(maio).
# --------------------------------------------------------------------------- #
def test_caso_l_financeiro_com_lacuna():
    fin = [
        {"competencia": date(2026, 3, 1), "ciclo": "C3", "valor": 10000.0},
        # abril ausente (vazio != zero)
        {"competencia": date(2026, 5, 1), "ciclo": "C3", "valor": 12000.0},
    ]
    r = montar_cobertura_temporal(_res(financeiro=fin))
    assert r.financeiro_ultima_evidencia == date(2026, 5, 1)
    assert r.financeiro_cobertura_inferida_ate == date(2026, 3, 1)   # so ate marco
    assert r.financeiro_cobertura_adotada_ate == date(2026, 3, 1)
    # projecao nao salta para junho por causa do MAX(maio).
    assert r.projecao_autorizada_a_partir_de.month == 3


# --------------------------------------------------------------------------- #
# CASO M — confirmacao GCC supera lacuna conhecida.
# GCC confirma financeiro ate 31/05 apesar de abril vazio: adota confirmacao,
# mas preserva a lacuna como alerta/conferencia.
# --------------------------------------------------------------------------- #
def test_caso_m_confirmacao_gcc_supera_lacuna():
    fin = [
        {"competencia": date(2026, 3, 1), "ciclo": "C3", "valor": 10000.0},
        # abril ausente
        {"competencia": date(2026, 5, 1), "ciclo": "C3", "valor": 12000.0},
    ]
    r = montar_cobertura_temporal(_res(
        financeiro=fin, confirmacao_gcc={"financeiro_ate": date(2026, 5, 31)}))
    assert r.financeiro_cobertura_confirmada_ate == date(2026, 5, 31)
    assert r.financeiro_cobertura_adotada_ate == date(2026, 5, 31)  # GCC prevalece
    assert r.financeiro_cobertura_inferida_ate == date(2026, 3, 1)  # lacuna preservada
    assert any(a["codigo"] == "FINANCEIRO_LACUNA_SOB_CONFIRMACAO" for a in r.alertas)


# --------------------------------------------------------------------------- #
# Invariantes gerais
# --------------------------------------------------------------------------- #
def test_datas_de_fontes_nao_se_confundem():
    """As tres datas (fisica, financeiro, PC) sao reportadas separadamente."""
    pcs = [{"numero_pc": "PC-1", "data_pc": date(2026, 5, 20), "valor_pc": 1.0}]
    fin = [{"competencia": date(2026, 4, 1), "ciclo": "C3", "valor": 1.0}]
    r = montar_cobertura_temporal(_res(financeiro=fin, itens_pc=pcs))
    ue = r.ultima_evidencia_por_fonte
    assert ue["fisica"] == MARCO.isoformat()
    assert ue["financeiro"] == "2026-04-01"
    assert ue["pc"] == "2026-05-20"
    assert len({ue["fisica"], ue["financeiro"], ue["pc"]}) == 3


def test_pc_nunca_infere_completude():
    """MAX(DATA_PC) jamais vira cobertura adotada sem confirmacao GCC."""
    pcs = [{"numero_pc": "PC-1", "data_pc": date(2026, 5, 20), "valor_pc": 1.0}]
    r = montar_cobertura_temporal(_res(itens_pc=pcs))
    assert r.pc_ultima_evidencia == date(2026, 5, 20)
    assert r.pc_cobertura_adotada_ate is None


def test_indisponivel_quando_sem_foto_e_sem_atual():
    r = montar_cobertura_temporal(_res(fotografias=()))
    assert r.ciclo_referencia is None
    assert r.reconciliacao_fisica["reconciliado"] is False
    assert any(a["codigo"] == "POSICAO_FISICA_INDISPONIVEL" for a in r.alertas)
