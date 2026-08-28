from copy import deepcopy
from pathlib import Path

import pytest

from _resultado_consolidado import (
    STATUS_BLOQUEADO,
    STATUS_CONFIAVEL,
    STATUS_PENDENTE,
    STATUS_RESSALVAS,
    montar_resultado_consolidado,
)


ROOT = Path(__file__).resolve().parents[1]


def _resultado_base(metodo="pc", status_oficial="VALIDADO"):
    return {
        # STATUS-CANON-1: a conclusao oficial da aba RESULTADOS (B3) e a fonte
        # canonica do status apresentado. Sem ela o consolidado atua em
        # fail-closed, entao o payload base precisa carrega-la.
        "diagnostico_coleta": {
            "metadados": {"status_resultados": {"geral": status_oficial}}
        },
        "valor_atualizado_contrato": 1_000.0,
        "valor_represado_a_pagar": 125.0,
        "controle": {
            "modo": metodo,
            "ciclo_vigente": "C2",
            "data_corte": "31/12/2025",
        },
        "memoria_por_ciclo": {"vta": {"metodo": metodo}},
        "referencias_vta": {
            "posicao_atual_disponivel": True,
            "forma1_posicao_atual": 1_000.0,
            "forma2_ultima_abertura": 900.0,
        },
        "totais_canonicos_pc": {
            "data_corte": "31/12/2025",
            "ate_o_corte": {
                "retroativo": 125.0,
                "valor_atualizado_em_analise": 0.0,
                "delta_potencial": 0.0,
            },
            "posterior_ao_corte": {
                "quantidade": 0,
                "valor_pc": 0.0,
            },
        },
        "composicao_vta": {
            "disponivel": True,
            "bloqueia_formalizacao": False,
            "linhas": [
                {
                    "descricao": "Execução atualizada",
                    "ciclo": "C1",
                    "valor_base": 900.0,
                    "fator_acumulado": 1.1,
                    "valor_atualizado": 1_000.0,
                    "fonte": "pc",
                }
            ],
            "aditivos_nao_computados": [],
            "alertas": [],
        },
        "politica_entrega_segura": {
            "status": "PRONTO_PARA_VALIDACAO_FISCAL",
            "pendencias": [],
            "retroativo": {"metodo": metodo},
            "pode_formalizar": False,
        },
        "reconciliacao_xls_python": {"status_geral": "CONCILIADO"},
        "campos_nao_confiaveis_documentos": [],
        "formalizacao_bloqueada": False,
        "bloqueios_formalizacao": [],
    }


def test_caso_a_resultado_limpo_e_confiavel():
    resultado = _resultado_base()
    antes = deepcopy(resultado)

    consolidado = montar_resultado_consolidado(resultado)

    assert consolidado["status_confiabilidade"] == STATUS_CONFIAVEL
    assert consolidado["status_apuracao"]["codigo"] == "VALIDADO"
    assert consolidado["status_apuracao"]["origem"] == "resultados_xls"
    assert consolidado["vta"] == 1_000.0
    assert consolidado["retroativo_reconhecido"] == 125.0
    assert resultado == antes  # função pura: não altera nem recalcula a origem


def test_caso_b_potencial_gera_ressalva_sem_alterar_vta():
    resultado = _resultado_base()
    resultado["totais_canonicos_pc"]["ate_o_corte"]["delta_potencial"] = 40.0

    consolidado = montar_resultado_consolidado(resultado)

    # STATUS-CANON-1: ressalva nao rebaixa a conclusao oficial da apuracao.
    assert consolidado["status_confiabilidade"] == STATUS_CONFIAVEL
    assert consolidado["retroativo_potencial"] == 40.0
    assert consolidado["vta"] == resultado["valor_atualizado_contrato"] == 1_000.0
    assert any(
        "aceitação pela área gestora" in ressalva
        for ressalva in consolidado["ressalvas"]
    )


def test_caso_c_pc_posterior_aparece_somente_fora_do_corte():
    resultado = _resultado_base()
    resultado["totais_canonicos_pc"]["posterior_ao_corte"] = {
        "quantidade": 2,
        "valor_pc": 350.0,
        "retroativo": 99.0,
        "delta_potencial": 88.0,
    }

    consolidado = montar_resultado_consolidado(resultado)

    assert consolidado["status_confiabilidade"] == STATUS_CONFIAVEL
    assert "Há pedido(s) de compra posterior(es) à data de corte." in (
        consolidado["ressalvas"]
    )
    assert consolidado["fora_do_corte"] == {
        "aplicavel": True,
        "quantidade": 2,
        "valor_informado": 350.0,
        "data_corte": "31/12/2025",
    }
    assert consolidado["retroativo_reconhecido"] == 125.0
    assert consolidado["retroativo_potencial"] == 0.0
    assert consolidado["vta"] == 1_000.0


def test_caso_d_informacao_insuficiente_fica_pendente_sem_criar_bloqueio():
    resultado = _resultado_base()
    resultado["valor_atualizado_contrato"] = None
    resultado["politica_entrega_segura"]["status"] = "INFORMACAO_INSUFICIENTE"
    resultado["composicao_vta"] = {"disponivel": False, "linhas": []}

    consolidado = montar_resultado_consolidado(resultado)

    assert consolidado["status_confiabilidade"] == STATUS_PENDENTE
    assert consolidado["formalizacao"]["bloqueada"] is False
    assert consolidado["vta"] is None


def test_caso_e_bloqueio_real_tem_prioridade():
    resultado = _resultado_base()
    resultado["formalizacao_bloqueada"] = True
    resultado["bloqueios_formalizacao"] = ["Divergência bloqueadora já classificada."]

    consolidado = montar_resultado_consolidado(resultado)

    # STATUS-CANON-1: bloqueio e estado da FORMALIZACAO. A apuracao mantem a
    # conclusao oficial do XLS; o bloqueio aparece com causa objetiva ao lado.
    assert consolidado["status_confiabilidade"] == STATUS_CONFIAVEL
    assert consolidado["formalizacao"]["bloqueada"] is True
    assert consolidado["formalizacao"]["status"] == "BLOQUEADA"
    assert consolidado["formalizacao"]["mensagem"] == (
        "Divergência bloqueadora já classificada."
    )
    assert consolidado["bloqueios"] == ["Divergência bloqueadora já classificada."]


def test_caso_f_none_nao_e_zero():
    zero = _resultado_base()
    zero["valor_atualizado_contrato"] = 0.0
    zero["totais_canonicos_pc"]["ate_o_corte"]["retroativo"] = 0.0
    zero["valor_represado_a_pagar"] = None
    zero["composicao_vta"]["linhas"][0]["valor_atualizado"] = 0.0
    ausente = deepcopy(zero)
    ausente["valor_atualizado_contrato"] = None
    ausente["totais_canonicos_pc"]["ate_o_corte"]["retroativo"] = None

    consolidado_zero = montar_resultado_consolidado(zero)
    consolidado_ausente = montar_resultado_consolidado(ausente)

    assert consolidado_zero["vta"] == 0.0
    assert consolidado_zero["retroativo_reconhecido"] == 0.0
    assert consolidado_ausente["vta"] is None
    assert consolidado_ausente["retroativo_reconhecido"] is None
    assert consolidado_ausente["status_confiabilidade"] == STATUS_PENDENTE


@pytest.mark.parametrize("metodo", ["principal", "financeiro", "d", "consumidos"])
def test_caso_g_metodo_sem_pc_nao_aplica_medidas_pc(metodo):
    resultado = _resultado_base(metodo)
    resultado["totais_canonicos_pc"]["ate_o_corte"].update({
        "valor_atualizado_em_analise": 500.0,
        "delta_potencial": 200.0,
    })
    resultado["totais_canonicos_pc"]["posterior_ao_corte"] = {
        "quantidade": 3,
        "valor_pc": 700.0,
    }

    consolidado = montar_resultado_consolidado(resultado)

    assert consolidado["medidas_pc_aplicaveis"] is False
    assert consolidado["valor_atualizado_em_analise"] is None
    assert consolidado["retroativo_potencial"] is None
    assert consolidado["fora_do_corte"]["aplicavel"] is False
    assert consolidado["fora_do_corte"]["quantidade"] is None
    assert consolidado["fora_do_corte"]["valor_informado"] is None


def test_divergencia_so_bloqueia_quando_a_politica_produz_bloqueio_explicito():
    resultado = _resultado_base()
    resultado["reconciliacao_xls_python"]["status_geral"] = "DIVERGENCIA_RELEVANTE"

    consolidado = montar_resultado_consolidado(resultado)

    assert consolidado["status_confiabilidade"] == STATUS_CONFIAVEL
    assert "Há divergência relevante pendente de confirmação." in (
        consolidado["ressalvas"]
    )
    assert consolidado["formalizacao"]["bloqueada"] is False


def test_composicao_copia_linhas_sem_criar_total_concorrente():
    resultado = _resultado_base()
    consolidado = montar_resultado_consolidado(resultado)

    composicao = consolidado["composicao_vta"]
    assert composicao["linhas"] == resultado["composicao_vta"]["linhas"]
    assert "vta_composicao" not in composicao
    assert consolidado["vta"] == 1_000.0


def test_runtime_anexa_objeto_consolidado_e_pagina_preserva_cards_originais():
    runtime = (ROOT / "_coleta_reajuste_documentos.py").read_text(encoding="utf-8")
    pagina = (ROOT / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")

    assert 'resultado["resultado_consolidado"] = montar_resultado_consolidado' in runtime
    assert "render_resultado_consolidado(resultado, diagnostico_coleta)" in pagina
    assert "Resultado da apuração" in pagina
    assert "COMPOSIÇÃO DO VTA" in pagina
    cards = [
        'resumo_indice.metric("Índice"',
        'resumo_ciclos.metric("Ciclos analisados"',
        'resumo_retro.metric("Retroativo reconhecido"',
        'resumo_acum.metric("Percentual acumulado"',
    ]
    posicoes = [pagina.index(card) for card in cards]
    assert posicoes == sorted(posicoes)
    assert posicoes[-1] < pagina.index(
        "render_resultado_consolidado(resultado, diagnostico_coleta)"
    )
