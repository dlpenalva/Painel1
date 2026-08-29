"""STATUS-CANON-1.1 — execucao zero conhecida e informacao, nao pendencia.

Defeito corrigido (caso real ICTI, C1 recem-iniciado, metodo PC): o painel
exibia como ressalva

    "Sem evidencia de execucao registrada em C1; o fiscal deve confirmar que
     nao houve execucao ou complementar o XLS."

mesmo com a apuracao VALIDADA e SEM BLOQUEIO. Nao faltava informacao nenhuma:
a aba `itens_PC` estava preenchida (PC historico em C0) e simplesmente nenhum
PC caia em C1, porque o ciclo mal comecou. Execucao do C1 = R$ 0,00 apurado.

`avaliar_entrega_segura` jogava todo ciclo sem registro em `ciclos_sem_evidencia`
sem olhar se a FONTE de execucao do metodo estava legivel — confundindo

    ZERO CONHECIDO   (a fonte esta preenchida e nada cai neste ciclo)
com
    EXECUCAO DESCONHECIDA (a fonte esta vazia/ilegivel: nao da para saber).

Contrato fixado aqui:

* fonte de execucao do metodo legivel + ciclo sem registro -> INFORMACAO;
* fonte vazia/ilegivel/metodo indeterminado               -> ALERTA preservado;
* bloqueio material                                       -> ALERTA (fail-closed);
* status da apuracao e formalizacao NAO mudam por causa disto.
"""

from __future__ import annotations

import pytest

from _politica_entrega_segura import avaliar_entrega_segura
from _resultado_consolidado import (
    STATUS_CONFIAVEL,
    STATUS_PENDENTE,
    montar_resultado_consolidado,
)


PEDIDO_CONFIRMACAO = "o fiscal deve confirmar"
PEDIDO_COMPLEMENTO = "complementar o XLS"
SEM_METODO = "Nenhuma evidência financeira, PC pago definitivo ou consumo pago"


def _informacao(ciclo: str) -> str:
    return (
        f"Não houve execução registrada em {ciclo}. Para a apuração, foi "
        "considerado R$ 0,00 no ciclo."
    )


def _ciclo_memoria(nome: str, evidencias_pc: int = 0, itens_residuais: int = 6):
    zerado = {
        "base_original": 0.0,
        "valor_atualizado": 0.0,
        "retroativo": 0.0,
        "evidencias": 0,
    }
    return {
        "ciclo": nome,
        "retroativo": {
            "financeiro": dict(zerado),
            "pc": dict(
                zerado,
                evidencias=evidencias_pc,
                base_original=100.0 if evidencias_pc else 0.0,
                valor_atualizado=104.0 if evidencias_pc else 0.0,
                retroativo=4.0 if evidencias_pc else 0.0,
            ),
            "consumidos": dict(zerado),
        },
        "residuais": {
            "itens": itens_residuais,
            "valor_atualizado": 3_174_825.64 if itens_residuais else 0.0,
        },
    }


def _leitura(
    *,
    modo="pc",
    fonte_pc_preenchida=True,
    fonte_financeiro_preenchida=False,
    fonte_consumidos_preenchida=False,
    evidencias_pc_no_ciclo=0,
    ciclos=("C1",),
    advertencias=(),
    cache_posicao_ausente=False,
):
    """Leitura minima e controlada, com uma fonte de execucao por metodo."""
    por_ciclo = {
        nome: {
            "computar_nesta_apuracao": "Sim",
            "percentual_reajuste": 0.04052187881255853,
            "fator_acumulado": 1.0405218788125585,
        }
        for nome in ciclos
    }
    memoria_ciclos = [_ciclo_memoria("C0")] + [
        _ciclo_memoria(nome, evidencias_pc=evidencias_pc_no_ciclo) for nome in ciclos
    ]
    return {
        "ok": True,
        "controle": {"modo": modo, "ciclo_vigente": ciclos[-1]},
        # Fonte de execucao do PC: 1 PC historico legivel (o caso real ICTI).
        "itens_pc_v10": {
            "ok": fonte_pc_preenchida,
            "itens": [{"numero_pc": "4100138412", "valor_pc": 2_066_146.03}]
            if fonte_pc_preenchida
            else [],
        },
        # Fonte de execucao do Financeiro: `informado` separa zero de vazio.
        "financeiro": [
            {
                "competencia": "2026-01",
                "ciclo": "C0",
                "valor": 1_000.0,
                "informado": fonte_financeiro_preenchida,
            },
        ],
        "itens_consumidos_v10": {
            "ok": fonte_consumidos_preenchida,
            "itens": [{"item": 1}] if fonte_consumidos_preenchida else [],
        },
        "parametros_v10": {"por_ciclo": por_ciclo},
        "posicao_contratual": {"cache_ausente": cache_posicao_ausente},
        "objeto_processo": {
            "memoria_por_ciclo": {
                "ciclos": memoria_ciclos,
                "vta": {"valor_total_atualizado": 5_240_971.67, "metodo": modo},
            },
            "pendencias": {"advertencias": list(advertencias)},
        },
    }


# ---------------------------------------------------------------------------
# CENARIO A — ZERO LEGITIMO
# ---------------------------------------------------------------------------

def test_a_zero_legitimo_vira_informacao_e_nao_pendencia():
    politica = avaliar_entrega_segura(_leitura())

    assert politica["informacoes"] == [_informacao("C1")]
    assert politica["retroativo"]["ciclos_execucao_zero"] == ["C1"]
    assert politica["retroativo"]["ciclos_sem_evidencia"] == []
    texto = " | ".join(politica["pendencias"])
    assert PEDIDO_CONFIRMACAO not in texto
    assert PEDIDO_COMPLEMENTO not in texto


def test_a_zero_legitimo_nao_alega_falta_de_evidencia_para_o_retroativo():
    """Com execucao zero em todos os ciclos, nao ha lacuna a declarar."""
    politica = avaliar_entrega_segura(_leitura())

    assert not any(SEM_METODO in p for p in politica["pendencias"])
    assert politica["bloqueios"] == []


def test_a_o_ciclo_da_mensagem_e_dinamico():
    politica = avaliar_entrega_segura(_leitura(ciclos=("C1", "C2")))

    assert politica["informacoes"] == [_informacao("C1"), _informacao("C2")]


# ---------------------------------------------------------------------------
# CENARIO B — AUSENCIA REAL DE INFORMACAO
# ---------------------------------------------------------------------------

def test_b_fonte_de_execucao_vazia_preserva_o_alerta():
    politica = avaliar_entrega_segura(_leitura(fonte_pc_preenchida=False))

    assert politica["informacoes"] == []
    assert politica["retroativo"]["ciclos_sem_evidencia"] == ["C1"]
    texto = " | ".join(politica["pendencias"])
    assert PEDIDO_CONFIRMACAO in texto
    assert PEDIDO_COMPLEMENTO in texto


def test_b_metodo_indeterminado_preserva_o_alerta():
    """Sem saber o metodo nao se sabe qual fonte olhar: fail-closed."""
    politica = avaliar_entrega_segura(_leitura(modo=""))

    assert politica["informacoes"] == []
    assert politica["retroativo"]["ciclos_sem_evidencia"] == ["C1"]
    assert any(PEDIDO_CONFIRMACAO in p for p in politica["pendencias"])


def test_b_bloqueio_material_devolve_o_ciclo_para_o_alerta():
    """Zero conhecido exige apuracao integra; com bloqueio a duvida volta."""
    politica = avaliar_entrega_segura(_leitura(cache_posicao_ausente=True))

    assert politica["bloqueios"]  # cache da posicao contratual ausente
    assert politica["informacoes"] == []
    assert politica["retroativo"]["ciclos_execucao_zero"] == []
    assert politica["retroativo"]["ciclos_sem_evidencia"] == ["C1"]
    assert any(PEDIDO_CONFIRMACAO in p for p in politica["pendencias"])


# ---------------------------------------------------------------------------
# CENARIO C — EXECUCAO POSITIVA
# ---------------------------------------------------------------------------

def test_c_execucao_positiva_nao_gera_mensagem_de_zero():
    politica = avaliar_entrega_segura(_leitura(evidencias_pc_no_ciclo=2))

    assert politica["informacoes"] == []
    assert politica["retroativo"]["ciclos_execucao_zero"] == []
    assert politica["retroativo"]["ciclos_sem_evidencia"] == []
    assert politica["retroativo"]["metodo"] == "pc"
    assert not any(PEDIDO_CONFIRMACAO in p for p in politica["pendencias"])


# ---------------------------------------------------------------------------
# CENARIO D — OUTRA RESSALVA REAL SOBREVIVE
# ---------------------------------------------------------------------------

def test_d_outra_ressalva_real_permanece_intacta():
    """VALIDADO nunca e motivo generico para apagar ressalvas verdadeiras."""
    politica = avaliar_entrega_segura(
        _leitura(advertencias=["Há aditivo pendente de conferência."])
    )

    assert "Há aditivo pendente de conferência." in politica["pendencias"]
    assert politica["informacoes"] == [_informacao("C1")]


# ---------------------------------------------------------------------------
# CENARIOS E e F — Financeiro e Consumido preservados
# ---------------------------------------------------------------------------

def test_e_financeiro_usa_a_propria_fonte_e_nao_exige_pc():
    zero_conhecido = avaliar_entrega_segura(
        _leitura(
            modo="principal",
            fonte_pc_preenchida=False,
            fonte_financeiro_preenchida=True,
        )
    )
    assert zero_conhecido["informacoes"] == [_informacao("C1")]

    ausencia = avaliar_entrega_segura(
        _leitura(
            modo="principal",
            fonte_pc_preenchida=True,
            fonte_financeiro_preenchida=False,
        )
    )
    # PC preenchido nao vale como execucao do Financeiro.
    assert ausencia["informacoes"] == []
    assert any(PEDIDO_CONFIRMACAO in p for p in ausencia["pendencias"])


def test_f_consumido_usa_a_propria_fonte_e_nao_exige_pc():
    zero_conhecido = avaliar_entrega_segura(
        _leitura(
            modo="d",
            fonte_pc_preenchida=False,
            fonte_consumidos_preenchida=True,
        )
    )
    assert zero_conhecido["informacoes"] == [_informacao("C1")]

    ausencia = avaliar_entrega_segura(
        _leitura(
            modo="d",
            fonte_pc_preenchida=True,
            fonte_consumidos_preenchida=False,
        )
    )
    assert ausencia["informacoes"] == []
    assert any(PEDIDO_CONFIRMACAO in p for p in ausencia["pendencias"])


# ---------------------------------------------------------------------------
# CONSOLIDADO — informacao separada da ressalva, sem tocar status/formalizacao
# ---------------------------------------------------------------------------

def _consolidado(status_oficial="VALIDADO", **kwargs):
    politica = avaliar_entrega_segura(_leitura(**kwargs))
    diagnostico = {
        "metadados": {"status_resultados": {"geral": status_oficial, "valores": {}}}
    }
    modo = kwargs.get("modo", "pc")
    resultado = {
        "valor_atualizado_contrato": 5_240_971.67,
        "controle": {"modo": modo, "ciclo_vigente": "C1"},
        "memoria_por_ciclo": {"vta": {"metodo": modo}},
        "totais_canonicos_pc": {
            "ate_o_corte": {
                "retroativo": 0.0,
                "valor_atualizado_em_analise": 0.0,
                "delta_potencial": 0.0,
            },
            "posterior_ao_corte": {"quantidade": 0, "valor_pc": 0.0},
        },
        "composicao_vta": {
            "disponivel": True,
            "bloqueia_formalizacao": False,
            "linhas": [
                {
                    "descricao": "Execução C0 atualizada",
                    "ciclo": "C0",
                    "valor_base": 2_066_146.03,
                    "fator_acumulado": 1.0405218788125585,
                    "valor_atualizado": 2_066_146.03,
                    "fonte": "pc",
                }
            ],
            "aditivos_nao_computados": [],
            "alertas": [],
        },
        "politica_entrega_segura": politica,
        "reconciliacao_xls_python": {"status_geral": "CONCILIADO"},
        "formalizacao_bloqueada": bool(politica["bloqueios"]),
        "bloqueios_formalizacao": list(politica["bloqueios"]),
    }
    return montar_resultado_consolidado(resultado, diagnostico)


def test_consolidado_expoe_informacao_fora_das_ressalvas():
    consolidado = _consolidado()

    assert consolidado["informacoes"] == [_informacao("C1")]
    assert consolidado["ressalvas"] == []
    # STATUS-CANON-1 preservado integralmente.
    assert consolidado["status_confiabilidade"] == STATUS_CONFIAVEL == "VALIDADO"
    assert consolidado["formalizacao"]["status"] == "SEM BLOQUEIO"
    assert consolidado["vta"] == 5_240_971.67
    assert consolidado["retroativo_reconhecido"] == 0.0
    assert consolidado["retroativo_potencial"] == 0.0


@pytest.mark.parametrize(
    "status_oficial", ["REVISE", None], ids=["revise", "indisponivel"]
)
def test_consolidado_sem_conclusao_oficial_devolve_o_fato_para_ressalva(status_oficial):
    """A ultima condicao do contrato: resultado oficial disponivel."""
    consolidado = _consolidado(status_oficial)

    assert consolidado["status_confiabilidade"] == STATUS_PENDENTE
    assert consolidado["informacoes"] == []
    assert _informacao("C1") in consolidado["ressalvas"]


def test_consolidado_preserva_ressalva_real_ao_lado_da_informacao():
    consolidado = _consolidado(advertencias=["Há aditivo pendente de conferência."])

    assert consolidado["informacoes"] == [_informacao("C1")]
    assert "Há aditivo pendente de conferência." in consolidado["ressalvas"]
    assert consolidado["status_confiabilidade"] == STATUS_CONFIAVEL
