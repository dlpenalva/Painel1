"""STATUS-CANON-1 — a aba RESULTADOS e a fonte canonica do status da apuracao.

Defeito corrigido (caso real ICTI, Coleta_Reajuste_C1_ICTI_27-08-2026):

    XLS  RESULTADOS!B3 = VALIDADO
    WEB  status         = PENDENTE DE CONFIRMACAO

A causa nao era divergencia material nenhuma: o metodo era PC, o ciclo em
analise era o C1, havia PC historico em C0 e nenhum PC em C1 — porque a
execucao do C1 foi legitimamente zero. `avaliar_entrega_segura` classificava
C1 em `ciclos_sem_evidencia`, nao encontrava metodo com evidencia no ciclo e
devolvia INFORMACAO_INSUFICIENTE; `montar_resultado_consolidado` promovia esse
status interno a PENDENTE DE CONFIRMACAO, substituindo silenciosamente a
conclusao oficial do XLS.

Contrato fixado aqui:

* STATUS DA APURACAO   -> espelha RESULTADOS!B3 (VALIDADO/ESTIMADO/REVISE);
* RESSALVAS            -> informam, nunca rebaixam o status oficial;
* FORMALIZACAO         -> eixo separado, pode bloquear com apuracao validada;
* ZERO REAL            -> valor apurado (0.0), jamais lacuna;
* AUSENCIA             -> None -> fail-closed, nunca VALIDADO fabricado.
"""

from __future__ import annotations

import pytest

from _resultado_consolidado import (
    STATUS_CONFIAVEL,
    STATUS_ESTIMADO,
    STATUS_PENDENTE,
    montar_resultado_consolidado,
)

_SEM_BLOCO_DE_STATUS = object()


def _diagnostico(status_oficial="VALIDADO"):
    metadados = {"ciclos_em_analise": ["C1"]}
    if status_oficial is not _SEM_BLOCO_DE_STATUS:
        metadados["status_resultados"] = {
            "geral": status_oficial,
            "valores": {"retroativo_oficial": 0, "vta_oficial": 5_240_971.67},
        }
    return {"metadados": metadados}


def _caso_icti(status_oficial="VALIDADO"):
    """Reproduz semanticamente o caso real: metodo PC, C1 com execucao zero.

    Espelha o payload observado no arquivo real — inclusive o veredito interno
    INFORMACAO_INSUFICIENTE da politica, que era exatamente o gatilho do
    defeito. Nenhum bloqueio, nenhuma divergencia, tudo conciliado.
    """
    return {
        "valor_atualizado_contrato": 5_240_971.67,
        "valor_represado_a_pagar": 0.0,
        "controle": {"modo": "pc", "ciclo_vigente": "C1", "data_corte": "28/08/2026"},
        "memoria_por_ciclo": {"vta": {"metodo": "pc"}},
        "totais_canonicos_pc": {
            "data_corte": "28/08/2026",
            # Execucao do C1 legitimamente igual a zero: valores APURADOS,
            # informados como 0.0 — nunca None.
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
        # Veredito interno da politica no caso real: nenhum PC no C1 fazia o
        # metodo do ciclo ficar indefinido.
        "politica_entrega_segura": {
            "status": "INFORMACAO_INSUFICIENTE",
            "pode_formalizar": False,
            "retroativo": {
                "metodo": None,
                "ciclos_sem_evidencia": ["C1"],
                "evidencias_por_metodo": {"financeiro": 0, "pc": 1, "consumidos": 0},
            },
            "pendencias": [
                "Sem evidência de execução registrada em C1; o fiscal deve "
                "confirmar que não houve execução ou complementar o XLS.",
            ],
        },
        "reconciliacao_xls_python": {"status_geral": "CONCILIADO"},
        "campos_nao_confiaveis_documentos": [],
        "formalizacao_bloqueada": False,
        "bloqueios_formalizacao": [],
        "diagnostico_coleta": _diagnostico(status_oficial),
    }


# ---------------------------------------------------------------------------
# CASO ICTI — o defeito que originou a etapa
# ---------------------------------------------------------------------------

def test_icti_ciclo_sem_pc_com_resultados_validado_nao_vira_pendente():
    consolidado = montar_resultado_consolidado(_caso_icti(), _diagnostico())

    assert consolidado["status_confiabilidade"] == STATUS_CONFIAVEL == "VALIDADO"
    assert consolidado["status_confiabilidade"] != "PENDENTE DE CONFIRMAÇÃO"
    assert consolidado["status_apuracao"]["codigo"] == "VALIDADO"
    assert consolidado["status_apuracao"]["origem"] == "resultados_xls"
    assert consolidado["formalizacao"]["bloqueada"] is False
    assert consolidado["formalizacao"]["status"] == "SEM BLOQUEIO"
    assert consolidado["bloqueios"] == []


def test_icti_preserva_as_grandezas_sem_recalcular():
    consolidado = montar_resultado_consolidado(_caso_icti(), _diagnostico())

    assert consolidado["vta"] == 5_240_971.67
    assert consolidado["retroativo_reconhecido"] == 0.0
    assert consolidado["retroativo_potencial"] == 0.0
    assert consolidado["metodo"]["codigo"] == "pc"
    assert consolidado["ciclo_vigente"] == "C1"
    assert consolidado["composicao_vta"]["exibivel"] is True


def test_icti_diagnostico_da_politica_sobrevive_como_ressalva():
    """O sinal de seguranca nao foi apagado — foi rebaixado ao lugar correto."""
    consolidado = montar_resultado_consolidado(_caso_icti(), _diagnostico())

    assert any("Sem evidência de execução" in r for r in consolidado["ressalvas"])
    # ...mas nao contamina o status oficial nem o eixo de formalizacao.
    assert consolidado["status_confiabilidade"] == STATUS_CONFIAVEL
    assert consolidado["status_apuracao"]["status_politica"] == (
        "INFORMACAO_INSUFICIENTE"
    )


# ---------------------------------------------------------------------------
# ZERO REAL != AUSENCIA
# ---------------------------------------------------------------------------

def test_zero_apurado_e_valor_e_nao_lacuna():
    consolidado = montar_resultado_consolidado(_caso_icti(), _diagnostico())

    for chave in ("retroativo_reconhecido", "retroativo_potencial"):
        assert consolidado[chave] == 0.0
        assert consolidado[chave] is not None


def test_ausencia_do_vta_continua_fail_closed_mesmo_com_resultados_validado():
    """VTA ausente nao pode ser apresentado como apuracao validada."""
    caso = _caso_icti()
    caso["valor_atualizado_contrato"] = None

    consolidado = montar_resultado_consolidado(caso, _diagnostico())

    assert consolidado["status_confiabilidade"] == STATUS_PENDENTE
    assert consolidado["vta"] is None


# ---------------------------------------------------------------------------
# TESTES NEGATIVOS — a correcao nao pode transformar tudo em VALIDADO
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "status_oficial",
    [None, "", "   ", "QUALQUER COISA"],
    ids=["ausente", "vazio", "espacos", "vocabulario-desconhecido"],
)
def test_sem_status_oficial_o_painel_nao_fabrica_validado(status_oficial):
    """A indisponibilidade vem da AUSENCIA DO STATUS OFICIAL, nao de PC."""
    caso = _caso_icti(status_oficial)

    consolidado = montar_resultado_consolidado(caso, _diagnostico(status_oficial))

    assert consolidado["status_confiabilidade"] == STATUS_PENDENTE
    assert consolidado["status_apuracao"]["disponivel"] is False
    assert consolidado["status_apuracao"]["origem"] == "indisponivel"
    assert "aba RESULTADOS" in consolidado["mensagem_status"]
    assert consolidado["formalizacao"]["status"] == "AGUARDA CONFIRMAÇÃO"


def test_bloco_de_status_totalmente_ausente_tambem_e_fail_closed():
    caso = _caso_icti(_SEM_BLOCO_DE_STATUS)

    consolidado = montar_resultado_consolidado(caso, {"metadados": {}})

    assert consolidado["status_confiabilidade"] == STATUS_PENDENTE
    assert consolidado["status_apuracao"]["disponivel"] is False


def test_revise_no_xls_permanece_pendente_no_painel():
    """Quando o proprio XLS pede revisao, o painel nao promove a VALIDADO."""
    consolidado = montar_resultado_consolidado(
        _caso_icti("REVISE"), _diagnostico("REVISE")
    )

    assert consolidado["status_confiabilidade"] == STATUS_PENDENTE
    assert consolidado["status_apuracao"]["codigo"] == "REVISE"
    assert consolidado["status_apuracao"]["conclusivo"] is False
    assert consolidado["formalizacao"]["status"] == "AGUARDA CONFIRMAÇÃO"


def test_estimado_no_xls_e_reproduzido_com_a_nomenclatura_oficial():
    consolidado = montar_resultado_consolidado(
        _caso_icti("ESTIMADO"), _diagnostico("ESTIMADO")
    )

    assert consolidado["status_confiabilidade"] == STATUS_ESTIMADO == "ESTIMADO"
    assert consolidado["formalizacao"]["status"] == "SEM BLOQUEIO"


def test_metodo_indeterminado_continua_pendente():
    """Sem metodo nao ha apuracao apresentavel."""
    caso = _caso_icti()
    caso["controle"]["modo"] = None
    caso["memoria_por_ciclo"] = {"vta": {"metodo": None}}

    consolidado = montar_resultado_consolidado(caso, _diagnostico())

    assert consolidado["metodo"]["codigo"] == "indeterminado"
    assert consolidado["status_confiabilidade"] == STATUS_PENDENTE


def test_divergencia_material_bloqueia_a_formalizacao_sem_negar_a_apuracao():
    """Bloqueio verdadeiro vive no eixo de formalizacao, com causa objetiva."""
    caso = _caso_icti()
    caso["formalizacao_bloqueada"] = True
    caso["bloqueios_formalizacao"] = [
        "Divergência relevante XLS × Python em VTA_FINAL; equalizar antes de formalizar."
    ]

    consolidado = montar_resultado_consolidado(caso, _diagnostico())

    assert consolidado["status_confiabilidade"] == STATUS_CONFIAVEL
    assert consolidado["formalizacao"]["bloqueada"] is True
    assert consolidado["formalizacao"]["status"] == "BLOQUEADA"
    assert "Divergência relevante" in consolidado["formalizacao"]["mensagem"]
    assert consolidado["bloqueios"] == caso["bloqueios_formalizacao"]


# ---------------------------------------------------------------------------
# OS TRES METODOS — PC nao pode contaminar Financeiro nem Consumido
# ---------------------------------------------------------------------------

def test_financeiro_nao_passa_a_exigir_pc():
    caso = _caso_icti()
    caso["controle"]["modo"] = "principal"
    caso["memoria_por_ciclo"] = {"vta": {"metodo": "financeiro"}}
    caso["valor_represado_a_pagar"] = 0.0

    consolidado = montar_resultado_consolidado(caso, _diagnostico())

    assert consolidado["metodo"]["rotulo"] == "Financeiro"
    assert consolidado["medidas_pc_aplicaveis"] is False
    assert consolidado["retroativo_potencial"] is None  # nao aplicavel, nao zero
    assert consolidado["retroativo_reconhecido"] == 0.0
    assert consolidado["status_confiabilidade"] == STATUS_CONFIAVEL


def test_consumido_nao_passa_a_exigir_pc_e_mantem_supressao_do_reconhecido():
    caso = _caso_icti()
    caso["controle"]["modo"] = "d"
    caso["memoria_por_ciclo"] = {"vta": {"metodo": "consumidos"}}

    consolidado = montar_resultado_consolidado(caso, _diagnostico())

    assert consolidado["metodo"]["rotulo"] == "Itens consumidos"
    assert consolidado["medidas_pc_aplicaveis"] is False
    # VTA-C2 preservado: supressao deliberada continua devolvendo ausencia...
    assert consolidado["retroativo_reconhecido"] is None
    # ...e agora aparece como ressalva, sem derrubar a apuracao inteira.
    assert any("Itens consumidos" in r for r in consolidado["ressalvas"])
    assert consolidado["status_confiabilidade"] == STATUS_CONFIAVEL
