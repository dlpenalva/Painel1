"""Prova de que a Adequacao usa o retroativo reconhecido canonico.

Cenario que motivou a correcao (auditoria P3A): no metodo PC, a cadeia bruta
valor_represado_a_pagar/delta_total (soma _retroativo_python por ciclo, sem
excluir PC posterior a data de corte ou sem efeito financeiro) pode divergir
do resultado_consolidado.retroativo_reconhecido (que usa
totais_canonicos_pc.ate_o_corte.retroativo, ja com essas exclusoes). A
Adequacao deve seguir o consolidado quando ele existir.
"""

from _adequacao_ui import extrair_contexto_valores


RECONHECIDO_CANONICO = 57_701.49
BRUTO_SEM_EXCLUSAO_DE_CORTE = 91_234.56


def _payload_pc_com_pc_excluido_pelo_corte():
    """Simula um PC posterior a data de corte: o consolidado (ate_o_corte)
    ja o excluiu, mas a cadeia bruta valor_represado_a_pagar/delta_total
    ainda o carrega."""
    return {
        "modo_apuracao": "Completo",
        "valor_represado_a_pagar": BRUTO_SEM_EXCLUSAO_DE_CORTE,
        "delta_total": BRUTO_SEM_EXCLUSAO_DE_CORTE,
        "resultado_consolidado": {
            "medidas_pc_aplicaveis": True,
            "retroativo_reconhecido": RECONHECIDO_CANONICO,
            "retroativo_potencial": 120_016.52,
        },
    }


def test_adequacao_usa_reconhecido_canonico_quando_diverge_da_cadeia_bruta():
    payload = _payload_pc_com_pc_excluido_pelo_corte()
    assert RECONHECIDO_CANONICO != BRUTO_SEM_EXCLUSAO_DE_CORTE

    ctx = extrair_contexto_valores(payload)

    assert ctx["valor_represado"] == RECONHECIDO_CANONICO


def test_potencial_permanece_intacto_e_independente_do_reconhecido():
    payload = _payload_pc_com_pc_excluido_pelo_corte()
    ctx = extrair_contexto_valores(payload)

    consolidado = ctx["resultado"]["resultado_consolidado"]
    assert consolidado["retroativo_potencial"] == 120_016.52
    assert ctx["valor_represado"] != consolidado["retroativo_potencial"]


def test_payload_legado_sem_resultado_consolidado_preserva_fallback_antigo():
    """Compatibilidade: payload sem resultado_consolidado (fixtures/testes
    legados existentes) continua usando valor_represado_a_pagar/delta_total."""
    payload = {
        "modo_apuracao": "Completo",
        "valor_represado_a_pagar": 16_888.59,
    }
    ctx = extrair_contexto_valores(payload)
    assert ctx["valor_represado"] == 16_888.59


def test_consolidado_presente_mas_sem_a_chave_preserva_fallback_antigo():
    """resultado_consolidado existe mas nao tem retroativo_reconhecido
    (payload legado parcial): mantem a cadeia antiga."""
    payload = {
        "modo_apuracao": "Completo",
        "valor_represado_a_pagar": 16_888.59,
        "resultado_consolidado": {"vta": 1_000.0},
    }
    ctx = extrair_contexto_valores(payload)
    assert ctx["valor_represado"] == 16_888.59
