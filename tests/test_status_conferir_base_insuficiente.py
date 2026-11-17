"""Regressao Cl8us 12.16 / Coleta 12.8 — status CONFERIR e BASE INSUFICIENTE.

Defeito 1: RESULTADOS_DETALHE!B3 (Coleta 12.7, PR #199) passou a emitir
CONFERIR e BASE INSUFICIENTE, mas o mapa de status do Python so conhecia
VALIDADO / VALIDADO COM RESSALVAS / ESTIMADO / REVISE. Um arquivo recalculado
no Excel aparecia como "status oficial indisponivel... recalcule e reenvie" e a
formalizacao como AGUARDA CONFIRMACAO (caso real 1-luc.xlsx).

Defeito 2: com o mesmo status, RESULTADOS!C13 continuava dizendo "Nenhuma
pendencia registrada pelas validacoes desta apuracao".

Os valores (VTA, retroativo reconhecido/potencial/em analise) nao dependem do
status: os testes comparam cada metodo com o mesmo payload em VALIDADO. A
igualdade exata da formula de C13 com o template e com a Coleta gerada fica com
tests/test_resultados_retro_reconhecido_potencial.py (descritor).
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from _conferencias_execucao import FORMULAS_MENSAGENS_RESULTADOS
from _resultado_consolidado import (
    ORIGEM_STATUS_RESULTADOS,
    STATUS_BASE_INSUFICIENTE,
    STATUS_CONFERIR,
    STATUS_CONFIAVEL,
    STATUS_PENDENTE,
    montar_resultado_consolidado,
)

RAIZ = Path(__file__).resolve().parents[1]
STATUS_CONFERENCIA = (STATUS_CONFERIR, STATUS_BASE_INSUFICIENTE)
METODOS = ("financeiro", "pc", "consumidos")
_MODO_CONTROLE = {"financeiro": "financeiro", "pc": "pc", "consumidos": "d"}


def _payload(metodo: str, status: str, **extra) -> dict:
    resultado = {
        "diagnostico_coleta": {
            "metadados": {
                "status_resultados": {
                    "geral": status,
                    "valores": {"retroativo_oficial": 125.0},
                },
                "formula_status_resultados_presente": True,
            }
        },
        "valor_atualizado_contrato": 1_000.0,
        "valor_represado_a_pagar": 125.0 if metodo != "consumidos" else None,
        "controle": {"modo": _MODO_CONTROLE[metodo], "ciclo_vigente": "C1"},
        "memoria_por_ciclo": {"vta": {"metodo": metodo}},
        "totais_canonicos_pc": {
            "ate_o_corte": {
                "retroativo": 125.0,
                "valor_atualizado_em_analise": 40.0,
                "delta_potencial": 0.0,
            },
            "posterior_ao_corte": {"quantidade": 0, "valor_pc": 0.0},
        },
        "politica_entrega_segura": {
            "status": "PRONTO_PARA_VALIDACAO_FISCAL",
            "pode_confirmar": True,
            "pendencias": [],
            "bloqueios": [],
        },
        "reconciliacao_xls_python": {"status_geral": "CONCILIADO"},
        "formalizacao_bloqueada": False,
        "bloqueios_formalizacao": [],
    }
    resultado.update(extra)
    return resultado


def _valores(consolidado: dict) -> tuple:
    return (
        consolidado["vta"],
        consolidado["retroativo_reconhecido"],
        consolidado["retroativo_potencial"],
        consolidado["valor_atualizado_em_analise"],
    )


@pytest.mark.parametrize("metodo", METODOS)
@pytest.mark.parametrize("status", STATUS_CONFERENCIA)
def test_status_de_conferencia_e_reconhecido_sem_virar_outro(status, metodo):
    c = montar_resultado_consolidado(_payload(metodo, status))
    apuracao = c["status_apuracao"]
    assert apuracao["codigo"] == status
    assert apuracao["rotulo"] == c["status_confiabilidade"] == status
    assert apuracao["disponivel"] is True and apuracao["conclusivo"] is True
    assert apuracao["origem"] == ORIGEM_STATUS_RESULTADOS
    # Nao e VALIDADO, nao e REVISE (PENDENTE) e nao e indisponibilidade.
    assert c["status_confiabilidade"] not in (STATUS_CONFIAVEL, STATUS_PENDENTE)
    assert "não está disponível" not in c["mensagem_status"]
    assert "recalcule" not in c["mensagem_status"]
    assert status in c["mensagem_status"]
    assert any("não altera o VTA nem bloqueia" in r for r in c["ressalvas"])
    assert c["formalizacao"] == {
        "bloqueada": False,
        "status": "SEM BLOQUEIO",
        "mensagem": "Sem bloqueio explícito à formalização.",
    }


@pytest.mark.parametrize("metodo", METODOS)
@pytest.mark.parametrize("status", STATUS_CONFERENCIA)
def test_valores_por_metodo_nao_dependem_do_status(status, metodo):
    validado = montar_resultado_consolidado(_payload(metodo, "VALIDADO"))
    conferencia = montar_resultado_consolidado(_payload(metodo, status))
    assert _valores(conferencia) == _valores(validado)


def test_retroativo_reconhecido_por_metodo_preservado():
    por_metodo = {
        m: montar_resultado_consolidado(_payload(m, STATUS_CONFERIR))
        for m in METODOS
    }
    assert por_metodo["financeiro"]["retroativo_reconhecido"] == 125.0
    assert por_metodo["pc"]["retroativo_reconhecido"] == 125.0
    assert por_metodo["pc"]["valor_atualizado_em_analise"] == 40.0
    # VTA-C2: Consumidos continua sem fonte independente de pagamento.
    assert por_metodo["consumidos"]["retroativo_reconhecido"] is None
    assert por_metodo["financeiro"]["retroativo_potencial"] is None


@pytest.mark.parametrize("status", STATUS_CONFERENCIA)
def test_zero_legitimo_continua_zero(status):
    payload = _payload("financeiro", status, valor_represado_a_pagar=0.0)
    c = montar_resultado_consolidado(payload)
    assert c["retroativo_reconhecido"] == 0.0
    assert c["retroativo_reconhecido"] is not None


@pytest.mark.parametrize("status", STATUS_CONFERENCIA)
def test_ausencia_continua_ausencia(status):
    payload = _payload("financeiro", status, valor_represado_a_pagar=None)
    payload["diagnostico_coleta"]["metadados"]["status_resultados"]["valores"] = {}
    c = montar_resultado_consolidado(payload)
    assert c["retroativo_reconhecido"] is None
    # Status reconhecido nao fabrica VTA: sem VTA o nucleo fica indisponivel.
    sem_vta = montar_resultado_consolidado(
        _payload("financeiro", status, valor_atualizado_contrato=None)
    )
    assert sem_vta["status_confiabilidade"] == STATUS_PENDENTE
    assert sem_vta["status_apuracao"]["conclusivo"] is False


@pytest.mark.parametrize("metodo", METODOS)
@pytest.mark.parametrize("status", STATUS_CONFERENCIA)
def test_bloqueio_economico_preservado(status, metodo):
    motivo = "PC sem início de efeito financeiro: formalização bloqueada."
    payload = _payload(
        metodo, status,
        formalizacao_bloqueada=True, bloqueios_formalizacao=[motivo],
    )
    c = montar_resultado_consolidado(payload)
    assert c["formalizacao"]["bloqueada"] is True
    assert c["formalizacao"]["status"] == "BLOQUEADA"
    assert c["formalizacao"]["mensagem"] == motivo
    assert c["bloqueios"] == [motivo]
    # O eixo da apuracao continua sendo a conclusao do XLS.
    assert c["status_confiabilidade"] == status


def test_revise_e_texto_desconhecido_continuam_fail_closed():
    revise = montar_resultado_consolidado(_payload("financeiro", "REVISE"))
    assert revise["status_confiabilidade"] == STATUS_PENDENTE
    assert revise["status_apuracao"]["codigo"] == "REVISE"
    for texto in ("CONFERIR PARCIAL", "BASE", "EM CONFERÊNCIA"):
        c = montar_resultado_consolidado(_payload("financeiro", texto))
        assert c["status_apuracao"]["disponivel"] is False, texto
        assert c["status_confiabilidade"] == STATUS_PENDENTE, texto


def test_painel_tem_classe_visual_de_ressalva_para_os_dois_status():
    pagina = (RAIZ / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")
    assert '"CONFERIR": "ressalvas"' in pagina
    assert '"BASE INSUFICIENTE": "ressalvas"' in pagina


def test_c13_so_troca_o_texto_quando_o_status_pede_conferencia():
    formula = FORMULAS_MENSAGENS_RESULTADOS["C13"]
    # Gate: apuracao sem pendencias (regra de A7) E status de conferencia.
    assert formula.startswith(
        '=IF(AND(LEFT(RESULTADOS_DETALHE!$A$7,4)="APUR",'
        'OR(STATUS_RESULTADOS="CONFERIR",STATUS_RESULTADOS="BASE INSUFICIENTE")),'
    )
    # Fora do gate, o texto anterior (T67) sai intacto.
    assert formula.endswith(",MEMORIA_RESULTADOS!$T$67)")
    # O texto de potencial (T67 diferente de B7) continua anexado.
    assert ('&IF(MEMORIA_RESULTADOS!$T$67=RESULTADOS_DETALHE!$B$7,"",'
            '" "&MEMORIA_RESULTADOS!$T$67)') in formula
    assert "Nenhuma pendência registrada" not in formula
    literais = re.findall(r'"((?:[^"]|"")*)"', formula)
    assert all(len(t) <= 255 for t in literais)  # limite do Excel por literal
    assert sum("não altera o VTA nem bloqueia a formalização" in t
               for t in literais) == 2
    sem_literais = re.sub(r'"(?:[^"]|"")*"', '""', formula)
    assert sem_literais.isascii()
    assert sem_literais.count("(") == sem_literais.count(")")
