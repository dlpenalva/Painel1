"""VTA-POT-1 — o retroativo POTENCIAL compoe o VTA no metodo PC.

Regra prudencial (somente Pedidos de Compra):

    VTA = execucao considerada (base + retroativo reconhecido)
        + remanescente atualizado do ciclo vigente
        + retroativo POTENCIAL dos MESMOS PCs

A parcela potencial e incorporada UMA UNICA VEZ e permanece identificada como
POTENCIAL. Financeiro e Itens Consumidos ficam inalterados.

Prova aritmetica que sustenta a ausencia de dupla contagem: em ``itens_PC``,
``H`` (RETROATIVO_RECONHECIDO_A_PAGAR) e ``J`` (DELTA_POTENCIAL) sao mutuamente
exclusivos por construcao — ``H`` so e diferente de zero quando ``G="Sim"`` e
``J`` so quando ``G<>"Sim"``. Logo, para cada PC:

    valor_pc + reconhecido + potencial = VALOR_ATUALIZADO do PC

Os cenarios abaixo sao os seis exigidos pela tarefa, sem excedentes.
"""
from __future__ import annotations

import _motor_composicao_vta as C
from _resultado_consolidado import montar_resultado_consolidado

# Base itemizada minima (2 itens) — o suficiente para o remanescente do ciclo
# vigente existir sem depender do caso real do usuario.
_ITENS = [
    (1, 100.0, 110.0, 10, 6, 4),   # (item, VU_C0, VU_C2, qtd_C0, qtd_C1, qtd_C2)
    (2, 200.0, 220.0, 5, 3, 1),
]

REMANESCENTE_C2 = round(110.0 * 4 + 220.0 * 1, 2)   # 660,00


def _leitura(pcs, *, modo="pc", vigente="C2"):
    posicao = [{
        "ITEM": item, "VU_ORIGINAL": vu0,
        "QTD_REM_AJUSTADA_C0": q0, "QTD_REM_AJUSTADA_C1": q1,
        "QTD_REM_AJUSTADA_C2": q2, "QTD_REM_AJUSTADA_C3": 0,
        "QTD_REM_AJUSTADA_C4": 0,
    } for (item, vu0, vu2, q0, q1, q2) in _ITENS]
    historico = [{
        "item": item,
        "vu_ciclos": {"VU_C0": vu0, "VU_C1": vu0, "VU_C2": vu2,
                      "VU_C3": vu2, "VU_C4": vu2},
    } for (item, vu0, vu2, q0, q1, q2) in _ITENS]
    return {
        "controle": {"modo": modo, "ciclo_vigente": vigente},
        "parametros_v10": {"por_ciclo": {}},
        "posicao_contratual": {"itens": posicao},
        "historico_vu": {"itens": historico},
        "itens_pc_v10": {"itens": list(pcs)},
    }


def _pc(ciclo, base, *, reconhecido=0.0, potencial=0.0, dentro_do_corte=True,
        entra="Sim"):
    """Um PC ja normalizado pelo leitor.

    ``valor_historico_considerado`` = base + reconhecido e exatamente o que o
    leitor grava (``itens_PC!U``); ``delta_potencial`` e ``itens_PC!J``.
    """
    return {
        "ciclo": ciclo,
        "valor_pc": base,
        "valor_atualizado": round(base + reconhecido + potencial, 2),
        "valor_historico_considerado": round(base + reconhecido, 2),
        "retroativo_reconhecido_a_pagar": reconhecido,
        "delta_potencial": potencial,
        "dentro_do_corte": dentro_do_corte,
        "entra_no_calculo": entra,
    }


def _soma_linhas(comp):
    return round(sum(l["valor_atualizado"] for l in comp["linhas"]), 2)


# --------------------------------------------------------------------------
# ASSERTIVA MATEMATICA OBRIGATORIA (secao 19) — vale para TODO cenario PC.
# --------------------------------------------------------------------------
def _provar_composicao(comp):
    """O VTA exibido e a soma das parcelas exibidas, sem sobreposicao."""
    assert comp["vta_composicao"] == _soma_linhas(comp), (
        "VTA divergiu da soma das parcelas exibidas"
    )
    assert comp["vta_composicao"] == round(
        comp["vta_sem_potencial"] + comp["retroativo_potencial_vta"], 2
    ), "vta_total != vta_sem_potencial + retroativo_potencial_vta"
    # Trava contra base + potencial + potencial: a parcela potencial aparece
    # em UMA unica linha, e nenhuma linha de execucao/saldo a carrega junto.
    potenciais = [l for l in comp["linhas"] if l.get("natureza") == "POTENCIAL"]
    assert len(potenciais) <= 1, "parcela potencial duplicada na composicao"
    if potenciais:
        assert potenciais[0]["valor_atualizado"] == comp["retroativo_potencial_vta"]
        assert potenciais[0]["tipo"] == "potencial"
    outras = [l for l in comp["linhas"] if l.get("natureza") != "POTENCIAL"]
    assert round(sum(l["valor_atualizado"] for l in outras), 2) == (
        comp["vta_sem_potencial"]
    ), "o potencial vazou para dentro de uma parcela economica"


# ============================ CENARIO 1 ====================================
# PC com potencial = 0 -> novo VTA identico ao VTA anterior.
def test_cenario1_sem_potencial_preserva_vta_anterior():
    pcs = [
        _pc("C0", 1000.0),
        _pc("C1", 500.0, reconhecido=50.0),
    ]
    comp = C.montar_composicao_vta(_leitura(pcs))
    assert comp["disponivel"] is True
    esperado = round(1000.0 + 550.0 + REMANESCENTE_C2, 2)
    assert comp["retroativo_potencial_vta"] == 0.0
    assert comp["tem_parcela_potencial"] is False
    assert comp["vta_sem_potencial"] == esperado
    assert comp["vta_composicao"] == esperado          # == comportamento main
    assert not [l for l in comp["linhas"] if l.get("natureza") == "POTENCIAL"]
    _provar_composicao(comp)


# ============================ CENARIO 2 ====================================
# PC com potencial > 0 -> o VTA reflete EXATAMENTE uma parcela potencial.
def test_cenario2_potencial_entra_uma_unica_vez():
    pcs = [
        _pc("C0", 1000.0),
        _pc("C1", 500.0, potencial=30.0),
    ]
    comp = C.montar_composicao_vta(_leitura(pcs))
    sem_potencial = round(1000.0 + 500.0 + REMANESCENTE_C2, 2)
    assert comp["vta_sem_potencial"] == sem_potencial
    assert comp["retroativo_potencial_vta"] == 30.0
    assert comp["tem_parcela_potencial"] is True
    assert comp["vta_composicao"] == round(sem_potencial + 30.0, 2)
    # A base do PC nao pago entra pelo valor original: o potencial NAO esta
    # embutido nela (senao a soma daria 30,00 a mais).
    execucao_c1 = next(l for l in comp["linhas"] if l["ciclo"] == "C1")
    assert execucao_c1["valor_atualizado"] == 500.0
    _provar_composicao(comp)


def test_cenario2_potencial_por_ciclo_identificado():
    pcs = [_pc("C1", 500.0, potencial=30.0), _pc("C0", 1000.0)]
    comp = C.montar_composicao_vta(_leitura(pcs))
    assert comp["potencial_por_ciclo"] == [{"ciclo": "C1", "valor": 30.0}]
    linha = next(l for l in comp["linhas"] if l.get("natureza") == "POTENCIAL")
    assert "POTENCIAL" in linha["descricao"]
    assert "criterio prudencial" in linha["observacao"]
    assert "nao representa" in linha["observacao"]


# ============================ CENARIO 3 ====================================
# Reconhecido > 0 E potencial > 0 -> ambos aparecem separados, sem sobreposicao.
def test_cenario3_reconhecido_e_potencial_separados():
    pcs = [
        _pc("C0", 1000.0),
        _pc("C1", 500.0, reconhecido=50.0),   # PC pago
        _pc("C1", 400.0, potencial=40.0),     # PC em analise
    ]
    comp = C.montar_composicao_vta(_leitura(pcs))
    # Execucao de C1 = 500 + 50 (reconhecido) + 400 (base do nao pago) = 950.
    execucao_c1 = next(l for l in comp["linhas"] if l["ciclo"] == "C1")
    assert execucao_c1["valor_atualizado"] == 950.0
    assert comp["retroativo_potencial_vta"] == 40.0
    sem_potencial = round(1000.0 + 950.0 + REMANESCENTE_C2, 2)
    assert comp["vta_sem_potencial"] == sem_potencial
    assert comp["vta_composicao"] == round(sem_potencial + 40.0, 2)
    _provar_composicao(comp)


def test_cenario3_soma_das_partes_fecha_com_valor_atualizado_dos_pcs():
    """base + reconhecido + potencial == VALOR_ATUALIZADO, uma unica vez."""
    pcs = [
        _pc("C0", 1000.0),
        _pc("C1", 500.0, reconhecido=50.0),
        _pc("C1", 400.0, potencial=40.0),
    ]
    comp = C.montar_composicao_vta(_leitura(pcs))
    elegiveis = [p for p in pcs if p["ciclo"] in ("C0", "C1")]
    assert comp["total_execucao_atualizada"] + comp["retroativo_potencial_vta"] == (
        round(sum(p["valor_atualizado"] for p in elegiveis), 2)
    )


# ============================ CENARIO 4 ====================================
# PC posterior ao corte / nao elegivel -> nao entra no potencial do VTA.
def test_cenario4_pc_posterior_ao_corte_fora_do_potencial():
    pcs = [
        _pc("C0", 1000.0),
        _pc("C1", 500.0, potencial=30.0),
        _pc("C1", 999.0, potencial=99.0, dentro_do_corte=False),
    ]
    comp = C.montar_composicao_vta(_leitura(pcs))
    assert comp["retroativo_potencial_vta"] == 30.0          # 99,00 fora
    assert comp["vta_sem_potencial"] == round(1000.0 + 500.0 + REMANESCENTE_C2, 2)
    _provar_composicao(comp)


def test_cenario4_pc_do_ciclo_vigente_fora_do_potencial():
    """Regra petrea do mesmo corte: se a base nao entra, o potencial tambem nao."""
    pcs = [
        _pc("C0", 1000.0),
        _pc("C2", 700.0, potencial=70.0),   # C2 e o ciclo vigente
    ]
    comp = C.montar_composicao_vta(_leitura(pcs))
    assert comp["retroativo_potencial_vta"] == 0.0
    assert comp["vta_sem_potencial"] == round(1000.0 + REMANESCENTE_C2, 2)
    _provar_composicao(comp)


def test_cenario4_pc_excluido_do_calculo_fora_do_potencial():
    pcs = [
        _pc("C0", 1000.0),
        _pc("C1", 500.0, potencial=30.0),
        _pc("C1", 800.0, potencial=80.0, entra="Nao"),
    ]
    comp = C.montar_composicao_vta(_leitura(pcs))
    assert comp["retroativo_potencial_vta"] == 30.0
    _provar_composicao(comp)


# ================= CENARIOS 5 e 6 — CONTROLES DE REGRESSAO =================
# Financeiro e Itens Consumidos: VTA identico a main, sem linha POTENCIAL.
def _leitura_outro_metodo(modo):
    return {
        "controle": {"modo": modo, "ciclo_vigente": "C2"},
        "parametros_v10": {"por_ciclo": {"C1": {"fator_acumulado": 1.1}}},
        "reconciliacao": {"registros": [{
            "ciclo": "C1", "valor_computado": 1000.0, "fonte_principal": "financeiro",
            "status_reconciliacao": "OK",
        }]},
        "potencial_futuro": {
            "saldo_remanescente_base": 500.0,
            "fator_vigente": 1.1,
            "valor_atualizado_vigente": 550.0,
            "ciclo_vigente": "C2",
        },
    }


def test_cenario5_financeiro_sem_parcela_potencial():
    comp = C.montar_composicao_vta(_leitura_outro_metodo("principal"))
    assert comp.get("metodo") == "financeiro"
    assert "retroativo_potencial_vta" not in comp
    assert not [l for l in comp.get("linhas") or []
                if l.get("natureza") == "POTENCIAL"]


def test_cenario6_consumidos_sem_parcela_potencial():
    comp = C.montar_composicao_vta(_leitura_outro_metodo("d"))
    assert "retroativo_potencial_vta" not in comp
    assert not [l for l in comp.get("linhas") or []
                if l.get("natureza") == "POTENCIAL"]
    # O caminho generico segue somando execucao + saldo, sem terceira parcela.
    assert comp["vta_composicao"] == _soma_linhas(comp)


# ================ FONTE UNICA DE VERDADE PARA OS CONSUMIDORES ==============
def _resultado(comp, *, modo="pc"):
    return {
        "controle": {"modo": modo, "ciclo_vigente": "C2"},
        "composicao_vta": comp,
        "valor_atualizado_contrato": comp.get("vta_composicao"),
        "totais_canonicos_pc": {"ate_o_corte": {
            "quantidade": 2, "retroativo": 50.0, "delta_potencial": 30.0,
            "valor_atualizado_em_analise": 530.0,
        }},
    }


def test_consolidado_publica_os_tres_conceitos_verificaveis():
    pcs = [_pc("C0", 1000.0), _pc("C1", 500.0, potencial=30.0)]
    comp = C.montar_composicao_vta(_leitura(pcs))
    consolidado = montar_resultado_consolidado(_resultado(comp), {})
    assert consolidado["tem_parcela_potencial"] is True
    assert consolidado["retroativo_potencial_vta"] == comp["retroativo_potencial_vta"]
    assert consolidado["vta_sem_potencial"] == comp["vta_sem_potencial"]
    assert consolidado["vta"] == round(
        consolidado["vta_sem_potencial"] + consolidado["retroativo_potencial_vta"], 2
    )
    frase = consolidado["frase_parcela_potencial"]
    assert "critério prudencial" in frase
    assert "não representa" in frase
    assert "retroativo reconhecido a pagar" in frase


def test_consolidado_nao_confunde_reconhecido_com_potencial():
    pcs = [_pc("C0", 1000.0), _pc("C1", 500.0, potencial=30.0)]
    comp = C.montar_composicao_vta(_leitura(pcs))
    consolidado = montar_resultado_consolidado(_resultado(comp), {})
    assert consolidado["retroativo_reconhecido"] == 50.0
    assert consolidado["retroativo_potencial_vta"] == 30.0
    assert consolidado["retroativo_reconhecido"] != consolidado["retroativo_potencial_vta"]


def test_consolidado_fora_do_metodo_pc_nao_publica_parcela_potencial():
    comp = C.montar_composicao_vta(_leitura_outro_metodo("principal"))
    consolidado = montar_resultado_consolidado(_resultado(comp, modo="principal"), {})
    assert consolidado["tem_parcela_potencial"] is False
    assert consolidado["retroativo_potencial_vta"] is None
    assert consolidado["vta_sem_potencial"] is None
    assert consolidado["frase_parcela_potencial"] == ""


# ============ DOCUMENTOS — QUADRO DO VTA COM A PARCELA POTENCIAL ============
# Deterministico: exercita a camada de apresentacao com o payload canonico ja
# pronto, sem depender do selo do XLS (RESULTADOS!H8), cuja indisponibilidade
# sem a posicao fisica e dívida PRE-EXISTENTE, alheia ao VTA-POT-1.
def _comp_pc_com_potencial():
    pcs = [_pc("C0", 1000.0), _pc("C1", 500.0, reconhecido=50.0),
           _pc("C1", 400.0, potencial=40.0)]
    return C.montar_composicao_vta(_leitura(pcs))


def test_sumario_expoe_componente_potencial_e_fecha_com_o_vta():
    from _sumario_executivo import _montar_composicao_vta

    comp = _comp_pc_com_potencial()
    quadro = _montar_composicao_vta(
        {"composicao_vta": comp}, {"vta": comp["vta_composicao"]}
    )
    assert quadro["exibivel"] is True
    assert quadro["total"] == comp["vta_composicao"]
    assert quadro["retroativo_potencial"] == comp["retroativo_potencial_vta"]
    assert quadro["sem_potencial"] == comp["vta_sem_potencial"]
    assert quadro["tem_parcela_potencial"] is True
    potenciais = [c for c in quadro["componentes"] if c["potencial"]]
    assert len(potenciais) == 1
    assert potenciais[0]["valor"] == comp["retroativo_potencial_vta"]
    # A soma dos componentes fecha ao centavo com o VTA exibido.
    assert round(sum(c["valor"] for c in quadro["componentes"]), 2) == quadro["total"]


def test_apostila_e_saneador_quebram_o_vta_em_tres_parcelas():
    from _templates_documentos import (
        ROTULO_PARCELA_POTENCIAL, _composicao_didatica_vta,
        _texto_parcela_potencial,
    )

    comp = _comp_pc_com_potencial()
    dados = {
        "vta": comp["vta_composicao"],
        "vta_execucao_atualizada": comp["total_execucao_atualizada"],
        "vta_saldo_remanescente_atualizado": (
            comp["saldo_remanescente"]["valor_atualizado"]
        ),
        "vta_retroativo_potencial": comp["retroativo_potencial_vta"],
        "vta_sem_potencial": comp["vta_sem_potencial"],
        "vta_tem_parcela_potencial": comp["tem_parcela_potencial"],
    }
    linhas = _composicao_didatica_vta(dados)
    rotulos = [d for d, _ in linhas]
    assert ROTULO_PARCELA_POTENCIAL in rotulos
    assert rotulos.index(ROTULO_PARCELA_POTENCIAL) == len(linhas) - 1
    assert round(sum(v for _, v in linhas), 2) == comp["vta_composicao"]
    valor_potencial = dict(linhas)[ROTULO_PARCELA_POTENCIAL]
    assert valor_potencial == comp["retroativo_potencial_vta"]

    frase = _texto_parcela_potencial(dados)
    assert "critério prudencial" in frase
    assert "não representa" in frase
    assert "retroativo reconhecido a pagar" in frase


def test_documentos_sem_potencial_mantem_duas_parcelas():
    from _templates_documentos import _composicao_didatica_vta

    pcs = [_pc("C0", 1000.0), _pc("C1", 500.0, reconhecido=50.0)]
    comp = C.montar_composicao_vta(_leitura(pcs))
    linhas = _composicao_didatica_vta({
        "vta": comp["vta_composicao"],
        "vta_execucao_atualizada": comp["total_execucao_atualizada"],
        "vta_saldo_remanescente_atualizado": (
            comp["saldo_remanescente"]["valor_atualizado"]
        ),
        "vta_retroativo_potencial": comp["retroativo_potencial_vta"],
        "vta_tem_parcela_potencial": comp["tem_parcela_potencial"],
    })
    assert len(linhas) == 2                      # == comportamento main
    assert round(sum(v for _, v in linhas), 2) == comp["vta_composicao"]
