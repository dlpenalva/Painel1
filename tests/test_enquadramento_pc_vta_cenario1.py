# -*- coding: utf-8 -*-
"""Enquadramento temporal dos PCs, valor historico considerado e VTA (Cenario 1).

Cobre os itens obrigatorios da correcao do metodo PC:

  1  total informado de PCs (inventario integral do arquivo);
  2  total enquadrado nos ciclos;
  3  total com efeito financeiro;
  4  PCs dentro do ciclo antes do inicio do efeito (retroativo zero);
  5  PCs no intervalo precluso entre dois ciclos conhecidos;
  6  totais financeiros finais (considerado e retroativo);
  7  PC posterior a data de corte;
  8  posicao fisica completa prevalece sobre estimativa por PCs;
  9  aditivo do meio de C1 nao altera a abertura de C1;
 10  supressao do meio de C2 nao altera a abertura de C2;
 11  o VTA inclui a execucao do intervalo, sem reajuste;
 12  PCs sem efeito nao recebem fator efetivo no VTA;
 16  modelos sem intervalo entre ciclos nao regridem;
 17  ciclos contiguos preservam o comportamento anterior;
 18  os metodos Financeiro e Itens nao sao afetados.

Os cenarios sao sinteticos e parametrizados: nenhum valor do enunciado e
gravado como constante de calculo — os totais esperados sao derivados dos
proprios fatores da linha temporal montada aqui.
"""
from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from _leitor_masterfile_v10 import (  # noqa: E402
    _materializar_aberturas_temporais,
    _pc_dentro_do_corte,
    _totais_canonicos_pc,
)
from _motor_composicao_vta import montar_composicao_vta  # noqa: E402
from _motor_temporal import (  # noqa: E402
    ENQ_CICLO,
    ENQ_INDETERMINADO,
    ENQ_INTERVALO_PRECLUSO,
    classificar_enquadramento_pc,
)

# --------------------------------------------------------------------------- #
# Linha temporal do Cenario 1: C1 fecha em 31/12/2025 e C2 so abre em 01/03/2026
# --------------------------------------------------------------------------- #
POR_CICLO_CANONICO = {
    "C0": {"data_inicio": date(2024, 1, 1), "data_fim": date(2024, 12, 31),
           "fator_acumulado": 1.0, "computar_nesta_apuracao": "Nao"},
    "C1": {"data_inicio": date(2025, 1, 1), "data_fim": date(2025, 12, 31),
           "fator_acumulado": 1.0449137427397692,
           "computar_nesta_apuracao": "Sim",
           "inicio_efeito_financeiro": date(2025, 3, 1)},
    "C2": {"data_inicio": date(2026, 1, 1), "data_fim": date(2026, 12, 31),
           "fator_acumulado": 1.0778978351973156,
           "computar_nesta_apuracao": "Sim",
           "inicio_efeito_financeiro": date(2026, 3, 1)},
    "C3": {"data_inicio": date(2027, 1, 1), "data_fim": date(2027, 12, 31),
           "fator_acumulado": 1.0778978351973156},
    "C4": {"data_inicio": date(2028, 1, 1), "data_fim": date(2028, 12, 31),
           "fator_acumulado": 1.0778978351973156},
}

# Calendario CORROMPIDO (a regressao que esta correcao elimina): C2
# materializado a partir do inicio do efeito financeiro, abrindo lacuna de
# jan-fev/2026 e deslocando C3 e C4.
POR_CICLO_LACUNA = {
    **{c: dict(r) for c, r in POR_CICLO_CANONICO.items()},
    "C2": {**POR_CICLO_CANONICO["C2"], "data_inicio": date(2026, 3, 1),
           "data_fim": date(2027, 2, 28)},
}

# Alias historico: os testes de contiguidade agora usam o proprio canonico.
POR_CICLO_CONTIGUO = POR_CICLO_CANONICO

VALOR_PC = 10000.0


def _fator(ciclo: str, por_ciclo=POR_CICLO_CANONICO) -> float:
    return float(por_ciclo[ciclo]["fator_acumulado"])


def _pc(numero: str, data_pc: date, por_ciclo=POR_CICLO_CANONICO,
        valor: float = VALOR_PC, data_corte: date | None = None) -> dict:
    """Monta um registro de PC ja com as medidas canonicas, como o leitor faz."""
    enq = classificar_enquadramento_pc(data_pc, por_ciclo)
    ciclo = enq.ciclo
    registro = por_ciclo.get(ciclo or "") or {}
    inicio = registro.get("inicio_efeito_financeiro")
    efeito = "Sim" if (ciclo and inicio and data_pc >= inicio) else "Nao"
    # VALOR_ATUALIZADO (26C): valor-base x fator historico integral do ciclo,
    # desacoplado do efeito. Quem governa o resultado e o valor considerado.
    atualizado = round(valor * _fator(ciclo, por_ciclo), 2) if ciclo else valor
    # Arredondamento por PC: o retroativo e a diferenca do valor ja arredondado
    # a centavos, nunca o produto bruto.
    retroativo = round(atualizado - valor, 2) if efeito == "Sim" else 0.0
    return {
        "numero_pc": numero,
        "data_pc": data_pc,
        "ciclo": ciclo,
        "enquadramento": enq.tipo,
        "enquadramento_rotulo": enq.rotulo,
        "valor_pc": valor,
        "valor_atualizado": atualizado,
        "valor_historico_considerado": round(valor + retroativo, 2),
        "fator_efetivo_considerado": (
            _fator(ciclo, por_ciclo) if efeito == "Sim" else 1.0
        ),
        "retroativo_reconhecido_a_pagar": retroativo,
        "efeito_financeiro_pc": efeito,
        "entra_no_calculo": "Sim",
        "dentro_do_corte": _pc_dentro_do_corte(data_pc, data_corte),
        "campos_vta": {"status_consolidacao": "COMPUTADO"},
    }


def _vinte_pcs(por_ciclo=POR_CICLO_CANONICO, data_corte: date | None = None):
    """20 PCs mensais: 12 em C1 (jan-dez/2025) e 8 em C2 (jan-ago/2026)."""
    datas = [date(2025, m, 15) for m in range(1, 13)]
    datas += [date(2026, 1, 15), date(2026, 2, 15)]
    datas += [date(2026, m, 15) for m in range(3, 9)]
    return [
        _pc(f"PC-{i:03d}", d, por_ciclo, data_corte=data_corte)
        for i, d in enumerate(datas, 1)
    ]


# --------------------------------------------------------------------------- #
# 1-3, 6 — totais canonicos separados por significado
# --------------------------------------------------------------------------- #
def test_total_informado_soma_todos_os_pcs():
    """Item 1: o total informado e o inventario integral (20 PCs)."""
    tc = _totais_canonicos_pc(_vinte_pcs())
    assert tc["informado"]["quantidade"] == 20
    assert tc["informado"]["valor_pc"] == pytest.approx(20 * VALOR_PC, abs=0.01)


def test_total_enquadrado_nos_ciclos_abrange_todos_os_pcs():
    """Item 2: com ciclos contiguos, os 20 PCs estao enquadrados em C1 ou C2.

    Nenhum PC fica "Fora dos ciclos" por consequencia do inicio do efeito
    financeiro: jan-fev/2026 pertencem a C2 como qualquer outra competencia.
    """
    tc = _totais_canonicos_pc(_vinte_pcs())
    assert tc["enquadrado_ciclos"]["quantidade"] == 20
    assert tc["enquadrado_ciclos"]["valor_pc"] == pytest.approx(20 * VALOR_PC, abs=0.01)
    assert tc["enquadrado_ciclos"]["valor_pc"] == tc["informado"]["valor_pc"]
    assert tc["intervalo_precluso"]["quantidade"] == 0
    assert tc["indeterminado"]["quantidade"] == 0


def test_total_com_efeito_financeiro():
    """Item 3: 16 PCs alcancam o inicio do efeito (10 em C1 + 6 em C2)."""
    tc = _totais_canonicos_pc(_vinte_pcs())
    assert tc["com_efeito"]["quantidade"] == 16
    assert tc["com_efeito"]["valor_pc"] == pytest.approx(16 * VALOR_PC, abs=0.01)


def test_totais_financeiros_finais_fecham():
    """Item 6: considerado = pago + retroativo, e o retroativo so vem do efeito."""
    tc = _totais_canonicos_pc(_vinte_pcs())
    informado = tc["informado"]
    esperado_retro = round(
        10 * (round(VALOR_PC * _fator("C1"), 2) - VALOR_PC)
        + 6 * (round(VALOR_PC * _fator("C2"), 2) - VALOR_PC), 2
    )
    assert informado["retroativo"] == pytest.approx(esperado_retro, abs=0.01)
    assert informado["valor_considerado"] == pytest.approx(
        informado["valor_pc"] + esperado_retro, abs=0.01
    )
    # Os PCs sem efeito nao contribuem com retroativo algum.
    assert tc["sem_efeito_ciclo"]["retroativo"] == 0.0
    assert tc["intervalo_precluso"]["retroativo"] == 0.0


# --------------------------------------------------------------------------- #
# 4 — dentro do ciclo, antes do inicio do efeito financeiro
# --------------------------------------------------------------------------- #
@pytest.mark.parametrize("data_pc", [date(2025, 1, 15), date(2025, 2, 15)])
def test_pc_dentro_de_c1_antes_do_efeito(data_pc):
    """Item 4: pertence a C1, efeito Nao, retroativo zero, valor sem reajuste."""
    pc = _pc("PC-X", data_pc)
    assert pc["ciclo"] == "C1"
    assert pc["enquadramento"] == ENQ_CICLO
    assert pc["efeito_financeiro_pc"] == "Nao"
    assert pc["retroativo_reconhecido_a_pagar"] == 0.0
    assert pc["fator_efetivo_considerado"] == 1.0
    assert pc["valor_historico_considerado"] == pytest.approx(VALOR_PC, abs=0.01)


def test_pcs_sem_efeito_nao_recebem_delta():
    """Item 4/6: o bloco sem efeito nao gera delta algum.

    Sao 4 PCs: jan-fev/2025 (antes do efeito de C1) e jan-fev/2026 (antes do
    efeito de C2) — todos DENTRO dos respectivos ciclos.
    """
    tc = _totais_canonicos_pc(_vinte_pcs())
    bloco = tc["sem_efeito_ciclo"]
    assert bloco["quantidade"] == 4
    assert bloco["valor_pc"] == pytest.approx(4 * VALOR_PC, abs=0.01)
    assert bloco["valor_considerado"] == pytest.approx(bloco["valor_pc"], abs=0.01)
    assert bloco["retroativo"] == 0.0


def test_valor_considerado_de_pc_sem_efeito_e_o_valor_original():
    """Item 6: sem efeito financeiro, o fator efetivo e 1 e o valor nao muda.

    VALOR_ATUALIZADO (26C) segue sendo o valor-base x fator historico integral
    e apenas nao governa mais o resultado.
    """
    pc = _pc("PC-X", date(2025, 1, 15))
    assert pc["fator_efetivo_considerado"] == 1.0
    assert pc["valor_historico_considerado"] == pytest.approx(VALOR_PC, abs=0.01)
    assert pc["valor_atualizado"] == pytest.approx(
        round(VALOR_PC * _fator("C1"), 2), abs=0.01
    )
    assert pc["valor_historico_considerado"] < pc["valor_atualizado"]


# --------------------------------------------------------------------------- #
# 5 — jan-fev do ciclo pertencem ao ciclo; nao existe intervalo derivado
# --------------------------------------------------------------------------- #
@pytest.mark.parametrize("data_pc", [date(2026, 1, 15), date(2026, 2, 15)])
def test_pc_anterior_ao_efeito_de_c2_pertence_a_c2(data_pc):
    """Item 5: PC-013 e PC-014 sao C2, efeito Nao, fator 1, retroativo 0."""
    pc = _pc("PC-X", data_pc)
    assert pc["ciclo"] == "C2"
    assert pc["enquadramento"] == ENQ_CICLO
    assert pc["efeito_financeiro_pc"] == "Nao"
    assert pc["fator_efetivo_considerado"] == 1.0
    assert pc["valor_historico_considerado"] == pytest.approx(VALOR_PC, abs=0.01)
    assert pc["retroativo_reconhecido_a_pagar"] == 0.0
    assert pc["enquadramento_rotulo"] == ""


def test_nenhum_pc_fica_fora_dos_ciclos_entre_ciclos_consecutivos():
    """Item 5: o total enquadrado nos ciclos e o proprio total informado."""
    tc = _totais_canonicos_pc(_vinte_pcs())
    assert tc["intervalo_precluso"]["quantidade"] == 0
    assert tc["indeterminado"]["quantidade"] == 0
    assert tc["enquadrado_ciclos"]["valor_pc"] == pytest.approx(
        tc["informado"]["valor_pc"], abs=0.01
    )


def test_fronteiras_dos_ciclos_pertencem_aos_ciclos():
    """Item 5/17: 31/12 fecha o ciclo e 01/01 abre o seguinte."""
    assert classificar_enquadramento_pc(
        date(2025, 12, 31), POR_CICLO_CANONICO).ciclo == "C1"
    assert classificar_enquadramento_pc(
        date(2026, 1, 1), POR_CICLO_CANONICO).ciclo == "C2"


def test_lacuna_no_calendario_nunca_vira_categoria_de_intervalo():
    """ETAPA 31: gap entre janelas de reajuste nao reclassifica o PC como
    intervalo/precluso — a data enquadra pela CRONOLOGIA fixa da execucao."""
    enq = classificar_enquadramento_pc(date(2026, 1, 15), POR_CICLO_LACUNA)
    assert enq.tipo == ENQ_CICLO
    assert enq.ciclo == "C2"
    assert enq.tipo != ENQ_INTERVALO_PRECLUSO
    assert not enq.e_precluso


def test_data_sem_faixa_justificavel_continua_indeterminada():
    """Item 5: data fora de qualquer ciclo segue indeterminada."""
    enq = classificar_enquadramento_pc(date(2035, 1, 1), POR_CICLO_CANONICO)
    assert enq.tipo == ENQ_INDETERMINADO
    assert not enq.e_precluso


# --------------------------------------------------------------------------- #
# 7 — data de corte
# --------------------------------------------------------------------------- #
def test_pc_posterior_ao_corte_fica_no_inventario_e_fora_dos_resultados():
    """Item 7: permanece no total informado, sai de tudo que e 'ate o corte'."""
    corte = date(2026, 8, 5)
    tc = _totais_canonicos_pc(_vinte_pcs(data_corte=corte), data_corte=corte)
    assert tc["informado"]["quantidade"] == 20
    assert tc["ate_o_corte"]["quantidade"] == 19
    assert tc["posterior_ao_corte"]["quantidade"] == 1
    assert tc["enquadrado_ciclos"]["quantidade"] == 19
    assert tc["com_efeito"]["quantidade"] == 15
    assert tc["sem_efeito_ciclo"]["quantidade"] == 4
    assert tc["informado"]["valor_pc"] == pytest.approx(20 * VALOR_PC, abs=0.01)
    assert tc["ate_o_corte"]["valor_pc"] == pytest.approx(19 * VALOR_PC, abs=0.01)


def test_corte_ausente_nao_exclui_nenhum_pc():
    """Item 7/16: sem data de corte, o comportamento anterior e preservado."""
    tc = _totais_canonicos_pc(_vinte_pcs(), data_corte=None)
    assert tc["corte_aplicado"] is False
    assert tc["ate_o_corte"]["quantidade"] == 20
    assert tc["posterior_ao_corte"]["quantidade"] == 0


def test_corte_nao_exclui_por_falta_de_informacao():
    """Exclusao exige prova de posterioridade, nunca ausencia de dado."""
    assert _pc_dentro_do_corte(None, date(2026, 8, 5)) is True
    assert _pc_dentro_do_corte("data ilegivel", date(2026, 8, 5)) is True
    assert _pc_dentro_do_corte(date(2026, 8, 5), date(2026, 8, 5)) is True
    assert _pc_dentro_do_corte(date(2026, 8, 6), date(2026, 8, 5)) is False


# --------------------------------------------------------------------------- #
# 9-10 — temporalidade dos aditivos nas aberturas
# --------------------------------------------------------------------------- #
def test_aditivo_do_meio_do_ciclo_nao_altera_a_abertura_de_c1():
    """Item 9: acrescimo de 20 com efeito em 05/05/2025 nao vira abertura 70."""
    registro = {
        "QTD_REM_AJUSTADA_C0": 60, "DELTA_POSTERIOR_ABERTURA_C0": 0,
        "QTD_REM_AJUSTADA_C1": 70, "DELTA_POSTERIOR_ABERTURA_C1": 20,
        "QTD_REM_AJUSTADA_C2": 52, "DELTA_POSTERIOR_ABERTURA_C2": 0,
        "QTD_REM_AJUSTADA_C3": None, "DELTA_POSTERIOR_ABERTURA_C3": None,
        "QTD_REM_AJUSTADA_C4": None, "DELTA_POSTERIOR_ABERTURA_C4": None,
    }
    _materializar_aberturas_temporais(registro)
    assert registro["QTD_REM_ABERTURA_C1"] == 50
    assert registro["QTD_REM_ABERTURA_C0"] == 60
    # O movimento existe, mas como alteracao DENTRO do periodo.
    assert registro["ALTERACAO_POSTERIOR_ABERTURA_C1"] == 20
    # A abertura seguinte pode refletir o evento, sem retroacao.
    assert registro["QTD_REM_ABERTURA_C2"] == 52


def test_supressao_do_meio_do_ciclo_nao_altera_a_abertura_de_c2():
    """Item 10: supressao de 15 com efeito em 05/05/2026 nao vira abertura 95."""
    registro = {
        "QTD_REM_AJUSTADA_C0": 150, "DELTA_POSTERIOR_ABERTURA_C0": 0,
        "QTD_REM_AJUSTADA_C1": 130, "DELTA_POSTERIOR_ABERTURA_C1": 0,
        "QTD_REM_AJUSTADA_C2": 95, "DELTA_POSTERIOR_ABERTURA_C2": -15,
        "QTD_REM_AJUSTADA_C3": None, "DELTA_POSTERIOR_ABERTURA_C3": None,
        "QTD_REM_AJUSTADA_C4": None, "DELTA_POSTERIOR_ABERTURA_C4": None,
    }
    _materializar_aberturas_temporais(registro)
    assert registro["QTD_REM_ABERTURA_C2"] == 110
    assert registro["ALTERACAO_POSTERIOR_ABERTURA_C2"] == -15
    assert registro["QTD_REM_ABERTURA_C1"] == 130


def test_abertura_sem_camada_temporal_preserva_comportamento_anterior():
    """Item 16: arquivo antigo, sem as colunas auxiliares, nao regride."""
    registro = {f"QTD_REM_AJUSTADA_C{n}": 40 + n for n in range(5)}
    registro.update({f"DELTA_POSTERIOR_ABERTURA_C{n}": None for n in range(5)})
    _materializar_aberturas_temporais(registro)
    for n in range(5):
        assert registro[f"QTD_REM_ABERTURA_C{n}"] == 40 + n
        assert registro[f"ALTERACAO_POSTERIOR_ABERTURA_C{n}"] == 0.0


def test_aditivo_na_propria_data_de_abertura_conta_na_abertura():
    """Item 9: evento com DATA_EFEITO == abertura pertence aquela abertura.

    A coluna auxiliar so acumula deltas POSTERIORES a abertura; um evento na
    data exata da abertura nao aparece nela e por isso permanece na fotografia.
    """
    registro = {"QTD_REM_AJUSTADA_C1": 70, "DELTA_POSTERIOR_ABERTURA_C1": 0}
    registro.update({f"QTD_REM_AJUSTADA_C{n}": None for n in (0, 2, 3, 4)})
    registro.update({f"DELTA_POSTERIOR_ABERTURA_C{n}": None for n in (0, 2, 3, 4)})
    _materializar_aberturas_temporais(registro)
    assert registro["QTD_REM_ABERTURA_C1"] == 70


# --------------------------------------------------------------------------- #
# 8, 11, 12, 15 — composicao do VTA
# --------------------------------------------------------------------------- #
def _leitura_vta(*, fisico: bool, data_corte: date | None = date(2026, 8, 5)):
    """Leitura minima para o motor: 1 item, C0 executa 10 unidades a VU 200."""
    posicao = {
        "ITEM": 3, "VU_ORIGINAL": 200,
        "QTD_REM_AJUSTADA_C0": 60, "DELTA_POSTERIOR_ABERTURA_C0": 0,
        "QTD_REM_AJUSTADA_C1": 70, "DELTA_POSTERIOR_ABERTURA_C1": 20,
        "QTD_REM_AJUSTADA_C2": 52, "DELTA_POSTERIOR_ABERTURA_C2": 0,
        "QTD_REM_AJUSTADA_C3": None, "DELTA_POSTERIOR_ABERTURA_C3": None,
        "QTD_REM_AJUSTADA_C4": None, "DELTA_POSTERIOR_ABERTURA_C4": None,
    }
    _materializar_aberturas_temporais(posicao)
    leitura = {
        "controle": {"modo": "PC", "ciclo_vigente": "C2"},
        "parametros_v10": {"por_ciclo": POR_CICLO_CANONICO},
        "posicao_contratual": {"ok": True, "itens": [posicao]},
        "historico_vu": {"itens": [{"item": 3, "vu_ciclos": {
            "VU_C0": 200.0, "VU_C1": 208.98, "VU_C2": 215.58}}]},
        "itens_pc_v10": {"itens": _vinte_pcs(data_corte=data_corte)},
    }
    if fisico:
        leitura["ciclo_em_execucao"] = {
            "disponivel": True, "completo": True, "valido": True,
            "data_posicao": date(2026, 8, 5),
            "total_valor_consumido": 3018.12,
            "total_valor_remanescente": 8192.04,
        }
    return leitura


def test_vta_nao_cria_parcela_historica_externa_ao_ciclo_vigente():
    """Item 11: jan-fev/2026 pertencem a C2 e nao viram parcela historica.

    Sem a categoria de intervalo, esses R$ 20.000,00 nao podem ser somados de
    novo ao VTA: ja estao abrangidos pelo componente fisico de C2.
    """
    c = montar_composicao_vta(_leitura_vta(fisico=True))
    assert c["disponivel"] is True
    fontes = [l["fonte"] for l in c["execucao_por_ciclo"]]
    assert "pc_intervalo_precluso" not in fontes
    assert not any(l["ciclo"] == "C2" and l["fonte"].startswith("pc")
                   for l in c["execucao_por_ciclo"])


def test_vta_nao_aplica_fator_efetivo_a_pc_sem_efeito():
    """Item 12: C1 entra pelo considerado, nao pelo fator integral."""
    c = montar_composicao_vta(_leitura_vta(fisico=True))
    c1 = next(l for l in c["execucao_por_ciclo"] if l["ciclo"] == "C1")
    integral = round(12 * VALOR_PC * _fator("C1"), 2)
    considerado = round(2 * VALOR_PC + 10 * round(VALOR_PC * _fator("C1"), 2), 2)
    assert c1["valor_atualizado"] == pytest.approx(considerado, abs=0.02)
    assert c1["valor_atualizado"] < integral


def test_posicao_fisica_completa_prevalece_sobre_estimativa_por_pcs():
    """Item 8: presente e futuro vem do fiscal, nao dos PCs do ciclo vigente."""
    c = montar_composicao_vta(_leitura_vta(fisico=True))
    fontes = [l["fonte"] for l in c["execucao_por_ciclo"]]
    assert "posicao_fisica" in fontes
    assert c["saldo_remanescente"]["fonte"] == "posicao_fisica"
    assert c["saldo_remanescente"]["valor_atualizado"] == pytest.approx(
        8192.04, abs=0.01
    )
    presente = next(
        l for l in c["execucao_por_ciclo"] if l["fonte"] == "posicao_fisica"
    )
    assert presente["valor_atualizado"] == pytest.approx(3018.12, abs=0.01)
    assert any("PREVALECE" in a for a in c["alertas"])
    # Nenhum PC do ciclo vigente foi somado a execucao.
    assert not any(
        l["ciclo"] == "C2" and l["fonte"] == "pc" for l in c["execucao_por_ciclo"]
    )


def test_sem_posicao_fisica_mantem_a_estimativa_anterior():
    """Item 8/16: sem CICLO_EM_EXECUCAO valido, o comportamento anterior segue."""
    c = montar_composicao_vta(_leitura_vta(fisico=False))
    assert c["saldo_remanescente"]["fonte"] == "remanescente"
    assert not any(l["fonte"] == "posicao_fisica" for l in c["execucao_por_ciclo"])


def test_c0_executado_usa_a_abertura_temporalmente_correta():
    """Item 9: o aditivo do meio de C1 nao pode encolher a execucao de C0."""
    c = montar_composicao_vta(_leitura_vta(fisico=True))
    c0 = next(l for l in c["execucao_por_ciclo"] if l["ciclo"] == "C0")
    # abertura C0 (60) - abertura C1 (50) = 10 unidades a VU_C0 200.
    assert c0["valor_atualizado"] == pytest.approx(2000.0, abs=0.01)


def test_vta_exclui_pc_posterior_ao_corte():
    """Item 7: o PC de 15/08/2026 nao compoe o VTA com corte em 05/08/2026."""
    c = montar_composicao_vta(_leitura_vta(fisico=True, data_corte=date(2026, 8, 5)))
    assert any("posterior a data de corte" in a for a in c["alertas"])


def test_reconciliacao_das_parcelas_fecha_com_o_total():
    """Item 15: a soma das parcelas e exatamente o VTA composto."""
    c = montar_composicao_vta(_leitura_vta(fisico=True))
    soma = sum(l["valor_atualizado"] for l in c["execucao_por_ciclo"])
    soma += c["saldo_remanescente"]["valor_atualizado"]
    assert round(soma, 2) == pytest.approx(c["vta_composicao"], abs=0.01)


# --------------------------------------------------------------------------- #
# 16-18 — nao regressao
# --------------------------------------------------------------------------- #
def test_ciclos_contiguos_nao_produzem_intervalo():
    """Item 17: sem lacuna entre C1 e C2, todo PC continua enquadrado."""
    tc = _totais_canonicos_pc(_vinte_pcs(POR_CICLO_CONTIGUO))
    assert tc["intervalo_precluso"]["quantidade"] == 0
    assert tc["enquadrado_ciclos"]["quantidade"] == 20
    assert tc["enquadrado_ciclos"]["valor_pc"] == pytest.approx(
        tc["informado"]["valor_pc"], abs=0.01
    )


def test_modelo_sem_medida_canonica_cai_no_comportamento_anterior():
    """Item 16: registro antigo (sem valor_historico_considerado) nao quebra."""
    antigo = {
        "numero_pc": "PC-LEGADO", "data_pc": date(2025, 6, 15), "ciclo": "C1",
        "valor_pc": VALOR_PC, "valor_atualizado": 10449.14,
        "retroativo_reconhecido_a_pagar": 449.14,
        "efeito_financeiro_pc": "Sim", "entra_no_calculo": "Sim",
    }
    tc = _totais_canonicos_pc([antigo])
    assert tc["informado"]["quantidade"] == 1
    assert tc["enquadrado_ciclos"]["valor_considerado"] == pytest.approx(
        VALOR_PC, abs=0.01
    )


def test_metodos_financeiro_e_itens_nao_passam_pela_composicao_pc():
    """Item 18: so o modo PC entra no ramo corrigido."""
    leitura = _leitura_vta(fisico=True)
    for modo in ("financeiro", "itens"):
        leitura_modo = {**leitura, "controle": {**leitura["controle"], "modo": modo}}
        c = montar_composicao_vta(leitura_modo)
        assert c.get("metodo") != "pc"
        assert not any(
            l.get("fonte") in ("pc_intervalo_precluso", "posicao_fisica")
            for l in c.get("execucao_por_ciclo", [])
        )


def test_pc_descartado_por_duplicidade_nao_soma():
    """Duplicidade ja resolvida continua fora de todos os totais canonicos."""
    pcs = _vinte_pcs()
    pcs.append({
        **pcs[0], "numero_pc": "PC-001-DUP",
        "campos_vta": {"status_consolidacao": "DESCARTADO_DUPLICIDADE"},
    })
    tc = _totais_canonicos_pc(pcs)
    assert tc["informado"]["quantidade"] == 20
