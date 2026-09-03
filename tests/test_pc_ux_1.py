"""PC-UX-1: contrato de apresentação sem alteração da metodologia PC."""
from __future__ import annotations

from copy import deepcopy
from datetime import date
from io import BytesIO
from pathlib import Path

import pytest
from docx import Document
from openpyxl import load_workbook

from _leitor_masterfile_v10 import _totais_canonicos_pc
from _templates_documentos import (
    _situacao_retroativos_pc,
    gerar_despacho_saneador,
    gerar_termo_apostila,
)
from test_sumario_executivo import (
    leitura_multiciclo_pc,
    leitura_simples_financeiro,
)

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"
CORTE = date(2026, 8, 31)


def _pc(
    numero: str,
    ciclo: str,
    valor: float,
    *,
    pago: bool,
    efeito: bool,
    retro: float = 0.0,
    analise: float = 0.0,
    potencial: float = 0.0,
    dentro: bool = True,
) -> dict:
    return {
        "numero_pc": numero,
        "ciclo": ciclo,
        "valor_pc": valor,
        "valor_historico_considerado": round(valor + retro, 2),
        "retroativo_reconhecido_a_pagar": retro,
        "valor_atualizado_em_analise": analise,
        "delta_potencial": potencial,
        "pc_pago_a_contratada": "Sim" if pago else "Nao",
        "efeito_financeiro_pc": "Sim" if efeito else "Nao",
        "dentro_do_corte": dentro,
        "entra_no_calculo": "Sim",
        "campos_vta": {"status_consolidacao": "COMPUTADO"},
    }


def _casos_basicos() -> list[dict]:
    return [
        _pc("C0-PAGO", "C0", 100.0, pago=True, efeito=False),
        _pc("C0-ANALISE", "C0", 50.0, pago=False, efeito=False, analise=50.0),
        _pc("C1-PAGO-EFEITO", "C1", 200.0, pago=True, efeito=True, retro=20.0),
        _pc(
            "C1-ANALISE-EFEITO", "C1", 300.0, pago=False, efeito=True,
            analise=330.0, potencial=30.0,
        ),
        _pc("C1-PAGO-SEM-EFEITO", "C1", 80.0, pago=True, efeito=False),
        _pc("C1-FORA", "C1", 70.0, pago=True, efeito=True, retro=7.0, dentro=False),
    ]


def _caso_real_referencia() -> list[dict]:
    # Entradas decomponíveis que reproduzem o benchmark; os totais são sempre
    # produzidos por _totais_canonicos_pc, nunca gravados na implementação.
    return [
        _pc("C0-PAGO", "C0", 13_559_396.46, pago=True, efeito=False),
        _pc(
            "C0-ANALISE", "C0", 242_807.91, pago=False, efeito=False,
            analise=242_807.91,
        ),
        _pc(
            "C1-PAGO-EFEITO", "C1", 1_697_102.65, pago=True, efeito=True,
            retro=57_701.49,
        ),
        _pc("C1-PAGO-SEM-EFEITO", "C1", 177_217.78, pago=True, efeito=False),
        _pc(
            "C1-ANALISE-EFEITO", "C1", 3_529_897.65, pago=False, efeito=True,
            analise=3_649_914.17, potencial=120_016.52,
        ),
        _pc(
            "C1-ANALISE-SEM-EFEITO", "C1", 377_476.28, pago=False,
            efeito=False, analise=377_476.28,
        ),
    ]


def _texto_docx(conteudo: bytes) -> str:
    doc = Document(BytesIO(conteudo))
    partes = [p.text for p in doc.paragraphs]
    for tabela in doc.tables:
        for linha in tabela.rows:
            partes.extend(celula.text for celula in linha.cells)
    return "\n".join(partes)


def _leitura_documental(itens: list[dict]) -> dict:
    leitura = deepcopy(leitura_multiciclo_pc())
    leitura["controle"]["data_corte"] = CORTE
    leitura["itens_pc_v10"]["itens"] = deepcopy(itens)
    leitura["itens_pc_v10"]["totais_canonicos"] = _totais_canonicos_pc(
        itens, CORTE
    )
    return leitura


def test_a_f_regras_pc_permanecem_separadas_por_estado_e_corte():
    totais = _totais_canonicos_pc(_casos_basicos(), CORTE)
    c0, c1 = totais["por_ciclo"]["C0"], totais["por_ciclo"]["C1"]
    assert c0["quantidade_reconhecida"] == 1
    assert c0["valor_original_reconhecido"] == 100.0
    assert c0["valor_atualizado_reconhecido"] == 100.0
    assert c0["retroativo"] == 0.0
    assert c0["valor_atualizado_em_analise"] == 50.0
    assert c0["delta_potencial"] == 0.0
    assert c1["quantidade_reconhecida"] == 2
    assert c1["valor_original_reconhecido"] == 280.0
    assert c1["valor_atualizado_reconhecido"] == 300.0
    assert c1["retroativo"] == 20.0
    assert c1["valor_atualizado_em_analise"] == 330.0
    assert c1["delta_potencial"] == 30.0
    assert totais["posterior_ao_corte"]["valor_pc"] == 70.0
    assert totais["posterior_ao_corte_por_ciclo"]["C1"]["valor_pc"] == 70.0


def test_g_ausencia_nao_vira_zero_no_payload_documental():
    ausente = {"controle": {"modo": "PC"}, "totais_canonicos_pc": {
        "ate_o_corte": {"quantidade": 1, "retroativo": None},
    }}
    zero = {"controle": {"modo": "PC"}, "totais_canonicos_pc": {
        "ate_o_corte": {"quantidade": 1, "retroativo": 0.0},
    }}
    assert _situacao_retroativos_pc(ausente)["reconhecido"] is None
    assert _situacao_retroativos_pc(zero)["reconhecido"] == 0.0


def test_h_quadro_itens_pc_le_exclusivamente_colunas_canonicas():
    ws = load_workbook(TEMPLATE, data_only=False)["itens_PC"]
    assert ws["M10"].value == "PCs CONSIDERADOS NA APURAÇÃO ATÉ A DATA DE CORTE"
    assert [ws.cell(11, c).value for c in range(13, 21)] == [
        "Ciclo", "PCs pagos/reconhecidos", "Valor original", "Valor atualizado",
        "Retroativo reconhecido", "Valor em análise (área gest.)",
        "Retroativo potencial", "Fora da data de corte",
    ]
    for linha in range(12, 17):
        assert "$D$2:$D$5001" in ws.cell(linha, 15).value
        assert "$U$2:$U$5001" in ws.cell(linha, 16).value
        assert "$H$2:$H$5001" in ws.cell(linha, 17).value
        assert "$I$2:$I$5001" in ws.cell(linha, 18).value
        assert "$J$2:$J$5001" in ws.cell(linha, 19).value
        assert 'MEMORIA_RESULTADOS!$T$31' in ws.cell(linha, 20).value


def test_i_j_resultados_preserva_formulas_e_vta():
    ws = load_workbook(TEMPLATE, data_only=False)["RESULTADOS"]
    assert ws["A9"].value == "1. COMO O VTA FOI CALCULADO"
    assert ws["A15"].value == "2. EXECUÇÃO E RETROATIVO POR CICLO"
    assert ws["B22"].value == '=IF(COUNT(B16:B20)=0,"",ROUND(SUM(B16:B20),2))'
    assert ws["B38"].value == '=IF(OR(B36="",B37=""),"",ROUND(B37-B36,2))'
    assert ws["B86"].value == '=IF(VTA_FINAL="","",VTA_FINAL)'
    assert ws["A70"].value is None and ws["A78"].value is None


def test_o_formulas_e_referencias_foram_reancoradas_sem_tocar_a_l():
    wb = load_workbook(TEMPLATE, data_only=False)
    ws, mem = wb["itens_PC"], wb["MEMORIA_RESULTADOS"]
    assert [ws.cell(1, c).value for c in range(1, 13)] == [
        "NUMERO_PC", "DATA_PC", "CICLO_PC", "VALOR_PC", "FATOR_ACUMULADO",
        "VALOR_ATUALIZADO", "PC_PAGO_A_CONTRATADA",
        "RETROATIVO_RECONHECIDO_A_PAGAR", "VALOR_ATUALIZADO_EM_ANALISE",
        "DELTA_POTENCIAL", "CHECK_PC_FINANCEIRO", "EFEITO_FINANCEIRO_PC",
    ]
    assert "itens_PC!$O$3" in mem["T21"].value
    assert "ROW(itens_PC!$O$3:$O$7)-3" in mem["T22"].value
    assert "itens_PC!$P$3" in mem["T25"].value
    assert "SUM(itens_PC!$O$3:$O$7)" in mem["T35"].value
    assert "ROW(itens_PC!$O$3:$O$7)-3" in mem["W67"].value
    assert wb.defined_names["VTA_FINAL"].value == "MEMORIA_RESULTADOS!$B$26"
    assert wb.defined_names["EXECUTADO_APURADO"].value == "RESULTADOS!$B$83"


def test_caso_real_de_referencia_fecha_sem_hardcode_na_implementacao():
    totais = _totais_canonicos_pc(_caso_real_referencia(), CORTE)
    c0, c1 = totais["por_ciclo"]["C0"], totais["por_ciclo"]["C1"]
    assert c0["valor_pc"] == 13_802_204.37
    assert c0["valor_atualizado_reconhecido"] == 13_559_396.46
    assert c0["valor_atualizado_em_analise"] == 242_807.91
    assert c0["retroativo"] == 0.0
    assert c1["valor_original_reconhecido"] == 1_874_320.43
    assert c1["valor_atualizado_reconhecido"] == 1_932_021.92
    assert c1["retroativo"] == 57_701.49
    assert c1["valor_atualizado_em_analise"] == 4_027_390.45
    assert c1["delta_potencial"] == 120_016.52
    reconhecido = round(sum(x["valor_atualizado_reconhecido"] for x in (c0, c1)), 2)
    analise = round(sum(x["valor_atualizado_em_analise"] for x in (c0, c1)), 2)
    assert reconhecido == 15_491_418.38
    assert analise == 4_270_198.36
    assert round(reconhecido + analise, 2) == 19_761_616.74


def test_k_l_q_r_apostila_e_saneador_usam_os_mesmos_valores_e_abrem():
    leitura = _leitura_documental(_caso_real_referencia())
    termo = gerar_termo_apostila(leitura)
    saneador = gerar_despacho_saneador(leitura)
    for conteudo in (termo, saneador):
        Document(BytesIO(conteudo))
        texto = _texto_docx(conteudo)
        assert "R$ 57.701,49" in texto
        assert "R$ 120.016,52" in texto
        assert "valor em análise pela área gestora" in texto.lower()
        assert "31/08/2026" in texto
    texto_termo = _texto_docx(termo)
    assert texto_termo.count("SITUAÇÃO DOS VALORES RETROATIVOS") == 1
    assert "Quadro 2 — Execução considerada até a data de corte" in texto_termo
    assert "Valor original" in texto_termo
    assert "Valor atualizado" in texto_termo
    assert "Retroativo reconhecido" in texto_termo
    texto_saneador = _texto_docx(saneador)
    assert "PENDÊNCIA TÉCNICA" in texto_saneador
    assert "PROVIDÊNCIA DA ÁREA GESTORA" in texto_saneador


@pytest.mark.parametrize("metodo", ["financeiro", "consumido"])
def test_m_n_outros_metodos_nao_recebem_texto_pc(metodo):
    leitura = leitura_simples_financeiro()
    if metodo == "consumido":
        leitura["controle"]["modo"] = "D"
    for gerador in (gerar_termo_apostila, gerar_despacho_saneador):
        texto = _texto_docx(gerador(leitura))
        assert "Metodologia utilizada: Pedidos de Compra (PC)" not in texto
        assert "PROVIDÊNCIA DA ÁREA GESTORA" not in texto
