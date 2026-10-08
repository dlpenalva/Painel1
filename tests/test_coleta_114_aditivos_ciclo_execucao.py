"""Coleta 11.4 — fator dos aditivos, ciclo em execucao e novo item.

Cenario focal (espelho de Coleta_Reajuste_C1_IST_08-10-2026.xlsx): somente C1
apurado (fator 1,0381), data de corte 08/10/2026 (janela de C4), aditivo em
30/09/2025 com N001:N005. Os testes rapidos conferem a estrutura; o bloco
Excel COM (RUN_EXCEL_INTEGRATION=1) recalcula o cenario no Excel real.
"""
from __future__ import annotations

import os
import re
from datetime import date, datetime
from pathlib import Path

import pytest

from _ciclo_em_execucao import (
    _formula_vu,
    _periodo_por_ciclo,
    ciclo_em_execucao_por_data_corte,
)
from _motor_posicao_contratual import normalizar_tipo_movimento
from tests._fabrica_coleta import bytes_coleta_oficial, workbook_coleta_oficial
from tools.aplicar_coleta_114_aditivos_ciclo_execucao import (
    FORMULA_CICLO_VIGENTE,
    FORMULA_FATOR,
    FORMULA_VALOR,
    OPCOES_TIPO,
    validar_formula,
)

RAIZ = Path(__file__).resolve().parents[1]

PAYLOAD_C1 = {
    "origem": "Reajuste Simples",
    "tipo": "Simples",
    "indice": "IST",
    "data_base_original": "01/03/2022",
    "ciclos": [{
        "ciclo": "C1",
        "data_base": "01/03/2022",
        "data_inicio": "01/03/2023",
        "data_pedido": "15/06/2023",
        "financeiro_inicio": "01/06/2023",
        "inicio_efeito_financeiro": "01/06/2023",
        "percentual_aplicado": 0.0381,
        "situacao": "TEMPESTIVO",
        "objeto_analise_atual": True,
    }],
}


@pytest.fixture(scope="module")
def wb_gerado():
    return workbook_coleta_oficial(PAYLOAD_C1)


def _deslocar(formula: str, linha: int) -> str:
    """Formula da linha 2 levada a outra linha (so linhas relativas, fora de aspas)."""
    return re.sub(r'(?<![A-Z$"])(\$?[A-M])2(?!\d)', lambda m: f"{m.group(1)}{linha}", formula)


# --- estrutura do XLS gerado ---------------------------------------------

def test_versoes_no_xls_gerado(wb_gerado):
    assert wb_gerado["CONTROLE"]["B24"].value == "11.4"
    assert wb_gerado["CONTROLE"]["B25"].value == "11.9"


def test_ciclo_em_execucao_e_formula_da_data_de_corte(wb_gerado):
    # O gerador nao sobrescreve B2 com o ciclo analisado.
    assert wb_gerado["CONTROLE"]["B2"].value == FORMULA_CICLO_VIGENTE
    assert wb_gerado["CICLO_EM_EXECUCAO"]["C3"].value == "=CONTROLE!$B$2"
    # O ciclo analisado continua registrado em parametros/B12.
    assert wb_gerado["parametros"]["A3"].value == "Sim"
    assert str(wb_gerado["CONTROLE"]["B12"].value).startswith("=")


def test_formulas_de_aditivos_em_todas_as_linhas(wb_gerado):
    ws = wb_gerado["aditivos"]
    for linha in (2, 57, 200):
        assert ws[f"I{linha}"].value == _deslocar(FORMULA_FATOR, linha)
        assert ws[f"J{linha}"].value == _deslocar(FORMULA_VALOR, linha)
        assert "ALERTA: NOVO ITEM INVALIDO" in ws[f"M{linha}"].value
    # Sem lookup na tabela de ciclos apurados (fator vazio em C2-C4).
    assert "parametros!$B:$F" not in ws["I2"].value


def test_formulas_novas_sao_ascii_e_balanceadas(wb_gerado):
    for formula in (FORMULA_CICLO_VIGENTE, FORMULA_FATOR, FORMULA_VALOR,
                    wb_gerado["aditivos"]["M2"].value, _formula_vu(13, 2)):
        validar_formula(formula)


def test_dropdown_tipo_com_novo_item(wb_gerado):
    listas = {
        str(dv.sqref): dv.formula1
        for dv in wb_gerado["aditivos"].data_validations.dataValidation
    }
    assert listas["D2:D200"] == '"' + ",".join(OPCOES_TIPO) + '"'
    assert OPCOES_TIPO == ("Acrescimo", "Acréscimo - novo item", "Supressao")


def test_vu_do_ciclo_em_execucao_carrega_ultimo_fator():
    formula = _formula_vu(13, 2)
    # Primeiro o historico_VU do ciclo; so na lacuna, o fator carregado a
    # partir do nascimento do item.
    assert formula.startswith('=IF(A13="","",IF(ISNUMBER(IF($C$3="C0",historico_VU!$C2')
    assert "posicao_contratual!$Y2" in formula
    assert "COUNT(parametros!$F$2:$F$6)" in formula


@pytest.mark.parametrize("tipo, esperado", [
    ("Acrescimo", "ACRESCIMO"),
    ("Acréscimo", "ACRESCIMO"),
    ("Acréscimo - novo item", "ACRESCIMO"),
    ("Supressao", "SUPRESSAO"),
    ("Supressão", "SUPRESSAO"),
])
def test_tipos_aceitos_pelo_motor(tipo, esperado):
    assert normalizar_tipo_movimento(tipo) == esperado


# --- ciclo em execucao sem cache (espelho Python da formula de B2) ---------
# Cronologia FIXA da execucao (ancora C0 = 01/03/2022): C1 03/2023-02/2024,
# C2 03/2024-02/2025, C3 03/2025-02/2026, C4 03/2026-02/2027. As janelas de
# reajuste (C2 a partir de 06/2024) NAO enquadram a execucao.

@pytest.mark.parametrize("corte, esperado", [
    (datetime(2026, 10, 8), "C4"),   # cenario focal
    (datetime(2024, 4, 15), "C2"),   # intervalo das janelas de reajuste
    (datetime(2024, 2, 29), "C1"),   # corte padrao = fim do ciclo analisado
    (datetime(2024, 3, 1), "C2"),
    (datetime(2025, 9, 30), "C3"),
    (datetime(2027, 2, 28), "C4"),   # ultimo dia da cronologia suportada
    (datetime(2027, 3, 1), ""),      # depois de C4: nunca C4 indefinidamente
    (datetime(2022, 2, 28), ""),     # antes de C0
    (None, ""),
])
def test_ciclo_em_execucao_por_data_corte(wb_gerado, corte, esperado):
    wb_gerado["CONTROLE"]["B3"].value = corte
    assert ciclo_em_execucao_por_data_corte(wb_gerado) == esperado


def test_inicio_e_fim_do_ciclo_em_execucao_seguem_a_cronologia(wb_gerado):
    ws = wb_gerado["CICLO_EM_EXECUCAO"]
    assert "EDATE(" in ws["F3"].value
    assert "parametros!$B$2:$B$6" not in ws["F3"].value  # nao e a janela C:D
    assert ws["H3"].value == '=IF(ISNUMBER($F$3),EDATE($F$3,12)-1,"")'
    assert _periodo_por_ciclo(wb_gerado, "C2") == (date(2024, 3, 1), date(2025, 2, 28))
    assert _periodo_por_ciclo(wb_gerado, "C4") == (date(2026, 3, 1), date(2027, 2, 28))


def test_janelas_do_cenario_nao_mudam(wb_gerado):
    ws = wb_gerado["parametros"]
    janelas = [(ws[f"C{r}"].value.date(), ws[f"D{r}"].value.date()) for r in range(3, 7)]
    assert janelas == [
        (date(2023, 3, 1), date(2024, 2, 29)),
        (date(2024, 6, 1), date(2025, 5, 31)),
        (date(2025, 6, 1), date(2026, 5, 31)),
        (date(2026, 6, 1), date(2027, 5, 31)),
    ]


# --- ICTI ------------------------------------------------------------------

def test_workflow_icti_instala_dependencias_do_projeto():
    workflow = (RAIZ / ".github/workflows/atualizar-icti-ipeadata.yml").read_text(encoding="utf-8")
    assert "-r requirements.txt pytest==8.4.1" in workflow
    assert "python-dateutil" in (RAIZ / "requirements.txt").read_text(encoding="utf-8")


# --- Excel real ------------------------------------------------------------

@pytest.mark.skipif(
    os.environ.get("RUN_EXCEL_INTEGRATION") != "1",
    reason="defina RUN_EXCEL_INTEGRATION=1 para executar o Excel COM",
)
@pytest.mark.parametrize("corte, ciclo, inicio, fim", [
    (datetime(2026, 10, 8), "C4", date(2026, 3, 1), date(2027, 2, 28)),
    (datetime(2024, 4, 15), "C2", date(2024, 3, 1), date(2025, 2, 28)),
])
def test_cenario_focal_no_excel(tmp_path, corte, ciclo, inicio, fim):
    import pythoncom
    import win32com.client as win32

    caminho = tmp_path / "Coleta_Reajuste_C1_IST_08-10-2026.xlsx"
    caminho.write_bytes(bytes_coleta_oficial(PAYLOAD_C1))
    itens = [("1.1", 1000, 10.0), ("N001", None, 5.5), ("N002", None, 6.0),
             ("N003", None, 7.25), ("N004", None, 8.0), ("N005", None, 9.0)]
    aditivos = [("1.1", "Acrescimo", 100)] + [
        (item, "Acréscimo - novo item", qtd)
        for item, qtd in (("N001", 10), ("N002", 20), ("N003", 1851), ("N004", 40), ("N005", 50))
    ]
    pythoncom.CoInitialize()
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(caminho))
        wb.Worksheets("CONTROLE").Range("B3").Value = corte
        rem = wb.Worksheets("itens_Remanesc")
        for linha, (item, base, vu) in enumerate(itens, start=2):
            rem.Range(f"A{linha}").Value = item
            if base is not None:
                rem.Range(f"B{linha}").Value = base
            rem.Range(f"C{linha}").Value = vu
            rem.Range(f"K{linha}").Value = 1851 if item == "N003" else 10
        adi = wb.Worksheets("aditivos")
        for linha, (item, tipo, qtd) in enumerate(aditivos, start=2):
            adi.Range(f"A{linha}").Value = item
            adi.Range(f"B{linha}").Value = datetime(2025, 9, 30)
            adi.Range(f"D{linha}").Value = tipo
            adi.Range(f"E{linha}").Value = qtd
            adi.Range(f"H{linha}").Value = "Sim"
        cee = wb.Worksheets("CICLO_EM_EXECUCAO")
        cee.Range("D5").Value = corte
        xl.CalculateFull()

        assert wb.Worksheets("CONTROLE").Range("B2").Value == ciclo
        assert wb.Worksheets("CONTROLE").Range("B12").Value == "C1"
        assert cee.Range("C3").Value == ciclo
        assert cee.Range("F3").Value.date() == inicio
        assert cee.Range("H3").Value.date() == fim
        assert cee.Range("D5").Validation.Value  # D5 aceita a data de corte
        assert cee.Range("K13").Value != "ERRO: DATA FORA DO CICLO"
        # Janelas de reajuste intactas (C2 segue iniciando em 06/2024).
        assert wb.Worksheets("parametros").Range("C4").Value.date() == date(2024, 6, 1)
        assert wb.Worksheets("parametros").Range("D4").Value.date() == date(2025, 5, 31)

        assert adi.Range("I2").Value == pytest.approx(1.0381)
        assert adi.Range("J2").Value == pytest.approx(1038.00)
        for linha in range(3, 8):
            assert adi.Range(f"I{linha}").Value == pytest.approx(1.0)
            assert adi.Range(f"M{linha}").Value == "OK"
        assert adi.Range("J5").Value == pytest.approx(round(1851 * 7.25, 2))

        if ciclo != "C4":
            wb.Close(False)
            return
        pos = wb.Worksheets("posicao_contratual")
        assert pos.Range("C5").Value == 0  # QTD_BASE_ORIGINAL de N003
        assert pos.Range("U5").Value == 1851  # nunca 3.702

        listados = [cee.Range(f"A{r}").Value for r in range(13, 19)]
        assert listados[1:] == ["N001", "N002", "N003", "N004", "N005"]
        assert cee.Range("B16").Value == 1851  # N003 na abertura de C4
        assert cee.Range("E13").Value == pytest.approx(10.38)
        assert cee.Range("E16").Value == pytest.approx(7.25)

        adi.Range("D2").Value = "Acréscimo - novo item"  # item ja existente
        xl.CalculateFull()
        assert str(adi.Range("M2").Value).startswith("ALERTA: NOVO ITEM INVALIDO")
        wb.Close(False)
    finally:
        xl.Quit()
        pythoncom.CoUninitialize()
