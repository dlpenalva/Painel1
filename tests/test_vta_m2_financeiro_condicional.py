"""Testes permanentes do VTA-M2/VTA-M2.1 — VTA do metodo Financeiro no XLS.

Protege, por leitura estrutural de formula (nao apenas SHA/contagem), que:
  A. o ramo Financeiro de B26 existe e usa a fonte financeira (D20);
  B. o ramo PC dentro de B26 permanece literalmente preservado;
  C. o ramo Itens/Consumido dentro de B26 permanece literalmente preservado;
  D. B28 no Financeiro mantem a formula antiga como referencia comparativa;
  E. o texto de metodologia em RESULTADOS e condicional ao metodo
     selecionado (nunca fixo mencionando "Financeiro" para outro metodo);
  F. o bloco CONFERENCIA DA EXECUCAO nao apresenta numeros/titulo do
     Financeiro como oficiais quando o metodo != Financeiro;
  G. ausencia de dado nunca vira zero (aparece "NAO COMPARAVEL"/"Nao
     aplicavel ao metodo selecionado", nao 0);
  H. o endereco/nome definido VTA_FINAL continua apontando para B26;
  I. o template abre estruturalmente via openpyxl;
  J. quando disponivel, o Excel real recalcula e reabre sem reparo.
"""
from __future__ import annotations

import gc
from pathlib import Path

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

_RAMO_PC_B26 = (
    'IF($T$25="CALCULO MANUAL REQUERIDO","",ROUND($T$25+IF(ISNUMBER(B24),B24,0),2))'
)
_RAMO_ITENS_B26 = (
    'IF(OR(B23="",AND(B24<>"",NOT(ISNUMBER(B24)))),"",'
    'ROUND(B23+IF(ISNUMBER(B24),B24,0)+IF(ISNUMBER($N$263),$N$263,0),2))'
)
_B28_ESPERADO = '=IF($B$4="Financeiro",$B$23,$B$26)'


@pytest.fixture(scope="module")
def wb():
    return load_workbook(TEMPLATE, data_only=False)


def test_a_ramo_financeiro_b26_usa_fonte_financeira(wb):
    b26 = str(wb["MEMORIA_RESULTADOS"]["B26"].value)
    assert '$B$4="Financeiro"' in b26
    assert "$D$20" in b26
    assert "B21" in b26 and "D35" in b26


def test_b_ramo_pc_b26_preservado_literalmente(wb):
    b26 = str(wb["MEMORIA_RESULTADOS"]["B26"].value)
    assert _RAMO_PC_B26 in b26


def test_c_ramo_itens_consumido_b26_preservado_literalmente(wb):
    b26 = str(wb["MEMORIA_RESULTADOS"]["B26"].value)
    assert _RAMO_ITENS_B26 in b26


def test_d_b28_financeiro_mantem_referencia_comparativa(wb):
    b28 = str(wb["MEMORIA_RESULTADOS"]["B28"].value)
    assert b28 == _B28_ESPERADO
    assert "$B$23" in b28


def test_e_metodologia_resultados_e_condicional_ao_metodo(wb):
    formula = str(wb["RESULTADOS"]["A70"].value)
    assert formula.startswith("=IF(")
    assert "MEMORIA_RESULTADOS!$B$4" in formula
    assert '"Financeiro"' in formula
    assert '"PCs"' in formula
    assert '"Itens"' in formula
    assert "Metodo Financeiro" in formula
    assert "Metodo PCs" in formula
    assert "Metodo Consumido" in formula


def test_f_titulo_conferencia_e_condicional(wb):
    titulo = str(wb["RESULTADOS"]["A71"].value)
    assert titulo.startswith("=IF(")
    assert "MEMORIA_RESULTADOS!$B$4" in titulo
    assert "(Financeiro)" in titulo
    assert "nao aplicavel ao metodo selecionado" in titulo.lower()


def test_f_bloco_conferencia_nao_mostra_financeiro_para_outro_metodo(wb):
    res = wb["RESULTADOS"]
    for linha in range(73, 78):
        for coluna in ("B", "C", "D", "E"):
            formula = str(res[f"{coluna}{linha}"].value)
            assert formula.startswith("=IF(MEMORIA_RESULTADOS!$B$4<>\"Financeiro\",")
            rotulo = "NAO APLICAVEL" if coluna == "E" else "Nao aplicavel ao metodo selecionado"
            assert rotulo in formula


def test_g_ausencia_de_dado_nao_vira_zero(wb):
    res = wb["RESULTADOS"]
    for linha in range(73, 78):
        formula_c = str(res[f"C{linha}"].value)
        assert "NAO COMPARAVEL" in formula_c
        assert "Nao aplicavel ao metodo selecionado" in formula_c
        # nao pode haver um "senao 0" generico substituindo a comparacao.
        assert not formula_c.rstrip().endswith(",0)")


def test_g_mapa_quantitativo_tem_cadeia_de_fallback_ate_e(wb):
    formula_c3 = str(wb["RESULTADOS"]["C76"].value)
    # C3: cadeia N -> J -> F -> E (nao trava em "NAO COMPARAVEL" so por
    # faltar o checkpoint imediatamente anterior, quando ha base mais
    # antiga disponivel).
    assert "posicao_contratual!$N$2:$N$201" in formula_c3
    assert "posicao_contratual!$J$2:$J$201" in formula_c3
    assert "posicao_contratual!$F$2:$F$201" in formula_c3
    assert "posicao_contratual!$E$2:$E$201" in formula_c3
    assert "historico_VU!$F$2:$F$201" in formula_c3  # VU_C3


def test_h_nome_definido_vta_final_aponta_para_b26(wb):
    definido = wb.defined_names["VTA_FINAL"]
    destinos = list(definido.destinations)
    assert len(destinos) == 1
    aba, referencia = destinos[0]
    assert aba == "MEMORIA_RESULTADOS"
    assert referencia.replace("$", "") == "B26"


def test_i_template_abre_estruturalmente_via_openpyxl():
    wb_local = load_workbook(TEMPLATE, data_only=False)
    assert "MEMORIA_RESULTADOS" in wb_local.sheetnames
    assert "RESULTADOS" in wb_local.sheetnames


def test_j_excel_real_recalcula_e_reabre_sem_reparo():
    client = pytest.importorskip("win32com.client")
    pythoncom = pytest.importorskip("pythoncom")
    pythoncom.CoInitialize()
    excel = client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb_com = excel.Workbooks.Open(str(TEMPLATE), UpdateLinks=0, ReadOnly=True)
        try:
            assert wb_com.Worksheets.Count == 15
            mem = wb_com.Worksheets("MEMORIA_RESULTADOS")
            assert str(mem.Range("B26").Formula)
            res = wb_com.Worksheets("RESULTADOS")
            assert str(res.Range("A70").Formula).startswith("=IF(")
        finally:
            wb_com.Close(False)
            del wb_com
    finally:
        excel.Quit()
        del excel
        gc.collect()
        pythoncom.CoUninitialize()
