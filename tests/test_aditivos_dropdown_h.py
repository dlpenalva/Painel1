"""Regressão cirúrgica das validações e fórmulas de aditivos."""

from io import BytesIO

import pytest
from openpyxl import load_workbook

from _coleta_oficial import (
    TEMPLATE_COLETA_OFICIAL,
    _garantir_apresentacao_retroativos_e_aditivos,
    obter_coleta_oficial_bytes,
)


@pytest.fixture(scope="module")
def workbooks():
    base = load_workbook(TEMPLATE_COLETA_OFICIAL, data_only=False)
    runtime = load_workbook(BytesIO(obter_coleta_oficial_bytes()), data_only=False)
    yield base, runtime
    base.close()
    runtime.close()


def _validacoes_por_faixa(wb):
    validacoes = {}
    for ws in wb.worksheets:
        for dv in ws.data_validations.dataValidation:
            assinatura = (
                dv.type,
                dv.formula1,
                dv.formula2,
                dv.allowBlank,
                dv.showDropDown,
                dv.errorStyle,
                dv.error,
                dv.errorTitle,
                dv.prompt,
                dv.promptTitle,
            )
            for faixa in dv.sqref.ranges:
                validacoes[(ws.title, str(faixa))] = assinatura
    return validacoes


def _formulas(ws):
    return {
        cell.coordinate: cell.value
        for row in ws.iter_rows()
        for cell in row
        if isinstance(cell.value, str) and cell.value.startswith("=")
    }


def test_a_aditivos_h_tem_lista_sim_nao(workbooks):
    _, runtime = workbooks
    ws = runtime["aditivos"]
    validacoes = [
        dv for dv in ws.data_validations.dataValidation
        if "H2:H200" in str(dv.sqref)
    ]
    assert ws["H1"].value == "Aplicar reajuste? (Sim/Nao)"
    assert len(validacoes) == 1
    assert validacoes[0].type == "list"
    assert validacoes[0].formula1 == '"Sim,Nao"'
    assert str(validacoes[0].sqref) == "H2:H200"


def test_b_aditivos_k_sem_validacao_automatico_e_oculto(workbooks):
    _, runtime = workbooks
    ws = runtime["aditivos"]
    assert ws.column_dimensions["K"].hidden is True
    assert ws["K2"].value == '=IF(A2="","",IF(M2="OK","Nao",""))'
    assert not any("K2:K200" in str(dv.sqref) for dv in ws.data_validations.dataValidation)


def test_c_aditivos_d_preserva_validacao_existente(workbooks):
    base, runtime = workbooks
    assert _validacoes_por_faixa(runtime)[("aditivos", "D2:D200")] == (
        _validacoes_por_faixa(base)[("aditivos", "D2:D200")]
    )


def test_d_todas_as_demais_validacoes_do_template_sao_preservadas(workbooks):
    base, runtime = workbooks
    esperadas = _validacoes_por_faixa(base)
    esperadas.pop(("aditivos", "K2:K200"))
    encontradas = _validacoes_por_faixa(runtime)
    assert not (set(esperadas) - set(encontradas))
    assert {
        chave: encontradas[chave][0]
        for chave in esperadas
    } == {
        chave: assinatura[0]
        for chave, assinatura in esperadas.items()
    }


def test_e_f_j_consumindo_h_preserva_os_dois_ramos(workbooks):
    base, runtime = workbooks
    formula = runtime["aditivos"]["J2"].value
    assert formula == base["aditivos"]["J2"].value
    assert 'UPPER(H2)="SIM"' in formula.upper()
    assert formula.endswith(",F2),2))")


def test_g_h_nao_altera_delta_de_quantidade_l(workbooks):
    base, runtime = workbooks
    formula = runtime["aditivos"]["L2"].value
    assert formula == base["aditivos"]["L2"].value
    assert "H2" not in formula.upper()


def test_h_funcao_nao_altera_formulas_de_resultados_memoria_e_vta():
    wb = load_workbook(TEMPLATE_COLETA_OFICIAL, data_only=False)
    antes = {
        aba: _formulas(wb[aba])
        for aba in ("RESULTADOS", "MEMORIA_RESULTADOS", "comparativo_VTA")
    }
    _garantir_apresentacao_retroativos_e_aditivos(wb)
    for aba in ("RESULTADOS", "MEMORIA_RESULTADOS", "comparativo_VTA"):
        assert _formulas(wb[aba]) == antes[aba]
    wb.close()


def test_normaliza_nao_com_til_sem_tocar_sim_e_nao():
    wb = load_workbook(TEMPLATE_COLETA_OFICIAL, data_only=False)
    ws = wb["aditivos"]
    ws["H2"] = "Não"
    ws["H3"] = "Sim"
    ws["H4"] = "Nao"
    _garantir_apresentacao_retroativos_e_aditivos(wb)
    assert [ws[f"H{linha}"].value for linha in range(2, 5)] == ["Nao", "Sim", "Nao"]
    wb.close()
