"""Etapa 1.1 - Execucao Historica Fora da Analise (bloco EIXO 5).

Cobre a estrutura estatica do bloco no template oficial e a invariancia do VTA
(B23/B26 nao dependem do bloco). Os cenarios numericos completos com Microsoft
Excel real ficam no teste opt-in ao final (RUN_EXCEL_INTEGRATION=1).
"""
from __future__ import annotations

import datetime
import os
from pathlib import Path

import pytest
from openpyxl import load_workbook

TEMPLATE = Path(__file__).resolve().parent.parent / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

LIN_HDR = 257
LIN_C0 = 258

CAB_BLOCO = [
    "CICLO", "FORA DA ANALISE?", "FONTE PRIMARIA", "COBERTURA DA FONTE",
    "QTD EXECUTADA RECONSTRUIDA", "QTD COBERTA PELA FONTE", "QTD NAO COBERTA",
    "FATOR INCREMENTAL DA APURACAO", "DIFERENCIAL POTENCIAL CALCULADO",
    "VALOR ADOTADO", "STATUS / CHECK",
]
CAB_REMANESC = [
    "QTD_COBERTA_C0", "QTD_NAO_COBERTA_C0", "DIFERENCIAL_C0",
    "QTD_COBERTA_C1", "QTD_NAO_COBERTA_C1", "DIFERENCIAL_C1",
    "QTD_COBERTA_C2", "QTD_NAO_COBERTA_C2", "DIFERENCIAL_C2",
    "QTD_COBERTA_C3", "QTD_NAO_COBERTA_C3", "DIFERENCIAL_C3",
    "QTD_COBERTA_C4", "QTD_NAO_COBERTA_C4", "DIFERENCIAL_C4",
]

VTA_HOMOLOGADO = {
    "B20": '=IF(COUNTIF(itens_Remanesc!$A$2:$A$200,"<>")=0,"",ROUND(SUMIF(itens_Remanesc!$A$2:$A$200,"<>",itens_Remanesc!$D$2:$D$200),2))',
    "B21": '=IF(B16="","",B16)',
    "B22": '=IF(OR(C35="",D35=""),"",ROUND(D35-C35,2))',
    "B23": '=IF(OR(B20="",B21="",B22=""),"",ROUND(B20+B21+B22,2))',
}


@pytest.fixture(scope="module")
def wb():
    return load_workbook(TEMPLATE, data_only=False)


def test_bloco_cabecalhos(wb):
    ws = wb["RESULTADOS"]
    assert str(ws["A256"].value).startswith("EIXO 5")
    achado = [ws.cell(LIN_HDR, c).value for c in range(1, 12)]
    assert achado == CAB_BLOCO
    assert [ws.cell(LIN_C0 + i, 1).value for i in range(5)] == ["C0", "C1", "C2", "C3", "C4"]


def test_helper_itens_remanesc_cabecalhos(wb):
    ws = wb["itens_Remanesc"]
    achado = [ws.cell(1, c).value for c in range(41, 56)]  # AO..BC
    assert achado == CAB_REMANESC


def test_vta_formulas_identicas_ao_homologado(wb):
    ws = wb["RESULTADOS"]
    for cel, esperado in VTA_HOMOLOGADO.items():
        assert ws[cel].value == esperado, f"{cel} divergiu do homologado"
    assert str(ws["B26"].value).startswith('=IF(AND(B24<>"",B25<>"")')


def test_vta_nao_depende_do_bloco_eixo5(wb):
    """Invariancia: a cadeia do VTA nunca referencia o bloco nem os auxiliares."""
    ws = wb["RESULTADOS"]
    chain = ["B20", "B21", "B22", "B23", "B26", "B16", "E15", "C35", "D35",
             "B32", "C32", "D32", "B33", "C33", "D33", "E10", "E11", "E12", "E13", "E14"]
    aux = ["AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY",
           "AZ", "BA", "BB", "BC"]
    for cel in chain:
        f = str(ws[cel].value or "")
        for p in aux:
            assert f"itens_Remanesc!{p}" not in f, f"{cel} referencia auxiliar {p}"
        for lin in range(256, 264):
            assert f"E{lin}" not in f and f"I{lin}" not in f


def test_valor_adotado_vazio_por_padrao(wb):
    ws = wb["RESULTADOS"]
    for i in range(5):
        assert ws.cell(LIN_C0 + i, 10).value in (None, "")  # coluna J


def test_valor_adotado_validacao_nao_negativa(wb):
    ws = wb["RESULTADOS"]
    encontrou = False
    for dv in ws.data_validations.dataValidation:
        sq = str(dv.sqref)
        if any(f"J{LIN_C0 + i}" in sq for i in range(5)):
            encontrou = True
            assert dv.type == "decimal"
            assert dv.operator == "greaterThanOrEqual"
            assert "0" in str(dv.formula1)
    assert encontrou, "validacao de VALOR ADOTADO (J258:J262) ausente"


def test_fator_incremental_reutiliza_canonico(wb):
    """DIFERENCIAL_Cn usa VU*(F_n - F_n/D_n), o mesmo fator do retroativo D11."""
    ws = wb["itens_Remanesc"]
    assert "parametros!$F$3-parametros!$F$3/parametros!$D$12" in str(ws["AT2"].value)
    assert "parametros!$F$2-parametros!$F$2/parametros!$D$11" in str(ws["AQ2"].value)
    assert "parametros!$F$3-parametros!$F$3/parametros!$D$12" in str(wb["RESULTADOS"]["H259"].value)


def test_reconciliavel_usa_verificacao_positiva_itens(wb):
    """Regra POSITIVA canonica: reconciliavel <=> $B$4="Itens" (dropdown
    Financeiro,PCs,Itens). Vazio/PCs/Financeiro/desconhecido nao reconciliaveis."""
    ws = wb["RESULTADOS"]
    for cel in ("D259", "F259", "G259", "I259", "K259"):
        assert '$B$4<>"Itens"' in str(ws[cel].value)
        assert '$B$4="Financeiro"' not in str(ws[cel].value)
        assert '$B$4="PCs"' not in str(ws[cel].value)


def test_cobertura_superior_classificada(wb):
    ws = wb["RESULTADOS"]
    assert "COBERTURA SUPERIOR" in str(ws["D259"].value)


def test_nao_coberta_usa_max_execucao_menos_cobertura(wb):
    ws = wb["itens_Remanesc"]
    assert ws["AS2"].value == '=IF(OR($A2="",M2=""),"",ROUND(MAX(M2-AR2,0),2))'


def test_colunas_auxiliares_ocultas_reexibiveis(wb):
    """AO:BC ocultas normalmente (hidden), nunca veryHidden."""
    ws = wb["itens_Remanesc"]
    from openpyxl.utils import column_index_from_string as cidx
    ocultas = [d for d in ws.column_dimensions.values() if d.hidden]
    # a faixa AO(41)..BC(55) deve estar coberta por dimensao(oes) hidden
    for col in range(cidx("AO"), cidx("BC") + 1):
        assert any(d.min <= col <= d.max for d in ocultas), f"coluna {col} nao oculta"
    # nunca veryHidden (atributo hidden simples)
    for ws2 in wb.worksheets:
        assert ws2.sheet_state != "veryHidden"


# --------------------------------------------------------------------------
# Cenarios numericos completos com Microsoft Excel real (opt-in).
# --------------------------------------------------------------------------
EXCEL = pytest.mark.skipif(
    os.environ.get("RUN_EXCEL_INTEGRATION") != "1",
    reason="defina RUN_EXCEL_INTEGRATION=1 para executar o Excel COM",
)


def _abrir_editar_ler(caminho, editar, ler):
    import gc
    import time
    import pythoncom
    import win32com.client

    def tentar(acao):
        ultimo = None
        for _ in range(30):
            try:
                return acao()
            except Exception as exc:  # RPC_E_CALL_REJECTED transitorio
                ultimo = exc
                cod = getattr(exc, "hresult", None) or (exc.args[0] if exc.args else None)
                if cod != -2147418111:
                    raise
                pythoncom.PumpWaitingMessages()
                time.sleep(0.2)
        raise ultimo

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    pasta = None
    try:
        pasta = tentar(lambda: excel.Workbooks.Open(str(Path(caminho).resolve()), UpdateLinks=0))
        time.sleep(0.4)
        tentar(lambda: editar(pasta))
        tentar(excel.CalculateFullRebuild)
        res = tentar(lambda: ler(pasta))
        tentar(lambda: pasta.Close(SaveChanges=False))
        pasta = None
        return res
    finally:
        if pasta is not None:
            pasta.Close(SaveChanges=False)
        excel.Quit()
        gc.collect()
        pythoncom.CoUninitialize()


def _cenario_base(tmp_path):
    from _coleta_oficial import gerar_coleta_oficial_preenchida
    dados = {
        "origem": "Etapa 1.1 COM", "indice": "IST",
        "data_base_original": "01/01/2023", "data_corte": datetime.date(2025, 12, 31),
        "ciclos": [
            {"ciclo": "C1", "data_inicio": datetime.date(2024, 1, 1), "data_fim": datetime.date(2024, 12, 31), "data_pedido": datetime.date(2024, 1, 1), "financeiro_inicio": datetime.date(2024, 1, 1), "percentual": 0.10},
            {"ciclo": "C2", "data_inicio": datetime.date(2025, 1, 1), "data_fim": datetime.date(2025, 12, 31), "data_pedido": datetime.date(2025, 1, 1), "financeiro_inicio": datetime.date(2025, 1, 1), "percentual": 0.10},
        ],
    }
    conteudo = gerar_coleta_oficial_preenchida(dados)
    caminho = tmp_path / "cenario.xlsx"
    caminho.write_bytes(conteudo)
    return caminho


def _editar_item(pasta, delta_c1=0.0, metodo=None, cons_c1=None):
    if metodo is not None:
        pasta.Worksheets("RESULTADOS").Range("B4").Value = metodo
    ir = pasta.Worksheets("itens_Remanesc")
    ir.Range("A2").Value = "ITEM-1"
    ir.Range("B2").Value = 100.0   # QTD_BASE_ORIGINAL
    ir.Range("C2").Value = 10.0    # VU_ORIGINAL
    ir.Range("E2").Value = 60.0    # QTD_REM_BASE_C1
    ir.Range("G2").Value = 30.0    # QTD_REM_BASE_C2
    if cons_c1 is not None:
        ic = pasta.Worksheets("itens_Consumidos")
        ic.Range("A2").Value = "ITEM-1"
        ic.Range("B2").Value = 100.0   # QTD_CONTRATADA
        ic.Range("C2").Value = 10.0    # VU_ORIGINAL
        ic.Range("G2").Value = float(cons_c1)  # QTD_CONS_C1
    if delta_c1:
        import pywintypes
        ad = pasta.Worksheets("aditivos")
        ad.Range("A2").Value = "ITEM-1"
        ad.Range("B2").Value = pywintypes.Time(datetime.datetime(2024, 6, 1))
        ad.Range("D2").Value = "ACRESCIMO" if delta_c1 > 0 else "SUPRESSAO"
        ad.Range("E2").Value = abs(delta_c1)
        ad.Range("H2").Value = "Nao"


def _ler_bloco_e_vta(pasta):
    r = pasta.Worksheets("RESULTADOS")
    linhas = {}
    for i, nome in enumerate(["C0", "C1", "C2", "C3", "C4"]):
        row = 258 + i
        linhas[nome] = {
            "fora": r.Range(f"B{row}").Value, "cobertura": r.Range(f"D{row}").Value,
            "exec": r.Range(f"E{row}").Value, "nao_coberta": r.Range(f"G{row}").Value,
            "fator": r.Range(f"H{row}").Value, "diferencial": r.Range(f"I{row}").Value,
            "adotado": r.Range(f"J{row}").Value,
        }
    return {"linhas": linhas, "B20": r.Range("B20").Value,
            "B23": r.Range("B23").Value, "B26": r.Range("B26").Value}


# C1 esta na apuracao (computado); exec C1 = REM_C1(60) - REM_C2(30) = 30.
# percentual C1 = 10% -> fator incremental = F3(1,10) - F3/D12(1,0) = 0,10.


@EXCEL
def test_com_cobertura_completa(tmp_path):
    """A: exec 30, consumo 30 -> COMPLETA, nao coberta 0, diferencial 0."""
    caminho = _cenario_base(tmp_path)
    res = _abrir_editar_ler(
        caminho, lambda p: _editar_item(p, metodo="Itens", cons_c1=30.0), _ler_bloco_e_vta)
    c1 = res["linhas"]["C1"]
    assert abs(c1["exec"] - 30.0) < 0.01
    assert c1["cobertura"] == "COMPLETA"
    assert abs(c1["nao_coberta"] - 0.0) < 0.01
    assert abs(c1["diferencial"] - 0.0) < 0.01
    assert c1["adotado"] in (None, "")                       # VALOR ADOTADO vazio
    assert res["B23"] == "" or isinstance(res["B23"], (int, float))
    assert res["B26"] == "" or isinstance(res["B26"], (int, float))


@EXCEL
def test_com_cobertura_parcial(tmp_path):
    """B: exec 30, consumo 10 -> PARCIAL, nao coberta 20, diferencial 20*10*0,10=20."""
    caminho = _cenario_base(tmp_path)
    res = _abrir_editar_ler(
        caminho, lambda p: _editar_item(p, metodo="Itens", cons_c1=10.0), _ler_bloco_e_vta)
    c1 = res["linhas"]["C1"]
    assert c1["cobertura"] == "PARCIAL"
    assert abs(c1["nao_coberta"] - 20.0) < 0.01
    assert abs(c1["diferencial"] - 20.0) < 0.01              # somente sobre 20 nao cobertas


@EXCEL
def test_com_cobertura_superior(tmp_path):
    """C: exec 30, consumo 35 -> nao coberta 0 (nunca negativa); status sinaliza."""
    caminho = _cenario_base(tmp_path)
    res = _abrir_editar_ler(
        caminho, lambda p: _editar_item(p, metodo="Itens", cons_c1=35.0), _ler_bloco_e_vta)
    c1 = res["linhas"]["C1"]
    assert abs(c1["nao_coberta"] - 0.0) < 0.01
    assert c1["cobertura"] == "COBERTURA SUPERIOR"


@EXCEL
def test_com_metodo_pcs_nao_reconciliavel(tmp_path):
    """D: metodo PCs -> NAO RECONCILIAVEL, nao coberta n/d, diferencial vazio."""
    caminho = _cenario_base(tmp_path)
    res = _abrir_editar_ler(
        caminho, lambda p: _editar_item(p, metodo="PCs", cons_c1=10.0), _ler_bloco_e_vta)
    c1 = res["linhas"]["C1"]
    assert c1["cobertura"] == "NAO RECONCILIAVEL"
    assert c1["nao_coberta"] == "n/d"
    assert c1["diferencial"] in (None, "")


@EXCEL
def test_com_metodo_vazio_ou_invalido_nao_reconciliavel(tmp_path):
    """E: metodo vazio/desconhecido -> nao reconciliavel; nunca assume Itens."""
    caminho = _cenario_base(tmp_path)
    for metodo in ("", "Itens consumidos", "XYZ"):
        res = _abrir_editar_ler(
            caminho, lambda p, m=metodo: _editar_item(p, metodo=m, cons_c1=10.0), _ler_bloco_e_vta)
        assert res["linhas"]["C1"]["cobertura"] == "NAO RECONCILIAVEL", metodo
        assert res["linhas"]["C1"]["diferencial"] in (None, ""), metodo


@EXCEL
def test_com_cenario_acrescimo(tmp_path):
    caminho = _cenario_base(tmp_path)
    res = _abrir_editar_ler(caminho, lambda p: _editar_item(p, delta_c1=+20.0, metodo="Itens"), _ler_bloco_e_vta)
    # execucao C1 = REM_C1 + DELTA - REM_C2 = 60 + 20 - 30 = 50 (nao 30)
    assert abs(res["linhas"]["C1"]["exec"] - 50.0) < 0.01


@EXCEL
def test_com_cenario_supressao(tmp_path):
    caminho = _cenario_base(tmp_path)
    res = _abrir_editar_ler(caminho, lambda p: _editar_item(p, delta_c1=-20.0, metodo="Itens"), _ler_bloco_e_vta)
    # execucao C1 = MAX(60 - 20 - 30, 0) = 10 ; supressao nao vira execucao negativa
    assert abs(res["linhas"]["C1"]["exec"] - 10.0) < 0.01


@EXCEL
def test_com_aplicador_fail_closed(tmp_path):
    """Segunda aplicacao sobre template ja modificado deve falhar (fail-closed)."""
    import shutil
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "aplicar_eh", TEMPLATE.parent.parent / "tools" / "aplicar_execucao_historica_template.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    aplicar = mod.aplicar
    alvo = tmp_path / "ja_aplicado.xlsx"
    shutil.copyfile(TEMPLATE, alvo)  # TEMPLATE ja tem o bloco
    import time
    antes = alvo.read_bytes()
    erro = None
    for _ in range(5):  # tolera RPC_E_CALL_REJECTED transitorio entre instancias Excel
        try:
            aplicar(alvo, alvo)
            erro = "SEM ERRO"  # nao deveria concluir silenciosamente
            break
        except (ValueError, RuntimeError) as exc:
            erro = exc  # fail-closed detectou bloco ja aplicado
            break
        except Exception as exc:  # transitorio de COM: tentar de novo
            if getattr(exc, "args", (None,))[0] == -2147418111:
                time.sleep(1.0)
                continue
            raise
    assert isinstance(erro, (ValueError, RuntimeError)), f"esperava fail-closed, obtive {erro!r}"
    assert alvo.read_bytes() == antes  # template intacto (nao duplicou/re-salvou)
