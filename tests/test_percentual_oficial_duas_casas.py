"""REGRA PETREA — percentual oficial do ciclo fechado em 2 casas.

Caso real (C1/ICTI): variacao bruta 4,052187881255853% -> percentual oficial
4,05% -> fator proprio oficial 1,0405. Depois do fechamento, nenhum calculo
financeiro pode usar o fator bruto 1,0405218788125585; a variacao bruta
sobrevive apenas como memoria tecnica da apuracao.

Multiciclo: cada ciclo e fechado individualmente e o acumulado e o PRODUTO dos
fatores oficiais — nunca soma de percentuais nem composicao de fatores brutos.
"""
from __future__ import annotations

import io
import os
import sys
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from _coleta_oficial import _percentual, gerar_coleta_oficial_preenchida  # noqa: E402
from _coleta_reajuste import _percentual_ciclo  # noqa: E402
from _leitor_masterfile_v10 import ler_masterfile_v10  # noqa: E402
from _memoria_calculo import (  # noqa: E402
    TIPO_VARIACAO_BRUTA,
    ler_memoria_calculo,
    normalizar_memoria_calculo,
)
from _objeto_processo_reajuste import _fatores_acumulados_por_percentual  # noqa: E402
from _reajuste_utils import (  # noqa: E402
    APLICAR_VARIACAO_NEGATIVA,
    NEUTRALIZAR_VARIACAO_NEGATIVA,
    fator_acumulado_oficial,
    fator_oficial,
    fechar_percentual_oficial,
    resolver_tratamento_variacao_negativa,
)
from _templates_documentos import _extrair_dados  # noqa: E402
from test_reajuste_negativo import PAGINA_MULTIPLA, PAGINA_SIMPLES  # noqa: E402

RAW_C1 = 0.04052187881255853
FATOR_BRUTO_C1 = 1.0405218788125585
RAW_C2 = 0.04496
RAW_C3 = 0.03814

# (item, VU C0, VU esperado com 4,05%, VU antigo com o fator bruto)
ITENS_CASO_REAL = (
    ("1", 1043304.26, 1085558.08, 1085580.91),
    ("2", 737983.34, 767871.67, 767887.81),
    ("3", 31551.30, 32829.13, 32829.82),
    ("10", 196942.70, 204918.88, 204923.19),
    ("11", 24813.13, 25818.06, 25818.60),
    ("12", 31551.30, 32829.13, 32829.82),
)


def _round2(valor: float) -> float:
    """ROUND(x, 2) do template/aplicacao (meio para cima)."""
    return float(Decimal(repr(valor)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


# ------------------------------------------------------------------ unitarios

def test_fechamento_do_percentual_do_caso_real():
    assert fechar_percentual_oficial(RAW_C1) == 0.0405
    assert fator_oficial(RAW_C1) == 1.0405


def test_fator_oficial_nao_e_o_fator_bruto():
    assert 1.0 + RAW_C1 == FATOR_BRUTO_C1
    assert fator_oficial(RAW_C1) != FATOR_BRUTO_C1
    resolvido = resolver_tratamento_variacao_negativa(RAW_C1)
    assert resolvido["percentual_indice"] == RAW_C1  # memoria bruta preservada
    assert resolvido["percentual_aplicado"] == 0.0405
    assert resolvido["fator"] == 1.0405


@pytest.mark.parametrize(
    ("bruto", "oficial"),
    (
        (0.04045, 0.0405),        # meio para cima em pontos percentuais
        (0.040449999, 0.0404),
        (0.0375, 0.0375),         # ja oficial: idempotente
        (0.0, 0.0),
        (-0.00004, 0.0),          # sem -0,00%
        (-0.020349, -0.0203),
        (0.04496, 0.045),
        (0.03814, 0.0381),
    ),
)
def test_regra_decimal_meio_para_cima_e_idempotente(bruto, oficial):
    assert fechar_percentual_oficial(bruto) == oficial
    assert fechar_percentual_oficial(fechar_percentual_oficial(bruto)) == oficial
    assert str(fator_oficial(bruto)) == str(Decimal(1) + Decimal(repr(oficial)))


def test_ausencia_permanece_ausencia():
    assert fechar_percentual_oficial(None) is None
    assert fator_oficial(None) is None
    assert fator_acumulado_oficial([RAW_C1, None]) is None


def test_negativo_preserva_os_tres_estados_e_fecha_depois():
    pendente = resolver_tratamento_variacao_negativa(-0.020349)
    assert pendente["pendente"] is True
    assert pendente["percentual_aplicado"] is None
    assert pendente["fator"] is None

    neutro = resolver_tratamento_variacao_negativa(
        -0.020349, NEUTRALIZAR_VARIACAO_NEGATIVA
    )
    assert neutro["percentual_aplicado"] == 0.0
    assert neutro["fator"] == 1.0

    aplicado = resolver_tratamento_variacao_negativa(
        -0.020349, APLICAR_VARIACAO_NEGATIVA
    )
    assert aplicado["percentual_indice"] == -0.020349
    assert aplicado["percentual_aplicado"] == -0.0203
    assert aplicado["fator"] == 0.9797


def test_multiciclo_fecha_cada_ciclo_e_compoe_fatores_oficiais():
    acumulado = fator_acumulado_oficial([RAW_C1, RAW_C2, RAW_C3])
    assert acumulado == 1.0405 * 1.045 * 1.0381


def test_multiciclo_proibe_soma_de_percentuais():
    acumulado = fator_acumulado_oficial([RAW_C1, RAW_C2, RAW_C3])
    soma = 1.0 + 0.0405 + 0.045 + 0.0381
    assert abs(acumulado - soma) > 1e-3


def test_multiciclo_proibe_fatores_brutos_e_acumulado_fechado():
    acumulado = fator_acumulado_oficial([RAW_C1, RAW_C2, RAW_C3])
    brutos = (1 + RAW_C1) * (1 + RAW_C2) * (1 + RAW_C3)
    assert abs(acumulado - brutos) > 1e-6
    # fechar o ACUMULADO e reaplicar tambem e proibido
    acumulado_fechado = fator_oficial(brutos - 1)
    assert abs(acumulado - acumulado_fechado) > 1e-6


@pytest.mark.parametrize(("item", "vu_c0", "esperado", "antigo"), ITENS_CASO_REAL)
def test_vu_c1_do_caso_real_com_4_05(item, vu_c0, esperado, antigo):
    fator = resolver_tratamento_variacao_negativa(RAW_C1)["fator"]
    assert _round2(vu_c0 * fator) == esperado
    assert _round2(vu_c0 * FATOR_BRUTO_C1) == antigo  # o valor que sai de cena
    assert esperado != antigo


# ------------------------------------------------------- fronteira payload/XLS

def test_fronteira_do_xls_fecha_payload_so_com_variacao_bruta():
    assert _percentual({"percentual_indice": RAW_C1}) == 0.0405
    assert _percentual({"fator": FATOR_BRUTO_C1}) == 0.0405
    assert _percentual({"percentual_aplicado": 4.052187881255853}) == 0.0405
    assert _percentual_ciclo({"percentual_indice": RAW_C1}) == 0.0405
    assert _percentual_ciclo({"fator": FATOR_BRUTO_C1}) == 0.0405


def test_consumidor_do_objeto_compoe_sobre_percentuais_oficiais():
    fatores = _fatores_acumulados_por_percentual([
        {"ciclo": "C1", "indice_percentual": 0.0405},
        {"ciclo": "C2", "indice_percentual": 0.045},
    ])
    assert fatores["C1"] == 1.0405
    assert fatores["C2"] == pytest.approx(1.0405 * 1.045, abs=1e-15)


# --------------------------------------------------------------- memoria

def _res_icti(variacao: float) -> dict:
    import pandas as pd

    return {
        "variacao": variacao,
        "var": variacao,
        "metodo": "ICTI",
        "dados": pd.DataFrame({
            "data": [pd.Timestamp("2024-08-01")],
            "valor": [variacao * 100],
            "fator_mensal": [1 + variacao],
            "fator_acumulado_progressivo": [1 + variacao],
        }),
    }


def test_memoria_preserva_bruto_e_resultado_e_oficial():
    oficial = resolver_tratamento_variacao_negativa(RAW_C1)
    memoria = normalizar_memoria_calculo(
        _res_icti(RAW_C1), oficial["fator"], oficial["percentual_aplicado"]
    )
    assert [m["tipo"] for m in memoria] == ["MES", TIPO_VARIACAO_BRUTA, "RESULTADO"]
    bruta, resultado = memoria[-2], memoria[-1]
    assert bruta["variacao_final"] == RAW_C1
    assert bruta["fator_acumulado"] == FATOR_BRUTO_C1
    assert "não alimenta" in bruta["metodo_fonte"]
    assert resultado["variacao_final"] == 0.0405
    assert resultado["fator_acumulado"] == 1.0405
    assert [m["ordem"] for m in memoria] == [1, 2, 3]


def test_memoria_sem_linha_bruta_quando_nada_foi_fechado_ou_ha_tratamento():
    # ja oficial: nada a declarar
    memoria = normalizar_memoria_calculo(_res_icti(0.05), 1.05, 0.05)
    assert [m["tipo"] for m in memoria] == ["MES", "RESULTADO"]
    # precluso/neutralizado: o 0,00% nao e fechamento do bruto
    memoria = normalizar_memoria_calculo(_res_icti(RAW_C1), 1.0, 0.0)
    assert [m["tipo"] for m in memoria] == ["MES", "RESULTADO"]


# --------------------------------------------------------------- payloads

def _ciclo(nome: str, ano: int, bruto: float) -> dict:
    resolvido = resolver_tratamento_variacao_negativa(bruto)
    return {
        "ciclo": nome,
        "data_base": f"01/08/{ano}",
        "periodo_inicio": f"01/08/{ano}",
        "periodo_fim": f"31/07/{ano + 1}",
        "financeiro_inicio": f"01/08/{ano + 1}",
        "percentual_indice": bruto,
        "percentual_aplicado": resolvido["percentual_aplicado"],
        "variacao": resolvido["percentual_aplicado"],
        "fator": resolvido["fator"],
        "situacao": "TEMPESTIVO",
        "situacao_aplicada": "TEMPESTIVO",
        "memoria_calculo": normalizar_memoria_calculo(
            _res_icti(bruto), resolvido["fator"], resolvido["percentual_aplicado"]
        ),
    }


def _dados(*brutos: float) -> dict:
    ciclos = [
        _ciclo(f"C{i}", 2024 + i - 1, bruto) for i, bruto in enumerate(brutos, start=1)
    ]
    acumulado = fator_acumulado_oficial(brutos)
    fator = 1.0
    for ciclo in ciclos:
        fator *= ciclo["fator"]
        ciclo["fator_acumulado"] = fator
    return {
        "origem": "Reajuste Simples" if len(brutos) == 1 else "Reajuste Represado",
        "indice": "ICTI",
        "data_base_original": "01/08/2024",
        "fator": acumulado,
        "fator_acumulado": acumulado,
        "variacao": acumulado - 1,
        "variacao_acumulada": acumulado - 1,
        "ciclos": ciclos,
    }


@pytest.fixture(scope="module")
def coleta_c1() -> bytes:
    return gerar_coleta_oficial_preenchida(_dados(RAW_C1))


@pytest.fixture(scope="module")
def coleta_multiciclo() -> bytes:
    return gerar_coleta_oficial_preenchida(_dados(RAW_C1, RAW_C2, RAW_C3))


def test_xlsx_parametros_recebe_percentual_oficial(coleta_c1):
    wb = load_workbook(io.BytesIO(coleta_c1))
    par = wb["parametros"]
    assert par["E3"].value == 0.0405
    # a cadeia do XLS continua 100% formula sobre parametros!E
    assert par["F3"].value.startswith("=") and "(1+E3)" in par["F3"].value
    assert par["D12"].value == '=IF(C12="","",1+C12)'
    assert wb["historico_VU"]["K3"].value.startswith("=")


def test_xlsx_legado_so_com_bruto_tambem_sai_oficial():
    dados = _dados(RAW_C1)
    for chave in ("percentual_aplicado", "variacao", "fator", "fator_acumulado"):
        dados["ciclos"][0].pop(chave)
    wb = load_workbook(io.BytesIO(gerar_coleta_oficial_preenchida(dados)))
    assert wb["parametros"]["E3"].value == 0.0405


def test_xlsx_memoria_demonstra_bruto_e_oficial(coleta_c1):
    memoria = ler_memoria_calculo(load_workbook(io.BytesIO(coleta_c1))["parametros"])
    tipos = [r["tipo"] for r in memoria["C1"]]
    assert tipos[-2:] == [TIPO_VARIACAO_BRUTA, "RESULTADO"]
    assert memoria["C1"][-2]["variacao_final"] == RAW_C1
    assert memoria["C1"][-1]["variacao_final"] == 0.0405
    assert memoria["C1"][-1]["fator_acumulado"] == 1.0405


def test_xlsx_multiciclo_grava_cada_ciclo_fechado(coleta_multiciclo):
    par = load_workbook(io.BytesIO(coleta_multiciclo))["parametros"]
    assert [par[f"E{r}"].value for r in (3, 4, 5)] == [0.0405, 0.045, 0.0381]


def _com_pc(conteudo: bytes, data_pc: datetime, valor: float) -> bytes:
    wb = load_workbook(io.BytesIO(conteudo))
    pc = wb["itens_PC"]
    pc["A2"], pc["B2"], pc["D2"], pc["G2"] = "PC-0001", data_pc, valor, "Sim"
    wb["CONTROLE"]["B1"] = "PCs"
    saida = io.BytesIO()
    wb.save(saida)
    return saida.getvalue()


def _pc_lido(conteudo: bytes) -> dict:
    leitura = ler_masterfile_v10(conteudo)
    return next(
        p for p in leitura["itens_pc_v10"]["itens"] if p.get("numero_pc") == "PC-0001"
    )


def test_leitor_e_pc_usam_percentual_oficial(coleta_c1):
    leitura = ler_masterfile_v10(coleta_c1)
    c1 = leitura["parametros_v10"]["por_ciclo"]["C1"]
    assert c1["percentual_reajuste"] == 0.0405
    pc = _pc_lido(_com_pc(coleta_c1, datetime(2025, 9, 15), 737983.34))
    assert pc["ciclo"] == "C1"
    assert pc["fator_acumulado"] == 1.0405
    assert pc["valor_atualizado"] == 767871.67


def test_pc_multiciclo_infere_produto_dos_fatores_oficiais(coleta_multiciclo):
    pc = _pc_lido(_com_pc(coleta_multiciclo, datetime(2026, 9, 15), 1000.0))
    assert pc["ciclo"] == "C2"
    assert pc["fator_acumulado"] == pytest.approx(1.0405 * 1.045, abs=1e-12)
    assert pc["valor_atualizado"] == _round2(1000.0 * 1.0405 * 1.045)
    assert pc["valor_atualizado"] != _round2(1000.0 * (1 + RAW_C1) * (1 + RAW_C2))


def test_documentos_recebem_o_mesmo_percentual_oficial(coleta_c1):
    dados = _extrair_dados(ler_masterfile_v10(coleta_c1), None)
    ciclo = next(c for c in dados["ciclos"] if c["ciclo"] == "C1")
    assert ciclo["percentual_reajuste"] == 0.0405


# ------------------------------------------------------ calculadoras (origem)

def _wrapper_bruto(pagina: Path, por_ano: dict[int, float]) -> str:
    return f'''
import runpy
import sys

import pandas as pd

sys.path.insert(0, {str(ROOT)!r})
import _indice_utils

POR_ANO = {por_ano!r}

def _ist_bruto(data_inicio, *args, **kwargs):
    inicio = pd.Timestamp(data_inicio).replace(day=1)
    fim = inicio + pd.DateOffset(months=12)
    variacao = POR_ANO[inicio.year]
    return {{
        "variacao": variacao, "var": variacao,
        "i_ini": 100.0, "i_fim": 100.0 * (1.0 + variacao),
        "d_ini": inicio, "d_fim": fim, "p_ini": inicio, "p_fim": fim,
        "metodo": "IST determinístico de teste", "serie": "IST-TESTE",
        "dados": pd.DataFrame({{"data": [inicio, fim],
                               "indice": [100.0, 100.0 * (1.0 + variacao)]}}),
    }}

_original = _indice_utils.calcular_ist_numero_indice
try:
    _indice_utils.calcular_ist_numero_indice = _ist_bruto
    runpy.run_path({str(pagina)!r}, run_name="__main__")
finally:
    _indice_utils.calcular_ist_numero_indice = _original
'''


def _e_da_coleta(at) -> list:
    """parametros!E3:E6 da Coleta gerada a partir do estado da Calculadora."""
    assert at.get("download_button"), "botao da Coleta ausente"
    conteudo = at.session_state["dados_admissibilidade"]
    par = load_workbook(io.BytesIO(gerar_coleta_oficial_preenchida(conteudo)))["parametros"]
    return [par[f"E{r}"].value for r in range(3, 7)]


def test_calculadora_simples_fecha_na_origem():
    from streamlit.testing.v1 import AppTest

    at = AppTest.from_string(
        _wrapper_bruto(PAGINA_SIMPLES, {2023: RAW_C1, 2024: RAW_C1}),
        default_timeout=180,
    ).run()
    at.date_input[0].set_value(date(2023, 2, 1))
    at.run()
    at.date_input[1].set_value(date(2024, 2, 1))
    at.run()
    next(b for b in at.button if "Processar" in str(b.label)).click()
    at.run()

    adm = at.session_state["dados_admissibilidade"]
    ciclo = adm["ciclos"][0]
    assert ciclo["percentual_indice"] == RAW_C1
    assert ciclo["percentual_aplicado"] == 0.0405
    assert ciclo["variacao"] == 0.0405
    assert ciclo["fator"] == 1.0405
    assert adm["fator"] == adm["fator_acumulado"] == 1.0405
    tipos = [r["tipo"] for r in ciclo["memoria_calculo"]]
    assert tipos[-2:] == [TIPO_VARIACAO_BRUTA, "RESULTADO"]
    assert _e_da_coleta(at)[0] == 0.0405


def test_calculadora_multiciclo_fecha_cada_ciclo_antes_de_compor():
    from streamlit.testing.v1 import AppTest

    at = AppTest.from_string(
        _wrapper_bruto(PAGINA_MULTIPLA, {2022: RAW_C1, 2023: RAW_C2, 2024: RAW_C3}),
        default_timeout=180,
    ).run()
    at.selectbox(key="rep_ciclo_final_analise").select("C3")
    at.run()
    next(b for b in at.button if "Processar" in str(b.label)).click()
    at.run()

    adm = at.session_state["dados_admissibilidade"]
    ciclos = {c["ciclo"]: c for c in adm["ciclos"]}
    assert [ciclos[n]["percentual_aplicado"] for n in ("C1", "C2", "C3")] == [
        0.0405, 0.045, 0.0381
    ]
    assert [ciclos[n]["fator"] for n in ("C1", "C2", "C3")] == [1.0405, 1.045, 1.0381]
    assert [ciclos[n]["percentual_indice"] for n in ("C1", "C2", "C3")] == [
        RAW_C1, RAW_C2, RAW_C3
    ]
    oficial = 1.0405 * 1.045 * 1.0381
    assert ciclos["C3"]["fator_acumulado"] == pytest.approx(oficial, abs=1e-15)
    assert adm["fator_acumulado"] == pytest.approx(oficial, abs=1e-15)
    bruto = (1 + RAW_C1) * (1 + RAW_C2) * (1 + RAW_C3)
    assert abs(adm["fator_acumulado"] - bruto) > 1e-6
    assert _e_da_coleta(at)[:3] == [0.0405, 0.045, 0.0381]


# ------------------------------------------------ XLSX real recalculado (Excel)

@pytest.fixture(scope="module")
def excel():
    client = pytest.importorskip("win32com.client")
    pythoncom = pytest.importorskip("pythoncom")
    pythoncom.CoInitialize()
    app = client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    yield app
    app.Quit()
    del app
    pythoncom.CoUninitialize()


def _cenario_excel(conteudo: bytes, destino: Path) -> Path:
    wb = load_workbook(io.BytesIO(conteudo))
    rem = wb["itens_Remanesc"]
    for linha, (item, vu_c0, _esperado, _antigo) in enumerate(ITENS_CASO_REAL, start=2):
        rem[f"A{linha}"], rem[f"B{linha}"], rem[f"C{linha}"] = item, 10, vu_c0
        rem[f"E{linha}"] = 4
    fin = wb["financeiro"]
    for linha in range(2, 80):
        competencia = fin[f"A{linha}"].value
        if isinstance(competencia, datetime) and competencia <= datetime(2026, 7, 1):
            fin[f"C{linha}"] = 187345.67
    wb["CONTROLE"]["B1"] = "Financeiro"
    wb.save(destino)
    return destino


@pytest.mark.skipif(
    os.environ.get("RUN_EXCEL_INTEGRATION") != "1",
    reason="defina RUN_EXCEL_INTEGRATION=1 para executar Excel COM",
)
def test_xlsx_real_recalculado_propaga_4_05_por_toda_a_cadeia(
    excel, coleta_c1, tmp_path
):
    caminho = _cenario_excel(coleta_c1, tmp_path / "c1_oficial.xlsx")
    book = excel.Workbooks.Open(str(caminho), UpdateLinks=0, ReadOnly=True, CorruptLoad=0)
    try:
        excel.CalculateFullRebuild()

        def v(aba, celula):
            return book.Worksheets(aba).Range(celula).Value

        assert v("parametros", "E3") == 0.0405
        assert v("parametros", "F3") == 1.0405
        assert v("parametros", "D12") == 1.0405
        assert v("RESULTADOS", "H5") == 1.0405
        for linha, (item, _vu, esperado, antigo) in enumerate(ITENS_CASO_REAL, start=2):
            assert str(v("historico_VU", f"A{linha}")) == item
            assert v("historico_VU", f"D{linha}") == esperado        # VU_C1
            assert v("historico_VU", f"D{linha}") != antigo
            assert v("itens_RC", f"E{linha + 1}") == esperado        # mesmo VU canonico
            assert v("itens_Remanesc", f"F{linha}") == _round2(4 * esperado)
        # retroativo C1 (financeiro) e remanescente atualizado com 4,05%
        assert float(v("RESULTADOS", "D17")) == _round2(12 * _round2(187345.67 * 0.0405))
        assert float(v("RESULTADOS", "C27")) == _round2(
            sum(_round2(4 * esperado) for *_x, esperado, _a in ITENS_CASO_REAL)
        )
    finally:
        book.Close(False)
