# -*- coding: utf-8 -*-
"""Regressao focal: referencia exata separada da competencia mensal."""
from __future__ import annotations

import io
import sys
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

RAIZ = Path(__file__).resolve().parents[1]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from _coleta_oficial import gerar_coleta_oficial_preenchida  # noqa: E402
from _reajuste_utils import (  # noqa: E402
    classificar_pedido_por_data_exata,
    referencia_exata_pedido_subsequente,
)


REFERENCIA = date(2025, 8, 23)
LIMITE = date(2025, 11, 21)


@pytest.mark.parametrize(
    "pedido",
    [
        date(2025, 1, 10),
        date(2025, 3, 15),
        date(2025, 5, 12),
        date(2025, 7, 31),
        date(2025, 8, 1),
        date(2025, 8, 22),
    ],
)
def test_todo_pedido_anterior_a_referencia_exata_e_adiantado(pedido):
    assert classificar_pedido_por_data_exata(pedido, REFERENCIA, LIMITE) == "ADIANTADO"


@pytest.mark.parametrize("pedido", [date(2025, 8, 23), date(2025, 8, 30)])
def test_data_igual_ou_posterior_na_janela_nao_e_adiantado(pedido):
    assert classificar_pedido_por_data_exata(pedido, REFERENCIA, LIMITE) == "TEMPESTIVO"


def test_precluso_preserva_janela_existente():
    assert (
        classificar_pedido_por_data_exata(date(2025, 11, 22), REFERENCIA, LIMITE)
        == "PRECLUSO"
    )


@pytest.mark.parametrize(
    ("pedido", "esperado"),
    [
        (date(2024, 8, 23), date(2025, 8, 23)),
        (date(2025, 5, 12), date(2026, 5, 12)),
        (date(2024, 2, 29), date(2025, 2, 28)),
    ],
)
def test_proxima_referencia_preserva_dia_com_convencao_calendaria(pedido, esperado):
    assert referencia_exata_pedido_subsequente(pedido) == esperado


def test_cadeia_reportada_propaga_pedido_tempestivo_e_depois_adiantado():
    """Tempestivo propaga o pedido; adiantado propaga a referencia atual.

    O pedido antecipado e recebido e computado normalmente, mas nao pode
    antecipar o nascimento da anualidade seguinte: o C4 nasce de 23/08/2025
    (referencia do C3), nao de 01/08/2025 (data do pedido antecipado).
    """
    referencia_c2 = referencia_exata_pedido_subsequente(date(2023, 8, 15))
    pedido_c2 = date(2024, 8, 23)
    referencia_c3 = referencia_exata_pedido_subsequente(pedido_c2)
    pedido_c3 = date(2025, 8, 1)
    referencia_c4 = referencia_exata_pedido_subsequente(referencia_c3)

    assert referencia_c2 == date(2024, 8, 15)
    assert classificar_pedido_por_data_exata(pedido_c2, referencia_c2, date(2024, 11, 13)) == "TEMPESTIVO"
    assert referencia_c3 == date(2025, 8, 23)
    assert classificar_pedido_por_data_exata(pedido_c3, referencia_c3, LIMITE) == "ADIANTADO"
    assert referencia_c4 == date(2026, 8, 23)


def test_fluxo_real_multiciclo_exibe_e_propaga_as_referencias_exatas():
    from streamlit.testing.v1 import AppTest

    pagina = RAIZ / "pages" / "02_Calculo_Represados.py"
    at = AppTest.from_file(str(pagina), default_timeout=300)
    at.run()
    assert not at.exception

    at.date_input[2].set_value(date(2023, 8, 15))
    at.selectbox(key="rep_ciclo_inicial_analise").select("C2")
    at.run()
    at.selectbox(key="rep_ciclo_final_analise").select("C4")
    at.run()

    at.date_input(key="p2_20240815").set_value(date(2024, 8, 23))
    at.run()
    at.date_input(key="p3_20250823").set_value(date(2025, 8, 1))
    at.run()

    assert not at.exception
    # O C3 foi pedido de forma antecipada (01/08/2025): o C4 continua nascendo
    # da referencia do C3 (23/08/2025 + 12m), sem antecipacao da anualidade.
    assert at.date_input(key="p4_20260823").value == date(2026, 8, 23)
    resumo = at.dataframe[0].value.to_dict("records")
    assert [linha["Referência exata"] for linha in resumo] == [
        "15/08/2024",
        "23/08/2025",
        "23/08/2026",
    ]
    assert resumo[1]["Situação preliminar"] == "⚠️ ADIANTADO"
    assert [linha["Início financeiro"] for linha in resumo] == [
        "01/08/2024",
        "01/08/2025",
        "01/08/2026",
    ]


def _resumo_referencia_23_08_2025(pedido):
    """Roda a pagina real com referencia exata 23/08/2025 (C1) e C2 visivel."""
    from streamlit.testing.v1 import AppTest

    pagina = RAIZ / "pages" / "02_Calculo_Represados.py"
    at = AppTest.from_file(str(pagina), default_timeout=300)
    at.run()
    assert not at.exception

    for campo in at.date_input:
        if "Data-base do ciclo" in str(campo.label):
            campo.set_value(date(2024, 8, 23))
            break
    at.run()
    at.selectbox(key="rep_ciclo_final_analise").select("C2")
    at.run()

    at.date_input(key="p1_20250823").set_value(pedido)
    at.run()
    assert not at.exception
    return at.dataframe[0].value.to_dict("records")


@pytest.mark.parametrize(
    ("pedido", "situacao", "inicio_financeiro", "proxima_referencia"),
    [
        # Adiantado (mes anterior, mesmo mes e vespera): a antecipacao do
        # pedido nunca antecipa a referencia do ciclo seguinte.
        (date(2025, 7, 15), "⚠️ ADIANTADO", "01/08/2025", "23/08/2026"),
        (date(2025, 8, 1), "⚠️ ADIANTADO", "01/08/2025", "23/08/2026"),
        (date(2025, 8, 22), "⚠️ ADIANTADO", "01/08/2025", "23/08/2026"),
        # Tempestivo: a data efetiva do pedido alimenta o ciclo seguinte.
        (date(2025, 8, 23), "✅ TEMPESTIVO", "01/08/2025", "23/08/2026"),
        (date(2025, 8, 30), "✅ TEMPESTIVO", "01/08/2025", "30/08/2026"),
        # Tempestivo*: idem, preservando o retardo dos efeitos financeiros.
        (date(2025, 9, 1), "✅ TEMPESTIVO*", "01/09/2025", "01/09/2026"),
        (date(2025, 10, 15), "✅ TEMPESTIVO*", "01/10/2025", "15/10/2026"),
    ],
)
def test_referencia_seguinte_so_e_alimentada_por_pedido_tempestivo(
    pedido, situacao, inicio_financeiro, proxima_referencia
):
    resumo = _resumo_referencia_23_08_2025(pedido)
    assert resumo[0]["Referência exata"] == "23/08/2025"
    assert resumo[0]["Situação preliminar"] == situacao
    assert resumo[0]["Início financeiro"] == inicio_financeiro
    assert resumo[1]["Referência exata"] == proxima_referencia


def test_pedido_adiantado_nao_faz_o_ciclo_seguinte_nascer_no_mes_antecipado():
    """Sem sobreposicao: pedido de julho nao puxa o C2 para julho/2026."""
    resumo = _resumo_referencia_23_08_2025(date(2025, 7, 15))
    assert resumo[0]["Situação preliminar"] == "⚠️ ADIANTADO"
    assert resumo[1]["Referência exata"] == "23/08/2026"
    assert not resumo[1]["Referência exata"].startswith("15/07")


def test_tempestivo_asterisco_preserva_competencias_sem_efeito_no_xls():
    """Pedido 15/10/2025 sobre referencia 23/08/2025: 08 e 09/2025 sem efeito."""
    payload = _payload_xls("✅ TEMPESTIVO*")
    payload["ciclos"][0]["data_pedido"] = "15/10/2025"
    payload["ciclos"][0]["financeiro_inicio"] = "01/10/2025"
    payload["ciclos"][0]["efeito_financeiro_retardado"] = True
    wb = load_workbook(
        io.BytesIO(gerar_coleta_oficial_preenchida(payload)), data_only=False
    )
    try:
        parametros = wb["parametros"]
        assert "TEMPESTIVO*" in str(parametros["G3"].value)
        assert parametros["A3"].value == "Sim"
        assert parametros["E3"].value == pytest.approx(0.05)
        assert str(parametros["F3"].value).startswith("=")
        inicio = parametros["H3"].value
        assert (inicio.date() if hasattr(inicio, "date") else inicio) == date(2025, 10, 1)

        financeiro = wb["financeiro"]
        grade = {}
        for linha in range(2, financeiro.max_row + 1):
            competencia = financeiro[f"A{linha}"].value
            if competencia is not None:
                grade[(competencia.year, competencia.month)] = financeiro[f"G{linha}"].value
        assert grade[(2025, 8)] == "Nao"
        assert grade[(2025, 9)] == "Nao"
        assert grade[(2025, 10)] == "Sim"
        assert grade[(2025, 11)] == "Sim"
    finally:
        wb.close()


def test_as_duas_calculadoras_usam_a_classificacao_unica_de_adiantado():
    simples = (RAIZ / "pages" / "01_Calculo_Simples.py").read_text(encoding="utf-8")
    multiplos = (RAIZ / "pages" / "02_Calculo_Represados.py").read_text(encoding="utf-8")

    for fonte in (simples, multiplos):
        assert "classificar_pedido_por_data_exata(" in fonte
        assert "ADMISSÍVEL - RESSALVA" not in fonte


def _payload_xls(situacao: str) -> dict:
    return {
        "origem": "Reajustes Múltiplos",
        "indice": "IPCA",
        "data_base_original": "23/08/2024",
        "ciclos": [
            {
                "ciclo": "C1",
                "data_base": "23/08/2024",
                "data_pedido": "01/08/2025",
                "percentual_aplicado": 0.05,
                "fator": 1.05,
                "fator_acumulado": 1.05,
                "financeiro_inicio": "01/08/2025",
                "objeto_analise_atual": True,
                "situacao": situacao,
            }
        ],
    }


def _validacoes(ws):
    return tuple(
        (dv.type, str(dv.sqref), dv.formula1, dv.formula2)
        for dv in ws.data_validations.dataValidation
    )


def test_adiantado_permanece_computavel_e_nao_altera_xls_ou_financeiro():
    wb_tempestivo = load_workbook(
        io.BytesIO(gerar_coleta_oficial_preenchida(_payload_xls("✅ TEMPESTIVO"))),
        data_only=False,
    )
    wb_adiantado = load_workbook(
        io.BytesIO(gerar_coleta_oficial_preenchida(_payload_xls("⚠️ ADIANTADO"))),
        data_only=False,
    )
    try:
        assert wb_tempestivo.sheetnames == wb_adiantado.sheetnames
        diferencas = []
        for nome in wb_tempestivo.sheetnames:
            antes = wb_tempestivo[nome]
            depois = wb_adiantado[nome]
            assert antes.max_row == depois.max_row
            assert antes.max_column == depois.max_column
            assert tuple(antes.merged_cells.ranges) == tuple(depois.merged_cells.ranges)
            assert _validacoes(antes) == _validacoes(depois)
            assert antes.protection.sheet == depois.protection.sheet
            for linha in antes.iter_rows():
                for celula_antes in linha:
                    celula_depois = depois[celula_antes.coordinate]
                    assert celula_antes.style_id == celula_depois.style_id
                    if celula_antes.value != celula_depois.value:
                        diferencas.append((nome, celula_antes.coordinate))

        assert diferencas == [("parametros", "G3")]
        parametros = wb_adiantado["parametros"]
        assert parametros["A3"].value == "Sim"
        assert parametros["E3"].value == pytest.approx(0.05)
        assert str(parametros["F3"].value).startswith("=")
        assert "ADIANTADO" in str(parametros["G3"].value)
        assert parametros["B12"].value == "=A3"
        assert parametros["D12"].value == '=IF(C12="","",1+C12)'
        assert parametros["E12"].value == '=IF(B12="Sim",E11*D12,E11)'

        for nome in ("financeiro", "CONTROLE"):
            antes = wb_tempestivo[nome]
            depois = wb_adiantado[nome]
            assert [c.value for row in antes.iter_rows() for c in row] == [
                c.value for row in depois.iter_rows() for c in row
            ]
    finally:
        wb_tempestivo.close()
        wb_adiantado.close()
