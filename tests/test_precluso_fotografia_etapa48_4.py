# -*- coding: utf-8 -*-
"""HOTFIX 48.4 — fotografia fisica do ciclo PRECLUSO.

O PRECLUSO sem acordo continua SEM efeito financeiro
(`referencia_exata_efeito = None`, `inicio_efeito_financeiro = None`), mas o
ciclo EXISTE fisicamente na analise: sua fotografia e a referencia exata ja
calculada pela pagina (`data_referencia_exata` = d_aniv), nunca a data mensal
de `parametros!C`.

Cenario do enunciado (data-base lateral 01/02/2022):

    C1  apto 01/02/2023  pedido 01/02/2023  TEMPESTIVO   -> foto 01/02/2023
    C2  apto 01/02/2024  pedido 10/03/2024  TEMPESTIVO*  -> foto 10/03/2024
    C3  apto 10/03/2025  pedido 10/06/2025  PRECLUSO     -> foto 10/03/2025
    C4  fora da analise                                  -> parametros!I vazio
"""
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

DATA_BASE_LATERAL = date(2022, 2, 1)
PEDIDOS = (
    ("p1_20230201", date(2023, 2, 1)),
    ("p2_20240201", date(2024, 3, 10)),
    ("p3_20250310", date(2025, 6, 10)),
)


@pytest.fixture(scope="module")
def payload() -> dict:
    """Roda a pagina real com C1+C2+C3 (C3 PRECLUSO) e devolve o payload."""
    pytest.importorskip("streamlit")
    from streamlit.testing.v1 import AppTest

    at = AppTest.from_file(
        str(RAIZ / "pages" / "02_Calculo_Represados.py"), default_timeout=600
    )
    at.run()
    assert not at.exception

    for campo in at.date_input:
        if "Data-base de referência" in str(campo.label):
            campo.set_value(DATA_BASE_LATERAL)
            break
    at.run()
    at.selectbox(key="rep_ciclo_final_analise").select("C3")
    at.run()
    for chave, data in PEDIDOS:
        at.date_input(key=chave).set_value(data)
        at.run()
    # o payload so nasce depois do comando explicito do usuario
    for botao in at.button:
        if "Processar Análise" in str(botao.label):
            botao.click()
            break
    at.run()
    assert not at.exception

    try:
        dados = at.session_state["dados_admissibilidade"]
    except KeyError:  # pragma: no cover - indice sem cobertura no periodo
        pytest.skip("payload indisponivel (indice sem cobertura no periodo)")
    return dados


@pytest.fixture(scope="module")
def parametros(payload):
    wb = load_workbook(
        io.BytesIO(gerar_coleta_oficial_preenchida(payload)), data_only=False
    )
    return wb, wb["parametros"]


def _d(celula):
    v = celula.value
    return v.date() if hasattr(v, "date") else v


def _ciclo(payload, nome):
    return next(c for c in payload["ciclos"] if c["ciclo"] == nome)


# ------------------------------------------------------------- C3 (PRECLUSO)
def test_precluso_transporta_a_referencia_exata_como_fotografia(payload):
    c3 = _ciclo(payload, "C3")
    assert "PRECLUSO" in c3["situacao"]
    assert c3["data_abertura_fisica_exata"] == "10/03/2025"


def test_precluso_continua_sem_efeito_financeiro(payload):
    c3 = _ciclo(payload, "C3")
    # a classificacao e a ausencia de efeito financeiro nao sao tocadas
    assert "PRECLUSO" in c3["situacao"]
    assert not c3["financeiro_inicio"]
    assert c3.get("efeito_financeiro_retardado") is False


def test_ciclos_com_efeito_seguem_transportando_a_propria_referencia(payload):
    assert _ciclo(payload, "C1")["data_abertura_fisica_exata"] == "01/02/2023"
    assert _ciclo(payload, "C2")["data_abertura_fisica_exata"] == "10/03/2024"
    # TEMPESTIVO* preserva o dia exato do pedido, sem mensalizar
    assert _ciclo(payload, "C2")["financeiro_inicio"] == "01/03/2024"


# --------------------------------------------------------------- XLS gerado
def test_parametros_i_recebe_a_fotografia_de_cada_ciclo_real(parametros):
    _, par = parametros
    assert _d(par["I3"]) == date(2023, 2, 1)     # C1
    assert _d(par["I4"]) == date(2024, 3, 10)    # C2
    assert _d(par["I5"]) == date(2025, 3, 10)    # C3 PRECLUSO
    # a cadeia mensal do PRECLUSO segue existindo e nao vira fotografia
    assert _d(par["I5"]) != _d(par["C5"])


def test_c4_fora_da_analise_permanece_sem_fotografia(parametros):
    _, par = parametros
    assert par["I6"].value is None


def test_cabecalhos_de_itens_remanesc_no_cenario(parametros):
    wb, _ = parametros
    rem = wb["itens_Remanesc"]
    # as formulas leem parametros!I; a prova de valor esta no smoke Excel.
    # A data ficou nas colunas de entrada manual (QTD. REMANESCENTE); as de
    # execucao passaram a se declarar calculadas.
    for coord, linha in (("E1", 3), ("G1", 4), ("I1", 5), ("K1", 6)):
        assert f"parametros!$I${linha}" in str(rem[coord].value)
        assert "parametros!$C$" not in str(rem[coord].value)
    for coord in ("M1", "O1", "Q1", "S1"):
        assert "automaticamente" in str(rem[coord].value)
