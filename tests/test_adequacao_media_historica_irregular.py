"""Adequacao Orcamentaria — projecao Automatica com historico irregular.

Protege esta entrega (Cl8us 12.18):
  * historico valido porem irregular: a Automatica usa a media mensal
    historica ja calculada (meses sem PC no divisor) em todos os meses
    restantes da vigencia, como estimativa de baixa confiabilidade;
  * "Valor a usar" nasce preenchido com a sugestao; campo vazio = R$ 0,00
    (sem recuperar a media); valor informado prevalece;
  * resultado identificado como estimado ("Complementacao orcamentaria
    estimada"), inclusive no XLSX;
  * sem historico valido nada e inventado;
  * as demais premissas (Mensal, Por ciclo, Manual) seguem inalteradas.
"""
from __future__ import annotations

import sys
from io import BytesIO
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from _adequacao_ui import (  # noqa: E402
    NOTA_RESUMO_CONFIRMADA,
    NOTA_RESUMO_ESTIMADA,
    ROTULO_MEDIA_IRREGULAR,
    calcular_projecao,
    gerar_xlsx_projecao,
    montar_base_editor,
)
from test_adequacao_pc_janela_potencial import (  # noqa: E402
    _blob,
    _run,
    _sessao_pc,
    _valor_do_card,
)

MEDIA = 2166.6666666666665  # 39.000,00 / 18 meses (mar/25 a ago/26)
FATOR = 1.0289
PERIODOS = list(pd.period_range("2026-09", "2026-11", freq="M"))
MAPA = {p.to_timestamp().date(): MEDIA for p in PERIODOS}
ESTIMADA = "COMPLEMENTAÇÃO ORÇAMENTÁRIA ESTIMADA"
CONFIRMADA = "COMPLEMENTAÇÃO CONFIRMADA"


def _editor_irregular():
    return montar_base_editor(PERIODOS, MEDIA, base_por_competencia=MAPA,
                              preencher_valor_a_usar=True)


def _proj(df):
    return calcular_projecao(df, MEDIA, FATOR, base_por_competencia=MAPA,
                             origem_automatica=ROTULO_MEDIA_IRREGULAR,
                             vazio_igual_zero=True)


# ------------------------------------------------------------ editor / regras

def test_valor_a_usar_nasce_com_a_sugestao_da_media():
    df = _editor_irregular()
    assert list(df["Valor informado pelo fiscal"]) == ["2.166,67"] * 3
    assert list(df["Base automática pela média"]) == ["2.166,67"] * 3


def test_sugestao_intacta_reproduz_exatamente_a_premissa_mensal():
    proj = _proj(_editor_irregular())
    legado = calcular_projecao(montar_base_editor(PERIODOS, MEDIA), MEDIA, FATOR)
    assert list(proj["Valor base considerado"]) == list(legado["Valor base considerado"])
    assert list(proj["Diferença futura a adequar"]) == list(legado["Diferença futura a adequar"])
    assert set(proj["Origem"]) == {ROTULO_MEDIA_IRREGULAR}


def test_campo_vazio_vale_zero_sem_recuperar_a_media():
    df = _editor_irregular()
    df.loc[0, "Valor informado pelo fiscal"] = ""
    df.loc[1, "Valor informado pelo fiscal"] = None
    proj = _proj(df)
    assert proj.loc[0, "Valor base considerado"] == 0.0
    assert proj.loc[0, "Diferença futura a adequar"] == 0.0
    assert proj.loc[1, "Valor base considerado"] == 0.0
    assert proj.loc[0, "Premissa usada"] == "Campo vazio = R$ 0,00"
    assert proj.loc[2, "Valor base considerado"] == pytest.approx(2166.67)


def test_valor_informado_prevalece_sobre_a_sugestao():
    df = _editor_irregular()
    df.loc[0, "Valor informado pelo fiscal"] = "5.000,00"
    df.loc[1, "Valor informado pelo fiscal"] = "0"
    proj = _proj(df)
    assert proj.loc[0, "Valor base considerado"] == 5000.0
    assert proj.loc[0, "Diferença futura a adequar"] == pytest.approx(144.50)
    assert proj.loc[0, "Origem"] == "Valor informado pelo fiscal"
    assert proj.loc[1, "Valor base considerado"] == 0.0


def test_caminhos_existentes_nao_mudam_sem_a_flag():
    """Sem a flag, vazio continua = projecao automatica (Mensal/Por ciclo/Manual)."""
    df = montar_base_editor(PERIODOS, MEDIA, base_por_competencia=MAPA)
    assert list(df["Valor informado pelo fiscal"]) == [""] * 3
    proj = calcular_projecao(df, MEDIA, FATOR, base_por_competencia=MAPA)
    assert proj.loc[0, "Valor base considerado"] == pytest.approx(2166.67)
    legado = calcular_projecao(montar_base_editor(PERIODOS, MEDIA), MEDIA, FATOR)
    assert set(legado["Origem"]) == {"Média dos últimos 6 meses"}


# ------------------------------------------------------------------- XLSX

def _resumo_xlsx(xlsx):
    ws = load_workbook(BytesIO(xlsx))["RESUMO"]
    return ws["A2"].value, {ws.cell(r, 1).value: ws.cell(r, 2).value
                            for r in range(5, ws.max_row + 1)}


def test_xlsx_identifica_a_complementacao_estimada():
    proj = _proj(_editor_irregular())
    resumo = [("Premissa de projeção", ROTULO_MEDIA_IRREGULAR),
              ("Complementação orçamentária estimada", 1234.56)]
    nota, linhas = _resumo_xlsx(gerar_xlsx_projecao(pd.DataFrame(), proj, resumo,
                                                    nota_resumo=NOTA_RESUMO_ESTIMADA))
    assert nota == NOTA_RESUMO_ESTIMADA
    assert "não obrigação de pagamento comprovada" in nota
    assert linhas["Premissa de projeção"] == ROTULO_MEDIA_IRREGULAR
    assert linhas["Complementação orçamentária estimada"] == 1234.56
    ws_p = load_workbook(BytesIO(gerar_xlsx_projecao(pd.DataFrame(), proj, resumo)))["PROJECAO"]
    assert ws_p["D2"].value == pytest.approx(2166.67)


def test_xlsx_padrao_preserva_a_nota_confirmada():
    nota, _ = _resumo_xlsx(gerar_xlsx_projecao(pd.DataFrame(), pd.DataFrame(), []))
    assert nota == NOTA_RESUMO_CONFIRMADA


# --------------------------------------------------------- pagina (AppTest)

def test_pagina_automatica_irregular_usa_a_media_e_rotula_estimativa():
    at = _run(_sessao_pc())
    assert not at.exception, at.exception
    blob = _blob(at)
    assert ROTULO_MEDIA_IRREGULAR in blob
    assert ESTIMADA in blob and CONFIRMADA not in blob
    dif = _valor_do_card(blob, "Diferença futura projetada")
    retro = _valor_do_card(blob, "Retroativo reconhecido considerado")
    assert dif == pytest.approx(1126.98)
    assert _valor_do_card(blob, ESTIMADA) == pytest.approx(retro + dif, abs=0.01)
    avisos = " ".join(w.value for w in at.warning)
    assert "não representa obrigação de pagamento comprovada" in avisos
    # retroativo potencial continua separado (nao integra a estimada)
    assert _valor_do_card(blob, "Cenário com potencial") - _valor_do_card(blob, ESTIMADA) \
        == pytest.approx(12000.0, abs=0.01)


def test_pagina_campo_vazio_zera_a_competencia():
    at = _run(_sessao_pc())
    chave = next(k for k in at.session_state.filtered_state
                 if str(k).startswith("adequacao_v3_editor"))
    assert chave.endswith("_media_irregular")
    at.session_state[chave] = {"edited_rows": {0: {"Valor informado pelo fiscal": ""}},
                               "added_rows": [], "deleted_rows": []}
    at.run()
    assert not at.exception, at.exception
    dif = _valor_do_card(_blob(at), "Diferença futura projetada")
    assert dif == pytest.approx(1126.98 - 62.61, abs=0.01)


def test_pagina_mensal_permanece_confirmada():
    at = _run(_sessao_pc())
    at.radio(key="adequacao_v3_premissa").set_value("Mensal (média)")
    at.run()
    assert not at.exception, at.exception
    blob = _blob(at)
    assert CONFIRMADA in blob and ESTIMADA not in blob
    assert ROTULO_MEDIA_IRREGULAR not in blob
    assert _valor_do_card(blob, "Diferença futura projetada") == pytest.approx(1126.98)


@pytest.mark.parametrize("premissa", ["Por ciclo", "Manual"])
def test_pagina_demais_premissas_sem_media_irregular(premissa):
    at = _run(_sessao_pc())
    at.radio(key="adequacao_v3_premissa").set_value(premissa)
    at.run()
    assert not at.exception, at.exception
    blob = _blob(at)
    assert CONFIRMADA in blob and ESTIMADA not in blob
    assert ROTULO_MEDIA_IRREGULAR not in blob


def test_pagina_sem_historico_valido_nao_inventa_valores():
    sessao = _sessao_pc()
    sessao["itens_pc_v10"]["itens"] = []
    at = _run(sessao)
    assert not at.exception, at.exception
    blob = _blob(at)
    assert ROTULO_MEDIA_IRREGULAR not in blob
    assert _valor_do_card(blob, "Diferença futura projetada") == 0
