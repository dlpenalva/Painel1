"""Adequacao Orcamentaria — historico irregular e premissas explicitas.

Os testes de motor do legado Cl8us 12.18 preservam o comportamento das
funcoes puras de media e override. A pagina Cl8us 12.21 NAO os usa
automaticamente para historico irregular:
  * Automatica sem cadencia confiavel: futuro nao determinado;
  * Mensal (media) selecionada expressamente: conserva o cenario de media;
  * Manual e Por ciclo seguem separados;
  * retroativo reconhecido e potencial nunca se confundem.
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
    NAO_DETERMINADA,
    ROTULO_MEDIA_IRREGULAR,
    calcular_projecao,
    consolidar_resultado_adequacao,
    explicar_resultado_adequacao,
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
# 12.20: o rotulo "COMPLEMENTACAO CONFIRMADA" deixou de existir (todas as
# premissas apresentam a complementacao ESTIMADA ou PARCIAL).
CONFIRMADA = "COMPLEMENTAÇÃO CONFIRMADA"
PARCIAL = "COMPLEMENTAÇÃO ORÇAMENTÁRIA PARCIAL"


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


# ------------------------------------------- narrativa com valores canonicos

def test_narrativa_mensal_explicita_soma_so_diferenca_futura_ao_retroativo():
    media = 892877.92
    periodos = list(pd.period_range("2026-11", periods=6, freq="M"))
    proj = calcular_projecao(
        montar_base_editor(periodos, media), media, 1.0381,
        saldo_elegivel=39364302.31,
    )
    cons = consolidar_resultado_adequacao(
        retroativo=36397.14, df_projecao=proj, meses_restantes=6,
        saldo_info={"valor": 39364302.31}, retroativo_potencial=119.20,
    )
    assert cons["diferenca_futura"] == pytest.approx(204111.90)
    assert cons["complementacao"] == pytest.approx(240509.04)
    narrativa = explicar_resultado_adequacao(
        cons, "Mensal (média)", media, 926896.57, 0.0381, origem_pc=True)
    for esperado in ("R$ 36.397,14", "6 meses", "R$ 892.877,92",
                     "3,81%", "R$ 926.896,57", "R$ 34.018,65",
                     "R$ 204.111,90", "R$ 240.509,04"):
        assert esperado in narrativa
    assert "somente a diferença atribuída ao reajuste" in narrativa
    assert "R$ 119,20" not in narrativa  # potencial nao entra na soma


def test_narrativa_automatica_irregular_nao_inventa_diferenca_futura():
    cons = consolidar_resultado_adequacao(
        retroativo=36397.14, df_projecao=pd.DataFrame(), meses_restantes=6,
        saldo_info={"valor": 39364302.31},
        motivo_indeterminado="histórico sem cadência confiável",
    )
    narrativa = explicar_resultado_adequacao(
        cons, "Automática (cadência histórica)", 892877.92, 926896.57,
        0.0381, origem_pc=True)
    assert "R$ 36.397,14" in narrativa
    assert "R$ 892.877,92" in narrativa
    assert "total parcial" in narrativa
    assert "não determinada" in narrativa
    assert "Mensal (média)" in narrativa
    assert "R$ 204.111,90" not in narrativa


# --------------------------------------------------------- pagina (AppTest)

def test_pagina_automatica_irregular_nao_presume_mensalidade():
    at = _run(_sessao_pc())
    assert not at.exception, at.exception
    blob = _blob(at)
    assert "Automática — histórico irregular; futuro não determinado" in blob
    assert ROTULO_MEDIA_IRREGULAR not in blob
    assert PARCIAL in blob and ESTIMADA not in blob
    idx = blob.index("Diferença futura projetada</div>")
    assert NAO_DETERMINADA in blob[idx:idx + 350]
    retro = _valor_do_card(blob, "Retroativo reconhecido considerado")
    assert _valor_do_card(blob, PARCIAL) == pytest.approx(retro, abs=0.01)
    alertas = " ".join(e.value for e in at.error)
    assert "HISTÓRICO SEM PERIODICIDADE SUFICIENTE" in alertas
    assert "não presumirá uma mensalidade" in alertas
    # A fixture nao possui calendario de ciclos: nesse caso, o texto
    # explica a insuficiencia sem inventar contagens por ciclo.
    assert "não mostrou frequência suficiente" in alertas
    assert "A projeção automática foi limitada ao saldo" not in alertas
    assert _valor_do_card(blob, "Cenário com potencial") - _valor_do_card(blob, PARCIAL) \
        == pytest.approx(12000.0, abs=0.01)


def test_pagina_automatica_irregular_ignora_override_de_sessao_antiga():
    at = _run(_sessao_pc())
    chave = next(k for k in at.session_state.filtered_state
                 if str(k).startswith("adequacao_v3_editor"))
    assert not chave.endswith("_media_irregular")
    at.session_state[chave] = {"edited_rows": {0: {"Valor informado pelo fiscal": "5.000,00"}},
                               "added_rows": [], "deleted_rows": []}
    at.run()
    assert not at.exception, at.exception
    blob = _blob(at)
    assert NAO_DETERMINADA in blob[blob.index("Diferença futura projetada</div>") :
                                  blob.index("Diferença futura projetada</div>") + 350]
    assert _valor_do_card(blob, PARCIAL) == pytest.approx(
        _valor_do_card(blob, "Retroativo reconhecido considerado"))


def test_pagina_mensal_nao_usa_o_rotulo_de_media_irregular():
    at = _run(_sessao_pc())
    at.radio(key="adequacao_v3_premissa").set_value("Mensal (média)")
    at.run()
    assert not at.exception, at.exception
    blob = _blob(at)
    assert ESTIMADA in blob and CONFIRMADA not in blob
    assert ROTULO_MEDIA_IRREGULAR not in blob
    assert _valor_do_card(blob, "Diferença futura projetada") == pytest.approx(1126.98)
    assert "Entenda este resultado" in blob
    assert "Você escolheu projetar o consumo pela média histórica" in blob
    assert "somente a diferença atribuída ao reajuste" in blob
    assert not any("HISTÓRICO SEM PERIODICIDADE SUFICIENTE" in e.value
                   for e in at.error)


@pytest.mark.parametrize("premissa", ["Por ciclo", "Manual"])
def test_pagina_demais_premissas_sem_media_irregular(premissa):
    at = _run(_sessao_pc())
    at.radio(key="adequacao_v3_premissa").set_value(premissa)
    at.run()
    assert not at.exception, at.exception
    blob = _blob(at)
    assert CONFIRMADA not in blob
    # Manual sem valor informado: futuro nao determinado (total parcial).
    assert (PARCIAL if premissa == "Manual" else ESTIMADA) in blob
    assert ROTULO_MEDIA_IRREGULAR not in blob


def test_pagina_sem_historico_valido_nao_inventa_valores():
    sessao = _sessao_pc()
    sessao["itens_pc_v10"]["itens"] = []
    at = _run(sessao)
    assert not at.exception, at.exception
    blob = _blob(at)
    assert ROTULO_MEDIA_IRREGULAR not in blob
    # 12.20: ausencia de historico nao vira diferenca futura de R$ 0,00.
    idx = blob.index("Diferença futura projetada</div>")
    assert "Não determinada" in blob[idx:idx + 300]
    assert PARCIAL in blob
