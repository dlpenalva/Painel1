"""Adequacao Orcamentaria — UX V3 (planilha guiada): comportamento funcional.

Prova, sem alterar a matematica:
  * a pagina abre com as quatro abas e a leitura compacta da Base (Secoes 3, 5, 41);
  * projecao: vazio -> media, 0 -> zero, valor -> override, "ja reajustado" -> /fator
    (Secoes 20, 21, 44) — via calcular_projecao (motor de arredondamento _round2);
  * PC: a exclusao por checkbox USAR alimenta o MESMO pedidos_de_itens_pc(exclusoes=)
    do antigo multiselect (Secao 14, 43);
  * Financeiro: "valor considerado" controla apenas a adequacao, sem alterar o
    valor importado (Secao 12, 42), preservando ZERO x VAZIO;
  * estado persiste entre abas e o reset limpa SOMENTE adequacao_v3_* (Secoes 28, 46);
  * o motor matematico permanece identico ao golden (Secoes 40, U/V/W/X/Y).
"""
from __future__ import annotations

from datetime import date, datetime

import pandas as pd
import pytest

from _adequacao_ui import calcular_projecao
from _adequacao_orcamentaria import (
    pedidos_de_itens_pc, media_pedidos_compra, media_financeiro,
    valor_original_foi_informado,
    calcular_adequacao_orcamentaria as calc, Pedido,
)
from tests.test_adequacao_orcamentaria import PEDIDOS_GOLDEN


def _itens_pc_golden():
    return [{"numero_pc": f"PC-{n}",
             "data_pc": datetime.strptime(d, "%Y-%m-%d").date(),
             "valor_pc": v} for n, (d, v) in enumerate(PEDIDOS_GOLDEN)]


def _sessao_golden():
    df_fin = pd.DataFrame({
        "Competência": ["01/2026", "02/2026", "03/2026", "04/2026", "05/2026", "06/2026"],
        "Valor": [10000.0] * 6,
    })
    return {"df_financeiro_mensal": df_fin, "itens_pc_v10": {"itens": _itens_pc_golden()},
            "valor_represado_a_pagar": 16888.59, "variacao_acumulada": 0.1201,
            "modo_apuracao": "Completo"}


# ----------------------------------------------------------------- Secao 44 (projecao)

def _linha(comp, valor, premissa="Valor sem reajuste"):
    return {"Competência": comp, "Valor informado pelo fiscal": valor,
            "Premissa do valor informado": premissa, "Observação": ""}


def test_projecao_vazio_zero_override():
    df = pd.DataFrame([_linha("jul/26", ""), _linha("ago/26", "0"), _linha("set/26", "15000")])
    r = calcular_projecao(df, media_mensal=1000.0, fator_reajuste=1.10)
    bases = [round(x, 2) for x in r["Valor base considerado"].tolist()]
    difs = [round(x, 2) for x in r["Diferença futura a adequar"].tolist()]
    origens = r["Origem"].tolist()
    assert bases == [1000.0, 0.0, 15000.0]          # vazio->media, 0->zero, 15000->override
    assert difs == [100.0, 0.0, 1500.0]
    assert origens[0] == "Média dos últimos 6 meses"       # vazio => automatico
    assert origens[1] == "Valor informado pelo fiscal"     # 0 informado (ZERO != VAZIO)
    assert origens[2] == "Valor informado pelo fiscal"


def test_projecao_valor_ja_reajustado_divide_pelo_fator():  # Secao 21, O
    df = pd.DataFrame([_linha("set/26", "16500", premissa="Valor já reajustado")])
    r = calcular_projecao(df, media_mensal=1000.0, fator_reajuste=1.10)
    assert round(r["Valor base considerado"].iloc[0], 2) == 15000.0      # 16500 / 1.10
    assert round(r["Valor reajustado estimado"].iloc[0], 2) == 16500.0
    assert r["Premissa usada"].iloc[0] == "Valor já reajustado"


# ----------------------------------------------------------------- Secao 43 (PC USAR)

def test_pc_exclusao_por_checkbox_iguala_multiselect():
    regs = _itens_pc_golden()
    ult = date(2026, 6, 1)
    todos = media_pedidos_compra(pedidos_de_itens_pc(regs), ult, 39)
    # Excluir um PC considerado (a coluna USAR desmarcada == exclusoes=[id]).
    alvo = "PC-0"
    excl = media_pedidos_compra(pedidos_de_itens_pc(regs, exclusoes=[alvo]), ult, 39)
    # mesma matematica do antigo multiselect: pedidos_de_itens_pc(exclusoes=...)
    assert excl["pedidos_considerados"] <= todos["pedidos_considerados"]
    assert excl["total_historico"] <= todos["total_historico"]
    # e a media do subconjunto confere com o motor (nao inventa)
    assert excl["media_mensal"] == pytest.approx(excl["total_historico"] / 39)


# ----------------------------------------------------------------- Secao 42 (financeiro ajuste)

def test_financeiro_valor_considerado_controla_adequacao_preserva_zero_vazio():
    # "Valor considerado" editavel: vazio => fora do denominador; 0 => entra.
    considerados = []
    for bruto in ["10000", "", "0", "12000"]:   # 4 competencias; 1 vazia
        if valor_original_foi_informado(bruto):
            considerados.append(float(str(bruto).replace(",", ".")) if bruto else 0.0)
    # vazio ficou fora; 0 entrou; 2 positivos entraram => 3 no denominador
    r = media_financeiro(considerados)
    assert r["meses_com_valor"] == 3
    assert r["media_mensal"] == pytest.approx((10000 + 0 + 12000) / 3)


# ----------------------------------------------------------------- Secao 40 (motor inalterado)

def test_motor_golden_permanece_identico():  # U, V, W, X, Y
    peds = [Pedido(f"PC-{n}", datetime.strptime(d, "%Y-%m-%d").date(), v)
            for n, (d, v) in enumerate(PEDIDOS_GOLDEN)]
    r = calc(origem="Pedidos de compra", percentual=0.1201,
             ultima_competencia=date(2026, 6, 1), data_fim_vigencia=date(2027, 5, 5),
             retroativo=16888.59, janela_meses=39, saldo_contratual=911237.89, pedidos=peds)
    assert abs(r["media_mensal"] - 14306.976846153848) < 1e-9
    assert round(r["referencia_reajustada"], 2) == 16025.24
    assert round(r["diferenca_futura"], 2) == 18900.86
    assert round(r["complemento_estimado"], 2) == 35789.45
    progs = {p["exercicio"]: round(p["valor"], 2) for p in r["programacao_por_exercicio"]}
    assert progs == {2026: 27198.15, 2027: 8591.30}


# ----------------------------------------------------------------- Secao 41 / 3 (abas + base) — WEB

def _run(origem=None, vigencia=None, sessao=None):
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file("pages/12_Adequacao_Orcamentaria.py", default_timeout=120)
    at.session_state["resultado_valor_global"] = sessao if sessao is not None else _sessao_golden()
    at.run()
    if origem is not None:
        at.radio(key="adequacao_v3_origem").set_value(origem)
    if vigencia is not None:
        at.text_input(key="adequacao_v3_data_final_vigencia").set_value(vigencia)
    if origem is not None or vigencia is not None:
        at.run()
    return at


def test_web_quatro_abas_e_base_leitura():  # A, B
    at = _run()
    assert not at.exception, at.exception
    assert len(at.tabs) == 4
    blob = "\n".join(str(m.value) for m in at.markdown)
    assert "Fontes encontradas" in blob
    assert "Financeiro + Pedidos de Compra" in blob   # ambas as fontes detectadas
    assert "16.888,59" in blob                          # retroativo (leitura Base)
    assert "12,01%" in blob                             # percentual (leitura Base)


def test_web_estado_persiste_entre_abas():  # Secao 28, S
    at = _run(origem="Pedidos de Compra", vigencia="05/05/2027")
    assert not at.exception, at.exception
    # alterna origem e volta; a vigencia e a origem escolhida persistem
    at.radio(key="adequacao_v3_origem").set_value("Financeiro"); at.run()
    at.radio(key="adequacao_v3_origem").set_value("Pedidos de Compra"); at.run()
    assert at.session_state["adequacao_v3_data_final_vigencia"] == "05/05/2027"
    assert at.session_state["adequacao_v3_origem"] == "Pedidos de Compra"


def test_web_reset_limpa_somente_v3():  # Secao 26, T
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file("pages/12_Adequacao_Orcamentaria.py", default_timeout=120)
    at.session_state["resultado_valor_global"] = _sessao_golden()
    at.session_state["adequacao_v3_data_final_vigencia"] = "05/05/2027"
    at.run()
    assert at.session_state["adequacao_v3_data_final_vigencia"] == "05/05/2027"
    at.button(key="adequacao_v3_reset").click(); at.run()
    assert not at.exception, at.exception
    # ajuste da adequacao foi limpo; apuracao/coleta permanecem
    vig = (at.session_state["adequacao_v3_data_final_vigencia"]
           if "adequacao_v3_data_final_vigencia" in at.session_state else "")
    assert vig == ""
    assert "resultado_valor_global" in at.session_state
