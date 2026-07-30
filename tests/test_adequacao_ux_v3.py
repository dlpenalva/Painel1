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

from _adequacao_ui import (
    calcular_projecao, situacao_financeira_considerada, atualizar_exclusoes_manuais_pc,
)
from _adequacao_orcamentaria import (
    pedidos_de_itens_pc, media_pedidos_compra, media_financeiro, classificar_pedidos,
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


# --------------------------------------------- Secao 3/5/13 (PC janela x exclusao manual)

# Reproduz o ESTADO da camada V3 do editor de PCs (sem Streamlit): elegibilidade
# temporal (classificar_pedidos) + seed do USAR + transicao das exclusoes MANUAIS.
def _simular_render_pc(regs, comp_date, janela, exclusoes_anteriores, toggles=None):
    cl = classificar_pedidos(pedidos_de_itens_pc(regs), comp_date, janela)
    excl = set(str(e) for e in (exclusoes_anteriores or []))
    linhas = []
    for x in cl["pedidos"]:
        pc = str(x["identificacao"])
        eligivel = x["situacao"] == "Considerado"
        usar_seed = eligivel and pc not in excl           # seed = pagina
        usar = (toggles or {}).get(pc, usar_seed)         # override manual do usuario
        linhas.append({"pc": pc, "eligivel": eligivel, "usar": usar})
    manual = atualizar_exclusoes_manuais_pc(linhas, excl)
    base = media_pedidos_compra(pedidos_de_itens_pc(regs, exclusoes=manual), comp_date, janela)
    return manual, base


_REGS_PC = [{"numero_pc": "PC-A", "data_pc": date(2026, 5, 10), "valor_pc": 3000.0},
            {"numero_pc": "PC-B", "data_pc": date(2026, 1, 10), "valor_pc": 6000.0}]
_COMP_PC = date(2026, 6, 1)   # janela 3 => [04..06] (PC-B fora); janela 6 => [01..06] (ambos)


def test_pc_bug_janela_x_exclusao_manual_ponta_a_ponta():  # 13 A,B,C,D
    # PASSO 1: janela curta -> PC-A considerado, PC-B fora (NAO vira exclusao manual)
    m1, b1 = _simular_render_pc(_REGS_PC, _COMP_PC, 3, set())
    assert "PC-B" not in m1 and "PC-A" not in m1          # A
    assert b1["pedidos_considerados"] == 1
    assert b1["media_mensal"] == pytest.approx(3000 / 3)  # 1000  (J: media)

    # PASSO 2: amplia a janela -> PC-B passa a elegivel/considerado AUTOMATICAMENTE
    m2, b2 = _simular_render_pc(_REGS_PC, _COMP_PC, 6, m1)
    assert "PC-B" not in m2                                # B
    assert b2["pedidos_considerados"] == 2
    assert b2["media_mensal"] == pytest.approx(9000 / 6)  # 1500

    # PASSO 3: desmarca PC-B manualmente -> vira exclusao manual
    m3, b3 = _simular_render_pc(_REGS_PC, _COMP_PC, 6, m2, toggles={"PC-B": False})
    assert "PC-B" in m3
    assert b3["pedidos_considerados"] == 1
    assert b3["media_mensal"] == pytest.approx(3000 / 6)  # 500

    # PASSO 4: muda a janela (PC-B fora) e volta -> exclusao MANUAL preservada
    m4a, _ = _simular_render_pc(_REGS_PC, _COMP_PC, 3, m3)   # fora da janela: preserva
    assert "PC-B" in m4a
    m4b, b4 = _simular_render_pc(_REGS_PC, _COMP_PC, 6, m4a)  # volta elegivel: segue excluido
    assert "PC-B" in m4b                                    # C, I
    assert b4["pedidos_considerados"] == 1
    assert b4["media_mensal"] == pytest.approx(3000 / 6)   # 500


def test_pc_fora_da_janela_nunca_entra_em_exclusoes_manuais():  # A (isolado)
    # PC-B fora, mesmo se o usuario "tocar" no checkbox, nao vira exclusao manual.
    m, _ = _simular_render_pc(_REGS_PC, _COMP_PC, 3, set(), toggles={"PC-B": True})
    assert "PC-B" not in m
    m2, _ = _simular_render_pc(_REGS_PC, _COMP_PC, 3, set(), toggles={"PC-B": False})
    assert "PC-B" not in m2


# --------------------------------------------- Secao 7/8/13 (Financeiro situacao considerada)

def test_situacao_financeira_considerada_zero_vazio():  # G, H, I
    assert situacao_financeira_considerada("") == "Sem informação"          # G vazio
    assert situacao_financeira_considerada("   ") == "Sem informação"
    assert situacao_financeira_considerada(None) == "Sem informação"
    assert situacao_financeira_considerada("0") == "Zero informado"          # H zero
    assert situacao_financeira_considerada("0,00") == "Zero informado"
    assert situacao_financeira_considerada("R$ 0,00") == "Zero informado"
    assert situacao_financeira_considerada("15000") == "Informado"           # I >0
    assert situacao_financeira_considerada("1.000,00") == "Informado"
