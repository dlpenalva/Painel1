"""Cobertura temporal — bug P0 da evidencia Financeiro (fonte Excel) + rotulos.

O motor Python ja e correto (ver test_cobertura_temporal.py, casos O/P/Q). Aqui
guardamos a correcao na FONTE EXCEL (a formula escrita no template) e a nova
nomenclatura, alem de reafirmar o cenario "nenhuma evidencia".
"""
from __future__ import annotations

from datetime import date
from pathlib import Path

from _motor_cobertura_temporal import montar_cobertura_temporal

_TOOL = (Path(__file__).resolve().parents[1]
         / "tools" / "aplicar_cobertura_temporal_template.py").read_text(
    encoding="utf-8")


# ---- Fonte Excel corrigida (nao basta o motor) -----------------------------
def test_formula_financeiro_nao_usa_max_da_coluna_a():
    # A grade (coluna A) vai ate o termino contratual; MAX(A) daria falsa
    # evidencia. A formula deve olhar o VALOR (coluna C) informado.
    assert "MAX(financeiro!$A$2:$A$200)" not in _TOOL
    # Robusto: evidencia exige, NA MESMA LINHA, A (competencia) E C (valor)
    # numericos — exclui a linha TOTAL C74=SUM(C2:C73) cuja A e vazia. As duas
    # metades ISNUMBER podem estar em linhas-fonte distintas (string concatenada).
    assert "LOOKUP(2,1/(ISNUMBER(financeiro!$A$2:$A$200)" in _TOOL
    assert "ISNUMBER(financeiro!$C$2:$C$200)" in _TOOL


def test_projecao_temporal_automatica_independe_da_analise():
    # Inicio da projecao = dia seguinte ao maior marco seguro, sem depender de B4.
    assert 'COB_ADOTADA = "MAX($B$11,$B$13,$B$15)"' in _TOOL
    assert '=IF({COB_ADOTADA}=0,"",{COB_ADOTADA}+1)' in _TOOL
    # nao pode mais depender da data de analise (B4) para autorizar a projecao
    assert 'IF($B$4>{COB_ADOTADA}' not in _TOOL


# ---- Nomenclatura: metodo de apuracao x fonte temporal ---------------------
def test_rotulos_renomeados():
    assert "Fonte temporal principal" in _TOOL
    assert "Fonte temporal de conferencia" in _TOOL
    assert "Inicio da projecao temporal" in _TOOL
    assert "Metodo de apuracao (CONTROLE!B1)" in _TOOL
    # rotulos antigos removidos
    assert '"Fonte principal"' not in _TOOL
    assert '"Fontes de conferencia"' not in _TOOL
    assert '"Projecao autorizada a partir de"' not in _TOOL


def test_metodo_apuracao_derivado_de_controle_b1():
    assert "CONTROLE!$B$1" in _TOOL
    assert "SELECIONE O METODO EM CONTROLE!B1" in _TOOL


# ---- Motor: nenhuma evidencia financeira (cenario 3) -----------------------
def _res(financeiro):
    return {
        "ok": True,
        "controle": {"ciclo_vigente": "C4"},
        "financeiro": financeiro,
        "posicao_contratual": {"itens": []},
    }


def test_nenhuma_evidencia_financeira():
    # grade futura pre-montada, todos os valores vazios => nenhuma evidencia
    fin = [{"competencia": date(2026, m, 1), "valor": None} for m in range(1, 13)]
    r = montar_cobertura_temporal(_res(fin))
    assert r.financeiro_ultima_evidencia is None
