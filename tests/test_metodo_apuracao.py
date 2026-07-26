"""Metodo de apuracao — fonte unica CONTROLE!B1, RESULTADOS!B4 derivado.

Cobre Secao 5 / Secao 13 (MÉTODO): normalizador canonico compartilhado,
derivacao read-only, bloqueio em SELECIONE AQUI e aliases legados.
"""
from __future__ import annotations

import _metodo_apuracao as M


# ---- CONTROLE!B1 como unica fonte (normalizacao canonica) ------------------
def test_dropdown_labels_para_codigo():
    assert M.normalizar_metodo(M.ROTULO_PRINCIPAL) == "principal"
    assert M.normalizar_metodo(M.ROTULO_PC) == "pc"
    assert M.normalizar_metodo(M.ROTULO_CONSUMIDOS) == "d"
    assert M.normalizar_metodo(M.SELECIONE_AQUI) == ""


def test_opcoes_dropdown_ordem_oficial():
    assert M.OPCOES_DROPDOWN == [
        "SELECIONE AQUI",
        "Principal (Financeiro)",
        "PC (Pedidos de Compra)",
        "Itens Consumidos",
    ]


# ---- RESULTADOS!B4 derivado (read-only) ------------------------------------
def test_b4_derivado_de_b1():
    assert M.rotulo_resultados_b4(M.ROTULO_PC) == M.ROTULO_PC
    assert M.rotulo_resultados_b4("Pedidos de Compras") == M.ROTULO_PC
    assert M.rotulo_resultados_b4("Financeiro") == M.ROTULO_PRINCIPAL


# ---- SELECIONE AQUI bloqueia o resultado dependente ------------------------
def test_selecione_aqui_bloqueia():
    assert M.metodo_selecionado(M.SELECIONE_AQUI) is False
    assert M.metodo_selecionado("") is False
    assert M.metodo_selecionado(None) is False
    assert M.rotulo_resultados_b4(M.SELECIONE_AQUI) == M.MENSAGEM_PENDENTE
    # metodo valido nao bloqueia
    assert M.metodo_selecionado(M.ROTULO_PC) is True


# ---- Aliases legados -------------------------------------------------------
def test_aliases_legados():
    principais = ["Principal", "Financeiro", "principal (financeiro)"]
    pcs = ["Pedidos de Compras", "Pedidos de Compra", "PC", "PCs",
           "PC (Pedidos de Compra)"]
    itens = ["Itens", "Itens Consumidos"]
    for v in principais:
        assert M.normalizar_metodo(v) == "principal", v
    for v in pcs:
        assert M.normalizar_metodo(v) == "pc", v
    for v in itens:
        assert M.normalizar_metodo(v) == "d", v


def test_leitor_delega_ao_normalizador_canonico():
    """O leitor deve reconhecer os mesmos aliases (fonte unica)."""
    import _leitor_masterfile_v10 as L
    assert L._normalizar_modo("Pedidos de Compras") == "pc"
    assert L._normalizar_modo("Principal (Financeiro)") == "principal"
    assert L._normalizar_modo("SELECIONE AQUI") == ""
