"""Fonte unica de verdade do metodo de apuracao (CONTROLE!B1).

CONTROLE!B1 e a UNICA decisao editavel do usuario. RESULTADOS!B4 e derivado
(read-only). Este modulo centraliza os rotulos do dropdown, a normalizacao de
aliases legados e os codigos internos — para nao espalhar condicionais
diferentes por varias abas/modulos.

Codigos internos:
    "principal"  -> Principal (Financeiro)
    "pc"         -> PC (Pedidos de Compra)
    "d"          -> Itens Consumidos
    ""           -> nenhum metodo selecionado (SELECIONE AQUI / pendente)
"""
from __future__ import annotations

import unicodedata
from typing import Any

# Rotulos oficiais do dropdown de CONTROLE!B1 (ordem e a exibida ao usuario).
SELECIONE_AQUI = "SELECIONE AQUI"
ROTULO_PRINCIPAL = "Principal (Financeiro)"
ROTULO_PC = "PC (Pedidos de Compra)"
ROTULO_CONSUMIDOS = "Itens Consumidos"
OPCOES_DROPDOWN = [SELECIONE_AQUI, ROTULO_PRINCIPAL, ROTULO_PC, ROTULO_CONSUMIDOS]

# Mensagem em RESULTADOS!B4 enquanto CONTROLE!B1 = SELECIONE AQUI.
MENSAGEM_PENDENTE = "SELECIONE O METODO EM CONTROLE!B1"

_ROTULO_POR_CODIGO = {
    "principal": ROTULO_PRINCIPAL,
    "pc": ROTULO_PC,
    "d": ROTULO_CONSUMIDOS,
    "": SELECIONE_AQUI,
}

# Aliases legados + novos rotulos -> codigo interno. Chaves ja normalizadas
# (minusculas, sem acento). Nao quebrar Coletas anteriormente geradas.
_ALIASES = {
    # Principal / Financeiro
    "principal": "principal",
    "principal (financeiro)": "principal",
    "financeiro": "principal",
    # PC / Pedidos de Compra
    "pc": "pc",
    "pcs": "pc",
    "pc (pedidos de compra)": "pc",
    "pedido de compra": "pc",
    "pedidos de compra": "pc",
    "pedidos de compras": "pc",
    # Itens Consumidos
    "d": "d",
    "itens": "d",
    "itens consumidos": "d",
    # Pendente / nao selecionado
    "": "",
    "selecione": "",
    "selecione aqui": "",
}


def _norm(txt: Any) -> str:
    if txt is None:
        return ""
    t = str(txt).strip().lower()
    t = unicodedata.normalize("NFKD", t)
    return "".join(ch for ch in t if not unicodedata.combining(ch))


def normalizar_metodo(valor: Any) -> str:
    """Retorna o codigo interno canonico ('principal'|'pc'|'d'|'').

    Aliases desconhecidos caem no proprio texto normalizado (compatibilidade
    com o comportamento legado do leitor), exceto SELECIONE AQUI -> ''.
    """
    n = _norm(valor)
    if n in _ALIASES:
        return _ALIASES[n]
    return n


def metodo_selecionado(valor: Any) -> bool:
    """True quando ha metodo valido escolhido (nao SELECIONE AQUI / vazio)."""
    return normalizar_metodo(valor) in ("principal", "pc", "d")


def rotulo_metodo(valor: Any) -> str:
    """Rotulo user-facing canonico a partir de qualquer alias/codigo."""
    codigo = normalizar_metodo(valor)
    return _ROTULO_POR_CODIGO.get(codigo, SELECIONE_AQUI)


def rotulo_resultados_b4(valor_b1: Any) -> str:
    """RESULTADOS!B4 derivado de CONTROLE!B1 (read-only).

    Enquanto CONTROLE!B1 = SELECIONE AQUI -> mensagem de alerta.
    """
    if not metodo_selecionado(valor_b1):
        return MENSAGEM_PENDENTE
    return rotulo_metodo(valor_b1)
