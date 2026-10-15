"""Geometria da aba ``aditivos`` — Coleta 12.1 x layouts anteriores.

A Coleta 12.1 reorganiza a aba em blocos de leitura (identificacao, base
economica, aplicacao do reajuste, totais, controles tecnicos) e expoe o VALOR
UNITARIO ATUALIZADO que a formula do valor atualizado da alteracao ja usava.
Nenhuma regra muda: so a letra fisica de cada coluna.

Arquivos de versoes anteriores (11.x/12.0, ainda aceitos no upload) mantem o
layout antigo. Todo leitor ou gerador que precisa de uma coluna de aditivos
pergunta a letra aqui, a partir do proprio workbook — nunca a fixa.

    chave             12.1   anterior
    item               A       A
    data               B       B
    ciclo              C       C
    tipo               D       D
    quantidade         E       E
    vu_original        F       F
    base_vu            G       N    ultimo reajuste ja incorporado ao VU
    aplicar            H       H
    fator              I       I
    vu_atualizado      J       -    (novo; exposicao de ROUND(F*I,2) / F)
    valor_original     K       G
    valor_atualizado   L       J
    considerado        M       K
    delta              N       L    DELTA_QTD_CONTRATUAL
    check              O       M    CHECK_POSICAO_CONTRATUAL
    indice_base        P       O    INDICE_BASE_ECONOMICA_VU (oculta)
"""
from __future__ import annotations

import re
import unicodedata
from typing import Any

ABA_ADITIVOS = "aditivos"

COLUNAS_ADITIVOS_12_1: dict[str, str] = {
    "item": "A",
    "data": "B",
    "ciclo": "C",
    "tipo": "D",
    "quantidade": "E",
    "vu_original": "F",
    "base_vu": "G",
    "aplicar": "H",
    "fator": "I",
    "vu_atualizado": "J",
    "valor_original": "K",
    "valor_atualizado": "L",
    "considerado": "M",
    "delta": "N",
    "check": "O",
    "indice_base": "P",
}

COLUNAS_ADITIVOS_LEGADO: dict[str, str | None] = {
    "item": "A",
    "data": "B",
    "ciclo": "C",
    "tipo": "D",
    "quantidade": "E",
    "vu_original": "F",
    "base_vu": "N",
    "aplicar": "H",
    "fator": "I",
    "vu_atualizado": None,
    "valor_original": "G",
    "valor_atualizado": "J",
    "considerado": "K",
    "delta": "L",
    "check": "M",
    "indice_base": "O",
}

# Letra anterior -> letra da 12.1 (traducao fisica usada na migracao do
# template e na comparacao A/B das formulas que leem aditivos).
TRADUCAO_LEGADO_PARA_12_1: dict[str, str] = {
    antiga: COLUNAS_ADITIVOS_12_1[chave]
    for chave, antiga in COLUNAS_ADITIVOS_LEGADO.items()
    if antiga is not None
}

CABECALHOS_ADITIVOS_12_1: dict[str, str] = {
    "A": "ITEM",
    "B": "Data do aditivo",
    "C": "Ciclo/Marco",
    "D": "Tipo de alteração formalizada",
    "E": "Quantidade acrescida/suprimida",
    "F": "Valor unitário original",
    "G": "Último reajuste já incorporado ao VU",
    "H": "Aplicar reajuste? (Sim/Nao)",
    "I": "Fator acumulado aplicável",
    "J": "Valor unitário atualizado",
    "K": "Valor original da alteração",
    "L": "Valor atualizado da alteração",
    "M": "Considerado no cálculo financeiro? (Sim/Nao)",
    "N": "DELTA_QTD_CONTRATUAL",
    "O": "CHECK_POSICAO_CONTRATUAL",
    "P": "INDICE_BASE_ECONOMICA_VU",
}

_MARCADOR_12_1 = "valor unitario atualizado"


def _normalizar(valor: Any) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or "").strip().lower())
    return " ".join("".join(c for c in texto if not unicodedata.combining(c)).split())


def eh_layout_aditivos_12_1(ws) -> bool:
    """True quando a aba traz o VU atualizado em J (Coleta 12.1)."""
    return _normalizar(ws.cell(row=1, column=10).value) == _MARCADOR_12_1


def colunas_aditivos(ws) -> dict[str, str | None]:
    """Letra de cada coluna semantica na aba ``aditivos`` recebida."""
    if eh_layout_aditivos_12_1(ws):
        return dict(COLUNAS_ADITIVOS_12_1)
    return dict(COLUNAS_ADITIVOS_LEGADO)


def colunas_aditivos_do_workbook(wb) -> dict[str, str | None]:
    """Igual a ``colunas_aditivos``; sem a aba, assume o layout da 12.1."""
    if ABA_ADITIVOS not in wb.sheetnames:
        return dict(COLUNAS_ADITIVOS_12_1)
    return colunas_aditivos(wb[ABA_ADITIVOS])


def indice_coluna(letra: str) -> int:
    """Indice 1-based da letra de coluna (A=1)."""
    from openpyxl.utils import column_index_from_string

    return column_index_from_string(letra)


_REF = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d*)")


def traduzir_formula_legado_para_12_1(formula: str, *, aba_da_formula: str) -> str:
    """Reescreve as referencias a ``aditivos`` de uma formula do layout anterior.

    So operandos do tipo RANGE sao tocados: os da propria aba ``aditivos``
    (sem prefixo, quando a formula mora nela) e os prefixados ``aditivos!``.
    Textos, funcoes e referencias a outras abas ficam intactos. Usado para
    PROVAR (A/B) que a migracao fisica do template so acompanhou a geometria.
    """
    from openpyxl.formula.tokenizer import Token, Tokenizer

    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    tok = Tokenizer(formula)
    saida = ["="]
    for item in tok.items:
        valor = item.value
        if item.type == Token.OPERAND and item.subtype == Token.RANGE:
            valor = _traduzir_operando(valor, aba_da_formula)
        saida.append(valor)
    return "".join(saida)


def _traduzir_operando(operando: str, aba_da_formula: str) -> str:
    if "!" in operando:
        aba, ref = operando.rsplit("!", 1)
        if aba.strip("'") != ABA_ADITIVOS:
            return operando
        return f"{aba}!{_traduzir_ref(ref)}"
    if aba_da_formula != ABA_ADITIVOS:
        return operando
    return _traduzir_ref(operando)


def _traduzir_ref(ref: str) -> str:
    def troca(m: re.Match) -> str:
        dolar_col, col, dolar_lin, lin = m.groups()
        return f"{dolar_col}{TRADUCAO_LEGADO_PARA_12_1.get(col, col)}{dolar_lin}{lin}"

    return ":".join(_REF.sub(troca, parte, count=1) for parte in ref.split(":"))
