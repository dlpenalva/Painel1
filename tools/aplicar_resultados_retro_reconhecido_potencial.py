# -*- coding: utf-8 -*-
"""RESULTADOS — card RETRO. RECONHECIDO + POTENCIAL e SITUACAO DA APURACAO.

Somente APRESENTACAO da aba RESULTADOS executiva (Coleta 12.1). Nenhum calculo
financeiro muda: VTA, retroativo reconhecido, retroativo potencial, saldo e a
regra de pendencias continuam exatamente onde estavam.

O QUE MUDA
  RESULTADOS          cards (linhas 8-10) na ordem VTA OFICIAL | RETROATIVO
                      RECONHECIDO | RETROATIVO POTENCIAL | RETRO. RECONHECIDO +
                      POTENCIAL (novo, lilas) | SALDO REMANESCENTE | INDICE
                      ACUMULADO. A SITUACAO DA APURACAO sai do card estreito G e
                      vira uma faixa de largura total (linhas 12-13, tres linhas
                      inseridas pelo proprio Excel antes da COMPOSICAO DO VTA,
                      que reajusta referencias, formatacao condicional, mesclas e
                      area de impressao). Tres estados: concluida (verde),
                      concluida com potencial (amarelo-esverdeado) e REVISE
                      (rosa) — a regra de pendencia continua sendo
                      RESULTADOS_DETALHE!A7 ($J$5 = 0).
  MEMORIA_RESULTADOS  AF65:AG70 (textos acentuados) e S66:T67 (camada tecnica
                      da apresentacao): T66 = soma-resumo; T67 = texto da faixa.

A SOMA E SO RESUMO
  T66 = RESULTADOS_DETALHE!D22 (a MESMA fonte do card RETROATIVO RECONHECIDO)
        + RETROATIVO_POTENCIAL_VTA (a MESMA fonte do card RETROATIVO POTENCIAL),
  com as mesmas condicoes de disponibilidade dos dois cards: reconhecido
  indisponivel -> vazio; metodo PCs com potencial indisponivel -> vazio;
  metodo sem parcela potencial -> o proprio reconhecido. Sem nome definido:
  nao e lida pelo VTA, por documentos nem pelo aplicativo.

  A aba executiva continua sem motor de calculo (so espelha): a soma vive na
  camada tecnica, como as demais grandezas.

REGRA ZERO CORRUPCAO XLSX
  Somente Excel COM (openpyxl destroi a formatacao condicional x14); formulas
  em ingles, ASCII e parenteses balanceados (verificado antes da escrita);
  formatacao condicional so le a propria aba (regra que le outra aba vai para a
  extensao x14, que o openpyxl descarta na geracao da Coleta). Qualquer celula
  fora da forma 12.0 esperada aborta sem gravar.

uso: python tools/aplicar_resultados_retro_reconhecido_potencial.py [caminho.xlsx]
     (exige pywin32 + Excel real; sem argumento, altera o template oficial)
"""
from __future__ import annotations

import gc
import re
import sys
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
for _caminho in (RAIZ, RAIZ / "tools"):
    if str(_caminho) not in sys.path:
        sys.path.insert(0, str(_caminho))

import aplicar_resultados_executivo_v3 as v3  # noqa: E402

TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"
EXE, DET, MEM = v3.EXE, v3.DET, v3.MEM

# ------------------------------------------------------------- geometria
# Tres linhas inseridas antes da COMPOSICAO DO VTA (linha 12 da 12.0).
LINHA_INSERCAO, LINHAS_INSERIDAS = 12, 3
L_SIT_ROT, L_SIT_TXT, L_SIT_ESPACO = 12, 13, 14
ALTURAS_FAIXA = {L_SIT_ROT: 22, L_SIT_TXT: 34, L_SIT_ESPACO: 12}
ULTIMA_LINHA = v3.ULTIMA_LINHA + LINHAS_INSERIDAS
L_CARD_ROT, L_CARD_VAL, L_CARD_SUB = v3.L_CARD_ROT, v3.L_CARD_VAL, v3.L_CARD_SUB

# Cards na ordem pedida (coluna: rotulo, fundo, texto, friso superior).
CARDS = {
    "B": ("VTA OFICIAL", *v3.CARDS["B"]),
    "C": ("RETROATIVO RECONHECIDO", *v3.CARDS["C"]),
    "D": ("RETROATIVO POTENCIAL", *v3.CARDS["D"]),
    # Lilas muito claro: sintese das duas parcelas anteriores.
    "E": ("RETRO. RECONHECIDO + POTENCIAL", "EDE7F6", "5B3F8F", "A994D6"),
    "F": ("SALDO REMANESCENTE", *v3.CARDS["E"]),
    "G": ("ÍNDICE ACUMULADO", *v3.CARDS["F"]),
}
# Rotulo mais longo da linha: a 9pt precisa de 26,73 de largura (coluna 26);
# a 8,5pt cabe sem mudar larguras nem o texto pedido.
TAMANHO_ROTULO = {"E": 8.5}
FAIXA_FUNDO, FAIXA_TEXTO, FAIXA_FRISO = "F7F7F7", "262626", "BFBFBF"
# Situacao: concluida (verde), concluida com potencial (amarelo-esverdeado,
# distinto do verde e do ambar do card potencial) e revise (rosa).
SIT_CORES = {
    "verde": v3.STATUS_CORES["verde"],
    "potencial": ("EEF3CF", "4F6228"),
    "rosa": v3.STATUS_CORES["rosa"],
}

# ------------------------------------------------------------- textos (AF:AG)
PRIMEIRA_LINHA_TEXTO = 65
TEXTOS = [
    ("SOMA_SUB", "Soma do retroativo reconhecido com a parcela potencial."),
    ("SOMA_SUB_NA", "Neste método, não há parcela potencial; o total coincide "
                    "com o retroativo reconhecido."),
    ("SIT_REVISE", "REVISE A APURAÇÃO"),
    ("SIT_REVISE_TXT", "Existem pendências que precisam ser corrigidas antes "
                       "da conclusão."),
    ("SIT_POT_1", "Nenhuma pendência impeditiva identificada. Há "),
    ("SIT_POT_2", " de retroativo potencial, sujeito à confirmação."),
]
LINHA_TEXTO = {chave: PRIMEIRA_LINHA_TEXTO + i for i, (chave, _) in enumerate(TEXTOS)}


def T(chave: str, *, local: bool = False) -> str:
    """Texto constante novo (local=True: a partir da propria MEMORIA)."""
    ref = f"${v3.COL_TEXTO}${LINHA_TEXTO[chave]}"
    return ref if local else f"{MEM}!{ref}"


# ------------------------------------------------------------- camada tecnica
RECONHECIDO = f"{DET}!$D$22"                   # mesma fonte do card C9
POTENCIAL = "RETROATIVO_POTENCIAL_VTA"         # mesma fonte do card D9
PCS = v3.PCS
CELULA_SOMA, CELULA_SITUACAO = "T66", "T67"

FORMULAS_MEMORIA = {
    "S66": ("Retroativo reconhecido + potencial (soma-resumo da RESULTADOS; "
            "NAO entra no VTA, em documentos nem em outro calculo)"),
    CELULA_SOMA: (
        f'=IF(NOT(ISNUMBER({RECONHECIDO})),"",IF({PCS},'
        f'IF(ISNUMBER({POTENCIAL}),ROUND({RECONHECIDO}+{POTENCIAL},2),""),'
        f"{RECONHECIDO}))"
    ),
    "S67": ("Texto da SITUACAO DA APURACAO (RESULTADOS): concluida / concluida "
            "com potencial / revise (regra de pendencia: RESULTADOS_DETALHE!A7)"),
    CELULA_SITUACAO: (
        f'=IF(LEFT({DET}!$A$7,4)<>"APUR",{T("SIT_REVISE_TXT", local=True)}'
        f'&" "&{DET}!$B$7,IF(NOT({PCS}),{DET}!$B$7,'
        f"IF(NOT(ISNUMBER({POTENCIAL})),{DET}!$B$7,IF({POTENCIAL}>0,"
        f'{T("SIT_POT_1", local=True)}&"R$ "&FIXED({POTENCIAL},2)'
        f'&{T("SIT_POT_2", local=True)},{DET}!$B$7))))'
    ),
}


def _abs(celula: str) -> str:
    return f"{MEM}!${celula[0]}${celula[1:]}"


# ------------------------------------------------------------- aba executiva
def _deslocar(endereco: str) -> str:
    col = "".join(ch for ch in endereco if ch.isalpha())
    lin = int("".join(ch for ch in endereco if ch.isdigit()))
    if lin >= LINHA_INSERCAO:
        lin += LINHAS_INSERIDAS
    return f"{col}{lin}"


def formulas_executivo() -> dict[str, str]:
    """Todas as formulas da aba RESULTADOS executiva na Coleta 12.1.

    Parte do descritor homologado da v3 (`aplicar_resultados_executivo_v3`),
    desloca as linhas a partir da 12 (as tres inseridas) e reorganiza os
    cards. As unicas referencias internas da v3 (F9 = F6; F10 le F9) ficam
    acima da insercao e seguem o card do indice para a coluna G.
    """
    original = v3.formulas_executivo()
    base = dict(original)
    for col in "EFG":
        for linha in (L_CARD_VAL, L_CARD_SUB):
            del base[f"{col}{linha}"]
    f = {_deslocar(endereco): formula for endereco, formula in base.items()}
    traco = v3.T("TRACO")
    soma = _abs(CELULA_SOMA)
    f[f"E{L_CARD_VAL}"] = f"=IF(ISNUMBER({soma}),{soma},{traco})"
    f[f"E{L_CARD_SUB}"] = f'=IF({PCS},{T("SOMA_SUB")},{T("SOMA_SUB_NA")})'
    f[f"F{L_CARD_VAL}"] = original[f"E{L_CARD_VAL}"]
    f[f"F{L_CARD_SUB}"] = original[f"E{L_CARD_SUB}"]
    f[f"G{L_CARD_VAL}"] = original[f"F{L_CARD_VAL}"]               # =F6
    f[f"G{L_CARD_SUB}"] = original[f"F{L_CARD_SUB}"].replace(
        f"ISNUMBER(F{L_CARD_VAL})", f"ISNUMBER(G{L_CARD_VAL})")
    f[f"C{L_SIT_ROT}"] = (
        f'=IF(LEFT({DET}!$A$7,4)="APUR",{DET}!$A$7,{T("SIT_REVISE")})'
    )
    f[f"C{L_SIT_TXT}"] = f"={_abs(CELULA_SITUACAO)}"
    return f


def constantes_executivo() -> dict[str, str]:
    """Textos constantes que esta frente grava na aba executiva."""
    c = {f"{col}{L_CARD_ROT}": rotulo for col, (rotulo, *_r) in CARDS.items()}
    c[f"B{L_SIT_ROT}"] = "SITUAÇÃO DA APURAÇÃO"
    return c


# Regras condicionais da faixa (so a propria aba: nada de x14).
_STATUS = f"$C${L_SIT_ROT}"
_POT_CARD = f"$D${L_CARD_VAL}"
REGRAS_FAIXA = (
    (f'=LEFT({_STATUS},4)="REVI"', "rosa"),
    (f'=AND(LEFT({_STATUS},4)="APUR",N({_POT_CARD})>0)', "potencial"),
    (f'=AND(LEFT({_STATUS},4)="APUR",N({_POT_CARD})<=0)', "verde"),
)
FAIXA_SITUACAO = f"B{L_SIT_ROT}:G{L_SIT_TXT}"


# ------------------------------------------------------------- conferencias
# Forma 12.0 exigida antes de qualquer escrita (fail-closed).
ESPERADO_EXE = {
    "B8": "VTA OFICIAL", "C8": "RETROATIVO RECONHECIDO",
    "D8": "RETROATIVO POTENCIAL", "E8": "SALDO REMANESCENTE",
    "F8": "ÍNDICE ACUMULADO", "G8": "SITUAÇÃO DA APURAÇÃO",
    "B12": "COMPOSIÇÃO DO VTA", "B2": v3.TITULO_EXECUTIVO,
}


def validar_ascii_e_parenteses(formulas: dict[str, str]) -> None:
    v3.validar_ascii_e_parenteses(formulas)


def conferir(wb) -> None:
    exe = wb.Worksheets(EXE)
    for endereco, texto in ESPERADO_EXE.items():
        if exe.Range(endereco).Value != texto:
            raise RuntimeError(f"{EXE}!{endereco} fora da forma 12.0; nada aplicado")
    for endereco, formula in v3.formulas_executivo().items():
        if exe.Range(endereco).Formula != formula:
            raise RuntimeError(f"{EXE}!{endereco} fora da forma 12.0; nada aplicado")
    mem = wb.Worksheets(MEM)
    for linha in range(PRIMEIRA_LINHA_TEXTO, PRIMEIRA_LINHA_TEXTO + len(TEXTOS)):
        for col in (v3.COL_CHAVE, v3.COL_TEXTO):
            if mem.Range(f"{col}{linha}").Formula != "":
                raise RuntimeError(f"{MEM}!{col}{linha} nao esta livre; nada aplicado")
    for endereco in FORMULAS_MEMORIA:
        if mem.Range(endereco).Formula != "":
            raise RuntimeError(f"{MEM}!{endereco} nao esta livre; nada aplicado")
    if exe.ProtectContents or mem.ProtectContents:
        raise RuntimeError("Aba protegida inesperada; nada aplicado")


# ------------------------------------------------------------- etapas COM
def frente_memoria(wb) -> None:
    mem = wb.Worksheets(MEM)
    for chave, texto in TEXTOS:
        linha = LINHA_TEXTO[chave]
        mem.Range(f"{v3.COL_CHAVE}{linha}").Value = chave
        celula = mem.Range(f"{v3.COL_TEXTO}{linha}")
        celula.NumberFormat = "@"
        celula.Value = texto
    for endereco, conteudo in FORMULAS_MEMORIA.items():
        if conteudo.startswith("="):
            validar_ascii_e_parenteses({endereco: conteudo})
            mem.Range(endereco).Formula = conteudo
        else:
            mem.Range(endereco).Value = conteudo


def _estilo_card(ws, col: str, fundo: str, texto: str, friso: str) -> None:
    faixa = ws.Range(f"{col}{L_CARD_ROT}:{col}{L_CARD_SUB}")
    faixa.ClearFormats()
    v3._preencher(faixa, fundo)
    v3._borda(ws.Range(f"{col}{L_CARD_ROT}"), v3.XL_EDGE_TOP, friso, v3.XL_THICK)
    rot = ws.Range(f"{col}{L_CARD_ROT}")
    rot.NumberFormat = "@"
    v3._fonte(rot, tamanho=TAMANHO_ROTULO.get(col, 9), negrito=True, cor=texto)
    v3._alinhar(rot, recuo=1)
    val = ws.Range(f"{col}{L_CARD_VAL}")
    v3._fonte(val, tamanho=15, negrito=True, cor=texto)
    v3._alinhar(val)
    sub = ws.Range(f"{col}{L_CARD_SUB}")
    v3._fonte(sub, tamanho=9, cor=texto)
    v3._alinhar(sub, v=v3.XL_TOP, quebra=True, recuo=1)


def frente_cards(ws) -> None:
    # As regras da situacao (G8:G10) saem com o card; nenhuma outra regra
    # condicional toca E8:G10 (conferido pela contagem).
    antes = ws.Cells.FormatConditions.Count
    ws.Range(f"E{L_CARD_ROT}:G{L_CARD_SUB}").FormatConditions.Delete()
    if ws.Cells.FormatConditions.Count != antes - 2:
        raise RuntimeError("Formatacao condicional inesperada em E8:G10")
    for col in "EFG":
        fundo, texto, friso = CARDS[col][1:]
        _estilo_card(ws, col, fundo, texto, friso)
    for col in "EF":
        v3._moeda(ws.Range(f"{col}{L_CARD_VAL}"))
        ws.Range(f"{col}{L_CARD_VAL}").IndentLevel = 1
    g = ws.Range(f"G{L_CARD_VAL}")
    v3._formato(g, v3.PCT_LOC, v3.PCT_INV)
    g.HorizontalAlignment = v3.XL_RIGHT
    g.IndentLevel = 1


def frente_faixa(ws) -> None:
    ws.Range(f"{LINHA_INSERCAO}:{LINHA_INSERCAO + LINHAS_INSERIDAS - 1}").Insert()
    for linha, altura in ALTURAS_FAIXA.items():
        ws.Rows(linha).RowHeight = altura
    novas = ws.Range(f"A{LINHA_INSERCAO}:H{LINHA_INSERCAO + LINHAS_INSERIDAS - 1}")
    novas.ClearFormats()
    v3._mesclar(ws, f"C{L_SIT_ROT}:G{L_SIT_ROT}")
    v3._mesclar(ws, f"C{L_SIT_TXT}:G{L_SIT_TXT}")
    faixa = ws.Range(FAIXA_SITUACAO)
    v3._preencher(faixa, FAIXA_FUNDO)
    v3._fonte(faixa, cor=FAIXA_TEXTO)
    v3._borda(ws.Range(f"B{L_SIT_ROT}:G{L_SIT_ROT}"), v3.XL_EDGE_TOP,
              FAIXA_FRISO, v3.XL_THICK)
    rot = ws.Range(f"B{L_SIT_ROT}")
    rot.NumberFormat = "@"
    v3._fonte(rot, tamanho=9, negrito=True)
    v3._alinhar(rot, recuo=1)
    status = ws.Range(f"C{L_SIT_ROT}:G{L_SIT_ROT}")
    v3._fonte(status, tamanho=12, negrito=True)
    v3._alinhar(status, recuo=1)
    texto = ws.Range(f"C{L_SIT_TXT}:G{L_SIT_TXT}")
    v3._fonte(texto, tamanho=9.5)
    v3._alinhar(texto, v=v3.XL_TOP, quebra=True, recuo=1)
    for formula, cor in REGRAS_FAIXA:
        fundo, letra = SIT_CORES[cor]
        v3._cf_cor(faixa, formula, v3._bgr(fundo), v3._bgr(letra))


def _escrever(ws) -> None:
    formulas = formulas_executivo()
    validar_ascii_e_parenteses(formulas)
    for endereco, texto in constantes_executivo().items():
        celula = ws.Range(endereco)
        celula.NumberFormat = "@"
        celula.Value = texto
    novas = [f"{col}{linha}" for col in "EFG" for linha in (L_CARD_VAL, L_CARD_SUB)]
    novas += [f"C{L_SIT_ROT}", f"C{L_SIT_TXT}"]
    for endereco in novas:
        ws.Range(endereco).Formula = formulas[endereco]


def conferir_resultado(wb) -> None:
    """Depois da insercao: a aba executiva bate, celula a celula, com o
    descritor (inclusive o deslocamento feito pelo proprio Excel)."""
    ws = wb.Worksheets(EXE)
    esperado = formulas_executivo()
    for celula in ws.UsedRange:
        formula = celula.Formula
        if isinstance(formula, str) and formula.startswith("="):
            endereco = celula.Address.replace("$", "")
            if esperado.get(endereco) != formula:
                raise RuntimeError(f"{EXE}!{endereco} diverge do descritor: {formula}")
            esperado.pop(endereco)
    if esperado:
        raise RuntimeError(f"Formulas ausentes na executiva: {sorted(esperado)[:8]}")
    area = str(ws.PageSetup.PrintArea).replace("$", "")
    if not re.fullmatch(rf"B1:G{ULTIMA_LINHA}", area):
        raise RuntimeError(f"Area de impressao inesperada: {area}")


def aplicar(caminho: Path) -> None:
    import pythoncom
    import win32com.client as com

    pythoncom.CoInitialize()
    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = ws = None
    try:
        wb = excel.Workbooks.Open(str(caminho))
        excel.Calculation = v3.XL_CALCULO_MANUAL
        conferir(wb)
        frente_memoria(wb)
        ws = wb.Worksheets(EXE)
        frente_faixa(ws)
        frente_cards(ws)
        _escrever(ws)
        conferir_resultado(wb)
        excel.Calculation = v3.XL_CALCULO_AUTOMATICO
        excel.CalculateFullRebuild()

        ws.Activate()
        janela = excel.ActiveWindow
        janela.ScrollRow = 1
        janela.ScrollColumn = 1
        ws.Range(f"B{L_CARD_VAL}").Select()
        wb.Worksheets("CONTROLE").Activate()
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1

        wb.Save()
        v3._fechar(wb)
        wb = None
    finally:
        if wb is not None:
            v3._fechar(wb)
        for _ in range(10):
            try:
                excel.Quit()
                break
            except Exception:
                time.sleep(1.0)
        wb = ws = excel = None
        gc.collect()
        pythoncom.CoUninitialize()


def main() -> int:
    caminho = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else TEMPLATE
    if not caminho.exists():
        print("ERRO: arquivo nao encontrado: %s" % caminho)
        return 1
    aplicar(caminho)
    print("OK: RETRO. RECONHECIDO + POTENCIAL aplicado em %s" % caminho)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
