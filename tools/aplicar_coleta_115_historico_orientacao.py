# -*- coding: utf-8 -*-
"""Coleta 11.5 — historico necessario + orientacao de itens novos por aditivo.

Aplica no template oficial, via Excel COM (openpyxl destroi a CF x14), somente:

* parametros!A8 — o aviso "HISTORICO INCOMPLETO" passa a pedir apenas o
  percentual de ciclo anterior a um ciclo COMPUTADO (A="Sim"): a cadeia
  parametros!F so precisa dele para chegar ao fator do ultimo ciclo apurado.
  Antes o limite era CONTROLE!B2, que desde a Coleta 11.4 e o ciclo em
  EXECUCAO pela data de corte: analise so de C1 com corte em C4 pedia C2/C3.
  Fonte da formula: `aplicar_ajustes_xls_ux_pos174.formula_aviso_historico`.
* parametros!E3:E6 — a formatacao condicional de "percentual historico
  ausente" ganha a mesma condicao (so a regra; o formato e mantido).
* itens_Remanesc!A:C e aditivos!A/D/E (linhas 2:200) — mensagem de entrada
  (caixa que o Excel mostra ao selecionar a celula) explicando item existente,
  item novo por aditivo e acrescimo posterior. Nao restringe valores: onde nao
  havia validacao entra uma do tipo "qualquer valor"; em aditivos!D a lista
  suspensa existente e preservada e so ganha a mensagem.

Nada mais e tocado: CONTROLE!B2, CICLO_EM_EXECUCAO e as formulas economicas
seguem identicos. REGRA ZERO CORRUPCAO XLSX: formulas em ingles, ASCII e com
parenteses balanceados.

uso: python tools/aplicar_coleta_115_historico_orientacao.py [caminho.xlsx]
"""
from __future__ import annotations

import gc
import sys
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
for _caminho in (RAIZ, RAIZ / "tools"):
    if str(_caminho) not in sys.path:
        sys.path.insert(0, str(_caminho))

import aplicar_ajustes_xls_ux_pos174 as ux  # noqa: E402

TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

XL_EXPRESSION = 2
XL_VALIDAR_QUALQUER_VALOR = 0  # xlValidateInputOnly
XL_ALERTA_PARAR = 1
XL_CALCULO_MANUAL, XL_CALCULO_AUTOMATICO = -4135, -4105

FAIXA_CF_HISTORICO = "E3:E6"
FORMULA_CF_ANTERIOR = 'AND($A3="Nao",$C3<>"",$E3="")'
FORMULA_CF_HISTORICO = 'AND($A3="Nao",$C3<>"",$E3="",COUNTIF($A4:$A$7,"Sim")>0)'

# (aba, faixa, titulo, mensagem). Limites do Excel: titulo 32, mensagem 255.
ORIENTACOES = (
    ("itens_Remanesc", "A2:A200", "ITEM NOVO POR ADITIVO",
     "Cadastre o item aqui PRIMEIRO (ex.: N001), com quantidade-base original "
     "= 0 e o VU de inclusão. A quantidade do item nasce na aba aditivos.\n"
     "N001 = código do item | 0 = quantidade-base original.\n"
     "Item do contrato: código normal."),
    ("itens_Remanesc", "B2:B200", "QUANTIDADE-BASE ORIGINAL",
     "Item do contrato: quantidade original contratada.\n"
     "Item novo por aditivo: 0 (com N001, N002... o 0 aparece sozinho). "
     "Não some aqui a quantidade do aditivo: ela é lançada na aba aditivos."),
    ("itens_Remanesc", "C2:C200", "VALOR UNITÁRIO ORIGINAL/BASE",
     "Item do contrato: VU original do contrato.\n"
     "Item novo por aditivo: VU do item no momento em que foi incluído no "
     "contrato."),
    ("aditivos", "A2:A200", "ITEM DO ADITIVO",
     "Use o mesmo código de itens_Remanesc.\n"
     "Item novo: cadastre-o ANTES em itens_Remanesc (ex.: N001, "
     "quantidade-base 0, VU de inclusão) e só depois registre aqui o aditivo "
     "em que ele nasce."),
    ("aditivos", "D2:D200", "ITEM NOVO",
     "Use \"Acréscimo - novo item\" somente no aditivo em que o item nasce no "
     "contrato, com toda a quantidade inicial.\n"
     "Novo aumento desse item depois: use apenas \"Acrescimo\".\n"
     "Item já existente: \"Acrescimo\" ou \"Supressao\"."),
    ("aditivos", "E2:E200", "QUANTIDADE DO ADITIVO",
     "Informe só a VARIAÇÃO deste aditivo (ex.: base 100 + Acrescimo 20 = 120).\n"
     "Nascimento: toda a quantidade inicial (N003: 1.851, nunca 3.702).\n"
     "Aumento posterior: só o adicional (+200 -> 2.051)."),
)


def validar_orientacoes() -> None:
    for aba, faixa, titulo, mensagem in ORIENTACOES:
        if len(titulo) > 32 or len(mensagem) > 255:
            raise ValueError(f"{aba}!{faixa}: titulo/mensagem acima do limite do Excel")


# ------------------------------------------------------------- frentes
def frente_parametros(wb) -> None:
    ws = wb.Worksheets("parametros")
    aviso = ws.Range(ux.CELULA_AVISO)
    if f"{ux.MEM}!$AG$41" not in str(aviso.Formula):
        raise RuntimeError(f"parametros!{ux.CELULA_AVISO} nao e o aviso de historico")
    formula = ux.formula_aviso_historico()
    ux.validar_ascii_e_parenteses({ux.CELULA_AVISO: formula,
                                   FAIXA_CF_HISTORICO: "=" + FORMULA_CF_HISTORICO})
    aviso.Formula = formula

    rng = ws.Range(FAIXA_CF_HISTORICO)
    anterior = ux._local(rng, "=" + FORMULA_CF_ANTERIOR)
    nova = ux._local(rng, "=" + FORMULA_CF_HISTORICO)
    alvo = [
        rng.FormatConditions(i)
        for i in range(1, rng.FormatConditions.Count + 1)
        if rng.FormatConditions(i).AppliesTo.Address == "$E$3:$E$6"
        and rng.FormatConditions(i).Formula1 in (anterior, nova)
    ]
    if len(alvo) != 1:
        raise RuntimeError("CF de historico em parametros!E3:E6 nao encontrada")
    alvo[0].Modify(XL_EXPRESSION, None, nova)


def frente_orientacao(wb) -> None:
    validar_orientacoes()
    for aba, faixa, titulo, mensagem in ORIENTACOES:
        validacao = wb.Worksheets(aba).Range(faixa).Validation
        try:
            lista = validacao.Formula1  # existe validacao (ex.: dropdown)
        except Exception:
            lista = None
            validacao.Add(XL_VALIDAR_QUALQUER_VALOR, XL_ALERTA_PARAR)
        validacao.InputTitle = titulo
        validacao.InputMessage = mensagem
        validacao.ShowInput = True
        if lista is not None and validacao.Formula1 != lista:
            raise RuntimeError(f"{aba}!{faixa}: lista suspensa alterada")


def aplicar(caminho: Path) -> None:
    import pythoncom
    import win32com.client as com

    pythoncom.CoInitialize()
    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(caminho))
        excel.Calculation = XL_CALCULO_MANUAL
        frente_parametros(wb)
        frente_orientacao(wb)
        excel.Calculation = XL_CALCULO_AUTOMATICO
        excel.CalculateFullRebuild()
        wb.Worksheets("CONTROLE").Activate()
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1
        wb.Save()
        ux._fechar(wb)
        wb = None
    finally:
        if wb is not None:
            ux._fechar(wb)
        for _ in range(10):
            try:
                excel.Quit()
                break
            except Exception:
                time.sleep(1.0)
        wb = excel = None
        gc.collect()
        pythoncom.CoUninitialize()


def main() -> int:
    caminho = Path(sys.argv[1] if len(sys.argv) > 1 else TEMPLATE).resolve()
    if not caminho.exists():
        print("ERRO: arquivo nao encontrado: %s" % caminho)
        return 1
    aplicar(caminho)
    print("OK: Coleta 11.5 aplicada em %s" % caminho)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
