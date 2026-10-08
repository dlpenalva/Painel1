"""Coleta 11.4 — aditivos com fator vigente e ciclo em execucao pela data de corte.

Aplica no template oficial, via Excel COM (openpyxl destroi a CF x14), somente:

* CONTROLE!B2 — "Ciclo vigente (em execucao)" passa a ser DERIVADO da data de
  corte (CONTROLE!B3) e das janelas ja calculadas em parametros!C2:C6: o ultimo
  ciclo cujo inicio e <= data de corte. Dentro das janelas equivale a
  inicio <= corte <= fim; nenhuma data de ciclo e recalculada. O ciclo
  ANALISADO continua em parametros!A (COMPUTAR) e CONTROLE!B12.
* aditivos!I2:I200 — fator efetivo do item = fator vigente no marco / fator
  vigente no nascimento (posicao_contratual!Y). O fator vigente e o ultimo
  fator conhecido da cadeia parametros!F (F2:F6 formam prefixo numerico), o
  mesmo carregamento da MEMORIA DO FATOR APLICAVEL, sem preencher
  parametros!F.
* aditivos!J2:J200 — mesmo encadeamento de VU do historico_VU (identico a ele
  sempre que o fator do ciclo existe); so preenche a lacuna do fator vazio.
* aditivos!M2:M200 — alerta para "Acrescimo - novo item" fora do nascimento.
* aditivos!D2:D200 — dropdown ganha "Acrescimo - novo item".

uso: python tools/aplicar_coleta_114_aditivos_ciclo_execucao.py [caminho.xlsx]
"""
from __future__ import annotations

import gc
import sys
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

OPCOES_TIPO = ("Acrescimo", "Acréscimo - novo item", "Supressao")

FORMULA_CICLO_VIGENTE = (
    '=IF(NOT(ISNUMBER($B$3)),"",'
    'IF(AND(ISNUMBER(parametros!$C$6),$B$3>=parametros!$C$6),"C4",'
    'IF(AND(ISNUMBER(parametros!$C$5),$B$3>=parametros!$C$5),"C3",'
    'IF(AND(ISNUMBER(parametros!$C$4),$B$3>=parametros!$C$4),"C2",'
    'IF(AND(ISNUMBER(parametros!$C$3),$B$3>=parametros!$C$3),"C1",'
    'IF(AND(ISNUMBER(parametros!$C$2),$B$3>=parametros!$C$2),"C0",""))))))'
)

_N = '(MATCH($C2,{"C0","C1","C2","C3","C4"},0)-1)'
_Y = 'INDEX(posicao_contratual!$Y$2:$Y$200,MATCH($A2,posicao_contratual!$A$2:$A$200,0))'


def _fv(k: str) -> str:
    """Fator vigente no ciclo k: ultimo fator conhecido de parametros!F."""
    return f'INDEX(parametros!$F$2:$F$6,MIN({k}+1,COUNT(parametros!$F$2:$F$6)))'


FORMULA_FATOR = (
    f'=IFERROR(IF(OR($A2="",$C2=""),"",IF(NOT(ISNUMBER({_Y})),"",'
    f'IF({_Y}>{_N},"",{_fv(_N)}/{_fv(_Y)}))),"")'
)
FORMULA_VALOR = (
    f'=IF(OR(L2="",F2=""),"",ROUND(L2*IF(AND(UPPER(H2)="SIM",ISNUMBER(I2)),'
    f'IF({_Y}={_N},F2,ROUND(F2*{_fv(_N)}/{_fv(_Y)},2)),F2),2))'
)

ALERTA_NOVO_ITEM = (
    "ALERTA: NOVO ITEM INVALIDO - use Acrescimo se o item ja existia; "
    "novo item exige QTD_BASE_ORIGINAL 0, VU_ORIGINAL e ser o 1o acrescimo do item."
)
_FIM_M_ANTIGO = '"ALERTA: TIPO_INVALIDO","OK"))))))'
_FIM_M_NOVO = (
    '"ALERTA: TIPO_INVALIDO",IF(AND(ISNUMBER(SEARCH("NOVO",D2)),'
    'OR(INDEX(posicao_contratual!$C$2:$C$200,MATCH(A2,posicao_contratual!$A$2:$A$200,0))<>0,'
    'F2="",COUNTIFS($A$2:$A$200,A2,$B$2:$B$200,"<"&B2,$L$2:$L$200,">0")>0)),'
    f'"{ALERTA_NOVO_ITEM}","OK")))))))'
)

XL_CALCULO_MANUAL = -4135
XL_CALCULO_AUTOMATICO = -4105
XL_VALIDATE_LIST = 3


def validar_formula(formula: str) -> None:
    if not formula.isascii():
        raise ValueError("formula com caractere nao ASCII")
    if formula.count("(") != formula.count(")"):
        raise ValueError("parenteses desbalanceados")


def _formula_m(formula_atual: str) -> str:
    if _FIM_M_NOVO in formula_atual:
        return formula_atual
    if not formula_atual.endswith(_FIM_M_ANTIGO):
        raise RuntimeError("aditivos!M2 fora do formato esperado; nada aplicado")
    return formula_atual[: -len(_FIM_M_ANTIGO)] + _FIM_M_NOVO


def aplicar(caminho: Path) -> None:
    import pythoncom
    import win32com.client as com

    for formula in (FORMULA_CICLO_VIGENTE, FORMULA_FATOR, FORMULA_VALOR):
        validar_formula(formula)

    pythoncom.CoInitialize()
    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = ws = None
    try:
        wb = excel.Workbooks.Open(str(caminho))
        excel.Calculation = XL_CALCULO_MANUAL

        ws = wb.Worksheets("CONTROLE")
        protegida = bool(ws.ProtectContents)
        if protegida:
            ws.Unprotect()
        ws.Range("B2").Formula = FORMULA_CICLO_VIGENTE
        if protegida:
            ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True)

        ws = wb.Worksheets("aditivos")
        if ws.ProtectContents:
            raise RuntimeError("aba aditivos protegida inesperadamente")
        formula_m = _formula_m(str(ws.Range("M2").Formula))
        validar_formula(formula_m)
        ws.Range("I2:I200").Formula = FORMULA_FATOR
        ws.Range("J2:J200").Formula = FORMULA_VALOR
        ws.Range("M2:M200").Formula = formula_m
        validacao = ws.Range("D2:D200").Validation
        separador = excel.International[5 - 1]  # xlListSeparator (pt-BR: ";")
        validacao.Modify(
            Type=XL_VALIDATE_LIST,
            AlertStyle=validacao.AlertStyle,
            Operator=validacao.Operator,
            Formula1=separador.join(OPCOES_TIPO),
        )

        excel.Calculation = XL_CALCULO_AUTOMATICO
        wb.Worksheets("CONTROLE").Activate()
        wb.Save()
        wb.Close(False)
        wb = None
    finally:
        if wb is not None:
            wb.Close(False)
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
    print("OK: Coleta 11.4 (aditivos + ciclo em execucao) aplicada em %s" % caminho)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
