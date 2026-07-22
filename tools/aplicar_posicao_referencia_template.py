"""Aplicador final da Etapa 3 (Posicao de Referencia do Contrato) ao template.

Aplicador UNICO e cirurgico (evolucao do aplicador da posicao atual). Alteracao
estritamente ADITIVA/governada:

  1. CONTROLE!B3 (rotulo "Data de corte (unica p/ contrato)"): formaliza como
     DATA_POSICAO_ATUAL (opcional) — desbloqueio, formato dd/mm/aaaa e validacao
     de data. Ja lida pelo Python como controle["data_corte"].

  2. Nova aba operacional "posicao_referencia" (apos "aditivos"): tela simples.
     Visiveis A:F — ITEM (ligado a itens_Remanesc), QTD_REM_ATUAL (unico campo
     manual, opcional), QTD_CONTRATADA_NA_DATA (auto, por data), QTD_REM_REFERENCIA
     (auto: atual OU fotografia), ORIGEM/SITUACAO e CHECK. Painel de referencia em
     H1:I8. Colunas tecnicas K:R ocultas (nunca veryHidden).

  3. RESULTADOS: bloco "POSICAO DE REFERENCIA DO CONTRATO" (A267:B277). O valor
     remanescente atualizado na posicao de referencia NAO e o VTA — B23 (VTA
     calculado) e B26 (VTA FINAL) permanecem o VTA oficial e inalterados.

CONCEITO: a posicao de referencia representa o remanescente mais recente e
comprovado disponivel e A QUE DATA ele se refere.
  - Se o fiscal informa a posicao ATUAL de forma COMPLETA (B3 valida no horizonte
    e QTD_REM_ATUAL para todos os itens): a referencia e a posicao atual.
  - Caso contrario (nada informado, parcial, data sem quantidades, quantidades
    sem data): FALLBACK automatico para a ULTIMA FOTOGRAFIA HISTORICA VALIDA — o
    ultimo ciclo cronologico com DATA_INICIO definida e QTD_REM_AJUSTADA valida.
    DATA_POSICAO_REFERENCIA = DATA_INICIO (abertura) desse ciclo. Sem posicao
    hibrida: a referencia consolidada usa integralmente a fotografia; os valores
    atuais informados sao preservados apenas para conferencia.

Nao usa MAX(C0:C4), nem ciclo futuro, nem TODAY()/HOJE(); reutiliza a linha
temporal canonica (parametros!C/D) e o motor posicao_contratual (QTD_REM_AJUSTADA).

CICLO da posicao atual: regra canonica de enquadramento (financeiro!B2) sobre
CONTROLE!B3. QTD_CONTRATADA_NA_DATA = base + deltas de aditivos com
"Data do aditivo" (aditivos!B) <= DATA_POSICAO_REFERENCIA — semantica temporal
comprovada (o motor homologado enquadra o aditivo pela mesma janela parametros!C/D).

QTD_EXEC_DESDE_ABERTURA = MAX(REM_ABERTURA_REF + DELTA_APOS - QTD_REM_REFERENCIA, 0).
No fallback a referencia E a abertura: DELTA_APOS=0 e QTD_REM_REFERENCIA=abertura,
logo execucao=0 (a fotografia representa exatamente aquela data). Execucao nunca
inferida pelo tempo transcorrido.

Gravador unico: Microsoft Excel real via COM (copia temporaria; promove so se
salvar sem erros de formula). FAIL-CLOSED: recusa reaplicacao. NAO altera
B23/B25/B26, aditivos, template legado, Python de producao nem paginas.
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

XL_PASTE_FORMATS = -4122
XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16
XL_VALIDATE_DECIMAL = 2
XL_VALIDATE_DATE = 4
XL_VALID_ALERT_STOP = 1
XL_GREATER_EQUAL = 7
XL_BETWEEN = 1
XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105

FIM = 200
ABA = "posicao_referencia"
FMT_DATA = "dd/mm/aaaa"   # pt-BR: token de ano "aaaa" (aplicar em celula livre)
FMT_NUM = "#.##0,00;-#.##0,00"

CTRL_DATA = "B3"  # DATA_POSICAO_ATUAL (rotulo "Data de corte (unica p/ contrato)")

R_TIT = 267
R_FIM = 277

CAB = ["ITEM", "QTD_REM_ATUAL", "QTD_CONTRATADA_NA_DATA",
       "QTD_REM_REFERENCIA", "ORIGEM / SITUACAO", "CHECK"]

# Mapas ciclo -> referencia canonica
_FATOR = {"C0": "parametros!$F$2", "C1": "parametros!$F$3", "C2": "parametros!$F$4",
          "C3": "parametros!$F$5", "C4": "parametros!$F$6"}
_INICIO = {"C0": "parametros!$C$2", "C1": "parametros!$C$3", "C2": "parametros!$C$4",
           "C3": "parametros!$C$5", "C4": "parametros!$C$6"}
# QTD_REM_AJUSTADA por ciclo (fotografia historica), row-relative (linha 2)
_FOTO = {"C0": "posicao_contratual!G2", "C1": "posicao_contratual!K2",
         "C2": "posicao_contratual!O2", "C3": "posicao_contratual!S2",
         "C4": "posicao_contratual!W2"}


def _pick(cell: str, mapa: dict, default: str = '""') -> str:
    """Nested-IF: escolhe o valor de `mapa` conforme o ciclo em `cell`."""
    expr = default
    for cic in ("C4", "C3", "C2", "C1", "C0"):
        expr = f'IF({cell}="{cic}",{mapa[cic]},{expr})'
    return expr


def _dstr(ref: str) -> str:
    """dd/mm/aaaa locale-proof (DAY/MONTH/YEAR) para rotulos textuais."""
    return (f'(RIGHT("0"&DAY({ref}),2)&"/"&RIGHT("0"&MONTH({ref}),2)&"/"&YEAR({ref}))')


# CICLO da posicao atual (enquadramento canonico sobre CONTROLE!B3) — espelha
# financeiro!B2. Retorna C0..C4, "" (vazio/invalida) ou "FORA DO HORIZONTE".
CIC_ATUAL = (
    '=IF(CONTROLE!$B$3="","",IF(NOT(ISNUMBER(CONTROLE!$B$3)),"",'
    'IF(AND(ISNUMBER(parametros!$C$2),ISNUMBER(parametros!$D$2),CONTROLE!$B$3>=parametros!$C$2,CONTROLE!$B$3<=parametros!$D$2),"C0",'
    'IF(AND(ISNUMBER(parametros!$C$3),ISNUMBER(parametros!$D$3),CONTROLE!$B$3>=parametros!$C$3,CONTROLE!$B$3<=parametros!$D$3),"C1",'
    'IF(AND(ISNUMBER(parametros!$C$4),ISNUMBER(parametros!$D$4),CONTROLE!$B$3>=parametros!$C$4,CONTROLE!$B$3<=parametros!$D$4),"C2",'
    'IF(AND(ISNUMBER(parametros!$C$5),ISNUMBER(parametros!$D$5),CONTROLE!$B$3>=parametros!$C$5,CONTROLE!$B$3<=parametros!$D$5),"C3",'
    'IF(AND(ISNUMBER(parametros!$C$6),ISNUMBER(parametros!$D$6),CONTROLE!$B$3>=parametros!$C$6,CONTROLE!$B$3<=parametros!$D$6),"C4",'
    '"FORA DO HORIZONTE")))))))'
)

# Ultima fotografia historica valida: maior ciclo (cronologico) com DATA_INICIO
# definida e QTD_REM_AJUSTADA presente. Nao usa MAX de valores nem ciclo futuro
# (ciclo futuro nao possui fotografia preenchida).
CICLO_FALLBACK = (
    '=IF(AND(ISNUMBER(parametros!$C$6),COUNT(posicao_contratual!$W$2:$W$200)>0),"C4",'
    'IF(AND(ISNUMBER(parametros!$C$5),COUNT(posicao_contratual!$S$2:$S$200)>0),"C3",'
    'IF(AND(ISNUMBER(parametros!$C$4),COUNT(posicao_contratual!$O$2:$O$200)>0),"C2",'
    'IF(AND(ISNUMBER(parametros!$C$3),COUNT(posicao_contratual!$K$2:$K$200)>0),"C1",'
    'IF(AND(ISNUMBER(parametros!$C$2),COUNT(posicao_contratual!$G$2:$G$200)>0),"C0",'
    '"")))))'
)


def _validar_layout(wb, excel) -> None:
    wf = excel.WorksheetFunction
    nomes = [ws.Name for ws in wb.Worksheets]
    if ABA in nomes or "remanescente_atual" in nomes:
        raise ValueError("Aba de posicao ja existe; bloco Etapa 3 ja aplicado?")
    ws_r = wb.Worksheets("RESULTADOS")
    if wf.CountA(ws_r.Range(f"A{R_TIT}:B{R_FIM}")) != 0:
        raise ValueError(f"RESULTADOS A{R_TIT}:B{R_FIM} nao vazia; bloco ja aplicado?")
    for aba in ("CONTROLE", "parametros", "itens_Remanesc", "aditivos",
                "posicao_contratual", "RESULTADOS"):
        if aba not in nomes:
            raise ValueError(f"Aba canonica ausente: {aba}")
    ctrl = wb.Worksheets("CONTROLE")
    if "Data de corte" not in str(ctrl.Range("A3").Value or ""):
        raise ValueError("CONTROLE!A3 nao e o rotulo de data de corte.")
    if "Data do aditivo" not in str(wb.Worksheets("aditivos").Range("B1").Value or ""):
        raise ValueError("aditivos!B1 (Data do aditivo) ausente; sem granularidade temporal.")


_PROT_FLAGS = (
    "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows",
    "AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks",
    "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting",
    "AllowFiltering", "AllowUsingPivotTables",
)


def _formalizar_controle(ws) -> None:
    """DATA_POSICAO_ATUAL = CONTROLE!B3: desbloqueio + dd/mm/aaaa + validacao.

    CONTROLE e protegida (sem senha) com B3 bloqueada; desbloqueamos B3 (como B1)
    para o fiscal poder informar a data, e restauramos a protecao original.
    """
    protegida = bool(ws.ProtectContents)
    estado, sel = {}, None
    if protegida:
        p = ws.Protection
        for f in _PROT_FLAGS:
            try:
                estado[f] = getattr(p, f)
            except Exception:
                estado[f] = True
        try:
            sel = ws.EnableSelection
        except Exception:
            sel = None
        ws.Unprotect()

    cel = ws.Range(CTRL_DATA)
    cel.Locked = False
    cel.NumberFormatLocal = FMT_DATA
    val = cel.Validation
    try:
        val.Delete()
    except Exception:
        pass
    import datetime as _dt
    _base = _dt.date(1899, 12, 30)
    _lo = (_dt.date(1990, 1, 1) - _base).days
    _hi = (_dt.date(2199, 12, 31) - _base).days
    val.Add(Type=XL_VALIDATE_DATE, AlertStyle=XL_VALID_ALERT_STOP,
            Operator=XL_BETWEEN, Formula1=_lo, Formula2=_hi)
    val.IgnoreBlank = True
    val.InputTitle = "Data da posicao atual (opcional)"
    val.InputMessage = ("Data da fotografia da posicao ATUAL (dd/mm/aaaa). Opcional: "
                        "se vazia, o Claus usa a ultima fotografia historica valida.")
    val.ErrorMessage = "Informe uma data valida no formato dd/mm/aaaa."

    if protegida:
        ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True, **estado)
        if sel is not None:
            try:
                ws.EnableSelection = sel
            except Exception:
                pass


def _criar_aba(wb, excel):
    ws = wb.Worksheets.Add(After=wb.Worksheets("aditivos"))
    ws.Name = ABA
    modelo = wb.Worksheets("itens_Remanesc")

    modelo.Range("A1").Copy()
    ws.Range("A1:F1").PasteSpecial(XL_PASTE_FORMATS)
    excel.CutCopyMode = False
    ws.Range("A1:F1").Value = [CAB]

    foto = _pick("$I$4", _FOTO)

    # Painel de referencia (H1:I8) — informativo, visivel.
    painel = [
        ("CICLO DA POSICAO ATUAL (B3)", CIC_ATUAL, None),
        ("POSICAO ATUAL COMPLETA?",
         '=AND(ISNUMBER(CONTROLE!$B$3),OR($I$1="C0",$I$1="C1",$I$1="C2",$I$1="C3",$I$1="C4"),'
         'COUNTIF(itens_Remanesc!$A$2:$A$200,"<>")>0,'
         'COUNT($B$2:$B$200)>=COUNTIF(itens_Remanesc!$A$2:$A$200,"<>"))', None),
        ("CICLO DA ULTIMA FOTOGRAFIA VALIDA", CICLO_FALLBACK, None),
        ("CICLO DE REFERENCIA", '=IF($I$2,$I$1,$I$3)', None),
        ("DATA DA POSICAO DE REFERENCIA", '=IF($I$4="","",IF($I$2,CONTROLE!$B$3,$I$6))', FMT_DATA),
        ("DATA DE ABERTURA DO CICLO DE REFERENCIA", "=" + _pick("$I$4", _INICIO), FMT_DATA),
        ("FATOR ACUM. DO CICLO DE REFERENCIA", "=" + _pick("$I$4", _FATOR), "0,000000"),
        ("ORIGEM DA POSICAO",
         '=IF($I$4="","POSICAO DE REFERENCIA INDISPONIVEL",'
         'IF($I$2,"POSICAO ATUAL INFORMADA - "&' + _dstr("$I$5") + ','
         'IF(COUNT($B$2:$B$200)>0,"POSICAO ATUAL INCOMPLETA - UTILIZADA ABERTURA "&$I$4,'
         '"ABERTURA DO CICLO "&$I$4&" - "&' + _dstr("$I$5") + ')))', None),
    ]
    for i, (rot, formula, fmt) in enumerate(painel, start=1):
        ws.Range(f"H{i}").Value = rot
        ws.Range(f"I{i}").Formula = formula
        if fmt:
            ws.Range(f"I{i}").NumberFormatLocal = fmt
    ws.Range("H1:H8").Font.Bold = True

    modelo.Range("A2").Copy()
    ws.Range(f"A2:F{FIM}").PasteSpecial(XL_PASTE_FORMATS)
    excel.CutCopyMode = False

    colunas = {
        "A": '=IF(itens_Remanesc!A2="","",itens_Remanesc!A2)',
        "C": ('=IF(A2="","",IF($I$5="","POSICAO DE REFERENCIA INDISPONIVEL",'
              'ROUND(itens_Remanesc!B2+SUMIFS(aditivos!$L$2:$L$200,aditivos!$A$2:$A$200,A2,'
              'aditivos!$B$2:$B$200,"<="&$I$5),2)))'),
        "D": (f'=IF(A2="","",IF($I$4="","",IF($I$2,IF(ISNUMBER(B2),B2,""),{foto})))'),
        "K": '=IF(A2="","",itens_Remanesc!C2)',
        "L": f'=IF(A2="","",{foto})',
        "M": ('=IF(OR(A2="",$I$5="",$I$6=""),"",'
              'ROUND(SUMIFS(aditivos!$L$2:$L$200,aditivos!$A$2:$A$200,A2,'
              'aditivos!$B$2:$B$200,">"&$I$6,aditivos!$B$2:$B$200,"<="&$I$5),2))'),
        "N": ('=IF(OR(A2="",L2="",D2=""),"",'
              'ROUND(MAX(L2+IF(ISNUMBER(M2),M2,0)-D2,0),2))'),
        "O": '=IF(OR(A2="",D2="",K2="",$I$7=""),"",ROUND(D2*K2*$I$7,2))',
        "P": '=IF(OR(A2="",N2="",K2="",$I$7=""),"",ROUND(N2*K2*$I$7,2))',
        "Q": '=IF(OR(A2="",D2="",K2=""),"",ROUND(D2*K2,2))',
        "R": '=IF(OR(A2="",L2="",K2="",$I$7=""),"",ROUND(L2*K2*$I$7,2))',
        "E": ('=IF(A2="","",IF($I$4="","POSICAO DE REFERENCIA INDISPONIVEL",'
              'IF($I$2,"ATUAL INFORMADA",'
              'IF(ISNUMBER(B2),"ATUAL IGNORADA (POSICAO INCOMPLETA) - FOTOGRAFIA "&$I$4,'
              '"FOTOGRAFIA "&$I$4))))'),
        "F": ('=IF(A2="","",'
              'IF(COUNTIF($A$2:$A$200,A2)>1,"ITEM DUPLICADO",'
              'IF($I$4="","POSICAO DE REFERENCIA INDISPONIVEL",'
              'IF(AND(B2<>"",NOT(ISNUMBER(B2))),"QTD NAO NUMERICA",'
              'IF(AND(ISNUMBER(B2),B2<0),"NEGATIVO NAO PERMITIDO",'
              'IF(NOT(ISNUMBER(C2)),"POSICAO CONTRATUAL INDISPONIVEL",'
              'IF(D2="","REMANESCENTE DE REFERENCIA INDISPONIVEL",'
              'IF(AND(ISNUMBER(D2),ISNUMBER(C2),D2>C2),"REMANESCENTE DE REFERENCIA SUPERA POSICAO CONTRATUAL",'
              '"OK"))))))))'),
    }
    for col, formula in colunas.items():
        ws.Range(f"{col}2:{col}{FIM}").Formula = formula

    ws.Range(f"C2:F{FIM}").NumberFormatLocal = FMT_NUM
    ws.Range(f"B2:B{FIM}").NumberFormatLocal = FMT_NUM
    ws.Range(f"K2:R{FIM}").NumberFormatLocal = FMT_NUM
    ws.Range(f"E2:F{FIM}").NumberFormatLocal = "@"  # E/F sao textos

    val = ws.Range(f"B2:B{FIM}").Validation
    try:
        val.Delete()
    except Exception:
        pass
    val.Add(Type=XL_VALIDATE_DECIMAL, AlertStyle=XL_VALID_ALERT_STOP,
            Operator=XL_GREATER_EQUAL, Formula1="0")
    val.IgnoreBlank = True
    val.InputTitle = "QTD_REM_ATUAL (opcional)"
    val.InputMessage = ("Quantidade remanescente ATUAL na data informada. Opcional; "
                        "zero permitido. Se nao informada para todos os itens, usa-se "
                        "a fotografia historica.")
    val.ErrorMessage = "QTD_REM_ATUAL deve ser >= 0 (negativo nao permitido)."

    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B:F").ColumnWidth = 22
    ws.Columns("H").ColumnWidth = 38
    ws.Columns("I").ColumnWidth = 18
    ws.Columns("K:R").Hidden = True  # tecnicas; reexibiveis, nunca veryHidden
    return ws


def _aplicar_bloco_resultados(ws, excel) -> None:
    titulo = ("POSICAO DE REFERENCIA DO CONTRATO | valor de referencia - NAO e o VTA; "
              "B23 (VTA calculado) e B26 (VTA FINAL) permanecem o VTA oficial")
    ws.Range(f"A{R_TIT}").Value = titulo
    ws.Range("A19").Copy()
    ws.Range(f"A{R_TIT}").PasteSpecial(XL_PASTE_FORMATS)
    excel.CutCopyMode = False

    status = (
        f'=IF({ABA}!$I$4="","POSICAO DE REFERENCIA INDISPONIVEL",'
        f'IF(COUNTIF({ABA}!$F$2:$F$200,"REMANESCENTE DE REFERENCIA SUPERA POSICAO CONTRATUAL")>0,'
        '"INCONSISTENTE - remanescente supera posicao contratual",'
        f'IF({ABA}!$I$2,"POSICAO ATUAL INFORMADA - RECONCILIAVEL",'
        f'IF(COUNT({ABA}!$B$2:$B$200)>0,"POSICAO ATUAL INCOMPLETA - REFERENCIA HISTORICA ABERTURA "&{ABA}!$I$4,'
        f'"REFERENCIA HISTORICA - ABERTURA "&{ABA}!$I$4))))'
    )
    conserv = (
        f'=IF({ABA}!$I$4="","",'
        f'IF(ABS((SUM({ABA}!$L$2:$L$200)+SUM({ABA}!$M$2:$M$200))-'
        f'(SUM({ABA}!$N$2:$N$200)+SUM({ABA}!$D$2:$D$200)))<=0.01,'
        '"OK - conservacao demonstrada","ALERTA - conservacao nao fecha (revisar)"))'
    )
    linhas = [
        (268, "Data da posicao de referencia", f'=IF({ABA}!$I$5="","",{ABA}!$I$5)', FMT_DATA),
        (269, "Origem da posicao de referencia", f'={ABA}!$I$8', None),
        (270, "Ciclo de referencia", f'=IF({ABA}!$I$4="","INDISPONIVEL",{ABA}!$I$4)', None),
        (271, "Valor remanescente na abertura do ciclo de referencia",
         f'=IF(COUNT({ABA}!$R$2:$R$200)=0,"",ROUND(SUM({ABA}!$R$2:$R$200),2))', FMT_NUM),
        (272, "Valor remanescente atualizado na posicao de referencia",
         f'=IF(COUNT({ABA}!$O$2:$O$200)=0,"",ROUND(SUM({ABA}!$O$2:$O$200),2))', FMT_NUM),
        (273, "Qtd executada desde a abertura (reconciliavel)",
         f'=IF(COUNT({ABA}!$N$2:$N$200)=0,"",ROUND(SUM({ABA}!$N$2:$N$200),2))', FMT_NUM),
        (274, "Valor executado desde a abertura (reconciliavel)",
         f'=IF(COUNT({ABA}!$P$2:$P$200)=0,"",ROUND(SUM({ABA}!$P$2:$P$200),2))', FMT_NUM),
        (275, "Status da reconciliacao", status, None),
        (276, "Identidade de conservacao (abertura+delta = exec+referencia)", conserv, None),
    ]
    for lin, rot, formula, fmt in linhas:
        ws.Range(f"A{lin}").Value = rot
        ws.Range(f"B{lin}").Formula = formula
        if fmt:
            ws.Range(f"B{lin}").NumberFormatLocal = fmt

    nota = ("Resultado de referencia (Etapa 3). O valor remanescente atualizado na "
            "posicao de referencia NAO e o VTA: B23 e B26 permanecem o VTA oficial. "
            "Sem posicao atual completa, usa-se a ultima fotografia historica valida "
            "(abertura do ciclo) e a execucao desde a abertura e 0. Manual completo "
            "prevalece; execucao nao inferida por tempo; sem TODAY()/HOJE().")
    ws.Range(f"A{R_FIM}").Value = nota


def _verificar_sem_erros(wb) -> None:
    import pywintypes
    problemas = []
    for ws in wb.Worksheets:
        for tipo in (XL_CELLTYPE_FORMULAS, XL_CELLTYPE_CONSTANTS):
            try:
                celulas = ws.UsedRange.SpecialCells(tipo, XL_ERRORS)
                problemas.append(f"{ws.Name}!{celulas.Address}")
            except pywintypes.com_error:
                continue
    if problemas:
        raise RuntimeError(f"Erros de formula apos recalculo: {problemas}")


def aplicar(origem: Path, destino: Path) -> None:
    origem, destino = Path(origem), Path(destino)
    if not origem.is_file():
        raise FileNotFoundError(origem)

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_etapa3_"))
    tmp_xlsx = tmp_dir / origem.name
    shutil.copyfile(origem, tmp_xlsx)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = True
    wb = None
    salvo = False
    try:
        wb = excel.Workbooks.Open(str(tmp_xlsx), UpdateLinks=0)
        excel.ScreenUpdating = False
        excel.Calculation = XL_CALC_MANUAL
        aba_ativa = wb.ActiveSheet.Name
        _validar_layout(wb, excel)
        _formalizar_controle(wb.Worksheets("CONTROLE"))
        _criar_aba(wb, excel)
        _aplicar_bloco_resultados(wb.Worksheets("RESULTADOS"), excel)
        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()
        _verificar_sem_erros(wb)
        wb.Worksheets(aba_ativa).Activate()
        wb.Save()
        salvo = True
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()
        del wb
        del excel
        pythoncom.CoUninitialize()

    if not salvo:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise RuntimeError("Excel nao salvou; destino preservado.")

    destino.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(tmp_xlsx, destino)
    shutil.rmtree(tmp_dir, ignore_errors=True)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("origem", type=Path)
    parser.add_argument("destino", type=Path)
    args = parser.parse_args()
    aplicar(args.origem, args.destino)


if __name__ == "__main__":
    main()
