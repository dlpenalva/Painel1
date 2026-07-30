# -*- coding: utf-8 -*-
"""Etapa 26J: integra aditivos uma vez na posicao contratual/remanescente.

Gravador unico do template oficial. Usa Microsoft Excel real em copia
temporaria, recalcula, verifica erros, salva, reabre sem reparo e somente entao
promove o destino.

Semantica:
* QTD_CONTRATADA_Cn continua sendo base + deltas acumulados.
* QTD_REM_BASE_Cn e a fotografia anterior as alteracoes do proprio ciclo.
* QTD_REM_AJUSTADA_Cn = fotografia base + DELTA_Cn.
* Para item que nasce em Cn, a fotografia anterior e zero canonico e o
  remanescente ajustado nasce do proprio DELTA_Cn.
* Antes do nascimento, quantidade, VU e total ficam vazios.
* Depois do nascimento, quantidade vigente continua contratual; remanescente
  posterior continua fail-closed quando nao ha fotografia base.
* itens_RC consome VU temporal do historico_VU e remanescente ajustado da
  posicao_contratual. Nenhum aditivo e somado novamente no VTA.
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16
XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105
FIM = 200
_RPC_E_CALL_REJECTED = -2147418111

_PROT_FLAGS = (
    "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows",
    "AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks",
    "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting",
    "AllowFiltering", "AllowUsingPivotTables",
)


def _retry(funcao, tentativas: int = 15, espera: float = 2.0):
    for tentativa in range(tentativas):
        try:
            return funcao()
        except pywintypes.com_error as exc:
            if exc.hresult == _RPC_E_CALL_REJECTED and tentativa < tentativas - 1:
                time.sleep(espera)
                continue
            raise


def _capturar_protecao(ws):
    if not bool(ws.ProtectContents):
        return None, None
    estado = {}
    for flag in _PROT_FLAGS:
        try:
            estado[flag] = getattr(ws.Protection, flag)
        except Exception:
            estado[flag] = True
    try:
        selecao = ws.EnableSelection
    except Exception:
        selecao = None
    ws.Unprotect()
    return estado, selecao


def _restaurar_protecao(ws, estado, selecao) -> None:
    if estado is None:
        return
    ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True, **estado)
    if selecao is not None:
        try:
            ws.EnableSelection = selecao
        except Exception:
            pass


def _formula_rem_ajustada(
    nascimento_maximo: int, rem_base: str, delta: str
) -> str:
    return (
        '=IF(A2="","",IF($Y2="","",'
        f'IF($Y2>{nascimento_maximo},"",'
        f'IF(ISNUMBER({rem_base}2),ROUND({rem_base}2+{delta}2,2),'
        f'IF($Y2={nascimento_maximo},ROUND({delta}2,2),"")))))'
    )


def _aplicar_posicao_contratual(ws) -> None:
    estado, selecao = _capturar_protecao(ws)
    esperados = {
        "G": '=IF(A2="","",E2)',
        "K": '=IF(OR(A2="",J2=""),"",ROUND(J2,2))',
        "O": '=IF(OR(A2="",N2=""),"",ROUND(N2,2))',
        "S": '=IF(OR(A2="",R2=""),"",ROUND(R2,2))',
        "W": '=IF(OR(A2="",V2=""),"",ROUND(V2,2))',
    }
    for col, antigo in esperados.items():
        atual = str(ws.Range(f"{col}2").Formula or "")
        if atual != antigo and "IF($Y2" not in atual:
            raise RuntimeError(
                f"posicao_contratual!{col}2 fora do padrao 26H/26J: {atual!r}"
            )

    ws.Range(f"G2:G{FIM}").Formula = (
        '=IF(A2="","",IF($Y2="","",IF($Y2>0,"",ROUND(E2,2))))'
    )
    for col, n, rem_base, delta in (
        ("K", 1, "J", "H"),
        ("O", 2, "N", "L"),
        ("S", 3, "R", "P"),
        ("W", 4, "V", "T"),
    ):
        ws.Range(f"{col}2:{col}{FIM}").Formula = _formula_rem_ajustada(
            n, rem_base, delta
        )
    _restaurar_protecao(ws, estado, selecao)


def _aplicar_historico_vu(ws) -> None:
    estado, selecao = _capturar_protecao(ws)
    ws.Range(f"B2:B{FIM}").Formula = (
        '=IF(A2="","",IF(posicao_contratual!$Y2="",'
        '"",IF(posicao_contratual!$Y2>0,"",posicao_contratual!G2)))'
    )
    for col, n, ref in (
        ("N", 0, "E"), ("O", 1, "I"), ("P", 2, "M"),
        ("Q", 3, "Q"), ("R", 4, "U"),
    ):
        ws.Range(f"{col}2:{col}{FIM}").Formula = (
            '=IF(A2="","",IF(posicao_contratual!$Y2="",'
            f'"",IF(posicao_contratual!$Y2>{n},"",'
            f'posicao_contratual!{ref}2)))'
        )
    _restaurar_protecao(ws, estado, selecao)


def _aplicar_itens_rc(ws) -> None:
    estado, selecao = _capturar_protecao(ws)
    for col, vu_col in (
        ("B", "C"), ("E", "D"), ("H", "E"), ("K", "F"), ("N", "G")
    ):
        ws.Range(f"{col}3:{col}{FIM + 1}").Formula = (
            f'=IF(OR(historico_VU!A2="",historico_VU!{vu_col}2=""),'
            f'"",historico_VU!{vu_col}2)'
        )
    for total_col, vu_col, qtd_col in (
        ("D", "B", "C"), ("G", "E", "F"), ("J", "H", "I"),
        ("M", "K", "L"), ("P", "N", "O"),
    ):
        ws.Range(f"{total_col}3:{total_col}{FIM + 1}").Formula = (
            f'=IF(A3="TOTAL",ROUND(SUM(${total_col}$3:{total_col}2),2),'
            f'IF(OR(A3="",{vu_col}3="",{qtd_col}3=""),"",'
            f'ROUND({vu_col}3*{qtd_col}3,2)))'
        )
    _restaurar_protecao(ws, estado, selecao)


def _verificar_sem_erros(wb) -> None:
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


def _verificar_formulas(wb) -> None:
    pc = wb.Worksheets("posicao_contratual")
    hv = wb.Worksheets("historico_VU")
    rc = wb.Worksheets("itens_RC")
    if "J2+H2" not in str(pc.Range("K2").Formula or ""):
        raise RuntimeError("posicao_contratual!K2 nao integra DELTA_C1.")
    if "IF($Y2=1" not in str(pc.Range("K2").Formula or ""):
        raise RuntimeError("posicao_contratual!K2 nao trata nascimento em C1.")
    if "posicao_contratual!$Y2>0" not in str(hv.Range("N2").Formula or ""):
        raise RuntimeError("historico_VU!N2 nao deixa pre-nascimento vazio.")
    if "historico_VU!C2" not in str(rc.Range("B3").Formula or ""):
        raise RuntimeError("itens_RC!B3 nao consome o VU temporal.")
    if 'OR(A3="",B3="",C3="")' not in str(rc.Range("D3").Formula or ""):
        raise RuntimeError("itens_RC!D3 nao preserva vazio antes do nascimento.")
    if str(wb.Worksheets("itens_Remanesc").Range("B201").Formula or ""):
        raise RuntimeError("B201 voltou a receber formula.")
    if "LEN(TRIM($A2))=4" not in str(pc.Range("Z2").Formula or ""):
        raise RuntimeError("Regex Nxxx 26H.2 foi alterada.")


def _verificar_destino(excel, caminho: Path) -> None:
    wb = excel.Workbooks.Open(
        str(caminho), UpdateLinks=0, ReadOnly=True, CorruptLoad=0
    )
    try:
        _verificar_formulas(wb)
        _verificar_sem_erros(wb)
    finally:
        wb.Close(SaveChanges=False)


def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem).resolve()
    destino = Path(destino).resolve()
    if not origem.is_file():
        raise FileNotFoundError(origem)

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_26j_"))
    tmp_xlsx = tmp_dir / origem.name
    shutil.copyfile(origem, tmp_xlsx)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    salvo = False
    try:
        wb = excel.Workbooks.Open(
            str(tmp_xlsx), UpdateLinks=0, ReadOnly=False, CorruptLoad=0
        )

        def _preparar():
            excel.ScreenUpdating = False
            excel.Calculation = XL_CALC_MANUAL
            return wb.ActiveSheet.Name

        aba_ativa = _retry(_preparar)
        _aplicar_posicao_contratual(wb.Worksheets("posicao_contratual"))
        _aplicar_historico_vu(wb.Worksheets("historico_VU"))
        _aplicar_itens_rc(wb.Worksheets("itens_RC"))
        try:
            wb.Worksheets(aba_ativa).Activate()
        except Exception:
            pass
        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()
        _verificar_formulas(wb)
        _verificar_sem_erros(wb)
        wb.Save()
        salvo = True
        try:
            wb.Close(SaveChanges=False)
        except TypeError:
            pass
        wb = None
        _verificar_destino(excel, tmp_xlsx)
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
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
    print("Etapa 26J aplicada:", args.destino)


if __name__ == "__main__":
    main()
