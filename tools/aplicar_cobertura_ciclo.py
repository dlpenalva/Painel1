# -*- coding: utf-8 -*-
"""Cobertura fisica ATUAL confirmada — origem automatica em CICLO_EM_EXECUCAO.

Substitui, na aba cobertura_temporal, a redacao e a fonte do campo B8:

* A8 (rotulo)  -> "COBERTURA FISICA ATUAL CONFIRMADA ATE"
* B8 (valor)   -> automatico: data de CICLO_EM_EXECUCAO!D5 quando a posicao
                  itemizada esta COMPLETA e VALIDA (sinal: CICLO_EM_EXECUCAO!A9
                  nao vazio — o total so aparece com data + itens + zero erros/
                  incompletos). Caso contrario, vazio. Via INDIRECT+ISERROR:
                  compativel com arquivos antigos sem a aba (permanece vazio,
                  sem #REF!, sem reparo).
* C8 (ajuda)   -> explicacao do comportamento automatico; remove a instrucao
                  antiga (QTD_REM_ATUAL em posicao_referencia + CONTROLE!B3).

NAO altera VTA/retroativo/B26/T25 nem as 3 referencias. B8 nao tem consumidores
(nem formula nem leitor por rotulo). Mutacao via Excel COM (padrao zero-corrupcao).
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105
XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16

ABA = "cobertura_temporal"
DATA_BR = "dd/mm/aaaa"

A8_TEXTO = "COBERTURA FISICA ATUAL CONFIRMADA ATE"
B8_FORMULA = (
    '=IF(ISERROR(INDIRECT("CICLO_EM_EXECUCAO!A9")),"",'
    'IF(INDIRECT("CICLO_EM_EXECUCAO!A9")="","",'
    'INDIRECT("CICLO_EM_EXECUCAO!D5")))'
)
C8_TEXTO = (
    "Automatico: apresenta a data informada em CICLO_EM_EXECUCAO quando a "
    "posicao itemizada estiver completa e valida. Caso contrario, permanece vazio."
)


def _nomes_abas(wb) -> list[str]:
    return [ws.Name for ws in wb.Worksheets]


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


def _snapshot_travas(wb) -> dict[str, str]:
    mem = wb.Worksheets("MEMORIA_RESULTADOS")
    return {e: str(mem.Range(e).Formula) for e in ("B26", "T25")}


def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem).resolve()
    destino = Path(destino).resolve()
    if not origem.is_file():
        raise FileNotFoundError(origem)
    if origem == destino:
        raise ValueError("Origem e destino devem ser diferentes.")

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_cobertura_"))
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
        excel.ScreenUpdating = False
        excel.Calculation = XL_CALC_MANUAL
        aba_ativa = wb.ActiveSheet.Name
        if ABA not in _nomes_abas(wb):
            raise ValueError(f"Aba {ABA} ausente; layout inesperado.")
        travas_antes = _snapshot_travas(wb)

        ws = wb.Worksheets(ABA)
        if bool(ws.ProtectContents):
            ws.Unprotect()
        ws.Range("A8").Value = A8_TEXTO
        ws.Range("B8").Formula = B8_FORMULA
        ws.Range("B8").NumberFormatLocal = DATA_BR
        ws.Range("C8").Value = C8_TEXTO

        travas_depois = _snapshot_travas(wb)
        if travas_antes != travas_depois:
            raise RuntimeError(f"TRAVA VIOLADA B26/T25: {travas_antes} -> {travas_depois}")

        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()
        _verificar_sem_erros(wb)
        if aba_ativa in _nomes_abas(wb):
            wb.Worksheets(aba_ativa).Activate()
        wb.Save()
        salvo = True
        wb.Close(SaveChanges=False)
        wb = None

        # Reabre sem reparo.
        wb = excel.Workbooks.Open(
            str(tmp_xlsx), UpdateLinks=0, ReadOnly=True, CorruptLoad=0
        )
        if str(wb.Worksheets(ABA).Range("A8").Value or "") != A8_TEXTO:
            raise RuntimeError("A8 nao persistiu apos reabertura.")
        wb.Close(SaveChanges=False)
        wb = None
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
    print("cobertura_temporal (origem CICLO_EM_EXECUCAO) aplicada:", args.destino)


if __name__ == "__main__":
    main()
