"""Ajuste pontual: identificacao inequivoca de itens novos por aditivo.

Substitui a orientacao ambigua "novo item = 0" por: ITEM = N001, N002... e
QTD_BASE_ORIGINAL = 0. Alteracao MINIMA de texto (orientacao + mensagem do check),
via Microsoft Excel real (COM), sem openpyxl para gravar. Nao altera matematica,
VTA, layout, cores, EIXOS, dropdowns nem qualquer logica.

Edicoes (somente texto):
  1. itens_Remanesc!A1  — cabecalho ITEM com orientacao de novo item (N001...).
  2. aditivos!A1        — idem no cabecalho ITEM da aba aditivos.
  3. aditivos!M2:M200   — mensagem do check de item ausente cita N001/N002.

O valor "ITEM" permanece no inicio dos cabecalhos (chave de mapeamento do leitor).
FAIL-CLOSED: recusa reaplicacao (A1 ja cita N001). Promove so se salvar sem erros.
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16
XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105

IR_A1 = ("ITEM\n(NOVO ITEM: use N001, N002...; QTD_BASE_ORIGINAL=0 e informe o VU_ORIGINAL)")
AD_A1 = ("ITEM\n(NOVO ITEM: cadastrar em itens_Remanesc como N001, N002...; BASE 0 + VU)")
MSG_ANTIGA = '"NOVO ITEM NAO CADASTRADO - CADASTRAR EM itens_Remanesc (QTD BASE 0 + VU)"'
MSG_NOVA = '"NOVO ITEM NAO CADASTRADO - CADASTRAR EM itens_Remanesc COMO N001, N002... (BASE 0 + VU)"'


def _validar(wb) -> None:
    nomes = [ws.Name for ws in wb.Worksheets]
    for aba in ("itens_Remanesc", "aditivos"):
        if aba not in nomes:
            raise ValueError(f"Aba ausente: {aba}")
    if "N001" in str(wb.Worksheets("itens_Remanesc").Range("A1").Value or ""):
        raise ValueError("itens_Remanesc!A1 ja cita N001; ajuste ja aplicado?")


def _aplicar(wb) -> None:
    wb.Worksheets("itens_Remanesc").Range("A1").Value = IR_A1
    ad = wb.Worksheets("aditivos")
    ad.Range("A1").Value = AD_A1
    f = str(ad.Range("M2").Formula)
    novo = f.replace(MSG_ANTIGA, MSG_NOVA)
    if novo == f:
        raise ValueError("Mensagem antiga do check nao encontrada em aditivos!M2.")
    ad.Range("M2:M200").Formula = novo


def _verificar_sem_erros(wb) -> None:
    import pywintypes
    problemas = []
    for ws in wb.Worksheets:
        for tipo in (XL_CELLTYPE_FORMULAS, XL_CELLTYPE_CONSTANTS):
            try:
                cels = ws.UsedRange.SpecialCells(tipo, XL_ERRORS)
                problemas.append(f"{ws.Name}!{cels.Address}")
            except pywintypes.com_error:
                continue
    if problemas:
        raise RuntimeError(f"Erros de formula apos recalculo: {problemas}")


def aplicar(origem: Path, destino: Path) -> None:
    origem, destino = Path(origem), Path(destino)
    if not origem.is_file():
        raise FileNotFoundError(origem)
    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_novoitem_"))
    tmp = tmp_dir / origem.name
    shutil.copyfile(origem, tmp)
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = True
    wb = None
    salvo = False
    try:
        wb = excel.Workbooks.Open(str(tmp), UpdateLinks=0)
        excel.ScreenUpdating = False
        excel.Calculation = XL_CALC_MANUAL
        aba_ativa = wb.ActiveSheet.Name
        _validar(wb)
        _aplicar(wb)
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
    shutil.copyfile(tmp, destino)
    shutil.rmtree(tmp_dir, ignore_errors=True)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("origem", type=Path)
    ap.add_argument("destino", type=Path)
    args = ap.parse_args()
    aplicar(args.origem, args.destino)


if __name__ == "__main__":
    main()
