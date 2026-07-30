"""Edicao CIRURGICA do template oficial via Microsoft Excel COM (Pacote pos-5-casos).

Aplica APENAS, no template real, sem reconstruir o workbook nem substituir formulas
inteiras (preserva textos de mensagem, validacoes e demais estilos):
  - §10: itens_PC!B2:B100  -> NumberFormatLocal dd/mm/aaaa (corrige literal "yyyy").
  - §11: aditivos!B2:B200   -> NumberFormatLocal dd/mm/aaaa.
  - §10: itens_PC!B2:B100 borda esquerda/direita 'thin' p/ coincidir com a coluna A
         (harmoniza a borda ja existente na coluna A; consistencia de coluna).
  - §12: aditivos!L2:L200 e M2:M200 -> SUBSTITUICAO CIRURGICA dos prefixos de tipo
         (ACRES->ACR, SUPRES->SUPR) preservando todo o restante da formula, inclusive
         a mensagem "NOVO ITEM NAO CADASTRADO ...".

NAO usa openpyxl para gravar. NAO altera a validacao (dropdown) de D. NAO toca RESULTADOS.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pythoncom
import win32com.client as win32

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

XL_EDGE_LEFT = 7
XL_EDGE_RIGHT = 10
XL_CONTINUOUS = 1
XL_THIN = 2


def _corrigir_prefixos(formula: str) -> str:
    """Troca so os prefixos de tipo, preservando o resto da formula."""
    return (
        formula
        .replace(',5)="ACRES"', ',3)="ACR"')
        .replace(',6)="SUPRES"', ',4)="SUPR"')
        .replace(',5)<>"ACRES"', ',3)<>"ACR"')
        .replace(',6)<>"SUPRES"', ',4)<>"SUPR"')
    )


def main() -> int:
    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    try:
        wb = excel.Workbooks.Open(str(TEMPLATE))
        ipc = wb.Worksheets("itens_PC")
        adi = wb.Worksheets("aditivos")

        # §10 / §11 — datas dd/mm/aaaa (token de ano local "a").
        ipc.Range("B2:B100").NumberFormatLocal = "dd/mm/aaaa"
        adi.Range("B2:B200").NumberFormatLocal = "dd/mm/aaaa"

        # §10 — borda da coluna B coincidente com a coluna A (esq/dir 'thin').
        for edge in (XL_EDGE_LEFT, XL_EDGE_RIGHT):
            borda = ipc.Range("B2:B100").Borders(edge)
            borda.LineStyle = XL_CONTINUOUS
            borda.Weight = XL_THIN

        # §12 — prefixos tolerantes a acento, preservando mensagem/estrutura.
        for r in range(2, 201):
            for col in ("L", "M"):
                cel = adi.Range(f"{col}{r}")
                f = cel.Formula
                if isinstance(f, str) and f.startswith("="):
                    nova = _corrigir_prefixos(f)
                    if nova != f:
                        cel.Formula = nova

        wb.Save()
        wb.Close(SaveChanges=False)
        print("COM_EDIT_OK")
        return 0
    finally:
        excel.Quit()


if __name__ == "__main__":
    sys.exit(main())
