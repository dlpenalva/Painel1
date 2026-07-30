"""Teste funcional §12.3 via Excel COM sobre uma COPIA do template.

Para cada variante de tipo (com/sem acento) injeta um aditivo minimo valido e
confere DELTA (aditivos!L2) e CHECK_POSICAO_CONTRATUAL (aditivos!M2):
  Acréscimo/Acrescimo -> DELTA > 0 e CHECK = OK
  Supressão/Supressao -> DELTA < 0 e CHECK = OK
Nao altera o template real.
"""
from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

import pythoncom
import win32com.client as win32

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

CASOS = [
    ("Acréscimo", +1),
    ("Acrescimo", +1),
    ("Supressão", -1),
    ("Supressao", -1),
]


def main() -> int:
    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    ok = True
    tmp = Path(tempfile.mkdtemp(prefix="cl8us_acc_")) / "func.xlsx"
    shutil.copy2(TEMPLATE, tmp)
    try:
        wb = excel.Workbooks.Open(str(tmp))
        rem = wb.Worksheets("itens_Remanesc")
        adi = wb.Worksheets("aditivos")
        rem.Range("A2").Value = "ITX"          # item existente p/ CHECK
        adi.Range("A2").Value = "ITX"
        adi.Range("C2").Value = "C1"           # ciclo valido (string)
        adi.Range("E2").Value = 10.0           # quantidade
        for tipo, sinal in CASOS:
            adi.Range("D2").Value = tipo
            excel.CalculateFull()
            delta = adi.Range("L2").Value
            check = adi.Range("M2").Value
            sinal_ok = (delta is not None) and (
                (sinal > 0 and delta > 0) or (sinal < 0 and delta < 0)
            )
            check_ok = check == "OK"
            print(f"{tipo:12s} DELTA={delta!r:>8} CHECK={check!r} -> "
                  f"{'OK' if (sinal_ok and check_ok) else 'FALHA'}")
            if not (sinal_ok and check_ok):
                ok = False
        wb.Close(SaveChanges=False)
        print("FUNC_ACENTO_OK" if ok else "FUNC_ACENTO_FALHOU")
        return 0 if ok else 1
    finally:
        excel.Quit()
        try:
            tmp.unlink()
            tmp.parent.rmdir()
        except OSError:
            pass


if __name__ == "__main__":
    sys.exit(main())
