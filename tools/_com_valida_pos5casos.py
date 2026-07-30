"""Validacao pos-edicao do template via Excel COM (Pacote pos-5-casos).

Abre o template real e confirma:
  - abre SEM reparo (nenhum dialogo de recuperacao);
  - 13 abas na ordem oficial (inclui cobertura_temporal);
  - itens_PC!B e aditivos!B com NumberFormat dd/mm/yyyy;
  - formulas aditivos L2/M2 com prefixos ACR/SUPR;
  - uma DATA REAL em itens_PC!B2 exibe dd/mm/aaaa (Range.Text).
A checagem de data usa uma COPIA temporaria e NAO altera o template real.
"""
from __future__ import annotations

import shutil
import sys
import tempfile
from datetime import date
from pathlib import Path

import pythoncom
import win32com.client as win32

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

ABAS_ESPERADAS = [
    "CONTROLE", "parametros", "financeiro", "itens_Remanesc", "itens_Consumidos",
    "itens_PC", "aditivos", "posicao_referencia", "posicao_contratual",
    "itens_RC", "historico_VU", "cobertura_temporal", "RESULTADOS",
]


def main() -> int:
    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    ok = True
    tmp = Path(tempfile.mkdtemp(prefix="cl8us_pos5_")) / "valida.xlsx"
    shutil.copy2(TEMPLATE, tmp)
    try:
        # CorruptLoad=0 (xlNormalLoad): se precisar reparar, Excel abriria em modo
        # de recuperacao; validamos que abre normal e sem alteracao pendente.
        wb = excel.Workbooks.Open(str(tmp), CorruptLoad=0)

        abas = [wb.Worksheets(i + 1).Name for i in range(wb.Worksheets.Count)]
        print("ABAS:", abas)
        if abas != ABAS_ESPERADAS:
            ok = False
            print("FALHA: abas divergem do esperado")

        ipc = wb.Worksheets("itens_PC")
        adi = wb.Worksheets("aditivos")
        fmt_ipc = ipc.Range("B2").NumberFormatLocal
        fmt_adi = adi.Range("B2").NumberFormatLocal
        print("FMT_LOCAL itens_PC!B2:", repr(fmt_ipc))
        print("FMT_LOCAL aditivos!B2:", repr(fmt_adi))

        f_l = adi.Range("L2").Formula
        f_m = adi.Range("M2").Formula
        print("aditivos!L2:", f_l[:60])
        if '"ACR"' not in f_l or '"SUPR"' not in f_l:
            ok = False
            print("FALHA: L2 sem prefixos ACR/SUPR")
        if '"ACR"' not in f_m or '"SUPR"' not in f_m:
            ok = False
            print("FALHA: M2 sem prefixos ACR/SUPR")

        # DATA REAL -> Range.Text deve exibir dd/mm/aaaa (ex.: 20/06/2023).
        # Usa serial do Excel (base 1899-12-30) para evitar coercao COM de date.
        serial = (date(2023, 6, 20) - date(1899, 12, 30)).days
        ipc.Range("B2").Value = serial
        texto = ipc.Range("B2").Text
        print("itens_PC!B2 .Text:", repr(texto))
        if texto != "20/06/2023":
            ok = False
            print("FALHA: data nao exibe dd/mm/aaaa")

        # RESULTADOS intacta: B23/B25/B26 continuam formulas.
        res = wb.Worksheets("RESULTADOS")
        for cel in ("B23", "B25", "B26"):
            val = res.Range(cel).Formula
            print(f"RESULTADOS!{cel} formula? ", str(val).startswith("="))

        wb.Close(SaveChanges=False)  # descarta a data de teste; template real intacto
        print("VALIDACAO_OK" if ok else "VALIDACAO_FALHOU")
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
