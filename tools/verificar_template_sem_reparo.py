"""Abre o template no Excel com alertas normais para confirmar ausencia de reparo.

Uso: python tools/verificar_template_sem_reparo.py
Resultado: ABERTURA SEM REPARO: OK ou FALHOU (excecao/timeout)
"""
from __future__ import annotations
import gc
import sys
import time
from pathlib import Path

import pythoncom
import win32com.client

TEMPLATE = Path(__file__).resolve().parent.parent / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"


def main() -> None:
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    try:
        excel.Visible = True
    except AttributeError:
        pass
    try:
        excel.AskToUpdateLinks = False
    except AttributeError:
        pass

    pasta = None
    try:
        print(f"Abrindo: {TEMPLATE}")
        pasta = excel.Workbooks.Open(str(TEMPLATE.resolve()), UpdateLinks=0, CorruptLoad=0)
        time.sleep(2)

        n_abas = pasta.Sheets.Count
        nome_ativo = pasta.ActiveSheet.Name if pasta.ActiveSheet else "(desconhecido)"
        print(f"Aberto com sucesso: {n_abas} abas, aba ativa: {nome_ativo}")
        print("ABERTURA SEM REPARO: OK")

        pasta.Close(SaveChanges=False)
        pasta = None
        excel.Quit()
        print("Excel fechado.")
    except Exception as exc:
        print(f"ABERTURA FALHOU: {exc}")
        if pasta is not None:
            try:
                pasta.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass
        sys.exit(1)
    finally:
        gc.collect()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
