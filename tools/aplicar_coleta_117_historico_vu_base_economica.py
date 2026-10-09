# -*- coding: utf-8 -*-
"""Coleta 11.7 — historico_VU segue a base economica do VU (aditivos!N).

Aplica no template oficial, via Excel COM (openpyxl destroi a CF x14), somente:

* historico_VU!D2:G200 (VU_C1..VU_C4) — o VU de cada ciclo parte do ultimo
  reajuste ja incorporado ao VU (aditivos!O), e nao mais do nascimento fisico
  (posicao_contratual!Y). Antes do nascimento a celula segue vazia; item sem
  base economica (original ou Nxxx sem aditivos!N) mantem o resultado anterior.

As formulas vem de `_coleta_oficial` (fonte unica com a migracao runtime).
Cada celula precisa estar na forma 11.6 ou ja na 11.7; qualquer outra coisa
aborta sem gravar. Nada mais e tocado. REGRA ZERO CORRUPCAO XLSX: formulas em
ingles, ASCII e com parenteses balanceados.

uso: python tools/aplicar_coleta_117_historico_vu_base_economica.py [caminho.xlsx]
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

import _coleta_oficial as co  # noqa: E402
import aplicar_ajustes_xls_ux_pos174 as ux  # noqa: E402
from aplicar_coleta_114_aditivos_ciclo_execucao import validar_formula  # noqa: E402

TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

XL_CALCULO_MANUAL, XL_CALCULO_AUTOMATICO = -4135, -4105
PRIMEIRA, ULTIMA = co._LINHAS_HISTORICO_VU[0], co._LINHAS_HISTORICO_VU[-1]


def frente_historico_vu(wb) -> None:
    ws = wb.Worksheets("historico_VU")
    for coluna, indice in co._COLUNAS_HISTORICO_VU:
        faixa = ws.Range(f"{coluna}{PRIMEIRA}:{coluna}{ULTIMA}")
        atuais = faixa.Formula
        novas = []
        for deslocamento, (atual,) in enumerate(atuais):
            linha = PRIMEIRA + deslocamento
            nova = co._formula_historico_vu_base_economica(linha, indice)
            validar_formula(nova)
            if atual not in (co._formula_historico_vu_nascimento(linha, indice), nova):
                raise RuntimeError(
                    f"historico_VU!{coluna}{linha} fora do formato esperado; nada aplicado"
                )
            novas.append((nova,))
        faixa.Formula = tuple(novas)


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
        frente_historico_vu(wb)
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
    print("OK: Coleta 11.7 aplicada em %s" % caminho)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
