# -*- coding: utf-8 -*-
"""Hotfix 26C via Excel COM real — NUNCA openpyxl para salvar.

Desacopla VALOR HISTORICO ATUALIZADO de EFEITO FINANCEIRO no template oficial
e aplica os ajustes de UX aprovados no diagnostico pos-26B:

  1. itens_PC!E2:E100 — FATOR_ACUMULADO passa a ser o fator HISTORICO do ciclo
     (memoria parametros!A11:E15), sem depender de L (EFEITO_FINANCEIRO_PC).
     Fonte unica da formula: _formula_e() do tool canonico da Etapa 3
     (tools/aplicar_efeitos_financeiros_itens_pc.py), ja atualizado.
     L continua governando retroativo (H/J) e a cor vermelha ($L2="Nao").
  2. parametros!G12 — dropdown passa a aceitar N/A (novo nome definido
     OPCOES_SIM_NAO_NA -> parametros!T2:T4, coluna tecnica ja oculta).
     G13:G15 permanecem com OPCOES_SIM_NAO (Sim/Nao).
  3. RESULTADOS!B259 — status do C1 no Eixo 5 ganha ramo proprio para N/A
     ("NAO SE APLICA - ANTERIOR E C0"), sem tratar N/A como formalizacao
     (I259 continua testando somente "Sim").
  4. RESULTADOS — colunas S:T (bloco tecnico VTA-PC S19:T25) ocultas, como
     G:O e X:Y; formulas preservadas.
  5. RESULTADOS!A28:C29 — comparacao executiva de referencia por REFERENCIA
     (sem matematica nova): B28 ==$B$26 (VTA aplicado) e
     B29 ==comparativo_VTA!$B$208 (contrato integralmente reajustado).

Idempotente. FAIL-CLOSED: valida layout/celulas antes de escrever; qualquer
falha interrompe sem promover o destino.
"""
from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

sys.path.insert(0, str(Path(__file__).resolve().parent))
from aplicar_efeitos_financeiros_itens_pc import LIMITE_PC, _formula_e

XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105
XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16
XL_VALIDATE_LIST = 3
XL_VALID_ALERT_STOP = 1
XL_BETWEEN = 1

MOEDA_BR = "R$ #.##0,00"
NOME_OPCOES_NA = "OPCOES_SIM_NAO_NA"

ROTULO_COMP_A28 = "COMPARACAO DE REFERENCIA - VTA pelo metodo aplicado"
ROTULO_COMP_A29 = (
    "COMPARACAO DE REFERENCIA - Contrato original integralmente reajustado"
)
NOTA_COMP = (
    "Valor meramente comparativo. Nao utilizar como VTA, retroativo ou base "
    "de pagamento."
)

B259_FALLBACK_ANTIGO = '"FORA - NAO COMPROVADO"'
B259_RAMO_NA = (
    'IF(parametros!$G$12="N/A","NAO SE APLICA - ANTERIOR E C0",'
    '"FORA - NAO COMPROVADO")'
)


def _aplicar_formula_e(wb) -> None:
    ws = wb.Worksheets("itens_PC")
    if str(ws.Cells(1, 5).Value or "") != "FATOR_ACUMULADO":
        raise ValueError("itens_PC!E1 nao e FATOR_ACUMULADO; layout inesperado.")
    destino = ws.Range(f"E2:E{LIMITE_PC}")
    destino.Formula = [[_formula_e(r)] for r in range(2, LIMITE_PC + 1)]


def _aplicar_na_g12(wb) -> None:
    p = wb.Worksheets("parametros")
    t2 = str(p.Range("T2").Value or "").strip()
    t3 = str(p.Range("T3").Value or "").strip()
    if (t2, t3) != ("Sim", "Nao"):
        raise ValueError(
            f"parametros!T2:T3 inesperado ({t2!r},{t3!r}); "
            "esperado intervalo tecnico Sim/Nao da rodada de homologacao."
        )
    p.Range("T4").Value = "N/A"
    nomes = [n.Name for n in wb.Names]
    if NOME_OPCOES_NA not in nomes:
        wb.Names.Add(Name=NOME_OPCOES_NA, RefersTo="=parametros!$T$2:$T$4")
    val = p.Range("G12").Validation
    try:
        val.Delete()
    except Exception:
        pass
    val.Add(XL_VALIDATE_LIST, XL_VALID_ALERT_STOP, XL_BETWEEN,
            f"={NOME_OPCOES_NA}")
    val.IgnoreBlank = True
    val.InCellDropdown = True
    # G13:G15 permanecem com OPCOES_SIM_NAO (nao tocar).


def _patch_b259(wb) -> None:
    r = wb.Worksheets("RESULTADOS")
    formula = str(r.Range("B259").Formula or "")
    if 'parametros!$G$12="N/A"' in formula:
        return  # idempotente
    if B259_FALLBACK_ANTIGO not in formula:
        raise ValueError(
            "RESULTADOS!B259 sem o fallback 'FORA - NAO COMPROVADO'; "
            "layout inesperado."
        )
    r.Range("B259").Formula = formula.replace(
        B259_FALLBACK_ANTIGO, B259_RAMO_NA, 1
    )


def _ocultar_st(wb) -> None:
    r = wb.Worksheets("RESULTADOS")
    r.Columns("S:T").Hidden = True


def _comparativo_executivo(wb, excel) -> None:
    r = wb.Worksheets("RESULTADOS")
    a28 = str(r.Range("A28").Value or "")
    if a28 == ROTULO_COMP_A28:
        return  # idempotente
    ocupadas = excel.WorksheetFunction.CountA(r.Range("A28:W29"))
    if ocupadas != 0:
        raise ValueError(
            "RESULTADOS!A28:W29 nao esta livre; comparativo executivo "
            "nao aplicado (nenhuma linha e inserida)."
        )
    r.Range("A28").Value = ROTULO_COMP_A28
    r.Range("B28").Formula = "=$B$26"
    r.Range("A29").Value = ROTULO_COMP_A29
    r.Range("B29").Formula = "=comparativo_VTA!$B$208"
    r.Range("C28").Value = NOTA_COMP
    r.Range("B28:B29").NumberFormatLocal = MOEDA_BR
    r.Range("A28:A29").Font.Italic = True
    r.Range("C28").Font.Italic = True


def _verificar_sem_erros(wb) -> None:
    import pywintypes
    problemas = []
    for ws in wb.Worksheets:
        for tipo in (XL_CELLTYPE_FORMULAS, XL_CELLTYPE_CONSTANTS):
            try:
                cel = ws.UsedRange.SpecialCells(tipo, XL_ERRORS)
                problemas.append(f"{ws.Name}!{cel.Address}")
            except pywintypes.com_error:
                continue
    if problemas:
        raise RuntimeError(f"Erros de formula apos recalculo: {problemas}")


def aplicar(origem: Path, destino: Path) -> None:
    origem, destino = Path(origem), Path(destino)
    if not origem.is_file():
        raise FileNotFoundError(origem)
    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_26c_"))
    tmp_xlsx = tmp_dir / origem.name
    shutil.copyfile(origem, tmp_xlsx)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    salvo = False
    try:
        wb = excel.Workbooks.Open(str(tmp_xlsx), UpdateLinks=0)
        excel.ScreenUpdating = False
        excel.Calculation = XL_CALC_MANUAL
        aba_ativa = wb.ActiveSheet.Name
        _aplicar_formula_e(wb)
        _aplicar_na_g12(wb)
        _patch_b259(wb)
        _ocultar_st(wb)
        _comparativo_executivo(wb, excel)
        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()
        _verificar_sem_erros(wb)
        if aba_ativa in [ws.Name for ws in wb.Worksheets]:
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
    print("HOTFIX 26C aplicado:", args.destino)


if __name__ == "__main__":
    main()
