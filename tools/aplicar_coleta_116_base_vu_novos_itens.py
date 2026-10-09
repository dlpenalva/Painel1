# -*- coding: utf-8 -*-
"""Coleta 11.6 — base economica do VU em aditivos + destaque dos itens Nxxx.

Aplica no template oficial, via Excel COM (openpyxl destroi a CF x14), somente:

* aditivos!N1:N200 — "ULTIMO REAJUSTE JA INCORPORADO AO VU" (C0..C4), lista
  suspensa com mensagem de entrada; cinza nas linhas que nao sao novo item.
* aditivos!O1:O200 — INDICE_BASE_ECONOMICA_VU (coluna tecnica oculta).
* aditivos!I/J/M2:M200 — fator = fator-alvo / fator ja incorporado ao VU;
  valor atualizado e status passam a respeitar a base economica.
* itens_Remanesc!A2:AC200 — fonte verde-escura para Nxxx, com a MENOR
  prioridade (erros e alertas prevalecem); nenhum preenchimento e tocado.

As formulas vem de `_coleta_oficial` (fonte unica com a migracao runtime).
Nada mais e tocado: A:M seguem no lugar, CICLO_EM_EXECUCAO e a cadeia oficial
do VTA seguem identicas. REGRA ZERO CORRUPCAO XLSX: formulas em ingles, ASCII
e com parenteses balanceados.

uso: python tools/aplicar_coleta_116_base_vu_novos_itens.py [caminho.xlsx]
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

XL_EXPRESSION = 2
XL_VALIDAR_LISTA = 3
XL_ALERTA_PARAR = 1
XL_ENTRE = 1
XL_SEM_PADRAO = -4142  # xlNone
XL_CALCULO_MANUAL, XL_CALCULO_AUTOMATICO = -4135, -4105
BORDAS = (7, 8, 9, 10, 11, 12)  # esquerda, topo, base, direita, internas
PRIMEIRA, ULTIMA = 2, 200


def _bgr(rgb_hex: str) -> int:
    """'RRGGBB' -> inteiro BGR do Excel COM."""
    r, g, b = (int(rgb_hex[i:i + 2], 16) for i in (0, 2, 4))
    return (b << 16) | (g << 8) | r


def _copiar_formato(origem, destino) -> None:
    """Copia fonte, preenchimento, bordas e alinhamento SEM levar CF/validacao."""
    destino.Font.Name = origem.Font.Name
    destino.Font.Size = origem.Font.Size
    destino.Font.Bold = origem.Font.Bold
    destino.Font.Color = origem.Font.Color
    destino.Interior.Pattern = origem.Interior.Pattern
    if origem.Interior.Pattern != XL_SEM_PADRAO:
        destino.Interior.Color = origem.Interior.Color
    destino.HorizontalAlignment = origem.HorizontalAlignment
    destino.VerticalAlignment = origem.VerticalAlignment
    destino.WrapText = origem.WrapText
    destino.NumberFormat = origem.NumberFormat
    destino.Locked = origem.Locked
    for indice in BORDAS:
        try:
            borda, alvo = origem.Borders(indice), destino.Borders(indice)
            alvo.LineStyle = borda.LineStyle
            if borda.LineStyle != XL_SEM_PADRAO:
                alvo.Weight = borda.Weight
                alvo.Color = borda.Color
        except Exception:  # bordas internas nao existem em celula unica
            continue


def _formulas_coluna(funcao) -> tuple[tuple[str], ...]:
    formulas = tuple((funcao(linha),) for linha in range(PRIMEIRA, ULTIMA + 1))
    for (formula,) in formulas:
        validar_formula(formula)
    return formulas


def frente_aditivos(wb) -> None:
    ws = wb.Worksheets("aditivos")
    for linha in range(1, ULTIMA + 1):
        valor = ws.Range(f"N{linha}").Formula
        if valor not in ("", co._CABECALHO_BASE_ECONOMICA_VU):
            raise RuntimeError(f"aditivos!N{linha} nao esta livre: {valor!r}")

    ws.Range("N1").Value = co._CABECALHO_BASE_ECONOMICA_VU
    _copiar_formato(ws.Range("H1"), ws.Range("N1"))
    _copiar_formato(ws.Range(f"H{PRIMEIRA}:H{ULTIMA}"), ws.Range(f"N{PRIMEIRA}:N{ULTIMA}"))
    ws.Range("O1").Value = co._CABECALHO_INDICE_BASE_VU
    _copiar_formato(ws.Range("M1"), ws.Range("O1"))
    _copiar_formato(ws.Range(f"M{PRIMEIRA}:M{ULTIMA}"), ws.Range(f"O{PRIMEIRA}:O{ULTIMA}"))
    ws.Columns("N").ColumnWidth = 34
    ws.Columns("O").Hidden = True

    colunas = (
        ("O", co._formula_indice_base_economica),
        ("I", co._formula_fator_base_economica),
        ("J", co._formula_valor_aditivo),
        ("M", co._formula_status_aditivo),
    )
    for coluna, funcao in colunas:
        ws.Range(f"{coluna}{PRIMEIRA}:{coluna}{ULTIMA}").Formula = _formulas_coluna(funcao)

    alvo = ws.Range(co._FAIXA_BASE_ECONOMICA_VU)
    alvo.Validation.Delete()
    # Lista no separador da interface (pt-BR: ";"); com "," o Excel grava UM
    # item "C0,C1,..." como x12ac:list, que o openpyxl descarta.
    separador = wb.Application.International[5 - 1]  # xlListSeparator
    alvo.Validation.Add(
        XL_VALIDAR_LISTA, XL_ALERTA_PARAR, XL_ENTRE,
        separador.join(co._CICLOS_BASE_ECONOMICA),
    )
    alvo.Validation.IgnoreBlank = True
    alvo.Validation.InCellDropdown = True
    alvo.Validation.InputTitle = co._DV_BASE_VU_TITULO
    alvo.Validation.InputMessage = co._DV_BASE_VU_MENSAGEM
    alvo.Validation.ErrorTitle = co._DV_BASE_VU_ERRO_TITULO
    alvo.Validation.ErrorMessage = co._DV_BASE_VU_ERRO
    alvo.Validation.ShowInput = True
    alvo.Validation.ShowError = True

    formula_cinza = "=" + co._FORMULA_CF_BASE_VU_NAO_APLICAVEL
    validar_formula(formula_cinza)
    alvo.FormatConditions.Delete()
    regra = alvo.FormatConditions.Add(XL_EXPRESSION, None, ux._local(alvo, formula_cinza))
    regra.Interior.Color = _bgr(co._COR_BASE_VU_NAO_APLICAVEL)
    regra.SetLastPriority()


def frente_novos_itens(wb) -> None:
    ws = wb.Worksheets("itens_Remanesc")
    rng = ws.Range(co._FAIXA_DESTAQUE_NOVOS_ITENS)
    formula = "=" + co._FORMULA_DESTAQUE_NOVOS_ITENS
    validar_formula(formula)
    # itens_Remanesc!BH2 (celula auxiliar de _local) e ocupada; a formula nao
    # cita aba e a linha 2 e a mesma, logo a traducao em aditivos e identica.
    local = ux._local(wb.Worksheets("aditivos").Range("A2"), formula)
    for i in range(rng.FormatConditions.Count, 0, -1):  # reaplicacao idempotente
        regra = rng.FormatConditions(i)
        if regra.AppliesTo.Address == rng.Address and regra.Formula1 == local:
            regra.Delete()
    regra = rng.FormatConditions.Add(XL_EXPRESSION, None, local)
    regra.Font.Color = _bgr(co._COR_FONTE_NOVOS_ITENS)
    regra.SetLastPriority()


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
        frente_aditivos(wb)
        frente_novos_itens(wb)
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
    print("OK: Coleta 11.6 aplicada em %s" % caminho)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
