# -*- coding: utf-8 -*-
"""Hotfix RESULTADOS (ETAPA 26B) via Excel COM real — NUNCA openpyxl para salvar.

Corrige a migracao meia-feita do metodo (CONTROLE!B1 com rotulos amigaveis
"Principal (Financeiro)" / "PC (Pedidos de Compra)" / "Itens Consumidos") que
deixava a aba RESULTADOS quebrada sob PC:

  - B4 vira o TOKEN CANONICO NORMALIZADO (Financeiro / PCs / Itens), robusto a
    rotulos antigos ("Principal","Pedidos de Compras") e novos (SEARCH
    case-insensitive). Fonte unica: tudo referencia $B$4.
  - H4 passa a ler $B$4 (nao mais CONTROLE!B1 com rotulo antigo).
  - Eixo remanescente (B35/C35/D35/F35/F36): adiciona ramo PC. PC usa a MESMA
    posicao fisica do Financeiro (PC nao e fonte de quantidade fisica). D35 sob
    PC usa o remanescente item-a-item do bloco VTA-PC ($T$23), que espelha o
    motor Python, para bater o golden 854564.71.
  - B26 (VTA final): ramo PC passa a testar $B$4="PCs" (antes o rotulo literal).
  - Memoria O/P/Q (linhas 54..253): substitui rotulos antigos por $B$4 e faz PC
    valorar pela coluna fisica K (como Financeiro), sem status "MANUAL".

  Layout executivo: esconde colunas auxiliares G:O e X:Y; agrupa (outline) a
  memoria item-a-item (54..253) e o bloco tecnico Eixo 5 (256..263). Formulas
  preservadas (nada e apagado). Resumo executivo (B4/B16/E16/B26/B35/D35/J4)
  permanece no topo.

Idempotente. Preserva formatacao condicional, validacoes, estilos e demais abas.

Uso:
    python tools/aplicar_hotfix_resultados_clemar.py <origem.xlsx> <destino.xlsx>
"""
from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105
XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16
XL_UP = -4162  # xlUp

# --- B4 token canonico (case-insensitive; cobre rotulos antigos e novos) -----
FORMULA_B4 = (
    '=IF(OR(CONTROLE!$B$1="SELECIONE AQUI",CONTROLE!$B$1=""),'
    '"SELECIONE O METODO EM CONTROLE!B1",'
    'IF(OR(ISNUMBER(SEARCH("Financeiro",CONTROLE!$B$1)),'
    'ISNUMBER(SEARCH("Principal",CONTROLE!$B$1))),"Financeiro",'
    'IF(OR(ISNUMBER(SEARCH("Pedido",CONTROLE!$B$1)),'
    'ISNUMBER(SEARCH("PC",CONTROLE!$B$1))),"PCs",'
    'IF(ISNUMBER(SEARCH("Itens",CONTROLE!$B$1)),"Itens",'
    '"SELECIONE O METODO EM CONTROLE!B1"))))'
)

FORMULA_H4 = '=IF($B$4="Itens","Itens",IF($B$4="PCs","PCs","Informado"))'

# --- Eixo remanescente: ramo PC = mesma posicao fisica do Financeiro ---------
FORMULA_B35 = '=IF(OR($B$4="Financeiro",$B$4="PCs"),B32,IF($B$4="Itens",B33,""))'
FORMULA_C35 = '=IF(OR($B$4="Financeiro",$B$4="PCs"),C32,IF($B$4="Itens",C33,""))'
# D35 sob PC usa T23 (remanescente item-a-item do bloco VTA-PC, espelha o motor)
FORMULA_D35 = (
    '=IF($B$4="PCs",$T$23,'
    'IF($B$4="Financeiro",D32,IF($B$4="Itens",D33,"")))'
)
FORMULA_F35 = (
    '=IF($B$4="Itens","ITENS",'
    'IF(OR($B$4="Financeiro",$B$4="PCs"),"INFORMADO",""))'
)
# F36: editado CIRURGICAMENTE em runtime a partir da formula original lida via
# COM (preserva acentos/travessoes exatos). Ver _fix_f36().  A mudanca:
#   1) remove o clause inicial IF(CONTROLE!$B$1="Pedidos de Compras","...",<resto>)
#      -> PC deixa de ser forcado a "CALCULO MANUAL"; segue as checagens do resto.
#   2) troca os testes de metodo remanescentes por $B$4 e adiciona ramo PC:
#         CONTROLE!$B$1="Principal"        -> OR($B$4="Financeiro",$B$4="PCs")
#         CONTROLE!$B$1="Itens Consumidos" -> $B$4="Itens"

# --- B26 (VTA final): ramo PC testa $B$4="PCs" ---
# Editado cirurgicamente: troca so o literal CONTROLE!$B$1="PC (Pedidos de
# Compra)" (template novo) OU CONTROLE!$B$1="Pedidos de Compras" (legado) por
# $B$4="PCs". Ver _fix_b26().  Fallback: se nenhum literal for encontrado (ex.:
# arquivo antigo sem ramo PC), aplica a formula canonica completa abaixo.
FORMULA_B26_FALLBACK = (
    '=IF(AND(B24<>"",B25<>""),"",IF(ISNUMBER(B25),B25,'
    'IF($B$4="PCs",'
    'IF($T$25="CALCULO MANUAL REQUERIDO","",'
    'ROUND($T$25+IF(ISNUMBER(B24),B24,0),2)),'
    'IF(OR(B23="",AND(B24<>"",NOT(ISNUMBER(B24)))),"",'
    'ROUND(B23+IF(ISNUMBER(B24),B24,0)+IF(ISNUMBER($N$263),$N$263,0),2)))))'
)


def _fix_b26(orig: str) -> str:
    """Troca o literal do ramo PC por $B$4="PCs"; senao usa o fallback canonico."""
    for lit in ('CONTROLE!$B$1="PC (Pedidos de Compra)"',
                'CONTROLE!$B$1="Pedidos de Compras"',
                'CONTROLE!$B$1="PC (Pedidos de Compra)s"'):
        if lit in orig:
            return orig.replace(lit, '$B$4="PCs"')
    return FORMULA_B26_FALLBACK


def _fix_f36(orig: str) -> str:
    """Remove o clause inicial que forca PC->MANUAL e migra testes de metodo p/ $B$4.

    orig (template atual):
      =IF(CONTROLE!$B$1="Pedidos de Compras","CALCULO MANUAL REQUERIDO",<RESTO>)
    Retorna <RESTO> (agora como formula, com "=" na frente) apos migrar os
    literais de metodo internos. Se o clause inicial nao existir, so migra os
    literais (idempotente).
    """
    f = orig
    # 1) remover o wrapper "=IF(CONTROLE!$B$1=\"Pedidos de Compras\",\"...\"," ... final ")"
    head = '=IF(CONTROLE!$B$1="Pedidos de Compras",'
    if f.startswith(head):
        # apos o head vem  "<msg manual>", depois virgula, depois o RESTO ...  )
        rest = f[len(head):]
        # pula o primeiro literal string "..." (a msg de calculo manual)
        assert rest[0] == '"'
        end = rest.index('"', 1)
        rest = rest[end + 1:]
        assert rest[0] == ','
        rest = rest[1:]           # remove a virgula separadora
        # rest termina com ")" extra do IF externo -> remove o ultimo ")"
        assert rest.endswith(')')
        rest = rest[:-1]
        f = '=' + rest
    # 2) migrar literais de metodo internos para $B$4 (+ ramo PC no Financeiro)
    f = f.replace('AND(CONTROLE!$B$1="Principal",$H$44>0)',
                  'AND(OR($B$4="Financeiro",$B$4="PCs"),$H$44>0)')
    f = f.replace('AND(CONTROLE!$B$1="Itens Consumidos",$F$44>0)',
                  'AND($B$4="Itens",$F$44>0)')
    return f


def _memoria_O(r: int) -> str:
    """O{r}: valor sem reajuste. PC valora pela coluna fisica K (como Financeiro)."""
    return (
        f'=IF(J{r}="","",IF(OR($B$4="Financeiro",$B$4="PCs"),'
        f'IF(OR(K{r}="",M{r}=""),"",ROUND(K{r}*M{r},2)),'
        f'IF($B$4="Itens",IF(OR(L{r}="",M{r}=""),"",ROUND(L{r}*M{r},2)),"")))'
    )


def _memoria_P(r: int) -> str:
    """P{r}: valor com reajuste. PC valora pela coluna fisica K (como Financeiro)."""
    return (
        f'=IF(J{r}="","",IF(OR($B$4="Financeiro",$B$4="PCs"),'
        f'IF(OR(K{r}="",N{r}=""),"",ROUND(K{r}*N{r},2)),'
        f'IF($B$4="Itens",IF(OR(L{r}="",N{r}=""),"",ROUND(L{r}*N{r},2)),"")))'
    )


def _memoria_Q(r: int) -> str:
    """Q{r}: status. PC segue a checagem fisica do Financeiro (K), nunca MANUAL."""
    return (
        f'=IF(J{r}="","",IF(OR(M{r}="",AND(OR($B$4="Financeiro",$B$4="PCs"),K{r}=""),'
        f'AND($B$4="Itens",L{r}="")),"INCOMPLETO",'
        f'IF(AND(ISNUMBER(K{r}),ISNUMBER(L{r}),'
        f'ABS(K{r}-L{r})/MAX(ABS(IF(OR($B$4="Financeiro",$B$4="PCs"),K{r},L{r})),1)>$D$4),'
        f'"DIVERGENTE","OK")))'
    )


MEM_INI, MEM_FIM = 54, 253       # memoria item-a-item O/P/Q
EIXO5_INI, EIXO5_FIM = 256, 263  # bloco tecnico Eixo 5


def _patch_resultados(wb) -> None:
    r = wb.Worksheets("RESULTADOS")

    r.Range("B4").Formula = FORMULA_B4
    r.Range("H4").Formula = FORMULA_H4
    # B26 e F36: edicao cirurgica a partir da formula original (preserva acentos).
    r.Range("B26").Formula = _fix_b26(r.Range("B26").Formula)
    r.Range("F36").Formula = _fix_f36(r.Range("F36").Formula)
    r.Range("B35").Formula = FORMULA_B35
    r.Range("C35").Formula = FORMULA_C35
    r.Range("D35").Formula = FORMULA_D35
    r.Range("F35").Formula = FORMULA_F35

    for row in range(MEM_INI, MEM_FIM + 1):
        r.Range(f"O{row}").Formula = _memoria_O(row)
        r.Range(f"P{row}").Formula = _memoria_P(row)
        r.Range(f"Q{row}").Formula = _memoria_Q(row)


def _layout(wb) -> None:
    r = wb.Worksheets("RESULTADOS")
    # Esconde colunas auxiliares (formulas preservadas).
    for col in ("G:O", "X:Y"):
        r.Columns(col).Hidden = True
    # Agrupa (outline) memoria item-a-item e bloco Eixo 5; colapsa.
    r.Outline.SummaryRow = 0  # xlSummaryAbove (resumo no topo)
    try:
        r.Rows(f"{MEM_INI}:{MEM_FIM}").Group()
        r.Rows(f"{EIXO5_INI}:{EIXO5_FIM}").Group()
    except Exception:
        pass
    try:
        r.Outline.ShowLevels(RowLevels=1)
    except Exception:
        pass


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
    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_hotfix_res_"))
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
        _patch_resultados(wb)
        _layout(wb)
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
    print("HOTFIX RESULTADOS aplicado:", args.destino)


if __name__ == "__main__":
    main()
