"""Corretor definitivo da linha dinamica TOTAL de itens_Remanesc.

Bug objetivo observado em XLS real de producao: a linha dinamica TOTAL (a
primeira linha vazia imediatamente apos o ultimo ITEM) so totalizava a coluna
D (VALOR_TOTAL). As demais colunas de VALOR financeiro — F, H, J, L, N, P, R e
AC — ficaram com a totalizacao DESATIVADA por um IF(FALSE,...) herdado do
aplicador de arredondamento itemizado, e a coluna T (VALOR_EXECUTADO_C4)
sequer possuia formula nas linhas 2:200.

REGRA PERMANENTE (registrada tambem na regra local de Masterfile/XLS):
a primeira linha vazia apos o ultimo ITEM em itens_Remanesc e a linha dinamica
TOTAL; TODAS as colunas de VALOR financeiro aplicaveis devem totalizar os
itens ate essa linha, com deteccao dinamica (1, 3, 20, 199 itens...). D nao
pode ser a unica coluna que totaliza. VU_ORIGINAL (C) NAO e totalizavel.

A correcao replica a MESMA logica que ja funciona em D:

  linha r (3..200):
    =IF(AND(A{r}="",A{r-1}<>"",COUNTIF(A{r+1}:$A$200,"<>")=0),
        <TOTAL>, <calculo normal do item>)

com <TOTAL> guardado por COUNT: coluna integralmente vazia/nao aplicavel
permanece vazia (nunca inventa 0,00 financeiro):

    IF(COUNT({col}$2:{col}{r-1})=0,"",
       ROUND(SUMIF($A$2:A{r-1},"<>",{col}$2:{col}{r-1}),2))

  linha 2: apenas o calculo normal (a linha 2 nunca e a linha TOTAL);
  linha 201: fallback fixo para a lotacao maxima (199 itens), tambem guardado
  por COUNT — mesmo padrao ja existente em D201/U201.

Gravador unico: Microsoft Excel real via COM (copia temporaria; promove so se
salvar sem erros de formula). FAIL-CLOSED: recusa reaplicacao.

REGRA ZERO CORRUPCAO XLSX: toda formula e ASCII puro e os parenteses sao
conferidos antes da gravacao.
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

FIM = 200          # ultima linha de itens; 201 e o fallback de lotacao maxima

# VALOR_REM_INICIO_Cn = ARRED(QTD_REM_AJUSTADA_Cn x VU_ATUALIZADO_Cn; 2)
# (col qtd em posicao_contratual, VU canonico em historico_VU, indice do ciclo)
REMANESCENTE = {
    "F": ("posicao_contratual!K", "historico_VU!D", 1),
    "H": ("posicao_contratual!O", "historico_VU!E", 2),
    "J": ("posicao_contratual!S", "historico_VU!F", 3),
    "L": ("posicao_contratual!W", "historico_VU!G", 4),
}

# VALOR_EXECUTADO_Cn = ARRED(QTD_EXECUTADA_Cn x VU_ATUALIZADO_Cn; 2)
# T (C4) estava sem formula no template: entra aqui com o mesmo padrao.
EXECUTADO = {
    "N": ("M", "historico_VU!D"),
    "P": ("O", "historico_VU!E"),
    "R": ("Q", "historico_VU!F"),
    "T": ("S", "historico_VU!G"),
    "AC": ("AB", "historico_VU!C"),
}

COLUNAS_FINANCEIRAS = tuple(REMANESCENTE) + tuple(EXECUTADO)

_PROT_FLAGS = (
    "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows",
    "AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks",
    "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting",
    "AllowFiltering", "AllowUsingPivotTables",
)


def _calculo_normal(col: str, linha: int) -> str:
    """Calculo do item na linha (sem wrapper de TOTAL) — identico ao vigente."""
    if col in REMANESCENTE:
        qtd, vu, indice = REMANESCENTE[col]
        return (
            f'IF(OR(A{linha}="",AND(ISNUMBER(posicao_contratual!$AL{linha}),'
            f'posicao_contratual!$AL{linha}>{indice}),{qtd}{linha}="",'
            f'NOT(ISNUMBER({vu}{linha}))),"",ROUND({qtd}{linha}*{vu}{linha},2))'
        )
    qtd, vu = EXECUTADO[col]
    return (
        f'IF(OR({qtd}{linha}="",NOT(ISNUMBER({vu}{linha}))),"",'
        f'ROUND({qtd}{linha}*{vu}{linha},2))'
    )


def _total_guardado(col: str, ate: int) -> str:
    """Soma dos itens com guarda de aplicabilidade (vazio nunca vira 0,00)."""
    return (
        f'IF(COUNT({col}$2:{col}{ate})=0,"",'
        f'ROUND(SUMIF($A$2:A{ate},"<>",{col}$2:{col}{ate}),2))'
    )


def _formula_linha2(col: str) -> str:
    return f"={_calculo_normal(col, 2)}"


def _formula_linha3(col: str) -> str:
    """Linha 3, preenchida ate a 200 com ajuste relativo — mesma logica de D3."""
    return (
        f'=IF(AND(A3="",A2<>"",COUNTIF(A4:$A${FIM},"<>")=0),'
        f"{_total_guardado(col, 2)},{_calculo_normal(col, 3)})"
    )


def _formula_fallback_201(col: str) -> str:
    """Lotacao maxima (199 itens): TOTAL fixo em 201, como D201/U201."""
    return (
        f'=IF($A${FIM}<>"",IF(COUNT({col}$2:{col}${FIM})=0,"",'
        f'ROUND(SUMIF($A$2:$A${FIM},"<>",{col}$2:{col}${FIM}),2)),"")'
    )


def _sem_protecao(ws):
    """Suspende a protecao da aba preservando as permissoes originais."""
    if not bool(ws.ProtectContents):
        return None
    p = ws.Protection
    estado = {}
    for f in _PROT_FLAGS:
        try:
            estado[f] = getattr(p, f)
        except Exception:
            estado[f] = True
    try:
        sel = ws.EnableSelection
    except Exception:
        sel = None
    ws.Unprotect()
    return estado, sel


def _restaurar_protecao(ws, estado) -> None:
    if estado is None:
        return
    flags, sel = estado
    ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True, **flags)
    if sel is not None:
        try:
            ws.EnableSelection = sel
        except Exception:
            pass


def _conferir_ascii_e_parenteses(*formulas: str) -> None:
    """REGRA ZERO CORRUPCAO: ASCII puro e parenteses balanceados."""
    for f in formulas:
        if not f.startswith("="):
            continue
        try:
            f.encode("ascii")
        except UnicodeEncodeError as exc:
            raise ValueError(f"Formula com caractere nao-ASCII: {f[:80]}") from exc
        if f.count("(") != f.count(")"):
            raise ValueError(f"Parenteses desbalanceados: {f[:80]}")


def _exigir_formulas(ws, faixas) -> None:
    """Fail-closed: prova que cada faixa gravada ficou como FORMULA, nao texto."""
    for faixa in faixas:
        alvo = ws.Range(faixa)
        if not bool(alvo.Cells(1, 1).HasFormula):
            raise RuntimeError(
                f"{ws.Name}!{faixa} nao ficou como formula (formato Texto?)."
            )


def _validar_layout(ws) -> None:
    """Confere o pre-estado esperado e recusa reaplicacao (fail-closed)."""
    d3 = str(ws.Range("D3").Formula or "")
    if 'IF(AND(A3="",A2<>"",COUNTIF(A4:$A$200,"<>")=0)' not in d3:
        raise ValueError(
            "itens_Remanesc!D3 nao contem a deteccao dinamica da linha TOTAL; "
            "template fora da linhagem esperada."
        )
    u3 = str(ws.Range("U3").Formula or "")
    if '"TOTAL"' not in u3:
        raise ValueError("itens_Remanesc!U3 nao contem o rotulo dinamico TOTAL.")
    f3 = str(ws.Range("F3").Formula or "")
    if 'AND(A3="",A2<>""' in f3:
        raise ValueError(
            "itens_Remanesc!F3 ja contem a deteccao dinamica; pacote ja aplicado?"
        )
    if "IF(FALSE" not in f3:
        raise ValueError(
            "itens_Remanesc!F3 sem o IF(FALSE,...) esperado do pre-estado; "
            "conferir a linhagem do template antes de aplicar."
        )


def _aplicar_itens_remanesc(ws) -> None:
    estado = _sem_protecao(ws)
    for col in COLUNAS_FINANCEIRAS:
        f2 = _formula_linha2(col)
        f3 = _formula_linha3(col)
        f201 = _formula_fallback_201(col)
        _conferir_ascii_e_parenteses(f2, f3, f201)
        ws.Range(f"{col}2").Formula = f2
        # Gravar a formula da linha 3 na faixa inteira: o Excel ajusta as
        # referencias relativas linha a linha, exatamente como a coluna D.
        ws.Range(f"{col}3:{col}{FIM}").Formula = f3
        ws.Range(f"{col}{FIM + 1}").Formula = f201
    _exigir_formulas(
        ws,
        [f"{col}2:{col}{FIM + 1}" for col in COLUNAS_FINANCEIRAS],
    )
    _restaurar_protecao(ws, estado)


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


def aplicar(origem: Path, destino: Path) -> None:
    origem, destino = Path(origem), Path(destino)
    if not origem.is_file():
        raise FileNotFoundError(origem)

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_total_remanesc_"))
    tmp_xlsx = tmp_dir / origem.name
    shutil.copyfile(origem, tmp_xlsx)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = True
    wb = None
    salvo = False
    try:
        wb = excel.Workbooks.Open(str(tmp_xlsx), UpdateLinks=0)
        excel.ScreenUpdating = False
        excel.Calculation = XL_CALC_MANUAL
        aba_ativa = wb.ActiveSheet.Name
        ws = wb.Worksheets("itens_Remanesc")
        _validar_layout(ws)
        _aplicar_itens_remanesc(ws)
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
    shutil.copyfile(tmp_xlsx, destino)
    shutil.rmtree(tmp_dir, ignore_errors=True)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("origem", type=Path)
    parser.add_argument("destino", type=Path)
    args = parser.parse_args()
    aplicar(args.origem, args.destino)


if __name__ == "__main__":
    main()
