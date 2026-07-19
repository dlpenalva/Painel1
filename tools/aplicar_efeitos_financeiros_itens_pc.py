"""Aplica, de forma reproduzivel, a Etapa 3 ao template oficial da Coleta.

Gravador unico: Microsoft Excel real via COM (DisplayAlerts=True).
Este script nao usa openpyxl, LibreOffice nem manipulacao de ZIP/XML para
gravar. Trabalha sobre copia temporaria e somente promove o resultado ao
destino depois de o proprio Excel salvar e fechar sem erros. Qualquer falha
interrompe o fluxo sem salvar e sem tocar o destino.
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

LIMITE_PC = 100

XL_PASTE_FORMATS = -4122
XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16
XL_EXPRESSION = 2

COR_AMARELO = 153 * 65536 + 230 * 256 + 255  # RGB(255,230,153)
COR_VERMELHO = 204 * 65536 + 204 * 256 + 244  # RGB(244,204,204)

CABECALHO_ITENS_PC = [
    "NUMERO_PC", "DATA_PC", "CICLO_PC", "VALOR_PC", "FATOR_ACUMULADO",
    "VALOR_ATUALIZADO", "PC_PAGO_A_CONTRATADA",
    "RETROATIVO_RECONHECIDO_A_PAGAR", "VALOR_ATUALIZADO_EM_ANALISE",
    "DELTA_POTENCIAL", "CHECK_PC_FINANCEIRO", None,
]
CABECALHO_PARAMETROS = [
    "COMPUTAR_NESTA_APURACAO", "CICLO", "DATA_INICIO", "DATA_FIM",
    "PERCENTUAL_DO_CICLO", "FATOR_ACUMULADO", "SITUACAO",
]


def _formula_l(linha: int) -> str:
    r = linha
    inicio = (
        f'IFERROR(INDEX(parametros!$H$2:$H$6,'
        f'MATCH(C{r},parametros!$B$2:$B$6,0)),"")'
    )
    computar = (
        f'IFERROR(INDEX(parametros!$A$2:$A$6,'
        f'MATCH(C{r},parametros!$B$2:$B$6,0)),"")'
    )
    return (
        f'=IF(AND(A{r}="",B{r}="",D{r}="",G{r}=""),"",'
        f'IF(OR(B{r}="",NOT(ISNUMBER(B{r})),C{r}="",C{r}="Fora dos ciclos"),"",'
        f'IF(C{r}="C0","Nao",IF({computar}<>"Sim","Nao",'
        f'IF({inicio}="","",IF(B{r}>={inicio},"Sim","Nao"))))))'
    )


def _formula_k(linha: int) -> str:
    r = linha
    inicio = (
        f'IFERROR(INDEX(parametros!$H$2:$H$6,'
        f'MATCH(C{r},parametros!$B$2:$B$6,0)),"")'
    )
    computar = (
        f'IFERROR(INDEX(parametros!$A$2:$A$6,'
        f'MATCH(C{r},parametros!$B$2:$B$6,0)),"")'
    )
    return (
        f'=IF(AND(A{r}="",B{r}="",D{r}="",G{r}=""),"",'
        f'IF(AND(A{r}<>"",SUMPRODUCT(--(UPPER(TRIM($A$2:$A$100))='
        f'UPPER(TRIM(A{r}))))>1),"NUMERO_PC duplicado",'
        f'IF(B{r}="","DATA_PC vazia",IF(NOT(ISNUMBER(B{r})),"DATA_PC invalida",'
        f'IF(OR(C{r}="",C{r}="Fora dos ciclos"),"CICLO_PC nao identificado",'
        f'IF(OR(D{r}="",D{r}=0),"VALOR_PC vazio ou zero",'
        f'IF(AND(G{r}<>"Sim",G{r}<>"Nao"),"PC_PAGO_A_CONTRATADA invalido",'
        f'IF(AND(C{r}<>"C0",{computar}="Sim",{inicio}=""),'
        f'"INICIO_EFEITO ausente: PC "&A{r}&" - "&C{r},'
        f'IF(L{r}="","EFEITO_PC nao calculado: PC "&A{r}&" - "&C{r},"OK")))))))))'
    )


def _formula_e(linha: int) -> str:
    r = linha
    return (
        f'=IF(C{r}="","",IF(L{r}="","",IF(L{r}="Nao",1,'
        f'IFERROR(VLOOKUP(C{r},parametros!$A$11:$E$15,5,0),""))))'
    )


def _formula_h(linha: int) -> str:
    r = linha
    return (
        f'=IF(G{r}="Sim",IF(OR(F{r}="",D{r}="",L{r}=""),"",'
        f'IF(L{r}="Sim",ROUND(F{r}-D{r},2),0)),0)'
    )


def _formula_i(linha: int) -> str:
    r = linha
    return f'=IF(G{r}="Nao",IF(OR(F{r}="",L{r}=""),"",F{r}),0)'


def _formula_j(linha: int) -> str:
    r = linha
    return (
        f'=IF(G{r}="Nao",IF(OR(F{r}="",D{r}="",L{r}=""),"",'
        f'IF(L{r}="Sim",ROUND(F{r}-D{r},2),0)),0)'
    )


def _validar_layout(ws_itens, ws_par) -> None:
    encontrados = [ws_itens.Cells(1, c).Value for c in range(1, 13)]
    if encontrados != CABECALHO_ITENS_PC:
        raise ValueError(f"Layout A:L inesperado em itens_PC: {encontrados!r}")
    encontrados = [ws_par.Cells(1, c).Value for c in range(1, 8)]
    if encontrados != CABECALHO_PARAMETROS:
        raise ValueError(f"Layout A:G inesperado em parametros: {encontrados!r}")


def _aplicar_parametros(ws_par, excel) -> None:
    ws_par.Range("G1").Copy()
    ws_par.Range("H1").PasteSpecial(XL_PASTE_FORMATS)
    ws_par.Range("D2:D6").Copy()
    ws_par.Range("H2:H6").PasteSpecial(XL_PASTE_FORMATS)
    excel.CutCopyMode = False
    ws_par.Range("H1").Value = "INICIO_EFEITO_FINANCEIRO"
    # Em Excel pt-BR, NumberFormat="dd/mm/yyyy" grava "y" literal (defeito
    # comprovado); o caminho local com ";@" armazena o codigo customizado
    # deterministico dd/mm/yyyy;@ no OOXML.
    try:
        ws_par.Range("H2:H6").NumberFormatLocal = "dd/mm/aaaa;@"
    except Exception:
        ws_par.Range("H2:H6").NumberFormat = "dd/mm/yyyy;@"
    ws_par.Range("H2").Value2 = 45400  # 18/04/2024, prova de renderizacao
    renderizado = ws_par.Range("H2").Text
    ws_par.Range("H2").ClearContents()
    if renderizado != "18/04/2024":
        raise RuntimeError(
            f"Formato de data nao aplicado em parametros!H2:H6: "
            f"data serial 45400 renderizou {renderizado!r}"
        )
    ws_par.Columns("H").ColumnWidth = 25


def _aplicar_itens_pc(ws, excel) -> None:
    ws.Range("K1").Copy()
    ws.Range("L1").PasteSpecial(XL_PASTE_FORMATS)
    ws.Range(f"K2:K{LIMITE_PC}").Copy()
    ws.Range(f"L2:L{LIMITE_PC}").PasteSpecial(XL_PASTE_FORMATS)
    excel.CutCopyMode = False
    ws.Range("L1").Value = "EFEITO_FINANCEIRO_PC"
    ws.Columns("L").ColumnWidth = 24

    linhas = range(2, LIMITE_PC + 1)
    for coluna, fabrica in (
        ("E", _formula_e), ("H", _formula_h), ("I", _formula_i),
        ("J", _formula_j), ("K", _formula_k), ("L", _formula_l),
    ):
        destino = ws.Range(f"{coluna}2:{coluna}{LIMITE_PC}")
        destino.Formula = [[fabrica(r)] for r in linhas]

    # A primeira linha alem do limite fica fora da grade de entrada.
    ultima = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    if ultima > LIMITE_PC:
        alem = ws.Range(f"A{LIMITE_PC + 1}:L{ultima}")
        if excel.WorksheetFunction.CountA(alem) != 0:
            raise RuntimeError(
                f"itens_PC possui conteudo inesperado em A{LIMITE_PC + 1}:L{ultima}."
            )
        alem.ClearFormats()


def _para_formula_local(ws, formula_en: str) -> str:
    """Converte sintaxe en-US para a sintaxe local exigida por FormatConditions."""
    apoio = ws.Range("AZ2")
    apoio.Formula = formula_en
    local = apoio.FormulaLocal
    apoio.ClearContents()
    return local


def _aplicar_formatacao_condicional(ws) -> None:
    grade = ws.Range(f"A2:L{LIMITE_PC}")
    grade.FormatConditions.Delete()
    # Referencias relativas de FormatConditions ancoram na celula ativa.
    ws.Activate()
    ws.Range("A2").Select()
    amarela = grade.FormatConditions.Add(
        Type=XL_EXPRESSION,
        Formula1=_para_formula_local(
            ws, '=AND(OR($A2<>"",$B2<>"",$D2<>"",$G2<>""),$L2="")'
        ),
    )
    amarela.Interior.Color = COR_AMARELO
    amarela.StopIfTrue = True
    vermelha = grade.FormatConditions.Add(
        Type=XL_EXPRESSION, Formula1=_para_formula_local(ws, '=$L2="Nao"')
    )
    vermelha.Interior.Color = COR_VERMELHO
    vermelha.StopIfTrue = False


def _verificar_sem_erros(wb) -> None:
    import pywintypes

    problemas = []
    for ws in wb.Worksheets:
        for tipo in (XL_CELLTYPE_FORMULAS, XL_CELLTYPE_CONSTANTS):
            try:
                celulas = ws.UsedRange.SpecialCells(tipo, XL_ERRORS)
                problemas.append(f"{ws.Name}!{celulas.Address}")
            except pywintypes.com_error:
                continue  # nenhuma celula de erro deste tipo
    if problemas:
        raise RuntimeError(f"Erros de formula apos recalculo: {problemas}")


def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem)
    destino = Path(destino)
    if not origem.is_file():
        raise FileNotFoundError(origem)

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_etapa3_"))
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
        ws_itens = wb.Worksheets("itens_PC")
        ws_par = wb.Worksheets("parametros")

        _validar_layout(ws_itens, ws_par)
        _aplicar_parametros(ws_par, excel)
        _aplicar_itens_pc(ws_itens, excel)
        _aplicar_formatacao_condicional(ws_itens)

        ws_itens.Activate()
        ws_itens.Range("A1").Select()
        ws_par.Activate()
        ws_par.Range("A1").Select()
        wb.Worksheets(1).Activate()

        excel.CalculateFullRebuild()
        _verificar_sem_erros(wb)

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
