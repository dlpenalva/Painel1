"""Aplica a Etapa 1.1 (Execucao Historica Fora da Analise) ao template oficial.

Alteracao ESTRITAMENTE ADITIVA e cirurgica sobre o template homologado:

  1. itens_Remanesc: novas colunas AO:BC (15) com a reconciliacao por
     item x ciclo -- QTD_COBERTA / QTD_NAO_COBERTA / DIFERENCIAL para C0..C4.
     Reutiliza o motor de execucao existente (AB/M/O/Q/S) e a fonte
     quantitativa itens_Consumidos. NAO cria motor paralelo.

  2. RESULTADOS: novo bloco "EIXO 5 - EXECUCAO HISTORICA FORA DA ANALISE"
     a partir da linha 256 (regiao livre), com uma linha por ciclo C0..C4,
     consolidando os auxiliares e classificando cobertura/status. O campo
     VALOR ADOTADO permanece VAZIO (decisao humana) e NAO alimenta B23/B26.

Fator incremental: reutiliza a celula canonica do retroativo (RESULTADOS!D11),
  diferencial por unidade = VU * (parametros!F_n - parametros!F_n/parametros!D_n).

Gravador unico: Microsoft Excel real via COM (DisplayAlerts=True). Trabalha
sobre copia temporaria e so promove ao destino se o proprio Excel salvar sem
erros de formula. NAO altera B20:B26, VTA, aditivos, template legado, Python
de producao nem paginas.
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

XL_PASTE_FORMATS = -4122
XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16
XL_VALIDATE_DECIMAL = 2
XL_VALID_ALERT_STOP = 1
XL_GREATER_EQUAL = 7
XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105

FIM = 200  # ultima linha de item em itens_Remanesc (201 e a linha TOTAL)

# ciclo, exec_col(itens_Remanesc), cons_col(itens_Consumidos), cob, nc, dif,
# F(parametros fator acumulado), D(parametros fator apuracao), status(F11..F15)
CICLOS = [
    ("C0", "AB", "E", "AO", "AP", "AQ", "$F$2", "$D$11", "$F$11"),
    ("C1", "M",  "G", "AR", "AS", "AT", "$F$3", "$D$12", "$F$12"),
    ("C2", "O",  "I", "AU", "AV", "AW", "$F$4", "$D$13", "$F$13"),
    ("C3", "Q",  "K", "AX", "AY", "AZ", "$F$5", "$D$14", "$F$14"),
    ("C4", "S",  "M", "BA", "BB", "BC", "$F$6", "$D$15", "$F$15"),
]

CAB_REMANESC = [
    "QTD_COBERTA_C0", "QTD_NAO_COBERTA_C0", "DIFERENCIAL_C0",
    "QTD_COBERTA_C1", "QTD_NAO_COBERTA_C1", "DIFERENCIAL_C1",
    "QTD_COBERTA_C2", "QTD_NAO_COBERTA_C2", "DIFERENCIAL_C2",
    "QTD_COBERTA_C3", "QTD_NAO_COBERTA_C3", "DIFERENCIAL_C3",
    "QTD_COBERTA_C4", "QTD_NAO_COBERTA_C4", "DIFERENCIAL_C4",
]

# RESULTADOS
LIN_TITULO = 256
LIN_HDR = 257
LIN_C0 = 258
LIN_NOTA = 263
CAB_BLOCO = [
    "CICLO", "FORA DA ANALISE?", "FONTE PRIMARIA", "COBERTURA DA FONTE",
    "QTD EXECUTADA RECONSTRUIDA", "QTD COBERTA PELA FONTE", "QTD NAO COBERTA",
    "FATOR INCREMENTAL DA APURACAO", "DIFERENCIAL POTENCIAL CALCULADO",
    "VALOR ADOTADO", "STATUS / CHECK",
]
TITULO = ("EIXO 5 - EXECUCAO HISTORICA FORA DA ANALISE | apoio a decisao; "
          "NAO altera o VTA (B23/B26) automaticamente")
NOTA = ("O valor calculado e apoio; o VALOR ADOTADO (coluna J) e decisao humana "
        "explicita, aceita zero, rejeita negativo e NAO alimenta B23/B26 nesta etapa.")


def _validar_layout(wb, excel) -> None:
    wf = excel.WorksheetFunction
    ws_i = wb.Worksheets("itens_Remanesc")
    if wf.CountA(ws_i.Range(f"AO1:BC{FIM + 1}")) != 0:
        raise ValueError("itens_Remanesc AO1:BC201 nao esta vazia; bloco ja aplicado?")
    ws_r = wb.Worksheets("RESULTADOS")
    if wf.CountA(ws_r.Range(f"A{LIN_TITULO}:K{LIN_NOTA + 2}")) != 0:
        raise ValueError("RESULTADOS A256:K265 nao esta vazia; bloco ja aplicado?")
    if ws_i.Range("AB1").Value is None or "EXEC" not in str(ws_i.Range("AA1").Value):
        raise ValueError("Ancora de execucao C0 (AA1/AB1) ausente em itens_Remanesc.")


def _aplicar_remanesc(ws, excel) -> None:
    ws.Range("Z1").Copy()
    ws.Range("AO1:BC1").PasteSpecial(XL_PASTE_FORMATS)
    excel.CutCopyMode = False
    ws.Range("AO1:BC1").Value = [CAB_REMANESC]
    for _c, exe, cons, cob, nc, dif, F, D, _s in CICLOS:
        ws.Range(f"{cob}2:{cob}{FIM}").Formula = (
            f'=IF($A2="","",ROUND(SUMIFS(itens_Consumidos!${cons}$2:${cons}$200,'
            f'itens_Consumidos!$A$2:$A$200,$A2),2))'
        )
        ws.Range(f"{nc}2:{nc}{FIM}").Formula = (
            f'=IF(OR($A2="",{exe}2=""),"",ROUND(MAX({exe}2-{cob}2,0),2))'
        )
        ws.Range(f"{dif}2:{dif}{FIM}").Formula = (
            f'=IF(OR($A2="",{nc}2="",NOT(ISNUMBER(parametros!{F})),'
            f'NOT(ISNUMBER(parametros!{D})),parametros!{D}=0),"",'
            f'ROUND({nc}2*$C2*(parametros!{F}-parametros!{F}/parametros!{D}),2))'
        )
    ws.Columns("AO:BC").ColumnWidth = 16
    ws.Columns("AO:BC").Hidden = True  # tecnicas; reexibiveis, nunca veryHidden


def _aplicar_bloco(ws, excel) -> None:
    ws.Range(f"A{LIN_TITULO}").Value = TITULO
    ws.Range("A19").Copy()
    ws.Range(f"A{LIN_TITULO}").PasteSpecial(XL_PASTE_FORMATS)
    ws.Range("A9:K9").Copy()
    ws.Range(f"A{LIN_HDR}:K{LIN_HDR}").PasteSpecial(XL_PASTE_FORMATS)
    excel.CutCopyMode = False
    ws.Range(f"A{LIN_HDR}:K{LIN_HDR}").Value = [CAB_BLOCO]

    # Reconciliavel <=> metodo quantitativo canonico "Itens" (dropdown
    # Financeiro,PCs,Itens). Verificacao POSITIVA: vazio/PCs/Financeiro/texto
    # desconhecido nunca sao tratados como quantitativos.
    nao_recon = '$B$4<>"Itens"'
    for i, (nome, exe, cons, cob, nc, dif, F, D, statusF) in enumerate(CICLOS):
        r = LIN_C0 + i
        ws.Range(f"A{r}").Value = nome
        ws.Range(f"B{r}").Formula = f'=IF(parametros!{statusF}="Aplicado","Nao","Sim")'
        ws.Range(f"C{r}").Formula = "=$B$4"
        ws.Range(f"E{r}").Formula = (
            f'=IF(COUNT(itens_Remanesc!{exe}$2:{exe}$200)=0,"",'
            f'ROUND(SUM(itens_Remanesc!{exe}$2:{exe}$200),2))'
        )
        # F = quantidade coberta BRUTA pela fonte (mesma populacao do retroativo)
        ws.Range(f"F{r}").Formula = (
            f'=IF(OR(E{r}="",{nao_recon}),"n/d",'
            f'ROUND(SUM(itens_Remanesc!{cob}$2:{cob}$200),2))'
        )
        ws.Range(f"G{r}").Formula = (
            f'=IF(OR(E{r}="",{nao_recon}),"n/d",'
            f'ROUND(SUM(itens_Remanesc!{nc}$2:{nc}$200),2))'
        )
        ws.Range(f"H{r}").Formula = (
            f'=IF(OR(NOT(ISNUMBER(parametros!{F})),NOT(ISNUMBER(parametros!{D})),'
            f'parametros!{D}=0),"",parametros!{F}-parametros!{F}/parametros!{D})'
        )
        ws.Range(f"I{r}").Formula = (
            f'=IF(OR(E{r}="",{nao_recon}),"",'
            f'ROUND(SUM(itens_Remanesc!{dif}$2:{dif}$200),2))'
        )
        ws.Range(f"D{r}").Formula = (
            f'=IF({nao_recon},"NAO RECONCILIAVEL",IF(E{r}="","SEM EXECUCAO",'
            f'IF(AND(ISNUMBER(F{r}),F{r}>E{r}),"COBERTURA SUPERIOR",'
            f'IF(G{r}=0,"COMPLETA",IF(G{r}=E{r},"AUSENTE","PARCIAL")))))'
        )
        ws.Range(f"K{r}").Formula = (
            f'=IF(AND(J{r}<>"",NOT(ISNUMBER(J{r}))),"VALOR ADOTADO INVALIDO",'
            f'IF(AND(ISNUMBER(J{r}),J{r}<0),"NEGATIVO - CORRIGIR",'
            f'IF(ISNUMBER(J{r}),"ADOTADO MANUAL - nao altera VTA",'
            f'IF({nao_recon},"MANUAL - fonte nao reconciliavel",'
            f'IF(E{r}="","sem execucao reconstruida",'
            f'IF(D{r}="COBERTURA SUPERIOR","COBERTURA SUPERIOR A EXECUCAO - conferir",'
            f'IF(B{r}="Nao","DENTRO DA ANALISE - diferencial ja no retroativo (nao adotar)",'
            f'IF(D{r}="COMPLETA","coberto - sem diferencial",'
            f'"APOIO - diferencial potencial; adocao manual"))))))))'
        )

    ws.Range(f"A{LIN_NOTA}").Value = NOTA
    ws.Range(f"A{LIN_NOTA}").WrapText = False
    ws.Range(f"E{LIN_C0}:I{LIN_C0 + 4}").NumberFormatLocal = "#.##0,00;-#.##0,00"
    ws.Range(f"J{LIN_C0}:J{LIN_C0 + 4}").NumberFormatLocal = "#.##0,00;-#.##0,00"
    val = ws.Range(f"J{LIN_C0}:J{LIN_C0 + 4}").Validation
    try:
        val.Delete()
    except Exception:
        pass
    val.Add(Type=XL_VALIDATE_DECIMAL, AlertStyle=XL_VALID_ALERT_STOP,
            Operator=XL_GREATER_EQUAL, Formula1="0")
    val.IgnoreBlank = True
    val.ErrorMessage = "VALOR ADOTADO deve ser >= 0 (zero permitido)."


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
    origem = Path(origem)
    destino = Path(destino)
    if not origem.is_file():
        raise FileNotFoundError(origem)

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_etapa11_"))
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
        excel.Calculation = XL_CALC_MANUAL  # evita recalculo entre cada escrita
        _validar_layout(wb, excel)
        _aplicar_remanesc(wb.Worksheets("itens_Remanesc"), excel)
        _aplicar_bloco(wb.Worksheets("RESULTADOS"), excel)

        excel.Calculation = XL_CALC_AUTOMATIC
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
