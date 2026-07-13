from __future__ import annotations

import argparse
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation


INPUT = PatternFill("solid", fgColor="FFFFF2CC")
AUTO = PatternFill("solid", fgColor="FFEDEDED")
HEADER = PatternFill("solid", fgColor="FF1F4E79")
CHECK = PatternFill("solid", fgColor="FFE2EFDA")
WHITE = "FFFFFFFF"
NAVY = "FF1F4E79"
GRAY = "FF595959"
MONEY = 'R$ #,##0.00'
QTY = '#,##0.00'
PERCENT = '0.00%'
FACTOR = '#,##0.0000'
THIN = Side(style="thin", color="FFD9D9D9")
GRID = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _style_formula_range(ws, row_start: int, row_end: int, formulas: dict[str, str]) -> None:
    for row in range(row_start, row_end + 1):
        for col, pattern in formulas.items():
            ws[f"{col}{row}"] = pattern.format(r=row)
            ws[f"{col}{row}"].fill = AUTO
            ws[f"{col}{row}"].font = Font(name="Calibri", size=10, color=GRAY)
            ws[f"{col}{row}"].border = GRID


def _reset_controle(wb) -> None:
    ws = wb["CONTROLE"]
    ws["B1"] = "Principal"
    ws["B2"] = "C0"
    ws["B3"] = None
    ws["B7"] = None
    ws["B8"] = None
    ws["A7"] = "Índice utilizado"
    ws["A8"] = "Data-base original"
    ws["A9"] = "Quantidade de ciclos desta análise"
    ws["A10"] = "Variação acumulada final"
    ws["A11"] = "Fator acumulado total"
    ws["A12"] = "Último ciclo considerado nesta apuração"
    ws["A13"] = "Início do último ciclo considerado"
    for ref in ("B3", "B8", "B13", "B14"):
        ws[ref].number_format = "mm/yyyy"
    for row in (10, 11, 12, 13):
        ws.row_dimensions[row].hidden = False
    ws.row_dimensions[14].hidden = True
    ws.sheet_view.showGridLines = False


def _reset_parametros(wb) -> None:
    ws = wb["parametros"]
    if ws.max_column > 8:
        ws.delete_cols(9, ws.max_column - 8)
    headers = [
        "COMPUTAR_NESTA_APURACAO", "CICLO", "PERIODO_DO_CICLO", "DATA_INICIO",
        "DATA_FIM", "PERCENTUAL_DO_CICLO", "FATOR_ACUMULADO", "SITUACAO",
    ]
    for col, value in enumerate(headers, 1):
        cell = ws.cell(1, col, value)
        cell.fill = HEADER
        cell.font = Font(name="Calibri", size=10, bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = GRID
    for row, cycle in enumerate(("C0", "C1", "C2", "C3", "C4"), 2):
        ws.cell(row, 1, "Nao").fill = INPUT
        ws.cell(row, 2, cycle)
        for col in (3, 4, 5):
            ws.cell(row, col).value = None
            ws.cell(row, col).number_format = "@" if col == 3 else "mm/yyyy"
        ws.cell(row, 6).value = 0 if cycle == "C0" else None
        ws.cell(row, 6).fill = INPUT
        ws.cell(row, 6).number_format = PERCENT
        if cycle == "C0":
            ws.cell(row, 7, "=1")
        else:
            prev = row - 1
            ws.cell(row, 7, f'=IF(F{row}="","",IF(NOT(ISNUMBER(F{row})),"",IF(NOT(ISNUMBER(G{prev})),"",G{prev}*(1+F{row}))))')
        ws.cell(row, 7).fill = AUTO
        ws.cell(row, 7).number_format = FACTOR
        ws.cell(row, 8, "Base" if cycle == "C0" else "Fora desta apuracao")
        for col in range(1, 9):
            ws.cell(row, col).font = Font(name="Calibri", size=10, color=GRAY)
            ws.cell(row, col).border = GRID
            ws.cell(row, col).alignment = Alignment(vertical="center")

    if ws.merged_cells:
        for merged in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(merged))
    ws["A9"] = "MEMORIA DO FATOR APLICAVEL A APURACAO"
    ws.merge_cells("A9:F9")
    ws["A9"].font = Font(name="Calibri", size=10, bold=True, color=NAVY)
    for col, value in enumerate(("CICLO", "COMPUTA?", "PERCENTUAL", "FATOR", "FATOR_ACUMULADO_APURACAO", "STATUS"), 1):
        c = ws.cell(10, col, value)
        c.fill = HEADER
        c.font = Font(name="Calibri", size=9, bold=True, color=WHITE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = GRID
    for i, cycle in enumerate(("C0", "C1", "C2", "C3", "C4")):
        row = 11 + i
        src = 2 + i
        ws.cell(row, 1, cycle)
        ws.cell(row, 2, f"=A{src}")
        ws.cell(row, 3, f'=IF(B{row}="Sim",F{src},"")')
        ws.cell(row, 4, "=1" if cycle == "C0" else f'=IF(C{row}="","",1+C{row})')
        ws.cell(row, 5, "=1" if cycle == "C0" else f'=IF(B{row}="Sim",E{row-1}*D{row},E{row-1})')
        ws.cell(row, 6, "Base" if cycle == "C0" else f'=IF(B{row}="Sim","Aplicado","Fora da apuracao")')
        for col in range(1, 7):
            c = ws.cell(row, col)
            c.fill = CHECK if cycle != "C0" else AUTO
            c.font = Font(name="Calibri", size=9, color=GRAY)
            c.border = GRID
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row, 3).number_format = PERCENT
        ws.cell(row, 4).number_format = FACTOR
        ws.cell(row, 5).number_format = FACTOR
    for col in range(1, 7):
        ws.cell(16, col).fill = PatternFill("solid", fgColor="FFD9E1F2")
        ws.cell(16, col).font = Font(name="Calibri", size=9, bold=True, color=GRAY)
        ws.cell(16, col).border = GRID
    ws["A16"] = "TOTAL APURACAO"
    ws["E16"] = "=E15"
    ws["E16"].number_format = FACTOR
    ws["F16"] = "=E16-1"
    ws["F16"].number_format = PERCENT
    ws.data_validations.dataValidation.clear()
    dv = DataValidation(type="list", formula1='"Sim,Nao"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("A2:A6")
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    widths = {"A": 30, "B": 10, "C": 24, "D": 16, "E": 16, "F": 23, "G": 20, "H": 28}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def _reset_financeiro(wb) -> None:
    ws = wb["financeiro"]
    if ws.max_column > 7:
        ws.delete_cols(8, ws.max_column - 7)
    if ws.max_row > 61:
        ws.delete_rows(62, ws.max_row - 61)
    for col, value in enumerate(
        ("COMPETENCIA", "CICLO", "VALOR_PAGO", "FATOR_APLICAVEL", "VALOR_ATUALIZADO", "DELTA", "EFEITO_FINANCEIRO"),
        1,
    ):
        ws.cell(1, col, value)
    formulas = {
        "D": '=IF(B{r}="","",IF(B{r}="c0",1,IF(B{r}="c1",IF(parametros!$A$3<>"Sim","",IF(NOT(ISNUMBER(parametros!$F$3)),"",1+parametros!$F$3)),IF(B{r}="c2",IF(parametros!$A$4<>"Sim","",IF(NOT(ISNUMBER(parametros!$F$4)),"",IF(parametros!$A$3="Sim",1+parametros!$F$3,1)*(1+parametros!$F$4))),IF(B{r}="c3",IF(parametros!$A$5<>"Sim","",IF(NOT(ISNUMBER(parametros!$F$5)),"",IF(parametros!$A$3="Sim",1+parametros!$F$3,1)*IF(parametros!$A$4="Sim",1+parametros!$F$4,1)*(1+parametros!$F$5))),IF(B{r}="c4",IF(parametros!$A$6<>"Sim","",IF(NOT(ISNUMBER(parametros!$F$6)),"",IF(parametros!$A$3="Sim",1+parametros!$F$3,1)*IF(parametros!$A$4="Sim",1+parametros!$F$4,1)*IF(parametros!$A$5="Sim",1+parametros!$F$5,1)*(1+parametros!$F$6))),""))))))',
        "E": '=IF(OR(C{r}="",NOT(ISNUMBER(D{r}))),"",ROUND(C{r}*D{r},2))',
        "F": '=IF(OR(C{r}="",E{r}=""),"",IF(G{r}<>"Sim",0,ROUND(E{r}-C{r},2)))',
    }
    for row in range(2, 62):
        for col in ("A", "B", "C", "G"):
            ws[f"{col}{row}"] = None
            ws[f"{col}{row}"].border = GRID
            ws[f"{col}{row}"].font = Font(name="Calibri", size=10, color=GRAY)
        ws[f"A{row}"].number_format = "mm/yyyy"
        ws[f"C{row}"].number_format = MONEY
        ws[f"C{row}"].fill = INPUT
        ws[f"G{row}"].alignment = Alignment(horizontal="center")
    _style_formula_range(ws, 2, 61, formulas)
    for row in range(2, 62):
        ws[f"D{row}"].number_format = FACTOR
        ws[f"E{row}"].number_format = MONEY
        ws[f"F{row}"].number_format = MONEY
    ws.freeze_panes = None
    ws.sheet_view.showGridLines = False
    for col, width in {"A": 16, "B": 14, "C": 18, "D": 17, "E": 19, "F": 18, "G": 19}.items():
        ws.column_dimensions[col].width = width


def _reset_itens_remanesc(wb) -> None:
    ws = wb["itens_Remanesc"]
    for row in range(2, 201):
        for col in ("A", "B", "C", "E", "G", "I", "K", "S", "T"):
            ws[f"{col}{row}"] = None
    formulas = {
        "D": '=IF(OR(A{r}="",B{r}="",C{r}=""),"",ROUND(B{r}*C{r},2))',
        "F": '=IF(OR(A{r}="",E{r}="",C{r}="",NOT(ISNUMBER($Z$3))),"",ROUND(E{r}*C{r}*$Z$3,2))',
        "H": '=IF(OR(A{r}="",G{r}="",C{r}="",NOT(ISNUMBER($Z$4))),"",ROUND(G{r}*C{r}*$Z$4,2))',
        "J": '=IF(OR(A{r}="",I{r}="",C{r}="",NOT(ISNUMBER($Z$5))),"",ROUND(I{r}*C{r}*$Z$5,2))',
        "L": '=IF(OR(A{r}="",K{r}="",C{r}="",NOT(ISNUMBER($Z$6))),"",ROUND(K{r}*C{r}*$Z$6,2))',
        "M": '=IF(OR(E{r}="",G{r}=""),"",ROUND(MAX(E{r}-G{r},0),2))',
        "N": '=IF(OR(M{r}="",C{r}="",NOT(ISNUMBER($Z$3))),"",ROUND(M{r}*C{r}*$Z$3,2))',
        "O": '=IF(OR(G{r}="",I{r}=""),"",ROUND(MAX(G{r}-I{r},0),2))',
        "P": '=IF(OR(O{r}="",C{r}="",NOT(ISNUMBER($Z$4))),"",ROUND(O{r}*C{r}*$Z$4,2))',
        "Q": '=IF(OR(I{r}="",K{r}=""),"",ROUND(MAX(I{r}-K{r},0),2))',
        "R": '=IF(OR(Q{r}="",C{r}="",NOT(ISNUMBER($Z$5))),"",ROUND(Q{r}*C{r}*$Z$5,2))',
        "U": '=IF(A{r}="","",IF(AND(G{r}<>"",E{r}<>"",G{r}>E{r}),"ALERTA: REM_C2>REM_C1",IF(AND(I{r}<>"",G{r}<>"",I{r}>G{r}),"ALERTA: REM_C3>REM_C2",IF(AND(K{r}<>"",I{r}<>"",K{r}>I{r}),"ALERTA: REM_C4>REM_C3","OK"))))',
        "V": '=IF(A{r}="","",IF(B{r}="","",IF(OR(B{r}<0,E{r}>B{r},G{r}>B{r},I{r}>B{r},K{r}>B{r}),"ALERTA: QTD_INVALIDA","OK")))',
        "AB": '=IF(OR(A{r}="",E{r}=""),"",ROUND(B{r}-E{r},2))',
        "AC": '=IF(OR(A{r}="",E{r}=""),"",ROUND(D{r}-E{r}*C{r},2))',
    }
    _style_formula_range(ws, 2, 200, formulas)
    for row in range(2, 201):
        for col in ("A", "B", "C", "E", "G", "I", "K"):
            ws[f"{col}{row}"].fill = INPUT
            ws[f"{col}{row}"].border = GRID
        for col in ("B", "E", "G", "I", "K", "M", "O", "Q", "S", "AB"):
            ws[f"{col}{row}"].number_format = QTY
        for col in ("C", "D", "F", "H", "J", "L", "N", "P", "R", "T", "AC"):
            ws[f"{col}{row}"].number_format = MONEY
    for i, cycle in enumerate(("C0", "C1", "C2", "C3", "C4"), 2):
        ws[f"W{i}"] = cycle
        ws[f"X{i}"] = f"=parametros!$A${i}"
        ws[f"Y{i}"] = f'=IF(ISNUMBER(parametros!$F${i}),parametros!$F${i},"")'
        ws[f"Z{i}"] = f'=IF(ISNUMBER(parametros!$G${i}),parametros!$G${i},"")'
        for col in ("W", "X", "Y", "Z"):
            ws[f"{col}{i}"].fill = AUTO
            ws[f"{col}{i}"].border = GRID
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False


def _reset_itens_consumidos(wb) -> None:
    ws = wb["itens_Consumidos"]
    ws["Q1"] = "CHECK"
    formulas = {
        "D": '=IF(OR(A{r}="",B{r}="",C{r}=""),"",ROUND(B{r}*C{r},2))',
        "F": '=IF(OR(A{r}="",E{r}="",C{r}="",NOT(ISNUMBER($U$2))),"",ROUND(E{r}*C{r}*$U$2,2))',
        "H": '=IF(OR(A{r}="",G{r}="",C{r}="",NOT(ISNUMBER($U$3))),"",ROUND(G{r}*C{r}*$U$3,2))',
        "J": '=IF(OR(A{r}="",I{r}="",C{r}="",NOT(ISNUMBER($U$4))),"",ROUND(I{r}*C{r}*$U$4,2))',
        "L": '=IF(OR(A{r}="",K{r}="",C{r}="",NOT(ISNUMBER($U$5))),"",ROUND(K{r}*C{r}*$U$5,2))',
        "N": '=IF(OR(A{r}="",M{r}="",C{r}="",NOT(ISNUMBER($U$6))),"",ROUND(M{r}*C{r}*$U$6,2))',
        "O": '=IF(A{r}="","",SUM(E{r},G{r},I{r},K{r},M{r}))',
        "P": '=IF(A{r}="","",SUM(F{r},H{r},J{r},L{r},N{r}))',
        "Q": '=IF(A{r}="","",IF(O{r}>B{r},"DIVERGENCIA: CONSUMO MAIOR QUE CONTRATADO","OK"))',
    }
    for row in range(2, 201):
        for col in ("A", "B", "C", "E", "G", "I", "K", "M"):
            ws[f"{col}{row}"] = None
            ws[f"{col}{row}"].fill = INPUT
            ws[f"{col}{row}"].border = GRID
    _style_formula_range(ws, 2, 200, formulas)
    for row in range(2, 201):
        for col in ("B", "E", "G", "I", "K", "M", "O"):
            ws[f"{col}{row}"].number_format = QTY
        for col in ("C", "D", "F", "H", "J", "L", "N", "P"):
            ws[f"{col}{row}"].number_format = MONEY
    for i, cycle in enumerate(("C0", "C1", "C2", "C3", "C4"), 2):
        ws[f"R{i}"] = cycle
        ws[f"S{i}"] = f"=parametros!$A${i}"
        ws[f"T{i}"] = f'=IF(ISNUMBER(parametros!$F${i}),parametros!$F${i},"")'
        ws[f"U{i}"] = f'=IF(ISNUMBER(parametros!$G${i}),parametros!$G${i},"")'
    widths = {"A": 14, "B": 20, "C": 18, "D": 18, "E": 16, "F": 17, "G": 16, "H": 17, "I": 16, "J": 17, "K": 16, "L": 17, "M": 16, "N": 17, "O": 18, "P": 20, "Q": 34}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for cell in ws[1]:
        if cell.value is not None:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
            font = copy(cell.font)
            font.name = "Calibri"
            font.sz = 10
            font.b = True
            cell.font = font
    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "D2"
    ws.sheet_view.showGridLines = False


def _reset_itens_pc(wb) -> None:
    ws = wb["itens_PC"]
    ws.column_dimensions["A"].hidden = False
    for table_name in list(ws.tables):
        del ws.tables[table_name]
    for row in range(101, ws.max_row + 1):
        for col in range(1, 13):
            ws.cell(row, col).value = None
    for row in range(2, 101):
        for col in ("A", "B", "C", "E", "H"):
            ws[f"{col}{row}"] = None
            ws[f"{col}{row}"].fill = INPUT
            ws[f"{col}{row}"].border = GRID
        ws[f"C{row}"].number_format = "dd/mm/yyyy"
        ws[f"E{row}"].number_format = MONEY
    formulas = {
        "D": '=IF(C{r}="","",IF(OR(C{r}<MIN(parametros!$D$2:$D$6),C{r}>MAX(parametros!$E$2:$E$6)),"Fora dos ciclos",LOOKUP(C{r},parametros!$D$2:$D$6,parametros!$B$2:$B$6)))',
        "F": '=IF(D{r}="","",IFERROR(VLOOKUP(D{r},parametros!$A$11:$E$15,5,0),""))',
        "G": '=IF(OR(E{r}="",F{r}=""),"",ROUND(E{r}*F{r},2))',
        "I": '=IF(H{r}="Sim",IF(OR(G{r}="",E{r}=""),"",ROUND(G{r}-E{r},2)),0)',
        "J": '=IF(H{r}="Nao",IF(G{r}="","",G{r}),0)',
        "K": '=IF(H{r}="Nao",IF(OR(G{r}="",E{r}=""),"",ROUND(G{r}-E{r},2)),0)',
        "L": '=IF(COUNTA(A{r}:H{r})=0,"",IF(C{r}="","DATA_PC vazia",IF(OR(D{r}="",D{r}="Fora dos ciclos"),"CICLO_PC nao identificado",IF(OR(E{r}="",E{r}=0),"VALOR_PC vazio ou zero",IF(AND(H{r}<>"Sim",H{r}<>"Nao"),"PC_PAGO_A_CONTRATADA invalido","OK")))))',
    }
    _style_formula_range(ws, 2, 100, formulas)
    for row in range(2, 101):
        ws[f"F{row}"].number_format = FACTOR
        for col in ("G", "I", "J", "K"):
            ws[f"{col}{row}"].number_format = MONEY
    ws.data_validations.dataValidation.clear()
    dv = DataValidation(type="list", formula1='"Sim,Nao"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("H2:H100")
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False


def _reset_aditivos(wb) -> None:
    ws = wb["aditivos"]
    formulas = {
        "C": '=IF(B{r}="","",IFERROR(IF(AND(B{r}>=parametros!$D$2,B{r}<=parametros!$E$2),"C0",IF(AND(B{r}>=parametros!$D$3,B{r}<=parametros!$E$3),"C1",IF(AND(B{r}>=parametros!$D$4,B{r}<=parametros!$E$4),"C2",IF(AND(B{r}>=parametros!$D$5,B{r}<=parametros!$E$5),"C3",IF(AND(B{r}>=parametros!$D$6,B{r}<=parametros!$E$6),"C4","Fora dos ciclos"))))),"Fora dos ciclos"))',
        "F": '=IF(A{r}="","",IFERROR(VLOOKUP(A{r},itens_Remanesc!$A:$C,3,0),""))',
        "G": '=IF(OR(E{r}="",F{r}=""),"",ROUND(E{r}*F{r},2))',
        "I": '=IFERROR(VLOOKUP(C{r},parametros!$B:$G,6,0),"")',
        "J": '=IF(G{r}="","",ROUND(IF(OR(UPPER(D{r})="SUPRESSAO",UPPER(D{r})="DECRESCIMO"),-1,1)*G{r}*IF(AND(UPPER(H{r})="SIM",ISNUMBER(I{r})),I{r},1),2))',
    }
    for row in range(2, 201):
        for col in ("A", "B", "D", "E", "H", "K"):
            ws[f"{col}{row}"] = None
            ws[f"{col}{row}"].fill = INPUT
            ws[f"{col}{row}"].border = GRID
        ws[f"B{row}"].number_format = "mm/yyyy"
        ws[f"E{row}"].number_format = QTY
    _style_formula_range(ws, 2, 200, formulas)
    for row in range(2, 201):
        for col in ("F", "G", "J"):
            ws[f"{col}{row}"].number_format = MONEY
        ws[f"I{row}"].number_format = FACTOR
    ws.data_validations.dataValidation.clear()
    for rng, values in (("D2:D200", "Acrescimo,Supressao"), ("H2:H200", "Sim,Nao"), ("K2:K200", "Sim,Nao")):
        dv = DataValidation(type="list", formula1=f'"{values}"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(rng)
    red_fill = PatternFill("solid", fgColor="FFFFC7CE")
    red_font = Font(color="FF9C0006")
    ws.conditional_formatting.add("A2:K200", FormulaRule(formula=['OR(LEFT(UPPER($D2),6)="SUPRES",LEFT(UPPER($D2),4)="DECR")'], fill=red_fill, font=red_font))
    widths = {"A": 12, "B": 18, "C": 16, "D": 30, "E": 25, "F": 22, "G": 24, "H": 24, "I": 22, "J": 25, "K": 25}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.row_dimensions[1].height = 40
    for cell in ws[1]:
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False


def _reset_historico_vu(wb) -> None:
    ws = wb["historico_VU"]
    formulas = {
        "A": '=IF(itens_Remanesc!A{r}="","",itens_Remanesc!A{r})',
        "B": '=IF(itens_Remanesc!A{r}="","",itens_Remanesc!B{r})',
        "C": '=IF(itens_Remanesc!A{r}="","",itens_Remanesc!C{r})',
        "D": '=IF(OR(A{r}="",C{r}="",NOT(ISNUMBER($L$3))),"",ROUND(C{r}*$L$3,2))',
        "E": '=IF(OR(A{r}="",C{r}="",NOT(ISNUMBER($L$4))),"",ROUND(C{r}*$L$4,2))',
        "F": '=IF(OR(A{r}="",C{r}="",NOT(ISNUMBER($L$5))),"",ROUND(C{r}*$L$5,2))',
        "G": '=IF(OR(A{r}="",C{r}="",NOT(ISNUMBER($L$6))),"",ROUND(C{r}*$L$6,2))',
        "H": '=IF(OR(A{r}="",C{r}="",C{r}=0,NOT(ISNUMBER($L$6))),"",ROUND($L$6-1,4))',
    }
    _style_formula_range(ws, 2, 200, formulas)
    for row in range(2, 201):
        ws[f"B{row}"].number_format = QTY
        for col in ("C", "D", "E", "F", "G"):
            ws[f"{col}{row}"].number_format = MONEY
        ws[f"H{row}"].number_format = PERCENT
    for i, cycle in enumerate(("C0", "C1", "C2", "C3", "C4"), 2):
        ws[f"J{i}"] = cycle
        ws[f"K{i}"] = f'=IF(ISNUMBER(parametros!$F${i}),parametros!$F${i},"")'
        ws[f"L{i}"] = f'=IF(ISNUMBER(parametros!$G${i}),parametros!$G${i},"")'
        ws[f"K{i}"].number_format = PERCENT
        ws[f"L{i}"].number_format = FACTOR
        for col in ("J", "K", "L"):
            ws[f"{col}{i}"].fill = AUTO
            ws[f"{col}{i}"].border = GRID
    ws.freeze_panes = "B2"
    ws.sheet_view.showGridLines = False


def _reset_itens_rc(wb) -> None:
    ws = wb["itens_RC"]
    for row in range(3, 203):
        src = row - 1
        formulas = {
            "A": f'=IF(itens_Remanesc!A{src}="","",itens_Remanesc!A{src})',
            "B": f'=IF(OR(itens_Remanesc!A{src}="",NOT(ISNUMBER(parametros!$G$2))),"",ROUND(itens_Remanesc!C{src}*parametros!$G$2,2))',
            "C": f'=IF(itens_Remanesc!A{src}="","",itens_Remanesc!B{src})',
            "D": f'=IF(A{row}="","",ROUND(B{row}*C{row},2))',
            "E": f'=IF(OR(itens_Remanesc!A{src}="",NOT(ISNUMBER(parametros!$G$3))),"",ROUND(itens_Remanesc!C{src}*parametros!$G$3,2))',
            "F": f'=IF(itens_Remanesc!A{src}="","",itens_Remanesc!E{src})',
            "G": f'=IF(OR(A{row}="",E{row}="",F{row}=""),"",ROUND(E{row}*F{row},2))',
            "H": f'=IF(OR(itens_Remanesc!A{src}="",NOT(ISNUMBER(parametros!$G$4))),"",ROUND(itens_Remanesc!C{src}*parametros!$G$4,2))',
            "I": f'=IF(itens_Remanesc!A{src}="","",itens_Remanesc!G{src})',
            "J": f'=IF(OR(A{row}="",H{row}="",I{row}=""),"",ROUND(H{row}*I{row},2))',
            "K": f'=IF(OR(itens_Remanesc!A{src}="",NOT(ISNUMBER(parametros!$G$5))),"",ROUND(itens_Remanesc!C{src}*parametros!$G$5,2))',
            "L": f'=IF(itens_Remanesc!A{src}="","",itens_Remanesc!I{src})',
            "M": f'=IF(OR(A{row}="",K{row}="",L{row}=""),"",ROUND(K{row}*L{row},2))',
            "N": f'=IF(OR(itens_Remanesc!A{src}="",NOT(ISNUMBER(parametros!$G$6))),"",ROUND(itens_Remanesc!C{src}*parametros!$G$6,2))',
            "O": f'=IF(itens_Remanesc!A{src}="","",itens_Remanesc!K{src})',
            "P": f'=IF(OR(A{row}="",N{row}="",O{row}=""),"",ROUND(N{row}*O{row},2))',
        }
        for col, formula in formulas.items():
            ws[f"{col}{row}"] = formula
            ws[f"{col}{row}"].border = GRID
    for row in range(3, 203):
        for col in ("B", "D", "E", "G", "H", "J", "K", "M", "N", "P"):
            ws[f"{col}{row}"].number_format = MONEY
        for col in ("C", "F", "I", "L", "O"):
            ws[f"{col}{row}"].number_format = QTY
    for col in range(1, 17):
        ws.cell(203, col).value = None
    ws["A203"] = "TOTAL"
    for col in ("D", "G", "J", "M", "P"):
        ws[f"{col}203"] = f"=SUM({col}3:{col}202)"
        ws[f"{col}203"].number_format = MONEY
    for cell in ws[203]:
        cell.fill = PatternFill("solid", fgColor="FFC0C0C0")
        cell.font = Font(name="Calibri", size=10, bold=True)
        cell.border = GRID
    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False


def _reset_resultados(wb) -> None:
    idx = wb.sheetnames.index("historico")
    old = wb["historico"]
    wb.remove(old)
    ws = wb.create_sheet("RESULTADOS", idx)
    ws.sheet_view.showGridLines = False


def _normalize(wb) -> None:
    for obsolete in ("itens_Execucao_Saldo", "REGRA_NEGOCIO_CLAUS"):
        if obsolete in wb.sheetnames:
            wb.remove(wb[obsolete])
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                cell.comment = None
                if cell.value is not None and ws.title != "CONTROLE":
                    font = copy(cell.font)
                    font.name = "Calibri"
                    font.sz = 10
                    cell.font = font
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.active = 0
    view = wb.views[0]
    view.visibility = "visible"
    view.minimized = False
    view.xWindow = 0
    view.yWindow = 0
    view.windowWidth = 22000
    view.windowHeight = 12000
    view.firstSheet = 0
    view.activeTab = 0


def build(source: Path, destination: Path) -> None:
    wb = load_workbook(source, data_only=False)
    required = {
        "CONTROLE", "parametros", "financeiro", "itens_Remanesc",
        "itens_Consumidos", "itens_PC", "aditivos", "historico",
        "itens_RC", "historico_VU",
    }
    missing = sorted(required.difference(wb.sheetnames))
    if missing:
        raise ValueError(f"Modelo de origem sem abas obrigatorias: {', '.join(missing)}")
    _reset_controle(wb)
    _reset_parametros(wb)
    _reset_financeiro(wb)
    _reset_itens_remanesc(wb)
    _reset_itens_consumidos(wb)
    _reset_itens_pc(wb)
    _reset_aditivos(wb)
    _reset_resultados(wb)
    _reset_itens_rc(wb)
    _reset_historico_vu(wb)
    _normalize(wb)
    destination.parent.mkdir(parents=True, exist_ok=True)
    wb.save(destination)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("destination", type=Path)
    args = parser.parse_args()
    build(args.source, args.destination)


if __name__ == "__main__":
    main()
