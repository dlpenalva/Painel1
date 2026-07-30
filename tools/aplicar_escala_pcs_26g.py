# -*- coding: utf-8 -*-
"""Etapa 26G: escala canonica de PCs + fail-closed do VTA, via Excel COM.

Gravador unico: Microsoft Excel real (padrao dos aplicadores anteriores).
Trabalha em copia temporaria e so promove ao destino apos recalculo
completo sem erros e reabertura verificada.

O que este owner aplica ao template oficial:

1. itens_PC — grade completa ate a capacidade canonica (_capacidade_pcs):
   formulas C/E/F/H/I/J/K/L, formatos, bordas, fills, Data Validation
   (G2:G{cap}) e formatacao condicional (A2:L{cap}). O CHECK de duplicidade
   migra de SUMPRODUCT(UPPER(TRIM(...))) — O(N^2), inviavel em 5.000 linhas —
   para COUNTIF nativo (case-insensitive); a robustez TRIM e garantida pela
   validacao Python do upload em toda a base.
2. Resumo lateral M:S consolidando TODA a faixa ($C$2:$C${cap}).
3. cobertura_temporal!B14 (ultima evidencia PC) sobre toda a faixa.
4. MEMORIA_RESULTADOS/RESULTADOS: toda referencia itens_PC!$X$2:$X$100
   passa para a capacidade canonica (Tabela 2, execucao do ciclo, D44/D45).
5. Fail-closed do VTA (VAZIO NAO E ZERO): X/Y por item deixam de produzir
   #VALUE! quando o remanescente obrigatorio esta ausente (retornam ""),
   e novos contadores T26/T27 tornam T25/C32 indisponiveis ("CALCULO
   MANUAL REQUERIDO" / "") — o STATUS GLOBAL resulta REVISE, nunca #VALUE!.
6. Visao executiva curta dos PCs sem efeito financeiro (RESULTADOS linha 23,
   associada a Tabela 2; helpers em MEMORIA!T28:T30). RESULTADOS segue limpa.
7. itens_Remanesc linha 101 ("item 100") reestilizada igual as vizinhas.
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from _capacidade_pcs import CAPACIDADE_PCS, ULTIMA_LINHA_PCS  # noqa: E402

CAP = ULTIMA_LINHA_PCS  # 5001: ultima linha da grade de dados de itens_PC

XL_PASTE_FORMATS = -4122
XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16
XL_EXPRESSION = 2
XL_VALIDATE_LIST = 3
XL_VALID_ALERT_STOP = 1
XL_BETWEEN = 1
XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105

COR_AMARELO = 153 * 65536 + 230 * 256 + 255  # RGB(255,230,153)
COR_VERMELHO = 204 * 65536 + 204 * 256 + 244  # RGB(244,204,204)
COR_CINZA_TEXTO = 0x59 + (0x59 << 8) + (0x59 << 16)
MOEDA_BR = "R$ #.##0,00"

ABA_MEMORIA = "MEMORIA_RESULTADOS"
ABA_RESULTADOS = "RESULTADOS"

_RE_FAIXA_ITENS_PC = re.compile(r"(itens_PC!\$[A-Z]{1,2}\$2:\$[A-Z]{1,2}\$)(?:100|200)\b")

# Trecho exato do CHECK atual (K) a substituir pelo COUNTIF escalavel.
_DUP_ANTIGA = (
    'SUMPRODUCT(--(UPPER(TRIM($A$2:$A$100))=UPPER(TRIM(A2))))>1'
)
_DUP_NOVA = f'COUNTIF($A$2:$A${CAP},A2)>1'

_PROT_FLAGS = (
    "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows",
    "AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks",
    "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting",
    "AllowFiltering", "AllowUsingPivotTables",
)


def _capturar_protecao(ws):
    if not bool(ws.ProtectContents):
        return None, None
    protecao = ws.Protection
    estado = {}
    for flag in _PROT_FLAGS:
        try:
            estado[flag] = getattr(protecao, flag)
        except Exception:
            estado[flag] = True
    try:
        selecao = ws.EnableSelection
    except Exception:
        selecao = None
    ws.Unprotect()
    return estado, selecao


def _restaurar_protecao(ws, estado, selecao) -> None:
    if estado is None:
        return
    ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True, **estado)
    if selecao is not None:
        try:
            ws.EnableSelection = selecao
        except Exception:
            pass


def _para_formula_local(ws, formula_en: str) -> str:
    apoio = ws.Range("AZ2")
    apoio.Formula = formula_en
    local = apoio.FormulaLocal
    apoio.ClearContents()
    return local


def _substituir_faixas(ws, enderecos: list[str]) -> int:
    """Reescreve refer. itens_PC!$X$2:$X$100|200 -> capacidade canonica."""
    trocas = 0
    for endereco in enderecos:
        celula = ws.Range(endereco)
        formula = str(celula.Formula or "")
        nova, n = _RE_FAIXA_ITENS_PC.subn(rf"\g<1>{CAP}", formula)
        if n:
            celula.Formula = nova
            trocas += n
    return trocas


def _aplicar_itens_pc(ws, excel) -> None:
    if str(ws.Range("A1").Value or "") != "NUMERO_PC" or str(
        ws.Range("L1").Value or ""
    ) != "EFEITO_FINANCEIRO_PC":
        raise ValueError("itens_PC fora do layout oficial A:L (NUMERO_PC..EFEITO_FINANCEIRO_PC).")

    # Nada pode existir alem da capacidade (sem truncamento silencioso).
    ultima_usada = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    if ultima_usada > CAP:
        alem = ws.Range(f"A{CAP + 1}:L{ultima_usada}")
        if excel.WorksheetFunction.CountA(alem) != 0:
            raise RuntimeError(f"itens_PC possui conteudo alem da capacidade em A{CAP + 1}:L{ultima_usada}.")

    # 1) Formatos (fontes, fills, bordas, number formats) da linha 2 para toda
    #    a grade. PasteSpecial(Formats) nao toca formulas nem validacao.
    ws.Range("A2:L2").Copy()
    ws.Range(f"A3:L{CAP}").PasteSpecial(XL_PASTE_FORMATS)
    excel.CutCopyMode = False
    # Altura uniforme: sem quebra visual apos a linha 100.
    ws.Rows(f"3:{CAP}").RowHeight = ws.Rows(2).RowHeight

    # 2) Formulas: as da linha 2 sao a fonte de verdade (zero drift semantico);
    #    somente K muda (duplicidade COUNTIF escalavel).
    formulas_linha2 = {c: str(ws.Range(f"{c}2").Formula or "") for c in "CEFHIJKL"}
    for coluna in "CEFHIJL":
        if not formulas_linha2[coluna].startswith("="):
            raise ValueError(f"itens_PC!{coluna}2 sem formula; template inesperado.")
    k2 = formulas_linha2["K"]
    if _DUP_ANTIGA not in k2:
        raise ValueError("CHECK de duplicidade (K2) fora do padrao esperado.")
    formulas_linha2["K"] = k2.replace(_DUP_ANTIGA, _DUP_NOVA)
    for coluna in "CEFHIJKL":
        ws.Range(f"{coluna}2:{coluna}{CAP}").Formula = formulas_linha2[coluna]

    # 3) Data Validation em toda a coluna G.
    validacao = ws.Range(f"G2:G{CAP}").Validation
    try:
        validacao.Delete()
    except Exception:
        pass
    validacao.Add(
        Type=XL_VALIDATE_LIST,
        AlertStyle=XL_VALID_ALERT_STOP,
        Operator=XL_BETWEEN,
        Formula1="Sim,Nao",
    )
    validacao.IgnoreBlank = True
    validacao.InCellDropdown = True
    validacao.ErrorMessage = "Selecione Sim ou Nao."

    # 4) Formatacao condicional em toda a grade.
    grade = ws.Range(f"A2:L{CAP}")
    grade.FormatConditions.Delete()
    ws.Activate()
    ws.Range("A2").Select()
    # Chamada POSICIONAL (Type, Operator ignorado, Formula1): argumentos
    # nomeados neste metodo falham com DISP_E_PARAMNOTFOUND no dispatch
    # dinamico do pywin32.
    amarela = grade.FormatConditions.Add(
        XL_EXPRESSION, 1,
        _para_formula_local(ws, '=AND(OR($A2<>"",$B2<>"",$D2<>"",$G2<>""),$L2="")'),
    )
    amarela.Interior.Color = COR_AMARELO
    amarela.StopIfTrue = True
    vermelha = grade.FormatConditions.Add(
        XL_EXPRESSION, 1, _para_formula_local(ws, '=$L2="Nao"')
    )
    vermelha.Interior.Color = COR_VERMELHO
    vermelha.StopIfTrue = False

    # 5) Resumo lateral N2:T6 sobre TODA a faixa (T = QTD_COM_CHECK).
    fontes = {"N": None, "O": "$D", "P": "$F", "Q": "$H", "R": "$I", "S": "$J"}
    for linha in range(2, 7):
        for coluna, origem in fontes.items():
            if origem is None:
                formula = f'=COUNTIF($C$2:$C${CAP},M{linha})'
            else:
                formula = (
                    f'=SUMIF($C$2:$C${CAP},M{linha},{origem}$2:{origem}${CAP})'
                )
            ws.Range(f"{coluna}{linha}").Formula = formula
        ws.Range(f"T{linha}").Formula = (
            f'=COUNTIFS($C$2:$C${CAP},M{linha},'
            f'$K$2:$K${CAP},"<>OK",$K$2:$K${CAP},"<>")'
        )

    # 6) Metadados constantes das colunas ocultas V:AC (Fase 1 do VTA):
    #    PCs alem da linha 100 devem carregar exatamente os mesmos campos.
    ws.Range("V2:AC2").Copy()
    ws.Range(f"V3:AC{CAP}").PasteSpecial(XL_PASTE_FORMATS)
    excel.CutCopyMode = False
    for coluna in ("V", "W", "X", "Y", "Z", "AA", "AB", "AC"):
        valor = ws.Range(f"{coluna}2").Value
        if valor not in (None, ""):
            ws.Range(f"{coluna}3:{coluna}{CAP}").Value = valor


def _aplicar_cobertura(ws) -> None:
    estado, selecao = _capturar_protecao(ws)
    ws.Range("B14").Formula = (
        f'=IF(COUNT(itens_PC!$B$2:$B${CAP})=0,"",MAX(itens_PC!$B$2:$B${CAP}))'
    )
    _restaurar_protecao(ws, estado, selecao)


def _aplicar_memoria(ws) -> None:
    estado, selecao = _capturar_protecao(ws)

    trocas = _substituir_faixas(ws, ["C10", "C11", "C12", "C13", "C14", "D44", "D45"])
    if trocas < 7:
        raise RuntimeError(f"MEMORIA: apenas {trocas} faixas itens_PC reescritas (esperado >= 7).")

    # Fail-closed por item: operando ausente devolve "" (nunca 0, nunca #VALUE!).
    ws.Range("X2:X201").Formula = (
        '=IF(posicao_contratual!$A2="",0,'
        'IF(AND(ISNUMBER(posicao_contratual!$G2),ISNUMBER(posicao_contratual!$K2),'
        'ISNUMBER(historico_VU!$C2)),'
        '(posicao_contratual!$G2-posicao_contratual!$K2)*historico_VU!$C2,""))'
    )
    _vu = (
        'CHOOSE($T$20+1,historico_VU!$C2,historico_VU!$D2,historico_VU!$E2,'
        'historico_VU!$F2,historico_VU!$G2)'
    )
    _qtd = (
        'CHOOSE($T$20+1,posicao_contratual!$G2,posicao_contratual!$K2,'
        'posicao_contratual!$O2,posicao_contratual!$S2,posicao_contratual!$W2)'
    )
    ws.Range("Y2:Y201").Formula = (
        f'=IF(OR(posicao_contratual!$A2="",$T$20=""),0,'
        f'IF(AND(ISNUMBER({_vu}),ISNUMBER({_qtd})),'
        f'ROUND(ROUND({_vu},2)*{_qtd},2),""))'
    )

    # Contadores de completude: itens com remanescente obrigatorio ausente.
    for endereco in ("T26", "T27", "T28", "T29", "T30"):
        atual = ws.Range(endereco).Formula
        if atual not in (None, ""):
            raise RuntimeError(f"MEMORIA!{endereco} ja possui conteudo: {atual!r}")
    _qtd_rng = (
        'CHOOSE($T$20+1,posicao_contratual!$G$2:$G$201,posicao_contratual!$K$2:$K$201,'
        'posicao_contratual!$O$2:$O$201,posicao_contratual!$S$2:$S$201,'
        'posicao_contratual!$W$2:$W$201)'
    )
    ws.Range("T26").Formula = (
        f'=IF($T$20="",0,SUMPRODUCT((posicao_contratual!$A$2:$A$201<>"")*'
        f'(1-ISNUMBER({_qtd_rng}))))'
    )
    ws.Range("T27").Formula = (
        '=SUMPRODUCT((posicao_contratual!$A$2:$A$201<>"")*'
        '(1-ISNUMBER(posicao_contratual!$G$2:$G$201)*'
        'ISNUMBER(posicao_contratual!$K$2:$K$201)*'
        'ISNUMBER(historico_VU!$C$2:$C$201)))'
    )

    # VTA por PCs indisponivel quando a base obrigatoria esta incompleta.
    ws.Range("T25").Formula = (
        '=IF(OR($T$24=0,$T$20="",$T$26>0,AND(NOT(itens_PC!$P$2>0),$T$27>0)),'
        '"CALCULO MANUAL REQUERIDO",ROUND($T$21+$T$22+$T$23,2))'
    )

    # C32 (remanescente atualizado informado): indisponivel sem base completa.
    c32 = str(ws.Range("C32").Formula or "")
    marcador = '"<>")=0,"",ROUND('
    if marcador not in c32 or "$T$26" in c32:
        raise RuntimeError("MEMORIA!C32 fora do padrao esperado para o guard fail-closed.")
    ws.Range("C32").Formula = (
        c32.replace(marcador, '"<>")=0,"",IF($T$26>0,"",ROUND(', 1)[:-1] + '))'
    )

    # Helpers executivos: PCs sem efeito financeiro nos ciclos em analise
    # (C0 nao entra: sem efeito por definicao; ciclo fora da apuracao idem).
    def _por_ciclo(func: str, faixa_soma: str | None) -> str:
        parcelas = []
        for ciclo, linha_par in (("C1", 3), ("C2", 4), ("C3", 5), ("C4", 6)):
            criterios = (
                f'itens_PC!$C$2:$C${CAP},"{ciclo}",'
                f'itens_PC!$L$2:$L${CAP},"Nao",'
                f'itens_PC!$D$2:$D${CAP},">0"'
            )
            if func == "COUNTIFS":
                nucleo = f'COUNTIFS({criterios})'
            else:
                nucleo = f'SUMIFS(itens_PC!{faixa_soma}$2:{faixa_soma}${CAP},{criterios})'
            parcelas.append(f'{nucleo}*(parametros!$A${linha_par}="Sim")')
        return "+".join(parcelas)

    ws.Range("T28").Formula = f'={_por_ciclo("COUNTIFS", None)}'
    ws.Range("T29").Formula = f'=ROUND({_por_ciclo("SUMIFS", "$D")},2)'
    ws.Range("T30").Formula = (
        f'=ROUND({_por_ciclo("SUMIFS", "$F")}-$T$29-({_por_ciclo("SUMIFS", "$H")}),2)'
    )
    rotulos = {
        "S26": "Itens sem remanescente do ciclo vigente",
        "S27": "Itens sem base p/ C0 fisico (X)",
        "S28": "PCs sem efeito financeiro (qtd)",
        "S29": "PCs sem efeito financeiro (valor-base)",
        "S30": "PCs sem efeito financeiro (dif. nao reconhecida)",
    }
    for endereco, texto in rotulos.items():
        if ws.Range(endereco).Formula in (None, ""):
            ws.Range(endereco).Value = texto

    _restaurar_protecao(ws, estado, selecao)


def _aplicar_resultados(ws) -> None:
    estado, selecao = _capturar_protecao(ws)

    trocas = _substituir_faixas(ws, ["B16", "B17", "B18", "B19", "B20", "B36"])
    if trocas < 6:
        raise RuntimeError(f"RESULTADOS: apenas {trocas} faixas itens_PC reescritas (esperado >= 6).")

    # Linha executiva unica associada a Tabela 2 (linha 23 e livre no 26F).
    for endereco in ("A23", "B23", "C23", "D23", "E23"):
        atual = ws.Range(endereco).Formula
        if atual not in (None, ""):
            raise RuntimeError(f"RESULTADOS!{endereco} ja possui conteudo: {atual!r}")
    # 26G.1: a linha so aparece quando o metodo oficial e PCs E ha PCs do
    # objeto da apuracao sem efeito financeiro (T28>0). Em Financeiro/Itens
    # a informacao nao pertence a Tabela 2 e permanece oculta.
    vazio = f'OR({ABA_MEMORIA}!$B$4<>"PCs",{ABA_MEMORIA}!$T$28=0)'
    ws.Range("A23").Formula = (
        f'=IF({vazio},"","PCs sem efeito financeiro: "&{ABA_MEMORIA}!$T$28&" PC(s)")'
    )
    ws.Range("B23").Formula = f'=IF({vazio},"",{ABA_MEMORIA}!$T$29)'
    ws.Range("C23").Formula = (
        f'=IF({vazio},"",ROUND({ABA_MEMORIA}!$T$29+{ABA_MEMORIA}!$T$30,2))'
    )
    ws.Range("D23").Formula = f'=IF({vazio},"",{ABA_MEMORIA}!$T$30)'
    ws.Range("E23:H23").Merge()
    ws.Range("E23").Formula = (
        f'=IF({vazio},"","Sem retroativo pela regra de DATA_PC. '
        'Detalhes: itens_PC -> EFEITO_FINANCEIRO_PC = Nao")'
    )
    rng = ws.Range("A23:H23")
    rng.Font.Italic = True
    rng.Font.Color = COR_CINZA_TEXTO
    rng.WrapText = True
    ws.Rows(23).RowHeight = 30
    ws.Range("B23:D23").NumberFormatLocal = MOEDA_BR
    for indice in (7, 8, 9, 10, 11, 12):
        borda = ws.Range("A23:H23").Borders(indice)
        borda.LineStyle = 1
        borda.Weight = 2

    _restaurar_protecao(ws, estado, selecao)


def _aplicar_itens_remanesc(ws, excel) -> None:
    # Linha 101 ("item 100") herda integralmente o estilo da linha 100.
    ws.Rows(100).Copy()
    ws.Rows(101).PasteSpecial(XL_PASTE_FORMATS)
    excel.CutCopyMode = False
    ws.Rows(101).RowHeight = ws.Rows(100).RowHeight


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


def _verificar_destino(excel, caminho: Path) -> None:
    wb = excel.Workbooks.Open(str(caminho), UpdateLinks=0, ReadOnly=True, CorruptLoad=0)
    try:
        ipc = wb.Worksheets("itens_PC")
        k_ultima = str(ipc.Range(f"K{CAP}").Formula or "")
        if f'COUNTIF($A$2:$A${CAP}' not in k_ultima:
            raise RuntimeError(f"itens_PC!K{CAP} sem o CHECK escalavel.")
        if str(ipc.Range("N2").Formula or "") != f'=COUNTIF($C$2:$C${CAP},M2)':
            raise RuntimeError("Resumo lateral N2 nao foi reescrito.")
        b14 = str(wb.Worksheets("cobertura_temporal").Range("B14").Formula or "")
        if f'$B$2:$B${CAP}' not in b14:
            raise RuntimeError("cobertura_temporal!B14 nao cobre toda a faixa.")
        t25 = str(wb.Worksheets(ABA_MEMORIA).Range("T25").Formula or "")
        if "$T$26" not in t25:
            raise RuntimeError("MEMORIA!T25 sem o guard fail-closed.")
        a23 = str(wb.Worksheets(ABA_RESULTADOS).Range("A23").Formula or "")
        if "T$28" not in a23:
            raise RuntimeError("RESULTADOS!A23 (PCs sem efeito) ausente.")
    finally:
        wb.Close(SaveChanges=False)


def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem).resolve()
    destino = Path(destino).resolve()
    if not origem.is_file():
        raise FileNotFoundError(origem)

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_26g_"))
    tmp_xlsx = tmp_dir / origem.name
    shutil.copyfile(origem, tmp_xlsx)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    salvo = False
    try:
        wb = excel.Workbooks.Open(str(tmp_xlsx), UpdateLinks=0, ReadOnly=False, CorruptLoad=0)
        excel.ScreenUpdating = False
        excel.Calculation = XL_CALC_MANUAL
        aba_ativa = wb.ActiveSheet.Name

        _aplicar_itens_pc(wb.Worksheets("itens_PC"), excel)
        _aplicar_cobertura(wb.Worksheets("cobertura_temporal"))
        _aplicar_memoria(wb.Worksheets(ABA_MEMORIA))
        _aplicar_resultados(wb.Worksheets(ABA_RESULTADOS))
        _aplicar_itens_remanesc(wb.Worksheets("itens_Remanesc"), excel)

        wb.Worksheets("itens_PC").Activate()
        wb.Worksheets("itens_PC").Range("A1").Select()
        if aba_ativa not in (ABA_MEMORIA,):
            try:
                wb.Worksheets(aba_ativa).Activate()
            except Exception:
                pass

        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()
        _verificar_sem_erros(wb)
        wb.Save()
        salvo = True
        wb.Close(SaveChanges=False)
        wb = None
        _verificar_destino(excel, tmp_xlsx)
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
    print(f"Escala 26G aplicada (capacidade {CAPACIDADE_PCS} PCs):", args.destino)


if __name__ == "__main__":
    main()
