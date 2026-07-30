"""Aplicador final da Etapa 1 (Execucao Historica) ao template oficial.

Aplicador UNICO e cirurgico. Alteracao estritamente ADITIVA/governada:

  1. parametros: marcador manual G11:G15 (dropdown Sim/Nao) para
     "reajuste anterior ja formalizado?" nos ciclos fora da apuracao. Vazio
     por padrao -> bloqueia calculo automatico do fator base ambiguo.

  2. itens_Remanesc: colunas AO:BH (20) com reconciliacao e as duas parcelas
     por item x ciclo: QTD_COBERTA, QTD_NAO_COBERTA, PARCELA_A (atualizacao
     anterior praticada) e PARCELA_B (atualizacao atual nao coberta), C0..C4.

  3. RESULTADOS: bloco "EIXO 5" (A256:O264) por ciclo, com fator cheio, fator
     base praticado, delta da apuracao, atualizacao anterior, atualizacao atual
     nao coberta, complemento potencial, VALOR ADOTADO (vazio) e status; linha
     de totais (263) com TOTAL ADOTADO em N263; status de integracao (A265).

  4. Integracao governada: B26 (VTA FINAL) incorpora N263 SOMENTE no ramo
     calculado (sem override). B25 (override total) permanece soberano. B23
     permanece INALTERADO.

Fator base praticado (canonico, reutiliza F42=B42/D42):
  - computado (D definido): F_cheio/F_apuracao;
  - fora + marcador "Sim": F_cheio (reajuste ja concedido) -> parcela atual 0;
  - fora sem comprovacao: vazio -> decisao manual.
Parcela A = QTD_EXEC*VU*(F_base-1); Parcela B = QTD_NAO_COBERTA*VU*(F_cheio-F_base).

Gravador unico: Microsoft Excel real via COM. Trabalha em copia temporaria e so
promove se o Excel salvar sem erros de formula. FAIL-CLOSED: recusa reaplicacao.
NAO altera aditivos (principal segue no gate manual B45/E26/B24/B25), template
legado, Python de producao nem paginas.
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
XL_VALIDATE_LIST = 3
XL_VALID_ALERT_STOP = 1
XL_GREATER_EQUAL = 7
XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105

FIM = 200

# ciclo, exec, cons, cob, nc, pa, pb, Fcheio, Dapur, statusF, marker, Fbase_cell
CICLOS = [
    ("C0", "AB", "E", "AO", "AP", "AQ", "AR", "$F$2", "$D$11", "$F$11", "$G$11", "$I$258"),
    ("C1", "M",  "G", "AS", "AT", "AU", "AV", "$F$3", "$D$12", "$F$12", "$G$12", "$I$259"),
    ("C2", "O",  "I", "AW", "AX", "AY", "AZ", "$F$4", "$D$13", "$F$13", "$G$13", "$I$260"),
    ("C3", "Q",  "K", "BA", "BB", "BC", "BD", "$F$5", "$D$14", "$F$14", "$G$14", "$I$261"),
    ("C4", "S",  "M", "BE", "BF", "BG", "BH", "$F$6", "$D$15", "$F$15", "$G$15", "$I$262"),
]
CAB_REMANESC = []
for _c in ("C0", "C1", "C2", "C3", "C4"):
    CAB_REMANESC += [f"QTD_COBERTA_{_c}", f"QTD_NAO_COBERTA_{_c}",
                     f"PARCELA_A_{_c}", f"PARCELA_B_{_c}"]

LIN_TITULO, LIN_HDR, LIN_C0, LIN_TOT, LIN_NOTA, LIN_INTEG = 256, 257, 258, 263, 264, 265
CAB_BLOCO = [
    "CICLO", "SITUACAO DO CICLO", "FONTE", "COBERTURA", "QTD EXECUTADA RECONSTRUIDA",
    "QTD COBERTA", "QTD NAO COBERTA", "FATOR CHEIO", "FATOR BASE PRATICADO",
    "DELTA DO FATOR DA APURACAO", "ATUALIZACAO ANTERIOR POTENCIAL",
    "ATUALIZACAO ATUAL NAO COBERTA", "COMPLEMENTO HISTORICO POTENCIAL",
    "VALOR ADOTADO", "STATUS / CHECK",
]
TITULO = ("EIXO 5 - EXECUCAO HISTORICA FORA DA ANALISE | apoio a decisao; "
          "so integra o VTA via VALOR ADOTADO (governado, nunca automatico)")
NOTA = ("PARCELA A = atualizacao anterior praticada (QTD_EXEC*VU*(fator_base-1)); "
        "PARCELA B = atualizacao atual nao coberta (QTD_NAO_COB*VU*(fator_cheio-fator_base)). "
        "VALOR ADOTADO e decisao humana (vazio padrao, zero ok, negativo nao).")
CAB_MARCADOR = "REAJUSTE ANTERIOR JA FORMALIZADO? (Sim/Nao; vazio=nao comprovado)"


def _validar_layout(wb, excel) -> None:
    wf = excel.WorksheetFunction
    ws_i = wb.Worksheets("itens_Remanesc")
    if wf.CountA(ws_i.Range(f"AO1:BH{FIM + 1}")) != 0:
        raise ValueError("itens_Remanesc AO1:BH201 nao vazia; bloco ja aplicado?")
    ws_r = wb.Worksheets("RESULTADOS")
    if wf.CountA(ws_r.Range(f"A{LIN_TITULO}:O{LIN_INTEG}")) != 0:
        raise ValueError("RESULTADOS A256:O265 nao vazia; bloco ja aplicado?")
    ws_p = wb.Worksheets("parametros")
    if wf.CountA(ws_p.Range("G10:G15")) != 0:
        raise ValueError("parametros G10:G15 nao vazia; marcador ja aplicado?")
    if ws_i.Range("AB1").Value is None or "EXEC" not in str(ws_i.Range("AA1").Value):
        raise ValueError("Ancora de execucao C0 ausente em itens_Remanesc.")


def _aplicar_marcador(ws, excel) -> None:
    ws.Range("F10").Copy()
    ws.Range("G10").PasteSpecial(XL_PASTE_FORMATS)
    excel.CutCopyMode = False
    ws.Range("G10").Value = CAB_MARCADOR
    ws.Columns("G").ColumnWidth = 42
    val = ws.Range("G12:G15").Validation
    try:
        val.Delete()
    except Exception:
        pass
    val.Add(Type=XL_VALIDATE_LIST, AlertStyle=XL_VALID_ALERT_STOP, Formula1="Sim,Nao")
    val.IgnoreBlank = True
    val.InCellDropdown = True


def _aplicar_remanesc(ws, excel) -> None:
    ws.Range("Z1").Copy()
    ws.Range("AO1:BH1").PasteSpecial(XL_PASTE_FORMATS)
    excel.CutCopyMode = False
    ws.Range("AO1:BH1").Value = [CAB_REMANESC]
    for (_c, exe, cons, cob, nc, pa, pb, Fcheio, _D, _s, _m, Fb) in CICLOS:
        ws.Range(f"{cob}2:{cob}{FIM}").Formula = (
            f'=IF($A2="","",ROUND(SUMIFS(itens_Consumidos!${cons}$2:${cons}$200,'
            f'itens_Consumidos!$A$2:$A$200,$A2),2))'
        )
        ws.Range(f"{nc}2:{nc}{FIM}").Formula = (
            f'=IF(OR($A2="",{exe}2=""),"",ROUND(MAX({exe}2-{cob}2,0),2))'
        )
        ws.Range(f"{pa}2:{pa}{FIM}").Formula = (
            f'=IF(OR($A2="",{exe}2="",RESULTADOS!{Fb}=""),"",'
            f'ROUND({exe}2*$C2*(RESULTADOS!{Fb}-1),2))'
        )
        ws.Range(f"{pb}2:{pb}{FIM}").Formula = (
            f'=IF(OR($A2="",{nc}2="",RESULTADOS!{Fb}=""),"",'
            f'ROUND({nc}2*$C2*(parametros!{Fcheio}-RESULTADOS!{Fb}),2))'
        )
    ws.Columns("AO:BH").ColumnWidth = 15
    ws.Columns("AO:BH").Hidden = True  # tecnicas; reexibiveis, nunca veryHidden


def _aplicar_bloco(ws, excel) -> None:
    ws.Range(f"A{LIN_TITULO}").Value = TITULO
    ws.Range("A19").Copy()
    ws.Range(f"A{LIN_TITULO}").PasteSpecial(XL_PASTE_FORMATS)
    ws.Range("A9:O9").Copy()
    ws.Range(f"A{LIN_HDR}:O{LIN_HDR}").PasteSpecial(XL_PASTE_FORMATS)
    excel.CutCopyMode = False
    ws.Range(f"A{LIN_HDR}:O{LIN_HDR}").Value = [CAB_BLOCO]

    nr = '$B$4<>"Itens"'  # nao reconciliavel (verificacao positiva do metodo)
    for i, (nome, exe, cons, cob, nc, pa, pb, Fc, Dp, sF, mk, Fb) in enumerate(CICLOS):
        r = LIN_C0 + i
        ws.Range(f"A{r}").Value = nome
        ws.Range(f"B{r}").Formula = (
            f'=IF(parametros!{sF}="Base","BASE",IF(parametros!{sF}="Aplicado",'
            f'"APLICADO NA APURACAO",IF(parametros!{mk}="Sim","FORA - FORMALIZADO",'
            f'IF(parametros!{mk}="Nao","FORA - NAO FORMALIZADO","FORA - NAO COMPROVADO"))))'
        )
        ws.Range(f"C{r}").Formula = "=$B$4"
        ws.Range(f"E{r}").Formula = (
            f'=IF(COUNT(itens_Remanesc!{exe}$2:{exe}$200)=0,"",'
            f'ROUND(SUM(itens_Remanesc!{exe}$2:{exe}$200),2))'
        )
        ws.Range(f"F{r}").Formula = (
            f'=IF(OR(E{r}="",{nr}),"n/d",ROUND(SUM(itens_Remanesc!{cob}$2:{cob}$200),2))'
        )
        ws.Range(f"G{r}").Formula = (
            f'=IF(OR(E{r}="",{nr}),"n/d",ROUND(SUM(itens_Remanesc!{nc}$2:{nc}$200),2))'
        )
        ws.Range(f"H{r}").Formula = (
            f'=IF(ISNUMBER(parametros!{Fc}),parametros!{Fc},"")'
        )
        # Fator base praticado (canonico): F/D quando computado; F_cheio quando
        # fora+formalizado; vazio quando ambiguo (decisao manual).
        ws.Range(f"I{r}").Formula = (
            f'=IF(ISNUMBER(parametros!{Dp}),parametros!{Fc}/parametros!{Dp},'
            f'IF(parametros!{mk}="Sim",parametros!{Fc},""))'
        )
        ws.Range(f"J{r}").Formula = (
            f'=IF(OR(H{r}="",I{r}=""),"",ROUND(H{r}-I{r},2))'
        )
        # K: atualizacao anterior (parcela A) - NAO depende de cobertura/metodo
        ws.Range(f"K{r}").Formula = (
            f'=IF(I{r}="","",ROUND(SUM(itens_Remanesc!{pa}$2:{pa}$200),2))'
        )
        # L: atualizacao atual nao coberta (parcela B) - exige metodo reconciliavel
        ws.Range(f"L{r}").Formula = (
            f'=IF(OR(I{r}="",{nr}),"",ROUND(SUM(itens_Remanesc!{pb}$2:{pb}$200),2))'
        )
        ws.Range(f"M{r}").Formula = (
            f'=IF(AND(K{r}="",L{r}=""),"",ROUND(IF(ISNUMBER(K{r}),K{r},0)+'
            f'IF(ISNUMBER(L{r}),L{r},0),2))'
        )
        ws.Range(f"D{r}").Formula = (
            f'=IF({nr},"NAO RECONCILIAVEL",IF(E{r}="","SEM EXECUCAO",'
            f'IF(AND(ISNUMBER(F{r}),F{r}>E{r}),"COBERTURA SUPERIOR",'
            f'IF(G{r}=0,"COMPLETA",IF(G{r}=E{r},"AUSENTE","PARCIAL")))))'
        )
        ws.Range(f"O{r}").Formula = (
            f'=IF(AND(N{r}<>"",NOT(ISNUMBER(N{r}))),"VALOR ADOTADO INVALIDO",'
            f'IF(AND(ISNUMBER(N{r}),N{r}<0),"NEGATIVO - CORRIGIR",'
            f'IF(AND(ISNUMBER(N{r}),ISNUMBER(M{r}),N{r}>M{r}),"ADOTADO SUPERIOR AO POTENCIAL - justificar",'
            f'IF(ISNUMBER(N{r}),"ADOTADO MANUAL",'
            f'IF(I{r}="","FATOR HISTORICO NAO COMPROVADO - decisao manual",'
            f'IF(AND({nr},E{r}<>""),"PARCELA ATUAL NAO RECONCILIAVEL - manual",'
            f'IF(M{r}="","sem complemento","COMPLEMENTO POTENCIAL - adocao manual")))))))'
        )

    # linha de totais
    ws.Range(f"A{LIN_TOT}").Value = "TOTAIS"
    ws.Range(f"K{LIN_TOT}").Formula = f'=IF(COUNT(K{LIN_C0}:K{LIN_C0+4})=0,"",ROUND(SUM(K{LIN_C0}:K{LIN_C0+4}),2))'
    ws.Range(f"L{LIN_TOT}").Formula = f'=IF(COUNT(L{LIN_C0}:L{LIN_C0+4})=0,"",ROUND(SUM(L{LIN_C0}:L{LIN_C0+4}),2))'
    ws.Range(f"M{LIN_TOT}").Formula = f'=IF(COUNT(M{LIN_C0}:M{LIN_C0+4})=0,"",ROUND(SUM(M{LIN_C0}:M{LIN_C0+4}),2))'
    ws.Range(f"N{LIN_TOT}").Formula = f'=ROUND(SUM(N{LIN_C0}:N{LIN_C0+4}),2)'  # TOTAL ADOTADO
    ws.Range(f"J{LIN_TOT}").Value = "TOTAL ADOTADO ->"

    ws.Range(f"A{LIN_NOTA}").Value = NOTA
    # status de integracao ao VTA (alerta de override; sem editar E26 homologado)
    ws.Range(f"A{LIN_INTEG}").Formula = (
        f'=IF(AND(ISNUMBER(B25),$N${LIN_TOT}>0),'
        f'"INTEGRACAO: OVERRIDE MANUAL (B25) VIGENTE - confirmar se complemento historico ja incluido; NAO somado automaticamente",'
        f'IF($N${LIN_TOT}>0,"INTEGRACAO: complemento adotado incorporado ao VTA FINAL (B26)",'
        f'IF($M${LIN_TOT}>0,"INTEGRACAO: complemento potencial disponivel - nao adotado",'
        f'"INTEGRACAO: sem complemento historico")))'
    )

    ws.Range(f"E{LIN_C0}:M{LIN_C0+5}").NumberFormatLocal = "#.##0,00;-#.##0,00"
    ws.Range(f"N{LIN_C0}:N{LIN_TOT}").NumberFormatLocal = "#.##0,00;-#.##0,00"
    ws.Range(f"H{LIN_C0}:J{LIN_C0+4}").NumberFormatLocal = "0,000000"
    val = ws.Range(f"N{LIN_C0}:N{LIN_C0+4}").Validation
    try:
        val.Delete()
    except Exception:
        pass
    val.Add(Type=XL_VALIDATE_DECIMAL, AlertStyle=XL_VALID_ALERT_STOP,
            Operator=XL_GREATER_EQUAL, Formula1="0")
    val.IgnoreBlank = True
    val.ErrorMessage = "VALOR ADOTADO deve ser >= 0 (zero permitido)."


def _integrar_vta(ws) -> None:
    """Adiciona o TOTAL ADOTADO (N263) ao B26 apenas no ramo calculado.

    B23 inalterado; B25 (override total) permanece soberano.
    """
    b26_atual = str(ws.Range("B26").Formula)
    esperado_fim = 'ROUND(B23+IF(ISNUMBER(B24),B24,0),2))))'
    if esperado_fim not in b26_atual:
        raise RuntimeError(f"B26 homologado inesperado: {b26_atual!r}")
    novo = b26_atual.replace(
        'ROUND(B23+IF(ISNUMBER(B24),B24,0),2)',
        f'ROUND(B23+IF(ISNUMBER(B24),B24,0)+IF(ISNUMBER($N${LIN_TOT}),$N${LIN_TOT},0),2)',
    )
    ws.Range("B26").Formula = novo


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

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_etapa1_"))
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
        _validar_layout(wb, excel)
        _aplicar_marcador(wb.Worksheets("parametros"), excel)
        _aplicar_bloco(wb.Worksheets("RESULTADOS"), excel)
        _aplicar_remanesc(wb.Worksheets("itens_Remanesc"), excel)
        _integrar_vta(wb.Worksheets("RESULTADOS"))
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
