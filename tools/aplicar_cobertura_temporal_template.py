"""Aplicador da Etapa — aba de diagnostico "cobertura_temporal" ao template.

Alteracao estritamente ADITIVA e GOVERNADA: (re)cria UMA aba diagnostica
(GCC/automatica) que consolida os marcos temporais e separa, no BLOCO B,
ULTIMA EVIDENCIA (automatica, MAX da data) de COBERTURA CONFIRMADA COMPLETA
(entrada GCC). SEM tocar o VTA oficial (B23/B25/B26), sem alterar itens_Remanesc
nem posicao_referencia e sem criar campo fiscal novo. Idempotente: se a aba ja
existir, e substituida (hotfix re-aplicavel).

Reutiliza integralmente:
  * a decisao homologada de `posicao_referencia` (painel I2/I5/I6/I8): posicao
    atual completa vs. fallback global para a fotografia historica;
  * a linha temporal canonica (parametros!C2:C6) para o inicio do ciclo atual;
  * as fontes concretas ja lidas: financeiro (COMPETENCIA, col A) e itens_PC
    (DATA_PC, col B) — apenas para a ULTIMA EVIDENCIA (nunca "completo ate").

BLOCO B (hotfix): ultima evidencia (auto) x confirmado completo ate (GCC).
  Posicao fisica conhecida ate      (auto)
  Ultima evidencia Financeiro       (auto = MAX competencia)
  Financeiro confirmado completo ate (GCC, opcional)
  Ultima evidencia PC               (auto = MAX DATA_PC)
  PC confirmado completo ate        (GCC, opcional)
  Projecao autorizada a partir de   (auto, FAIL-CLOSED: dia seguinte a
                                     max(fisica, confirmadas GCC); NUNCA a partir
                                     da ultima evidencia nao confirmada)

Legenda FISCAL / GCC / AUTOMATICO / PROJECAO. Entradas GCC (amarelo de entrada):
B4 (data da analise), B13 (financeiro confirmado), B15 (PC confirmado). Tudo o
mais e formula (AUTOMATICO) ou diagnostico de PROJECAO.

Cores REUTILIZADAS (nao inventa tonalidade para categoria existente):
  cabecalho/banda = CONTROLE!A6; rotulo = CONTROLE!A7; valor auto = CONTROLE!B9;
  entrada GCC = CONTROLE!B3 (amarelo). Categoria PROJECAO = laranja claro FCE4D6.

Gravador unico: Microsoft Excel real via COM. NAO usa openpyxl para salvar.
"""
from __future__ import annotations

import argparse
import datetime as _dt
import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

XL_PASTE_FORMATS = -4122
XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16
XL_VALIDATE_DATE = 4
XL_VALID_ALERT_STOP = 1
XL_BETWEEN = 1
XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105

ABA = "cobertura_temporal"
FMT_DATA = "dd/mm/aaaa"   # pt-BR: token de ano "aaaa"
COR_PROJECAO = 14083324   # BGR de FCE4D6 (laranja claro) — categoria PROJECAO

# posicao_referencia — painel homologado reutilizado.
PR = "posicao_referencia"
I_COMPLETA = f"{PR}!$I$2"     # POSICAO ATUAL COMPLETA?
I_DATA_REF = f"{PR}!$I$5"     # DATA DA POSICAO DE REFERENCIA
I_ABERTURA = f"{PR}!$I$6"     # DATA DE ABERTURA DO CICLO DE REFERENCIA
I_ORIGEM = f"{PR}!$I$8"       # ORIGEM DA POSICAO

# Inicio do ciclo atual (CONTROLE!B2) sobre a linha temporal parametros!C.
INICIO_CICLO_ATUAL = (
    '=IF($B$5="C4",parametros!$C$6,IF($B$5="C3",parametros!$C$5,'
    'IF($B$5="C2",parametros!$C$4,IF($B$5="C1",parametros!$C$3,'
    'IF($B$5="C0",parametros!$C$2,"")))))'
)
# Ultima evidencia (MAX). NUNCA rotular como "completo ate".
FIN_ULTIMA = ('=IF(COUNT(financeiro!$A$2:$A$200)=0,"",MAX(financeiro!$A$2:$A$200))')
PC_ULTIMA = ('=IF(COUNT(itens_PC!$B$2:$B$200)=0,"",MAX(itens_PC!$B$2:$B$200))')

# Cobertura ADOTADA (fail-closed): posicao fisica (B11) + confirmadas GCC
# (B13 financeiro, B15 PC). NAO usa a ultima evidencia nao confirmada.
COB_ADOTADA = "MAX($B$11,$B$13,$B$15)"
# Projecao autorizada = dia seguinte a cobertura adotada, se a analise (B4)
# ultrapassa essa cobertura. Caso contrario, vazio.
PROJ_AUTORIZADA = (
    f'=IF(OR($B$4="",NOT(ISNUMBER($B$4))),"",'
    f'IF({COB_ADOTADA}=0,"",'
    f'IF($B$4>{COB_ADOTADA},{COB_ADOTADA}+1,"")))'
)

# Posterioridade da ULTIMA EVIDENCIA em relacao a posicao fisica (B11).
POST_FIN = '(AND($B$12<>"",$B$11<>"",$B$12>$B$11))'
POST_PC = '(AND($B$14<>"",$B$11<>"",$B$14>$B$11))'
MODO = (
    f'=IF(AND({I_COMPLETA},NOT({POST_FIN}),NOT({POST_PC})),"POSICAO_ATUAL",'
    f'IF(AND({POST_FIN},{POST_PC}),"HIBRIDO_TEMPORAL",'
    f'IF({POST_FIN},"FINANCEIRO_POSTERIOR",'
    f'IF({POST_PC},"PC_POSTERIOR","POSICAO_DE_CORTE"))))'
)


def _dstr(ref: str) -> str:
    """dd/mm/aaaa locale-proof (DAY/MONTH/YEAR) para rotulos textuais."""
    return (f'(RIGHT("0"&DAY({ref}),2)&"/"&RIGHT("0"&MONTH({ref}),2)&"/"&YEAR({ref}))')


def _validar_layout(wb) -> None:
    nomes = [ws.Name for ws in wb.Worksheets]
    for aba in ("CONTROLE", "parametros", "financeiro", "itens_PC",
                "posicao_referencia", "RESULTADOS"):
        if aba not in nomes:
            raise ValueError(f"Aba canonica ausente: {aba}")
    pr = wb.Worksheets(PR)
    if "DATA DA POSICAO DE REFERENCIA" not in str(pr.Range("H5").Value or ""):
        raise ValueError("posicao_referencia!H5 nao e o painel de referencia esperado.")


def _remover_aba_existente(wb, excel) -> None:
    """Idempotencia: remove a aba antiga antes de recriar (hotfix)."""
    for ws in wb.Worksheets:
        if ws.Name == ABA:
            alertas = excel.DisplayAlerts
            excel.DisplayAlerts = False
            ws.Delete()
            excel.DisplayAlerts = alertas
            break


def _fmt(ws, modelo, origem_addr, destino_addr, excel) -> None:
    modelo.Range(origem_addr).Copy()
    ws.Range(destino_addr).PasteSpecial(XL_PASTE_FORMATS)
    excel.CutCopyMode = False


def _validacao_data(ws, addr, titulo, msg) -> None:
    cel = ws.Range(addr)
    cel.NumberFormatLocal = FMT_DATA
    val = cel.Validation
    try:
        val.Delete()
    except Exception:
        pass
    base = _dt.date(1899, 12, 30)
    lo = (_dt.date(1990, 1, 1) - base).days
    hi = (_dt.date(2199, 12, 31) - base).days
    val.Add(Type=XL_VALIDATE_DATE, AlertStyle=XL_VALID_ALERT_STOP,
            Operator=XL_BETWEEN, Formula1=lo, Formula2=hi)
    val.IgnoreBlank = True
    val.InputTitle = titulo
    val.InputMessage = msg
    val.ErrorMessage = "Informe uma data valida (dd/mm/aaaa)."


def _criar_aba(wb, excel):
    # Inserida ANTES de RESULTADOS: RESULTADOS permanece a ultima aba (exigido
    # por _coleta_reajuste.py e pelos testes de integridade estrutural).
    ws = wb.Worksheets.Add(Before=wb.Worksheets("RESULTADOS"))
    ws.Name = ABA
    ctrl = wb.Worksheets("CONTROLE")

    titulo = ("COBERTURA TEMPORAL - POSICAO FISICA NO CICLO EM EXECUCAO  |  "
              "diagnostico GCC/automatico; NAO altera o VTA (B23/B25/B26)")

    # (linha, rotulo, formula|None, formato_valor, categoria)
    # categoria: "auto" | "gcc" | "proj" | "banner"
    linhas = [
        (1, titulo, None, None, "banner"),
        (3, "BLOCO A - MARCOS", None, None, "banner"),
        (4, "Data da analise (GCC, opcional)", None, FMT_DATA, "gcc"),
        (5, "Ciclo atual (em execucao)",
         '=IF(CONTROLE!$B$2="","",CONTROLE!$B$2)', None, "auto"),
        (6, "Inicio do ciclo atual", INICIO_CICLO_ATUAL, FMT_DATA, "auto"),
        (7, "Data da fotografia fisica do corte (abertura)", f"={I_ABERTURA}", FMT_DATA, "auto"),
        (8, "Data da fotografia fisica mais recente",
         f'=IF({I_COMPLETA},{I_DATA_REF},"")', FMT_DATA, "auto"),
        (10, "BLOCO B - ULTIMA EVIDENCIA x COBERTURA CONFIRMADA", None, None, "banner"),
        (11, "Posicao fisica conhecida ate", f"={I_DATA_REF}", FMT_DATA, "auto"),
        (12, "Ultima evidencia Financeiro (nao e completo ate)", FIN_ULTIMA, FMT_DATA, "auto"),
        (13, "Financeiro confirmado completo ate (GCC, opcional)", None, FMT_DATA, "gcc"),
        (14, "Ultima evidencia PC (nao e completo ate)", PC_ULTIMA, FMT_DATA, "auto"),
        (15, "PC confirmado completo ate (GCC, opcional)", None, FMT_DATA, "gcc"),
        (16, "Projecao autorizada a partir de", PROJ_AUTORIZADA, FMT_DATA, "proj"),
        (18, "BLOCO C - DECISAO", None, None, "banner"),
        (19, "Modo temporal", MODO, None, "auto"),
        (20, "Fonte principal",
         '=IF($B$12<>"","Financeiro",IF($B$14<>"","PC",""))', None, "auto"),
        (21, "Fontes de conferencia",
         '=IF(AND($B$12<>"",$B$14<>""),"PC","")', None, "auto"),
        (22, "Posicao observada (data)", f"={I_DATA_REF}", FMT_DATA, "auto"),
        (23, "Posicao observada (origem)", f"={I_ORIGEM}", None, "auto"),
        (24, "Posicao projetada",
         f'=IF($B$16="","NAO PROJETADO (posicao observada mantida)",'
         f'"ESTIMATIVA a partir de "&{_dstr("$B$16")}&'
         '" - nao observada, nao cria retroativo a pagar")', None, "proj"),
    ]

    for lin, rot, formula, fmt, cat in linhas:
        cel_val = ws.Range(f"B{lin}")
        if cat == "banner":
            ws.Range(f"A{lin}:C{lin}").Merge()
            _fmt(ws, ctrl, "A6", f"A{lin}", excel)
            ws.Range(f"A{lin}").Value = rot
            continue
        _fmt(ws, ctrl, "A7", f"A{lin}", excel)      # rotulo (azul claro)
        ws.Range(f"A{lin}").Value = rot
        if cat == "gcc":
            _fmt(ws, ctrl, "B3", f"B{lin}", excel)  # amarelo de entrada
        else:
            _fmt(ws, ctrl, "B9", f"B{lin}", excel)  # azul muito claro / formula
            if cat == "proj":
                cel_val.Interior.Color = COR_PROJECAO
        if formula:
            cel_val.Formula = formula
        if fmt:
            cel_val.NumberFormatLocal = fmt

    # Entradas GCC (datas opcionais).
    _validacao_data(
        ws, "B4", "Data da analise (GCC, opcional)",
        "Data de referencia da analise. Opcional; se vazia, nao ha diagnostico "
        "de projecao. NAO altera a posicao observada.")
    _validacao_data(
        ws, "B13", "Financeiro confirmado (GCC)",
        "Data ate a qual a cobertura FINANCEIRA e CONFIRMADA completa (GCC). "
        "Vazio = cobertura completa NAO confirmada; usa-se so a ultima evidencia.")
    _validacao_data(
        ws, "B15", "PC confirmado (GCC)",
        "Data ate a qual a cobertura de PCs e CONFIRMADA completa (GCC). PC nao "
        "admite inferencia automatica: vazio = completude NAO confirmada.")

    # Legenda — reutiliza a banda de banner e amostra as cores por categoria.
    _fmt(ws, ctrl, "A6", "A26", excel)
    ws.Range("A26:C26").Merge()
    ws.Range("A26").Value = "LEGENDA - QUEM PREENCHE O QUE"
    legenda = [
        (27, "FISCAL", "B3", "Amarelo: itens_Remanesc, posicao_referencia (QTD_REM_ATUAL) e CONTROLE!B3."),
        (28, "GCC", "B3", "Amarelo: Data da analise (B4) e as confirmacoes de cobertura completa (B13/B15)."),
        (29, "AUTOMATICO", "B9", "Derivado de CONTROLE/parametros/posicao_referencia/financeiro/itens_PC (MAX = ultima evidencia)."),
        (30, "PROJECAO", None, "Laranja claro: estimativa (nunca fato observado; so a partir da cobertura confirmada/fisica)."),
    ]
    for lin, rotulo, modelo_fmt, desc in legenda:
        if modelo_fmt:
            _fmt(ws, ctrl, modelo_fmt, f"A{lin}", excel)
        else:
            _fmt(ws, ctrl, "B9", f"A{lin}", excel)
            ws.Range(f"A{lin}").Interior.Color = COR_PROJECAO
        ws.Range(f"A{lin}").Value = rotulo
        ws.Range(f"B{lin}:C{lin}").Merge()
        ws.Range(f"B{lin}").Value = desc

    ws.Columns("A").ColumnWidth = 46
    ws.Columns("B").ColumnWidth = 26
    ws.Columns("C").ColumnWidth = 40
    return ws


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

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_cobertura_"))
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
        _validar_layout(wb)
        _remover_aba_existente(wb, excel)
        _criar_aba(wb, excel)
        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()
        _verificar_sem_erros(wb)
        if aba_ativa != ABA:
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
