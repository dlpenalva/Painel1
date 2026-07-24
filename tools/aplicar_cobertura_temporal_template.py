"""Aplicador da Etapa — aba de diagnostico "cobertura_temporal" ao template.

Alteracao estritamente ADITIVA e GOVERNADA: cria UMA nova aba diagnostica
(GCC/automatica) que consolida os marcos temporais ja existentes e adiciona as
duas fronteiras que faltavam (Financeiro-completo-ate / PC-completo-ate), SEM
tocar o VTA oficial (B23/B25/B26), sem alterar itens_Remanesc e sem criar campo
fiscal novo. Reutiliza integralmente:

  * a decisao homologada de `posicao_referencia` (painel I2/I5/I6/I8): posicao
    atual completa vs. fallback global para a fotografia historica;
  * a linha temporal canonica (parametros!C2:C6) para o inicio do ciclo atual;
  * as fontes concretas ja lidas: financeiro (COMPETENCIA, col A) e itens_PC
    (DATA_PC, col B) para as datas de ultima evidencia.

BLOCO A (marcos), BLOCO B (cobertura das fontes), BLOCO C (decisao) + legenda
FISCAL / GCC / AUTOMATICO / PROJECAO. Um unico campo de entrada (GCC, opcional):
"Data da analise" (B4). Tudo o mais e formula (AUTOMATICO) ou diagnostico de
PROJECAO. Projecao e sempre estimativa: nunca vira fato observado.

Cores REUTILIZADAS do template (nao inventa tonalidade para categoria existente):
  * cabecalho/banda        = formato de CONTROLE!A6 (azul FF1F4E79, fonte branca);
  * rotulo (coluna A)      = formato de CONTROLE!A7 (azul claro);
  * valor automatico (B)   = formato de CONTROLE!B9 (azul muito claro / formula);
  * entrada GCC (B4)       = formato de CONTROLE!B3 (amarelo de entrada).
Unica categoria NOVA (permitida pela etapa p/ projecao): PROJECAO = laranja
claro FCE4D6 — tonalidade clara e distinta, sem confundir com entrada manual.

Gravador unico: Microsoft Excel real via COM (copia temporaria; promove so se
salvar sem erros de formula). FAIL-CLOSED: recusa reaplicacao. NAO usa openpyxl
para salvar.
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
# MAX das competencias financeiras (col A) e das datas de PC (col B).
FIN_ATE = ('=IF(COUNT(financeiro!$A$2:$A$200)=0,"",MAX(financeiro!$A$2:$A$200))')
PC_ATE = ('=IF(COUNT(itens_PC!$B$2:$B$200)=0,"",MAX(itens_PC!$B$2:$B$200))')

# Posterioridade (fonte alem da posicao fisica). B11 = posicao fisica ate,
# B12 = financeiro ate, B13 = PC ate.
POST_FIN = '(AND($B$12<>"",$B$11<>"",$B$12>$B$11))'
POST_PC = '(AND($B$13<>"",$B$11<>"",$B$13>$B$11))'

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
    if ABA in nomes:
        raise ValueError("Aba cobertura_temporal ja existe; bloco ja aplicado?")
    for aba in ("CONTROLE", "parametros", "financeiro", "itens_PC",
                "posicao_referencia", "RESULTADOS"):
        if aba not in nomes:
            raise ValueError(f"Aba canonica ausente: {aba}")
    pr = wb.Worksheets(PR)
    if "DATA DA POSICAO DE REFERENCIA" not in str(pr.Range("H5").Value or ""):
        raise ValueError("posicao_referencia!H5 nao e o painel de referencia esperado.")


def _fmt(ws, modelo, origem_addr, destino_addr, excel) -> None:
    modelo.Range(origem_addr).Copy()
    ws.Range(destino_addr).PasteSpecial(XL_PASTE_FORMATS)
    excel.CutCopyMode = False


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
        (10, "BLOCO B - COBERTURA DAS FONTES", None, None, "banner"),
        (11, "Posicao fisica conhecida ate", f"={I_DATA_REF}", FMT_DATA, "auto"),
        (12, "Financeiro conhecido/completo ate", FIN_ATE, FMT_DATA, "auto"),
        (13, "PC conhecido/completo ate", PC_ATE, FMT_DATA, "auto"),
        (14, "Ultima evidencia concreta (geral)",
         '=IF(COUNT($B$11:$B$13)=0,"",MAX($B$11:$B$13))', FMT_DATA, "auto"),
        (15, "Projecao necessaria a partir de",
         '=IF(AND(ISNUMBER($B$4),$B$14<>"",$B$4>$B$14),$B$14,"")', FMT_DATA, "proj"),
        (17, "BLOCO C - DECISAO", None, None, "banner"),
        (18, "Modo temporal", MODO, None, "auto"),
        (19, "Fonte principal",
         '=IF($B$12<>"","Financeiro",IF($B$13<>"","PC",""))', None, "auto"),
        (20, "Fontes de conferencia",
         '=IF(AND($B$12<>"",$B$13<>""),"PC","")', None, "auto"),
        (21, "Posicao observada (data)", f"={I_DATA_REF}", FMT_DATA, "auto"),
        (22, "Posicao observada (origem)", f"={I_ORIGEM}", None, "auto"),
        (23, "Posicao projetada",
         f'=IF($B$15="","NAO PROJETADO (posicao observada mantida)",'
         f'"ESTIMATIVA a partir de "&{_dstr("$B$15")}&'
         '" - nao observada, nao cria retroativo a pagar")', None, "proj"),
    ]

    for lin, rot, formula, fmt, cat in linhas:
        cel_rot = ws.Range(f"A{lin}")
        cel_val = ws.Range(f"B{lin}")
        if cat == "banner":
            ws.Range(f"A{lin}:C{lin}").Merge()
            _fmt(ws, ctrl, "A6", f"A{lin}", excel)
            cel_rot.Value = rot
            continue
        # rotulo (coluna A) — azul claro de CONTROLE!A7
        _fmt(ws, ctrl, "A7", f"A{lin}", excel)
        cel_rot.Value = rot
        # valor (coluna B)
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

    # Entrada GCC (B4): data opcional, validacao de data.
    b4 = ws.Range("B4")
    b4.NumberFormatLocal = FMT_DATA
    val = b4.Validation
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
    val.InputTitle = "Data da analise (GCC, opcional)"
    val.InputMessage = ("Data de referencia da analise. Opcional; se vazia, nao "
                        "ha diagnostico de projecao. NAO altera a posicao observada.")
    val.ErrorMessage = "Informe uma data valida (dd/mm/aaaa)."

    # Legenda — reutiliza a banda de banner e amostra as cores por categoria.
    _fmt(ws, ctrl, "A6", "A25", excel)
    ws.Range("A25:C25").Merge()
    ws.Range("A25").Value = "LEGENDA - QUEM PREENCHE O QUE"
    legenda = [
        (26, "FISCAL", "B3", "Amarelo: itens_Remanesc, posicao_referencia (QTD_REM_ATUAL) e CONTROLE!B3."),
        (27, "GCC", "B3", "Amarelo: Data da analise (B4) desta aba; confirmacao de cobertura."),
        (28, "AUTOMATICO", "B9", "Derivado de CONTROLE/parametros/posicao_referencia/financeiro/itens_PC."),
        (29, "PROJECAO", None, "Laranja claro: estimativa (nunca fato observado; nao cria retroativo)."),
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

    ws.Columns("A").ColumnWidth = 42
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
        _criar_aba(wb, excel)
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
