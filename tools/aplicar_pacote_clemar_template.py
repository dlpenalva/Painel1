"""Rodada coordenada Excel COM — pacote pos-CLEMAR sobre o template oficial.

Aplica, numa unica sessao COM (Microsoft Excel real; NAO usa openpyxl para
salvar), todas as alteracoes estruturais/visuais ja aprovadas:

  CONTROLE      : B1 = SELECIONE AQUI + dropdown de 4 opcoes (fonte unica).
  RESULTADOS    : B4 derivado/read-only de CONTROLE!B1 (elimina 2o seletor).
  parametros    : linhas 5:6 no cinza estrutural (remove distincao C3/C4);
                  destaque vermelho de E e texto de G ficam a cargo do gerador.
  posicao_contratual : col B (VU_ORIGINAL) como moeda R$.
  historico_VU  : S:W (QTD_REM_AJUSTADA_C0:C4) com preenchimento suave distinto.
  posicao_referencia : col I alinhada a esquerda e com largura adequada.
  comparativo_VTA : nova aba de referencia (Bloco 1 precos originais; Bloco 2
                  VTA aplicado x contrato integralmente reajustado, fator
                  historico integral; alerta NAO E VTA).
  cobertura_temporal : reaplica a aba com a formula corrigida da evidencia
                  Financeira e a nova nomenclatura (via o tool ja atualizado).

Idempotente. Preserva formulas/validacoes/estilos/named ranges das demais abas.
"""
from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

sys.path.insert(0, str(Path(__file__).resolve().parent))  # tools/ no path
import aplicar_cobertura_temporal_template as COB  # reutiliza aba corrigida

XL_VALIDATE_LIST = 3
XL_VALID_ALERT_STOP = 1
XL_BETWEEN = 1
XL_LEFT = -4131
XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105
XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16

CINZA = 0xEDEDED           # BGR de EDEDED (cinza estrutural das linhas de dados)
SUAVE_REMANESC = 0xDAEFE2  # BGR de E2EFDA (verde muito claro, distinto)
MOEDA_BR = "R$ #.##0,00"   # pt-BR (NumberFormatLocal: '.' milhar, ',' decimal)

OPCOES_B1 = ["SELECIONE AQUI", "Principal (Financeiro)",
             "PC (Pedidos de Compra)", "Itens Consumidos"]

# Opcoes de protecao preservadas ao re-proteger a aba CONTROLE.
_ALLOW = [
    "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows",
    "AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks",
    "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting",
    "AllowFiltering", "AllowUsingPivotTables",
]
FORMULA_B4 = (
    '=IF(OR(CONTROLE!$B$1="SELECIONE AQUI",CONTROLE!$B$1=""),'
    '"SELECIONE O METODO EM CONTROLE!B1",CONTROLE!$B$1)'
)

ABA_COMP = "comparativo_VTA"
ALERTA_COMP = (
    "O valor integralmente reajustado e apenas um cenario comparativo. Ele "
    "pressupoe que todo o contrato original permanecesse sujeito ao fator "
    "acumulado final, desconsiderando a execucao ocorrida em ciclos distintos "
    "e a reducao do saldo remanescente. Nao utilizar este valor como VTA, "
    "retroativo ou base de pagamento."
)


def _validar_layout_itens_pc(wb) -> None:
    """Falha fechado se o quadro lateral nao estiver no layout PC-UX-1."""
    ws = wb.Worksheets("itens_PC")
    esperado = (
        "TODOS OS PCs CADASTRADOS POR CICLO",
        "CICLO",
        "C0",
        "VALOR DOS PCs COM FATOR DO CICLO",
    )
    obtido = (
        ws.Range("M1").Value,
        ws.Range("M2").Value,
        ws.Range("M3").Value,
        ws.Range("P2").Value,
    )
    if obtido != esperado:
        raise RuntimeError(
            "Layout itens_PC incompativel com PC-UX-1; aplicacao cancelada "
            f"antes de regravar o template: {obtido!r}"
        )


def _dv_lista(cel, itens: list[str], sep: str) -> None:
    val = cel.Validation
    try:
        val.Delete()
    except Exception:
        pass
    # win32com late-bind: Add exige args posicionais (Type, AlertStyle, Operator,
    # Formula1) — a forma por keyword sem Operator falha (0x800A0BEC).
    val.Add(XL_VALIDATE_LIST, XL_VALID_ALERT_STOP, XL_BETWEEN, sep.join(itens))
    val.IgnoreBlank = True
    val.InCellDropdown = True


def _controle(wb, excel) -> None:
    c = wb.Worksheets("CONTROLE")
    protegida = bool(c.ProtectContents)
    opts = {a: getattr(c.Protection, a) for a in _ALLOW} if protegida else {}
    if protegida:
        c.Unprotect()
    sep = excel.International[4]   # xlListSeparator (pt-BR: ';')
    _dv_lista(c.Range("B1"), OPCOES_B1, sep)
    c.Range("B1").Value = "SELECIONE AQUI"
    # E3:E6 / G3:G6 sem negrito (padrao uniforme).
    for addr in ("E3:E6", "G3:G6"):
        c.Range(addr).Font.Bold = False
    if protegida:
        # Re-protege preservando as opcoes originais (DrawingObjects/Contents/
        # Scenarios padrao; sem senha, como no template original).
        c.Protect(DrawingObjects=True, Contents=True, Scenarios=True, **opts)


def _resultados(wb) -> None:
    r = wb.Worksheets("RESULTADOS")
    b4 = r.Range("B4")
    try:
        b4.Validation.Delete()   # remove o 2o seletor independente
    except Exception:
        pass
    b4.Formula = FORMULA_B4      # derivado/read-only de CONTROLE!B1


def _resultados_vta_pc(wb) -> None:
    """VTA-PC oficial materializado no XLS (RESULTADOS!B26), espelhando o motor
    Python _composicao_vta_pc. Bloco auxiliar auditavel em RESULTADOS!S:T; so
    compoe quando CONTROLE!B1 = 'PC (Pedidos de Compra)'. Sem valor hardcoded;
    recalcula sozinho a partir das abas oficiais. Base insuficiente -> CALCULO
    MANUAL REQUERIDO (nao fabrica zero como VTA valido).

    Componentes (mesmo corte temporal; sem dupla contagem):
      C0  = PC C0 (itens_PC!P3) se houver; senao movimentacao fisica
            (QTD_REM_AJUSTADA_C0 - _C1) x VU_C0.
      PCn = itens_PC.VALOR_ATUALIZADO_TOTAL por ciclo (col P), ciclos
            1 <= n < vigente (exclui vigente/posterior).
      Rem = item a item, ROUND(ROUND(VU_Cvig,2) x QTD_REM_AJUSTADA_Cvig, 2).
    """
    r = wb.Worksheets("RESULTADOS")
    # Helpers POR ITEM (colunas X:Y, linhas 2..201) — formulas ESCALARES por
    # linha (sem depender de array em SUMPRODUCT, que aqui nao propaga
    # IFERROR/MID/CHOOSE). Alinhadas por ITEM a posicao_contratual/historico_VU.
    r.Range("X1").Value = "VTA-PC aux: C0 fisico por item ((QREM_C0-QREM_C1)*VU_C0)"
    r.Range("Y1").Value = "VTA-PC aux: remanescente por item (ROUND(VU_vig,2)*QREM_vig)"
    r.Range("X2:X201").Formula = (
        '=IF(posicao_contratual!$A2="",0,'
        '(posicao_contratual!$G2-posicao_contratual!$K2)*historico_VU!$C2)')
    r.Range("Y2:Y201").Formula = (
        '=IF(OR(posicao_contratual!$A2="",$T$20=""),0,'
        'ROUND(ROUND(CHOOSE($T$20+1,historico_VU!$C2,historico_VU!$D2,'
        'historico_VU!$E2,historico_VU!$F2,historico_VU!$G2),2)'
        '*CHOOSE($T$20+1,posicao_contratual!$G2,posicao_contratual!$K2,'
        'posicao_contratual!$O2,posicao_contratual!$S2,posicao_contratual!$W2),2))')
    aux = [
        ("S19", "VTA-PC (metodo PC) - auxiliar (espelha o motor Python)"),
        ("S20", "Ciclo vigente (numero)"),
        ("T20", '=IFERROR(VALUE(MID(CONTROLE!$B$2,2,3)),"")'),
        ("S21", "C0 executado (PC C0 ou movimentacao fisica x VU_C0)"),
        # PC C0 se houver (itens_PC!P3); senao soma dos C0 fisicos por item.
        ("T21", '=IF(itens_PC!$P$3>0,itens_PC!$P$3,SUM($X$2:$X$201))'),
        ("S22", "Execucao PC (ciclos 1..vigente-1, VALOR_ATUALIZADO)"),
        # ROW(P3:P7)-3 = numero do ciclo (0..4); soma P dos ciclos 1..vigente-1.
        ("T22", '=SUMPRODUCT((ROW(itens_PC!$P$3:$P$7)-3>=1)*(ROW(itens_PC!$P$3:$P$7)-3<$T$20)*itens_PC!$P$3:$P$7)'),
        ("S23", "Remanescente fisico atualizado (item a item, ciclo vigente)"),
        ("T23", '=SUM($Y$2:$Y$201)'),
        ("S24", "Base itemizada presente?"),
        ("T24", '=IF(COUNTIF(posicao_contratual!$A$2:$A$201,"<>")=0,0,1)'),
        ("S25", "VTA-PC (metodo aplicado)"),
        ("T25", '=IF(OR($T$24=0,$T$20=""),"CALCULO MANUAL REQUERIDO",ROUND($T$21+$T$22+$T$23,2))'),
    ]
    for addr, val in aux:
        cel = r.Range(addr)
        if isinstance(val, str) and val.startswith("="):
            cel.Formula = val
        else:
            cel.Value = val
    for addr in ("T21", "T22", "T23", "T25"):
        r.Range(addr).NumberFormatLocal = MOEDA_BR
    # B26 (VTA FINAL): caminho PC via bloco auxiliar; preserva Principal/Itens
    # Consumidos e o override manual (B24/B25) e os aditivos ($N$263).
    r.Range("B26").Formula = (
        '=IF(AND(B24<>"",B25<>""),"",IF(ISNUMBER(B25),B25,'
        'IF(CONTROLE!$B$1="PC (Pedidos de Compra)",'
        'IF($T$25="CALCULO MANUAL REQUERIDO","",'
        'ROUND($T$25+IF(ISNUMBER(B24),B24,0),2)),'
        'IF(OR(B23="",AND(B24<>"",NOT(ISNUMBER(B24)))),"",'
        'ROUND(B23+IF(ISNUMBER(B24),B24,0)+IF(ISNUMBER($N$263),$N$263,0),2)))))'
    )


def _parametros(wb) -> None:
    p = wb.Worksheets("parametros")
    # Linhas 5:6 no mesmo cinza estrutural das demais (sem distincao C3/C4).
    # O vermelho de E (fora da apuracao) e o texto de G sao aplicados pelo
    # gerador por contrato; aqui so uniformizamos o fundo estatico do template.
    p.Range("A5:G6").Interior.Color = CINZA


def _posicao_contratual(wb) -> None:
    pc = wb.Worksheets("posicao_contratual")
    pc.Range("B2:B201").NumberFormatLocal = MOEDA_BR   # VU_ORIGINAL como moeda R$


def _historico_vu(wb) -> None:
    h = wb.Worksheets("historico_VU")
    # S:W = QTD_REM_AJUSTADA_C0:C4 — preenchimento suave distinto (so visual).
    h.Range("S2:W201").Interior.Color = SUAVE_REMANESC


def _posicao_referencia(wb) -> None:
    pr = wb.Worksheets("posicao_referencia")
    col_i = pr.Range("I1:I40")
    col_i.HorizontalAlignment = XL_LEFT
    col_i.WrapText = False
    pr.Columns("I").ColumnWidth = 26


def _comparativo_vta(wb, excel) -> None:
    # (re)cria a aba de referencia.
    for ws in wb.Worksheets:
        if ws.Name == ABA_COMP:
            alertas = excel.DisplayAlerts
            excel.DisplayAlerts = False
            ws.Delete()
            excel.DisplayAlerts = alertas
            break
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = ABA_COMP

    ws.Range("A1").Value = (
        "comparativo_VTA - REFERENCIA (nao aparece na web; nao e VTA de pagamento)")
    ws.Range("A2").Value = (
        "BLOCO 1 - EVOLUCAO DO REMANESCENTE A PRECOS ORIGINAIS (sem reajuste)")
    hdr = ["ITEM", "VU_ORIGINAL", "C0", "C1", "C2", "C3", "C4"]
    for i, h in enumerate(hdr):
        ws.Cells(3, i + 1).Value = h
    # Formulas relativas (Excel ajusta a linha automaticamente na faixa).
    # comparativo linha 4 -> posicao_contratual linha 2 (offset -2).
    P = "posicao_contratual"
    ws.Range("A4:A203").Formula = f'=IF({P}!A2="","",{P}!A2)'
    ws.Range("B4:B203").Formula = f'=IF({P}!A2="","",{P}!B2)'
    # VU_ORIGINAL x QTD_REM_AJUSTADA_Cn (G,K,O,S,W). Ciclo ausente
    # (QTD_REM_AJUSTADA vazia/string) ou VU nao numerico -> vazio (nunca
    # #VALUE!). QTD=0 NUMERICO -> R$ 0,00 (zero != ausencia).
    for col, qcol in (("C", "G"), ("D", "K"), ("E", "O"), ("F", "S"), ("G", "W")):
        ws.Range(f"{col}4:{col}203").Formula = (
            f'=IF(OR({P}!A2="",NOT(ISNUMBER({P}!B2)),NOT(ISNUMBER({P}!{qcol}2))),'
            f'"",ROUND({P}!B2*{P}!{qcol}2,2))')
    ws.Cells(204, 1).Value = "TOTAL"
    for col in ("C", "D", "E", "F", "G"):
        ws.Range(f"{col}204").Formula = f"=ROUND(SUM({col}4:{col}203),2)"
    ws.Range("B4:G204").NumberFormatLocal = MOEDA_BR

    ws.Range("A206").Value = (
        "BLOCO 2 - COMPARACAO DE REFERENCIA - NAO UTILIZAR COMO VTA OU "
        "VALOR DE PAGAMENTO")
    ws.Range("A207").Value = "VTA pelo metodo aplicado"
    ws.Range("B207").Formula = "=RESULTADOS!$B$26"
    ws.Range("A208").Value = (
        "Contrato original integralmente reajustado ate o ciclo atual - "
        "REFERENCIA APENAS")
    # valor_original (VU_ORIGINAL x QTD_BASE_ORIGINAL) x fator historico integral.
    ws.Range("B208").Formula = (
        f'=IFERROR(ROUND(SUMPRODUCT({P}!$B$2:$B$201,{P}!$C$2:$C$201)'
        '*CONTROLE!$B$11,2),"")')
    ws.Range("B207:B208").NumberFormatLocal = MOEDA_BR
    ws.Range("A210").Value = ALERTA_COMP

    ws.Columns("A").ColumnWidth = 40
    for col in ("B", "C", "D", "E", "F", "G"):
        ws.Columns(col).ColumnWidth = 16


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
    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_pacote_"))
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
        COB._validar_layout(wb)
        _validar_layout_itens_pc(wb)
        _controle(wb, excel)
        _resultados(wb)
        _resultados_vta_pc(wb)
        _parametros(wb)
        _posicao_contratual(wb)
        _historico_vu(wb)
        _posicao_referencia(wb)
        _comparativo_vta(wb, excel)
        # cobertura_temporal: reaplica com formula corrigida + nova nomenclatura.
        COB._remover_aba_existente(wb, excel)
        COB._criar_aba(wb, excel)
        # A ordem fisica das abas segue o que o Excel COM produz (comparativo_VTA
        # como primeira aba; a aba ativa original e preservada e a aba nao aparece
        # na web). ABAS_COLETA_OFICIAL reflete essa ordem. Reordenar por COM (14
        # Move) mostrou-se instavel; nao vale o risco para um artefato cosmetico.
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


if __name__ == "__main__":
    main()
