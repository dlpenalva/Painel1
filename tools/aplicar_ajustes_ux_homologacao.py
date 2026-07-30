"""Ajustes de UX/validacao/seguranca do template oficial (rodada de homologacao).

Aplicador UNICO e cirurgico via Microsoft Excel real (COM). NAO usa openpyxl para
gravar (evita a corrupcao de estilos por reserializacao). Preserva formulas,
estilos nao relacionados, conditional formatting, validacoes nao corrigidas,
nomes definidos, protecao, print areas e a estrutura OOXML (sem repairLoad).

Alteracoes (todas cirurgicas; NAO tocam matematica homologada):

  1. parametros!G12:G15 (REAJUSTE ANTERIOR JA FORMALIZADO?): corrige o dropdown
     quebrado. Cria intervalo tecnico parametros!T2:T3 = {Sim, Nao} (coluna T
     oculta, nunca veryHidden) + nome definido OPCOES_SIM_NAO, e revalida
     G12:G15 como lista referenciando o nome (robusto ao separador regional).
     Vazio permitido. Valores "Sim"/"Nao" (ASCII) — casam com as formulas do
     bloco EIXO 5 (parametros!Gn="Sim"/"Nao"), sem alterar matematica.

  2. CONTROLE!C1: remove o texto explicativo (sem dependencia funcional).

  3. itens_Consumidos!Q2:Q200: remove o preenchimento estatico (a aba nao possui
     conditional formatting; so ha fill fixo). Alertas condicionais, se houvesse,
     nao seriam tocados.

  4. aditivos!M2:M200 (CHECK): mensagem de item ausente mais explicita
     ("NOVO ITEM NAO CADASTRADO..."). O check de existencia (COUNTIF em
     itens_Remanesc) ja existe; aqui apenas clareamos o texto.

  5. posicao_referencia!H1:I8: leiaute suave do painel (fundo dos rotulos e dos
     valores + borda externa discreta). Nao pinta colunas inteiras; nao toca
     formulas/valores/hidden helpers.

  6. RESULTADOS:
     - A17 (EIXO 2): padroniza o cabecalho a identidade dos demais EIXOS (copia
       o formato de A30) e reescreve o texto sem a referencia obsoleta a "linha 62".
     - B4: garante destaque de dropdown (somente B4).
     - A48/A49: notas metodologicas movidas para o bloco NOTAS TECNICAS ao fim da
       aba (A281/A282), SEM deslocar linhas (limpa o conteudo original e recria no
       fim) — preserva todas as referencias absolutas (EIXO 5 em 256, POSICAO em
       267 etc.). A48/A49 ficam vazias.

FAIL-CLOSED: recusa reaplicacao (nome OPCOES_SIM_NAO ja presente). Gravador unico:
Excel real; promove so se salvar sem erros de formula.
NAO altera B23/B25/B26, a matematica, o template legado nem o Python de producao.
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
XL_VALIDATE_LIST = 3
XL_VALID_ALERT_STOP = 1
XL_NONE = -4142
XL_CONTINUOUS = 1
XL_THIN = 2
XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105

NOME_OPCOES = "OPCOES_SIM_NAO"
R_NOTAS = 280  # bloco NOTAS TECNICAS ao fim de RESULTADOS (apos POSICAO em 277)


def _validar_layout(wb) -> None:
    nomes = [n.Name for n in wb.Names]
    if NOME_OPCOES in nomes:
        raise ValueError(f"{NOME_OPCOES} ja existe; ajustes de UX ja aplicados?")
    for aba in ("CONTROLE", "parametros", "itens_Consumidos", "aditivos",
                "posicao_referencia", "RESULTADOS"):
        if aba not in [ws.Name for ws in wb.Worksheets]:
            raise ValueError(f"Aba ausente: {aba}")


def _fix_dropdown_sim_nao(wb) -> None:
    ws = wb.Worksheets("parametros")
    ws.Range("T2").Value = "Sim"
    ws.Range("T3").Value = "Nao"
    ws.Columns("T").Hidden = True  # tecnica; reexibivel, nunca veryHidden
    wb.Names.Add(Name=NOME_OPCOES, RefersTo="=parametros!$T$2:$T$3")
    val = ws.Range("G12:G15").Validation
    try:
        val.Delete()
    except Exception:
        pass
    val.Add(Type=XL_VALIDATE_LIST, AlertStyle=XL_VALID_ALERT_STOP,
            Formula1="=" + NOME_OPCOES)
    val.IgnoreBlank = True
    val.InCellDropdown = True
    val.ErrorMessage = "Selecione Sim ou Nao (ou deixe vazio se ainda nao comprovado)."


_PROT_FLAGS = (
    "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows",
    "AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks",
    "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting",
    "AllowFiltering", "AllowUsingPivotTables",
)


def _com_protecao(ws):
    """Captura opcoes de protecao e desprotege (sem senha). Retorna estado + sel."""
    if not bool(ws.ProtectContents):
        return None, None
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


def _restaurar_protecao(ws, estado, sel) -> None:
    if estado is None:
        return
    ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True, **estado)
    if sel is not None:
        try:
            ws.EnableSelection = sel
        except Exception:
            pass


def _remover_c1(wb) -> None:
    ws = wb.Worksheets("CONTROLE")
    estado, sel = _com_protecao(ws)
    ws.Range("C1").ClearContents()
    _restaurar_protecao(ws, estado, sel)


def _limpar_fill_q(wb) -> None:
    ws = wb.Worksheets("itens_Consumidos")
    ws.Range("Q2:Q200").Interior.Pattern = XL_NONE


def _msg_aditivos(wb) -> None:
    ws = wb.Worksheets("aditivos")
    f = str(ws.Range("M2").Formula)
    novo = f.replace(
        '"ALERTA: ITEM_AUSENTE"',
        '"NOVO ITEM NAO CADASTRADO - CADASTRAR EM itens_Remanesc (QTD BASE 0 + VU)"')
    if novo != f:
        ws.Range("M2:M200").Formula = novo


def _estilizar_painel(wb) -> None:
    ws = wb.Worksheets("posicao_referencia")
    ws.Range("H1:H8").Interior.Color = 0xF2F2F2   # rotulos: cinza muito suave (BGR)
    ws.Range("I1:I8").Interior.Color = 0xFBFBFB   # valores: quase branco
    ws.Range("H1:I8").BorderAround(LineStyle=XL_CONTINUOUS, Weight=XL_THIN)


def _ajustar_resultados(wb, excel) -> None:
    """Ajustes de RESULTADOS com FORMATACAO DIRETA (sem Copy/PasteSpecial entre
    celulas mescladas e sem novas mesclagens — essas operacoes corrompem o
    arquivo, exigindo reparo do Excel na reabertura)."""
    ws = wb.Worksheets("RESULTADOS")
    MAROON = 0x38158A   # FF8A1538 em BGR (identidade dos demais EIXOS)
    BRANCO = 0xFFFFFF
    # A17 (EIXO 2): identidade maroon dos demais EIXOS + texto sem "linha 62".
    a17 = ws.Range("A17")
    a17.Interior.Color = MAROON
    a17.Font.Color = BRANCO
    a17.Font.Bold = True
    a17.Value = ("EIXO 2 - VU REAJUSTADO: evolucao cronologica item a item "
                 "disponivel na memoria detalhada abaixo.")
    # B4: destaque de dropdown (somente B4).
    ws.Range("B4").Interior.Color = 0xB2E7F7  # dourado de input (FFF7E7B2 em BGR)
    ws.Range("B4").Font.Bold = True
    # A48/A49 -> NOTAS TECNICAS ao fim (sem deslocar linhas; sem mesclar). As notas
    # ficam em celula unica e transbordam visualmente sobre B:J (vazias).
    nota48 = ws.Range("A48").Value
    nota49 = ws.Range("A49").Value
    hdr = ws.Range(f"A{R_NOTAS}")
    hdr.Value = "NOTAS TECNICAS"
    hdr.Interior.Color = MAROON
    hdr.Font.Color = BRANCO
    hdr.Font.Bold = True
    ws.Range(f"A{R_NOTAS + 1}").Value = nota48
    ws.Range(f"A{R_NOTAS + 2}").Value = nota49
    ws.Range("A48:J48").ClearContents()
    ws.Range("A49:J49").ClearContents()


def _verificar_sem_erros(wb) -> None:
    import pywintypes
    problemas = []
    for ws in wb.Worksheets:
        for tipo in (XL_CELLTYPE_FORMULAS, XL_CELLTYPE_CONSTANTS):
            try:
                cels = ws.UsedRange.SpecialCells(tipo, XL_ERRORS)
                problemas.append(f"{ws.Name}!{cels.Address}")
            except pywintypes.com_error:
                continue
    if problemas:
        raise RuntimeError(f"Erros de formula apos recalculo: {problemas}")


def aplicar(origem: Path, destino: Path) -> None:
    origem, destino = Path(origem), Path(destino)
    if not origem.is_file():
        raise FileNotFoundError(origem)
    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_uxhml_"))
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
        _fix_dropdown_sim_nao(wb)
        _remover_c1(wb)
        _limpar_fill_q(wb)
        _msg_aditivos(wb)
        _estilizar_painel(wb)
        _ajustar_resultados(wb, excel)
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
