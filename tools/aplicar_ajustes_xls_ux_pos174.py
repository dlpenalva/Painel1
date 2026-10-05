# -*- coding: utf-8 -*-
"""AJUSTES-XLS-UX pos-PR #174: tres ajustes VISUAIS focais no template oficial.

Somente apresentacao. Nenhuma formula economica, validacao, nome definido ou
endereco existente e alterado; cada frente escreve apenas na sua allowlist e
confere, antes, que a area de destino esta livre.

  parametros  aviso dinamico de percentuais HISTORICOS ausentes em A8:F8
              (ciclos C1..C(vigente-1) sem E; o vigente vem de CONTROLE!B2,
              a mesma referencia canonica de CONTROLE!B11) e nota curta em A17
              sob a MEMORIA DO FATOR. Textos acentuados ficam em celulas
              constantes de MEMORIA_RESULTADOS!AF41:AG42 (formula ASCII).
  consumidos  identidade visual do bloco opcional itens_Consumidos!X1:AG13:
              Z/AA com cor de entrada manual, demais colunas com aparencia de
              campo automatico, orientacao com quebra de linha.
  detalhe     RESULTADOS_DETALHE passa a hidden (nunca veryHidden) e o
              hiperlink de RESULTADOS!D31 vira texto normal.

A memoria IST (parametros!J:R) e gravada em Python na geracao da Coleta; o
destaque da fronteira entre ciclos fica em `_memoria_calculo`.

REGRA ZERO CORRUPCAO XLSX: so Excel COM (openpyxl destroi a CF x14 do
template); formulas em ingles, ASCII e com parenteses balanceados.

Uso:  python tools/aplicar_ajustes_xls_ux_pos174.py --frente F [--frente F] [xlsx]
"""
from __future__ import annotations

import argparse
import gc
import math
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

XL_EXPRESSION = 2
XL_LEFT, XL_CENTER = -4131, -4108
XL_VCENTER = -4108
XL_CONTINUOUS, XL_THIN = 1, 2
XL_EDGES = (7, 8, 9, 10, 11, 12)  # esquerda, topo, base, direita, internas V/H
XL_SHEET_HIDDEN = 0
XL_SHEET_VERY_HIDDEN = 2
XL_CALCULO_MANUAL, XL_CALCULO_AUTOMATICO = -4135, -4105

# ------------------------------------------------------------- parametros
MEM = "MEMORIA_RESULTADOS"
TEXTOS_HISTORICO = (
    (41, "HIST_INC_1", "HISTÓRICO INCOMPLETO — informe o percentual de "),
    (42, "HIST_INC_N", "HISTÓRICO INCOMPLETO — informe os percentuais de "),
)
CELULA_AVISO = "A8"
FAIXA_AVISO = "A8:F8"
CELULA_NOTA_FATOR = "A17"
NOTA_FATOR = "Este quadro mostra apenas os ciclos computados nesta apuração."


def formula_aviso_historico() -> str:
    """Lista C1..C(vigente-1) sem percentual numerico em parametros!E.

    Historico = ciclos ANTERIORES ao vigente (CONTROLE!B2). C0 nunca e pedido;
    o proprio vigente e os futuros (C > vigente, que tambem tem datas
    derivadas e E vazio) nunca entram no aviso. C4 nunca e historico.
    """
    vig = "IFERROR(VALUE(MID(TRIM(CONTROLE!$B$2),2,1)),0)"
    a = f"AND({vig}>1,NOT(ISNUMBER($E$3)))"
    b = f"AND({vig}>2,NOT(ISNUMBER($E$4)))"
    c = f"AND({vig}>3,NOT(ISNUMBER($E$5)))"
    nomes = (
        f'IF({a},"C1","")'
        f'&IF({b},IF({a},IF({c},", "," e "),"")&"C2","")'
        f'&IF({c},IF(OR({a},{b})," e ","")&"C3","")'
    )
    quantidade = f"(({a})+({b})+({c}))"
    return (
        f'=IF({nomes}="","",IF({quantidade}=1,{MEM}!$AG$41,{MEM}!$AG$42)'
        f'&{nomes}&".")'
    )


# ------------------------------------------------------------- consumidos
CONS = "itens_Consumidos"
COLUNAS_ENTRADA = ("Z", "AA")
COLUNAS_AUTOMATICAS = ("X", "Y", "AB", "AC", "AD", "AE", "AF", "AG")
TITULO_AJUSTE = "AJUSTE OPCIONAL — preencher somente quando necessário (valor pago / glosa)"
ENTRADAS_AJUSTE = (
    "Entradas manuais: somente as colunas amarelas AJUSTE_TIPO (Z) e "
    "AJUSTE_VALOR_INFORMADO (AA); as demais colunas são automáticas. "
)
TEXTO_X9_BASE = (
    "Preencha somente quando o valor pago considerado na execucao do ciclo "
    "diferir do valor calculado pelos itens consumidos."
)
COR_BLOCO_CAB = "FFE4DFEC"     # lilas claro: identidade do bloco opcional
COR_BLOCO_TEXTO = "FF3F3151"
COR_ENTRADA = "FFFFF2CC"       # ambar claro: entrada manual (mesma do workbook)
COR_ENTRADA_CAB = "FFFFE699"
COR_AUTOMATICA = "FFF2F2F2"
COR_TEXTO_AUTO = "FF595959"
COR_BORDA = "FFBFBFBF"
COR_ORIENTACAO = "FFF7F5FA"

# ------------------------------------------------------------- detalhe
EXE, DET = "RESULTADOS", "RESULTADOS_DETALHE"
TEXTO_AJUSTES_D31 = (
    "Na aba técnica RESULTADOS_DETALHE (bloco 5), oculta por padrão: "
    "use Reexibir no Excel para acessá-la."
)
SUBTITULO_B3 = (
    "Síntese executiva da apuração do reajuste. A memória técnica, as "
    "conferências e os ajustes manuais estão na aba RESULTADOS_DETALHE "
    "(oculta por padrão; use Reexibir no Excel)."
)


def validar_ascii_e_parenteses(formulas: dict[str, str]) -> None:
    for endereco, formula in formulas.items():
        if any(ord(ch) > 127 for ch in formula):
            raise ValueError(f"Formula nao-ASCII em {endereco}: {formula!r}")
        nivel, aspas = 0, False
        for ch in formula:
            if ch == '"':
                aspas = not aspas
            elif not aspas and ch == "(":
                nivel += 1
            elif not aspas and ch == ")":
                nivel -= 1
                if nivel < 0:
                    break
        if nivel != 0 or aspas:
            raise ValueError(f"Parenteses/aspas desbalanceados em {endereco}")


# ------------------------------------------------------------- COM helpers
def _bgr(rgb: str) -> int:
    rgb = rgb[-6:]
    r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
    return r + (g << 8) + (b << 16)


def _fechar(wb, tentativas: int = 10) -> None:
    for _ in range(tentativas):
        try:
            wb.Close(SaveChanges=False)
            return
        except Exception:
            time.sleep(1.0)


def _preencher(rng, rgb: str) -> None:
    rng.Interior.Pattern = 1
    rng.Interior.Color = _bgr(rgb)


def _fonte(rng, *, tamanho=None, negrito=None, cor=None, italico=None) -> None:
    if tamanho is not None:
        rng.Font.Size = tamanho
    if negrito is not None:
        rng.Font.Bold = negrito
    if italico is not None:
        rng.Font.Italic = italico
    if cor is not None:
        rng.Font.Color = _bgr(cor)


def _alinhar(rng, h=XL_LEFT, v=XL_VCENTER, quebra=False, recuo=0) -> None:
    rng.HorizontalAlignment = h
    rng.VerticalAlignment = v
    rng.WrapText = quebra
    rng.IndentLevel = recuo


def _bordas(rng, rgb: str) -> None:
    for indice in XL_EDGES:
        b = rng.Borders(indice)
        b.LineStyle = XL_CONTINUOUS
        b.Weight = XL_THIN
        b.Color = _bgr(rgb)


def _mesclar(ws, endereco: str) -> None:
    rng = ws.Range(endereco)
    rng.UnMerge()
    rng.Merge()


def _local(rng, formula: str) -> str:
    """Formula de CF no idioma da interface (Excel pt-BR usa ';' e nomes locais)."""
    ws = rng.Worksheet
    aux = ws.Cells(rng.Cells(1, 1).Row, 60)
    if aux.Formula not in ("", None):
        raise RuntimeError(f"Celula auxiliar {aux.Address} ocupada em {ws.Name}")
    aux.Formula = formula
    local = aux.FormulaLocal
    aux.ClearContents()
    return local


def _exigir_vazias(ws, endereco: str) -> None:
    for cel in ws.Range(endereco).Cells:
        if cel.Formula not in ("", None):
            raise RuntimeError(f"{ws.Name}!{cel.Address} nao esta livre: {cel.Formula!r}")


# ------------------------------------------------------------- frentes
def frente_parametros(wb) -> None:
    mem = wb.Worksheets(MEM)
    for linha, chave, texto in TEXTOS_HISTORICO:
        atual = mem.Range(f"AF{linha}").Value
        if atual not in (None, "", chave):
            raise RuntimeError(f"{MEM}!AF{linha} ocupada: {atual!r}")
        if atual in (None, ""):
            _exigir_vazias(mem, f"AF{linha}:AG{linha}")
        mem.Range(f"AF{linha}").Value = chave
        mem.Range(f"AG{linha}").Value = texto

    ws = wb.Worksheets("parametros")
    formula = formula_aviso_historico()
    validar_ascii_e_parenteses({CELULA_AVISO: formula})
    if ws.Range(CELULA_AVISO).Formula != formula:
        _exigir_vazias(ws, FAIXA_AVISO)
    rng = ws.Range(FAIXA_AVISO)
    rng.FormatConditions.Delete()
    _mesclar(ws, FAIXA_AVISO)
    ws.Range(CELULA_AVISO).Formula = formula
    _fonte(rng, tamanho=10, negrito=True, cor="FF9C0006", italico=False)
    _alinhar(rng, h=XL_LEFT, v=XL_VCENTER, quebra=True, recuo=1)
    ws.Rows(8).RowHeight = 18
    regra = rng.FormatConditions.Add(XL_EXPRESSION, None, _local(rng, "=LEN($A$8)>0"))
    regra.StopIfTrue = False
    regra.Interior.Color = _bgr("FFFFC7CE")

    if ws.Range(CELULA_NOTA_FATOR).Value not in (None, "", NOTA_FATOR):
        raise RuntimeError(f"parametros!{CELULA_NOTA_FATOR} ocupada")
    _exigir_vazias(ws, "B17:F17")
    nota = ws.Range(CELULA_NOTA_FATOR)
    nota.Value = NOTA_FATOR
    _fonte(nota, tamanho=9, negrito=False, italico=True, cor="FF595959")
    _alinhar(nota, h=XL_LEFT, v=XL_VCENTER)


def _altura_texto(texto: str, largura_caracteres: float, tamanho: float = 10) -> float:
    """Altura de linha para texto com quebra numa faixa mesclada (AutoFit nao
    funciona em celula mesclada). Margem de 15% sobre a largura util."""
    por_linha = max(1.0, largura_caracteres * 0.85 * (11.0 / tamanho))
    linhas = max(1, math.ceil(len(texto) / por_linha))
    return round(linhas * tamanho * 1.45 + 4, 1)


def frente_consumidos(wb) -> None:
    ws = wb.Worksheets(CONS)
    # Fotografia das formulas do bloco: nada pode mudar alem de estilo/texto.
    antes = {c.Address: c.Formula for c in ws.Range("X1:AG6").Cells}

    largura = sum(ws.Columns(c).ColumnWidth for c in
                  ("X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG"))

    # Cabecalhos (linha 1): identidade propria; Z1/AA1 em ambar de entrada.
    cab = ws.Range("X1:AG1")
    _preencher(cab, COR_BLOCO_CAB)
    _fonte(cab, tamanho=10, negrito=True, cor=COR_BLOCO_TEXTO)
    _alinhar(cab, h=XL_CENTER, v=XL_VCENTER, quebra=True)
    for col in COLUNAS_ENTRADA:
        _preencher(ws.Range(f"{col}1"), COR_ENTRADA_CAB)
    if ws.Rows(1).RowHeight < 30:
        ws.Rows(1).RowHeight = 30

    # Dados C0..C4 (linhas 2:6).
    for col in COLUNAS_AUTOMATICAS:
        rng = ws.Range(f"{col}2:{col}6")
        _preencher(rng, COR_AUTOMATICA)
        _fonte(rng, tamanho=10, cor=COR_TEXTO_AUTO)
    for col in COLUNAS_ENTRADA:
        rng = ws.Range(f"{col}2:{col}6")
        _preencher(rng, COR_ENTRADA)
        _fonte(rng, tamanho=10, cor="FF000000")
    _alinhar(ws.Range("X2:X6"), h=XL_CENTER)
    _alinhar(ws.Range("AG2:AG6"), h=XL_LEFT, quebra=True)
    _bordas(ws.Range("X1:AG6"), COR_BORDA)

    # Orientacao (X8:X13): titulo do bloco + explicacoes ja existentes.
    if ws.Range("X9").Value == TEXTO_X9_BASE:
        ws.Range("X9").Value = ENTRADAS_AJUSTE + TEXTO_X9_BASE
    elif not str(ws.Range("X9").Value or "").startswith(ENTRADAS_AJUSTE):
        raise RuntimeError(f"{CONS}!X9 com texto inesperado: {ws.Range('X9').Value!r}")
    ws.Range("X8").Value = TITULO_AJUSTE
    _exigir_vazias(ws, "Y8:AG13")
    for linha in range(8, 14):
        faixa = f"X{linha}:AG{linha}"
        _mesclar(ws, faixa)
        rng = ws.Range(faixa)
        _alinhar(rng, h=XL_LEFT, v=XL_VCENTER, quebra=True, recuo=1)
        if linha == 8:
            _preencher(rng, COR_BLOCO_CAB)
            _fonte(rng, tamanho=11, negrito=True, cor=COR_BLOCO_TEXTO)
            ws.Rows(linha).RowHeight = 20
        else:
            _preencher(rng, COR_ORIENTACAO)
            _fonte(rng, tamanho=10, negrito=False, cor="FF404040")
            texto = str(ws.Range(f"X{linha}").Value or "")
            ws.Rows(linha).RowHeight = max(15.0, _altura_texto(texto, largura))
    _bordas(ws.Range("X8:AG13"), COR_BORDA)

    depois = {c.Address: c.Formula for c in ws.Range("X1:AG6").Cells}
    if antes != depois:
        mudou = [k for k in antes if antes[k] != depois.get(k)]
        raise RuntimeError(f"Bloco X1:AG6 teve conteudo alterado: {mudou[:5]}")


def frente_detalhe(wb) -> None:
    exe = wb.Worksheets(EXE)
    d31 = exe.Range("D31")
    if d31.Hyperlinks.Count:
        d31.Hyperlinks.Delete()
    d31.Value = TEXTO_AJUSTES_D31
    vizinho = exe.Range("D30")  # mesma familia visual das demais descricoes
    d31.Font.Name = vizinho.Font.Name
    d31.Font.Size = vizinho.Font.Size
    d31.Font.Bold = False
    d31.Font.Italic = False
    d31.Font.Underline = vizinho.Font.Underline
    d31.Font.Color = vizinho.Font.Color
    exe.Range("B3").Value = SUBTITULO_B3

    det = wb.Worksheets(DET)
    if det.Visible == XL_SHEET_VERY_HIDDEN:
        raise RuntimeError("RESULTADOS_DETALHE esta veryHidden (proibido)")
    det.Visible = XL_SHEET_HIDDEN


FRENTES = {
    "parametros": frente_parametros,
    "consumidos": frente_consumidos,
    "detalhe": frente_detalhe,
}


def aplicar(caminho: Path, frentes: list[str]) -> None:
    import pythoncom
    import win32com.client as com

    pythoncom.CoInitialize()
    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(caminho))
        excel.Calculation = XL_CALCULO_MANUAL
        for nome in frentes:
            FRENTES[nome](wb)
            print("frente aplicada:", nome)
        excel.Calculation = XL_CALCULO_AUTOMATICO
        excel.CalculateFullRebuild()
        controle = wb.Worksheets("CONTROLE")
        controle.Activate()
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1
        wb.Save()
        _fechar(wb)
        wb = None
    finally:
        if wb is not None:
            _fechar(wb)
        for _ in range(10):
            try:
                excel.Quit()
                break
            except Exception:
                time.sleep(1.0)
        wb = excel = None
        gc.collect()
        pythoncom.CoUninitialize()


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--frente", action="append", choices=sorted(FRENTES), required=True)
    p.add_argument("xlsx", nargs="?", type=Path, default=TEMPLATE)
    args = p.parse_args()
    caminho = args.xlsx.resolve()
    if not caminho.exists():
        print("ERRO: arquivo nao encontrado: %s" % caminho)
        return 1
    import pywintypes

    for tentativa in range(1, 6):
        try:
            aplicar(caminho, args.frente)
            break
        except pywintypes.com_error as erro:
            if erro.args[0] != -2147418111 or tentativa == 5:
                raise
            print("Excel ocupado (tentativa %d); reexecutando do zero..." % tentativa)
            time.sleep(8.0)
    print("OK: %s aplicado(s) em %s" % (", ".join(args.frente), caminho))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
