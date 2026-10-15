# -*- coding: utf-8 -*-
"""Coleta 12.1 — orientacao visual no topo das abas de metodo.

Somente APRESENTACAO (prevencao de erro). Nenhuma regra financeira, nenhuma
fonte de dados e nenhuma validacao muda: a coerencia metodo x fonte da Coleta
12.0 (MEMORIA_RESULTADOS!T54:T62, bloqueio/aviso em RESULTADOS e CONTROLE!C1)
continua sendo a unica regra; aqui ela so e REEXIBIDA no topo de cada aba.

ONDE
  As tres abas tem o cabecalho na linha 1 e os dados a partir da linha 2 — o
  gerador, o validador do upload e o leitor dependem disso, entao nenhuma
  linha e inserida. Em cada aba:

  * cabecalho da area de dados (linha 1) pintado conforme o estado, por
    formatacao condicional — visivel assim que a aba abre;
  * faixa de texto (4 colunas mescladas) na linha 1, logo apos a ultima
    coluna usada, com titulo e orientacao;
  * celula de estado (1 a 4) logo apos a faixa, com formato ;;; (invisivel):
    a formatacao condicional le essa celula da PROPRIA aba (regra que le
    outra aba vai para a extensao x14, que o openpyxl descarta na geracao).

  financeiro        cabecalho A1:G1  | faixa H1:K1   | estado L1
  itens_Consumidos  cabecalho A1:AG1 | faixa AH1:AK1 | estado AL1
  itens_PC          cabecalho A1:L1  | faixa AE1:AH1 | estado AI1

ESTADOS (mesmos indicadores da validacao metodo x fonte da Coleta 12.0)
  metodo  = MEMORIA_RESULTADOS!B4 (Financeiro / PCs / Itens / nao selecionado)
  dados   = MEMORIA_RESULTADOS!T54 (financeiro), T55 (itens_PC), T56
            (itens_Consumidos); T57 numerico = ha metodo selecionado.
  1 verde    a aba e a fonte do metodo selecionado
  2 amarelo  outro metodo selecionado e a aba sem dados
  3 laranja  outro metodo selecionado e a aba COM dados (so aviso)
  4 azul     nenhum metodo selecionado
  O nome exibido e o texto real de CONTROLE!B1.

REGRA ZERO CORRUPCAO XLSX
  Somente Excel COM (openpyxl destroi a formatacao condicional x14 de
  RESULTADOS_DETALHE); formulas em ingles, ASCII e parenteses balanceados;
  textos acentuados so como constantes em MEMORIA_RESULTADOS!AG71:AG79.
  Qualquer celula fora da forma 12.1 (Pacote 1) aborta sem gravar.

uso: python tools/aplicar_coleta_orientacao_abas_metodo.py [caminho.xlsx]
     (exige pywin32 + Excel real; sem argumento, altera o template oficial)
"""
from __future__ import annotations

import gc
import sys
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
for _caminho in (RAIZ, RAIZ / "tools"):
    if str(_caminho) not in sys.path:
        sys.path.insert(0, str(_caminho))

import _coleta_oficial as co  # noqa: E402
import aplicar_resultados_executivo_v3 as v3  # noqa: E402

TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"
MEM = v3.MEM

# ------------------------------------------------------------- textos (AF:AG)
PRIMEIRA_LINHA_TEXTO = 71
TEXTOS = [
    ("ABA_ATIVO_TIT", "MÉTODO ATIVO — "),
    ("ABA_ATIVO_TXT", "Preencha os dados desta aba para a apuração."),
    ("ABA_OUTRO_TIT", "ESTA ABA NÃO CORRESPONDE AO MÉTODO SELECIONADO"),
    ("ABA_DADOS_TIT", "HÁ DADOS NESTA ABA, MAS ELA NÃO É O MÉTODO ATIVO"),
    ("ABA_METODO_ATUAL", "Método atual: "),
    ("ABA_OUTRO_TXT", ". Os dados desta aba não serão utilizados como fonte "
                      "principal desta apuração."),
    ("ABA_DADOS_TXT", ". Confirme em CONTROLE se a escolha do método está correta."),
    ("ABA_SEM_TIT", "SELECIONE PRIMEIRO O MÉTODO DE APURAÇÃO"),
    ("ABA_SEM_TXT", "Selecione o método na aba CONTROLE. Depois, preencha a aba "
                    "correspondente."),
]
LINHA_TEXTO = {chave: PRIMEIRA_LINHA_TEXTO + i for i, (chave, _) in enumerate(TEXTOS)}


def T(chave: str) -> str:
    return f"{MEM}!${v3.COL_TEXTO}${LINHA_TEXTO[chave]}"


# ------------------------------------------------------------- geometria
ESTADO_ATIVO, ESTADO_OUTRO, ESTADO_OUTRO_COM_DADOS, ESTADO_SEM_METODO = 1, 2, 3, 4

# aba: (codigo em MEMORIA_RESULTADOS!B4, contagem da fonte, cabecalho, faixa,
#       celula de estado)
ABAS = {
    "financeiro": ("Financeiro", "$T$54", "A1:G1", "H1:K1", "L1"),
    "itens_Consumidos": ("Itens", "$T$56", "A1:AG1", "AH1:AK1", "AL1"),
    "itens_PC": ("PCs", "$T$55", "A1:L1", "AE1:AH1", "AI1"),
}
# Cabecalho esperado (A1 e ultima coluna) — a faixa nunca cobre a grade.
CABECALHO_ESPERADO = {
    "financeiro": {"A1": "COMPETENCIA", "G1": "EFEITO_FINANCEIRO"},
    "itens_Consumidos": {"A1": "ITEM", "AG1": "AJUSTE_STATUS"},
    "itens_PC": {"A1": "NUMERO_PC", "L1": "EFEITO_FINANCEIRO_PC",
                 "AD1": co.CABECALHO_PC_PAGO_NORMALIZADO},
}
LARGURA_COLUNA_FAIXA = 28
TAMANHO_FONTE_FAIXA = 9
COR_TEXTO_BASE = "262626"
FORMATO_ESTADO = ";;;"

# estado: (fundo, texto) — claros; o laranja e mais forte que o amarelo.
CORES = {
    ESTADO_ATIVO: ("E2EFDA", "375623"),
    ESTADO_OUTRO: ("FFF2CC", "7F6000"),
    ESTADO_OUTRO_COM_DADOS: ("F8CBAD", "843C0C"),
    ESTADO_SEM_METODO: ("DDEBF7", "1F4E79"),
}

# Indicadores da Coleta 12.0 (fonte unica: _coleta_oficial._CELULAS_COLETA_120).
INDICADORES_120 = ("T54", "T55", "T56", "T57")


def _absoluta(celula: str) -> str:
    coluna = celula.rstrip("0123456789")
    return f"${coluna}${celula[len(coluna):]}"


def formula_estado(aba: str) -> str:
    codigo, contagem, *_ = ABAS[aba]
    m = f"{MEM}!"
    return (
        f'=IF({m}$B$4="{codigo}",{ESTADO_ATIVO},'
        f"IF(ISNUMBER({m}$T$57),IF({m}{contagem}>0,{ESTADO_OUTRO_COM_DADOS},"
        f"{ESTADO_OUTRO}),{ESTADO_SEM_METODO}))"
    )


def formula_faixa(aba: str) -> str:
    estado = _absoluta(ABAS[aba][4])
    metodo = "CONTROLE!$B$1"
    quebra = "CHAR(10)"
    return (
        f"=CHOOSE({estado},"
        f'{T("ABA_ATIVO_TIT")}&{metodo}&{quebra}&{T("ABA_ATIVO_TXT")},'
        f'{T("ABA_OUTRO_TIT")}&{quebra}&{T("ABA_METODO_ATUAL")}&{metodo}&{T("ABA_OUTRO_TXT")},'
        f'{T("ABA_DADOS_TIT")}&{quebra}&{T("ABA_METODO_ATUAL")}&{metodo}&{T("ABA_DADOS_TXT")},'
        f'{T("ABA_SEM_TIT")}&{quebra}&{T("ABA_SEM_TXT")})'
    )


def regras(aba: str) -> list[tuple[str, int]]:
    """Regras de cor (formula, estado): leem so a celula de estado da aba."""
    ref = _absoluta(ABAS[aba][4])
    return [(f"={ref}={codigo}", codigo) for codigo in CORES]


def faixa_regras(aba: str) -> str:
    _, _, cabecalho, faixa, _ = ABAS[aba]
    return f"{cabecalho},{faixa}"


def formulas() -> dict[tuple[str, str], str]:
    resultado = {}
    for aba, (*_r, faixa, estado) in ABAS.items():
        resultado[(aba, estado)] = formula_estado(aba)
        resultado[(aba, faixa.split(":")[0])] = formula_faixa(aba)
    return resultado


def validar_ascii_e_parenteses() -> None:
    v3.validar_ascii_e_parenteses(
        {f"{aba}!{celula}": f for (aba, celula), f in formulas().items()}
    )
    for aba in ABAS:
        v3.validar_ascii_e_parenteses({f"{aba}!CF{c}": f for f, c in regras(aba)})


def colunas_faixa(aba: str) -> list[str]:
    from openpyxl.utils import column_index_from_string, get_column_letter

    inicio, fim = ABAS[aba][3].split(":")
    a = column_index_from_string(inicio.rstrip("0123456789"))
    b = column_index_from_string(fim.rstrip("0123456789"))
    return [get_column_letter(i) for i in range(a, b + 1)]


# ------------------------------------------------------------- conferencias
def conferir(wb) -> None:
    """Forma 12.1 (Pacote 1) exigida antes de qualquer escrita (fail-closed)."""
    mem = wb.Worksheets(MEM)
    for celula in INDICADORES_120:
        esperado = co._CELULAS_COLETA_120[celula][1]
        if mem.Range(celula).Formula != esperado.replace("_xlfn.", ""):
            raise RuntimeError(f"{MEM}!{celula} fora da forma 12.0; nada aplicado")
    if "CONTROLE!$B$1" not in str(mem.Range("B4").Formula):
        raise RuntimeError(f"{MEM}!B4 nao le CONTROLE!B1; nada aplicado")
    for linha in range(PRIMEIRA_LINHA_TEXTO, PRIMEIRA_LINHA_TEXTO + len(TEXTOS)):
        for col in (v3.COL_CHAVE, v3.COL_TEXTO):
            if mem.Range(f"{col}{linha}").Formula != "":
                raise RuntimeError(f"{MEM}!{col}{linha} nao esta livre; nada aplicado")
    for aba, (_c, _n, cabecalho, faixa, estado) in ABAS.items():
        ws = wb.Worksheets(aba)
        if ws.ProtectContents:
            raise RuntimeError(f"{aba} protegida inesperadamente; nada aplicado")
        for endereco, texto in CABECALHO_ESPERADO[aba].items():
            if ws.Range(endereco).Value != texto:
                raise RuntimeError(f"{aba}!{endereco} fora do cabecalho esperado; nada aplicado")
        for endereco in (faixa, estado):
            rng = ws.Range(endereco)
            if rng.MergeCells or any(c.Formula != "" for c in rng.Cells):
                raise RuntimeError(f"{aba}!{endereco} nao esta livre; nada aplicado")
        if ws.Range(cabecalho).FormatConditions.Count:
            raise RuntimeError(f"{aba}!{cabecalho} ja tem formatacao condicional; nada aplicado")


# ------------------------------------------------------------- etapas COM
def frente_memoria(wb) -> None:
    mem = wb.Worksheets(MEM)
    for chave, texto in TEXTOS:
        linha = LINHA_TEXTO[chave]
        mem.Range(f"{v3.COL_CHAVE}{linha}").Value = chave
        celula = mem.Range(f"{v3.COL_TEXTO}{linha}")
        celula.NumberFormat = "@"
        celula.Value = texto


def frente_aba(wb, aba: str) -> None:
    ws = wb.Worksheets(aba)
    _c, _n, _cab, faixa, estado = ABAS[aba]
    for coluna in colunas_faixa(aba):
        ws.Columns(coluna).ColumnWidth = LARGURA_COLUNA_FAIXA

    celula_estado = ws.Range(estado)
    celula_estado.NumberFormat = FORMATO_ESTADO
    celula_estado.Formula = formula_estado(aba)

    v3._mesclar(ws, faixa)
    rng = ws.Range(faixa)
    v3._fonte(rng, tamanho=TAMANHO_FONTE_FAIXA, negrito=True, cor=COR_TEXTO_BASE)
    v3._alinhar(rng, quebra=True, recuo=1)
    rng.Cells(1, 1).Formula = formula_faixa(aba)

    antes = ws.Cells.FormatConditions.Count
    # Union em vez de Range("a,b"): o separador de areas segue o idioma local.
    alvo = ws.Application.Union(*(ws.Range(p) for p in faixa_regras(aba).split(",")))
    for formula, codigo in regras(aba):
        fundo, texto = CORES[codigo]
        # Formula sem funcoes nem separadores: identica no idioma local.
        regra = alvo.FormatConditions.Add(v3.XL_EXPRESSION, None, formula)
        regra.StopIfTrue = False
        regra.Interior.Color = v3._bgr(fundo)
        regra.Font.Color = v3._bgr(texto)
    if ws.Cells.FormatConditions.Count != antes + len(CORES):
        raise RuntimeError(f"{aba}: formatacao condicional inesperada")


def conferir_resultado(wb) -> None:
    for (aba, celula), formula in formulas().items():
        atual = wb.Worksheets(aba).Range(celula).Formula
        if atual != formula:
            raise RuntimeError(f"{aba}!{celula} diverge do descritor: {atual}")


def aplicar(caminho: Path) -> None:
    import pythoncom
    import win32com.client as com

    validar_ascii_e_parenteses()
    pythoncom.CoInitialize()
    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(caminho))
        excel.Calculation = v3.XL_CALCULO_MANUAL
        conferir(wb)
        frente_memoria(wb)
        for aba in ABAS:
            frente_aba(wb, aba)
        conferir_resultado(wb)
        excel.Calculation = v3.XL_CALCULO_AUTOMATICO
        excel.CalculateFullRebuild()
        for aba in ABAS:
            wb.Worksheets(aba).Activate()
            excel.ActiveWindow.ScrollRow = 1
            excel.ActiveWindow.ScrollColumn = 1
        wb.Worksheets("CONTROLE").Activate()
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1
        wb.Save()
        v3._fechar(wb)
        wb = None
    finally:
        if wb is not None:
            v3._fechar(wb)
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
    caminho = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else TEMPLATE
    if not caminho.exists():
        print("ERRO: arquivo nao encontrado: %s" % caminho)
        return 1
    aplicar(caminho)
    print("OK: orientacao das abas de metodo aplicada em %s" % caminho)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
