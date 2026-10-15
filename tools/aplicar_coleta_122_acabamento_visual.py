# -*- coding: utf-8 -*-
"""Coleta 12.2 — acabamento visual de itens_PC, aditivos e CONTROLE.

Somente APRESENTACAO. Nenhuma formula, valor, coordenada, validacao funcional
ou regra financeira muda; nenhuma linha ou coluna e inserida/removida.

itens_PC
  Borda vertical dupla, discreta, na direita de L1:L5001: separa a area de
  insercao (A:L) da area de resultados/calculos (M em diante) sem coluna nova.

aditivos
  A validacao JA EXISTENTE em A2:A200 (qualquer valor, so mensagem de entrada)
  recebe o titulo NOVO ITEM e o texto de cadastro do item Nxxx. O cabecalho A1
  continua apenas ITEM.

CONTROLE
  * coluna B mais larga, para o dropdown de metodo (B1) caber inteiro;
  * C1 nao esta livre (alerta de coerencia da 12.0, =MEMORIA_RESULTADOS!$T$62,
    que transborda para D1:G1 — texto em D1 o cortaria). A orientacao usa a
    propria C1 sem tocar no valor: formatacao condicional ativa so enquanto B1
    esta em "SELECIONE AQUI" (ou vazio) — estado em que C1 e sempre "" — com
    fundo amarelo claro e formato de numero que exibe "<- Selecione o metodo".
    Apos uma selecao valida a regra desliga e C1 volta a mostrar so o alerta.

REGRA ZERO CORRUPCAO XLSX
  Somente Excel COM (openpyxl destroi a formatacao condicional x14 de
  RESULTADOS_DETALHE). A formula da regra nova nao usa funcoes nem
  separadores (identica em qualquer idioma) e le so a propria aba (sem x14).
  Qualquer celula fora da forma 12.1 aborta sem gravar.

uso: python tools/aplicar_coleta_122_acabamento_visual.py [caminho.xlsx]
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

import aplicar_resultados_executivo_v3 as v3  # noqa: E402

TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

# ------------------------------------------------------------- itens_PC
FAIXA_SEPARACAO_PC = "L1:L5001"
COR_SEPARACAO_PC = "8497B0"          # azul-acinzentado: claro, mas discreto
XL_EDGE_RIGHT = 10
XL_DOUBLE = -4119
CABECALHO_PC_ESPERADO = {
    "A1": "NUMERO_PC",
    "L1": "EFEITO_FINANCEIRO_PC",
    "M1": "TODOS OS PCs CADASTRADOS POR CICLO",
}

# ------------------------------------------------------------- aditivos
FAIXA_ITEM_ADITIVOS = "A2:A200"
TITULO_ITEM_ANTERIOR = "ITEM DO ADITIVO"
TITULO_NOVO_ITEM = "NOVO ITEM"
MENSAGEM_NOVO_ITEM = (
    "Cadastre primeiro em itens_Remanesc como N001, N002...\n"
    "A QTD_BASE_ORIGINAL ficará 0 automaticamente.\n"
    "Informe o VU_ORIGINAL e, depois, registre aqui a data e a quantidade "
    "do aditivo."
)
XL_VALIDATE_INPUT_ONLY = 0

# ------------------------------------------------------------- CONTROLE
LARGURA_ANTERIOR_B = 22.0
LARGURA_CONTROLE_B = 34.0
FORMULA_C1 = "=MEMORIA_RESULTADOS!$T$62"
OPCOES_METODO = "SELECIONE AQUI,Principal (Financeiro),PC (Pedidos de Compra),Itens Consumidos"
TEXTO_ORIENTACAO_METODO = "← Selecione o método de apuração"
# Sem funcoes/separadores: identica no idioma local; le so a propria aba.
FORMULA_CF_ORIENTACAO = '=($B$1="SELECIONE AQUI")+($B$1="")'
# Secao de TEXTO (4a) so com o literal: C1 vale "" nesse estado.
FORMATO_ORIENTACAO = ';;;"%s"' % TEXTO_ORIENTACAO_METODO
COR_ORIENTACAO_FUNDO = "FFF2CC"
COR_ORIENTACAO_TEXTO = "7F6000"

if len(MENSAGEM_NOVO_ITEM) > 255:
    raise SystemExit("mensagem de entrada acima do limite do Excel (255)")


def _opcoes(formula1: str) -> list[str]:
    """Itens da lista; o COM devolve o separador do idioma local (; ou ,)."""
    return str(formula1).replace(";", ",").split(",")


# ------------------------------------------------------------- conferencias
def conferir(wb) -> None:
    """Forma 12.1 exigida antes de qualquer escrita (fail-closed)."""
    pc = wb.Worksheets("itens_PC")
    for endereco, texto in CABECALHO_PC_ESPERADO.items():
        if pc.Range(endereco).Value != texto:
            raise RuntimeError(f"itens_PC!{endereco} fora do layout 12.1; nada aplicado")

    adt = wb.Worksheets("aditivos")
    if adt.Range("A1").Value != "ITEM":
        raise RuntimeError("aditivos!A1 fora do layout 12.1; nada aplicado")
    validacao = adt.Range(FAIXA_ITEM_ADITIVOS).Validation
    try:
        tipo, titulo = validacao.Type, validacao.InputTitle
    except Exception as exc:  # faixa sem validacao unica
        raise RuntimeError("aditivos!A2:A200 sem validacao unica; nada aplicado") from exc
    if tipo != XL_VALIDATE_INPUT_ONLY or titulo not in (TITULO_ITEM_ANTERIOR, TITULO_NOVO_ITEM):
        raise RuntimeError("aditivos!A2:A200 fora da validacao 12.1; nada aplicado")

    ctl = wb.Worksheets("CONTROLE")
    if ctl.Range("C1").Formula != FORMULA_C1:
        raise RuntimeError("CONTROLE!C1 fora da forma 12.0; nada aplicado")
    if _opcoes(ctl.Range("B1").Validation.Formula1) != _opcoes(OPCOES_METODO):
        raise RuntimeError("CONTROLE!B1 fora do dropdown 12.1; nada aplicado")
    if ctl.Range("C1").FormatConditions.Count != 1:
        raise RuntimeError("CONTROLE!C1 com formatacao condicional inesperada; nada aplicado")
    # ColumnWidth (COM) desconta o padding gravado no arquivo (~0,7).
    if abs(float(ctl.Columns("B").ColumnWidth) - LARGURA_ANTERIOR_B) > 1.5:
        raise RuntimeError("CONTROLE!B com largura inesperada; nada aplicado")


# ------------------------------------------------------------- etapas COM
def frente_itens_pc(wb) -> None:
    borda = wb.Worksheets("itens_PC").Range(FAIXA_SEPARACAO_PC).Borders(XL_EDGE_RIGHT)
    borda.LineStyle = XL_DOUBLE
    borda.Color = v3._bgr(COR_SEPARACAO_PC)


def frente_aditivos(wb) -> None:
    validacao = wb.Worksheets("aditivos").Range(FAIXA_ITEM_ADITIVOS).Validation
    validacao.InputTitle = TITULO_NOVO_ITEM
    validacao.InputMessage = MENSAGEM_NOVO_ITEM
    validacao.ShowInput = True


def frente_controle(wb) -> None:
    # Import tardio: a ferramenta 12.0 exige pywin32, e os testes (CI Linux)
    # importam este modulo so para ler as constantes.
    import aplicar_coleta_120_itens_pc_simnao_coerencia as c120

    ws = wb.Worksheets("CONTROLE")
    estado, selecao = c120._capturar_protecao(ws)
    ws.Columns("B").ColumnWidth = LARGURA_CONTROLE_B
    alvo = ws.Range("C1")
    regra = alvo.FormatConditions.Add(v3.XL_EXPRESSION, None, FORMULA_CF_ORIENTACAO)
    regra.StopIfTrue = False
    regra.NumberFormat = FORMATO_ORIENTACAO
    regra.Interior.Color = v3._bgr(COR_ORIENTACAO_FUNDO)
    regra.Font.Color = v3._bgr(COR_ORIENTACAO_TEXTO)
    regra.Font.Bold = True
    c120._restaurar_protecao(ws, estado, selecao)


def conferir_resultado(wb) -> None:
    pc = wb.Worksheets("itens_PC").Range(FAIXA_SEPARACAO_PC).Borders(XL_EDGE_RIGHT)
    if pc.LineStyle != XL_DOUBLE:
        raise RuntimeError("itens_PC: separacao L|M nao aplicada")
    validacao = wb.Worksheets("aditivos").Range(FAIXA_ITEM_ADITIVOS).Validation
    if validacao.InputTitle != TITULO_NOVO_ITEM or validacao.Type != XL_VALIDATE_INPUT_ONLY:
        raise RuntimeError("aditivos!A2:A200: mensagem de novo item nao aplicada")
    ctl = wb.Worksheets("CONTROLE")
    if ctl.Range("C1").Formula != FORMULA_C1 or ctl.Range("C1").FormatConditions.Count != 2:
        raise RuntimeError("CONTROLE!C1: orientacao do metodo inesperada")
    if _opcoes(ctl.Range("B1").Validation.Formula1) != _opcoes(OPCOES_METODO):
        raise RuntimeError("CONTROLE!B1: dropdown alterado")


def aplicar(caminho: Path) -> None:
    import pythoncom
    import win32com.client as com

    pythoncom.CoInitialize()
    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(caminho))
        excel.Calculation = v3.XL_CALCULO_MANUAL
        conferir(wb)
        frente_itens_pc(wb)
        frente_aditivos(wb)
        frente_controle(wb)
        conferir_resultado(wb)
        excel.Calculation = v3.XL_CALCULO_AUTOMATICO
        excel.CalculateFullRebuild()
        for aba in ("itens_PC", "aditivos", "CONTROLE"):
            wb.Worksheets(aba).Activate()
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
    print("OK: acabamento visual da Coleta 12.2 aplicado em %s" % caminho)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
