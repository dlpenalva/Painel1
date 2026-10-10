# -*- coding: utf-8 -*-
"""Coleta 11.8 — ajustes da Coleta em uso.

Aplica no template oficial, via Excel COM (openpyxl destroi a CF x14), somente:

* itens_PC!K2:L5001 — "Nao"/"Não" equivalentes em PC_PAGO_A_CONTRATADA; PC de
  ciclo PRECLUSO sem INICIO_EFEITO_FINANCEIRO fica sem efeito proprio (L=Nao).
* MEMORIA_RESULTADOS!S49:T53 (novos) e T23, T25, T40, B35, C35, D35, W50, W52,
  E26 — VTA com a posicao atual da CICLO_EM_EXECUCAO quando falta o residual
  do ciclo vigente (Financeiro e PCs).
* RESULTADOS_DETALHE!B36:B38 e A39 — bloco do ciclo atual coerente com o VTA.
* parametros — P2:P80 com 4 casas, Q2:Q80 em xx,xx%, A3:E6 sem azul/negrito
  fixo (CF: azul/negrito so em TEMPESTIVO) e G12:G15 com o preenchimento de
  campo de entrada.
* posicao_referencia — largura de F, H e I ajustada ao maior texto exibivel.
* fonte verde Nxxx em posicao_contratual, posicao_referencia, aditivos e
  historico_VU (CICLO_EM_EXECUCAO e criada em runtime: _ciclo_em_execucao).
* cobertura_temporal — B8 com a data real da posicao fisica, coluna B mais
  larga e com quebra, legenda A25:C29 excluida (sem dependentes).

As formulas vem de `_coleta_oficial` (fonte unica). Cada celula precisa estar
na forma 11.7 ou ja na 11.8; qualquer outra coisa aborta sem gravar.
REGRA ZERO CORRUPCAO XLSX: formulas em ingles e com parenteses balanceados.

uso: python tools/aplicar_coleta_118_ajustes_em_uso.py [caminho.xlsx]
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
import aplicar_ajustes_xls_ux_pos174 as ux  # noqa: E402
from aplicar_26h_template import _capturar_protecao, _restaurar_protecao  # noqa: E402

TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

XL_CALCULO_MANUAL, XL_CALCULO_AUTOMATICO = -4135, -4105
PRIMEIRA_PC, ULTIMA_PC = 2, 5001

# Larguras medidas com AutoFit do Excel sobre o maior texto que cada coluna
# exibe: F = "REMANESCENTE DE REFERENCIA SUPERA POSICAO CONTRATUAL" (Aptos 10),
# H = maior rotulo do painel, I = "POSICAO FISICA INCOMPLETA - QUANTIDADES
# PENDENTES EM CICLO_EM_EXECUCAO." (Calibri 11).
LARGURAS_POSICAO_REFERENCIA = {"F": 54.18, "H": 42.82, "I": 72.91}
LARGURA_COBERTURA_B = 50.0
LINHAS_COBERTURA_DUAS_LINHAS = (22, 23)
FAIXA_COBERTURA_B = "B2:B23"
LEGENDA_COBERTURA = "25:29"
FORMATO_P_PARAMETROS = ("0,0000", "0.0000")
FORMATO_Q_PARAMETROS = ("0,00%", "0.00%")
FORMULA_TEMPESTIVO = (
    '=AND(ISNUMBER(SEARCH("TEMPESTIVO",$G2)),ISERROR(SEARCH("INTEMPESTIVO",$G2)))'
)
COR_TEMPESTIVO = "123B63"
COR_ENTRADA = "FFF2CC"


def validar_formula(formula: str) -> None:
    if formula.count("(") != formula.count(")"):
        raise ValueError("formula com parenteses desbalanceados")
    if not formula.startswith("="):
        raise ValueError("formula sem '='")


def _com(formula: str) -> str:
    """Forma lida/gravada pelo COM: sem o prefixo _xlfn. do arquivo."""
    return formula.replace("_xlfn.", "")


def frente_itens_pc(wb) -> None:
    ws = wb.Worksheets("itens_PC")
    estado, selecao = _capturar_protecao(ws)
    for coluna, anterior, nova in (
        ("K", co._formula_check_pc_11_7, co._formula_check_pc),
        ("L", co._formula_efeito_pc_11_7, co._formula_efeito_pc),
    ):
        faixa = ws.Range(f"{coluna}{PRIMEIRA_PC}:{coluna}{ULTIMA_PC}")
        novas = []
        for deslocamento, (atual,) in enumerate(faixa.Formula):
            linha = PRIMEIRA_PC + deslocamento
            formula = _com(nova(linha))
            validar_formula(formula)
            if atual not in (_com(anterior(linha)), formula):
                raise RuntimeError(f"itens_PC!{coluna}{linha} fora do formato esperado; nada aplicado")
            novas.append((formula,))
        faixa.Formula = tuple(novas)
    _restaurar_protecao(ws, estado, selecao)


def _forma_11_8(aba: str, celula: str, atual: str) -> str:
    """Forma 11.8 da celula (idempotente: a forma 11.8 volta como esta)."""
    trocas = co._TRECHOS_POSICAO_ATUAL_VTA[(aba, celula)]
    if all(troca in atual for _, troca in trocas):
        return atual
    try:
        return co._formula_posicao_atual_vta(aba, celula, atual)
    except ValueError:
        raise RuntimeError(f"{aba}!{celula} fora do formato 11.7; nada aplicado") from None


def frente_vta_posicao_atual(wb) -> None:
    mem = wb.Worksheets("MEMORIA_RESULTADOS")
    for celula, (rotulo, formula) in co._CELULAS_POSICAO_ATUAL_VTA.items():
        linha = celula[1:]
        destino_rotulo, destino = mem.Range(f"S{linha}"), mem.Range(celula)
        if destino_rotulo.Formula not in ("", rotulo) or destino.Formula not in ("", formula):
            raise RuntimeError(f"MEMORIA_RESULTADOS!S{linha}:{celula} nao esta livre; nada aplicado")
        validar_formula(formula)
    novas = {}
    for (aba, celula) in co._TRECHOS_POSICAO_ATUAL_VTA:
        nova = _forma_11_8(aba, celula, str(wb.Worksheets(aba).Range(celula).Formula))
        validar_formula(nova)
        novas[(aba, celula)] = nova
    formato_valor = mem.Range("T23").NumberFormat
    for celula, (rotulo, formula) in co._CELULAS_POSICAO_ATUAL_VTA.items():
        linha = celula[1:]
        mem.Range(f"S{linha}").Value = rotulo
        _copiar_formato_simples(mem.Range("S47"), mem.Range(f"S{linha}"))
        mem.Range(celula).Formula = formula
        _copiar_formato_simples(mem.Range("T47"), mem.Range(celula))
        mem.Range(celula).NumberFormat = formato_valor
    mem.Range("T49").NumberFormat = "0"
    mem.Range("T52").NumberFormat = mem.Range("B35").NumberFormat
    for (aba, celula), nova in novas.items():
        wb.Worksheets(aba).Range(celula).Formula = nova


def _copiar_formato_simples(origem, destino) -> None:
    destino.Font.Name = origem.Font.Name
    destino.Font.Size = origem.Font.Size
    destino.Font.Bold = origem.Font.Bold
    destino.Font.Color = origem.Font.Color
    destino.HorizontalAlignment = origem.HorizontalAlignment
    destino.VerticalAlignment = origem.VerticalAlignment
    destino.WrapText = origem.WrapText
    destino.Locked = origem.Locked


def _cf_fonte_ultima(faixa, ancora, formula: str, cor: str, *, negrito: bool = False) -> None:
    """CF de fonte com a MENOR prioridade (alertas existentes prevalecem)."""
    validar_formula(formula)
    local = ux._local(ancora, formula)
    # Reaplicacao idempotente: a formula e exclusiva desta regra na aba (o
    # endereco de uma uniao nao se compara de forma estavel via COM).
    for i in range(faixa.FormatConditions.Count, 0, -1):
        regra = faixa.FormatConditions(i)
        if regra.Formula1 == local:
            regra.Delete()
    regra = faixa.FormatConditions.Add(ux.XL_EXPRESSION, None, local)
    if negrito:
        regra.Font.Bold = True
    regra.Font.Color = ux._bgr(cor)
    regra.SetLastPriority()


def frente_parametros(wb) -> None:
    ws = wb.Worksheets("parametros")
    estado, selecao = _capturar_protecao(ws)
    ux._formato(ws.Range("P2:P80"), *FORMATO_P_PARAMETROS)
    ux._formato(ws.Range("Q2:Q80"), *FORMATO_Q_PARAMETROS)
    # A3:E4 carregavam azul/negrito fixo; agora a cor segue a situacao (G).
    neutro = ws.Range("A5")
    linhas = ws.Range("A3:E6")
    linhas.Font.Bold = neutro.Font.Bold
    linhas.Font.Color = neutro.Font.Color
    faixa = ws.Range("A2:E6")
    _cf_fonte_ultima(faixa, faixa, FORMULA_TEMPESTIVO, COR_TEMPESTIVO, negrito=True)
    ux._preencher(ws.Range("G12:G15"), COR_ENTRADA)
    _restaurar_protecao(ws, estado, selecao)


def frente_posicao_referencia(wb) -> None:
    ws = wb.Worksheets("posicao_referencia")
    estado, selecao = _capturar_protecao(ws)
    for coluna, largura in LARGURAS_POSICAO_REFERENCIA.items():
        ws.Columns(coluna).ColumnWidth = largura
    _restaurar_protecao(ws, estado, selecao)


def frente_novos_itens(wb) -> None:
    formula = "=" + co.formula_destaque_novo_item(2)
    for aba, endereco in co._FAIXAS_DESTAQUE_NOVOS_ITENS_118.items():
        ws = wb.Worksheets(aba)
        estado, selecao = _capturar_protecao(ws)
        # Uniao via Application.Union: o separador de Range("A1,B1") segue o
        # idioma da interface (pt-BR: ";").
        areas = [ws.Range(area) for area in endereco.split()]
        faixa = areas[0]
        for area in areas[1:]:
            faixa = wb.Application.Union(faixa, area)
        _cf_fonte_ultima(faixa, ws.Range("A2"), formula, co._COR_FONTE_NOVOS_ITENS)
        _restaurar_protecao(ws, estado, selecao)


def frente_cobertura(wb) -> None:
    ws = wb.Worksheets("cobertura_temporal")
    estado, selecao = _capturar_protecao(ws)
    validar_formula(co._FORMULA_DATA_POSICAO_COBERTURA)
    ws.Range("B8").Formula = co._FORMULA_DATA_POSICAO_COBERTURA
    ws.Range("C8").Value = co._AJUDA_DATA_POSICAO_COBERTURA
    ws.Columns("B").ColumnWidth = LARGURA_COBERTURA_B
    ws.Range(FAIXA_COBERTURA_B).WrapText = True
    for linha in LINHAS_COBERTURA_DUAS_LINHAS:
        ws.Rows(linha).RowHeight = max(ws.Rows(linha).RowHeight, 29)
    legenda = ws.Range("A25").Value
    if legenda not in (None, "") and "LEGENDA" not in str(legenda):
        raise RuntimeError("cobertura_temporal!A25 nao e a legenda; nada excluido")
    if legenda not in (None, ""):
        ws.Rows(LEGENDA_COBERTURA).Delete()
    _restaurar_protecao(ws, estado, selecao)


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
        excel.Calculation = XL_CALCULO_MANUAL
        frente_itens_pc(wb)
        frente_vta_posicao_atual(wb)
        frente_parametros(wb)
        frente_posicao_referencia(wb)
        frente_novos_itens(wb)
        frente_cobertura(wb)
        excel.Calculation = XL_CALCULO_AUTOMATICO
        excel.CalculateFullRebuild()
        wb.Worksheets("CONTROLE").Activate()
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1
        wb.Save()
        ux._fechar(wb)
        wb = None
    finally:
        if wb is not None:
            ux._fechar(wb)
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
    caminho = Path(sys.argv[1] if len(sys.argv) > 1 else TEMPLATE).resolve()
    if not caminho.exists():
        print("ERRO: arquivo nao encontrado: %s" % caminho)
        return 1
    aplicar(caminho)
    print("OK: Coleta 11.8 aplicada em %s" % caminho)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
