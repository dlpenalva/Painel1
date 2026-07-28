# -*- coding: utf-8 -*-
"""Etapa 26F: RESULTADOS executiva e memoria tecnica separada, via Excel COM.

O aplicador preserva integralmente a planilha de calculo existente:

* a RESULTADOS anterior e renomeada para MEMORIA_RESULTADOS e fica oculta;
* os nomes definidos continuam apontando para a memoria, pois o proprio Excel
  atualiza as referencias durante a renomeacao;
* a nova RESULTADOS contem somente quatro tabelas executivas, um status global
  e uma tabela unica de ajustes manuais;
* entradas manuais antigas sao migradas para a nova interface e a memoria passa
  a referenciar essas entradas, sem duplicidade de campos;
* CONTROLE!B11 passa a falhar em branco se o historico ate o ciclo vigente nao
  estiver completo;
* cobertura_temporal!A8/C8 e o formato de financeiro!D2:D73 sao corrigidos.

O arquivo e sempre trabalhado em copia temporaria. A promocao do destino so
ocorre depois de recalculo completo, ausencia de erros Excel e reabertura.
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client


XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105
XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16
XL_VALIDATE_LIST = 3
XL_VALIDATE_DECIMAL = 2
XL_VALID_ALERT_STOP = 1
XL_BETWEEN = 1
XL_GREATER_EQUAL = 7
XL_EXPRESSION = 2
XL_CELL_VALUE = 1
XL_EQUAL = 3
XL_LESS = 6
XL_CONTINUOUS = 1
XL_THIN = 2
XL_MEDIUM = -4138
XL_LANDSCAPE = 2
XL_SHEET_HIDDEN = 0
XL_SHEET_VISIBLE = -1
XL_NONE = -4142
XL_UNLOCKED_CELLS = 1

ABA_MEMORIA = "MEMORIA_RESULTADOS"
ABA_RESULTADOS = "RESULTADOS"
NOME_OPCOES_MANUAL = "OPCOES_APLICAR_MANUAL"

MOEDA_BR = "R$ #.##0,00"
PERCENTUAL = "0,0000%"
FATOR = "0,000000"
DATA_BR = "dd/mm/aaaa"


def _cor(hex_rgb: str) -> int:
    """Converte RRGGBB para o inteiro BGR usado pelo Excel COM."""
    texto = hex_rgb.strip().lstrip("#")
    r, g, b = (int(texto[i : i + 2], 16) for i in (0, 2, 4))
    return r + (g << 8) + (b << 16)


CORES = {
    "vinho": _cor("8A1538"),
    "azul": _cor("1F4E78"),
    "azul_escuro": _cor("17365D"),
    "azul_claro": _cor("D9EAF7"),
    "azul_muito_claro": _cor("EEF4F8"),
    "cinza": _cor("F2F2F2"),
    "cinza_texto": _cor("595959"),
    "branco": _cor("FFFFFF"),
    "amarelo_input": _cor("FFF2CC"),
    "amarelo_status": _cor("FFE699"),
    "verde_status": _cor("C6E0B4"),
    "vermelho_status": _cor("F4CCCC"),
    "borda": _cor("B0C4D8"),
}


def _nomes_abas(wb) -> list[str]:
    return [ws.Name for ws in wb.Worksheets]


def _nomes_definidos(wb) -> set[str]:
    return {str(nome.Name).split("!")[-1] for nome in wb.Names}


def _validar_origem(wb) -> None:
    abas = _nomes_abas(wb)
    obrigatorias = {
        "comparativo_VTA",
        "CONTROLE",
        "parametros",
        "financeiro",
        "itens_Remanesc",
        "itens_Consumidos",
        "itens_PC",
        "posicao_referencia",
        "cobertura_temporal",
        ABA_RESULTADOS,
    }
    ausentes = sorted(obrigatorias.difference(abas))
    if ausentes:
        raise ValueError(f"Abas obrigatorias ausentes: {', '.join(ausentes)}")
    if ABA_MEMORIA in abas:
        raise ValueError(
            "MEMORIA_RESULTADOS ja existe; a Etapa 26F parece ter sido aplicada."
        )
    r = wb.Worksheets(ABA_RESULTADOS)
    if str(r.Range("A1").Value or "") != (
        "RESULTADOS CONSOLIDADOS — REAJUSTE CONTRATUAL"
    ):
        raise ValueError("RESULTADOS de origem nao corresponde ao layout homologado.")
    for endereco in ("B16", "B23", "B26", "B35", "D35", "F36", "T25"):
        if r.Range(endereco).Value is None and not str(r.Range(endereco).Formula or ""):
            raise ValueError(f"RESULTADOS!{endereco} esta ausente; layout inesperado.")


def _capturar_manuais(ws) -> dict[str, object]:
    enderecos = (
        "B5",
        "D5",
        "B24",
        "D24",
        "B25",
        "D25",
        "N258",
        "N259",
        "N260",
        "N261",
        "N262",
    )
    return {endereco: ws.Range(endereco).Value for endereco in enderecos}


_PROT_FLAGS = (
    "AllowFormattingCells",
    "AllowFormattingColumns",
    "AllowFormattingRows",
    "AllowInsertingColumns",
    "AllowInsertingRows",
    "AllowInsertingHyperlinks",
    "AllowDeletingColumns",
    "AllowDeletingRows",
    "AllowSorting",
    "AllowFiltering",
    "AllowUsingPivotTables",
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


def _formula_fator_historico() -> str:
    """Fator historico integral fail-closed ate o ciclo indicado em B2."""
    return (
        '=IFERROR(IF(UPPER($B$2)="C0",1,'
        'IF(UPPER($B$2)="C1",IF(COUNT(parametros!$E$3:$E$3)=1,parametros!$F$3,""),'
        'IF(UPPER($B$2)="C2",IF(COUNT(parametros!$E$3:$E$4)=2,parametros!$F$4,""),'
        'IF(UPPER($B$2)="C3",IF(COUNT(parametros!$E$3:$E$5)=3,parametros!$F$5,""),'
        'IF(UPPER($B$2)="C4",IF(COUNT(parametros!$E$3:$E$6)=4,parametros!$F$6,""),'
        '""))))),"")'
    )


def _ajustes_complementares(wb) -> None:
    controle = wb.Worksheets("CONTROLE")
    estado, selecao = _capturar_protecao(controle)
    controle.Range("A10").Value = "Variação histórica integral até o ciclo vigente"
    controle.Range("B10").Formula = '=IF(ISNUMBER(B11),B11-1,"")'
    controle.Range("A11").Value = "Fator histórico integral (até o ciclo vigente)"
    controle.Range("B11").Formula = _formula_fator_historico()
    controle.Range("B10").NumberFormatLocal = PERCENTUAL
    controle.Range("B11").NumberFormatLocal = FATOR
    _restaurar_protecao(controle, estado, selecao)

    cobertura = wb.Worksheets("cobertura_temporal")
    estado, selecao = _capturar_protecao(cobertura)
    cobertura.Range("A8").Value = (
        "Data da posição física ATUAL confirmada (vazio = não confirmada integralmente)"
    )
    cobertura.Range("C8").Value = (
        "Para confirmar: preencher QTD_REM_ATUAL em posicao_referencia + "
        "data de corte em CONTROLE!B3"
    )
    cobertura.Range("C8").WrapText = True
    cobertura.Columns("C").ColumnWidth = 52
    _restaurar_protecao(cobertura, estado, selecao)

    wb.Worksheets("financeiro").Range("D2:D73").NumberFormatLocal = "0,0000"


def _estilo_fonte(rng, *, branco: bool = False, negrito: bool = False, tamanho=10):
    rng.Font.Name = "Calibri"
    rng.Font.Size = tamanho
    rng.Font.Bold = negrito
    rng.Font.Color = CORES["branco"] if branco else CORES["cinza_texto"]


def _bordas_tabela(rng) -> None:
    for indice in (7, 8, 9, 10, 11, 12):
        borda = rng.Borders(indice)
        borda.LineStyle = XL_CONTINUOUS
        borda.Weight = XL_THIN
        borda.Color = CORES["borda"]


def _titulo_secao(ws, linha: int, texto: str) -> None:
    rng = ws.Range(f"A{linha}:H{linha}")
    rng.Interior.Color = CORES["vinho"]
    rng.RowHeight = 23
    _estilo_fonte(rng, branco=True, negrito=True, tamanho=11)
    ws.Range(f"A{linha}").Value = texto


def _cabecalho_tabela(ws, endereco: str) -> None:
    rng = ws.Range(endereco)
    rng.Interior.Color = CORES["azul"]
    _estilo_fonte(rng, branco=True, negrito=True)
    rng.HorizontalAlignment = -4108
    rng.VerticalAlignment = -4108
    _bordas_tabela(rng)


def _formula_pago(ciclo: str, linha: int) -> str:
    """Tri-state 26F.2: movimento -> valor; ZERO CONFIRMADO -> 0; ausente -> "".

    Zero confirmado exige que a cobertura confirmada GCC (cobertura_temporal
    B13 Financeiro / B15 PCs — nunca a mera "ultima evidencia" B12/B14)
    alcance integralmente a JANELA NECESSARIA do ciclo:
    MIN(fim nominal parametros!D, corte CONTROLE!B3 quando informado).
    Para Itens nao existe evidencia canonica capaz de provar consumo zero
    (posicao fisica so e conhecida em fotografias); sem movimento positivo a
    celula permanece vazia (DESCONHECIDO) — limitacao documentada.
    AUSENCIA DE LINHA != ZERO: vazio nunca e trocado por 0.
    """
    ciclo_min = ciclo.lower()
    par_linha = 2 + int(ciclo[1])
    qtd_itens = {
        "C0": "E",
        "C1": "G",
        "C2": "I",
        "C3": "K",
        "C4": "M",
    }[ciclo]
    janela = (
        f'IF(CONTROLE!$B$3="",parametros!$D${par_linha},'
        f'MIN(parametros!$D${par_linha},CONTROLE!$B$3))'
    )
    fin_completo = (
        f'AND(ISNUMBER(cobertura_temporal!$B$13),'
        f'ISNUMBER(parametros!$D${par_linha}),'
        f'cobertura_temporal!$B$13>={janela})'
    )
    pc_completo = (
        f'AND(ISNUMBER(cobertura_temporal!$B$15),'
        f'ISNUMBER(parametros!$D${par_linha}),'
        f'cobertura_temporal!$B$15>={janela})'
    )
    return (
        f'=IF($B$5="Financeiro",'
        f'IF(COUNTIFS(financeiro!$B$2:$B$73,"{ciclo_min}",'
        'financeiro!$C$2:$C$73,">0")>0,'
        f'ROUND(SUMIFS(financeiro!$C$2:$C$73,financeiro!$B$2:$B$73,"{ciclo_min}",'
        'financeiro!$C$2:$C$73,">0"),2),'
        f'IF({fin_completo},0,"")),'
        'IF($B$5="PCs",'
        f'IF(COUNTIFS(itens_PC!$C$2:$C$100,"{ciclo}",'
        'itens_PC!$D$2:$D$100,">0",itens_PC!$G$2:$G$100,"Sim")>0,'
        f'ROUND(SUMIFS(itens_PC!$D$2:$D$100,itens_PC!$C$2:$C$100,"{ciclo}",'
        'itens_PC!$G$2:$G$100,"Sim"),2),'
        f'IF({pc_completo},0,"")),'
        'IF($B$5="Itens",'
        f'IF(COUNTIFS(itens_Consumidos!${qtd_itens}$2:${qtd_itens}$200,">0")=0,"",'
        f'ROUND(SUMPRODUCT(itens_Consumidos!${qtd_itens}$2:${qtd_itens}$200,'
        'itens_Consumidos!$C$2:$C$200),2)),"")))'
    )


def _formula_rem_atualizado(ciclo: str) -> str:
    qtd = {"C0": "G", "C1": "K", "C2": "O", "C3": "S", "C4": "W"}[ciclo]
    vu = {"C0": "C", "C1": "D", "C2": "E", "C3": "F", "C4": "G"}[ciclo]
    return (
        f'=IF(B{{linha}}="","",ROUND(SUMPRODUCT('
        f'posicao_contratual!${qtd}$2:${qtd}$201,'
        f'historico_VU!${vu}$2:${vu}$201),2))'
    )


def _formula_data_referencia() -> str:
    """Data de referencia fail-closed: ausencia permanece vazia.

    26F.1: nenhum fallback pode devolver a celula bruta de um marco vazio —
    referencia direta a celula vazia vira serial 0 (00/01/1900) e mascara a
    ausencia. Todo ramo termina em data valida ou em "".
    """
    return (
        '=IF($B$5="Financeiro",'
        'IF(cobertura_temporal!$B$13<>"",cobertura_temporal!$B$13,'
        'IF(cobertura_temporal!$B$12<>"",cobertura_temporal!$B$12,"")),'
        'IF($B$5="PCs",'
        'IF(cobertura_temporal!$B$15<>"",cobertura_temporal!$B$15,'
        'IF(cobertura_temporal!$B$14<>"",cobertura_temporal!$B$14,"")),'
        'IF(posicao_referencia!$I$5<>"",posicao_referencia!$I$5,'
        'IF(CONTROLE!$B$3="","",CONTROLE!$B$3))))'
    )


def _formula_execucao_atual() -> str:
    itens_por_ciclo = {
        "C0": ("E", "C"),
        "C1": ("G", "D"),
        "C2": ("I", "E"),
        "C3": ("K", "F"),
        "C4": ("M", "G"),
    }
    ramo_itens = '""'
    for ciclo in reversed(tuple(itens_por_ciclo)):
        qtd, vu = itens_por_ciclo[ciclo]
        calculo = (
            f'ROUND(SUMPRODUCT(itens_Consumidos!${qtd}$2:${qtd}$200,'
            f'historico_VU!${vu}$2:${vu}$200),2)'
        )
        ramo_itens = f'IF(UPPER(CONTROLE!$B$2)="{ciclo}",{calculo},{ramo_itens})'
    return (
        '=IF(OR(CONTROLE!$B$2="",$B$35=""),"",'
        'IF($B$5="Financeiro",ROUND(SUMIFS(financeiro!$E$2:$E$73,'
        'financeiro!$B$2:$B$73,LOWER(CONTROLE!$B$2),'
        'financeiro!$A$2:$A$73,"<="&$B$35),2),'
        'IF($B$5="PCs",ROUND(SUMIFS(itens_PC!$F$2:$F$100,'
        'itens_PC!$C$2:$C$100,UPPER(CONTROLE!$B$2),'
        'itens_PC!$B$2:$B$100,"<="&$B$35),2),'
        f'IF($B$5="Itens",{ramo_itens},""))))'
    )


def _criar_nova_resultados(wb, manuais: dict[str, object]):
    memoria = wb.Worksheets(ABA_RESULTADOS)
    memoria.Name = ABA_MEMORIA
    nova = wb.Worksheets.Add(After=memoria)
    nova.Name = ABA_RESULTADOS
    # Alguns bindings COM ignoram o argumento nomeado After no Add. O Move
    # posicional garante deterministicamente MEMORIA_RESULTADOS, RESULTADOS no
    # fim do arquivo.
    nova.Move(None, memoria)
    nova.Visible = XL_SHEET_VISIBLE

    # Os nomes do template historico foram criados com referencias relativas
    # (ex.: RESULTADOS!B26). Isso era suficiente para leitura direta, mas nao
    # para formulas executivas em outra aba. A Etapa 26F os torna absolutos.
    nomes_legados = {
        "METODO_RETROATIVO": "B4",
        "TOLERANCIA_DIVERGENCIA": "D4",
        "VALOR_MANUAL_RETRO": "B5",
        "JUSTIFICATIVA_RETRO": "D5",
        "RETRO_FIN": "B15",
        "RETRO_PC": "C15",
        "RETRO_ITENS": "D15",
        "RETRO_OFICIAL": "B16",
        "VTA_CALCULADO": "B23",
        "AJUSTE_MANUAL_VTA": "B24",
        "VTA_MANUAL_OFICIAL": "B25",
        "VTA_FINAL": "B26",
        "QTD_REM_OFICIAL": "B35",
        "REM_BASE_OFICIAL": "C35",
        "REM_ATUALIZADO_OFICIAL": "D35",
    }
    existentes = _nomes_definidos(wb)
    for nome, endereco in nomes_legados.items():
        if nome not in existentes:
            raise ValueError(f"Nome definido obrigatorio ausente: {nome}")
        wb.Names(nome).Delete()
        wb.Names.Add(
            Name=nome,
            RefersTo=f"={ABA_MEMORIA}!${endereco[0]}${endereco[1:]}",
        )

    # Aparencia geral.
    nova.Cells.Font.Name = "Calibri"
    nova.Cells.Font.Size = 10
    nova.Cells.Font.Color = CORES["cinza_texto"]
    nova.Columns("A").ColumnWidth = 31
    nova.Columns("B:C").ColumnWidth = 22
    nova.Columns("D").ColumnWidth = 34
    nova.Columns("E").ColumnWidth = 22
    nova.Columns("F").ColumnWidth = 14
    nova.Columns("G:H").ColumnWidth = 13
    nova.Rows("1:52").RowHeight = 20
    nova.Range("A1:H52").VerticalAlignment = -4108

    # Titulo e status unico.
    nova.Range("A1:H1").Interior.Color = CORES["azul_escuro"]
    nova.Range("A1").Value = "RESULTADOS CONSOLIDADOS — REAJUSTE CONTRATUAL"
    _estilo_fonte(nova.Range("A1:H1"), branco=True, negrito=True, tamanho=15)
    nova.Rows(1).RowHeight = 30
    nova.Range("A2").Value = (
        "Leitura executiva. A memória de cálculo integral foi preservada na "
        "aba oculta MEMORIA_RESULTADOS."
    )
    nova.Range("A2:H2").Font.Italic = True
    nova.Range("A2:H2").Font.Color = CORES["cinza_texto"]
    nova.Range("A3").Value = "STATUS GLOBAL"
    nova.Range("A3").Interior.Color = CORES["azul"]
    _estilo_fonte(nova.Range("A3"), branco=True, negrito=True)
    # O selo global e o pior caso dos selos por tabela (H8/H14/H24/H33) e da
    # tabela manual. Precedencia: REVISE > ESTIMADO > VALIDADO.
    nova.Range("B3").Formula = (
        '=IF(OR($H$8="REVISE",$H$14="REVISE",$H$24="REVISE",'
        '$H$33="REVISE",COUNTIF($H$43:$H$50,"REVISE")>0),"REVISE",'
        'IF($H$33="ESTIMADO","ESTIMADO","VALIDADO"))'
    )
    nova.Range("B3").HorizontalAlignment = -4108
    _estilo_fonte(nova.Range("B3"), negrito=True, tamanho=12)
    nova.Range("C3:H3").Merge()
    nova.Range("C3").Value = (
        "VALIDADO = bases completas | ESTIMADO = posição física parcial | "
        "REVISE = entrada ou decisão pendente"
    )
    nova.Range("C3").WrapText = True
    nova.Range("C3").Font.Italic = True
    _bordas_tabela(nova.Range("A3:H3"))

    # Parametros que precisam aparecer uma unica vez.
    rotulos = {
        "A5": "Método oficial",
        "C5": "Ciclo atual",
        "E5": "Data de corte",
        "G5": "Fator histórico integral",
        "A6": "Fator da apuração atual",
        "C6": "Variação histórica integral",
    }
    for celula, valor in rotulos.items():
        nova.Range(celula).Value = valor
        nova.Range(celula).Font.Bold = True
        nova.Range(celula).Interior.Color = CORES["cinza"]
    nova.Range("B5").Formula = f"={ABA_MEMORIA}!$B$4"
    nova.Range("D5").Formula = '=UPPER(CONTROLE!$B$2)'
    nova.Range("F5").Formula = '=IF(CONTROLE!$B$3="","",CONTROLE!$B$3)'
    nova.Range("H5").Formula = '=IF(CONTROLE!$B$11="","",CONTROLE!$B$11)'
    nova.Range("B6").Formula = '=IF(parametros!$E$16="","",parametros!$E$16)'
    nova.Range("D6").Formula = '=IF($H$5="","",$H$5-1)'
    nova.Range("E6:H6").Merge()
    nova.Range("E6").Value = (
        "Os fatores da apuração atual e do histórico integral são deliberadamente separados."
    )
    nova.Range("E6").Font.Italic = True
    nova.Range("A5:H6").WrapText = True
    nova.Rows("5:6").RowHeight = 30
    nova.Range("F5").NumberFormatLocal = DATA_BR
    nova.Range("H5").NumberFormatLocal = FATOR
    nova.Range("B6").NumberFormatLocal = FATOR
    nova.Range("D6").NumberFormatLocal = PERCENTUAL
    _bordas_tabela(nova.Range("A5:H6"))

    # Tabela 1.
    _titulo_secao(nova, 8, "1. VALOR TOTAL DO CONTRATO — VTA x ATUALIZAÇÃO CHEIA")
    nova.Range("A9").Value = "Referência"
    nova.Range("B9").Value = "Valor"
    nova.Range("C9:H9").Merge()
    nova.Range("C9").Value = "Como interpretar"
    _cabecalho_tabela(nova, "A9:H9")
    nova.Range("A10").Value = "VTA pelo método aplicado"
    nova.Range("B10").Formula = '=IFERROR(VTA_FINAL,"")'
    nova.Range("C10:H10").Merge()
    nova.Range("C10").Value = (
        "Executado + atualização reconhecida + remanescente atualizado, "
        "conforme o método escolhido."
    )
    nova.Range("A11").Value = "Contrato original integralmente reajustado"
    nova.Range("B11").Formula = '=IFERROR(comparativo_VTA!$B$208,"")'
    nova.Range("C11:H11").Merge()
    nova.Range("C11").Value = (
        "Referência comparativa: valor original multiplicado pelo fator histórico integral."
    )
    nova.Range("B10:B11").NumberFormatLocal = MOEDA_BR
    nova.Range("A10:H10").Interior.Color = CORES["azul_muito_claro"]
    nova.Range("A10:B10").Font.Bold = True
    nova.Range("A11:H11").Interior.Color = CORES["cinza"]
    nova.Range("A11:H11").Font.Italic = True
    nova.Range("B11").Font.Bold = True
    _bordas_tabela(nova.Range("A10:H11"))

    # Tabela 2.
    _titulo_secao(nova, 14, "2. VALOR RETROATIVO A PAGAR")
    for celula, valor in zip(
        ("A15", "B15", "C15", "D15"),
        ("Ciclo", "Valor pago / considerado", "Valor reajustado", "Delta"),
    ):
        nova.Range(celula).Value = valor
    _cabecalho_tabela(nova, "A15:D15")
    nova.Range("A15:D15").WrapText = True
    nova.Rows(15).RowHeight = 30
    for indice, ciclo in enumerate(("C0", "C1", "C2", "C3", "C4"), start=16):
        nova.Range(f"A{indice}").Value = ciclo
        nova.Range(f"B{indice}").Formula = _formula_pago(ciclo, indice)
        memoria_linha = 10 + (indice - 16)
        nova.Range(f"D{indice}").Formula = (
            f'=IF(B{indice}="","",IF({ABA_MEMORIA}!$E${memoria_linha}="",0,'
            f'{ABA_MEMORIA}!$E${memoria_linha}))'
        )
        nova.Range(f"C{indice}").Formula = (
            f'=IF(OR(B{indice}="",D{indice}=""),"",ROUND(B{indice}+D{indice},2))'
        )
    nova.Range("A21").Value = "Ajuste manual não rateado"
    nova.Range("D21").Formula = (
        f'=IF(ISNUMBER({ABA_MEMORIA}!$B$5),'
        'ROUND(RETRO_OFICIAL-SUM($D$16:$D$20),2),"")'
    )
    nova.Range("A22").Value = "TOTAL"
    nova.Range("B22").Formula = '=IF(COUNT(B16:B20)=0,"",ROUND(SUM(B16:B20),2))'
    nova.Range("D22").Formula = '=IFERROR(RETRO_OFICIAL,"")'
    nova.Range("C22").Formula = (
        '=IF(OR(B22="",D22=""),"",ROUND(B22+D22,2))'
    )
    nova.Range("B16:D22").NumberFormatLocal = MOEDA_BR
    nova.Range("A16:D22").Interior.Color = CORES["branco"]
    nova.Range("A21:D21").Interior.Color = CORES["amarelo_input"]
    nova.Range("A22:D22").Interior.Color = CORES["azul_claro"]
    nova.Range("A22:D22").Font.Bold = True
    _bordas_tabela(nova.Range("A16:D22"))

    # Tabela 3.
    _titulo_secao(nova, 24, "3. VALOR REMANESCENTE TOTAL POR CICLO")
    for celula, valor in zip(
        ("A25", "B25", "C25", "D25"),
        ("Ciclo", "Remanescente sem reajuste", "Remanescente atualizado", "Delta"),
    ):
        nova.Range(celula).Value = valor
    _cabecalho_tabela(nova, "A25:D25")
    nova.Range("A25:D25").WrapText = True
    nova.Rows(25).RowHeight = 32
    base_cols = {"C0": "C", "C1": "D", "C2": "E", "C3": "F", "C4": "G"}
    # Zero nao e fabricado: ciclo futuro (apos o vigente) e base itemizada
    # ausente ficam vazios em vez de herdarem 0 das fontes tecnicas.
    for indice_ciclo, (linha, ciclo) in enumerate(
        zip(range(26, 31), ("C0", "C1", "C2", "C3", "C4"))
    ):
        nova.Range(f"A{linha}").Value = ciclo
        nova.Range(f"B{linha}").Formula = (
            f'=IF(IFERROR(VALUE(MID(UPPER(CONTROLE!$B$2),2,1)),-1)<{indice_ciclo},"",'
            'IF(COUNTIF(itens_Remanesc!$A$2:$A$200,"<>")=0,"",'
            f'IFERROR(comparativo_VTA!${base_cols[ciclo]}$204,"")))'
        )
        nova.Range(f"C{linha}").Formula = _formula_rem_atualizado(ciclo).format(
            linha=linha
        )
        nova.Range(f"D{linha}").Formula = (
            f'=IF(OR(B{linha}="",C{linha}=""),"",ROUND(C{linha}-B{linha},2))'
        )
    nova.Range("B26:D30").NumberFormatLocal = MOEDA_BR
    _bordas_tabela(nova.Range("A26:D30"))

    # Tabela 4.
    _titulo_secao(nova, 33, "4. CICLO ATUAL EM EXECUÇÃO")
    nova.Range("A34").Value = "Ciclo em execução"
    nova.Range("B34").Formula = '=UPPER(CONTROLE!$B$2)'
    nova.Range("A35").Value = "Data de referência"
    nova.Range("B35").Formula = _formula_data_referencia()
    nova.Range("A36").Value = "Valor executado atualizado até a data"
    nova.Range("B36").Formula = _formula_execucao_atual()
    nova.Range("A37").Value = "Remanescente atualizado na abertura do ciclo"
    nova.Range("B37").Formula = (
        '=IFERROR(INDEX($C$26:$C$30,MATCH(UPPER(CONTROLE!$B$2),'
        '$A$26:$A$30,0)),"")'
    )
    nova.Range("A38").Value = "Saldo remanescente atualizado"
    nova.Range("B38").Formula = (
        '=IF(OR(B36="",B37=""),"",ROUND(B37-B36,2))'
    )
    nova.Range("C34:H34").Merge()
    nova.Range("C34").Value = (
        "O saldo é a diferença entre o remanescente atualizado na abertura "
        "do ciclo e o executado atualizado até a data."
    )
    nova.Range("C34").WrapText = True
    nova.Range("A34:A38").Font.Bold = True
    nova.Range("A34:A38").WrapText = True
    nova.Rows("34:38").RowHeight = 30
    nova.Range("A34:A38").Interior.Color = CORES["cinza"]
    nova.Range("B35").NumberFormatLocal = DATA_BR
    nova.Range("B36:B38").NumberFormatLocal = MOEDA_BR
    nova.Range("A38:B38").Interior.Color = CORES["azul_claro"]
    nova.Range("A38:B38").Font.Bold = True
    _bordas_tabela(nova.Range("A34:H38"))

    # Selos discretos por tabela: localizam imediatamente o bloco em REVISE/
    # ESTIMADO sem exigir leitura da memoria. Estados-bons da memoria sao
    # detectados por substrings ASCII estaveis dos status homologados.
    nova.Range("H8").Formula = (
        '=IF(OR(VTA_FINAL="",NOT(ISNUMBER(VTA_FINAL)),CONTROLE!$B$11="",'
        f'NOT(OR(ISNUMBER(SEARCH("CALCULADO",{ABA_MEMORIA}!$E$26)),'
        f'ISNUMBER(SEARCH("MANUAL VALIDADO",{ABA_MEMORIA}!$E$26)),'
        f'ISNUMBER(SEARCH("COM AJUSTE MANUAL",{ABA_MEMORIA}!$E$26))))),'
        '"REVISE","VALIDADO")'
    )
    # 26F.2: a Tabela 2 e o selo consomem a MESMA semantica tri-state da
    # coluna de valor (B16:B20): numero = estado conhecido (inclusive ZERO
    # CONFIRMADO pela cobertura GCC, que a propria coluna materializa como 0);
    # vazio = fonte ausente/incompleta. Ciclo obrigatorio = COMPUTAR_NESTA_
    # APURACAO="Sim"; fora do escopo nao e exigido. MANUAL VALIDADO encerra o
    # retroativo no TOTAL oficial sem fabricar rateio por ciclo (regra
    # historica da planilha: "o valor manual prevalece apenas no total
    # oficial, sem criar rateio ficticio por ciclo" — MEMORIA!A6).
    _obrig = "".join(
        f'AND(parametros!$A${lin_par}="Sim",NOT(ISNUMBER($B${lin_res}))),'
        for lin_par, lin_res in zip(range(2, 7), range(16, 21))
    )
    nova.Range("H14").Formula = (
        f'=IF(ISNUMBER(SEARCH("MANUAL VALIDADO",{ABA_MEMORIA}!$F$16)),'
        '"VALIDADO",'
        f'IF(OR({_obrig}'
        f'NOT(ISNUMBER(SEARCH("CALCULADO",{ABA_MEMORIA}!$F$16)))),'
        '"REVISE","VALIDADO"))'
    )
    nova.Range("H24").Formula = (
        f'=IF(ISNUMBER(SEARCH("CONFERIR",{ABA_MEMORIA}!$F$36)),'
        '"VALIDADO","REVISE")'
    )
    nova.Range("H33").Formula = (
        '=IF(OR($B$35="",$B$36="",$B$37="",$B$38="",$B$38<0),"REVISE",'
        'IF(posicao_referencia!$I$2=TRUE,"VALIDADO","ESTIMADO"))'
    )
    for selo in ("H8", "H14", "H24", "H33"):
        rng = nova.Range(selo)
        rng.HorizontalAlignment = -4108
        _estilo_fonte(rng, negrito=True)

    # Premissa curta e inequivoca da estimativa da Tabela 4 (posicao fisica
    # nao confirmada integralmente -> execucao estimada pela fonte financeira).
    nova.Range("A39:H39").Merge()
    # 26F.1: a execucao de B36 segue o metodo oficial (Financeiro, PCs ou
    # Itens); a redacao nao pode afirmar "financeira" quando a fonte e outra.
    nova.Range("A39").Formula = (
        '=IF(AND($B$36<>"",posicao_referencia!$I$2<>TRUE),'
        '"Estimativa: assume execucao registrada pelo metodo oficial ('
        '"&$B$5&") aproximadamente igual ao consumo fisico do ciclo.","")'
    )
    nova.Range("A39").Font.Italic = True

    # Tabela unica de ajustes manuais.
    _titulo_secao(nova, 41, "AJUSTES MANUAIS — USAR SOMENTE QUANDO NECESSÁRIO")
    cabecalhos = (
        "Tipo",
        "Ciclo",
        "Valor (+/-)",
        "Justificativa",
        "Responsável",
        "Data",
        "Aplicar?",
        "Situação",
    )
    for coluna, valor in enumerate(cabecalhos, start=1):
        nova.Cells(42, coluna).Value = valor
    _cabecalho_tabela(nova, "A42:H42")
    tipos = (
        ("Retroativo manual oficial", "TOTAL"),
        ("Ajuste do VTA", "TOTAL"),
        ("VTA manual substitutivo", "TOTAL"),
        ("Complemento histórico", "C0"),
        ("Complemento histórico", "C1"),
        ("Complemento histórico", "C2"),
        ("Complemento histórico", "C3"),
        ("Complemento histórico", "C4"),
    )
    for linha, (tipo, ciclo) in enumerate(tipos, start=43):
        nova.Range(f"A{linha}").Value = tipo
        nova.Range(f"B{linha}").Value = ciclo
        # 26F.1: complementos historicos (linhas 46-50) herdam a regra do
        # baseline N258:N262 — valor >= 0. Negativo (mesmo colado por cima da
        # Data Validation) nunca recebe VALIDADO. Retroativo manual e ajustes
        # de VTA (43-45) seguem sem essa restricao, como no baseline.
        _regra_valor = f"ISNUMBER(C{linha})"
        if tipo == "Complemento histórico":
            _regra_valor = f"AND(ISNUMBER(C{linha}),C{linha}>=0)"
        nova.Range(f"H{linha}").Formula = (
            f'=IF(COUNTA(C{linha}:G{linha})=0,"",'
            f'IF(G{linha}<>"Sim","REVISE",'
            f'IF(AND({_regra_valor},D{linha}<>"",E{linha}<>"",'
            f'ISNUMBER(F{linha})),"VALIDADO","REVISE")))'
        )
        nova.Rows(linha).RowHeight = 32
    nova.Range("C43:G50").Interior.Color = CORES["amarelo_input"]
    nova.Range("C43:C50").NumberFormatLocal = MOEDA_BR
    nova.Range("F43:F50").NumberFormatLocal = DATA_BR
    nova.Range("D43:E50").WrapText = True
    nova.Range("H43:H50").HorizontalAlignment = -4108
    _bordas_tabela(nova.Range("A43:H50"))

    # Valores existentes sao migrados para a interface unica.
    migracao = {
        43: ("B5", "D5"),
        44: ("B24", "D24"),
        45: ("B25", "D25"),
        46: ("N258", None),
        47: ("N259", None),
        48: ("N260", None),
        49: ("N261", None),
        50: ("N262", None),
    }
    for linha, (end_valor, end_justificativa) in migracao.items():
        valor = manuais.get(end_valor)
        justificativa = manuais.get(end_justificativa) if end_justificativa else None
        if valor not in (None, ""):
            nova.Range(f"C{linha}").Value = valor
            nova.Range(f"G{linha}").Value = "Sim"
        if justificativa not in (None, ""):
            nova.Range(f"D{linha}").Value = justificativa

    # Lista robusta para Aplicar?, sem depender do separador regional.
    nova.Range("J2").Value = "Sim"
    nova.Range("J3").Value = "Nao"
    nomes = _nomes_definidos(wb)
    if NOME_OPCOES_MANUAL in nomes:
        wb.Names(NOME_OPCOES_MANUAL).Delete()
    wb.Names.Add(
        Name=NOME_OPCOES_MANUAL,
        RefersTo=f"={ABA_RESULTADOS}!$J$2:$J$3",
    )
    validacao = nova.Range("G43:G50").Validation
    try:
        validacao.Delete()
    except Exception:
        pass
    validacao.Add(
        Type=XL_VALIDATE_LIST,
        AlertStyle=XL_VALID_ALERT_STOP,
        Operator=XL_BETWEEN,
        Formula1=f"={NOME_OPCOES_MANUAL}",
    )
    validacao.IgnoreBlank = True
    validacao.InCellDropdown = True
    validacao.ErrorMessage = "Selecione Sim ou Nao."

    # 26F.1: mesma validacao decimal >= 0 que o baseline aplicava em
    # N258:N262, agora na entrada efetiva dos complementos historicos.
    validacao_complemento = nova.Range("C46:C50").Validation
    try:
        validacao_complemento.Delete()
    except Exception:
        pass
    validacao_complemento.Add(
        Type=XL_VALIDATE_DECIMAL,
        AlertStyle=XL_VALID_ALERT_STOP,
        Operator=XL_GREATER_EQUAL,
        Formula1="0",
    )
    validacao_complemento.IgnoreBlank = True
    validacao_complemento.ErrorMessage = (
        "Complemento historico deve ser maior ou igual a 0."
    )

    # Colunas tecnicas ocultas (J2:J3 guarda as opcoes do dropdown manual).
    nova.Columns("J:N").Hidden = True

    # A memoria passa a consumir a tabela manual visivel. ZERO != VAZIO:
    # Aplicar?="Sim" com valor vazio NAO materializa 0 — o override so existe
    # quando a celula de valor esta efetivamente preenchida (a Situacao da
    # linha marca REVISE ate la). Zero explicito digitado segue sendo valido.
    estado, selecao = _capturar_protecao(memoria)

    def _override(linha_res: int, coluna_res: str) -> str:
        origem = f"{ABA_RESULTADOS}!${coluna_res}${linha_res}"
        return (
            f'=IF(AND({ABA_RESULTADOS}!$G${linha_res}="Sim",'
            f'{origem}<>""),{origem},"")'
        )

    memoria.Range("B5").Formula = _override(43, "C")
    memoria.Range("D5").Formula = _override(43, "D")
    memoria.Range("B24").Formula = _override(44, "C")
    memoria.Range("D24").Formula = _override(44, "D")
    memoria.Range("B25").Formula = _override(45, "C")
    memoria.Range("D25").Formula = _override(45, "D")
    # 26F.1: complemento historico preserva a regra do baseline (>= 0).
    # Negativo nao flui para N258:N262 — jamais reduz o VTA silenciosamente.
    # Vazio nao vira zero; zero explicito e valido.
    for memoria_linha, resultados_linha in zip(range(258, 263), range(46, 51)):
        origem = f"{ABA_RESULTADOS}!$C${resultados_linha}"
        memoria.Range(f"N{memoria_linha}").Formula = (
            f'=IF(AND({ABA_RESULTADOS}!$G${resultados_linha}="Sim",'
            f'ISNUMBER({origem}),{origem}>=0),{origem},"")'
        )
    _restaurar_protecao(memoria, estado, selecao)

    # Nomes novos de leitura executiva.
    for nome, referencia in {
        "STATUS_RESULTADOS": f"={ABA_RESULTADOS}!$B$3",
        "VTA_ATUALIZACAO_CHEIA": f"={ABA_RESULTADOS}!$B$11",
        "EXECUCAO_ATUALIZADA_CICLO": f"={ABA_RESULTADOS}!$B$36",
        "SALDO_REMANESCENTE_ATUAL": f"={ABA_RESULTADOS}!$B$38",
    }.items():
        if nome in _nomes_definidos(wb):
            wb.Names(nome).Delete()
        wb.Names.Add(Name=nome, RefersTo=referencia)

    # Condicionais: selo global, selos por tabela e excecoes materiais.
    for endereco_selo in ("B3", "H8", "H14", "H24", "H33"):
        status = nova.Range(endereco_selo)
        status.FormatConditions.Delete()
        for texto, cor in (
            ("REVISE", CORES["vermelho_status"]),
            ("ESTIMADO", CORES["amarelo_status"]),
            ("VALIDADO", CORES["verde_status"]),
        ):
            cond = status.FormatConditions.Add(
                Type=XL_CELL_VALUE, Operator=XL_EQUAL, Formula1=f'="{texto}"'
            )
            cond.Interior.Color = cor
    negativo = nova.Range("B38").FormatConditions.Add(
        Type=XL_CELL_VALUE, Operator=XL_LESS, Formula1="=0"
    )
    negativo.Interior.Color = CORES["vermelho_status"]
    negativo.Font.Bold = True

    # Protege formulas, mantendo somente C:G da tabela manual editavel.
    nova.Cells.Locked = True
    nova.Range("C43:G50").Locked = False
    nova.Protect(DrawingObjects=True, Contents=True, Scenarios=True)
    nova.EnableSelection = XL_UNLOCKED_CELLS

    # Impressao e navegacao.
    nova.PageSetup.Orientation = XL_LANDSCAPE
    nova.PageSetup.Zoom = False
    nova.PageSetup.FitToPagesWide = 1
    nova.PageSetup.FitToPagesTall = False
    nova.PageSetup.PrintArea = "$A$1:$H$50"
    nova.Application.ActiveWindow.DisplayGridlines = False
    nova.Activate()
    nova.Range("A8").Select()
    nova.Application.ActiveWindow.FreezePanes = True
    nova.Range("A1").Select()

    # Guia executiva unica colorida; memorias podem ser reexibidas para auditoria.
    try:
        memoria.Tab.ColorIndex = XL_NONE
    except Exception:
        pass
    nova.Tab.Color = CORES["vinho"]
    memoria.Visible = XL_SHEET_HIDDEN
    wb.Worksheets("comparativo_VTA").Visible = XL_SHEET_HIDDEN
    return nova, memoria


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


def _verificar_destino_excel(excel, caminho: Path) -> None:
    wb = excel.Workbooks.Open(
        str(caminho), UpdateLinks=0, ReadOnly=True, CorruptLoad=0
    )
    try:
        abas = _nomes_abas(wb)
        if abas[-2:] != [ABA_MEMORIA, ABA_RESULTADOS]:
            raise RuntimeError(f"Ordem final de abas inesperada: {abas[-3:]}")
        if wb.Worksheets(ABA_MEMORIA).Visible != XL_SHEET_HIDDEN:
            raise RuntimeError("MEMORIA_RESULTADOS nao ficou oculta.")
        if wb.Worksheets(ABA_RESULTADOS).Visible != XL_SHEET_VISIBLE:
            raise RuntimeError("RESULTADOS executiva nao ficou visivel.")
        if str(wb.Worksheets(ABA_RESULTADOS).Range("A1").Value or "") != (
            "RESULTADOS CONSOLIDADOS — REAJUSTE CONTRATUAL"
        ):
            raise RuntimeError("Titulo da nova RESULTADOS nao foi preservado.")
        obrigatorios = {
            "VTA_FINAL",
            "RETRO_OFICIAL",
            "REM_BASE_OFICIAL",
            "REM_ATUALIZADO_OFICIAL",
            "STATUS_RESULTADOS",
        }
        faltantes = sorted(obrigatorios.difference(_nomes_definidos(wb)))
        if faltantes:
            raise RuntimeError(f"Nomes definidos ausentes: {faltantes}")
    finally:
        wb.Close(SaveChanges=False)


def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem).resolve()
    destino = Path(destino).resolve()
    if not origem.is_file():
        raise FileNotFoundError(origem)
    if origem == destino:
        raise ValueError("Origem e destino devem ser diferentes.")

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_26f_"))
    tmp_xlsx = tmp_dir / origem.name
    shutil.copyfile(origem, tmp_xlsx)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    salvo = False
    try:
        wb = excel.Workbooks.Open(
            str(tmp_xlsx), UpdateLinks=0, ReadOnly=False, CorruptLoad=0
        )
        excel.ScreenUpdating = False
        excel.Calculation = XL_CALC_MANUAL
        aba_ativa = wb.ActiveSheet.Name
        _validar_origem(wb)
        manuais = _capturar_manuais(wb.Worksheets(ABA_RESULTADOS))
        _ajustes_complementares(wb)
        nova, _memoria = _criar_nova_resultados(wb, manuais)
        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()
        _verificar_sem_erros(wb)
        if aba_ativa == ABA_RESULTADOS:
            nova.Activate()
        elif aba_ativa in _nomes_abas(wb) and aba_ativa != ABA_MEMORIA:
            wb.Worksheets(aba_ativa).Activate()
        else:
            wb.Worksheets("CONTROLE").Activate()
        wb.Save()
        salvo = True
        wb.Close(SaveChanges=False)
        wb = None
        _verificar_destino_excel(excel, tmp_xlsx)
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
    print("RESULTADOS 26F aplicada:", args.destino)


if __name__ == "__main__":
    main()
