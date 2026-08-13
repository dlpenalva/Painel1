# -*- coding: utf-8 -*-
"""Etapa 50: novo leiaute executivo da aba RESULTADOS, via Excel COM.

ESCOPO: APRESENTACAO. Nenhuma regra de negocio, fator, ciclo, temporalidade,
admissibilidade, criterio de efeito financeiro, tratamento de PC/Financeiro,
posicao contratual ou aditivo e alterada. Nenhum numero e recalculado: os tres
cards executivos sao ESPELHOS de celulas oficiais que ja existiam
(B10 = VTA oficial, D22 = retroativo oficial, B38 = saldo remanescente).

CONTRATO ESTRUTURAL PRESERVADO (inventario da Fase A):

  * A1  -> texto literal conferido por _coleta_reajuste._validar_resultados_integra
  * B3  -> nome definido STATUS_RESULTADOS e leitura _coleta_reajuste.py:854
  * B5  -> metodo oficial; base de $B$5 em toda a aba
  * H5  -> fator historico integral; _coleta_oficial._garantir_fator_historico_
           desacoplado reescreve comparativo_VTA!B208 para RESULTADOS!$H$5
  * H8  -> selo cuja formula recebe a troca CONTROLE!$B$11 -> $H$5
  * C12 -> texto auditavel que recebe a troca CONTROLE!B11 -> RESULTADOS!H5
  * D6  -> variacao historica integral (guarda da Etapa 31)
  * B10, B11, B12, B13, H10, H11, H13 -> _leitor_masterfile_v10._ler_referencias_vta
  * B12 -> nome definido VTA_ATUALIZACAO_CHEIA
  * B36 -> nome definido EXECUCAO_ATUALIZADA_CICLO
  * B38 -> nome definido SALDO_REMANESCENTE_ATUAL
  * J2:J3 -> nome definido OPCOES_APLICAR_MANUAL (fonte do dropdown)
  * C43:C50, D43:D45, G43:G50 -> consumidos por MEMORIA_RESULTADOS
    (B5/D5/B24/D24/B25/D25/N258:N262); a MEMORIA nao pode ser tocada
  * A23:E23 -> linha "PCs sem efeito financeiro" (guarda da Etapa 26G)
  * A53:C66 -> tabela de medidas canonicas de PCs

B3 continua sendo a celula canonica do status: a formula, o nome definido e a
leitura Python permanecem intactos. O status passa a ser exibido no selo do
cabecalho (G1:H1) e B3 recebe o formato ";;;" para nao repetir a mesma
informacao duas vezes na mesma tela.

O arquivo e sempre trabalhado em copia temporaria. A promocao do destino so
ocorre depois de recalculo completo, ausencia de erros Excel, verificacao de
que MEMORIA_RESULTADOS ficou identica e reabertura do resultado.
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

# O Excel devolve RPC_E_CALL_REJECTED (-2147418111) quando esta ocupado no
# instante da chamada COM. Como o aplicador trabalha SEMPRE em copia
# temporaria descartavel, a recuperacao segura e reexecutar a aplicacao
# inteira do zero (ver aplicar_com_retentativas).
RPC_E_CALL_REJECTED = -2147418111


XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105
XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16
XL_EXPRESSION = 2
XL_CELL_VALUE = 1
XL_EQUAL = 3
XL_LESS = 6
XL_CONTINUOUS = 1
XL_THIN = 2
XL_LANDSCAPE = 2
XL_SHEET_HIDDEN = 0
XL_SHEET_VISIBLE = -1
XL_UNLOCKED_CELLS = 1
XL_CENTER = -4108
XL_LEFT = -4131
XL_RIGHT = -4152
XL_TOP = -4160
XL_NONE = -4142

ABA_MEMORIA = "MEMORIA_RESULTADOS"
ABA_RESULTADOS = "RESULTADOS"

MOEDA_BR = "R$ #.##0,00"
# Secao de texto "—": formula que devolve "" mostra travessao em vez de vazio.
MOEDA_BR_TRACO = 'R$ #.##0,00;-R$ #.##0,00;R$ 0,00;"—"'
PERCENTUAL = "0,0000%"
FATOR = "0,000000"
DATA_BR = "dd/mm/aaaa"
DATA_BR_TRACO = 'dd/mm/aaaa;;"—";"—"'
OCULTO = ";;;"
# Excel pt-BR rejeita a propriedade NumberFormat em ingles; o reset de formato
# das celulas reaproveitadas usa o token localizado.
GERAL = "Padrão"

# Linhas separadoras (visualmente brancas em todos os cenarios) e linhas
# excedentes ocultas do leiaute final homologado (Etapa 50.3).
LINHAS_BRANCAS = (8, 14, 23, 32, 39, 52)
LINHAS_OCULTAS = (31, 40, 51)

TITULO_OFICIAL = "RESULTADOS CONSOLIDADOS — REAJUSTE CONTRATUAL"
SUBTITULO = "Síntese executiva e memória auditável da apuração"


def _cor(hex_rgb: str) -> int:
    """Converte RRGGBB para o inteiro BGR usado pelo Excel COM."""
    texto = hex_rgb.strip().lstrip("#")
    r, g, b = (int(texto[i : i + 2], 16) for i in (0, 2, 4))
    return r + (g << 8) + (b << 16)


# Paleta funcional e restrita (§14): cada cor tem significado.
CORES = {
    "azul_escuro": _cor("1F4E78"),    # cabecalhos e titulos de secao
    "azul": _cor("2E75B6"),           # cabecalho das tabelas
    "azul_claro": _cor("D9EAF7"),     # totais e destaque secundario
    "azul_palido": _cor("EEF4F8"),    # linha oficial
    "cinza": _cor("F2F2F2"),          # informacao secundaria
    "cinza_borda": _cor("D6DCE4"),
    "cinza_texto": _cor("595959"),
    "cinza_fraco": _cor("8497B0"),
    "branco": _cor("FFFFFF"),
    "branco_azulado": _cor("DCE9F5"),
    "amarelo_input": _cor("FFF2CC"),   # entrada manual (contrato de teste)
    "amarelo_status": _cor("FFE699"),  # dado necessario ainda ausente
    "verde_status": _cor("C6E0B4"),    # concluido / valido
    "vermelho_status": _cor("F4CCCC"),  # erro real
    "borda": _cor("B0C4D8"),
    # Paleta do topo homologado (Etapas 50.1-50.3).
    "cinza_card": _cor("FAFAFA"),      # fundo dos cards 2 e 3
    "contexto": _cor("F7F9FB"),        # faixa de contexto
    "amarelo_suave": _cor("FFF2CC"),   # pendencia (suave)
    "verde_suave": _cor("E2EFDA"),     # conclusao (suave)
}


# --------------------------------------------------------------------------- #
# Utilitarios de estilo
# --------------------------------------------------------------------------- #

def _fonte(rng, *, cor=None, negrito=False, tamanho=10, italico=False):
    rng.Font.Name = "Calibri"
    rng.Font.Size = tamanho
    rng.Font.Bold = negrito
    rng.Font.Italic = italico
    rng.Font.Color = CORES["cinza_texto"] if cor is None else cor


def _bordas(rng, cor=None) -> None:
    for indice in (7, 8, 9, 10, 11, 12):
        borda = rng.Borders(indice)
        borda.LineStyle = XL_CONTINUOUS
        borda.Weight = XL_THIN
        borda.Color = CORES["borda"] if cor is None else cor


def _sem_bordas(rng) -> None:
    for indice in (7, 8, 9, 10, 11, 12):
        try:
            rng.Borders(indice).LineStyle = XL_NONE
        except Exception:
            pass


def _desmesclar(ws, endereco: str) -> None:
    try:
        ws.Range(endereco).UnMerge()
    except Exception:
        pass


def _mesclar(ws, endereco: str) -> None:
    try:
        ws.Range(endereco).Merge()
    except Exception:
        pass


def _duas_linhas(celula, principal: str, apoio: str, cor_principal, cor_apoio,
                 tam_principal=10, tam_apoio=8) -> None:
    """Rotulo de card em duas linhas, com hierarquia tipografica real."""
    celula.Value = f"{principal}\n{apoio}"
    celula.WrapText = True
    celula.Font.Name = "Calibri"
    principal_len, apoio_len = len(principal), len(apoio)
    # O binding tipado do pywin32 expoe a propriedade parametrizada Characters
    # como GetCharacters; o acesso direto nao e chamavel.
    trecho = getattr(celula, "GetCharacters", None)
    if trecho is None:
        _fonte(celula, cor=cor_principal, negrito=True, tamanho=tam_principal)
        return
    cab = trecho(1, principal_len)
    cab.Font.Size = tam_principal
    cab.Font.Bold = True
    cab.Font.Color = cor_principal
    sub = trecho(principal_len + 2, apoio_len)
    sub.Font.Size = tam_apoio
    sub.Font.Bold = False
    sub.Font.Color = cor_apoio


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


# --------------------------------------------------------------------------- #
# Guardas de contrato
# --------------------------------------------------------------------------- #

# Celulas cuja FORMULA nao pode mudar: o leiaute so altera estilo/posicao visual.
ANCORAS_FORMULA = (
    "B3", "B5", "D5", "F5", "H5", "B6", "D6",
    "H8", "B10", "B11", "B12", "B13", "C10", "C11", "C13",
    "F10", "F11", "F12", "F13", "H10", "H11", "H12", "H13",
    "H14", "B16", "B17", "B18", "B19", "B20", "C16", "C17", "C18", "C19", "C20",
    "D16", "D17", "D18", "D19", "D20", "D21", "B22", "C22", "D22",
    "A23", "B23", "C23", "D23", "E23",
    "H24", "B26", "B27", "B28", "B29", "B30",
    "C26", "C27", "C28", "C29", "C30", "D26", "D27", "D28", "D29", "D30",
    "H33", "B34", "B35", "B36", "B37", "B38", "A39",
    "H43", "H44", "H45", "H46", "H47", "H48", "H49", "H50",
    "B55", "B56", "B57", "B58", "B59", "B60", "B61", "B62", "B63", "B64",
    "B65", "B66",
    "J2", "J3",
)

# Rotulos que outros componentes leem literalmente.
ANCORAS_TEXTO_LITERAL = {
    "A1": TITULO_OFICIAL,
    "C12": (
        "SUMPRODUCT(posicao_contratual!B x C) x CONTROLE!B11 "
        "(fator historico integral)"
    ),
}

# Rotulos numerados do anexo tecnico (contrato da Etapa 35C).
ANCORAS_ANEXO = {
    55: "1. Total cadastrado de PCs",
    56: "2. Total considerado ate a data de corte",
    57: "3. Total enquadrado nos ciclos",
    58: "4. Total com efeito financeiro",
    59: "5. Total sem efeito financeiro",
    60: "6. Retroativo reconhecido",
    61: "7. Execucao fisica do ciclo atual",
    62: "8. Remanescente fisico atual",
    63: "9. VTA oficial",
    64: "10. Referencias auditaveis",
    65: "11. Reconciliacao",
    66: "12. Status",
}


def _snapshot(ws, enderecos) -> dict[str, object]:
    return {e: ws.Range(e).Formula for e in enderecos}


def _snapshot_aba(ws) -> dict[str, object]:
    usado = ws.UsedRange
    linhas, colunas = usado.Rows.Count, usado.Columns.Count
    dados = usado.Formula
    primeira_linha, primeira_coluna = usado.Row, usado.Column
    if linhas == 1 and colunas == 1:
        return {f"{primeira_linha}:{primeira_coluna}": dados}
    achatado: dict[str, object] = {}
    for i in range(linhas):
        linha = dados[i] if linhas > 1 else dados
        for j in range(colunas):
            valor = linha[j] if colunas > 1 else linha
            if valor not in (None, ""):
                achatado[f"{primeira_linha + i}:{primeira_coluna + j}"] = valor
    return achatado


def _validar_origem(wb) -> None:
    abas = [ws.Name for ws in wb.Worksheets]
    if ABA_RESULTADOS not in abas or ABA_MEMORIA not in abas:
        raise ValueError("Template sem RESULTADOS/MEMORIA_RESULTADOS (pre-26F).")
    if abas[-1] != ABA_RESULTADOS:
        raise ValueError(f"RESULTADOS deixou de ser a ultima aba: {abas[-3:]}")
    res = wb.Worksheets(ABA_RESULTADOS)
    if str(res.Range("A1").Value or "") != TITULO_OFICIAL:
        raise ValueError("RESULTADOS de origem nao corresponde ao layout 26F+.")
    for endereco in ("B3", "B5", "H5", "H8", "B10", "B11", "B12", "B13",
                     "B36", "B38", "H43", "A53"):
        if not str(res.Range(endereco).Formula or ""):
            raise ValueError(f"Ancora RESULTADOS!{endereco} ausente na origem.")
    # Linha 4 e o respiro; linha 7 recebe os cards. Livres antes da Etapa 50 e,
    # depois dela, ocupadas pelo proprio leiaute — o aplicador e reexecutavel.
    if str(res.Range("A4").Formula or ""):
        raise ValueError("Linha 4 deveria estar livre; leiaute inesperado.")
    a7 = str(res.Range("A7").Formula or "")
    if a7 and not a7.startswith("VTA OFICIAL"):
        raise ValueError(f"Linha 7 ocupada por conteudo inesperado: {a7!r}")
    if res.Range("B5").Formula != FORMULA_METODO_ORIGEM:
        raise ValueError(
            "RESULTADOS!B5 nao e a leitura canonica do metodo oficial: "
            f"{res.Range('B5').Formula!r}"
        )


# --------------------------------------------------------------------------- #
# Formulas de APRESENTACAO acrescentadas (nenhuma calcula negocio)
# --------------------------------------------------------------------------- #

# Contador de pendencias: le exclusivamente os selos que ja existiam.
FORMULA_CONTAGEM = (
    '=($H$8<>"VALIDADO")+($H$14<>"VALIDADO")+($H$24<>"VALIDADO")'
    '+($H$33<>"VALIDADO")+COUNTIF($H$43:$H$50,"REVISE")'
)

FORMULA_CONTAGEM_TEXTO = (
    '=IF($J$5=0,"Sem pendências",$J$5&IF($J$5=1," pendência"," pendências"))'
)

FORMULA_TITULO_PENDENCIAS = (
    '=IF($J$5=0,"APURAÇÃO CONCLUÍDA","PENDÊNCIAS PARA CONCLUSÃO")'
)

# Mensagem consolidada: area unica, derivada dos mesmos selos. O ciclo vem de
# CONTROLE!B2 — nenhum ciclo e fixado em codigo.
FORMULA_PENDENCIAS = (
    '=IF($J$5=0,'
    '"Nenhuma pendência registrada pelas validações desta apuração.",'
    '"Informar / concluir: "'
    '&IF($H$8<>"VALIDADO","composição do VTA; ","")'
    '&IF($H$14<>"VALIDADO","retroativo por ciclo; ","")'
    '&IF($H$24<>"VALIDADO","remanescente por ciclo; ","")'
    '&IF($H$33<>"VALIDADO","situação do ciclo atual"'
    '&IF(CONTROLE!$B$2="",""," ("&UPPER(CONTROLE!$B$2)&")")&"; ","")'
    '&IF(COUNTIF($H$43:$H$50,"REVISE")>0,"ajustes manuais; ",""))'
)

FORMULA_TITULO_CICLO = (
    '=IF(UPPER(CONTROLE!$B$2)="","4. CICLO ATUAL EM EXECUÇÃO",'
    '"4. CICLO ATUAL EM EXECUÇÃO ("&UPPER(CONTROLE!$B$2)&")")'
)


def _aguardando(origem: str) -> str:
    return f'=IF({origem}="","Aguardando dado","")'


# RESULTADOS!B5 e celula de NEGOCIO: 14 formulas da propria aba (B16:B20, D21,
# B35, B36, A39, B55:B59) ramificam por $B$5="Financeiro"/"PCs"/"Itens". Sua
# formula NAO pode ser tocada por motivo de UX. §6, porem, proibe exibir
# "SELECIONE O METODO EM CONTROLE!B1" ao usuario. A solucao fica inteiramente
# na camada de apresentacao:
#   * J6 (coluna oculta) sinaliza se B5 esta em um dos tres valores canonicos;
#   * A5, que sempre foi apenas rotulo, passa a dizer o que fazer quando nao ha
#     metodo escolhido;
#   * uma formatacao condicional pinta o texto de B5 com a cor do fundo nesse
#     caso — a celula continua legivel na barra de formulas e para o Python.
FORMULA_METODO_ORIGEM = "=MEMORIA_RESULTADOS!$B$4"
FORMULA_METODO_CANONICO = (
    '=IF(OR($B$5="Financeiro",$B$5="PCs",$B$5="Itens"),1,0)'
)
FORMULA_ROTULO_METODO = (
    '=IF($J$6=1,"MÉTODO OFICIAL","MÉTODO OFICIAL — selecionar na aba CONTROLE")'
)


# --------------------------------------------------------------------------- #
# Blocos do novo leiaute
# --------------------------------------------------------------------------- #

def _limpar_faixa(ws, endereco: str) -> None:
    rng = ws.Range(endereco)
    rng.Interior.ColorIndex = XL_NONE
    _sem_bordas(rng)
    rng.Font.Italic = False


def _cabecalho(ws) -> None:
    """Faixa azul institucional: titulo, subtitulo, selo de status e contagem."""
    for endereco in ("A1:H1", "A2:H2"):
        _desmesclar(ws, endereco)
    for endereco in ("A1:F1", "G1:H1", "A2:F2", "G2:H2"):
        _mesclar(ws, endereco)

    faixa = ws.Range("A1:H2")
    faixa.Interior.Color = CORES["azul_escuro"]
    _sem_bordas(faixa)
    faixa.VerticalAlignment = XL_CENTER
    faixa.WrapText = False

    ws.Range("A1").Value = TITULO_OFICIAL
    _fonte(ws.Range("A1"), cor=CORES["branco"], negrito=True, tamanho=15)
    ws.Range("A1").HorizontalAlignment = XL_LEFT
    ws.Range("A1").IndentLevel = 1

    ws.Range("A2").Value = SUBTITULO
    _fonte(ws.Range("A2"), cor=CORES["branco_azulado"], tamanho=10)
    ws.Range("A2").HorizontalAlignment = XL_LEFT
    ws.Range("A2").IndentLevel = 1

    # Selo do status global no canto direito. B3 continua sendo a fonte unica.
    ws.Range("G1").Formula = "=$B$3"
    ws.Range("G1").NumberFormatLocal = "@"
    _fonte(ws.Range("G1"), cor=CORES["azul_escuro"], negrito=True, tamanho=13)
    ws.Range("G1").HorizontalAlignment = XL_CENTER
    ws.Range("G1").Interior.Color = CORES["cinza"]

    ws.Range("G2").Formula = FORMULA_CONTAGEM_TEXTO
    _fonte(ws.Range("G2"), cor=CORES["branco_azulado"], tamanho=9)
    ws.Range("G2").HorizontalAlignment = XL_CENTER

    ws.Rows(1).RowHeight = 34
    ws.Rows(2).RowHeight = 22

    ws.Range("G1").FormatConditions.Delete()
    for texto, fundo in (
        ("REVISE", CORES["vermelho_status"]),
        ("ESTIMADO", CORES["amarelo_status"]),
        ("VALIDADO", CORES["verde_status"]),
    ):
        cond = ws.Range("G1").FormatConditions.Add(
            Type=XL_CELL_VALUE, Operator=XL_EQUAL, Formula1=f'="{texto}"'
        )
        cond.Interior.Color = fundo
        cond.Font.Bold = True


def _pendencias(ws) -> None:
    """Area unica de pendencias — nenhuma validacao nova e criada."""
    ws.Range("J5").Formula = FORMULA_CONTAGEM
    _fonte(ws.Range("J5"), tamanho=9)

    ws.Range("A3").Formula = FORMULA_TITULO_PENDENCIAS
    _fonte(ws.Range("A3"), cor=CORES["azul_escuro"], negrito=True, tamanho=10)
    ws.Range("A3").HorizontalAlignment = XL_LEFT
    ws.Range("A3").IndentLevel = 1

    # B3 permanece canonico (nome definido + leitor Python). O formato ";;;"
    # apenas evita repetir o status ao lado do selo do cabecalho; a formatacao
    # condicional migrou junto com a exibicao para G1, senao B3 pintaria um
    # bloco colorido sem texto no meio da faixa de pendencias.
    ws.Range("B3").NumberFormatLocal = OCULTO
    ws.Range("B3").HorizontalAlignment = XL_CENTER
    ws.Range("B3").FormatConditions.Delete()

    _desmesclar(ws, "C3:H3")
    _mesclar(ws, "C3:H3")
    ws.Range("C3").Formula = FORMULA_PENDENCIAS
    _fonte(ws.Range("C3"), tamanho=9)
    ws.Range("C3").WrapText = False
    ws.Range("C3").HorizontalAlignment = XL_LEFT
    ws.Range("C3").VerticalAlignment = XL_CENTER

    faixa = ws.Range("A3:H3")
    faixa.Interior.Color = CORES["cinza"]
    _bordas(faixa, CORES["cinza_borda"])
    ws.Rows(3).RowHeight = 22

    ws.Range("A3").FormatConditions.Delete()
    concluida = ws.Range("A3").FormatConditions.Add(
        Type=XL_EXPRESSION, Formula1="=$J$5=0"
    )
    concluida.Interior.Color = CORES["verde_status"]
    pendente = ws.Range("A3").FormatConditions.Add(
        Type=XL_EXPRESSION, Formula1="=$J$5>0"
    )
    pendente.Interior.Color = CORES["amarelo_status"]

    # Linha 4: respiro entre o cabecalho e a faixa de contexto.
    _limpar_faixa(ws, "A4:H4")
    ws.Rows(4).RowHeight = 6


def _contexto(ws) -> None:
    """Faixa compacta: metodo, ciclo, data de corte e variacao historica.

    Os fatores continuam visiveis como informacao tecnica secundaria (§6).
    Nenhuma referencia de celula e mostrada ao usuario.
    """
    _desmesclar(ws, "E6:H6")
    faixa = ws.Range("A5:H6")
    faixa.Interior.Color = CORES["cinza"]
    _bordas(faixa, CORES["cinza_borda"])
    faixa.VerticalAlignment = XL_CENTER
    faixa.WrapText = False

    # A5 e rotulo puro (nenhum consumidor) e agora carrega a orientacao humana
    # quando o metodo ainda nao foi escolhido — B5 permanece intocada.
    ws.Range("J6").Formula = FORMULA_METODO_CANONICO
    _fonte(ws.Range("J6"), tamanho=9)
    ws.Range("A5").Formula = FORMULA_ROTULO_METODO
    _fonte(ws.Range("A5"), cor=CORES["cinza_fraco"], negrito=True, tamanho=8)
    ws.Range("A5").HorizontalAlignment = XL_LEFT
    ws.Range("A5").IndentLevel = 1

    for celula, texto in {
        "C5": "CICLO ATUAL",
        "E5": "DATA DE CORTE DA APURAÇÃO",
        "C6": "VARIAÇÃO HISTÓRICA INTEGRAL",
    }.items():
        ws.Range(celula).Value = texto
        _fonte(ws.Range(celula), cor=CORES["cinza_fraco"], negrito=True, tamanho=8)
        ws.Range(celula).HorizontalAlignment = XL_LEFT
        ws.Range(celula).IndentLevel = 1

    for celula, texto in {
        "A6": "Fator da apuração atual",
        "G5": "Fator histórico integral",
    }.items():
        ws.Range(celula).Value = texto
        _fonte(ws.Range(celula), cor=CORES["cinza_fraco"], tamanho=8)
        ws.Range(celula).HorizontalAlignment = XL_LEFT
        ws.Range(celula).IndentLevel = 1

    for celula in ("B5", "D5", "F5", "D6"):
        _fonte(ws.Range(celula), cor=CORES["azul_escuro"], negrito=True, tamanho=11)
        ws.Range(celula).HorizontalAlignment = XL_LEFT
    for celula in ("B6", "H5"):
        _fonte(ws.Range(celula), cor=CORES["cinza_fraco"], tamanho=9)
        ws.Range(celula).HorizontalAlignment = XL_LEFT

    ws.Range("B5").NumberFormatLocal = "@"
    ws.Range("D5").NumberFormatLocal = "@"
    ws.Range("F5").NumberFormatLocal = DATA_BR_TRACO
    ws.Range("H5").NumberFormatLocal = FATOR
    ws.Range("B6").NumberFormatLocal = FATOR
    ws.Range("D6").NumberFormatLocal = PERCENTUAL

    _mesclar(ws, "E6:H6")
    ws.Range("E6").Value = (
        "Informação técnica: o fator da apuração atual e o fator histórico "
        "integral são deliberadamente distintos."
    )
    _fonte(ws.Range("E6"), cor=CORES["cinza_fraco"], tamanho=8, italico=True)
    ws.Range("E6").HorizontalAlignment = XL_LEFT
    ws.Range("E6").IndentLevel = 1
    ws.Range("E6").WrapText = False

    # Sem metodo canonico, o texto tecnico de B5 some visualmente (fonte na cor
    # do fundo). A formula, o valor e a leitura por Python continuam intactos.
    ws.Range("B5").FormatConditions.Delete()
    tecnico = ws.Range("B5").FormatConditions.Add(
        Type=XL_EXPRESSION, Formula1="=$J$6=0"
    )
    tecnico.Font.Color = CORES["cinza"]

    ws.Rows(5).RowHeight = 26
    ws.Rows(6).RowHeight = 22


def _cards(ws) -> None:
    """Tres cards executivos — espelhos dos resultados oficiais existentes."""
    _desmesclar(ws, "A7:H7")
    _mesclar(ws, "B7:C7")
    _mesclar(ws, "E7:F7")

    ws.Rows(7).RowHeight = 54
    ws.Range("A7:H7").VerticalAlignment = XL_CENTER

    # CARD 1 — VTA OFICIAL: maior destaque, azul institucional.
    cartao = ws.Range("A7:C7")
    cartao.Interior.Color = CORES["azul_escuro"]
    _bordas(cartao, CORES["azul_escuro"])
    _duas_linhas(
        ws.Range("A7"), "VTA OFICIAL", "Selo OFICIAL · posição física atual",
        CORES["branco"], CORES["branco_azulado"], tam_principal=11, tam_apoio=8,
    )
    ws.Range("A7").HorizontalAlignment = XL_LEFT
    ws.Range("A7").IndentLevel = 1
    ws.Range("B7").Formula = "=$B$10"
    ws.Range("B7").NumberFormatLocal = MOEDA_BR_TRACO
    _fonte(ws.Range("B7"), cor=CORES["branco"], negrito=True, tamanho=16)
    ws.Range("B7").HorizontalAlignment = XL_RIGHT
    ws.Range("B7").IndentLevel = 1

    # CARD 2 — RETROATIVO TOTAL A PAGAR: verde quando concluido.
    cartao = ws.Range("D7:F7")
    cartao.Interior.Color = CORES["cinza"]
    _bordas(cartao, CORES["cinza_borda"])
    _duas_linhas(
        ws.Range("D7"), "RETROATIVO TOTAL A PAGAR",
        "Total oficial apurado nos ciclos",
        CORES["cinza_texto"], CORES["cinza_fraco"], tam_principal=10, tam_apoio=8,
    )
    ws.Range("D7").HorizontalAlignment = XL_LEFT
    ws.Range("D7").IndentLevel = 1
    ws.Range("E7").Formula = "=$D$22"
    ws.Range("E7").NumberFormatLocal = MOEDA_BR_TRACO
    _fonte(ws.Range("E7"), cor=CORES["azul_escuro"], negrito=True, tamanho=15)
    ws.Range("E7").HorizontalAlignment = XL_RIGHT
    ws.Range("E7").IndentLevel = 1

    # CARD 3 — REMANESCENTE ATUALIZADO: amarelo so com dado ausente.
    cartao = ws.Range("G7:H7")
    cartao.Interior.Color = CORES["cinza"]
    _bordas(cartao, CORES["cinza_borda"])
    _duas_linhas(
        ws.Range("G7"), "REMANESCENTE ATUALIZADO", "Saldo do ciclo em execução",
        CORES["cinza_texto"], CORES["cinza_fraco"], tam_principal=9, tam_apoio=8,
    )
    ws.Range("G7").HorizontalAlignment = XL_LEFT
    ws.Range("G7").IndentLevel = 1
    ws.Range("H7").Formula = "=$B$38"
    ws.Range("H7").NumberFormatLocal = MOEDA_BR_TRACO
    _fonte(ws.Range("H7"), cor=CORES["azul_escuro"], negrito=True, tamanho=15)
    ws.Range("H7").HorizontalAlignment = XL_RIGHT
    ws.Range("H7").IndentLevel = 1

    # R$ 0,00 nunca e tratado como problema: a cor deriva do SELO da secao,
    # nunca do valor.
    for endereco, selo in (("D7:F7", "$H$14"), ("G7:H7", "$H$33")):
        rng = ws.Range(endereco)
        rng.FormatConditions.Delete()
        concluido = rng.FormatConditions.Add(
            Type=XL_EXPRESSION, Formula1=f'={selo}="VALIDADO"'
        )
        concluido.Interior.Color = CORES["verde_status"]
    # Uma regra por estado: FormatConditions.Add interpreta a formula no
    # separador REGIONAL, entao nenhuma condicao usa funcao com dois argumentos.
    for estado in ("REVISE", "ESTIMADO"):
        ausente = ws.Range("G7:H7").FormatConditions.Add(
            Type=XL_EXPRESSION, Formula1=f'=$H$33="{estado}"'
        )
        ausente.Interior.Color = CORES["amarelo_status"]


def _titulo_secao(ws, linha: int, conteudo: str, *, formula=False, cor=None) -> None:
    faixa = ws.Range(f"A{linha}:H{linha}")
    faixa.Interior.Color = CORES["azul_escuro"] if cor is None else cor
    _sem_bordas(faixa)
    faixa.VerticalAlignment = XL_CENTER
    faixa.WrapText = False
    _fonte(faixa, cor=CORES["branco"], negrito=True, tamanho=11)
    alvo = ws.Range(f"A{linha}")
    if formula:
        alvo.Formula = conteudo
    else:
        alvo.Value = conteudo
    alvo.HorizontalAlignment = XL_LEFT
    alvo.IndentLevel = 1
    ws.Rows(linha).RowHeight = 22


def _cabecalho_tabela(ws, endereco: str) -> None:
    rng = ws.Range(endereco)
    rng.Interior.Color = CORES["azul"]
    _fonte(rng, cor=CORES["branco"], negrito=True, tamanho=9)
    rng.HorizontalAlignment = XL_LEFT
    rng.VerticalAlignment = XL_CENTER
    rng.IndentLevel = 1
    rng.WrapText = True
    _bordas(rng)


def _estilo_selo(ws, endereco: str) -> None:
    """Selo discreto na faixa de titulo da secao (contrato de fonte da 26H)."""
    rng = ws.Range(endereco)
    rng.HorizontalAlignment = XL_CENTER
    rng.VerticalAlignment = XL_CENTER
    rng.WrapText = False
    _fonte(rng, cor=CORES["cinza_texto"], negrito=True, tamanho=9)
    rng.Interior.Color = CORES["cinza"]
    _bordas(rng, CORES["cinza_borda"])


def _secao_vta(ws) -> None:
    """1. COMPOSIÇÃO DO VTA — mesmas formulas, hierarquia nova."""
    _titulo_secao(ws, 8, "1. COMPOSIÇÃO DO VTA")
    _estilo_selo(ws, "H8")

    ws.Range("A9").Value = "Referência"
    ws.Range("B9").Value = "Valor"
    _desmesclar(ws, "C9:E9")
    _mesclar(ws, "C9:E9")
    ws.Range("C9").Value = "Base de cálculo"
    _desmesclar(ws, "F9:G9")
    _mesclar(ws, "F9:G9")
    ws.Range("F9").Value = "Fontes"
    ws.Range("H9").Value = "Situação"
    _cabecalho_tabela(ws, "A9:H9")
    # Altura minima de 28 e contrato de legibilidade desde os ajustes finais.
    ws.Rows(9).RowHeight = 28
    ws.Range("B9").HorizontalAlignment = XL_RIGHT

    ws.Range("A10").Value = "VTA OFICIAL — POSIÇÃO FÍSICA ATUAL"
    ws.Range("A11").Value = "REFERENCIAL — ABERTURA DO CICLO DE REFERÊNCIA"
    ws.Range("A12").Value = "REFERENCIAL — CONTRATO ORIGINAL INTEGRALMENTE REAJUSTADO"
    ws.Range("A13").Value = "Reconciliação (posição atual − última abertura)"

    corpo = ws.Range("A10:H13")
    corpo.WrapText = True
    corpo.VerticalAlignment = XL_CENTER
    _bordas(corpo)
    for linha in (10, 11, 12, 13):
        ws.Rows(linha).RowHeight = 46
        ws.Range(f"A{linha}").HorizontalAlignment = XL_LEFT
        ws.Range(f"A{linha}").IndentLevel = 1
        ws.Range(f"B{linha}").NumberFormatLocal = MOEDA_BR_TRACO
        ws.Range(f"B{linha}").HorizontalAlignment = XL_RIGHT

    oficial = ws.Range("A10:H10")
    oficial.Interior.Color = CORES["azul_palido"]
    _fonte(oficial, tamanho=9)
    _fonte(ws.Range("A10"), cor=CORES["azul_escuro"], negrito=True, tamanho=10)
    _fonte(ws.Range("B10"), cor=CORES["azul_escuro"], negrito=True, tamanho=13)

    for linha in (11, 12, 13):
        faixa = ws.Range(f"A{linha}:H{linha}")
        faixa.Interior.Color = CORES["branco"]
        _fonte(faixa, tamanho=9)
        _fonte(ws.Range(f"A{linha}"), cor=CORES["cinza_fraco"], negrito=True, tamanho=9)
        _fonte(ws.Range(f"B{linha}"), negrito=True, tamanho=11)
    for linha in (10, 11, 12, 13):
        for coluna in ("C", "F", "H"):
            _fonte(ws.Range(f"{coluna}{linha}"), cor=CORES["cinza_fraco"], tamanho=8)


def _secao_retroativo(ws) -> None:
    """2. RETROATIVO POR CICLO — TOTAL destacado em azul claro."""
    _titulo_secao(ws, 14, "2. RETROATIVO POR CICLO")
    _estilo_selo(ws, "H14")

    for celula, valor in zip(
        ("A15", "B15", "C15", "D15"),
        ("Ciclo", "Valor pago / considerado", "Valor reajustado", "Delta"),
    ):
        ws.Range(celula).Value = valor
    _cabecalho_tabela(ws, "A15:D15")
    ws.Rows(15).RowHeight = 28
    ws.Range("B15:D15").HorizontalAlignment = XL_RIGHT

    corpo = ws.Range("A16:D22")
    corpo.Interior.Color = CORES["branco"]
    _fonte(corpo, tamanho=10)
    _bordas(corpo)
    ws.Range("B16:D22").NumberFormatLocal = MOEDA_BR_TRACO
    ws.Range("B16:D22").HorizontalAlignment = XL_RIGHT
    for linha in range(16, 23):
        ws.Rows(linha).RowHeight = 19
        ws.Range(f"A{linha}").HorizontalAlignment = XL_LEFT
        ws.Range(f"A{linha}").IndentLevel = 1

    ws.Range("A21").Value = "Ajuste manual não rateado"
    ws.Range("A21:D21").Interior.Color = CORES["cinza"]
    _fonte(ws.Range("A21:D21"), cor=CORES["cinza_fraco"], tamanho=9)

    ws.Range("A22").Value = "TOTAL"
    total = ws.Range("A22:D22")
    total.Interior.Color = CORES["azul_claro"]
    _fonte(total, cor=CORES["azul_escuro"], negrito=True, tamanho=11)
    ws.Rows(22).RowHeight = 24

    nota = ws.Range("A23:H23")
    nota.Interior.Color = CORES["cinza"]
    _fonte(nota, cor=CORES["cinza_fraco"], tamanho=8, italico=True)
    # A linha inteira e condicional (so existe no metodo PCs com T28>0): sem a
    # secao de travessao, ela permanece visualmente vazia quando nao se aplica.
    ws.Range("B23:D23").NumberFormatLocal = MOEDA_BR
    _bordas(nota, CORES["cinza_borda"])
    ws.Rows(23).RowHeight = 20


def _secao_remanescente(ws) -> None:
    """3. REMANESCENTE TOTAL POR CICLO — ciclo vigente destacado."""
    _titulo_secao(ws, 24, "3. REMANESCENTE TOTAL POR CICLO")
    _estilo_selo(ws, "H24")

    for celula, valor in zip(
        ("A25", "B25", "C25", "D25"),
        ("Ciclo", "Remanescente sem reajuste", "Remanescente atualizado", "Delta"),
    ):
        ws.Range(celula).Value = valor
    _cabecalho_tabela(ws, "A25:D25")
    ws.Rows(25).RowHeight = 28
    ws.Range("B25:D25").HorizontalAlignment = XL_RIGHT

    corpo = ws.Range("A26:D30")
    corpo.Interior.Color = CORES["branco"]
    _fonte(corpo, tamanho=10)
    _bordas(corpo)
    ws.Range("B26:D30").NumberFormatLocal = MOEDA_BR_TRACO
    ws.Range("B26:D30").HorizontalAlignment = XL_RIGHT
    for linha in range(26, 31):
        ws.Rows(linha).RowHeight = 19
        ws.Range(f"A{linha}").HorizontalAlignment = XL_LEFT
        ws.Range(f"A{linha}").IndentLevel = 1

    # Cada linha e a fotografia de um ciclo: somar ciclos criaria um numero que
    # o modelo nao possui. O destaque vai para o ciclo vigente, que e o
    # remanescente operante.
    corpo.FormatConditions.Delete()
    vigente = corpo.FormatConditions.Add(
        Type=XL_EXPRESSION, Formula1='=$A26=UPPER(CONTROLE!$B$2)'
    )
    vigente.Interior.Color = CORES["azul_claro"]
    vigente.Font.Bold = True

    _desmesclar(ws, "E25:H25")
    _mesclar(ws, "E25:H25")
    ws.Range("E25").Value = (
        "Cada linha é a fotografia de um ciclo — ciclos não se somam. "
        "O ciclo vigente aparece destacado."
    )
    _fonte(ws.Range("E25"), cor=CORES["cinza_fraco"], tamanho=8, italico=True)
    ws.Range("E25").HorizontalAlignment = XL_LEFT
    ws.Range("E25").WrapText = True
    ws.Range("E25").Interior.ColorIndex = XL_NONE
    _sem_bordas(ws.Range("E25:H25"))

    _limpar_faixa(ws, "A31:H32")
    ws.Rows(31).RowHeight = 8
    ws.Rows(32).RowHeight = 8


def _secao_ciclo_atual(ws) -> None:
    """4. CICLO ATUAL EM EXECUÇÃO (Cx) — travessao no lugar de selos amarelos."""
    _titulo_secao(ws, 33, FORMULA_TITULO_CICLO, formula=True)
    _estilo_selo(ws, "H33")

    ws.Range("A34").Value = "Ciclo em execução"
    ws.Range("A35").Value = "Data de referência utilizada"
    ws.Range("A36").Value = "Valor executado atualizado até a data"
    ws.Range("A37").Value = "Remanescente atualizado na abertura do ciclo"
    ws.Range("A38").Value = "Saldo remanescente atualizado"

    corpo = ws.Range("A34:B38")
    corpo.Interior.Color = CORES["branco"]
    _fonte(corpo, tamanho=10)
    _bordas(ws.Range("A34:D38"))
    for linha in range(34, 39):
        ws.Rows(linha).RowHeight = 22
        ws.Range(f"A{linha}").HorizontalAlignment = XL_LEFT
        ws.Range(f"A{linha}").IndentLevel = 1
        ws.Range(f"A{linha}").WrapText = False
        ws.Range(f"B{linha}").HorizontalAlignment = XL_RIGHT

    ws.Range("B34").NumberFormatLocal = "@"
    ws.Range("B35").NumberFormatLocal = DATA_BR_TRACO
    ws.Range("B36:B38").NumberFormatLocal = MOEDA_BR_TRACO

    # Ausencia de dado: travessao no valor e observacao discreta ao lado,
    # nunca um selo amarelo por linha.
    for linha in (35, 36, 37, 38):
        _desmesclar(ws, f"C{linha}:D{linha}")
        _mesclar(ws, f"C{linha}:D{linha}")
        ws.Range(f"C{linha}").Formula = _aguardando(f"$B${linha}")
        _fonte(ws.Range(f"C{linha}"), cor=CORES["cinza_fraco"], tamanho=8, italico=True)
        ws.Range(f"C{linha}").HorizontalAlignment = XL_LEFT
        ws.Range(f"C{linha}").IndentLevel = 1

    saldo = ws.Range("A38:B38")
    saldo.Interior.Color = CORES["azul_claro"]
    _fonte(saldo, cor=CORES["azul_escuro"], negrito=True, tamanho=11)

    _desmesclar(ws, "C34:H34")
    # O texto explicativo migrou para E34; C34:D34 nao pode guardar o residuo.
    ws.Range("C34:D34").ClearContents()
    _mesclar(ws, "C34:D34")
    _mesclar(ws, "E34:H34")
    ws.Range("E34").Value = (
        "Saldo = remanescente atualizado na abertura do ciclo − executado "
        "atualizado até a data."
    )
    _fonte(ws.Range("E34"), cor=CORES["cinza_fraco"], tamanho=8, italico=True)
    ws.Range("E34").WrapText = True
    ws.Range("E34").HorizontalAlignment = XL_LEFT
    _sem_bordas(ws.Range("E34:H34"))

    _desmesclar(ws, "A39:H39")
    _mesclar(ws, "A39:H39")
    _fonte(ws.Range("A39"), cor=CORES["cinza_fraco"], tamanho=8, italico=True)
    ws.Range("A39").WrapText = False
    ws.Range("A39").HorizontalAlignment = XL_LEFT
    ws.Rows(39).RowHeight = 18

    _limpar_faixa(ws, "A40:H40")
    ws.Rows(40).RowHeight = 8

    # Saldo negativo continua sendo condicao real de alerta.
    ws.Range("B38").FormatConditions.Delete()
    negativo = ws.Range("B38").FormatConditions.Add(
        Type=XL_CELL_VALUE, Operator=XL_LESS, Formula1="=0"
    )
    negativo.Interior.Color = CORES["vermelho_status"]
    negativo.Font.Bold = True


def _secao_ajustes_manuais(ws) -> None:
    """5. AJUSTES MANUAIS — area secundaria em cinza, sem faixa vermelha."""
    _titulo_secao(ws, 41, "5. AJUSTES MANUAIS", cor=CORES["cinza"])
    faixa = ws.Range("A41:H41")
    _fonte(faixa, cor=CORES["cinza_texto"], negrito=True, tamanho=10)
    _bordas(faixa, CORES["cinza_borda"])
    ws.Range("D41").Value = "USAR SOMENTE QUANDO NECESSÁRIO"
    _fonte(ws.Range("D41"), cor=CORES["cinza_fraco"], tamanho=8, italico=True)
    ws.Range("D41").HorizontalAlignment = XL_LEFT

    cabecalhos = (
        "Tipo", "Ciclo", "Valor (+/-)", "Justificativa",
        "Responsável", "Data", "Aplicar?", "Situação",
    )
    for coluna, valor in enumerate(cabecalhos, start=1):
        ws.Cells(42, coluna).Value = valor
    cab = ws.Range("A42:H42")
    cab.Interior.Color = CORES["cinza_borda"]
    _fonte(cab, cor=CORES["cinza_texto"], negrito=True, tamanho=8)
    cab.HorizontalAlignment = XL_LEFT
    cab.VerticalAlignment = XL_CENTER
    cab.IndentLevel = 1
    cab.WrapText = True
    _bordas(cab, CORES["cinza_borda"])
    ws.Rows(42).RowHeight = 24

    corpo = ws.Range("A43:H50")
    corpo.Interior.Color = CORES["branco"]
    _fonte(corpo, tamanho=9)
    _bordas(corpo, CORES["cinza_borda"])
    for linha in range(43, 51):
        ws.Rows(linha).RowHeight = 24
        ws.Range(f"A{linha}").IndentLevel = 1

    # A entrada manual mantem o amarelo de input do modelo.
    ws.Range("C43:G50").Interior.Color = CORES["amarelo_input"]
    ws.Range("C43:C50").NumberFormatLocal = MOEDA_BR
    ws.Range("C43:C50").HorizontalAlignment = XL_RIGHT
    ws.Range("F43:F50").NumberFormatLocal = DATA_BR
    ws.Range("D43:E50").WrapText = True
    ws.Range("H43:H50").HorizontalAlignment = XL_CENTER
    _fonte(ws.Range("H43:H50"), negrito=True, tamanho=8)

    # Vermelho apenas quando existe de fato ajuste manual em REVISE.
    situacao = ws.Range("H43:H50")
    situacao.FormatConditions.Delete()
    revise = situacao.FormatConditions.Add(
        Type=XL_CELL_VALUE, Operator=XL_EQUAL, Formula1='="REVISE"'
    )
    revise.Interior.Color = CORES["vermelho_status"]
    validado = situacao.FormatConditions.Add(
        Type=XL_CELL_VALUE, Operator=XL_EQUAL, Formula1='="VALIDADO"'
    )
    validado.Interior.Color = CORES["verde_status"]

    _limpar_faixa(ws, "A51:H52")
    ws.Rows(51).RowHeight = 8
    ws.Rows(52).RowHeight = 8


def _anexo_tecnico(ws) -> None:
    """6. Medidas canonicas de PCs — anexo tecnico, fora da area de impressao."""
    ws.Range("A53").Value = "6. TOTAIS CANONICOS DE PCs — MEDIDAS COM NOMES CLAROS"
    faixa = ws.Range("A53:H53")
    faixa.Interior.Color = CORES["cinza"]
    _fonte(faixa, negrito=True, tamanho=9)
    ws.Range("A53").HorizontalAlignment = XL_LEFT
    ws.Range("A53").IndentLevel = 1
    ws.Rows(53).RowHeight = 20
    _bordas(faixa, CORES["cinza_borda"])

    cab = ws.Range("A54:C54")
    cab.Interior.Color = CORES["cinza_borda"]
    _fonte(cab, negrito=True, tamanho=8)
    cab.HorizontalAlignment = XL_LEFT
    cab.IndentLevel = 1

    corpo = ws.Range("A55:C66")
    corpo.Interior.Color = CORES["branco"]
    _fonte(corpo, cor=CORES["cinza_fraco"], tamanho=8)
    for linha in range(55, 67):
        ws.Rows(linha).RowHeight = 16
        ws.Range(f"A{linha}").IndentLevel = 1
    # Bordas finas em A54:C66 sao contrato de teste desde a Etapa 26G.
    _bordas(ws.Range("A54:C66"))


def _acabamento(ws) -> None:
    ws.Columns("A").ColumnWidth = 31.63
    ws.Columns("B").ColumnWidth = 22.63
    ws.Columns("C").ColumnWidth = 22.63
    ws.Columns("D").ColumnWidth = 34.63
    ws.Columns("E").ColumnWidth = 22.63
    ws.Columns("F").ColumnWidth = 14.63
    ws.Columns("G").ColumnWidth = 18.63
    ws.Columns("H").ColumnWidth = 24.63
    ws.Columns("J:N").Hidden = True

    ws.PageSetup.Orientation = XL_LANDSCAPE
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    ws.PageSetup.PrintArea = "$A$1:$H$50"

    ws.Activate()
    janela = ws.Application.ActiveWindow
    janela.DisplayGridlines = False
    janela.FreezePanes = False
    janela.SplitRow = 0
    janela.SplitColumn = 0
    janela.Zoom = 90
    ws.Range("A1").Select()


# =========================================================================== #
# ESTAGIOS DO TOPO HOMOLOGADO (Etapas 50.1, 50.2 e 50.3)
#
# O leiaute base acima (Etapa 50) foi refinado em tres rodadas de homologacao
# visual. As funcoes abaixo replicam EXATAMENTE as transformacoes aplicadas
# nas copias aprovadas — a cadeia inteira roda em uma unica sessao Excel,
# produzindo o leiaute final aprovado (50.3) diretamente do template pre-50.
# =========================================================================== #

def _contorno(rng, cor) -> None:
    """Somente a borda externa — moldura discreta, sem grade interna."""
    for indice in (7, 8, 9, 10):
        borda = rng.Borders(indice)
        borda.LineStyle = XL_CONTINUOUS
        borda.Weight = XL_THIN
        borda.Color = cor
    for indice in (11, 12):
        try:
            rng.Borders(indice).LineStyle = XL_NONE
        except Exception:
            pass


def _t1_ocultar_ancora(ws, endereco: str, fundo) -> None:
    """Preserva a celula canonica e apenas suprime sua renderizacao."""
    cel = ws.Range(endereco)
    cel.NumberFormatLocal = OCULTO
    cel.Interior.Color = fundo
    _fonte(cel, cor=fundo, tamanho=9)


T1_FORMULA_METODO_HUMANO = '=IF($J$6=1,$B$5,"selecionar na aba CONTROLE")'
T1_FORMULA_CICLO = '=IF($D$5="","—",UPPER($D$5))'
T1_FORMULA_TITULO_PENDENCIAS = (
    '=IF($J$5=0,"APURAÇÃO CONCLUÍDA","PENDÊNCIAS PARA CONCLUSÃO:")'
)
T1_FORMULA_PENDENCIAS = (
    '=IF($J$5=0,'
    '"Nenhuma pendência registrada pelas validações desta apuração.",'
    '"Informar / concluir: "'
    '&IF($H$8<>"VALIDADO","composição do VTA; ","")'
    '&IF($H$14<>"VALIDADO","retroativo por ciclo; ","")'
    '&IF($H$24<>"VALIDADO","remanescente por ciclo; ","")'
    '&IF($H$33<>"VALIDADO","situação do ciclo atual"'
    '&IF(CONTROLE!$B$2="",""," ("&UPPER(CONTROLE!$B$2)&")")&"; ","")'
    '&IF(COUNTIF($H$43:$H$50,"REVISE")>0,"ajustes manuais; ",""))'
)


def _t1_limpar_topo(ws) -> None:
    """Desfaz o topo da Etapa 50 sem tocar em nenhuma ancora."""
    for endereco in ("A1:H1", "A2:H2", "C3:H3", "E6:H6", "B7:C7", "E7:F7",
                     "A3:H3", "A4:H4", "A5:H5", "A6:H6", "A7:H7"):
        _desmesclar(ws, endereco)
    livres = ("A3", "C3", "D3", "E3", "F3", "G3", "H3",
              "A4", "B4", "C4", "D4", "E4", "F4", "G4", "H4",
              "A5", "C5", "E5", "G5",
              "A6", "C6", "E6", "F6", "G6", "H6",
              "A7", "B7", "C7", "D7", "E7", "F7", "G7", "H7")
    for endereco in livres:
        cel = ws.Range(endereco)
        cel.ClearContents()
        # Formatos herdados corrompem a reatribuicao: "@" tornaria a proxima
        # formula texto literal; formato monetario com secao "—" renderizaria
        # travessao sobre strings.
        cel.NumberFormatLocal = GERAL
    ws.Range("G1").NumberFormatLocal = GERAL
    ws.Range("G2").NumberFormatLocal = GERAL
    faixa = ws.Range("A3:H7")
    faixa.Interior.ColorIndex = XL_NONE
    _sem_bordas(faixa)
    faixa.Font.Italic = False
    faixa.IndentLevel = 0
    faixa.WrapText = False
    for endereco in ("A3", "B5", "D5", "F5", "H5", "B6", "D6"):
        ws.Range(endereco).FormatConditions.Delete()


def _t1_cabecalho(ws) -> None:
    _mesclar(ws, "A1:F1")
    _mesclar(ws, "G1:H1")
    _mesclar(ws, "A2:F2")
    _mesclar(ws, "G2:H2")

    faixa = ws.Range("A1:H2")
    faixa.Interior.Color = CORES["azul_escuro"]
    _sem_bordas(faixa)
    faixa.VerticalAlignment = XL_CENTER

    ws.Range("A1").Value = TITULO_OFICIAL
    _fonte(ws.Range("A1"), cor=CORES["branco"], negrito=True, tamanho=15)
    ws.Range("A1").HorizontalAlignment = XL_LEFT
    ws.Range("A1").IndentLevel = 1

    ws.Range("A2").Value = SUBTITULO
    _fonte(ws.Range("A2"), cor=CORES["branco_azulado"], tamanho=10)
    ws.Range("A2").HorizontalAlignment = XL_LEFT
    ws.Range("A2").IndentLevel = 1

    # Selo: espelho puro de B3, que continua sendo a celula canonica.
    ws.Range("G1").Formula = "=$B$3"
    _fonte(ws.Range("G1"), cor=CORES["azul_escuro"], negrito=True, tamanho=13)
    ws.Range("G1").HorizontalAlignment = XL_CENTER
    ws.Range("G1").Interior.Color = CORES["cinza"]

    ws.Range("G2").Formula = (
        '=IF($J$5=0,"Sem pendências",$J$5&IF($J$5=1," pendência"," pendências"))'
    )
    _fonte(ws.Range("G2"), cor=CORES["branco_azulado"], tamanho=9)
    ws.Range("G2").HorizontalAlignment = XL_CENTER

    ws.Rows(1).RowHeight = 34
    ws.Rows(2).RowHeight = 22

    ws.Range("G1").FormatConditions.Delete()
    for texto, fundo in (
        ("REVISE", CORES["vermelho_status"]),
        ("ESTIMADO", CORES["amarelo_suave"]),
        ("VALIDADO", CORES["verde_status"]),
    ):
        cond = ws.Range("G1").FormatConditions.Add(
            Type=XL_CELL_VALUE, Operator=XL_EQUAL, Formula1=f'="{texto}"'
        )
        cond.Interior.Color = fundo
        cond.Font.Bold = True


def _t1_contexto(ws) -> None:
    """Linha 3: quatro blocos. B3 continua intacta, apenas oculta. Os rotulos
    entram como PREFIXO LITERAL do formato numerico (o bloco 1 so tem A3
    livre, pois B3 e a ancora do status)."""
    faixa = ws.Range("A3:H3")
    faixa.Interior.Color = CORES["cinza"]
    faixa.VerticalAlignment = XL_CENTER
    _fonte(faixa, cor=CORES["cinza_texto"], tamanho=9)
    ws.Rows(3).RowHeight = 20

    ws.Range("A3").Formula = T1_FORMULA_METODO_HUMANO
    ws.Range("A3").NumberFormatLocal = '"MÉTODO   "@'
    ws.Range("C3").Formula = T1_FORMULA_CICLO
    ws.Range("C3").NumberFormatLocal = '"CICLO ATUAL   "@'
    ws.Range("E3").Formula = "=$F$5"
    ws.Range("E3").NumberFormatLocal = (
        '"DATA DE CORTE   "dd/mm/aaaa;;;"DATA DE CORTE   —"'
    )
    ws.Range("G3").Formula = "=$D$6"
    ws.Range("G3").NumberFormatLocal = (
        '"VARIAÇÃO ACUMULADA   "0,00%;;;"VARIAÇÃO ACUMULADA   —"'
    )
    for endereco in ("A3", "C3", "E3", "G3"):
        _fonte(ws.Range(endereco), cor=CORES["cinza_texto"], negrito=True,
               tamanho=9)
        ws.Range(endereco).HorizontalAlignment = XL_LEFT
        ws.Range(endereco).IndentLevel = 1

    _t1_ocultar_ancora(ws, "B3", CORES["cinza"])
    for endereco in ("A3:B3", "C3:D3", "E3:F3", "G3:H3"):
        _contorno(ws.Range(endereco), CORES["cinza_borda"])


def _t1_cards(ws) -> None:
    ws.Rows(4).RowHeight = 18
    ws.Rows(5).RowHeight = 38
    ws.Rows(6).RowHeight = 18
    ws.Range("A4:H6").VerticalAlignment = XL_CENTER

    _mesclar(ws, "A4:C4")
    _mesclar(ws, "D4:E4")
    _mesclar(ws, "F4:H4")
    _mesclar(ws, "F6:H6")

    cards = (
        ("A4:C6", "VTA OFICIAL", "A5", "=$B$10", 18, "A6",
         '="Posição física atual — "&$H$8', ("B5", "B6"),
         CORES["azul_escuro"], CORES["branco_azulado"], CORES["branco"]),
        ("D4:E6", "RETROATIVO TOTAL A PAGAR", "E5", "=$D$22", 15, "E6",
         '="Ciclos apurados — "&$H$14', ("D5", "D6"),
         CORES["cinza_card"], CORES["cinza_fraco"], CORES["azul_escuro"]),
        ("F4:H6", "REMANESCENTE ATUALIZADO", "G5", "=$B$38", 14, "F6",
         '="Saldo do ciclo em execução — "&$H$33', ("F5", "H5"),
         CORES["cinza_card"], CORES["cinza_fraco"], CORES["azul_escuro"]),
    )
    for (faixa, titulo, valor_em, formula, tamanho_valor, rodape_em, rodape,
         ancoras, fundo, tinta_titulo, tinta_valor) in cards:
        rng = ws.Range(faixa)
        rng.Interior.Color = fundo
        _contorno(rng, CORES["cinza_borda"])

        topo = ws.Range(faixa.split(":")[0])
        topo.Value = titulo
        _fonte(topo, cor=tinta_titulo, negrito=True, tamanho=9)
        topo.HorizontalAlignment = XL_LEFT
        topo.IndentLevel = 1

        valor = ws.Range(valor_em)
        valor.Formula = formula
        valor.NumberFormatLocal = MOEDA_BR_TRACO
        _fonte(valor, cor=tinta_valor, negrito=True, tamanho=tamanho_valor)
        valor.HorizontalAlignment = XL_RIGHT
        valor.IndentLevel = 1

        rod = ws.Range(rodape_em)
        rod.Formula = rodape
        _fonte(rod, cor=tinta_titulo, tamanho=8)
        rod.HorizontalAlignment = XL_LEFT
        rod.IndentLevel = 1

        for ancora in ancoras:
            _t1_ocultar_ancora(ws, ancora, fundo)

    cond2 = ws.Range("D4:E6").FormatConditions
    cond2.Delete()
    ok2 = cond2.Add(Type=XL_EXPRESSION, Formula1='=$H$14="VALIDADO"')
    ok2.Interior.Color = CORES["verde_suave"]
    cond3 = ws.Range("F4:H6").FormatConditions
    cond3.Delete()
    ok3 = cond3.Add(Type=XL_EXPRESSION, Formula1='=$H$33="VALIDADO"')
    ok3.Interior.Color = CORES["verde_suave"]
    for estado in ("REVISE", "ESTIMADO"):
        pend = ws.Range("F4:H6").FormatConditions.Add(
            Type=XL_EXPRESSION, Formula1=f'=$H$33="{estado}"'
        )
        pend.Interior.Color = CORES["amarelo_suave"]


def _t1_pendencias(ws) -> None:
    _mesclar(ws, "B7:H7")
    ws.Range("A7").Formula = T1_FORMULA_TITULO_PENDENCIAS
    _fonte(ws.Range("A7"), cor=CORES["cinza_texto"], negrito=True, tamanho=9)
    ws.Range("A7").HorizontalAlignment = XL_LEFT
    ws.Range("A7").IndentLevel = 1

    ws.Range("B7").Formula = T1_FORMULA_PENDENCIAS
    _fonte(ws.Range("B7"), cor=CORES["cinza_texto"], tamanho=9)
    ws.Range("B7").HorizontalAlignment = XL_LEFT
    ws.Range("B7").IndentLevel = 1

    faixa = ws.Range("A7:H7")
    faixa.VerticalAlignment = XL_CENTER
    ws.Rows(7).RowHeight = 20
    _contorno(faixa, CORES["cinza_borda"])

    faixa.FormatConditions.Delete()
    concluida = faixa.FormatConditions.Add(Type=XL_EXPRESSION, Formula1="=$J$5=0")
    concluida.Interior.Color = CORES["verde_suave"]
    pendente = faixa.FormatConditions.Add(Type=XL_EXPRESSION, Formula1="=$J$5>0")
    pendente.Interior.Color = CORES["amarelo_suave"]


def _t1_respiro(ws) -> None:
    for linha in (13, 23):
        ws.Rows(linha).RowHeight = ws.Rows(linha).RowHeight + 14
        alvo = ws.Range(f"A{linha}:H{linha}")
        alvo.VerticalAlignment = XL_TOP
    for linha, altura in ((31, 7), (32, 14), (40, 14), (51, 7), (52, 14)):
        ws.Rows(linha).RowHeight = altura
        vazio = ws.Range(f"A{linha}:H{linha}")
        vazio.Interior.ColorIndex = XL_NONE
        _sem_bordas(vazio)


def _chip(ws, endereco: str, fonte_selo: str) -> None:
    """Chip de status do card: espelho do selo canonico, com CF de cores."""
    cel = ws.Range(endereco)
    cel.NumberFormatLocal = GERAL
    cel.Formula = f"={fonte_selo}"
    _fonte(cel, cor=CORES["cinza_fraco"], negrito=True, tamanho=8)
    cel.HorizontalAlignment = XL_CENTER
    cel.VerticalAlignment = XL_CENTER
    cel.FormatConditions.Delete()
    for texto, fundo in (
        ("REVISE", CORES["vermelho_status"]),
        ("ESTIMADO", CORES["amarelo_suave"]),
        ("VALIDADO", CORES["verde_status"]),
    ):
        cond = cel.FormatConditions.Add(
            Type=XL_CELL_VALUE, Operator=XL_EQUAL, Formula1=f'="{texto}"'
        )
        cond.Interior.Color = fundo
        cond.Font.Color = CORES["cinza_texto"]
        cond.Font.Bold = True


def _t2_contexto(ws) -> None:
    """Linha 3: fundo unico muito claro, sem molduras internas."""
    faixa = ws.Range("A3:H3")
    faixa.Interior.Color = CORES["contexto"]
    _sem_bordas(faixa)
    faixa.VerticalAlignment = XL_CENTER
    ws.Rows(3).RowHeight = 20
    for endereco in ("A3", "C3", "E3", "G3"):
        _fonte(ws.Range(endereco), cor=CORES["cinza_texto"], negrito=True,
               tamanho=9)
        ws.Range(endereco).HorizontalAlignment = XL_LEFT
        ws.Range(endereco).IndentLevel = 1
    b3 = ws.Range("B3")
    b3.NumberFormatLocal = OCULTO
    b3.Interior.Color = CORES["contexto"]
    _fonte(b3, cor=CORES["contexto"], tamanho=9)


def _t2_cards(ws) -> None:
    """Linhas 4-6: tres cards do mesmo conjunto, valores uniformes de 15pt."""
    for endereco in ("A4:C4", "D4:E4", "F4:H4", "A5:H5", "F6:H6", "A6:H6"):
        _desmesclar(ws, endereco)

    ws.Rows(4).RowHeight = 18
    ws.Rows(5).RowHeight = 34
    ws.Rows(6).RowHeight = 18
    zona = ws.Range("A4:H6")
    zona.VerticalAlignment = XL_CENTER
    zona.WrapText = False

    for endereco in ("A4", "B4", "C4", "D4", "E4", "F4", "G4", "H4",
                     "A5", "C5", "E5", "G5", "A6", "C6", "E6", "F6",
                     "G6", "H6"):
        cel = ws.Range(endereco)
        cel.ClearContents()
        cel.NumberFormatLocal = GERAL

    cards = (
        ("A4:C6", CORES["azul_escuro"], "A4", "VTA OFICIAL", "C4", "$H$8",
         "C5", "=$B$10", "A6", "Posição física atual",
         CORES["branco_azulado"], CORES["branco"]),
        ("D4:E6", CORES["cinza_card"], "D4", "RETROATIVO TOTAL A PAGAR",
         "E4", "$H$14", "E5", "=$D$22", "E6", "Ciclos apurados",
         CORES["cinza_fraco"], CORES["azul_escuro"]),
        ("F4:H6", CORES["cinza_card"], "F4", "REMANESCENTE ATUALIZADO",
         "H4", "$H$33", "G5", "=$B$38", "F6", "Saldo do ciclo em execução",
         CORES["cinza_fraco"], CORES["azul_escuro"]),
    )
    for (faixa, fundo, titulo_em, titulo, chip_em, selo, valor_em, formula,
         rodape_em, rodape, tinta_titulo, tinta_valor) in cards:
        rng = ws.Range(faixa)
        rng.FormatConditions.Delete()
        rng.Interior.Color = fundo
        _contorno(rng, CORES["cinza_borda"])

        alvo = ws.Range(titulo_em)
        alvo.Value = titulo
        _fonte(alvo, cor=tinta_titulo, negrito=True, tamanho=9)
        alvo.HorizontalAlignment = XL_LEFT
        alvo.IndentLevel = 1

        valor = ws.Range(valor_em)
        valor.Formula = formula
        valor.NumberFormatLocal = MOEDA_BR_TRACO
        _fonte(valor, cor=tinta_valor, negrito=True, tamanho=15)
        valor.HorizontalAlignment = XL_RIGHT
        valor.IndentLevel = 1

        rod = ws.Range(rodape_em)
        rod.Value = rodape
        _fonte(rod, cor=tinta_titulo, tamanho=8)
        rod.HorizontalAlignment = XL_LEFT
        rod.IndentLevel = 1

    for endereco, fundo in (("B5", CORES["azul_escuro"]),
                            ("B6", CORES["azul_escuro"]),
                            ("D5", CORES["cinza_card"]),
                            ("D6", CORES["cinza_card"]),
                            ("F5", CORES["cinza_card"]),
                            ("H5", CORES["cinza_card"])):
        cel = ws.Range(endereco)
        cel.NumberFormatLocal = OCULTO
        cel.Interior.Color = fundo
        _fonte(cel, cor=fundo, tamanho=9)

    ok2 = ws.Range("D4:E6").FormatConditions.Add(
        Type=XL_EXPRESSION, Formula1='=$H$14="VALIDADO"'
    )
    ok2.Interior.Color = CORES["verde_suave"]
    ok3 = ws.Range("F4:H6").FormatConditions.Add(
        Type=XL_EXPRESSION, Formula1='=$H$33="VALIDADO"'
    )
    ok3.Interior.Color = CORES["verde_suave"]
    for estado in ("REVISE", "ESTIMADO"):
        pend = ws.Range("F4:H6").FormatConditions.Add(
            Type=XL_EXPRESSION, Formula1=f'=$H$33="{estado}"'
        )
        pend.Interior.Color = CORES["amarelo_suave"]
    _chip(ws, "C4", "$H$8")
    _chip(ws, "E4", "$H$14")
    _chip(ws, "H4", "$H$33")


def _t2_pendencias(ws) -> None:
    faixa = ws.Range("A7:H7")
    _sem_bordas(faixa)
    faixa.VerticalAlignment = XL_CENTER
    ws.Rows(7).RowHeight = 20
    _fonte(ws.Range("A7"), cor=CORES["cinza_texto"], negrito=True, tamanho=9)
    _fonte(ws.Range("B7"), cor=CORES["cinza_texto"], tamanho=9)
    faixa.FormatConditions.Delete()
    concluida = faixa.FormatConditions.Add(Type=XL_EXPRESSION, Formula1="=$J$5=0")
    concluida.Interior.Color = CORES["verde_suave"]
    pendente = faixa.FormatConditions.Add(Type=XL_EXPRESSION, Formula1="=$J$5>0")
    pendente.Interior.Color = CORES["amarelo_suave"]


def _t2_linha_branca(ws, linha: int, ancoras_na_linha=()) -> None:
    """Linha visualmente vazia real: sem cor, sem borda, sem texto visivel."""
    faixa = ws.Range(f"A{linha}:H{linha}")
    faixa.Interior.ColorIndex = XL_NONE
    _sem_bordas(faixa)
    faixa.WrapText = False
    ws.Rows(linha).RowHeight = 15
    for endereco in ancoras_na_linha:
        cel = ws.Range(endereco)
        cel.FormatConditions.Delete()
        cel.NumberFormatLocal = OCULTO
        cel.Interior.ColorIndex = XL_NONE
        _fonte(cel, cor=CORES["branco"], tamanho=9)


def _t2_banda_titulo_combinada(ws, linha: int, titulo: str, captions: dict) -> None:
    """Titulo da secao coabitando a linha de cabecalho da tabela."""
    faixa = ws.Range(f"A{linha}:H{linha}")
    faixa.Interior.Color = CORES["azul_escuro"]
    _sem_bordas(faixa)
    faixa.VerticalAlignment = XL_CENTER

    alvo = ws.Range(f"A{linha}")
    alvo.NumberFormatLocal = GERAL
    alvo.Value = titulo
    _fonte(alvo, cor=CORES["branco"], negrito=True, tamanho=11)
    alvo.HorizontalAlignment = XL_LEFT
    alvo.IndentLevel = 1
    alvo.WrapText = False

    for endereco, (texto, alinhamento) in captions.items():
        cel = ws.Range(endereco)
        cel.Value = texto
        _fonte(cel, cor=CORES["branco_azulado"], negrito=True, tamanho=9)
        cel.HorizontalAlignment = alinhamento
        cel.VerticalAlignment = XL_CENTER
        cel.WrapText = True


def _t2_transicoes(ws) -> None:
    # Remove o espacamento simulado do estagio anterior.
    ws.Rows(13).RowHeight = 46
    ws.Range("A13:H13").VerticalAlignment = XL_CENTER

    # topo -> Bloco 1: linha 8 branca; titulo passa para a banda da linha 9.
    ws.Range("A8").ClearContents()
    _t2_linha_branca(ws, 8, ancoras_na_linha=("H8",))
    _t2_banda_titulo_combinada(
        ws, 9, "1. COMPOSIÇÃO DO VTA",
        {"B9": ("Valor", XL_RIGHT), "C9": ("Base de cálculo", XL_LEFT),
         "F9": ("Fontes", XL_LEFT), "H9": ("Situação", XL_CENTER)},
    )
    ws.Rows(9).RowHeight = 26

    # Bloco 1 -> Bloco 2: linha 14 branca; titulo na banda da linha 15.
    ws.Range("A14").ClearContents()
    _t2_linha_branca(ws, 14, ancoras_na_linha=("H14",))
    _t2_banda_titulo_combinada(
        ws, 15, "2. RETROATIVO POR CICLO",
        {"B15": ("Valor pago / considerado", XL_RIGHT),
         "C15": ("Valor reajustado", XL_RIGHT),
         "D15": ("Delta", XL_RIGHT)},
    )
    ws.Rows(15).RowHeight = 28

    # Bloco 2 -> Bloco 3: a linha 23 vira o separador.
    _t2_linha_branca(ws, 23)
    _fonte(ws.Range("A23:H23"), cor=CORES["cinza_fraco"], tamanho=8,
           italico=True)

    # Bloco 3 -> 4 / 4 -> 5 / 5 -> 6: uma linha branca visivel; a excedente
    # fica oculta (ocultar nao desloca enderecos).
    _t2_linha_branca(ws, 32)
    _t2_linha_branca(ws, 39)
    _fonte(ws.Range("A39"), cor=CORES["cinza_fraco"], tamanho=8, italico=True)
    _t2_linha_branca(ws, 52)
    for linha in LINHAS_OCULTAS:
        ws.Rows(linha).Hidden = True


# Espelho da nota de PCs sem efeito (bloco 2). Concatenacao pura — TEXT() e
# evitado (armadilha de locale registrada na v10.5.2).
FORMULA_ESPELHO_PCS = (
    '=IF($A$23="","",$A$23&CHAR(10)'
    '&"Pago: R$ "&ROUND($B$23,2)&"  ·  Reajustado: R$ "&ROUND($C$23,2)'
    '&"  ·  Delta: R$ "&ROUND($D$23,2)&CHAR(10)&$E$23)'
)
FORMULA_ESPELHO_NOTA = '=IF($A$39="","",$A$39)'


def _t3_espelho(ws, faixa: str, formula: str) -> None:
    try:
        ws.Range(faixa).UnMerge()
    except Exception:
        pass
    zona = ws.Range(faixa)
    zona.ClearContents()
    zona.Merge()
    alvo = ws.Range(faixa.split(":")[0])
    alvo.NumberFormatLocal = GERAL
    alvo.Formula = formula
    alvo.Font.Name = "Calibri"
    alvo.Font.Size = 8
    alvo.Font.Italic = True
    alvo.Font.Bold = False
    alvo.Font.Color = CORES["cinza_fraco"]
    alvo.WrapText = True
    alvo.HorizontalAlignment = XL_LEFT
    alvo.VerticalAlignment = XL_TOP
    alvo.IndentLevel = 1


def _t3_ocultar_e_espelhar(ws) -> None:
    """Separadores brancos em TODOS os cenarios (Etapa 50.3).

    A23:E23 e A39 mantem as formulas canonicas e recebem ";;;"; a informacao
    visivel ganha espelhos DENTRO do proprio bloco (E16:H21 e E35:H38).
    """
    for endereco in ("A23", "B23", "C23", "D23", "E23"):
        ws.Range(endereco).NumberFormatLocal = OCULTO
    _t3_espelho(ws, "E16:H21", FORMULA_ESPELHO_PCS)

    ws.Range("A39").NumberFormatLocal = OCULTO
    _t3_espelho(ws, "E35:H38", FORMULA_ESPELHO_NOTA)


def _verificar_linhas_brancas(ws) -> None:
    """Sem cor, sem borda e sem texto visivel em cada linha separadora."""
    for linha in LINHAS_BRANCAS:
        for coluna in "ABCDEFGH":
            cel = ws.Range(f"{coluna}{linha}")
            if int(cel.DisplayFormat.Interior.ColorIndex) != XL_NONE:
                raise RuntimeError(f"{coluna}{linha} ainda tem preenchimento.")
            if str(cel.Text or "") != "":
                raise RuntimeError(
                    f"{coluna}{linha} ainda exibe texto: {cel.Text!r}"
                )
            for indice in (7, 8, 9, 10):
                if cel.DisplayFormat.Borders(indice).LineStyle != XL_NONE:
                    raise RuntimeError(f"{coluna}{linha} ainda tem borda.")


def _acabamento_final(ws) -> None:
    ws.Activate()
    janela = ws.Application.ActiveWindow
    janela.DisplayGridlines = False
    janela.FreezePanes = False
    janela.SplitRow = 0
    janela.SplitColumn = 0
    janela.Zoom = 90
    ws.Range("A1").Select()


def _aplicar_layout(ws) -> None:
    # Estagio 1 — leiaute base da Etapa 50 (secoes, tabelas, selos, anexos).
    _cabecalho(ws)
    _pendencias(ws)
    _contexto(ws)
    _cards(ws)
    _secao_vta(ws)
    _secao_retroativo(ws)
    _secao_remanescente(ws)
    _secao_ciclo_atual(ws)
    _secao_ajustes_manuais(ws)
    _anexo_tecnico(ws)
    _acabamento(ws)
    # Estagio 2 — topo homologado (rodada 50.1).
    _t1_limpar_topo(ws)
    _t1_cabecalho(ws)
    _t1_contexto(ws)
    _t1_cards(ws)
    _t1_pendencias(ws)
    _t1_respiro(ws)
    # Estagio 3 — linhas brancas e bandas combinadas (rodada 50.2).
    _t2_contexto(ws)
    _t2_cards(ws)
    _t2_pendencias(ws)
    _t2_transicoes(ws)
    # Estagio 4 — separadores brancos em todos os cenarios (rodada 50.3).
    _t3_ocultar_e_espelhar(ws)
    _acabamento_final(ws)


# --------------------------------------------------------------------------- #
# Verificacoes
# --------------------------------------------------------------------------- #

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


def _verificar_ancoras(ws, antes: dict[str, object]) -> None:
    divergentes = [
        endereco
        for endereco, formula in antes.items()
        if ws.Range(endereco).Formula != formula
    ]
    if divergentes:
        raise RuntimeError(
            "Ancoras de RESULTADOS foram alteradas: " + ", ".join(divergentes)
        )
    for endereco, texto in ANCORAS_TEXTO_LITERAL.items():
        if str(ws.Range(endereco).Value or "") != texto:
            raise RuntimeError(
                f"Texto ancorado divergente em {endereco}: "
                f"{ws.Range(endereco).Value!r}"
            )
    for linha, rotulo in ANCORAS_ANEXO.items():
        if str(ws.Range(f"A{linha}").Value or "") != rotulo:
            raise RuntimeError(f"Rotulo do anexo tecnico divergente em A{linha}.")


def _verificar_memoria(ws, antes: dict[str, object], w48_antes) -> None:
    depois = _snapshot_aba(ws)
    if depois != antes:
        divergentes = sorted(
            k for k in set(antes) | set(depois) if antes.get(k) != depois.get(k)
        )
        raise RuntimeError(
            f"MEMORIA_RESULTADOS foi modificada ({len(divergentes)} celulas): "
            + ", ".join(divergentes[:10])
        )
    if ws.Range("W48").Formula != w48_antes:
        raise RuntimeError("MEMORIA_RESULTADOS!W48 foi alterada; abortando.")


def _verificar_destino_excel(excel, caminho: Path) -> None:
    wb = excel.Workbooks.Open(str(caminho), UpdateLinks=0, ReadOnly=True, CorruptLoad=0)
    try:
        abas = [ws.Name for ws in wb.Worksheets]
        if abas[-2:] != [ABA_MEMORIA, ABA_RESULTADOS]:
            raise RuntimeError(f"Ordem final de abas inesperada: {abas[-3:]}")
        if wb.Worksheets(ABA_MEMORIA).Visible != XL_SHEET_HIDDEN:
            raise RuntimeError("MEMORIA_RESULTADOS nao ficou oculta.")
        res = wb.Worksheets(ABA_RESULTADOS)
        if res.Visible != XL_SHEET_VISIBLE:
            raise RuntimeError("RESULTADOS nao ficou visivel.")
        if str(res.Range("A1").Value or "") != TITULO_OFICIAL:
            raise RuntimeError("Titulo da RESULTADOS nao foi preservado.")
        nomes = {str(n.Name).split("!")[-1] for n in wb.Names}
        obrigatorios = {
            "VTA_FINAL", "RETRO_OFICIAL", "REM_BASE_OFICIAL",
            "REM_ATUALIZADO_OFICIAL", "STATUS_RESULTADOS",
            "VTA_ATUALIZACAO_CHEIA", "EXECUCAO_ATUALIZADA_CICLO",
            "SALDO_REMANESCENTE_ATUAL", "OPCOES_APLICAR_MANUAL",
        }
        faltantes = sorted(obrigatorios.difference(nomes))
        if faltantes:
            raise RuntimeError(f"Nomes definidos ausentes: {faltantes}")
    finally:
        wb.Close(SaveChanges=False)


# --------------------------------------------------------------------------- #
# Orquestracao
# --------------------------------------------------------------------------- #

def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem).resolve()
    destino = Path(destino).resolve()
    if not origem.is_file():
        raise FileNotFoundError(origem)
    if origem == destino:
        raise ValueError("Origem e destino devem ser diferentes.")

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_50_"))
    tmp_xlsx = tmp_dir / origem.name
    shutil.copyfile(origem, tmp_xlsx)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    # Blindagem contra interferencia externa na instancia dedicada: nenhum
    # evento de macro e nenhuma entrada interativa durante a formatacao.
    try:
        excel.Interactive = False
        excel.EnableEvents = False
    except Exception:
        pass
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

        res = wb.Worksheets(ABA_RESULTADOS)
        memoria = wb.Worksheets(ABA_MEMORIA)
        ancoras_antes = _snapshot(res, ANCORAS_FORMULA)
        memoria_antes = _snapshot_aba(memoria)
        w48_antes = memoria.Range("W48").Formula

        estado, selecao = _capturar_protecao(res)
        _aplicar_layout(res)
        _verificar_ancoras(res, ancoras_antes)

        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()
        _verificar_sem_erros(wb)
        _verificar_linhas_brancas(res)
        _verificar_memoria(memoria, memoria_antes, w48_antes)

        res.Cells.Locked = True
        res.Range("C43:G50").Locked = False
        _restaurar_protecao(res, estado, selecao)

        if aba_ativa in [ws.Name for ws in wb.Worksheets] and aba_ativa != ABA_MEMORIA:
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


def aplicar_com_retentativas(origem: Path, destino: Path,
                             tentativas: int = 6) -> None:
    """Reexecuta a aplicacao inteira quando o Excel rejeita chamadas COM.

    RPC_E_CALL_REJECTED e transitorio (Excel ocupado). Cada tentativa parte
    de uma copia temporaria nova, entao descartar e recomecar e seguro: o
    destino so e gravado apos a cadeia completa de verificacoes.
    """
    import time
    import pywintypes

    for tentativa in range(1, tentativas + 1):
        try:
            aplicar(origem, destino)
            return
        except pywintypes.com_error as exc:
            transitorio = RPC_E_CALL_REJECTED in (
                getattr(exc, "hresult", None), *(exc.args or ())
            )
            if not transitorio or tentativa == tentativas:
                raise
            print(f"Excel ocupado (tentativa {tentativa}/{tentativas}); "
                  "aguardando 20s e reexecutando do zero...")
            time.sleep(20)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("origem", type=Path)
    parser.add_argument("destino", type=Path)
    args = parser.parse_args()
    aplicar_com_retentativas(args.origem, args.destino)
    print("RESULTADOS (dashboard 50) aplicada:", args.destino)


if __name__ == "__main__":
    main()
