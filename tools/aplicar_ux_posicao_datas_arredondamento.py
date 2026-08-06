"""Aplicador unico: posicao fisica unificada, datas padronizadas e arredondamento.

Alteracao cirurgica e governada sobre o template oficial. Sete frentes:

  1. FONTE UNICA DA POSICAO FISICA — posicao_referencia!B (QTD_REM_ATUAL) deixa
     de ser entrada manual e passa a buscar, POR ITEM (INDEX/MATCH, nunca por
     linha), a quantidade declarada em CICLO_EM_EXECUCAO!C13:C211. Data, ciclo,
     completude e origem passam a derivar de CICLO_EM_EXECUCAO!D5. O fiscal
     informa a posicao fisica uma unica vez.

  2. DATAS — um rotulo por conceito: DATA DE CORTE DA APURACAO (CONTROLE!B3),
     DATA DA POSICAO FISICA (CICLO_EM_EXECUCAO!D5), DATA DE GERACAO/ANALISE
     (cobertura_temporal!B4, automatica) e as coberturas opcionais
     FINANCEIRO/PCS CONFERIDOS ATE.

  3. VALIDACAO TEMPORAL — posicao_referencia!I10 classifica a posicao fisica em
     VALIDADO (na data de corte), ESTIMADO (anterior) ou REVISE (posterior). Uma
     fotografia POSTERIOR ao corte nunca representa a apuracao encerrada: ela
     deixa de alimentar o VTA (MEMORIA_RESULTADOS!W49) sem alterar a data de
     corte em silencio.

  4. PCs POSTERIORES AO CORTE — permanecem cadastrados; a cobertura oficial
     passa a ser o ULTIMO PC CONSIDERADO ATE O CORTE e um campo proprio informa
     se existem PCs posteriores.

  5. RESULTADOS — as duas referencias do VTA continuam calculadas do mesmo
     jeito; muda a apresentacao (H10/H11) e a selecao correta da fonte.

  6. itens_RC — apenas titulos e orientacoes.

  7. ARREDONDAMENTO ITEMIZADO — regra canonica unica
     VU_ATUALIZADO_Cn = ARRED(VU_ORIGINAL x FATOR_ACUMULADO_Cn; 2) e
     VALOR_ITEM_Cn   = ARRED(QTD x VU_ATUALIZADO_Cn; 2), com historico_VU como
     fonte canonica do VU. Elimina QTD x VU_ORIGINAL x FATOR NAO ARREDONDADO em
     itens_Remanesc, aditivos e posicao_referencia.

Gravador unico: Microsoft Excel real via COM (copia temporaria; promove so se
salvar sem erros de formula). FAIL-CLOSED: recusa reaplicacao.

REGRA ZERO CORRUPCAO XLSX: toda string DENTRO de formula e ASCII puro e os
parenteses sao conferidos antes da gravacao.
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16
XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105
XL_PASTE_FORMATS = -4122
XL_CENTER = -4108

FIM = 200
FMT_DATA = "dd/mm/aaaa"          # pt-BR (NumberFormatLocal)
FMT_NUM = "#.##0,00;-#.##0,00"

# Dourado/amarelo ja associado a campos de entrada (CICLO_EM_EXECUCAO!D5).
COR_ABA_ENTRADA = 0x00C0FF       # BGR de FFC000

ABAS_DE_ENTRADA = (
    "CONTROLE", "financeiro", "itens_Remanesc", "itens_Consumidos",
    "itens_PC", "aditivos", "cobertura_temporal",
)

# Referencias indiretas: CICLO_EM_EXECUCAO so existe no arquivo gerado.
CEE_D5 = 'INDIRECT("CICLO_EM_EXECUCAO!$D$5")'
CEE_A9 = 'INDIRECT("CICLO_EM_EXECUCAO!$A$9")'
CEE_ITENS = 'INDIRECT("CICLO_EM_EXECUCAO!$A$13:$A$211")'
CEE_QTDS = 'INDIRECT("CICLO_EM_EXECUCAO!$C$13:$C$211")'

_HIST = {"C0": "C", "C1": "D", "C2": "E", "C3": "F", "C4": "G"}


def _dstr(ref: str) -> str:
    """dd/mm/aaaa locale-proof (DAY/MONTH/YEAR) para rotulos textuais."""
    return f'(RIGHT("0"&DAY({ref}),2)&"/"&RIGHT("0"&MONTH({ref}),2)&"/"&YEAR({ref}))'


def _pick(cell: str, mapa: dict, default: str = '""') -> str:
    expr = default
    for cic in ("C4", "C3", "C2", "C1", "C0"):
        expr = f'IF({cell}="{cic}",{mapa[cic]},{expr})'
    return expr


def _vu_referencia(linha: int) -> str:
    """VU atualizado do ciclo de referencia, por item (historico_VU)."""
    return _pick("$I$4", {c: f"historico_VU!{col}{linha}" for c, col in _HIST.items()})


def _vu_por_ciclo(cel_ciclo: str, cel_item: str) -> str:
    """VU atualizado por (item, ciclo) via historico_VU — fonte canonica."""
    expr = '""'
    for cic in ("C4", "C3", "C2", "C1", "C0"):
        col = 3 + int(cic[1])          # C0 -> coluna 3 (C) ... C4 -> 7 (G)
        expr = (
            f'IF({cel_ciclo}="{cic}",'
            f'IFERROR(VLOOKUP({cel_item},historico_VU!$A:$G,{col},0),""),{expr})'
        )
    return expr


# CICLO enquadrado a partir da DATA DA POSICAO FISICA (nao mais de CONTROLE!B3).
CICLO_DA_POSICAO_FISICA = (
    '=IF($I$9="","",IF(NOT(ISNUMBER($I$9)),"",'
    'IF(AND(ISNUMBER(parametros!$C$2),ISNUMBER(parametros!$D$2),$I$9>=parametros!$C$2,$I$9<=parametros!$D$2),"C0",'
    'IF(AND(ISNUMBER(parametros!$C$3),ISNUMBER(parametros!$D$3),$I$9>=parametros!$C$3,$I$9<=parametros!$D$3),"C1",'
    'IF(AND(ISNUMBER(parametros!$C$4),ISNUMBER(parametros!$D$4),$I$9>=parametros!$C$4,$I$9<=parametros!$D$4),"C2",'
    'IF(AND(ISNUMBER(parametros!$C$5),ISNUMBER(parametros!$D$5),$I$9>=parametros!$C$5,$I$9<=parametros!$D$5),"C3",'
    'IF(AND(ISNUMBER(parametros!$C$6),ISNUMBER(parametros!$D$6),$I$9>=parametros!$C$6,$I$9<=parametros!$D$6),"C4",'
    '"FORA DO HORIZONTE")))))))'
)

# A posicao fisica so e adotada quando informada, completa E nao posterior ao
# corte. Fotografia futura nao representa apuracao encerrada em data anterior.
POSICAO_ATUAL_COMPLETA = (
    '=AND(ISNUMBER($I$9),ISNUMBER(CONTROLE!$B$3),$I$9<=CONTROLE!$B$3,'
    'OR($I$1="C0",$I$1="C1",$I$1="C2",$I$1="C3",$I$1="C4"),'
    'COUNTIF(itens_Remanesc!$A$2:$A$200,"<>")>0,'
    f'IFERROR({CEE_A9}<>"",FALSE))'
)

STATUS_REFERENCIA_FISICA = (
    '=IF($I$9="","NAO INFORMADA",'
    'IF(NOT(ISNUMBER(CONTROLE!$B$3)),"REVISE",'
    'IF($I$9>CONTROLE!$B$3,"REVISE",'
    f'IF(IFERROR({CEE_A9}="",TRUE),"INCOMPLETA",'
    'IF($I$9=CONTROLE!$B$3,"VALIDADO","ESTIMADO")))))'
)

MENSAGEM_VALIDACAO_TEMPORAL = (
    '=IF($I$9="","POSICAO FISICA NAO INFORMADA EM CICLO_EM_EXECUCAO.",'
    'IF(NOT(ISNUMBER(CONTROLE!$B$3)),"DATA DE CORTE DA APURACAO NAO INFORMADA.",'
    'IF($I$9>CONTROLE!$B$3,"POSICAO FISICA POSTERIOR A DATA DE CORTE.",'
    'IF($I$10="INCOMPLETA","POSICAO FISICA INCOMPLETA - QUANTIDADES PENDENTES EM CICLO_EM_EXECUCAO.",'
    'IF($I$10="VALIDADO","POSICAO FISICA NA DATA DE CORTE.",'
    '"POSICAO FISICA ANTERIOR A DATA DE CORTE - UTILIZADA A DATA DE "&'
    + _dstr("$I$9") + '&".")))))'
)

# QTD_REM_ATUAL automatica: correspondencia por ITEM, nunca por linha; vazia
# quando nao ha posicao valida; jamais soma as duas abas.
QTD_REM_ATUAL_AUTOMATICA = (
    '=IF($A2="","",IF(NOT($I$2),"",'
    f'IFERROR(IF(INDEX({CEE_QTDS},MATCH($A2,{CEE_ITENS},0))="","",'
    f'INDEX({CEE_QTDS},MATCH($A2,{CEE_ITENS},0))),"")))'
)

ULTIMO_PC_ATE_CORTE = (
    '=IFERROR(IF(COUNT(itens_PC!$B$2:$B$5001)=0,"",'
    'IF(NOT(ISNUMBER(CONTROLE!$B$3)),MAX(itens_PC!$B$2:$B$5001),'
    'IF(COUNTIFS(itens_PC!$B$2:$B$5001,"<="&CONTROLE!$B$3,itens_PC!$B$2:$B$5001,">0")=0,"",'
    'SUMPRODUCT(MAX((itens_PC!$B$2:$B$5001<=CONTROLE!$B$3)*'
    '(itens_PC!$B$2:$B$5001>0)*itens_PC!$B$2:$B$5001))))),"")'
)

EXISTEM_PCS_POSTERIORES = (
    '=IFERROR(IF(COUNT(itens_PC!$B$2:$B$5001)=0,"NAO",'
    'IF(NOT(ISNUMBER(CONTROLE!$B$3)),"NAO",'
    'IF(COUNTIFS(itens_PC!$B$2:$B$5001,">"&CONTROLE!$B$3)=0,"NAO",'
    '"SIM - ULTIMO PC CADASTRADO EM "&'
    + _dstr("MAX(itens_PC!$B$2:$B$5001)") + '))),"NAO")'
)

# --------------------------------------------------------------------------- #
# RESULTADOS — apresentacao das duas referencias do VTA
# --------------------------------------------------------------------------- #
SITUACAO_VTA_POSICAO_ATUAL = (
    '=IF(AND(ISNUMBER(posicao_referencia!$I$9),ISNUMBER(CONTROLE!$B$3),'
    'posicao_referencia!$I$9>CONTROLE!$B$3),'
    '"REVISE - POSICAO POSTERIOR A DATA DE CORTE",'
    'IF(MEMORIA_RESULTADOS!$W$50="",'
    '"NAO DISPONIVEL - POSICAO FISICA NAO INFORMADA OU INCOMPLETA",'
    'IF(posicao_referencia!$I$10="VALIDADO",'
    '"UTILIZADA - POSICAO FISICA DE "&' + _dstr("posicao_referencia!$I$9") + ','
    '"ESTIMADA - POSICAO FISICA DE "&' + _dstr("posicao_referencia!$I$9") + ')))'
)

SITUACAO_VTA_ABERTURA = (
    '=IF(MEMORIA_RESULTADOS!$W$48="","INCOMPLETO",'
    '"REFERENCIA - ABERTURA DO CICLO C"&MEMORIA_RESULTADOS!$W$46)'
)

SITUACAO_CICLO_ATUAL = (
    '=IF(OR($B$35="",$B$36="",$B$37="",$B$38="",$B$38<0),"REVISE",'
    'IF(posicao_referencia!$I$10="REVISE","REVISE",'
    'IF(posicao_referencia!$I$10="VALIDADO","VALIDADO",'
    'IF(posicao_referencia!$I$10="ESTIMADO","ESTIMADO",'
    'IF(MEMORIA_RESULTADOS!$W$49=1,"VALIDADO","ESTIMADO")))))'
)

# CICLO_EM_EXECUCAO so alimenta o VTA quando a fotografia nao e posterior ao corte.
CEE_DISPONIVEL = (
    f'=IF(ISERROR({CEE_A9}),0,IF({CEE_A9}="",0,'
    f'IF(AND(ISNUMBER({CEE_D5}),ISNUMBER(CONTROLE!$B$3),{CEE_D5}>CONTROLE!$B$3),0,1)))'
)


# --------------------------------------------------------------------------- #
# Rotulos e orientacoes (texto puro — acentuacao permitida fora de formulas)
# --------------------------------------------------------------------------- #
ROTULO_CORTE = "DATA DE CORTE DA APURAÇÃO"
ORIENTACAO_CORTE = (
    "Última data considerada nos cálculos, PCs, competências e resultados "
    "da apuração."
)

ROTULOS_COBERTURA = {
    4: "DATA DE GERAÇÃO/ANÁLISE — AUTO",
    7: "DATA DE ABERTURA DO CICLO DE REFERÊNCIA — AUTO",
    8: "DATA DA POSIÇÃO FÍSICA — AUTO (CICLO_EM_EXECUCAO)",
    11: "POSIÇÃO FÍSICA CONSIDERADA ATÉ — AUTO",
    12: "ÚLTIMA COMPETÊNCIA FINANCEIRA COM PAGAMENTO — AUTO",
    13: "FINANCEIRO CONFERIDO ATÉ — OPCIONAL",
    14: "ÚLTIMO PC CONSIDERADO ATÉ O CORTE — AUTO",
    15: "PCS CONFERIDOS ATÉ — OPCIONAL",
    17: "EXISTEM PCS POSTERIORES AO CORTE?",
}

LEGENDA_FISCAL = (
    "Amarelo: itens_Remanesc, itens_Consumidos, itens_PC, aditivos, financeiro, "
    "CICLO_EM_EXECUCAO (data e quantidades da posição física) e CONTROLE!B3 "
    "(data de corte da apuração)."
)
LEGENDA_GCC = (
    "Amarelo: FINANCEIRO CONFERIDO ATÉ (B13) e PCS CONFERIDOS ATÉ (B15). "
    "A data de geração/análise (B4) é automática."
)

TITULO_RC_CICLOS = (
    "POSIÇÃO AJUSTADA POR CICLO "
    "(AUTO — INCLUI ALTERAÇÕES CONTRATUAIS APLICÁVEIS)"
)
TITULO_RC_ATUAL = "POSIÇÃO FÍSICA ATUAL (AUTO — ORIGEM: CICLO_EM_EXECUCAO)"
ROTULO_RC_DATA_ADITIVO = "DATA DE EFEITO DA ALTERAÇÃO CONTRATUAL — AUTO"
NOTA_RC_DATA_ADITIVO = (
    "Esta data pertence ao acréscimo, à supressão ou ao novo item. "
    "Não corresponde ao início do efeito financeiro do reajuste."
)

_PROT_FLAGS = (
    "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows",
    "AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks",
    "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting",
    "AllowFiltering", "AllowUsingPivotTables",
)


def _sem_protecao(ws):
    """Suspende a protecao da aba preservando as permissoes originais."""
    if not bool(ws.ProtectContents):
        return None
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


def _restaurar_protecao(ws, estado) -> None:
    if estado is None:
        return
    flags, sel = estado
    ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True, **flags)
    if sel is not None:
        try:
            ws.EnableSelection = sel
        except Exception:
            pass


def _conferir_ascii_e_parenteses(*formulas: str) -> None:
    """REGRA ZERO CORRUPCAO: ASCII puro e parenteses balanceados."""
    for f in formulas:
        if not f.startswith("="):
            continue
        try:
            f.encode("ascii")
        except UnicodeEncodeError as exc:
            raise ValueError(f"Formula com caractere nao-ASCII: {f[:80]}") from exc
        if f.count("(") != f.count(")"):
            raise ValueError(f"Parenteses desbalanceados: {f[:80]}")


def _exigir_formulas(ws, faixas) -> None:
    """Fail-closed: prova que cada faixa gravada ficou como FORMULA, nao texto.

    Atribuir `.Formula` a uma celula ja formatada como Texto ("@") faz o Excel
    gravar a expressao como texto literal — a coluna some do calculo em
    silencio. Esta guarda impede que isso chegue ao template.
    """
    for faixa in faixas:
        alvo = ws.Range(faixa)
        if not bool(alvo.Cells(1, 1).HasFormula):
            raise RuntimeError(
                f"{ws.Name}!{faixa} nao ficou como formula (formato Texto?)."
            )


def _validar_layout(wb) -> None:
    nomes = [ws.Name for ws in wb.Worksheets]
    for aba in ("CONTROLE", "parametros", "itens_Remanesc", "itens_Consumidos",
                "itens_PC", "aditivos", "posicao_referencia",
                "posicao_contratual", "itens_RC", "historico_VU",
                "cobertura_temporal", "MEMORIA_RESULTADOS", "RESULTADOS"):
        if aba not in nomes:
            raise ValueError(f"Aba canonica ausente: {aba}")
    pr = wb.Worksheets("posicao_referencia")
    if str(pr.Range("B2").Formula or "").startswith("="):
        raise ValueError("posicao_referencia!B2 ja e formula; pacote ja aplicado?")
    if str(pr.Range("H9").Value or "").strip():
        raise ValueError("posicao_referencia!H9 ocupada; pacote ja aplicado?")
    cob = wb.Worksheets("cobertura_temporal")
    if str(cob.Range("A17").Value or "").strip():
        raise ValueError("cobertura_temporal!A17 ocupada; pacote ja aplicado?")
    if "Data de corte" not in str(wb.Worksheets("CONTROLE").Range("A3").Value or ""):
        raise ValueError("CONTROLE!A3 nao e o rotulo de data de corte.")


# --------------------------------------------------------------------------- #
# 1) posicao_referencia — memoria automatica alimentada por CICLO_EM_EXECUCAO
# --------------------------------------------------------------------------- #
def _aplicar_posicao_referencia(ws) -> None:
    _sem_protecao(ws)

    painel = [
        (1, "CICLO DA POSICAO FISICA (CICLO_EM_EXECUCAO)",
         CICLO_DA_POSICAO_FISICA, None),
        (2, "POSICAO ATUAL COMPLETA?", POSICAO_ATUAL_COMPLETA, None),
        (5, "DATA DA POSICAO DE REFERENCIA",
         '=IF($I$4="","",IF($I$2,$I$9,$I$6))', FMT_DATA),
        (8, "ORIGEM DA POSICAO",
         '=IF($I$4="","POSICAO DE REFERENCIA INDISPONIVEL",'
         'IF($I$2,"POSICAO FISICA INFORMADA - "&' + _dstr("$I$5") + ','
         'IF($I$9<>"","POSICAO FISICA NAO UTILIZAVEL - UTILIZADA ABERTURA "&$I$4,'
         '"ABERTURA DO CICLO "&$I$4&" - "&' + _dstr("$I$5") + ')))', None),
        (9, "DATA DA POSICAO FISICA (CICLO_EM_EXECUCAO)",
         f'=IFERROR(IF({CEE_D5}="","",{CEE_D5}),"")', FMT_DATA),
        (10, "STATUS DA REFERENCIA FISICA", STATUS_REFERENCIA_FISICA, None),
        (11, "VALIDACAO TEMPORAL DA POSICAO", MENSAGEM_VALIDACAO_TEMPORAL, None),
    ]
    _conferir_ascii_e_parenteses(*[f for _, _, f, _ in painel])
    for linha, rotulo, formula, fmt in painel:
        ws.Range(f"H{linha}").Value = rotulo
        ws.Range(f"I{linha}").NumberFormatLocal = "Geral"
        ws.Range(f"I{linha}").Formula = formula
        if fmt:
            ws.Range(f"I{linha}").NumberFormatLocal = fmt
    ws.Range("H1:H11").Font.Bold = True
    _exigir_formulas(ws, [f"I{linha}" for linha, _, _, _ in painel])

    vu = _vu_referencia(2)
    colunas = {
        "B": QTD_REM_ATUAL_AUTOMATICA,
        "E": ('=IF(A2="","",IF($I$4="","POSICAO DE REFERENCIA INDISPONIVEL",'
              'IF($I$2,"POSICAO FISICA (CICLO_EM_EXECUCAO)",'
              'IF(ISNUMBER(B2),"POSICAO FISICA NAO UTILIZAVEL - FOTOGRAFIA "&$I$4,'
              '"FOTOGRAFIA "&$I$4))))'),
        # Arredondamento canonico: QTD x VU_ATUALIZADO (historico_VU), nunca
        # QTD x VU_ORIGINAL x FATOR NAO ARREDONDADO.
        "O": f'=IF(OR(A2="",D2="",NOT(ISNUMBER({vu}))),"",ROUND(D2*{vu},2))',
        "P": f'=IF(OR(A2="",N2="",NOT(ISNUMBER({vu}))),"",ROUND(N2*{vu},2))',
        "R": f'=IF(OR(A2="",L2="",NOT(ISNUMBER({vu}))),"",ROUND(L2*{vu},2))',
    }
    _conferir_ascii_e_parenteses(*colunas.values())
    for col, formula in colunas.items():
        alvo = ws.Range(f"{col}2:{col}{FIM}")
        # E esta formatada como Texto ("@"): atribuir .Formula a uma celula ja
        # em formato Texto faz o Excel gravar a formula como TEXTO LITERAL.
        # Neutraliza-se o formato, grava-se a formula e restaura-se o formato.
        formato = alvo.NumberFormatLocal
        alvo.NumberFormatLocal = "Geral"
        alvo.Formula = formula
        alvo.NumberFormatLocal = formato
    _exigir_formulas(ws, [f"{col}2:{col}{FIM}" for col in colunas])

    # Deixa de ser campo de entrada: sem validacao, sem destaque manual.
    alvo = ws.Range(f"B2:B{FIM}")
    try:
        alvo.Validation.Delete()
    except Exception:
        pass
    alvo.Interior.Color = 0xEDEDED          # cinza das colunas automaticas
    alvo.NumberFormatLocal = FMT_NUM

    # Toda a aba passa a ser automatica: celulas protegidas.
    ws.Cells.Locked = True
    ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True,
               AllowFormattingCells=True, AllowFormattingColumns=True,
               AllowFormattingRows=True, AllowSorting=False,
               AllowFiltering=True, AllowUsingPivotTables=False)


# --------------------------------------------------------------------------- #
# 2) CONTROLE — DATA DE CORTE DA APURACAO
# --------------------------------------------------------------------------- #
def _aplicar_controle(ws) -> None:
    estado = _sem_protecao(ws)
    ws.Range("A3").Value = ROTULO_CORTE
    ws.Range("C3").Value = ORIENTACAO_CORTE
    try:
        val = ws.Range("B3").Validation
        val.InputTitle = "DATA DE CORTE DA APURACAO"
        val.InputMessage = (
            "Ultima data considerada nos calculos, PCs, competencias e "
            "resultados da apuracao."
        )
        val.ErrorMessage = "Informe uma data valida no formato dd/mm/aaaa."
    except Exception:
        pass
    _restaurar_protecao(ws, estado)


# --------------------------------------------------------------------------- #
# 3) cobertura_temporal — datas padronizadas e PCs posteriores ao corte
# --------------------------------------------------------------------------- #
def _aplicar_cobertura_temporal(ws) -> None:
    estado = _sem_protecao(ws)
    excel = ws.Application

    # Linha 17 nasce com o visual das linhas automaticas do proprio quadro.
    ws.Range("A12").Copy()
    ws.Range("A17").PasteSpecial(XL_PASTE_FORMATS)
    ws.Range("B12").Copy()
    ws.Range("B17").PasteSpecial(XL_PASTE_FORMATS)
    excel.CutCopyMode = False
    ws.Range("B17").NumberFormatLocal = "Geral"

    for linha, rotulo in ROTULOS_COBERTURA.items():
        ws.Range(f"A{linha}").Value = rotulo

    # Data de geracao/analise: automatica (finalidade administrativa; pode ser
    # posterior a data de corte sem contaminar a apuracao).
    ws.Range("B4").Formula = '=IF(CONTROLE!$B$14="","",CONTROLE!$B$14)'
    ws.Range("B4").NumberFormatLocal = FMT_DATA
    try:
        ws.Range("B4").Validation.Delete()
    except Exception:
        pass
    ws.Range("B4").Interior.Color = 0xFBF3EB   # azul claro dos campos AUTO

    _conferir_ascii_e_parenteses(ULTIMO_PC_ATE_CORTE, EXISTEM_PCS_POSTERIORES)
    ws.Range("B14").Formula = ULTIMO_PC_ATE_CORTE
    ws.Range("B14").NumberFormatLocal = FMT_DATA
    ws.Range("B17").Formula = EXISTEM_PCS_POSTERIORES
    _exigir_formulas(ws, ["B4", "B14", "B17"])
    ws.Range("C14").Value = (
        "Automático: maior DATA_PC menor ou igual à data de corte da apuração. "
        "PCs posteriores permanecem cadastrados e não alteram a cobertura oficial."
    )
    ws.Range("C17").Value = (
        "PCs posteriores ao corte não entram no retroativo nem no total "
        "considerado até o corte; a existência deles não é erro."
    )
    ws.Range("B26").Value = LEGENDA_FISCAL
    ws.Range("B27").Value = LEGENDA_GCC
    _restaurar_protecao(ws, estado)


# --------------------------------------------------------------------------- #
# 4) itens_RC — apenas titulos e orientacoes
# --------------------------------------------------------------------------- #
def _aplicar_itens_rc(ws) -> None:
    estado = _sem_protecao(ws)
    ciclos = {"B": "C0", "E": "C1", "H": "C2", "K": "C3", "N": "C4"}
    for col in ciclos:
        ws.Range(f"{col}1:{chr(ord(col) + 2)}1").UnMerge()
    # O ciclo passa a identificar cada coluna na linha 2; a linha 1 vira o
    # titulo unico do bloco. Limpar antes de mesclar evita o alerta do Excel.
    ws.Range("B1:P1").ClearContents()
    ws.Range("B1:P1").Merge()
    ws.Range("B1").Value = TITULO_RC_CICLOS
    ws.Range("B1").WrapText = True
    ws.Range("B1").HorizontalAlignment = XL_CENTER
    for col, ciclo in ciclos.items():
        c1, c2, c3 = col, chr(ord(col) + 1), chr(ord(col) + 2)
        ws.Range(f"{c1}2").Value = f"VU ATUALIZADO {ciclo}"
        ws.Range(f"{c2}2").Value = f"QTD RESTANTE {ciclo}"
        ws.Range(f"{c3}2").Value = f"TOTAL R$ {ciclo}"
    ws.Range("Q1").Value = TITULO_RC_ATUAL
    ws.Range("Q2").Value = "DATA DA POSIÇÃO FÍSICA (AUTO)"
    ws.Range("U2").Value = "QTD REMANESCENTE NA DATA DA POSIÇÃO (AUTO)"
    ws.Range("Z1").Value = (
        "APLICABILIDADE TEMPORAL NA ABERTURA "
        "(AUTO — ORIGEM: DATA DE EFEITO DA ALTERAÇÃO CONTRATUAL)"
    )
    ws.Range("Z2").Value = ROTULO_RC_DATA_ADITIVO
    ws.Range("AE1:AE2").Merge()
    ws.Range("AE1").Value = NOTA_RC_DATA_ADITIVO
    ws.Range("AE1").WrapText = True
    ws.Columns("AE").ColumnWidth = 46
    _restaurar_protecao(ws, estado)


# --------------------------------------------------------------------------- #
# 5) RESULTADOS + MEMORIA_RESULTADOS — apresentacao e selecao da fonte
# --------------------------------------------------------------------------- #
def _aplicar_resultados(wb) -> None:
    _conferir_ascii_e_parenteses(
        SITUACAO_VTA_POSICAO_ATUAL, SITUACAO_VTA_ABERTURA,
        SITUACAO_CICLO_ATUAL, CEE_DISPONIVEL,
    )
    mem = wb.Worksheets("MEMORIA_RESULTADOS")
    estado_mem = _sem_protecao(mem)
    mem.Range("W49").Formula = CEE_DISPONIVEL
    _restaurar_protecao(mem, estado_mem)

    ws = wb.Worksheets("RESULTADOS")
    estado = _sem_protecao(ws)
    ws.Range("H10").Formula = SITUACAO_VTA_POSICAO_ATUAL
    ws.Range("H11").Formula = SITUACAO_VTA_ABERTURA
    ws.Range("H33").Formula = SITUACAO_CICLO_ATUAL
    _exigir_formulas(ws, ["H10", "H11", "H33"])
    _exigir_formulas(mem, ["W49"])
    ws.Range("A10").Value = "VTA PELA POSIÇÃO FÍSICA ATUAL"
    ws.Range("A11").Value = "VTA PELA ABERTURA DO CICLO DE REFERÊNCIA"
    ws.Range("E5").Value = "Data de corte da apuração"
    ws.Range("A35").Value = "Data de referência utilizada"
    _restaurar_protecao(ws, estado)


# --------------------------------------------------------------------------- #
# 6) Arredondamento itemizado — historico_VU como fonte canonica do VU
# --------------------------------------------------------------------------- #
def _aplicar_arredondamento(wb) -> None:
    ir = wb.Worksheets("itens_Remanesc")
    estado_ir = _sem_protecao(ir)
    # VALOR_REM_INICIO_Cn = ARRED(QTD_REM_AJUSTADA_Cn x VU_ATUALIZADO_Cn; 2)
    remanescente = {
        "F": ("posicao_contratual!K2", "historico_VU!D2", 1),
        "H": ("posicao_contratual!O2", "historico_VU!E2", 2),
        "J": ("posicao_contratual!S2", "historico_VU!F2", 3),
        "L": ("posicao_contratual!W2", "historico_VU!G2", 4),
    }
    formulas: dict[str, str] = {}
    for col, (qtd, vu, indice) in remanescente.items():
        formulas[col] = (
            f'=IF(FALSE,ROUND(SUMIF($A1:A$2,"<>",${col}1:{col}$2),2),'
            f'IF(OR(A2="",AND(ISNUMBER(posicao_contratual!$AL2),'
            f'posicao_contratual!$AL2>{indice}),{qtd}="",'
            f'NOT(ISNUMBER({vu}))),"",ROUND({qtd}*{vu},2)))'
        )
    # VALOR_EXECUTADO_Cn = ARRED(QTD_EXECUTADA_Cn x VU_ATUALIZADO_Cn; 2)
    executado = {
        "N": ("M2", "historico_VU!D2"),
        "P": ("O2", "historico_VU!E2"),
        "R": ("Q2", "historico_VU!F2"),
        "AC": ("AB2", "historico_VU!C2"),
    }
    for col, (qtd, vu) in executado.items():
        formulas[col] = (
            f'=IF(FALSE,ROUND(SUMIF($A1:A$2,"<>",${col}1:{col}$2),2),'
            f'IF(OR({qtd}="",NOT(ISNUMBER({vu}))),"",ROUND({qtd}*{vu},2)))'
        )
    _conferir_ascii_e_parenteses(*formulas.values())
    for col, formula in formulas.items():
        ir.Range(f"{col}2:{col}{FIM}").Formula = formula
    _exigir_formulas(ir, [f"{col}2:{col}{FIM}" for col in formulas])
    _restaurar_protecao(ir, estado_ir)

    ad = wb.Worksheets("aditivos")
    estado_ad = _sem_protecao(ad)
    vu_aditivo = _vu_por_ciclo("$C2", "$A2")
    formula_j = (
        f'=IF(OR(L2="",F2=""),"",ROUND(L2*IF(AND(UPPER(H2)="SIM",'
        f'ISNUMBER({vu_aditivo})),{vu_aditivo},F2),2))'
    )
    _conferir_ascii_e_parenteses(formula_j)
    ad.Range(f"J2:J{FIM}").Formula = formula_j
    _exigir_formulas(ad, [f"J2:J{FIM}"])
    ad.Range("J1").Value = "Valor atualizado da alteracao (QTD x VU atualizado)"
    _restaurar_protecao(ad, estado_ad)

    # RESULTADOS!C26:C30 (remanescente atualizado por ciclo) permanecem como
    # ROUND(SUMPRODUCT(QTD, VU_ATUALIZADO), 2), inalteradas. Sao AGREGADOS, nao
    # valores itemizados: o VU ja e o canonico de historico_VU e a unica
    # diferenca possivel seria o ponto de arredondamento do somatorio. A versao
    # por item — SUMPRODUCT(ROUND(IFERROR(qtd*vu,0),2)) — foi testada no Excel
    # real e retornou 0,00 (IFERROR nao sobrevive como array dentro de
    # SUMPRODUCT nesta pasta), entao a formula homologada foi preservada.
    # O arredondamento POR ITEM continua garantido onde o item e calculado:
    # itens_RC, CICLO_EM_EXECUCAO, itens_Remanesc e MEMORIA_RESULTADOS (que e
    # quem alimenta o VTA oficial, W48/W50).


# --------------------------------------------------------------------------- #
# 7) Cores das abas de preenchimento
# --------------------------------------------------------------------------- #
def _aplicar_cores(wb) -> None:
    nomes = [ws.Name for ws in wb.Worksheets]
    for aba in ABAS_DE_ENTRADA:
        if aba in nomes:
            wb.Worksheets(aba).Tab.Color = COR_ABA_ENTRADA


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

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_ux_posicao_"))
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
        _aplicar_controle(wb.Worksheets("CONTROLE"))
        _aplicar_posicao_referencia(wb.Worksheets("posicao_referencia"))
        _aplicar_cobertura_temporal(wb.Worksheets("cobertura_temporal"))
        _aplicar_itens_rc(wb.Worksheets("itens_RC"))
        _aplicar_resultados(wb)
        _aplicar_arredondamento(wb)
        _aplicar_cores(wb)
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
