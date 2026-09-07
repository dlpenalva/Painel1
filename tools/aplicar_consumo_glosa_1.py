# -*- coding: utf-8 -*-
"""CONSUMO-GLOSA-1 — ajuste OPCIONAL da execucao por valor pago / glosa.

Aplica, via Excel COM em copia temporaria (padrao zero-corrupcao ja usado por
`aplicar_vta_consumido_canonico.py`), a camada economica opcional do metodo
Itens Consumidos.

REGRA PETREA: sem nenhum campo novo preenchido, todo o workbook precisa
produzir exatamente o mesmo resultado do checkpoint homologado. As formulas
reescritas (F20 e D10:D14) foram construidas para colapsar de volta, termo a
termo, na formula anterior quando nao ha ajuste — nunca em uma reconstrucao
equivalente "por outra via" que pudesse mudar arredondamento.

* itens_Consumidos (area lateral livre X:AG, linhas 1..6 = cabecalho + C0..C4;
  W fica vazia como separador):
    - X CICLO (literal), Y AJUSTE_VALOR_CALCULADO (auto),
      Z AJUSTE_TIPO (MANUAL, dropdown vazio/"Valor pago"/"Glosa"),
      AA AJUSTE_VALOR_INFORMADO (MANUAL),
      AB AJUSTE_VALOR_PAGO_CONSIDERADO, AC AJUSTE_GLOSA, AD AJUSTE_FATOR,
      AE AJUSTE_VALOR_PAGO_ATUALIZADO, AF AJUSTE_RETROATIVO,
      AG AJUSTE_STATUS (todas automaticas).
    - A base economica do ciclo e SUMPRODUCT(QTD_CONS_Cn, VU_ORIGINAL), a
      mesma ja usada por MEMORIA_RESULTADOS!D11:D14 — NUNCA VALOR_CONS_Cn,
      que ja embute o proprio reajuste em apuracao.
* MEMORIA_RESULTADOS:
    - F20 (execucao consumida atualizada): cada ciclo passa a contribuir com
      AE (valor pago atualizado) quando ha ajuste valido; sem ajuste, com a
      mesma SUM(coluna VALOR_CONS_Cn) de antes. Qualquer REVISAR fecha a
      medida inteira (fail-closed), nunca vira zero.
    - D10:D14 (retroativo do metodo Itens): idem — quando ha ajuste valido, a
      base SUMPRODUCT e trocada pelo valor pago considerado; a expressao do
      fator (F - F/D) fica intacta.
    - S69:T75 (bloco novo): medidas canonicas agregadas dos ajustes.
* RESULTADOS NAO e alterada: a aba nao tem uma unica linha visivel livre
  entre 1 e 87 (as vazias sao separadores geridos, linhas ocultas ou ancoras
  testadas) e abaixo da 87 esta a camada que o rollback da UX2 removeu. Ela
  ja reflete o ajuste sozinha, porque F20 alimenta o executado apurado, o
  retroativo e o VTA exibidos ali. Ver o comentario do bloco RESULTADOS.

NAO altera: quantidades consumidas, remanescente fisico (C33/D33/D35),
Financeiro, PCs, aditivos, ou qualquer ramo de B26 alem do que ja existia.
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
XL_VALIDATE_LIST = 3
XL_VALID_ALERT_STOP = 1
XL_BETWEEN = 1

ABA_MEMORIA = "MEMORIA_RESULTADOS"
ABA_CONSUMIDOS = "itens_Consumidos"
ABA_RESULTADOS = "RESULTADOS"

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

CICLOS = ("C0", "C1", "C2", "C3", "C4")
# linha do bloco lateral por ciclo (espelha R2:U6, ja existente na aba)
LINHA_CICLO = {nome: 2 + i for i, nome in enumerate(CICLOS)}
COL_QTD = {"C0": "E", "C1": "G", "C2": "I", "C3": "K", "C4": "M"}
COL_VALOR = {"C0": "F", "C1": "H", "C2": "J", "C3": "L", "C4": "N"}

CABECALHOS = {
    "X": "AJUSTE_CICLO",
    "Y": "AJUSTE_VALOR_CALCULADO",
    "Z": "AJUSTE_TIPO",
    "AA": "AJUSTE_VALOR_INFORMADO",
    "AB": "AJUSTE_VALOR_PAGO_CONSIDERADO",
    "AC": "AJUSTE_GLOSA",
    "AD": "AJUSTE_FATOR",
    "AE": "AJUSTE_VALOR_PAGO_ATUALIZADO",
    "AF": "AJUSTE_RETROATIVO",
    "AG": "AJUSTE_STATUS",
}

LEGENDA = (
    "AJUSTES DA EXECUCAO - VALOR PAGO / GLOSAS (OPCIONAL)",
    "Preencha somente quando o valor pago considerado na execucao do ciclo"
    " diferir do valor calculado pelos itens consumidos.",
    "AJUSTE_TIPO vazio = comportamento normal. Escolha 'Valor pago' para"
    " informar o valor bruto reconhecido apos a glosa, ou 'Glosa' para"
    " informar o total glosado no ciclo.",
    "Valor pago = valor bruto da execucao economicamente reconhecida, antes do"
    " reajuste retroativo em apuracao. Nao e valor liquido: IR, ISS, INSS e"
    " demais retencoes tributarias nao sao glosa.",
    "A glosa e financeira: nao altera quantidade consumida nem devolve saldo"
    " ao remanescente. Uma glosa que alcance dois ciclos deve ser segregada"
    " entre eles.",
)

STATUS_APLICADO = "AJUSTE APLICADO"

_AJUSTE_APLICADO_EM = {
    ciclo: f'{ABA_CONSUMIDOS}!$AG${LINHA_CICLO[ciclo]}="{STATUS_APLICADO}"'
    for ciclo in CICLOS
}
# Qualquer ciclo em REVISAR fecha as medidas derivadas (fail-closed): dado
# invalido nunca vira zero nem some silenciosamente.
_HA_REVISAR = f'COUNTIF({ABA_CONSUMIDOS}!$AG$2:$AG$6,"REVISAR*")>0'


def _formula_valor_calculado(ciclo: str) -> str:
    qtd = COL_QTD[ciclo]
    return (
        f'=IF(COUNT(${qtd}$2:${qtd}$200)=0,"",'
        f'ROUND(SUMPRODUCT(${qtd}$2:${qtd}$200,$C$2:$C$200),2))'
    )


def _formula_status(linha: int) -> str:
    return (
        f'=IF(AND($Z{linha}="",$AA{linha}=""),"",'
        f'IF($Z{linha}="","REVISAR: VALOR INFORMADO SEM TIPO DE AJUSTE",'
        f'IF(AND($Z{linha}<>"Valor pago",$Z{linha}<>"Glosa"),'
        f'"REVISAR: TIPO DE AJUSTE INVALIDO",'
        f'IF($AA{linha}="","REVISAR: TIPO DE AJUSTE SEM VALOR",'
        f'IF(NOT(ISNUMBER($AA{linha})),"REVISAR: VALOR INFORMADO NAO NUMERICO",'
        f'IF($AA{linha}<0,IF($Z{linha}="Glosa","REVISAR: GLOSA NEGATIVA",'
        f'"REVISAR: VALOR PAGO NEGATIVO"),'
        f'IF(NOT(ISNUMBER($Y{linha})),"REVISAR: SEM EXECUCAO CALCULADA NO CICLO",'
        f'IF(NOT(ISNUMBER($AD{linha})),"REVISAR: FATOR DO CICLO INDISPONIVEL",'
        f'IF(ROUND($AA{linha},2)>ROUND($Y{linha},2),'
        f'IF($Z{linha}="Glosa","REVISAR: GLOSA MAIOR QUE O CALCULADO",'
        f'"REVISAR: VALOR PAGO MAIOR QUE O CALCULADO"),'
        f'"{STATUS_APLICADO}")))))))))'
    )


def _formulas_linha(ciclo: str) -> dict[str, str]:
    linha = LINHA_CICLO[ciclo]
    gate = f'$AG{linha}<>"{STATUS_APLICADO}"'
    return {
        "Y": _formula_valor_calculado(ciclo),
        # Valor pago considerado: medida canonica unica. "Valor pago" e
        # "Glosa" sao apenas duas formas de informar a MESMA grandeza.
        "AB": (
            f'=IF({gate},"",IF($Z{linha}="Valor pago",ROUND($AA{linha},2),'
            f'ROUND($Y{linha}-$AA{linha},2)))'
        ),
        "AC": f'=IF({gate},"",ROUND($Y{linha}-$AB{linha},2))',
        "AD": f'=IF(ISNUMBER($U${linha}),$U${linha},"")',
        "AE": f'=IF({gate},"",ROUND($AB{linha}*$AD{linha},2))',
        "AF": f'=IF({gate},"",ROUND($AE{linha}-$AB{linha},2))',
        "AG": _formula_status(linha),
    }


# --------------------------------------------------------------------------
# MEMORIA_RESULTADOS!F20 — execucao consumida atualizada
# --------------------------------------------------------------------------
_F20_GATES = (
    'IF(OR('
    f'COUNT({ABA_CONSUMIDOS}!$E$2:$E$200)<>COUNT({ABA_CONSUMIDOS}!$F$2:$F$200),'
    f'COUNT({ABA_CONSUMIDOS}!$G$2:$G$200)<>COUNT({ABA_CONSUMIDOS}!$H$2:$H$200),'
    f'COUNT({ABA_CONSUMIDOS}!$I$2:$I$200)<>COUNT({ABA_CONSUMIDOS}!$J$2:$J$200),'
    f'COUNT({ABA_CONSUMIDOS}!$K$2:$K$200)<>COUNT({ABA_CONSUMIDOS}!$L$2:$L$200),'
    f'COUNT({ABA_CONSUMIDOS}!$M$2:$M$200)<>COUNT({ABA_CONSUMIDOS}!$N$2:$N$200)'
    '),"",'
    f'IF(COUNT({ABA_CONSUMIDOS}!$E$2:$E$200,{ABA_CONSUMIDOS}!$G$2:$G$200,'
    f'{ABA_CONSUMIDOS}!$I$2:$I$200,{ABA_CONSUMIDOS}!$K$2:$K$200,'
    f'{ABA_CONSUMIDOS}!$M$2:$M$200)=0,"",'
)

_F20_ANTES = (
    "=" + _F20_GATES
    + f'ROUND(SUM({ABA_CONSUMIDOS}!$F$2:$F$200,{ABA_CONSUMIDOS}!$H$2:$H$200,'
    f'{ABA_CONSUMIDOS}!$J$2:$J$200,{ABA_CONSUMIDOS}!$L$2:$L$200,'
    f'{ABA_CONSUMIDOS}!$N$2:$N$200),2)))'
)

_F20_PARCELAS = "+".join(
    f'IF({_AJUSTE_APLICADO_EM[ciclo]},'
    f'{ABA_CONSUMIDOS}!$AE${LINHA_CICLO[ciclo]},'
    f'SUM({ABA_CONSUMIDOS}!${COL_VALOR[ciclo]}$2:${COL_VALOR[ciclo]}$200))'
    for ciclo in CICLOS
)

_F20_NOVA = (
    "=" + _F20_GATES
    + f'IF({_HA_REVISAR},"",'
    + f'ROUND({_F20_PARCELAS},2))))'
)


# --------------------------------------------------------------------------
# MEMORIA_RESULTADOS!D10:D14 — retroativo do metodo Itens
# --------------------------------------------------------------------------
def _d10_antes() -> str:
    return f'=IF(COUNTIFS({ABA_CONSUMIDOS}!$E$2:$E$200,">0")=0,"",0)'


def _dn_antes(ciclo: str) -> str:
    """Formula homologada de D11..D14 (indice do ciclo 1..4)."""
    n = int(ciclo[1])
    qtd = COL_QTD[ciclo]
    fator = f"parametros!$F{n + 2}"       # C1 -> F3 ... C4 -> F6
    apuracao = f"parametros!$D{11 + n}"   # C1 -> D12 ... C4 -> D15
    return (
        f'=IF(OR(COUNTIFS({ABA_CONSUMIDOS}!${qtd}$2:${qtd}$200,">0")=0,'
        f'NOT(ISNUMBER({fator})),NOT(ISNUMBER({apuracao})),{apuracao}=0),"",'
        f'ROUND(SUMPRODUCT({ABA_CONSUMIDOS}!${qtd}$2:${qtd}$200,'
        f'{ABA_CONSUMIDOS}!$C$2:$C$200)*({fator}-{fator}/{apuracao}),2))'
    )


def _d10_nova() -> str:
    return f'=IF({_HA_REVISAR},"",{_d10_antes()[1:]})'


def _dn_nova(ciclo: str) -> str:
    n = int(ciclo[1])
    linha = LINHA_CICLO[ciclo]
    fator = f"parametros!$F{n + 2}"
    apuracao = f"parametros!$D{11 + n}"
    ramo_ajustado = (
        f'IF(OR(NOT(ISNUMBER({fator})),NOT(ISNUMBER({apuracao})),{apuracao}=0),"",'
        f'ROUND({ABA_CONSUMIDOS}!$AB${linha}*({fator}-{fator}/{apuracao}),2))'
    )
    return (
        f'=IF({_HA_REVISAR},"",'
        f'IF({_AJUSTE_APLICADO_EM[ciclo]},{ramo_ajustado},{_dn_antes(ciclo)[1:]}))'
    )


# --------------------------------------------------------------------------
# MEMORIA_RESULTADOS!S69:T75 — medidas canonicas agregadas
# --------------------------------------------------------------------------
_T75 = (
    f'=IF({_HA_REVISAR},"REVISAR",'
    f'IF(COUNTIF({ABA_CONSUMIDOS}!$AG$2:$AG$6,"{STATUS_APLICADO}")=0,"SEM AJUSTE",'
    f'"{STATUS_APLICADO}"))'
)


def _agregado(expressao: str) -> str:
    return f'=IF($T$75<>"{STATUS_APLICADO}","",{expressao})'


_MEMORIA_AGREGADOS = {
    "S69": "AJUSTES DA EXECUCAO - VALOR PAGO / GLOSA (metodo Itens)",
    "S70": "Valor calculado da execucao (ciclos ajustados)",
    "T70": _agregado(
        f'ROUND(SUM({ABA_CONSUMIDOS}!$AB$2:$AB$6)'
        f'+SUM({ABA_CONSUMIDOS}!$AC$2:$AC$6),2)'
    ),
    "S71": "Glosa total",
    "T71": _agregado(f'ROUND(SUM({ABA_CONSUMIDOS}!$AC$2:$AC$6),2)'),
    "S72": "Valor pago considerado",
    "T72": _agregado(f'ROUND(SUM({ABA_CONSUMIDOS}!$AB$2:$AB$6),2)'),
    "S73": "Valor pago atualizado",
    "T73": _agregado(f'ROUND(SUM({ABA_CONSUMIDOS}!$AE$2:$AE$6),2)'),
    "S74": "Retroativo do valor pago",
    "T74": _agregado(f'ROUND(SUM({ABA_CONSUMIDOS}!$AF$2:$AF$6),2)'),
    "S75": "Status dos ajustes",
    "T75": _T75,
}


# --------------------------------------------------------------------------
# RESULTADOS — deliberadamente NAO recebe faixa nova.
#
# A aba nao tem uma unica linha visivel livre entre 1 e 87. Todas as seis
# linhas vazias sao ancoras de leiaute defendidas por testes de frentes
# anteriores:
#   - 8/14/23/32/39/52 sao separadores brancos geridos pela Etapa 50.3 (o
#     conteudo canonico que mora neles usa o formato ";;;", invisivel);
#   - 31/40/51 sao linhas OCULTAS de 7pt (mesma etapa);
#   - 78 fecha a tabela 8 e e verificada como vazia por
#     tests/test_resultados_final_1.py::test_tabela_8_padronizada;
#   - abaixo da 87 esta a camada que o rollback da UX2 removeu, com
#     `max_row == 87` e 88:200 vazio travados em dois testes.
# Escrever em qualquer uma delas ou seria invisivel para o usuario, ou
# reabriria a porta que a UX2 fechou.
#
# A aba, porem, JA reflete o ajuste sozinha: com glosa valida, F20 muda e com
# ele o executado apurado (B36/B83), o retroativo (D22) e o VTA (C5/B65/B86).
# As seis medidas do ajuste ficam publicadas em MEMORIA_RESULTADOS!S69:T75 (a
# aba de memoria auditavel, onde as medidas canonicas do projeto vivem) e no
# card discreto da web.
# --------------------------------------------------------------------------
_RESULTADOS_BLOCO: dict[str, str] = {}


# --------------------------------------------------------------------------
# Infraestrutura COM
# --------------------------------------------------------------------------
def _checar_parenteses(rotulo: str, formula: str) -> None:
    saldo = 0
    dentro_texto = False
    for ch in formula:
        if ch == '"':
            dentro_texto = not dentro_texto
            continue
        if dentro_texto:
            continue
        if ch == "(":
            saldo += 1
        elif ch == ")":
            saldo -= 1
        if saldo < 0:
            raise ValueError(f"{rotulo}: parenteses desbalanceados (fecha demais).")
    if saldo != 0:
        raise ValueError(f"{rotulo}: parenteses desbalanceados (saldo {saldo}).")
    if dentro_texto:
        raise ValueError(f"{rotulo}: aspas desbalanceadas.")
    try:
        formula.encode("ascii")
    except UnicodeEncodeError as exc:
        raise ValueError(f"{rotulo}: formula com caractere nao-ASCII.") from exc


def _validar_formulas_estaticas() -> None:
    _checar_parenteses("F20", _F20_NOVA)
    _checar_parenteses("F20_ANTES", _F20_ANTES)
    _checar_parenteses("D10", _d10_nova())
    for ciclo in ("C1", "C2", "C3", "C4"):
        _checar_parenteses(f"D{10 + int(ciclo[1])}", _dn_nova(ciclo))
    for ciclo in CICLOS:
        for coluna, formula in _formulas_linha(ciclo).items():
            _checar_parenteses(f"{coluna}{LINHA_CICLO[ciclo]}", formula)
    for endereco, valor in _MEMORIA_AGREGADOS.items():
        if str(valor).startswith("="):
            _checar_parenteses(endereco, valor)
    for endereco, valor in _RESULTADOS_BLOCO.items():
        _checar_parenteses(endereco, valor)
    for texto in LEGENDA:
        texto.encode("ascii")


def _nomes_abas(wb) -> list[str]:
    return [ws.Name for ws in wb.Worksheets]


def _nomes_definidos(wb) -> set[str]:
    return {str(nome.Name).split("!")[-1] for nome in wb.Names}


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


def _validar_origem(wb) -> None:
    abas = _nomes_abas(wb)
    obrigatorias = {
        ABA_MEMORIA, ABA_CONSUMIDOS, ABA_RESULTADOS, "parametros", "CONTROLE",
    }
    ausentes = sorted(obrigatorias.difference(abas))
    if ausentes:
        raise ValueError(f"Abas obrigatorias ausentes: {', '.join(ausentes)}")

    mem = wb.Worksheets(ABA_MEMORIA)
    atual_f20 = str(mem.Range("F20").Formula)
    if atual_f20 != _F20_ANTES:
        raise ValueError(
            "MEMORIA_RESULTADOS!F20 nao corresponde ao checkpoint homologado.\n"
            f"Atual:    {atual_f20}\nEsperada: {_F20_ANTES}"
        )
    esperadas_d = {"D10": _d10_antes()}
    for ciclo in ("C1", "C2", "C3", "C4"):
        esperadas_d[f"D{10 + int(ciclo[1])}"] = _dn_antes(ciclo)
    for endereco, esperada in esperadas_d.items():
        atual = str(mem.Range(endereco).Formula)
        if atual != esperada:
            raise ValueError(
                f"MEMORIA_RESULTADOS!{endereco} nao corresponde ao checkpoint.\n"
                f"Atual:    {atual}\nEsperada: {esperada}"
            )
    for endereco in _MEMORIA_AGREGADOS:
        if mem.Range(endereco).Value not in (None, ""):
            raise ValueError(f"MEMORIA_RESULTADOS!{endereco} ja ocupada.")

    consumidos = wb.Worksheets(ABA_CONSUMIDOS)
    for coluna in CABECALHOS:
        if consumidos.Range(f"{coluna}1").Value not in (None, ""):
            raise ValueError(f"itens_Consumidos!{coluna}1 ja ocupada.")
    if consumidos.Range("W1").Value not in (None, ""):
        raise ValueError("itens_Consumidos!W1 deveria ficar vazia (separador).")
    for ciclo in CICLOS:
        linha = LINHA_CICLO[ciclo]
        rotulo = str(consumidos.Range(f"R{linha}").Value or "").strip().upper()
        if rotulo != ciclo:
            raise ValueError(
                f"itens_Consumidos!R{linha} deveria conter {ciclo}; achei "
                f"{rotulo!r}. A tabela lateral de ciclos mudou de lugar."
            )

    resultados = wb.Worksheets(ABA_RESULTADOS)
    for endereco in _RESULTADOS_BLOCO:
        linha = int("".join(ch for ch in endereco if ch.isdigit()))
        if resultados.Range(endereco).Value not in (None, ""):
            raise ValueError(f"RESULTADOS!{endereco} ja ocupada.")
        # Toda linha "vazia" de RESULTADOS entre 1 e 87 e ancora de leiaute:
        # 8/14/23/32/39/52 sao separadores brancos geridos pela Etapa 50.3,
        # 31/40/51 sao linhas ocultas de 7pt e 78 fecha a tabela 8. Escrever
        # em qualquer uma delas ficaria invisivel ou quebraria uma trava — por
        # isso _RESULTADOS_BLOCO esta vazio; estes gates ficam de sentinela
        # caso alguem volte a povoa-lo.
        if bool(resultados.Rows(linha).Hidden):
            raise ValueError(
                f"RESULTADOS linha {linha} esta oculta; ficaria invisivel."
            )
        if linha in {8, 14, 23, 32, 39, 52, 78}:
            raise ValueError(
                f"RESULTADOS linha {linha} e ancora de leiaute (Etapa 50.3 / "
                "tabela 8); nao pode receber conteudo novo."
            )

    if "VTA_FINAL" not in _nomes_definidos(wb):
        raise ValueError("Nome definido VTA_FINAL ausente.")


def _snapshot_travas(wb) -> dict[str, str]:
    mem = wb.Worksheets(ABA_MEMORIA)
    enderecos = (
        # VTA/PC e remanescente fisico: nada aqui pode mudar.
        "B15", "B16", "B20", "B21", "B22", "B23", "B26", "B28",
        "C33", "D33", "D35", "T21", "T22", "T23", "T25", "T39", "T40", "T41",
    )
    travas = {f"{ABA_MEMORIA}!{e}": str(mem.Range(e).Formula) for e in enderecos}
    consumidos = wb.Worksheets(ABA_CONSUMIDOS)
    for endereco in ("D2", "F2", "H2", "J2", "L2", "N2", "O2", "P2", "Q2", "V2",
                     "D200", "N200", "O200", "P200", "V200"):
        travas[f"{ABA_CONSUMIDOS}!{endereco}"] = str(
            consumidos.Range(endereco).Formula
        )
    return travas


def _aplicar_itens_consumidos(wb) -> None:
    ws = wb.Worksheets(ABA_CONSUMIDOS)
    estado, selecao = _capturar_protecao(ws)
    try:
        for coluna, titulo in CABECALHOS.items():
            ws.Range(f"{coluna}1").Value = titulo
        for ciclo in CICLOS:
            linha = LINHA_CICLO[ciclo]
            ws.Range(f"X{linha}").Value = ciclo
            for coluna, formula in _formulas_linha(ciclo).items():
                ws.Range(f"{coluna}{linha}").Formula = formula

        alvo = ws.Range("Z2:Z6")
        try:
            alvo.Validation.Delete()
        except Exception:
            pass
        alvo.Validation.Add(
            Type=XL_VALIDATE_LIST,
            AlertStyle=XL_VALID_ALERT_STOP,
            Operator=XL_BETWEEN,
            Formula1="Valor pago,Glosa",
        )
        alvo.Validation.IgnoreBlank = True
        alvo.Validation.InCellDropdown = True

        for offset, texto in enumerate(LEGENDA):
            ws.Range(f"X{8 + offset}").Value = texto

        try:
            ws.Range("X1:AG1").Font.Bold = True
            ws.Range("X8").Font.Bold = True
            for coluna in ("Y", "AA", "AB", "AC", "AE", "AF"):
                ws.Range(f"{coluna}2:{coluna}6").NumberFormat = "#,##0.00"
            ws.Range("AD2:AD6").NumberFormat = "0.000000"
            for coluna, largura in (
                ("X", 10), ("Y", 22), ("Z", 16), ("AA", 20), ("AB", 26),
                ("AC", 16), ("AD", 14), ("AE", 26), ("AF", 20), ("AG", 42),
            ):
                ws.Columns(coluna).ColumnWidth = largura
        except Exception:
            pass
    finally:
        _restaurar_protecao(ws, estado, selecao)


def _aplicar_memoria(wb) -> None:
    mem = wb.Worksheets(ABA_MEMORIA)
    estado, selecao = _capturar_protecao(mem)
    try:
        mem.Range("F20").Formula = _F20_NOVA
        mem.Range("D10").Formula = _d10_nova()
        for ciclo in ("C1", "C2", "C3", "C4"):
            mem.Range(f"D{10 + int(ciclo[1])}").Formula = _dn_nova(ciclo)
        for endereco, valor in _MEMORIA_AGREGADOS.items():
            if str(valor).startswith("="):
                mem.Range(endereco).Formula = valor
            else:
                mem.Range(endereco).Value = valor
        try:
            mem.Range("S69").Font.Bold = True
            mem.Range("T70:T74").NumberFormat = "#,##0.00"
        except Exception:
            pass
    finally:
        _restaurar_protecao(mem, estado, selecao)


def _aplicar_resultados(wb) -> None:
    """No-op enquanto _RESULTADOS_BLOCO estiver vazio (ver o comentario la).

    Mantida para que a aba nao seja nem aberta para escrita sem necessidade:
    sem celulas a escrever, a protecao da planilha nem chega a ser removida.
    """
    if not _RESULTADOS_BLOCO:
        return
    ws = wb.Worksheets(ABA_RESULTADOS)
    estado, selecao = _capturar_protecao(ws)
    try:
        for endereco, formula in _RESULTADOS_BLOCO.items():
            ws.Range(endereco).Formula = formula
    finally:
        _restaurar_protecao(ws, estado, selecao)


def aplicar(origem: Path, destino: Path) -> None:
    _validar_formulas_estaticas()

    origem = Path(origem).resolve()
    destino = Path(destino).resolve()
    if not origem.is_file():
        raise FileNotFoundError(origem)
    if origem == destino:
        raise ValueError("Origem e destino devem ser diferentes.")

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_consumo_glosa_"))
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
        travas_antes = _snapshot_travas(wb)

        _aplicar_itens_consumidos(wb)
        _aplicar_memoria(wb)
        _aplicar_resultados(wb)

        # Trava do rodape de RESULTADOS: a aba nao pode ganhar camada nova
        # abaixo da linha 87 (porta que o rollback da UX2 fechou).
        usada = wb.Worksheets(ABA_RESULTADOS).UsedRange
        ultima_linha = int(usada.Row) + int(usada.Rows.Count) - 1
        if ultima_linha > 87:
            raise RuntimeError(
                "TRAVA VIOLADA: RESULTADOS passou a usar a linha "
                f"{ultima_linha}; o rodape homologado termina na 87."
            )

        travas_depois = _snapshot_travas(wb)
        if travas_antes != travas_depois:
            difs = {
                k: (travas_antes[k], travas_depois[k])
                for k in travas_antes if travas_antes[k] != travas_depois[k]
            }
            raise RuntimeError(f"TRAVA VIOLADA: {difs}")

        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()

        if aba_ativa in _nomes_abas(wb):
            wb.Worksheets(aba_ativa).Activate()
        wb.Save()
        salvo = True
        wb.Close(SaveChanges=False)
        wb = None

        # Reabre sem reparo para provar zero-corrupcao.
        wb = excel.Workbooks.Open(
            str(tmp_xlsx), UpdateLinks=0, ReadOnly=True, CorruptLoad=0
        )
        abas = _nomes_abas(wb)
        for obrig in (ABA_MEMORIA, ABA_CONSUMIDOS, ABA_RESULTADOS):
            if obrig not in abas:
                raise RuntimeError(f"Aba {obrig} ausente apos reabertura.")
        if "VTA_FINAL" not in _nomes_definidos(wb):
            raise RuntimeError("Nome VTA_FINAL ausente apos reabertura.")
        consumidos = wb.Worksheets(ABA_CONSUMIDOS)
        for coluna, titulo in CABECALHOS.items():
            if str(consumidos.Range(f"{coluna}1").Value or "") != titulo:
                raise RuntimeError(f"Cabecalho {coluna}1 perdido apos reabertura.")
        wb.Close(SaveChanges=False)
        wb = None
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
    print("CONSUMO-GLOSA-1 aplicado:", args.destino)


if __name__ == "__main__":
    main()
