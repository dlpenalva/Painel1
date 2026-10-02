# -*- coding: utf-8 -*-
"""RESULTADOS-SEM-EFEITO-150: execucao anterior ao inicio dos efeitos financeiros.

Quadro exclusivamente EXPLICATIVO na aba RESULTADOS, para os tres metodos
(Financeiro, PCs, Itens). Responde: "quanto da execucao ocorreu antes do inicio
dos efeitos financeiros e qual seria a diferenca caso essa execucao tivesse
sido alcancada pelo reajuste?". Nao e retroativo, nao e potencial, nao e valor
a pagar e nao entra em nenhuma soma, no VTA, em documento, card, Garantia ou
DOU.

COMO O BLOCO ANTIGO DO PC FUNCIONAVA (e por que foi preservado)
  MEMORIA_RESULTADOS!T28:T30 contam/somam, por ciclo C1..C4 computado
  (`parametros!A = "Sim"`), os PCs com `EFEITO_FINANCEIRO_PC = "Nao"` e
  VALOR_PC > 0. C0 e ciclo nao computado ja ficam FORA por construcao. Dentro
  desse universo, `L = "Nao"` equivale a `DATA_PC < INICIO_EFEITO_FINANCEIRO`
  (itens_PC!L). Esses helpers e a linha oculta A23:E23 permanecem intactos;
  este aplicador NAO cria um segundo calculo concorrente: ele quebra o MESMO
  predicado por ciclo e explicita a condicao de data (DATA_PC < H).

PERIODO SEM EFEITO = SO DATAS (revisao do PR #172): existe periodo quando
  INICIO_EFEITO > INICIO DO CICLO (no Financeiro, em nivel de MES). As linhas de
  `financeiro` NAO decidem se o periodo existe — apenas medem a execucao nele:
  periodo sem nenhuma linha/pagamento => estado B (R$ 0,00 conhecido), nunca D.

REGRAS POR METODO
  PCs        itens_PC: CICLO_PC = Cn, VALOR_PC > 0, EFEITO_FINANCEIRO_PC = Nao,
             DATA_PC < parametros!H(n). Valor original = soma de VALOR_PC;
             valor com reajuste = soma de VALOR_ATUALIZADO (DATA EXATA, fator
             do proprio ciclo). Sem fator numerico => "nao mensuravel".
  Financeiro financeiro: CICLO = cn, EFEITO_FINANCEIRO (G) = Nao, competencia
             anterior ao mes de parametros!H(n), VALOR_PAGO > 0. Valor com
             reajuste = soma de ROUND(VALOR_PAGO * FATOR_APLICAVEL(D), 2) por
             competencia (mesmo arredondamento de financeiro!E).
  Itens      itens_Consumidos so guarda quantidade POR CICLO, sem data de
             consumo. Nao ha como separar o que foi consumido antes do inicio
             dos efeitos: nada e rateado, estimado ou presumido. Ha consumo no
             ciclo e ha periodo sem efeito => "nao mensuravel com seguranca".
             Consumo informado igual a zero => sem execucao (R$ 0,00 conhecido).

ESTADOS POR CICLO (MEMORIA_RESULTADOS linha 91)
  A valores mensuraveis | B periodo sem efeito e SEM execucao (0,00 conhecido)
  C nao mensuravel com seguranca | D sem periodo anterior aos efeitos

ONDE FICA (nenhuma linha e inserida ou movida)
  RESULTADOS!E15:H21, ao lado da tabela 2 e alinhado as linhas C1..C4 (17..20);
  ate aqui era o texto pequeno da linha "PCs sem efeito financeiro".
  Helpers: MEMORIA_RESULTADOS!S78:W120 (area livre da aba oculta).

ALLOWLIST (alteracao cirurgica): RESULTADOS!E15:H21 (+ altura das linhas
16-21) e MEMORIA_RESULTADOS!S78:W120. Nada mais e tocado.

REGRA ZERO CORRUPCAO XLSX: aplicacao por Excel COM (openpyxl destroi a
formatacao condicional x14 da aba); formulas ASCII, em ingles, com parenteses
balanceados; textos acentuados ficam em CELULAS CONSTANTES de
MEMORIA_RESULTADOS, referenciadas pelas formulas.

Uso:  python tools/aplicar_resultados_sem_efeito_150.py [caminho.xlsx]
      (exige pywin32 + Excel real; sem argumento, altera o template oficial)
"""
from __future__ import annotations

import gc
import sys
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

ABA_RESULTADOS = "RESULTADOS"
ABA_MEMORIA = "MEMORIA_RESULTADOS"
MEM = ABA_MEMORIA

# Capacidades do template (iguais as usadas pelos helpers T28:T30 e por
# financeiro!B:G).
CAP_PC = 5001
ULT_FIN = 73
ULT_CONS = 200

# Ciclo -> (coluna na MEMORIA, linha em parametros, coluna QTD_CONS em
# itens_Consumidos, linha correspondente na tabela 2 de RESULTADOS).
CICLOS = {
    1: ("T", 3, "G", 17),
    2: ("U", 4, "I", 18),
    3: ("V", 5, "K", 19),
    4: ("W", 6, "M", 20),
}

# Linhas dos helpers em MEMORIA_RESULTADOS.
L_TITULO, L_CICLO, L_COMPUTADO, L_INI_CICLO, L_INI_EFEITO = 78, 79, 80, 81, 82
L_LIMITE, L_HA_PERIODO, L_QTD, L_QTD_FATOR = 83, 84, 85, 86
L_ORIGINAL, L_TERIAM, L_DIFERENCA, L_CONSUMO, L_ESTADO = 87, 88, 89, 90, 91
L_PRIMEIRA, L_ULTIMA, L_PERIODO_TXT, L_EFEITO_TXT = 92, 93, 94, 95
L_LINHA1, L_LINHA2 = 96, 97
L_FLAG_VISIVEL, L_FLAG_VALORES, L_FLAG_C, L_FLAG_D = 99, 100, 101, 102
L_TEXTOS = 104

# Textos constantes (acentuados) — celulas T105:T120.
TEXTOS = [
    ("T_TITULO", "EXECUÇÃO SEM EFEITO FINANCEIRO"),
    ("T_HDR_ORIGINAL", "Valor original"),
    ("T_HDR_PAGO", "Valor pago"),
    ("T_HDR_TERIAM", "Valor que teriam com o reajuste"),
    ("T_HDR_DIFERENCA", "Diferença sem efeito financeiro"),
    ("T_NOTA", "A diferença é apenas informativa. Não constitui retroativo "
               "reconhecido ou potencial e não representa valor a pagar."),
    ("T_CASO_C", "Execução sem efeito financeiro identificada, mas o impacto "
                 "monetário não pode ser mensurado com segurança com os dados "
                 "disponíveis."),
    ("T_CASO_D", "Não foi identificada execução anterior ao início dos efeitos "
                 "financeiros."),
    ("T_EFEITOS_A_PARTIR", " — efeitos a partir de "),
    ("T_EFEITOS_DESDE", " — efeitos desde o início do ciclo"),
    ("T_PCS", " PC(s) · "),
    ("T_COMPETENCIAS", " competência(s) · "),
    ("T_SEM_EXECUCAO", "Sem execução · "),
    ("T_PERIODO_SEM_EFEITO", "Sem efeito: "),
    ("T_NAO_MENSURAVEL", "Não mensurável"),
]
LINHA_TEXTO = {chave: L_TEXTOS + 1 + i for i, (chave, _) in enumerate(TEXTOS)}


def ref_texto(chave: str) -> str:
    """Referencia absoluta (na propria MEMORIA) de um texto constante."""
    return f"$T${LINHA_TEXTO[chave]}"


def ref_texto_ext(chave: str) -> str:
    """Referencia a um texto constante a partir de outra aba."""
    return f"{MEM}!{ref_texto(chave)}"


# ------------------------------------------------------------------ formulas
def _criterios_pc(n: int, c: str) -> str:
    return (
        f'itens_PC!$C$2:$C${CAP_PC},"C{n}",'
        f'itens_PC!$D$2:$D${CAP_PC},">0",'
        f'itens_PC!$B$2:$B${CAP_PC},"<"&{c}{L_INI_EFEITO},'
        f'itens_PC!$L$2:$L${CAP_PC},"Nao"'
    )


def _criterios_fin(n: int, c: str, com_valor: bool) -> str:
    base = (
        f'financeiro!$B$2:$B${ULT_FIN},"c{n}",'
        f'financeiro!$A$2:$A${ULT_FIN},"<"&{c}{L_LIMITE},'
        f'financeiro!$G$2:$G${ULT_FIN},"Nao"'
    )
    return base + (f',financeiro!$C$2:$C${ULT_FIN},">0"' if com_valor else "")


def _data_texto(ref: str) -> str:
    return (
        f'RIGHT("0"&DAY({ref}),2)&"/"&RIGHT("0"&MONTH({ref}),2)&"/"&YEAR({ref})'
    )


def _mes_texto(ref: str) -> str:
    return f'RIGHT("0"&MONTH({ref}),2)&"/"&YEAR({ref})'


def formulas_ciclo(n: int) -> dict[int, str]:
    """Formulas dos helpers de um ciclo (linha -> formula), em ASCII."""
    c, p, col_cons, _ = CICLOS[n]
    ini, efe, lim = f"{c}{L_INI_CICLO}", f"{c}{L_INI_EFEITO}", f"{c}{L_LIMITE}"
    qtd, fator = f"{c}{L_QTD}", f"{c}{L_QTD_FATOR}"
    orig, teriam = f"{c}{L_ORIGINAL}", f"{c}{L_TERIAM}"
    dif, cons, est = f"{c}{L_DIFERENCA}", f"{c}{L_CONSUMO}", f"{c}{L_ESTADO}"
    prim, ult = f"{c}{L_PRIMEIRA}", f"{c}{L_ULTIMA}"
    per_txt, efe_txt = f"{c}{L_PERIODO_TXT}", f"{c}{L_EFEITO_TXT}"
    ha = f"{c}{L_HA_PERIODO}"
    crit_pc = _criterios_pc(n, c)
    crit_fin_v = _criterios_fin(n, c, True)
    crit_fin = _criterios_fin(n, c, False)
    consumo_rng = f"itens_Consumidos!${col_cons}$2:${col_cons}${ULT_CONS}"
    metodo = "$B$4"

    f: dict[int, str] = {}
    f[L_COMPUTADO] = f'=IF(parametros!$A${p}="Sim",1,0)'
    f[L_INI_CICLO] = f'=IF(ISNUMBER(parametros!$C${p}),parametros!$C${p},"")'
    f[L_INI_EFEITO] = f'=IF(ISNUMBER(parametros!$H${p}),parametros!$H${p},"")'
    f[L_LIMITE] = f'=IF(ISNUMBER({efe}),EOMONTH({efe},-1)+1,"")'
    f[L_HA_PERIODO] = (
        f'=IF(OR({c}{L_COMPUTADO}=0,NOT(ISNUMBER({ini})),NOT(ISNUMBER({efe})),'
        f'AND({metodo}<>"Financeiro",{metodo}<>"PCs",{metodo}<>"Itens")),"",'
        f'IF({metodo}="Financeiro",IF(EOMONTH({efe},-1)>EOMONTH({ini},-1),1,0),'
        f'IF({efe}>{ini},1,0)))'
    )
    f[L_QTD] = (
        f'=IF({ha}<>1,"",IF({metodo}="PCs",COUNTIFS({crit_pc}),'
        f'IF({metodo}="Financeiro",COUNTIFS({crit_fin_v}),"")))'
    )
    f[L_QTD_FATOR] = (
        f'=IF({ha}<>1,"",IF({metodo}="PCs",'
        f'COUNTIFS({crit_pc},itens_PC!$F$2:$F${CAP_PC},">0"),'
        f'IF({metodo}="Financeiro",IFERROR(INDEX(financeiro!$D$2:$D${ULT_FIN},'
        f'MATCH("c{n}",financeiro!$B$2:$B${ULT_FIN},0)),""),"")))'
    )
    f[L_ORIGINAL] = (
        f'=IF({ha}<>1,"",IF({metodo}="PCs",'
        f'ROUND(SUMIFS(itens_PC!$D$2:$D${CAP_PC},{crit_pc}),2),'
        f'IF({metodo}="Financeiro",'
        f'ROUND(SUMIFS(financeiro!$C$2:$C${ULT_FIN},{crit_fin_v}),2),'
        f'IF(AND({metodo}="Itens",ISNUMBER({cons})),IF({cons}=0,0,""),""))))'
    )
    f[L_TERIAM] = (
        f'=IF({ha}<>1,"",IF({metodo}="PCs",'
        f'IF({qtd}=0,0,IF({fator}<{qtd},"",'
        f'ROUND(SUMIFS(itens_PC!$F$2:$F${CAP_PC},{crit_pc}),2))),'
        f'IF({metodo}="Financeiro",IF({qtd}=0,0,IF(NOT(ISNUMBER({fator})),"",'
        f'IFERROR(ROUND(SUMPRODUCT((financeiro!$B$2:$B${ULT_FIN}="c{n}")*'
        f'(financeiro!$A$2:$A${ULT_FIN}<{lim})*'
        f'(financeiro!$G$2:$G${ULT_FIN}="Nao")*'
        f'(financeiro!$C$2:$C${ULT_FIN}>0)*'
        f'ROUND(financeiro!$C$2:$C${ULT_FIN}*{fator},2)),2),""))),'
        f'IF(AND({metodo}="Itens",ISNUMBER({cons})),IF({cons}=0,0,""),""))))'
    )
    f[L_DIFERENCA] = f'=IF(OR({orig}="",{teriam}=""),"",ROUND({teriam}-{orig},2))'
    f[L_CONSUMO] = (
        f'=IF({metodo}="Itens",IF(COUNT({consumo_rng})=0,"",'
        f'SUM({consumo_rng})),"")'
    )
    f[L_ESTADO] = (
        f'=IF({ha}="","",IF({ha}=0,"D",IF({metodo}="Itens",'
        f'IF({cons}="","",IF({cons}=0,"B","C")),'
        f'IF({qtd}=0,"B",IF({dif}="","C","A")))))'
    )
    f[L_PRIMEIRA] = (
        f'=IF(AND({ha}=1,{metodo}="Financeiro"),'
        f'IF(COUNTIFS({crit_fin})>0,'
        f'MINIFS(financeiro!$A$2:$A${ULT_FIN},{crit_fin}),EOMONTH({ini},-1)+1),"")'
    )
    f[L_ULTIMA] = (
        f'=IF(AND({ha}=1,{metodo}="Financeiro"),'
        f'IF(COUNTIFS({crit_fin})>0,'
        f'MAXIFS(financeiro!$A$2:$A${ULT_FIN},{crit_fin}),EOMONTH({efe},-2)+1),"")'
    )
    f[L_PERIODO_TXT] = (
        f'=IF({ha}<>1,"",IF({metodo}="Financeiro",'
        f'{_mes_texto(prim)}&IF({prim}={ult},""," a "&{_mes_texto(ult)}),'
        f'{_data_texto(ini)}&" a "&{_data_texto(efe + "-1")}))'
    )
    f[L_EFEITO_TXT] = (
        f'=IF(NOT(ISNUMBER({efe})),"",IF({metodo}="Financeiro",'
        f'{_mes_texto(efe)},{_data_texto(efe)}))'
    )
    f[L_LINHA1] = (
        f'=IF({est}="","",{c}{L_CICLO}&IF({est}="D",{ref_texto("T_EFEITOS_DESDE")},'
        f'{ref_texto("T_EFEITOS_A_PARTIR")}&{efe_txt}))'
    )
    f[L_LINHA2] = (
        f'=IF(OR({est}="",{est}="D"),"",'
        f'IF({est}="B",{ref_texto("T_SEM_EXECUCAO")}&{per_txt},'
        f'IF({est}="C",{ref_texto("T_PERIODO_SEM_EFEITO")}&{per_txt},'
        f'{qtd}&IF({metodo}="PCs",{ref_texto("T_PCS")},'
        f'{ref_texto("T_COMPETENCIAS")})&{per_txt})))'
    )
    return f


def formulas_flags() -> dict[int, str]:
    est = f"T{L_ESTADO}:W{L_ESTADO}"
    return {
        L_FLAG_VISIVEL: f'=IF(COUNTIF({est},"?*")>0,1,0)',
        L_FLAG_VALORES: (
            f'=IF(COUNTIF({est},"A")+COUNTIF({est},"B")>0,1,0)'
        ),
        L_FLAG_C: f'=IF(COUNTIF({est},"C")>0,1,0)',
        L_FLAG_D: (
            f'=IF(AND(T{L_FLAG_VISIVEL}=1,T{L_FLAG_VALORES}=0,'
            f'T{L_FLAG_C}=0),1,0)'
        ),
    }


def formulas_resultados() -> dict[str, str]:
    """Formulas das celulas de RESULTADOS!E15:H21 (ASCII)."""
    vis = f"{MEM}!$T${L_FLAG_VISIVEL}"
    val = f"{MEM}!$T${L_FLAG_VALORES}"
    f: dict[str, str] = {
        "E15": f'=IF({vis}=1,{ref_texto_ext("T_TITULO")},"")',
        "F15": (
            f'=IF({val}=1,IF({MEM}!$B$4="Financeiro",'
            f'{ref_texto_ext("T_HDR_PAGO")},{ref_texto_ext("T_HDR_ORIGINAL")}),"")'
        ),
        "G15": f'=IF({val}=1,{ref_texto_ext("T_HDR_TERIAM")},"")',
        "H15": f'=IF({val}=1,{ref_texto_ext("T_HDR_DIFERENCA")},"")',
        "E16": f'=IF({vis}=1,{ref_texto_ext("T_NOTA")},"")',
        "E21": (
            f'=IF({MEM}!$T${L_FLAG_C}=1,{ref_texto_ext("T_CASO_C")},'
            f'IF({MEM}!$T${L_FLAG_D}=1,{ref_texto_ext("T_CASO_D")},""))'
        ),
    }
    for n, (c, _p, _cc, linha) in CICLOS.items():
        est = f"{MEM}!{c}${L_ESTADO}"
        ab = f'OR({est}="A",{est}="B")'
        l1, l2 = f"{MEM}!{c}${L_LINHA1}", f"{MEM}!{c}${L_LINHA2}"
        f[f"E{linha}"] = (
            f'=IF({est}="","",{l1}&IF({l2}="","",CHAR(10)&{l2}))'
        )
        f[f"F{linha}"] = f'=IF({ab},{MEM}!{c}${L_ORIGINAL},"")'
        f[f"G{linha}"] = f'=IF({ab},{MEM}!{c}${L_TERIAM},"")'
        f[f"H{linha}"] = (
            f'=IF({ab},{MEM}!{c}${L_DIFERENCA},'
            f'IF({est}="C",{ref_texto_ext("T_NAO_MENSURAVEL")},""))'
        )
    return f


def formulas_helpers() -> dict[str, str]:
    """Todas as formulas de MEMORIA_RESULTADOS (endereco -> formula)."""
    saida: dict[str, str] = {}
    for n, (c, *_resto) in CICLOS.items():
        for linha, formula in formulas_ciclo(n).items():
            saida[f"{c}{linha}"] = formula
    for linha, formula in formulas_flags().items():
        saida[f"T{linha}"] = formula
    return saida


# ------------------------------------------------------------------ constantes
XL_LEFT, XL_RIGHT, XL_CENTER, XL_VCENTER = -4131, -4152, -4108, -4108
XL_EXPRESSION = 2
XL_PASTE_FORMATS = -4122
XL_CONTINUOUS, XL_MEDIUM = 1, -4138
XL_CALCULO_MANUAL, XL_CALCULO_AUTOMATICO = -4135, -4105

AZUL_ESCURO = "1F4E78"
CINZA_TEXTO = "595959"
AMBAR_FUNDO = "FFF2CC"
AMBAR_TEXTO = "7F6000"

MOEDA_LOCAL_TEXTO = '"R$" #.##0,00;-"R$" #.##0,00;"R$" 0,00;@'
MOEDA_INVARIANTE_TEXTO = '"R$" #,##0.00;-"R$" #,##0.00;"R$" 0.00;@'

ALTURA_NOTA = 20.0
ALTURA_CICLO = 27.0
ALTURA_CASO = 26.0


def _bgr(rgb: str) -> int:
    r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
    return b * 65536 + g * 256 + r


def _fechar(wb, tentativas: int = 10) -> None:
    """Excel recusa Close logo apos um Save pesado (RPC_E_CALL_REJECTED)."""
    for _ in range(tentativas):
        try:
            wb.Close(SaveChanges=False)
            return
        except Exception:
            time.sleep(1.0)


def _formato_data(rng) -> None:
    for formato in ("dd/mm/aaaa", "dd/mm/yyyy"):
        try:
            rng.NumberFormatLocal = formato
            return
        except Exception:
            continue


def _copiar_formato(origem, destino) -> None:
    origem.Copy()
    destino.PasteSpecial(XL_PASTE_FORMATS)


def _exigir_area_livre(ws) -> None:
    """A area S78:W120 precisa estar livre OU conter o que este aplicador grava."""
    esperados = set(formulas_helpers())
    ultima = L_TEXTOS + len(TEXTOS) + 1
    for linha in range(L_TITULO, ultima + 1):
        for coluna in "STUVW":
            celula = ws.Range(f"{coluna}{linha}")
            conteudo = celula.Formula
            if conteudo in (None, ""):
                continue
            if coluna == "S" or f"{coluna}{linha}" in esperados:
                continue
            if coluna == "T" and linha > L_TEXTOS:
                continue
            if linha == L_CICLO:
                continue
            raise RuntimeError(
                f"{ABA_MEMORIA}!{coluna}{linha} ja possui conteudo: {conteudo!r}"
            )


def _aplicar_helpers(wb) -> None:
    ws = wb.Worksheets(ABA_MEMORIA)
    _exigir_area_livre(ws)

    ws.Range(f"S{L_TITULO}").Value = (
        "SEM EFEITO FINANCEIRO — apoio informativo (não compõe VTA, retroativo "
        "nem total)"
    )
    ws.Range(f"S{L_CICLO}").Value = "Ciclo"
    for n, (c, *_resto) in CICLOS.items():
        ws.Range(f"{c}{L_CICLO}").Value = f"C{n}"

    rotulos = {
        L_COMPUTADO: "Ciclo computado nesta apuração? (1/0)",
        L_INI_CICLO: "Início do ciclo (parametros!C)",
        L_INI_EFEITO: "Início dos efeitos financeiros (parametros!H)",
        L_LIMITE: "Limite: 1º dia do mês do início dos efeitos",
        L_HA_PERIODO: "Há período anterior aos efeitos? (1/0; vazio = indeterminado)",
        L_QTD: "Qtd. de PCs / competências com execução sem efeito",
        L_QTD_FATOR: "PCs com fator numérico (qtd) / fator do ciclo (financeiro!D)",
        L_ORIGINAL: "Valor original da execução sem efeito",
        L_TERIAM: "Valor que teriam com o reajuste",
        L_DIFERENCA: "Diferença sem efeito financeiro (informativa)",
        L_CONSUMO: "Consumo informado no ciclo (qtd; método Itens)",
        L_ESTADO: "Estado: A valores | B sem execução | C não mensurável | D sem período",
        L_PRIMEIRA: "Primeira competência sem efeito (Financeiro)",
        L_ULTIMA: "Última competência sem efeito (Financeiro)",
        L_PERIODO_TXT: "Período sem efeito (texto)",
        L_EFEITO_TXT: "Início dos efeitos (texto)",
        L_LINHA1: "Linha 1 do quadro (texto)",
        L_LINHA2: "Linha 2 do quadro (texto)",
        L_FLAG_VISIVEL: "Quadro visível? (1/0)",
        L_FLAG_VALORES: "Cabeçalho de valores visível? (1/0)",
        L_FLAG_C: "Há ciclo não mensurável? (1/0)",
        L_FLAG_D: "Somente ciclos sem período anterior? (1/0)",
    }
    for linha, texto in rotulos.items():
        ws.Range(f"S{linha}").Value = texto
    ws.Range(f"S{L_TEXTOS}").Value = (
        "TEXTOS DO QUADRO (constantes; as fórmulas acima são ASCII)"
    )
    for chave, texto in TEXTOS:
        linha = LINHA_TEXTO[chave]
        ws.Range(f"S{linha}").Value = chave
        ws.Range(f"T{linha}").NumberFormat = "@"
        ws.Range(f"T{linha}").Value = texto

    for endereco, formula in formulas_helpers().items():
        ws.Range(endereco).Formula = formula

    _formato_data(ws.Range(f"T{L_INI_CICLO}:W{L_LIMITE}"))
    _formato_data(ws.Range(f"T{L_PRIMEIRA}:W{L_ULTIMA}"))


def _aplicar_resultados(wb) -> None:
    ws = wb.Worksheets(ABA_RESULTADOS)

    # E16:H21 era um unico merge com o texto antigo. Desfaz e remonta.
    ws.Range("E16:H21").UnMerge()
    ws.Range("E16:H21").ClearContents()

    for endereco, formula in formulas_resultados().items():
        ws.Range(endereco).Formula = formula

    # Cabecalho: E15 herda o formato do titulo (A15); F15:H15, o das colunas.
    _copiar_formato(ws.Range("A15"), ws.Range("E15"))
    _copiar_formato(ws.Range("B15"), ws.Range("F15:H15"))
    wb.Application.CutCopyMode = False
    e15 = ws.Range("E15")
    e15.HorizontalAlignment = XL_LEFT
    e15.WrapText = True
    ws.Range("F15:H15").WrapText = True

    # Linha 16 (alinhada a C0): nota informativa em faixa unica.
    ws.Range("E16:H16").Merge()
    nota = ws.Range("E16")
    nota.Font.Italic = True
    nota.Font.Size = 8
    nota.Font.Color = _bgr(CINZA_TEXTO)
    nota.WrapText = True
    nota.HorizontalAlignment = XL_LEFT
    nota.VerticalAlignment = XL_VCENTER
    ws.Rows("16:16").RowHeight = ALTURA_NOTA

    # Linhas 17..20 (C1..C4): descricao + tres valores, no padrao da tabela 2.
    for _n, (_c, _p, _cc, linha) in CICLOS.items():
        _copiar_formato(ws.Range(f"A{linha}"), ws.Range(f"E{linha}"))
        _copiar_formato(ws.Range(f"B{linha}"), ws.Range(f"F{linha}:H{linha}"))
        ws.Rows(f"{linha}:{linha}").RowHeight = ALTURA_CICLO
    wb.Application.CutCopyMode = False
    descricao = ws.Range("E17:E20")
    descricao.Font.Size = 8
    descricao.Font.Italic = False
    descricao.Font.Color = _bgr(CINZA_TEXTO)
    descricao.WrapText = True
    descricao.HorizontalAlignment = XL_LEFT
    descricao.VerticalAlignment = XL_VCENTER
    valores = ws.Range("F17:H20")
    # Moeda brasileira com secao de TEXTO preservada (@): o formato copiado da
    # tabela 2 troca texto por "-" e esconderia "Nao mensuravel". Ciclo sem uso
    # (texto vazio) fica em branco.
    try:
        valores.NumberFormatLocal = MOEDA_LOCAL_TEXTO
    except Exception:
        valores.NumberFormat = MOEDA_INVARIANTE_TEXTO
    valores.HorizontalAlignment = XL_RIGHT
    valores.VerticalAlignment = XL_VCENTER
    valores.Font.Size = 10

    # Destaque ambar claro (nunca vermelho) na coluna da diferenca.
    diferenca = ws.Range("H17:H20")
    diferenca.FormatConditions.Delete()
    regra = diferenca.FormatConditions.Add(XL_EXPRESSION, None, '=$H17<>""')
    regra.Interior.Color = _bgr(AMBAR_FUNDO)
    regra.Font.Color = _bgr(AMBAR_TEXTO)

    # Separador do quadro em relacao a tabela 2: friso branco no cabecalho e
    # friso azul nas linhas de ciclo (a faixa azul da linha 15 e continua).
    separador_cab = ws.Range("E15").Borders(7)
    separador_cab.LineStyle = XL_CONTINUOUS
    separador_cab.Weight = XL_MEDIUM
    separador_cab.Color = _bgr("FFFFFF")
    separador_corpo = ws.Range("E17:E20").Borders(7)
    separador_corpo.LineStyle = XL_CONTINUOUS
    separador_corpo.Weight = XL_MEDIUM
    separador_corpo.Color = _bgr(AZUL_ESCURO)

    # Linha 21: mensagem de caso (nao mensuravel / sem periodo anterior).
    ws.Range("E21:H21").Merge()
    caso = ws.Range("E21")
    caso.Font.Italic = True
    caso.Font.Size = 8
    caso.Font.Color = _bgr(CINZA_TEXTO)
    caso.WrapText = True
    caso.HorizontalAlignment = XL_LEFT
    caso.VerticalAlignment = XL_VCENTER
    ws.Rows("21:21").RowHeight = ALTURA_CASO


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
        # Recalculo manual durante a escrita (evita Excel "ocupado" a cada
        # celula); restaurado para automatico ANTES do Save, para nao gravar o
        # modo manual no template.
        excel.Calculation = XL_CALCULO_MANUAL
        _aplicar_helpers(wb)
        _aplicar_resultados(wb)
        excel.Calculation = XL_CALCULO_AUTOMATICO

        ws = wb.Worksheets(ABA_RESULTADOS)
        ws.Activate()
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1
        ws.Range("D5").Select()
        wb.Worksheets("CONTROLE").Activate()
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
        # Solta os proxies COM antes de desinicializar (evita 0x80010108).
        wb = ws = excel = None
        gc.collect()
        pythoncom.CoUninitialize()


def main() -> int:
    caminho = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else TEMPLATE
    if not caminho.exists():
        print("ERRO: arquivo nao encontrado: %s" % caminho)
        return 1
    import pywintypes

    # RPC_E_CALL_REJECTED (Excel ocupado): a falha ocorre antes do Save, entao
    # reexecutar do zero e seguro e nao deixa o arquivo parcialmente alterado.
    for tentativa in range(1, 6):
        try:
            aplicar(caminho)
            break
        except pywintypes.com_error as erro:
            if erro.args[0] != -2147418111 or tentativa == 5:
                raise
            print("Excel ocupado (tentativa %d); reexecutando do zero..." % tentativa)
            time.sleep(8.0)
    print("OK: RESULTADOS-SEM-EFEITO-150 aplicado em %s" % caminho)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
