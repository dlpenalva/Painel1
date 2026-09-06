# -*- coding: utf-8 -*-
"""PC-VTA-POT-TOTAL-1: regra patria do VTA no metodo PC.

Regra de negocio (decidida; nao reaberta aqui):

    RETROATIVO CONSIDERADO NO VTA
        = RETROATIVO RECONHECIDO + SOMA DOS POTENCIAIS POSITIVOS
    VTA OFICIAL
        = VTA SEM POTENCIAL + SOMA DOS POTENCIAIS POSITIVOS

O ciclo do PC deixa de limitar a incorporacao: o potencial do ciclo VIGENTE
tambem integra o VTA. O objetivo prudencial e nao SUBdimensionar o contrato.

Potencial NEGATIVO continua visivel e identificado como POTENCIAL, mas nao
reduz o VTA e nao compensa parcela positiva — a incorporacao soma as parcelas
POSITIVAS uma a uma, nunca ``MAX(soma_liquida, 0)``.

FONTE CANONICA (unica): ``MEMORIA_RESULTADOS!$T$39`` (named range
``RETROATIVO_POTENCIAL_VTA``). Todos os consumidores ja LEEM esse named range
— nenhum deles recalcula o potencial. Este aplicador so corrige a fonte e os
textos que descreviam a regra antiga.

MEMORIA_RESULTADOS (aba oculta de apoio)
    T42:T46  NOVAS — potencial POSITIVO por ciclo (C0..C4), PC a PC, ate o
             corte. Sao elas que impedem o negativo de compensar o positivo.
    T39      passa a somar T42:T46 (era ``MAX(T41,0)``, restrito aos ciclos
             ja encerrados).
    T41      passa a ser o apurado LIQUIDO de todos os ciclos ate o corte
             (informativo; pode ser negativo).
    T47      NOVA — a parcela potencial NEGATIVA que ficou de fora (<= 0).

itens_PC
    19       rotulo do fechamento reescrito pela regra patria (S19, o valor,
             ja soma Q18 + T39 e nao muda de formula).
    9/18/19  destaque de TOTAL/fechamento na paleta da propria aba; a coluna
             S (potencial) recebe o ambar ja adotado no produto.
    V:AC     reasseguradas ocultas.

RESULTADOS
    C61/C84/C86  textos: o discriminante do potencial negativo passa a ser
                 T47<0 (existe parcela negativa) e nao mais T41<0 — no caso
                 MISTO o VTA incorpora o positivo E ha negativo a declarar.

itens_PC!A:L, financeiro, itens_Consumidos, aditivos e as ancoras
EXECUTADO_APURADO (B83), AJUSTES_DEVIDOS (B84), VTA_FINAL (B86) e
CONFERENCIA_FORMACAO_VTA (B87) ficam intactas.

O aplicador usa Excel COM para preservar recursos x14 que o openpyxl remove.

Uso: <python> tools/aplicar_pc_vta_pot_total_1.py [caminho.xlsx]
"""
from __future__ import annotations

import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

XL_LEFT, XL_RIGHT, XL_CENTER = -4131, -4152, -4108
XL_VCENTER = -4108
XL_CONTINUOUS, XL_THIN, XL_MEDIUM = 1, 2, -4138
XL_EDGE_TOP = 8

AZUL_CABECALHO = "D9EAF7"     # mesma faixa dos cabecalhos M2:T2 / M11:T11
AZUL_BORDA_TOTAL = "9DC3E6"
TEXTO_PADRAO = "1F3864"
CINZA_BORDA = "BFBFBF"

# Par unico da parcela POTENCIAL no produto inteiro (XLS, web e documentos).
POTENCIAL_BG = "FFF4CC"
POTENCIAL_TEXTO = "7F6000"

MOEDA_LOCAL = '"R$" #.##0,00;-"R$" #.##0,00;"R$" 0,00;"-"'
MOEDA_INVARIANTE = '"R$" #,##0.00;-"R$" #,##0.00;"R$" 0.00;"-"'

METODO = 'MEMORIA_RESULTADOS!$B$4'
E_PC = f'{METODO}="PCs"'

# ------------------------------------------------- MEMORIA_RESULTADOS (fonte)
S39_NOVO = (
    "Retroativo potencial INCORPORADO ao VTA (soma das parcelas POSITIVAS, "
    "PC a PC, de todos os ciclos ate a data de corte)"
)
F_T39 = "=ROUND(SUM($T$42:$T$46),2)"

S41_NOVO = (
    "Retroativo potencial APURADO liquido (informativo; pode ser negativo e "
    "nunca e o valor somado ao VTA)"
)
F_T41 = "=ROUND(SUM(itens_PC!$S$12:$S$16),2)"

# Uma linha por ciclo: SOMENTE itens_PC!J > 0, PC a PC. E o que impede um
# potencial negativo de compensar um positivo dentro do mesmo ciclo.
CICLOS_POTENCIAL = ("C0", "C1", "C2", "C3", "C4")
LINHA_POTENCIAL_INICIAL = 42


def _formula_positivo(ciclo: str) -> str:
    return (
        "=ROUND(SUMIFS(itens_PC!$J$2:$J$5001,"
        f'itens_PC!$C$2:$C$5001,"{ciclo}",'
        'itens_PC!$J$2:$J$5001,">0",'
        'itens_PC!$B$2:$B$5001,"<="&$T$31),2)'
    )


S47_NOVO = (
    "Parcela potencial NEGATIVA apurada (permanece visivel; nao reduz o VTA "
    "nem compensa parcela positiva)"
)
F_T47 = "=ROUND($T$41-$T$39,2)"

NOME_NEGATIVO = "RETROATIVO_POTENCIAL_NEGATIVO"
REF_NEGATIVO = "=MEMORIA_RESULTADOS!$T$47"

# ------------------------------------------------------------------ itens_PC
F_PC_19_ROTULO = (
    '="RETROATIVO CONSIDERADO NO VTA = RECONHECIDO R$ "&'
    'TEXT(N($Q$18),"#.##0,00")&"  +  POTENCIAL POSITIVO INCORPORADO R$ "&'
    'TEXT(N(MEMORIA_RESULTADOS!$T$39),"#.##0,00")&'
    "IF(N(MEMORIA_RESULTADOS!$T$47)<0,"
    '"   (há ainda R$ "&TEXT(-N(MEMORIA_RESULTADOS!$T$47),"#.##0,00")&'
    '" de potencial negativo: permanece visível e não reduz o VTA)",'
    '"   (todo o potencial positivo apurado até a data de corte integra o VTA)")'
)
F_PC_19_VALOR = "=ROUND(N($Q$18)+N(MEMORIA_RESULTADOS!$T$39),2)"

LINHAS_TOTAL_PC = (9, 18, 19)

# --------------------------------------------------------------- RESULTADOS
C61_NOVO = (
    "Soma de TODAS as parcelas potenciais positivas dos PCs em análise até a "
    "data de corte — inclusive as do ciclo vigente. O VTA a incorpora por "
    "critério prudencial. Não é retroativo reconhecido nem valor a pagar."
)

F_C84 = (
    f'=IF({METODO}="Financeiro",'
    '"Reajuste ja reconhecido e ainda nao contido no valor pago.",'
    f"IF({E_PC},"
    "IF(MEMORIA_RESULTADOS!$T$47<0,"
    '"Retroativo POTENCIAL positivo dos PCs em analise, incorporado ao VTA '
    'por criterio prudencial. Ha ainda R$ "&'
    'TEXT(-MEMORIA_RESULTADOS!$T$47,"#.##0,00")&'
    '" de parcela potencial NEGATIVA apurada: ela permanece visivel, nao '
    'reduz o VTA e nao compensa a parcela positiva.",'
    '"Retroativo POTENCIAL dos PCs ainda em analise pela area gestora, '
    "incorporado ao VTA por criterio prudencial. Nao e retroativo reconhecido "
    'a pagar."),'
    f'IF({METODO}="Itens",'
    '"Nao aplicavel: o reajuste ja esta dentro da execucao atualizada.","")))'
)

F_C86 = (
    f"=IF(AND({E_PC},ISNUMBER(MEMORIA_RESULTADOS!$T$40),"
    "ISNUMBER(MEMORIA_RESULTADOS!$T$39),"
    "OR(MEMORIA_RESULTADOS!$T$39<>0,MEMORIA_RESULTADOS!$T$47<>0)),"
    '"VTA antes da parcela potencial: R$ "&TEXT(MEMORIA_RESULTADOS!$T$40,'
    '"#.##0,00")&"  +  Parcela potencial no VTA: R$ "&'
    'TEXT(MEMORIA_RESULTADOS!$T$39,"#.##0,00")&"  =  VALOR TOTAL ATUALIZADO "&'
    '"— VTA."&IF(MEMORIA_RESULTADOS!$T$47<0,'
    '"  Há ainda R$ "&TEXT(-MEMORIA_RESULTADOS!$T$47,'
    '"#.##0,00")&" de parcela potencial negativa apurada: por critério "&'
    '"prudencial ela permanece visível, não reduz o VTA e não compensa a "&'
    '"parcela positiva.",'
    '"  O VTA inclui essa parcela por critério prudencial; ela permanece "&'
    '"sujeita à confirmação pela área gestora e não representa, nesta data, "&'
    '"retroativo reconhecido a pagar."),'
    '"Resultado final do método de apuração selecionado.")'
)


def _bgr(hex_rgb: str) -> int:
    return int(hex_rgb[4:6] + hex_rgb[2:4] + hex_rgb[0:2], 16)


def _fechar(wb) -> None:
    try:
        wb.Close(SaveChanges=False)
    except Exception:
        pass


def _moeda(rng) -> None:
    try:
        rng.NumberFormatLocal = MOEDA_LOCAL
    except Exception:
        rng.NumberFormat = MOEDA_INVARIANTE


def _bordas(rng) -> None:
    for indice in (7, 8, 9, 10, 11, 12):
        borda = rng.Borders(indice)
        borda.LineStyle = XL_CONTINUOUS
        borda.Weight = XL_THIN
        borda.Color = _bgr(CINZA_BORDA)


# ------------------------------------------------------- MEMORIA_RESULTADOS
def _aplicar_memoria(mem) -> None:
    """Fonte canonica unica do potencial incorporado."""
    mem.Range("S39").Value = S39_NOVO
    mem.Range("T39").Formula = F_T39
    mem.Range("S41").Value = S41_NOVO
    mem.Range("T41").Formula = F_T41

    for indice, ciclo in enumerate(CICLOS_POTENCIAL):
        linha = LINHA_POTENCIAL_INICIAL + indice
        mem.Range(f"S{linha}").Value = (
            f"Potencial POSITIVO {ciclo} (itens_PC!J > 0, ate a data de corte)"
        )
        mem.Range(f"T{linha}").Formula = _formula_positivo(ciclo)

    mem.Range("S47").Value = S47_NOVO
    mem.Range("T47").Formula = F_T47

    primeira = LINHA_POTENCIAL_INICIAL
    ultima = LINHA_POTENCIAL_INICIAL + len(CICLOS_POTENCIAL) - 1
    _moeda(mem.Range(f"T{primeira}:T{ultima}"))
    _moeda(mem.Range("T47"))
    _moeda(mem.Range("T39"))
    _moeda(mem.Range("T41"))


def _aplicar_nome(wb) -> None:
    """Named range da parcela negativa (leitura, nunca recalculo)."""
    for indice in range(wb.Names.Count, 0, -1):
        try:
            if str(wb.Names.Item(indice).Name) == NOME_NEGATIVO:
                wb.Names.Item(indice).Delete()
        except Exception:
            continue
    wb.Names.Add(Name=NOME_NEGATIVO, RefersTo=REF_NEGATIVO)


# ------------------------------------------------------------------ itens_PC
def _destacar_total(ws, linha: int) -> None:
    """Destaque de TOTAL/fechamento na paleta da propria aba (M..T)."""
    faixa = ws.Range(f"M{linha}:T{linha}")
    faixa.Interior.Color = _bgr(AZUL_CABECALHO)
    faixa.Font.Bold = True
    faixa.Font.Color = _bgr(TEXTO_PADRAO)
    _bordas(faixa)
    topo = faixa.Borders(XL_EDGE_TOP)
    topo.LineStyle = XL_CONTINUOUS
    topo.Weight = XL_MEDIUM
    topo.Color = _bgr(AZUL_BORDA_TOTAL)


def _destacar_potencial(ws, celula: str) -> None:
    """Celula que representa ESPECIFICAMENTE potencial: ambar suave."""
    alvo = ws.Range(celula)
    alvo.Interior.Color = _bgr(POTENCIAL_BG)
    alvo.Font.Color = _bgr(POTENCIAL_TEXTO)
    alvo.Font.Bold = True


def _aplicar_itens_pc(ws) -> None:
    principal_antes = ws.Range("A1:L5001").Formula

    # Rotulo do fechamento (a formula do VALOR nao muda de significado).
    ws.Range("M19").Formula = F_PC_19_ROTULO
    ws.Range("S19").Formula = F_PC_19_VALOR

    for linha in LINHAS_TOTAL_PC:
        _destacar_total(ws, linha)

    # S9 e S18 sao os TOTAIS da coluna RETROATIVO POTENCIAL: ambar, como os
    # cabecalhos S2/S11 ja adotam. A linha 19 e retroativo CONSIDERADO
    # (reconhecido + potencial): cor de total, nunca ambar.
    _destacar_potencial(ws, "S9")
    _destacar_potencial(ws, "S18")

    ws.Range("M19:R19").WrapText = True
    ws.Range("M19:R19").HorizontalAlignment = XL_LEFT
    ws.Range("M19:R19").VerticalAlignment = XL_VCENTER
    ws.Range("M19").Font.Size = 9
    ws.Range("S19:T19").HorizontalAlignment = XL_RIGHT
    _moeda(ws.Range("S19:T19"))

    # V:AC seguem ocultas (template e copia entregue).
    ws.Range("V:AC").EntireColumn.Hidden = True

    if ws.Range("A1:L5001").Formula != principal_antes:
        raise SystemExit("ABORTADO: itens_PC!A:L foi alterada.")


# ----------------------------------------------------------------- RESULTADOS
def _aplicar_resultados(ws) -> None:
    ws.Range("C61").Value = C61_NOVO
    ws.Range("C84").Formula = F_C84
    ws.Range("C86").Formula = F_C86


def main(caminho: Path) -> None:
    import win32com.client as com

    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(caminho))
        _aplicar_memoria(wb.Worksheets("MEMORIA_RESULTADOS"))
        _aplicar_nome(wb)
        _aplicar_itens_pc(wb.Worksheets("itens_PC"))
        _aplicar_resultados(wb.Worksheets("RESULTADOS"))
        excel.CalculateFullRebuild()
        wb.Save()
        print(f"OK: {caminho}")
    finally:
        _fechar(wb)
        try:
            excel.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    alvo = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else TEMPLATE
    if not alvo.exists():
        raise SystemExit(f"Arquivo nao encontrado: {alvo}")
    main(alvo)
