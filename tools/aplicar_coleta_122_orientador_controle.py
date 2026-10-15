# -*- coding: utf-8 -*-
"""Coleta 12.2 — ORIENTADOR DA COLETA na aba CONTROLE.

Somente ORIENTACAO/COLETA. Nenhuma formula economica, fator, VTA,
retroativo ou admissibilidade muda; nenhuma aba de calculo e tocada (o quadro
so LE CONTROLE!B1:B3, MEMORIA_RESULTADOS!B4, parametros!U3:U6 e aditivos!A).

LAYOUT DA CONTROLE
  * bloco A:B preservado (B4 so ganha quebra de texto: o texto legado
    transbordava por C:G e D4 deixa de estar livre). As linhas 4 e 15:18,
    ocultas na 12.1, voltam a ser exibidas (D4, diagnostico, marcadores);
    a 14 (data de geracao) segue oculta;
  * coluna C reduzida a separador; a explicacao da data de corte (antes em C3)
    passa a D4; sem linhas de grade;
  * a orientacao do metodo em C1 (Ajuste 3) passa a pintar C1:H1, pois o texto
    da C1 estreita transborda por D1:H1 (mesma regra, mesma formula);
  * marcadores de versao: A16:B17 (gravados em runtime por
    _coleta_oficial.registrar_versoes_coleta; antes A24:B25).

ORIENTADOR (D5:H34)
  D5   titulo + introducao
  D7   PARAMETROS COMPLEMENTARES: historico completo? (E8) e aditivos (E9)
  D10  CICLOS COM PEDIDO DE REAJUSTE: leitura automatica de parametros!U
       (DATA_PEDIDO) e, so na falta dela, dropdown Sim/Nao (E12:H12)
  D15  DIAGNOSTICO — INFORMACAO / STATUS / POR QUE
  D25  O QUE VOCE VAI PEDIR AO FISCAL (texto dinamico, itens numerados)
  D34  objetivos da coleta

AUXILIARES (colunas J:L ocultas)
  J17:J23 codigo de status (1..5); J37:J49 estado; J40:L46 itens candidatos
  (flag, ordem, texto); J50:K89 constantes de texto (acentos so como VALOR —
  regra ZERO CORRUPCAO: formulas ASCII, parenteses balanceados).

uso: python tools/aplicar_coleta_122_orientador_controle.py [caminho.xlsx]
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
ABA = "CONTROLE"

# ------------------------------------------------------------- constantes (J:K)
CONSTANTES = {
    50: ("ST_OBRIGATORIO", "OBRIGATÓRIO"),
    51: ("ST_COMPLEMENTAR", "COMPLEMENTAR"),
    52: ("ST_DISPENSAVEL", "DISPENSÁVEL"),
    53: ("ST_VERIFICAR", "VERIFICAR"),
    54: ("ST_NAO_SE_APLICA", "NÃO SE APLICA"),
    56: ("HIST_SIM", "Sim"),
    57: ("HIST_PARCIAL", "Parcialmente"),
    58: ("HIST_NAO_CONFIRMADO", "Não confirmado"),
    60: ("ADT_NAO", "Não"),
    61: ("ADT_SIM", "Sim"),
    62: ("ADT_NAO_SEI", "Não sei"),
    64: ("PED_SIM", "Sim"),
    65: ("PED_NAO", "Não"),
    66: ("PED_NAO_INFORMADO", "Não informado"),
    67: ("CICLO_PADRAO", "último ciclo"),
    70: ("SAIDA_ABERTURA", "Para concluir a apuração, favor informar:"),
    71: ("SAIDA_SEM_METODO", "Selecione primeiro o método de apuração em CONTROLE!B1."),
    72: ("SAIDA_SEM_CORTE", "Informe a data de corte da apuração em CONTROLE!B3."),
    73: ("ITEM_PCS", "os Pedidos de Compra mais recentes disponíveis até a data de corte "
                    "da apuração, indicando quais foram pagos/aceitos e quais não foram"),
    74: ("ITEM_FIN", "os valores do financeiro por competência, efetivamente pagos, até a "
                    "data de corte da apuração"),
    75: ("ITEM_ITENS_A", "as quantidades remanescentes por item na abertura de cada ciclo "
                        "anterior ao "),
    76: ("ITEM_ITENS_B", ", para reconstruir a execução com o VU de cada período"),
    77: ("ITEM_HIST_PARCIAL", "posição física ou evidência equivalente apenas para os "
                             "períodos não cobertos pelo histórico da execução"),
    78: ("ITEM_HIST_CONF_A", "a confirmação de que o histórico da execução cobre todo o "
                            "período até a abertura do "),
    79: ("ITEM_HIST_CONF_B", "; nos períodos não cobertos, posição física ou evidência "
                            "equivalente"),
    80: ("ITEM_ABERTURA", "as quantidades remanescentes por item na abertura do "),
    81: ("ITEM_ADT_A", "os acréscimos e supressões ocorridos desde a abertura do "),
    82: ("ITEM_ADT_B", ", com item, quantidade e data de efeito"),
    83: ("ITEM_ADT_CONF_A", "a confirmação sobre acréscimos ou supressões desde a abertura do "),
    84: ("ITEM_ADT_CONF_B", " (havendo, item, quantidade e data de efeito)"),
    85: ("ITEM_ATUAL_A", "as quantidades remanescentes por item na data de corte ("),
    86: ("ITEM_ATUAL_B", "), correspondente à posição atual"),
    87: ("ADT_LANCADOS", "Lançamentos já registrados na aba aditivos: "),
    88: ("DIAG_ABERTURA", "Posição física na abertura do "),
    89: ("SAIDA_CORTE_FORA", "A data de corte não está dentro dos ciclos da apuração; "
                            "confira CONTROLE!B3."),
}
K = {chave: f"$K${linha}" for linha, (chave, _t) in CONSTANTES.items()}

# ------------------------------------------------------------- textos visiveis
TITULO = "ORIENTADOR DA COLETA — O QUE PEDIR AO FISCAL"
INTRODUCAO = ("Use este quadro para definir o mínimo necessário à apuração: retroativo, "
              "VU por ciclo, VTA e saldo atualizado na abertura do último ciclo.")
EXPLICACAO_CORTE = "Data-limite considerada nos cálculos e resultados."
SECAO_PARAMETROS = "PARÂMETROS COMPLEMENTARES"
ROTULO_HISTORICO = "Histórico da execução está completo?"
DICA_HISTORICO = ("Sim: a fonte do método (PCs ou financeiro) cobre todo o período até a "
                  "abertura do ciclo em execução.")
ROTULO_ADITIVOS = "Aditivos/supressões no período"
SECAO_CICLOS = "CICLOS COM PEDIDO DE REAJUSTE"
CICLOS = ("C1", "C2", "C3", "C4")            # parametros!B3:B6
COLUNAS_CICLOS = ("E", "F", "G", "H")
ROTULOS_CICLOS = ("Leitura automática (parametros)", "Informar, se não houver leitura",
                  "Considerado")
SECAO_DIAGNOSTICO = "DIAGNÓSTICO — INFORMAÇÕES A SOLICITAR"
CABECALHO_DIAGNOSTICO = ("INFORMAÇÃO", "STATUS", "POR QUE / PARA QUE SERVE")
SECAO_SAIDA = "O QUE VOCÊ VAI PEDIR AO FISCAL"
OBJETIVOS = ("Objetivos da coleta: apurar retroativo, VU por ciclo, VTA e o valor "
             "remanescente atualizado na abertura do último ciclo.")

# ------------------------------------------------------------- estado (J37:J49)
M = "$J$39"            # codigo do metodo (MEMORIA_RESULTADOS!B4)
METODO_OK = "$J$37"
GATE = "$J$48"
CICLO = "$J$49"
DDMM = "$J$38"
TOTAL = "$J$47"
ESTADO = {
    "J37": f'=OR({M}="Financeiro",{M}="PCs",{M}="Itens")',
    "J38": ('=IF(ISNUMBER($B$3),RIGHT("0"&DAY($B$3),2)&"/"&'
            'RIGHT("0"&MONTH($B$3),2),"")'),
    "J39": "=MEMORIA_RESULTADOS!$B$4",
    "J47": "=SUM($J$40:$J$46)",
    "J48": f'=AND({METODO_OK},ISNUMBER($B$3),$B$2<>"")',
    "J49": f'=IF($B$2<>"",$B$2,{K["CICLO_PADRAO"]})',
}

# Itens candidatos do pedido ao fiscal, NA ORDEM de exibicao: (flag, texto).
ITENS = [
    (f'=AND({GATE},{M}="PCs")*1', f'={K["ITEM_PCS"]}'),
    (f'=AND({GATE},{M}="Financeiro")*1', f'={K["ITEM_FIN"]}'),
    (f'=AND({GATE},{M}="Itens")*1', f'={K["ITEM_ITENS_A"]}&{CICLO}&{K["ITEM_ITENS_B"]}'),
    (f'=AND({GATE},OR({M}="PCs",{M}="Financeiro"),$E$8<>{K["HIST_SIM"]})*1',
     f'=IF($E$8={K["HIST_PARCIAL"]},{K["ITEM_HIST_PARCIAL"]},'
     f'{K["ITEM_HIST_CONF_A"]}&{CICLO}&{K["ITEM_HIST_CONF_B"]})'),
    (f"={GATE}*1", f'={K["ITEM_ABERTURA"]}&{CICLO}'),
    (f'=AND({GATE},$E$9<>{K["ADT_NAO"]})*1',
     f'=IF($E$9={K["ADT_SIM"]},{K["ITEM_ADT_A"]}&{CICLO}&{K["ITEM_ADT_B"]},'
     f'{K["ITEM_ADT_CONF_A"]}&{CICLO}&{K["ITEM_ADT_CONF_B"]})'),
    (f"={GATE}*1", f'={K["ITEM_ATUAL_A"]}&{DDMM}&{K["ITEM_ATUAL_B"]}'),
]
PRIMEIRA_LINHA_ITEM = 40
LINHAS_SAIDA = range(27, 27 + len(ITENS))      # D27:D33

# Diagnostico (linhas 17..23): (informacao, codigo de status, por que).
_CICLO_N = '(MATCH($B$2,{"C0","C1","C2","C3","C4"},0)-1)'
DIAGNOSTICO = [
    ("Pedidos de Compra até a data de corte (pagos/aceitos e não pagos)",
     f'=IF({M}="PCs",1,IF({METODO_OK},5,4))',
     "Fonte da execução histórica no método PC."),
    ("Financeiro por competência até a data de corte",
     f'=IF({M}="Financeiro",1,IF({METODO_OK},5,4))',
     "Fonte da execução histórica no método Financeiro."),
    ("Posições físicas dos ciclos anteriores ao ciclo em execução",
     f'=IF({M}="Itens",1,IF({METODO_OK},IF($E$8={K["HIST_SIM"]},3,'
     f'IF($E$8={K["HIST_PARCIAL"]},2,4)),4))',
     "Consumido: reconstrói a execução com o VU de cada período. PC/Financeiro: "
     "dispensáveis com histórico completo; só cobrem lacunas do histórico."),
    (f'={K["DIAG_ABERTURA"]}&{CICLO}',
     "=1",
     "Saldo remanescente atualizado no marco e VU pertinente; base de comparação "
     "com a posição atual. Nunca é dispensada."),
    ("Posição física na data de corte (posição atual)",
     '=IF($B$2<>"",1,4)',
     "Comparada à abertura e ajustada por acréscimos/supressões, mede o consumo "
     "do ciclo em execução."),
    ("Acréscimos e supressões desde a abertura do ciclo",
     f'=IF($E$9={K["ADT_SIM"]},1,IF($E$9={K["ADT_NAO"]},5,4))',
     "Ajustam a diferença entre a abertura e a posição atual."),
    ("Ciclos com pedido de reajuste",
     f'=IFERROR(IF(SUMPRODUCT(($E$13:$H$13={K["PED_NAO_INFORMADO"]})*'
     f'((COLUMN($E$13:$H$13)-COLUMN($E$13)+1)<={_CICLO_N}))>0,4,3),4)',
     "Definem os ciclos com retroativo e efeitos financeiros a apurar; lidos de "
     "parametros (DATA_PEDIDO) quando disponíveis."),
]
PRIMEIRA_LINHA_DIAGNOSTICO = 17

# Cores por codigo de status (fundo, texto) — claras, como no resto da Coleta.
CORES_STATUS = {
    1: ("FCE4D6", "843C0C"),   # OBRIGATORIO
    2: ("EDE7F6", "5B3F8F"),   # COMPLEMENTAR
    3: ("E2EFDA", "375623"),   # DISPENSAVEL
    4: ("FBE5E8", "9B2C3F"),   # VERIFICAR
    5: ("EDEDED", "595959"),   # NAO SE APLICA
}

AZUL_CLARO, AZUL_SUAVE, CINZA, AMARELO = "DDEBF7", "F2F7FC", "EDEDED", "FFF2CC"
VERDE_FORTE, VERDE_CLARO = "A9D08E", "E2EFDA"
TEXTO, TEXTO_TITULO, TEXTO_CALCULO, TEXTO_VERDE = "1F1F1F", "1F4E79", "404040", "274E13"
BORDA = "BFBFBF"

LARGURAS = {"C": 2.0, "D": 40.0, "E": 18.0, "F": 30.0, "G": 16.0, "H": 16.0}
COLUNAS_OCULTAS = ("J", "K", "L")
ALTURAS = {4: 64.0, 5: 26.0, 6: 34.0, 7: 20.0, 13: 22.0, 15: 22.0, 16: 26.0,
           17: 30.0, 18: 30.0, 19: 30.0, 20: 30.0, 21: 30.0, 22: 30.0, 23: 30.0,
           25: 22.0, 26: 22.0, **{linha: 30.0 for linha in LINHAS_SAIDA}, 34: 22.0}
# Linhas ocultas no template 12.1 que passam a ser usadas (D4, diagnostico e
# marcadores de versao em A16:B17). A 14 (data de geracao) segue oculta.
LINHAS_REEXIBIDAS = (4, 15, 16, 17, 18)
LINHAS_OCULTAS = (14,)

XL_VALIDATE_LIST, XL_VALID_ALERT_STOP, XL_BETWEEN = 3, 1, 1
XL_EDGES = (7, 8, 9, 10)            # left, top, bottom, right
XL_INSIDE = (11, 12)                # vertical, horizontal


def formula_saida(i: int) -> str:
    return (f'=IF({TOTAL}<{i},"",{i}&". "&INDEX($L$40:$L$46,MATCH({i},$K$40:$K$46,0))'
            f'&IF({i}={TOTAL},".",";"))')


FORMULA_ABERTURA_SAIDA = (
    f'=IF(NOT({METODO_OK}),{K["SAIDA_SEM_METODO"]},IF(NOT(ISNUMBER($B$3)),'
    f'{K["SAIDA_SEM_CORTE"]},IF($B$2="",{K["SAIDA_CORTE_FORA"]},{K["SAIDA_ABERTURA"]})))'
)


def formula_ciclo_automatico(indice: int) -> str:
    return f'=IF(ISNUMBER(parametros!$U${indice + 3}),{K["PED_SIM"]},"")'


def formula_ciclo_considerado(coluna: str) -> str:
    return (f'=IF({coluna}11<>"",{coluna}11,IF({coluna}12<>"",{coluna}12,'
            f'{K["PED_NAO_INFORMADO"]}))')


def formulas() -> dict[str, str]:
    """Todas as formulas novas da CONTROLE (celula -> formula)."""
    resultado = dict(ESTADO)
    for i, (flag, texto) in enumerate(ITENS):
        linha = PRIMEIRA_LINHA_ITEM + i
        resultado[f"J{linha}"] = flag
        resultado[f"K{linha}"] = f'=IF(J{linha}=1,SUM($J${PRIMEIRA_LINHA_ITEM}:J{linha}),"")'
        resultado[f"L{linha}"] = texto
    for i, (info, codigo, _porque) in enumerate(DIAGNOSTICO):
        linha = PRIMEIRA_LINHA_DIAGNOSTICO + i
        resultado[f"J{linha}"] = codigo
        resultado[f"E{linha}"] = f"=INDEX($K$50:$K$54,J{linha})"
        if info.startswith("="):
            resultado[f"D{linha}"] = info
    for indice, coluna in enumerate(COLUNAS_CICLOS):
        resultado[f"{coluna}11"] = formula_ciclo_automatico(indice)
        resultado[f"{coluna}13"] = formula_ciclo_considerado(coluna)
    resultado["F9"] = f'={K["ADT_LANCADOS"]}&COUNTA(aditivos!$A$2:$A$200)'
    resultado["D26"] = FORMULA_ABERTURA_SAIDA
    for i, linha in enumerate(LINHAS_SAIDA, start=1):
        resultado[f"D{linha}"] = formula_saida(i)
    return resultado


def validar_ascii_e_parenteses() -> None:
    v3.validar_ascii_e_parenteses({f"{ABA}!{c}": f for c, f in formulas().items()})


# ------------------------------------------------------------- conferencias
AREA_NOVA = "D4:H34"
AREA_AUXILIAR = "J17:L89"
C3_ANTERIOR = ("Última data considerada nos cálculos, PCs, competências e resultados "
               "da apuração.")


def conferir(wb) -> None:
    """Forma 12.2 (PR #190) exigida antes de qualquer escrita (fail-closed)."""
    ws = wb.Worksheets(ABA)
    if ws.Range("C1").Formula != "=MEMORIA_RESULTADOS!$T$62":
        raise RuntimeError("CONTROLE!C1 fora da forma 12.2; nada aplicado")
    if ws.Range("C1").FormatConditions.Count != 2:
        raise RuntimeError("CONTROLE!C1 sem a orientacao do Ajuste 3; nada aplicado")
    if ws.Range("C3").Value != C3_ANTERIOR:
        raise RuntimeError("CONTROLE!C3 fora do texto 12.1; nada aplicado")
    for faixa in (AREA_NOVA, AREA_AUXILIAR, "A16:B17", "C4:C34", "D1:H3"):
        rng = ws.Range(faixa)
        if rng.MergeCells or any(c.Formula != "" for c in rng.Cells):
            raise RuntimeError(f"CONTROLE!{faixa} nao esta livre; nada aplicado")
    for endereco, rotulo in (("A1", "MODO DE LEITURA"), ("A3", "DATA DE CORTE DA APURAÇÃO")):
        if ws.Range(endereco).Value != rotulo:
            raise RuntimeError(f"CONTROLE!{endereco} fora do layout; nada aplicado")


# ------------------------------------------------------------- helpers COM
_COM_OCUPADO = -2147418111          # RPC_E_CALL_REJECTED: Excel ocupado


def _com_seguro(funcao):
    """Repete a chamada enquanto o Excel recusa por estar ocupado."""
    def chamada(*args, **kwargs):
        for _ in range(60):
            try:
                return funcao(*args, **kwargs)
            except Exception as exc:  # pragma: no cover - somente Excel COM
                if getattr(exc, "hresult", None) != _COM_OCUPADO:
                    raise
                time.sleep(0.3)
        return funcao(*args, **kwargs)
    return chamada


@_com_seguro
def _estilo(rng, *, fundo=None, cor=TEXTO, tamanho=10, negrito=False, italico=False,
            h=v3.XL_LEFT, quebra=True, borda=True) -> None:
    if fundo:
        v3._preencher(rng, fundo)
    rng.Font.Name = "Calibri"
    v3._fonte(rng, tamanho=tamanho, negrito=negrito, cor=cor, italico=italico)
    v3._alinhar(rng, h=h, quebra=quebra, recuo=1 if h == v3.XL_LEFT else 0)
    if borda:
        for indice in XL_EDGES + XL_INSIDE:
            try:
                v3._borda(rng, indice, BORDA)
            except Exception:
                pass                       # faixa de uma celula nao tem "inside"


@_com_seguro
def _bloco(ws, faixa: str, valor=None, **estilo):
    if ":" in faixa and not ws.Range(faixa).MergeCells:
        v3._mesclar(ws, faixa)
    rng = ws.Range(faixa)
    _estilo(rng, **estilo)
    if valor is not None:
        alvo = rng.Cells(1, 1)
        if isinstance(valor, str) and valor.startswith("="):
            alvo.Formula = valor
        else:
            alvo.Value = valor
    return rng


def _secao(ws, linha: int, texto: str) -> None:
    _bloco(ws, f"D{linha}:H{linha}", texto, fundo=AZUL_CLARO, cor=TEXTO_TITULO, negrito=True)


@_com_seguro
def _lista(ws, faixa: str, fonte: str, titulo: str, mensagem: str) -> None:
    rng = ws.Range(faixa)
    rng.Validation.Delete()
    rng.Validation.Add(XL_VALIDATE_LIST, XL_VALID_ALERT_STOP, XL_BETWEEN, fonte)
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
    rng.Validation.InputTitle = titulo
    rng.Validation.InputMessage = mensagem
    rng.Validation.ShowInput = True
    rng.Locked = False
    _estilo(rng, fundo=AMARELO, h=v3.XL_CENTER, quebra=False)


# ------------------------------------------------------------- frente COM
def frente_controle(wb) -> None:
    import aplicar_coleta_120_itens_pc_simnao_coerencia as c120
    import aplicar_coleta_122_acabamento_visual as av

    ws = wb.Worksheets(ABA)
    estado, selecao = c120._capturar_protecao(ws)

    # constantes (valores) nas colunas ocultas
    for linha, (chave, texto) in CONSTANTES.items():
        ws.Range(f"J{linha}").Value = chave
        celula = ws.Range(f"K{linha}")
        celula.NumberFormat = "@"
        celula.Value = texto

    # layout A:C e D4
    ws.Range("C3").ClearContents()
    for coluna, largura in LARGURAS.items():
        ws.Columns(coluna).ColumnWidth = largura
    for coluna in COLUNAS_OCULTAS:
        ws.Columns(coluna).Hidden = True
    for linha in LINHAS_REEXIBIDAS:
        ws.Rows(linha).Hidden = False
    for linha, altura in ALTURAS.items():
        ws.Rows(linha).RowHeight = altura
    for linha in LINHAS_OCULTAS:
        ws.Rows(linha).Hidden = True
    ws.Range("B4").WrapText = True
    for celula in ("A4", "B4"):
        ws.Range(celula).VerticalAlignment = v3.XL_TOP
    # C1 estreita: recuo para o alerta nao nascer sob o botao do dropdown de B1.
    ws.Range("C1").IndentLevel = 1
    _bloco(ws, "D4:H4", EXPLICACAO_CORTE, cor="595959", italico=True, borda=False)

    # Ajuste 3: a mesma regra de orientacao pinta C1:H1 (C1 agora e estreita).
    regras = ws.Range("C1").FormatConditions
    regra = next(regras(i) for i in range(1, regras.Count + 1)
                 if "SELECIONE AQUI" in str(regras(i).Formula1))
    regra.ModifyAppliesToRange(ws.Range("C1:H1"))

    # titulo e introducao
    _bloco(ws, "D5:H5", TITULO, fundo=AZUL_CLARO, cor=TEXTO_TITULO, tamanho=12, negrito=True)
    _bloco(ws, "D6:H6", INTRODUCAO, fundo=AZUL_SUAVE, cor=TEXTO_TITULO)

    # parametros complementares
    _secao(ws, 7, SECAO_PARAMETROS)
    _bloco(ws, "D8", ROTULO_HISTORICO, fundo=AZUL_SUAVE, negrito=True)
    _lista(ws, "E8", f'={K["HIST_SIM"]}:{K["HIST_NAO_CONFIRMADO"]}', "HISTÓRICO DA EXECUÇÃO",
           "Sim: a fonte do método cobre todo o período até a abertura do ciclo em execução.\n"
           "Parcialmente: há lacunas.\nNão confirmado: ainda não verificado.")
    _bloco(ws, "F8:H8", DICA_HISTORICO, cor="595959", italico=True)
    _bloco(ws, "D9", ROTULO_ADITIVOS, fundo=AZUL_SUAVE, negrito=True)
    _lista(ws, "E9", f'={K["ADT_NAO"]}:{K["ADT_NAO_SEI"]}', "ADITIVOS/SUPRESSÕES",
           "Houve acréscimos ou supressões de quantidade desde a abertura do ciclo em "
           "execução?")
    _bloco(ws, "F9:H9", None, fundo=CINZA, cor=TEXTO_CALCULO)

    # ciclos com pedido de reajuste
    _bloco(ws, "D10", SECAO_CICLOS, fundo=AZUL_CLARO, cor=TEXTO_TITULO, negrito=True)
    for coluna, ciclo in zip(COLUNAS_CICLOS, CICLOS):
        _bloco(ws, f"{coluna}10", ciclo, fundo=AZUL_CLARO, cor=TEXTO_TITULO, negrito=True,
               h=v3.XL_CENTER)
    for linha, rotulo in zip((11, 12, 13), ROTULOS_CICLOS):
        _bloco(ws, f"D{linha}", rotulo, fundo=AZUL_SUAVE)
    _estilo(ws.Range("E11:H11"), fundo=CINZA, cor=TEXTO_CALCULO, h=v3.XL_CENTER)
    _lista(ws, "E12:H12", f'={K["PED_SIM"]}:{K["PED_NAO"]}', "PEDIDO DE REAJUSTE",
           "Use só quando a leitura automática (linha acima) estiver vazia.")
    _estilo(ws.Range("E13:H13"), fundo=CINZA, cor=TEXTO_CALCULO, negrito=True, h=v3.XL_CENTER)

    # diagnostico
    _secao(ws, 15, SECAO_DIAGNOSTICO)
    for faixa, texto in zip(("D16", "E16", "F16:H16"), CABECALHO_DIAGNOSTICO):
        _bloco(ws, faixa, texto, fundo=AZUL_CLARO, cor=TEXTO, negrito=True, h=v3.XL_CENTER)
    for i, (info, _codigo, porque) in enumerate(DIAGNOSTICO):
        linha = PRIMEIRA_LINHA_DIAGNOSTICO + i
        _bloco(ws, f"D{linha}", None if info.startswith("=") else info, fundo=AZUL_SUAVE)
        _bloco(ws, f"E{linha}", None, fundo=CINZA, negrito=True, h=v3.XL_CENTER)
        _bloco(ws, f"F{linha}:H{linha}", porque, cor=TEXTO_CALCULO, tamanho=9)

    # o que voce vai pedir ao fiscal
    _bloco(ws, "D25:H25", SECAO_SAIDA, fundo=VERDE_FORTE, cor=TEXTO_VERDE, tamanho=11,
           negrito=True)
    _bloco(ws, "D26:H26", None, fundo=VERDE_CLARO, cor=TEXTO_VERDE, negrito=True, borda=False)
    for linha in LINHAS_SAIDA:
        _bloco(ws, f"D{linha}:H{linha}", None, fundo=VERDE_CLARO, cor=TEXTO, borda=False)
    contorno = ws.Range(f"D26:H{LINHAS_SAIDA[-1]}")
    for indice in XL_EDGES:
        v3._borda(contorno, indice, BORDA)
    _bloco(ws, "D34:H34", OBJETIVOS, cor="595959", tamanho=9, italico=True, borda=False)

    # formulas (depois das mesclas: ancora de cada bloco)
    for celula, formula in formulas().items():
        ws.Range(celula).Formula = formula

    # cores do STATUS por codigo oculto (formula sem funcoes: igual em qualquer idioma)
    ws.Activate()
    primeira = PRIMEIRA_LINHA_DIAGNOSTICO
    ws.Range(f"E{primeira}").Select()          # ancora a referencia relativa
    alvo = ws.Range(f"E{primeira}:E{primeira + len(DIAGNOSTICO) - 1}")
    for codigo, (fundo, texto) in CORES_STATUS.items():
        r = alvo.FormatConditions.Add(v3.XL_EXPRESSION, None, f"=$J{primeira}={codigo}")
        r.StopIfTrue = False
        r.Interior.Color = v3._bgr(fundo)
        r.Font.Color = v3._bgr(texto)
    ws.Range("B1").Select()
    wb.Application.ActiveWindow.DisplayGridlines = False
    c120._restaurar_protecao(ws, estado, selecao)
    if ws.Range("C1").Formula != av.FORMULA_C1:
        raise RuntimeError("CONTROLE!C1 alterada; nada gravado")


def conferir_resultado(wb) -> None:
    ws = wb.Worksheets(ABA)
    for celula, formula in formulas().items():
        if ws.Range(celula).Formula != formula:
            raise RuntimeError(f"CONTROLE!{celula} diverge do descritor: {ws.Range(celula).Formula}")
    for celula in ("E8", "E9", "E12", "F12", "G12", "H12", "B1", "B3"):
        if ws.Range(celula).Locked:
            raise RuntimeError(f"CONTROLE!{celula} deveria ser editavel")


def aplicar(caminho: Path) -> None:
    import pythoncom
    import win32com.client as com

    validar_ascii_e_parenteses()
    pythoncom.CoInitialize()
    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(caminho))
        excel.Calculation = v3.XL_CALCULO_MANUAL
        conferir(wb)
        frente_controle(wb)
        conferir_resultado(wb)
        excel.Calculation = v3.XL_CALCULO_AUTOMATICO
        excel.CalculateFullRebuild()
        wb.Worksheets(ABA).Activate()
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
    print("OK: orientador da coleta aplicado em %s" % caminho)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
