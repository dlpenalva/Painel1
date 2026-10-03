# -*- coding: utf-8 -*-
"""RESULTADOS-EXECUTIVO-V3: nova aba RESULTADOS executiva + RESULTADOS_DETALHE.

ARQUITETURA
  RESULTADOS_DETALHE = a aba RESULTADOS homologada ate o PR #172, RENOMEADA pelo
                       proprio Excel. Nenhuma formula, nome definido, validacao,
                       formatacao condicional ou entrada manual e movida: o
                       Excel reescreve sozinho os 33 nomes definidos e as
                       referencias de MEMORIA_RESULTADOS (ajustes manuais
                       C43:G50) e de comparativo_VTA (H5). Continua sendo a
                       camada tecnica e o UNICO lugar dos ajustes manuais.
  RESULTADOS         = aba nova, executiva, ultima do arquivo. So ESPELHA nomes
                       definidos e celulas canonicas (RESULTADOS_DETALHE,
                       MEMORIA_RESULTADOS, parametros, CONTROLE). Nao calcula
                       VTA, retroativo, remanescente nem qualquer grandeza
                       economica: nao e um segundo motor.

POR QUE RENOMEAR (E NAO RECONSTRUIR)
  A auditoria do template mostrou que o acoplamento XLS->XLS da aba antiga e
  so de entradas manuais (C43:G50, lidas por MEMORIA_RESULTADOS) e do fator
  historico H5 (lido por comparativo_VTA!B208); nao ha INDIRECT textual
  apontando para "RESULTADOS". Renomear pelo Excel preserva tudo isso por
  construcao. Os leitores Python por coordenada passam a resolver a aba
  tecnica por `_resultados_abas.aba_resultados_tecnica` (arquivo novo ->
  RESULTADOS_DETALHE; arquivo anterior -> RESULTADOS).

REGRA ZERO CORRUPCAO XLSX
  Somente Excel COM (openpyxl destroi a formatacao condicional x14); formulas
  ASCII, em ingles, com parenteses balanceados (verificado antes da escrita);
  textos acentuados ficam em CELULAS CONSTANTES de MEMORIA_RESULTADOS!AF:AG.
  Mesclagens: apenas faixas de uma linha, criadas pelo Excel, sem sobreposicao.
  A aba nova so usa recursos que sobrevivem ao load/save do openpyxl na geracao
  da Coleta (formatacao condicional classica, mesclagem, hiperlink interno).

Uso:  python tools/aplicar_resultados_executivo_v3.py [caminho.xlsx]
      (exige pywin32 + Excel real; sem argumento, altera o template oficial)
"""
from __future__ import annotations

import gc
import sys
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

EXE = "RESULTADOS"
DET = "RESULTADOS_DETALHE"
MEM = "MEMORIA_RESULTADOS"

TITULO_LEGADO = "RESULTADOS CONSOLIDADOS — REAJUSTE CONTRATUAL"
TITULO_EXECUTIVO = "RESULTADO DA APURAÇÃO"
TITULO_DETALHE = "RESULTADOS — DETALHE TÉCNICO DA APURAÇÃO"
SUBTITULO_DETALHE = (
    "Memória auditável, conferências e AJUSTES MANUAIS. "
    "A síntese executiva está na aba RESULTADOS."
)

# ------------------------------------------------------------- textos (AF:AG)
COL_CHAVE, COL_TEXTO = "AF", "AG"
LINHA_TITULO_TEXTOS = 1
TEXTOS = [
    ("TRACO", "—"),
    ("MET_FIN", "Financeiro (Mensalidade)"),
    ("MET_PC", "Pedidos de Compra"),
    ("MET_ITENS", "Itens Consumidos"),
    ("MET_NAO", "Método não selecionado"),
    ("VTA_SUB", "Valor Total Atualizado do contrato pelo método selecionado."),
    ("VTA_SUB_POT", "Inclui a parcela de retroativo POTENCIAL (card âmbar)."),
    ("RETRO_SUB", "Diferença reconhecida a pagar à contratada."),
    ("POT_SUB", "Incorporado ao VTA por prudência. Sujeito à confirmação da "
                "área gestora; não é valor a pagar."),
    ("POT_NA", "Não se aplica ao método selecionado."),
    ("SALDO_SUB", "Saldo atualizado que ainda falta executar."),
    ("IND_SUB", "Variação histórica integral até o ciclo vigente."),
    ("IND_SUB_NA", "Aguardando percentuais completos dos ciclos."),
    ("COMPUTADO", "Computado nesta apuração"),
    ("NAO_COMPUTADO", "Não computado nesta apuração"),
    ("SEM_AJUSTES", "SEM AJUSTES"),
    ("SEM_QUADRO3", "Não há ciclo reajustado computado nesta apuração."),
    ("HDR_PERIODO_SEM", "Período sem efeito e início dos efeitos"),
    ("DET_TITULO", "DETALHES DO MÉTODO — PEDIDOS DE COMPRA"),
    ("DET_MEDIDA", "Medida"),
    ("DET_VALOR", "Valor"),
    ("DET_REPRESENTA", "O que este valor representa"),
    ("ORIG_EXEC_PC", "Execução reconhecida em PCs, já atualizada: contém o "
                     "retroativo reconhecido incorporado."),
    ("ORIG_EXEC_FIN", "Valores efetivamente pagos, conforme a aba financeiro."),
    ("ORIG_EXEC_ITENS", "Consumo informado, atualizado pelo fator de cada ciclo."),
    ("ORIG_POT_PC", "Parcela POTENCIAL dos PCs em análise, incorporada por "
                    "prudência. Não é retroativo reconhecido nem valor a pagar."),
    ("ORIG_AJ_FIN", "Reajuste já reconhecido e ainda não contido no valor pago."),
    ("ORIG_AJ_ITENS", "Não aplicável: o reajuste já está dentro da execução "
                      "atualizada."),
    ("ORIG_REM_PC", "Saldo remanescente final atualizado usado no VTA (os "
                    "saldos históricos não se somam)."),
    ("ORIG_REM", "Saldo que ainda falta executar, já atualizado."),
    ("ORIG_VTA", "Resultado final do método de apuração selecionado."),
    ("LBL_VTA_SEM_POT", "VTA sem a parcela potencial (referência)"),
    ("ORIG_VTA_SEM_POT", "Referência: VTA antes de incorporar a parcela "
                         "potencial. Não substitui o VTA oficial."),
    ("LBL_POT_NEG", "Parcela potencial negativa apurada"),
    ("ORIG_POT_NEG", "Permanece visível por prudência: não reduz o VTA e não "
                     "compensa a parcela positiva."),
    ("ORIG_POT_NEG_ZERO", "Não há parcela potencial negativa apurada."),
    ("LBL_CONFERENCIA", "Conferência da formação (deve ser R$ 0,00)"),
    ("ORIG_AGUARDANDO", "Aguardando base para conferir."),
    ("ORIG_NA", "Não se aplica ao método selecionado."),
]
LINHA_TEXTO = {chave: LINHA_TITULO_TEXTOS + 1 + i for i, (chave, _) in enumerate(TEXTOS)}


def T(chave: str) -> str:
    """Referencia absoluta a um texto constante (a partir de outra aba)."""
    return f"{MEM}!${COL_TEXTO}${LINHA_TEXTO[chave]}"


def D(celula: str) -> str:
    """Referencia absoluta a uma celula de RESULTADOS_DETALHE."""
    col = "".join(ch for ch in celula if ch.isalpha())
    lin = "".join(ch for ch in celula if ch.isdigit())
    return f"{DET}!${col}${lin}"


def espelho(ref: str, vazio: str | None = None) -> str:
    """Espelho puro: vazio na origem -> traco (ou o texto informado)."""
    alternativa = vazio if vazio is not None else T("TRACO")
    return f'=IF({ref}="",{alternativa},{ref})'


PCS = 'METODO_RETROATIVO="PCs"'

# ------------------------------------------------------------- geometria
LARGURAS = {"A": 2.0, "B": 34.0, "C": 26.0, "D": 26.0, "E": 26.0, "F": 26.0,
            "G": 38.0, "H": 2.0}

L_TITULO, L_SUBTITULO = 2, 3
L_CTX_ROT, L_CTX_VAL = 5, 6
L_CARD_ROT, L_CARD_VAL, L_CARD_SUB = 8, 9, 10
L_COMP_SEC, L_COMP_CAB = 12, 13
L_COMP = {"exec": 14, "pot": 15, "rem": 16, "vta": 17, "sem_pot": 18,
          "pot_neg": 19, "conf": 20}
L_VER_SEC, L_VER_CAB = 22, 23
L_VER = {"comp": 24, "retro": 25, "rem": 26, "ciclo": 27, "data_ref": 28,
         "sit": 29, "dif": 30, "ajustes": 31}
L_Q1_SEC, L_Q1_CAB, L_Q1_C0 = 33, 34, 35          # 35..39
L_Q2_SEC, L_Q2_CAB, L_Q2_C0 = 41, 42, 43          # 43..47
L_Q2_AJUSTE, L_Q2_TOTAL = 48, 49
L_Q3_SEC, L_Q3_NOTA, L_Q3_CAB, L_Q3_C1, L_Q3_CASO = 51, 52, 53, 54, 58
# Rodape logo apos o quadro 3; "Detalhes do metodo" e a ULTIMA secao: fora de
# PCs ela fica vazia no fim da pagina, sem abrir um vao no meio da leitura.
L_RODAPE = 60
L_DM_SEC, L_DM_CAB, L_DM_1 = 62, 63, 64           # 64..71
ULTIMA_LINHA = L_DM_1 + 7

ALTURAS = {1: 6, L_TITULO: 30, L_SUBTITULO: 18, 4: 8, L_CTX_ROT: 16,
           L_CTX_VAL: 22, 7: 10, L_CARD_ROT: 18, L_CARD_VAL: 32,
           L_CARD_SUB: 44, 11: 12, 21: 12, 32: 12, 40: 12, 50: 12, 59: 8,
           61: 12, L_RODAPE: 30}

# ------------------------------------------------------------- cores (RGB)
TITULO_FUNDO, TITULO_TEXTO, SUBTITULO_TEXTO = "1F3864", "FFFFFF", "D9E1F2"
SECAO_FUNDO, SECAO_TEXTO = "1F4E78", "FFFFFF"
CAB_FUNDO, CAB_TEXTO = "DCE6F1", "1F3864"
ROTULO_TEXTO, CORPO_TEXTO, CINZA_TEXTO = "595959", "262626", "595959"
BORDA_LINHA = "D9D9D9"
CARDS = {
    # coluna: (fundo, texto, friso superior)
    "B": ("DEEAF6", "1F4E78", "2E75B6"),     # VTA oficial (azul institucional)
    "C": ("FBE5E8", "9B2C3F", "E6A1AD"),     # retroativo reconhecido (rosa)
    "D": ("FFF2CC", "7F6000", "FFC000"),     # potencial (ambar)
    "E": ("E2EFDA", "375623", "70AD47"),     # saldo remanescente
    "F": ("EDEDED", "404040", "A6A6A6"),     # indice acumulado
    "G": ("F7F7F7", "262626", "BFBFBF"),     # situacao/pendencias
}
STATUS_CORES = {
    "verde": ("C6EFCE", "006100"),
    "ambar": ("FFEB9C", "9C5700"),
    "rosa": ("FFC7CE", "9C0006"),
}

# ------------------------------------------------------------- formatos
MOEDA_INV = '"R$" #,##0.00;-"R$" #,##0.00;"R$" 0.00;@'
MOEDA_LOC = '"R$" #.##0,00;-"R$" #.##0,00;"R$" 0,00;@'
PCT_INV = "0.00%;-0.00%;0.00%;@"
PCT_LOC = "0,00%;-0,00%;0,00%;@"
DATA_INV = "dd/mm/yyyy;@"
DATA_LOC = "dd/mm/aaaa;@"

XL_LEFT, XL_RIGHT, XL_CENTER = -4131, -4152, -4108
XL_TOP, XL_VCENTER, XL_BOTTOM = -4160, -4108, -4107
XL_EXPRESSION = 2
XL_CONTINUOUS, XL_THIN, XL_MEDIUM, XL_THICK = 1, 2, -4138, 4
XL_EDGE_TOP, XL_EDGE_BOTTOM = 8, 9
XL_CF_BOTTOM = -4107                 # xlBottom (bordas de formatacao condicional)
XL_CALCULO_MANUAL, XL_CALCULO_AUTOMATICO = -4135, -4105
XL_LANDSCAPE = 2
XL_SHEET_VISIBLE = -1


def _bgr(rgb: str) -> int:
    r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
    return b * 65536 + g * 256 + r


# ------------------------------------------------------------- formulas
def formulas_executivo() -> dict[str, str]:
    """Todas as formulas da aba RESULTADOS executiva (endereco -> formula)."""
    f: dict[str, str] = {}
    traco = T("TRACO")

    # Contexto.
    f[f"B{L_CTX_VAL}"] = (
        f'=IF(METODO_RETROATIVO="Financeiro",{T("MET_FIN")},'
        f'IF({PCS},{T("MET_PC")},IF(METODO_RETROATIVO="Itens",{T("MET_ITENS")},'
        f'{T("MET_NAO")})))'
    )
    f[f"C{L_CTX_VAL}"] = f'=IF(CONTROLE!$B$2="",{traco},UPPER(CONTROLE!$B$2))'
    f[f"D{L_CTX_VAL}"] = f'=IF(ISNUMBER(CONTROLE!$B$3),CONTROLE!$B$3,{traco})'
    f[f"E{L_CTX_VAL}"] = f'=IF(CONTROLE!$B$7="",{traco},CONTROLE!$B$7)'
    f[f"F{L_CTX_VAL}"] = f'=IF(ISNUMBER({D("D6")}),{D("D6")},{traco})'
    f[f"G{L_CTX_VAL}"] = f'=IF(STATUS_RESULTADOS="",{traco},STATUS_RESULTADOS)'

    # Cards (espelhos das MESMAS fontes dos cards da aba anterior).
    f[f"B{L_CARD_VAL}"] = f"=IF(ISNUMBER(VTA_FINAL),VTA_FINAL,{traco})"
    f[f"C{L_CARD_VAL}"] = f'=IF(ISNUMBER({D("D22")}),{D("D22")},{traco})'
    f[f"D{L_CARD_VAL}"] = (
        f"=IF(NOT({PCS}),{traco},IF(ISNUMBER(RETROATIVO_POTENCIAL_VTA),"
        f"RETROATIVO_POTENCIAL_VTA,{traco}))"
    )
    f[f"E{L_CARD_VAL}"] = (
        f"=IF(ISNUMBER(SALDO_REMANESCENTE_ATUAL),SALDO_REMANESCENTE_ATUAL,{traco})"
    )
    f[f"F{L_CARD_VAL}"] = f"=F{L_CTX_VAL}"
    f[f"G{L_CARD_VAL}"] = f'={D("A7")}'
    f[f"B{L_CARD_SUB}"] = (
        f'=IF(AND({PCS},N(RETROATIVO_POTENCIAL_VTA)<>0),{T("VTA_SUB_POT")},'
        f'{T("VTA_SUB")})'
    )
    f[f"C{L_CARD_SUB}"] = f'={T("RETRO_SUB")}'
    f[f"D{L_CARD_SUB}"] = f'=IF({PCS},{T("POT_SUB")},{T("POT_NA")})'
    f[f"E{L_CARD_SUB}"] = f'={T("SALDO_SUB")}'
    f[f"F{L_CARD_SUB}"] = (
        f'=IF(ISNUMBER(F{L_CARD_VAL}),{T("IND_SUB")},{T("IND_SUB_NA")})'
    )
    f[f"G{L_CARD_SUB}"] = f'={D("B7")}'

    # Composicao do VTA (rotulos e valores espelhados da tabela 9 do detalhe).
    r = L_COMP
    f[f"B{r['exec']}"] = f'={D("A83")}'
    f[f"C{r['exec']}"] = espelho(D("B83"))
    f[f"D{r['exec']}"] = (
        f'=IF({PCS},{T("ORIG_EXEC_PC")},IF(METODO_RETROATIVO="Itens",'
        f'{T("ORIG_EXEC_ITENS")},{T("ORIG_EXEC_FIN")}))'
    )
    f[f"B{r['pot']}"] = f'={D("A84")}'
    f[f"C{r['pot']}"] = espelho(D("B84"))
    f[f"D{r['pot']}"] = (
        f'=IF({PCS},{T("ORIG_POT_PC")},IF(METODO_RETROATIVO="Itens",'
        f'{T("ORIG_AJ_ITENS")},{T("ORIG_AJ_FIN")}))'
    )
    f[f"B{r['rem']}"] = f'={D("A85")}'
    f[f"C{r['rem']}"] = espelho(D("B85"))
    f[f"D{r['rem']}"] = f'=IF({PCS},{T("ORIG_REM_PC")},{T("ORIG_REM")})'
    f[f"B{r['vta']}"] = f'={D("A86")}'
    f[f"C{r['vta']}"] = espelho(D("B86"))
    f[f"D{r['vta']}"] = f'={T("ORIG_VTA")}'
    f[f"B{r['sem_pot']}"] = f'={T("LBL_VTA_SEM_POT")}'
    f[f"C{r['sem_pot']}"] = (
        f"=IF(NOT({PCS}),{traco},IF(ISNUMBER(VTA_SEM_POTENCIAL),"
        f"VTA_SEM_POTENCIAL,{traco}))"
    )
    f[f"D{r['sem_pot']}"] = f'=IF({PCS},{T("ORIG_VTA_SEM_POT")},{T("ORIG_NA")})'
    f[f"B{r['pot_neg']}"] = f'={T("LBL_POT_NEG")}'
    f[f"C{r['pot_neg']}"] = (
        f"=IF(NOT({PCS}),{traco},IF(N(RETROATIVO_POTENCIAL_NEGATIVO)<0,"
        f"RETROATIVO_POTENCIAL_NEGATIVO,0))"
    )
    f[f"D{r['pot_neg']}"] = (
        f'=IF(NOT({PCS}),{T("ORIG_NA")},IF(N(RETROATIVO_POTENCIAL_NEGATIVO)<0,'
        f'{T("ORIG_POT_NEG")},{T("ORIG_POT_NEG_ZERO")}))'
    )
    f[f"B{r['conf']}"] = f'={T("LBL_CONFERENCIA")}'
    f[f"C{r['conf']}"] = espelho("CONFERENCIA_FORMACAO_VTA")
    f[f"D{r['conf']}"] = espelho(D("C87"), T("ORIG_AGUARDANDO"))

    # Verificacoes (status canonicos ja existentes; nenhum classificador novo).
    v = L_VER
    f[f"C{v['comp']}"] = espelho(D("H8"))
    f[f"C{v['retro']}"] = espelho(D("H14"))
    f[f"C{v['rem']}"] = espelho(D("H24"))
    f[f"C{v['ciclo']}"] = espelho(D("H33"))
    f[f"C{v['data_ref']}"] = f'=IF(ISNUMBER({D("B35")}),{D("B35")},{traco})'
    f[f"C{v['sit']}"] = espelho(D("H13"))
    f[f"C{v['dif']}"] = espelho(D("B13"))
    # Mesmo predicado do STATUS_RESULTADOS (COUNTIF REVISE em H43:H50).
    f[f"C{v['ajustes']}"] = (
        f'=IF(COUNTA({DET}!$C$43:$G$50)=0,{T("SEM_AJUSTES")},'
        f'IF(COUNTIF({DET}!$H$43:$H$50,"REVISE")>0,"REVISE","VALIDADO"))'
    )

    # Quadro 1 — evolucao por ciclo (parametros, linhas 2..6).
    for n in range(5):
        lin, p = L_Q1_C0 + n, 2 + n
        f[f"B{lin}"] = f'="C{n}"&IF(UPPER(CONTROLE!$B$2)="C{n}"," (vigente)","")'
        f[f"C{lin}"] = f"=IF(ISNUMBER(parametros!$C${p}),parametros!$C${p},{traco})"
        f[f"D{lin}"] = f"=IF(ISNUMBER(parametros!$D${p}),parametros!$D${p},{traco})"
        f[f"E{lin}"] = f"=IF(ISNUMBER(parametros!$E${p}),parametros!$E${p},{traco})"
        f[f"F{lin}"] = (
            f'=IF(ISNUMBER(parametros!$H${p}),parametros!$H${p},'
            f'IF(parametros!$H${p}="",{traco},parametros!$H${p}))'
        )
        f[f"G{lin}"] = (
            f'=IF(parametros!$G${p}<>"",parametros!$G${p},'
            f'IF(parametros!$A${p}="Sim",{T("COMPUTADO")},'
            f'IF(ISNUMBER(parametros!$C${p}),{T("NAO_COMPUTADO")},{traco})))'
        )

    # Quadro 2 — espelho da tabela 2 do detalhe (linhas 15..22).
    f[f"B{L_Q2_SEC}"] = f'={D("A15")}'
    f[f"C{L_Q2_CAB}"] = f'={D("B15")}'
    f[f"D{L_Q2_CAB}"] = f'={D("C15")}'
    f[f"E{L_Q2_CAB}"] = f'={D("D15")}'
    for n in range(5):
        lin, origem = L_Q2_C0 + n, 16 + n
        for col_exe, col_det in (("C", "B"), ("D", "C"), ("E", "D")):
            f[f"{col_exe}{lin}"] = espelho(D(f"{col_det}{origem}"))
    f[f"E{L_Q2_AJUSTE}"] = espelho(D("D21"))
    for col_exe, col_det in (("C", "B"), ("D", "C"), ("E", "D")):
        f[f"{col_exe}{L_Q2_TOTAL}"] = espelho(D(f"{col_det}22"))

    # Quadro 3 — espelho do quadro do PR #172 (detalhe E15:H21).
    f[f"B{L_Q3_NOTA}"] = f'=IF({D("E15")}="",{T("SEM_QUADRO3")},{D("E16")})'
    f[f"B{L_Q3_CAB}"] = f'=IF({D("F15")}="","","Ciclo")'
    f[f"C{L_Q3_CAB}"] = f'={D("F15")}&""'
    f[f"D{L_Q3_CAB}"] = f'={D("G15")}&""'
    f[f"E{L_Q3_CAB}"] = f'={D("H15")}&""'
    f[f"F{L_Q3_CAB}"] = f'=IF({D("F15")}="","",{T("HDR_PERIODO_SEM")})'
    for n in range(1, 5):
        lin, origem = L_Q3_C1 + n - 1, 16 + n
        f[f"B{lin}"] = f'=IF({D(f"E{origem}")}="","","C{n}")'
        for col_exe, col_det in (("C", "F"), ("D", "G"), ("E", "H")):
            ref = D(f"{col_det}{origem}")
            f[f"{col_exe}{lin}"] = f'=IF({ref}="","",{ref})'
        f[f"F{lin}"] = f'={D(f"E{origem}")}&""'
    f[f"B{L_Q3_CASO}"] = f'={D("E21")}&""'

    # Detalhes do metodo (somente PCs; espelho da tabela 6 do detalhe).
    f[f"B{L_DM_SEC}"] = f'=IF({PCS},{T("DET_TITULO")},"")'
    f[f"B{L_DM_CAB}"] = f'=IF({PCS},{T("DET_MEDIDA")},"")'
    f[f"C{L_DM_CAB}"] = f'=IF({PCS},{T("DET_VALOR")},"")'
    f[f"D{L_DM_CAB}"] = f'=IF({PCS},{T("DET_REPRESENTA")},"")'
    for i in range(8):
        lin, origem = L_DM_1 + i, 55 + i                  # tabela 6: linhas 55..62
        f[f"B{lin}"] = f'=IF({PCS},{D(f"A{origem}")},"")'
        f[f"C{lin}"] = (
            f'=IF({PCS},IF({D(f"B{origem}")}="",{traco},{D(f"B{origem}")}),"")'
        )
        f[f"D{lin}"] = f'=IF({PCS},{D(f"C{origem}")},"")'
    return f


CONSTANTES_EXECUTIVO = {
    f"B{L_TITULO}": TITULO_EXECUTIVO,
    f"B{L_SUBTITULO}": (
        "Síntese executiva da apuração do reajuste. A memória técnica, as "
        "conferências e os ajustes manuais estão na aba RESULTADOS_DETALHE."
    ),
    f"B{L_CTX_ROT}": "MÉTODO",
    f"C{L_CTX_ROT}": "CICLO VIGENTE",
    f"D{L_CTX_ROT}": "DATA DE CORTE",
    f"E{L_CTX_ROT}": "ÍNDICE",
    f"F{L_CTX_ROT}": "ÍNDICE ACUMULADO",
    f"G{L_CTX_ROT}": "STATUS DA APURAÇÃO",
    f"B{L_CARD_ROT}": "VTA OFICIAL",
    f"C{L_CARD_ROT}": "RETROATIVO RECONHECIDO",
    f"D{L_CARD_ROT}": "RETROATIVO POTENCIAL",
    f"E{L_CARD_ROT}": "SALDO REMANESCENTE",
    f"F{L_CARD_ROT}": "ÍNDICE ACUMULADO",
    f"G{L_CARD_ROT}": "SITUAÇÃO DA APURAÇÃO",
    f"B{L_COMP_SEC}": "COMPOSIÇÃO DO VTA",
    f"B{L_COMP_CAB}": "Parcela",
    f"C{L_COMP_CAB}": "Valor",
    f"D{L_COMP_CAB}": "Origem",
    f"B{L_VER_SEC}": "VERIFICAÇÕES",
    f"B{L_VER_CAB}": "Verificação",
    f"C{L_VER_CAB}": "Situação",
    f"D{L_VER_CAB}": "O que é verificado",
    f"B{L_VER['comp']}": "Composição do VTA",
    f"D{L_VER['comp']}": "VTA oficial calculado e reconhecido pelo método selecionado.",
    f"B{L_VER['retro']}": "Retroativo por ciclo",
    f"D{L_VER['retro']}": "Retroativo apurado em todos os ciclos computados nesta apuração.",
    f"B{L_VER['rem']}": "Remanescente por ciclo",
    f"D{L_VER['rem']}": "Remanescente conferido entre as bases disponíveis.",
    f"B{L_VER['ciclo']}": "Ciclo em execução",
    f"D{L_VER['ciclo']}": "Execução e saldo do ciclo vigente na data de referência.",
    f"B{L_VER['data_ref']}": "Data de referência do ciclo em execução",
    f"D{L_VER['data_ref']}": (
        "Data da posição usada para medir a execução e o saldo do ciclo atual."
    ),
    f"B{L_VER['sit']}": "Situação atual × referência anterior",
    f"D{L_VER['sit']}": (
        "Reconciliação entre a posição atual e a última abertura de ciclo "
        "(não é VTA)."
    ),
    f"B{L_VER['dif']}": "Diferença entre posição atual e última abertura",
    f"D{L_VER['dif']}": (
        "R$ 0,00 indica que as duas decomposições chegaram ao mesmo valor."
    ),
    f"B{L_VER['ajustes']}": "Ajustes manuais da apuração",
    f"B{L_Q1_SEC}": "1. EVOLUÇÃO POR CICLO",
    f"B{L_Q1_CAB}": "Ciclo",
    f"C{L_Q1_CAB}": "Início do ciclo",
    f"D{L_Q1_CAB}": "Fim do ciclo",
    f"E{L_Q1_CAB}": "Variação do ciclo",
    f"F{L_Q1_CAB}": "Início do efeito financeiro",
    f"G{L_Q1_CAB}": "Situação",
    f"B{L_Q2_CAB}": "Ciclo",
    f"B{L_Q2_AJUSTE}": "Ajuste manual não rateado",
    f"B{L_Q2_TOTAL}": "TOTAL",
    f"B{L_Q3_SEC}": "3. EXECUÇÃO SEM EFEITO FINANCEIRO",
    f"B{L_RODAPE}": (
        "Esta aba apenas apresenta valores. Todos os números são espelhos das "
        "fontes canônicas (RESULTADOS_DETALHE, MEMORIA_RESULTADOS, parametros e "
        "CONTROLE); nenhum valor é recalculado aqui. A execução sem efeito "
        "financeiro é exclusivamente informativa: não é retroativo, não é valor "
        "a pagar e não integra o VTA."
    ),
}
for _n in range(5):
    CONSTANTES_EXECUTIVO[f"B{L_Q2_C0 + _n}"] = f"C{_n}"

TEXTO_LINK_AJUSTES = (
    "Editar na aba RESULTADOS_DETALHE — bloco 5. AJUSTES MANUAIS (clique aqui)"
)
DESTINO_LINK_AJUSTES = f"'{DET}'!A41"


def validar_ascii_e_parenteses(formulas: dict[str, str]) -> None:
    for endereco, formula in formulas.items():
        if any(ord(ch) > 127 for ch in formula):
            raise ValueError(f"Formula nao-ASCII em {endereco}: {formula!r}")
        nivel = 0
        aspas = False
        for ch in formula:
            if ch == '"':
                aspas = not aspas
            elif not aspas and ch == "(":
                nivel += 1
            elif not aspas and ch == ")":
                nivel -= 1
                if nivel < 0:
                    break
        if nivel != 0 or aspas:
            raise ValueError(f"Parenteses/aspas desbalanceados em {endereco}")


# ------------------------------------------------------------- COM helpers
def _fechar(wb, tentativas: int = 10) -> None:
    for _ in range(tentativas):
        try:
            wb.Close(SaveChanges=False)
            return
        except Exception:
            time.sleep(1.0)


def _formato(rng, local: str, invariante: str) -> None:
    try:
        rng.NumberFormatLocal = local
    except Exception:
        rng.NumberFormat = invariante


def _preencher(rng, rgb: str) -> None:
    rng.Interior.Pattern = 1
    rng.Interior.Color = _bgr(rgb)


def _fonte(rng, *, tamanho=None, negrito=None, cor=None, italico=None) -> None:
    if tamanho is not None:
        rng.Font.Size = tamanho
    if negrito is not None:
        rng.Font.Bold = negrito
    if italico is not None:
        rng.Font.Italic = italico
    if cor is not None:
        rng.Font.Color = _bgr(cor)


def _alinhar(rng, h=XL_LEFT, v=XL_VCENTER, quebra=False, recuo=0) -> None:
    rng.HorizontalAlignment = h
    rng.VerticalAlignment = v
    rng.WrapText = quebra
    if recuo:
        rng.IndentLevel = recuo


def _borda(rng, indice: int, rgb: str, peso=XL_THIN) -> None:
    b = rng.Borders(indice)
    b.LineStyle = XL_CONTINUOUS
    b.Weight = peso
    b.Color = _bgr(rgb)


def _mesclar(ws, endereco: str) -> None:
    rng = ws.Range(endereco)
    rng.UnMerge()
    rng.Merge()


def _local(rng, formula: str) -> str:
    """Formula de formatacao condicional no idioma da interface do Excel.

    `FormatConditions.Add` interpreta a expressao no idioma LOCAL (no Excel
    pt-BR: OU, ENUM, ';'). O proprio Excel traduz: grava-se a formula em
    ingles numa celula auxiliar da MESMA LINHA da ancora (coluna BH, fora da
    pagina) e le-se `FormulaLocal`. Todas as regras desta aba usam coluna
    absoluta, entao a traducao preserva as referencias de linha relativas.
    """
    ws = rng.Worksheet
    aux = ws.Cells(rng.Cells(1, 1).Row, 60)
    aux.Formula = formula
    local = aux.FormulaLocal
    aux.ClearContents()
    return local


def _status_cf(rng, primeira: str) -> None:
    """Cores semanticas por TEXTO do status (sem reclassificar nada)."""
    regras = (
        (f'=OR({primeira}="VALIDADO",{primeira}="RECONCILIADO")', "verde"),
        (f'=ISNUMBER(SEARCH("ESTIMAD",{primeira}))', "ambar"),
        (f'=ISNUMBER(SEARCH("REVISE",{primeira}))', "rosa"),
    )
    for formula, cor in regras:
        fundo, texto = STATUS_CORES[cor]
        regra = rng.FormatConditions.Add(XL_EXPRESSION, None, _local(rng, formula))
        regra.Interior.Color = _bgr(fundo)
        regra.Font.Color = _bgr(texto)


# ------------------------------------------------------------- etapas
def _renomear_para_detalhe(wb) -> None:
    nomes = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
    if DET in nomes:
        # Reaplicacao: a aba executiva e reconstruida do zero.
        if EXE in nomes:
            wb.Worksheets(EXE).Delete()
        return
    if EXE not in nomes:
        raise RuntimeError("Template sem a aba RESULTADOS")
    det = wb.Worksheets(EXE)
    if str(det.Range("A1").Value or "") != TITULO_LEGADO:
        raise RuntimeError("Aba RESULTADOS inesperada: titulo A1 diferente do homologado")
    det.Name = DET


def _ajustar_titulos_detalhe(wb) -> None:
    det = wb.Worksheets(DET)
    det.Range("A1").Value = TITULO_DETALHE
    det.Range("A2").Value = SUBTITULO_DETALHE


def _exigir_area_textos_livre(ws) -> None:
    ultima = LINHA_TITULO_TEXTOS + len(TEXTOS) + 5
    chaves = {chave for chave, _ in TEXTOS}
    for linha in range(1, ultima + 1):
        for col in (COL_CHAVE, COL_TEXTO, "AH"):
            valor = ws.Range(f"{col}{linha}").Formula
            if valor in (None, ""):
                continue
            if linha == LINHA_TITULO_TEXTOS:
                continue
            if col == COL_CHAVE and valor in chaves:
                continue
            if col == COL_TEXTO:
                continue
            raise RuntimeError(f"{MEM}!{col}{linha} ja possui conteudo: {valor!r}")


def _aplicar_textos(wb) -> None:
    ws = wb.Worksheets(MEM)
    _exigir_area_textos_livre(ws)
    ws.Range(f"{COL_CHAVE}{LINHA_TITULO_TEXTOS}").Value = (
        "RESULTADOS EXECUTIVA — textos constantes (as fórmulas da aba são ASCII)"
    )
    for chave, texto in TEXTOS:
        linha = LINHA_TEXTO[chave]
        ws.Range(f"{COL_CHAVE}{linha}").Value = chave
        ws.Range(f"{COL_TEXTO}{linha}").NumberFormat = "@"
        ws.Range(f"{COL_TEXTO}{linha}").Value = texto


def _criar_aba_executiva(wb):
    n = wb.Worksheets.Count
    ws = wb.Worksheets.Add(None, wb.Worksheets(n))
    ws.Name = EXE
    ws.Visible = XL_SHEET_VISIBLE
    ws.Tab.Color = wb.Worksheets(DET).Tab.Color
    return ws


def _geometria(ws) -> None:
    ws.Cells.Font.Size = 10
    ws.Cells.VerticalAlignment = XL_VCENTER
    for col, largura in LARGURAS.items():
        ws.Columns(f"{col}:{col}").ColumnWidth = largura
    ws.Rows(f"1:{ULTIMA_LINHA + 2}").RowHeight = 20
    for linha in list(L_COMP.values()) + list(L_VER.values()):
        ws.Rows(f"{linha}:{linha}").RowHeight = 30
    for linha in range(L_Q1_C0, L_Q1_C0 + 5):
        ws.Rows(f"{linha}:{linha}").RowHeight = 30
    for linha in range(L_Q3_C1, L_Q3_C1 + 4):
        ws.Rows(f"{linha}:{linha}").RowHeight = 32
    for linha in (L_Q3_NOTA, L_Q3_CASO):
        ws.Rows(f"{linha}:{linha}").RowHeight = 30
    for linha in range(L_DM_1, L_DM_1 + 8):
        ws.Rows(f"{linha}:{linha}").RowHeight = 30
    for linha in (L_COMP_CAB, L_VER_CAB, L_Q1_CAB, L_Q2_CAB, L_Q3_CAB, L_DM_CAB):
        ws.Rows(f"{linha}:{linha}").RowHeight = 22
    for linha in (L_COMP_SEC, L_VER_SEC, L_Q1_SEC, L_Q2_SEC, L_Q3_SEC, L_DM_SEC):
        ws.Rows(f"{linha}:{linha}").RowHeight = 24
    for linha, altura in ALTURAS.items():
        ws.Rows(f"{linha}:{linha}").RowHeight = altura


def _secao(ws, linha: int) -> None:
    faixa = ws.Range(f"B{linha}:G{linha}")
    _preencher(faixa, SECAO_FUNDO)
    _fonte(faixa, tamanho=11, negrito=True, cor=SECAO_TEXTO)
    _alinhar(faixa, recuo=1)


def _cabecalho(ws, linha: int, ate: str = "G") -> None:
    faixa = ws.Range(f"B{linha}:{ate}{linha}")
    _preencher(faixa, CAB_FUNDO)
    _fonte(faixa, tamanho=9.5, negrito=True, cor=CAB_TEXTO)
    _alinhar(faixa, quebra=True)
    _borda(faixa, XL_EDGE_BOTTOM, "9DB3D3")
    ws.Range(f"B{linha}").IndentLevel = 1


def _linhas_dados(ws, primeira: int, ultima: int, ate: str = "G") -> None:
    faixa = ws.Range(f"B{primeira}:{ate}{ultima}")
    _fonte(faixa, tamanho=10, cor=CORPO_TEXTO)
    for linha in range(primeira, ultima + 1):
        _borda(ws.Range(f"B{linha}:{ate}{linha}"), XL_EDGE_BOTTOM, BORDA_LINHA)
    ws.Range(f"B{primeira}:B{ultima}").IndentLevel = 1


def _moeda(rng) -> None:
    _formato(rng, MOEDA_LOC, MOEDA_INV)
    rng.HorizontalAlignment = XL_RIGHT


def _cf_cor(rng, formula: str, fundo: int | None, texto: int | None = None):
    regra = rng.FormatConditions.Add(XL_EXPRESSION, None, _local(rng, formula))
    if fundo is not None:
        regra.Interior.Color = fundo
    if texto is not None:
        regra.Font.Color = texto
    return regra


def _layout(ws, cor_pot_fundo: int, cor_pot_texto: int) -> None:
    _geometria(ws)
    rosa_fundo = _bgr(STATUS_CORES["rosa"][0])
    rosa_texto = _bgr(STATUS_CORES["rosa"][1])

    # Titulo.
    _preencher(ws.Range(f"B{L_TITULO}:G{L_SUBTITULO}"), TITULO_FUNDO)
    _fonte(ws.Range(f"B{L_TITULO}"), tamanho=18, negrito=True, cor=TITULO_TEXTO)
    _fonte(ws.Range(f"B{L_SUBTITULO}"), tamanho=9.5, cor=SUBTITULO_TEXTO)
    _alinhar(ws.Range(f"B{L_TITULO}:B{L_SUBTITULO}"), recuo=1)

    # Contexto.
    rot = ws.Range(f"B{L_CTX_ROT}:G{L_CTX_ROT}")
    _fonte(rot, tamanho=8.5, negrito=True, cor=ROTULO_TEXTO)
    _alinhar(rot, v=XL_BOTTOM)
    val = ws.Range(f"B{L_CTX_VAL}:G{L_CTX_VAL}")
    _fonte(val, tamanho=11, negrito=True, cor=TITULO_FUNDO)
    _alinhar(val)
    _borda(val, XL_EDGE_BOTTOM, "BFBFBF")
    _formato(ws.Range(f"D{L_CTX_VAL}"), DATA_LOC, DATA_INV)
    _formato(ws.Range(f"F{L_CTX_VAL}"), PCT_LOC, PCT_INV)
    ws.Range(f"B{L_CTX_ROT}:B{L_CTX_VAL}").IndentLevel = 1
    _alinhar(ws.Range(f"G{L_CTX_ROT}:G{L_CTX_VAL}"), h=XL_CENTER,
             v=XL_VCENTER)
    ws.Range(f"G{L_CTX_ROT}").VerticalAlignment = XL_BOTTOM
    _status_cf(ws.Range(f"G{L_CTX_VAL}"), f"$G${L_CTX_VAL}")

    # Cards.
    for col, (fundo, texto, friso) in CARDS.items():
        _preencher(ws.Range(f"{col}{L_CARD_ROT}:{col}{L_CARD_SUB}"), fundo)
        _borda(ws.Range(f"{col}{L_CARD_ROT}"), XL_EDGE_TOP, friso, XL_THICK)
        _fonte(ws.Range(f"{col}{L_CARD_ROT}"), tamanho=9, negrito=True, cor=texto)
        _alinhar(ws.Range(f"{col}{L_CARD_ROT}"), recuo=1)
        _fonte(ws.Range(f"{col}{L_CARD_VAL}"), tamanho=15, negrito=True, cor=texto)
        _fonte(ws.Range(f"{col}{L_CARD_SUB}"), tamanho=9, cor=texto)
        _alinhar(ws.Range(f"{col}{L_CARD_SUB}"), v=XL_TOP, quebra=True, recuo=1)
    for col in "BCDE":
        _moeda(ws.Range(f"{col}{L_CARD_VAL}"))
        ws.Range(f"{col}{L_CARD_VAL}").IndentLevel = 1
    _formato(ws.Range(f"F{L_CARD_VAL}"), PCT_LOC, PCT_INV)
    ws.Range(f"F{L_CARD_VAL}").HorizontalAlignment = XL_RIGHT
    ws.Range(f"F{L_CARD_VAL}").IndentLevel = 1
    g = ws.Range(f"G{L_CARD_VAL}")
    _fonte(g, tamanho=11, negrito=True)
    _alinhar(g, quebra=True, recuo=1)
    regra = _cf_cor(ws.Range(f"B{L_CARD_SUB}"),
                    f"=AND({PCS},N(RETROATIVO_POTENCIAL_VTA)<>0)",
                    cor_pot_fundo, cor_pot_texto)
    regra.Font.Bold = True
    # Situacao: rosa suave quando ha pendencia (mesmo contador de A7/B7).
    _cf_cor(ws.Range(f"G{L_CARD_ROT}:G{L_CARD_SUB}"), f"={DET}!$J$5>0",
            rosa_fundo, rosa_texto)
    _cf_cor(ws.Range(f"G{L_CARD_ROT}:G{L_CARD_SUB}"), f"={DET}!$J$5=0",
            _bgr(STATUS_CORES["verde"][0]), _bgr(STATUS_CORES["verde"][1]))

    # Composicao do VTA.
    _secao(ws, L_COMP_SEC)
    _cabecalho(ws, L_COMP_CAB)
    primeira, ultima = min(L_COMP.values()), max(L_COMP.values())
    _linhas_dados(ws, primeira, ultima)
    for linha in range(L_COMP_CAB, ultima + 1):
        _mesclar(ws, f"D{linha}:G{linha}")
    _alinhar(ws.Range(f"B{primeira}:B{ultima}"), quebra=True, recuo=1)
    _moeda(ws.Range(f"C{primeira}:C{ultima}"))
    _alinhar(ws.Range(f"D{primeira}:G{ultima}"), quebra=True, recuo=1)
    _fonte(ws.Range(f"D{primeira}:G{ultima}"), tamanho=9, cor=CINZA_TEXTO)
    vta = ws.Range(f"B{L_COMP['vta']}:G{L_COMP['vta']}")
    _preencher(vta, CARDS["B"][0])
    _fonte(ws.Range(f"B{L_COMP['vta']}:C{L_COMP['vta']}"), negrito=True,
           cor=TITULO_FUNDO)
    _borda(vta, XL_EDGE_TOP, CARDS["B"][2], XL_MEDIUM)
    _fonte(ws.Range(f"B{L_COMP['sem_pot']}:C{L_COMP['sem_pot']}"), italico=True,
           cor=CINZA_TEXTO)
    _cf_cor(ws.Range(f"B{L_COMP['pot']}:G{L_COMP['pot']}"),
            f"=AND({PCS},N(RETROATIVO_POTENCIAL_APURADO)<>0)",
            cor_pot_fundo, cor_pot_texto)
    _cf_cor(ws.Range(f"B{L_COMP['pot_neg']}:G{L_COMP['pot_neg']}"),
            f"=AND({PCS},N(RETROATIVO_POTENCIAL_NEGATIVO)<0)",
            cor_pot_fundo, cor_pot_texto)
    c = f"$C${L_COMP['conf']}"
    _cf_cor(ws.Range(f"C{L_COMP['conf']}"), f"=AND(ISNUMBER({c}),{c}<>0)",
            rosa_fundo, rosa_texto)

    # Verificacoes.
    _secao(ws, L_VER_SEC)
    _cabecalho(ws, L_VER_CAB)
    primeira, ultima = min(L_VER.values()), max(L_VER.values())
    _linhas_dados(ws, primeira, ultima)
    for linha in range(L_VER_CAB, ultima + 1):
        _mesclar(ws, f"D{linha}:G{linha}")
    _alinhar(ws.Range(f"B{primeira}:B{ultima}"), quebra=True, recuo=1)
    _alinhar(ws.Range(f"C{L_VER_CAB}:C{ultima}"), h=XL_CENTER, quebra=True)
    _fonte(ws.Range(f"C{primeira}:C{ultima}"), tamanho=9, negrito=True)
    _alinhar(ws.Range(f"D{primeira}:G{ultima}"), quebra=True, recuo=1)
    _fonte(ws.Range(f"D{primeira}:G{ultima}"), tamanho=9, cor=CINZA_TEXTO)
    _formato(ws.Range(f"C{L_VER['data_ref']}"), DATA_LOC, DATA_INV)
    _moeda(ws.Range(f"C{L_VER['dif']}"))
    ws.Range(f"C{L_VER['dif']}").HorizontalAlignment = XL_CENTER
    ws.Range(f"C{L_VER['data_ref']}:C{L_VER['dif']}").Font.Bold = False
    ws.Range(f"C{L_VER['data_ref']}").Font.Bold = False
    for chave in ("comp", "retro", "rem", "ciclo", "sit", "ajustes"):
        _status_cf(ws.Range(f"C{L_VER[chave]}"), f"$C${L_VER[chave]}")
    ws.Range(f"C{L_VER['sit']}").Font.Bold = True

    # Link para os ajustes manuais (unico lugar de entrada manual).
    ancora = ws.Range(f"D{L_VER['ajustes']}")
    ws.Hyperlinks.Add(ancora, "", DESTINO_LINK_AJUSTES, "", TEXTO_LINK_AJUSTES)
    _fonte(ancora, tamanho=9, negrito=True, cor="0563C1")
    ancora.Font.Underline = 2
    _alinhar(ws.Range(f"D{L_VER['ajustes']}:G{L_VER['ajustes']}"), quebra=True,
             recuo=1)

    # Quadro 1.
    _secao(ws, L_Q1_SEC)
    _cabecalho(ws, L_Q1_CAB)
    ultima = L_Q1_C0 + 4
    _linhas_dados(ws, L_Q1_C0, ultima)
    _fonte(ws.Range(f"B{L_Q1_C0}:B{ultima}"), negrito=True)
    _formato(ws.Range(f"C{L_Q1_C0}:D{ultima}"), DATA_LOC, DATA_INV)
    _formato(ws.Range(f"F{L_Q1_C0}:F{ultima}"), DATA_LOC, DATA_INV)
    _formato(ws.Range(f"E{L_Q1_C0}:E{ultima}"), PCT_LOC, PCT_INV)
    _alinhar(ws.Range(f"C{L_Q1_CAB}:F{ultima}"), h=XL_CENTER, quebra=True)
    _alinhar(ws.Range(f"G{L_Q1_C0}:G{ultima}"), quebra=True, recuo=1)
    _fonte(ws.Range(f"G{L_Q1_C0}:G{ultima}"), tamanho=9)
    _cf_cor(ws.Range(f"B{L_Q1_C0}:G{ultima}"),
            f'=ISNUMBER(SEARCH("vigente",$B{L_Q1_C0}))', _bgr(CARDS["B"][0]))

    # Quadro 2.
    _secao(ws, L_Q2_SEC)
    _cabecalho(ws, L_Q2_CAB, ate="E")
    _linhas_dados(ws, L_Q2_C0, L_Q2_TOTAL, ate="E")
    _fonte(ws.Range(f"B{L_Q2_C0}:B{L_Q2_TOTAL}"), negrito=True)
    _moeda(ws.Range(f"C{L_Q2_C0}:E{L_Q2_TOTAL}"))
    _alinhar(ws.Range(f"C{L_Q2_CAB}:E{L_Q2_CAB}"), h=XL_RIGHT, quebra=True)
    _fonte(ws.Range(f"B{L_Q2_AJUSTE}"), negrito=False, italico=True, cor=CINZA_TEXTO)
    total = ws.Range(f"B{L_Q2_TOTAL}:E{L_Q2_TOTAL}")
    _fonte(total, negrito=True, cor=TITULO_FUNDO)
    _preencher(total, CAB_FUNDO)
    _borda(total, XL_EDGE_TOP, SECAO_FUNDO, XL_MEDIUM)

    # Quadro 3.
    _secao(ws, L_Q3_SEC)
    for linha in (L_Q3_NOTA, L_Q3_CASO):
        _mesclar(ws, f"B{linha}:G{linha}")
        alvo = ws.Range(f"B{linha}")
        _alinhar(alvo, quebra=True, recuo=1)
        _fonte(alvo, tamanho=9, italico=True, cor=CINZA_TEXTO)
    # Cabecalho e linhas so ganham fundo/bordas quando o quadro existe (mesmo
    # gatilho do PR #172: cabecalho de valores visivel em F15 do detalhe).
    existe = f'={DET}!$F$15<>""'
    cab = ws.Range(f"B{L_Q3_CAB}:G{L_Q3_CAB}")
    _fonte(cab, tamanho=9.5, negrito=True, cor=CAB_TEXTO)
    _alinhar(cab, quebra=True)
    ws.Range(f"B{L_Q3_CAB}").IndentLevel = 1
    regra = _cf_cor(cab, existe, _bgr(CAB_FUNDO))
    regra.Borders(XL_CF_BOTTOM).LineStyle = XL_CONTINUOUS
    regra.Borders(XL_CF_BOTTOM).Color = _bgr("9DB3D3")
    _mesclar(ws, f"F{L_Q3_CAB}:G{L_Q3_CAB}")
    ultima = L_Q3_C1 + 3
    corpo = ws.Range(f"B{L_Q3_C1}:G{ultima}")
    _fonte(corpo, tamanho=10, cor=CORPO_TEXTO)
    ws.Range(f"B{L_Q3_C1}:B{ultima}").IndentLevel = 1
    regra = _cf_cor(corpo, f'={DET}!$E$17&{DET}!$E$18&{DET}!$E$19&{DET}!$E$20<>""', None)
    regra.Borders(XL_CF_BOTTOM).LineStyle = XL_CONTINUOUS
    regra.Borders(XL_CF_BOTTOM).Color = _bgr(BORDA_LINHA)
    _fonte(ws.Range(f"B{L_Q3_C1}:B{ultima}"), negrito=True)
    _alinhar(ws.Range(f"C{L_Q3_CAB}:E{L_Q3_CAB}"), h=XL_RIGHT, quebra=True)
    _moeda(ws.Range(f"C{L_Q3_C1}:E{ultima}"))
    for linha in range(L_Q3_C1, ultima + 1):
        _mesclar(ws, f"F{linha}:G{linha}")
    _alinhar(ws.Range(f"F{L_Q3_C1}:G{ultima}"), quebra=True, recuo=1)
    _fonte(ws.Range(f"F{L_Q3_C1}:G{ultima}"), tamanho=9, cor=CINZA_TEXTO)
    _cf_cor(ws.Range(f"E{L_Q3_C1}:E{ultima}"), f'=$E{L_Q3_C1}<>""',
            cor_pot_fundo, cor_pot_texto)

    # Detalhes do metodo (so PCs; nos demais metodos o bloco fica vazio e sem
    # formatacao, no FIM da pagina — nenhum buraco no meio da leitura).
    ultima = L_DM_1 + 7
    for linha in range(L_DM_CAB, ultima + 1):
        _mesclar(ws, f"D{linha}:G{linha}")
    _fonte(ws.Range(f"B{L_DM_SEC}"), tamanho=11, negrito=True, cor=SECAO_TEXTO)
    _alinhar(ws.Range(f"B{L_DM_SEC}"), recuo=1)
    _fonte(ws.Range(f"B{L_DM_CAB}:G{L_DM_CAB}"), tamanho=9.5, negrito=True,
           cor=CAB_TEXTO)
    _alinhar(ws.Range(f"B{L_DM_CAB}:G{L_DM_CAB}"), quebra=True)
    ws.Range(f"B{L_DM_CAB}").IndentLevel = 1
    _alinhar(ws.Range(f"B{L_DM_1}:B{ultima}"), quebra=True, recuo=1)
    _moeda(ws.Range(f"C{L_DM_1}:C{ultima}"))
    _alinhar(ws.Range(f"D{L_DM_1}:G{ultima}"), quebra=True, recuo=1)
    _fonte(ws.Range(f"D{L_DM_1}:G{ultima}"), tamanho=9, cor=CINZA_TEXTO)
    _cf_cor(ws.Range(f"B{L_DM_SEC}:G{L_DM_SEC}"), f"={PCS}", _bgr(SECAO_FUNDO))
    _cf_cor(ws.Range(f"B{L_DM_CAB}:G{L_DM_CAB}"), f"={PCS}", _bgr(CAB_FUNDO))
    regra = _cf_cor(ws.Range(f"B{L_DM_1}:G{ultima}"), f"={PCS}", None)
    regra.Borders(XL_CF_BOTTOM).LineStyle = XL_CONTINUOUS
    regra.Borders(XL_CF_BOTTOM).Color = _bgr(BORDA_LINHA)
    _cf_cor(ws.Range(f"B{L_DM_1 + 6}:G{L_DM_1 + 6}"),
            f"=AND({PCS},N(RETROATIVO_POTENCIAL_VTA)<>0)",
            cor_pot_fundo, cor_pot_texto)

    # Rodape.
    _mesclar(ws, f"B{L_RODAPE}:G{L_RODAPE}")
    rod = ws.Range(f"B{L_RODAPE}")
    _alinhar(rod, quebra=True, recuo=1)
    _fonte(rod, tamanho=8.5, italico=True, cor=CINZA_TEXTO)


def _escrever(ws) -> None:
    for endereco, texto in CONSTANTES_EXECUTIVO.items():
        celula = ws.Range(endereco)
        celula.NumberFormat = "@"
        celula.Value = texto
    formulas = formulas_executivo()
    validar_ascii_e_parenteses(formulas)
    for endereco, formula in formulas.items():
        ws.Range(endereco).Formula = formula


def _configurar_impressao(ws) -> None:
    ps = ws.PageSetup
    ps.PrintArea = f"$B$1:$G${ULTIMA_LINHA}"
    ps.Orientation = XL_LANDSCAPE
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = False
    ps.CenterHorizontally = True


def aplicar(caminho: Path) -> None:
    import pythoncom
    import win32com.client as com

    pythoncom.CoInitialize()
    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = ws = det = None
    try:
        wb = excel.Workbooks.Open(str(caminho))
        excel.Calculation = XL_CALCULO_MANUAL
        _renomear_para_detalhe(wb)
        _ajustar_titulos_detalhe(wb)
        _aplicar_textos(wb)
        det = wb.Worksheets(DET)
        # Mesmo ambar da linha POTENCIAL da composicao do VTA no detalhe.
        regra_pot = det.Range("A84").FormatConditions(1)
        cor_pot_fundo = int(regra_pot.Interior.Color)
        cor_pot_texto = int(regra_pot.Font.Color)
        ws = _criar_aba_executiva(wb)
        _escrever(ws)
        _layout(ws, cor_pot_fundo, cor_pot_texto)
        _configurar_impressao(ws)
        excel.Calculation = XL_CALCULO_AUTOMATICO
        excel.CalculateFullRebuild()

        ws.Activate()
        janela = excel.ActiveWindow
        janela.DisplayGridlines = False
        janela.Zoom = 100
        janela.ScrollRow = 1
        janela.ScrollColumn = 1
        ws.Range(f"B{L_CARD_VAL}").Select()
        det.Activate()
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1
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
        wb = ws = det = excel = None
        gc.collect()
        pythoncom.CoUninitialize()


def main() -> int:
    caminho = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else TEMPLATE
    if not caminho.exists():
        print("ERRO: arquivo nao encontrado: %s" % caminho)
        return 1
    import pywintypes

    for tentativa in range(1, 6):
        try:
            aplicar(caminho)
            break
        except pywintypes.com_error as erro:
            if erro.args[0] != -2147418111 or tentativa == 5:
                raise
            print("Excel ocupado (tentativa %d); reexecutando do zero..." % tentativa)
            time.sleep(8.0)
    print("OK: RESULTADOS-EXECUTIVO-V3 aplicado em %s" % caminho)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
