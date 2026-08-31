# -*- coding: utf-8 -*-
"""Gera o plano declarativo da nova apresentacao da aba RESULTADOS (PR 2).

O plano e um JSON UTF-8 consumido por tools/aplicar_resultados_ux2_pr2.ps1,
que o aplica no template via Excel COM. A separacao existe por um motivo
pratico: o PowerShell 5.1 desta maquina nao le acentos de forma confiavel a
partir do proprio .ps1, entao todo o texto acentuado vive aqui, em UTF-8.

PRINCIPIO: o motor tecnico das linhas 1-87 nao e tocado. A camada humana
nasce nas linhas 90-166 e le o motor por defined name ou por espelho simples.
Nenhuma formula economica e movida, reancorada ou reescrita.
"""
from __future__ import annotations

import json
from pathlib import Path

# ----------------------------------------------------------------- paleta
# (§23) azul escuro = titulo/VTA; azul claro = cabecalho de secao;
# verde = validacao; cinza = referencia secundaria; ambar = potencial.
AZUL_ESCURO = "1F3864"
AZUL_CLARO = "D9E2F3"
VERDE = "C6E0B4"
VERDE_TXT = "375623"
CINZA = "F2F2F2"
CINZA_TXT = "595959"
AMBAR = "FFE699"
AMBAR_TXT = "9A6700"
VERMELHO = "F8CBAD"
VERMELHO_TXT = "9C0006"
BRANCO = "FFFFFF"

# ARMADILHA DE LOCALE: a propriedade .NumberFormat do Excel pt-BR reinterpreta
# padroes escritos em ingles — "0.000000" vira "0,000,000", "mm/yyyy" vira
# "mm/\y\y\y\y" e "0.00%" vira "#,000%". Por isso todo formato daqui e
# expresso em pt-BR e aplicado por .NumberFormatLocal; o monetario e copiado
# de uma celula ja homologada (RESULTADOS!C5), o que dispensa traducao.
MOEDA = "@COPIAR_DE:C5"
PCT = '0,00%;-0,00%;0,00%;"—"'
FATOR = "0,000000"
COMPETENCIA = "mm/aaaa"
# O ";@" e obrigatorio: "dd/mm/aaaa" puro e o NOME LOCALIZADO do formato de
# data embutido, e o Excel o grava como formato de sistema — que fora do
# pt-BR reaparece como "mm-dd-yy". Com a secao de texto vira formato
# customizado e o padrao brasileiro fica gravado no arquivo. E a mesma causa
# raiz do defeito pre-existente em CONTROLE!B3.
DATA = "dd/mm/aaaa;@"
GERAL = "Geral"

# Condicao unica do bloco de potencial (§11): metodo PC E potencial material.
COND_POT = 'AND(MEMORIA_RESULTADOS!$B$4="PCs",N(RETROATIVO_POTENCIAL_PC)>0)'

celulas: list[dict] = []
merges: list[str] = []
alturas: dict[str, float] = {}
condicionais: list[dict] = []


def cel(ref, *, valor=None, formula=None, fmt=None, tam=None, negrito=False,
        italico=False, cor=None, fundo=None, wrap=False):
    item = {"ref": ref}
    if valor is not None:
        item["valor"] = valor
    if formula is not None:
        item["formula"] = formula
    # "fmt" = copiar de celula existente; "fmtl" = padrao pt-BR aplicado por
    # NumberFormatLocal. Sem formato declarado a celula recebe "Geral", para
    # nunca herdar o ";;;" (texto invisivel) de vizinhas ocultas.
    if fmt is None:
        item["fmtl"] = GERAL
    elif fmt.startswith("@COPIAR_DE"):
        item["fmt"] = fmt
    else:
        item["fmtl"] = fmt
    if tam is not None:
        item["tam"] = tam
    if negrito:
        item["negrito"] = True
    if italico:
        item["italico"] = True
    if cor is not None:
        item["cor"] = cor
    if fundo is not None:
        item["fundo"] = fundo
    if wrap:
        item["wrap"] = True
    celulas.append(item)


def secao(ref_merge, ref, texto):
    """Cabecalho de secao: azul muito claro, texto azul escuro (§23)."""
    merges.append(ref_merge)
    cel(ref, valor=texto, tam=12, negrito=True, cor=AZUL_ESCURO, fundo=AZUL_CLARO)


def rotulo(ref, texto, *, fundo=None, wrap=False):
    cel(ref, valor=texto, tam=9, negrito=True, cor=CINZA_TXT, fundo=fundo,
        wrap=wrap)


# =========================================================== PAGINA 1
# "O QUE FOI APURADO?"

merges.append("A90:H90")
cel("A90", valor="RESULTADO DA APURAÇÃO CONTRATUAL", tam=16, negrito=True,
    cor=BRANCO, fundo=AZUL_ESCURO)
alturas["90"] = 30

# --- faixa compacta de contexto (§9B)
for m in ("C91:D91", "E91:F91", "G91:H91", "C92:D92", "E92:F92", "G92:H92"):
    merges.append(m)
rotulo("A91", "STATUS", fundo=AZUL_CLARO)
rotulo("B91", "MÉTODO", fundo=AZUL_CLARO)
rotulo("C91", "CICLOS", fundo=AZUL_CLARO)
rotulo("E91", "DATA DE CORTE", fundo=AZUL_CLARO)
rotulo("G91", "ÍNDICE", fundo=AZUL_CLARO)
cel("A92", formula="=$B$3", tam=11, negrito=True)          # STATUS_RESULTADOS
cel("B92", formula='=IF($B$5="","—",$B$5)', tam=11)
cel("C92", formula='=IF($J$11="","—",$J$11)', tam=11)
cel("E92", formula='=IF(CONTROLE!$B$3="","—",CONTROLE!$B$3)', fmt=DATA, tam=11)
cel("G92", formula='=IF(parametros!$E$16="","—",parametros!$E$16)', fmt=FATOR,
    tam=11)
alturas["91"] = 16
alturas["92"] = 20
alturas["93"] = 8

# STATUS com cor derivada do proprio status canonico (§24).
condicionais.append({"faixa": "A92", "expr": '=$A$92="VALIDADO"',
                     "fundo": VERDE, "cor": VERDE_TXT})
condicionais.append({"faixa": "A92", "expr": '=$A$92="ESTIMADO"',
                     "fundo": AMBAR, "cor": AMBAR_TXT})
condicionais.append({"faixa": "A92", "expr": '=$A$92="REVISE"',
                     "fundo": VERMELHO, "cor": VERMELHO_TXT})

# --- HERO VALUES (§9C): os dois numeros de maior hierarquia
for m in ("A94:D94", "E94:H94", "A95:D95", "E95:H95"):
    merges.append(m)
cel("A94", valor="VTA OFICIAL", tam=10, negrito=True, cor=AZUL_ESCURO,
    fundo=AZUL_CLARO)
cel("E94", valor="RETROATIVO TOTAL A PAGAR", tam=10, negrito=True,
    cor=AZUL_ESCURO, fundo=AZUL_CLARO)
cel("A95", formula='=IF(VTA_FINAL="","",VTA_FINAL)', fmt=MOEDA, tam=20,
    negrito=True, cor=AZUL_ESCURO)
cel("E95", formula="=$D$22", fmt=MOEDA, tam=20, negrito=True, cor=AZUL_ESCURO)
alturas["94"] = 16
alturas["95"] = 32
alturas["96"] = 8

# --- indicadores secundarios (§9D): hierarquia menor que o hero
for m in ("A97:B97", "C97:D97", "E97:F97", "G97:H97",
          "A98:B98", "C98:D98", "E98:F98", "G98:H98"):
    merges.append(m)
rotulo("A97", "REMANESCENTE ATUALIZADO")
rotulo("C97", "SALDO DO CICLO")
rotulo("E97", "VARIAÇÃO ACUMULADA")
rotulo("G97", "SITUAÇÃO")
cel("A98", formula='=IF(REM_ATUALIZADO_OFICIAL="","",REM_ATUALIZADO_OFICIAL)',
    fmt=MOEDA, tam=12, negrito=True)
cel("C98", formula='=IF(SALDO_REMANESCENTE_ATUAL="","",SALDO_REMANESCENTE_ATUAL)',
    fmt=MOEDA, tam=12, negrito=True)
cel("E98", formula='=IF($D$6="","",$D$6)', fmt=PCT, tam=12, negrito=True)
cel("G98", formula='=IF($H$33="","",$H$33)', tam=12, negrito=True)
alturas["97"] = 14
alturas["98"] = 20
alturas["99"] = 8

# --- INSERCAO 2 (§11): retroativo potencial, ambar, so quando material
for m in ("A100:C100", "E100:H100", "A101:H101"):
    merges.append(m)
cel("A100", formula=f'=IF({COND_POT},"RETROATIVO POTENCIAL — EM ANÁLISE","")',
    tam=10, negrito=True, cor=AMBAR_TXT)
cel("D100", formula=f'=IF({COND_POT},RETROATIVO_POTENCIAL_PC,"")',
    fmt=MOEDA, tam=13, negrito=True, cor=AMBAR_TXT)
cel("A101", formula=(
    f'=IF({COND_POT},"Valor potencial relacionado a PCs ainda em análise. '
    'Não compõe os valores oficiais enquanto permanecer em análise.","")'),
    tam=9, cor=AMBAR_TXT)
alturas["100"] = 18
alturas["101"] = 14
alturas["102"] = 8
# ARMADILHA: FormatConditions.Add usa o separador de lista LOCAL (";" em
# pt-BR), ao contrario de .Formula, que usa virgula. Uma expressao com
# AND(a,b) e recusada com "valor fora do intervalo esperado". A guia abaixo
# concentra a logica numa celula normal (via .Formula, virgula) e deixa a
# regra condicional com uma comparacao unica, imune ao locale.
cel("I100", formula=f"=IF({COND_POT},1,0)")
condicionais.append({"faixa": "A100:H101", "expr": "=$I$100=1",
                     "fundo": AMBAR, "cor": AMBAR_TXT})

# --- INSERCAO 1 (§10): aviso de revisao, destaque discreto
merges.append("A103:H103")
cel("A103", valor=(
    "IMPORTANTE — Esta ferramenta auxilia a apuração e a conferência dos "
    "cálculos, mas não substitui a revisão do responsável pela análise. "
    "Antes de qualquer aprovação ou formalização, confirme os dados de "
    "entrada, os critérios aplicados e os resultados apresentados."),
    tam=9, italico=True, cor=CINZA_TXT, fundo=CINZA, wrap=True)
alturas["103"] = 28
alturas["104"] = 8

# --- ciclos apurados (§9E). Fatores NAO aparecem nesta pagina.
secao("A105:H105", "A105", "CICLOS APURADOS")
merges.append("G106:H106")
# "INÍCIO DO EFEITO FINANCEIRO" nomeia o que parametros!H de fato guarda: a
# competencia em que o reajuste passa a produzir efeitos. O rotulo antigo
# ("Efeito financeiro") era ambiguo — podia ser lido como Sim/Nao ou como
# valor monetario. A fonte e o formato mm/aaaa nao mudam.
for ref, txt in (("A106", "Ciclo"), ("B106", "Situação"), ("C106", "Período"),
                 ("D106", "Variação apurada"), ("E106", "Percentual aplicado"),
                 ("F106", "INÍCIO DO EFEITO FINANCEIRO"),
                 ("G106", "Observação")):
    rotulo(ref, txt, fundo=CINZA, wrap=(ref == "F106"))
# O cabecalho passa a caber em duas linhas sem truncar na coluna F.
alturas["106"] = 26
for i in range(5):
    r = 107 + i          # linha de destino
    p = 2 + i            # parametros: cadastro do ciclo (B/C/D/E/G/H)
    m = 11 + i           # parametros: memoria do fator aplicavel (C/F)
    merges.append(f"G{r}:H{r}")
    cel(f"A{r}", formula=f'=IF(parametros!$B${p}="","",parametros!$B${p})',
        tam=10)
    cel(f"B{r}", formula=(
        f'=IF(parametros!$B${p}="","",'
        f'IF(parametros!$G${p}="","—",parametros!$G${p}))'), tam=10)
    # Periodo montado com DAY/MONTH/YEAR: mesmo idioma ja usado na aba, imune
    # ao locale do TEXT().
    cel(f"C{r}", formula=(
        f'=IF(OR(parametros!$C${p}="",parametros!$D${p}=""),"",'
        f'RIGHT("0"&DAY(parametros!$C${p}),2)&"/"&'
        f'RIGHT("0"&MONTH(parametros!$C${p}),2)&"/"&YEAR(parametros!$C${p})'
        f'&" a "&'
        f'RIGHT("0"&DAY(parametros!$D${p}),2)&"/"&'
        f'RIGHT("0"&MONTH(parametros!$D${p}),2)&"/"&YEAR(parametros!$D${p}))'),
        tam=10)
    cel(f"D{r}", formula=f'=IF(parametros!$E${p}="","",parametros!$E${p})',
        fmt=PCT, tam=10)
    # Neutralizado aparece como 0,00% e nao como vazio (§9).
    cel(f"E{r}", formula=(
        f'=IF(parametros!$B${p}="","",'
        f'IF(parametros!$C${m}="",0,parametros!$C${m}))'), fmt=PCT, tam=10)
    cel(f"F{r}", formula=f'=IF(parametros!$H${p}="","",parametros!$H${p})',
        fmt=COMPETENCIA, tam=10)
    cel(f"G{r}", formula=(
        f'=IF(parametros!$B${p}="","",'
        f'IF(parametros!$F${m}="","",parametros!$F${m}))'), tam=9,
        cor=CINZA_TXT)
    alturas[str(r)] = 16
alturas["112"] = 8

# --- conclusao / pendencias (§12): potencial nao pode virar "sem pendencias"
merges.append("A113:H113")
merges.append("A114:H114")
cel("A113", formula=(
    f'=IF($J$5=0,IF({COND_POT},'
    '"APURAÇÃO CONCLUÍDA — COM VALOR POTENCIAL EM ANÁLISE",'
    '"APURAÇÃO CONCLUÍDA"),"PENDÊNCIAS PARA CONCLUSÃO")'),
    tam=10, negrito=True)
cel("A114", formula=(
    f'=IF($J$5=0,IF({COND_POT},'
    '"Os valores oficiais estão apurados, mas ainda há retroativo potencial '
    'em análise (bloco em âmbar acima). Ele não compõe o VTA nem o retroativo '
    'oficial.",'
    '"Nenhuma pendência registrada pelas validações desta apuração."),$B$7)'),
    tam=9, wrap=True)
alturas["113"] = 16
alturas["114"] = 26
alturas["115"] = 8
alturas["116"] = 8

# =========================================================== PAGINA 2
# "COMO ESSE RESULTADO FOI FORMADO?"

secao("A117:H117", "A117",
      "VALOR TOTAL ATUALIZADO DO CONTRATO — METODOLOGIA E FORMAÇÃO")
merges.append("A118:H118")
merges.append("A119:H119")
cel("A118", formula="=$A$69", tam=9, cor=CINZA_TXT)   # identidade conceitual
cel("A119", formula="=$A$70", tam=9, cor=CINZA_TXT)   # frase do metodo
merges.append("C120:H120")
for ref, txt in (("A120", "Parcela"), ("B120", "Valor"), ("C120", "Origem")):
    rotulo(ref, txt, fundo=CINZA)
formacao = (
    (121, "Executado apurado",
     '=IF(EXECUTADO_APURADO="","",EXECUTADO_APURADO)', "=$C$83"),
    (122, "(+) Ajustes ainda devidos",
     '=IF(AJUSTES_DEVIDOS="","",AJUSTES_DEVIDOS)', "=$C$84"),
    (123, "(+) Remanescente atualizado", '=IF($B$85="","",$B$85)', "=$C$85"),
    (124, "(=) VTA OFICIAL", '=IF(VTA_FINAL="","",VTA_FINAL)', "=$C$86"),
)
for r, parcela, valor, origem in formacao:
    merges.append(f"C{r}:H{r}")
    cel(f"A{r}", valor=parcela, tam=10, negrito=(r == 124))
    cel(f"B{r}", formula=valor, fmt=MOEDA, tam=10, negrito=(r == 124))
    cel(f"C{r}", formula=origem, tam=9, cor=CINZA_TXT)
    alturas[str(r)] = 16
# Conferencia da formacao: "DE ACORDO" quando fecha (§13).
merges.append("C125:H125")
cel("A125", valor="Conferência da formação", tam=10)
cel("B125",
    formula='=IF(CONFERENCIA_FORMACAO_VTA="","",CONFERENCIA_FORMACAO_VTA)',
    fmt=MOEDA, tam=10)
cel("C125", formula=(
    '=IF(CONFERENCIA_FORMACAO_VTA="","Aguardando base para conferir.",'
    'IF(ABS(CONFERENCIA_FORMACAO_VTA)<=MEMORIA_RESULTADOS!$D$4,'
    '"DE ACORDO — as parcelas fecham com o VTA Oficial.",$C$87))'), tam=9)
alturas["125"] = 16
alturas["126"] = 8
cel("I125", formula=('=IF(AND(CONFERENCIA_FORMACAO_VTA<>"",'
                     'ABS(CONFERENCIA_FORMACAO_VTA)<=MEMORIA_RESULTADOS!$D$4)'
                     ',1,0)'))
condicionais.append({"faixa": "C125", "expr": "=$I$125=1",
                     "fundo": VERDE, "cor": VERDE_TXT})

# --- retroativo por ciclo (§14): "Diferença", nunca "Delta"
secao("A127:H127", "A127", "RETROATIVO POR CICLO")
for ref, txt in (("A128", "Ciclo"), ("B128", "Pago / considerado"),
                 ("C128", "Reajustado"), ("D128", "Diferença")):
    rotulo(ref, txt, fundo=CINZA)
for i in range(5):
    r, o = 129 + i, 16 + i
    cel(f"A{r}", formula=f"=$A${o}", tam=10)
    cel(f"B{r}", formula=f'=IF($B${o}="","",$B${o})', fmt=MOEDA, tam=10)
    cel(f"C{r}", formula=f'=IF($C${o}="","",$C${o})', fmt=MOEDA, tam=10)
    cel(f"D{r}", formula=f'=IF($D${o}="","",$D${o})', fmt=MOEDA, tam=10)
    alturas[str(r)] = 15
cel("A134", valor="TOTAL", tam=10, negrito=True)
cel("B134", formula='=IF($B$22="","",$B$22)', fmt=MOEDA, tam=10, negrito=True)
cel("C134", formula='=IF($C$22="","",$C$22)', fmt=MOEDA, tam=10, negrito=True)
cel("D134", formula='=IF($D$22="","",$D$22)', fmt=MOEDA, tam=10, negrito=True)
alturas["134"] = 16
alturas["135"] = 8

# --- remanescente por ciclo (§15): fator com no maximo 6 casas na exibicao
secao("A136:H136", "A136", "REMANESCENTE POR CICLO")
for ref, txt in (("A137", "Ciclo"), ("B137", "Remanescente base"),
                 ("C137", "Fator"), ("D137", "Remanescente atualizado")):
    rotulo(ref, txt, fundo=CINZA)
for i in range(5):
    r, o, p = 138 + i, 26 + i, 2 + i
    cel(f"A{r}", formula=f"=$A${o}", tam=10)
    cel(f"B{r}", formula=f'=IF($B${o}="","",$B${o})', fmt=MOEDA, tam=10)
    cel(f"C{r}", formula=(
        f'=IF($B${o}="","",IF(parametros!$F${p}="","",parametros!$F${p}))'),
        fmt=FATOR, tam=10)
    cel(f"D{r}", formula=f'=IF($C${o}="","",$C${o})', fmt=MOEDA, tam=10)
    alturas[str(r)] = 15
alturas["143"] = 8

# --- ciclo em execucao (§16): "SITUAÇÃO ATUAL DO CONTRATO"
secao("A144:H144", "A144", "SITUAÇÃO ATUAL DO CONTRATO")
situacao = (
    (145, "Data da referência", '=IF($B$35="","",$B$35)', DATA),
    (146, "Execução atualizada",
     '=IF(EXECUCAO_ATUALIZADA_CICLO="","",EXECUCAO_ATUALIZADA_CICLO)', MOEDA),
    (147, "Remanescente atualizado", '=IF($B$37="","",$B$37)', MOEDA),
    (148, "Saldo do ciclo",
     '=IF(SALDO_REMANESCENTE_ATUAL="","",SALDO_REMANESCENTE_ATUAL)', MOEDA),
)
for r, lab, form, fmt in situacao:
    merges.append(f"C{r}:H{r}")
    cel(f"A{r}", valor=lab, tam=10)
    cel(f"B{r}", formula=form, fmt=fmt, tam=10, negrito=True)
    alturas[str(r)] = 15
alturas["149"] = 8

# --- INSERCAO 3 (§17): referencias com FINALIDADE legivel
secao("A150:H150", "A150",
      "REFERÊNCIAS PARA CONFERÊNCIA — NÃO SÃO O VTA OFICIAL")
merges.append("C151:H151")
for ref, txt in (("A151", "Referência"), ("B151", "Valor"),
                 ("C151", "FINALIDADE")):
    rotulo(ref, txt, fundo=CINZA)
referencias = (
    (152, "Situação atual do contrato", "$B$10", "REFERÊNCIA ATUAL"),
    (153, "Última referência de abertura", "$B$11", "REFERÊNCIA DE ABERTURA"),
    (154, "Contrato original integralmente reajustado", "$B$12",
     "COMPARATIVO TEÓRICO"),
)
for r, lab, origem, finalidade in referencias:
    merges.append(f"C{r}:H{r}")
    cel(f"A{r}", valor=lab, tam=10, cor=CINZA_TXT)
    cel(f"B{r}", formula=f'=IF({origem}="","",{origem})', fmt=MOEDA, tam=10)
    cel(f"C{r}", valor=finalidade, tam=9, negrito=True, cor=CINZA_TXT,
        fundo=CINZA)
    alturas[str(r)] = 15
alturas["155"] = 8

# --- conferencia entre referencias (§18): "DIFERENÇA", nao jargao antigo
secao("A156:H156", "A156", "CONFERÊNCIA ENTRE REFERÊNCIAS DO CONTRATO")
merges.append("C157:H157")
cel("A157", valor="DIFERENÇA", tam=10, negrito=True, cor=CINZA_TXT)
cel("B157", formula='=IF($B$13="","",$B$13)', fmt=MOEDA, tam=10)
cel("C157", formula='=IF($H$13="","",$H$13)', tam=9, cor=CINZA_TXT)
alturas["157"] = 15
alturas["158"] = 8

# --- conferencia da execucao (§19): tabela util so no Financeiro
secao("A159:H159", "A159", "CONFERÊNCIA DA EXECUÇÃO")
merges.append("F160:H160")
for ref, txt in (("A160", "Ciclo"), ("B160", "Desembolsado informado"),
                 ("C160", "Execução estimada pelo quantitativo"),
                 ("D160", "Diferença"), ("E160", "Conferência")):
    rotulo(ref, txt, fundo=CINZA)
cel("F160", formula=(
    '=IF(MEMORIA_RESULTADOS!$B$4="Financeiro","",'
    '"Não aplicável ao método selecionado.")'), tam=9, italico=True,
    cor=CINZA_TXT)
for i in range(5):
    r, o = 161 + i, 73 + i
    for col in ("A", "B", "C", "D", "E"):
        cel(f"{col}{r}", formula=(
            f'=IF(MEMORIA_RESULTADOS!$B$4<>"Financeiro","",${col}${o})'),
            fmt=(MOEDA if col in ("B", "C", "D") else None), tam=10)
    alturas[str(r)] = 15
merges.append("A166:H166")
cel("A166", formula='=IF(MEMORIA_RESULTADOS!$B$4<>"Financeiro","",$A$78)',
    tam=9, italico=True, cor=CINZA_TXT)
alturas["166"] = 14

plano = {
    "aba": "RESULTADOS",
    "memoria": {
        "S38": "Retroativo potencial de PCs em analise (ate o corte)",
        "T38": ('=ROUND(SUMIFS(itens_PC!$J$2:$J$5001,'
                'itens_PC!$B$2:$B$5001,"<="&$T$31),2)'),
    },
    "name_novo": {"nome": "RETROATIVO_POTENCIAL_PC",
                  "refere": "=MEMORIA_RESULTADOS!$T$38"},
    "merges": merges,
    "celulas": celulas,
    "alturas": alturas,
    "condicionais": condicionais,
    # A camada tecnica inteira sai da tela e da impressao (§20, §21).
    "ocultar_linhas": [[1, 89]],
    # Coluna I hospeda apenas as guias de formatacao condicional.
    "ocultar_colunas": ["I"],
    # (§25) Dois acertos de formato ja mapeados, atingidos por este leiaute.
    # CONTROLE!B3 estava em "mm-dd-yy" (padrao americano) pela mesma armadilha
    # do formato embutido; B55:B59 estava em "Geral" onde as vizinhas B60:B63
    # ja eram monetarias.
    "formatos_extra": [
        {"aba": "CONTROLE", "ref": "B3", "fmtl": DATA},
        {"aba": "RESULTADOS", "ref": "B55:B59", "fmt": MOEDA},
    ],
    "impressao": {
        "print_area": "$A$90:$H$166",
        "quebra_antes_da_linha": 117,
        # Escala calibrada em Excel real: FitToPagesTall=2 nao convive com a
        # quebra manual (o Excel calcula a escala pela altura total e ignora o
        # corte, devolvendo 3 paginas). Com zoom explicito, 75% e o maior
        # valor que fecha em exatamente 2 paginas — acima do 68% do leiaute
        # anterior, portanto sem miniaturizar mais do que ja se fazia.
        "zoom": 75,
    },
}

destino = Path(__file__).resolve().parent / "resultados_ux2_plano.json"
destino.write_text(
    json.dumps(plano, ensure_ascii=False, indent=1), encoding="utf-8"
)
print(f"plano gravado: {destino}")
print(f"  celulas={len(celulas)} merges={len(merges)} "
      f"alturas={len(alturas)} condicionais={len(condicionais)}")
