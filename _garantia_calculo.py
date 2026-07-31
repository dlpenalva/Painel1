"""Motor de cálculo da Garantia Contratual (método único).

Regra pétrea (Cláusula Décima): a garantia corresponde a 5% do VALOR TOTAL DO
CONTRATO após cada instrumento. O endosso/adequação de cada instrumento é a
diferença entre a garantia daquele instrumento e a garantia do instrumento
imediatamente anterior.

    GARANTIA DO INSTRUMENTO = ARREDONDAR(VALOR TOTAL DO CONTRATO x 5%; 2)
    ENDOSSO/ADEQUAÇÃO       = GARANTIA ATUAL - GARANTIA ANTERIOR

Nunca se usa como base: saldo remanescente, valor executado, valor pago, valor
retroativo isolado, valor apenas do acréscimo, garantia informada manualmente,
apólice/endosso anteriormente apresentado. A base é sempre o valor total
consolidado do contrato após o instrumento.

Arredondamento financeiro: Decimal + ROUND_HALF_UP, duas casas decimais. O float
nunca é a base do cálculo monetário — apenas a formatação/serialização final.

Este módulo é puro (sem Streamlit) para permitir testes focais das funções de
conversão, cálculo, diferença, arredondamento, normalização do histórico, linha
do tempo e geração dos textos de comunicação.
"""
from __future__ import annotations

from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from zoneinfo import ZoneInfo

try:  # ReportLab é opcional; a página avisa quando ausente.
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from io import BytesIO
    from html import escape as _escape
    REPORTLAB_OK = True
except Exception:  # pragma: no cover - depende do ambiente
    REPORTLAB_OK = False


PERCENTUAL_GARANTIA = Decimal("0.05")   # 5% — regra contratual absoluta
PERCENTUAL_GARANTIA_PCT = 5.0
CENTAVO = Decimal("0.01")
TRACO = "—"  # em dash usado para "sem endosso anterior"


# ============================================================
# Conversão monetária e arredondamento financeiro
# ============================================================

def parse_moeda_br(valor):
    """Converte entrada monetária brasileira em ``Decimal``.

    Aceita: 2968866, 2968866,00, 2.968.866,00, R$ 2.968.866,00, 2968866.00 e
    valores numéricos (int/float/Decimal). Retorna ``None`` quando o valor não
    puder ser interpretado (para permitir distinção entre "zero" e "inválido").
    """
    if valor is None:
        return None
    if isinstance(valor, Decimal):
        return valor
    if isinstance(valor, bool):
        return None
    if isinstance(valor, int):
        return Decimal(valor)
    if isinstance(valor, float):
        if valor != valor:  # NaN
            return None
        return Decimal(str(valor))

    texto = str(valor).strip()
    if not texto:
        return None
    texto = (
        texto.replace("R$", "")
        .replace("r$", "")
        .replace(" ", "")
        .replace(" ", "")
    )
    if not texto:
        return None

    if "," in texto:
        # Formato BR: ponto é separador de milhar, vírgula é decimal.
        texto = texto.replace(".", "").replace(",", ".")
    elif texto.count(".") > 1:
        # Múltiplos pontos sem vírgula: pontos são separadores de milhar.
        texto = texto.replace(".", "")
    # Caso restante: inteiro ou decimal com ponto único.

    try:
        return Decimal(texto)
    except (InvalidOperation, ValueError):
        return None


def arredondar_financeiro(valor):
    """Arredonda para 2 casas com ROUND_HALF_UP (arredondamento financeiro)."""
    if not isinstance(valor, Decimal):
        valor = parse_moeda_br(valor)
        if valor is None:
            valor = Decimal("0")
    return valor.quantize(CENTAVO, rounding=ROUND_HALF_UP)


def calcular_garantia(valor_total_contrato):
    """Garantia = ARREDONDAR(valor total do contrato x 5%; 2)."""
    base = valor_total_contrato
    if not isinstance(base, Decimal):
        base = parse_moeda_br(base)
        if base is None:
            base = Decimal("0")
    return arredondar_financeiro(base * PERCENTUAL_GARANTIA)


def moeda(valor, com_prefixo=True):
    """Formata como moeda brasileira: R$ 2.968.866,00."""
    dec = parse_moeda_br(valor)
    if dec is None:
        dec = Decimal("0")
    dec = arredondar_financeiro(dec)
    texto = f"{dec:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {texto}" if com_prefixo else texto


def formatar_brl(valor):
    """Função única de formatação monetária brasileira: R$ 2.968.866,00.

    Toda exibição de valor na página, no TXT e no PDF passa por aqui (ou por
    ``moeda``, que é a mesma implementação). Ex.: Decimal("3117252.48") ->
    "R$ 3.117.252,48".
    """
    return moeda(valor, com_prefixo=True)


# ============================================================
# VTA (Valor Total Atualizado) fornecido pelo módulo Valor Global
# ============================================================

def extrair_vta(resultado_valor_global):
    """Recupera o VTA da análise atual do resultado do módulo Valor Global.

    Chave canônica: ``valor_atualizado_contrato`` (o módulo Valor Global grava o
    mesmo número também em ``valor_global_financeiro``). Retorna ``Decimal`` > 0
    ou ``None`` quando o VTA não estiver disponível/for inválido — nunca zero
    silencioso.
    """
    if not isinstance(resultado_valor_global, dict) or not resultado_valor_global:
        return None
    bruto = resultado_valor_global.get("valor_atualizado_contrato")
    if bruto is None:
        bruto = resultado_valor_global.get("valor_global_financeiro")
    valor = parse_moeda_br(bruto)
    if valor is None or valor <= 0:
        return None
    return arredondar_financeiro(valor)


# ============================================================
# Normalização do histórico e linha do tempo
# ============================================================

def normalizar_linhas_historico(registros):
    """Normaliza as linhas do histórico informado manualmente.

    ``registros``: iterável de dicts com "Instrumento" e "Valor total do
    contrato". Retorna ``(linhas_validas, avisos)`` onde cada linha válida é
    ``{"instrumento": str, "valor_total": Decimal}`` e ``avisos`` é a lista de
    mensagens de validação. Linhas totalmente vazias são ignoradas em silêncio.
    """
    linhas = []
    avisos = []
    for idx, reg in enumerate(registros, start=1):
        instrumento_bruto = reg.get("Instrumento")
        instrumento = "" if instrumento_bruto is None else str(instrumento_bruto).strip()
        valor_bruto = reg.get("Valor total do contrato")
        tem_texto_valor = valor_bruto is not None and str(valor_bruto).strip() != ""

        if not instrumento and not tem_texto_valor:
            continue  # linha vazia

        if instrumento and not tem_texto_valor:
            avisos.append(f"Linha {idx} ({instrumento}): instrumento sem valor total informado.")
            continue
        if tem_texto_valor and not instrumento:
            avisos.append(f"Linha {idx}: valor informado sem o nome do instrumento.")
            continue

        valor = parse_moeda_br(valor_bruto)
        if valor is None:
            avisos.append(f"Linha {idx} ({instrumento}): valor \"{str(valor_bruto).strip()}\" não pôde ser interpretado.")
            continue
        if valor <= 0:
            avisos.append(f"Linha {idx} ({instrumento}): o valor total do contrato deve ser maior que zero.")
            continue

        linhas.append({"instrumento": instrumento, "valor_total": arredondar_financeiro(valor)})

    return linhas, avisos


def classificar_endosso(endosso):
    """Classifica a diferença de garantia entre instrumentos consecutivos."""
    if endosso is None:
        return "inicial"
    if endosso > 0:
        return "aumento"
    if endosso < 0:
        return "reducao"
    return "sem_alteracao"


def construir_linha_do_tempo(itens):
    """Constrói a linha do tempo da garantia.

    ``itens``: lista ordenada de ``{"instrumento": str, "valor_total": Decimal}``
    (histórico manual + instrumento da análise atual). A ordem informada é a
    ordem cronológica — nunca é reordenada por nome.

    Cada linha recebe:
        garantia      = 5% do valor total daquela linha
        endosso       = garantia - garantia_da_linha_anterior (None na 1ª linha)
        tipo_endosso  = inicial | aumento | reducao | sem_alteracao
    """
    linha_do_tempo = []
    garantia_anterior = None
    for item in itens:
        valor_total = item["valor_total"]
        if not isinstance(valor_total, Decimal):
            valor_total = arredondar_financeiro(parse_moeda_br(valor_total) or Decimal("0"))
        garantia = calcular_garantia(valor_total)
        if garantia_anterior is None:
            endosso = None
        else:
            endosso = arredondar_financeiro(garantia - garantia_anterior)
        linha_do_tempo.append(
            {
                "instrumento": item["instrumento"],
                "valor_total": valor_total,
                "garantia": garantia,
                "endosso": endosso,
                "tipo_endosso": classificar_endosso(endosso),
            }
        )
        garantia_anterior = garantia
    return linha_do_tempo


def formatar_endosso(endosso, tipo_endosso):
    """Texto da coluna Endosso/Adequação conforme o sinal da diferença."""
    if tipo_endosso == "inicial" or endosso is None:
        return TRACO
    if tipo_endosso == "reducao":
        return f"Redução de {moeda(abs(endosso))}"
    # aumento e sem_alteracao (R$ 0,00)
    return moeda(endosso)


# ============================================================
# Comunicação à contratada
# ============================================================

def gerar_texto_comunicacao(numero_contrato, vta_atual, garantia_total, endosso, tipo_endosso, contratada=""):
    """Gera o texto da comunicação à contratada conforme o tipo de resultado.

    Um único parágrafo (o "parágrafo do endosso") varia conforme o resultado:
    garantia inicial, endosso positivo, ausência de alteração ou redução. O
    parágrafo de prazo só aparece quando há efetivamente algo a apresentar
    (garantia inicial ou endosso complementar).
    """
    contrato = (str(numero_contrato).strip() if numero_contrato else "") or "[número do contrato]"
    nome_contratada = str(contratada).strip() if contratada else ""
    vta_txt = moeda(vta_atual)
    garantia_txt = moeda(garantia_total)

    p_intro = f"Em razão da alteração do valor do Contrato nº {contrato}"
    if nome_contratada:
        p_intro += f", celebrado com {nome_contratada},"
    p_intro += f" o valor total contratual passa a corresponder a {vta_txt}."

    p_regra = (
        "Nos termos da Cláusula Décima, a garantia contratual deve corresponder a 5% do valor "
        f"total do contrato, totalizando {garantia_txt}."
    )

    pede_prazo = False
    if tipo_endosso == "inicial":
        p_variavel = f"Deverá ser apresentada garantia contratual no valor total de {garantia_txt}."
        pede_prazo = True
    elif tipo_endosso == "reducao":
        reducao_txt = moeda(abs(endosso)) if endosso is not None else moeda(0)
        p_variavel = (
            f"O novo valor de referência da garantia corresponde a {garantia_txt}, representando "
            f"redução de {reducao_txt} em relação ao valor anterior. Eventual adequação do instrumento "
            "deverá ser submetida à análise e à aceitação da Telebras."
        )
    elif tipo_endosso == "sem_alteracao":
        p_variavel = (
            f"A cobertura total exigida permanece em {garantia_txt}, não havendo complementação "
            "financeira decorrente desta alteração. Deverá ser verificada, contudo, a necessidade de "
            "adequação do prazo de validade da garantia."
        )
    else:  # aumento
        endosso_txt = moeda(endosso) if endosso is not None else moeda(0)
        p_variavel = (
            f"Dessa forma, deverá ser apresentado endosso complementar no valor de {endosso_txt}, de "
            f"modo que a cobertura total da garantia passe a corresponder a {garantia_txt}."
        )
        pede_prazo = True

    p_prazo = (
        "Solicitamos o envio do endosso no prazo de até 5 dias úteis, contado do recebimento desta "
        "comunicação. Eventual prorrogação deverá ser previamente solicitada, mediante justificativa, e "
        "dependerá de aceitação da Telebras."
    )
    p_validade = (
        "A garantia deverá permanecer válida durante todo o período de vigência contratual e por mais 3 "
        "meses após o seu término, com possibilidade de prorrogação em caso de ocorrência de sinistro."
    )
    p_clausula = "O documento deverá observar as demais condições estabelecidas na Cláusula Décima do contrato."

    paragrafos = [
        f"Assunto: Garantia contratual – Contrato nº {contrato}",
        "",
        "Prezados,",
        "",
        p_intro,
        "",
        p_regra,
        "",
        p_variavel,
        "",
    ]
    if pede_prazo:
        paragrafos.extend([p_prazo, ""])
    paragrafos.extend(
        [
            p_validade,
            "",
            p_clausula,
            "",
            "Atenciosamente,",
            "",
            "Gerência de Compras e Contratos",
            "TELEBRAS",
        ]
    )
    return "\n".join(paragrafos)


def montar_txt_bytes(texto):
    """TXT em UTF-8 com BOM (abre corretamente no Windows/Bloco de Notas)."""
    return texto.encode("utf-8-sig")


# ============================================================
# Utilitários de apoio
# ============================================================

def data_hora_brasilia():
    return datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M")


def resumo_resultado(linha_do_tempo):
    """Extrai (vta_atual, garantia_total, endosso, tipo) da última linha."""
    if not linha_do_tempo:
        return None, None, None, "inicial"
    ultima = linha_do_tempo[-1]
    return (
        ultima["valor_total"],
        ultima["garantia"],
        ultima["endosso"],
        ultima["tipo_endosso"],
    )


# ============================================================
# Relatório PDF (memória de cálculo completa)
# ============================================================

def gerar_pdf_garantia(dados):
    """Gera o PDF "Memória de Cálculo da Garantia Contratual".

    ``dados`` deve conter: numero_contrato, contratada, vta_atual, garantia_total,
    endosso, tipo_endosso, linha_do_tempo (lista de dicts) e texto_comunicacao.
    Retorna ``bytes`` ou ``None`` se o ReportLab não estiver disponível.
    """
    if not REPORTLAB_OK:
        return None

    numero_contrato = str(dados.get("numero_contrato") or "").strip() or "[número do contrato]"
    contratada = str(dados.get("contratada") or "").strip()
    vta_atual = dados.get("vta_atual")
    garantia_total = dados.get("garantia_total")
    endosso = dados.get("endosso")
    tipo_endosso = dados.get("tipo_endosso", "inicial")
    linha_do_tempo = dados.get("linha_do_tempo") or []
    texto_comunicacao = dados.get("texto_comunicacao") or ""

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=1.6 * cm, leftMargin=1.6 * cm, topMargin=1.5 * cm, bottomMargin=1.5 * cm,
    )
    styles = getSampleStyleSheet()
    titulo = ParagraphStyle("TituloGarantia", parent=styles["Title"], alignment=TA_CENTER, fontSize=13, leading=16, spaceAfter=6)
    subtitulo = ParagraphStyle("SubtituloGarantia", parent=styles["Normal"], alignment=TA_CENTER, fontSize=9, leading=12, textColor=colors.HexColor("#334155"), spaceAfter=10)
    h2 = ParagraphStyle("H2Garantia", parent=styles["Heading2"], fontSize=10.5, leading=13, spaceBefore=9, spaceAfter=6, textColor=colors.HexColor("#123B63"))
    normal = ParagraphStyle("NormalGarantia", parent=styles["Normal"], fontSize=9, leading=12.5, alignment=TA_JUSTIFY)
    corpo_comunicacao = ParagraphStyle("ComunicacaoGarantia", parent=styles["Normal"], fontSize=9, leading=13, alignment=TA_JUSTIFY, spaceAfter=6)

    def _p(txt):
        return _escape(str(txt))

    elementos = []
    elementos.append(Paragraph("Memória de Cálculo da Garantia Contratual", titulo))
    elementos.append(Paragraph(f"Contrato nº {_p(numero_contrato)}", subtitulo))
    if contratada:
        elementos.append(Paragraph(f"Contratada: {_p(contratada)}", subtitulo))
    elementos.append(Paragraph(f"Gerado em: {data_hora_brasilia()}", subtitulo))

    elementos.append(Paragraph("1. Regra e fórmula aplicada", h2))
    elementos.append(Paragraph("Percentual aplicado: 5% sobre o valor total do contrato após cada instrumento.", normal))
    elementos.append(Paragraph("Garantia do instrumento = ARREDONDAR(valor total do contrato x 5%; 2).", normal))
    elementos.append(Paragraph("Endosso/adequação = garantia do instrumento atual menos a garantia do instrumento anterior.", normal))
    elementos.append(Spacer(1, 6))

    elementos.append(Paragraph("2. Memória de cálculo (todos os instrumentos)", h2))
    linhas_tabela = [["Instrumento", "Valor total do contrato", "Garantia (5%)", "Endosso / Adequação"]]
    for linha in linha_do_tempo:
        linhas_tabela.append([
            _p(linha.get("instrumento", "")),
            moeda(linha.get("valor_total", 0)),
            moeda(linha.get("garantia", 0)),
            formatar_endosso(linha.get("endosso"), linha.get("tipo_endosso", "inicial")),
        ])
    tabela = Table(linhas_tabela, colWidths=[5.2 * cm, 4.1 * cm, 3.6 * cm, 3.6 * cm], repeatRows=1, hAlign="CENTER")
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9E2F3")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F8FAFC")),
    ]))
    elementos.append(tabela)
    elementos.append(Spacer(1, 8))

    elementos.append(Paragraph("3. Resultado da análise atual", h2))
    if tipo_endosso == "reducao":
        rotulo_endosso = "Adequação apurada"
        valor_endosso = f"Redução de {moeda(abs(endosso))}" if endosso is not None else "Redução de " + moeda(0)
    elif tipo_endosso == "inicial":
        rotulo_endosso = "Garantia inicial"
        valor_endosso = moeda(garantia_total)
    else:
        rotulo_endosso = "Endosso necessário"
        valor_endosso = moeda(endosso if endosso is not None else 0)
    resultado_tabela = [
        ["Indicador", "Valor"],
        ["VTA atual (Valor Total Atualizado do Contrato)", moeda(vta_atual)],
        ["Garantia total exigida (5%)", moeda(garantia_total)],
        [rotulo_endosso, valor_endosso],
    ]
    tab_res = Table(resultado_tabela, colWidths=[9.5 * cm, 6.5 * cm], repeatRows=1, hAlign="CENTER")
    tab_res.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9E2F3")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ("BACKGROUND", (0, 3), (-1, 3), colors.HexColor("#EAF2F8")),
    ]))
    elementos.append(tab_res)
    elementos.append(Spacer(1, 10))

    elementos.append(Paragraph("4. Comunicação à contratada", h2))
    for bloco in texto_comunicacao.split("\n"):
        if bloco.strip() == "":
            elementos.append(Spacer(1, 4))
        else:
            elementos.append(Paragraph(_p(bloco), corpo_comunicacao))

    doc.build(elementos)
    buffer.seek(0)
    return buffer.getvalue()
