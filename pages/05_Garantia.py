from datetime import datetime, date
from io import BytesIO
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from dateutil.relativedelta import relativedelta

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    REPORTLAB_OK = True
except Exception:
    REPORTLAB_OK = False

st.set_page_config(page_title="Análises de Reajustes - Garantia", layout="wide")


def render_marca_topo():
    """Identidade visual própria do sistema, sem uso de logomarca institucional."""
    st.markdown(
        """
        <style>
        .tlb-cl8us-brand {
            display: inline-flex;
            flex-direction: column;
            gap: 1px;
            margin: 0 0 0.70rem 0;
            padding: 0;
        }
        .tlb-cl8us-brand-main {
            display: flex;
            align-items: baseline;
            gap: 0.45rem;
            line-height: 1.05;
            letter-spacing: -0.02em;
        }
        .tlb-cl8us-tlb {
            color: #123B63;
            font-size: 1.38rem;
            font-weight: 750;
            font-family: "Inter", "Segoe UI", Arial, sans-serif;
        }
        .tlb-cl8us-dot {
            color: #C0842B;
            font-size: 1.18rem;
            font-weight: 700;
        }
        .tlb-cl8us-name {
            color: #0F172A;
            font-size: 1.42rem;
            font-weight: 800;
            font-family: "Consolas", "SFMono-Regular", "Cascadia Mono", "Courier New", monospace;
            letter-spacing: -0.04em;
        }
        .tlb-cl8us-subtitle {
            color: #64748B;
            font-size: 0.74rem;
            font-weight: 500;
            font-family: "Inter", "Segoe UI", Arial, sans-serif;
            margin-top: 0.12rem;
            letter-spacing: 0.01em;
        }
        </style>
        <div class="tlb-cl8us-brand" aria-label="TLB cl8us - apoio à gestão de contratos">
            <div class="tlb-cl8us-brand-main">
                <span class="tlb-cl8us-tlb">TLB</span>
                <span class="tlb-cl8us-dot">·</span>
                <span class="tlb-cl8us-name">cl8us</span>
            </div>
            <div class="tlb-cl8us-subtitle">apoio à gestão de contratos</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# Utilitários
# ============================================================

def moeda(valor, com_prefixo=True):
    try:
        valor = round(float(valor), 2)
    except Exception:
        valor = 0.0
    texto = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {texto}" if com_prefixo else texto


def parse_moeda_br(valor):
    """Aceita entradas como 85771019,12, 85.771.019,12 ou R$ 85.771.019,12."""
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip()
    if not texto:
        return 0.0
    texto = texto.replace("R$", "").replace(" ", "")
    texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except Exception:
        return 0.0


def numero_para_input(valor):
    try:
        return float(valor)
    except Exception:
        return 0.0


def data_hora_brasilia():
    return datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M")


def obter_contexto_valor_global(resultado_valor_global):
    """Extrai, de forma defensiva, os principais valores vindos do módulo Valor Global."""
    resultado_valor_global = resultado_valor_global or {}
    contexto = resultado_valor_global.get("contexto_contratual_anterior", {}) or {}

    valor_original = numero_para_input(resultado_valor_global.get("valor_original_contrato", 0.0))
    valor_total_atualizado = numero_para_input(
        resultado_valor_global.get(
            "valor_atualizado_contrato",
            resultado_valor_global.get("valor_global_contrato", resultado_valor_global.get("valor_global_estoque", 0.0)),
        )
    )
    valor_formalizado_anterior = numero_para_input(
        resultado_valor_global.get(
            "valor_formalizado_anterior",
            contexto.get("valor_formalizado_anterior", 0.0),
        )
    )
    valor_executado_atualizado = numero_para_input(resultado_valor_global.get("valor_executado_atualizado", 0.0))
    remanescente_atualizado = numero_para_input(resultado_valor_global.get("remanescente_reajustado", 0.0))
    total_aditivos = numero_para_input(resultado_valor_global.get("total_aditivos_atualizados", 0.0))
    quantidade_aditivos = int(numero_para_input(resultado_valor_global.get("quantidade_aditivos_total", 0)))
    data_processamento = limpar_texto(resultado_valor_global.get("data_processamento", ""))

    return {
        "valor_original": valor_original,
        "valor_total_atualizado": valor_total_atualizado,
        "valor_formalizado_anterior": valor_formalizado_anterior,
        "valor_executado_atualizado": valor_executado_atualizado,
        "remanescente_atualizado": remanescente_atualizado,
        "total_aditivos": total_aditivos,
        "quantidade_aditivos": quantidade_aditivos,
        "data_processamento": data_processamento,
        "contexto": contexto,
    }


def limpar_texto(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and pd.isna(valor):
        return ""
    return str(valor).strip()


def normalizar_historico(df_hist):
    """Limpa a tabela de histórico e retorna apenas linhas efetivamente preenchidas."""
    if df_hist is None or df_hist.empty:
        return pd.DataFrame(columns=[
            "Data", "Evento", "Valor contratual de referência", "Garantia exigida",
            "Garantia apresentada/endossada", "Observação"
        ])

    df = df_hist.copy()
    colunas = [
        "Data", "Evento", "Valor contratual de referência", "Garantia exigida",
        "Garantia apresentada/endossada", "Observação"
    ]
    for col in colunas:
        if col not in df.columns:
            df[col] = "" if col in ["Data", "Evento", "Observação"] else 0.0

    for col in ["Data", "Evento", "Observação"]:
        df[col] = df[col].apply(limpar_texto)

    for col in ["Valor contratual de referência", "Garantia exigida", "Garantia apresentada/endossada"]:
        df[col] = df[col].apply(parse_moeda_br)

    mask = (
        df["Data"].ne("") |
        df["Evento"].ne("") |
        df["Observação"].ne("") |
        df["Valor contratual de referência"].ne(0) |
        df["Garantia exigida"].ne(0) |
        df["Garantia apresentada/endossada"].ne(0)
    )
    return df.loc[mask, colunas].reset_index(drop=True)


def css():
    st.markdown(
        """
        <style>
        .garantia-card {
            background: #F6F8FA;
            border: 1px solid #E1E6EB;
            border-radius: 14px;
            padding: 18px 20px;
            margin: 6px 0 14px 0;
        }
        .garantia-card-destaque {
            background: #EAF2F8;
            border: 1px solid #C8D9E8;
            border-radius: 14px;
            padding: 20px 22px;
            margin: 8px 0 16px 0;
        }
        .garantia-label {
            color: #475569;
            font-size: 0.92rem;
            margin-bottom: 4px;
        }
        .garantia-valor {
            color: #1F2937;
            font-size: 1.75rem;
            font-weight: 700;
            line-height: 1.2;
        }
        .garantia-valor-destaque {
            color: #123B63;
            font-size: 2.1rem;
            font-weight: 800;
            line-height: 1.2;
        }
        .garantia-nota {
            color: #64748B;
            font-size: 0.9rem;
            margin-top: 6px;
        }
        .valor-formatado-apoio {
            color: #64748B;
            font-size: 0.88rem;
            margin-top: -8px;
            margin-bottom: 8px;
        }
        .historico-garantia-box {
            background: #F4F8FB;
            border: 1px solid #D7E3EE;
            border-radius: 12px;
            padding: 14px 16px;
            margin: 10px 0 8px 0;
        }
        .historico-garantia-titulo {
            color: #123B63;
            font-size: 1rem;
            font-weight: 700;
            margin-bottom: 4px;
        }
        .historico-garantia-texto {
            color: #475569;
            font-size: 0.9rem;
            margin-bottom: 0;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def card(label, valor, nota=None, destaque=False):
    classe = "garantia-card-destaque" if destaque else "garantia-card"
    valor_classe = "garantia-valor-destaque" if destaque else "garantia-valor"
    nota_html = f'<div class="garantia-nota">{nota}</div>' if nota else ""
    st.markdown(
        f"""
        <div class="{classe}">
            <div class="garantia-label">{label}</div>
            <div class="{valor_classe}">{valor}</div>
            {nota_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def montar_texto_instrucao(
    valor_original,
    percentual_garantia,
    garantia_original,
    valor_atualizado,
    garantia_constituida,
    garantia_exigida,
    endosso,
    prazo_dias,
    data_fim_vigencia,
    data_validade_minima,
    historico_usado=False,
    origem_base="Manual",
    valor_formalizado_anterior=0.0,
    valor_executado_atualizado=0.0,
    remanescente_atualizado=0.0,
):
    percentual = percentual_garantia * 100
    valor_formalizado_anterior = numero_para_input(valor_formalizado_anterior)
    valor_executado_atualizado = numero_para_input(valor_executado_atualizado)
    remanescente_atualizado = numero_para_input(remanescente_atualizado)

    if endosso > 0:
        conclusao = (
            f"Considerando a garantia atualmente constituída no valor de {moeda(garantia_constituida)}, "
            f"faz-se necessário o endosso complementar no montante de {moeda(endosso)}."
        )
    else:
        conclusao = (
            f"Considerando a garantia atualmente constituída no valor de {moeda(garantia_constituida)}, "
            "não foi identificada necessidade de endosso complementar, desde que esse valor esteja efetivamente vigente e aceito."
        )

    origem_garantia = (
        "A garantia atualmente constituída considerada nesta análise decorre do histórico de garantias e endossos informado no módulo, contemplando a garantia original e os endossos anteriores registrados."
        if historico_usado else
        "Para fins de conferência, o campo “Garantia atualmente constituída” deve refletir o valor total da garantia já apresentada e aceita, incluindo a garantia original e eventuais endossos anteriores decorrentes de reajustes, repactuações, acréscimos ou outros aditivos contratuais."
    )

    texto_base = (
        "O valor-base utilizado para a nova garantia foi importado do módulo Valor Global como Valor Total Atualizado do Contrato. "
        "Esse valor corresponde à soma da execução atualizada por ciclo com o saldo remanescente atualizado, sem soma autônoma dos aditivos/supressões, os quais permanecem como eventos de governança e rastreabilidade."
        if origem_base == "Valor Global" else
        "O valor-base utilizado para a nova garantia foi informado manualmente neste módulo."
    )

    detalhe_composicao = ""
    if origem_base == "Valor Global" and (valor_executado_atualizado > 0 or remanescente_atualizado > 0):
        detalhe_composicao = (
            f" Para conferência, a composição importada indica execução atualizada por ciclo de {moeda(valor_executado_atualizado)} "
            f"e saldo remanescente atualizado de {moeda(remanescente_atualizado)}."
        )

    referencia_formalizada = ""
    if valor_formalizado_anterior > 0 and abs(valor_formalizado_anterior - valor_atualizado) > 0.01:
        referencia_formalizada = (
            f" Como referência histórica, o valor contratual formalizado antes desta análise consta como {moeda(valor_formalizado_anterior)}. "
            "Esse valor é apresentado para memória administrativa e não substitui o valor-base selecionado para o cálculo da garantia nesta tela."
        )

    return f"""Considerando o valor original do contrato de {moeda(valor_original)}, a garantia contratual original correspondente a {percentual:.2f}% equivale a {moeda(garantia_original)}.

{texto_base}{detalhe_composicao}{referencia_formalizada}

Após a definição do valor-base para garantia em {moeda(valor_atualizado)}, a garantia contratual exigida, calculada pelo mesmo percentual de {percentual:.2f}%, passa a ser de {moeda(garantia_exigida)}.

{conclusao}

{origem_garantia}

Nos termos da Cláusula Décima, a garantia contratual deve ser apresentada no prazo de até {prazo_dias} dias úteis contados do recebimento da convocação pela TELEBRAS, prorrogável por igual período mediante solicitação justificada e aceita pela Gerência de Compras e Contratos.

Ainda nos termos da Cláusula Décima, a validade da garantia, qualquer que seja a modalidade escolhida, deverá abranger período adicional de 3 meses após o término da vigência contratual. Assim, considerando o encerramento da vigência em {data_fim_vigencia.strftime('%d/%m/%Y')}, a garantia deverá permanecer válida, no mínimo, até {data_validade_minima.strftime('%d/%m/%Y')}.
"""


def _paragrafo_tabela(valor, estilo):
    return Paragraph(limpar_texto(valor).replace("&", "&amp;"), estilo)


def gerar_pdf_garantia(dados, texto_instrucao, historico_df=None):
    if not REPORTLAB_OK:
        return None

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.6 * cm,
        leftMargin=1.6 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm,
    )

    styles = getSampleStyleSheet()
    titulo = ParagraphStyle(
        "TituloGarantia",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontSize=13,
        leading=16,
        spaceAfter=8,
    )
    subtitulo = ParagraphStyle(
        "SubtituloGarantia",
        parent=styles["Normal"],
        alignment=TA_CENTER,
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#334155"),
        spaceAfter=14,
    )
    h2 = ParagraphStyle(
        "H2Garantia",
        parent=styles["Heading2"],
        fontSize=10,
        leading=13,
        spaceBefore=8,
        spaceAfter=6,
        textColor=colors.HexColor("#123B63"),
    )
    normal = ParagraphStyle(
        "NormalGarantia",
        parent=styles["Normal"],
        fontSize=8.5,
        leading=12,
        alignment=TA_JUSTIFY,
    )
    celula = ParagraphStyle(
        "CelulaGarantia",
        parent=styles["Normal"],
        fontSize=6.8,
        leading=8,
    )

    elementos = []
    elementos.append(Paragraph("Garantia Contratual", titulo))
    elementos.append(Paragraph("Monitoramento", subtitulo))
    elementos.append(Paragraph(f"Gerado em: {data_hora_brasilia()}", subtitulo))

    elementos.append(Paragraph("1. Memória de cálculo", h2))
    tabela_dados = [
        ["Indicador", "Valor"],
        ["Valor original do contrato", moeda(dados["valor_original"])],
        ["Percentual da garantia", f"{dados['percentual_garantia_pct']:.2f}%".replace(".", ",")],
        ["Garantia original", moeda(dados["garantia_original"])],
        ["Valor-base da garantia", moeda(dados["valor_atualizado"])],
        ["Origem do valor-base", dados.get("origem_base", "Manual")],
        ["Valor formalizado anterior", moeda(dados.get("valor_formalizado_anterior", 0.0)) if dados.get("valor_formalizado_anterior", 0.0) else "Não informado"],
        ["Nova garantia exigida", moeda(dados["garantia_exigida"])],
        ["Garantia atualmente constituída", moeda(dados["garantia_constituida"])],
        ["Endosso necessário", moeda(dados["endosso_necessario"])],
        ["Encerramento da vigência", dados["data_fim_vigencia"].strftime("%d/%m/%Y")],
        ["Validade mínima da garantia", dados["data_validade_minima"].strftime("%d/%m/%Y")],
    ]
    tabela = Table(tabela_dados, colWidths=[8.2 * cm, 7.8 * cm], repeatRows=1, hAlign="CENTER")
    tabela.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9E2F3")),
                ("BACKGROUND", (0, 1), (-1, -1), colors.white),
                ("BACKGROUND", (0, 7), (-1, 7), colors.HexColor("#EAF2F8")),
                ("FONTNAME", (0, 7), (-1, 7), "Helvetica-Bold"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    elementos.append(tabela)
    elementos.append(Spacer(1, 10))

    if historico_df is not None and not historico_df.empty:
        elementos.append(Paragraph("2. Histórico de garantias e endossos considerados", h2))
        dados_hist = [[
            _paragrafo_tabela("Data", celula),
            _paragrafo_tabela("Evento", celula),
            _paragrafo_tabela("Valor ref.", celula),
            _paragrafo_tabela("Garantia exigida", celula),
            _paragrafo_tabela("Apresentada/endossada", celula),
            _paragrafo_tabela("Observação", celula),
        ]]
        for _, row in historico_df.iterrows():
            dados_hist.append([
                _paragrafo_tabela(row.get("Data", ""), celula),
                _paragrafo_tabela(row.get("Evento", ""), celula),
                _paragrafo_tabela(moeda(row.get("Valor contratual de referência", 0.0)), celula),
                _paragrafo_tabela(moeda(row.get("Garantia exigida", 0.0)), celula),
                _paragrafo_tabela(moeda(row.get("Garantia apresentada/endossada", 0.0)), celula),
                _paragrafo_tabela(row.get("Observação", ""), celula),
            ])
        tabela_hist = Table(
            dados_hist,
            colWidths=[1.8 * cm, 3.2 * cm, 2.7 * cm, 2.7 * cm, 3.0 * cm, 2.6 * cm],
            repeatRows=1,
            hAlign="CENTER",
        )
        tabela_hist.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9E2F3")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 3),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        elementos.append(tabela_hist)
        elementos.append(Spacer(1, 10))
        secao_texto = "3. Informações para instrução processual"
    else:
        secao_texto = "2. Informações para instrução processual"

    elementos.append(Paragraph(secao_texto, h2))
    for paragrafo in texto_instrucao.strip().split("\n\n"):
        elementos.append(Paragraph(paragrafo.replace("\n", " "), normal))
        elementos.append(Spacer(1, 5))

    doc.build(elementos)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# Interface
# ============================================================

css()
render_marca_topo()
st.title("Garantia Contratual")
st.write(
    "Este módulo calcula a garantia contratual original, a garantia exigida após atualização do valor do contrato "
    "e o endosso complementar necessário."
)

resultado_valor_global = st.session_state.get("resultado_valor_global", {}) or {}
ctx_valor_global = obter_contexto_valor_global(resultado_valor_global)

valor_original_padrao = ctx_valor_global["valor_original"]
valor_atualizado_padrao = ctx_valor_global["valor_total_atualizado"]
valor_formalizado_anterior_padrao = ctx_valor_global["valor_formalizado_anterior"]
valor_executado_atualizado_padrao = ctx_valor_global["valor_executado_atualizado"]
remanescente_atualizado_padrao = ctx_valor_global["remanescente_atualizado"]

with st.expander("Contexto importado do Valor Global", expanded=True):
    usar_valor_global = False
    if resultado_valor_global:
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.metric("Valor original identificado", moeda(valor_original_padrao))
        with col_b:
            st.metric("Valor Total Atualizado do Contrato", moeda(valor_atualizado_padrao))
        with col_c:
            st.metric("Valor formalizado anterior", moeda(valor_formalizado_anterior_padrao) if valor_formalizado_anterior_padrao > 0 else "Não informado")

        if valor_executado_atualizado_padrao > 0 or remanescente_atualizado_padrao > 0:
            composicao_html = (
                "Composição importada do Valor Global: execução atualizada por ciclo "
                f"{moeda(valor_executado_atualizado_padrao).replace('$', '&#36;')} + saldo remanescente atualizado "
                f"{moeda(remanescente_atualizado_padrao).replace('$', '&#36;')}."
            )
            st.markdown(
                f"<div style='color:#6B7280; font-size:0.92rem; margin-top:4px; margin-bottom:12px;'>{composicao_html}</div>",
                unsafe_allow_html=True,
            )
        if ctx_valor_global.get("quantidade_aditivos", 0):
            st.caption(
                f"Aditivos registrados no Valor Global: {ctx_valor_global.get('quantidade_aditivos', 0)}. "
                "Eles permanecem como eventos de governança/rastreabilidade e não são somados autonomamente ao Valor Total Atualizado."
            )

        usar_valor_global = st.checkbox(
            "Usar o Valor Total Atualizado do Contrato como base da garantia",
            value=True,
            help=(
                "Quando marcado, o módulo Garantia usa o Valor Total Atualizado do Contrato consolidado no módulo Valor Global. "
                "Esse valor corresponde à execução atualizada por ciclo + saldo remanescente atualizado."
            ),
        )
        st.caption(
            "Dados herdados da sessão atual do módulo Valor Global. A opção acima integra a garantia ao valor consolidado da análise."
        )
    else:
        st.info(
            "Não há dados do Valor Global disponíveis na sessão atual. Informe os valores manualmente para calcular a garantia."
        )

st.subheader("Dados para cálculo")
st.caption("Informe valores monetários no padrão brasileiro, por exemplo: 85.771.019,12.")

col1, col2, col3 = st.columns(3)
with col1:
    valor_original_txt = st.text_input(
        "Valor original do contrato",
        value=moeda(valor_original_padrao, com_prefixo=False),
        help="Use ponto para milhares e vírgula para centavos. Ex.: 85.771.019,12",
    )
    valor_original = parse_moeda_br(valor_original_txt)
    st.markdown(f"<div class='valor-formatado-apoio'>{moeda(valor_original)}</div>", unsafe_allow_html=True)

with col2:
    percentual_garantia_pct = st.number_input(
        "Percentual da garantia (%)",
        min_value=0.0,
        max_value=100.0,
        value=5.0,
        step=0.1,
        format="%.2f",
    )

with col3:
    valor_atualizado_txt = st.text_input(
        "Valor-base para cálculo da garantia",
        value=moeda(valor_atualizado_padrao, com_prefixo=False),
        help="Use ponto para milhares e vírgula para centavos. Ex.: 91.609.000,27",
        disabled=bool(resultado_valor_global and usar_valor_global),
    )
    valor_atualizado = valor_atualizado_padrao if (resultado_valor_global and usar_valor_global) else parse_moeda_br(valor_atualizado_txt)
    st.markdown(f"<div class='valor-formatado-apoio'>{moeda(valor_atualizado)}</div>", unsafe_allow_html=True)
    if resultado_valor_global and usar_valor_global:
        st.caption("Valor Total Atualizado importado automaticamente do módulo Valor Global.")

percentual_garantia = percentual_garantia_pct / 100

garantia_original = round(valor_original * percentual_garantia, 2)
garantia_exigida = round(valor_atualizado * percentual_garantia, 2)

default_garantia_constituida = garantia_original

col4, col5, col6 = st.columns(3)
with col4:
    garantia_constituida_txt = st.text_input(
        "Garantia atualmente constituída",
        value=moeda(default_garantia_constituida, com_prefixo=False),
        help=(
            "Informe o valor total da garantia já apresentada e aceita, incluindo garantia original "
            "e eventuais endossos anteriores de reajustes, repactuações ou aditivos."
        ),
    )
    garantia_constituida_manual = parse_moeda_br(garantia_constituida_txt)
    st.markdown(f"<div class='valor-formatado-apoio'>{moeda(garantia_constituida_manual)}</div>", unsafe_allow_html=True)

with col5:
    prazo_dias = st.number_input(
        "Prazo para apresentação/endosso (dias úteis)",
        min_value=1,
        max_value=60,
        value=5,
        step=1,
    )

with col6:
    data_fim_vigencia = st.date_input(
        "Encerramento da vigência contratual",
        value=date.today(),
        format="DD/MM/YYYY",
    )
    data_validade_minima = data_fim_vigencia + relativedelta(months=3)
    st.markdown(
        f"<div class='valor-formatado-apoio'>Validade mínima: {data_validade_minima.strftime('%d/%m/%Y')}</div>",
        unsafe_allow_html=True,
    )

st.info(
    "No campo ‘Garantia atualmente constituída’, considere a garantia original e todos os endossos anteriores já aceitos. "
    "O sistema calculará apenas o endosso complementar necessário para atingir a nova garantia exigida."
)

# ============================================================
# Histórico opcional de garantias e endossos
# ============================================================

st.markdown(
    """
    <div class="historico-garantia-box">
        <div class="historico-garantia-titulo">Histórico de garantias e endossos anteriores</div>
        <p class="historico-garantia-texto">
            Use esta opção quando quiser demonstrar a garantia original e os endossos já apresentados por reajustes, repactuações ou aditivos.
            Se preenchido, o total do histórico substituirá o campo manual de garantia atualmente constituída.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

usar_historico = st.checkbox(
    "Detalhar histórico",
    value=False,
    help="Se marcado, a tabela abaixo será usada para calcular a garantia atualmente constituída.",
)

historico_limpo = pd.DataFrame()
if usar_historico:
    with st.expander("Histórico de Garantias e Endossos", expanded=True):
        st.caption(
            "Preencha os eventos já constituídos/aceitos. A coluna ‘Garantia apresentada/endossada’ será somada para definir a garantia atualmente constituída."
        )
        eventos = [
            "Garantia original",
            "Endosso por reajuste",
            "Endosso por repactuação",
            "Endosso por aditivo",
            "Endosso por acréscimo",
            "Redução por supressão",
            "Outro",
        ]
        historico_base = pd.DataFrame(
            [
                {
                    "Data": "",
                    "Evento": "Garantia original",
                    "Valor contratual de referência": valor_original,
                    "Garantia exigida": garantia_original,
                    "Garantia apresentada/endossada": garantia_original,
                    "Observação": "Garantia original contratual",
                },
                {
                    "Data": "",
                    "Evento": "Endosso por reajuste",
                    "Valor contratual de referência": 0.0,
                    "Garantia exigida": 0.0,
                    "Garantia apresentada/endossada": 0.0,
                    "Observação": "",
                },
                {
                    "Data": "",
                    "Evento": "Endosso por aditivo",
                    "Valor contratual de referência": 0.0,
                    "Garantia exigida": 0.0,
                    "Garantia apresentada/endossada": 0.0,
                    "Observação": "",
                },
            ]
        )
        colunas_monetarias_historico = [
            "Valor contratual de referência",
            "Garantia exigida",
            "Garantia apresentada/endossada",
        ]
        historico_base_editor = historico_base.copy()
        for col_moeda in colunas_monetarias_historico:
            if col_moeda in historico_base_editor.columns:
                historico_base_editor[col_moeda] = historico_base_editor[col_moeda].apply(moeda)

        historico_editado = st.data_editor(
            historico_base_editor,
            hide_index=True,
            use_container_width=True,
            num_rows="dynamic",
            key="garantia_historico_editor",
            column_config={
                "Data": st.column_config.TextColumn("Data", help="Informe no formato dd/mm/aaaa."),
                "Evento": st.column_config.SelectboxColumn("Evento", options=eventos, required=False),
                "Valor contratual de referência": st.column_config.TextColumn(
                    "Valor contratual de referência",
                    help="Informe em formato de moeda. Ex.: R$ 1.000,00.",
                ),
                "Garantia exigida": st.column_config.TextColumn(
                    "Garantia exigida",
                    help="Informe em formato de moeda. Ex.: R$ 50.000,00.",
                ),
                "Garantia apresentada/endossada": st.column_config.TextColumn(
                    "Garantia apresentada/endossada",
                    help="Informe em formato de moeda. Ex.: R$ 50.000,00.",
                ),
                "Observação": st.column_config.TextColumn("Observação"),
            },
        )
        historico_limpo = normalizar_historico(historico_editado)

        if not historico_limpo.empty:
            total_historico = float(historico_limpo["Garantia apresentada/endossada"].sum())
            st.success(f"Garantia atualmente constituída pelo histórico: {moeda(total_historico)}")
        else:
            st.warning("Histórico selecionado, mas sem lançamentos válidos. O sistema usará o valor manual informado acima.")

historico_usado = usar_historico and not historico_limpo.empty
garantia_constituida = round(
    float(historico_limpo["Garantia apresentada/endossada"].sum()) if historico_usado else garantia_constituida_manual,
    2,
)

endosso_necessario = round(max(garantia_exigida - garantia_constituida, 0.0), 2)
excesso_garantia = round(max(garantia_constituida - garantia_exigida, 0.0), 2)

st.divider()
st.subheader("Resultado")

colr1, colr2, colr3 = st.columns(3)
with colr1:
    card("Garantia original", moeda(garantia_original), f"{percentual_garantia_pct:.2f}% sobre o valor original")
with colr2:
    card("Nova garantia exigida", moeda(garantia_exigida), f"{percentual_garantia_pct:.2f}% sobre o valor-base")
with colr3:
    card(
        "Endosso necessário",
        moeda(endosso_necessario),
        "Diferença entre a nova garantia exigida e a garantia atualmente constituída",
        destaque=True,
    )

if historico_usado:
    st.caption(
        f"Garantia atualmente constituída calculada pelo histórico: {moeda(garantia_constituida)}."
    )

if endosso_necessario > 0:
    st.warning(
        f"Será necessário solicitar endosso complementar de {moeda(endosso_necessario)}, observado o prazo contratual de {prazo_dias} dias úteis."
    )
elif excesso_garantia > 0:
    st.success(
        f"A garantia atualmente constituída supera a garantia exigida em {moeda(excesso_garantia)}. Verifique se a garantia vigente está válida e aceita."
    )
else:
    st.success("A garantia atualmente constituída corresponde exatamente à garantia exigida.")

st.subheader("Memória de cálculo")
memoria = [
    {"Indicador": "Valor original do contrato", "Valor": moeda(valor_original)},
    {"Indicador": "Percentual da garantia", "Valor": f"{percentual_garantia_pct:.2f}%".replace(".", ",")},
    {"Indicador": "Garantia original", "Valor": moeda(garantia_original)},
    {"Indicador": "Valor-base da garantia", "Valor": moeda(valor_atualizado)},
    {"Indicador": "Origem do valor-base", "Valor": "Valor Global" if (resultado_valor_global and usar_valor_global) else "Manual"},
    {"Indicador": "Valor formalizado anterior", "Valor": moeda(valor_formalizado_anterior_padrao) if valor_formalizado_anterior_padrao > 0 else "Não informado"},
    {"Indicador": "Execução atualizada por ciclo", "Valor": moeda(valor_executado_atualizado_padrao) if (resultado_valor_global and usar_valor_global) else "Não importado"},
    {"Indicador": "Saldo remanescente atualizado", "Valor": moeda(remanescente_atualizado_padrao) if (resultado_valor_global and usar_valor_global) else "Não importado"},
    {"Indicador": "Nova garantia exigida", "Valor": moeda(garantia_exigida)},
    {"Indicador": "Garantia atualmente constituída", "Valor": moeda(garantia_constituida)},
    {"Indicador": "Origem da garantia constituída", "Valor": "Histórico detalhado" if historico_usado else "Campo manual"},
    {"Indicador": "Endosso necessário", "Valor": moeda(endosso_necessario)},
    {"Indicador": "Encerramento da vigência contratual", "Valor": data_fim_vigencia.strftime("%d/%m/%Y")},
    {"Indicador": "Validade mínima da garantia", "Valor": data_validade_minima.strftime("%d/%m/%Y")},
]
st.dataframe(memoria, use_container_width=True, hide_index=True)

st.subheader("Informações para instrução processual")
texto_instrucao = montar_texto_instrucao(
    valor_original,
    percentual_garantia,
    garantia_original,
    valor_atualizado,
    garantia_constituida,
    garantia_exigida,
    endosso_necessario,
    prazo_dias,
    data_fim_vigencia,
    data_validade_minima,
    historico_usado=historico_usado,
    origem_base="Valor Global" if (resultado_valor_global and usar_valor_global) else "Manual",
    valor_formalizado_anterior=valor_formalizado_anterior_padrao,
    valor_executado_atualizado=valor_executado_atualizado_padrao if (resultado_valor_global and usar_valor_global) else 0.0,
    remanescente_atualizado=remanescente_atualizado_padrao if (resultado_valor_global and usar_valor_global) else 0.0,
)

st.text_area(
    "Texto sugerido",
    value=texto_instrucao,
    height=330,
)

st.subheader("Relatório")
dados_pdf = {
    "valor_original": valor_original,
    "percentual_garantia_pct": percentual_garantia_pct,
    "garantia_original": garantia_original,
    "valor_atualizado": valor_atualizado,
    "origem_base": "Valor Global" if (resultado_valor_global and usar_valor_global) else "Manual",
    "valor_formalizado_anterior": valor_formalizado_anterior_padrao,
    "garantia_exigida": garantia_exigida,
    "garantia_constituida": garantia_constituida,
    "endosso_necessario": endosso_necessario,
    "data_fim_vigencia": data_fim_vigencia,
    "data_validade_minima": data_validade_minima,
}

st.session_state["resultado_garantia"] = {
    "valor_original": valor_original,
    "valor_atualizado_base": valor_atualizado,
    "valor_total_atualizado_contrato": valor_atualizado,
    "valor_formalizado_anterior": valor_formalizado_anterior_padrao,
    "valor_executado_atualizado": valor_executado_atualizado_padrao if (resultado_valor_global and usar_valor_global) else 0.0,
    "remanescente_atualizado": remanescente_atualizado_padrao if (resultado_valor_global and usar_valor_global) else 0.0,
    "percentual_garantia": percentual_garantia,
    "garantia_original": garantia_original,
    "garantia_exigida": garantia_exigida,
    "garantia_constituida": garantia_constituida,
    "endosso_necessario": endosso_necessario,
    "origem_valor_base": "Valor Global" if (resultado_valor_global and usar_valor_global) else "Manual",
    "data_fim_vigencia": data_fim_vigencia,
    "data_validade_minima": data_validade_minima,
}

if REPORTLAB_OK:
    pdf_bytes = gerar_pdf_garantia(
        dados_pdf,
        texto_instrucao,
        historico_df=historico_limpo if historico_usado else None,
    )
    st.session_state["arquivo_garantia_pdf"] = pdf_bytes
    st.download_button(
        "Baixar Relatório de Garantia em PDF",
        data=pdf_bytes,
        file_name="relatorio_garantia_contratual.pdf",
        mime="application/pdf",
        type="primary",
        use_container_width=False,
    )
else:
    st.error("A biblioteca reportlab não está disponível. Inclua reportlab no requirements.txt para habilitar o PDF.")
