from datetime import datetime, date
from io import BytesIO
from zoneinfo import ZoneInfo

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


# ============================================================
# Utilitários
# ============================================================

def moeda(valor, com_prefixo=True):
    try:
        valor = float(valor)
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
):
    percentual = percentual_garantia * 100
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

    return f"""Considerando o valor original do contrato de {moeda(valor_original)}, a garantia contratual original correspondente a {percentual:.2f}% equivale a {moeda(garantia_original)}.

Após a atualização do valor contratual para {moeda(valor_atualizado)}, a garantia contratual exigida, calculada pelo mesmo percentual de {percentual:.2f}%, passa a ser de {moeda(garantia_exigida)}.

{conclusao}

Para fins de conferência, o campo “Garantia atualmente constituída” deve refletir o valor total da garantia já apresentada e aceita, incluindo a garantia original e eventuais endossos anteriores decorrentes de reajustes, repactuações, acréscimos ou outros aditivos contratuais.

Nos termos da Cláusula Décima, a garantia contratual deve ser apresentada no prazo de até {prazo_dias} dias úteis contados do recebimento da convocação pela TELEBRAS, prorrogável por igual período mediante solicitação justificada e aceita pela Gerência de Compras e Contratos.

Ainda nos termos da Cláusula Décima, a validade da garantia, qualquer que seja a modalidade escolhida, deverá abranger período adicional de 3 meses após o término da vigência contratual. Assim, considerando o encerramento da vigência em {data_fim_vigencia.strftime('%d/%m/%Y')}, a garantia deverá permanecer válida, no mínimo, até {data_validade_minima.strftime('%d/%m/%Y')}.
"""


def gerar_pdf_garantia(dados, texto_instrucao):
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

    elementos = []
    elementos.append(Paragraph("TELEBRAS - Análise de Garantia Contratual", titulo))
    elementos.append(Paragraph("Relatório Executivo - GCC", subtitulo))
    elementos.append(Paragraph(f"Gerado em: {data_hora_brasilia()}", subtitulo))

    elementos.append(Paragraph("1. Memória de cálculo", h2))
    tabela_dados = [
        ["Indicador", "Valor"],
        ["Valor original do contrato", moeda(dados["valor_original"])],
        ["Percentual da garantia", f"{dados['percentual_garantia_pct']:.2f}%".replace(".", ",")],
        ["Garantia original", moeda(dados["garantia_original"])],
        ["Valor atualizado do contrato", moeda(dados["valor_atualizado"])],
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

    elementos.append(Paragraph("2. Informações para instrução processual", h2))
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

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Garantia Contratual")
st.write(
    "Este módulo calcula a garantia contratual original, a garantia exigida após atualização do valor do contrato "
    "e o endosso complementar necessário."
)

resultado_valor_global = st.session_state.get("resultado_valor_global", {}) or {}

valor_original_padrao = numero_para_input(resultado_valor_global.get("valor_original_contrato", 0.0))
valor_atualizado_padrao = numero_para_input(
    resultado_valor_global.get(
        "valor_atualizado_contrato",
        resultado_valor_global.get("valor_global_contrato", resultado_valor_global.get("valor_global_estoque", 0.0)),
    )
)

with st.expander("Contexto importado do Valor Global", expanded=True):
    if resultado_valor_global:
        col_a, col_b = st.columns(2)
        with col_a:
            st.metric("Valor original identificado", moeda(valor_original_padrao))
        with col_b:
            st.metric("Valor atualizado identificado", moeda(valor_atualizado_padrao))
        st.caption("Dados herdados da sessão atual do módulo Valor Global. Os campos abaixo permanecem editáveis para conferência.")
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
        "Valor atualizado do contrato",
        value=moeda(valor_atualizado_padrao, com_prefixo=False),
        help="Use ponto para milhares e vírgula para centavos. Ex.: 91.609.000,27",
    )
    valor_atualizado = parse_moeda_br(valor_atualizado_txt)
    st.markdown(f"<div class='valor-formatado-apoio'>{moeda(valor_atualizado)}</div>", unsafe_allow_html=True)

percentual_garantia = percentual_garantia_pct / 100

garantia_original = valor_original * percentual_garantia
garantia_exigida = valor_atualizado * percentual_garantia

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
    garantia_constituida = parse_moeda_br(garantia_constituida_txt)
    st.markdown(f"<div class='valor-formatado-apoio'>{moeda(garantia_constituida)}</div>", unsafe_allow_html=True)

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

endosso_necessario = max(garantia_exigida - garantia_constituida, 0.0)
excesso_garantia = max(garantia_constituida - garantia_exigida, 0.0)

st.divider()
st.subheader("Resultado")

colr1, colr2, colr3 = st.columns(3)
with colr1:
    card("Garantia original", moeda(garantia_original), f"{percentual_garantia_pct:.2f}% sobre o valor original")
with colr2:
    card("Nova garantia exigida", moeda(garantia_exigida), f"{percentual_garantia_pct:.2f}% sobre o valor atualizado")
with colr3:
    card(
        "Endosso necessário",
        moeda(endosso_necessario),
        "Diferença entre a nova garantia exigida e a garantia atualmente constituída",
        destaque=True,
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
    {"Indicador": "Valor atualizado do contrato", "Valor": moeda(valor_atualizado)},
    {"Indicador": "Nova garantia exigida", "Valor": moeda(garantia_exigida)},
    {"Indicador": "Garantia atualmente constituída", "Valor": moeda(garantia_constituida)},
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
    "garantia_exigida": garantia_exigida,
    "garantia_constituida": garantia_constituida,
    "endosso_necessario": endosso_necessario,
    "data_fim_vigencia": data_fim_vigencia,
    "data_validade_minima": data_validade_minima,
}

if REPORTLAB_OK:
    pdf_bytes = gerar_pdf_garantia(dados_pdf, texto_instrucao)
    st.download_button(
        "Baixar Relatório de Garantia em PDF",
        data=pdf_bytes,
        file_name="relatorio_garantia_contratual.pdf",
        mime="application/pdf",
    )
else:
    st.error("A biblioteca reportlab não está disponível. Inclua reportlab no requirements.txt para habilitar o PDF.")
