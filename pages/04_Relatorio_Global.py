from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Relatório Global", layout="wide")


# ============================================================
# Formatação
# ============================================================

def moeda(valor):
    try:
        valor = float(valor)
    except Exception:
        valor = 0.0
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def percentual(valor, casas=2):
    try:
        valor = float(valor)
    except Exception:
        valor = 0.0
    return f"{valor * 100:.{casas}f}%".replace(".", ",")


def fator_fmt(valor):
    try:
        valor = float(valor)
    except Exception:
        valor = 1.0
    return f"{valor:.4f}".replace(".", ",")


def normalizar_status(status):
    texto = str(status or "").upper()
    if "PRECLUS" in texto:
        return "PRECLUSO"
    if "RESSALVA" in texto:
        return "ADMISSÍVEL COM RESSALVA"
    if "TEMPEST" in texto:
        return "TEMPESTIVO"
    if "ADIANT" in texto:
        return "ADIANTADO"
    return texto or "NÃO INFORMADO"


def texto_clausula_oito(adm):
    ciclos = []
    if adm:
        ciclos = adm.get("ciclos") or adm.get("detalhamento_ciclos") or []

    status_gerais = []
    for c in ciclos:
        status_gerais.append(normalizar_status(c.get("situacao") or c.get("Situação") or c.get("status")))

    texto_status = " ".join(status_gerais)

    if "PRECLUSO" in texto_status:
        return (
            "A análise registra ciclo classificado como precluso, em razão da ausência de solicitação "
            "dentro do prazo de 90 dias previsto no Parágrafo Quinto da Cláusula Oitava, observando-se "
            "também a regra do Parágrafo Sétimo quanto à possibilidade de novo pleito após ultrapassados "
            "12 meses da data em que poderia ter sido requerido."
        )

    if "ADMISSÍVEL COM RESSALVA" in texto_status:
        return (
            "A análise registra pleito admissível com ressalva, por ter sido apresentado no mesmo mês de "
            "implemento da anualidade, porém antes do dia exato de completude dos 12 meses. Os efeitos "
            "financeiros devem observar os Parágrafos Primeiro e Segundo da Cláusula Oitava."
        )

    if "ADIANTADO" in texto_status:
        return (
            "A análise registra pleito apresentado antes do implemento da anualidade contratual. Nos termos "
            "dos Parágrafos Primeiro e Segundo da Cláusula Oitava, eventual reconhecimento de efeitos "
            "financeiros deve observar a completude dos 12 meses e a data juridicamente apta para o pedido."
        )

    return (
        "O pleito foi classificado como tempestivo, considerando a anualidade prevista no Parágrafo Primeiro "
        "da Cláusula Oitava e a apresentação da solicitação dentro da janela contratual de 90 dias prevista "
        "no Parágrafo Quinto. Os efeitos financeiros devem observar o Parágrafo Segundo da mesma cláusula."
    )


def dataframe_para_texto(df, max_linhas=20):
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return "Sem dados disponíveis."
    return df.head(max_linhas).to_string(index=False)


def gerar_minuta_sei(adm, res):
    origem = (adm or {}).get("origem") or (adm or {}).get("tipo") or "Não informado"
    indice = res.get("indice", (adm or {}).get("indice", "Não informado"))
    fator = res.get("fator_acumulado", (adm or {}).get("fator_acumulado", (adm or {}).get("fator", 1.0)))

    texto = f"""
RELATÓRIO EXECUTIVO DE VALOR GLOBAL DO CONTRATO

1. Contexto da análise

A presente análise consolida os resultados da etapa de admissibilidade do reajuste contratual e da etapa de quantificação do impacto financeiro, com base nos dados constantes do Arquivo de Coleta preenchido.

Origem da análise: {origem}
Índice utilizado: {indice}
Fator acumulado aplicado: {fator_fmt(fator)}

2. Fundamentação contratual

{texto_clausula_oito(adm)}

3. Resultado financeiro consolidado

Valor original do contrato: {moeda(res.get("valor_original_contrato", 0))}
Total pago/faturado: {moeda(res.get("total_pago_faturado", 0))}
Total devido reajustado: {moeda(res.get("total_devido_reajustado", 0))}
Delta acumulado financeiro: {moeda(res.get("delta_acumulado", 0))}
Remanescente original: {moeda(res.get("remanescente_original", 0))}
Remanescente atualizado: {moeda(res.get("remanescente_reajustado", 0))}

4. Valor Global

Valor Global Financeiro: {moeda(res.get("valor_global_financeiro", 0))}
Valor Global por Estoque/Itens: {moeda(res.get("valor_global_estoque", 0))}
Diferença entre métodos: {moeda(res.get("diferenca_metodos", 0))}
Percentual de divergência: {percentual(res.get("percentual_divergencia", 0))}

5. Observação executiva

A visão financeira considera os valores pagos ou faturados por ciclo e pode refletir glosas, descontos, retenções, pagamentos parciais ou diferenças de competência. A visão por estoque/itens considera as quantidades e remanescentes físicos/contratuais, aplicando os fatores de reajuste aos saldos remanescentes informados. Eventual diferença entre as duas visões deve ser analisada pela fiscalização e pela gestão contratual.
"""
    return texto.strip()


# ============================================================
# PDF com ReportLab
# ============================================================

def criar_pdf_relatorio(adm, res):
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    except Exception as exc:
        raise RuntimeError(
            "A biblioteca reportlab não está instalada. Inclua 'reportlab' no requirements.txt."
        ) from exc

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.5 * cm,
        leftMargin=1.5 * cm,
        topMargin=1.3 * cm,
        bottomMargin=1.3 * cm,
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="TituloTelebras",
        parent=styles["Title"],
        fontSize=15,
        leading=18,
        spaceAfter=10,
    ))
    styles.add(ParagraphStyle(
        name="Subtitulo",
        parent=styles["Heading2"],
        fontSize=11,
        leading=14,
        spaceBefore=8,
        spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        name="Texto",
        parent=styles["BodyText"],
        fontSize=8.8,
        leading=11,
        alignment=4,
    ))

    story = []

    logo_paths = [
        Path("assets/telebras_logo.png"),
        Path("assets/logo.png"),
        Path("telebras_logo.png"),
        Path("logo.png"),
    ]
    logo_adicionado = False
    for caminho in logo_paths:
        if caminho.exists():
            try:
                story.append(Image(str(caminho), width=4.0 * cm, height=1.2 * cm))
                logo_adicionado = True
                break
            except Exception:
                pass

    if not logo_adicionado:
        story.append(Paragraph("TELEBRAS — Ferramenta de Análise de Reajuste Contratual", styles["TituloTelebras"]))

    story.append(Paragraph("Relatório Executivo de Valor Global do Contrato", styles["TituloTelebras"]))

    story.append(Paragraph("1. Identificação da Análise", styles["Subtitulo"]))
    origem = (adm or {}).get("origem") or (adm or {}).get("tipo") or "Não informado"
    indice = res.get("indice", (adm or {}).get("indice", "Não informado"))
    fator = res.get("fator_acumulado", (adm or {}).get("fator_acumulado", (adm or {}).get("fator", 1.0)))

    tabela_ident = [
        ["Origem da análise", origem],
        ["Índice aplicado", indice],
        ["Fator acumulado", fator_fmt(fator)],
        ["Data de processamento", res.get("data_processamento", "Não informado")],
    ]
    story.append(tabela_pdf(tabela_ident, col_widths=[6 * cm, 10 * cm]))
    story.append(Spacer(1, 8))

    story.append(Paragraph("2. Análise Contratual — Cláusula Oitava", styles["Subtitulo"]))
    story.append(Paragraph(texto_clausula_oito(adm), styles["Texto"]))

    story.append(Paragraph("3. Indicadores Executivos", styles["Subtitulo"]))
    indicadores = [
        ["Indicador", "Valor"],
        ["Valor original do contrato", moeda(res.get("valor_original_contrato", 0))],
        ["Total pago/faturado", moeda(res.get("total_pago_faturado", 0))],
        ["Total devido reajustado", moeda(res.get("total_devido_reajustado", 0))],
        ["Delta acumulado", moeda(res.get("delta_acumulado", 0))],
        ["Remanescente original", moeda(res.get("remanescente_original", 0))],
        ["Remanescente atualizado", moeda(res.get("remanescente_reajustado", 0))],
        ["Valor Global Financeiro", moeda(res.get("valor_global_financeiro", 0))],
        ["Valor Global por Estoque/Itens", moeda(res.get("valor_global_estoque", 0))],
        ["Diferença entre métodos", moeda(res.get("diferenca_metodos", 0))],
        ["Percentual de divergência", percentual(res.get("percentual_divergencia", 0))],
    ]
    story.append(tabela_pdf(indicadores, header=True, col_widths=[8 * cm, 8 * cm]))

    story.append(Paragraph("4. Quadro Comparativo — Antes x Depois", styles["Subtitulo"]))
    df_comp = res.get("df_comparativo")
    story.append(tabela_dataframe_pdf(df_comp, max_linhas=12))

    story.append(Paragraph("5. Financeiro por Ciclo", styles["Subtitulo"]))
    story.append(tabela_dataframe_pdf(res.get("df_financeiro_por_ciclo"), max_linhas=15))

    story.append(Paragraph("6. Remanescentes por Ciclo", styles["Subtitulo"]))
    story.append(tabela_dataframe_pdf(res.get("df_remanescentes"), max_linhas=15))

    story.append(Paragraph("7. Validação do Gestor/Fiscal do Contrato", styles["Subtitulo"]))
    assinatura = [
        ["Nome:", ""],
        ["Matrícula:", ""],
        ["Unidade:", ""],
        ["Data:", ""],
        ["Assinatura:", ""],
    ]
    story.append(tabela_pdf(assinatura, col_widths=[4 * cm, 12 * cm]))

    doc.build(story)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf


def tabela_pdf(dados, header=False, col_widths=None):
    from reportlab.lib import colors
    from reportlab.platypus import Table, TableStyle

    table = Table(dados, colWidths=col_widths)
    estilo = [
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]
    if header:
        estilo.extend([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E9EEF6")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ])
    table.setStyle(TableStyle(estilo))
    return table


def tabela_dataframe_pdf(df, max_linhas=12):
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return tabela_pdf([["Informação", "Sem dados disponíveis"]], col_widths=[5 * 2.83, 11 * 2.83])

    dados = [list(df.columns)]
    for _, row in df.head(max_linhas).iterrows():
        linha = []
        for valor in row.tolist():
            if isinstance(valor, float):
                # Fatores ficam menores; valores monetários já podem ter sido formatados na tela.
                linha.append(f"{valor:,.4f}".replace(",", "X").replace(".", ",").replace("X", "."))
            else:
                linha.append(str(valor))
        dados.append(linha)

    # Largura simples distribuída.
    return tabela_pdf(dados, header=True)


def df_visual(df, moeda_cols=None, fator_cols=None, pct_cols=None):
    if df is None or not isinstance(df, pd.DataFrame):
        return pd.DataFrame()
    visual = df.copy()
    for col in moeda_cols or []:
        if col in visual.columns:
            visual[col] = visual[col].apply(moeda)
    for col in fator_cols or []:
        if col in visual.columns:
            visual[col] = visual[col].apply(fator_fmt)
    for col in pct_cols or []:
        if col in visual.columns:
            visual[col] = visual[col].apply(percentual)
    return visual


# ============================================================
# Interface
# ============================================================

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Relatório Global")

adm = st.session_state.get("dados_admissibilidade")
res = st.session_state.get("resultado_valor_global")

if not res:
    st.warning(
        "Ainda não há dados processados para o Relatório Global. "
        "Acesse o módulo Valor Global, carregue o Arquivo de Coleta e processe a análise."
    )
    st.stop()

st.subheader("Resumo Executivo")

col1, col2, col3, col4 = st.columns(4)
col1.metric("Índice", res.get("indice", "Não informado"))
col2.metric("Fator acumulado", fator_fmt(res.get("fator_acumulado", 1.0)))
col3.metric("Valor Global Financeiro", moeda(res.get("valor_global_financeiro", 0)))
col4.metric("Delta acumulado", moeda(res.get("delta_acumulado", 0)))

col5, col6, col7 = st.columns(3)
col5.metric("Valor Global por Estoque", moeda(res.get("valor_global_estoque", 0)))
col6.metric("Diferença entre métodos", moeda(res.get("diferenca_metodos", 0)))
col7.metric("Divergência", percentual(res.get("percentual_divergencia", 0)))

st.divider()

tab1, tab2, tab3 = st.tabs(["Relatório Executivo", "Tabelas Comparativas", "PDF"])

with tab1:
    st.markdown("### Fundamentação Contratual")
    st.info(texto_clausula_oito(adm))

    st.markdown("### Minuta para Nota Técnica / SEI")
    texto_sei = gerar_minuta_sei(adm, res)
    st.text_area("Copie para o SEI:", texto_sei, height=420)

with tab2:
    st.markdown("### Quadro Comparativo — Antes x Depois")
    st.dataframe(
        df_visual(
            res.get("df_comparativo"),
            moeda_cols=["Antes do Reajuste", "Após Reajuste", "Diferença"],
        ),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("### Financeiro por Ciclo")
    st.dataframe(
        df_visual(
            res.get("df_financeiro_por_ciclo"),
            moeda_cols=["Valor pago/faturado", "Valor devido reajustado", "Delta do ciclo", "Delta acumulado"],
            fator_cols=["Fator aplicado"],
        ),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("### Remanescentes")
    st.dataframe(
        df_visual(
            res.get("df_remanescentes"),
            moeda_cols=["Remanescente original", "Remanescente reajustado"],
            fator_cols=["Fator aplicado"],
        ),
        use_container_width=True,
        hide_index=True,
    )

with tab3:
    st.markdown("### Baixar Relatório Executivo em PDF")
    try:
        pdf_bytes = criar_pdf_relatorio(adm, res)
        st.download_button(
            label="Baixar Relatório Executivo em PDF",
            data=pdf_bytes,
            file_name="relatorio_global_valor_contrato.pdf",
            mime="application/pdf",
            type="primary",
        )
    except Exception as exc:
        st.error(f"Não foi possível gerar o PDF: {exc}")
        st.caption("Verifique se a biblioteca reportlab foi incluída no requirements.txt.")
