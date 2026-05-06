from io import BytesIO
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Análises de Reajustes - Relatório Global", layout="wide")


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
        return "▲ PRECLUSO"
    if "RESSALVA" in texto:
        return "● ADMISSÍVEL COM RESSALVA"
    if "TEMPEST" in texto:
        return "■ TEMPESTIVO"
    if "ADIANT" in texto:
        return "▲ ADIANTADO"
    return texto or "NÃO INFORMADO"


def texto_clausula_oito(adm):
    ciclos = []
    if adm:
        ciclos = adm.get("ciclos") or adm.get("detalhamento_ciclos") or []

    status_gerais = []
    for c in ciclos:
        status_gerais.append(str(c.get("situacao") or c.get("Situação") or c.get("status") or "").upper())
    texto_status = " ".join(status_gerais)

    if "PRECLUS" in texto_status:
        return (
            "A análise registra ciclo classificado como precluso, em razão da ausência de solicitação "
            "dentro do prazo de 90 dias previsto no Parágrafo Quinto da Cláusula Oitava, observando-se "
            "também a regra do Parágrafo Sétimo quanto à possibilidade de novo pleito após ultrapassados "
            "12 meses da data em que poderia ter sido requerido."
        )
    if "RESSALVA" in texto_status:
        return (
            "A análise registra pleito admissível com ressalva, por ter sido apresentado no mesmo mês de "
            "implemento da anualidade, porém antes do dia exato de completude dos 12 meses. Os efeitos "
            "financeiros devem observar os Parágrafos Primeiro e Segundo da Cláusula Oitava."
        )
    if "ADIANT" in texto_status:
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


def gerar_texto_instrucao(adm, res):
    origem = (adm or {}).get("origem") or (adm or {}).get("tipo") or "Não informado"
    indice = res.get("indice", (adm or {}).get("indice", "Não informado"))
    fator = res.get("fator_acumulado", (adm or {}).get("fator_acumulado", (adm or {}).get("fator", 1.0)))
    return f"""
RELATÓRIO EXECUTIVO — VALOR ATUALIZADO DO CONTRATO

1. Contexto da análise

A presente análise consolida os resultados da etapa de admissibilidade do reajuste contratual e da etapa de quantificação do impacto financeiro, com base nos dados constantes do Arquivo de Coleta preenchido.

Origem da análise: {origem}
Índice utilizado: {indice}
Fator acumulado considerado: {fator_fmt(fator)}
Data/hora de geração: {datetime.now(ZoneInfo('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M')}

2. Fundamentação contratual

{texto_clausula_oito(adm)}

3. Resultado financeiro consolidado

Valor original do contrato: {moeda(res.get('valor_original_contrato', 0))}
Valor pago efetivo: {moeda(res.get('total_pago_faturado', 0))}
Valor teórico calculado: {moeda(res.get('total_devido_reajustado', 0))}
Valor represado a pagar: {moeda(res.get('valor_represado_a_pagar', 0))}
Delta total: {moeda(res.get('delta_acumulado', 0))}
Saldo remanescente atualizado: {moeda(res.get('remanescente_reajustado', 0))}
Aditivos/Supressões atualizados: {moeda(res.get('total_aditivos_atualizados', 0))}
Valor atualizado do contrato: {moeda(res.get('valor_atualizado_contrato', res.get('valor_global_estoque', 0)))}

4. Observação executiva

Os ciclos classificados como preclusos permanecem registrados para fins de memória e rastreabilidade, mas não compõem o impacto financeiro do reajuste acumulado nem o valor represado a pagar. O valor atualizado do contrato reflete a composição financeira consolidada a partir da execução, do saldo remanescente atualizado e dos aditivos ou supressões informados, quando aplicável.
""".strip()


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
    if "Situação" in visual.columns:
        visual["Situação"] = visual["Situação"].apply(normalizar_status)
    return visual


def criar_pdf_relatorio(adm, res):
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle, PageBreak
    except Exception as exc:
        raise RuntimeError("A biblioteca reportlab não está instalada. Inclua 'reportlab' no requirements.txt.") from exc

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.05 * cm,
        leftMargin=1.05 * cm,
        topMargin=1.0 * cm,
        bottomMargin=1.0 * cm,
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Titulo", parent=styles["Title"], fontSize=13, leading=16, alignment=1, spaceAfter=6))
    styles.add(ParagraphStyle(name="Subtitulo", parent=styles["Heading2"], fontSize=10, leading=12, spaceBefore=6, spaceAfter=4))
    styles.add(ParagraphStyle(name="Texto", parent=styles["BodyText"], fontSize=8.2, leading=10.5, alignment=4))
    styles.add(ParagraphStyle(name="Celula", parent=styles["BodyText"], fontSize=6.8, leading=8.2))
    styles.add(ParagraphStyle(name="CelulaCab", parent=styles["BodyText"], fontSize=6.8, leading=8.2, fontName="Helvetica-Bold"))

    story = []
    story.append(Paragraph("TELEBRAS — Análise de Reajuste Contratual", styles["Titulo"]))
    story.append(Paragraph("Relatório Executivo - GCC", styles["Titulo"]))
    story.append(Spacer(1, 6))

    origem = (adm or {}).get("origem") or (adm or {}).get("tipo") or "Não informado"
    indice = res.get("indice", (adm or {}).get("indice", "Não informado"))
    fator = res.get("fator_acumulado", (adm or {}).get("fator_acumulado", (adm or {}).get("fator", 1.0)))

    story.append(Paragraph("1. Identificação da Análise", styles["Subtitulo"]))
    story.append(tabela_pdf([
        ["Origem da análise", origem],
        ["Índice aplicado", indice],
        ["Fator acumulado", fator_fmt(fator)],
        ["Data de geração", datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M")],
    ], col_widths=[6 * cm, 11 * cm]))

    story.append(Paragraph("2. Fundamentação Contratual — Cláusula Oitava", styles["Subtitulo"]))
    story.append(Paragraph(texto_clausula_oito(adm), styles["Texto"]))

    story.append(Paragraph("3. Indicadores Executivos", styles["Subtitulo"]))
    story.append(tabela_pdf([
        ["Indicador", "Valor"],
        ["Valor original", moeda(res.get("valor_original_contrato", 0))],
        ["Valor pago efetivo", moeda(res.get("total_pago_faturado", 0))],
        ["Valor teórico calculado", moeda(res.get("total_devido_reajustado", 0))],
        ["Valor represado a pagar", moeda(res.get("valor_represado_a_pagar", 0))],
        ["Delta total", moeda(res.get("delta_acumulado", 0))],
        ["Saldo remanescente atualizado", moeda(res.get("remanescente_reajustado", 0))],
        ["Aditivos/Supressões atualizados", moeda(res.get("total_aditivos_atualizados", 0))],
        ["Valor atualizado do contrato", moeda(res.get("valor_atualizado_contrato", res.get("valor_global_estoque", 0)))],
    ], header=True, col_widths=[8.5 * cm, 8.5 * cm]))

    story.append(Paragraph("4. Ciclos, Percentuais e Efeitos Financeiros", styles["Subtitulo"]))
    df_ciclos = res.get("df_ciclos")
    if isinstance(df_ciclos, pd.DataFrame) and not df_ciclos.empty:
        cols = [c for c in ["Ciclo", "Data-base", "Data do pedido", "Situação", "Variação", "Fator acumulado efetivo"] if c in df_ciclos.columns]
        df_c = df_ciclos[cols].copy()
        if "Situação" in df_c.columns:
            df_c["Situação"] = df_c["Situação"].apply(normalizar_status)
        if "Variação" in df_c.columns:
            df_c["Variação"] = df_c["Variação"].apply(lambda x: percentual(x, 2))
        if "Fator acumulado efetivo" in df_c.columns:
            df_c["Fator acumulado efetivo"] = df_c["Fator acumulado efetivo"].apply(fator_fmt)
        story.append(tabela_dataframe_pdf(df_c, max_linhas=12))

    story.append(Paragraph("5. Financeiro por Ciclo", styles["Subtitulo"]))
    df_fin = df_visual(
        res.get("df_financeiro_por_ciclo"),
        moeda_cols=["Valor pago/faturado", "Valor devido reajustado", "Delta do ciclo", "Delta acumulado"],
        fator_cols=["Fator aplicado"],
    )
    keep_fin = [c for c in ["Ciclo", "Situação", "Fator aplicado", "Valor pago/faturado", "Valor devido reajustado", "Delta do ciclo"] if c in df_fin.columns]
    story.append(tabela_dataframe_pdf(df_fin[keep_fin] if keep_fin else df_fin, max_linhas=20))

    story.append(Paragraph("6. Valor Atualizado do Contrato", styles["Subtitulo"]))
    story.append(tabela_pdf([
        ["Componente", "Valor"],
        ["Valor executado atualizado", moeda(res.get("valor_executado_atualizado", 0))],
        ["Saldo remanescente atualizado", moeda(res.get("remanescente_reajustado", 0))],
        ["Aditivos/Supressões atualizados", moeda(res.get("total_aditivos_atualizados", 0))],
        ["Valor atualizado do contrato", moeda(res.get("valor_atualizado_contrato", res.get("valor_global_estoque", 0)))],
    ], header=True, col_widths=[10 * cm, 7 * cm]))

    df_ad = res.get("df_aditivos")
    if isinstance(df_ad, pd.DataFrame) and not df_ad.empty:
        story.append(Paragraph("7. Aditivos e Supressões", styles["Subtitulo"]))
        df_adv = df_visual(
            df_ad,
            moeda_cols=["Valor original da alteração", "Valor atualizado da alteração"],
            fator_cols=["Fator aplicado"],
        )
        keep_ad = [c for c in ["Item", "Data do aditivo", "Ciclo/Marco", "Tipo de alteração", "Valor atualizado da alteração"] if c in df_adv.columns]
        story.append(tabela_dataframe_pdf(df_adv[keep_ad], max_linhas=12))

    story.append(Paragraph("8. Informações para instrução processual", styles["Subtitulo"]))
    story.append(Paragraph(gerar_texto_instrucao(adm, res).replace("\n", "<br/>"), styles["Texto"]))

    doc.build(story)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf


def tabela_pdf(dados, header=False, col_widths=None):
    from reportlab.lib import colors
    from reportlab.platypus import Table, TableStyle

    table = Table(dados, colWidths=col_widths, hAlign="CENTER")
    estilo = [
        ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.4),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]
    if header:
        estilo.extend([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E9EEF6")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ])
    table.setStyle(TableStyle(estilo))
    return table


def tabela_dataframe_pdf(df, max_linhas=12):
    from reportlab.platypus import Paragraph
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm

    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return tabela_pdf([["Informação", "Sem dados disponíveis"]], col_widths=[5 * cm, 11 * cm])

    styles = getSampleStyleSheet()
    cel = ParagraphStyle("Cel", parent=styles["BodyText"], fontSize=6.6, leading=8.0)
    cab = ParagraphStyle("Cab", parent=styles["BodyText"], fontSize=6.6, leading=8.0, fontName="Helvetica-Bold")
    dados = [[Paragraph(str(c), cab) for c in df.columns]]
    for _, row in df.head(max_linhas).iterrows():
        linha = []
        for valor in row.tolist():
            linha.append(Paragraph(str(valor), cel))
        dados.append(linha)

    ncols = max(len(df.columns), 1)
    largura_total = 17.0 * cm
    col_widths = [largura_total / ncols] * ncols
    return tabela_pdf(dados, header=True, col_widths=col_widths)


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
col3.metric("Valor atualizado do contrato", moeda(res.get("valor_atualizado_contrato", res.get("valor_global_estoque", 0))))
col4.metric("Valor represado a pagar", moeda(res.get("valor_represado_a_pagar", 0)))

col5, col6, col7 = st.columns(3)
col5.metric("Valor pago efetivo", moeda(res.get("total_pago_faturado", 0)))
col6.metric("Valor teórico calculado", moeda(res.get("total_devido_reajustado", 0)))
col7.metric("Delta total", moeda(res.get("delta_acumulado", 0)))

st.divider()

tab1, tab2, tab3 = st.tabs(["Relatório Executivo", "Tabelas", "PDF"])

with tab1:
    st.markdown("### Fundamentação Contratual")
    st.info(texto_clausula_oito(adm))

    st.markdown("### Informações para instrução processual")
    texto = gerar_texto_instrucao(adm, res)
    st.text_area("Copie para o processo:", texto, height=420)

with tab2:
    st.markdown("### Quadro Executivo")
    st.dataframe(
        df_visual(res.get("df_comparativo"), moeda_cols=["Antes do Reajuste", "Após Reajuste", "Diferença"]),
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

    df_ad = res.get("df_aditivos")
    if isinstance(df_ad, pd.DataFrame) and not df_ad.empty:
        st.markdown("### Aditivos e Supressões")
        st.dataframe(
            df_visual(df_ad, moeda_cols=["Valor original da alteração", "Valor atualizado da alteração"], fator_cols=["Fator aplicado"]),
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
