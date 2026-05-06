from io import BytesIO
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Análises de Reajustes - Relatório Global", layout="wide")


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


def agora_brasilia():
    return datetime.now(ZoneInfo("America/Sao_Paulo"))


def normalizar_status(status):
    texto = str(status or "").upper()
    if "PRECLUS" in texto:
        return "[!] PRECLUSO"
    if "RESSALVA" in texto:
        return "[R] ADMISSÍVEL COM RESSALVA"
    if "TEMPEST" in texto:
        return "[OK] TEMPESTIVO"
    if "ADIANT" in texto:
        return "[A] ADIANTADO"
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
            "A análise registra ciclo classificado como precluso, razão pela qual a variação do índice pode "
            "ser exibida para memória, mas não compõe o efeito financeiro do acumulado nem o retroativo a pagar, "
            "observados os Parágrafos Quinto e Sétimo da Cláusula Oitava."
        )
    if "RESSALVA" in texto_status:
        return (
            "A análise registra pleito admissível com ressalva, por ter sido apresentado no mesmo mês de "
            "implemento da anualidade, porém antes do dia exato de completude dos 12 meses. Os efeitos financeiros "
            "devem observar os Parágrafos Primeiro e Segundo da Cláusula Oitava."
        )
    if "ADIANTADO" in texto_status:
        return (
            "A análise registra pleito apresentado antes do implemento da anualidade contratual. Eventual "
            "reconhecimento de efeitos financeiros deve observar a completude dos 12 meses e a data juridicamente "
            "apta para o pedido, nos termos da Cláusula Oitava."
        )
    return (
        "O pleito foi classificado como tempestivo, considerando a anualidade prevista no Parágrafo Primeiro "
        "da Cláusula Oitava e a apresentação da solicitação dentro da janela contratual de 90 dias prevista "
        "no Parágrafo Quinto. Os efeitos financeiros observam o Parágrafo Segundo da mesma cláusula."
    )


def gerar_informacoes_processuais(adm, res):
    origem = (adm or {}).get("origem") or (adm or {}).get("tipo") or "Não informado"
    indice = res.get("indice", (adm or {}).get("indice", "Não informado"))
    fator = res.get("fator_acumulado", (adm or {}).get("fator_acumulado", (adm or {}).get("fator", 1.0)))
    texto = f"""
RELATÓRIO EXECUTIVO DE ANÁLISE DE REAJUSTE CONTRATUAL

1. Contexto da análise

A presente análise consolida os resultados da etapa de admissibilidade do reajuste contratual e da etapa de quantificação financeira, com base nos dados constantes do Arquivo de Coleta preenchido.

Origem da análise: {origem}
Índice utilizado: {indice}
Fator acumulado considerado: {fator_fmt(fator)}
Quantidade de ciclos identificados: {res.get('quantidade_ciclos', 0)}

2. Fundamentação contratual

{texto_clausula_oito(adm)}

3. Resultado financeiro consolidado

Valor original do contrato: {moeda(res.get('valor_original_contrato', 0))}
Valor pago efetivo: {moeda(res.get('valor_pago_efetivo', res.get('total_pago_faturado', 0)))}
Valor teórico calculado: {moeda(res.get('valor_teorico_calculado', res.get('total_devido_reajustado', 0)))}
Delta total apurado: {moeda(res.get('delta_total', res.get('delta_acumulado', 0)))}
Saldo remanescente atualizado: {moeda(res.get('remanescente_reajustado', 0))}
Valor atualizado do contrato: {moeda(res.get('valor_atualizado_contrato', res.get('valor_global_financeiro', 0)))}

4. Observação executiva

O painel considera o que foi efetivamente pago, o que deveria ter sido pago conforme os ciclos financeiros válidos e o saldo remanescente atualizado. Ciclos preclusos ou sem efeito financeiro não geram retroativo, sem prejuízo de constarem na memória histórica da análise.
"""
    return texto.strip()


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
# PDF com ReportLab
# ============================================================

def preparar_df_pdf(df, colunas=None, moeda_cols=None, fator_cols=None, pct_cols=None, max_linhas=18):
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()
    base = df.copy()
    if colunas:
        existentes = [c for c in colunas if c in base.columns]
        base = base[existentes]
    base = df_visual(base, moeda_cols=moeda_cols, fator_cols=fator_cols, pct_cols=pct_cols)
    return base.head(max_linhas)


def criar_pdf_relatorio(adm, res):
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    except Exception as exc:
        raise RuntimeError("A biblioteca reportlab não está instalada. Inclua 'reportlab' no requirements.txt.") from exc

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=1.0 * cm,
        leftMargin=1.0 * cm,
        topMargin=1.0 * cm,
        bottomMargin=1.0 * cm,
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Titulo", parent=styles["Title"], fontSize=14, leading=16, alignment=1, spaceAfter=6))
    styles.add(ParagraphStyle(name="Subtitulo", parent=styles["Heading2"], fontSize=10, leading=12, spaceBefore=6, spaceAfter=4))
    styles.add(ParagraphStyle(name="Texto", parent=styles["BodyText"], fontSize=8, leading=10, alignment=4))

    story = []
    logo_paths = [Path("assets/telebras_logo.png"), Path("assets/logo.png"), Path("telebras_logo.png"), Path("logo.png")]
    for caminho in logo_paths:
        if caminho.exists():
            try:
                story.append(Image(str(caminho), width=3.8 * cm, height=1.1 * cm))
                break
            except Exception:
                pass

    story.append(Paragraph("TELEBRAS — Análise de Reajuste Contratual", styles["Titulo"]))
    story.append(Paragraph("Relatório Executivo - GCC", styles["Titulo"]))
    story.append(Spacer(1, 6))

    origem = (adm or {}).get("origem") or (adm or {}).get("tipo") or "Não informado"
    indice = res.get("indice", (adm or {}).get("indice", "Não informado"))
    fator = res.get("fator_acumulado", (adm or {}).get("fator_acumulado", (adm or {}).get("fator", 1.0)))
    data_proc = res.get("data_processamento") or agora_brasilia().strftime("%d/%m/%Y %H:%M")

    story.append(Paragraph("1. Identificação da Análise", styles["Subtitulo"]))
    story.append(tabela_pdf([
        ["Origem", origem, "Índice", indice],
        ["Fator acumulado", fator_fmt(fator), "Data de processamento", data_proc],
        ["Ciclos identificados", str(res.get("quantidade_ciclos", 0)), "Valor atualizado do contrato", moeda(res.get("valor_atualizado_contrato", res.get("valor_global_financeiro", 0)))],
    ], col_widths=[4.0*cm, 6.0*cm, 5.0*cm, 7.0*cm]))

    story.append(Paragraph("2. Análise Contratual — Cláusula Oitava", styles["Subtitulo"]))
    story.append(Paragraph(texto_clausula_oito(adm), styles["Texto"]))

    story.append(Paragraph("3. Painel Executivo", styles["Subtitulo"]))
    indicadores = [
        ["Indicador", "Valor"],
        ["Valor original", moeda(res.get("valor_original_contrato", 0))],
        ["Valor pago efetivo", moeda(res.get("valor_pago_efetivo", res.get("total_pago_faturado", 0)))],
        ["Valor teórico calculado", moeda(res.get("valor_teorico_calculado", res.get("total_devido_reajustado", 0)))],
        ["Delta total", moeda(res.get("delta_total", res.get("delta_acumulado", 0)))],
        ["Saldo remanescente atualizado", moeda(res.get("remanescente_reajustado", 0))],
        ["Aditivos/Supressões atualizados", moeda(res.get("total_aditivos_atualizados", 0))],
        ["Valor atualizado do contrato", moeda(res.get("valor_atualizado_contrato", res.get("valor_global_financeiro", 0)))],
    ]
    story.append(tabela_pdf(indicadores, header=True, col_widths=[8.0*cm, 7.0*cm]))

    story.append(Paragraph("4. Ciclos, Percentuais e Efeitos Financeiros", styles["Subtitulo"]))
    df_ciclos = preparar_df_pdf(
        res.get("df_ciclos"),
        colunas=["Ciclo", "Data-base", "Data do pedido", "Situação", "Tratamento financeiro do ciclo", "Variação", "Fator acumulado efetivo"],
        pct_cols=["Variação"], fator_cols=["Fator acumulado efetivo"], max_linhas=12,
    )
    if not df_ciclos.empty and "Situação" in df_ciclos.columns:
        df_ciclos["Situação"] = df_ciclos["Situação"].apply(normalizar_status)
    story.append(tabela_dataframe_pdf(df_ciclos))

    story.append(Paragraph("5. Financeiro por Ciclo", styles["Subtitulo"]))
    df_fin = preparar_df_pdf(
        res.get("df_financeiro_por_ciclo"),
        colunas=["Ciclo", "Situação", "Tratamento financeiro", "Valor pago efetivo", "Valor teórico calculado", "Delta do ciclo"],
        moeda_cols=["Valor pago efetivo", "Valor teórico calculado", "Delta do ciclo"], max_linhas=14,
    )
    if not df_fin.empty and "Situação" in df_fin.columns:
        df_fin["Situação"] = df_fin["Situação"].apply(normalizar_status)
    story.append(tabela_dataframe_pdf(df_fin))

    story.append(Paragraph("6. Valor Atualizado do Contrato", styles["Subtitulo"]))
    df_exec = preparar_df_pdf(
        res.get("df_execucao_atualizada"),
        colunas=["Ciclo", "Status financeiro", "Valor executado original", "Percentual acumulado aplicado", "Valor executado atualizado"],
        moeda_cols=["Valor executado original", "Valor executado atualizado"], pct_cols=["Percentual acumulado aplicado"], max_linhas=12,
    )
    story.append(tabela_dataframe_pdf(df_exec))

    story.append(Paragraph("7. Informações para instrução processual", styles["Subtitulo"]))
    story.append(Paragraph(gerar_informacoes_processuais(adm, res).replace("\n", "<br/>"), styles["Texto"]))

    doc.build(story)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf


def tabela_pdf(dados, header=False, col_widths=None):
    from reportlab.lib import colors
    from reportlab.platypus import Table, TableStyle
    table = Table(dados, colWidths=col_widths, hAlign="CENTER")
    estilo = [
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#B7B7B7")),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
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


def tabela_dataframe_pdf(df):
    from reportlab.lib.units import cm
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return tabela_pdf([["Informação", "Sem dados disponíveis"]], col_widths=[6*cm, 10*cm])
    dados = [list(df.columns)]
    for _, row in df.iterrows():
        dados.append([str(v) for v in row.tolist()])
    largura_total = 25.0 * cm
    ncols = max(len(dados[0]), 1)
    col_width = largura_total / ncols
    return tabela_pdf(dados, header=True, col_widths=[col_width] * ncols)


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
col2.metric("Ciclos", res.get("quantidade_ciclos", 0))
col3.metric("Valor Atualizado do Contrato", moeda(res.get("valor_atualizado_contrato", res.get("valor_global_financeiro", 0))))
col4.metric("Delta Total", moeda(res.get("delta_total", res.get("delta_acumulado", 0))))

st.divider()

tab1, tab2, tab3 = st.tabs(["Relatório Executivo", "Tabelas Financeiras", "PDF"])

with tab1:
    st.markdown("### Fundamentação Contratual")
    st.info(texto_clausula_oito(adm))
    st.markdown("### Informações para instrução processual")
    texto_sei = gerar_informacoes_processuais(adm, res)
    st.text_area("Copie as informações para a instrução processual:", texto_sei, height=420)

with tab2:
    st.markdown("### Painel Executivo")
    st.dataframe(
        df_visual(res.get("df_comparativo"), moeda_cols=["Valor"]),
        use_container_width=True,
        hide_index=True,
    )
    st.markdown("### Financeiro por Ciclo")
    st.dataframe(
        df_visual(
            res.get("df_financeiro_por_ciclo"),
            moeda_cols=["Valor pago efetivo", "Valor teórico calculado", "Delta do ciclo"],
            fator_cols=["Fator aplicado ao retroativo"],
        ),
        use_container_width=True,
        hide_index=True,
    )
    st.markdown("### Valor Atualizado do Contrato")
    st.dataframe(
        df_visual(
            res.get("df_execucao_atualizada"),
            moeda_cols=["Valor executado original", "Valor executado atualizado"],
            pct_cols=["Percentual acumulado aplicado"],
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
            file_name="relatorio_executivo_reajuste.pdf",
            mime="application/pdf",
            type="primary",
        )
    except Exception as exc:
        st.error(f"Não foi possível gerar o PDF: {exc}")
        st.caption("Verifique se a biblioteca reportlab foi incluída no requirements.txt.")
