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


def df_visual(df, moeda_cols=None, pct_cols=None):
    if df is None or not isinstance(df, pd.DataFrame):
        return pd.DataFrame()
    visual = df.copy()
    for col in moeda_cols or []:
        if col in visual.columns:
            visual[col] = visual[col].apply(moeda)
    for col in pct_cols or []:
        if col in visual.columns:
            visual[col] = visual[col].apply(percentual)
    return visual


# ============================================================
# Textos executivos
# ============================================================

def ciclos_do_resultado_ou_adm(adm, res):
    df_ciclos = res.get("df_ciclos") if isinstance(res, dict) else None
    if isinstance(df_ciclos, pd.DataFrame) and not df_ciclos.empty:
        return df_ciclos.to_dict(orient="records")
    if adm:
        return adm.get("ciclos") or adm.get("detalhamento_ciclos") or []
    return []


def texto_clausula_oito(adm, res=None):
    ciclos = ciclos_do_resultado_ou_adm(adm, res or {})
    status_gerais = []
    for c in ciclos:
        status_gerais.append(normalizar_status(c.get("Situação") or c.get("situacao") or c.get("status")))
    texto_status = " ".join(status_gerais)

    if "PRECLUSO" in texto_status:
        return (
            "A análise registra ciclo classificado como precluso, em razão da ausência de solicitação dentro "
            "do prazo de 90 dias previsto no Parágrafo Quinto da Cláusula Oitava, observando-se também a regra "
            "do Parágrafo Sétimo quanto à possibilidade de novo pleito após ultrapassados 12 meses da data em que "
            "poderia ter sido requerido."
        )
    if "ADMISSÍVEL COM RESSALVA" in texto_status:
        return (
            "A análise registra pleito admissível com ressalva, por ter sido apresentado no mesmo mês de implemento "
            "da anualidade, porém antes do dia exato de completude dos 12 meses. Os efeitos financeiros devem observar "
            "os Parágrafos Primeiro e Segundo da Cláusula Oitava."
        )
    if "ADIANTADO" in texto_status:
        return (
            "A análise registra pleito apresentado antes do implemento da anualidade contratual. Nos termos dos "
            "Parágrafos Primeiro e Segundo da Cláusula Oitava, eventual reconhecimento de efeitos financeiros deve "
            "observar a completude dos 12 meses e a data juridicamente apta para o pedido."
        )
    return (
        "O pleito foi classificado como tempestivo, considerando a anualidade prevista no Parágrafo Primeiro da "
        "Cláusula Oitava e a apresentação da solicitação dentro da janela contratual de 90 dias prevista no Parágrafo "
        "Quinto. Os efeitos financeiros devem observar o Parágrafo Segundo da mesma cláusula."
    )


def gerar_texto_instrucao(adm, res):
    origem = (adm or {}).get("origem") or (adm or {}).get("tipo") or res.get("origem_ciclos", "Não informado")
    indice = res.get("indice", (adm or {}).get("indice", "Não informado"))
    ciclos = res.get("df_ciclos")
    qtd_ciclos = len(ciclos) if isinstance(ciclos, pd.DataFrame) else 0

    return f"""
RELATÓRIO EXECUTIVO — VALOR GLOBAL DO CONTRATO

1. Contexto da análise

A presente análise consolida os resultados da etapa de admissibilidade do reajuste contratual e da etapa de quantificação financeira, com base nos dados constantes do Arquivo de Coleta preenchido.

Origem da análise: {origem}
Índice utilizado: {indice}
Quantidade de ciclos considerados: {qtd_ciclos}
Percentual acumulado aplicado: {percentual(res.get('percentual_acumulado', 0))}

2. Fundamentação contratual

{texto_clausula_oito(adm, res)}

3. Regra de efeito financeiro

A apuração financeira considerou que os efeitos financeiros do reajuste são reconhecidos a partir da data da solicitação da contratada, quando atendidos os requisitos contratuais. Assim, competências anteriores ao mês do pedido não foram consideradas para fins de retroativo, ainda que integrem o ciclo teórico de reajuste.

4. Resultado financeiro consolidado

Valor original inicial do contrato: {moeda(res.get('valor_original_contrato', 0))}
Total pago/faturado: {moeda(res.get('total_pago_faturado', 0))}
Executado atualizado: {moeda(res.get('total_devido_reajustado', 0))}
Total retroativo a pagar: {moeda(res.get('delta_acumulado', 0))}
Remanescente original no último ciclo: {moeda(res.get('remanescente_original', 0))}
Remanescente atualizado no último ciclo: {moeda(res.get('remanescente_reajustado', 0))}
Valor Global Contrato: {moeda(res.get('valor_global_contrato', 0))}

5. Encaminhamento

As informações acima consolidam a apuração financeira por ciclo, os meses com efeito financeiro, os valores pagos/faturados, os valores atualizados e o retroativo total estimado, servindo como subsídio para instrução processual, validação pela fiscalização/gestão contratual e eventual formalização do reajuste.
""".strip()


# ============================================================
# PDF com ReportLab
# ============================================================

def tabela_pdf(dados, header=False, col_widths=None):
    from reportlab.lib import colors
    from reportlab.platypus import Table, TableStyle

    table = Table(dados, colWidths=col_widths, repeatRows=1 if header else 0)
    table.hAlign = "LEFT"
    estilo = [
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
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


def tabela_dataframe_pdf(df, colunas=None, max_linhas=20):
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return tabela_pdf([["Informação", "Sem dados disponíveis"]], col_widths=[5 * 28.3, 11 * 28.3])
    base = df.copy()
    if colunas:
        colunas_existentes = [c for c in colunas if c in base.columns]
        base = base[colunas_existentes]
    dados = [list(base.columns)]
    for _, row in base.head(max_linhas).iterrows():
        linha = []
        for valor in row.tolist():
            if isinstance(valor, float):
                linha.append(f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            else:
                linha.append(str(valor))
        dados.append(linha)
    return tabela_pdf(dados, header=True)


def criar_pdf_relatorio(adm, res):
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer
    except Exception as exc:
        raise RuntimeError("A biblioteca reportlab não está instalada. Inclua 'reportlab' no requirements.txt.") from exc

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.4 * cm,
        leftMargin=1.4 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TituloTelebras", parent=styles["Title"], fontSize=14, leading=17, spaceAfter=8, alignment=1))
    styles.add(ParagraphStyle(name="Subtitulo", parent=styles["Heading2"], fontSize=10.5, leading=13, spaceBefore=8, spaceAfter=5, alignment=1))
    styles.add(ParagraphStyle(name="Texto", parent=styles["BodyText"], fontSize=8.7, leading=11, alignment=4))

    story = []
    logo_paths = [Path("assets/telebras_logo.png"), Path("assets/logo.png"), Path("telebras_logo.png"), Path("logo.png")]
    logo_adicionado = False
    for caminho in logo_paths:
        if caminho.exists():
            try:
                logo = Image(str(caminho), width=4.0 * cm, height=1.2 * cm)
                logo.hAlign = "CENTER"
                story.append(logo)
                logo_adicionado = True
                break
            except Exception:
                pass
    if not logo_adicionado:
        story.append(Paragraph("TELEBRAS — Análise de Reajuste Contratual", styles["TituloTelebras"]))
    else:
        story.append(Paragraph("TELEBRAS — Análise de Reajuste Contratual", styles["TituloTelebras"]))
    story.append(Paragraph("Relatório Executivo - GCC", styles["TituloTelebras"]))

    story.append(Paragraph("1. Identificação da Análise", styles["Subtitulo"]))
    tabela_ident = [
        ["Índice aplicado", res.get("indice", "Não informado")],
        ["Percentual acumulado", percentual(res.get("percentual_acumulado", 0))],
        ["Data de processamento", res.get("data_processamento", "Não informado")],
    ]
    story.append(tabela_pdf(tabela_ident, col_widths=[6 * cm, 10 * cm]))
    story.append(Spacer(1, 6))

    story.append(Paragraph("2. Análise Contratual — Cláusula Oitava", styles["Subtitulo"]))
    story.append(Paragraph(texto_clausula_oito(adm, res), styles["Texto"]))

    story.append(Paragraph("3. Indicadores Executivos", styles["Subtitulo"]))
    indicadores = [
        ["Indicador", "Valor"],
        ["Valor original inicial do contrato", moeda(res.get("valor_original_contrato", 0))],
        ["Total pago/faturado", moeda(res.get("total_pago_faturado", 0))],
        ["Executado atualizado", moeda(res.get("total_devido_reajustado", 0))],
        ["Total retroativo a pagar", moeda(res.get("delta_acumulado", 0))],
        ["Remanescente original", moeda(res.get("remanescente_original", 0))],
        ["Remanescente atualizado", moeda(res.get("remanescente_reajustado", 0))],
        ["Valor Global Contrato", moeda(res.get("valor_global_contrato", 0))],
        ["Aditivos quantitativos atualizados", moeda(res.get("valor_aditivos_atualizados", 0))],
        ["Valor Global Contrato com aditivos", moeda(res.get("valor_global_contrato_com_aditivos", 0))],
    ]
    story.append(tabela_pdf(indicadores, header=True, col_widths=[8 * cm, 8 * cm]))

    story.append(Paragraph("4. Ciclos, Percentuais e Efeitos Financeiros", styles["Subtitulo"]))
    story.append(tabela_dataframe_pdf(
        res.get("df_ciclos"),
        colunas=["Ciclo", "Mês início do ciclo", "Data do pedido", "Mês início efeito financeiro", "Situação", "Percentual do ciclo", "Percentual acumulado"],
        max_linhas=12,
    ))

    story.append(Paragraph("5. Apuração de Retroativos", styles["Subtitulo"]))
    story.append(tabela_dataframe_pdf(
        res.get("df_financeiro_por_ciclo"),
        colunas=["Ciclo", "Meses com efeito financeiro", "Meses sem efeito financeiro", "Valor pago/faturado", "Valor executado atualizado", "Retroativo do ciclo"],
        max_linhas=15,
    ))

    story.append(Paragraph("6. Valor Total Atualizado do Contrato por Ciclo", styles["Subtitulo"]))
    story.append(tabela_dataframe_pdf(
        res.get("df_valor_total_por_ciclo"),
        colunas=["Ciclo/Marco", "Mês de referência", "Composição", "Valor Global Contrato"],
        max_linhas=15,
    ))

    df_aditivos = res.get("df_aditivos")
    if isinstance(df_aditivos, pd.DataFrame) and not df_aditivos.empty:
        story.append(Paragraph("7. Aditivos quantitativos", styles["Subtitulo"]))
        story.append(tabela_dataframe_pdf(
            df_aditivos,
            colunas=["Descrição do aditivo", "Data do aditivo", "Ciclo/Marco", "Valor original do acréscimo", "Valor atualizado do acréscimo"],
            max_linhas=10,
        ))
        secao_info = "8. Informações para instrução processual"
    else:
        secao_info = "7. Informações para instrução processual"

    story.append(Paragraph(secao_info, styles["Subtitulo"]))
    story.append(Paragraph(gerar_texto_instrucao(adm, res).replace("\n", "<br/>").replace("&", "&amp;"), styles["Texto"]))

    doc.build(story)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf


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
col2.metric("Percentual acumulado", percentual(res.get("percentual_acumulado", 0)))
col3.metric("Valor Global Contrato", moeda(res.get("valor_global_contrato", 0)))
col4.metric("Total retroativo a pagar", moeda(res.get("delta_acumulado", 0)))

col5, col6, col7 = st.columns(3)
col5.metric("Valor original", moeda(res.get("valor_original_contrato", 0)))
col6.metric("Total pago/faturado", moeda(res.get("total_pago_faturado", 0)))
col7.metric("Executado atualizado", moeda(res.get("total_devido_reajustado", 0)))

col8, col9 = st.columns(2)
col8.metric("Aditivos atualizados", moeda(res.get("valor_aditivos_atualizados", 0)))
col9.metric("Valor Global Contrato com aditivos", moeda(res.get("valor_global_contrato_com_aditivos", res.get("valor_global_contrato", 0))))

st.divider()

tab1, tab2, tab3 = st.tabs(["Relatório Executivo", "Tabelas", "PDF"])

with tab1:
    st.markdown("### Fundamentação contratual")
    st.info(texto_clausula_oito(adm, res))

    st.markdown("### Informações para instrução processual")
    texto_instrucao = gerar_texto_instrucao(adm, res)
    st.text_area("Texto para instrução processual:", texto_instrucao, height=420)

with tab2:
    st.markdown("### Ciclos, percentuais e efeitos financeiros")
    st.dataframe(
        df_visual(res.get("df_ciclos"), pct_cols=["Percentual do ciclo", "Percentual acumulado", "Variação"]),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("### Apuração de Retroativos")
    st.dataframe(
        df_visual(
            res.get("df_financeiro_por_ciclo"),
            moeda_cols=["Valor pago/faturado", "Valor executado atualizado", "Retroativo do ciclo", "Retroativo acumulado"],
            pct_cols=["Percentual do ciclo", "Percentual acumulado aplicado"],
        ),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("### Remanescentes")
    st.dataframe(
        df_visual(
            res.get("df_remanescentes"),
            moeda_cols=["Remanescente original", "Remanescente atualizado"],
            pct_cols=["Percentual acumulado aplicado"],
        ),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("### Aditivos quantitativos")
    df_aditivos = res.get("df_aditivos")
    if isinstance(df_aditivos, pd.DataFrame) and not df_aditivos.empty:
        st.dataframe(
            df_visual(df_aditivos, moeda_cols=["Valor original do acréscimo", "Valor atualizado do acréscimo"], pct_cols=["Percentual aplicado"]),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("Nenhum aditivo quantitativo foi informado.")

    st.markdown("### Valor Global Contrato por ciclo")
    st.dataframe(
        df_visual(
            res.get("df_valor_total_por_ciclo"),
            moeda_cols=["Executado atualizado até ciclo anterior", "Remanescente atualizado", "Valor Global Contrato"],
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
            file_name="relatorio_executivo_gcc_valor_global.pdf",
            mime="application/pdf",
            type="primary",
        )
    except Exception as exc:
        st.error(f"Não foi possível gerar o PDF: {exc}")
        st.caption("Verifique se a biblioteca reportlab foi incluída no requirements.txt.")
