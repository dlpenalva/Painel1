from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape

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


def primeiro_valor(dicionario, chaves, padrao=0.0):
    for chave in chaves:
        valor = dicionario.get(chave)
        if valor not in (None, ""):
            return valor
    return padrao


def valor_global_contrato(res):
    return primeiro_valor(
        res,
        [
            "valor_global_contrato",
            "valor_global_estoque",
            "valor_global_financeiro",
        ],
        0.0,
    )


def valor_aditivos(res):
    return primeiro_valor(
        res,
        [
            "aditivos_quantitativos_atualizados",
            "valor_aditivos_atualizados",
            "aditivos_atualizados",
        ],
        0.0,
    )


def valor_global_com_aditivos(res):
    valor = primeiro_valor(
        res,
        [
            "valor_global_contrato_com_aditivos",
            "valor_global_com_aditivos",
        ],
        None,
    )
    if valor is not None:
        return valor
    return float(valor_global_contrato(res) or 0.0) + float(valor_aditivos(res) or 0.0)


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


def gerar_informacoes_processuais(adm, res):
    origem = (adm or {}).get("origem") or (adm or {}).get("tipo") or res.get("origem_ciclos", "Não informado")
    indice = res.get("indice", (adm or {}).get("indice", "Não informado"))
    fator = res.get("fator_acumulado", (adm or {}).get("fator_acumulado", (adm or {}).get("fator", 1.0)))

    texto = f"""
RELATÓRIO EXECUTIVO — VALOR GLOBAL DO CONTRATO

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
Retroativo/Delta acumulado: {moeda(res.get("delta_acumulado", 0))}
Remanescente original: {moeda(res.get("remanescente_original", 0))}
Remanescente atualizado: {moeda(res.get("remanescente_reajustado", 0))}

4. Valor Global Contrato

Valor Global Contrato: {moeda(valor_global_contrato(res))}
Aditivos quantitativos atualizados: {moeda(valor_aditivos(res))}
Valor Global Contrato com aditivos: {moeda(valor_global_com_aditivos(res))}

5. Observação executiva

As competências anteriores ao início dos efeitos financeiros do respectivo ciclo não devem ser consideradas para fins de retroativo. Ciclos classificados como preclusos, já concedidos ou sem efeito financeiro devem ser preservados para fins de histórico contratual, mas não devem compor o retroativo a pagar, salvo orientação específica em contrário.
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
            visual[col] = visual[col].apply(lambda x: percentual(x, 2))
    return visual


# ============================================================
# PDF com ReportLab
# ============================================================

def _escape_pdf(valor):
    return escape("" if valor is None else str(valor))


def _formatar_valor_pdf(valor, nome_coluna=""):
    nome = str(nome_coluna).lower()
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    if isinstance(valor, (int, float)):
        if any(ch in nome for ch in ["valor", "total", "pago", "devido", "delta", "remanescente", "aditivo", "contrato"]):
            return moeda(valor)
        if "fator" in nome:
            return fator_fmt(valor)
        if "variação" in nome or "percentual" in nome or "%" in nome:
            return percentual(valor, 2)
        return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return str(valor)


def _colunas_existentes(df, especificacao):
    existentes = []
    for item in especificacao:
        origem = item[0]
        if origem in df.columns:
            existentes.append(item)
    return existentes


def _dados_df_para_pdf(df, especificacao=None, max_linhas=20):
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return [["Informação", "Sem dados disponíveis"]], False

    base = df.copy()
    if especificacao:
        specs = _colunas_existentes(base, especificacao)
        if specs:
            dados = [[rotulo for _, rotulo in specs]]
            for _, row in base.head(max_linhas).iterrows():
                dados.append([_formatar_valor_pdf(row.get(origem, ""), origem) for origem, _ in specs])
            return dados, True

    # Fallback: limita a quantidade de colunas para evitar estouro de página.
    colunas = list(base.columns[:7])
    dados = [colunas]
    for _, row in base[colunas].head(max_linhas).iterrows():
        dados.append([_formatar_valor_pdf(row.get(col, ""), col) for col in colunas])
    return dados, True


def tabela_pdf(dados, header=False, col_widths=None, available_width=None, font_size=6.4, h_align="CENTER"):
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import Paragraph, Table, TableStyle

    if not dados:
        dados = [["Informação", "Sem dados disponíveis"]]
        header = False

    ncols = max(len(linha) for linha in dados)
    dados_norm = []
    for linha in dados:
        dados_norm.append(list(linha) + [""] * (ncols - len(linha)))

    if available_width is None:
        available_width = 700

    if not col_widths:
        pesos = []
        for col in range(ncols):
            max_len = max(len(str(linha[col])) for linha in dados_norm[: min(len(dados_norm), 15)])
            pesos.append(min(max(max_len, 8), 32))
        total = sum(pesos) or ncols
        col_widths = [available_width * p / total for p in pesos]

    soma = sum(col_widths)
    if soma > available_width:
        escala = available_width / soma
        col_widths = [w * escala for w in col_widths]

    body_style = ParagraphStyle(
        name="TabelaBodyCompacta",
        fontName="Helvetica",
        fontSize=font_size,
        leading=font_size + 1.8,
        wordWrap="CJK",
        spaceAfter=0,
        spaceBefore=0,
    )
    header_style = ParagraphStyle(
        name="TabelaHeaderCompacta",
        parent=body_style,
        fontName="Helvetica-Bold",
        fontSize=font_size,
        leading=font_size + 1.8,
        wordWrap="CJK",
    )

    dados_pdf = []
    for r, linha in enumerate(dados_norm):
        estilo = header_style if header and r == 0 else body_style
        dados_pdf.append([Paragraph(_escape_pdf(valor), estilo) for valor in linha])

    table = Table(dados_pdf, colWidths=col_widths, repeatRows=1 if header else 0, splitByRow=1)
    table.hAlign = h_align
    style = [
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#9CA3AF")),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2.2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2.2),
        ("TOPPADDING", (0, 0), (-1, -1), 2.0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.0),
    ]
    if header:
        style += [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E9EEF6")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#111827")),
        ]
    table.setStyle(TableStyle(style))
    return table


def tabela_dataframe_pdf(df, especificacao=None, max_linhas=20, col_widths=None, available_width=None, font_size=6.4):
    dados, header = _dados_df_para_pdf(df, especificacao=especificacao, max_linhas=max_linhas)
    return tabela_pdf(
        dados,
        header=header,
        col_widths=col_widths,
        available_width=available_width,
        font_size=font_size,
        h_align="CENTER",
    )


def criar_pdf_relatorio(adm, res):
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer
    except Exception as exc:
        raise RuntimeError(
            "A biblioteca reportlab não está instalada. Inclua 'reportlab' no requirements.txt."
        ) from exc

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=1.0 * cm,
        leftMargin=1.0 * cm,
        topMargin=1.0 * cm,
        bottomMargin=1.0 * cm,
    )
    largura_util = float(doc.width)
    largura_tabela = largura_util * 0.90

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="TituloCentral",
        parent=styles["Title"],
        fontSize=14,
        leading=17,
        alignment=1,
        spaceAfter=4,
    ))
    styles.add(ParagraphStyle(
        name="SubtituloCentral",
        parent=styles["Heading2"],
        fontSize=10.5,
        leading=13,
        alignment=1,
        spaceBefore=8,
        spaceAfter=5,
    ))
    styles.add(ParagraphStyle(
        name="TextoJustificado",
        parent=styles["BodyText"],
        fontSize=8.2,
        leading=10.5,
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
                img = Image(str(caminho), width=3.7 * cm, height=1.1 * cm)
                img.hAlign = "CENTER"
                story.append(img)
                logo_adicionado = True
                break
            except Exception:
                pass

    story.append(Paragraph("TELEBRAS — Análise de Reajuste Contratual", styles["TituloCentral"]))
    story.append(Paragraph("Relatório Executivo - GCC", styles["TituloCentral"]))

    story.append(Paragraph("1. Identificação da Análise", styles["SubtituloCentral"]))
    origem = (adm or {}).get("origem") or (adm or {}).get("tipo") or res.get("origem_ciclos", "Não informado")
    indice = res.get("indice", (adm or {}).get("indice", "Não informado"))
    fator = res.get("fator_acumulado", (adm or {}).get("fator_acumulado", (adm or {}).get("fator", 1.0)))

    tabela_ident = [
        ["Origem da análise", origem],
        ["Índice aplicado", indice],
        ["Fator acumulado", fator_fmt(fator)],
        ["Data de processamento", res.get("data_processamento", "Não informado")],
    ]
    story.append(tabela_pdf(tabela_ident, col_widths=[0.32 * largura_util, 0.56 * largura_util], available_width=largura_tabela, font_size=7.0))

    story.append(Paragraph("2. Análise Contratual — Cláusula Oitava", styles["SubtituloCentral"]))
    story.append(Paragraph(texto_clausula_oito(adm), styles["TextoJustificado"]))

    story.append(Paragraph("3. Indicadores Executivos", styles["SubtituloCentral"]))
    indicadores = [
        ["Indicador", "Valor"],
        ["Valor original do contrato", moeda(res.get("valor_original_contrato", 0))],
        ["Total pago/faturado", moeda(res.get("total_pago_faturado", 0))],
        ["Total devido reajustado", moeda(res.get("total_devido_reajustado", 0))],
        ["Retroativo/Delta acumulado", moeda(res.get("delta_acumulado", 0))],
        ["Remanescente original", moeda(res.get("remanescente_original", 0))],
        ["Remanescente atualizado", moeda(res.get("remanescente_reajustado", 0))],
        ["Valor Global Contrato", moeda(valor_global_contrato(res))],
        ["Aditivos quantitativos atualizados", moeda(valor_aditivos(res))],
        ["Valor Global Contrato com aditivos", moeda(valor_global_com_aditivos(res))],
    ]
    story.append(tabela_pdf(indicadores, header=True, col_widths=[0.48 * largura_util, 0.40 * largura_util], available_width=largura_tabela, font_size=7.0))

    story.append(Paragraph("4. Ciclos, Percentuais e Efeitos Financeiros", styles["SubtituloCentral"]))
    ciclos_spec = [
        ("Ciclo", "Ciclo"),
        ("Data-base", "Início ciclo"),
        ("Data do pedido", "Pedido"),
        ("Início financeiro", "Efeito fin."),
        ("Situação", "Situação"),
        ("Tratamento financeiro do ciclo", "Tratamento"),
        ("Variação", "% ciclo"),
        ("Fator acumulado", "Fator acum."),
    ]
    story.append(tabela_dataframe_pdf(
        res.get("df_ciclos"),
        especificacao=ciclos_spec,
        max_linhas=25,
        col_widths=[1.1*cm, 2.1*cm, 2.1*cm, 2.1*cm, 2.4*cm, 3.1*cm, 1.7*cm, 2.0*cm],
        available_width=largura_tabela,
        font_size=5.9,
    ))

    story.append(Paragraph("5. Apuração de Retroativos", styles["SubtituloCentral"]))
    fin_spec = [
        ("Ciclo", "Ciclo"),
        ("Situação", "Situação"),
        ("Fator aplicado", "Fator"),
        ("Valor pago/faturado", "Pago/faturado"),
        ("Valor devido reajustado", "Devido reaj."),
        ("Delta do ciclo", "Retroativo"),
        ("Delta acumulado", "Acumulado"),
    ]
    story.append(tabela_dataframe_pdf(
        res.get("df_financeiro_por_ciclo"),
        especificacao=fin_spec,
        max_linhas=25,
        col_widths=[1.1*cm, 2.7*cm, 1.7*cm, 3.0*cm, 3.2*cm, 3.0*cm, 3.0*cm],
        available_width=largura_tabela,
        font_size=6.2,
    ))

    story.append(Paragraph("6. Valor Total Atualizado do Contrato por Ciclo", styles["SubtituloCentral"]))
    df_consumo = res.get("df_consumo_estoque")
    if isinstance(df_consumo, pd.DataFrame) and not df_consumo.empty:
        consumo_spec = [
            ("Ciclo", "Ciclo/Marco"),
            ("Critério", "Composição"),
            ("Valor consumido original", "Valor original"),
            ("Fator aplicado", "Fator"),
            ("Valor consumido reajustado", "Valor atualizado"),
        ]
        story.append(tabela_dataframe_pdf(
            df_consumo,
            especificacao=consumo_spec,
            max_linhas=20,
            col_widths=[1.6*cm, 7.5*cm, 3.3*cm, 2.2*cm, 3.5*cm],
            available_width=largura_tabela,
            font_size=6.1,
        ))
    else:
        dados_valor_total = [
            ["Indicador", "Valor"],
            ["Valor original inicial do contrato", moeda(res.get("valor_original_contrato", 0))],
            ["Executado atualizado", moeda(res.get("valor_consumido_estoque", 0))],
            ["Remanescente atualizado no último ciclo", moeda(res.get("remanescente_reajustado", 0))],
            ["Valor Global Contrato", moeda(valor_global_contrato(res))],
            ["Aditivos quantitativos atualizados", moeda(valor_aditivos(res))],
            ["Valor Global Contrato com aditivos", moeda(valor_global_com_aditivos(res))],
        ]
        story.append(tabela_pdf(
            dados_valor_total,
            header=True,
            col_widths=[0.48 * largura_tabela, 0.32 * largura_tabela],
            available_width=largura_tabela,
            font_size=6.8,
            h_align="CENTER",
        ))

    story.append(Paragraph("7. Informações para instrução processual", styles["SubtituloCentral"]))
    story.append(Paragraph(gerar_informacoes_processuais(adm, res).replace("\n", "<br/>").replace("&", "&amp;"), styles["TextoJustificado"]))

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
col2.metric("Fator acumulado", fator_fmt(res.get("fator_acumulado", 1.0)))
col3.metric("Valor Global Contrato", moeda(valor_global_contrato(res)))
col4.metric("Retroativo/Delta", moeda(res.get("delta_acumulado", 0)))

col5, col6, col7 = st.columns(3)
col5.metric("Aditivos atualizados", moeda(valor_aditivos(res)))
col6.metric("Valor com aditivos", moeda(valor_global_com_aditivos(res)))
col7.metric("Remanescente atualizado", moeda(res.get("remanescente_reajustado", 0)))

st.divider()

tab1, tab2, tab3 = st.tabs(["Relatório Executivo", "Tabelas", "PDF"])

with tab1:
    st.markdown("### Fundamentação Contratual")
    st.info(texto_clausula_oito(adm))

    st.markdown("### Informações para instrução processual")
    texto = gerar_informacoes_processuais(adm, res)
    st.text_area("Copie as informações:", texto, height=420)

with tab2:
    st.markdown("### Ciclos, percentuais e efeitos financeiros")
    df_ciclos = res.get("df_ciclos")
    if isinstance(df_ciclos, pd.DataFrame) and not df_ciclos.empty:
        st.dataframe(df_ciclos, use_container_width=True, hide_index=True)
    else:
        st.info("Sem dados de ciclos disponíveis.")

    st.markdown("### Apuração de retroativos")
    df_fin = res.get("df_financeiro_por_ciclo")
    if isinstance(df_fin, pd.DataFrame) and not df_fin.empty:
        st.dataframe(
            df_visual(
                df_fin,
                moeda_cols=["Valor pago/faturado", "Valor devido reajustado", "Delta do ciclo", "Delta acumulado"],
                fator_cols=["Fator aplicado"],
            ),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("Sem dados financeiros disponíveis.")

    st.markdown("### Remanescentes")
    df_rem = res.get("df_remanescentes")
    if isinstance(df_rem, pd.DataFrame) and not df_rem.empty:
        st.dataframe(
            df_visual(
                df_rem,
                moeda_cols=["Remanescente original", "Remanescente reajustado"],
                fator_cols=["Fator aplicado"],
            ),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("Sem dados de remanescentes disponíveis.")

with tab3:
    st.markdown("### Baixar Relatório Executivo em PDF")
    try:
        pdf_bytes = criar_pdf_relatorio(adm, res)
        st.download_button(
            label="Baixar Relatório Executivo em PDF",
            data=pdf_bytes,
            file_name="Relatorio_Executivo_Reajuste_GCC.pdf",
            mime="application/pdf",
            type="primary",
        )
    except Exception as exc:
        st.error(f"Não foi possível gerar o PDF: {exc}")
