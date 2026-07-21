from datetime import datetime, date
from io import BytesIO
from zoneinfo import ZoneInfo
from html import escape

import pandas as pd
import streamlit as st
from dateutil.relativedelta import relativedelta

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    REPORTLAB_OK = True
except Exception:
    REPORTLAB_OK = False

from _ui_utils import render_marca_topo, render_aviso_privacidade

st.set_page_config(page_title="Análises de Reajustes - Garantia", layout="wide")




# >>> UX_ADITIVOS_25_COMPACTO
def aplicar_css_aditivos25_compacto():
    st.markdown(
        """
        <style>
        div[data-testid="stMetric"] {
            min-height: 72px;
            padding: 8px 10px;
        }
        div[data-testid="stMetricValue"] {
            font-size: clamp(0.95rem, 1.55vw, 1.28rem) !important;
            line-height: 1.12 !important;
            white-space: normal !important;
            overflow-wrap: anywhere !important;
            word-break: normal !important;
        }
        div[data-testid="stMetricLabel"] p {
            font-size: clamp(0.70rem, 1.00vw, 0.86rem) !important;
            line-height: 1.15 !important;
            white-space: normal !important;
        }
        .aditivos25-ux-note {
            font-size: 0.86rem;
            color: #475569;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
# <<< UX_ADITIVOS_25_COMPACTO
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
    """Aceita 85771019,12, 85.771.019,12, 85771019.12 ou R$ 85.771.019,12."""
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        try:
            if pd.isna(valor):
                return 0.0
        except Exception:
            pass
        return float(valor)
    texto = str(valor).strip()
    if not texto:
        return 0.0
    texto = texto.replace("R$", "").replace(" ", "")
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except Exception:
        return 0.0


def numero_para_input(valor):
    return parse_moeda_br(valor)


def limpar_texto(valor):
    if valor is None:
        return ""
    try:
        if isinstance(valor, float) and pd.isna(valor):
            return ""
    except Exception:
        pass
    return str(valor).strip()


def data_hora_brasilia():
    return datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M")


def data_para_texto(valor):
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass
    if isinstance(valor, (datetime, date)):
        return valor.strftime("%d/%m/%Y")
    try:
        dt = pd.to_datetime(valor, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return limpar_texto(valor)


def primeira_coluna_existente(df, candidatos):
    if not isinstance(df, pd.DataFrame) or df.empty:
        return None
    mapa = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidatos:
        chave = str(cand).strip().lower()
        if chave in mapa:
            return mapa[chave]
    return None


def obter_contexto_valor_global(resultado_valor_global):
    """Extrai dados importados do módulo Valores, sem alterar a regra pétrea do Valor Total Atualizado."""
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
        resultado_valor_global.get("valor_formalizado_anterior", contexto.get("valor_formalizado_anterior", 0.0))
    )
    valor_executado_atualizado = numero_para_input(resultado_valor_global.get("valor_executado_atualizado", 0.0))
    remanescente_atualizado = numero_para_input(resultado_valor_global.get("remanescente_reajustado", 0.0))
    variacao_acumulada = numero_para_input(resultado_valor_global.get("variacao_acumulada", 0.0))
    quantidade_aditivos = int(numero_para_input(resultado_valor_global.get("quantidade_aditivos_total", 0)))
    modo_apuracao = limpar_texto(resultado_valor_global.get("modo_apuracao", ""))
    retroativo_estimado_itens = numero_para_input(resultado_valor_global.get("valor_retroativo_estimado_itens_estoque", 0.0))

    df_aditivos = resultado_valor_global.get("df_aditivos_executivo", pd.DataFrame())
    if not isinstance(df_aditivos, pd.DataFrame):
        df_aditivos = pd.DataFrame()

    df_ciclos = resultado_valor_global.get("df_ciclos", pd.DataFrame())
    if not isinstance(df_ciclos, pd.DataFrame):
        df_ciclos = pd.DataFrame()

    return {
        "valor_original": valor_original,
        "valor_total_atualizado": valor_total_atualizado,
        "valor_formalizado_anterior": valor_formalizado_anterior,
        "valor_executado_atualizado": valor_executado_atualizado,
        "remanescente_atualizado": remanescente_atualizado,
        "variacao_acumulada": variacao_acumulada,
        "quantidade_aditivos": quantidade_aditivos,
        "modo_apuracao": modo_apuracao,
        "retroativo_estimado_itens": retroativo_estimado_itens,
        "df_aditivos": df_aditivos,
        "df_ciclos": df_ciclos,
    }


def extrair_aditivos_para_garantia(df_aditivos):
    """Consolida aditivos para leitura orgânica da garantia.

    Mantém a lógica de instrumento: se houver várias linhas técnicas, agrupa por data + ciclo.
    """
    colunas_saida = ["Evento", "Data", "Ciclo", "Valor original", "Valor atualizado", "Endosso esperado"]
    if not isinstance(df_aditivos, pd.DataFrame) or df_aditivos.empty:
        return pd.DataFrame(columns=colunas_saida)

    df = df_aditivos.copy()
    col_data = primeira_coluna_existente(df, ["Data do aditivo", "Data", "Data assinatura", "Data de assinatura"])
    col_ciclo = primeira_coluna_existente(df, ["Ciclo/Marco", "Ciclo", "Ciclo de referência", "ciclos"])
    col_valor_original = primeira_coluna_existente(
        df,
        [
            "Valor original consolidado",
            "Valor do aditivo na assinatura",
            "Valor da alteração na assinatura",
            "Valor assinado",
            "Valor original",
            "Valor",
        ],
    )
    col_valor_atualizado = primeira_coluna_existente(
        df,
        [
            "Valor atualizado consolidado",
            "Valor do aditivo reajustado",
            "Valor atualizado do aditivo",
            "Valor da alteração atualizado",
            "Valor atualizado",
        ],
    )
    col_aditivo = primeira_coluna_existente(df, ["Aditivo", "Instrumento", "Identificação", "Identificacao"])

    if col_data:
        df["_data_ref"] = df[col_data].apply(data_para_texto)
    else:
        df["_data_ref"] = ""

    if col_ciclo:
        df["_ciclo_ref"] = df[col_ciclo].apply(limpar_texto)
    else:
        df["_ciclo_ref"] = ""

    if col_aditivo:
        df["_aditivo_ref"] = df[col_aditivo].apply(limpar_texto)
    else:
        df["_aditivo_ref"] = ""

    if col_valor_original:
        df["_valor_original"] = df[col_valor_original].apply(parse_moeda_br)
    else:
        df["_valor_original"] = 0.0

    if col_valor_atualizado:
        df["_valor_atualizado"] = df[col_valor_atualizado].apply(parse_moeda_br)
    else:
        df["_valor_atualizado"] = df["_valor_original"]

    # Chave defensiva: data + ciclo; quando a data não existir, usa aditivo + ciclo.
    df["_chave"] = df.apply(
        lambda r: "|".join([
            r.get("_data_ref") or r.get("_aditivo_ref") or "Aditivo sem data",
            r.get("_ciclo_ref") or "",
        ]),
        axis=1,
    )

    linhas = []
    for idx, (_, grupo) in enumerate(df.groupby("_chave", dropna=False), start=1):
        data_ref = limpar_texto(grupo["_data_ref"].iloc[0]) if "_data_ref" in grupo else ""
        ciclo_ref = limpar_texto(grupo["_ciclo_ref"].iloc[0]) if "_ciclo_ref" in grupo else ""
        aditivo_ref = limpar_texto(grupo["_aditivo_ref"].iloc[0]) if "_aditivo_ref" in grupo else ""
        evento = aditivo_ref if aditivo_ref and not aditivo_ref.isdigit() else f"Aditivo {idx}"
        valor_original = round(float(grupo["_valor_original"].sum()), 2)
        valor_atualizado = round(float(grupo["_valor_atualizado"].sum()), 2)
        linhas.append(
            {
                "Evento": evento,
                "Data": data_ref or "Não informada",
                "Ciclo": ciclo_ref or "Não informado",
                "Valor original": valor_original,
                "Valor atualizado": valor_atualizado,
                "Endosso esperado": 0.0,
            }
        )

    return pd.DataFrame(linhas, columns=colunas_saida)


def resumo_ciclos_texto(df_ciclos, variacao_acumulada):
    if isinstance(df_ciclos, pd.DataFrame) and not df_ciclos.empty and "Ciclo" in df_ciclos.columns:
        ciclos = []
        col_pct = primeira_coluna_existente(df_ciclos, ["Percentual aplicado", "Variação", "Variacao", "%", "Percentual"])
        for _, row in df_ciclos.iterrows():
            ciclo = limpar_texto(row.get("Ciclo", ""))
            if not ciclo or ciclo.upper() == "C0":
                continue
            if col_pct:
                pct = parse_moeda_br(row.get(col_pct, 0.0))
                if abs(pct) < 1:
                    pct *= 100
                ciclos.append(f"{ciclo} ({pct:.2f}%)".replace(".", ","))
            else:
                ciclos.append(ciclo)
        if ciclos:
            return ", ".join(ciclos)
    if variacao_acumulada:
        pct = variacao_acumulada * 100 if abs(variacao_acumulada) < 1 else variacao_acumulada
        return f"reajuste acumulado informado ({pct:.2f}%)".replace(".", ",")
    return "reajustes informados na análise"




# >>> METODOLOGIA_VALOR_IMPORTADO_CORTE_OPERACIONAL
def _metodologia_valor_importado_html(resultado_valor_global):
    """Box informativo sobre a metodologia do Valor Total Atualizado importado do módulo Valores.
    Não altera cálculo. Apenas documenta a origem/metodologia usada.
    """
    if not isinstance(resultado_valor_global, dict) or not resultado_valor_global:
        return """
        <div style="background:#F8FAFC; border:1px solid #CBD5E1; border-left:6px solid #64748B; border-radius:12px; padding:13px 15px; margin:10px 0 16px 0; color:#334155;">
            <div style="font-weight:800; margin-bottom:5px;">Metodologia do Valor Total Atualizado importado</div>
            <div style="font-size:0.92rem; line-height:1.45;">Não há resultado do módulo Valores carregado nesta sessão. Os campos desta página devem ser conferidos/preenchidos manualmente.</div>
        </div>
        """

    cfg = resultado_valor_global.get("config_ciclo_em_execucao", {}) or {}
    corte = bool(resultado_valor_global.get("corte_operacional_aplicado") or resultado_valor_global.get("corte_operacional_solicitado") or cfg.get("aplicar"))

    try:
        valor_total = moeda(resultado_valor_global.get("valor_atualizado_contrato", resultado_valor_global.get("valor_global_estoque", 0)))
    except Exception:
        valor_total = str(resultado_valor_global.get("valor_atualizado_contrato", "Não informado"))
    try:
        execucao = moeda(resultado_valor_global.get("valor_executado_atualizado", 0))
    except Exception:
        execucao = str(resultado_valor_global.get("valor_executado_atualizado", "Não informado"))
    try:
        remanescente = moeda(resultado_valor_global.get("remanescente_reajustado", 0))
    except Exception:
        remanescente = str(resultado_valor_global.get("remanescente_reajustado", "Não informado"))
    try:
        c0_manual = moeda(cfg.get("valor_c0_manual", 0)) if parse_moeda_br(cfg.get("valor_c0_manual", 0)) > 0 else "Não informado"
    except Exception:
        c0_manual = "Não informado"

    ciclo = str(cfg.get("ciclo") or resultado_valor_global.get("ciclo_ultimo_remanescente") or "Não informado")
    competencia = str(cfg.get("competencia_corte") or cfg.get("data_corte") or "Não informado")
    fonte = str(cfg.get("fonte") or "Base financeira preferencial, quando disponível")

    if corte:
        linhas = [
            "<b>Metodologia:</b> corte operacional no ciclo em execução.",
            f"<b>Ciclo em execução:</b> {ciclo}.",
            f"<b>Competência de corte:</b> {competencia}.",
            f"<b>Fonte da execução:</b> {fonte}.",
            f"<b>C0 financeiro manual:</b> {c0_manual}.",
            f"<b>Execução atualizada considerada:</b> {execucao}.",
            f"<b>Saldo remanescente atualizado considerado:</b> {remanescente}.",
            f"<b>Valor Total Atualizado importado:</b> {valor_total}.",
        ]
        return f"""
        <div style="background:#ECFEFF; border:1px solid #14B8A6; border-left:6px solid #0F766E; border-radius:12px; padding:13px 15px; margin:10px 0 16px 0; color:#134E4A;">
            <div style="font-weight:900; margin-bottom:6px;">Metodologia do Valor Total Atualizado importado</div>
            <div style="font-size:0.92rem; line-height:1.48;">{'<br>'.join(linhas)}</div>
        </div>
        """

    return f"""
    <div style="background:#F8FAFC; border:1px solid #CBD5E1; border-left:6px solid #64748B; border-radius:12px; padding:13px 15px; margin:10px 0 16px 0; color:#334155;">
        <div style="font-weight:900; margin-bottom:6px;">Metodologia do Valor Total Atualizado importado</div>
        <div style="font-size:0.92rem; line-height:1.48;">
            <b>Metodologia:</b> corte padrão, sem corte operacional específico no ciclo em execução.<br>
            <b>Composição:</b> execução atualizada por ciclo + saldo remanescente atualizado.<br>
            <b>Execução atualizada considerada:</b> {execucao}.<br>
            <b>Saldo remanescente atualizado considerado:</b> {remanescente}.<br>
            <b>Valor Total Atualizado importado:</b> {valor_total}.
        </div>
    </div>
    """
# <<< METODOLOGIA_VALOR_IMPORTADO_CORTE_OPERACIONAL

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
        .garantia-label { color: #475569; font-size: 0.92rem; margin-bottom: 4px; }
        .garantia-valor { color: #1F2937; font-size: 1.55rem; font-weight: 700; line-height: 1.2; }
        .garantia-valor-destaque { color: #123B63; font-size: 2rem; font-weight: 800; line-height: 1.2; }
        .garantia-nota { color: #64748B; font-size: 0.88rem; margin-top: 6px; }
        .valor-formatado-apoio { color: #64748B; font-size: 0.86rem; margin-top: -8px; margin-bottom: 8px; }
        .linha-tempo-box {
            background: #FBFCFE;
            border: 1px solid #E1E7EF;
            border-radius: 14px;
            padding: 16px 18px;
            margin: 10px 0 12px 0;
        }
        .linha-tempo-titulo { color: #123B63; font-weight: 800; font-size: 1rem; margin-bottom: 5px; }
        .linha-tempo-texto { color: #334155; font-size: 0.94rem; margin-bottom: 0; }
        .garantia-tabela-wrap { width: 100%; overflow-x: auto; margin: 10px 0 18px 0; }
        table.garantia-tabela { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 0.92rem; }
        table.garantia-tabela th { background: #E6F0F7; color: #173B5D; border: 1px solid #C5D6E2; padding: 9px 10px; text-align: left; font-weight: 700; }
        table.garantia-tabela td { border: 1px solid #E5EAF0; padding: 9px 10px; vertical-align: top; white-space: normal; overflow-wrap: anywhere; word-break: normal; line-height: 1.35; }
        table.garantia-tabela td.valor { white-space: nowrap; overflow-wrap: normal; text-align: left; }
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




def ordenar_linha_tempo_garantia(df):
    """Ordena a linha do tempo: contrato original primeiro, eventos por data e reajustes atuais ao final.

    Também cria a coluna Ordem para permitir ajuste manual pelo usuário no editor.
    """
    if not isinstance(df, pd.DataFrame) or df.empty:
        return df

    base = df.copy()

    def rank_evento(row, idx):
        evento = limpar_texto(row.get("Evento", "")).lower()
        data_txt = limpar_texto(row.get("Data", ""))

        if evento.startswith("contrato original"):
            return -10_000
        if evento.startswith("reajustes atuais"):
            return 10_000_000

        dt = pd.to_datetime(data_txt, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return int(dt.strftime("%Y%m%d"))
        return 5_000_000 + idx

    base["_ordem_auto"] = [rank_evento(row, idx) for idx, row in base.iterrows()]
    base = base.sort_values(by=["_ordem_auto", "Evento"], kind="stable").drop(columns=["_ordem_auto"])
    base = base.reset_index(drop=True)
    base.insert(0, "Ordem", range(1, len(base) + 1))
    return base


def render_linha_tempo_garantia(df):
    """Renderiza a linha do tempo com quebra de texto adequada nas células."""
    if not isinstance(df, pd.DataFrame) or df.empty:
        st.info("Nenhum evento informado para a linha do tempo da garantia.")
        return
    colunas = ["Evento", "Data", "Ciclo", "Valor atualizado", "Endosso esperado"]
    linhas = []
    for _, row in df.iterrows():
        evento = escape(limpar_texto(row.get("Evento", "")))
        data = escape(limpar_texto(row.get("Data", "")))
        ciclo = escape(limpar_texto(row.get("Ciclo", "")))
        valor = escape(moeda(row.get("Valor atualizado", row.get("Valor-base", 0.0))))
        endosso = escape(moeda(row.get("Endosso esperado", 0.0)))
        linhas.append(
            f"<tr><td>{evento}</td><td>{data}</td><td>{ciclo}</td><td class='valor'>{valor}</td><td class='valor'>{endosso}</td></tr>"
        )
    html = """
    <div class="garantia-tabela-wrap">
      <table class="garantia-tabela">
        <colgroup>
          <col style="width: 28%;">
          <col style="width: 18%;">
          <col style="width: 14%;">
          <col style="width: 20%;">
          <col style="width: 20%;">
        </colgroup>
        <thead>
          <tr>
            <th>Evento</th><th>Data</th><th>Ciclo</th><th>Valor-base</th><th>Endosso esperado</th>
          </tr>
        </thead>
        <tbody>
          {linhas}
        </tbody>
      </table>
    </div>
    """.format(linhas="\n".join(linhas))
    st.markdown(html, unsafe_allow_html=True)

def gerar_pdf_garantia(dados, df_linha_tempo):
    if not REPORTLAB_OK:
        return None
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=1.6*cm, leftMargin=1.6*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    titulo = ParagraphStyle("TituloGarantia", parent=styles["Title"], alignment=TA_CENTER, fontSize=13, leading=16, spaceAfter=8)
    subtitulo = ParagraphStyle("SubtituloGarantia", parent=styles["Normal"], alignment=TA_CENTER, fontSize=9, leading=12, textColor=colors.HexColor("#334155"), spaceAfter=14)
    h2 = ParagraphStyle("H2Garantia", parent=styles["Heading2"], fontSize=10, leading=13, spaceBefore=8, spaceAfter=6, textColor=colors.HexColor("#123B63"))
    normal = ParagraphStyle("NormalGarantia", parent=styles["Normal"], fontSize=8.5, leading=12, alignment=TA_JUSTIFY)

    elementos = []
    elementos.append(Paragraph("Garantia Contratual", titulo))
    elementos.append(Paragraph("Histórico da garantia e endosso complementar", subtitulo))
    elementos.append(Paragraph(f"Gerado em: {data_hora_brasilia()}", subtitulo))

    elementos.append(Paragraph("1. Resultado", h2))
    tabela_dados = [
        ["Indicador", "Valor"],
        ["Garantia/endossos esperados acumulados", moeda(dados["garantia_esperada_acumulada"])],
        ["Garantia/endossos já apresentados", moeda(dados["garantia_apresentada"])],
        ["Endosso complementar estimado", moeda(dados["endosso_complementar"])],
        ["Percentual da garantia", f"{dados['percentual_garantia_pct']:.2f}%".replace(".", ",")],
        ["Valor Total Atualizado do Contrato", moeda(dados["valor_total_atualizado"])],
    ]
    tabela = Table(tabela_dados, colWidths=[8.2*cm, 7.8*cm], repeatRows=1, hAlign="CENTER")
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#D9E2F3")),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ]))
    elementos.append(tabela)
    elementos.append(Spacer(1, 10))

    if isinstance(df_linha_tempo, pd.DataFrame) and not df_linha_tempo.empty:
        elementos.append(Paragraph("2. Linha do tempo da garantia", h2))
        linhas = [["Evento", "Data", "Ciclo", "Valor-base", "Endosso esperado"]]
        for _, row in df_linha_tempo.iterrows():
            linhas.append([
                limpar_texto(row.get("Evento", "")),
                limpar_texto(row.get("Data", "")),
                limpar_texto(row.get("Ciclo", "")),
                moeda(row.get("Valor atualizado", row.get("Valor-base", 0.0))),
                moeda(row.get("Endosso esperado", 0.0)),
            ])
        tabela_lt = Table(linhas, colWidths=[4.0*cm, 2.5*cm, 2.2*cm, 3.8*cm, 3.5*cm], repeatRows=1, hAlign="CENTER")
        tabela_lt.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F4E78")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE", (0,0), (-1,-1), 7),
            ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#D9E2F3")),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ]))
        elementos.append(tabela_lt)
        elementos.append(Spacer(1, 10))

    elementos.append(Paragraph("3. Observação", h2))
    elementos.append(Paragraph(
        "Os valores apresentados são auxiliares à conferência da garantia. A definição da base de cálculo deve observar a cláusula contratual de garantia e a orientação administrativa aplicável.",
        normal,
    ))
    doc.build(elementos)
    buffer.seek(0)
    return buffer.getvalue()




def gerar_pdf_garantia_delta(dados):
    """Gera PDF objetivo para o Método 1 — Delta da Garantia."""
    if not REPORTLAB_OK:
        return None

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
    titulo = ParagraphStyle(
        "TituloGarantiaDelta",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontSize=13,
        leading=16,
        spaceAfter=6,
        textColor=colors.HexColor("#0B1F3A"),
    )
    subtitulo = ParagraphStyle(
        "SubtituloGarantiaDelta",
        parent=styles["Normal"],
        alignment=TA_CENTER,
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#334155"),
        spaceAfter=10,
    )
    h2 = ParagraphStyle(
        "H2GarantiaDelta",
        parent=styles["Heading2"],
        fontSize=10,
        leading=13,
        spaceBefore=7,
        spaceAfter=5,
        textColor=colors.HexColor("#123B63"),
    )
    normal = ParagraphStyle(
        "NormalGarantiaDelta",
        parent=styles["Normal"],
        fontSize=8.2,
        leading=11.2,
        alignment=TA_JUSTIFY,
        textColor=colors.HexColor("#1F2937"),
    )
    celula = ParagraphStyle(
        "CelulaGarantiaDelta",
        parent=styles["Normal"],
        fontSize=7.2,
        leading=8.8,
        alignment=TA_LEFT,
        textColor=colors.HexColor("#1F2937"),
    )
    celula_branca = ParagraphStyle(
        "CelulaGarantiaDeltaBranca",
        parent=celula,
        textColor=colors.white,
        fontName="Helvetica-Bold",
    )
    celula_negrito = ParagraphStyle(
        "CelulaGarantiaDeltaNegrito",
        parent=celula,
        fontName="Helvetica-Bold",
    )
    obs = ParagraphStyle(
        "ObservacaoGarantiaDelta",
        parent=styles["Normal"],
        fontSize=7.6,
        leading=9.6,
        alignment=TA_JUSTIFY,
        textColor=colors.HexColor("#475569"),
    )

    def _p(valor, estilo=celula):
        texto = escape(limpar_texto(valor)).replace("\n", "<br/>")
        return Paragraph(texto or "-", estilo)

    tipo_evento = dados.get("tipo_evento", "Evento atual")
    valor_base_evento = parse_moeda_br(dados.get("valor_base_evento", 0.0))
    percentual_garantia_pct = parse_moeda_br(dados.get("percentual_garantia_pct", 0.0))
    endosso_esperado = parse_moeda_br(dados.get("endosso_esperado_evento", 0.0))
    endosso_apresentado = parse_moeda_br(dados.get("endosso_apresentado_evento", 0.0))
    endosso_complementar = parse_moeda_br(dados.get("endosso_complementar_evento", 0.0))
    excesso = max(endosso_apresentado - endosso_esperado, 0.0)

    if endosso_complementar > 0.004:
        leitura = f"O evento analisado gera endosso complementar estimado de {moeda(endosso_complementar)}."
    elif excesso > 0.004:
        leitura = f"O endosso já apresentado para o evento supera o esperado em {moeda(excesso)}. Recomenda-se conferir validade e aceitação formal."
    else:
        leitura = "O endosso já apresentado para o evento corresponde ao valor esperado."

    elementos = []
    elementos.append(Paragraph("Garantia Contratual", titulo))
    elementos.append(Paragraph("Método 1 — Delta da Garantia", subtitulo))
    elementos.append(Paragraph(f"Gerado em: {data_hora_brasilia()}", subtitulo))

    elementos.append(Paragraph("1. Resultado executivo", h2))
    tabela_resultado = [
        [_p("Indicador", celula_branca), _p("Valor", celula_branca)],
        [_p("Tipo de evento"), _p(tipo_evento, celula_negrito)],
        [_p("Valor-base do evento atual"), _p(moeda(valor_base_evento), celula_negrito)],
        [_p("Percentual da garantia"), _p(f"{percentual_garantia_pct:.2f}%".replace(".", ","))],
        [_p("Endosso esperado do evento"), _p(moeda(endosso_esperado), celula_negrito)],
        [_p("Endosso já apresentado para o evento"), _p(moeda(endosso_apresentado), celula_negrito)],
        [_p("Endosso complementar do evento"), _p(moeda(endosso_complementar), celula_negrito)],
        [_p("Prazo de referência para apresentação/endosso"), _p(f"{int(dados.get('prazo_dias', 0) or 0)} dias úteis")],
        [_p("Encerramento da vigência contratual"), _p(dados.get("data_fim_vigencia", "[campo a preencher]"))],
        [_p("Validade mínima sugerida da garantia"), _p(dados.get("data_validade_minima", "[campo a preencher]"))],
    ]
    tabela = Table(tabela_resultado, colWidths=[8.2 * cm, 7.6 * cm], repeatRows=1, hAlign="CENTER")
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.4),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9E2F3")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F8FAFC")),
        ("BACKGROUND", (0, 6), (-1, 6), colors.HexColor("#EAF2F8")),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    elementos.append(tabela)
    elementos.append(Spacer(1, 7))
    elementos.append(Paragraph(leitura, normal))

    elementos.append(Paragraph("2. Memória de cálculo", h2))
    elementos.append(Paragraph(
        "Endosso esperado do evento = valor-base do evento atual × percentual da garantia.",
        normal,
    ))
    elementos.append(Paragraph(
        "Endosso complementar do evento = endosso esperado do evento − endosso já apresentado para esse evento, limitado a zero quando o resultado for negativo.",
        normal,
    ))

    elementos.append(Paragraph("3. Observações", h2))
    elementos.append(Paragraph(
        "O Método 1 — Delta da Garantia calcula apenas o reforço de garantia gerado pelo evento específico informado. Ele pressupõe que a garantia anterior estava regular e não substitui, quando necessário, a conferência histórica completa da garantia contratual.",
        obs,
    ))
    elementos.append(Paragraph(
        "Para auditoria da suficiência global da garantia, recomenda-se utilizar o Método 2 — Linha do Tempo Completa.",
        obs,
    ))

    doc.build(elementos)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# Interface
# ============================================================

css()
render_marca_topo()
if st.button("← Voltar para Central", key="voltar_central_garantia"):
    st.switch_page("pages/03_Valor_Global.py")
st.title("Garantia Contratual")
st.write("Escolha o método de cálculo da garantia. O Método Delta calcula o reforço do evento atual; a Linha do Tempo Completa serve para conferência histórica.")

render_aviso_privacidade(tem_download=True)
resultado_valor_global = st.session_state.get("resultado_valor_global", {}) or {}
modo_apuracao_vg = resultado_valor_global.get("modo_apuracao", "Completo") if isinstance(resultado_valor_global, dict) else "Completo"
ctx = obter_contexto_valor_global(resultado_valor_global)
st.markdown(_metodologia_valor_importado_html(resultado_valor_global), unsafe_allow_html=True)

valor_original_padrao = ctx["valor_original"]
valor_total_atualizado_padrao = ctx["valor_total_atualizado"]
variacao_acumulada_padrao = ctx["variacao_acumulada"]
df_aditivos_importado = ctx["df_aditivos"]
df_ciclos_importado = ctx["df_ciclos"]
modo_reduzido_estoque = ctx.get("modo_apuracao") == "Reduzido por Itens/Estoque"
modo_consumo_itens_ciclo = ctx.get("modo_apuracao") == "Consumo por Itens/Ciclo"
if modo_consumo_itens_ciclo:
    st.markdown(
        """
        <div style="background:#F6F3EE; border:1px solid #7A8F63; border-left:6px solid #4E6E58; border-radius:12px; padding:14px 16px; margin:10px 0 16px 0; color:#2F3E2F;">
            <div style="font-weight:800; margin-bottom:4px;">Modo Consumo por Itens/Ciclo</div>
            <div style="font-size:0.95rem; line-height:1.45;">A base da garantia decorre do Valor Total Atualizado do Contrato apurado por consumo itemizado e saldo remanescente atualizado. Use a base importada como referência operacional, mantendo a validação fiscal da equivalência consumo/execução/faturamento.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
elif modo_reduzido_estoque:
    st.markdown(
        """
        <div style="background:#F3E8FF; border:1px solid #A855F7; border-left:6px solid #7E22CE; border-radius:12px; padding:14px 16px; margin:10px 0 16px 0; color:#581C87;">
            <div style="font-weight:800; margin-bottom:4px;">Modo Reduzido por Itens/Estoque</div>
            <div style="font-size:0.95rem; line-height:1.45;">A base da garantia pode decorrer de estimativa por itens/remanescentes, sem conciliação com a base mensal de execução. Use como apoio operacional e valide a base antes da formalização.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

metodo_garantia = st.radio(
    "Método de cálculo da garantia",
    options=["Método 1 — Delta da Garantia", "Método 2 — Linha do Tempo Completa"],
    index=0,
    horizontal=True,
    help="Use o Método Delta para calcular o endosso de um evento específico. Use a Linha do Tempo Completa para conferir todo o histórico da garantia.",
)

if metodo_garantia == "Método 1 — Delta da Garantia":
    st.subheader("Método 1 — Delta da Garantia")
    st.caption(
        "Este método calcula o endosso complementar gerado por um evento específico. "
        "Ele pressupõe que a garantia anterior estava regular. Para conferência histórica completa, use o Método 2."
    )

    with st.expander("Contexto importado da análise atual", expanded=False):
        if resultado_valor_global:
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                st.metric("Valor original identificado", moeda(valor_original_padrao))
            with col_b:
                st.metric("Valor Total Atualizado do Contrato", moeda(valor_total_atualizado_padrao))
            with col_c:
                st.metric("Aditivos identificados", int(ctx.get("quantidade_aditivos", 0)))
            if modo_consumo_itens_ciclo:
                st.metric("Retroativo (itens consumidos/ciclo)", moeda(ctx.get("retroativo_estimado_itens", 0.0)))
                st.caption("Esses dados decorrem de consumo itemizado por ciclo. O Método Delta usa o valor-base do evento informado abaixo.")
            elif modo_reduzido_estoque:
                st.metric("Retroativo estimado por itens/estoque", moeda(ctx.get("retroativo_estimado_itens", 0.0)))
                st.caption("Esses dados são estimativos quando o Valor Global foi processado sem base de execução mensal por competência. O Método Delta usa o valor-base do evento informado abaixo.")
            else:
                st.caption("Esses dados são apenas referências. O Método Delta usa o valor-base do evento informado abaixo.")
        else:
            st.info("Não há dados do módulo Valores na sessão atual. Informe os dados manualmente.")

    st.markdown("### Dados do evento atual")
    tipo_evento = st.selectbox(
        "Tipo do evento",
        [
            "Reajuste",
            "Aditivo de acréscimo",
            "Repactuação",
            "Reequilíbrio econômico-financeiro",
            "Outro evento com aumento de base",
        ],
        index=0,
    )

    valor_formalizado_anterior = parse_moeda_br(ctx.get("valor_formalizado_anterior", 0.0))
    sugestao_delta = max(valor_total_atualizado_padrao - valor_formalizado_anterior, 0.0)
    if sugestao_delta <= 0 and valor_total_atualizado_padrao > 0 and valor_original_padrao > 0:
        sugestao_delta = max(valor_total_atualizado_padrao - valor_original_padrao, 0.0)

    col_d1, col_d2, col_d3 = st.columns(3)
    with col_d1:
        valor_base_evento_txt = st.text_input(
            "Valor-base do evento atual",
            value=moeda(sugestao_delta, com_prefixo=False),
            help="Informe o aumento de base gerado pelo evento atual. Ex.: valor do reajuste, valor do aditivo ou delta contratual do evento.",
        )
        valor_base_evento = parse_moeda_br(valor_base_evento_txt)
        st.markdown(f"<div class='valor-formatado-apoio'>{moeda(valor_base_evento)}</div>", unsafe_allow_html=True)

    with col_d2:
        percentual_garantia_pct = st.number_input(
            "Percentual da garantia (%)",
            min_value=0.0,
            max_value=100.0,
            value=5.0,
            step=0.1,
            format="%.2f",
        )

    with col_d3:
        endosso_apresentado_txt = st.text_input(
            "Endosso já apresentado para este evento",
            value="0,00",
            help="Informe apenas eventual endosso já apresentado especificamente para o evento atual.",
        )
        endosso_apresentado_evento = parse_moeda_br(endosso_apresentado_txt)
        st.markdown(f"<div class='valor-formatado-apoio'>{moeda(endosso_apresentado_evento)}</div>", unsafe_allow_html=True)

    col_p1, col_p2 = st.columns(2)
    with col_p1:
        prazo_dias = st.number_input("Prazo para apresentação/endosso (dias úteis)", min_value=1, max_value=60, value=5, step=1)
    with col_p2:
        data_fim_vigencia = st.date_input("Encerramento da vigência contratual", value=date.today(), format="DD/MM/YYYY")

    data_validade_minima = data_fim_vigencia + relativedelta(months=3)
    st.caption(f"Validade mínima sugerida da garantia: {data_validade_minima.strftime('%d/%m/%Y')}")

    percentual_garantia = percentual_garantia_pct / 100
    endosso_esperado_evento = round(valor_base_evento * percentual_garantia, 2)
    endosso_complementar_evento = round(max(endosso_esperado_evento - endosso_apresentado_evento, 0.0), 2)
    excesso_evento = round(max(endosso_apresentado_evento - endosso_esperado_evento, 0.0), 2)

    st.divider()
    st.subheader("Resultado do delta")
    colr1, colr2, colr3 = st.columns(3)
    with colr1:
        card("Endosso esperado do evento", moeda(endosso_esperado_evento), "Valor-base do evento × percentual da garantia.")
    with colr2:
        card("Endosso já apresentado para o evento", moeda(endosso_apresentado_evento), "Valor informado pelo usuário.")
    with colr3:
        card("Endosso complementar do evento", moeda(endosso_complementar_evento), f"Prazo de referência: {prazo_dias} dias úteis.", destaque=True)

    if endosso_complementar_evento > 0:
        st.warning(f"O evento analisado gera endosso complementar estimado de {moeda(endosso_complementar_evento)}.")
    elif excesso_evento > 0:
        st.success(f"O endosso já apresentado para o evento supera o esperado em {moeda(excesso_evento)}. Confira validade e aceitação.")
    else:
        st.success("O endosso já apresentado para o evento corresponde ao valor esperado.")

    with st.expander("Memória de cálculo", expanded=False):
        memoria_delta = pd.DataFrame([
            {"Indicador": "Método", "Valor": "Método 1 — Delta da Garantia"},
            {"Indicador": "Tipo do evento", "Valor": tipo_evento},
            {"Indicador": "Valor-base do evento atual", "Valor": moeda(valor_base_evento)},
            {"Indicador": "Percentual da garantia", "Valor": f"{percentual_garantia_pct:.2f}%".replace(".", ",")},
            {"Indicador": "Endosso esperado do evento", "Valor": moeda(endosso_esperado_evento)},
            {"Indicador": "Endosso já apresentado para o evento", "Valor": moeda(endosso_apresentado_evento)},
            {"Indicador": "Endosso complementar do evento", "Valor": moeda(endosso_complementar_evento)},
            {"Indicador": "Encerramento da vigência", "Valor": data_fim_vigencia.strftime("%d/%m/%Y")},
            {"Indicador": "Validade mínima da garantia", "Valor": data_validade_minima.strftime("%d/%m/%Y")},
        ])
        st.dataframe(memoria_delta, use_container_width=True, hide_index=True)

    dados_pdf_delta = {
        "tipo_evento": tipo_evento,
        "valor_base_evento": valor_base_evento,
        "percentual_garantia_pct": percentual_garantia_pct,
        "endosso_esperado_evento": endosso_esperado_evento,
        "endosso_apresentado_evento": endosso_apresentado_evento,
        "endosso_complementar_evento": endosso_complementar_evento,
        "prazo_dias": prazo_dias,
        "data_fim_vigencia": data_fim_vigencia.strftime("%d/%m/%Y"),
        "data_validade_minima": data_validade_minima.strftime("%d/%m/%Y"),
    }

    st.session_state["resultado_garantia"] = {
        "metodo_garantia": "Delta da Garantia",
        "modo_apuracao_valor_global": modo_apuracao_vg,
        "base_estimativa_por_itens_estoque": modo_apuracao_vg == "Reduzido por Itens/Estoque",
        "tipo_evento": tipo_evento,
        "valor_base_evento": valor_base_evento,
        "percentual_garantia": percentual_garantia,
        "endosso_esperado_evento": endosso_esperado_evento,
        "endosso_apresentado_evento": endosso_apresentado_evento,
        "endosso_necessario": endosso_complementar_evento,
        "data_fim_vigencia": data_fim_vigencia,
        "data_validade_minima": data_validade_minima,
    }

    st.subheader("Relatório")
    if REPORTLAB_OK:
        pdf_bytes = gerar_pdf_garantia_delta(dados_pdf_delta)
        st.session_state["arquivo_garantia_pdf"] = pdf_bytes
        st.download_button(
            "Baixar relatório de garantia (PDF)",
            data=pdf_bytes,
            file_name="relatorio_garantia_delta.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=False,
        )
    else:
        st.info("Instale reportlab para gerar o PDF: pip install reportlab")

else:
    with st.expander("Contexto importado da análise atual", expanded=True):
        if resultado_valor_global:
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                st.metric("Valor original identificado", moeda(valor_original_padrao))
            with col_b:
                st.metric("Valor Total Atualizado do Contrato", moeda(valor_total_atualizado_padrao))
            with col_c:
                st.metric("Aditivos identificados", int(ctx.get("quantidade_aditivos", 0)))
            st.caption("O Valor Total Atualizado é importado apenas como base de conferência da garantia. A regra do módulo Valores não é alterada.")
        else:
            st.info("Não há dados do módulo Valores na sessão atual. Informe os dados manualmente.")

    st.subheader("Dados-base")
    st.caption("O sistema sugere valores quando houver análise carregada, mas você pode ajustar manualmente.")

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
        valor_total_atualizado_txt = st.text_input(
            "Valor Total Atualizado do Contrato",
            value=moeda(valor_total_atualizado_padrao, com_prefixo=False),
            help="Valor importado do módulo Valores ou informado manualmente para conferência da garantia.",
        )
        valor_total_atualizado = parse_moeda_br(valor_total_atualizado_txt)
        st.markdown(f"<div class='valor-formatado-apoio'>{moeda(valor_total_atualizado)}</div>", unsafe_allow_html=True)

    percentual_garantia = percentual_garantia_pct / 100
    valor_garantia_original = round(valor_original * percentual_garantia, 2)

    # Aditivos importados ou manuais
    st.subheader("Histórico da garantia")
    st.caption("Confira os eventos identificados e ajuste os valores quando necessário.")

    df_aditivos_base = extrair_aditivos_para_garantia(df_aditivos_importado)
    if not df_aditivos_base.empty:
        df_aditivos_base["Endosso esperado"] = df_aditivos_base["Valor atualizado"].apply(lambda v: round(float(v) * percentual_garantia, 2))

    linha_contrato = {
        "Evento": "Contrato original",
        "Data": "Assinatura/base inicial",
        "Ciclo": "C0",
        "Valor original": valor_original,
        "Valor atualizado": valor_original,
        "Endosso esperado": valor_garantia_original,
    }

    df_linha_tempo_base = pd.concat([pd.DataFrame([linha_contrato]), df_aditivos_base], ignore_index=True)

    # Reajuste atual como evento de fechamento da linha do tempo.
    texto_ciclos = resumo_ciclos_texto(df_ciclos_importado, variacao_acumulada_padrao)
    garantia_exigida_total = round(valor_total_atualizado * percentual_garantia, 2)
    endossos_parciais = round(float(df_linha_tempo_base["Endosso esperado"].sum()), 2) if not df_linha_tempo_base.empty else 0.0
    endosso_reajustes_atuais = round(garantia_exigida_total - endossos_parciais, 2)
    linha_reajustes = {
        "Evento": f"Reajustes atuais — {texto_ciclos}",
        "Data": "Análise atual",
        "Ciclo": "Atual",
        "Valor original": 0.0,
        "Valor atualizado": max(valor_total_atualizado - float(df_linha_tempo_base["Valor atualizado"].sum()), 0.0) if not df_linha_tempo_base.empty else valor_total_atualizado,
        "Endosso esperado": endosso_reajustes_atuais,
    }

    df_linha_tempo = pd.concat([df_linha_tempo_base, pd.DataFrame([linha_reajustes])], ignore_index=True)
    df_linha_tempo = ordenar_linha_tempo_garantia(df_linha_tempo)

    st.markdown("### Linha do tempo da garantia")
    st.caption(
        "Os eventos são ordenados automaticamente por data. O contrato original fica no início e os reajustes atuais ficam ao final. "
        "Se precisar, ajuste a coluna Ordem."
    )

    # Editor manual da linha do tempo
    editor_df = df_linha_tempo.copy()
    for col in ["Valor original", "Valor atualizado", "Endosso esperado"]:
        editor_df[col] = editor_df[col].apply(moeda)

    with st.expander("Ajustar linha do tempo da garantia", expanded=True):
        st.caption(
            "Edite os valores se o histórico real do contrato for diferente do que foi importado. "
            "A coluna Ordem permite reorganizar manualmente a sequência. "
            "O campo ‘Endosso esperado’ alimenta o contador acumulado."
        )
        editado = st.data_editor(
            editor_df,
            hide_index=True,
            use_container_width=True,
            num_rows="dynamic",
            key="garantia_linha_tempo_editor",
            column_config={
                "Ordem": st.column_config.NumberColumn("Ordem", min_value=1, step=1),
                "Evento": st.column_config.TextColumn("Evento"),
                "Data": st.column_config.TextColumn("Data"),
                "Ciclo": st.column_config.TextColumn("Ciclo"),
                "Valor original": st.column_config.TextColumn("Valor original"),
                "Valor atualizado": st.column_config.TextColumn("Valor atualizado"),
                "Endosso esperado": st.column_config.TextColumn("Endosso esperado"),
            },
        )

    for col in ["Valor original", "Valor atualizado", "Endosso esperado"]:
        editado[col] = editado[col].apply(parse_moeda_br)

    if "Ordem" in editado.columns:
        editado["Ordem"] = pd.to_numeric(editado["Ordem"], errors="coerce").fillna(9999).astype(int)
        editado = editado.sort_values(by=["Ordem"], kind="stable").reset_index(drop=True)

    render_linha_tempo_garantia(editado)

    garantia_esperada_acumulada = round(float(editado["Endosso esperado"].sum()), 2) if not editado.empty else 0.0

    st.markdown("### Garantia/endossos já apresentados")
    col_ap1, col_ap2 = st.columns(2)
    with col_ap1:
        garantia_apresentada_txt = st.text_input(
            "Total de garantia/endossos já apresentados (R$)",
            value=moeda(valor_garantia_original, com_prefixo=False),
            placeholder="Ex.: 7.914.629,92",
            help="Informe o valor monetário total da garantia original e dos endossos já aceitos pela Administração. Não é quantidade de endossos.",
        )
        garantia_apresentada = parse_moeda_br(garantia_apresentada_txt)
        st.caption("Informe o valor em reais. Ex.: 7.914.629,92, e não a quantidade de endossos.")
        st.markdown(f"<div class='valor-formatado-apoio'>{moeda(garantia_apresentada)}</div>", unsafe_allow_html=True)
    with col_ap2:
        prazo_dias = st.number_input("Prazo para apresentação/endosso (dias úteis)", min_value=1, max_value=60, value=5, step=1)

    data_fim_vigencia = st.date_input("Encerramento da vigência contratual", value=date.today(), format="DD/MM/YYYY")
    data_validade_minima = data_fim_vigencia + relativedelta(months=3)
    st.caption(f"Validade mínima sugerida da garantia: {data_validade_minima.strftime('%d/%m/%Y')}")

    endosso_complementar = round(max(garantia_esperada_acumulada - garantia_apresentada, 0.0), 2)
    excesso_garantia = round(max(garantia_apresentada - garantia_esperada_acumulada, 0.0), 2)

    st.divider()
    st.subheader("Resultado acumulado")
    colr1, colr2, colr3 = st.columns(3)
    with colr1:
        card("Garantia/endossos esperados acumulados", moeda(garantia_esperada_acumulada), "Soma dos endossos esperados na linha do tempo.")
    with colr2:
        card("Garantia/endossos já apresentados", moeda(garantia_apresentada), "Valor informado pelo usuário.")
    with colr3:
        card("Endosso complementar estimado", moeda(endosso_complementar), f"Prazo de referência: {prazo_dias} dias úteis.", destaque=True)

    if endosso_complementar > 0:
        st.warning(f"Pelo histórico informado, ainda haveria endosso complementar estimado de {moeda(endosso_complementar)}.")
    elif excesso_garantia > 0:
        st.success(f"Pelo histórico informado, a garantia/endossos apresentados superam o esperado em {moeda(excesso_garantia)}. Confira validade e aceitação.")
    else:
        st.success("Pelo histórico informado, a garantia/endossos apresentados correspondem ao esperado.")

    with st.expander("Calcule de outra forma", expanded=False):
        st.caption(
            "Use este bloco quando quiser montar uma linha do tempo alternativa. "
            "Adicione as linhas necessárias; o endosso é calculado automaticamente por valor-base × percentual."
        )
        df_alt_base = pd.DataFrame([
            {"Evento": "Valor original", "Valor-base": moeda(valor_original, com_prefixo=False), "Percentual da garantia (%)": f"{percentual_garantia_pct:.2f}".replace(".", ",")},
            {"Evento": "Aditivo 1", "Valor-base": "0,00", "Percentual da garantia (%)": f"{percentual_garantia_pct:.2f}".replace(".", ",")},
        ])
        df_alt = st.data_editor(
            df_alt_base,
            hide_index=True,
            use_container_width=True,
            num_rows="dynamic",
            key="garantia_calculo_alternativo_linhas",
            column_config={
                "Evento": st.column_config.TextColumn("Evento"),
                "Valor-base": st.column_config.TextColumn("Valor-base"),
                "Percentual da garantia (%)": st.column_config.TextColumn("Percentual da garantia (%)"),
            },
        )
        if isinstance(df_alt, pd.DataFrame) and not df_alt.empty:
            df_alt_calc = df_alt.copy()
            df_alt_calc["_valor_base"] = df_alt_calc["Valor-base"].apply(parse_moeda_br)
            df_alt_calc["_percentual"] = df_alt_calc["Percentual da garantia (%)"].apply(parse_moeda_br)
            df_alt_calc["Endosso esperado"] = df_alt_calc.apply(
                lambda r: round(float(r.get("_valor_base", 0.0)) * float(r.get("_percentual", 0.0)) / 100, 2),
                axis=1,
            )
            total_alt = round(float(df_alt_calc["Endosso esperado"].sum()), 2)
            col_alt_r1, col_alt_r2 = st.columns(2)
            col_alt_r1.metric("Total de endossos pela linha alternativa", moeda(total_alt))
            col_alt_r2.metric("Diferença frente ao apresentado", moeda(max(total_alt - garantia_apresentada, 0.0)))
            df_alt_view = df_alt_calc[["Evento", "Valor-base", "Percentual da garantia (%)", "Endosso esperado"]].copy()
            df_alt_view["Endosso esperado"] = df_alt_view["Endosso esperado"].apply(moeda)
            st.dataframe(df_alt_view, use_container_width=True, hide_index=True)

    with st.expander("Memória de cálculo", expanded=False):
        memoria = [
            {"Indicador": "Valor original do contrato", "Valor": moeda(valor_original)},
            {"Indicador": "Percentual da garantia", "Valor": f"{percentual_garantia_pct:.2f}%".replace(".", ",")},
            {"Indicador": "Garantia original", "Valor": moeda(valor_garantia_original)},
            {"Indicador": "Valor Total Atualizado do Contrato", "Valor": moeda(valor_total_atualizado)},
            {"Indicador": "Garantia exigida pelo Valor Total Atualizado", "Valor": moeda(garantia_exigida_total)},
            {"Indicador": "Garantia/endossos esperados pela linha do tempo", "Valor": moeda(garantia_esperada_acumulada)},
            {"Indicador": "Garantia/endossos já apresentados", "Valor": moeda(garantia_apresentada)},
            {"Indicador": "Endosso complementar estimado", "Valor": moeda(endosso_complementar)},
            {"Indicador": "Encerramento da vigência", "Valor": data_fim_vigencia.strftime("%d/%m/%Y")},
            {"Indicador": "Validade mínima da garantia", "Valor": data_validade_minima.strftime("%d/%m/%Y")},
        ]
        st.dataframe(memoria, use_container_width=True, hide_index=True)
        st.markdown("**Linha do tempo considerada**")
        st.dataframe(editado.assign(
            **{
                "Valor original": editado["Valor original"].apply(moeda),
                "Valor atualizado": editado["Valor atualizado"].apply(moeda),
                "Endosso esperado": editado["Endosso esperado"].apply(moeda),
            }
        ), use_container_width=True, hide_index=True)

    st.subheader("Relatório")
    dados_pdf = {
        "garantia_esperada_acumulada": garantia_esperada_acumulada,
        "garantia_apresentada": garantia_apresentada,
        "endosso_complementar": endosso_complementar,
        "percentual_garantia_pct": percentual_garantia_pct,
        "valor_total_atualizado": valor_total_atualizado,
    }

    st.session_state["resultado_garantia"] = {
        "metodo_garantia": "Linha do Tempo Completa",
        "valor_original": valor_original,
        "valor_total_atualizado_contrato": valor_total_atualizado,
        "percentual_garantia": percentual_garantia,
        "garantia_original": valor_garantia_original,
        "garantia_exigida": garantia_esperada_acumulada,
        "garantia_constituida": garantia_apresentada,
        "endosso_necessario": endosso_complementar,
        "data_fim_vigencia": data_fim_vigencia,
        "data_validade_minima": data_validade_minima,
        "linha_tempo_garantia": editado,
    }

    if REPORTLAB_OK:
        pdf_bytes = gerar_pdf_garantia(dados_pdf, editado)
        st.session_state["arquivo_garantia_pdf"] = pdf_bytes
        st.download_button(
            "Baixar relatório de garantia (PDF)",
            data=pdf_bytes,
            file_name="relatorio_garantia.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=False,
        )
    else:
        st.info("Instale reportlab para gerar o PDF: pip install reportlab")
