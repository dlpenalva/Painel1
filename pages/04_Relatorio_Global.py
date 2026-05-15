from io import BytesIO
import re
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Análises de Reajustes - Relatório Global", layout="wide")


from _ui_utils import render_marca_topo

def moeda(valor):
    try:
        valor = round(float(valor), 2)
    except Exception:
        valor = 0.0
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def texto_seguro(valor, padrao="Não"):
    if valor is None:
        return padrao
    try:
        if pd.isna(valor):
            return padrao
    except Exception:
        pass
    texto = str(valor).strip()
    if texto.lower() in ["", "nan", "none", "nat", "<na>"]:
        return padrao
    return texto


def moeda_ou_texto(valor):
    """Formata números como moeda e preserva textos explicativos."""
    if isinstance(valor, str):
        texto = valor.strip()
        if texto and not any(ch.isdigit() for ch in texto):
            return texto
        if texto and any(ch.isdigit() for ch in texto):
            # Permite números em formato brasileiro, preservando textos mistos não numéricos.
            limpo = texto.replace("R$", "").replace(".", "").replace(",", ".").strip()
            try:
                return moeda(float(limpo))
            except Exception:
                return texto
    try:
        return moeda(float(valor))
    except Exception:
        return "" if pd.isna(valor) else str(valor)


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



def _normalizar_modo_apuracao(valor):
    texto = texto_seguro(valor, "")
    texto = texto.lower()
    mapa = str.maketrans("áàâãéêíóôõúç", "aaaaeeiooouc")
    texto = texto.translate(mapa)
    return re.sub(r"\s+", " ", texto).strip()


def eh_modo_reduzido_itens(res):
    if not isinstance(res, dict):
        return False
    modo = _normalizar_modo_apuracao(res.get("modo_apuracao", ""))
    return "reduzido" in modo and ("item" in modo or "estoque" in modo)


def valor_retroativo_estimado_itens(res):
    if not isinstance(res, dict):
        return 0.0
    for chave in [
        "valor_retroativo_estimado_itens_estoque",
        "retroativo_estimado_itens_estoque",
        "valor_retroativo_itens_estoque",
        "retroativo_itens_estoque",
    ]:
        if chave in res:
            try:
                return float(res.get(chave) or 0)
            except Exception:
                return 0.0
    return 0.0


def df_retroativo_estimado_itens(res):
    if not isinstance(res, dict):
        return pd.DataFrame()
    df = res.get("df_retroativo_estimado_itens_estoque")
    if isinstance(df, pd.DataFrame):
        return df
    return pd.DataFrame()


def indicadores_executivos_relatorio(res):
    if eh_modo_reduzido_itens(res):
        return [
            ["Indicador", "Valor"],
            ["Modo de apuração", texto_seguro(res.get("modo_apuracao"), "Reduzido por Itens/Estoque")],
            ["Valor original", moeda(res.get("valor_original_contrato", 0))],
            ["Valor Total Atualizado do Contrato", moeda(res.get("valor_atualizado_contrato", res.get("valor_global_estoque", 0)))],
            ["Retroativo financeiro definitivo", "Não calculado"],
            ["Retroativo estimado por itens/estoque", moeda(valor_retroativo_estimado_itens(res))],
            ["Execução estimada por itens/estoque", moeda(res.get("valor_executado_atualizado", 0))],
            ["Saldo remanescente atualizado", moeda(res.get("remanescente_reajustado", 0))],
            ["Aditivos/supressões registrados", moeda(res.get("total_aditivos_atualizados", 0))],
            ["Base mensal por competência", "Não informada"],
        ]
    return [
        ["Indicador", "Valor"],
        ["Valor original", moeda(res.get("valor_original_contrato", 0))],
        ["Valor pago efetivo", moeda(res.get("total_pago_faturado", 0))],
        ["Valor teórico calculado", moeda(res.get("total_devido_reajustado", 0))],
        ["Valor represado a pagar", moeda(res.get("valor_represado_a_pagar", 0))],
        ["Valor executado atualizado por ciclos", moeda(res.get("valor_executado_atualizado", 0))],
        ["Saldo remanescente atualizado", moeda(res.get("remanescente_reajustado", 0))],
        ["Valor total de aditivos/supressões", moeda(res.get("total_aditivos_atualizados", 0))],
        ["Valor Total Atualizado do Contrato", moeda(res.get("valor_atualizado_contrato", res.get("valor_global_estoque", 0)))],
    ]


def aviso_modo_reduzido_html():
    return """
    <div class="modo-reduzido-box">
        <div class="modo-reduzido-titulo">Modo Reduzido por Itens/Estoque</div>
        <div class="modo-reduzido-texto">
            A análise foi processada sem base mensal por competência. O retroativo financeiro definitivo não é calculado neste modo.
            O valor exibido como retroativo estimado por itens/estoque possui natureza estimativa e deve ser validado antes de qualquer formalização de pagamento.
        </div>
    </div>
    """


def formatar_data_br(valor):
    data = pd.to_datetime(valor, dayfirst=True, errors="coerce")
    if pd.isna(data):
        return "" if pd.isna(valor) else str(valor)
    return data.strftime("%d/%m/%Y")


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


def _status_relatorio(valor):
    """Status textual, sem emojis e sem caixa alta, para relatórios executivos."""
    texto = texto_seguro(valor, "")
    texto = re.sub(r"^[▲■●•\-\s]+", "", str(texto)).strip()
    if not texto:
        return "não informado"
    t = texto.upper()
    if "NEGOC" in t or "ACORDO" in t:
        return "aceito por negociação entre as partes"
    if "PRECLUS" in t:
        return "precluso"
    if "RESSALVA" in t:
        return "admissível com ressalva"
    if "TEMPEST" in t:
        return "tempestivo"
    if "ADIANT" in t:
        return "adiantado"
    return texto.lower()


def _ciclos_df_para_relatorio(adm, res):
    """Monta quadro de ciclos para relatório, incluindo C0 como ciclo-base."""
    df = None
    if isinstance(res, dict) and isinstance(res.get("df_ciclos"), pd.DataFrame) and not res.get("df_ciclos").empty:
        df = res.get("df_ciclos").copy()
    else:
        ciclos = []
        if adm:
            ciclos = adm.get("ciclos") or adm.get("detalhamento_ciclos") or []
        df = pd.DataFrame(ciclos)

    linhas = [{
        "Ciclo": "C0",
        "Data-base": "ciclo-base inicial",
        "Data do pedido": "não se aplica",
        "Classificação": "base sem reajuste",
        "Percentual aplicado": "0,00%",
        "Observação": "Período inicial do contrato, sem aplicação de reajuste.",
    }]

    if isinstance(df, pd.DataFrame) and not df.empty:
        for _, row in df.iterrows():
            ciclo = texto_seguro(row.get("Ciclo", row.get("ciclo", "")), "")
            if not ciclo or str(ciclo).upper() == "C0":
                continue
            acordo = texto_seguro(row.get("Acordo negocial", row.get("acordo_negocial", "")), "")
            situacao = row.get("Situação aplicada", row.get("situacao_aplicada", row.get("Situação", row.get("situacao", ""))))
            if str(acordo).strip().lower() in ["sim", "s", "true", "1"]:
                classificacao = "aceito por negociação entre as partes"
            else:
                classificacao = _status_relatorio(situacao)

            percentual_val = row.get("Percentual aplicado", row.get("percentual_aplicado", row.get("Variação", row.get("variacao", 0))))
            try:
                percentual_txt = percentual(float(percentual_val), 2)
            except Exception:
                percentual_txt = texto_seguro(percentual_val, "0,00%")

            obs = []
            if classificacao == "aceito por negociação entre as partes":
                obs.append("Tratamento negocial informado na análise.")
            justificativa = texto_seguro(row.get("Justificativa negocial", row.get("justificativa_negocial", "")), "")
            if justificativa:
                obs.append(justificativa)

            linhas.append({
                "Ciclo": ciclo,
                "Data-base": formatar_data_br(row.get("Data-base", row.get("data_base", ""))),
                "Data do pedido": formatar_data_br(row.get("Data do pedido", row.get("data_pedido", ""))),
                "Classificação": classificacao,
                "Percentual aplicado": percentual_txt,
                "Observação": " ".join(obs) if obs else "",
            })

    return pd.DataFrame(linhas)


def texto_clausula_oito(adm):
    """Síntese enxuta, sem afirmar tempestividade quando houver tratamento negocial."""
    ciclos = []
    if adm:
        ciclos = adm.get("ciclos") or adm.get("detalhamento_ciclos") or []

    if not ciclos:
        return (
            "A análise deve ser lida a partir da Cláusula Oitava, separando a existência do ciclo, "
            "a admissibilidade do pedido e o início dos efeitos financeiros."
        )

    qtd = len(ciclos)
    plural = "pleitos" if qtd > 1 else "pleito"
    status_gerais = []
    tem_negocial = False
    for c in ciclos:
        status_gerais.append(str(c.get("situacao") or c.get("Situação") or c.get("status") or c.get("situacao_aplicada") or "").upper())
        if c.get("superacao_negocial") or str(c.get("Acordo negocial", "")).strip().lower() in ["sim", "s", "true", "1"]:
            tem_negocial = True
    texto_status = " ".join(status_gerais)

    partes = [
        f"A análise contempla {qtd} {plural} de reajuste, devendo distinguir ciclo, admissibilidade e efeitos financeiros."
    ]
    if "PRECLUS" in texto_status:
        partes.append("Há ciclo classificado como precluso, mantido para memória e rastreabilidade, sem gerar retroativo a pagar.")
    if tem_negocial or "NEGOC" in texto_status or "ACORDO" in texto_status:
        partes.append("Há ciclo aceito por negociação entre as partes, razão pela qual o relatório não deve descrevê-lo como simples pleito tempestivo.")
    if "RESSALVA" in texto_status:
        partes.append("Há pleito admissível com ressalva, com efeitos financeiros a serem conferidos conforme a cláusula contratual.")
    if not any(["PRECLUS" in texto_status, tem_negocial, "NEGOC" in texto_status, "ACORDO" in texto_status, "RESSALVA" in texto_status]):
        partes.append("Nos ciclos classificados como tempestivos, os efeitos financeiros devem observar a data juridicamente aplicável ao pedido e a cláusula de reajuste.")
    return " ".join(partes)


def texto_contexto_contrato(res):
    contexto = (res or {}).get("contexto_contratual_anterior", {}) or {}
    if not contexto or not contexto.get("contexto_informado"):
        return "Não houve contexto contratual anterior informado para esta análise."
    linhas = []
    if contexto.get("valor_formalizado_anterior"):
        linhas.append(f"Valor formalizado antes desta análise: {moeda(contexto.get('valor_formalizado_anterior', 0))}.")
    if contexto.get("ultimo_ciclo_concedido"):
        linhas.append(f"Último ciclo já concedido/formalizado: {contexto.get('ultimo_ciclo_concedido')}.")
    if texto_seguro(contexto.get("observacao_historico", ""), ""):
        linhas.append(f"Observação: {texto_seguro(contexto.get('observacao_historico'), 'Não')}.")
    eventos = contexto.get("eventos_historicos_anteriores", []) or []
    if eventos:
        linhas.append(f"Eventos históricos anteriores registrados: {len(eventos)}.")
    linhas.append("O contexto informado serve para memória e governança. Seus efeitos só impactam os valores quando estiverem refletidos na execução, nos itens ou no saldo remanescente do Arquivo de Coleta.")
    return "\n".join(linhas)


def texto_contexto_analise(adm, res):
    indice = res.get("indice", (adm or {}).get("indice", "Não informado"))
    fator = res.get("fator_acumulado", (adm or {}).get("fator_acumulado", (adm or {}).get("fator", 1.0)))
    df_ciclos = _ciclos_df_para_relatorio(adm, res)
    ciclos_validos = df_ciclos[df_ciclos["Ciclo"].astype(str).str.upper() != "C0"] if not df_ciclos.empty else pd.DataFrame()
    qtd_ciclos = len(ciclos_validos)
    classificacoes = []
    if not ciclos_validos.empty:
        for _, row in ciclos_validos.iterrows():
            ciclo = row.get("Ciclo", "")
            classificacao = row.get("Classificação", "")
            pct_apl = row.get("Percentual aplicado", "")
            trecho = f"{ciclo}: {classificacao}"
            if pct_apl:
                trecho += f", percentual aplicado {pct_apl}"
            classificacoes.append(trecho)
    resumo_ciclos = "; ".join(classificacoes) if classificacoes else "sem ciclos de reajuste informados"

    contexto_txt = texto_contexto_contrato(res)
    if eh_modo_reduzido_itens(res):
        modo_txt = (
            "A análise foi processada em Modo Reduzido por Itens/Estoque, sem base mensal por competência. "
            "Nesse cenário, o retroativo financeiro definitivo não é calculado; apresenta-se apenas o retroativo estimado por itens/estoque, com natureza estimativa. "
        )
    else:
        modo_txt = "O bloco financeiro serve para apurar o valor represado/retroativo. "
    return (
        "A análise consolida a admissibilidade dos pleitos de reajuste e a quantificação financeira apurada com base no Arquivo de Coleta. "
        "O C0 é tratado como ciclo-base inicial, sem aplicação de reajuste. "
        f"Foram considerados {qtd_ciclos} ciclo(s) de reajuste: {resumo_ciclos}. "
        f"Índice utilizado: {indice}. Fator acumulado considerado: {fator_fmt(fator)}. "
        "O Valor Total Atualizado do Contrato mantém a composição execução atualizada por ciclo + saldo remanescente atualizado. "
        f"{modo_txt}"
        "Aditivos e supressões são apresentados como eventos de controle e governança, sem soma autônoma quando seus efeitos já estiverem refletidos na execução ou no saldo remanescente. "
        f"{contexto_txt}"
    )


def gerar_texto_instrucao(adm, res):
    if eh_modo_reduzido_itens(res):
        df_ri = df_retroativo_estimado_itens(res)
        tabela_ri = "Sem tabela detalhada disponível."
        if isinstance(df_ri, pd.DataFrame) and not df_ri.empty:
            tabela_ri = df_visual(
                df_ri,
                moeda_cols=["Valor executado original", "Valor executado atualizado", "Retroativo estimado por itens/estoque"],
            ).to_string(index=False)

        return f"""
RELATÓRIO EXECUTIVO — VALOR ATUALIZADO DO CONTRATO

1. Síntese dos ciclos

{_ciclos_df_para_relatorio(adm, res).to_string(index=False)}

2. Modo de apuração

A análise foi processada em Modo Reduzido por Itens/Estoque, sem base mensal por competência. Por essa razão, o retroativo financeiro definitivo não foi calculado. O valor abaixo possui natureza estimativa e foi apurado a partir das informações de itens/remanescentes disponíveis.

Retroativo financeiro definitivo: Não calculado
Retroativo estimado por itens/estoque: {moeda(valor_retroativo_estimado_itens(res))}
Execução estimada por itens/estoque: {moeda(res.get('valor_executado_atualizado', 0))}
Saldo remanescente atualizado: {moeda(res.get('remanescente_reajustado', 0))}

3. Retroativo estimado por itens/estoque

{tabela_ri}

4. Consolidação contratual

Valor original do contrato: {moeda(res.get('valor_original_contrato', 0))}
Valor Total Atualizado do Contrato: {moeda(res.get('valor_atualizado_contrato', res.get('valor_global_estoque', 0)))}
Aditivos/supressões registrados para controle: {moeda(res.get('total_aditivos_atualizados', 0))}

Data/hora de geração: {datetime.now(ZoneInfo('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M')}
""".strip()

    return f"""
RELATÓRIO EXECUTIVO — VALOR ATUALIZADO DO CONTRATO

1. Síntese dos ciclos

{_ciclos_df_para_relatorio(adm, res).to_string(index=False)}

2. Síntese financeira

Valor original do contrato: {moeda(res.get('valor_original_contrato', 0))}
Valor financeiro pago até o mês mais recente: {moeda(res.get('total_pago_faturado', 0))}
Valor teórico calculado: {moeda(res.get('total_devido_reajustado', 0))}
Valor represado a pagar: {moeda(res.get('valor_represado_a_pagar', 0))}
Valor executado total atualizado, considerado até o início do ciclo atual: {moeda(res.get('valor_executado_atualizado', 0))}
Saldo remanescente atualizado: {moeda(res.get('remanescente_reajustado', 0))}
Aditivos da análise atual, para controle: {moeda(res.get('total_aditivos_atualizados', 0))}
Valor Total Atualizado do Contrato: {moeda(res.get('valor_atualizado_contrato', res.get('valor_global_estoque', 0)))}

Data/hora de geração: {datetime.now(ZoneInfo('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M')}
""".strip()

def df_visual(df, moeda_cols=None, fator_cols=None, pct_cols=None):
    if df is None or not isinstance(df, pd.DataFrame):
        return pd.DataFrame()
    visual = df.copy()
    for col in moeda_cols or []:
        if col in visual.columns:
            visual[col] = visual[col].apply(moeda_ou_texto)
    for col in fator_cols or []:
        if col in visual.columns:
            visual[col] = visual[col].apply(fator_fmt)
    for col in pct_cols or []:
        if col in visual.columns:
            visual[col] = visual[col].apply(percentual)
    for col in visual.columns:
        if "data" in str(col).lower():
            visual[col] = visual[col].apply(formatar_data_br)

    # Quadro Executivo costuma vir como Indicador/Valor. Formatar como moeda
    # apenas os indicadores financeiros, preservando textos e percentuais.
    if "Indicador" in visual.columns and "Valor" in visual.columns:
        termos_monetarios = (
            "valor", "saldo", "aditivo", "supress", "delta", "pago", "teórico", "teorico",
            "remanescente", "formalizado", "original", "atualizado", "represado"
        )
        def _formatar_valor_quadro(row):
            indicador = str(row.get("Indicador", "")).lower()
            valor = row.get("Valor", "")
            if any(t in indicador for t in termos_monetarios):
                return moeda_ou_texto(valor)
            return valor
        visual["Valor"] = visual.apply(_formatar_valor_quadro, axis=1)

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
    styles.add(ParagraphStyle(name="CelulaCab", parent=styles["BodyText"], fontSize=6.8, leading=8.2, fontName="Helvetica-Bold", textColor=colors.white))

    story = []
    story.append(Paragraph("Análise de Reajuste Contratual", styles["Titulo"]))
    story.append(Paragraph("Relatório Executivo", styles["Titulo"]))
    story.append(Paragraph(f"Gerado em: {datetime.now(ZoneInfo('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M')}", styles["Texto"]))
    story.append(Spacer(1, 6))

    origem = (adm or {}).get("origem") or (adm or {}).get("tipo") or "Não informado"
    indice = res.get("indice", (adm or {}).get("indice", "Não informado"))
    fator = res.get("fator_acumulado", (adm or {}).get("fator_acumulado", (adm or {}).get("fator", 1.0)))

    story.append(Paragraph("1. Identificação da Análise", styles["Subtitulo"]))
    dados_identificacao = []
    if str(origem).strip().lower() not in ["", "não informado", "nao informado"]:
        dados_identificacao.append(["Origem da análise", origem])
    dados_identificacao.extend([
        ["Índice aplicado", indice],
        ["Fator acumulado", fator_fmt(fator)],
        ["Data de geração", datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M")],
    ])
    story.append(tabela_pdf(dados_identificacao, col_widths=[6 * cm, 11 * cm]))

    df_ciclos_rel = _ciclos_df_para_relatorio(adm, res)
    if isinstance(df_ciclos_rel, pd.DataFrame) and not df_ciclos_rel.empty:
        story.append(Paragraph("2. Síntese dos ciclos", styles["Subtitulo"]))
        story.append(tabela_dataframe_pdf(df_ciclos_rel, max_linhas=20))

    story.append(Paragraph("3. Indicadores Executivos", styles["Subtitulo"]))
    story.append(tabela_pdf(indicadores_executivos_relatorio(res), header=True, col_widths=[8.5 * cm, 8.5 * cm]))

    if eh_modo_reduzido_itens(res):
        story.append(Paragraph("3.1. Modo Reduzido por Itens/Estoque", styles["Subtitulo"]))
        story.append(Paragraph(
            "A análise foi processada sem base mensal por competência. O retroativo financeiro definitivo não é calculado neste modo. "
            "O valor de retroativo estimado por itens/estoque possui natureza estimativa e deve ser validado antes de qualquer formalização de pagamento.",
            styles["Texto"],
        ))
        df_ri_pdf = df_retroativo_estimado_itens(res)
        if isinstance(df_ri_pdf, pd.DataFrame) and not df_ri_pdf.empty:
            df_ri_pdf = df_visual(
                df_ri_pdf,
                moeda_cols=["Valor executado original", "Valor executado atualizado", "Retroativo estimado por itens/estoque"],
            )
            story.append(tabela_dataframe_pdf(df_ri_pdf, max_linhas=20))

    story.append(Paragraph("4. Financeiro por Ciclo", styles["Subtitulo"]))
    df_fin = df_visual(
        res.get("df_financeiro_por_ciclo"),
        moeda_cols=["Valor pago efetivo", "Valor teórico calculado", "Valor pago/faturado", "Valor devido reajustado", "Delta do ciclo", "Delta acumulado"],
        fator_cols=["Fator aplicado ao retroativo", "Fator aplicado"],
    )
    keep_fin = [c for c in ["Ciclo", "Situação", "Tratamento financeiro", "Fator aplicado ao retroativo", "Fator aplicado", "Valor pago efetivo", "Valor teórico calculado", "Valor pago/faturado", "Valor devido reajustado", "Delta do ciclo"] if c in df_fin.columns]
    story.append(tabela_dataframe_pdf(df_fin[keep_fin] if keep_fin else df_fin, max_linhas=20))

    story.append(Paragraph("5. Composição do Valor Total Atualizado do Contrato", styles["Subtitulo"]))
    df_comp = res.get("df_composicao_valor_total")
    if isinstance(df_comp, pd.DataFrame) and not df_comp.empty:
        df_comp_pdf = df_comp.copy()
        if "Valor" in df_comp_pdf.columns:
            df_comp_pdf["Valor"] = df_comp_pdf["Valor"].apply(moeda)
        keep_comp = [c for c in ["Componente", "Ciclo/Referência", "Valor", "Observação"] if c in df_comp_pdf.columns]
        story.append(tabela_dataframe_pdf(df_comp_pdf[keep_comp], max_linhas=20))
    else:
        story.append(tabela_pdf([
            ["Componente", "Valor"],
            ["Valor executado atualizado", moeda(res.get("valor_executado_atualizado", 0))],
            ["Saldo remanescente atualizado", moeda(res.get("remanescente_reajustado", 0))],
            ["Valor Total Atualizado do Contrato", moeda(res.get("valor_atualizado_contrato", res.get("valor_global_estoque", 0)))],
        ], header=True, col_widths=[10 * cm, 7 * cm]))
        story.append(Paragraph(
            "Aditivos e supressões registrados são apresentados em seção própria para controle e não são somados como parcela autônoma ao Valor Total Atualizado quando já refletidos na execução ou no saldo remanescente.",
            styles["Texto"],
        ))

    df_ad = res.get("df_aditivos_executivo", res.get("df_aditivos"))
    if isinstance(df_ad, pd.DataFrame) and not df_ad.empty:
        story.append(Paragraph("6. Aditivos", styles["Subtitulo"]))
        df_adv = df_visual(
            df_ad,
            moeda_cols=["Valor do aditivo na assinatura", "Valor do aditivo reajustado", "Valor original da alteração", "Valor atualizado da alteração"],
            fator_cols=["Fator aplicado"],
        )
        keep_ad = [c for c in ["Aditivo", "Ciclo/Marco", "Tratamento do aditivo", "Quantidade de linhas", "Valor do aditivo na assinatura", "Fator aplicado", "Valor do aditivo reajustado"] if c in df_adv.columns]
        story.append(tabela_dataframe_pdf(df_adv[keep_ad], max_linhas=12))

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
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
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

    from reportlab.lib import colors

    styles = getSampleStyleSheet()
    cel = ParagraphStyle("Cel", parent=styles["BodyText"], fontSize=6.6, leading=8.0)
    cab = ParagraphStyle(
        "Cab",
        parent=styles["BodyText"],
        fontSize=6.6,
        leading=8.0,
        fontName="Helvetica-Bold",
        textColor=colors.white,
    )
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




# ============================================================
# Layout responsivo e Minuta de Apostilamento
# ============================================================

def aplicar_css_responsivo_relatorio():
    st.markdown(
        """
        <style>
        div[data-testid="stMetric"] {
            background: #FFFFFF;
            border: 1px solid #E5EAF0;
            border-radius: 12px;
            padding: 10px 12px;
            min-height: 82px;
        }
        div[data-testid="stMetricLabel"] p {
            color: #475569;
            font-size: clamp(0.74rem, 1.15vw, 0.90rem);
            line-height: 1.2;
            white-space: normal;
            word-break: normal;
        }
        div[data-testid="stMetricValue"] {
            font-size: clamp(1.02rem, 1.75vw, 1.50rem);
            line-height: 1.20;
            white-space: normal;
            overflow-wrap: anywhere;
        }
        div[data-testid="stMetricDelta"] {
            font-size: 0.78rem;
        }
        .telebras-kpi-destaque {
            background:#EAF2F8;
            border:1px solid #BFD7EA;
            border-radius:14px;
            padding:16px 20px;
            margin:10px 0 16px 0;
        }
        .telebras-kpi-destaque-label {
            font-size:0.92rem;
            color:#27496D;
            font-weight:600;
        }
        .telebras-kpi-destaque-valor {
            font-size:clamp(1.25rem, 2.45vw, 1.95rem);
            color:#0B1F3A;
            font-weight:800;
            line-height:1.25;
            word-break:break-word;
        }
        @media (max-width: 1200px) {
            div[data-testid="stMetricValue"] {
                font-size: 1.05rem;
            }
            div[data-testid="stMetric"] {
                padding: 8px 10px;
                min-height: 76px;
            }
        }

        .modo-reduzido-box {
            background: #F3E8FF;
            border: 1px solid #A855F7;
            border-left: 8px solid #7E22CE;
            border-radius: 14px;
            padding: 18px 22px;
            margin: 12px 0 18px 0;
        }
        .modo-reduzido-titulo {
            color: #581C87;
            font-weight: 900;
            font-size: 1.02rem;
            margin-bottom: 7px;
        }
        .modo-reduzido-texto {
            color: #3B0764;
            font-size: 0.95rem;
            line-height: 1.45;
        }
        .modo-reduzido-card {
            background: #F3E8FF;
            border: 1px solid #A855F7;
            border-radius: 14px;
            padding: 16px 18px;
            margin: 8px 0 16px 0;
        }
        .modo-reduzido-card-label {
            color: #581C87;
            font-weight: 800;
            font-size: 0.92rem;
            margin-bottom: 6px;
        }
        .modo-reduzido-card-valor {
            color: #0F172A;
            font-size: clamp(1.20rem, 2.10vw, 1.65rem);
            font-weight: 900;
            line-height: 1.2;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _valor_resumo(res, *chaves, padrao=0.0):
    for chave in chaves:
        if isinstance(res, dict) and chave in res:
            return res.get(chave)
    return padrao


def _placeholder(valor=None):
    if valor is None:
        return "[campo a preencher]"
    texto = str(valor).strip()
    return texto if texto else "[campo a preencher]"


def _limpar_texto_formal(valor):
    """Remove emojis e marcadores visuais informais de textos destinados à minuta formal."""
    if valor is None:
        return ""
    texto = str(valor)
    substituicoes = {
        "❌": "",
        "✅": "",
        "⚠️": "",
        "⚠": "",
        "🟡": "",
        "🔴": "",
        "🟢": "",
        "🔵": "",
        "🛡️": "",
        "🛡": "",
        "📊": "",
        "📝": "",
        "⚖️": "",
        "⚖": "",
        "🔄": "",
        "📥": "",
        "🔍": "",
    }
    for antigo, novo in substituicoes.items():
        texto = texto.replace(antigo, novo)
    # Remove caracteres fora do plano multilíngue básico, onde ficam a maior parte dos emojis.
    texto = "".join(ch for ch in texto if ord(ch) <= 0xFFFF)
    return " ".join(texto.split())


def _adicionar_item_numerado(document, numero, texto):
    p = document.add_paragraph()
    p.add_run(f"{numero}. ").bold = True
    p.add_run(_limpar_texto_formal(texto))
    return p


def _adicionar_subitem(document, marcador, texto):
    from docx.shared import Cm
    p = document.add_paragraph(f"{marcador} {_limpar_texto_formal(texto)}")
    p.paragraph_format.left_indent = Cm(0.7)
    return p


def _romano(numero):
    mapa = [
        (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i"),
    ]
    n = int(numero)
    saida = ""
    for valor, simbolo in mapa:
        while n >= valor:
            saida += simbolo
            n -= valor
    return saida


def _ciclos_para_minuta(adm, res):
    ciclos = []
    if adm:
        ciclos = adm.get("ciclos") or adm.get("detalhamento_ciclos") or []
    if not ciclos and isinstance(res.get("df_ciclos"), pd.DataFrame):
        ciclos = res.get("df_ciclos").to_dict("records")
    return ciclos or []


def _percentual_ciclo_minuta(ciclo):
    for chave in ["percentual_aplicado", "Percentual aplicado", "Variação", "variacao", "var"]:
        if chave in ciclo and ciclo.get(chave) not in [None, ""]:
            try:
                return percentual(float(ciclo.get(chave)), 2)
            except Exception:
                return str(ciclo.get(chave))
    return "[campo a preencher]"


def _efeito_ciclo_minuta(ciclo):
    for chave in [
        "data_inicio_efeito_financeiro",
        "inicio_financeiro_acordo",
        "financeiro_inicio",
        "Início financeiro",
        "Data de início dos efeitos financeiros",
        "data_pedido",
        "Data do pedido",
    ]:
        valor = ciclo.get(chave) if isinstance(ciclo, dict) else None
        if valor not in [None, ""]:
            return formatar_data_br(valor)
    return "[campo a preencher]"


def _nome_ciclo_minuta(ciclo, idx):
    return str(ciclo.get("ciclo") or ciclo.get("Ciclo") or f"C{idx}")


def _adicionar_paragrafo_justificado(document, texto):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    p = document.add_paragraph(texto)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p


def _aplicar_estilo_docx(document):
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    section = document.sections[0]
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)

    styles = document.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10)
    for nome in ["Title", "Heading 1", "Heading 2"]:
        if nome in styles:
            styles[nome].font.name = "Arial"
    styles["Title"].font.size = Pt(13)
    styles["Title"].font.bold = True


def _destacar_campos_preencher(document):
    """Destaca em amarelo os trechos [campo a preencher...] na minuta DOCX."""
    try:
        from copy import deepcopy
        from docx.enum.text import WD_COLOR_INDEX
        from docx.oxml import OxmlElement
        from docx.text.run import Run
    except Exception:
        return document

    padrao = re.compile(r"(\[campo a preencher[^\]]*\])", flags=re.IGNORECASE)
    for paragraph in document.paragraphs:
        runs_originais = list(paragraph.runs)
        for run in runs_originais:
            texto = run.text
            if not texto or not padrao.search(texto):
                continue
            partes = padrao.split(texto)
            run.text = partes[0]
            anterior = run._r
            for parte in partes[1:]:
                novo_r = OxmlElement('w:r')
                if run._r.rPr is not None:
                    novo_r.append(deepcopy(run._r.rPr))
                novo_t = OxmlElement('w:t')
                novo_t.text = parte
                novo_r.append(novo_t)
                anterior.addnext(novo_r)
                novo_run = Run(novo_r, paragraph)
                if padrao.fullmatch(parte):
                    novo_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                anterior = novo_r
    return document


def gerar_minuta_apostilamento_docx(adm, res):
    try:
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Pt
    except Exception as exc:
        raise RuntimeError("A biblioteca python-docx não está instalada. Inclua 'python-docx' no requirements.txt.") from exc

    document = Document()
    _aplicar_estilo_docx(document)

    titulo = document.add_paragraph()
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = titulo.add_run("MINUTA DE TERMO DE APOSTILAMENTO")
    run.bold = True
    run.font.size = Pt(13)

    contrato = _placeholder((adm or {}).get("contrato") or (adm or {}).get("numero_contrato"))
    p = document.add_paragraph(f"Contrato nº {contrato}")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    _adicionar_paragrafo_justificado(
        document,
        "A TELECOMUNICAÇÕES BRASILEIRAS S.A. - TELEBRAS, sociedade de economia mista, vinculada ao Ministério das Comunicações, com sede no SIG, Quadra 04, Bloco A, Salas 201 a 224, Edifício Capital Financial Center, CEP nº 70.610-440, inscrita no CNPJ sob o n.º 00.336.701/0001-04, com seus atos constitutivos devidamente arquivados na Junta Comercial do Distrito Federal, sob o nº 7.665, em 20/02/1978, publicada no Diário Oficial da União de 13/03/1978, doravante denominada TELEBRAS, neste ato representada por [campo a preencher], Matrícula [campo a preencher], e por seu [campo a preencher], Matrícula [campo a preencher], nos termos da Diretriz nº 229/2018, apostilam o Contrato nº [campo a preencher], celebrado com a empresa [campo a preencher], doravante denominada CONTRATADA, nos termos do parágrafo 7º, art. 81, da Lei nº 13.303, de 30 de junho de 2016, legislação complementar, e:"
    )

    document.add_paragraph("CONSIDERANDO:")
    considerandos = [
        "A Cláusula Oitava do Contrato nº [campo a preencher], que prevê o reajuste contratual e disciplina o marco para concessão, os efeitos financeiros, os reajustes subsequentes e a regra de preclusão;",
        "A deliberação da Diretoria Executiva da Telebras, consignada na Ata da 1869ª Reunião Ordinária, de 13 de janeiro de 2026, que revogou a suspensão anteriormente imposta à tramitação dos reajustes contratuais e restabeleceu a normalidade do respectivo processamento;",
        "O Despacho Saneador [campo a preencher], por meio do qual a Gerência de Compras e Contratos consolidou o histórico do pleito, a admissibilidade, a apuração dos índices, os efeitos financeiros e os documentos de suporte necessários à formalização do presente apostilamento;",
        f"A memória de cálculo constante em [campo a preencher], que apurou o reajuste acumulado de {percentual(res.get('variacao_acumulada', res.get('fator_acumulado', 1.0) - 1), 2)};",
        "A memória financeira disponível em [campo a preencher], que detalha os valores nominais pagos, os valores devidos e o saldo retroativo a pagar por ciclo, bem como os saldos remanescentes considerados no início de cada ciclo;",
        "As informações prestadas pela área gestora do contrato em [campo a preencher], quanto aos itens/unidades remanescentes, e em [campo a preencher], quanto aos valores financeiros utilizados na apuração;",
        "A manifestação da CONTRATADA constante de [campo a preencher], pela qual anuiu com os cálculos apresentados pela TELEBRAS;",
        "As certidões de regularidade da CONTRATADA juntadas em [campo a preencher] e a adequação orçamentária registrada em [campo a preencher].",
    ]
    df_ad_docx = res.get("df_aditivos_executivo", res.get("df_aditivos", pd.DataFrame())) if isinstance(res, dict) else pd.DataFrame()
    if isinstance(df_ad_docx, pd.DataFrame) and not df_ad_docx.empty:
        for _, ad in df_ad_docx.iterrows():
            identificacao = texto_seguro(ad.get("Aditivo", ad.get("Identificação", "Termo Aditivo")), "Termo Aditivo")
            data_ad = formatar_data_br(ad.get("Data do aditivo", ""))
            valor_ad = moeda(ad.get("Valor do aditivo na assinatura", ad.get("Valor original da alteração", 0)))
            considerandos.append(
                f"O {identificacao}, de {data_ad if data_ad else '[campo a preencher]'}, com valor de {valor_ad} na data de assinatura;"
            )

    for item in considerandos:
        document.add_paragraph(_limpar_texto_formal(item), style="List Bullet")

    ciclos = _ciclos_para_minuta(adm or {}, res or {})
    ciclos_validos = [c for c in ciclos if str(c.get("ciclo") or c.get("Ciclo") or "").strip()]

    _adicionar_item_numerado(
        document,
        1,
        "Procede-se ao apostilamento do Contrato nº [campo a preencher] para formalizar a concessão dos reajustes contratuais apurados, nos seguintes termos:"
    )
    if ciclos_validos:
        for idx, ciclo in enumerate(ciclos_validos, start=1):
            nome = _limpar_texto_formal(_nome_ciclo_minuta(ciclo, idx))
            pct = _limpar_texto_formal(_percentual_ciclo_minuta(ciclo))
            efeito = _limpar_texto_formal(_efeito_ciclo_minuta(ciclo))
            situacao = _limpar_texto_formal(
                ciclo.get("situacao_aplicada")
                or ciclo.get("Situação aplicada")
                or ciclo.get("situacao")
                or ciclo.get("Situação")
                or ""
            )
            complemento = f", com tratamento aplicado: {situacao}" if situacao else ""
            _adicionar_subitem(
                document,
                f"({_romano(idx)})",
                f"{nome}, no percentual de {pct}, com efeitos financeiros a partir de {efeito}{complemento};"
            )
        _adicionar_subitem(
            document,
            f"({_romano(len(ciclos_validos) + 1)})",
            f"O percentual acumulado apurado corresponde a {percentual(res.get('variacao_acumulada', res.get('fator_acumulado', 1.0) - 1), 2)}."
        )
    else:
        _adicionar_subitem(document, "(i)", "[campo a preencher: informar ciclos, percentuais e efeitos financeiros].")

    _adicionar_item_numerado(document, 2, "Para fins de apuração financeira e cálculo do retroativo a pagar, foram consolidados os seguintes valores:")
    df_fin = res.get("df_financeiro_por_ciclo")
    subitem_idx = 1
    if isinstance(df_fin, pd.DataFrame) and not df_fin.empty:
        for _, row in df_fin.iterrows():
            ciclo = str(row.get("Ciclo", "")).strip()
            if ciclo.upper() == "TOTAL" or not ciclo:
                continue
            pago = row.get("Valor pago efetivo", row.get("Valor pago/faturado", row.get("Valor nominal pago", 0)))
            devido = row.get("Valor teórico calculado", row.get("Valor devido reajustado", row.get("Valor devido", 0)))
            delta = row.get("Delta do ciclo", row.get("Delta acumulado", 0))
            _adicionar_subitem(
                document,
                f"({_romano(subitem_idx)})",
                f"{ciclo}: valor nominal pago de {moeda(pago)}, valor devido de {moeda(devido)} e saldo retroativo a pagar de {moeda(delta)};"
            )
            subitem_idx += 1
    else:
        _adicionar_subitem(document, f"({_romano(subitem_idx)})", "[campo a preencher: valores nominais pagos, valores devidos e saldo retroativo por ciclo].")
        subitem_idx += 1
    _adicionar_subitem(
        document,
        f"({_romano(subitem_idx)})",
        f"Total consolidado: valor pago efetivo de {moeda(res.get('total_pago_faturado', res.get('valor_pago_efetivo', 0)))}, valor teórico calculado de {moeda(res.get('total_devido_reajustado', res.get('valor_teorico_calculado', 0)))} e saldo retroativo total a pagar de {moeda(res.get('valor_represado_a_pagar', res.get('delta_acumulado', 0)))}."
    )

    _adicionar_item_numerado(document, 3, "Registra-se, ainda, que a memória de cálculo consignou os saldos remanescentes do contrato nos marcos de início dos ciclos de reajuste:")
    df_rem = res.get("df_remanescentes")
    if isinstance(df_rem, pd.DataFrame) and not df_rem.empty:
        subitem_idx = 1
        for _, row in df_rem.iterrows():
            ciclo = str(row.get("Ciclo", "")).strip() or "[ciclo]"
            rem = row.get("Remanescente atualizado", row.get("Remanescente original", 0))
            _adicionar_subitem(document, f"({_romano(subitem_idx)})", f"{ciclo}: saldo remanescente de {moeda(rem)} no marco de referência do ciclo;")
            subitem_idx += 1
    else:
        _adicionar_subitem(document, "(i)", "[campo a preencher: saldos remanescentes por ciclo].")

    _adicionar_item_numerado(document, 4, "Para fins de consolidação do Valor Total Atualizado do Contrato, registra-se que o total decorre da execução atualizada por ciclo somada ao saldo remanescente atualizado do ciclo mais recente, sem soma autônoma de aditivos ou supressões já incorporados à execução ou ao saldo:")
    _adicionar_subitem(
        document,
        "a)",
        f"Valor total atualizado do contrato: de {moeda(res.get('valor_original_contrato', 0))} para {moeda(res.get('valor_atualizado_contrato', res.get('valor_global_estoque', 0)))}."
    )
    _adicionar_subitem(
        document,
        "b)",
        f"Valor consolidado da execução atualizada por ciclos: {moeda(res.get('valor_executado_atualizado', res.get('total_devido_reajustado', 0)))}."
    )
    _adicionar_subitem(
        document,
        "c)",
        f"Saldo remanescente atualizado considerado na consolidação: {moeda(res.get('remanescente_reajustado', 0))}."
    )
    _adicionar_subitem(
        document,
        "d)",
        f"Valor total de aditivos/supressões registrados para controle: {moeda(res.get('total_aditivos_atualizados', 0))}. Esses valores não são somados como parcela autônoma ao Valor Total Atualizado do Contrato quando já estiverem refletidos na execução ou no saldo remanescente."
    )

    _adicionar_item_numerado(document, 5, "Permanecem inalteradas e em pleno vigor as demais cláusulas e condições do Contrato e de seus instrumentos posteriores não modificadas por este Termo de Apostila.")
    _adicionar_item_numerado(document, 6, "A CONTRATADA deverá atualizar a garantia contratual, prevista na Cláusula Décima do Contrato, no prazo contratualmente estabelecido, observado o novo valor após a formalização deste Termo de Apostila.")
    _adicionar_item_numerado(document, 7, "O presente apostilamento vincula-se, para todos os fins, aos documentos [campo a preencher] instruídos no Processo [campo a preencher].")

    document.add_paragraph("Brasília/DF, [Data].")
    document.add_paragraph("\nTELECOMUNICAÇÕES BRASILEIRAS S.A. - TELEBRAS")
    document.add_paragraph("Representante Legal 1")
    document.add_paragraph("\nTELECOMUNICAÇÕES BRASILEIRAS S.A. - TELEBRAS")
    document.add_paragraph("Representante Legal 2")

    _destacar_campos_preencher(document)

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
render_marca_topo()
st.title("Relatório Global")

adm = st.session_state.get("dados_admissibilidade")
res = st.session_state.get("resultado_valor_global")

if not res:
    st.warning(
        "Ainda não há dados processados para o Relatório Global. "
        "Acesse o módulo Valor Global, carregue o Arquivo de Coleta e processe a análise."
    )
    st.stop()

aplicar_css_responsivo_relatorio()

modo_reduzido = eh_modo_reduzido_itens(res)
retroativo_itens_valor = valor_retroativo_estimado_itens(res)

st.subheader("Resumo Executivo")
if modo_reduzido:
    st.markdown(aviso_modo_reduzido_html(), unsafe_allow_html=True)

col1, col2 = st.columns(2)
col1.metric("Índice", res.get("indice", "Não informado"))
col2.metric("Fator acumulado", fator_fmt(res.get("fator_acumulado", 1.0)))

st.markdown(
    f"""
    <div class="telebras-kpi-destaque">
        <div class="telebras-kpi-destaque-label">Valor Total Atualizado do Contrato</div>
        <div class="telebras-kpi-destaque-valor">{moeda(res.get("valor_atualizado_contrato", res.get("valor_global_estoque", 0)))}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

if modo_reduzido:
    col3, col4 = st.columns(2)
    col3.metric("Retroativo financeiro definitivo", "Não calculado")
    with col4:
        st.markdown(
            f"""
            <div class="modo-reduzido-card">
                <div class="modo-reduzido-card-label">Retroativo estimado por itens/estoque</div>
                <div class="modo-reduzido-card-valor">{moeda(retroativo_itens_valor)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    col5, col6 = st.columns(2)
    col5.metric("Base mensal por competência", "Não informada")
    col6.metric("Execução estimada por itens/estoque", moeda(res.get("valor_executado_atualizado", 0)))

    col7, col8 = st.columns(2)
    col7.metric("Saldo remanescente atualizado", moeda(res.get("remanescente_reajustado", 0)))
    col8.metric("Aditivos/supressões registrados", moeda(res.get("total_aditivos_atualizados", 0)))
else:
    col3, col4 = st.columns(2)
    col3.metric("Valor represado a pagar", moeda(res.get("valor_represado_a_pagar", 0)))
    col4.metric("Aditivos/supressões registrados", moeda(res.get("total_aditivos_atualizados", 0)))

    col5, col6 = st.columns(2)
    col5.metric("Valor pago efetivo", moeda(res.get("total_pago_faturado", 0)))
    col6.metric("Valor teórico calculado", moeda(res.get("total_devido_reajustado", 0)))

st.divider()

tab1, tab2, tab3, tab4 = st.tabs(["Relatório Executivo", "Tabelas", "PDF", "Minuta de Apostilamento"])

with tab1:
    texto = gerar_texto_instrucao(adm, res)
    st.text_area("Copie para o processo:", texto, height=420)

with tab2:
    st.markdown("### Quadro Executivo")
    st.dataframe(
        df_visual(res.get("df_comparativo"), moeda_cols=["Valor", "Antes do Reajuste", "Após Reajuste", "Diferença"]),
        use_container_width=True,
        hide_index=True,
    )

    if modo_reduzido:
        st.markdown("### Retroativo estimado por itens/estoque")
        st.markdown(aviso_modo_reduzido_html(), unsafe_allow_html=True)
        st.metric("Retroativo estimado por itens/estoque", moeda(retroativo_itens_valor))
        df_ri = df_retroativo_estimado_itens(res)
        if isinstance(df_ri, pd.DataFrame) and not df_ri.empty:
            st.dataframe(
                df_visual(
                    df_ri,
                    moeda_cols=["Valor executado original", "Valor executado atualizado", "Retroativo estimado por itens/estoque"],
                ),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("Tabela detalhada do retroativo estimado por itens/estoque não disponível nesta sessão.")

    st.markdown("### Composição do Valor Total Atualizado do Contrato")
    st.caption("Composição considerada: execução atualizada por ciclo + saldo remanescente atualizado. Aditivos/supressões são demonstrados separadamente para controle, sem soma autônoma ao total.")
    df_comp_valor = res.get("df_composicao_valor_total")
    if isinstance(df_comp_valor, pd.DataFrame) and not df_comp_valor.empty:
        st.dataframe(
            df_visual(df_comp_valor, moeda_cols=["Valor"]),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("Composição do valor total não disponível nesta sessão.")

    st.markdown("### Financeiro por Ciclo")
    st.dataframe(
        df_visual(
            res.get("df_financeiro_por_ciclo"),
            moeda_cols=["Valor pago efetivo", "Valor teórico calculado", "Valor pago/faturado", "Valor devido reajustado", "Delta do ciclo", "Delta acumulado"],
            fator_cols=["Fator aplicado ao retroativo", "Fator aplicado"],
        ),
        use_container_width=True,
        hide_index=True,
    )

    df_ad = res.get("df_aditivos_executivo", res.get("df_aditivos"))
    if isinstance(df_ad, pd.DataFrame) and not df_ad.empty:
        st.markdown("### Aditivos e Supressões")
        st.caption("Quadro de controle formal. Esses valores não são somados autonomamente ao Valor Total Atualizado quando já estiverem incorporados à execução ou ao saldo remanescente.")
        keep_ad = [c for c in ["Aditivo", "Ciclo/Marco", "Tratamento do aditivo", "Quantidade de linhas", "Valor do aditivo na assinatura", "Fator aplicado", "Valor do aditivo reajustado"] if c in df_ad.columns]
        st.dataframe(
            df_visual(df_ad[keep_ad].copy() if keep_ad else df_ad, moeda_cols=["Valor do aditivo na assinatura", "Valor do aditivo reajustado", "Valor original da alteração", "Valor atualizado da alteração"], fator_cols=["Fator aplicado"]),
            use_container_width=True,
            hide_index=True,
        )

with tab3:
    st.markdown("### Baixar Relatório Executivo em PDF")
    try:
        pdf_bytes = criar_pdf_relatorio(adm, res)
        st.session_state["arquivo_relatorio_executivo_pdf"] = pdf_bytes
        st.download_button(
            label="Baixar Relatório Executivo em PDF",
            data=pdf_bytes,
            file_name="Relatorio_Executivo_Analise_Reajuste.pdf",
            mime="application/pdf",
            type="primary",
        )
    except Exception as exc:
        st.error(f"Não foi possível gerar o PDF: {exc}")
        st.caption("Verifique se a biblioteca reportlab foi incluída no requirements.txt.")


with tab4:
    st.markdown("### Gerar Minuta de Termo de Apostilamento")
    if modo_reduzido:
        st.markdown(
            """
            <div class="modo-reduzido-box">
                <div class="modo-reduzido-titulo">Atenção: minuta em Modo Reduzido</div>
                <div class="modo-reduzido-texto">
                    A análise foi processada sem base mensal por competência. Recomenda-se validar a base financeira antes de usar a minuta para formalização de pagamento ou apostilamento.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    st.info(
        "A minuta é gerada em DOCX editável. Os dados disponíveis no sistema são preenchidos automaticamente; "
        "as informações ainda não cadastradas permanecem como [campo a preencher]."
    )
    try:
        docx_bytes = gerar_minuta_apostilamento_docx(adm, res)
        st.session_state["arquivo_minuta_apostilamento_docx"] = docx_bytes
        st.download_button(
            label="Baixar Minuta de Apostilamento em DOCX",
            data=docx_bytes,
            file_name="minuta_termo_apostilamento.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=False,
        )
    except Exception as exc:
        st.error(f"Não foi possível gerar a minuta de apostilamento: {exc}")
        st.caption("Verifique se a biblioteca python-docx foi incluída no requirements.txt.")
