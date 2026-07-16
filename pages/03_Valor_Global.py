import re
import html
import unicodedata
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from _coleta_reajuste import (
    CAMINHO_MODELO_COLETA,
    NOME_ARQUIVO_COLETA,
    eh_coleta_reajuste,
    ler_coleta_reajuste,
)
from _coleta_reajuste_documentos import adaptar_coleta_reajuste_para_documentos
from _capacidades_apuracao import avaliar_capacidades_apuracao
from _estado_upload import (
    DETALHES_ARQUIVO_NAO_RECONHECIDO,
    MENSAGEM_ARQUIVO_NAO_RECONHECIDO,
    ORIGEM_COLETA_OFICIAL,
    ORIGEM_NAO_RECONHECIDA,
    limpar_estados_derivados,
    procedencia_registrada,
    registrar_upload,
    sha256_do_arquivo,
    upload_ja_processado,
)
from _ui_capacidades import (
    render_resultados_progressivos,
    render_status_apuracao,
    render_status_documentos,
)

aditivos_somados_ao_valor_total = 0.0  # fallback: planilha sem aditivos computaveis
LEITOR_CONSUMO_ITENS_CICLO_VERSAO = "20260516_0207"

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    REPORTLAB_OK = True
except Exception:
    REPORTLAB_OK = False


st.set_page_config(page_title="Análises de Reajustes - Valor Global", layout="wide")




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
from _ui_utils import render_cabecalho_pagina

def agora_brasilia():
    return datetime.now(ZoneInfo("America/Sao_Paulo"))


def normalizar_texto(valor):
    """Normaliza textos para comparação robusta de nomes de colunas/abas."""
    if valor is None:
        return ""
    texto = str(valor).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = re.sub(r"[^a-z0-9]+", "_", texto)
    return texto.strip("_")


def moeda(valor):
    try:
        valor = round(float(valor), 2)
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


def numero_br(valor):
    """Converte valores em número, aceitando R$, ponto de milhar, vírgula decimal e percentuais."""
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip()
    if texto == "":
        return 0.0
    texto = texto.replace("R$", "").replace("\xa0", "").strip()
    texto = texto.replace("%", "").strip()

    # Se houver vírgula, assume padrão brasileiro: 1.234,56
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    else:
        # Se houver múltiplos pontos, remove separadores de milhar
        if texto.count(".") > 1:
            partes = texto.split(".")
            texto = "".join(partes[:-1]) + "." + partes[-1]

    try:
        return float(texto)
    except Exception:
        return 0.0


def percentual_para_decimal(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, str) and "%" in valor:
        return numero_br(valor) / 100
    n = numero_br(valor)
    if abs(n) > 1:
        return n / 100
    return n


def fator_de_valor(valor, variacao=None):
    """Retorna fator. Aceita fator direto ou variação."""
    if valor is not None and not pd.isna(valor):
        texto = str(valor)
        n = numero_br(valor)
        if "%" in texto:
            return 1 + (n / 100)
        if n > 0:
            # Se vier 4,68, trata como variação percentual; se vier 1,0468, trata como fator.
            if n >= 2:
                return 1 + (n / 100)
            return n
    if variacao is not None:
        return 1 + percentual_para_decimal(variacao)
    return 1.0



def fator_operacional(valor):
    """Fator financeiro operacional usado no Valor Global.

    Regra de estabilidade: preservar o padrão histórico da ferramenta,
    usando o fator de cada ciclo com 4 casas decimais. Isso evita divergências
    entre o XLS de coleta, a planilha executiva e o cálculo do site causadas
    por casas decimais residuais dos índices.
    """
    try:
        return round(float(valor), 4)
    except Exception:
        return 1.0

def localizar_coluna(df, opcoes):
    """Localiza uma coluna com base em alternativas normalizadas."""
    colunas_norm = {normalizar_texto(c): c for c in df.columns}
    opcoes_norm = [normalizar_texto(o) for o in opcoes]

    for alvo in opcoes_norm:
        if alvo in colunas_norm:
            return colunas_norm[alvo]

    for col_norm, col_original in colunas_norm.items():
        for alvo in opcoes_norm:
            if alvo and alvo in col_norm:
                return col_original
    return None


def normalizar_ciclo(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().upper()
    if texto == "":
        return ""
    m = re.search(r"C\s*([0-9]+)", texto)
    if m:
        return f"C{int(m.group(1))}"
    m = re.search(r"([0-9]+)", texto)
    if m:
        return f"C{int(m.group(1))}"
    return texto


def numero_ciclo(ciclo):
    m = re.search(r"([0-9]+)", str(ciclo))
    return int(m.group(1)) if m else 999


def numero_seguro(valor, padrao=0.0):
    try:
        n = float(valor)
        if pd.isna(n) or n in [float("inf"), float("-inf")]:
            return padrao
        return n
    except Exception:
        return padrao


def limpar_nan_inf_df(df):
    if df is None or not isinstance(df, pd.DataFrame):
        return pd.DataFrame()
    return df.replace([float("inf"), float("-inf")], pd.NA).fillna("")


def texto_seguro(valor, padrao="Não"):
    """Converte valores vazios/nan/None em texto seguro para interface e relatórios."""
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


def contem_preclusao_ou_adiantamento(situacao):
    s = normalizar_texto(situacao)
    return "preclus" in s or ("adiantado" in s and "ressalva" not in s)


# ============================================================
# Leitura de planilha
# ============================================================

def localizar_aba(xls, opcoes):
    mapa = {normalizar_texto(s): s for s in xls.sheet_names}
    for opcao in opcoes:
        n = normalizar_texto(opcao)
        if n in mapa:
            return mapa[n]
    for sheet_norm, sheet_original in mapa.items():
        for opcao in opcoes:
            if normalizar_texto(opcao) in sheet_norm:
                return sheet_original
    return None


def ler_aba_com_cabecalho(bytes_arquivo, aba, termos_obrigatorios=None):
    """Lê aba encontrando automaticamente a linha de cabeçalho.

    A detecção é tolerante a linhas superiores mescladas, espaços, acentos e caixa.
    Quando termos obrigatórios são informados, escolhe a linha que contém todos eles
    como cabeçalhos reais, evitando confundir textos como "Dados do item..." com a
    coluna "Item".
    """
    bruto = pd.read_excel(BytesIO(bytes_arquivo), sheet_name=aba, header=None)
    termos_obrigatorios = [normalizar_texto(t) for t in (termos_obrigatorios or [])]

    linha_cabecalho = None
    melhor_linha = None
    melhor_pontuacao = -1

    for idx, row in bruto.iterrows():
        valores = [normalizar_texto(v) for v in row.tolist()]
        valores_validos = [v for v in valores if v]
        if not valores_validos:
            continue

        if not termos_obrigatorios:
            linha_cabecalho = idx
            break

        # Pontua por termos encontrados como célula exata ou contidos em uma célula.
        pontuacao = 0
        for termo in termos_obrigatorios:
            if termo in valores_validos:
                pontuacao += 2
            elif any(termo and termo in v for v in valores_validos):
                pontuacao += 1

        if pontuacao > melhor_pontuacao:
            melhor_pontuacao = pontuacao
            melhor_linha = idx

        # Exige todos os termos, preferencialmente por célula exata ou por conteúdo inequívoco.
        if pontuacao >= len(termos_obrigatorios) * 2:
            linha_cabecalho = idx
            break
        if all(any(termo and (termo == v or termo in v) for v in valores_validos) for termo in termos_obrigatorios):
            linha_cabecalho = idx
            break

    if linha_cabecalho is None:
        linha_cabecalho = melhor_linha if melhor_linha is not None else 0

    df = pd.read_excel(BytesIO(bytes_arquivo), sheet_name=aba, header=linha_cabecalho)
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~pd.Series(df.columns).astype(str).str.startswith("Unnamed").values]
    return df

def ler_parametros(bytes_arquivo, xls):
    aba = localizar_aba(xls, ["PARAMETROS_REAJUSTE", "PARAMETROS", "Parâmetros"])
    if not aba:
        return {}

    # As planilhas-modelo do modo Consumo por Itens/Ciclo possuem linhas de título
    # antes da tabela Campo/Valor. Exigir esses termos evita interpretar o título
    # da aba como cabeçalho e perder parâmetros como índice e premissas fiscais.
    df = ler_aba_com_cabecalho(bytes_arquivo, aba, termos_obrigatorios=["Campo", "Valor"])
    if df.empty or len(df.columns) < 2:
        return {}

    # Tenta estrutura Campo/Valor; se não houver, usa as duas primeiras colunas.
    col_campo = localizar_coluna(df, ["Campo", "Parâmetro", "Parametro", "Indicador"])
    col_valor = localizar_coluna(df, ["Valor", "Resultado"])

    if col_campo is None or col_valor is None:
        col_campo, col_valor = df.columns[0], df.columns[1]

    params = {}
    for _, row in df.iterrows():
        chave = normalizar_texto(row.get(col_campo, ""))
        if chave:
            params[chave] = row.get(col_valor)
    return params



def ler_eventos_historicos_anteriores(bytes_arquivo, xls):
    """Lê a aba opcional EVENTOS_HISTORICOS_ANTERIORES.

    A aba é apenas memória formal do contrato. Seus valores não alteram o cálculo
    do Valor Total Atualizado, que permanece execução atualizada + saldo remanescente atualizado.
    """
    aba = localizar_aba(xls, ["EVENTOS_HISTORICOS_ANTERIORES", "EVENTOS HISTORICOS ANTERIORES", "HISTORICO_ANTERIOR"])
    if not aba:
        return []
    try:
        df = ler_aba_com_cabecalho(bytes_arquivo, aba, termos_obrigatorios=["Tipo"])
    except Exception:
        return []
    if df.empty:
        return []

    col_tipo = localizar_coluna(df, ["Tipo de evento", "Tipo"])
    col_ciclo = localizar_coluna(df, ["Ciclo"])
    col_data = localizar_coluna(df, ["Data"])
    col_valor = localizar_coluna(df, ["Valor formalizado/impacto", "Valor atualizado/formalizado", "Valor original", "Valor"])
    col_incorp = localizar_coluna(df, ["Incorporado ao valor formalizado?", "Incorporado", "Incorporado ao valor formalizado"])
    col_obs = localizar_coluna(df, ["Observação", "Observacao", "Obs"])

    eventos = []
    for _, row in df.iterrows():
        evento = {
            "Tipo de evento": str(row.get(col_tipo, "") if col_tipo else "").strip(),
            "Ciclo": str(row.get(col_ciclo, "") if col_ciclo else "").strip(),
            "Data": row.get(col_data, "") if col_data else "",
            "Valor formalizado/impacto": row.get(col_valor, "") if col_valor else "",
            "Incorporado ao valor formalizado?": str(row.get(col_incorp, "") if col_incorp else "").strip(),
            "Observação": str(row.get(col_obs, "") if col_obs else "").strip(),
        }
        tem_conteudo = any([
            evento["Tipo de evento"], str(evento["Data"]).strip(), str(evento["Valor formalizado/impacto"]).strip(), evento["Observação"]
        ])
        if tem_conteudo and evento["Tipo de evento"].upper() not in ["TOTAL", "ORIENTAÇÃO", "ORIENTACAO"]:
            eventos.append(evento)
    return eventos


def contexto_contratual_de_parametros(params, bytes_arquivo=None, xls=None):
    """Retorna o contexto contratual anterior informado nos módulos 01/02 ou no XLS.

    O valor formalizado anterior representa a fotografia consolidada do contrato
    antes da análise atual, incluindo reajustes, repactuações, aditivos e supressões
    já incorporados. Quando não houver contexto, todos os valores retornam vazios/zero.
    """
    contexto = (
        st.session_state.get("dados_admissibilidade", {}).get("contexto_contratual_anterior")
        or st.session_state.get("contexto_contratual_anterior")
        or {}
    )

    valor_original = numero_br(
        contexto.get("valor_original_contrato")
        or params.get("valor_original_do_contrato_contexto")
        or params.get("valor_original_do_contrato")
        or 0
    )
    valor_formalizado = numero_br(
        contexto.get("valor_formalizado_anterior")
        or params.get("valor_contratual_formalizado_antes_desta_analise")
        or params.get("valor_formalizado_antes_desta_analise")
        or params.get("valor_contratual_formalizado_atual")
        or 0
    )
    ultimo_ciclo = (
        contexto.get("ultimo_ciclo_concedido")
        or params.get("ultimo_ciclo_ja_concedido_formalizado")
        or params.get("ultimo_ciclo_concedido_formalizado")
        or ""
    )
    observacao = texto_seguro(
        contexto.get("observacao_historico")
        or params.get("observacao_sobre_historico_anterior")
        or params.get("observacao_sobre_o_historico_anterior")
        or params.get("observacao_historico_anterior")
        or "",
        ""
    )

    dados_admissibilidade = st.session_state.get("dados_admissibilidade", {}) or {}
    data_base_reajuste = (
        contexto.get("data_base_reajuste")
        or contexto.get("data_base_original")
        or contexto.get("data_base_anterior")
        or dados_admissibilidade.get("data_base_original")
        or dados_admissibilidade.get("data_base")
        or dados_admissibilidade.get("data_base_anterior")
        or params.get("data_base_original")
        or params.get("data_base_para_reajuste")
        or params.get("data_base_anterior")
        or params.get("data_base")
        or ""
    )

    eventos_historicos = contexto.get("eventos_historicos_anteriores") or []
    if not eventos_historicos and bytes_arquivo is not None and xls is not None:
        eventos_historicos = ler_eventos_historicos_anteriores(bytes_arquivo, xls)
    if not eventos_historicos and params.get("eventos_historicos_anteriores"):
        try:
            import json
            eventos_historicos = json.loads(str(params.get("eventos_historicos_anteriores")))
        except Exception:
            eventos_historicos = []
    eventos_normalizados = []
    for evento in eventos_historicos or []:
        if not isinstance(evento, dict):
            continue
        valor_evento = evento.get("Valor formalizado/impacto", evento.get("Valor atualizado/formalizado", evento.get("Valor original", "")))
        evento_norm = {
            "Tipo de evento": str(evento.get("Tipo de evento", "")).strip(),
            "Ciclo": str(evento.get("Ciclo", "")).strip(),
            "Data": evento.get("Data", ""),
            "Valor formalizado/impacto": valor_evento,
            "Incorporado ao valor formalizado?": str(evento.get("Incorporado ao valor formalizado?", "")).strip(),
            "Observação": texto_seguro(evento.get("Observação", ""), ""),
        }
        if any([evento_norm["Tipo de evento"], str(evento_norm["Data"]).strip(), str(evento_norm["Valor formalizado/impacto"]).strip(), evento_norm["Observação"]]):
            eventos_normalizados.append(evento_norm)
    eventos_historicos = eventos_normalizados

    return {
        "valor_original_contrato": float(valor_original or 0.0),
        "valor_formalizado_anterior": float(valor_formalizado or 0.0),
        "ultimo_ciclo_concedido": str(ultimo_ciclo).strip(),
        "observacao_historico": texto_seguro(observacao, ""),
        "eventos_historicos_anteriores": eventos_historicos,
        "data_base_reajuste": data_base_reajuste,
        "contexto_informado": bool(valor_original or valor_formalizado or str(ultimo_ciclo).strip() or str(observacao).strip() or str(data_base_reajuste).strip() or eventos_historicos),
    }


def aditivo_informativo_ja_incorporado(tratamento):
    texto = normalizar_texto(tratamento)
    return "informativo" in texto or "ja_incorporado" in texto or "ja_incluido" in texto or "formalizado_anterior" in texto or "valor_formalizado" in texto


def dataframe_ciclos_de_session_state():
    adm = st.session_state.get("dados_admissibilidade", {})
    ciclos = adm.get("ciclos") or adm.get("detalhamento_ciclos") or []
    if not ciclos:
        return pd.DataFrame()

    linhas = []
    for c in ciclos:
        ciclo = c.get("ciclo") or c.get("Ciclo") or c.get("nome") or ""
        ciclo = normalizar_ciclo(ciclo)
        variacao = (
            c.get("variacao")
            if "variacao" in c
            else c.get("Variação", c.get("var", c.get("percentual")))
        )
        fator = c.get("fator") or c.get("Fator") or c.get("fator_ciclo")
        fator_acum = c.get("fator_acumulado") or c.get("Fator acumulado") or c.get("fator_acum")

        var_dec = percentual_para_decimal(variacao)
        fator_calc = fator_de_valor(fator, variacao=var_dec)
        if fator_acum is None or fator_acum == "":
            fator_acum = fator_calc
        else:
            fator_acum = fator_de_valor(fator_acum)

        linhas.append({
            "Ciclo": ciclo,
            "Data-base": c.get("data_base") or c.get("Data-base") or c.get("Base") or "",
            "Intervalo do índice": c.get("intervalo_indice") or c.get("intervalo") or c.get("Janela") or "",
            "Janela de admissibilidade": c.get("janela_admissibilidade") or c.get("JanelaAdm") or "",
            "Data do pedido": c.get("data_pedido") or c.get("Pedido") or "",
            "Início financeiro": c.get("financeiro_inicio") or c.get("Inicio financeiro") or c.get("Início financeiro") or "",
            "Fim financeiro": c.get("financeiro_fim") or c.get("Fim financeiro") or "",
            "Situação": c.get("situacao") or c.get("Situação") or "",
            "Situação automática": c.get("situacao_automatica") or c.get("Situação automática") or c.get("situacao") or "",
            "Acordo negocial": "Sim" if c.get("superacao_negocial", False) else "Não",
            "Situação aplicada": c.get("situacao_aplicada") or c.get("Situação aplicada") or c.get("situacao") or "",
            "Percentual apurado pelo índice": c.get("percentual_indice", var_dec),
            "Percentual aplicado": c.get("percentual_aplicado", var_dec),
            "Ciclo negativo": "Sim" if c.get("ciclo_negativo", False) else "Não",
            "Tratamento ciclo negativo": c.get("tratamento_ciclo_negativo", ""),
            "Justificativa negocial": c.get("justificativa_negocial", ""),
            "Referência documental": c.get("referencia_documental", ""),
            "Variação": var_dec,
            "Fator": fator_calc,
            "Fator acumulado": fator_acum,
        })
    return pd.DataFrame(linhas)


def ler_ciclos(bytes_arquivo, xls):
    # Prioridade: dados da sessão. Caso não existam, usa Excel.
    df_session = dataframe_ciclos_de_session_state()
    if not df_session.empty:
        return padronizar_ciclos(df_session), "session_state"

    # Para evitar confusão com ITENS_CICLOS, aceitar apenas abas de ciclos por correspondência exata.
    # O modelo Consumo por Itens/Ciclo usa CICLOS_APURADOS.
    mapa_exato = {normalizar_texto(s): s for s in xls.sheet_names}
    aba = mapa_exato.get("ciclos") or mapa_exato.get("ciclo") or mapa_exato.get("ciclos_apurados")
    if not aba and mapa_exato.get("parametros"):
        df = pd.read_excel(BytesIO(bytes_arquivo), sheet_name=mapa_exato["parametros"], header=0)
        return padronizar_ciclos(df), "parametros"
    if not aba:
        return pd.DataFrame(), "indisponível"

    df = ler_aba_com_cabecalho(bytes_arquivo, aba, termos_obrigatorios=["Ciclo"])
    return padronizar_ciclos(df), "arquivo"


def padronizar_ciclos(df):
    if df.empty:
        return pd.DataFrame()

    col_ciclo = localizar_coluna(df, ["Ciclo"])
    col_base = localizar_coluna(df, ["Data-base", "Base", "Data início", "DATA_INICIO"])
    col_intervalo = localizar_coluna(df, ["Intervalo do índice", "Intervalo", "Janela"])
    col_janela_adm = localizar_coluna(df, ["Janela de admissibilidade", "JanelaAdm"])
    col_pedido = localizar_coluna(df, ["Data do pedido", "Pedido"])
    col_inicio_financeiro = localizar_coluna(df, ["Início financeiro", "Inicio financeiro", "Início dos efeitos financeiros", "Inicio dos efeitos financeiros", "Efeitos financeiros"])
    col_fim_financeiro = localizar_coluna(df, ["Fim financeiro", "Fim dos efeitos financeiros", "Fim efeito financeiro"])
    col_situacao = localizar_coluna(df, ["Situação", "Resultado", "SITUACAO"])
    col_variacao = localizar_coluna(df, ["Variação", "Variacao", "Percentual"])
    col_fator = localizar_coluna(df, ["Fator"])
    col_fator_acum = localizar_coluna(df, ["Fator acumulado", "Fator acumulado final", "Fator Acumulado"])
    col_tratamento = localizar_coluna(df, ["Tratamento financeiro do ciclo", "Tratamento financeiro", "Tratamento"])
    col_situacao_automatica = localizar_coluna(df, ["Situação automática", "Situacao automatica"])
    col_superacao = localizar_coluna(df, ["Acordo negocial", "Superação negocial", "Superacao negocial", "Superar preclusão"])
    col_situacao_aplicada = localizar_coluna(df, ["Situação aplicada", "Situacao aplicada"])
    col_percentual_indice = localizar_coluna(df, ["Percentual apurado pelo índice", "Percentual indice", "Percentual apurado"])
    col_percentual_aplicado = localizar_coluna(df, ["Percentual aplicado", "Percentual negocial"])
    col_justificativa_negocial = localizar_coluna(df, ["Justificativa negocial", "Justificativa"])
    col_referencia_documental = localizar_coluna(df, ["Referência documental", "Referencia documental"])
    col_ciclo_negativo = localizar_coluna(df, ["Ciclo negativo", "Ciclo negativo?"])
    col_tratamento_negativo = localizar_coluna(df, ["Tratamento ciclo negativo", "Tratamento negativo"])

    linhas = []
    fator_acumulado_calculado = 1.0

    for _, row in df.iterrows():
        ciclo = normalizar_ciclo(row.get(col_ciclo, "")) if col_ciclo else ""
        if not ciclo:
            continue

        situacao = row.get(col_situacao, "") if col_situacao else ""
        situacao_automatica = row.get(col_situacao_automatica, situacao) if col_situacao_automatica else situacao
        superacao_txt = row.get(col_superacao, "") if col_superacao else ""
        superacao_negocial = normalizar_texto(superacao_txt) in ["sim", "s", "true", "1", "yes"] or "acordo" in normalizar_texto(row.get(col_situacao_aplicada, "") if col_situacao_aplicada else "")
        situacao_aplicada = row.get(col_situacao_aplicada, situacao) if col_situacao_aplicada else situacao
        tratamento = row.get(col_tratamento, "A apurar") if col_tratamento else ("Precluso" if "PRECLUS" in str(situacao_automatica).upper() and not superacao_negocial else "A apurar")
        variacao_indice = percentual_para_decimal(row.get(col_percentual_indice, row.get(col_variacao, 0))) if (col_percentual_indice or col_variacao) else 0.0
        variacao = percentual_para_decimal(row.get(col_percentual_aplicado, row.get(col_variacao, 0))) if (col_percentual_aplicado or col_variacao) else 0.0
        ciclo_negativo_txt = row.get(col_ciclo_negativo, "") if col_ciclo_negativo else ""
        ciclo_negativo = normalizar_texto(ciclo_negativo_txt) in ["sim", "s", "true", "1", "yes"] or variacao_indice < 0 or (col_percentual_indice is None and variacao < 0)
        tratamento_negativo = row.get(col_tratamento_negativo, "") if col_tratamento_negativo else ""
        if ciclo_negativo and not superacao_negocial:
            variacao = 0.0
            if not tratamento_negativo:
                tratamento_negativo = "Ciclo negativo - percentual aplicado 0,00% no acumulado"
            if "ciclo negativo" not in normalizar_texto(str(situacao_aplicada)):
                situacao_aplicada = f"{situacao_aplicada} | CICLO NEGATIVO (APLICADO 0,00%)"
        fator = fator_operacional(fator_de_valor(row.get(col_fator, None) if col_fator else None, variacao=variacao))
        # Preserva o padrão histórico da ferramenta: percentuais/fatores operacionais
        # são calculados com o fator do ciclo em 4 casas decimais.
        variacao = fator - 1
        if ciclo_negativo and not superacao_negocial:
            fator = 1.0
            variacao = 0.0

        # Não usar o fator acumulado importado com casas residuais para cálculo financeiro.
        # A ferramenta deve reproduzir o padrão histórico da Planilha Executiva:
        # produto dos fatores de ciclo arredondados a 4 casas.
        fator_acum = fator_acumulado_calculado * fator

        # Fator acumulado efetivo para cálculos financeiros: não aplica ciclo precluso/adiantado comum.
        if (not superacao_negocial) and (contem_preclusao_ou_adiantamento(situacao_automatica or situacao) or tratamento_sem_retroativo(tratamento, situacao_automatica or situacao)):
            fator_acumulado_calculado = fator_acumulado_calculado
            fator_acum_efetivo = fator_acumulado_calculado
            fator_efetivo = 1.0
        else:
            fator_acumulado_calculado = fator_acumulado_calculado * fator
            fator_acum_efetivo = fator_acumulado_calculado
            fator_efetivo = fator

        linhas.append({
            "Ciclo": ciclo,
            "Data-base": row.get(col_base, "") if col_base else "",
            "Intervalo do índice": row.get(col_intervalo, "") if col_intervalo else "",
            "Janela de admissibilidade": row.get(col_janela_adm, "") if col_janela_adm else "",
            "Data do pedido": row.get(col_pedido, "") if col_pedido else "",
            "Início financeiro": row.get(col_inicio_financeiro, "") if col_inicio_financeiro else "",
            "Fim financeiro": row.get(col_fim_financeiro, "") if col_fim_financeiro else "",
            "Situação": situacao_aplicada if superacao_negocial else situacao,
            "Situação automática": situacao_automatica,
            "Acordo negocial": "Sim" if superacao_negocial else "Não",
            "Situação aplicada": situacao_aplicada,
            "Percentual apurado pelo índice": variacao_indice,
            "Percentual aplicado": variacao,
            "Ciclo negativo": "Sim" if ciclo_negativo else "Não",
            "Tratamento ciclo negativo": tratamento_negativo,
            "Justificativa negocial": row.get(col_justificativa_negocial, "") if col_justificativa_negocial else "",
            "Referência documental": row.get(col_referencia_documental, "") if col_referencia_documental else "",
            "Tratamento financeiro do ciclo": tratamento,
            "Variação": variacao,
            "Fator": fator,
            "Fator acumulado": fator_acum,
            "Fator acumulado efetivo": fator_acum_efetivo,
            "Fator ciclo efetivo": fator_efetivo,
        })

    ciclos = pd.DataFrame(linhas)
    if not ciclos.empty:
        ciclos = ciclos.sort_values(by="Ciclo", key=lambda s: s.map(numero_ciclo)).reset_index(drop=True)
    return ciclos


def ler_financeiro(bytes_arquivo, xls, ciclos):
    aba = localizar_aba(xls, ["BASE_EXECUCAO_MENSAL", "FINANCEIRO_MENSAL", "RETROATIVO", "Financeiro"])
    if not aba:
        raise ValueError("Aba BASE_EXECUCAO_MENSAL/FINANCEIRO_MENSAL não encontrada.")

    df = ler_aba_com_cabecalho(bytes_arquivo, aba, termos_obrigatorios=["Valor"])
    if df.empty:
        raise ValueError("Aba BASE_EXECUCAO_MENSAL/FINANCEIRO_MENSAL está vazia.")

    col_ciclo = localizar_coluna(df, ["Ciclo"])
    col_comp = localizar_coluna(df, ["Competência", "Competencia"])
    col_valor = localizar_coluna(df, ["Valor bruto medido/aprovado por competência", "Valor bruto medido", "Valor medido/aprovado", "Valor pago/faturado", "Valor bruto faturado", "Valor faturado", "Valor pago", "Valor medido", "Valor"])

    if col_valor is None:
        raise ValueError("Não foi encontrada coluna de valor bruto medido/aprovado na aba BASE_EXECUCAO_MENSAL/FINANCEIRO_MENSAL.")

    resultado = pd.DataFrame()
    resultado["Ciclo"] = df[col_ciclo].apply(normalizar_ciclo) if col_ciclo else ""
    resultado["Competência"] = df[col_comp] if col_comp else ""
    resultado["Valor pago/faturado"] = df[col_valor].apply(numero_br)

    if (resultado["Ciclo"] == "").all() and col_comp and not ciclos.empty:
        resultado["Ciclo"] = atribuir_ciclo_por_competencia(resultado["Competência"], ciclos)

    if (resultado["Ciclo"] == "").all():
        # Fallback: se não houver ciclo, assume C1.
        resultado["Ciclo"] = "C1"

    resultado = resultado[~resultado["Ciclo"].astype(str).str.upper().eq("TOTAL")].copy()
    resultado = resultado[resultado["Valor pago/faturado"].fillna(0) != 0].copy()
    return resultado.reset_index(drop=True)


def parse_intervalo_mensal(intervalo):
    if pd.isna(intervalo):
        return None, None
    texto = str(intervalo)
    matches = re.findall(r"(\d{1,2})[\/\-](\d{4})", texto)
    if len(matches) >= 2:
        ini = pd.Period(f"{int(matches[0][1])}-{int(matches[0][0]):02d}", freq="M")
        fim = pd.Period(f"{int(matches[1][1])}-{int(matches[1][0]):02d}", freq="M")
        return ini, fim
    return None, None


def normalizar_competencia_periodo(valor):
    """Converte datas/competências para Period mensal de forma robusta.

    Aceita: 04/2026, 01/04/2026, 30/04/2026, abr/2026,
    datetime/Timestamp do Excel e serial numérico do Excel.
    """
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except Exception:
        pass
    if isinstance(valor, pd.Period):
        return valor.asfreq("M")
    if isinstance(valor, (pd.Timestamp, datetime)):
        return pd.Timestamp(valor).to_period("M")
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        n = float(valor)
        if 20000 <= n <= 80000:
            try:
                return pd.to_datetime(n, unit="D", origin="1899-12-30").to_period("M")
            except Exception:
                pass
        if 190001 <= n <= 299912:
            inteiro = int(n)
            ano = inteiro // 100
            mes = inteiro % 100
            if 1 <= mes <= 12:
                return pd.Period(f"{ano}-{mes:02d}", freq="M")

    texto = str(valor).strip().lower()
    if not texto or texto in ["nan", "none", "nat", "total"]:
        return None
    texto = texto.replace(".", "/").replace("-", "/")

    if re.fullmatch(r"\d{5,6}(\.0)?", texto):
        try:
            n = float(texto)
            if 20000 <= n <= 80000:
                return pd.to_datetime(n, unit="D", origin="1899-12-30").to_period("M")
        except Exception:
            pass

    meses = {
        "jan": 1, "janeiro": 1, "fev": 2, "fevereiro": 2, "mar": 3, "marco": 3, "março": 3,
        "abr": 4, "abril": 4, "mai": 5, "maio": 5, "jun": 6, "junho": 6,
        "jul": 7, "julho": 7, "ago": 8, "agosto": 8, "set": 9, "setembro": 9,
        "out": 10, "outubro": 10, "nov": 11, "novembro": 11, "dez": 12, "dezembro": 12,
    }
    m = re.search(r"([a-zçãéíóú]+)\s*/\s*(\d{2,4})", texto, flags=re.IGNORECASE)
    if m:
        mes_txt = normalizar_texto(m.group(1))
        ano = int(m.group(2))
        if ano < 100:
            ano += 2000
        mes = meses.get(mes_txt)
        if mes:
            return pd.Period(f"{ano}-{mes:02d}", freq="M")

    m = re.fullmatch(r"(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{2,4})", texto)
    if m:
        dia = int(m.group(1)); mes = int(m.group(2)); ano = int(m.group(3))
        if ano < 100:
            ano += 2000
        if 1 <= mes <= 12 and 1 <= dia <= 31:
            return pd.Period(f"{ano}-{mes:02d}", freq="M")

    m = re.fullmatch(r"(\d{1,2})\s*/\s*(\d{2,4})", texto)
    if m:
        mes = int(m.group(1)); ano = int(m.group(2))
        if ano < 100:
            ano += 2000
        if 1 <= mes <= 12:
            return pd.Period(f"{ano}-{mes:02d}", freq="M")

    m = re.fullmatch(r"(\d{4})\s*/\s*(\d{1,2})(?:\s*/\s*(\d{1,2}))?", texto)
    if m:
        ano = int(m.group(1)); mes = int(m.group(2))
        if 1 <= mes <= 12:
            return pd.Period(f"{ano}-{mes:02d}", freq="M")

    try:
        dt = pd.to_datetime(valor, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.to_period("M")
    except Exception:
        pass
    return None

def periodo_para_label_br(periodo):
    if periodo is None:
        return ""
    try:
        p = pd.Period(periodo, freq="M")
        return f"{p.month:02d}/{p.year}"
    except Exception:
        return ""


def inicio_financeiro_periodo_por_ciclo(ciclos):
    mapa = {}
    if not isinstance(ciclos, pd.DataFrame) or ciclos.empty:
        return mapa
    for _, row in ciclos.iterrows():
        ciclo = normalizar_ciclo(row.get("Ciclo", ""))
        inicio = normalizar_competencia_periodo(row.get("Início financeiro", ""))
        if ciclo and inicio is not None:
            mapa[ciclo] = inicio
    return mapa


def atribuir_ciclo_por_competencia(competencias, ciclos):
    intervalos = []
    for _, row in ciclos.iterrows():
        ini, fim = parse_intervalo_mensal(row.get("Intervalo do índice", ""))
        if ini is not None:
            intervalos.append((row["Ciclo"], ini, fim))

    atribuidos = []
    for comp in competencias:
        periodo = normalizar_competencia_periodo(comp)
        if periodo is None:
            atribuidos.append("")
            continue

        ciclo_encontrado = ""
        for ciclo, ini, fim in intervalos:
            if ini <= periodo <= fim:
                ciclo_encontrado = ciclo
                break
        atribuidos.append(ciclo_encontrado)
    return atribuidos


def ler_itens(bytes_arquivo, xls):
    aba = localizar_aba(xls, ["ITENS_REMANESCENTES", "ITENS_REMANESC", "ITENS_CICLOS", "Itens"])
    if not aba:
        raise ValueError("Aba ITENS_REMANESCENTES não encontrada.")

    df = ler_aba_com_cabecalho(bytes_arquivo, aba, termos_obrigatorios=["Item"])
    if df.empty:
        raise ValueError("Aba ITENS_REMANESCENTES está vazia.")

    col_item = localizar_coluna(df, ["Item"])
    col_qtd = localizar_coluna(df, ["Quantidade contratada", "QTD_CONTRATADA", "Qtd C0", "Quantidade", "Qtd"])
    col_vu = localizar_coluna(df, ["Valor unitário original", "VU_ORIGINAL", "VU C0", "VU Original", "Valor Unitario"])
    col_total = localizar_coluna(df, ["Valor total", "VALOR_TOTAL", "TOTAL C0", "Total"])

    if col_item is None or col_qtd is None or col_vu is None:
        raise ValueError("Aba ITENS_REMANESCENTES precisa conter Item, Quantidade contratada e Valor unitário original.")

    valor_total_referencia = None
    if col_total is not None:
        total_rows = df[df[col_item].astype(str).str.strip().str.upper().eq("TOTAL")]
        if not total_rows.empty:
            totais_validos = total_rows[col_total].apply(numero_br)
            totais_validos = totais_validos[totais_validos > 0]
            if not totais_validos.empty:
                valor_total_referencia = float(totais_validos.iloc[-1])

    df = df[~df[col_item].astype(str).str.strip().str.upper().eq("TOTAL")].copy()
    df = df[df[col_item].notna()].copy()
    df = df[df[col_item].astype(str).str.strip() != ""].copy()

    itens = pd.DataFrame(index=df.index)
    itens["Item"] = df[col_item]
    itens["Quantidade contratada"] = df[col_qtd].apply(numero_br)
    itens["Valor unitário original"] = df[col_vu].apply(numero_br)

    if col_total is not None:
        total_informado = df[col_total].apply(numero_br)
        itens["Valor total original"] = total_informado.where(
            total_informado > 0,
            itens["Quantidade contratada"] * itens["Valor unitário original"],
        )
    else:
        itens["Valor total original"] = itens["Quantidade contratada"] * itens["Valor unitário original"]

    itens = itens[
        (itens["Quantidade contratada"] > 0)
        | (itens["Valor unitário original"] > 0)
        | (itens["Valor total original"] > 0)
    ].copy()

    colunas_rem = []
    for col in df.columns:
        n = normalizar_texto(col)
        nome_original = str(col)
        possui_ciclo = re.search(r"C\s*\d+", nome_original, flags=re.IGNORECASE) is not None
        eh_inicio_ciclo = "inicio" in n or "inicial" in n or possui_ciclo
        eh_coluna_excluida = "atual" in n or "data_corte" in n or "corte" in n
        eh_quantidade_remanescente = "remanescente" in n or n.startswith("qtd_rem_")
        if eh_quantidade_remanescente and "valor" not in n and "total" not in n and eh_inicio_ciclo and not eh_coluna_excluida:
            colunas_rem.append(col)

    col_consumido = localizar_coluna(df, ["Consumido no Ciclo", "Consumido"])
    if not colunas_rem:
        col_qtd_rem = localizar_coluna(df, ["Qtd Remanescente", "Quantidade Remanescente", "Remanescente"])
        if col_qtd_rem:
            colunas_rem = [col_qtd_rem]

    for col in colunas_rem:
        itens[col] = df.loc[itens.index, col].apply(numero_br) if col in df.columns else 0.0

    if col_consumido:
        itens["Consumido no Ciclo"] = df.loc[itens.index, col_consumido].apply(numero_br)

    if valor_total_referencia is not None:
        itens.attrs["valor_original_contrato_referencia"] = valor_total_referencia

    return itens, colunas_rem


# ============================================================
# Cálculos financeiros/executivos
# ============================================================

def fator_fmt(valor):
    try:
        valor = float(valor)
    except Exception:
        valor = 1.0
    return f"{valor:.4f}".replace('.', ',')


def obter_tratamento(row):
    for col in ["Tratamento financeiro do ciclo", "Tratamento", "Situação"]:
        if col in row.index and str(row.get(col, "")).strip():
            return str(row.get(col, "")).strip()
    return "A apurar"


def tratamento_sem_retroativo(tratamento, situacao=""):
    texto = normalizar_texto(f"{tratamento} {situacao}")
    return any(x in texto for x in ["ja_concedido", "preclus", "sem_efeito", "adiantado"])


def mapa_fatores(ciclos):
    mapa = {"C0": {
        "fator": 1.0,
        "fator_acumulado": 1.0,
        "fator_acumulado_efetivo": 1.0,
        "fator_retroativo": 1.0,
        "situacao": "BASE",
        "tratamento": "Base",
        "variacao": 0.0,
    }}
    if ciclos.empty:
        return mapa

    for _, row in ciclos.iterrows():
        ciclo = row.get("Ciclo", "")
        situacao = row.get("Situação", "")
        tratamento = row.get("Tratamento financeiro do ciclo", "A apurar")
        fator_efetivo = float(row.get("Fator acumulado efetivo", row.get("Fator acumulado", 1.0)) or 1.0)
        mapa[ciclo] = {
            "fator": float(row.get("Fator", 1.0) or 1.0),
            "fator_acumulado": float(row.get("Fator acumulado", 1.0) or 1.0),
            "fator_acumulado_efetivo": fator_efetivo,
            "fator_retroativo": 1.0 if tratamento_sem_retroativo(tratamento, situacao) else fator_efetivo,
            "situacao": situacao,
            "tratamento": tratamento,
            "variacao": float(row.get("Variação", 0.0) or 0.0),
        }
    return mapa


def financeiro_com_efeito_financeiro(df_financeiro, ciclos):
    """Aplica a regra pétrea de efeito financeiro por competência.

    Competências anteriores ao início financeiro do ciclo são mantidas em memória,
    mas não geram delta/retroativo a pagar. Isso evita pagar o lapso entre a data
    em que o reajuste poderia produzir efeitos e a data em que o pedido foi feito.
    """
    if df_financeiro is None or not isinstance(df_financeiro, pd.DataFrame) or df_financeiro.empty:
        return pd.DataFrame(columns=[
            "Ciclo", "Competência", "Valor pago/faturado", "Competência sem efeito financeiro?",
            "Fator aplicado", "Valor teórico calculado", "Delta computável", "Delta sem efeito financeiro"
        ])

    fatores = mapa_fatores(ciclos)
    inicios_fin = inicio_financeiro_periodo_por_ciclo(ciclos)
    linhas = []
    for _, row in df_financeiro.iterrows():
        ciclo = normalizar_ciclo(row.get("Ciclo", "")) or "C1"
        comp = row.get("Competência", "")
        periodo = normalizar_competencia_periodo(comp)
        valor = numero_seguro(row.get("Valor pago/faturado", 0.0), 0.0)
        info = fatores.get(ciclo, {})
        fator_retroativo = float(info.get("fator_retroativo", 1.0) or 1.0)
        inicio_fin = inicios_fin.get(ciclo)
        sem_efeito = False
        if periodo is not None and inicio_fin is not None and periodo < inicio_fin:
            sem_efeito = True
        valor_teorico = valor if sem_efeito else valor * fator_retroativo
        delta_computavel = 0.0 if sem_efeito else (valor_teorico - valor)
        delta_sem_efeito = (valor * fator_retroativo - valor) if sem_efeito else 0.0
        linhas.append({
            "Ciclo": ciclo,
            "Competência": periodo_para_label_br(periodo) or str(comp),
            "Valor pago/faturado": valor,
            "Competência sem efeito financeiro?": "Sim" if sem_efeito else "Não",
            "Fator aplicado": 1.0 if sem_efeito else fator_retroativo,
            "Valor teórico calculado": valor_teorico,
            "Delta computável": delta_computavel,
            "Delta sem efeito financeiro": delta_sem_efeito,
            "Situação": info.get("situacao", ""),
            "Tratamento financeiro": info.get("tratamento", ""),
        })
    return pd.DataFrame(linhas)


def meses_sem_efeito_financeiro(df_financeiro, ciclos):
    """Consolida competências que existem apenas como memória, sem efeito financeiro.

    Regra pétrea: quando o pedido foi apresentado depois do marco em que o ciclo
    poderia produzir efeitos, as competências do lapso são demonstradas, mas o
    delta teórico desses meses não compõe o retroativo a pagar.
    """
    df = financeiro_com_efeito_financeiro(df_financeiro, ciclos)
    colunas = [
        "Ciclo", "Competência", "Valor base", "Fator teórico", "Valor teórico se aplicado",
        "Delta não devido", "Fundamento"
    ]
    if df.empty:
        return pd.DataFrame(columns=colunas)
    df = df[df["Competência sem efeito financeiro?"].astype(str).str.upper().eq("SIM")].copy()
    if df.empty:
        return pd.DataFrame(columns=colunas)
    df["Valor base"] = df["Valor pago/faturado"]
    df["Fator teórico"] = df.apply(lambda r: (r["Valor pago/faturado"] + r["Delta sem efeito financeiro"]) / r["Valor pago/faturado"] if abs(r["Valor pago/faturado"]) > 0.004 else 1.0, axis=1)
    df["Valor teórico se aplicado"] = (df["Valor pago/faturado"] + df["Delta sem efeito financeiro"]).round(2)
    df["Delta não devido"] = df["Delta sem efeito financeiro"].round(2)
    df["Fundamento"] = "Competência anterior ao início dos efeitos financeiros do pedido; demonstrada apenas para memória, sem compor o retroativo a pagar."
    return df[colunas].reset_index(drop=True)


def calcular_financeiro_por_ciclo(df_financeiro, ciclos):
    df_mensal = financeiro_com_efeito_financeiro(df_financeiro, ciclos)
    if df_mensal.empty:
        return pd.DataFrame(columns=[
            "Ciclo", "Situação", "Tratamento financeiro", "Fator aplicado ao retroativo",
            "Valor pago efetivo", "Valor teórico calculado", "Delta do ciclo"
        ])

    linhas = []
    for ciclo, grupo in df_mensal.groupby("Ciclo", dropna=False):
        ciclo = normalizar_ciclo(ciclo) or "C1"
        total_pago = float(grupo["Valor pago/faturado"].sum())
        devido = float(grupo["Valor teórico calculado"].sum())
        delta = float(grupo["Delta computável"].sum())
        info = mapa_fatores(ciclos).get(ciclo, {})
        linhas.append({
            "Ciclo": ciclo,
            "Situação": info.get("situacao", grupo["Situação"].iloc[0] if "Situação" in grupo else ""),
            "Tratamento financeiro": info.get("tratamento", grupo["Tratamento financeiro"].iloc[0] if "Tratamento financeiro" in grupo else ""),
            "Fator aplicado ao retroativo": float(info.get("fator_retroativo", 1.0) or 1.0),
            "Valor pago efetivo": total_pago,
            "Valor teórico calculado": devido,
            "Delta do ciclo": delta,
        })

    df = pd.DataFrame(linhas).sort_values(by="Ciclo", key=lambda s: s.map(numero_ciclo)).reset_index(drop=True)
    df.loc[len(df)] = {
        "Ciclo": "TOTAL",
        "Situação": "",
        "Tratamento financeiro": "",
        "Fator aplicado ao retroativo": "",
        "Valor pago efetivo": df["Valor pago efetivo"].sum(),
        "Valor teórico calculado": df["Valor teórico calculado"].sum(),
        "Delta do ciclo": df["Delta do ciclo"].sum(),
    }
    return df


def ciclo_da_coluna_remanescente(coluna, fallback="C1"):
    texto = str(coluna).upper()
    m = re.search(r"C\s*([0-9]+)", texto)
    if m:
        return f"C{int(m.group(1))}"
    return fallback


def calcular_remanescentes_valor(itens, colunas_remanescentes, ciclos):
    fatores = mapa_fatores(ciclos)
    linhas = []
    if not colunas_remanescentes:
        return pd.DataFrame(), None

    for col in colunas_remanescentes:
        ciclo = ciclo_da_coluna_remanescente(col, fallback="C1")
        info = fatores.get(ciclo, {})
        fator = float(info.get("fator_acumulado_efetivo", info.get("fator_acumulado", 1.0)) or 1.0)
        qtd_rem = itens[col].apply(numero_br)
        valor_original = qtd_rem * itens["Valor unitário original"]
        valor_atualizado = valor_original * fator
        linhas.append({
            "Ciclo": ciclo,
            "Remanescente original": float(valor_original.sum()),
            "Fator aplicado": fator,
            "Remanescente atualizado": float(valor_atualizado.sum()),
        })

    df = pd.DataFrame(linhas).sort_values(by="Ciclo", key=lambda s: s.map(numero_ciclo)).reset_index(drop=True)
    ciclo_ultimo = df.iloc[-1]["Ciclo"] if not df.empty else None
    return df, ciclo_ultimo




def construir_valores_unitarios_totais(itens, colunas_remanescentes, ciclos):
    """Monta tabela analítica por item e por ciclo.

    Regras:
    - C0 usa a quantidade contratada original e o valor unitário original.
    - C1, C2, C3... usam o remanescente informado no início do respectivo ciclo.
    - Ciclos preclusos usam o fator acumulado efetivo já saneado na aba CICLOS, isto é, repetem o valor anterior.
    """
    if itens.empty:
        return pd.DataFrame(columns=[
            "Item", "Ciclo", "Valor unitário", "Quantidade", "Total R$", "Ciclo precluso"
        ])

    fatores = mapa_fatores(ciclos)
    linhas = []

    # C0: fotografia original do item.
    for _, row in itens.iterrows():
        item = row.get("Item", "")
        vu_original = numero_br(row.get("Valor unitário original", 0))
        qtd_original = numero_br(row.get("Quantidade contratada", 0))
        linhas.append({
            "Item": item,
            "Ciclo": "C0",
            "Valor unitário": vu_original,
            "Quantidade": qtd_original,
            "Total R$": vu_original * qtd_original,
            "Ciclo precluso": False,
        })

    # Ciclos seguintes: quantidade = remanescente no início do ciclo.
    rem_por_ciclo = sorted(
        [(ciclo_da_coluna_remanescente(col, fallback="C1"), col) for col in colunas_remanescentes],
        key=lambda x: numero_ciclo(x[0])
    )

    for ciclo, col in rem_por_ciclo:
        info = fatores.get(ciclo, {})
        fator = float(info.get("fator_acumulado_efetivo", info.get("fator_acumulado", 1.0)) or 1.0)
        eh_precluso = contem_preclusao_ou_adiantamento(info.get("situacao", "")) or tratamento_sem_retroativo(info.get("tratamento", ""), info.get("situacao", ""))

        for _, row in itens.iterrows():
            item = row.get("Item", "")
            vu_original = numero_br(row.get("Valor unitário original", 0))
            qtd = numero_br(row.get(col, 0))
            vu_ciclo = vu_original * fator
            linhas.append({
                "Item": item,
                "Ciclo": ciclo,
                "Valor unitário": vu_ciclo,
                "Quantidade": qtd,
                "Total R$": vu_ciclo * qtd,
                "Ciclo precluso": bool(eh_precluso),
            })

    return pd.DataFrame(linhas)


def gerar_excel_valores_unitarios_por_ciclo(df_valores, ciclos):
    """Gera Excel analítico com valores unitários e totais por ciclo.

    Preserva a aba detalhada VALORES_POR_CICLO e acrescenta uma visão matricial
    em MATRIZ_VALORES_CICLO, sem alterar a leitura do arquivo de coleta ou os
    cálculos do Valor Global.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        fmt_header = workbook.add_format({
            "bold": True,
            "font_color": "white",
            "bg_color": "#1F4E79",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        })
        fmt_subheader = workbook.add_format({
            "bold": True,
            "font_color": "white",
            "bg_color": "#5B9BD5",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        })
        fmt_money = workbook.add_format({"num_format": 'R$ #,##0.00', "border": 1})
        fmt_money_red = workbook.add_format({"num_format": 'R$ #,##0.00', "font_color": "#C00000", "border": 1})
        fmt_num = workbook.add_format({"num_format": '#,##0.00', "border": 1})
        fmt_num_red = workbook.add_format({"num_format": '#,##0.00', "font_color": "#C00000", "border": 1})
        fmt_text = workbook.add_format({"border": 1})
        fmt_text_red = workbook.add_format({"font_color": "#C00000", "border": 1})
        fmt_pct = workbook.add_format({"num_format": "0.00%", "border": 1})
        fmt_factor = workbook.add_format({"num_format": "0.00", "border": 1})

        ciclo_fills = ["#FFFFFF", "#DDEBF7", "#E2F0D9", "#FFF2CC", "#EADCF8", "#E7E6E6", "#DDEBF7", "#E2F0D9"]
        def fmt_ciclo(ciclo, tipo="text", precluso=False, header=False):
            idx_cor = numero_ciclo(ciclo)
            bg = ciclo_fills[idx_cor % len(ciclo_fills)]
            base = {"border": 1, "bg_color": bg}
            if header:
                base.update({"bold": True, "align": "center", "valign": "vcenter", "text_wrap": True})
            if precluso:
                base["font_color"] = "#C00000"
            if tipo == "money":
                base["num_format"] = 'R$ #,##0.00'
            elif tipo == "num":
                base["num_format"] = '#,##0.00'
            elif tipo == "pct":
                base["num_format"] = "0.00%"
            elif tipo == "factor":
                base["num_format"] = "0.00"
            return workbook.add_format(base)


        df_export = df_valores.copy()
        # No Modo Consumo por Itens/Ciclo, a conferência fica mais legível por blocos:
        # primeiro todos os itens do C0, depois todos do C1, C2 etc.
        if not df_export.empty and "Ciclo" in df_export.columns:
            colunas_ordenacao = ["Ciclo"] + (["Item"] if "Item" in df_export.columns else [])
            df_export = df_export.sort_values(
                by=colunas_ordenacao,
                key=lambda col: col.map(numero_ciclo) if col.name == "Ciclo" else col.astype(str),
                kind="stable",
            ).reset_index(drop=True)
        df_export.to_excel(writer, sheet_name="VALORES_POR_CICLO", index=False)
        ws = writer.sheets["VALORES_POR_CICLO"]

        for col_idx, title in enumerate(df_export.columns):
            ws.write(0, col_idx, title, fmt_header)
        if not df_export.empty:
            ws.autofilter(0, 0, len(df_export), max(0, len(df_export.columns) - 1))
        ws.freeze_panes(1, 0)

        colunas = {col: idx for idx, col in enumerate(df_export.columns)}
        ws.set_column("A:A", 16)
        ws.set_column("B:B", 10)
        ws.set_column("C:C", 20)
        ws.set_column("D:D", 16)
        ws.set_column("E:E", 22)
        ws.set_column("F:F", 14)

        for row_idx, (_, row) in enumerate(df_export.iterrows(), start=1):
            precluso = bool(row.get("Ciclo precluso", False))
            ciclo_linha = row.get("Ciclo", "")
            for col, col_idx in colunas.items():
                value = row.get(col, "")
                if col in ["Valor unitário", "Total R$"]:
                    ws.write_number(row_idx, col_idx, numero_seguro(value, 0.0), fmt_ciclo(ciclo_linha, "money", precluso))
                elif col == "Quantidade":
                    ws.write_number(row_idx, col_idx, numero_seguro(value, 0.0), fmt_ciclo(ciclo_linha, "num", precluso))
                elif col == "Ciclo precluso":
                    ws.write(row_idx, col_idx, "Sim" if precluso else "Não", fmt_ciclo(ciclo_linha, "text", precluso))
                else:
                    ws.write(row_idx, col_idx, value, fmt_ciclo(ciclo_linha, "text", precluso))

        # Aba matricial horizontal: uma linha por item e blocos por ciclo.
        ws_m = workbook.add_worksheet("MATRIZ_VALORES_CICLO")
        writer.sheets["MATRIZ_VALORES_CICLO"] = ws_m

        if not df_export.empty:
            ciclos_ordenados = sorted(df_export["Ciclo"].dropna().astype(str).unique(), key=numero_ciclo)
            itens_ordenados = list(df_export["Item"].dropna().drop_duplicates())
        else:
            ciclos_ordenados = []
            itens_ordenados = []

        ws_m.write(0, 0, "Item", fmt_header)
        col_atual = 1
        for ciclo in ciclos_ordenados:
            ws_m.merge_range(0, col_atual, 0, col_atual + 2, ciclo, fmt_ciclo(ciclo, "text", False, header=True))
            ws_m.write(1, col_atual, "Valor unitário", fmt_ciclo(ciclo, "text", False, header=True))
            ws_m.write(1, col_atual + 1, "Quantidade", fmt_ciclo(ciclo, "text", False, header=True))
            ws_m.write(1, col_atual + 2, "Total R$", fmt_ciclo(ciclo, "text", False, header=True))
            col_atual += 3

        ws_m.set_column(0, 0, 16)
        if ciclos_ordenados:
            ws_m.set_column(1, 1 + (len(ciclos_ordenados) * 3), 18)

        ultima_linha_itens = 1
        for row_idx, item in enumerate(itens_ordenados, start=2):
            ultima_linha_itens = row_idx
            ws_m.write(row_idx, 0, item, fmt_text)
            col_atual = 1
            for ciclo in ciclos_ordenados:
                linha = df_export[(df_export["Item"].astype(str) == str(item)) & (df_export["Ciclo"].astype(str) == str(ciclo))]
                if linha.empty:
                    vu = qtd = total = 0.0
                    precluso = False
                else:
                    rec = linha.iloc[0]
                    vu = numero_seguro(rec.get("Valor unitário", 0), 0.0)
                    qtd = numero_seguro(rec.get("Quantidade", 0), 0.0)
                    total = numero_seguro(rec.get("Total R$", 0), 0.0)
                    precluso = bool(rec.get("Ciclo precluso", False))
                ws_m.write_number(row_idx, col_atual, vu, fmt_ciclo(ciclo, "money", precluso))
                ws_m.write_number(row_idx, col_atual + 1, qtd, fmt_ciclo(ciclo, "num", precluso))
                ws_m.write_number(row_idx, col_atual + 2, total, fmt_ciclo(ciclo, "money", precluso))
                col_atual += 3

        # Linha dinâmica de total por ciclo, posicionada logo abaixo do último item.
        total_row = ultima_linha_itens + 1
        ws_m.write(total_row, 0, "TOTAL REMANESCENTE / TOTAL DO CICLO", fmt_header)
        col_atual = 1
        for ciclo in ciclos_ordenados:
            excel_col_total = col_atual + 2
            # Soma apenas a coluna "Total R$" de cada bloco de ciclo.
            primeira_linha_excel = 3
            ultima_linha_excel = ultima_linha_itens + 1
            col_letter = chr(ord('A') + excel_col_total) if excel_col_total < 26 else None
            if col_letter:
                formula = f"=ROUND(SUM({col_letter}{primeira_linha_excel}:{col_letter}{ultima_linha_excel}),2)"
                ws_m.write_blank(total_row, col_atual, None, fmt_ciclo(ciclo, "text", False))
                ws_m.write_blank(total_row, col_atual + 1, None, fmt_ciclo(ciclo, "text", False))
                ws_m.write_formula(total_row, col_atual + 2, formula, fmt_ciclo(ciclo, "money", False, header=True))
            else:
                total_valor = df_export[df_export["Ciclo"].astype(str) == str(ciclo)]["Total R$"].apply(numero_seguro).sum()
                ws_m.write_blank(total_row, col_atual, None, fmt_ciclo(ciclo, "text", False))
                ws_m.write_blank(total_row, col_atual + 1, None, fmt_ciclo(ciclo, "text", False))
                ws_m.write_number(total_row, col_atual + 2, total_valor, fmt_ciclo(ciclo, "money", False, header=True))
            col_atual += 3

        df_ciclos = ciclos.copy() if isinstance(ciclos, pd.DataFrame) else pd.DataFrame()
        if (not isinstance(df_ciclos, pd.DataFrame)) or df_ciclos.empty or "Ciclo" not in df_ciclos.columns:
            if not df_export.empty and "Ciclo" in df_export.columns:
                ciclos_derivados = sorted(df_export["Ciclo"].dropna().astype(str).unique(), key=numero_ciclo)
                df_ciclos = pd.DataFrame({
                    "Ciclo": ciclos_derivados,
                    "Observação": ["Ciclo derivado da aba VALORES_POR_CICLO para conferência do remanescente." for _ in ciclos_derivados],
                })
            else:
                df_ciclos = pd.DataFrame(columns=["Ciclo"])
        df_ciclos = df_ciclos[df_ciclos["Ciclo"].astype(str).str.strip().ne("")].copy()
        df_ciclos = df_ciclos[~df_ciclos["Ciclo"].astype(str).str.upper().isin(["TOTAL", "CICLO"])].copy()
        df_ciclos = df_ciclos.drop_duplicates(subset=["Ciclo"], keep="first").reset_index(drop=True)
        df_ciclos.to_excel(writer, sheet_name="CICLOS_CONSIDERADOS", index=False)
        ws2 = writer.sheets["CICLOS_CONSIDERADOS"]
        for col_idx, title in enumerate(df_ciclos.columns):
            ws2.write(0, col_idx, title, fmt_header)
        for idx, col in enumerate(df_ciclos.columns):
            largura = 24 if str(col).lower() in ["janela de admissibilidade", "referência para preenchimento", "observação"] else max(14, min(38, len(str(col)) + 4))
            ws2.set_column(idx, idx, largura)
        for row_idx in range(1, len(df_ciclos) + 1):
            for col_idx, col in enumerate(df_ciclos.columns):
                valor = df_ciclos.iloc[row_idx - 1, col_idx]
                if col in ["Fator", "Fator acumulado", "Fator acumulado efetivo", "Fator ciclo efetivo"]:
                    try:
                        ciclo_linha = df_ciclos.iloc[row_idx - 1].get("Ciclo", "") if "Ciclo" in df_ciclos.columns else ""
                        ws2.write_number(row_idx, col_idx, float(valor), fmt_ciclo(ciclo_linha, "factor", False))
                    except Exception:
                        ciclo_linha = df_ciclos.iloc[row_idx - 1].get("Ciclo", "") if "Ciclo" in df_ciclos.columns else ""
                        ws2.write(row_idx, col_idx, valor, fmt_ciclo(ciclo_linha, "text", False))
                elif col in ["Variação", "Percentual aplicado", "Percentual apurado pelo índice"]:
                    try:
                        ciclo_linha = df_ciclos.iloc[row_idx - 1].get("Ciclo", "")
                        ws2.write_number(row_idx, col_idx, float(valor), fmt_ciclo(ciclo_linha, "pct", False))
                    except Exception:
                        ciclo_linha = df_ciclos.iloc[row_idx - 1].get("Ciclo", "")
                        ws2.write(row_idx, col_idx, valor, fmt_ciclo(ciclo_linha, "text", False))
                else:
                    ciclo_linha = df_ciclos.iloc[row_idx - 1].get("Ciclo", "") if "Ciclo" in df_ciclos.columns else ""
                    ws2.write(row_idx, col_idx, "" if pd.isna(valor) else valor, fmt_ciclo(ciclo_linha, "text", False))

    output.seek(0)
    return output.getvalue()



def gerar_planilha_executiva(resultado):
    """Gera uma Planilha Executiva consolidada da análise.

    A proposta é entregar um XLS editável, com leitura próxima ao relatório final:
    - RESUMO_FINANCEIRO: visão macro do contrato e delta por ciclo;
    - DETALHAMENTO_ITENS: matriz item x ciclos com quantidade, VU e total;
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        # >>> XLS_EXECUTIVO_VISUAL_V2
        fonte_executiva_xls = "Aptos"
        # <<< XLS_EXECUTIVO_VISUAL_V2

        fmt_title = workbook.add_format({
            "bold": True, "font_size": 14, "font_color": "#0B1F3A",
            "align": "left", "valign": "vcenter"
        })
        fmt_subtitle = workbook.add_format({
            "bold": True, "font_color": "white", "bg_color": "#1F4E79",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True
        })
        fmt_section = workbook.add_format({
            "bold": True, "font_color": "#123B63", "bg_color": "#EAF2F8",
            "border": 1, "align": "left", "valign": "vcenter"
        })
        fmt_text = workbook.add_format({"border": 1, "valign": "vcenter"})
        fmt_text_center = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter"})
        fmt_money = workbook.add_format({"border": 1, "num_format": 'R$ #,##0.00', "valign": "vcenter"})
        fmt_money_bold = workbook.add_format({"border": 1, "num_format": 'R$ #,##0.00', "bold": True, "bg_color": "#E2F0D9"})
        fmt_pct = workbook.add_format({"border": 1, "num_format": "0.00%", "valign": "vcenter"})
        fmt_factor = workbook.add_format({"border": 1, "num_format": "0.00", "valign": "vcenter"})
        fmt_red_money = workbook.add_format({"border": 1, "font_color": "#C00000", "num_format": 'R$ #,##0.00'})
        fmt_red_text = workbook.add_format({"border": 1, "font_color": "#C00000"})
        fmt_note = workbook.add_format({"italic": True, "font_color": "#64748B"})
        fmt_note_wrap = workbook.add_format({"italic": True, "font_color": "#64748B", "text_wrap": True, "valign": "top"})
        fmt_text_wrap = workbook.add_format({"border": 1, "valign": "top", "text_wrap": True})

        # >>> XLS_EXECUTIVO_VISUAL_V2_HELPER
        # Padronização visual leve da Planilha Executiva.
        # Não altera cálculo, fórmulas, dados, nomes de abas ou validações.
        formatos_base_executivos = [
            fmt_title, fmt_subtitle, fmt_section, fmt_text, fmt_text_center,
            fmt_money, fmt_money_bold, fmt_pct, fmt_factor, fmt_red_money,
            fmt_red_text, fmt_note, fmt_note_wrap, fmt_text_wrap
        ]
        for _fmt in formatos_base_executivos:
            try:
                _fmt.set_font_name(fonte_executiva_xls)
                _fmt.set_font_size(11)
            except Exception:
                pass

        def _aplicar_visual_executivo_ws(ws, nome_aba):
            try:
                ws.hide_gridlines(2)
                ws.set_zoom(90)
            except Exception:
                pass

            try:
                if nome_aba == "CONFERENCIA":
                    ws.freeze_panes(4, 0)
                elif nome_aba in [
                    "RESUMO_FINANCEIRO", "VALOR_ATUALIZADO", "COMPOSICAO_VALOR_TOTAL",
                    "CICLOS", "FINANCEIRO_MENSAL", "FINANCEIRO_POR_CICLO",
                    "EXECUCAO_ATUALIZADA", "REMANESCENTES", "ADITIVOS",
                    "AUDITORIA", "VALORES_UNITARIOS_CICLO", "DELTA_POR_CICLO",
                    "MESES_SEM_EFEITO", "VALORES_POR_CICLO"
                ]:
                    ws.freeze_panes(1, 0)
            except Exception:
                pass

            try:
                if nome_aba == "VALOR_ATUALIZADO":
                    ws.set_column("A:A", 36)
                    ws.set_column("B:B", 24)
                    ws.set_column("C:C", 22)
                    ws.set_column("D:D", 88)
                elif nome_aba == "COMPOSICAO_VALOR_TOTAL":
                    ws.set_column("A:A", 44)
                    ws.set_column("B:B", 24)
                    ws.set_column("C:C", 22)
                    ws.set_column("D:D", 96)
                elif nome_aba == "CONFERENCIA":
                    ws.set_column("A:A", 38)
                    ws.set_column("B:B", 30)
                    ws.set_column("C:F", 24)
                    ws.set_column("G:G", 90)
                elif nome_aba == "RESUMO_FINANCEIRO":
                    ws.set_column("A:A", 42)
                    ws.set_column("B:B", 28)
                    ws.set_column("C:H", 24)
                elif nome_aba == "AUDITORIA":
                    ws.set_column("A:A", 40)
                    ws.set_column("B:B", 18)
                    ws.set_column("C:C", 24)
                    ws.set_column("D:D", 96)
                elif nome_aba in ["CICLOS", "FINANCEIRO_MENSAL", "FINANCEIRO_POR_CICLO", "EXECUCAO_ATUALIZADA"]:
                    ws.set_column("A:A", 18)
                    ws.set_column("B:Z", 24)
                elif nome_aba in ["REMANESCENTES", "ADITIVOS", "VALORES_UNITARIOS_CICLO", "VALORES_POR_CICLO"]:
                    ws.set_column("A:A", 24)
                    ws.set_column("B:Z", 24)
            except Exception:
                pass
        # <<< XLS_EXECUTIVO_VISUAL_V2_HELPER

        ciclo_fills = ["#FFFFFF", "#DDEBF7", "#E2F0D9", "#FFF2CC", "#EADCF8", "#E7E6E6", "#DDEBF7", "#E2F0D9"]
        def fmt_ciclo(ciclo, tipo="text", precluso=False, header=False):
            idx_cor = numero_ciclo(ciclo)
            bg = ciclo_fills[idx_cor % len(ciclo_fills)]
            base = {"border": 1, "bg_color": bg, "valign": "vcenter"}
            if header:
                base.update({"bold": True, "align": "center", "text_wrap": True})
            if precluso:
                base["font_color"] = "#C00000"
            if tipo == "money":
                base["num_format"] = 'R$ #,##0.00'
            elif tipo == "num":
                base["num_format"] = '#,##0.00'
            elif tipo == "pct":
                base["num_format"] = "0.00%"
            elif tipo == "factor":
                base["num_format"] = "0.00"
            return workbook.add_format(base)

        # ====================================================
        # Aba Valor Atualizado - Composição por Ciclo
        # ====================================================
        # >>> QUADRO_VALOR_ATUALIZADO_CICLOS_V1
        ws_va = workbook.add_worksheet("VALOR_ATUALIZADO")
        writer.sheets["VALOR_ATUALIZADO"] = ws_va
        ws_va.write(0, 0, "Valor Atualizado - Composição por Ciclo", fmt_title)
        ws_va.write(
            1,
            0,
            "Quadro de conferência da composição do Valor Total Atualizado. A execução e o saldo remanescente são apresentados separadamente por ciclo para facilitar a comparação com a planilha fiscal.",
            fmt_note_wrap,
        )

        headers_va = [
            "Ciclo",
            "Parcela",
            "Valor original/base",
            "Fator ou percentual",
            "Valor atualizado",
            "Origem / observação",
        ]
        for c_idx, titulo in enumerate(headers_va):
            ws_va.write(3, c_idx, titulo, fmt_subtitle)

        linha_va = 4
        soma_va = 0.0
        config_corte_va = resultado.get("config_ciclo_em_execucao", {}) or {}
        ciclo_corte_va = normalizar_ciclo(config_corte_va.get("ciclo", ""))
        corte_aplicado_va = bool(resultado.get("corte_operacional_aplicado", False))

        def _escrever_linha_valor_atualizado(ciclo, parcela, valor_base, fator, valor_atualizado, observacao):
            nonlocal linha_va, soma_va
            ciclo_norm = normalizar_ciclo(ciclo) or str(ciclo or "").strip() or "-"
            valor_base_num = numero_seguro(valor_base, 0.0)
            valor_atual_num = numero_seguro(valor_atualizado, 0.0)
            fator_num = numero_seguro(fator, 0.0)
            ws_va.write(linha_va, 0, ciclo_norm, fmt_ciclo(ciclo_norm, "text", False))
            ws_va.write(linha_va, 1, parcela, fmt_ciclo(ciclo_norm, "text", False))
            ws_va.write_number(linha_va, 2, valor_base_num, fmt_ciclo(ciclo_norm, "money", False))
            ws_va.write_number(linha_va, 3, fator_num, fmt_ciclo(ciclo_norm, "factor", False))
            ws_va.write_number(linha_va, 4, valor_atual_num, fmt_ciclo(ciclo_norm, "money", False))
            ws_va.write(linha_va, 5, observacao, fmt_ciclo(ciclo_norm, "text", False))
            soma_va += valor_atual_num
            linha_va += 1

        df_exec_quadro = limpar_nan_inf_df(resultado.get("df_execucao_atualizada", pd.DataFrame())).copy()
        if isinstance(df_exec_quadro, pd.DataFrame) and not df_exec_quadro.empty:
            if "Ciclo" in df_exec_quadro.columns:
                df_exec_quadro = df_exec_quadro.sort_values(by="Ciclo", key=lambda s: s.map(numero_ciclo), kind="stable")
            for _, row_exec in df_exec_quadro.iterrows():
                ciclo_exec = normalizar_ciclo(row_exec.get("Ciclo", "")) or str(row_exec.get("Ciclo", "")).strip()
                status_exec = normalizar_texto(row_exec.get("Status financeiro", ""))
                valor_base_exec = numero_seguro(row_exec.get("Valor executado original", 0.0), 0.0)
                valor_atual_exec = numero_seguro(row_exec.get("Valor executado atualizado", 0.0), 0.0)
                pct_exec = numero_seguro(row_exec.get("Percentual acumulado aplicado", 0.0), 0.0)
                fator_exec = 1.0 + pct_exec

                if ciclo_exec == "C0" and ("manual" in status_exec or "c0" in status_exec):
                    parcela_exec = "Execução financeira manual"
                    obs_exec = "Valor financeiro de C0 informado na aba CICLO_EM_EXECUCAO."
                elif corte_aplicado_va and ciclo_corte_va and ciclo_exec == ciclo_corte_va:
                    parcela_exec = "Execução financeira até o corte operacional"
                    competencia = config_corte_va.get("competencia_corte", config_corte_va.get("data_corte", ""))
                    obs_exec = f"Execução apurada pela base financeira até a competência de corte {competencia}."
                else:
                    parcela_exec = "Execução atualizada"
                    obs_exec = "Valor executado/consumido no ciclo, atualizado pelo fator aplicável."

                if abs(valor_base_exec) > 0.004 or abs(valor_atual_exec) > 0.004:
                    _escrever_linha_valor_atualizado(
                        ciclo_exec,
                        parcela_exec,
                        valor_base_exec,
                        fator_exec,
                        valor_atual_exec,
                        obs_exec,
                    )

        df_rem_quadro = limpar_nan_inf_df(resultado.get("df_remanescentes", pd.DataFrame())).copy()
        if isinstance(df_rem_quadro, pd.DataFrame) and not df_rem_quadro.empty:
            for _, row_rem in df_rem_quadro.iterrows():
                ciclo_rem_raw = str(row_rem.get("Ciclo", "")).strip()
                ciclo_rem = normalizar_ciclo(ciclo_rem_raw) or ciclo_rem_raw or "-"
                obs_rem_raw = str(row_rem.get("Observação", "")).strip()
                texto_rem = normalizar_texto(ciclo_rem_raw + " " + obs_rem_raw)
                valor_original_rem = numero_seguro(row_rem.get("Remanescente original", 0.0), 0.0)
                fator_rem = numero_seguro(row_rem.get("Fator aplicado", 1.0), 1.0)
                valor_atual_rem = numero_seguro(row_rem.get("Remanescente atualizado", 0.0), 0.0)

                if "corte_operacional" in texto_rem or "ciclo_em_execucao" in texto_rem:
                    parcela_rem = "Remanescente informado no corte operacional"
                    obs_rem = "Saldo futuro informado na aba CICLO_EM_EXECUCAO. Se já estava atualizado pela fiscalização, foi usado diretamente, sem nova atualização."
                else:
                    parcela_rem = "Remanescente itemizado automático"
                    obs_rem = "Saldo calculado automaticamente a partir dos itens/remanescentes informados."

                if abs(valor_original_rem) > 0.004 or abs(valor_atual_rem) > 0.004:
                    _escrever_linha_valor_atualizado(
                        ciclo_rem,
                        parcela_rem,
                        valor_original_rem,
                        fator_rem,
                        valor_atual_rem,
                        obs_rem,
                    )

        valor_total_va = numero_seguro(resultado.get("valor_atualizado_contrato", 0.0), 0.0)
        linha_va += 1
        ws_va.write(linha_va, 0, "TOTAL", fmt_subtitle)
        ws_va.write(linha_va, 1, "Valor Total Atualizado do Contrato", fmt_subtitle)
        ws_va.write(linha_va, 2, "", fmt_subtitle)
        ws_va.write(linha_va, 3, "", fmt_subtitle)
        ws_va.write_number(linha_va, 4, valor_total_va, fmt_money_bold)
        ws_va.write(linha_va, 5, "Resultado consolidado usado pelo cl8us.", fmt_text_wrap)

        diferenca_va = round(valor_total_va - soma_va, 2)
        if abs(diferenca_va) > 0.01:
            linha_va += 1
            ws_va.write(linha_va, 0, "CONTROLE", fmt_red_text)
            ws_va.write(linha_va, 1, "Diferença de conferência", fmt_red_text)
            ws_va.write(linha_va, 2, "", fmt_red_text)
            ws_va.write(linha_va, 3, "", fmt_red_text)
            ws_va.write_number(linha_va, 4, diferenca_va, fmt_red_money)
            ws_va.write(linha_va, 5, "Diferença entre o total e a soma das linhas acima; revisar apenas se houver valor material.", fmt_red_text)

        ws_va.set_column("A:A", 14)
        ws_va.set_column("B:B", 42)
        ws_va.set_column("C:E", 24)
        ws_va.set_column("F:F", 90)
        ws_va.set_row(1, 42)
        ws_va.freeze_panes(4, 0)
        # <<< QUADRO_VALOR_ATUALIZADO_CICLOS_V1

        # ====================================================
        # Aba 0 - Conferência Executiva
        # ====================================================
        ws_conf = workbook.add_worksheet("CONFERENCIA")
        writer.sheets["CONFERENCIA"] = ws_conf
        ws_conf.write(0, 0, "Conferência Executiva", fmt_title)
        ws_conf.write(1, 0, "Quadro executivo da apuração, com destaque obrigatório para competências sem efeito financeiro.", fmt_note_wrap)
        modo_consumo_conferencia = resultado.get("modo_apuracao") == "Consumo por Itens/Ciclo"
        label_retroativo_conferencia = "Retroativo (itens consumidos/ciclo)" if modo_consumo_conferencia else "Valor represado a pagar"
        valor_retroativo_conferencia = (
            resultado.get("valor_retroativo_consumo_itens_ciclo", 0.0)
            if modo_consumo_conferencia
            else resultado.get("valor_represado_a_pagar", 0.0)
        )
        conf_linhas = [
            ("Modo de apuração", resultado.get("modo_apuracao", "Completo"), "text"),
            ("Valor Total Atualizado do Contrato", resultado.get("valor_atualizado_contrato", 0.0), "money_bold"),
            (label_retroativo_conferencia, valor_retroativo_conferencia, "money"),
            ("Meses sem efeitos financeiros", int(resultado.get("quantidade_meses_sem_efeito_financeiro", 0) or 0), "text"),
            ("Valor total sem efeito financeiro", resultado.get("valor_total_sem_efeito_financeiro", 0.0), "money"),
            ("Observação", "Competências sem efeito financeiro são demonstradas para memória e transparência, mas o delta teórico não compõe o retroativo a pagar.", "text"),
        ]
        ws_conf.write(3, 0, "Indicador", fmt_subtitle)
        ws_conf.write(3, 1, "Valor", fmt_subtitle)
        for r, (label, value, tipo) in enumerate(conf_linhas, start=4):
            ws_conf.write(r, 0, label, fmt_text)
            if tipo == "money":
                ws_conf.write_number(r, 1, numero_seguro(value, 0.0), fmt_money)
            elif tipo == "money_bold":
                ws_conf.write_number(r, 1, numero_seguro(value, 0.0), fmt_money_bold)
            else:
                ws_conf.write(r, 1, str(value), fmt_text_wrap if label == "Observação" else fmt_text)
        df_ef_conf = limpar_nan_inf_df(resultado.get("df_meses_sem_efeito_financeiro", pd.DataFrame())).copy()
        start_ef = 12
        ws_conf.write(start_ef, 0, "Detalhamento dos meses sem efeito financeiro", fmt_section)
        if not df_ef_conf.empty:
            for c_idx, col in enumerate(df_ef_conf.columns):
                ws_conf.write(start_ef + 1, c_idx, col, fmt_subtitle)
            for r_idx, (_, row) in enumerate(df_ef_conf.iterrows(), start=start_ef + 2):
                for c_idx, col in enumerate(df_ef_conf.columns):
                    value = row.get(col, "")
                    if col in ["Valor base", "Valor teórico se aplicado", "Delta não devido"]:
                        ws_conf.write_number(r_idx, c_idx, numero_seguro(value, 0.0), fmt_red_money if col == "Delta não devido" else fmt_money)
                    elif col == "Fator teórico":
                        ws_conf.write_number(r_idx, c_idx, numero_seguro(value, 1.0), fmt_factor)
                    else:
                        ws_conf.write(r_idx, c_idx, texto_seguro(value, ""), fmt_text_wrap)
        else:
            ws_conf.write(start_ef + 1, 0, "Não foram identificadas competências sem efeito financeiro.", fmt_text)
        ws_conf.set_column("A:A", 36)
        ws_conf.set_column("B:B", 32)
        ws_conf.set_column("C:F", 24)
        ws_conf.set_column("G:G", 90)
        ws_conf.set_row(1, 42)

        # ====================================================
        # Aba 1 - Resumo Financeiro
        # ====================================================
        ws = workbook.add_worksheet("RESUMO_FINANCEIRO")
        writer.sheets["RESUMO_FINANCEIRO"] = ws
        ws.set_column("A:A", 38)
        ws.set_column("B:B", 24)
        ws.set_column("C:H", 22)
        ws.write("A1", "Planilha Executiva da Análise", fmt_title)
        ws.write("A2", "Sistema de Apoio à Gestão de Contratos", fmt_note)

        resumo = [
            ("Data de processamento", resultado.get("data_processamento", ""), "text"),
            ("Índice utilizado", resultado.get("indice", "Não informado"), "text"),
            ("Valor original do contrato", resultado.get("valor_original_contrato", 0.0), "money"),
            ("Valor formalizado antes desta análise", resultado.get("valor_formalizado_anterior", resultado.get("valor_original_contrato", 0.0)), "money"),
            ("Valor pago efetivo", resultado.get("valor_pago_efetivo", 0.0), "money"),
            ("Valor teórico calculado", resultado.get("valor_teorico_calculado", 0.0), "money"),
            ("Valor represado a pagar", resultado.get("valor_represado_a_pagar", 0.0), "money"),
            ("Meses sem efeitos financeiros", resultado.get("quantidade_meses_sem_efeito_financeiro", 0), "text"),
            ("Valor total sem efeito financeiro", resultado.get("valor_total_sem_efeito_financeiro", 0.0), "money"),
            ("Saldo remanescente atualizado", resultado.get("remanescente_reajustado", 0.0), "money"),
            ("Aditivos/Supressões registrados (informativo)", resultado.get("total_aditivos_atualizados", 0.0), "money"),
            ("Aditivos/Supressões informativos", resultado.get("total_aditivos_informativos", 0.0), "money"),
            ("Reajuste acumulado", resultado.get("variacao_acumulada", resultado.get("fator_acumulado", 1.0) - 1), "pct"),
            ("Valor Total Atualizado do Contrato", resultado.get("valor_atualizado_contrato", 0.0), "money_bold"),
        ]
        ws.write(3, 0, "Indicador", fmt_subtitle)
        ws.write(3, 1, "Valor", fmt_subtitle)
        for r, (label, value, tipo) in enumerate(resumo, start=4):
            ws.write(r, 0, label, fmt_text)
            if tipo == "money":
                ws.write_number(r, 1, numero_seguro(value, 0.0), fmt_money)
            elif tipo == "money_bold":
                ws.write_number(r, 1, numero_seguro(value, 0.0), fmt_money_bold)
            elif tipo == "pct":
                ws.write_number(r, 1, numero_seguro(value, 0.0), fmt_pct)
            else:
                ws.write(r, 1, value, fmt_text)

        row_base = 19
        ws.write(row_base, 0, "Financeiro por Ciclo / Deltas", fmt_section)
        df_delta = limpar_nan_inf_df(resultado.get("df_delta_por_ciclo", pd.DataFrame())).copy()
        if df_delta.empty:
            df_delta = limpar_nan_inf_df(resultado.get("df_financeiro_por_ciclo", pd.DataFrame())).copy()
        if not df_delta.empty:
            cols = [c for c in ["Ciclo", "Situação", "Tratamento financeiro", "Valor pago efetivo", "Valor teórico calculado", "Delta do ciclo", "Fator aplicado ao retroativo"] if c in df_delta.columns]
            df_delta = df_delta[cols].copy()
            for c_idx, col in enumerate(df_delta.columns):
                ws.write(row_base + 1, c_idx, col, fmt_subtitle)
            for r_idx, (_, row) in enumerate(df_delta.iterrows(), start=row_base + 2):
                ciclo = row.get("Ciclo", "")
                precluso = "preclus" in normalizar_texto(row.get("Situação", ""))
                for c_idx, col in enumerate(df_delta.columns):
                    value = row.get(col, "")
                    if col in ["Valor pago efetivo", "Valor teórico calculado", "Delta do ciclo"]:
                        ws.write_number(r_idx, c_idx, numero_seguro(value, 0.0), fmt_ciclo(ciclo, "money", precluso))
                    elif col == "Fator aplicado ao retroativo":
                        ws.write_number(r_idx, c_idx, numero_seguro(value, 1.0), fmt_ciclo(ciclo, "factor", precluso))
                    else:
                        ws.write(r_idx, c_idx, "" if pd.isna(value) else value, fmt_ciclo(ciclo, "text", precluso))

        # ====================================================
        # Aba 2 - Detalhamento de Itens
        # ====================================================
        ws_i = workbook.add_worksheet("DETALHAMENTO_ITENS")
        writer.sheets["DETALHAMENTO_ITENS"] = ws_i
        ws_i.write(0, 0, "Detalhamento de Itens por Ciclo", fmt_title)
        df_vu = limpar_nan_inf_df(resultado.get("df_valores_unitarios_ciclo", pd.DataFrame())).copy()
        if not df_vu.empty:
            ciclos_ordenados = sorted(df_vu["Ciclo"].dropna().astype(str).unique(), key=numero_ciclo)
            itens_ordenados = list(df_vu["Item"].dropna().drop_duplicates())
            ws_i.write(2, 0, "Item", fmt_subtitle)
            col_atual = 1
            for ciclo in ciclos_ordenados:
                ws_i.merge_range(1, col_atual, 1, col_atual + 2, ciclo, fmt_ciclo(ciclo, "text", False, header=True))
                ws_i.write(2, col_atual, "Quantidade", fmt_ciclo(ciclo, "text", False, header=True))
                ws_i.write(2, col_atual + 1, "Valor unitário", fmt_ciclo(ciclo, "text", False, header=True))
                ws_i.write(2, col_atual + 2, "Valor total", fmt_ciclo(ciclo, "text", False, header=True))
                col_atual += 3
            ws_i.set_column(0, 0, 16)
            ws_i.set_column(1, max(1, col_atual), 18)
            ultima_linha_item = 2
            for r_idx, item in enumerate(itens_ordenados, start=3):
                ultima_linha_item = r_idx
                ws_i.write(r_idx, 0, item, fmt_text)
                col_atual = 1
                for ciclo in ciclos_ordenados:
                    linha = df_vu[(df_vu["Item"].astype(str) == str(item)) & (df_vu["Ciclo"].astype(str) == str(ciclo))]
                    if linha.empty:
                        qtd = vu = total = 0.0
                        precluso = False
                    else:
                        rec = linha.iloc[0]
                        qtd = numero_seguro(rec.get("Quantidade", 0), 0.0)
                        vu = numero_seguro(rec.get("Valor unitário", 0), 0.0)
                        total = numero_seguro(rec.get("Total R$", 0), 0.0)
                        precluso = bool(rec.get("Ciclo precluso", False))
                    ws_i.write_number(r_idx, col_atual, qtd, fmt_ciclo(ciclo, "num", precluso))
                    ws_i.write_number(r_idx, col_atual + 1, vu, fmt_ciclo(ciclo, "money", precluso))
                    ws_i.write_number(r_idx, col_atual + 2, total, fmt_ciclo(ciclo, "money", precluso))
                    col_atual += 3

            # Linha de total por ciclo, logo abaixo dos itens, somando apenas as colunas "Valor total".
            total_row = ultima_linha_item + 1
            ws_i.write(total_row, 0, "TOTAL", fmt_subtitle)
            col_atual = 1
            for ciclo in ciclos_ordenados:
                primeira_linha_excel = 4
                ultima_linha_excel = ultima_linha_item + 1
                col_total_excel = col_atual + 2
                if col_total_excel < 26:
                    col_letter = chr(ord('A') + col_total_excel)
                    formula = f"=ROUND(SUM({col_letter}{primeira_linha_excel}:{col_letter}{ultima_linha_excel}),2)"
                    ws_i.write_blank(total_row, col_atual, None, fmt_ciclo(ciclo, "text", False))
                    ws_i.write_blank(total_row, col_atual + 1, None, fmt_ciclo(ciclo, "text", False))
                    ws_i.write_formula(total_row, col_atual + 2, formula, fmt_ciclo(ciclo, "money", False, header=True))
                else:
                    total_valor = df_vu[df_vu["Ciclo"].astype(str) == str(ciclo)]["Total R$"].apply(numero_seguro).sum()
                    ws_i.write_blank(total_row, col_atual, None, fmt_ciclo(ciclo, "text", False))
                    ws_i.write_blank(total_row, col_atual + 1, None, fmt_ciclo(ciclo, "text", False))
                    ws_i.write_number(total_row, col_atual + 2, total_valor, fmt_ciclo(ciclo, "money", False, header=True))
                col_atual += 3
        else:
            ws_i.write(2, 0, "Não há dados de itens disponíveis.", fmt_text)

        # ====================================================
        # Aba 3 - Aditivos e Supressões
        # ====================================================
        ws_a = workbook.add_worksheet("ADITIVOS_CONSOLIDADOS")
        writer.sheets["ADITIVOS_CONSOLIDADOS"] = ws_a
        ws_a.write(0, 0, "Aditivos e Supressões", fmt_title)
        df_ad = limpar_nan_inf_df(resultado.get("df_aditivos_executivo", resultado.get("df_aditivos", pd.DataFrame()))).copy()
        if not df_ad.empty:
            cols = [c for c in ["Aditivo", "Ciclo/Marco", "Tipo de alteração", "Tratamento do aditivo", "Quantidade de linhas", "Valor do aditivo na assinatura", "Fator aplicado", "Valor do aditivo reajustado", "Computa no Valor Global"] if c in df_ad.columns]
            df_ad = df_ad[cols].copy()
            if "Computa no Valor Global" in df_ad.columns:
                df_ad = df_ad.rename(columns={"Computa no Valor Global": "Marcado como computável no arquivo"})
            for c_idx, col in enumerate(df_ad.columns):
                ws_a.write(2, c_idx, col, fmt_subtitle)
            for r_idx, (_, row) in enumerate(df_ad.iterrows(), start=3):
                tipo_norm = normalizar_texto(row.get("Tipo de alteração", ""))
                vermelho = tipo_norm.startswith("decr") or tipo_norm.startswith("supres")
                for c_idx, col in enumerate(df_ad.columns):
                    value = row.get(col, "")
                    fmt_base_text = fmt_red_text if vermelho else fmt_text
                    fmt_base_money = fmt_red_money if vermelho else fmt_money
                    if col in ["Valor original da alteração", "Valor atualizado da alteração", "Valor do aditivo na assinatura", "Valor do aditivo reajustado"]:
                        ws_a.write_number(r_idx, c_idx, numero_seguro(value, 0.0), fmt_base_money)
                    elif col == "Fator aplicado":
                        ws_a.write_number(r_idx, c_idx, numero_seguro(value, 1.0), fmt_factor)
                    elif col == "Data do aditivo":
                        ws_a.write(r_idx, c_idx, formatar_data_br(value), fmt_base_text)
                    elif col in ["Computa no Valor Global", "Marcado como computável no arquivo"]:
                        ws_a.write(r_idx, c_idx, "Sim" if bool(value) else "Não", fmt_base_text)
                    else:
                        ws_a.write(r_idx, c_idx, "" if pd.isna(value) else value, fmt_base_text)
            for idx, col in enumerate(df_ad.columns):
                ws_a.set_column(idx, idx, max(14, min(34, len(str(col)) + 4)))
        else:
            ws_a.write(2, 0, "Não há aditivos ou supressões informados.", fmt_text)

        # ====================================================
        # Aba 4 - Composição do Valor Total Atualizado
        # ====================================================
        ws_c = workbook.add_worksheet("COMPOSICAO_VALOR_TOTAL")
        writer.sheets["COMPOSICAO_VALOR_TOTAL"] = ws_c
        ws_c.write(0, 0, "Composição do Valor Total Atualizado do Contrato", fmt_title)
        ws_c.write(1, 0, "Execução atualizada por ciclo + saldo remanescente atualizado + aditivos/supressões computáveis quando indicados como parcela da análise atual.", fmt_note_wrap)
        df_comp = limpar_nan_inf_df(resultado.get("df_composicao_valor_total", pd.DataFrame())).copy()
        if not df_comp.empty:
            for c_idx, col in enumerate(df_comp.columns):
                ws_c.write(3, c_idx, col, fmt_subtitle)
            for r_idx, (_, row) in enumerate(df_comp.iterrows(), start=4):
                componente = str(row.get("Componente", ""))
                for c_idx, col in enumerate(df_comp.columns):
                    value = row.get(col, "")
                    if col == "Valor":
                        fmt_valor = fmt_money_bold if "Valor Total Atualizado" in componente else fmt_money
                        ws_c.write_number(r_idx, c_idx, numero_seguro(value, 0.0), fmt_valor)
                    else:
                        ws_c.write(r_idx, c_idx, texto_seguro(value, ""), fmt_text_wrap)
            ws_c.set_column("A:A", 44)
            ws_c.set_column("B:B", 24)
            ws_c.set_column("C:C", 22)
            ws_c.set_column("D:D", 92)
            ws_c.set_row(1, 42)
        else:
            ws_c.write(3, 0, "Não há composição disponível.", fmt_text)

        # ====================================================
        # Aba - Retroativo por itens
        # ====================================================
        ws_ri = workbook.add_worksheet("RETROATIVO_ITENS")
        writer.sheets["RETROATIVO_ITENS"] = ws_ri
        modo_consumo_xls = resultado.get("modo_apuracao") == "Consumo por Itens/Ciclo"
        titulo_retro_xls = "Retroativo (itens consumidos/ciclo)" if modo_consumo_xls else "Retroativo estimado por itens/estoque"
        nota_retro_xls = "Apuração por consumo itemizado por ciclo, baseada nas quantidades consumidas informadas pela fiscalização." if modo_consumo_xls else "Apuração estimativa usada quando não há base de execução mensal por competência. Não substitui retroativo financeiro definitivo."
        ws_ri.write(0, 0, titulo_retro_xls, fmt_title)
        ws_ri.write(1, 0, nota_retro_xls, fmt_note_wrap)
        df_ri = limpar_nan_inf_df(resultado.get("df_retroativo_itemizado_por_ciclo", pd.DataFrame()) if modo_consumo_xls else resultado.get("df_retroativo_estimado_itens_estoque", pd.DataFrame())).copy()
        if not df_ri.empty:
            for c_idx, col in enumerate(df_ri.columns):
                ws_ri.write(3, c_idx, col, fmt_subtitle)
            for r_idx, (_, row) in enumerate(df_ri.iterrows(), start=4):
                for c_idx, col in enumerate(df_ri.columns):
                    value = row.get(col, "")
                    if col in ["Valor executado original", "Valor executado atualizado", "Retroativo estimado por itens/estoque", "Valor original consumido", "Valor atualizado consumido", "Retroativo"]:
                        ws_ri.write_number(r_idx, c_idx, numero_seguro(value, 0.0), fmt_money)
                    elif col in ["Fator acumulado", "Fator aplicado"]:
                        ws_ri.write_number(r_idx, c_idx, numero_seguro(value, 1.0), fmt_factor)
                    else:
                        ws_ri.write(r_idx, c_idx, texto_seguro(value, ""), fmt_text_wrap)
            total_row = 4 + len(df_ri)
            ws_ri.write(total_row, 0, "TOTAL", fmt_subtitle)
            if modo_consumo_xls and "Retroativo" in df_ri.columns:
                col_total = list(df_ri.columns).index("Retroativo")
                ws_ri.write_number(total_row, col_total, numero_seguro(resultado.get("valor_retroativo_consumo_itens_ciclo", 0.0), 0.0), fmt_money_bold)
            elif "Retroativo estimado por itens/estoque" in df_ri.columns:
                col_total = list(df_ri.columns).index("Retroativo estimado por itens/estoque")
                ws_ri.write_number(total_row, col_total, numero_seguro(resultado.get("valor_retroativo_estimado_itens_estoque", 0.0), 0.0), fmt_money_bold)
        else:
            ws_ri.write(3, 0, "Não há retroativo estimado por itens/estoque disponível.", fmt_text)
        ws_ri.set_column("A:A", 18)
        ws_ri.set_column("B:D", 28)
        ws_ri.set_column("E:F", 34)
        ws_ri.set_column("G:G", 38)
        ws_ri.set_column("H:H", 44)
        ws_ri.set_row(1, 44)

        # ====================================================
        # Aba 5 - Contexto do Contrato
        # ====================================================
        ws_ctx = workbook.add_worksheet("CONTEXTO_CONTRATO")
        writer.sheets["CONTEXTO_CONTRATO"] = ws_ctx
        ws_ctx.write(0, 0, "Contexto do Contrato", fmt_title)
        ws_ctx.write(1, 0, "Memória formal anterior. Os valores aqui registrados não são somados automaticamente ao Valor Total Atualizado. Quando aplicáveis, seus efeitos devem estar refletidos nas medições/execução financeira, nas quantidades dos itens ou nos saldos remanescentes informados.", fmt_note_wrap)
        contexto = resultado.get("contexto_contratual_anterior", {}) or {}
        linhas_ctx = [
            ["Valor original do contrato (contexto)", contexto.get("valor_original_contrato", "")],
            ["Valor formalizado antes desta análise", contexto.get("valor_formalizado_anterior", "")],
            ["Último ciclo já concedido/formalizado", contexto.get("ultimo_ciclo_concedido", "")],
            ["Observação sobre histórico anterior", contexto.get("observacao_historico", "")],
        ]
        ws_ctx.write(3, 0, "Campo", fmt_subtitle)
        ws_ctx.write(3, 1, "Valor", fmt_subtitle)
        for r_idx, linha in enumerate(linhas_ctx, start=4):
            ws_ctx.write(r_idx, 0, linha[0], fmt_text)
            if isinstance(linha[1], (int, float)):
                ws_ctx.write_number(r_idx, 1, numero_seguro(linha[1], 0.0), fmt_money)
            else:
                ws_ctx.write(r_idx, 1, texto_seguro(linha[1], "Não"), fmt_text_wrap)
        eventos_ctx = contexto.get("eventos_historicos_anteriores", []) if isinstance(contexto, dict) else []
        start_evt = 10
        ws_ctx.write(start_evt, 0, "Eventos históricos anteriores", fmt_section)
        if eventos_ctx:
            headers_evt = ["Tipo de evento", "Ciclo", "Data", "Valor formalizado/impacto", "Incorporado ao valor formalizado?", "Observação"]
            for c_idx, h in enumerate(headers_evt):
                ws_ctx.write(start_evt + 1, c_idx, h, fmt_subtitle)
            for r_idx, evento in enumerate(eventos_ctx, start=start_evt + 2):
                for c_idx, h in enumerate(headers_evt):
                    valor = evento.get(h, "") if isinstance(evento, dict) else ""
                    if h == "Valor formalizado/impacto" and str(valor).strip() != "":
                        ws_ctx.write_number(r_idx, c_idx, numero_br(valor), fmt_money)
                    elif h == "Data":
                        ws_ctx.write(r_idx, c_idx, formatar_data_br(valor), fmt_text)
                    else:
                        ws_ctx.write(r_idx, c_idx, texto_seguro(valor, "Não"), fmt_text_wrap)
        else:
            ws_ctx.write(start_evt + 1, 0, "Sem eventos históricos anteriores informados.", fmt_text)
        ws_ctx.set_column("A:A", 38)
        ws_ctx.set_column("B:B", 40)
        ws_ctx.set_column("C:C", 18)
        ws_ctx.set_column("D:D", 24)
        ws_ctx.set_column("E:E", 28)
        ws_ctx.set_column("F:F", 80)
        ws_ctx.set_row(1, 50)

        # ====================================================
        # Aba 6 - Efeitos Financeiros
        # ====================================================
        ws_ef = workbook.add_worksheet("EFEITOS_FINANCEIROS")
        writer.sheets["EFEITOS_FINANCEIROS"] = ws_ef
        ws_ef.write(0, 0, "Efeitos Financeiros — Competências sem efeito", fmt_title)
        ws_ef.write(1, 0, "Regra: competências anteriores ao início dos efeitos financeiros do pedido são demonstradas para memória, mas o delta teórico não compõe o retroativo a pagar.", fmt_note_wrap)
        df_ef = limpar_nan_inf_df(resultado.get("df_meses_sem_efeito_financeiro", pd.DataFrame())).copy()
        if not df_ef.empty:
            for c_idx, col in enumerate(df_ef.columns):
                ws_ef.write(3, c_idx, col, fmt_subtitle)
            for r_idx, (_, row) in enumerate(df_ef.iterrows(), start=4):
                for c_idx, col in enumerate(df_ef.columns):
                    value = row.get(col, "")
                    if col in ["Valor base", "Valor teórico se aplicado", "Delta não devido"]:
                        ws_ef.write_number(r_idx, c_idx, numero_seguro(value, 0.0), fmt_red_money if col == "Delta não devido" else fmt_money)
                    elif col == "Fator teórico":
                        ws_ef.write_number(r_idx, c_idx, numero_seguro(value, 1.0), fmt_factor)
                    else:
                        ws_ef.write(r_idx, c_idx, texto_seguro(value, ""), fmt_text_wrap)
            total_row = len(df_ef) + 5
            ws_ef.write(total_row, 0, "TOTAL SEM EFEITO FINANCEIRO", fmt_subtitle)
            try:
                col_delta_total = list(df_ef.columns).index("Delta não devido")
            except Exception:
                col_delta_total = 5
            ws_ef.write_number(total_row, col_delta_total, numero_seguro(resultado.get("valor_total_sem_efeito_financeiro", 0.0), 0.0), fmt_red_money)
        else:
            ws_ef.write(3, 0, "Não foram identificadas competências sem efeito financeiro.", fmt_text)
        ws_ef.set_column("A:A", 12)
        ws_ef.set_column("B:B", 16)
        ws_ef.set_column("C:C", 20)
        ws_ef.set_column("D:D", 18)
        ws_ef.set_column("E:F", 24)
        ws_ef.set_column("G:G", 90)
        ws_ef.set_row(1, 42)

        # ====================================================
        # Aba 7 - Auditoria de Consistência
        # ====================================================
        ws_aud = workbook.add_worksheet("AUDITORIA_CONSISTENCIA")
        writer.sheets["AUDITORIA_CONSISTENCIA"] = ws_aud
        ws_aud.write(0, 0, "Auditoria de Consistência", fmt_title)
        ws_aud.write(1, 0, "Checklist automático de fechamento da análise. Status ATENÇÃO exige revisão manual antes de instrução final.", fmt_note)
        df_aud = limpar_nan_inf_df(resultado.get("df_auditoria_consistencia", pd.DataFrame())).copy()
        if not df_aud.empty:
            for c_idx, col in enumerate(df_aud.columns):
                ws_aud.write(3, c_idx, col, fmt_subtitle)
            for r_idx, (_, row) in enumerate(df_aud.iterrows(), start=4):
                for c_idx, col in enumerate(df_aud.columns):
                    value = row.get(col, "")
                    if col == "Diferença/Valor" and isinstance(value, (int, float)) and not pd.isna(value):
                        ws_aud.write_number(r_idx, c_idx, numero_seguro(value, 0.0), fmt_money)
                    else:
                        ws_aud.write(r_idx, c_idx, "" if pd.isna(value) else value, fmt_text)
            for idx, col in enumerate(df_aud.columns):
                largura = 24 if col == "Diferença/Valor" else max(18, min(60, len(str(col)) + 8))
                ws_aud.set_column(idx, idx, largura)
        else:
            ws_aud.write(3, 0, "Auditoria não disponível.", fmt_text)

        # >>> XLS_EXECUTIVO_VISUAL_V2_APLICAR
        # Aplicação final nas abas geradas pela Planilha Executiva.
        for _nome_ws, _ws in writer.sheets.items():
            _aplicar_visual_executivo_ws(_ws, _nome_ws)
        # <<< XLS_EXECUTIVO_VISUAL_V2_APLICAR

    output.seek(0)
    return output.getvalue()

def calcular_execucao_por_diferenca(itens, colunas_remanescentes, ciclos):
    if not colunas_remanescentes:
        return pd.DataFrame()

    fatores = mapa_fatores(ciclos)
    rem_por_ciclo = sorted(
        [(ciclo_da_coluna_remanescente(col, fallback="C1"), col) for col in colunas_remanescentes],
        key=lambda x: numero_ciclo(x[0])
    )
    linhas = []

    primeiro_ciclo, primeira_col = rem_por_ciclo[0]
    qtd_consumida_c0 = itens["Quantidade contratada"] - itens[primeira_col].apply(numero_br)
    qtd_consumida_c0 = qtd_consumida_c0.clip(lower=0)
    valor_c0 = qtd_consumida_c0 * itens["Valor unitário original"]
    linhas.append({
        "Ciclo": "C0",
        "Status financeiro": "Base",
        "Valor executado original": float(valor_c0.sum()),
        "Percentual acumulado aplicado": 0.0,
        "Valor executado atualizado": float(valor_c0.sum()),
    })

    for idx in range(len(rem_por_ciclo) - 1):
        ciclo_atual, col_atual = rem_por_ciclo[idx]
        _, col_proxima = rem_por_ciclo[idx + 1]
        qtd_consumida = itens[col_atual].apply(numero_br) - itens[col_proxima].apply(numero_br)
        qtd_consumida = qtd_consumida.clip(lower=0)
        valor_original = qtd_consumida * itens["Valor unitário original"]
        info = fatores.get(ciclo_atual, {})
        fator = float(info.get("fator_acumulado_efetivo", 1.0) or 1.0)
        linhas.append({
            "Ciclo": ciclo_atual,
            "Status financeiro": info.get("tratamento", info.get("situacao", "")),
            "Valor executado original": float(valor_original.sum()),
            "Percentual acumulado aplicado": fator - 1,
            "Valor executado atualizado": float((valor_original * fator).sum()),
        })

    # Garante visibilidade de todos os ciclos identificados, inclusive o ciclo em execução
    # quando ainda não houver coluna posterior de remanescente para apurar consumo por diferença.
    ciclos_existentes = {str(l.get("Ciclo", "")) for l in linhas}
    if isinstance(ciclos, pd.DataFrame) and not ciclos.empty:
        fatores = mapa_fatores(ciclos)
        for _, ciclo_row in ciclos.iterrows():
            ciclo_nome = normalizar_ciclo(ciclo_row.get("Ciclo", ""))
            if ciclo_nome and ciclo_nome not in ciclos_existentes:
                info = fatores.get(ciclo_nome, {})
                fator = float(info.get("fator_acumulado_efetivo", 1.0) or 1.0)
                linhas.append({
                    "Ciclo": ciclo_nome,
                    "Status financeiro": info.get("tratamento", info.get("situacao", "")) or "A apurar",
                    "Valor executado original": 0.0,
                    "Percentual acumulado aplicado": fator - 1,
                    "Valor executado atualizado": 0.0,
                })

    df = pd.DataFrame(linhas)
    if not df.empty:
        df = df.sort_values(by="Ciclo", key=lambda s: s.map(numero_ciclo)).reset_index(drop=True)
    return df



def montar_retroativo_estimado_por_itens(df_execucao):
    """Calcula o Retroativo estimado por itens/estoque no Modo Reduzido.

    Conceito: quando não há base de execução mensal por competência, o sistema
    não calcula retroativo financeiro definitivo. Ainda assim, é possível
    demonstrar uma estimativa a partir da execução física apurada por diferença
    de estoque/remanescentes.

    Retroativo estimado por itens/estoque = valor executado atualizado -
    valor executado original, por ciclo, excluindo C0/base.
    """
    colunas = [
        "Ciclo", "Valor executado original", "Valor executado atualizado",
        "Retroativo estimado por itens/estoque", "Natureza da apuração", "Observação"
    ]
    if df_execucao is None or not isinstance(df_execucao, pd.DataFrame) or df_execucao.empty:
        return pd.DataFrame(columns=colunas)
    df = df_execucao.copy()
    if "Ciclo" not in df.columns:
        return pd.DataFrame(columns=colunas)
    df = df[~df["Ciclo"].astype(str).str.upper().eq("C0")].copy()
    if df.empty:
        return pd.DataFrame(columns=colunas)
    for col in ["Valor executado original", "Valor executado atualizado"]:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = df[col].apply(numero_seguro)
    df["Retroativo estimado por itens/estoque"] = (
        df["Valor executado atualizado"] - df["Valor executado original"]
    ).round(2)
    df["Natureza da apuração"] = "Estimativa por itens/estoque"
    df["Observação"] = (
        "Estimativa calculada pela diferença entre o valor executado atualizado e "
        "o valor executado original, com base nos itens/remanescentes informados. "
        "Não substitui a base de execução mensal por competência."
    )
    return df[colunas].reset_index(drop=True)

def _marcos_ciclos_para_aditivos(ciclos):
    """Retorna marcos de ciclo com início/fim em datetime para enquadrar aditivos.

    Regra: Data-base do ciclo <= Data do aditivo <= fim do ciclo. Quando o fim não
    estiver disponível, usa a véspera da Data-base do ciclo seguinte; para o último
    ciclo, considera janela aberta.
    """
    if ciclos is None or ciclos.empty:
        return []

    marcos = []
    for _, row in ciclos.iterrows():
        ciclo = normalizar_ciclo(row.get("Ciclo", ""))
        inicio = pd.to_datetime(row.get("Data-base"), dayfirst=True, errors="coerce")
        fim = pd.to_datetime(row.get("Fim financeiro"), dayfirst=True, errors="coerce")
        if pd.notna(inicio) and ciclo:
            marcos.append({"ciclo": ciclo, "inicio": inicio.normalize(), "fim": fim.normalize() if pd.notna(fim) else pd.NaT})

    marcos.sort(key=lambda x: x["inicio"])
    for idx, marco in enumerate(marcos):
        if pd.isna(marco["fim"]):
            if idx + 1 < len(marcos):
                marco["fim"] = marcos[idx + 1]["inicio"] - pd.Timedelta(days=1)
            else:
                marco["fim"] = pd.Timestamp.max.normalize()
    return marcos


def inferir_ciclo_por_data(data_aditivo, ciclos):
    """Infere o Ciclo/Marco do aditivo por comparação de datas, exclusivamente em Python.

    Regra obrigatória:
    - se Data_Aditivo for anterior à primeira Data-base, retorna C0;
    - se Data_Aditivo for válida, retorna o último ciclo cuja Data-base seja menor ou igual à Data do Aditivo;
    - se a data estiver vazia/inválida ou não houver marcos de ciclo, retorna "Fora de Ciclo".
    """
    data = pd.to_datetime(data_aditivo, dayfirst=True, errors="coerce")
    if pd.isna(data):
        return "Fora de Ciclo"
    data = data.normalize()

    marcos = _marcos_ciclos_para_aditivos(ciclos)
    if not marcos:
        return "Fora de Ciclo"

    marcos_validos = [m for m in marcos if pd.notna(m.get("inicio"))]
    if not marcos_validos:
        return "Fora de Ciclo"

    marcos_validos = sorted(marcos_validos, key=lambda x: x["inicio"])
    if data < marcos_validos[0]["inicio"]:
        return "C0"

    ciclo_encontrado = "Fora de Ciclo"
    for marco in marcos_validos:
        if data >= marco["inicio"]:
            ciclo_encontrado = marco["ciclo"]
        else:
            break
    return ciclo_encontrado

def ler_aditivos(bytes_arquivo, xls, ciclos):
    """Lê aditivos/supressões com proteção contra layout antigo/misalinhado."""
    fatores = mapa_fatores(ciclos)
    linhas = []

    def _eh_sim_nao(valor):
        return str(valor or "").strip().upper() in ["SIM", "S", "NÃO", "NAO", "N"]

    def _registrar(identificacao, data_aditivo, ciclo_informado, tipo, valor_assinatura, aplicar, fator, tratamento, observacao="", origem="Resumo"):
        data_aditivo = pd.to_datetime(data_aditivo, dayfirst=True, errors="coerce")
        ciclo = normalizar_ciclo(ciclo_informado) if str(ciclo_informado or "").strip() else inferir_ciclo_por_data(data_aditivo, ciclos)
        valor_assinatura = numero_br(valor_assinatura)
        if valor_assinatura == 0 and pd.isna(data_aditivo):
            return

        tipo_txt = str(tipo or "Acréscimo").strip() or "Acréscimo"
        if tipo_txt.lower() in ["nan", "none"]:
            return
        if "supress" in normalizar_texto(tipo_txt) or "decresc" in normalizar_texto(tipo_txt):
            valor_assinatura = -abs(valor_assinatura)
        else:
            valor_assinatura = abs(valor_assinatura)

        aplicar_txt = str(aplicar or "Sim").strip().upper()

        # PATCH_VERTIV_FATOR_ADITIVOS_V3
        # Regra: para aditivos computáveis, priorizar o fator acumulado efetivo
        # calculado pelo motor de ciclos do cl8us. Esse fator já segue o critério
        # operacional da ferramenta, com fatores de ciclo saneados/arredondados,
        # evitando divergência por casas residuais vindas da planilha.
        fator_planilha = numero_br(fator)
        fator_mapa = fatores.get(ciclo, {}).get("fator_acumulado_efetivo", 0.0)

        if fator_mapa and fator_mapa > 0:
            fator_final = fator_mapa
        elif fator_planilha > 0 and fator_planilha <= 10:
            fator_final = fator_planilha
        else:
            fator_final = 1.0

        valor_reajustado = valor_assinatura if aplicar_txt in ["NAO", "NÃO", "N"] else valor_assinatura * fator_final
        tratamento_txt = str(tratamento or "Computar nesta análise").strip() or "Computar nesta análise"
        identificacao = str(identificacao or "").strip()

        linhas.append({
            "Identificação": identificacao,
            "Origem do lançamento": origem,
            "Data do aditivo": data_aditivo,
            "Ciclo/Marco": ciclo,
            "Tipo de alteração": tipo_txt,
            "Valor do aditivo na assinatura": valor_assinatura,
            "Fator aplicado": fator_final,
            "Valor do aditivo reajustado": valor_reajustado,
            "Tratamento do aditivo": tratamento_txt,
            "Observação": observacao,
            "Computa no Valor Global": not aditivo_informativo_ja_incorporado(tratamento_txt),
            "Item": identificacao,
            "Valor original da alteração": valor_assinatura,
            "Valor atualizado da alteração": valor_reajustado,
        })

    # A aba resumida só é usada se existir com nome próprio.
    # Não usar busca ampla por "ADITIVO", pois isso pode capturar indevidamente a aba
    # ADITIVOS_QUANTITATIVOS e ler a linha TOTAL como um terceiro aditivo.
    mapa_abas_exato = {normalizar_texto(s): s for s in xls.sheet_names}
    aba_res = (
        mapa_abas_exato.get("aditivos_supressoes")
        or mapa_abas_exato.get("aditivos_supressoes_resumidos")
    )
    resumo_lido = False
    if aba_res:
        try:
            df_res = ler_aba_com_cabecalho(bytes_arquivo, aba_res, termos_obrigatorios=["Tipo"])
            if not df_res.empty:
                col_ident = localizar_coluna(df_res, ["Identificação", "Identificacao", "Descrição", "Descricao"])
                col_data = localizar_coluna(df_res, ["Data do aditivo", "Data de assinatura", "Data"])
                col_ciclo = localizar_coluna(df_res, ["Ciclo/Marco", "Ciclo", "Marco"])
                col_tipo = localizar_coluna(df_res, ["Tipo de alteração", "Tipo"])
                col_valor = localizar_coluna(df_res, ["Valor do aditivo na assinatura", "Valor total do aditivo", "Valor total", "Valor total original", "Valor original da alteração", "Valor original"])
                col_aplicar = localizar_coluna(df_res, ["Aplicar reajuste acumulado? (Sim/Não)", "Aplicar reajuste", "Aplicar"])
                col_fator = localizar_coluna(df_res, ["Fator acumulado aplicável", "Fator aplicado", "Fator"])
                col_trat = localizar_coluna(df_res, ["Tratamento do aditivo", "Tratamento"])
                col_obs = localizar_coluna(df_res, ["Observação", "Observacao"])

                validos = 0
                for _, row in df_res.iterrows():
                    ident = str(row.get(col_ident, "") if col_ident else "").strip().upper()
                    if ident in ["TOTAL", "ORIENTAÇÃO", "ORIENTACAO"]:
                        continue
                    valor = numero_br(row.get(col_valor, 0)) if col_valor else 0
                    aplicar_val = row.get(col_aplicar, "Sim") if col_aplicar else "Sim"
                    if valor != 0 and _eh_sim_nao(aplicar_val):
                        validos += 1

                if validos:
                    for _, row in df_res.iterrows():
                        ident = str(row.get(col_ident, "") if col_ident else "").strip()
                        if ident.upper() in ["TOTAL", "ORIENTAÇÃO", "ORIENTACAO"]:
                            continue
                        valor = row.get(col_valor, 0) if col_valor else 0
                        if numero_br(valor) == 0:
                            continue
                        _registrar(
                            ident,
                            row.get(col_data, "") if col_data else "",
                            row.get(col_ciclo, "") if col_ciclo else "",
                            row.get(col_tipo, "Acréscimo") if col_tipo else "Acréscimo",
                            valor,
                            row.get(col_aplicar, "Sim") if col_aplicar else "Sim",
                            row.get(col_fator, 0) if col_fator else 0,
                            row.get(col_trat, "Computar nesta análise") if col_trat else "Computar nesta análise",
                            row.get(col_obs, "") if col_obs else "",
                            origem="Resumo",
                        )
                    resumo_lido = True
        except Exception:
            resumo_lido = False

    # Se a resumida estiver inválida/misalinhada, usa a quantitativa e evita dupla contagem.
    if not resumo_lido:
        aba = localizar_aba(xls, ["ADITIVOS_QUANTITATIVOS", "ADITIVOS", "Aditivos"])
        if aba:
            try:
                df = ler_aba_com_cabecalho(bytes_arquivo, aba, termos_obrigatorios=["Item"])
            except Exception:
                df = pd.DataFrame()
            if not df.empty:
                col_item = localizar_coluna(df, ["Item"])
                col_data = localizar_coluna(df, ["Data do aditivo", "Data de assinatura", "Data"])
                col_ciclo = localizar_coluna(df, ["Ciclo/Marco", "Ciclo", "Marco"])
                col_tipo = localizar_coluna(df, ["Tipo de alteração", "Tipo", "Acréscimo/Supressão"])
                col_qtd = localizar_coluna(df, ["Quantidade acrescida/suprimida", "Quantidade", "Qtd"])
                col_vu = localizar_coluna(df, ["Valor unitário original", "VU", "Valor unitario"])
                col_valor = localizar_coluna(df, ["Valor original da alteração", "Valor original do acréscimo", "Valor original"])
                col_aplicar = localizar_coluna(df, ["Aplicar reajuste acumulado? (Sim/Não)", "Aplicar reajuste", "Aplicar"])
                col_fator = localizar_coluna(df, ["Fator acumulado aplicável", "Fator", "Fator acumulado"])
                col_trat = localizar_coluna(df, ["Tratamento do aditivo", "Tratamento", "Computar nesta análise"])
                for _, row in df.iterrows():
                    item = str(row.get(col_item, "") if col_item else "").strip()
                    if not item or item.upper() in ["TOTAL", "ORIENTAÇÃO", "ORIENTACAO"]:
                        continue
                    qtd = numero_br(row.get(col_qtd, 0)) if col_qtd else 0
                    vu = numero_br(row.get(col_vu, 0)) if col_vu else 0
                    valor = numero_br(row.get(col_valor, 0)) if col_valor else 0
                    if valor == 0 and qtd != 0 and vu != 0:
                        valor = qtd * vu
                    if valor == 0:
                        continue
                    _registrar(
                        item,
                        row.get(col_data, "") if col_data else "",
                        row.get(col_ciclo, "") if col_ciclo else "",
                        row.get(col_tipo, "Acréscimo") if col_tipo else "Acréscimo",
                        valor,
                        row.get(col_aplicar, "Sim") if col_aplicar else "Sim",
                        row.get(col_fator, 0) if col_fator else 0,
                        row.get(col_trat, "Computar nesta análise") if col_trat else "Computar nesta análise",
                        "",
                        origem="Quantitativo",
                    )

    return limpar_nan_inf_df(pd.DataFrame(linhas))


def consolidar_aditivos_executivo(df_aditivos):
    """Agrupa linhas de itens em instrumentos aditivos executivos.

    Regra operacional: o instrumento é consolidado por data + ciclo/marco + tratamento,
    independentemente de conter acréscimos e supressões no mesmo ato. Assim, um mesmo
    termo aditivo com itens de acréscimo e decréscimo aparece como um único aditivo,
    preservando a quantidade de linhas de cada natureza para auditoria.
    """
    if df_aditivos is None or not isinstance(df_aditivos, pd.DataFrame) or df_aditivos.empty:
        return pd.DataFrame(columns=[
            "Aditivo", "Data do aditivo", "Ciclo/Marco", "Tipo de alteração", "Tratamento do aditivo",
            "Quantidade de linhas", "Valor do aditivo na assinatura", "Fator aplicado", "Valor do aditivo reajustado", "Computa no Valor Global",
            "Quantidade de acréscimos", "Quantidade de supressões"
        ])
    df = df_aditivos.copy()
    # Segurança adicional: elimina linhas de total/orientação ou resíduos de leitura indevida.
    if "Identificação" in df.columns:
        df = df[~df["Identificação"].astype(str).str.strip().str.upper().isin(["TOTAL", "ORIENTAÇÃO", "ORIENTACAO", "NAN"])].copy()
    if "Item" in df.columns:
        df = df[~df["Item"].astype(str).str.strip().str.upper().isin(["TOTAL", "ORIENTAÇÃO", "ORIENTACAO", "NAN"])].copy()
    if "Tipo de alteração" in df.columns:
        df = df[~df["Tipo de alteração"].astype(str).str.strip().str.lower().isin(["", "nan", "none"])].copy()
    for col in ["Data do aditivo", "Ciclo/Marco", "Tipo de alteração", "Tratamento do aditivo", "Computa no Valor Global"]:
        if col not in df.columns:
            df[col] = ""
    df["_data"] = pd.to_datetime(df["Data do aditivo"], dayfirst=True, errors="coerce")
    df["_ciclo"] = df["Ciclo/Marco"].astype(str)
    df["_trat"] = df["Tratamento do aditivo"].astype(str)
    df["_comp"] = df["Computa no Valor Global"].astype(bool)
    df["_eh_supressao"] = df["Tipo de alteração"].astype(str).apply(lambda v: "supress" in normalizar_texto(v) or "decresc" in normalizar_texto(v))
    df["_eh_acrescimo"] = ~df["_eh_supressao"]

    # Consolida o instrumento pelo marco material do ato: mesma data + mesmo ciclo.
    # Não separa acréscimo e supressão em dois aditivos, pois ambos podem compor o
    # mesmo termo/instrumento formal. O tratamento e a marcação computável são
    # agregados para preservar governança sem duplicar a contagem.
    agrupado = (
        df.groupby(["_data", "_ciclo"], dropna=False)
        .agg(
            quantidade_linhas=("Valor do aditivo na assinatura", "size"),
            quantidade_acrescimos=("_eh_acrescimo", "sum"),
            quantidade_supressoes=("_eh_supressao", "sum"),
            valor_assinatura=("Valor do aditivo na assinatura", "sum"),
            valor_reajustado=("Valor do aditivo reajustado", "sum"),
            tratamentos=("_trat", lambda x: " / ".join(sorted({str(v).strip() for v in x if str(v).strip()}))),
            computa=("_comp", "max"),
        )
        .reset_index()
        .sort_values(["_data", "_ciclo"], na_position="last")
        .reset_index(drop=True)
    )
    linhas = []
    for idx, row in agrupado.iterrows():
        valor_ass = numero_seguro(row.get("valor_assinatura", 0), 0.0)
        valor_reaj = numero_seguro(row.get("valor_reajustado", 0), 0.0)
        fator = (valor_reaj / valor_ass) if abs(valor_ass) > 0.005 else 1.0
        qtd_ad = int(numero_seguro(row.get("quantidade_acrescimos", 0), 0))
        qtd_sup = int(numero_seguro(row.get("quantidade_supressoes", 0), 0))
        resumo_qtd = []
        if qtd_ad:
            resumo_qtd.append(f"{qtd_ad} itens ad.")
        if qtd_sup:
            resumo_qtd.append(f"{qtd_sup} itens exc.")
        qtd_total = int(numero_seguro(row.get("quantidade_linhas", 0), 0))
        linhas.append({
            "Aditivo": f"Aditivo {idx + 1}",
            "Data do aditivo": row.get("_data"),
            "Ciclo/Marco": normalizar_ciclo(row.get("_ciclo", "")),
            "Tipo de alteração": "Aditivo consolidado",
            "Tratamento do aditivo": row.get("tratamentos", ""),
            "Quantidade de linhas": " / ".join(resumo_qtd) if resumo_qtd else str(qtd_total),
            "Valor do aditivo na assinatura": valor_ass,
            "Fator aplicado": fator,
            "Valor do aditivo reajustado": valor_reaj,
            "Computa no Valor Global": bool(row.get("computa", True)),
            "Quantidade de acréscimos": qtd_ad,
            "Quantidade de supressões": qtd_sup,
        })
    return limpar_nan_inf_df(pd.DataFrame(linhas))

def montar_comparativo_executivo(
    valor_original,
    valor_formalizado_anterior,
    impacto_analise_atual,
    total_pago,
    total_devido,
    delta_total,
    rem_original,
    rem_atualizado,
    valor_atualizado_contrato,
    total_aditivos,
    total_aditivos_informativos=0.0,
    quantidade_aditivos_total=0,
    quantidade_meses_sem_efeito=0,
    valor_total_sem_efeito=0.0,
):
    valor_original_num = numero_seguro(valor_original, 0.0)
    valor_formalizado_num = numero_seguro(valor_formalizado_anterior, valor_original_num)

    linhas = [
        {"Indicador": "Valor original do contrato", "Valor": valor_original_num},
    ]

    if abs(valor_formalizado_num - valor_original_num) > 0.005:
        linhas.append({"Indicador": "Valor formalizado antes desta análise", "Valor": valor_formalizado_num})

    linhas.extend([
        {"Indicador": "Número de Aditivos", "Valor": int(quantidade_aditivos_total or 0)},
        {"Indicador": "Valor pago efetivo", "Valor": total_pago},
        {"Indicador": "Valor teórico calculado", "Valor": total_devido},
        {"Indicador": "Valor represado a pagar", "Valor": delta_total},
        {"Indicador": "Meses sem efeitos financeiros", "Valor": int(quantidade_meses_sem_efeito or 0)},
        {"Indicador": "Valor total sem efeito financeiro", "Valor": valor_total_sem_efeito},
        {"Indicador": "Saldo remanescente original", "Valor": rem_original},
        {"Indicador": "Saldo remanescente atualizado", "Valor": rem_atualizado},
        {"Indicador": "Aditivos/Supressões registrados (informativo)", "Valor": total_aditivos},
        {"Indicador": "Aditivos/Supressões informativos", "Valor": total_aditivos_informativos},
        {"Indicador": "Valor Total Atualizado do Contrato", "Valor": valor_atualizado_contrato},
    ])
    return pd.DataFrame(linhas)


def montar_delta_por_ciclo(df_financeiro_por_ciclo, df_execucao_atualizada, ciclos):
    """Monta quadro de Delta por Ciclo incluindo C0 e ciclos identificados sem lançamento financeiro."""
    linhas = []

    valor_c0 = 0.0
    if isinstance(df_execucao_atualizada, pd.DataFrame) and not df_execucao_atualizada.empty and "Ciclo" in df_execucao_atualizada.columns:
        df_c0 = df_execucao_atualizada[df_execucao_atualizada["Ciclo"].astype(str).str.upper().eq("C0")]
        if not df_c0.empty:
            valor_c0 = float(df_c0["Valor executado original"].sum())

    linhas.append({
        "Ciclo": "C0",
        "Situação": "BASE",
        "Tratamento financeiro": "Base",
        "Fator aplicado ao retroativo": 1.0,
        "Valor pago efetivo": valor_c0,
        "Valor teórico calculado": valor_c0,
        "Delta do ciclo": 0.0,
    })

    if isinstance(df_financeiro_por_ciclo, pd.DataFrame) and not df_financeiro_por_ciclo.empty:
        for _, row in df_financeiro_por_ciclo.iterrows():
            ciclo = normalizar_ciclo(row.get("Ciclo", ""))
            if not ciclo or ciclo == "TOTAL" or ciclo == "C0":
                continue
            linhas.append({
                "Ciclo": ciclo,
                "Situação": row.get("Situação", ""),
                "Tratamento financeiro": row.get("Tratamento financeiro", ""),
                "Fator aplicado ao retroativo": row.get("Fator aplicado ao retroativo", 1.0),
                "Valor pago efetivo": numero_seguro(row.get("Valor pago efetivo", 0.0), 0.0),
                "Valor teórico calculado": numero_seguro(row.get("Valor teórico calculado", 0.0), 0.0),
                "Delta do ciclo": numero_seguro(row.get("Delta do ciclo", 0.0), 0.0),
            })

    ciclos_existentes = {normalizar_ciclo(l.get("Ciclo", "")) for l in linhas}
    fatores = mapa_fatores(ciclos) if isinstance(ciclos, pd.DataFrame) else {}
    if isinstance(ciclos, pd.DataFrame) and not ciclos.empty:
        for _, ciclo_row in ciclos.iterrows():
            ciclo = normalizar_ciclo(ciclo_row.get("Ciclo", ""))
            if ciclo and ciclo not in ciclos_existentes:
                info = fatores.get(ciclo, {})
                linhas.append({
                    "Ciclo": ciclo,
                    "Situação": info.get("situacao", ciclo_row.get("Situação", "")),
                    "Tratamento financeiro": info.get("tratamento", ciclo_row.get("Tratamento financeiro do ciclo", "")),
                    "Fator aplicado ao retroativo": info.get("fator_retroativo", 1.0),
                    "Valor pago efetivo": 0.0,
                    "Valor teórico calculado": 0.0,
                    "Delta do ciclo": 0.0,
                })

    df = pd.DataFrame(linhas)
    if not df.empty:
        df = df.sort_values(by="Ciclo", key=lambda s: s.map(numero_ciclo)).reset_index(drop=True)
    return df


def arredondar_resultado_financeiro(resultado):
    """Aplica arredondamento operacional aos valores exibidos/exportados."""
    chaves_moeda = [
        "valor_original_contrato", "valor_formalizado_anterior", "impacto_analise_atual",
        "valor_pago_efetivo", "total_pago_faturado", "valor_teorico_calculado",
        "total_devido_reajustado", "delta_total", "delta_acumulado",
        "valor_represado_a_pagar", "remanescente_original", "remanescente_reajustado",
        "valor_executado_atualizado", "valor_calculado_sem_aditivos",
        "valor_atualizado_contrato", "valor_global_financeiro",
        "total_aditivos_atualizados", "total_aditivos_informativos",
    ]
    for chave in chaves_moeda:
        if chave in resultado:
            resultado[chave] = round(numero_seguro(resultado.get(chave), 0.0), 2)
    colunas_moeda_exatas = {
        "valor_pago_efetivo", "valor_teorico_calculado", "valor_pago_faturado",
        "valor_devido_reajustado", "delta_do_ciclo", "delta_acumulado",
        "valor_executado_original", "valor_executado_atualizado",
        "remanescente_original", "remanescente_atualizado",
        "valor_unitario", "total_r", "valor_do_aditivo_na_assinatura",
        "valor_do_aditivo_reajustado", "valor_original_da_alteracao",
        "valor_atualizado_da_alteracao", "valor_total_original", "valor",
    }
    colunas_fator_exatas = {"fator", "fator_aplicado", "fator_acumulado", "fator_acumulado_efetivo", "fator_ciclo_efetivo", "fator_aplicado_ao_retroativo"}
    colunas_pct_exatas = {"percentual_acumulado_aplicado", "variacao", "percentual_aplicado", "percentual_apurado_pelo_indice"}
    termos_moeda = ("valor", "total", "saldo", "remanescente", "aditivo", "supress", "delta", "pago", "teorico", "executado", "unitario")
    termos_excluir_moeda = ("fator", "percentual", "variacao", "quantidade", "qtd", "ciclo", "data", "status", "verificacao")
    colunas_texto_exatas = {
        "aditivo", "identificacao", "origem_do_lancamento", "tipo_de_alteracao",
        "tratamento_do_aditivo", "observacao", "computa_no_valor_global",
        "marcado_como_computavel_no_arquivo",
    }
    for chave, valor in list(resultado.items()):
        if isinstance(valor, pd.DataFrame) and not valor.empty:
            df = valor.copy()
            for col in df.columns:
                col_norm = normalizar_texto(col)
                eh_coluna_moeda = (
                    col_norm not in colunas_texto_exatas
                    and (
                        col_norm in colunas_moeda_exatas
                        or (any(t in col_norm for t in termos_moeda) and not any(t in col_norm for t in termos_excluir_moeda))
                    )
                )
                if col_norm in colunas_fator_exatas:
                    df[col] = df[col].apply(lambda x: round(numero_seguro(x, 1.0), 4) if str(x).strip() != "" else x)
                elif col_norm in colunas_pct_exatas:
                    df[col] = df[col].apply(lambda x: round(numero_seguro(x, 0.0), 4) if str(x).strip() != "" else x)
                elif eh_coluna_moeda:
                    df[col] = df[col].apply(lambda x: round(numero_seguro(x, 0.0), 2) if str(x).strip() != "" else x)
            resultado[chave] = df
    return resultado


def montar_composicao_valor_total(
    df_execucao,
    remanescente_atualizado,
    ciclo_ultimo_rem,
    valor_total_atualizado,
    df_aditivos_computaveis=None,
    total_aditivos_atualizados=0.0,
):
    """Monta quadro executivo da composição do Valor Total Atualizado do Contrato.

    Regra adotada: Valor Total Atualizado = execução atualizada por ciclo + saldo remanescente atualizado
    + aditivos/supressões computáveis quando indicados como parcela da análise atual.
    Aditivos informativos ou já incorporados permanecem rastreáveis, mas não entram como parcela autônoma.
    """
    linhas = []
    soma_componentes = 0.0

    if isinstance(df_execucao, pd.DataFrame) and not df_execucao.empty:
        df_exec = df_execucao.copy()
        if "Ciclo" in df_exec.columns:
            df_exec = df_exec.sort_values(by="Ciclo", key=lambda s: s.map(numero_ciclo))
        for _, row in df_exec.iterrows():
            ciclo = str(row.get("Ciclo", "")).strip() or "Ciclo não identificado"
            valor = numero_seguro(row.get("Valor executado atualizado", 0.0), 0.0)
            if abs(valor) > 0.004:
                linhas.append({
                    "Componente": f"{ciclo} - execução atualizada",
                    "Ciclo/Referência": ciclo,
                    "Valor": round(valor, 2),
                    "Observação": "Valor executado/consumido no ciclo, atualizado pelo fator aplicável à execução.",
                })
                soma_componentes += valor

    if abs(numero_seguro(remanescente_atualizado, 0.0)) > 0.004:
        valor = numero_seguro(remanescente_atualizado, 0.0)
        linhas.append({
            "Componente": "Saldo remanescente atualizado",
            "Ciclo/Referência": str(ciclo_ultimo_rem or "Último ciclo informado"),
            "Valor": round(valor, 2),
            "Observação": "Saldo remanescente já atualizado pelos reajustes aplicáveis ao ciclo de referência, exceto ciclos preclusos não admitidos por negociação.",
        })
        soma_componentes += valor

    valor_total = numero_seguro(valor_total_atualizado, soma_componentes)

    total_aditivos_comp = numero_seguro(total_aditivos_atualizados, 0.0)
    if abs(total_aditivos_comp) > 0.004:
        linhas.append({
            "Componente": "Aditivos/supressões computáveis atualizados",
            "Ciclo/Referência": "ADITIVOS_QUANTITATIVOS",
            "Valor": round(total_aditivos_comp, 2),
            "Observação": "Aditivos/supressões marcados como computáveis nesta análise e não incorporados ao saldo remanescente informado no corte.",
        })
        soma_componentes += total_aditivos_comp

    linhas.append({
        "Componente": "Valor Total Atualizado do Contrato",
        "Ciclo/Referência": "Total",
        "Valor": round(valor_total, 2),
        "Observação": "Resultado consolidado: execução atualizada por ciclo + saldo remanescente atualizado. Aditivos/supressões não são somados como parcela autônoma para evitar dupla contagem.",
    })

    diferenca = round(valor_total - soma_componentes, 2)
    if abs(diferenca) > 0.01:
        linhas.append({
            "Componente": "Diferença de conferência",
            "Ciclo/Referência": "Controle",
            "Valor": diferenca,
            "Observação": "Diferença entre o total consolidado e a soma execução + saldo; revisar se houver valor material.",
        })

    return pd.DataFrame(linhas)



def montar_auditoria_consistencia(resultado):
    """Monta quadro de auditoria objetiva para conferência final da análise."""
    linhas = []

    def add_check(item, status, detalhe, diferenca=None):
        if diferenca is None:
            valor_auditoria = ""
        elif isinstance(diferenca, str):
            valor_auditoria = diferenca
        else:
            valor_auditoria = round(numero_seguro(diferenca, 0.0), 2)
        linhas.append({
            "Verificação": item,
            "Status": status,
            "Diferença/Valor": valor_auditoria,
            "Detalhamento": detalhe,
        })

    valor_total = numero_seguro(resultado.get("valor_atualizado_contrato", 0.0), 0.0)
    df_comp = resultado.get("df_composicao_valor_total", pd.DataFrame())
    if isinstance(df_comp, pd.DataFrame) and not df_comp.empty and "Componente" in df_comp.columns and "Valor" in df_comp.columns:
        mask_total = df_comp["Componente"].astype(str).str.contains("Valor Total Atualizado", case=False, na=False)
        soma_componentes = df_comp.loc[~mask_total, "Valor"].apply(numero_seguro).sum()
        diferenca = round(valor_total - soma_componentes, 2)
        add_check(
            "Composição do Valor Total Atualizado",
            "OK" if abs(diferenca) <= 0.01 else "ATENÇÃO",
            "Compara o Valor Total Atualizado do Contrato com a soma dos componentes: execução atualizada por ciclo + saldo remanescente atualizado + aditivos/supressões computáveis.",
            diferenca,
        )
    else:
        add_check("Composição do Valor Total Atualizado", "ATENÇÃO", "Quadro de composição não localizado no resultado.")

    valor_represado = numero_seguro(resultado.get("valor_represado_a_pagar", 0.0), 0.0)
    add_check(
        "Valor Represado a Pagar",
        "OK" if valor_represado >= -0.01 else "ATENÇÃO",
        "Confere se o valor represado a pagar não está negativo.",
        valor_represado,
    )

    df_ad_exec = resultado.get("df_aditivos_executivo", pd.DataFrame())
    total_aditivos = numero_seguro(resultado.get("total_aditivos_atualizados", 0.0), 0.0)
    if isinstance(df_ad_exec, pd.DataFrame) and not df_ad_exec.empty and "Valor do aditivo reajustado" in df_ad_exec.columns:
        soma_aditivos = df_ad_exec["Valor do aditivo reajustado"].apply(numero_seguro).sum()
        diferenca_ad = round(total_aditivos - soma_aditivos, 2)
        add_check(
            "Aditivos registrados",
            "OK" if abs(diferenca_ad) <= 0.01 else "ATENÇÃO",
            "Confere o total de aditivos/supressões registrados no consolidado executivo. Valores computáveis entram no Valor Total Atualizado; valores informativos permanecem apenas como memória.",
            diferenca_ad,
        )
    else:
        add_check(
            "Aditivos registrados",
            "OK" if abs(total_aditivos) <= 0.01 else "ATENÇÃO",
            "Não há consolidado executivo de aditivos. Se houver aditivos no caso, revise a aba ADITIVOS_QUANTITATIVOS.",
            total_aditivos,
        )

    # Alerta de preenchimento: muitas datas distintas em aditivos podem indicar arraste acidental no Excel.
    df_ad_raw = resultado.get("df_aditivos", pd.DataFrame())
    if isinstance(df_ad_raw, pd.DataFrame) and not df_ad_raw.empty and "Data do aditivo" in df_ad_raw.columns:
        datas_ad = pd.to_datetime(df_ad_raw["Data do aditivo"], dayfirst=True, errors="coerce").dropna().dt.normalize().drop_duplicates().sort_values()
        qtd_datas = int(len(datas_ad))
        sequencia_curta = False
        if qtd_datas >= 3:
            difs = datas_ad.diff().dropna().dt.days.tolist()
            sequencia_curta = any(d in [1, 2] for d in difs)
        if qtd_datas > 3 or sequencia_curta:
            detalhe_datas = ", ".join([d.strftime("%d/%m/%Y") for d in datas_ad.head(8)])
            add_check(
                "Datas dos aditivos",
                "ATENÇÃO",
                "Foram identificadas várias datas distintas na aba de aditivos. Confirme se cada data corresponde efetivamente a um instrumento aditivo diferente. Em preenchimentos manuais, o Excel pode arrastar datas em sequência e gerar contagem indevida de aditivos.",
                f"{qtd_datas} data(s): {detalhe_datas}",
            )
        else:
            add_check(
                "Datas dos aditivos",
                "OK",
                "Quantidade de datas distintas em aditivos dentro do padrão esperado para conferência executiva.",
                f"{qtd_datas} data(s)",
            )
    else:
        add_check("Datas dos aditivos", "OK", "Não há datas de aditivos para auditoria específica.", "0 data(s)")

    df_ciclos = resultado.get("df_ciclos", pd.DataFrame())
    if isinstance(df_ciclos, pd.DataFrame) and not df_ciclos.empty:
        fatores_altos = []
        for col in ["Fator", "Fator acumulado", "Fator acumulado efetivo", "Fator aplicado ao retroativo"]:
            if col in df_ciclos.columns:
                serie = df_ciclos[col].apply(lambda x: numero_seguro(x, 1.0))
                fatores_altos.extend([v for v in serie.tolist() if abs(v) > 5])
        add_check(
            "Fatores de reajuste",
            "OK" if not fatores_altos else "ATENÇÃO",
            "Confere se não há fatores anormalmente altos, que poderiam indicar leitura indevida de moeda como fator.",
            max(fatores_altos) if fatores_altos else 0.0,
        )
    else:
        add_check("Fatores de reajuste", "ATENÇÃO", "Tabela de ciclos não localizada.")

    df_fin = resultado.get("df_financeiro_por_ciclo", pd.DataFrame())
    if isinstance(df_fin, pd.DataFrame) and not df_fin.empty:
        total_row = df_fin[df_fin.get("Ciclo", "").astype(str).str.upper().eq("TOTAL")] if "Ciclo" in df_fin.columns else pd.DataFrame()
        if not total_row.empty:
            base = df_fin[~df_fin["Ciclo"].astype(str).str.upper().eq("TOTAL")].copy()
            dif_pago = round(numero_seguro(total_row.iloc[0].get("Valor pago efetivo", 0.0), 0.0) - base["Valor pago efetivo"].apply(numero_seguro).sum(), 2) if "Valor pago efetivo" in base.columns else 0.0
            add_check(
                "Financeiro por Ciclo",
                "OK" if abs(dif_pago) <= 0.01 else "ATENÇÃO",
                "Confere se a linha TOTAL do financeiro por ciclo corresponde à soma dos ciclos informados.",
                dif_pago,
            )
        else:
            add_check("Financeiro por Ciclo", "OK", "Tabela financeira não possui linha TOTAL, mas há lançamentos por ciclo para conferência.")
    else:
        add_check("Financeiro por Ciclo", "ATENÇÃO", "Não há dados financeiros por ciclo no resultado.")

    # Auditoria objetiva do padrão de casas decimais após o arredondamento operacional.
    moeda_cols = {
        "valor_pago_efetivo", "valor_teorico_calculado", "valor_pago_faturado",
        "valor_devido_reajustado", "delta_do_ciclo", "delta_acumulado",
        "valor_executado_original", "valor_executado_atualizado",
        "remanescente_original", "remanescente_atualizado",
        "valor_unitario", "total_r", "valor_do_aditivo_na_assinatura",
        "valor_do_aditivo_reajustado", "valor_original_da_alteracao",
        "valor_atualizado_da_alteracao", "valor_total_original", "valor",
    }
    fator_cols = {
        "fator", "fator_aplicado", "fator_acumulado",
        "fator_acumulado_efetivo", "fator_ciclo_efetivo", "fator_aplicado_ao_retroativo",
    }
    problemas_moeda = 0
    problemas_fator = 0
    termos_moeda_aud = ("valor", "total", "saldo", "remanescente", "aditivo", "supress", "delta", "pago", "teorico", "executado", "unitario")
    termos_excluir_moeda_aud = ("fator", "percentual", "variacao", "quantidade", "qtd", "ciclo", "data", "status", "verificacao")
    colunas_texto_auditoria = {
        "aditivo", "identificacao", "origem_do_lancamento", "tipo_de_alteracao",
        "tratamento_do_aditivo", "observacao", "computa_no_valor_global",
        "marcado_como_computavel_no_arquivo",
    }
    for nome_df, df in resultado.items():
        if not isinstance(df, pd.DataFrame) or df.empty or nome_df == "df_auditoria_consistencia":
            continue
        for col in df.columns:
            col_norm = normalizar_texto(col)
            eh_coluna_moeda = (
                col_norm not in colunas_texto_auditoria
                and (
                    col_norm in moeda_cols
                    or (any(t in col_norm for t in termos_moeda_aud) and not any(t in col_norm for t in termos_excluir_moeda_aud))
                )
            )
            eh_coluna_fator = col_norm in fator_cols
            if not eh_coluna_moeda and not eh_coluna_fator:
                continue
            serie = pd.to_numeric(df[col], errors="coerce").dropna()
            if serie.empty:
                continue
            if eh_coluna_moeda:
                # Tolera apenas ruído binário residual de float, mas sinaliza qualquer valor com diferença material de centavo.
                problemas_moeda += int((serie - serie.round(2)).abs().gt(0.000001).sum())
            elif eh_coluna_fator:
                problemas_fator += int((serie - serie.round(4)).abs().gt(0.000001).sum())
    add_check(
        "Arredondamento monetário",
        "OK" if problemas_moeda == 0 else "ATENÇÃO",
        "Confere se valores financeiros, totais, unitários e saldos estão com duas casas decimais operacionais nas tabelas processadas.",
        f"{int(problemas_moeda)} ocorrência(s)",
    )
    add_check(
        "Arredondamento de fatores",
        "OK" if problemas_fator == 0 else "ATENÇÃO",
        "Confere se fatores foram preservados com quatro casas decimais nas tabelas processadas.",
        f"{int(problemas_fator)} ocorrência(s)",
    )

    return pd.DataFrame(linhas)


def _aba_exata(xls, nome):
    """Localiza aba por correspondência normalizada exata."""
    alvo = normalizar_texto(nome)
    for sheet in xls.sheet_names:
        if normalizar_texto(sheet) == alvo:
            return sheet
    return None


def _ler_ciclos_apurados_consumo(bytes_arquivo, xls):
    """Lê CICLOS_APURADOS preservando o fator acumulado informado no modelo.

    Este leitor é exclusivo do Modo Consumo por Itens/Ciclo. A planilha gerada
    para esse modo usa o fator acumulado da Etapa 1 para calcular os valores
    unitários atualizados; por isso, aqui não se recalcula nem arredonda o fator.
    """
    aba = _aba_exata(xls, "CICLOS_APURADOS") or _aba_exata(xls, "CICLOS") or _aba_exata(xls, "CICLO")
    if not aba:
        return pd.DataFrame(), {"C0": 1.0}

    bruto = pd.read_excel(BytesIO(bytes_arquivo), sheet_name=aba, header=None)
    linha_cabecalho = None
    for idx, row in bruto.iterrows():
        valores = [normalizar_texto(v) for v in row.tolist()]
        if any(v == "ciclo" for v in valores) and any(v == "fator_acumulado" for v in valores):
            linha_cabecalho = idx
            break
    if linha_cabecalho is None:
        return pd.DataFrame(), {"C0": 1.0}

    df = pd.read_excel(BytesIO(bytes_arquivo), sheet_name=aba, header=linha_cabecalho)
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~pd.Series(df.columns).astype(str).str.startswith("Unnamed").values]
    if df.empty:
        return pd.DataFrame(), {"C0": 1.0}

    col_ciclo = localizar_coluna(df, ["Ciclo"])
    col_base = localizar_coluna(df, ["Data-base", "Base"])
    col_janela = localizar_coluna(df, ["Janela de admissibilidade", "Janela"])
    col_pedido = localizar_coluna(df, ["Data do pedido", "Pedido"])
    col_inicio = localizar_coluna(df, ["Início financeiro", "Inicio financeiro"])
    col_percentual = localizar_coluna(df, ["Percentual aplicado", "Percentual", "Variação"])
    col_fator_acum = localizar_coluna(df, ["Fator acumulado", "Fator acumulado final", "Fator Acumulado"])
    col_fator = localizar_coluna(df, ["Fator do ciclo", "Fator", "Fator ciclo"])
    col_situacao = localizar_coluna(df, ["Situação", "Situacao", "Resultado"])
    col_ref = localizar_coluna(df, ["Referência para preenchimento", "Referencia para preenchimento", "Referência"])
    col_obs = localizar_coluna(df, ["Observação", "Observacao", "Obs"])
    if col_ciclo is None:
        return pd.DataFrame(), {"C0": 1.0}

    linhas = []
    fatores = {}
    fator_corrente = 1.0
    for _, row in df.iterrows():
        ciclo = normalizar_ciclo(row.get(col_ciclo, ""))
        if not ciclo or ciclo.upper() == "TOTAL":
            continue
        fator_acum = numero_br(row.get(col_fator_acum, "")) if col_fator_acum else 0.0
        if not fator_acum:
            fator_ciclo = numero_br(row.get(col_fator, "")) if col_fator else 1.0
            fator_acum = fator_corrente * (fator_ciclo if fator_ciclo else 1.0)
        if not fator_acum:
            fator_acum = 1.0
        fator_corrente = fator_acum
        fatores[ciclo] = float(fator_acum)
        linhas.append({
            "Ciclo": ciclo,
            "Data-base": row.get(col_base, "") if col_base else "",
            "Janela de admissibilidade": row.get(col_janela, "") if col_janela else "",
            "Data do pedido": row.get(col_pedido, "") if col_pedido else "",
            "Início financeiro": row.get(col_inicio, "") if col_inicio else "",
            "Percentual aplicado": row.get(col_percentual, "") if col_percentual else "",
            "Fator acumulado": float(fator_acum),
            "Situação": row.get(col_situacao, "") if col_situacao else "",
            "Referência para preenchimento": row.get(col_ref, "") if col_ref else "",
            "Observação": row.get(col_obs, "") if col_obs else "",
        })
    fatores.setdefault("C0", 1.0)
    return pd.DataFrame(linhas), fatores


def _localizar_colunas_consumo_por_ciclo(df):
    """Retorna mapa {C0: coluna, C1: coluna...} para CONSUMO_ITENS."""
    mapa = {}
    for col in df.columns:
        original = str(col).strip()
        n = normalizar_texto(original)
        if not n or "total" in n or "saldo" in n or "check" in n:
            continue
        m = re.search(r"(?:^|_)(c\d+)(?:$|_)", n)
        if not m:
            # Também aceita cabeçalho simples C0, C1, C2 etc.
            m = re.fullmatch(r"c\d+", n)
            ciclo = n.upper() if m else ""
        else:
            ciclo = m.group(1).upper()
        if not ciclo:
            continue
        if any(p in n for p in ["consum", "execut", "qtd", "quantidade"]) or re.fullmatch(r"c\d+", n):
            mapa[ciclo] = col
    return dict(sorted(mapa.items(), key=lambda kv: numero_ciclo(kv[0])))


def _ler_consumo_itens_matriz(bytes_arquivo, xls):
    """Lê o modelo estruturado do Modo Consumo por Itens/Ciclo.

    Formato esperado: aba CONSUMO_ITENS com uma matriz simples por item:
    Item | Quantidade contratada | Valor unitário original/base | Consumido C0 | Consumido C1...
    """
    aba = _aba_exata(xls, "CONSUMO_ITENS")
    if not aba:
        raise ValueError("Aba CONSUMO_ITENS não encontrada.")

    # A aba tem título e orientação nas primeiras linhas; localizar a linha real do cabeçalho
    # por correspondência forte, para não confundir o texto de orientação com cabeçalho.
    bruto = pd.read_excel(BytesIO(bytes_arquivo), sheet_name=aba, header=None)
    linha_cabecalho = None
    for idx, row in bruto.iterrows():
        valores = [normalizar_texto(v) for v in row.tolist()]
        tem_item = any(v == "item" for v in valores)
        tem_qtd = any(v in ["quantidade_contratada", "qtd_contratada"] for v in valores)
        tem_vu = any(v in ["valor_unitario_original_base", "valor_unitario_original", "vu_original", "vu_c0"] for v in valores)
        tem_consumo = any(re.search(r"(?:^|_)c\d+(?:$|_)", v) and ("consum" in v or "qtd" in v or "quantidade" in v or "execut" in v or re.fullmatch(r"c\d+", v)) for v in valores)
        if tem_item and tem_qtd and tem_vu and tem_consumo:
            linha_cabecalho = idx
            break
    if linha_cabecalho is None:
        df_tmp = ler_aba_com_cabecalho(bytes_arquivo, aba, termos_obrigatorios=["Item", "Quantidade contratada", "Valor unitário original"])
        colunas_tmp = ", ".join([str(c) for c in df_tmp.columns]) if not df_tmp.empty else "nenhuma coluna identificada"
        raise ValueError(
            "Leitor Consumo por Itens/Ciclo " + LEITOR_CONSUMO_ITENS_CICLO_VERSAO +
            ": não foi possível localizar a linha de cabeçalho da aba CONSUMO_ITENS. " +
            "Colunas detectadas pela leitura preliminar: " + colunas_tmp
        )
    df = pd.read_excel(BytesIO(bytes_arquivo), sheet_name=aba, header=linha_cabecalho)
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~pd.Series(df.columns).astype(str).str.startswith("Unnamed").values]

    col_item = localizar_coluna(df, ["Item"])
    col_qtd = localizar_coluna(df, ["Quantidade contratada", "Qtd contratada", "Quantidade", "Qtd"])
    col_vu = localizar_coluna(df, ["Valor unitário original/base", "Valor unitário original", "VU original", "VU C0", "Valor unitario original"])
    col_saldo = localizar_coluna(df, ["Saldo a faturar", "Saldo atual", "Quantidade saldo a faturar", "Saldo"])
    colunas_consumo = _localizar_colunas_consumo_por_ciclo(df)

    if col_item is None or col_qtd is None or col_vu is None:
        raise ValueError(
            "Leitor Consumo por Itens/Ciclo " + LEITOR_CONSUMO_ITENS_CICLO_VERSAO +
            ": a aba CONSUMO_ITENS não possui Item, Quantidade contratada e Valor unitário original/base. " +
            "Colunas encontradas: " + ", ".join([str(c) for c in df.columns])
        )
    if not colunas_consumo:
        raise ValueError(
            "Leitor Consumo por Itens/Ciclo " + LEITOR_CONSUMO_ITENS_CICLO_VERSAO +
            ": a aba CONSUMO_ITENS não possui colunas de consumo por ciclo. " +
            "Aceito: Consumido C0, Consumido C1, Consumo C1, Qtd consumida C1, Executado C1 ou C1. " +
            "Colunas encontradas: " + ", ".join([str(c) for c in df.columns])
        )

    linhas = []
    for _, row in df.iterrows():
        item = row.get(col_item, "")
        item_txt = str(item).strip()
        if not item_txt or item_txt.upper() == "TOTAL" or item_txt.lower() in ["nan", "none"]:
            continue
        qtd_contratada = numero_br(row.get(col_qtd, 0))
        vu_original = numero_br(row.get(col_vu, 0))
        if qtd_contratada == 0 and vu_original == 0:
            continue
        consumos = {ciclo: numero_br(row.get(col, 0)) for ciclo, col in colunas_consumo.items()}
        saldo_informado = numero_br(row.get(col_saldo, "")) if col_saldo else None
        linhas.append({
            "Item": item_txt,
            "Quantidade contratada": qtd_contratada,
            "Valor unitário original/base": vu_original,
            "Consumos": consumos,
            "Saldo a faturar informado": saldo_informado,
        })
    if not linhas:
        raise ValueError("A aba CONSUMO_ITENS não contém linhas de itens preenchidas.")
    return linhas, list(colunas_consumo.keys())



def construir_valores_unitarios_consumo_ciclo(itens_consumo, fatores, ciclos_consumo):
    """Monta a fotografia de remanescente por início de ciclo no Modo Consumo por Itens/Ciclo.

    A aba/tabela de Valores Unitários e Totais, neste modo, não representa o consumo do ciclo,
    mas sim o saldo remanescente existente no início de cada ciclo, calculado a partir das
    quantidades consumidas informadas pela fiscalização.
    """
    if not itens_consumo:
        return pd.DataFrame(columns=["Item", "Ciclo", "Quantidade", "Valor unitário", "Total R$", "Situação", "Observação"])

    ciclos_base = set(ciclos_consumo or []) | set(fatores.keys())
    ciclos_base.add("C0")
    ciclos_ordenados = sorted([c for c in ciclos_base if c], key=numero_ciclo)
    linhas = []

    for item in itens_consumo:
        item_id = item.get("Item", "")
        qtd_contratada = numero_seguro(item.get("Quantidade contratada", 0.0), 0.0)
        vu_original = numero_seguro(item.get("Valor unitário original/base", 0.0), 0.0)
        consumos = item.get("Consumos", {}) or {}

        for ciclo in ciclos_ordenados:
            n_ciclo = numero_ciclo(ciclo)
            if ciclo == "C0":
                qtd_inicio = qtd_contratada
                observacao = "Quantidade contratada original no ciclo-base."
            else:
                consumido_antes = 0.0
                for c_consumo, qtd in consumos.items():
                    if numero_ciclo(c_consumo) < n_ciclo:
                        consumido_antes += numero_seguro(qtd, 0.0)
                qtd_inicio = qtd_contratada - consumido_antes
                observacao = "Remanescente no início do ciclo, antes do consumo informado para o próprio ciclo."

            fator = numero_seguro(fatores.get(ciclo, 1.0), 1.0)
            vu_atualizado = round(vu_original * fator, 2)
            total = round(qtd_inicio * vu_atualizado, 2)
            situacao = "Remanescente por início de ciclo"
            if qtd_inicio < -0.000001:
                situacao = "DIVERGÊNCIA: consumo anterior maior que contratado"

            linhas.append({
                "Item": item_id,
                "Ciclo": ciclo,
                "Quantidade": round(qtd_inicio, 6),
                "Valor unitário": vu_atualizado,
                "Total R$": total,
                "Situação": situacao,
                "Observação": observacao,
            })

    return pd.DataFrame(linhas)

def processar_consumo_itens_ciclo(bytes_arquivo, xls, params, contexto_contratual, ciclos_padronizados, origem_ciclos, aviso_base_execucao=""):
    """Processa exclusivamente o Modo Consumo por Itens/Ciclo.

    Este modo é diferente do Modo Reduzido por Itens/Estoque. Aqui a base é a
    quantidade consumida/executada por item e por ciclo, validada pela fiscalização.
    """
    ciclos_apurados, fatores = _ler_ciclos_apurados_consumo(bytes_arquivo, xls)
    if ciclos_apurados.empty:
        ciclos_apurados = ciclos_padronizados.copy() if isinstance(ciclos_padronizados, pd.DataFrame) else pd.DataFrame()
        if isinstance(ciclos_apurados, pd.DataFrame) and not ciclos_apurados.empty and "Fator acumulado" in ciclos_apurados.columns:
            fatores = {normalizar_ciclo(r.get("Ciclo", "")): numero_seguro(r.get("Fator acumulado", 1.0), 1.0) for _, r in ciclos_apurados.iterrows()}
            fatores.setdefault("C0", 1.0)
    itens_consumo, ciclos_consumo = _ler_consumo_itens_matriz(bytes_arquivo, xls)

    ciclos_validos = [c for c in ciclos_consumo if c in fatores]
    if not ciclos_validos:
        ciclos_validos = ciclos_consumo
    ultimo_ciclo = sorted([c for c in fatores if c and c != "C0"], key=numero_ciclo)[-1] if any(c != "C0" for c in fatores) else "C0"
    fator_saldo = numero_seguro(fatores.get(ultimo_ciclo, 1.0), 1.0)

    linhas_consumo = []
    linhas_saldo = []
    for item in itens_consumo:
        item_id = item["Item"]
        qtd_contratada = numero_seguro(item["Quantidade contratada"], 0.0)
        vu_original = numero_seguro(item["Valor unitário original/base"], 0.0)
        consumos = item["Consumos"]
        total_consumido = sum(numero_seguro(v, 0.0) for v in consumos.values())
        saldo_qtd = item.get("Saldo a faturar informado")
        if saldo_qtd is None or str(saldo_qtd).strip() == "":
            saldo_qtd = qtd_contratada - total_consumido
        saldo_qtd = numero_seguro(saldo_qtd, 0.0)

        for ciclo in ciclos_consumo:
            qtd_consumida = numero_seguro(consumos.get(ciclo, 0.0), 0.0)
            fator = numero_seguro(fatores.get(ciclo, 1.0), 1.0)
            vu_atualizado = round(vu_original * fator, 2)
            valor_original = round(qtd_consumida * vu_original, 2)
            valor_atualizado = round(qtd_consumida * vu_atualizado, 2)
            retroativo = round(valor_atualizado - valor_original, 2)
            linhas_consumo.append({
                "Item": item_id,
                "Ciclo": ciclo,
                "Quantidade consumida": round(qtd_consumida, 6),
                "Valor unitário original/base": round(vu_original, 2),
                "Fator acumulado": fator,
                "Valor unitário atualizado": vu_atualizado,
                "Valor original consumido": valor_original,
                "Valor atualizado consumido": valor_atualizado,
                "Retroativo": retroativo,
            })

        vu_saldo_atualizado = round(vu_original * fator_saldo, 2)
        saldo_original = round(saldo_qtd * vu_original, 2)
        saldo_atualizado = round(saldo_qtd * vu_saldo_atualizado, 2)
        linhas_saldo.append({
            "Item": item_id,
            "Quantidade contratada": round(qtd_contratada, 6),
            "Quantidade consumida total": round(total_consumido, 6),
            "Quantidade saldo a faturar": round(saldo_qtd, 6),
            "Valor unitário original/base": round(vu_original, 2),
            "Fator saldo atual": fator_saldo,
            "Valor unitário atualizado saldo": vu_saldo_atualizado,
            "Saldo original": saldo_original,
            "Saldo a faturar atualizado": saldo_atualizado,
        })

    df_consumo_itemizado = pd.DataFrame(linhas_consumo)
    df_saldo_itemizado = pd.DataFrame(linhas_saldo)
    df_valores_unitarios_consumo = construir_valores_unitarios_consumo_ciclo(itens_consumo, fatores, ciclos_consumo)

    df_retroativo_por_ciclo = (
        df_consumo_itemizado.groupby("Ciclo", as_index=False)
        .agg({
            "Quantidade consumida": "sum",
            "Valor original consumido": "sum",
            "Valor atualizado consumido": "sum",
            "Retroativo": "sum",
        })
        .sort_values(by="Ciclo", key=lambda s: s.map(numero_ciclo))
        .reset_index(drop=True)
    )
    df_retroativo_por_ciclo["Fator acumulado"] = df_retroativo_por_ciclo["Ciclo"].map(lambda c: fatores.get(c, 1.0))
    df_execucao = df_retroativo_por_ciclo.rename(columns={
        "Valor original consumido": "Valor executado original",
        "Valor atualizado consumido": "Valor executado atualizado",
        "Retroativo": "Delta da execução",
    })[["Ciclo", "Valor executado original", "Valor executado atualizado", "Delta da execução"]]

    valor_original_contrato = round(sum(numero_seguro(i["Quantidade contratada"], 0.0) * numero_seguro(i["Valor unitário original/base"], 0.0) for i in itens_consumo), 2)
    execucao_original = round(df_consumo_itemizado["Valor original consumido"].sum(), 2)
    execucao_atualizada = round(df_consumo_itemizado["Valor atualizado consumido"].sum(), 2)
    retroativo_total = round(df_consumo_itemizado["Retroativo"].sum(), 2)
    saldo_original = round(df_saldo_itemizado["Saldo original"].sum(), 2)
    saldo_atualizado = round(df_saldo_itemizado["Saldo a faturar atualizado"].sum(), 2)
    valor_atualizado_contrato = round(execucao_atualizada + saldo_atualizado, 2)

    df_delta_por_ciclo = df_retroativo_por_ciclo.rename(columns={
        "Valor original consumido": "Valor pago efetivo",
        "Valor atualizado consumido": "Valor teórico calculado",
        "Retroativo": "Delta do ciclo",
    })[["Ciclo", "Valor pago efetivo", "Valor teórico calculado", "Delta do ciclo", "Fator acumulado"]]

    df_retro_compat = df_execucao.copy()
    df_retro_compat["Retroativo estimado por itens/estoque"] = df_retro_compat["Delta da execução"]
    df_retro_compat["Natureza da apuração"] = "Consumo por Itens/Ciclo"
    df_retro_compat["Observação"] = "Apuração itemizada por quantidade consumida/executada por ciclo."

    if isinstance(df_valores_unitarios_consumo, pd.DataFrame) and not df_valores_unitarios_consumo.empty:
        df_remanescentes = (
            df_valores_unitarios_consumo.groupby("Ciclo", as_index=False)
            .agg({"Total R$": "sum"})
            .rename(columns={"Total R$": "Remanescente atualizado"})
            .sort_values(by="Ciclo", key=lambda s: s.map(numero_ciclo))
            .reset_index(drop=True)
        )
        df_remanescentes["Remanescente original"] = pd.NA
        df_remanescentes["Fator aplicado"] = df_remanescentes["Ciclo"].map(lambda c: fatores.get(c, 1.0))
        df_remanescentes["Observação"] = "Remanescente estimado no início do ciclo, calculado a partir das quantidades consumidas informadas."
        df_remanescentes = df_remanescentes[["Ciclo", "Remanescente original", "Remanescente atualizado", "Fator aplicado", "Observação"]]
    else:
        df_remanescentes = pd.DataFrame([{
            "Ciclo": ultimo_ciclo,
            "Remanescente original": saldo_original,
            "Remanescente atualizado": saldo_atualizado,
            "Fator aplicado": fator_saldo,
        }])

    df_composicao_valor_total = montar_composicao_valor_total(
        df_execucao,
        saldo_atualizado,
        ultimo_ciclo,
        valor_atualizado_contrato,
    )

    fator_acumulado = fator_saldo
    indice = (
        st.session_state.get("dados_admissibilidade", {}).get("indice")
        or params.get("indice_utilizado")
        or params.get("indice")
        or params.get("indice_apurado_na_etapa_1")
        or params.get("indice_contratual")
        or params.get("indice_aplicado")
        or "Não informado no modelo"
    )

    df_comparativo = montar_comparativo_executivo(
        valor_original_contrato,
        valor_original_contrato,
        valor_atualizado_contrato - valor_original_contrato,
        0.0,
        0.0,
        0.0,
        saldo_original,
        saldo_atualizado,
        valor_atualizado_contrato,
        0.0,
        0.0,
        0,
        0,
        0.0,
    )

    resultado = {
        "data_processamento": agora_brasilia().strftime("%d/%m/%Y %H:%M"),
        "modo_apuracao": "Consumo por Itens/Ciclo",
        "base_execucao_mensal_disponivel": False,
        "base_itens_disponivel": True,
        "base_consumo_itens_ciclo_disponivel": True,
        "aviso_base_execucao": aviso_base_execucao,
        "ressalva_modo_apuracao": "Apuração por consumo itemizado por ciclo, sem base mensal financeira por competência. A força da apuração depende da premissa fiscal de equivalência entre consumo, medição/aprovação e faturamento devido.",
        "origem_ciclos": origem_ciclos,
        "indice": indice,
        "fator_acumulado": fator_acumulado,
        "variacao_acumulada": fator_acumulado - 1,
        "quantidade_ciclos": int(len([c for c in fatores if c != "C0"])),
        "valor_original_contrato": valor_original_contrato,
        "contexto_contratual_anterior": contexto_contratual,
        "valor_formalizado_anterior": valor_original_contrato,
        "impacto_analise_atual": valor_atualizado_contrato - valor_original_contrato,
        "valor_pago_efetivo": 0.0,
        "total_pago_faturado": 0.0,
        "valor_teorico_calculado": 0.0,
        "total_devido_reajustado": 0.0,
        "delta_total": 0.0,
        "delta_acumulado": 0.0,
        "valor_represado_a_pagar": 0.0,
        "valor_retroativo_estimado_itens_estoque": retroativo_total,
        "valor_retroativo_consumo_itens_ciclo": retroativo_total,
        "retroativo_estimado_itens_estoque_disponivel": True,
        "quantidade_meses_sem_efeito_financeiro": 0,
        "valor_total_sem_efeito_financeiro": 0.0,
        "remanescente_original": saldo_original,
        "remanescente_reajustado": saldo_atualizado,
        "saldo_a_faturar_atualizado": saldo_atualizado,
        "fator_remanescente": fator_saldo,
        "valor_executado_original": execucao_original,
        "valor_executado_atualizado": execucao_atualizada,
        "valor_calculado_sem_aditivos": valor_atualizado_contrato,
        "valor_atualizado_contrato": valor_atualizado_contrato,
        "valor_global_financeiro": valor_atualizado_contrato,
        "total_aditivos_atualizados": 0.0,
        "total_aditivos_informativos": 0.0,
        "aditivos_somados_ao_valor_total": bool(aditivos_somados_ao_valor_total),
        "formula_valor_total_atualizado": "execucao_atualizada_por_ciclo + saldo_a_faturar_atualizado",
        "quantidade_aditivos_total": 0,
        "quantidade_aditivos_marcados_computaveis": 0,
        "ciclo_ultimo_remanescente": ultimo_ciclo,
        # No Modo Consumo por Itens/Ciclo, a fonte correta é a aba CICLOS_APURADOS.
        # O leitor genérico pode captar linhas de orientação como ciclos vazios.
        "df_ciclos": ciclos_apurados if isinstance(ciclos_apurados, pd.DataFrame) and not ciclos_apurados.empty else ciclos_padronizados,
        "df_financeiro_mensal": pd.DataFrame(columns=["Ciclo", "Competência", "Valor pago/faturado"]),
        "df_financeiro_mensal_tratado": pd.DataFrame(),
        "df_meses_sem_efeito_financeiro": pd.DataFrame(),
        "df_financeiro_por_ciclo": pd.DataFrame(),
        "df_delta_por_ciclo": df_delta_por_ciclo,
        "df_execucao_atualizada": df_execucao,
        "df_consumo_itemizado_ciclo": df_consumo_itemizado,
        "df_retroativo_itemizado_por_ciclo": df_retroativo_por_ciclo,
        "df_saldo_itemizado": df_saldo_itemizado,
        "df_retroativo_estimado_itens_estoque": df_retro_compat,
        "df_composicao_valor_total": df_composicao_valor_total,
        "df_remanescentes": df_remanescentes,
        "df_valores_unitarios_ciclo": df_valores_unitarios_consumo,
        "df_aditivos": pd.DataFrame(),
        "df_aditivos_executivo": pd.DataFrame(),
        "df_aditivos_computaveis": pd.DataFrame(),
        "df_aditivos_informativos": pd.DataFrame(),
        "df_comparativo": df_comparativo,
    }
    resultado = arredondar_resultado_financeiro(resultado)
    resultado["df_auditoria_consistencia"] = montar_auditoria_consistencia(resultado)
    return resultado


def ler_ciclo_em_execucao_config_segura(bytes_arquivo):
    # Leitura interna e segura da aba opcional CICLO_EM_EXECUCAO.
    # Quando o corte operacional está como Não/vazio, a aba não altera o cálculo padrão.
    try:
        xls_local = pd.ExcelFile(BytesIO(bytes_arquivo))
        aba = localizar_aba(xls_local, ["CICLO_EM_EXECUCAO", "CICLO EM EXECUCAO", "CICLO_EM_EXECUÇÃO"])
        if not aba:
            return {"existe": False, "aplicar": False}

        df = ler_aba_com_cabecalho(bytes_arquivo, aba, termos_obrigatorios=["Campo", "Valor"])
        if df.empty or len(df.columns) < 2:
            return {"existe": True, "aplicar": False}

        col_campo = localizar_coluna(df, ["Campo", "Parâmetro", "Parametro"])
        col_valor = localizar_coluna(df, ["Valor"])
        if col_campo is None or col_valor is None:
            col_campo, col_valor = df.columns[0], df.columns[1]

        dados = {}
        for _, row in df.iterrows():
            campo = str(row.get(col_campo, "")).strip()
            chave = normalizar_texto(campo)
            if chave:
                dados[chave] = row.get(col_valor, "")

        aplicar_txt = (
            dados.get("aplicar_corte_operacional")
            or dados.get("aplicar_corte_operacional_sim_nao")
            or dados.get("aplicar_corte_operacional_no_ciclo_em_execucao")
            or ""
        )
        aplicar = normalizar_texto(aplicar_txt) in ["sim", "s", "true", "1", "yes"]

        valor_c0_manual = dados.get(
            "valor_financeiro_c0_manual_override",
            dados.get("valor_financeiro_c0_manual", "")
        )
        usa_c0_manual = abs(numero_br(valor_c0_manual)) > 0.004

        valor_corte = (
            dados.get("competencia_de_corte_operacional")
            or dados.get("competencia_corte_operacional")
            or dados.get("data_de_corte_operacional")
            or dados.get("data_corte_operacional")
            or ""
        )
        periodo_corte = normalizar_competencia_periodo(valor_corte)
        competencia_corte_label = periodo_para_label_br(periodo_corte) if periodo_corte is not None else texto_seguro(valor_corte, "")

        rem_original = (
            dados.get("valor_remanescente_original_no_corte_operacional")
            or dados.get("valor_remanescente_original_nominal_no_corte_operacional")
            or dados.get("valor_remanescente_original_corte_operacional")
            or dados.get("remanescente_original_no_corte_operacional")
            or ""
        )
        rem_atualizado = dados.get(
            "valor_remanescente_atualizado_no_corte_operacional",
            dados.get("valor_remanescente_atualizado_corte_operacional", dados.get("remanescente_atualizado_no_corte_operacional", ""))
        )

        return {
            "existe": True,
            "aplicar": aplicar,
            "ciclo": texto_seguro(dados.get("ciclo_em_execucao", ""), ""),
            "competencia_corte": competencia_corte_label,
            "data_corte": competencia_corte_label,
            "periodo_corte": periodo_corte,
            "fonte": texto_seguro(dados.get("fonte_da_execucao_realizada", dados.get("fonte_preferencial_da_execucao_realizada", "")), ""),
            "usar_c0_manual": "Sim" if usa_c0_manual else "Não",
            "valor_c0_manual": valor_c0_manual,
            "valor_remanescente_original_corte": rem_original,
            "valor_remanescente_atualizado_corte": rem_atualizado,
            "observacao": texto_seguro(dados.get("observacao_fiscal", ""), ""),
        }
    except Exception as exc:
        return {"existe": False, "aplicar": False, "erro": str(exc)}

def _corte_operacional_ativo_seguro(config):
    return bool(config and isinstance(config, dict) and config.get("aplicar"))


def filtrar_financeiro_por_corte_operacional(financeiro, config):
    # Corte financeiro: considera a base mensal até a competência de corte, inclusive.
    if financeiro is None or not isinstance(financeiro, pd.DataFrame) or financeiro.empty:
        return financeiro
    if not _corte_operacional_ativo_seguro(config):
        return financeiro

    periodo_corte = config.get("periodo_corte") or normalizar_competencia_periodo(config.get("competencia_corte", config.get("data_corte", "")))
    if periodo_corte is None or "Competência" not in financeiro.columns:
        return financeiro

    periodos = financeiro["Competência"].apply(normalizar_competencia_periodo)
    filtrado = financeiro[(periodos.notna()) & (periodos <= periodo_corte)].copy()
    if filtrado.empty:
        return financeiro.iloc[0:0].copy()
    return filtrado.reset_index(drop=True)

def aplicar_corte_operacional_execucao_v1(df_execucao, df_fin_por_ciclo, config, base_execucao_mensal_disponivel):
    # Regra v2:
    # - corte desativado: não altera a execução;
    # - corte ativado com base financeira: execução realizada vem da base financeira até a competência de corte;
    # - C0 manual é usado automaticamente se houver valor preenchido maior que zero.
    if not _corte_operacional_ativo_seguro(config):
        return df_execucao

    df_base = df_execucao.copy() if isinstance(df_execucao, pd.DataFrame) else pd.DataFrame()

    if base_execucao_mensal_disponivel and isinstance(df_fin_por_ciclo, pd.DataFrame) and not df_fin_por_ciclo.empty:
        linhas = []
        for _, row in df_fin_por_ciclo.iterrows():
            ciclo = normalizar_ciclo(row.get("Ciclo", ""))
            if not ciclo or ciclo.upper() == "TOTAL":
                continue
            pago = numero_seguro(row.get("Valor pago efetivo", 0.0), 0.0)
            teorico = numero_seguro(row.get("Valor teórico calculado", pago), pago)
            pct = ((teorico / pago) - 1) if abs(pago) > 0.004 else 0.0
            linhas.append({
                "Ciclo": ciclo,
                "Status financeiro": "Corte operacional — execução realizada pela base financeira até a competência de corte",
                "Valor executado original": float(pago),
                "Percentual acumulado aplicado": float(pct),
                "Valor executado atualizado": float(teorico),
            })
        if linhas:
            df_base = pd.DataFrame(linhas)

    valor_c0 = numero_br(config.get("valor_c0_manual", 0))
    usar_c0 = abs(valor_c0) > 0.004
    if usar_c0:
        if df_base.empty:
            df_base = pd.DataFrame(columns=[
                "Ciclo", "Status financeiro", "Valor executado original",
                "Percentual acumulado aplicado", "Valor executado atualizado"
            ])
        if "Ciclo" in df_base.columns:
            mask_c0 = df_base["Ciclo"].astype(str).str.upper().eq("C0")
        else:
            mask_c0 = pd.Series([False] * len(df_base))
        linha_c0 = {
            "Ciclo": "C0",
            "Status financeiro": "C0 - execução financeira manual informado na aba CICLO_EM_EXECUCAO",
            "Valor executado original": float(valor_c0),
            "Percentual acumulado aplicado": 0.0,
            "Valor executado atualizado": float(valor_c0),
        }
        if mask_c0.any():
            for col, val in linha_c0.items():
                df_base.loc[mask_c0, col] = val
        else:
            df_base = pd.concat([pd.DataFrame([linha_c0]), df_base], ignore_index=True)

    if not df_base.empty and "Ciclo" in df_base.columns:
        df_base = df_base.sort_values(by="Ciclo", key=lambda s: s.map(numero_ciclo)).reset_index(drop=True)
    return df_base

def validar_config_corte_operacional(config):
    if not _corte_operacional_ativo_seguro(config):
        return
    periodo_corte = config.get("periodo_corte") or normalizar_competencia_periodo(config.get("competencia_corte", config.get("data_corte", "")))
    if periodo_corte is None:
        raise ValueError(
            "Corte operacional incompleto: informe a Competência de corte operacional. "
            "Exemplos válidos: 04/2026, 01/04/2026 ou 30/04/2026."
        )
    rem_original = numero_br(config.get("valor_remanescente_original_corte", 0))
    rem_atualizado = numero_br(config.get("valor_remanescente_atualizado_corte", 0))
    if abs(rem_original) <= 0.004 and abs(rem_atualizado) <= 0.004:
        raise ValueError(
            "Corte operacional incompleto: informe o Valor remanescente original no corte operacional "
            "ou o Valor remanescente atualizado no corte operacional. O sistema não deduz saldo futuro automaticamente."
        )


def ajustar_remanescente_por_corte_operacional(df_rem, df_execucao, valor_original_contrato, fator_remanescente, ciclo_ultimo_rem, config, base_execucao_mensal_disponivel):
    # Ajusta o saldo remanescente no corte operacional somente quando houver informação expressa.
    # Não deduz automaticamente a execução financeira do saldo.
    if not _corte_operacional_ativo_seguro(config):
        return df_rem, ciclo_ultimo_rem, None, None

    validar_config_corte_operacional(config)

    valor_original_informado = numero_br(config.get("valor_remanescente_original_corte", 0))
    valor_atualizado_informado = numero_br(config.get("valor_remanescente_atualizado_corte", 0))

    fator = float(fator_remanescente or 1.0)
    if abs(valor_atualizado_informado) > 0.004:
        remanescente_atualizado_operacional = float(valor_atualizado_informado)
        if abs(valor_original_informado) > 0.004:
            remanescente_original_operacional = float(valor_original_informado)
        else:
            remanescente_original_operacional = float(valor_atualizado_informado) / fator if abs(fator) > 0.004 else float(valor_atualizado_informado)
    else:
        remanescente_original_operacional = float(valor_original_informado)
        remanescente_atualizado_operacional = remanescente_original_operacional * fator

    ciclo_corte = normalizar_ciclo(config.get("ciclo", "")) or ciclo_ultimo_rem or "Ciclo em execução"
    ciclo_label = f"{ciclo_corte} (corte operacional)"
    comp = config.get("competencia_corte") or config.get("data_corte") or ""
    linha = {
        "Ciclo": ciclo_label,
        "Remanescente original": float(remanescente_original_operacional),
        "Fator aplicado": fator,
        "Remanescente atualizado": float(remanescente_atualizado_operacional),
        "Observação": f"Saldo informado expressamente na aba CICLO_EM_EXECUCAO para a competência de corte {comp}.",
    }

    df_rem_ajustado = pd.DataFrame([linha])
    return df_rem_ajustado, ciclo_label, remanescente_original_operacional, remanescente_atualizado_operacional

def processar_arquivo_coleta(bytes_arquivo):
    """LEGADO — isolado do fluxo oficial, sem chamadores.

    Calculava a apuração em Python a partir de arquivos não canônicos. O Painel
    passou a admitir exclusivamente o Coleta_Reajuste.xlsx oficial, cujo
    resultado vem do próprio XLS; manter dois motores permitiria divergência
    entre o que o arquivo diz e o que a tela mostra. Preservada aqui apenas até
    a remoção controlada, junto de suas sub-rotinas (entre elas
    `processar_consumo_itens_ciclo`). Não religar ao upload.
    """

    aditivos_somados_ao_valor_total = 0.0  # fallback: planilha sem aditivos computaveis
    xls = pd.ExcelFile(BytesIO(bytes_arquivo))
    params = ler_parametros(bytes_arquivo, xls)
    contexto_contratual = contexto_contratual_de_parametros(params, bytes_arquivo, xls)
    ciclos, origem_ciclos = ler_ciclos(bytes_arquivo, xls)
    config_ciclo_em_execucao = ler_ciclo_em_execucao_config_segura(bytes_arquivo)
    corte_operacional_solicitado = bool(config_ciclo_em_execucao.get('aplicar', False))
    validar_config_corte_operacional(config_ciclo_em_execucao)

    try:
        financeiro = ler_financeiro(bytes_arquivo, xls, ciclos)
        base_execucao_mensal_disponivel = not financeiro.empty
        aviso_base_execucao = ""
    except Exception as exc:
        financeiro = pd.DataFrame(columns=["Ciclo", "Competência", "Valor pago/faturado"])
        base_execucao_mensal_disponivel = False
        aviso_base_execucao = str(exc)

    # Modo Consumo por Itens/Ciclo: terceiro modo, distinto do modo reduzido por remanescentes.
    # Deve ser priorizado antes de ler ITENS_REMANESCENTES, porque CONSUMO_ITENS também contém
    # dados de itens e não deve ser interpretada como estoque/remanescente.
    if (not base_execucao_mensal_disponivel) and _aba_exata(xls, "CONSUMO_ITENS"):
        return processar_consumo_itens_ciclo(
            bytes_arquivo, xls, params, contexto_contratual, ciclos, origem_ciclos, aviso_base_execucao
        )

    itens, colunas_remanescentes = ler_itens(bytes_arquivo, xls)
    base_itens_disponivel = (not itens.empty) and bool(colunas_remanescentes)

    if not base_execucao_mensal_disponivel and not base_itens_disponivel:
        raise ValueError("Processamento inviável: não foram localizadas base de execução mensal nem informações suficientes de itens/remanescentes.")

    if base_execucao_mensal_disponivel and base_itens_disponivel:
        modo_apuracao = "Completo"
    elif base_execucao_mensal_disponivel:
        modo_apuracao = "Financeiro/Base de execução mensal"
    else:
        modo_apuracao = "Reduzido por Itens/Estoque"

    ressalva_modo_apuracao = ""
    if modo_apuracao == "Reduzido por Itens/Estoque":
        ressalva_modo_apuracao = (
            "A apuração foi realizada sem base de execução mensal por competência. "
            "Os resultados financeiros possuem natureza estimativa, calculada a partir de itens/remanescentes, "
            "e devem ser validados antes da formalização de pagamento."
        )

    # Remove linhas de total ou vazias da base de itens.
    if "Item" in itens.columns:
        itens = itens[~itens["Item"].astype(str).str.strip().str.upper().eq("TOTAL")].copy()
        itens = itens[itens["Item"].astype(str).str.strip() != ""].copy()

    if ciclos.empty:
        perc_legacy = extrair_percentual_reajuste_legacy(bytes_arquivo, xls)
        fator = fator_de_valor(params.get("fator_acumulado_total") or params.get("fator_acumulado_final") or params.get("fator_acumulado") or params.get("fator"))
        if fator == 1.0 and perc_legacy is not None:
            fator = 1 + perc_legacy
        ciclos = pd.DataFrame([{
            "Ciclo": "C1",
            "Data-base": "",
            "Intervalo do índice": "",
            "Janela de admissibilidade": "",
            "Data do pedido": "",
            "Situação": "",
            "Tratamento financeiro do ciclo": "A apurar",
            "Variação": fator - 1,
            "Fator": fator,
            "Fator acumulado": fator,
            "Fator acumulado efetivo": fator,
            "Fator ciclo efetivo": fator,
        }])

    # Base mensal tratada pela regra pétrea de efeito financeiro.
    # Competências anteriores ao início financeiro do ciclo permanecem em memória,
    # mas não compõem o retroativo a pagar.
    financeiro_para_calculo = filtrar_financeiro_por_corte_operacional(financeiro, config_ciclo_em_execucao)
    df_financeiro_mensal_tratado = financeiro_com_efeito_financeiro(financeiro_para_calculo, ciclos)
    df_meses_sem_efeito = meses_sem_efeito_financeiro(financeiro_para_calculo, ciclos)
    quantidade_meses_sem_efeito = int(len(df_meses_sem_efeito)) if isinstance(df_meses_sem_efeito, pd.DataFrame) else 0
    valor_total_sem_efeito = float(df_meses_sem_efeito["Delta não devido"].apply(numero_seguro).sum()) if quantidade_meses_sem_efeito and "Delta não devido" in df_meses_sem_efeito.columns else 0.0

    df_fin_por_ciclo = calcular_financeiro_por_ciclo(financeiro_para_calculo, ciclos)
    df_rem, ciclo_ultimo_rem = calcular_remanescentes_valor(itens, colunas_remanescentes, ciclos)
    df_execucao = calcular_execucao_por_diferenca(itens, colunas_remanescentes, ciclos)
    df_execucao = aplicar_corte_operacional_execucao_v1(
        df_execucao,
        df_fin_por_ciclo,
        config_ciclo_em_execucao,
        base_execucao_mensal_disponivel,
    )
    df_retroativo_estimado_itens = montar_retroativo_estimado_por_itens(df_execucao)
    valor_retroativo_estimado_itens = float(df_retroativo_estimado_itens["Retroativo estimado por itens/estoque"].apply(numero_seguro).sum()) if not df_retroativo_estimado_itens.empty else 0.0
    df_valores_unitarios = construir_valores_unitarios_totais(itens, colunas_remanescentes, ciclos)
    df_aditivos = ler_aditivos(bytes_arquivo, xls, ciclos)
    df_aditivos_executivo = consolidar_aditivos_executivo(df_aditivos)

    valor_original_referencia = itens.attrs.get("valor_original_contrato_referencia") if hasattr(itens, "attrs") else None
    if valor_original_referencia is not None and float(valor_original_referencia) > 0:
        valor_original_contrato = float(valor_original_referencia)
    else:
        valor_original_contrato = float(itens["Valor total original"].sum()) if "Valor total original" in itens.columns else 0.0
    total_pago_efetivo = float(df_fin_por_ciclo.loc[df_fin_por_ciclo["Ciclo"] != "TOTAL", "Valor pago efetivo"].sum()) if not df_fin_por_ciclo.empty else 0.0
    total_teorico_calculado = float(df_fin_por_ciclo.loc[df_fin_por_ciclo["Ciclo"] != "TOTAL", "Valor teórico calculado"].sum()) if not df_fin_por_ciclo.empty else 0.0
    delta_total = float(df_fin_por_ciclo.loc[df_fin_por_ciclo["Ciclo"] != "TOTAL", "Delta do ciclo"].sum()) if not df_fin_por_ciclo.empty else 0.0
    if not df_fin_por_ciclo.empty:
        df_rep = df_fin_por_ciclo[df_fin_por_ciclo["Ciclo"] != "TOTAL"].copy()
        if "Tratamento financeiro" in df_rep.columns:
            mask_apurar = ~df_rep["Tratamento financeiro"].astype(str).apply(lambda x: tratamento_sem_retroativo(x, ""))
            df_rep = df_rep[mask_apurar]
        valor_represado_a_pagar = float(df_rep["Delta do ciclo"].clip(lower=0).sum()) if "Delta do ciclo" in df_rep.columns else 0.0
    else:
        valor_represado_a_pagar = 0.0

    if not df_rem.empty:
        remanescente_original = float(df_rem.iloc[-1]["Remanescente original"])
        remanescente_atualizado = float(df_rem.iloc[-1]["Remanescente atualizado"])
        fator_remanescente = float(df_rem.iloc[-1]["Fator aplicado"])
    else:
        remanescente_original = 0.0
        remanescente_atualizado = 0.0
        fator_remanescente = 1.0

    df_rem, ciclo_ultimo_rem, remanescente_original_corte, remanescente_atualizado_corte = ajustar_remanescente_por_corte_operacional(
        df_rem,
        df_execucao,
        valor_original_contrato,
        fator_remanescente,
        ciclo_ultimo_rem,
        config_ciclo_em_execucao,
        base_execucao_mensal_disponivel,
    )
    if remanescente_original_corte is not None:
        remanescente_original = float(remanescente_original_corte)
        remanescente_atualizado = float(remanescente_atualizado_corte)

    total_execucao_atualizada = float(df_execucao["Valor executado atualizado"].sum()) if not df_execucao.empty else max(valor_original_contrato - remanescente_original, 0.0)
    valor_calculado_sem_aditivos = total_execucao_atualizada + remanescente_atualizado

    if not df_aditivos.empty and "Computa no Valor Global" in df_aditivos.columns:
        df_aditivos_computaveis = df_aditivos[df_aditivos["Computa no Valor Global"].astype(bool)].copy()
        df_aditivos_informativos = df_aditivos[~df_aditivos["Computa no Valor Global"].astype(bool)].copy()
    else:
        df_aditivos_computaveis = df_aditivos.copy() if not df_aditivos.empty else pd.DataFrame()
        df_aditivos_informativos = pd.DataFrame()

    total_aditivos_atualizados = float(df_aditivos_computaveis["Valor atualizado da alteração"].sum()) if not df_aditivos_computaveis.empty else 0.0
    total_aditivos_informativos = float(df_aditivos_informativos["Valor atualizado da alteração"].sum()) if not df_aditivos_informativos.empty else 0.0

    valor_formalizado_anterior = float(contexto_contratual.get("valor_formalizado_anterior") or 0.0)
    if valor_formalizado_anterior <= 0:
        valor_formalizado_anterior = valor_original_contrato

    # Regra conceitual do Valor Total Atualizado:
    # o valor final é composto por execução atualizada por ciclo + saldo remanescente atualizado
    # + aditivos/supressões computáveis quando indicados como parcela da análise atual.
    # Aditivos informativos/já incorporados permanecem rastreáveis, mas não entram como parcela autônoma.
    aditivos_somados_ao_valor_total = abs(numero_seguro(total_aditivos_atualizados, 0.0)) > 0.004
    valor_atualizado_contrato = valor_calculado_sem_aditivos + (
        total_aditivos_atualizados if aditivos_somados_ao_valor_total else 0.0
    )
    impacto_analise_atual = valor_atualizado_contrato - valor_formalizado_anterior

    df_composicao_valor_total = montar_composicao_valor_total(
        df_execucao,
        remanescente_atualizado,
        ciclo_ultimo_rem,
        valor_atualizado_contrato,
        df_aditivos_computaveis=df_aditivos_computaveis,
        total_aditivos_atualizados=total_aditivos_atualizados,
    )

    fator_acumulado = float(ciclos["Fator acumulado efetivo"].iloc[-1]) if "Fator acumulado efetivo" in ciclos.columns and not ciclos.empty else fator_remanescente
    indice = (
        st.session_state.get("dados_admissibilidade", {}).get("indice")
        or params.get("indice_utilizado")
        or params.get("indice")
        or "Não informado"
    )

    df_comparativo = montar_comparativo_executivo(
        valor_original_contrato,
        valor_formalizado_anterior,
        impacto_analise_atual,
        total_pago_efetivo,
        total_teorico_calculado,
        delta_total,
        remanescente_original,
        remanescente_atualizado,
        valor_atualizado_contrato,
        total_aditivos_atualizados,
        total_aditivos_informativos,
        len(df_aditivos_executivo),
        quantidade_meses_sem_efeito,
        valor_total_sem_efeito,
    )

    df_delta_por_ciclo = montar_delta_por_ciclo(df_fin_por_ciclo, df_execucao, ciclos)

    df_aditivos = limpar_nan_inf_df(df_aditivos)
    df_aditivos_executivo = limpar_nan_inf_df(df_aditivos_executivo)
    df_valores_unitarios = limpar_nan_inf_df(df_valores_unitarios)
    df_delta_por_ciclo = limpar_nan_inf_df(df_delta_por_ciclo)

    resultado = {
        "data_processamento": agora_brasilia().strftime("%d/%m/%Y %H:%M"),
        "modo_apuracao": modo_apuracao,
        "base_execucao_mensal_disponivel": bool(base_execucao_mensal_disponivel),
        "base_itens_disponivel": bool(base_itens_disponivel),
        "aviso_base_execucao": aviso_base_execucao,
        "ressalva_modo_apuracao": ressalva_modo_apuracao,
        "config_ciclo_em_execucao": config_ciclo_em_execucao,
        "corte_operacional_solicitado": bool(corte_operacional_solicitado),
        "corte_operacional_aplicado": bool(corte_operacional_solicitado),
        "metodologia_corte_operacional": "financeiro_preferencial_ate_corte_e_c0_manual_quando_informado" if corte_operacional_solicitado else "",
        "origem_ciclos": origem_ciclos,
        "indice": indice,
        "fator_acumulado": fator_acumulado,
        "variacao_acumulada": fator_acumulado - 1,
        "quantidade_ciclos": int(len(ciclos)),
        "valor_original_contrato": valor_original_contrato,
        "contexto_contratual_anterior": contexto_contratual,
        "valor_formalizado_anterior": valor_formalizado_anterior,
        "impacto_analise_atual": impacto_analise_atual,
        "valor_pago_efetivo": total_pago_efetivo,
        "total_pago_faturado": total_pago_efetivo,
        "valor_teorico_calculado": total_teorico_calculado,
        "total_devido_reajustado": total_teorico_calculado,
        "delta_total": delta_total,
        "delta_acumulado": delta_total,
        "valor_represado_a_pagar": round(valor_represado_a_pagar, 2),
        "valor_retroativo_estimado_itens_estoque": round(valor_retroativo_estimado_itens, 2),
        "retroativo_estimado_itens_estoque_disponivel": bool(abs(valor_retroativo_estimado_itens) > 0.004),
        "quantidade_meses_sem_efeito_financeiro": quantidade_meses_sem_efeito,
        "valor_total_sem_efeito_financeiro": round(valor_total_sem_efeito, 2),
        "remanescente_original": round(remanescente_original, 2),
        "remanescente_reajustado": round(remanescente_atualizado, 2),
        "remanescente_recalculado_por_corte_operacional": bool(remanescente_original_corte is not None),
        "fator_remanescente": fator_remanescente,
        "valor_executado_atualizado": round(total_execucao_atualizada, 2),
        "valor_calculado_sem_aditivos": round(valor_calculado_sem_aditivos, 2),
        "valor_atualizado_contrato": round(valor_atualizado_contrato, 2),
        "valor_global_financeiro": valor_atualizado_contrato,
        "total_aditivos_atualizados": total_aditivos_atualizados,
        "total_aditivos_informativos": total_aditivos_informativos,
        "aditivos_somados_ao_valor_total": False,
        "formula_valor_total_atualizado": "execucao_atualizada_por_ciclo + saldo_remanescente_atualizado + aditivos_computaveis_atualizados",
        "quantidade_aditivos_total": int(len(df_aditivos_executivo)),
        "quantidade_aditivos_marcados_computaveis": int(len(df_aditivos_executivo[df_aditivos_executivo["Computa no Valor Global"].astype(bool)])) if not df_aditivos_executivo.empty and "Computa no Valor Global" in df_aditivos_executivo.columns else 0,
        "ciclo_ultimo_remanescente": ciclo_ultimo_rem,
        "df_ciclos": ciclos,
        "df_financeiro_mensal": financeiro,
        "df_financeiro_mensal_corte_operacional": financeiro_para_calculo,
        "df_financeiro_mensal_tratado": df_financeiro_mensal_tratado,
        "df_meses_sem_efeito_financeiro": df_meses_sem_efeito,
        "df_financeiro_por_ciclo": df_fin_por_ciclo,
        "df_delta_por_ciclo": df_delta_por_ciclo,
        "df_execucao_atualizada": df_execucao,
        "df_retroativo_estimado_itens_estoque": df_retroativo_estimado_itens,
        "df_composicao_valor_total": df_composicao_valor_total,
        "df_remanescentes": df_rem,
        "df_valores_unitarios_ciclo": df_valores_unitarios,
        "df_aditivos": df_aditivos,
        "df_aditivos_executivo": df_aditivos_executivo,
        "df_aditivos_computaveis": df_aditivos_computaveis,
        "df_aditivos_informativos": df_aditivos_informativos,
        "df_comparativo": df_comparativo,
    }
    resultado = arredondar_resultado_financeiro(resultado)
    resultado["df_auditoria_consistencia"] = montar_auditoria_consistencia(resultado)
    return resultado


def formatar_dataframe_moeda(df, colunas_moeda=None, colunas_fator=None, colunas_pct=None, colunas_data=None):
    visual = limpar_nan_inf_df(df).copy()
    for col in colunas_moeda or []:
        if col in visual.columns:
            visual[col] = visual[col].apply(moeda)
    for col in colunas_fator or []:
        if col in visual.columns:
            visual[col] = visual[col].apply(fator_fmt)
    for col in colunas_pct or []:
        if col in visual.columns:
            visual[col] = visual[col].apply(lambda x: percentual(x, 2))
    for col in colunas_data or []:
        if col in visual.columns:
            visual[col] = visual[col].apply(formatar_data_br)
    return visual


def formatar_data_br(valor):
    data = pd.to_datetime(valor, dayfirst=True, errors="coerce")
    if pd.isna(data):
        return "" if pd.isna(valor) else str(valor)
    return data.strftime("%d/%m/%Y")



# ============================================================
# Linha do Tempo do Contrato
# ============================================================

def _data_evento(valor):
    """Converte valores diversos em Timestamp normalizado para a timeline."""
    data = pd.to_datetime(valor, dayfirst=True, errors="coerce")
    if pd.isna(data):
        return pd.NaT
    return data.normalize()


def _texto_evento(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    return str(valor).strip()


def _limpar_marcadores_timeline(valor):
    """Remove marcadores informais/emojis para manter a timeline executiva."""
    texto = _texto_evento(valor)
    for marcador in ["✅", "❌", "⚠️", "⚠", "🟡", "🔴", "🟢", "🔵"]:
        texto = texto.replace(marcador, "")
    return " ".join(texto.split())


def _adicionar_evento_timeline(eventos, data, tipo, titulo, detalhe="", ciclo="", valor=None, prioridade=50):
    data_ts = _data_evento(data)
    if pd.isna(data_ts):
        return
    eventos.append({
        "Data": data_ts,
        "Tipo": tipo,
        "Evento": titulo,
        "Ciclo": ciclo,
        "Detalhe": detalhe,
        "Valor": valor,
        "Prioridade": prioridade,
    })


def montar_eventos_linha_tempo(resultado):
    """Consolida eventos contratuais para a Linha do Tempo do Contrato.

    A função usa apenas dados já disponíveis no módulo Valor Global:
    ciclos, pedidos, efeitos financeiros, aditivos/supressões e acordos negociais.
    Quando campos como fim da vigência ou assinatura não estiverem disponíveis,
    eles simplesmente não são exibidos, preservando o funcionamento atual.
    """
    eventos = []
    ciclos = resultado.get("df_ciclos", pd.DataFrame())
    aditivos = resultado.get("df_aditivos", pd.DataFrame())
    contexto = resultado.get("contexto_contratual_anterior", {}) or {}

    # Marco inicial obrigatório da análise: data-base para reajuste informada no Simples ou Múltiplos.
    # Esse marco é exibido de forma própria, porque representa a origem temporal do processamento.
    if isinstance(contexto, dict) and contexto.get("data_base_reajuste"):
        _adicionar_evento_timeline(
            eventos,
            contexto.get("data_base_reajuste"),
            "Data-base para reajuste",
            "Data-base para reajuste",
            "Marco inicial informado na Calculadora de Reajuste para apuração dos ciclos.",
            prioridade=1,
        )

    # Marcos contratuais opcionais, quando futuramente estiverem disponíveis no contexto/arquivo.
    for chave, titulo in [
        ("data_assinatura", "Assinatura do contrato"),
        ("inicio_vigencia", "Início da vigência"),
        ("fim_vigencia", "Fim da vigência"),
    ]:
        if isinstance(contexto, dict) and contexto.get(chave):
            tipo = "Fim da vigência" if chave == "fim_vigencia" else "Marco contratual"
            _adicionar_evento_timeline(eventos, contexto.get(chave), tipo, titulo, prioridade=5)

    for evento in (contexto.get("eventos_historicos_anteriores", []) if isinstance(contexto, dict) else []):
        if not isinstance(evento, dict):
            continue
        tipo_evento = str(evento.get("Tipo de evento", "Histórico anterior") or "Histórico anterior").strip()
        ciclo_evento = str(evento.get("Ciclo", "") or "").strip()
        incorporado = str(evento.get("Incorporado ao valor formalizado?", "") or "").strip()
        obs_evento = str(evento.get("Observação", "") or "").strip()
        valor_evento = numero_br(evento.get("Valor formalizado/impacto", 0))
        detalhe = "Evento histórico anterior informado no Contexto do Contrato."
        if ciclo_evento:
            detalhe += f" Ciclo: {ciclo_evento}."
        if incorporado:
            detalhe += f" Incorporado ao valor formalizado: {incorporado}."
        if obs_evento:
            detalhe += f" Observação: {obs_evento}"
        _adicionar_evento_timeline(
            eventos,
            evento.get("Data", ""),
            "Histórico anterior",
            tipo_evento,
            detalhe,
            ciclo_evento,
            valor_evento if abs(valor_evento) > 0.004 else None,
            prioridade=6,
        )

    if isinstance(ciclos, pd.DataFrame) and not ciclos.empty:
        ciclos_ord = ciclos.sort_values(by="Ciclo", key=lambda s: s.map(numero_ciclo)) if "Ciclo" in ciclos.columns else ciclos
        for _, row in ciclos_ord.iterrows():
            ciclo = normalizar_ciclo(row.get("Ciclo", ""))
            if not ciclo:
                continue
            situacao = _texto_evento(row.get("Situação", ""))
            situacao_auto = _texto_evento(row.get("Situação automática", situacao))
            acordo = _texto_evento(row.get("Acordo negocial", ""))
            perc = percentual_para_decimal(row.get("Percentual aplicado", row.get("Variação", 0)))
            situacao_timeline = _limpar_marcadores_timeline(situacao_auto or situacao)
            detalhe_base = f"{ciclo} | {situacao_timeline} | percentual aplicado: {percentual(perc, 2)}"

            _adicionar_evento_timeline(
                eventos,
                row.get("Data-base", ""),
                "Ciclo de reajuste",
                f"{ciclo} - Data-base",
                detalhe_base,
                ciclo,
                prioridade=10 + numero_ciclo(ciclo),
            )
            _adicionar_evento_timeline(
                eventos,
                row.get("Data do pedido", ""),
                "Pedido de reajuste",
                f"{ciclo} - Pedido de reajuste",
                f"Pedido relacionado ao {ciclo}. Resultado automático: {situacao_timeline}.",
                ciclo,
                prioridade=20 + numero_ciclo(ciclo),
            )
            _adicionar_evento_timeline(
                eventos,
                row.get("Início financeiro", ""),
                "Efeito financeiro",
                f"{ciclo} - Início dos efeitos financeiros",
                f"Efeitos financeiros do {ciclo}. Tratamento: {_texto_evento(row.get('Tratamento financeiro do ciclo', '')) or 'A apurar'}.",
                ciclo,
                prioridade=30 + numero_ciclo(ciclo),
            )
            if normalizar_texto(acordo) in ["sim", "s", "true", "1", "yes"]:
                _adicionar_evento_timeline(
                    eventos,
                    row.get("Início financeiro", row.get("Data do pedido", "")),
                    "Acordo negocial",
                    f"{ciclo} - Acordo negocial de admissão de reajuste",
                    f"Ciclo tecnicamente classificado como {situacao_timeline}, mas admitido por negociação entre as partes.",
                    ciclo,
                    prioridade=35 + numero_ciclo(ciclo),
                )

    if isinstance(aditivos, pd.DataFrame) and not aditivos.empty:
        aditivos_temp = aditivos.copy()
        if "Data do aditivo" in aditivos_temp.columns:
            aditivos_temp["_data_evento"] = aditivos_temp["Data do aditivo"].apply(_data_evento)
        else:
            aditivos_temp["_data_evento"] = pd.NaT
        aditivos_temp = aditivos_temp.dropna(subset=["_data_evento"]).copy()

        if not aditivos_temp.empty:
            def valores_aditivo(nome_coluna, padrao):
                if nome_coluna not in aditivos_temp.columns:
                    return [padrao] * len(aditivos_temp.index)
                coluna = aditivos_temp.loc[:, nome_coluna]
                if isinstance(coluna, pd.DataFrame):
                    coluna = coluna.iloc[:, 0]
                return coluna.tolist()

            aditivos_temp["_eh_supressao"] = [
                "supress" in normalizar_texto(valor) or "decresc" in normalizar_texto(valor)
                for valor in valores_aditivo("Tipo de alteração", "Aditivo")
            ]
            aditivos_temp["_eh_acrescimo"] = ~aditivos_temp["_eh_supressao"]
            aditivos_temp["_ciclo"] = [
                normalizar_ciclo(valor) for valor in valores_aditivo("Ciclo/Marco", "")
            ]
            aditivos_temp["_tratamento"] = [
                _texto_evento(valor) for valor in valores_aditivo("Tratamento do aditivo", "")
            ]
            aditivos_temp["_valor"] = [
                numero_seguro(valor, 0.0) for valor in valores_aditivo("Valor atualizado da alteração", 0.0)
            ]

            agrupados = (
                aditivos_temp
                .groupby(["_data_evento", "_ciclo"], dropna=False)
                .agg(
                    quantidade=("_valor", "size"),
                    quantidade_acrescimos=("_eh_acrescimo", "sum"),
                    quantidade_supressoes=("_eh_supressao", "sum"),
                    valor_total=("_valor", "sum"),
                    tratamentos=("_tratamento", lambda x: " / ".join(sorted({str(v).strip() for v in x if str(v).strip()}))),
                )
                .reset_index()
            )

            for _, row in agrupados.iterrows():
                tipo_evento = "Aditivo"
                ciclo = normalizar_ciclo(row.get("_ciclo", ""))
                tratamento = _texto_evento(row.get("tratamentos", "")) or "Computar nesta análise"
                qtd = int(numero_seguro(row.get("quantidade", 0), 0))
                qtd_ad = int(numero_seguro(row.get("quantidade_acrescimos", 0), 0))
                qtd_sup = int(numero_seguro(row.get("quantidade_supressoes", 0), 0))
                valor = numero_seguro(row.get("valor_total", 0.0), 0.0)
                partes_qtd = []
                if qtd_ad:
                    partes_qtd.append(f"{qtd_ad} itens ad.")
                if qtd_sup:
                    partes_qtd.append(f"{qtd_sup} itens exc.")
                titulo = "Aditivo"
                if partes_qtd:
                    titulo += " - " + " / ".join(partes_qtd)
                elif qtd > 1:
                    titulo += f" - {qtd} itens"
                elif qtd == 1:
                    titulo += " - 1 item"
                detalhe = f"Aditivo consolidado. Ciclo/Marco: {ciclo or 'não identificado'}. Tratamento: {tratamento}."
                _adicionar_evento_timeline(
                    eventos,
                    row.get("_data_evento", ""),
                    tipo_evento,
                    titulo,
                    detalhe,
                    ciclo,
                    valor=valor,
                    prioridade=40,
                )

    if not eventos:
        return pd.DataFrame(columns=["Data", "Tipo", "Evento", "Ciclo", "Detalhe", "Valor"])

    df = pd.DataFrame(eventos)
    df = df.dropna(subset=["Data"]).copy()
    df = df.sort_values(["Data", "Prioridade", "Evento"]).reset_index(drop=True)
    return df


def _cor_tipo_timeline(tipo):
    mapa = {
        "Data-base para reajuste": "#0B1F3A",
        "Marco contratual": "#475569",
        "Fim da vigência": "#334155",
        "Ciclo de reajuste": "#1F4E78",
        "Pedido de reajuste": "#5B9BD5",
        "Efeito financeiro": "#2E7D32",
        "Aditivo": "#C97A20",
        "Supressão": "#B42318",
        "Acordo negocial": "#6B5B95",
        "Outros reajustes": "#0F766E",
    }
    return mapa.get(tipo, "#64748B")


def render_linha_tempo_contrato(resultado):
    df_eventos = montar_eventos_linha_tempo(resultado)
    if df_eventos.empty:
        st.info("Ainda não há eventos suficientes para montar a Linha do Tempo do Contrato.")
        return

    cards = []
    for _, ev in df_eventos.iterrows():
        tipo = _limpar_marcadores_timeline(ev.get("Tipo", "Evento")) or "Evento"
        cor = _cor_tipo_timeline(tipo)
        data_txt = pd.to_datetime(ev["Data"]).strftime("%d/%m/%Y")
        valor_txt = ""
        if ev.get("Valor", None) is not None and abs(numero_seguro(ev.get("Valor"), 0.0)) > 0:
            valor_txt = f"<br><strong>Valor:</strong> {html.escape(moeda(ev.get('Valor')))}"
        detalhe = _limpar_marcadores_timeline(ev.get("Detalhe", ""))
        evento = _limpar_marcadores_timeline(ev.get("Evento", "Evento"))
        cards.append(
            "".join([
                '<div class="telebras-timeline-event">',
                f'<div class="telebras-timeline-dot" style="background:{cor};"></div>',
                f'<div class="telebras-timeline-card" style="border-top:4px solid {cor};">',
                f'<div class="telebras-timeline-date">{html.escape(data_txt)}</div>',
                f'<div class="telebras-timeline-type" style="color:{cor};">{html.escape(tipo)}</div>',
                f'<div class="telebras-timeline-event-title">{html.escape(evento)}</div>',
                f'<div class="telebras-timeline-detail">{html.escape(detalhe)}{valor_txt}</div>',
                '</div>',
                '</div>',
            ])
        )

    tipos_ordenados = ["Data-base para reajuste", "Ciclo de reajuste", "Pedido de reajuste", "Efeito financeiro", "Aditivo", "Supressão", "Acordo negocial", "Fim da vigência", "Marco contratual"]
    tipos_presentes = [t for t in tipos_ordenados if t in set(df_eventos["Tipo"].astype(str))]
    legend = "".join(
        f"<span class='telebras-legend-item'><span class='telebras-legend-dot' style='background:{_cor_tipo_timeline(t)};'></span>{html.escape(t)}</span>"
        for t in tipos_presentes
    )

    timeline_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        body {{ margin:0; font-family: Arial, Helvetica, sans-serif; background: transparent; }}
        .telebras-timeline-shell {{
            background:#F8FAFC;
            border:1px solid #E1E7EF;
            border-radius:16px;
            padding:18px 18px 14px 18px;
            box-sizing:border-box;
        }}
        .telebras-timeline-title {{
            color:#0B1F3A;
            font-size:18px;
            font-weight:800;
            margin-bottom:4px;
        }}
        .telebras-timeline-subtitle {{
            color:#64748B;
            font-size:13px;
            margin-bottom:14px;
        }}
        .telebras-timeline-track {{
            position:relative;
            display:flex;
            gap:16px;
            overflow-x:auto;
            overflow-y:hidden;
            padding:18px 4px 12px 4px;
            scroll-behavior:smooth;
        }}
        .telebras-timeline-track:before {{
            content:"";
            position:absolute;
            top:25px;
            left:12px;
            right:12px;
            height:2px;
            background:#D8E2EC;
            z-index:0;
        }}
        .telebras-timeline-event {{
            min-width:190px;
            max-width:230px;
            position:relative;
            z-index:1;
            flex:0 0 auto;
        }}
        .telebras-timeline-dot {{
            width:13px;
            height:13px;
            border-radius:50%;
            border:3px solid #FFFFFF;
            box-shadow:0 0 0 1px rgba(15, 23, 42, 0.12);
            margin:0 0 8px 10px;
        }}
        .telebras-timeline-card {{
            background:#FFFFFF;
            border:1px solid #E5EAF0;
            border-radius:12px;
            padding:10px 11px;
            box-shadow:0 1px 2px rgba(15, 23, 42, 0.04);
            min-height:112px;
            box-sizing:border-box;
        }}
        .telebras-timeline-date {{ color:#475569; font-size:12px; font-weight:700; margin-bottom:5px; }}
        .telebras-timeline-event-title {{ color:#0F172A; font-size:14px; font-weight:800; line-height:1.25; margin-bottom:6px; }}
        .telebras-timeline-type {{ font-size:11px; font-weight:700; line-height:1.1; margin-bottom:6px; }}
        .telebras-timeline-detail {{ color:#64748B; font-size:12px; line-height:1.28; }}
        .telebras-timeline-legend {{
            display:flex;
            flex-wrap:wrap;
            gap:8px 14px;
            margin-top:10px;
            color:#475569;
            font-size:12px;
        }}
        .telebras-legend-item {{ display:flex; align-items:center; gap:6px; }}
        .telebras-legend-dot {{ width:9px; height:9px; border-radius:50%; display:inline-block; }}
    </style>
    </head>
    <body>
        <div class="telebras-timeline-shell">
            <div class="telebras-timeline-title">Linha do Tempo do Contrato</div>
            <div class="telebras-timeline-subtitle">Visão executiva dos marcos de reajuste, pedidos, efeitos financeiros, aditivos e acordos negociais.</div>
            <div class="telebras-timeline-track">{''.join(cards)}</div>
            <div class="telebras-timeline-legend">{legend}</div>
        </div>
    </body>
    </html>
    """
    components.html(timeline_html, height=365, width=1200, scrolling=True)

    with st.expander("Ver eventos da linha do tempo em tabela"):
        df_visual = df_eventos[["Data", "Tipo", "Evento", "Ciclo", "Detalhe", "Valor"]].copy()
        df_visual["Data"] = df_visual["Data"].apply(formatar_data_br)
        for col in ["Tipo", "Evento", "Detalhe"]:
            if col in df_visual.columns:
                df_visual[col] = df_visual[col].apply(_limpar_marcadores_timeline)
        if "Valor" in df_visual.columns:
            df_visual["Valor"] = df_visual["Valor"].apply(lambda x: moeda(x) if numero_seguro(x, 0.0) != 0 else "")
        st.dataframe(df_visual, use_container_width=True, hide_index=True)



def _paragrafo_pdf(valor, estilo):
    texto = _limpar_marcadores_timeline(valor)
    texto = html.escape(texto).replace("\n", "<br/>")
    return Paragraph(texto or "-", estilo)


def gerar_pdf_linha_tempo_contrato(resultado):
    """Gera relatório executivo em PDF da Linha do Tempo do Contrato."""
    if not REPORTLAB_OK:
        return None

    df_eventos = montar_eventos_linha_tempo(resultado)
    if df_eventos.empty:
        return None

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=1.2 * cm,
        leftMargin=1.2 * cm,
        topMargin=1.0 * cm,
        bottomMargin=1.0 * cm,
    )

    styles = getSampleStyleSheet()
    titulo = ParagraphStyle(
        "telebrasTituloTimeline",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontSize=15,
        leading=18,
        textColor=colors.HexColor("#0B1F3A"),
        spaceAfter=4,
    )
    subtitulo = ParagraphStyle(
        "telebrasSubtituloTimeline",
        parent=styles["Normal"],
        alignment=TA_CENTER,
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#475569"),
        spaceAfter=10,
    )
    h2 = ParagraphStyle(
        "telebrasH2Timeline",
        parent=styles["Heading2"],
        fontSize=10.5,
        leading=13,
        textColor=colors.HexColor("#123B63"),
        spaceBefore=6,
        spaceAfter=5,
    )
    normal = ParagraphStyle(
        "telebrasNormalTimeline",
        parent=styles["Normal"],
        fontSize=8,
        leading=10.5,
        textColor=colors.HexColor("#1F2937"),
    )
    celula = ParagraphStyle(
        "telebrasCelulaTimeline",
        parent=styles["Normal"],
        fontSize=6.7,
        leading=8.2,
        textColor=colors.HexColor("#1F2937"),
        alignment=TA_LEFT,
    )
    celula_branca = ParagraphStyle(
        "telebrasCelulaBrancaTimeline",
        parent=celula,
        textColor=colors.white,
    )

    elementos = []
    elementos.append(Paragraph("Mapa dos Marcos Contratuais", titulo))
    elementos.append(Paragraph("Linha do Tempo do Contrato", subtitulo))
    elementos.append(Paragraph(f"Gerado em: {agora_brasilia().strftime('%d/%m/%Y %H:%M')}", subtitulo))

    # Resumo executivo do relatório.
    resumo = [
        ["Indicador", "Valor", "Indicador", "Valor"],
        ["Índice", str(resultado.get("indice", "-")), "Ciclos", str(resultado.get("quantidade_ciclos", "-"))],
        ["Reajuste acumulado", percentual(resultado.get("variacao_acumulada", 0.0), 2), "Valor Total Atualizado do Contrato", moeda(resultado.get("valor_atualizado_contrato", 0.0))],
        ["Valor represado a pagar", moeda(resultado.get("valor_represado_a_pagar", 0.0)), "Eventos na timeline", str(len(df_eventos))],
        ["Meses sem efeito financeiro", str(resultado.get("quantidade_meses_sem_efeito_financeiro", 0)), "Valor total sem efeito financeiro", moeda(resultado.get("valor_total_sem_efeito_financeiro", 0.0))],
    ]
    tabela_resumo = Table(resumo, colWidths=[4.2 * cm, 7.0 * cm, 4.8 * cm, 8.0 * cm], repeatRows=1)
    tabela_resumo.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9E2F3")),
        ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F8FAFC")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ]))
    elementos.append(tabela_resumo)
    elementos.append(Spacer(1, 8))

    elementos.append(Paragraph("1. Síntese visual dos eventos", h2))
    contagem = df_eventos["Tipo"].astype(str).value_counts().reset_index()
    contagem.columns = ["Tipo de evento", "Quantidade"]
    dados_contagem = [["Tipo de evento", "Quantidade"]]
    for _, row in contagem.iterrows():
        dados_contagem.append([_paragrafo_pdf(row["Tipo de evento"], celula), str(row["Quantidade"])])
    tabela_contagem = Table(dados_contagem, colWidths=[12.0 * cm, 3.0 * cm], hAlign="LEFT", repeatRows=1)
    tabela_contagem.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAF2F8")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#123B63")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9E2F3")),
        ("FONTSIZE", (0, 0), (-1, -1), 7.3),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
    ]))
    elementos.append(tabela_contagem)
    elementos.append(Spacer(1, 8))

    df_efeitos_pdf = resultado.get("df_meses_sem_efeito_financeiro", pd.DataFrame())
    if isinstance(df_efeitos_pdf, pd.DataFrame) and not df_efeitos_pdf.empty:
        elementos.append(Paragraph("2. Efeitos financeiros sem retroativo", h2))
        elementos.append(Paragraph(
            "As competências abaixo representam o lapso entre a competência em que o reajuste poderia produzir efeitos e o início financeiro decorrente do pedido. O delta teórico é demonstrado, mas não compõe o retroativo a pagar.",
            normal,
        ))
        dados_ef = [[
            _paragrafo_pdf("Ciclo", celula_branca),
            _paragrafo_pdf("Competência", celula_branca),
            _paragrafo_pdf("Valor base", celula_branca),
            _paragrafo_pdf("Fator teórico", celula_branca),
            _paragrafo_pdf("Delta não devido", celula_branca),
        ]]
        for _, row in df_efeitos_pdf.iterrows():
            dados_ef.append([
                _paragrafo_pdf(row.get("Ciclo", ""), celula),
                _paragrafo_pdf(row.get("Competência", ""), celula),
                _paragrafo_pdf(moeda(row.get("Valor base", 0.0)), celula),
                _paragrafo_pdf(fator_fmt(row.get("Fator teórico", 1.0)), celula),
                _paragrafo_pdf(moeda(row.get("Delta não devido", 0.0)), celula),
            ])
        tabela_ef = Table(dados_ef, colWidths=[2.0 * cm, 3.0 * cm, 4.0 * cm, 3.0 * cm, 4.2 * cm], repeatRows=1, hAlign="LEFT")
        tabela_ef.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#9C0006")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#E6B8B7")),
            ("FONTSIZE", (0, 0), (-1, -1), 7.0),
            ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#FCE4D6")),
        ]))
        elementos.append(tabela_ef)
        elementos.append(Spacer(1, 8))
        numero_linha_tempo = "3"
    else:
        numero_linha_tempo = "2"

    elementos.append(Paragraph(f"{numero_linha_tempo}. Linha do tempo executiva", h2))
    elementos.append(Paragraph(
        "Os eventos abaixo consolidam os marcos de reajuste, pedidos, efeitos financeiros, aditivos, supressões e acordos negociais identificados no processamento.",
        normal,
    ))
    elementos.append(Spacer(1, 5))

    dados = [[
        _paragrafo_pdf("Data", celula_branca),
        _paragrafo_pdf("Tipo", celula_branca),
        _paragrafo_pdf("Evento", celula_branca),
        _paragrafo_pdf("Ciclo", celula_branca),
        _paragrafo_pdf("Detalhe", celula_branca),
        _paragrafo_pdf("Valor", celula_branca),
    ]]

    df_pdf = df_eventos.copy()
    for _, ev in df_pdf.iterrows():
        valor = numero_seguro(ev.get("Valor"), 0.0)
        dados.append([
            _paragrafo_pdf(pd.to_datetime(ev["Data"]).strftime("%d/%m/%Y"), celula),
            _paragrafo_pdf(ev.get("Tipo", ""), celula),
            _paragrafo_pdf(ev.get("Evento", ""), celula),
            _paragrafo_pdf(ev.get("Ciclo", ""), celula),
            _paragrafo_pdf(ev.get("Detalhe", ""), celula),
            _paragrafo_pdf(moeda(valor) if abs(valor) > 0 else "", celula),
        ])

    tabela = Table(dados, colWidths=[2.2 * cm, 3.5 * cm, 5.0 * cm, 1.7 * cm, 10.0 * cm, 3.3 * cm], repeatRows=1)
    estilo = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9E2F3")),
        ("FONTSIZE", (0, 0), (-1, -1), 6.7),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ])
    for i in range(1, len(dados)):
        bg = colors.HexColor("#FFFFFF") if i % 2 else colors.HexColor("#F8FAFC")
        estilo.add("BACKGROUND", (0, i), (-1, i), bg)
        tipo_txt = _limpar_marcadores_timeline(df_pdf.iloc[i-1].get("Tipo", ""))
        cor = _cor_tipo_timeline(tipo_txt)
        estilo.add("TEXTCOLOR", (1, i), (1, i), colors.HexColor(cor))
        estilo.add("FONTNAME", (1, i), (1, i), "Helvetica-Bold")
    tabela.setStyle(estilo)
    elementos.append(tabela)

    elementos.append(Spacer(1, 8))
    elementos.append(Paragraph(
        "Observação: este relatório é executivo e reflete os dados disponíveis no processamento do Valor Global. Eventos históricos não informados no arquivo ou no contexto do contrato não serão exibidos.",
        normal,
    ))

    doc.build(elementos)
    buffer.seek(0)
    return buffer.getvalue()

def aplicar_css_responsivo_telebras():
    """Ajusta KPIs/cards para reduzir truncamento em telas menores."""
    st.markdown(
        """
        <style>
        div[data-testid="stMetric"] {
            background: #FFFFFF;
            border: 1px solid #E5EAF0;
            border-radius: 12px;
            padding: 10px 12px;
            min-height: 84px;
        }
        div[data-testid="stMetricLabel"] p {
            color: #475569;
            font-size: clamp(0.74rem, 1.2vw, 0.92rem);
            line-height: 1.2;
            white-space: normal;
            word-break: normal;
        }
        div[data-testid="stMetricValue"] {
            font-size: clamp(1.0rem, 1.8vw, 1.55rem);
            line-height: 1.2;
            white-space: normal;
            overflow-wrap: anywhere;
        }
        div[data-testid="stMetricDelta"] {
            font-size: 0.78rem;
        }
        .telebras-valor-destaque {
            background:#EAF2F8;
            border:1px solid #BFD7EA;
            border-radius:14px;
            padding:18px 22px;
            margin:12px 0 18px 0;
        }
        .telebras-valor-destaque-label {
            font-size:0.95rem;
            color:#27496D;
            font-weight:600;
        }
        .telebras-valor-destaque-valor {
            font-size:clamp(1.35rem, 2.8vw, 2.05rem);
            color:#0B1F3A;
            font-weight:800;
            line-height:1.25;
            word-break:break-word;
        }
        div[data-testid="stTabs"] div[role="tablist"] {
            overflow-x: auto;
            overflow-y: hidden;
            white-space: nowrap;
            scrollbar-width: thin;
            gap: 0.25rem;
        }
        div[data-testid="stTabs"] button[role="tab"] {
            flex: 0 0 auto;
            padding-left: 0.65rem;
            padding-right: 0.65rem;
            font-size: 0.92rem;
        }
        div[data-testid="stTabs"] div[role="tablist"]::-webkit-scrollbar {
            height: 6px;
        }
        div[data-testid="stTabs"] div[role="tablist"]::-webkit-scrollbar-thumb {
            background: #CBD5E1;
            border-radius: 999px;
        }
        [data-testid="stFileUploaderDropzone"] button {
            background: #FFFFFF !important;
            color: #C56A00 !important;
            border: 1px solid #A855F7 !important;
            border-radius: 8px !important;
            font-weight: 600 !important;
        }
        [data-testid="stFileUploaderDropzone"] button:hover {
            background: #FFF7ED !important;
            border-color: #EA580C !important;
            color: #9A3412 !important;
        }
        [data-testid="stFileUploaderDropzoneInstructions"] div:first-child {
            font-size: 0 !important;
        }
        [data-testid="stFileUploaderDropzoneInstructions"] div:first-child::after {
            content: "Arraste e solte o arquivo aqui";
            font-size: 0.95rem;
            color: #334155;
            font-weight: 500;
        }
        [data-testid="stFileUploaderDropzoneInstructions"] div:nth-child(2) {
            font-size: 0 !important;
        }
        [data-testid="stFileUploaderDropzoneInstructions"] div:nth-child(2)::after {
            content: "Limite 200 MB por arquivo • XLSX";
            font-size: 0.82rem;
            color: #64748B;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )



def _formatar_data_corte_br(valor):
    """Formata a data do corte operacional para dd/mm/aaaa, preservando texto inválido."""
    try:
        if valor is None:
            return "Não informada"
        texto = str(valor).strip()
        if texto == "" or texto.lower() in ["nan", "none", "nat"]:
            return "Não informada"
        data = pd.to_datetime(valor, dayfirst=True, errors="coerce")
        if pd.isna(data):
            return texto
        return data.strftime("%d/%m/%Y")
    except Exception:
        return str(valor or "Não informada")


def formatar_data_br_corte_operacional(valor):
    if valor is None:
        return "Não informada"
    texto = str(valor).strip()
    if texto.lower() in ["", "nan", "none", "nat", "<na>"]:
        return "Não informada"
    try:
        data = pd.to_datetime(valor, dayfirst=True, errors="coerce")
        if not pd.isna(data):
            return data.strftime("%d/%m/%Y")
    except Exception:
        pass
    return texto

def render_metodologia_corte_operacional_v3(resultado):
    # Aviso visual da metodologia usada no Valor Global.
    config = resultado.get('config_ciclo_em_execucao', {}) or {}
    solicitado = bool(resultado.get('corte_operacional_solicitado', False) or resultado.get('corte_operacional_aplicado', False))

    if not solicitado:
        st.markdown(
            '<div style="background:#EAF2F8; border:1px solid #93C5FD; border-left:6px solid #2563EB; '
            'border-radius:12px; padding:12px 14px; margin:8px 0 16px 0; color:#1E3A8A;">'
            '<div style="font-weight:800; margin-bottom:3px;">Metodologia aplicada: corte padrão no início dos ciclos</div>'
            '<div style="font-size:0.92rem; line-height:1.42;">'
            'A aba CICLO_EM_EXECUCAO está ausente, vazia ou marcada como Não. '
            'O Valor Total Atualizado permanece calculado pela metodologia padrão.'
            '</div></div>',
            unsafe_allow_html=True,
        )
        return

    ciclo = str(config.get('ciclo', '') or 'Não informado')
    data_corte = formatar_data_br_corte_operacional(config.get('data_corte', ''))
    fonte = str(config.get('fonte', '') or 'Não informada')
    usar_c0 = str(config.get('usar_c0_manual', '') or 'Não')
    valor_c0 = config.get('valor_c0_manual', '')
    rem_original = config.get('remanescente_original_corte', config.get('valor_remanescente_original_corte_operacional', ''))
    rem_atualizado = config.get('remanescente_atualizado_corte', config.get('valor_remanescente_atualizado_corte_operacional', ''))

    def _fmt_moeda_seguro(valor):
        try:
            if str(valor).strip() == '':
                return 'Não informado'
            return moeda(numero_br(valor))
        except Exception:
            return str(valor or 'Não informado')

    st.markdown(
        '<div style="background:#CCFBF1; border:1px solid #14B8A6; border-left:6px solid #0F766E; '
        'border-radius:12px; padding:12px 14px; margin:8px 0 16px 0; color:#134E4A;">'
        '<div style="font-weight:900; margin-bottom:4px;">Metodologia aplicada: corte operacional no ciclo em execução</div>'
        '<div style="font-size:0.92rem; line-height:1.45;">'
        f'<b>Ciclo em execução:</b> {ciclo}<br>'
        f'<b>Competência de corte:</b> {_formatar_data_corte_br(data_corte)}<br>'
        f'<b>Fonte da execução realizada:</b> {fonte}<br>'
        f'<b>Remanescente original informado no corte:</b> {_fmt_moeda_seguro(rem_original)}<br>'
        f'<b>Remanescente atualizado informado no corte:</b> {_fmt_moeda_seguro(rem_atualizado)}<br>'
        f'<b>C0 financeiro manual:</b> {usar_c0} — {_fmt_moeda_seguro(valor_c0)}'
        '</div></div>',
        unsafe_allow_html=True,
    )


def render_metodologia_corte_operacional_v4(resultado, modo_apuracao_ui="Completo"):
    """Renderiza box metodológico no Painel Executivo, sem alterar cálculo."""
    try:
        aplicado = bool(resultado.get("corte_operacional_aplicado", False))
        config = resultado.get("config_ciclo_em_execucao", {}) or {}
        existe_config = bool(config.get("existe", False))

        def _valor_config(*chaves):
            for chave in chaves:
                valor = config.get(chave)
                if valor is not None and str(valor).strip() not in ["", "nan", "None"]:
                    return valor
            return ""

        def _moeda_config(valor):
            try:
                return moeda(numero_br(valor)) if str(valor).strip() else "Não informado"
            except Exception:
                return str(valor or "Não informado")

        if aplicado:
            ciclo = _valor_config("ciclo", "ciclo_em_execucao") or "Não informado"
            data_corte = formatar_data_br_corte_operacional(_valor_config("data_corte", "data_de_corte_operacional", "data_corte_operacional"))
            fonte = _valor_config("fonte", "fonte_preferencial_da_execucao_realizada") or "Não informada"
            rem_original = _valor_config("valor_remanescente_original_corte", "remanescente_original_corte", "valor_remanescente_original_no_corte_operacional")
            rem_atualizado = _valor_config("valor_remanescente_atualizado_corte", "remanescente_atualizado_corte", "valor_remanescente_atualizado_no_corte_operacional")
            c0_manual = _valor_config("usar_c0_manual", "usar_c0_financeiro_manual") or "Não"
            valor_c0 = _valor_config("valor_c0_manual", "valor_financeiro_c0_manual_override", "valor_financeiro_c0_manual")

            html = (
                '<div style="background:#CCFBF1; border:1px solid #14B8A6; border-left:7px solid #0F766E; border-radius:12px; padding:14px 16px; margin:8px 0 16px 0; color:#134E4A;">'
                '<div style="font-weight:900; font-size:0.98rem; margin-bottom:6px;">Metodologia aplicada: corte operacional no ciclo em execução</div>'
                '<div style="font-size:0.92rem; line-height:1.45;">'
                f'<b>Ciclo em execução:</b> {ciclo}<br>'
                f'<b>Competência de corte:</b> {_formatar_data_corte_br(data_corte)}<br>'
                f'<b>Fonte da execução realizada:</b> {fonte}<br>'
                f'<b>Remanescente original no corte:</b> {_moeda_config(rem_original)}<br>'
                f'<b>Remanescente atualizado no corte:</b> {_moeda_config(rem_atualizado)}<br>'
                f'<b>C0 financeiro manual:</b> {c0_manual} — {_moeda_config(valor_c0)}'
                '</div>'
                '<div style="font-size:0.86rem; margin-top:8px; color:#0F766E;">Composição: execução atualizada até o corte + saldo remanescente informado no corte operacional.</div>'
                '</div>'
            )
            st.markdown(html, unsafe_allow_html=True)
        else:
            complemento = "A aba CICLO_EM_EXECUCAO foi detectada, mas o campo principal está como Não/vazio." if existe_config else "A aba CICLO_EM_EXECUCAO não foi utilizada nesta apuração."
            html = (
                '<div style="background:#EFF6FF; border:1px solid #93C5FD; border-left:7px solid #2563EB; border-radius:12px; padding:13px 16px; margin:8px 0 16px 0; color:#1E3A8A;">'
                '<div style="font-weight:900; font-size:0.98rem; margin-bottom:5px;">Metodologia aplicada: corte padrão no início dos ciclos</div>'
                f'<div style="font-size:0.92rem; line-height:1.45;">Modo de apuração: {modo_apuracao_ui}. {complemento} O Valor Total Atualizado permanece calculado por execução atualizada por ciclo + saldo remanescente atualizado.</div>'
                '</div>'
            )
            st.markdown(html, unsafe_allow_html=True)
    except Exception:
        render_metodologia_corte_operacional_v4(resultado, modo_apuracao_ui)


# ============================================================
# Interface
# ============================================================

aplicar_css_responsivo_telebras()
render_cabecalho_pagina(
    "Painel da Apuração Contratual",
    "Envie o Coleta_Reajuste.xlsx preenchido para validar cada bloco, acompanhar os resultados disponíveis e gerar documentos.",
)

st.markdown(
    '<div class="cl8us-docs-note">O Coleta_Reajuste.xlsx reúne os dados da apuração. '
    'A web valida a estrutura e aproveita, de forma independente, cada bloco seguro da apuração.</div>',
    unsafe_allow_html=True,
)

with st.container(border=True):
    st.markdown(
        '<span class="cl8us-docs-card-marker"></span>'
        '<div class="cl8us-docs-card-title">1 · Baixar arquivo de trabalho</div>',
        unsafe_allow_html=True,
    )
    st.caption("Use o modelo único com fórmulas para registrar a apuração contratual.")
    if CAMINHO_MODELO_COLETA.exists():
        st.download_button(
            "Baixar Coleta_Reajuste.xlsx",
            data=CAMINHO_MODELO_COLETA.read_bytes(),
            file_name=NOME_ARQUIVO_COLETA,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="download_coleta_documentos",
        )
    else:
        st.error("O modelo Coleta_Reajuste.xlsx não foi localizado.")

with st.container(border=True):
    st.markdown(
        '<span class="cl8us-docs-card-marker"></span>'
        '<div class="cl8us-docs-card-title">2 · Enviar Coleta_Reajuste.xlsx preenchido</div>',
        unsafe_allow_html=True,
    )
    arquivo = st.file_uploader(
        "Selecione o arquivo .xlsx preenchido",
        type=["xlsx"],
        key="upload_coleta_documentos",
    )

if arquivo is None:
    st.info("Envie o Coleta_Reajuste.xlsx preenchido. Cada bloco informado será processado de forma independente.")
    capacidades_iniciais = avaliar_capacidades_apuracao({}, {})
    render_status_apuracao(capacidades_iniciais)
    render_status_documentos(
        capacidades_iniciais,
        (
            "planilha_executiva", "valores_unitarios", "relatorio_executivo",
            "mapa_marcos", "minuta_apostilamento", "garantia_contratual",
            "dou", "checklist_processual",
        ),
    )
    st.stop()

adm = st.session_state.get("dados_admissibilidade")

with st.expander("Contexto da Admissibilidade", expanded=True):
    if adm:
        col1, col2, col3 = st.columns(3)
        col1.metric("Origem", adm.get("origem") or adm.get("tipo", "Não informado"))
        col2.metric("Índice", adm.get("indice", "Não informado"))
        col3.metric("Ciclos", len(adm.get("ciclos", [])))
        st.caption("Dados herdados da sessão atual. Caso o arquivo contenha parâmetros próprios, eles também serão lidos para conferência.")
    else:
        st.warning(
            "Os dados de admissibilidade não foram encontrados na sessão atual. "
            "A ferramenta utilizará os parâmetros constantes do Arquivo de Coleta."
        )

if arquivo is not None:
    if st.button("Validar Coleta Preenchida", type="primary", use_container_width=False):
        conteudo = arquivo.getvalue()
        sha256 = sha256_do_arquivo(conteudo)
        # Nada do arquivo anterior sobrevive à troca: o estado é apagado antes de
        # qualquer decisão sobre o novo upload.
        if not upload_ja_processado(st.session_state, sha256):
            limpar_estados_derivados(st.session_state)
        try:
            if not eh_coleta_reajuste(conteudo):
                registrar_upload(
                    st.session_state,
                    sha256=sha256,
                    origem=ORIGEM_NAO_RECONHECIDA,
                    aceito=False,
                    motivo=MENSAGEM_ARQUIVO_NAO_RECONHECIDO,
                )
            else:
                diagnostico = ler_coleta_reajuste(conteudo)
                st.session_state["diagnostico_coleta_v2"] = diagnostico
                if diagnostico.get("valido"):
                    resultado = adaptar_coleta_reajuste_para_documentos(conteudo)
                    st.session_state["resultado_valor_global"] = resultado
                    registrar_upload(
                        st.session_state,
                        sha256=sha256,
                        origem=ORIGEM_COLETA_OFICIAL,
                        aceito=True,
                    )
                else:
                    registrar_upload(
                        st.session_state,
                        sha256=sha256,
                        origem=ORIGEM_COLETA_OFICIAL,
                        aceito=False,
                        motivo="A coleta oficial foi reprovada no diagnóstico estrutural.",
                    )
        except Exception as exc:
            # Falha no meio do processamento deixaria resultado e diagnóstico
            # dessincronizados; o estado volta a ser ausência explícita.
            limpar_estados_derivados(st.session_state)
            registrar_upload(
                st.session_state,
                sha256=sha256,
                origem=ORIGEM_NAO_RECONHECIDA,
                aceito=False,
                motivo=f"Não foi possível processar o arquivo: {exc}",
            )

procedencia = procedencia_registrada(st.session_state)
if procedencia and procedencia.get("origem") == ORIGEM_NAO_RECONHECIDA:
    st.error(procedencia.get("motivo") or MENSAGEM_ARQUIVO_NAO_RECONHECIDO)
    for detalhe in DETALHES_ARQUIVO_NAO_RECONHECIDO:
        st.caption(detalhe)
    st.stop()

diagnostico_coleta = st.session_state.get("diagnostico_coleta_v2")
if diagnostico_coleta:
    if diagnostico_coleta.get("valido"):
        st.success("Coleta reconhecida: estrutura e fórmulas essenciais preservadas.")
    else:
        st.error("A coleta não está segura para prosseguir.")

    metadados = diagnostico_coleta.get("metadados", {})
    contagens = diagnostico_coleta.get("contagens", {})
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Índice", metadados.get("indice") or "Não informado")
    col2.metric("Ciclo vigente", metadados.get("ciclo_vigente") or "Não informado")
    col3.metric("Meses com valor", contagens.get("competencias_com_valor", 0))
    col4.metric("Itens remanescentes", contagens.get("itens_remanescentes", 0))

    capacidades_coleta = diagnostico_coleta.get("capacidades") or avaliar_capacidades_apuracao({}, {})
    render_status_apuracao(capacidades_coleta)
    render_status_documentos(
        capacidades_coleta,
        (
            "planilha_executiva", "valores_unitarios", "relatorio_executivo",
            "mapa_marcos", "minuta_apostilamento", "garantia_contratual",
            "dou", "checklist_processual",
        ),
    )

    for bloqueio in diagnostico_coleta.get("bloqueios_estruturais", []):
        st.error(bloqueio)
    for bloqueio in diagnostico_coleta.get("bloqueios_criticos", []):
        st.error(bloqueio)
    for lacuna in diagnostico_coleta.get("lacunas_apuracao", []):
        st.warning(lacuna)
    for aviso in diagnostico_coleta.get("avisos", []):
        st.warning(aviso)

resultado = st.session_state.get("resultado_valor_global")

if resultado:
    with st.container(border=True):
        st.markdown("### Ações sobre os documentos")
        st.caption("A Central mantém todos os documentos visíveis e explica individualmente eventuais dependências.")
        documentos_cap = (resultado.get("capacidades") or {}).get("documentos") or {}
        col_arquivos, col_relatorios, col_gestao = st.columns(3)
        with col_arquivos:
            st.page_link("pages/06_Central_Arquivos.py", label="Abrir Central de Arquivos", use_container_width=True)
        with col_relatorios:
            if (documentos_cap.get("relatorio_executivo") or {}).get("habilitado"):
                st.page_link("pages/04_Relatorio_Global.py", label="Gerar relatório e minuta", use_container_width=True)
            else:
                st.button("Relatório aguardando dados", disabled=True, use_container_width=True)
        with col_gestao:
            if (documentos_cap.get("garantia_contratual") or {}).get("habilitado"):
                st.page_link("pages/05_Garantia.py", label="Gerar garantia contratual", use_container_width=True)
            else:
                st.button("Garantia aguardando VTA", disabled=True, use_container_width=True)

    st.divider()
    render_resultados_progressivos(resultado)
    if not diagnostico_coleta.get("pronto_para_consolidar"):
        st.subheader("Arquivos disponíveis nesta etapa")
        col_planilha, col_itens, col_memoria = st.columns(3)
        with col_planilha:
            if (documentos_cap.get("planilha_executiva") or {}).get("habilitado"):
                try:
                    excel_executivo_parcial = gerar_planilha_executiva(resultado)
                    st.session_state["arquivo_planilha_executiva_xlsx"] = excel_executivo_parcial
                    st.download_button(
                        "Baixar Planilha Executiva",
                        data=excel_executivo_parcial,
                        file_name="Planilha_Executiva_Analise_Reajuste.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="download_planilha_executiva_progressiva",
                    )
                except Exception as exc:
                    st.error(f"Planilha Executiva não pôde ser gerada: {exc}")
            else:
                st.button("Planilha Executiva · Pendente de dados", disabled=True, use_container_width=True)
        with col_itens:
            df_vu_parcial = limpar_nan_inf_df(resultado.get("df_valores_unitarios_ciclo", pd.DataFrame()))
            if (documentos_cap.get("valores_unitarios") or {}).get("habilitado") and not df_vu_parcial.empty:
                excel_vu_parcial = gerar_excel_valores_unitarios_por_ciclo(df_vu_parcial, resultado["df_ciclos"])
                st.session_state["arquivo_valores_unitarios_xlsx"] = excel_vu_parcial
                st.download_button(
                    "Baixar Itens por Ciclo",
                    data=excel_vu_parcial,
                    file_name="Valores_Unitarios_e_Totais_por_Ciclo.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="download_itens_ciclo_progressivo",
                )
            else:
                st.button("Itens por Ciclo · Pendente de dados", disabled=True, use_container_width=True)
        with col_memoria:
            if (documentos_cap.get("mapa_marcos") or {}).get("habilitado"):
                pdf_memoria_parcial = gerar_pdf_linha_tempo_contrato(resultado)
                if pdf_memoria_parcial:
                    st.session_state["arquivo_mapa_marcos_pdf"] = pdf_memoria_parcial
                    st.download_button(
                        "Baixar Memória e Marcos",
                        data=pdf_memoria_parcial,
                        file_name="Mapa_Marcos_Contratuais_Linha_do_Tempo.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key="download_memoria_progressiva",
                    )
                else:
                    st.button("Memória · Disponível com ressalvas", disabled=True, use_container_width=True)
            else:
                st.button("Memória · Pendente de dados", disabled=True, use_container_width=True)
        st.info("A apuração seguirá disponível nesta etapa. Complete apenas os blocos necessários aos resultados ainda pendentes.")
        st.stop()

    st.subheader("Painel Executivo")

    modo_apuracao_ui = resultado.get("modo_apuracao", "Completo")
    base_execucao_ok_ui = bool(resultado.get("base_execucao_mensal_disponivel", True))
    if modo_apuracao_ui == "Reduzido por Itens/Estoque":
        st.markdown(
            """
            <div style="background:#F3E8FF; border:1px solid #A855F7; border-left:6px solid #7E22CE; border-radius:12px; padding:14px 16px; margin:10px 0 16px 0; color:#581C87;">
                <div style="font-weight:800; margin-bottom:4px;">Modo Reduzido por Itens/Estoque</div>
                <div style="font-size:0.95rem; line-height:1.45;">A base mensal por competência não foi informada. Os resultados por itens/estoque são estimativos e não substituem a validação financeira para pagamento.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    elif modo_apuracao_ui == "Consumo por Itens/Ciclo":
        st.markdown(
            """
            <div style="background:#F6F3EE; border:1px solid #7A8F63; border-left:6px solid #4E6E58; border-radius:12px; padding:14px 16px; margin:10px 0 16px 0; color:#2F3E2F;">
                <div style="font-weight:800; margin-bottom:4px;">Modo Consumo por Itens/Ciclo</div>
                <div style="font-size:0.95rem; line-height:1.45;">Apuração itemizada baseada nos quantitativos consumidos/executados por item e por ciclo, sem base financeira mensal por competência. Use quando a fiscalização confirmar equivalência entre consumo, medição/aprovação e faturamento devido.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        config_ce_ui = resultado.get("config_ciclo_em_execucao", {}) or {}
        corte_operacional_ui = bool(
            resultado.get("corte_operacional_aplicado", False)
            or resultado.get("corte_operacional_solicitado", False)
        )

        if corte_operacional_ui:
            ciclo_ui = config_ce_ui.get("ciclo") or "Não informado"
            data_corte_ui = config_ce_ui.get("data_corte") or "Não informada"
            fonte_ui = config_ce_ui.get("fonte") or "Não informada"

            rem_orig_num = numero_br(config_ce_ui.get("valor_remanescente_original_corte", 0))
            rem_atual_num = numero_br(config_ce_ui.get("valor_remanescente_atualizado_corte", 0))
            c0_num = numero_br(config_ce_ui.get("valor_c0_manual", 0))

            rem_orig_ui = moeda(rem_orig_num) if abs(rem_orig_num) > 0.004 else "Não informado"
            rem_atual_ui = moeda(rem_atual_num) if abs(rem_atual_num) > 0.004 else "Não informado"
            c0_ui = moeda(c0_num) if abs(c0_num) > 0.004 else "Não aplicado"

            st.markdown(
                f"""
                <div style="background:#CCFBF1; border:1px solid #14B8A6; border-left:6px solid #0F766E; border-radius:12px; padding:14px 16px; margin:10px 0 16px 0; color:#134E4A;">
                    <div style="font-weight:800; margin-bottom:4px;">Metodologia aplicada: corte operacional no ciclo em execução</div>
                    <div style="font-size:0.95rem; line-height:1.45;">
                        <b>Ciclo em execução:</b> {ciclo_ui}<br>
                        <b>Competência de corte:</b> {data_corte_ui}<br>
                        <b>Fonte da execução realizada:</b> {fonte_ui}<br>
                        <b>Remanescente original no corte:</b> {rem_orig_ui}<br>
                        <b>Remanescente atualizado no corte:</b> {rem_atual_ui}<br>
                        <b>C0 financeiro manual:</b> {c0_ui}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                f"""
                <div style="background:#EAF2F8; border:1px solid #7FB3D5; border-left:6px solid #1F4E79; border-radius:12px; padding:14px 16px; margin:10px 0 16px 0; color:#12355B;">
                    <div style="font-weight:800; margin-bottom:4px;">Metodologia aplicada: corte padrão no início dos ciclos</div>
                    <div style="font-size:0.95rem; line-height:1.45;">
                        Modo de apuração: {modo_apuracao_ui}. O saldo remanescente segue a fotografia padrão do início dos ciclos.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    df_sem_efeito_ui = resultado.get("df_meses_sem_efeito_financeiro", pd.DataFrame())
    if isinstance(df_sem_efeito_ui, pd.DataFrame) and not df_sem_efeito_ui.empty:
        st.warning(
            "Foram identificadas competências sem efeito financeiro no retroativo. "
            "Esses meses permanecem em memória, mas o respectivo delta não compõe o valor a pagar."
        )
        ef_col1, ef_col2 = st.columns(2)
        ef_col1.metric("Meses sem efeitos financeiros", int(resultado.get("quantidade_meses_sem_efeito_financeiro", len(df_sem_efeito_ui))))
        ef_col2.metric("Valor total sem efeito financeiro", moeda(resultado.get("valor_total_sem_efeito_financeiro", 0.0)))
        with st.expander("Meses sem efeito financeiro considerados na apuração", expanded=True):
            st.dataframe(
                formatar_dataframe_moeda(df_sem_efeito_ui, colunas_moeda=["Valor base", "Delta não devido"]),
                use_container_width=True,
                hide_index=True,
            )

    col1, col2 = st.columns(2)
    valor_original_painel = numero_seguro(resultado.get("valor_original_contrato", 0.0), 0.0)
    valor_formalizado_painel = numero_seguro(resultado.get("valor_formalizado_anterior", valor_original_painel), valor_original_painel)
    col1.metric("Valor Original", moeda(valor_original_painel))
    col2.metric("Número de Aditivos", int(resultado.get("quantidade_aditivos_total", 0)))
    if abs(valor_formalizado_painel - valor_original_painel) > 0.005:
        st.caption("Valor formalizado antes desta análise: " + moeda(valor_formalizado_painel))

    colp1, colp2 = st.columns(2)
    if base_execucao_ok_ui:
        colp1.metric("Valor bruto medido/aprovado", moeda(resultado["valor_pago_efetivo"]))
        colp2.metric("Valor teórico calculado", moeda(resultado["valor_teorico_calculado"]))
    else:
        if modo_apuracao_ui == "Consumo por Itens/Ciclo":
            colp1.metric("Execução atualizada por itens/ciclo", moeda(resultado.get("valor_executado_atualizado", 0.0)))
        else:
            colp1.metric("Execução estimada por estoque", moeda(resultado.get("valor_executado_atualizado", 0.0)))
        colp2.metric("Base mensal por competência", "Não informada")

    col4, col5 = st.columns(2)
    label_represado = "Valor Represado a Pagar" if base_execucao_ok_ui else "Retroativo financeiro definitivo"
    valor_represado_card = moeda(resultado.get("valor_represado_a_pagar", resultado.get("delta_total", 0))) if base_execucao_ok_ui else "Não calculado"
    col4.metric(label_represado, valor_represado_card)
    col5.metric("Ciclos", resultado.get("quantidade_ciclos", 0))

    if not base_execucao_ok_ui:
        if modo_apuracao_ui == "Consumo por Itens/Ciclo":
            retroativo_itens = numero_seguro(resultado.get("valor_retroativo_consumo_itens_ciclo", resultado.get("valor_retroativo_estimado_itens_estoque", 0.0)), 0.0)
            st.markdown(
                f"""
                <div style="background:#F6F3EE; border:1px solid #7A8F63; border-left:6px solid #4E6E58; border-radius:14px; padding:16px 18px; margin:10px 0 16px 0;">
                    <div style="color:#2F3E2F; font-weight:800; font-size:0.92rem; margin-bottom:6px;">Retroativo (itens consumidos/ciclo)</div>
                    <div style="color:#0F172A; font-size:1.65rem; font-weight:900; line-height:1.2;">{moeda(retroativo_itens)}</div>
                    <div style="color:#4E6E58; font-size:0.86rem; margin-top:8px;">Apuração calculada a partir das quantidades consumidas por item e por ciclo, com aplicação dos fatores acumulados informados em CICLOS_APURADOS.</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            df_retro_itens_ui = resultado.get("df_retroativo_itemizado_por_ciclo", pd.DataFrame())
            if isinstance(df_retro_itens_ui, pd.DataFrame) and not df_retro_itens_ui.empty:
                with st.expander("Detalhar retroativo itemizado por consumo/ciclo", expanded=False):
                    st.dataframe(
                        formatar_dataframe_moeda(
                            df_retro_itens_ui,
                            colunas_moeda=["Valor original consumido", "Valor atualizado consumido", "Retroativo"],
                            colunas_fator=["Fator acumulado"],
                        ),
                        use_container_width=True,
                        hide_index=True,
                    )
        else:
            retroativo_itens = numero_seguro(resultado.get("valor_retroativo_estimado_itens_estoque", 0.0), 0.0)
            st.markdown(
                f"""
                <div style="background:#F3E8FF; border:1px solid #A855F7; border-left:6px solid #7E22CE; border-radius:14px; padding:16px 18px; margin:10px 0 16px 0;">
                    <div style="color:#581C87; font-weight:800; font-size:0.92rem; margin-bottom:6px;">Retroativo estimado por itens/estoque</div>
                    <div style="color:#0F172A; font-size:1.65rem; font-weight:900; line-height:1.2;">{moeda(retroativo_itens)}</div>
                    <div style="color:#581C87; font-size:0.86rem; margin-top:8px;">Estimativa calculada pela diferença entre a execução atualizada por itens/estoque e a execução original estimada. Não substitui a apuração financeira mensal por competência.</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            df_retro_itens_ui = resultado.get("df_retroativo_estimado_itens_estoque", pd.DataFrame())
            if isinstance(df_retro_itens_ui, pd.DataFrame) and not df_retro_itens_ui.empty:
                with st.expander("Detalhar retroativo estimado por itens/estoque", expanded=False):
                    st.dataframe(
                        formatar_dataframe_moeda(
                            df_retro_itens_ui,
                            colunas_moeda=["Valor executado original", "Valor executado atualizado", "Retroativo estimado por itens/estoque"],
                        ),
                        use_container_width=True,
                        hide_index=True,
                    )

    st.markdown(
        f"""
        <div class="telebras-valor-destaque">
            <div class="telebras-valor-destaque-label">Valor Total Atualizado do Contrato</div>
            <div class="telebras-valor-destaque-valor">{moeda(resultado["valor_atualizado_contrato"])}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col6, col7, col8 = st.columns(3)
    col6.metric("Saldo Remanescente Atualizado", moeda(resultado["remanescente_reajustado"]))
    col7.metric("Aditivos/Supressões Registrados", moeda(resultado["total_aditivos_atualizados"]))
    col8.metric("Variação Acumulada", percentual(resultado.get("variacao_acumulada", resultado["fator_acumulado"] - 1), 2))

    if resultado.get("total_aditivos_informativos", 0.0):
        st.caption("Aditivos informativos já incorporados ao valor formalizado anterior: " + moeda(resultado.get("total_aditivos_informativos", 0.0)))

    tab_timeline, tab1, tab2, tab_aditivos, tab3, tab4, tab_auditoria, tab5 = st.tabs([
        "Linha do Tempo",
        "Painel Financeiro por Ciclo",
        "Valor Atualizado do Contrato",
        "Aditivos e Supressões",
        "Ciclos e Deltas",
        "Valores Unitários",
        "Auditoria",
        "Conferência",
    ])

    with tab_timeline:
        render_linha_tempo_contrato(resultado)
        pdf_timeline = gerar_pdf_linha_tempo_contrato(resultado)
        if pdf_timeline:
            st.session_state["arquivo_mapa_marcos_pdf"] = pdf_timeline
            st.download_button(
                label="Baixar Mapa dos Marcos em PDF",
                data=pdf_timeline,
                file_name="Mapa_Marcos_Contratuais_Linha_do_Tempo.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=False,
            )
        else:
            st.caption("O PDF da linha do tempo ficará disponível quando houver eventos processados e a biblioteca ReportLab estiver instalada.")

    with tab1:
        st.markdown("### Painel Financeiro por Ciclo")
        if not base_execucao_ok_ui:
            if modo_apuracao_ui == "Consumo por Itens/Ciclo":
                st.markdown(
                    """
                    <div style="background:#F6F3EE; border:1px solid #7A8F63; border-left:6px solid #4E6E58; border-radius:12px; padding:13px 15px; margin:8px 0 14px 0; color:#2F3E2F;">
                        Não há base mensal por competência. Neste modo, a apuração utiliza consumo por item/ciclo validado pela fiscalização; o retroativo financeiro mensal definitivo não é calculado.
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            else:
                st.warning(
                    "A base de execução mensal por competência não foi informada. "
                    "Por isso, não há cálculo definitivo de retroativo por competência neste processamento."
                )
        else:
            df_fin = formatar_dataframe_moeda(
                resultado["df_financeiro_por_ciclo"],
                colunas_moeda=["Valor pago efetivo", "Valor teórico calculado", "Delta do ciclo"],
                colunas_fator=["Fator aplicado ao retroativo"],
            )
            st.dataframe(df_fin, use_container_width=True, hide_index=True)

            with st.expander("Ver lançamentos mensais importados"):
                df_mensal = formatar_dataframe_moeda(
                    resultado["df_financeiro_mensal"],
                    colunas_moeda=["Valor pago/faturado"],
                )
                st.dataframe(df_mensal, use_container_width=True, hide_index=True)

    with tab2:
        st.markdown("### Valor Total Atualizado do Contrato")
        st.metric("Valor Total Atualizado do Contrato", moeda(resultado.get("valor_atualizado_contrato", 0.0)))
        st.caption(
            "Composição: execução atualizada por ciclo + saldo remanescente atualizado no último ciclo informado. Aditivos/supressões são apresentados para controle, sem soma autônoma ao total."
        )
        df_comp_valor = limpar_nan_inf_df(resultado.get("df_composicao_valor_total", pd.DataFrame()))
        if not df_comp_valor.empty:
            st.dataframe(
                formatar_dataframe_moeda(df_comp_valor, colunas_moeda=["Valor"]),
                use_container_width=True,
                hide_index=True,
            )
        st.markdown("### Composição do Valor Atualizado do Contrato")
        df_exec = formatar_dataframe_moeda(
            resultado["df_execucao_atualizada"],
            colunas_moeda=["Valor executado original", "Valor executado atualizado"],
            colunas_pct=["Percentual acumulado aplicado"],
        )
        st.dataframe(df_exec, use_container_width=True, hide_index=True)

        st.markdown("### Saldo Remanescente Atualizado")
        st.caption("Os saldos remanescentes atualizados já consideram os reajustes aplicáveis aos respectivos ciclos, exceto ciclos preclusos não admitidos por negociação.")
        df_rem = formatar_dataframe_moeda(
            resultado["df_remanescentes"],
            colunas_moeda=["Remanescente original", "Remanescente atualizado"],
            colunas_fator=["Fator aplicado"],
        )
        st.dataframe(df_rem, use_container_width=True, hide_index=True)

        st.info(f"Índice do contrato: {resultado.get('indice', 'Não informado')}")

    with tab_aditivos:
        st.markdown("### Aditivos")
        df_ad_exec = limpar_nan_inf_df(resultado.get("df_aditivos_executivo", pd.DataFrame()))
        colad1, colad2, colad3 = st.columns(3)
        colad1.metric("Número de Aditivos", int(resultado.get("quantidade_aditivos_total", 0)))
        colad2.metric("Valor Assinado", moeda(df_ad_exec.get("Valor do aditivo na assinatura", pd.Series(dtype=float)).apply(numero_seguro).sum() if not df_ad_exec.empty else 0.0))
        colad3.metric("Valor Atualizado", moeda(df_ad_exec.get("Valor do aditivo reajustado", pd.Series(dtype=float)).apply(numero_seguro).sum() if not df_ad_exec.empty else 0.0))
        st.caption("Aditivos e supressões são apresentados para controle formal e governança. O Valor Total Atualizado do Contrato é calculado por execução atualizada + saldo remanescente atualizado, sem soma autônoma dos aditivos para evitar dupla contagem.")
        if not df_ad_exec.empty:
            cols_exec = [c for c in ["Aditivo", "Ciclo/Marco", "Tratamento do aditivo", "Quantidade de linhas", "Valor do aditivo na assinatura", "Fator aplicado", "Valor do aditivo reajustado"] if c in df_ad_exec.columns]
            st.dataframe(
                formatar_dataframe_moeda(
                    df_ad_exec[cols_exec].copy(),
                    colunas_moeda=["Valor do aditivo na assinatura", "Valor do aditivo reajustado"],
                    colunas_fator=["Fator aplicado"],
                ),
                use_container_width=True,
                hide_index=True,
            )
            st.caption("A coluna 'Quantidade de linhas' mostra quantos itens/linhas do arquivo compõem cada aditivo consolidado.")
            with st.expander("Ver detalhamento técnico por item/linha"):
                df_ad_det = limpar_nan_inf_df(resultado.get("df_aditivos", pd.DataFrame()))
                cols_det = [c for c in ["Data do aditivo", "Identificação", "Origem do lançamento", "Ciclo/Marco", "Tipo de alteração", "Tratamento do aditivo", "Valor do aditivo na assinatura", "Fator aplicado", "Valor do aditivo reajustado", "Computa no Valor Global"] if c in df_ad_det.columns]
                df_ad_det_vis = df_ad_det[cols_det].copy()
                if "Computa no Valor Global" in df_ad_det_vis.columns:
                    df_ad_det_vis = df_ad_det_vis.rename(columns={"Computa no Valor Global": "Marcado como computável no arquivo"})
                st.dataframe(
                    formatar_dataframe_moeda(
                        df_ad_det_vis,
                        colunas_moeda=["Valor do aditivo na assinatura", "Valor do aditivo reajustado"],
                        colunas_fator=["Fator aplicado"],
                        colunas_data=["Data do aditivo"],
                    ),
                    use_container_width=True,
                    hide_index=True,
                )
        else:
            st.info("Não há aditivos ou supressões informados no arquivo processado.")

    with tab3:
        titulo_delta = "Retroativo por Ciclo" if modo_apuracao_ui == "Consumo por Itens/Ciclo" else "Delta por Ciclo"
        st.markdown(f"### {titulo_delta}")
        if not base_execucao_ok_ui:
            if modo_apuracao_ui == "Consumo por Itens/Ciclo":
                st.markdown(
                    """
                    <div style="background:#F6F3EE; border:1px solid #7A8F63; border-left:6px solid #4E6E58; border-radius:12px; padding:13px 15px; margin:8px 0 14px 0; color:#2F3E2F;">
                        Apuração por consumo itemizado. A tabela abaixo demonstra o retroativo por ciclo com base nas quantidades consumidas informadas pela fiscalização.
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            else:
                st.warning("Sem base mensal por competência, o delta/retroativo por ciclo não é definitivo. A tabela abaixo é apenas controle estimativo/estrutural.")
        df_delta = resultado.get("df_delta_por_ciclo", resultado["df_financeiro_por_ciclo"]).copy()
        if "Ciclo" in df_delta.columns:
            df_delta = df_delta[df_delta["Ciclo"] != "TOTAL"].copy()
        if modo_apuracao_ui == "Consumo por Itens/Ciclo" and "Delta do ciclo" in df_delta.columns:
            df_delta = df_delta.rename(columns={"Delta do ciclo": "Retroativo do ciclo"})
            moedas_delta = ["Valor pago efetivo", "Valor teórico calculado", "Retroativo do ciclo"]
        else:
            moedas_delta = ["Valor pago efetivo", "Valor teórico calculado", "Delta do ciclo"]
        st.dataframe(
            formatar_dataframe_moeda(
                df_delta,
                colunas_moeda=moedas_delta,
                colunas_fator=["Fator aplicado ao retroativo", "Fator acumulado"],
            ),
            use_container_width=True,
            hide_index=True,
        )
        st.caption("O C0 foi incluído como base de referência, com retroativo igual a R$ 0,00." if modo_apuracao_ui == "Consumo por Itens/Ciclo" else "O C0 foi incluído como base de referência, com delta igual a R$ 0,00.")

    with tab4:
        st.markdown("### Valores Unitários e Totais por Ciclo")
        if modo_apuracao_ui == "Consumo por Itens/Ciclo":
            st.markdown(
                """
                <div style="background:#F6F3EE; border:1px solid #7A8F63; border-left:6px solid #4E6E58; border-radius:12px; padding:12px 14px; margin:8px 0 12px 0; color:#2F3E2F;">
                    Neste modo, a tabela representa o <strong>remanescente no início de cada ciclo</strong>, calculado a partir da quantidade contratada e das quantidades consumidas já informadas.
                </div>
                """,
                unsafe_allow_html=True,
            )
        df_vu = limpar_nan_inf_df(resultado.get("df_valores_unitarios_ciclo", pd.DataFrame()))
        if df_vu.empty:
            st.info("Não há dados suficientes para montar os valores unitários por ciclo.")
        else:
            df_vu_vis = formatar_dataframe_moeda(
                df_vu,
                colunas_moeda=["Valor unitário", "Total R$"],
            )
            st.dataframe(df_vu_vis, use_container_width=True, hide_index=True)

            excel_vu = gerar_excel_valores_unitarios_por_ciclo(df_vu, resultado["df_ciclos"])
            st.session_state["arquivo_valores_unitarios_xlsx"] = excel_vu
            st.download_button(
                label="Baixar Valores Unitários e Totais por Ciclo",
                data=excel_vu,
                file_name="Valores_Unitarios_e_Totais_por_Ciclo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=False,
            )

    with tab_auditoria:
        st.markdown("### Auditoria de Consistência")
        st.caption("Checklist automático para fechamento da análise. Itens com status ATENÇÃO devem ser revisados antes da instrução processual final.")
        df_aud = limpar_nan_inf_df(resultado.get("df_auditoria_consistencia", pd.DataFrame()))
        if df_aud.empty:
            st.info("Auditoria de consistência não disponível para este processamento.")
        else:
            df_aud_vis = df_aud.copy()
            if "Diferença/Valor" in df_aud_vis.columns:
                def _formatar_diferenca_auditoria(valor):
                    if isinstance(valor, str):
                        return valor
                    if pd.isna(valor):
                        return ""
                    return moeda(valor)
                df_aud_vis["Diferença/Valor"] = df_aud_vis["Diferença/Valor"].apply(_formatar_diferenca_auditoria)
            def _estilo_linha_auditoria(row):
                status = str(row.get("Status", "")).upper()
                if "ATENÇÃO" in status or "ATENCAO" in status:
                    return ["background-color: #FFF2CC; color: #7A4A00; font-weight: 600" for _ in row]
                return ["" for _ in row]

            if "Status" in df_aud_vis.columns:
                st.dataframe(df_aud_vis.style.apply(_estilo_linha_auditoria, axis=1), use_container_width=True, hide_index=True)
            else:
                st.dataframe(df_aud_vis, use_container_width=True, hide_index=True)
            qtd_atencao = int((df_aud["Status"].astype(str).str.upper() == "ATENÇÃO").sum()) if "Status" in df_aud.columns else 0
            if qtd_atencao:
                st.warning(f"Foram identificados {qtd_atencao} ponto(s) para revisão manual.")
            else:
                st.success("Auditoria automática sem apontamentos relevantes.")

    with tab5:
        st.markdown("### Contexto do Contrato")
        contexto = resultado.get("contexto_contratual_anterior", {}) or {}
        st.caption("Memória formal anterior. Este bloco é informativo e não altera automaticamente o Valor Total Atualizado do Contrato. Quando houver valores já incorporados/retirados, eles devem estar refletidos na análise por meio das medições/execução financeira, das quantidades dos itens ou dos saldos remanescentes informados.")
        if contexto and contexto.get("contexto_informado"):
            colctx1, colctx2 = st.columns(2)
            colctx1.metric("Valor formalizado antes desta análise", moeda(contexto.get("valor_formalizado_anterior", 0.0)))
            colctx2.metric("Último ciclo formalizado", contexto.get("ultimo_ciclo_concedido", "Não informado") or "Não informado")
            if texto_seguro(contexto.get("observacao_historico", ""), ""):
                st.info(texto_seguro(contexto.get("observacao_historico", ""), "Não"))
            eventos_ctx = contexto.get("eventos_historicos_anteriores", [])
            if eventos_ctx:
                st.dataframe(pd.DataFrame(eventos_ctx), use_container_width=True, hide_index=True)
            else:
                st.caption("Sem eventos históricos anteriores informados.")
        else:
            st.info("Sem contexto contratual anterior informado.")

        if not base_execucao_ok_ui:
            if modo_apuracao_ui == "Consumo por Itens/Ciclo":
                st.markdown("### Retroativo (itens consumidos/ciclo)")
                st.markdown(
                    """
                    <div style="background:#F6F3EE; border:1px solid #7A8F63; border-left:6px solid #4E6E58; border-radius:12px; padding:12px 14px; margin:8px 0 12px 0; color:#2F3E2F;">
                        O valor abaixo é apurado por consumo itemizado por ciclo. A força da apuração depende da premissa fiscal de equivalência entre consumo, medição/aprovação e faturamento devido.
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                df_retro_itens_conf = limpar_nan_inf_df(resultado.get("df_retroativo_itemizado_por_ciclo", pd.DataFrame()))
                st.metric("Retroativo (itens consumidos/ciclo)", moeda(resultado.get("valor_retroativo_consumo_itens_ciclo", 0.0)))
                if not df_retro_itens_conf.empty:
                    st.dataframe(
                        formatar_dataframe_moeda(
                            df_retro_itens_conf,
                            colunas_moeda=["Valor original consumido", "Valor atualizado consumido", "Retroativo"],
                            colunas_fator=["Fator acumulado"],
                        ),
                        use_container_width=True,
                        hide_index=True,
                    )
            else:
                st.markdown("### Retroativo estimado por itens/estoque")
                st.markdown(
                    """
                    <div style="background:#F3E8FF; border:1px solid #A855F7; border-left:6px solid #7E22CE; border-radius:12px; padding:12px 14px; margin:8px 0 12px 0; color:#581C87;">
                        O valor abaixo é estimativo, pois não houve base mensal por competência. Ele deve ser validado antes de qualquer formalização de pagamento.
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                df_retro_itens_conf = limpar_nan_inf_df(resultado.get("df_retroativo_estimado_itens_estoque", pd.DataFrame()))
                st.metric("Retroativo estimado por itens/estoque", moeda(resultado.get("valor_retroativo_estimado_itens_estoque", 0.0)))
                if not df_retro_itens_conf.empty:
                    st.dataframe(
                        formatar_dataframe_moeda(
                            df_retro_itens_conf,
                            colunas_moeda=["Valor executado original", "Valor executado atualizado", "Retroativo estimado por itens/estoque"],
                        ),
                        use_container_width=True,
                        hide_index=True,
                    )

        st.markdown("### Efeitos financeiros sem retroativo")
        df_efeitos_conf = limpar_nan_inf_df(resultado.get("df_meses_sem_efeito_financeiro", pd.DataFrame()))
        if modo_apuracao_ui == "Consumo por Itens/Ciclo":
            st.info(
                "Não aplicável ao Modo Consumo por Itens/Ciclo. Neste modo, os efeitos financeiros são usados "
                "para orientar a distribuição das quantidades consumidas entre C0, C1, C2 etc., conforme a aba "
                "CICLOS_APURADOS. Não há apuração mensal por competência para separar meses sem retroativo."
            )
        else:
            cefe1, cefe2 = st.columns(2)
            cefe1.metric("Meses sem efeitos financeiros", int(resultado.get("quantidade_meses_sem_efeito_financeiro", 0)))
            cefe2.metric("Valor total sem efeito financeiro", moeda(resultado.get("valor_total_sem_efeito_financeiro", 0.0)))
            if not df_efeitos_conf.empty:
                st.dataframe(
                    formatar_dataframe_moeda(
                        df_efeitos_conf,
                        colunas_moeda=["Valor base", "Valor teórico se aplicado", "Delta não devido"],
                        colunas_fator=["Fator teórico"],
                    ),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.caption("Não foram identificadas competências sem efeito financeiro neste processamento.")

        st.markdown("### Quadro Executivo")
        st.dataframe(
            formatar_dataframe_moeda(resultado["df_comparativo"], colunas_moeda=["Valor"]),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("### Planilha Executiva")
        st.caption("Gera um XLS executivo com resumo financeiro, detalhamento de itens por ciclo e aditivos/supressões.")
        excel_executivo = gerar_planilha_executiva(resultado)
        st.session_state["arquivo_planilha_executiva_xlsx"] = excel_executivo
        st.download_button(
            label="Baixar Planilha Executiva",
            data=excel_executivo,
            file_name="Planilha_Executiva_Analise_Reajuste.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=False,
        )

else:
    pass
