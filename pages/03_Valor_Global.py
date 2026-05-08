import re
import html
import unicodedata
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

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


# ============================================================
# Utilitários de formatação e normalização
# ============================================================

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

    df = ler_aba_com_cabecalho(bytes_arquivo, aba)
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



def contexto_contratual_de_parametros(params):
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
    observacao = (
        contexto.get("observacao_historico")
        or params.get("observacao_sobre_historico_anterior")
        or params.get("observacao_sobre_o_historico_anterior")
        or params.get("observacao_historico_anterior")
        or ""
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

    return {
        "valor_original_contrato": float(valor_original or 0.0),
        "valor_formalizado_anterior": float(valor_formalizado or 0.0),
        "ultimo_ciclo_concedido": str(ultimo_ciclo).strip(),
        "observacao_historico": str(observacao).strip(),
        "data_base_reajuste": data_base_reajuste,
        "contexto_informado": bool(valor_original or valor_formalizado or str(ultimo_ciclo).strip() or str(observacao).strip() or str(data_base_reajuste).strip()),
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

    # Para evitar confusão com ITENS_CICLOS, aceitar apenas aba CICLOS/CICLO por correspondência exata.
    mapa_exato = {normalizar_texto(s): s for s in xls.sheet_names}
    aba = mapa_exato.get("ciclos") or mapa_exato.get("ciclo")
    if not aba:
        return pd.DataFrame(), "indisponível"

    df = ler_aba_com_cabecalho(bytes_arquivo, aba, termos_obrigatorios=["Ciclo"])
    return padronizar_ciclos(df), "arquivo"


def padronizar_ciclos(df):
    if df.empty:
        return pd.DataFrame()

    col_ciclo = localizar_coluna(df, ["Ciclo"])
    col_base = localizar_coluna(df, ["Data-base", "Base"])
    col_intervalo = localizar_coluna(df, ["Intervalo do índice", "Intervalo", "Janela"])
    col_janela_adm = localizar_coluna(df, ["Janela de admissibilidade", "JanelaAdm"])
    col_pedido = localizar_coluna(df, ["Data do pedido", "Pedido"])
    col_inicio_financeiro = localizar_coluna(df, ["Início financeiro", "Inicio financeiro", "Início dos efeitos financeiros", "Inicio dos efeitos financeiros", "Efeitos financeiros"])
    col_fim_financeiro = localizar_coluna(df, ["Fim financeiro", "Fim dos efeitos financeiros", "Fim efeito financeiro"])
    col_situacao = localizar_coluna(df, ["Situação", "Resultado"])
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
        fator = fator_de_valor(row.get(col_fator, None) if col_fator else None, variacao=variacao)

        if col_fator_acum:
            fator_acum = fator_de_valor(row.get(col_fator_acum, None), variacao=None)
            if fator_acum <= 0:
                fator_acum = fator_acumulado_calculado * fator
        else:
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
    aba = localizar_aba(xls, ["FINANCEIRO_MENSAL", "RETROATIVO", "Financeiro"])
    if not aba:
        raise ValueError("Aba FINANCEIRO_MENSAL não encontrada.")

    df = ler_aba_com_cabecalho(bytes_arquivo, aba, termos_obrigatorios=["Valor"])
    if df.empty:
        raise ValueError("Aba FINANCEIRO_MENSAL está vazia.")

    col_ciclo = localizar_coluna(df, ["Ciclo"])
    col_comp = localizar_coluna(df, ["Competência", "Competencia"])
    col_valor = localizar_coluna(df, ["Valor pago/faturado", "Valor bruto faturado", "Valor faturado", "Valor pago", "Valor"])

    if col_valor is None:
        raise ValueError("Não foi encontrada coluna de valor pago/faturado na aba FINANCEIRO_MENSAL.")

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


def atribuir_ciclo_por_competencia(competencias, ciclos):
    intervalos = []
    for _, row in ciclos.iterrows():
        ini, fim = parse_intervalo_mensal(row.get("Intervalo do índice", ""))
        if ini is not None:
            intervalos.append((row["Ciclo"], ini, fim))

    atribuidos = []
    for comp in competencias:
        try:
            periodo = pd.to_datetime(comp).to_period("M")
        except Exception:
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
    aba = localizar_aba(xls, ["ITENS_REMANESCENTES", "ITENS_CICLOS", "Itens"])
    if not aba:
        raise ValueError("Aba ITENS_REMANESCENTES não encontrada.")

    df = ler_aba_com_cabecalho(bytes_arquivo, aba, termos_obrigatorios=["Item", "Quantidade contratada", "Valor unitário original"])
    if df.empty:
        raise ValueError("Aba ITENS_REMANESCENTES está vazia.")

    col_item = localizar_coluna(df, ["Item"])
    col_qtd = localizar_coluna(df, ["Quantidade contratada", "Qtd C0", "Quantidade", "Qtd"])
    col_vu = localizar_coluna(df, ["Valor unitário original", "VU C0", "VU Original", "Valor Unitario"])
    col_total = localizar_coluna(df, ["Valor total", "TOTAL C0", "Total"])

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
        if "remanescente" in n and "valor" not in n and "total" not in n and eh_inicio_ciclo and not eh_coluna_excluida:
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


def calcular_financeiro_por_ciclo(df_financeiro, ciclos):
    fatores = mapa_fatores(ciclos)
    linhas = []

    for ciclo, grupo in df_financeiro.groupby("Ciclo", dropna=False):
        ciclo = normalizar_ciclo(ciclo) or "C1"
        total_pago = float(grupo["Valor pago/faturado"].sum())
        info = fatores.get(ciclo, {})
        fator_retroativo = float(info.get("fator_retroativo", 1.0))
        situacao = info.get("situacao", "")
        tratamento = info.get("tratamento", "")
        devido = total_pago * fator_retroativo
        delta = devido - total_pago
        linhas.append({
            "Ciclo": ciclo,
            "Situação": situacao,
            "Tratamento financeiro": tratamento,
            "Fator aplicado ao retroativo": fator_retroativo,
            "Valor pago efetivo": total_pago,
            "Valor teórico calculado": devido,
            "Delta do ciclo": delta,
        })

    if not linhas:
        return pd.DataFrame(columns=[
            "Ciclo", "Situação", "Tratamento financeiro", "Fator aplicado ao retroativo",
            "Valor pago efetivo", "Valor teórico calculado", "Delta do ciclo"
        ])

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
        fmt_factor = workbook.add_format({"num_format": "0.0000", "border": 1})

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
                base["num_format"] = "0.0000"
            return workbook.add_format(base)


        df_export = df_valores.copy()
        df_export.to_excel(writer, sheet_name="VALORES_POR_CICLO", index=False)
        ws = writer.sheets["VALORES_POR_CICLO"]

        for col_idx, title in enumerate(df_export.columns):
            ws.write(0, col_idx, title, fmt_header)

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

        for row_idx, item in enumerate(itens_ordenados, start=2):
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

        df_ciclos = ciclos.copy() if isinstance(ciclos, pd.DataFrame) else pd.DataFrame()
        df_ciclos.to_excel(writer, sheet_name="CICLOS_CONSIDERADOS", index=False)
        ws2 = writer.sheets["CICLOS_CONSIDERADOS"]
        for col_idx, title in enumerate(df_ciclos.columns):
            ws2.write(0, col_idx, title, fmt_header)
        for idx, col in enumerate(df_ciclos.columns):
            ws2.set_column(idx, idx, max(14, min(32, len(str(col)) + 4)))
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
                elif col == "Variação":
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
    - ADITIVOS_SUPRESSOES: lançamentos computáveis e informativos.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

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
        fmt_factor = workbook.add_format({"border": 1, "num_format": "0.0000", "valign": "vcenter"})
        fmt_red_money = workbook.add_format({"border": 1, "font_color": "#C00000", "num_format": 'R$ #,##0.00'})
        fmt_red_text = workbook.add_format({"border": 1, "font_color": "#C00000"})
        fmt_note = workbook.add_format({"italic": True, "font_color": "#64748B"})

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
                base["num_format"] = "0.0000"
            return workbook.add_format(base)

        # ====================================================
        # Aba 1 - Resumo Financeiro
        # ====================================================
        ws = workbook.add_worksheet("RESUMO_FINANCEIRO")
        writer.sheets["RESUMO_FINANCEIRO"] = ws
        ws.set_column("A:A", 38)
        ws.set_column("B:B", 24)
        ws.set_column("C:H", 22)
        ws.write("A1", "Planilha Executiva da Análise", fmt_title)
        ws.write("A2", "cl8us — Sistema de Apoio à Gestão de Contratos", fmt_note)

        resumo = [
            ("Data de processamento", resultado.get("data_processamento", ""), "text"),
            ("Índice utilizado", resultado.get("indice", "Não informado"), "text"),
            ("Valor original do contrato", resultado.get("valor_original_contrato", 0.0), "money"),
            ("Valor formalizado antes desta análise", resultado.get("valor_formalizado_anterior", resultado.get("valor_original_contrato", 0.0)), "money"),
            ("Valor pago efetivo", resultado.get("valor_pago_efetivo", 0.0), "money"),
            ("Valor teórico calculado", resultado.get("valor_teorico_calculado", 0.0), "money"),
            ("Valor represado a pagar", resultado.get("valor_represado_a_pagar", 0.0), "money"),
            ("Saldo remanescente atualizado", resultado.get("remanescente_reajustado", 0.0), "money"),
            ("Aditivos/Supressões computados nesta análise", resultado.get("total_aditivos_atualizados", 0.0), "money"),
            ("Aditivos/Supressões informativos", resultado.get("total_aditivos_informativos", 0.0), "money"),
            ("Reajuste acumulado", resultado.get("variacao_acumulada", resultado.get("fator_acumulado", 1.0) - 1), "pct"),
            ("Valor atualizado após esta análise", resultado.get("valor_atualizado_contrato", 0.0), "money_bold"),
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
            for r_idx, item in enumerate(itens_ordenados, start=3):
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
        else:
            ws_i.write(2, 0, "Não há dados de itens disponíveis.", fmt_text)

        # ====================================================
        # Aba 3 - Aditivos e Supressões
        # ====================================================
        ws_a = workbook.add_worksheet("ADITIVOS_SUPRESSOES")
        writer.sheets["ADITIVOS_SUPRESSOES"] = ws_a
        ws_a.write(0, 0, "Aditivos e Supressões", fmt_title)
        df_ad = limpar_nan_inf_df(resultado.get("df_aditivos", pd.DataFrame())).copy()
        if not df_ad.empty:
            cols = [c for c in ["Item", "Data do aditivo", "Ciclo/Marco", "Tipo de alteração", "Tratamento do aditivo", "Valor original da alteração", "Fator aplicado", "Valor atualizado da alteração", "Computa no Valor Global"] if c in df_ad.columns]
            df_ad = df_ad[cols].copy()
            for c_idx, col in enumerate(df_ad.columns):
                ws_a.write(2, c_idx, col, fmt_subtitle)
            for r_idx, (_, row) in enumerate(df_ad.iterrows(), start=3):
                tipo_norm = normalizar_texto(row.get("Tipo de alteração", ""))
                vermelho = tipo_norm.startswith("decr") or tipo_norm.startswith("supres")
                for c_idx, col in enumerate(df_ad.columns):
                    value = row.get(col, "")
                    fmt_base_text = fmt_red_text if vermelho else fmt_text
                    fmt_base_money = fmt_red_money if vermelho else fmt_money
                    if col in ["Valor original da alteração", "Valor atualizado da alteração"]:
                        ws_a.write_number(r_idx, c_idx, numero_seguro(value, 0.0), fmt_base_money)
                    elif col == "Fator aplicado":
                        ws_a.write_number(r_idx, c_idx, numero_seguro(value, 1.0), fmt_factor)
                    elif col == "Data do aditivo":
                        ws_a.write(r_idx, c_idx, formatar_data_br(value), fmt_base_text)
                    elif col == "Computa no Valor Global":
                        ws_a.write(r_idx, c_idx, "Sim" if bool(value) else "Não", fmt_base_text)
                    else:
                        ws_a.write(r_idx, c_idx, "" if pd.isna(value) else value, fmt_base_text)
            for idx, col in enumerate(df_ad.columns):
                ws_a.set_column(idx, idx, max(14, min(34, len(str(col)) + 4)))
        else:
            ws_a.write(2, 0, "Não há aditivos ou supressões informados.", fmt_text)

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
    aba = localizar_aba(xls, ["ADITIVOS_QUANTITATIVOS", "ADITIVOS", "Aditivos"])
    if not aba:
        return pd.DataFrame()
    try:
        df = ler_aba_com_cabecalho(bytes_arquivo, aba, termos_obrigatorios=["Item"])
    except Exception:
        return pd.DataFrame()
    if df.empty:
        return pd.DataFrame()

    col_item = localizar_coluna(df, ["Item"])
    col_data = localizar_coluna(df, ["Data do aditivo", "Data"])
    col_ciclo = localizar_coluna(df, ["Ciclo/Marco", "Ciclo", "Marco"])
    col_tipo = localizar_coluna(df, ["Tipo de alteração", "Tipo", "Acréscimo/Supressão"])
    col_qtd = localizar_coluna(df, ["Quantidade acrescida/suprimida", "Quantidade", "Qtd"])
    col_vu = localizar_coluna(df, ["Valor unitário original", "VU", "Valor unitario"])
    col_valor_original = localizar_coluna(df, ["Valor original da alteração", "Valor original do acréscimo", "Valor original"])
    col_aplicar = localizar_coluna(df, ["Aplicar reajuste acumulado? (Sim/Não)", "Aplicar reajuste", "Aplicar"])
    col_fator = localizar_coluna(df, ["Fator acumulado aplicável", "Fator", "Fator acumulado"])
    col_tratamento_aditivo = localizar_coluna(df, ["Tratamento do aditivo", "Tratamento", "Computar nesta análise"])

    fatores = mapa_fatores(ciclos)
    linhas = []
    for _, row in df.iterrows():
        item = row.get(col_item, "") if col_item else ""
        if not str(item).strip() or str(item).strip().upper() == "TOTAL":
            continue
        data_aditivo = pd.to_datetime(row.get(col_data, "") if col_data else "", dayfirst=True, errors="coerce")
        ciclo = inferir_ciclo_por_data(data_aditivo, ciclos)
        tipo = str(row.get(col_tipo, "Acréscimo") if col_tipo else "Acréscimo").strip() or "Acréscimo"
        qtd = numero_br(row.get(col_qtd, 0)) if col_qtd else 0.0
        vu = numero_br(row.get(col_vu, 0)) if col_vu else 0.0
        valor_original = numero_br(row.get(col_valor_original, 0)) if col_valor_original else 0.0
        if valor_original == 0 and qtd != 0 and vu != 0:
            valor_original = qtd * vu
        if "supress" in normalizar_texto(tipo) or "decresc" in normalizar_texto(tipo):
            valor_original = -abs(valor_original)
        else:
            valor_original = abs(valor_original)
        aplicar = str(row.get(col_aplicar, "Sim") if col_aplicar else "Sim").strip().upper()
        tratamento_aditivo = str(row.get(col_tratamento_aditivo, "Computar nesta análise") if col_tratamento_aditivo else "Computar nesta análise").strip() or "Computar nesta análise"
        fator = numero_br(row.get(col_fator, 0)) if col_fator else 0.0
        if fator <= 0:
            fator = fatores.get(ciclo, {}).get("fator_acumulado_efetivo", 1.0)
        valor_atualizado = valor_original if aplicar in ["NAO", "NÃO", "N"] else valor_original * fator
        linhas.append({
            "Item": item,
            "Data do aditivo": data_aditivo,
            "Ciclo/Marco": ciclo,
            "Tipo de alteração": tipo,
            "Valor original da alteração": valor_original,
            "Fator aplicado": fator,
            "Valor atualizado da alteração": valor_atualizado,
            "Tratamento do aditivo": tratamento_aditivo,
            "Computa no Valor Global": not aditivo_informativo_ja_incorporado(tratamento_aditivo),
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
):
    valor_original_num = numero_seguro(valor_original, 0.0)
    valor_formalizado_num = numero_seguro(valor_formalizado_anterior, valor_original_num)

    linhas = [
        {"Indicador": "Valor original do contrato", "Valor": valor_original_num},
    ]

    if abs(valor_formalizado_num - valor_original_num) > 0.005:
        linhas.append({"Indicador": "Valor formalizado antes desta análise", "Valor": valor_formalizado_num})
    else:
        linhas.append({"Indicador": "Histórico anterior", "Valor": "Sem alteração formalizada antes desta análise."})

    linhas.extend([
        {"Indicador": "Valor pago efetivo", "Valor": total_pago},
        {"Indicador": "Valor teórico calculado", "Valor": total_devido},
        {"Indicador": "Delta total", "Valor": delta_total},
        {"Indicador": "Saldo remanescente original", "Valor": rem_original},
        {"Indicador": "Saldo remanescente atualizado", "Valor": rem_atualizado},
        {"Indicador": "Aditivos/Supressões computados nesta análise", "Valor": total_aditivos},
        {"Indicador": "Aditivos/Supressões informativos", "Valor": total_aditivos_informativos},
        {"Indicador": "Valor atualizado após esta análise", "Valor": valor_atualizado_contrato},
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


def processar_arquivo_coleta(bytes_arquivo):
    xls = pd.ExcelFile(BytesIO(bytes_arquivo))
    params = ler_parametros(bytes_arquivo, xls)
    contexto_contratual = contexto_contratual_de_parametros(params)
    ciclos, origem_ciclos = ler_ciclos(bytes_arquivo, xls)
    financeiro = ler_financeiro(bytes_arquivo, xls, ciclos)
    itens, colunas_remanescentes = ler_itens(bytes_arquivo, xls)

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

    df_fin_por_ciclo = calcular_financeiro_por_ciclo(financeiro, ciclos)
    df_rem, ciclo_ultimo_rem = calcular_remanescentes_valor(itens, colunas_remanescentes, ciclos)
    df_execucao = calcular_execucao_por_diferenca(itens, colunas_remanescentes, ciclos)
    df_valores_unitarios = construir_valores_unitarios_totais(itens, colunas_remanescentes, ciclos)
    df_aditivos = ler_aditivos(bytes_arquivo, xls, ciclos)

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
    impacto_analise_atual = valor_calculado_sem_aditivos - valor_original_contrato

    if valor_formalizado_anterior > 0:
        valor_atualizado_contrato = valor_formalizado_anterior + impacto_analise_atual + total_aditivos_atualizados
    else:
        valor_formalizado_anterior = valor_original_contrato
        valor_atualizado_contrato = valor_calculado_sem_aditivos + total_aditivos_atualizados

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
    )

    df_delta_por_ciclo = montar_delta_por_ciclo(df_fin_por_ciclo, df_execucao, ciclos)

    df_aditivos = limpar_nan_inf_df(df_aditivos)
    df_valores_unitarios = limpar_nan_inf_df(df_valores_unitarios)
    df_delta_por_ciclo = limpar_nan_inf_df(df_delta_por_ciclo)

    resultado = {
        "data_processamento": agora_brasilia().strftime("%d/%m/%Y %H:%M"),
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
        "valor_represado_a_pagar": valor_represado_a_pagar,
        "remanescente_original": remanescente_original,
        "remanescente_reajustado": remanescente_atualizado,
        "fator_remanescente": fator_remanescente,
        "valor_atualizado_contrato": valor_atualizado_contrato,
        "valor_global_financeiro": valor_atualizado_contrato,
        "total_aditivos_atualizados": total_aditivos_atualizados,
        "total_aditivos_informativos": total_aditivos_informativos,
        "ciclo_ultimo_remanescente": ciclo_ultimo_rem,
        "df_ciclos": ciclos,
        "df_financeiro_mensal": financeiro,
        "df_financeiro_por_ciclo": df_fin_por_ciclo,
        "df_delta_por_ciclo": df_delta_por_ciclo,
        "df_execucao_atualizada": df_execucao,
        "df_remanescentes": df_rem,
        "df_valores_unitarios_ciclo": df_valores_unitarios,
        "df_aditivos": df_aditivos,
        "df_aditivos_computaveis": df_aditivos_computaveis,
        "df_aditivos_informativos": df_aditivos_informativos,
        "df_comparativo": df_comparativo,
    }
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
            aditivos_temp["_tipo_evento"] = aditivos_temp.get("Tipo de alteração", "Aditivo").apply(
                lambda v: "Supressão" if ("supress" in normalizar_texto(v) or "decresc" in normalizar_texto(v)) else "Aditivo"
            )
            aditivos_temp["_ciclo"] = aditivos_temp.get("Ciclo/Marco", "").apply(normalizar_ciclo)
            aditivos_temp["_tratamento"] = aditivos_temp.get("Tratamento do aditivo", "").apply(_texto_evento)
            aditivos_temp["_valor"] = aditivos_temp.get("Valor atualizado da alteração", 0.0).apply(lambda v: numero_seguro(v, 0.0))

            agrupados = (
                aditivos_temp
                .groupby(["_data_evento", "_tipo_evento", "_ciclo", "_tratamento"], dropna=False)
                .agg(
                    quantidade=("_tipo_evento", "size"),
                    valor_total=("_valor", "sum"),
                )
                .reset_index()
            )

            for _, row in agrupados.iterrows():
                tipo_evento = _texto_evento(row.get("_tipo_evento", "Aditivo")) or "Aditivo"
                ciclo = normalizar_ciclo(row.get("_ciclo", ""))
                tratamento = _texto_evento(row.get("_tratamento", "")) or "Computar nesta análise"
                qtd = int(numero_seguro(row.get("quantidade", 0), 0))
                valor = numero_seguro(row.get("valor_total", 0.0), 0.0)
                titulo = f"{tipo_evento}"
                if qtd > 1:
                    titulo += f" - {qtd} itens"
                elif qtd == 1:
                    titulo += " - 1 item"
                detalhe = f"{tipo_evento} consolidado. Ciclo/Marco: {ciclo or 'não identificado'}. Tratamento: {tratamento}."
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
                '<div class="cl8us-timeline-event">',
                f'<div class="cl8us-timeline-dot" style="background:{cor};"></div>',
                f'<div class="cl8us-timeline-card" style="border-top:4px solid {cor};">',
                f'<div class="cl8us-timeline-date">{html.escape(data_txt)}</div>',
                f'<div class="cl8us-timeline-type" style="color:{cor};">{html.escape(tipo)}</div>',
                f'<div class="cl8us-timeline-event-title">{html.escape(evento)}</div>',
                f'<div class="cl8us-timeline-detail">{html.escape(detalhe)}{valor_txt}</div>',
                '</div>',
                '</div>',
            ])
        )

    tipos_ordenados = ["Data-base para reajuste", "Ciclo de reajuste", "Pedido de reajuste", "Efeito financeiro", "Aditivo", "Supressão", "Acordo negocial", "Fim da vigência", "Marco contratual"]
    tipos_presentes = [t for t in tipos_ordenados if t in set(df_eventos["Tipo"].astype(str))]
    legend = "".join(
        f"<span class='cl8us-legend-item'><span class='cl8us-legend-dot' style='background:{_cor_tipo_timeline(t)};'></span>{html.escape(t)}</span>"
        for t in tipos_presentes
    )

    timeline_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        body {{ margin:0; font-family: Arial, Helvetica, sans-serif; background: transparent; }}
        .cl8us-timeline-shell {{
            background:#F8FAFC;
            border:1px solid #E1E7EF;
            border-radius:16px;
            padding:18px 18px 14px 18px;
            box-sizing:border-box;
        }}
        .cl8us-timeline-title {{
            color:#0B1F3A;
            font-size:18px;
            font-weight:800;
            margin-bottom:4px;
        }}
        .cl8us-timeline-subtitle {{
            color:#64748B;
            font-size:13px;
            margin-bottom:14px;
        }}
        .cl8us-timeline-track {{
            position:relative;
            display:flex;
            gap:16px;
            overflow-x:auto;
            overflow-y:hidden;
            padding:18px 4px 12px 4px;
            scroll-behavior:smooth;
        }}
        .cl8us-timeline-track:before {{
            content:"";
            position:absolute;
            top:25px;
            left:12px;
            right:12px;
            height:2px;
            background:#D8E2EC;
            z-index:0;
        }}
        .cl8us-timeline-event {{
            min-width:190px;
            max-width:230px;
            position:relative;
            z-index:1;
            flex:0 0 auto;
        }}
        .cl8us-timeline-dot {{
            width:13px;
            height:13px;
            border-radius:50%;
            border:3px solid #FFFFFF;
            box-shadow:0 0 0 1px rgba(15, 23, 42, 0.12);
            margin:0 0 8px 10px;
        }}
        .cl8us-timeline-card {{
            background:#FFFFFF;
            border:1px solid #E5EAF0;
            border-radius:12px;
            padding:10px 11px;
            box-shadow:0 1px 2px rgba(15, 23, 42, 0.04);
            min-height:112px;
            box-sizing:border-box;
        }}
        .cl8us-timeline-date {{ color:#475569; font-size:12px; font-weight:700; margin-bottom:5px; }}
        .cl8us-timeline-event-title {{ color:#0F172A; font-size:14px; font-weight:800; line-height:1.25; margin-bottom:6px; }}
        .cl8us-timeline-type {{ font-size:11px; font-weight:700; line-height:1.1; margin-bottom:6px; }}
        .cl8us-timeline-detail {{ color:#64748B; font-size:12px; line-height:1.28; }}
        .cl8us-timeline-legend {{
            display:flex;
            flex-wrap:wrap;
            gap:8px 14px;
            margin-top:10px;
            color:#475569;
            font-size:12px;
        }}
        .cl8us-legend-item {{ display:flex; align-items:center; gap:6px; }}
        .cl8us-legend-dot {{ width:9px; height:9px; border-radius:50%; display:inline-block; }}
    </style>
    </head>
    <body>
        <div class="cl8us-timeline-shell">
            <div class="cl8us-timeline-title">Linha do Tempo do Contrato</div>
            <div class="cl8us-timeline-subtitle">Visão executiva dos marcos de reajuste, pedidos, efeitos financeiros, aditivos e acordos negociais.</div>
            <div class="cl8us-timeline-track">{''.join(cards)}</div>
            <div class="cl8us-timeline-legend">{legend}</div>
        </div>
    </body>
    </html>
    """
    components.html(timeline_html, height=310, scrolling=False)

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
        "cl8usTituloTimeline",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontSize=15,
        leading=18,
        textColor=colors.HexColor("#0B1F3A"),
        spaceAfter=4,
    )
    subtitulo = ParagraphStyle(
        "cl8usSubtituloTimeline",
        parent=styles["Normal"],
        alignment=TA_CENTER,
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#475569"),
        spaceAfter=10,
    )
    h2 = ParagraphStyle(
        "cl8usH2Timeline",
        parent=styles["Heading2"],
        fontSize=10.5,
        leading=13,
        textColor=colors.HexColor("#123B63"),
        spaceBefore=6,
        spaceAfter=5,
    )
    normal = ParagraphStyle(
        "cl8usNormalTimeline",
        parent=styles["Normal"],
        fontSize=8,
        leading=10.5,
        textColor=colors.HexColor("#1F2937"),
    )
    celula = ParagraphStyle(
        "cl8usCelulaTimeline",
        parent=styles["Normal"],
        fontSize=6.7,
        leading=8.2,
        textColor=colors.HexColor("#1F2937"),
        alignment=TA_LEFT,
    )
    celula_branca = ParagraphStyle(
        "cl8usCelulaBrancaTimeline",
        parent=celula,
        textColor=colors.white,
    )

    elementos = []
    elementos.append(Paragraph("cl8us - Relatório Executivo", titulo))
    elementos.append(Paragraph("Linha do Tempo do Contrato", subtitulo))
    elementos.append(Paragraph(f"Gerado em: {agora_brasilia().strftime('%d/%m/%Y %H:%M')}", subtitulo))

    # Resumo executivo do relatório.
    resumo = [
        ["Indicador", "Valor", "Indicador", "Valor"],
        ["Índice", str(resultado.get("indice", "-")), "Ciclos", str(resultado.get("quantidade_ciclos", "-"))],
        ["Reajuste acumulado", percentual(resultado.get("variacao_acumulada", 0.0), 2), "Valor atualizado após análise", moeda(resultado.get("valor_atualizado_contrato", 0.0))],
        ["Valor represado a pagar", moeda(resultado.get("valor_represado_a_pagar", 0.0)), "Eventos na timeline", str(len(df_eventos))],
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

    elementos.append(Paragraph("2. Linha do tempo executiva", h2))
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

def aplicar_css_responsivo_cl8us():
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
        .cl8us-valor-destaque {
            background:#EAF2F8;
            border:1px solid #BFD7EA;
            border-radius:14px;
            padding:18px 22px;
            margin:12px 0 18px 0;
        }
        .cl8us-valor-destaque-label {
            font-size:0.95rem;
            color:#27496D;
            font-weight:600;
        }
        .cl8us-valor-destaque-valor {
            font-size:clamp(1.35rem, 2.8vw, 2.05rem);
            color:#0B1F3A;
            font-weight:800;
            line-height:1.25;
            word-break:break-word;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# Interface
# ============================================================

aplicar_css_responsivo_cl8us()

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Valor Global do Contrato")

st.markdown(
    """
    Este módulo processa o **Arquivo de Coleta** preenchido e apresenta uma visão executiva
    da execução financeira, dos deltas por ciclo e do valor atualizado do contrato.
    """
)

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

st.subheader("Carregar Arquivo de Coleta")
st.info(
    "Na aba ADITIVOS_QUANTITATIVOS, use **Computar nesta análise** para aditivos ou supressões que ainda devem impactar o Valor Global atual. "
    "Use **Informativo - já incluído no valor formalizado** quando o lançamento já estiver contemplado no campo **Valor formalizado antes desta análise**, evitando dupla contagem."
)
arquivo = st.file_uploader("Carregue aqui o Arquivo de Coleta preenchido (.xlsx)", type=["xlsx"])

if arquivo is not None:
    if st.button("Processar Coleta Preenchida", type="primary", use_container_width=False):
        try:
            resultado = processar_arquivo_coleta(arquivo.getvalue())
            st.session_state["resultado_valor_global"] = resultado
            st.success("Arquivo processado com sucesso. Resultados disponíveis abaixo e no módulo Relatório Global.")
        except Exception as exc:
            st.error(f"Não foi possível processar o arquivo: {exc}")

resultado = st.session_state.get("resultado_valor_global")

if resultado:
    st.divider()
    st.subheader("Painel Executivo")

    col1, col2 = st.columns(2)
    valor_original_painel = numero_seguro(resultado.get("valor_original_contrato", 0.0), 0.0)
    valor_formalizado_painel = numero_seguro(resultado.get("valor_formalizado_anterior", valor_original_painel), valor_original_painel)
    col1.metric("Valor Original", moeda(valor_original_painel))
    if abs(valor_formalizado_painel - valor_original_painel) > 0.005:
        col2.metric("Valor Formalizado Antes da Análise", moeda(valor_formalizado_painel))
    else:
        col2.metric("Histórico anterior", "Sem alteração")
        col2.caption("Valor formalizado antes desta análise igual ao valor original do contrato.")

    colp1, colp2 = st.columns(2)
    colp1.metric("Valor Pago Efetivo", moeda(resultado["valor_pago_efetivo"]))
    colp2.metric("Valor Teórico Calculado", moeda(resultado["valor_teorico_calculado"]))

    col4, col5 = st.columns(2)
    col4.metric("Valor Represado a Pagar", moeda(resultado.get("valor_represado_a_pagar", resultado.get("delta_total", 0))))
    col5.metric("Ciclos", resultado.get("quantidade_ciclos", 0))

    st.markdown(
        f"""
        <div class="cl8us-valor-destaque">
            <div class="cl8us-valor-destaque-label">Valor Atualizado Após Esta Análise</div>
            <div class="cl8us-valor-destaque-valor">{moeda(resultado["valor_atualizado_contrato"])}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col6, col7, col8 = st.columns(3)
    col6.metric("Saldo Remanescente Atualizado", moeda(resultado["remanescente_reajustado"]))
    col7.metric("Aditivos Computados nesta Análise", moeda(resultado["total_aditivos_atualizados"]))
    col8.metric("Variação Acumulada", percentual(resultado.get("variacao_acumulada", resultado["fator_acumulado"] - 1), 2))

    if resultado.get("total_aditivos_informativos", 0.0):
        st.caption("Aditivos informativos já incorporados ao valor formalizado anterior: " + moeda(resultado.get("total_aditivos_informativos", 0.0)))

    tab_timeline, tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "Linha do Tempo",
        "Painel Financeiro por Ciclo",
        "Valor Atualizado do Contrato",
        "Ciclos e Deltas",
        "Valores Unitários",
        "Conferência",
    ])

    with tab_timeline:
        render_linha_tempo_contrato(resultado)
        pdf_timeline = gerar_pdf_linha_tempo_contrato(resultado)
        if pdf_timeline:
            st.download_button(
                label="Baixar Relatório Executivo em PDF",
                data=pdf_timeline,
                file_name="Relatorio_Executivo_Linha_do_Tempo_cl8us.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=False,
            )
        else:
            st.caption("O PDF da linha do tempo ficará disponível quando houver eventos processados e a biblioteca ReportLab estiver instalada.")

    with tab1:
        st.markdown("### Painel Financeiro por Ciclo")
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
        st.markdown("### Composição do Valor Atualizado do Contrato")
        df_exec = formatar_dataframe_moeda(
            resultado["df_execucao_atualizada"],
            colunas_moeda=["Valor executado original", "Valor executado atualizado"],
            colunas_pct=["Percentual acumulado aplicado"],
        )
        st.dataframe(df_exec, use_container_width=True, hide_index=True)

        st.markdown("### Saldo Remanescente Atualizado")
        df_rem = formatar_dataframe_moeda(
            resultado["df_remanescentes"],
            colunas_moeda=["Remanescente original", "Remanescente atualizado"],
            colunas_fator=["Fator aplicado"],
        )
        st.dataframe(df_rem, use_container_width=True, hide_index=True)

        if not resultado["df_aditivos"].empty:
            st.markdown("### Aditivos e Supressões")
            df_ad = formatar_dataframe_moeda(
                resultado["df_aditivos"],
                colunas_moeda=["Valor original da alteração", "Valor atualizado da alteração"],
                colunas_fator=["Fator aplicado"],
                colunas_data=["Data do aditivo"],
            )
            st.dataframe(df_ad, use_container_width=True, hide_index=True)
            st.caption("Somente os lançamentos com tratamento 'Computar nesta análise' são somados ao valor atualizado após esta análise.")

    with tab3:
        st.markdown("### Delta por Ciclo")
        df_delta = resultado.get("df_delta_por_ciclo", resultado["df_financeiro_por_ciclo"]).copy()
        if "Ciclo" in df_delta.columns:
            df_delta = df_delta[df_delta["Ciclo"] != "TOTAL"].copy()
        st.dataframe(
            formatar_dataframe_moeda(
                df_delta,
                colunas_moeda=["Valor pago efetivo", "Valor teórico calculado", "Delta do ciclo"],
                colunas_fator=["Fator aplicado ao retroativo"],
            ),
            use_container_width=True,
            hide_index=True,
        )
        st.caption("O C0 foi incluído como base de referência, com delta igual a R$ 0,00.")

    with tab4:
        st.markdown("### Valores Unitários e Totais por Ciclo")
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
            st.download_button(
                label="Baixar Valores Unitários e Totais por Ciclo",
                data=excel_vu,
                file_name="Valores_Unitarios_e_Totais_por_Ciclo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=False,
            )

    with tab5:
        st.markdown("### Quadro Executivo")
        st.dataframe(
            formatar_dataframe_moeda(resultado["df_comparativo"], colunas_moeda=["Valor"]),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("### Planilha Executiva")
        st.caption("Gera um XLS executivo com resumo financeiro, detalhamento de itens por ciclo e aditivos/supressões.")
        excel_executivo = gerar_planilha_executiva(resultado)
        st.download_button(
            label="Baixar Planilha Executiva",
            data=excel_executivo,
            file_name="Planilha_Executiva_cl8us.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=False,
        )

else:
    st.info("Carregue e processe o Arquivo de Coleta para visualizar o Valor Global do contrato.")
