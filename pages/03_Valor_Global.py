import re
import unicodedata
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Análises de Reajustes - Valor Global", layout="wide")


# ============================================================
# Utilitários de formatação e normalização
# ============================================================

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
    """Lê aba encontrando automaticamente a linha de cabeçalho."""
    bruto = pd.read_excel(BytesIO(bytes_arquivo), sheet_name=aba, header=None)
    termos_obrigatorios = [normalizar_texto(t) for t in (termos_obrigatorios or [])]

    linha_cabecalho = 0
    for idx, row in bruto.iterrows():
        valores = [normalizar_texto(v) for v in row.tolist()]
        texto_linha = " ".join(valores)
        if not termos_obrigatorios:
            if any(v for v in valores):
                linha_cabecalho = idx
                break
        else:
            if all(t in texto_linha for t in termos_obrigatorios):
                linha_cabecalho = idx
                break

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
            "Situação": c.get("situacao") or c.get("Situação") or "",
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
    col_situacao = localizar_coluna(df, ["Situação", "Resultado"])
    col_variacao = localizar_coluna(df, ["Variação", "Variacao", "Percentual"])
    col_fator = localizar_coluna(df, ["Fator"])
    col_fator_acum = localizar_coluna(df, ["Fator acumulado", "Fator acumulado final", "Fator Acumulado"])

    linhas = []
    fator_acumulado_calculado = 1.0

    for _, row in df.iterrows():
        ciclo = normalizar_ciclo(row.get(col_ciclo, "")) if col_ciclo else ""
        if not ciclo:
            continue

        situacao = row.get(col_situacao, "") if col_situacao else ""
        variacao = percentual_para_decimal(row.get(col_variacao, 0)) if col_variacao else 0.0
        fator = fator_de_valor(row.get(col_fator, None) if col_fator else None, variacao=variacao)

        if col_fator_acum:
            fator_acum = fator_de_valor(row.get(col_fator_acum, None), variacao=None)
            if fator_acum <= 0:
                fator_acum = fator_acumulado_calculado * fator
        else:
            fator_acum = fator_acumulado_calculado * fator

        # Fator acumulado efetivo para cálculos financeiros: não aplica ciclo precluso/adiantado comum.
        if contem_preclusao_ou_adiantamento(situacao):
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
            "Situação": situacao,
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

    itens = pd.DataFrame()
    itens["Item"] = df[col_item]
    itens["Quantidade contratada"] = df[col_qtd].apply(numero_br)
    itens["Valor unitário original"] = df[col_vu].apply(numero_br)

    if col_total:
        total_informado = df[col_total].apply(numero_br)
        itens["Valor total original"] = total_informado.where(total_informado > 0, itens["Quantidade contratada"] * itens["Valor unitário original"])
    else:
        itens["Valor total original"] = itens["Quantidade contratada"] * itens["Valor unitário original"]

    # Detecta colunas de remanescente físico.
    colunas_rem = []
    for col in df.columns:
        n = normalizar_texto(col)
        if "remanescente" in n and "valor" not in n and "total" not in n:
            colunas_rem.append(col)

    # Compatibilidade com modelo antigo.
    col_consumido = localizar_coluna(df, ["Consumido no Ciclo", "Consumido"])
    if not colunas_rem:
        col_qtd_rem = localizar_coluna(df, ["Qtd Remanescente", "Quantidade Remanescente", "Remanescente"])
        if col_qtd_rem:
            colunas_rem = [col_qtd_rem]

    for col in colunas_rem:
        itens[col] = df[col].apply(numero_br)

    if col_consumido:
        itens["Consumido no Ciclo"] = df[col_consumido].apply(numero_br)

    return itens, colunas_rem




def data_base_para_datetime(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return None
    try:
        return pd.to_datetime(valor, dayfirst=True).to_pydatetime()
    except Exception:
        return None


def inferir_ciclo_por_data_aditivo(data_aditivo, ciclos):
    data = data_base_para_datetime(data_aditivo)
    if data is None or ciclos is None or ciclos.empty:
        return ""
    escolhido = ""
    melhor_data = None
    for _, row in ciclos.iterrows():
        ciclo = normalizar_ciclo(row.get("Ciclo", ""))
        base = data_base_para_datetime(row.get("Data-base", ""))
        if ciclo and base is not None and base <= data:
            if melhor_data is None or base >= melhor_data:
                melhor_data = base
                escolhido = ciclo
    return escolhido


def ler_aditivos(bytes_arquivo, xls):
    aba = localizar_aba(xls, ["ADITIVOS_QUANTITATIVOS", "ADITIVOS", "ACRESCIMOS", "ACRÉSCIMOS"])
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
    col_tipo = localizar_coluna(df, ["Tipo de alteração", "Tipo", "Acréscimo", "Acrescimo"])
    col_qtd = localizar_coluna(df, ["Quantidade acrescida/suprimida", "Quantidade", "Qtd"])
    col_vu = localizar_coluna(df, ["Valor unitário original", "Valor Unitario", "VU"])
    col_valor = localizar_coluna(df, ["Valor original da alteração", "Valor original", "Valor original do acréscimo", "Valor original do acrescimo"])
    col_aplicar = localizar_coluna(df, ["Aplicar reajuste acumulado? (Sim/Não)", "Aplicar reajuste acumulado", "Aplicar reajuste"])
    col_fator = localizar_coluna(df, ["Fator acumulado aplicável", "Fator acumulado", "Fator"])
    col_valor_atualizado = localizar_coluna(df, ["Valor atualizado da alteração", "Valor atualizado"])

    linhas = []
    for _, row in df.iterrows():
        item = row.get(col_item, "") if col_item else ""
        data = row.get(col_data, "") if col_data else ""
        ciclo = normalizar_ciclo(row.get(col_ciclo, "")) if col_ciclo else ""
        tipo = str(row.get(col_tipo, "") if col_tipo else "").strip()
        qtd = numero_br(row.get(col_qtd, 0)) if col_qtd else 0.0
        vu = numero_br(row.get(col_vu, 0)) if col_vu else 0.0
        valor_original = numero_br(row.get(col_valor, 0)) if col_valor else 0.0
        if valor_original == 0 and qtd != 0 and vu != 0:
            valor_original = qtd * vu
        aplicar = str(row.get(col_aplicar, "Sim") if col_aplicar else "Sim").strip() or "Sim"
        fator = fator_de_valor(row.get(col_fator, None) if col_fator else None)
        valor_atualizado = numero_br(row.get(col_valor_atualizado, 0)) if col_valor_atualizado else 0.0
        if str(item).strip() == "" and valor_original == 0 and str(data).strip() == "":
            continue
        linhas.append({
            "Item": item,
            "Data do aditivo": data,
            "Ciclo/Marco": ciclo,
            "Tipo de alteração": tipo,
            "Quantidade acrescida/suprimida": qtd,
            "Valor unitário original": vu,
            "Valor original da alteração": valor_original,
            "Aplicar reajuste acumulado?": aplicar,
            "Fator acumulado aplicável": fator,
            "Valor atualizado da alteração": valor_atualizado,
        })
    return pd.DataFrame(linhas)


def processar_aditivos(df_aditivos, ciclos):
    if df_aditivos is None or df_aditivos.empty:
        return pd.DataFrame(columns=[
            "Item", "Data do aditivo", "Ciclo/Marco", "Tipo de alteração", "Valor original da alteração",
            "Fator aplicado", "Valor atualizado da alteração"
        ])
    fatores = mapa_fatores(ciclos)
    linhas = []
    for _, row in df_aditivos.iterrows():
        ciclo = normalizar_ciclo(row.get("Ciclo/Marco", ""))
        if not ciclo:
            ciclo = inferir_ciclo_por_data_aditivo(row.get("Data do aditivo", ""), ciclos)
        tipo = str(row.get("Tipo de alteração", "") or "").strip()
        sinal = -1.0 if "decr" in normalizar_texto(tipo) or "supress" in normalizar_texto(tipo) else 1.0
        valor_original = float(row.get("Valor original da alteração", 0.0) or 0.0) * sinal
        aplicar = str(row.get("Aplicar reajuste acumulado?", "Sim") or "Sim").strip().lower()
        fator = float(row.get("Fator acumulado aplicável", 0.0) or 0.0)
        if fator <= 0:
            fator = fatores.get(ciclo, {}).get("fator_acumulado_efetivo", 1.0)
        if aplicar in ["não", "nao", "n"]:
            fator = 1.0
        valor_atualizado = float(row.get("Valor atualizado da alteração", 0.0) or 0.0)
        if valor_atualizado == 0 and valor_original != 0:
            valor_atualizado = valor_original * fator
        linhas.append({
            "Item": row.get("Item", ""),
            "Data do aditivo": row.get("Data do aditivo", ""),
            "Ciclo/Marco": ciclo,
            "Tipo de alteração": tipo or "Acréscimo",
            "Valor original da alteração": valor_original,
            "Fator aplicado": fator,
            "Valor atualizado da alteração": valor_atualizado,
        })
    return pd.DataFrame(linhas)

# ============================================================
# Cálculos
# ============================================================

def mapa_fatores(ciclos):
    mapa = {"C0": {"fator": 1.0, "fator_acumulado": 1.0, "fator_acumulado_efetivo": 1.0, "situacao": "BASE"}}
    if ciclos.empty:
        return mapa

    for _, row in ciclos.iterrows():
        ciclo = row["Ciclo"]
        mapa[ciclo] = {
            "fator": float(row.get("Fator", 1.0) or 1.0),
            "fator_acumulado": float(row.get("Fator acumulado", 1.0) or 1.0),
            "fator_acumulado_efetivo": float(row.get("Fator acumulado efetivo", row.get("Fator acumulado", 1.0)) or 1.0),
            "situacao": row.get("Situação", ""),
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
        fator_aplicado = info.get("fator_acumulado_efetivo", 1.0)
        situacao = info.get("situacao", "")

        if ciclo == "C0":
            fator_aplicado = 1.0

        devido = total_pago * fator_aplicado
        delta = devido - total_pago

        linhas.append({
            "Ciclo": ciclo,
            "Situação": situacao,
            "Fator aplicado": fator_aplicado,
            "Valor pago/faturado": total_pago,
            "Valor devido reajustado": devido,
            "Delta do ciclo": delta,
        })

    if not linhas:
        return pd.DataFrame(columns=[
            "Ciclo", "Situação", "Fator aplicado", "Valor pago/faturado",
            "Valor devido reajustado", "Delta do ciclo"
        ])

    df = pd.DataFrame(linhas).sort_values(by="Ciclo", key=lambda s: s.map(numero_ciclo)).reset_index(drop=True)
    df["Delta acumulado"] = df["Delta do ciclo"].cumsum()
    return df


def ciclo_da_coluna_remanescente(coluna, fallback="C1"):
    texto = str(coluna).upper()
    m = re.search(r"C\s*([0-9]+)", texto)
    if m:
        return f"C{int(m.group(1))}"
    return fallback


def calcular_remanescentes(itens, colunas_remanescentes, ciclos):
    fatores = mapa_fatores(ciclos)
    linhas = []
    if not colunas_remanescentes:
        return pd.DataFrame(), None

    for col in colunas_remanescentes:
        ciclo = ciclo_da_coluna_remanescente(col, fallback="C1")
        info = fatores.get(ciclo, {})
        fator = info.get("fator_acumulado_efetivo", info.get("fator_acumulado", 1.0))

        qtd_rem = itens[col].apply(numero_br)
        valor_original = qtd_rem * itens["Valor unitário original"]
        valor_reajustado = valor_original * fator

        linhas.append({
            "Ciclo": ciclo,
            "Coluna de origem": col,
            "Quantidade remanescente": float(qtd_rem.sum()),
            "Remanescente original": float(valor_original.sum()),
            "Fator aplicado": float(fator),
            "Remanescente reajustado": float(valor_reajustado.sum()),
        })

    df = pd.DataFrame(linhas).sort_values(by="Ciclo", key=lambda s: s.map(numero_ciclo)).reset_index(drop=True)
    ciclo_ultimo = df.iloc[-1]["Ciclo"] if not df.empty else None
    return df, ciclo_ultimo


def calcular_consumo_estoque_por_diferenca(itens, colunas_remanescentes, ciclos):
    """Calcula consumo por ciclo usando diferença entre remanescentes sucessivos."""
    if not colunas_remanescentes:
        return pd.DataFrame()

    fatores = mapa_fatores(ciclos)

    rem_por_ciclo = []
    for col in colunas_remanescentes:
        ciclo = ciclo_da_coluna_remanescente(col, fallback="C1")
        rem_por_ciclo.append((ciclo, col))

    rem_por_ciclo = sorted(rem_por_ciclo, key=lambda x: numero_ciclo(x[0]))

    linhas = []

    # C0 = contratado inicial - remanescente início C1
    primeiro_ciclo, primeira_col = rem_por_ciclo[0]
    consumido_c0_qtd = itens["Quantidade contratada"] - itens[primeira_col].apply(numero_br)
    valor_c0 = consumido_c0_qtd * itens["Valor unitário original"]
    linhas.append({
        "Ciclo": "C0",
        "Quantidade consumida estimada": float(consumido_c0_qtd.sum()),
        "Fator aplicado": 1.0,
        "Valor consumido original": float(valor_c0.sum()),
        "Valor consumido reajustado": float(valor_c0.sum()),
        "Critério": f"Quantidade contratada - {primeira_col}",
    })

    # Cn = remanescente início Cn - remanescente início C(n+1)
    for idx in range(len(rem_por_ciclo) - 1):
        ciclo_atual, col_atual = rem_por_ciclo[idx]
        proximo_ciclo, col_proxima = rem_por_ciclo[idx + 1]
        qtd_consumida = itens[col_atual].apply(numero_br) - itens[col_proxima].apply(numero_br)
        qtd_consumida = qtd_consumida.clip(lower=0)
        valor_original = qtd_consumida * itens["Valor unitário original"]
        fator = fatores.get(ciclo_atual, {}).get("fator_acumulado_efetivo", 1.0)
        linhas.append({
            "Ciclo": ciclo_atual,
            "Quantidade consumida estimada": float(qtd_consumida.sum()),
            "Fator aplicado": float(fator),
            "Valor consumido original": float(valor_original.sum()),
            "Valor consumido reajustado": float((valor_original * fator).sum()),
            "Critério": f"{col_atual} - {col_proxima}",
        })

    return pd.DataFrame(linhas)


def montar_comparativo(valor_original, total_pago, total_devido, rem_original, rem_reaj, valor_atualizado, delta_total, valor_represado):
    linhas = [
        {
            "Indicador": "Valor original do contrato",
            "Antes do Reajuste": valor_original,
            "Após Reajuste": valor_original,
            "Diferença": 0.0,
        },
        {
            "Indicador": "Valor pago efetivo",
            "Antes do Reajuste": total_pago,
            "Após Reajuste": total_pago,
            "Diferença": 0.0,
        },
        {
            "Indicador": "Valor teórico calculado",
            "Antes do Reajuste": total_pago,
            "Após Reajuste": total_devido,
            "Diferença": total_devido - total_pago,
        },
        {
            "Indicador": "Valor represado a pagar",
            "Antes do Reajuste": 0.0,
            "Após Reajuste": valor_represado,
            "Diferença": valor_represado,
        },
        {
            "Indicador": "Saldo remanescente",
            "Antes do Reajuste": rem_original,
            "Após Reajuste": rem_reaj,
            "Diferença": rem_reaj - rem_original,
        },
        {
            "Indicador": "Valor atualizado do contrato",
            "Antes do Reajuste": valor_original,
            "Após Reajuste": valor_atualizado,
            "Diferença": valor_atualizado - valor_original,
        },
        {
            "Indicador": "Delta total",
            "Antes do Reajuste": 0.0,
            "Após Reajuste": delta_total,
            "Diferença": delta_total,
        },
    ]
    return pd.DataFrame(linhas)


def extrair_percentual_reajuste_legacy(bytes_arquivo, xls):
    """Busca percentual de reajuste em modelos legados, como ITENS_CICLOS."""
    for aba in xls.sheet_names:
        try:
            bruto = pd.read_excel(BytesIO(bytes_arquivo), sheet_name=aba, header=None)
        except Exception:
            continue
        for i in range(bruto.shape[0]):
            for j in range(bruto.shape[1]):
                cel = bruto.iat[i, j]
                if "percentual_reajuste" in normalizar_texto(cel):
                    # Procura valor à direita e, se necessário, nas células próximas.
                    candidatos = []
                    for jj in range(j + 1, min(j + 4, bruto.shape[1])):
                        candidatos.append(bruto.iat[i, jj])
                    for ii in range(i + 1, min(i + 3, bruto.shape[0])):
                        candidatos.append(bruto.iat[ii, j])
                    for cand in candidatos:
                        if cand is not None and not pd.isna(cand):
                            n = numero_br(cand)
                            if n != 0:
                                return n if abs(n) <= 1 else n / 100
    return None


def processar_arquivo_coleta(bytes_arquivo):
    xls = pd.ExcelFile(BytesIO(bytes_arquivo))

    params = ler_parametros(bytes_arquivo, xls)
    ciclos, origem_ciclos = ler_ciclos(bytes_arquivo, xls)
    financeiro = ler_financeiro(bytes_arquivo, xls, ciclos)
    itens, colunas_remanescentes = ler_itens(bytes_arquivo, xls)
    aditivos_brutos = ler_aditivos(bytes_arquivo, xls)

    if ciclos.empty:
        # Fallback pelo parâmetro de fator ou pelo percentual existente em modelos legados.
        perc_legacy = extrair_percentual_reajuste_legacy(bytes_arquivo, xls)
        fator = fator_de_valor(params.get("fator_acumulado_final") or params.get("fator_acumulado") or params.get("fator"))
        if fator == 1.0 and perc_legacy is not None:
            fator = 1 + perc_legacy
        ciclos = pd.DataFrame([{
            "Ciclo": "C1",
            "Data-base": "",
            "Intervalo do índice": "",
            "Janela de admissibilidade": "",
            "Data do pedido": "",
            "Situação": "",
            "Variação": fator - 1,
            "Fator": fator,
            "Fator acumulado": fator,
            "Fator acumulado efetivo": fator,
            "Fator ciclo efetivo": fator,
        }])

    df_fin_por_ciclo = calcular_financeiro_por_ciclo(financeiro, ciclos)
    df_rem, ciclo_ultimo_rem = calcular_remanescentes(itens, colunas_remanescentes, ciclos)
    df_consumo_estoque = calcular_consumo_estoque_por_diferenca(itens, colunas_remanescentes, ciclos)

    valor_original_contrato = float(itens["Valor total original"].sum())
    total_pago_faturado = float(df_fin_por_ciclo["Valor pago/faturado"].sum()) if not df_fin_por_ciclo.empty else 0.0
    total_devido_reajustado = float(df_fin_por_ciclo["Valor devido reajustado"].sum()) if not df_fin_por_ciclo.empty else 0.0
    delta_acumulado = float(df_fin_por_ciclo["Delta do ciclo"].sum()) if not df_fin_por_ciclo.empty else 0.0

    if not df_rem.empty:
        remanescente_original = float(df_rem.iloc[-1]["Remanescente original"])
        remanescente_reajustado = float(df_rem.iloc[-1]["Remanescente reajustado"])
        fator_remanescente = float(df_rem.iloc[-1]["Fator aplicado"])
    else:
        remanescente_original = 0.0
        remanescente_reajustado = 0.0
        fator_remanescente = 1.0

    valor_global_financeiro = total_devido_reajustado + remanescente_reajustado

    if not df_consumo_estoque.empty:
        valor_executado_atualizado = float(df_consumo_estoque["Valor consumido reajustado"].sum())
    else:
        valor_executado_atualizado = max(valor_original_contrato - remanescente_original, 0.0)

    valor_atualizado_contrato_sem_aditivos = valor_executado_atualizado + remanescente_reajustado
    df_aditivos = processar_aditivos(aditivos_brutos, ciclos)
    total_aditivos_atualizados = float(df_aditivos["Valor atualizado da alteração"].sum()) if not df_aditivos.empty else 0.0
    valor_atualizado_contrato = valor_atualizado_contrato_sem_aditivos + total_aditivos_atualizados

    if not df_fin_por_ciclo.empty:
        mask_apuravel = ~df_fin_por_ciclo["Situação"].apply(contem_preclusao_ou_adiantamento)
        valor_represado_a_pagar = float(df_fin_por_ciclo.loc[mask_apuravel, "Delta do ciclo"].clip(lower=0).sum())
    else:
        valor_represado_a_pagar = 0.0

    fator_acumulado = float(ciclos["Fator acumulado efetivo"].iloc[-1]) if "Fator acumulado efetivo" in ciclos.columns and not ciclos.empty else fator_remanescente
    indice = (
        st.session_state.get("dados_admissibilidade", {}).get("indice")
        or params.get("indice_utilizado")
        or params.get("indice")
        or "Não informado"
    )

    df_comparativo = montar_comparativo(
        valor_original_contrato,
        total_pago_faturado,
        total_devido_reajustado,
        remanescente_original,
        remanescente_reajustado,
        valor_atualizado_contrato,
        delta_acumulado,
        valor_represado_a_pagar,
    )

    resultado = {
        "data_processamento": datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M"),
        "origem_ciclos": origem_ciclos,
        "indice": indice,
        "fator_acumulado": fator_acumulado,
        "valor_original_contrato": valor_original_contrato,
        "total_pago_faturado": total_pago_faturado,
        "total_devido_reajustado": total_devido_reajustado,
        "delta_acumulado": delta_acumulado,
        "valor_represado_a_pagar": valor_represado_a_pagar,
        "remanescente_original": remanescente_original,
        "remanescente_reajustado": remanescente_reajustado,
        "fator_remanescente": fator_remanescente,
        "valor_global_financeiro": valor_global_financeiro,
        "valor_executado_atualizado": valor_executado_atualizado,
        "valor_atualizado_contrato_sem_aditivos": valor_atualizado_contrato_sem_aditivos,
        "total_aditivos_atualizados": total_aditivos_atualizados,
        "valor_atualizado_contrato": valor_atualizado_contrato,
        # Campos legados preservados para compatibilidade com relatórios antigos.
        "valor_consumido_estoque": valor_executado_atualizado,
        "valor_global_estoque": valor_atualizado_contrato,
        "diferenca_metodos": 0.0,
        "percentual_divergencia": 0.0,
        "ciclo_ultimo_remanescente": ciclo_ultimo_rem,
        "df_ciclos": ciclos,
        "df_financeiro_mensal": financeiro,
        "df_financeiro_por_ciclo": df_fin_por_ciclo,
        "df_itens_processado": itens,
        "df_remanescentes": df_rem,
        "df_consumo_estoque": df_consumo_estoque,
        "df_aditivos": df_aditivos,
        "df_comparativo": df_comparativo,
    }
    return resultado


def formatar_dataframe_moeda(df, colunas_moeda=None, colunas_fator=None, colunas_pct=None):
    visual = df.copy()
    for col in colunas_moeda or []:
        if col in visual.columns:
            visual[col] = visual[col].apply(moeda)
    for col in colunas_fator or []:
        if col in visual.columns:
            visual[col] = visual[col].apply(fator_fmt)
    for col in colunas_pct or []:
        if col in visual.columns:
            visual[col] = visual[col].apply(lambda x: percentual(x, 2))
    return visual


# ============================================================
# Interface
# ============================================================

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Valor Global do Contrato")

st.markdown(
    """
    Este módulo processa o **Arquivo de Coleta** preenchido e apresenta uma visão executiva da execução financeira, dos deltas por ciclo e do valor atualizado do contrato.
    """
)

adm = st.session_state.get("dados_admissibilidade")

with st.expander("Contexto da Admissibilidade", expanded=True):
    if adm:
        col1, col2, col3 = st.columns(3)
        col1.metric("Origem", adm.get("origem") or adm.get("tipo", "Não informado"))
        col2.metric("Índice", adm.get("indice", "Não informado"))
        fator_adm = adm.get("fator_acumulado") or adm.get("fator") or 1.0
        col3.metric("Fator herdado", fator_fmt(fator_adm))
        st.caption("Dados herdados da sessão atual. Caso o arquivo contenha parâmetros próprios, eles também serão lidos para conferência.")
    else:
        st.warning(
            "Os dados de admissibilidade não foram encontrados na sessão atual. "
            "A ferramenta utilizará os parâmetros constantes do Arquivo de Coleta."
        )

st.subheader("Carregar Arquivo de Coleta")
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
    st.subheader("Dashboard Executivo")

    valor_destacado = moeda(resultado.get("valor_atualizado_contrato", resultado.get("valor_global_estoque", 0)))
    st.markdown(
        f"""
        <div style="background:#EAF2F8;border:1px solid #C8D9EA;border-radius:14px;padding:20px 24px;margin:10px 0 22px 0;">
            <div style="font-size:0.95rem;color:#334155;font-weight:600;margin-bottom:6px;">Valor Atualizado do Contrato</div>
            <div style="font-size:2.1rem;color:#163A5F;font-weight:800;line-height:1.2;">{valor_destacado}</div>
            <div style="font-size:0.85rem;color:#64748B;margin-top:6px;">Composição financeira: execução atualizada, remanescente atualizado e aditivos/supressões quando informados.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Valor original", moeda(resultado["valor_original_contrato"]))
    col2.metric("Valor pago efetivo", moeda(resultado["total_pago_faturado"]))
    col3.metric("Valor teórico calculado", moeda(resultado["total_devido_reajustado"]))
    col4.metric("Valor represado a pagar", moeda(resultado.get("valor_represado_a_pagar", 0)))

    col5, col6, col7 = st.columns(3)
    col5.metric("Delta total", moeda(resultado["delta_acumulado"]))
    col6.metric("Saldo remanescente atualizado", moeda(resultado["remanescente_reajustado"]))
    col7.metric("Aditivos/Supressões atualizados", moeda(resultado.get("total_aditivos_atualizados", 0)))

    tab1, tab2, tab3, tab4 = st.tabs([
        "Financeiro por Ciclo",
        "Valor Atualizado do Contrato",
        "Aditivos e Supressões",
        "Dados de Apoio",
    ])

    with tab1:
        st.markdown("### Painel Financeiro por Ciclo")
        df_fin = formatar_dataframe_moeda(
            resultado["df_financeiro_por_ciclo"],
            colunas_moeda=["Valor pago/faturado", "Valor devido reajustado", "Delta do ciclo", "Delta acumulado"],
            colunas_fator=["Fator aplicado"],
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
        composicao = pd.DataFrame([
            {"Componente": "Valor original do contrato", "Valor": resultado.get("valor_original_contrato", 0)},
            {"Componente": "Valor executado atualizado", "Valor": resultado.get("valor_executado_atualizado", 0)},
            {"Componente": "Saldo remanescente atualizado", "Valor": resultado.get("remanescente_reajustado", 0)},
            {"Componente": "Aditivos/Supressões atualizados", "Valor": resultado.get("total_aditivos_atualizados", 0)},
            {"Componente": "Valor Atualizado do Contrato", "Valor": resultado.get("valor_atualizado_contrato", 0)},
        ])
        st.dataframe(formatar_dataframe_moeda(composicao, colunas_moeda=["Valor"]), use_container_width=True, hide_index=True)

        st.markdown("### Remanescentes Financeiros por Ciclo")
        df_rem = formatar_dataframe_moeda(
            resultado["df_remanescentes"],
            colunas_moeda=["Remanescente original", "Remanescente reajustado"],
            colunas_fator=["Fator aplicado"],
        )
        st.dataframe(df_rem, use_container_width=True, hide_index=True)

    with tab3:
        df_ad = resultado.get("df_aditivos", pd.DataFrame())
        if df_ad is None or df_ad.empty:
            st.info("Nenhum aditivo/supressão informado no Arquivo de Coleta.")
        else:
            st.dataframe(
                formatar_dataframe_moeda(
                    df_ad,
                    colunas_moeda=["Valor original da alteração", "Valor atualizado da alteração"],
                    colunas_fator=["Fator aplicado"],
                ),
                use_container_width=True,
                hide_index=True,
            )

    with tab4:
        st.markdown("### Ciclos e Fatores Utilizados")
        df_ciclos = resultado["df_ciclos"].copy()
        if not df_ciclos.empty:
            df_ciclos_vis = formatar_dataframe_moeda(
                df_ciclos,
                colunas_fator=["Fator", "Fator acumulado", "Fator acumulado efetivo", "Fator ciclo efetivo"],
                colunas_pct=["Variação"],
            )
            st.dataframe(df_ciclos_vis, use_container_width=True, hide_index=True)

        st.markdown("### Quadro Executivo")
        df_comp = formatar_dataframe_moeda(
            resultado["df_comparativo"],
            colunas_moeda=["Antes do Reajuste", "Após Reajuste", "Diferença"],
        )
        st.dataframe(df_comp, use_container_width=True, hide_index=True)
else:
    st.info("Carregue e processe o Arquivo de Coleta para visualizar o Valor Global do contrato.")
