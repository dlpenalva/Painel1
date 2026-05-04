import re
import unicodedata
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Valor Global do Contrato", layout="wide")


# ============================================================
# Utilitários
# ============================================================

def agora_brasilia():
    return datetime.now(ZoneInfo("America/Sao_Paulo"))


def normalizar_texto(valor):
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


def numero_br(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip()
    if texto == "":
        return 0.0
    texto = texto.replace("R$", "").replace("\xa0", "").replace("%", "").strip()
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif texto.count(".") > 1:
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
    if valor is not None and not pd.isna(valor):
        texto = str(valor)
        n = numero_br(valor)
        if "%" in texto:
            return 1 + (n / 100)
        if n > 0:
            if n >= 2:
                return 1 + (n / 100)
            return n
    if variacao is not None:
        return 1 + percentual_para_decimal(variacao)
    return 1.0


def localizar_coluna(df, opcoes):
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
    if str(ciclo).upper().strip() == "C0":
        return 0
    m = re.search(r"([0-9]+)", str(ciclo))
    return int(m.group(1)) if m else 999


def situacao_bloqueia_reajuste(situacao):
    s = normalizar_texto(situacao)
    if "preclus" in s:
        return True
    if "adiantado" in s and "ressalva" not in s:
        return True
    return False


def periodo_competencia(valor):
    if pd.isna(valor):
        return None
    texto = str(valor).strip()
    if texto == "":
        return None
    m = re.search(r"(\d{1,2})[\/\-](\d{4})", texto)
    if m:
        return pd.Period(f"{int(m.group(2))}-{int(m.group(1)):02d}", freq="M")
    try:
        return pd.to_datetime(valor, dayfirst=True).to_period("M")
    except Exception:
        return None


def periodo_de_data(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return None
    try:
        return pd.to_datetime(valor, dayfirst=True).to_period("M")
    except Exception:
        return periodo_competencia(valor)


def mes_ano_de_data(valor):
    p = periodo_de_data(valor)
    if p is not None:
        return f"{p.month:02d}/{p.year}"
    return ""


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


# ============================================================
# Leitura do Excel
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
    return pd.DataFrame(ciclos)


def ler_ciclos(bytes_arquivo, xls):
    df_session = dataframe_ciclos_de_session_state()
    if not df_session.empty:
        return padronizar_ciclos(df_session), "session_state"
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
    col_base = localizar_coluna(df, ["Data-base", "Data Base", "Base", "data_base"])
    col_intervalo = localizar_coluna(df, ["Intervalo do índice", "Intervalo", "Janela", "intervalo_indice"])
    col_janela_adm = localizar_coluna(df, ["Janela de admissibilidade", "JanelaAdm", "janela_admissibilidade"])
    col_pedido = localizar_coluna(df, ["Data do pedido", "Pedido", "data_pedido"])
    col_situacao = localizar_coluna(df, ["Situação", "Resultado", "situacao"])
    col_variacao = localizar_coluna(df, ["Variação", "Variacao", "Percentual", "variacao"])
    col_fator = localizar_coluna(df, ["Fator", "fator"])
    col_fator_acum = localizar_coluna(df, ["Fator acumulado", "Fator acumulado final", "fator_acumulado"])

    linhas = []
    fator_acum_efetivo_calculado = 1.0
    fator_acum_nominal_calculado = 1.0

    for _, row in df.iterrows():
        ciclo = normalizar_ciclo(row.get(col_ciclo, "")) if col_ciclo else ""
        if not ciclo:
            continue
        situacao = row.get(col_situacao, "") if col_situacao else ""
        variacao = percentual_para_decimal(row.get(col_variacao, 0)) if col_variacao else 0.0
        fator = fator_de_valor(row.get(col_fator, None) if col_fator else None, variacao=variacao)

        fator_acum_nominal_calculado *= fator
        if col_fator_acum:
            fator_acum_nominal = fator_de_valor(row.get(col_fator_acum, None), variacao=None)
            if fator_acum_nominal <= 0:
                fator_acum_nominal = fator_acum_nominal_calculado
        else:
            fator_acum_nominal = fator_acum_nominal_calculado

        if situacao_bloqueia_reajuste(situacao):
            fator_efetivo = 1.0
        else:
            fator_efetivo = fator
            fator_acum_efetivo_calculado *= fator

        data_base = row.get(col_base, "") if col_base else ""
        pedido = row.get(col_pedido, "") if col_pedido else ""
        intervalo = row.get(col_intervalo, "") if col_intervalo else ""
        inicio_intervalo, fim_intervalo = parse_intervalo_mensal(intervalo)

        linhas.append({
            "Ciclo": ciclo,
            "Data-base": data_base,
            "Mês início do ciclo": mes_ano_de_data(data_base) or (f"{inicio_intervalo.month:02d}/{inicio_intervalo.year}" if inicio_intervalo else ""),
            "Intervalo do índice": intervalo,
            "Início intervalo": inicio_intervalo,
            "Fim intervalo": fim_intervalo,
            "Janela de admissibilidade": row.get(col_janela_adm, "") if col_janela_adm else "",
            "Data do pedido": pedido,
            "Mês início efeito financeiro": mes_ano_de_data(pedido),
            "Pedido período": periodo_de_data(pedido),
            "Situação": situacao,
            "Variação": variacao,
            "Percentual do ciclo": variacao,
            "Fator": fator,
            "Fator acumulado": fator_acum_nominal,
            "Percentual acumulado": fator_acum_efetivo_calculado - 1,
            "Fator acumulado efetivo": fator_acum_efetivo_calculado,
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
    resultado["Competência período"] = resultado["Competência"].apply(periodo_competencia)
    resultado["Valor pago/faturado"] = df[col_valor].apply(numero_br)
    if (resultado["Ciclo"] == "").all() and col_comp and not ciclos.empty:
        resultado["Ciclo"] = atribuir_ciclo_por_competencia(resultado["Competência"], ciclos)
    if (resultado["Ciclo"] == "").all():
        resultado["Ciclo"] = "C1"
    resultado = resultado[resultado["Valor pago/faturado"].fillna(0) != 0].copy()
    return resultado.reset_index(drop=True)


def atribuir_ciclo_por_competencia(competencias, ciclos):
    intervalos = []
    for _, row in ciclos.iterrows():
        ini = row.get("Início intervalo")
        fim = row.get("Fim intervalo")
        if ini is not None:
            intervalos.append((row["Ciclo"], ini, fim))
    atribuidos = []
    for comp in competencias:
        periodo = periodo_competencia(comp)
        ciclo_encontrado = ""
        if periodo is not None:
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
    df = ler_aba_com_cabecalho(
        bytes_arquivo,
        aba,
        termos_obrigatorios=["Item", "Quantidade contratada", "Valor unitário original"],
    )
    if localizar_coluna(df, ["Item"]) is None or localizar_coluna(df, ["Quantidade contratada", "Qtd C0", "Quantidade", "Qtd"]) is None or localizar_coluna(df, ["Valor unitário original", "VU C0", "VU Original", "Valor Unitario"]) is None:
        df = ler_aba_com_cabecalho(bytes_arquivo, aba, termos_obrigatorios=["Item", "Qtd", "VU"])
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
    colunas_rem = []
    for col in df.columns:
        n = normalizar_texto(col)
        nome_original = str(col)
        possui_ciclo = re.search(r"C\s*\d+", nome_original, flags=re.IGNORECASE) is not None
        eh_inicio_ciclo = "inicio" in n or "inicial" in n or possui_ciclo
        eh_coluna_excluida = "atual" in n or "data_corte" in n or "corte" in n
        if "remanescente" in n and "valor" not in n and "total" not in n and eh_inicio_ciclo and not eh_coluna_excluida:
            colunas_rem.append(col)
    if not colunas_rem:
        col_qtd_rem = localizar_coluna(df, ["Qtd Remanescente", "Quantidade Remanescente", "Remanescente"])
        if col_qtd_rem:
            colunas_rem = [col_qtd_rem]
    for col in colunas_rem:
        itens[col] = df[col].apply(numero_br)
    return itens, colunas_rem


# ============================================================
# Cálculos
# ============================================================

def mapa_fatores(ciclos):
    mapa = {"C0": {"fator_acumulado_efetivo": 1.0, "variacao": 0.0, "percentual_acumulado": 0.0, "situacao": "BASE", "pedido_periodo": None}}
    if ciclos.empty:
        return mapa
    for _, row in ciclos.iterrows():
        mapa[row["Ciclo"]] = {
            "fator_acumulado_efetivo": float(row.get("Fator acumulado efetivo", 1.0) or 1.0),
            "fator_ciclo_efetivo": float(row.get("Fator ciclo efetivo", 1.0) or 1.0),
            "variacao": float(row.get("Variação", 0.0) or 0.0),
            "percentual_acumulado": float(row.get("Percentual acumulado", 0.0) or 0.0),
            "situacao": row.get("Situação", ""),
            "pedido_periodo": row.get("Pedido período"),
            "mes_inicio": row.get("Mês início do ciclo", ""),
            "mes_efeito": row.get("Mês início efeito financeiro", ""),
        }
    return mapa


def aplicar_efeito_financeiro_mensal(financeiro, ciclos):
    fatores = mapa_fatores(ciclos)
    linhas = []
    for _, row in financeiro.iterrows():
        ciclo = normalizar_ciclo(row.get("Ciclo", "")) or "C1"
        valor_pago = float(row.get("Valor pago/faturado", 0.0) or 0.0)
        comp_periodo = row.get("Competência período")
        info = fatores.get(ciclo, {})
        pedido_periodo = info.get("pedido_periodo")
        fator = info.get("fator_acumulado_efetivo", 1.0)
        situacao = info.get("situacao", "")

        if ciclo == "C0" or situacao_bloqueia_reajuste(situacao):
            com_efeito = False if ciclo != "C0" else True
            fator_aplicado = 1.0
            motivo = "Base C0" if ciclo == "C0" else "Ciclo sem reajuste aplicável"
        elif comp_periodo is not None and pedido_periodo is not None and comp_periodo < pedido_periodo:
            com_efeito = False
            fator_aplicado = 1.0
            motivo = "Competência anterior ao pedido"
        else:
            com_efeito = True
            fator_aplicado = fator
            motivo = "Competência com efeito financeiro"

        devido = valor_pago * fator_aplicado
        linhas.append({
            "Ciclo": ciclo,
            "Competência": row.get("Competência", ""),
            "Valor pago/faturado": valor_pago,
            "Com efeito financeiro": com_efeito,
            "Motivo": motivo,
            "Percentual acumulado aplicado": fator_aplicado - 1,
            "Valor devido reajustado": devido,
            "Retroativo": devido - valor_pago,
        })
    return pd.DataFrame(linhas)


def calcular_financeiro_por_ciclo(df_mensal, ciclos):
    if df_mensal.empty:
        return pd.DataFrame(columns=[
            "Ciclo", "Mês início do ciclo", "Mês início efeito financeiro", "Situação",
            "Meses lançados", "Meses com efeito financeiro", "Meses sem efeito financeiro",
            "Valor pago/faturado", "Valor executado atualizado", "Retroativo do ciclo"
        ])
    fatores = mapa_fatores(ciclos)
    linhas = []
    for ciclo, grupo in df_mensal.groupby("Ciclo", dropna=False):
        ciclo = normalizar_ciclo(ciclo) or "C1"
        info = fatores.get(ciclo, {})
        meses_total = int(grupo["Competência período"].notna().sum()) if "Competência período" in grupo.columns else len(grupo)
        meses_efeito = int(grupo["Com efeito financeiro"].sum()) if "Com efeito financeiro" in grupo.columns else 0
        meses_sem = max(meses_total - meses_efeito, 0)
        pago = float(grupo["Valor pago/faturado"].sum())
        devido = float(grupo["Valor devido reajustado"].sum())
        retro = float(grupo["Retroativo"].sum())
        linhas.append({
            "Ciclo": ciclo,
            "Mês início do ciclo": info.get("mes_inicio", ""),
            "Mês início efeito financeiro": info.get("mes_efeito", ""),
            "Situação": info.get("situacao", ""),
            "Percentual do ciclo": info.get("variacao", 0.0),
            "Percentual acumulado aplicado": info.get("percentual_acumulado", 0.0),
            "Meses lançados": meses_total,
            "Meses com efeito financeiro": meses_efeito,
            "Meses sem efeito financeiro": meses_sem,
            "Valor pago/faturado": pago,
            "Valor executado atualizado": devido,
            "Retroativo do ciclo": retro,
        })
    df = pd.DataFrame(linhas).sort_values(by="Ciclo", key=lambda s: s.map(numero_ciclo)).reset_index(drop=True)
    df["Retroativo acumulado"] = df["Retroativo do ciclo"].cumsum()
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
    for col in colunas_remanescentes:
        ciclo = ciclo_da_coluna_remanescente(col, fallback="C1")
        info = fatores.get(ciclo, {})
        fator = info.get("fator_acumulado_efetivo", 1.0)
        qtd_rem = itens[col].apply(numero_br)
        valor_original = qtd_rem * itens["Valor unitário original"]
        valor_reajustado = valor_original * fator
        linhas.append({
            "Ciclo": ciclo,
            "Mês início do ciclo": info.get("mes_inicio", ""),
            "Coluna de origem": col,
            "Quantidade remanescente": float(qtd_rem.sum()),
            "Remanescente original": float(valor_original.sum()),
            "Percentual acumulado aplicado": fator - 1,
            "Remanescente atualizado": float(valor_reajustado.sum()),
        })
    if not linhas:
        return pd.DataFrame(), None
    df = pd.DataFrame(linhas).sort_values(by="Ciclo", key=lambda s: s.map(numero_ciclo)).reset_index(drop=True)
    return df, df.iloc[-1]["Ciclo"]


def calcular_valor_total_por_ciclo(df_fin_por_ciclo, df_remanescentes, valor_original):
    linhas = []
    if df_remanescentes.empty:
        return pd.DataFrame()
    for _, rem in df_remanescentes.iterrows():
        ciclo = rem["Ciclo"]
        n = numero_ciclo(ciclo)
        exec_ate_anterior = 0.0
        partes = []
        for _, fin in df_fin_por_ciclo.iterrows():
            if numero_ciclo(fin["Ciclo"]) < n:
                exec_ate_anterior += float(fin.get("Valor executado atualizado", 0.0) or 0.0)
                partes.append(f"{fin['Ciclo']} executado atualizado")
        total = exec_ate_anterior + float(rem.get("Remanescente atualizado", 0.0) or 0.0)
        comp = " + ".join(partes + [f"remanescente atualizado em {ciclo}"])
        linhas.append({
            "Ciclo/Marco": ciclo,
            "Mês de referência": rem.get("Mês início do ciclo", ""),
            "Composição": comp,
            "Executado atualizado até ciclo anterior": exec_ate_anterior,
            "Remanescente atualizado": float(rem.get("Remanescente atualizado", 0.0) or 0.0),
            "Valor total atualizado do contrato": total,
        })
    linhas.insert(0, {
        "Ciclo/Marco": "C0",
        "Mês de referência": "Original",
        "Composição": "Valor original inicial do contrato",
        "Executado atualizado até ciclo anterior": 0.0,
        "Remanescente atualizado": valor_original,
        "Valor total atualizado do contrato": valor_original,
    })
    return pd.DataFrame(linhas)


def montar_comparativo(valor_original, total_pago, total_devido, delta_acumulado, rem_original, rem_atualizado, valor_global):
    return pd.DataFrame([
        {"Indicador": "Valor original inicial do contrato", "Valor": valor_original},
        {"Indicador": "Total pago/faturado", "Valor": total_pago},
        {"Indicador": "Total executado atualizado", "Valor": total_devido},
        {"Indicador": "Total retroativo a pagar", "Valor": delta_acumulado},
        {"Indicador": "Remanescente original no último ciclo", "Valor": rem_original},
        {"Indicador": "Remanescente atualizado no último ciclo", "Valor": rem_atualizado},
        {"Indicador": "Valor global atualizado do contrato", "Valor": valor_global},
    ])


def extrair_percentual_reajuste_legacy(bytes_arquivo, xls):
    for aba in xls.sheet_names:
        try:
            bruto = pd.read_excel(BytesIO(bytes_arquivo), sheet_name=aba, header=None)
        except Exception:
            continue
        for i in range(bruto.shape[0]):
            for j in range(bruto.shape[1]):
                cel = bruto.iat[i, j]
                if "percentual_reajuste" in normalizar_texto(cel):
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

    if ciclos.empty:
        perc_legacy = extrair_percentual_reajuste_legacy(bytes_arquivo, xls)
        fator = fator_de_valor(params.get("fator_acumulado_final") or params.get("fator_acumulado") or params.get("fator"))
        if fator == 1.0 and perc_legacy is not None:
            fator = 1 + perc_legacy
        ciclos = pd.DataFrame([{
            "Ciclo": "C1", "Data-base": "", "Mês início do ciclo": "", "Intervalo do índice": "",
            "Início intervalo": None, "Fim intervalo": None, "Janela de admissibilidade": "",
            "Data do pedido": "", "Mês início efeito financeiro": "", "Pedido período": None,
            "Situação": "", "Variação": fator - 1, "Percentual do ciclo": fator - 1,
            "Fator": fator, "Fator acumulado": fator, "Percentual acumulado": fator - 1,
            "Fator acumulado efetivo": fator, "Fator ciclo efetivo": fator,
        }])

    financeiro_mensal = aplicar_efeito_financeiro_mensal(financeiro, ciclos)
    df_fin_por_ciclo = calcular_financeiro_por_ciclo(financeiro_mensal, ciclos)
    df_rem, ciclo_ultimo_rem = calcular_remanescentes(itens, colunas_remanescentes, ciclos)
    valor_original_contrato = float(itens["Valor total original"].sum())

    total_pago = float(df_fin_por_ciclo["Valor pago/faturado"].sum()) if not df_fin_por_ciclo.empty else 0.0
    total_devido = float(df_fin_por_ciclo["Valor executado atualizado"].sum()) if not df_fin_por_ciclo.empty else 0.0
    delta_acumulado = float(df_fin_por_ciclo["Retroativo do ciclo"].sum()) if not df_fin_por_ciclo.empty else 0.0

    if not df_rem.empty:
        rem_original = float(df_rem.iloc[-1]["Remanescente original"])
        rem_atualizado = float(df_rem.iloc[-1]["Remanescente atualizado"])
    else:
        rem_original = 0.0
        rem_atualizado = 0.0

    valor_total_por_ciclo = calcular_valor_total_por_ciclo(df_fin_por_ciclo, df_rem, valor_original_contrato)
    valor_global_atualizado = float(valor_total_por_ciclo.iloc[-1]["Valor total atualizado do contrato"]) if not valor_total_por_ciclo.empty else total_devido + rem_atualizado

    indice = (
        st.session_state.get("dados_admissibilidade", {}).get("indice")
        or params.get("indice_utilizado")
        or params.get("indice")
        or "Não informado"
    )
    fator_acumulado = float(ciclos["Fator acumulado efetivo"].iloc[-1]) if "Fator acumulado efetivo" in ciclos.columns and not ciclos.empty else 1.0

    df_comparativo = montar_comparativo(
        valor_original_contrato, total_pago, total_devido, delta_acumulado, rem_original, rem_atualizado, valor_global_atualizado
    )

    return {
        "data_processamento": agora_brasilia().strftime("%d/%m/%Y %H:%M"),
        "origem_ciclos": origem_ciclos,
        "indice": indice,
        "fator_acumulado": fator_acumulado,
        "percentual_acumulado": fator_acumulado - 1,
        "valor_original_contrato": valor_original_contrato,
        "total_pago_faturado": total_pago,
        "total_devido_reajustado": total_devido,
        "delta_acumulado": delta_acumulado,
        "remanescente_original": rem_original,
        "remanescente_reajustado": rem_atualizado,
        "valor_global_atualizado": valor_global_atualizado,
        "ciclo_ultimo_remanescente": ciclo_ultimo_rem,
        "df_ciclos": ciclos,
        "df_financeiro_mensal": financeiro_mensal,
        "df_financeiro_por_ciclo": df_fin_por_ciclo,
        "df_itens_processado": itens,
        "df_remanescentes": df_rem,
        "df_valor_total_por_ciclo": valor_total_por_ciclo,
        "df_comparativo": df_comparativo,
    }


def formatar_dataframe(df, colunas_moeda=None, colunas_pct=None):
    if df is None or not isinstance(df, pd.DataFrame):
        return pd.DataFrame()
    visual = df.copy()
    for col in colunas_moeda or []:
        if col in visual.columns:
            visual[col] = visual[col].apply(moeda)
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
    Este módulo processa o **Arquivo de Coleta** preenchido, aplica a regra de efeito financeiro
    e consolida execução, remanescentes, retroativos e Valor Global atualizado do contrato.
    """
)

adm = st.session_state.get("dados_admissibilidade")
with st.expander("Contexto da Admissibilidade", expanded=True):
    if adm:
        col1, col2, col3 = st.columns(3)
        col1.metric("Origem", adm.get("origem") or adm.get("tipo", "Não informado"))
        col2.metric("Índice", adm.get("indice", "Não informado"))
        fator_adm = adm.get("fator_acumulado") or adm.get("fator") or 1.0
        col3.metric("Percentual acumulado herdado", percentual(float(fator_adm) - 1))
        st.caption("Os dados da sessão são utilizados como referência quando disponíveis; o Arquivo de Coleta também contém os parâmetros necessários.")
    else:
        st.warning(
            "Os dados de admissibilidade não foram encontrados na sessão atual. "
            "A ferramenta utilizará os parâmetros constantes do Arquivo de Coleta."
        )

st.subheader("Carregar Arquivo de Coleta")
arquivo = st.file_uploader("Carregue aqui o Arquivo de Coleta preenchido (.xlsx)", type=["xlsx"])

if arquivo is not None:
    if st.button("Processar Coleta Preenchida", type="primary"):
        try:
            resultado = processar_arquivo_coleta(arquivo.getvalue())
            st.session_state["resultado_valor_global"] = resultado
            st.success("Arquivo processado com sucesso. Resultados disponíveis abaixo e no módulo Relatório Global.")
        except Exception as exc:
            st.error(f"Não foi possível processar o arquivo: {exc}")

resultado = st.session_state.get("resultado_valor_global")

if not resultado:
    st.info("Carregue e processe o Arquivo de Coleta para visualizar o Valor Global do contrato.")
    st.stop()

st.divider()
st.subheader("Dashboard Executivo")

col1, col2, col3, col4 = st.columns(4)
col1.metric("Índice", resultado.get("indice", "Não informado"))
col2.metric("Percentual acumulado", percentual(resultado.get("percentual_acumulado", 0)))
col3.metric("Valor original do contrato", moeda(resultado["valor_original_contrato"]))
col4.metric("Valor global atualizado", moeda(resultado["valor_global_atualizado"]))

col5, col6, col7, col8 = st.columns(4)
col5.metric("Total pago/faturado", moeda(resultado["total_pago_faturado"]))
col6.metric("Total executado atualizado", moeda(resultado["total_devido_reajustado"]))
col7.metric("Total retroativo a pagar", moeda(resultado["delta_acumulado"]))
col8.metric("Remanescente atualizado", moeda(resultado["remanescente_reajustado"]))

st.markdown("### Percentuais e ciclos considerados")
df_ciclos_vis = formatar_dataframe(
    resultado["df_ciclos"],
    colunas_pct=["Percentual do ciclo", "Percentual acumulado", "Variação"],
)
cols_ciclos = [c for c in [
    "Ciclo", "Mês início do ciclo", "Data do pedido", "Mês início efeito financeiro",
    "Situação", "Percentual do ciclo", "Percentual acumulado"
] if c in df_ciclos_vis.columns]
st.dataframe(df_ciclos_vis[cols_ciclos], use_container_width=True, hide_index=True)

tab1, tab2, tab3, tab4 = st.tabs([
    "Execução e Retroativos",
    "Remanescentes",
    "Valor Total do Contrato",
    "Dados Processados",
])

with tab1:
    st.markdown("### Apuração de Retroativos")
    df_fin = formatar_dataframe(
        resultado["df_financeiro_por_ciclo"],
        colunas_moeda=["Valor pago/faturado", "Valor executado atualizado", "Retroativo do ciclo", "Retroativo acumulado"],
        colunas_pct=["Percentual do ciclo", "Percentual acumulado aplicado"],
    )
    st.dataframe(df_fin, use_container_width=True, hide_index=True)

    with st.expander("Ver competências financeiras importadas e regra de efeito financeiro"):
        df_mensal = formatar_dataframe(
            resultado["df_financeiro_mensal"],
            colunas_moeda=["Valor pago/faturado", "Valor devido reajustado", "Retroativo"],
            colunas_pct=["Percentual acumulado aplicado"],
        )
        st.dataframe(df_mensal, use_container_width=True, hide_index=True)

with tab2:
    st.markdown("### Remanescente atualizado por ciclo")
    df_rem = formatar_dataframe(
        resultado["df_remanescentes"],
        colunas_moeda=["Remanescente original", "Remanescente atualizado"],
        colunas_pct=["Percentual acumulado aplicado"],
    )
    st.dataframe(df_rem, use_container_width=True, hide_index=True)
    with st.expander("Ver itens importados"):
        df_itens = formatar_dataframe(
            resultado["df_itens_processado"],
            colunas_moeda=["Valor unitário original", "Valor total original"],
        )
        st.dataframe(df_itens, use_container_width=True, hide_index=True)

with tab3:
    st.markdown("### Valor total atualizado do contrato por ciclo")
    df_total = formatar_dataframe(
        resultado["df_valor_total_por_ciclo"],
        colunas_moeda=[
            "Executado atualizado até ciclo anterior",
            "Remanescente atualizado",
            "Valor total atualizado do contrato",
        ],
    )
    st.dataframe(df_total, use_container_width=True, hide_index=True)

    st.markdown("### Resumo para instrução")
    df_comp = formatar_dataframe(resultado["df_comparativo"], colunas_moeda=["Valor"])
    st.dataframe(df_comp, use_container_width=True, hide_index=True)

with tab4:
    st.markdown("### Estrutura consolidada disponível para o Relatório Global")
    st.code("st.session_state['resultado_valor_global'] foi atualizado com os resultados processados.", language="python")
