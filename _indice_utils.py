from dateutil.relativedelta import relativedelta

import pandas as pd
import requests


ICTI_SERCODIGO = "DIMAC_ICTI2"
ICTI_API_BASES = [
    "https://www.ipeadata.gov.br/api/odata4",
    "http://www.ipeadata.gov.br/api/odata4",
]

MESES_PT_ABREV = {
    1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
    7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez",
}

MESES_PT_EXTENSO = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
    7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}


def carregar_ist_local(caminho="ist.csv"):
    """Carrega o IST local aceitando os dois layouts usados no projeto.

    Layout novo/atual: MES_ANO;INDICE_NIVEL, com competências como jan/22.
    Layout antigo: data;indice.
    Retorna DataFrame padronizado com colunas data e indice.
    """
    df = pd.read_csv(caminho, sep=";", decimal=",", encoding="utf-8-sig")
    df.columns = [str(col).strip().lower() for col in df.columns]

    if "data" in df.columns and "indice" in df.columns:
        df["data"] = pd.to_datetime(df["data"], dayfirst=True, errors="coerce").dt.normalize()
        df["indice"] = pd.to_numeric(df["indice"], errors="coerce")
    elif "mes_ano" in df.columns and "indice_nivel" in df.columns:
        meses = {
            "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
            "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
        }

        def converter_mes_ano(valor):
            texto = str(valor).strip().lower()
            if "/" not in texto:
                return pd.NaT
            mes_txt, ano_txt = texto.split("/", 1)
            mes = meses.get(mes_txt[:3])
            if mes is None:
                return pd.NaT
            ano = int(ano_txt)
            if ano < 100:
                ano += 2000
            return pd.Timestamp(ano, mes, 1)

        df["data"] = df["mes_ano"].apply(converter_mes_ano)
        df["indice"] = pd.to_numeric(df["indice_nivel"], errors="coerce")
    else:
        raise KeyError("O arquivo ist.csv deve conter as colunas data/indice ou MES_ANO/INDICE_NIVEL.")

    df = df.dropna(subset=["data", "indice"]).sort_values("data")
    if df.empty:
        raise ValueError("O arquivo ist.csv não contém dados válidos.")
    return df[["data", "indice"]]


def calcular_ist_numero_indice(data_inicio, caminho="ist.csv"):
    """Calcula IST por divisão de número-índice entre o mês-base e o mesmo mês 12 meses depois."""
    df = carregar_ist_local(caminho)

    r_ini = pd.Timestamp(data_inicio.year, data_inicio.month, 1).normalize()
    marco_final = data_inicio + relativedelta(years=1)
    r_fim = pd.Timestamp(marco_final.year, marco_final.month, 1).normalize()

    v_ini_rows = df[df["data"].dt.to_period("M") == r_ini.to_period("M")]
    v_fim_rows = df[df["data"].dt.to_period("M") == r_fim.to_period("M")]

    if v_ini_rows.empty or v_fim_rows.empty:
        return None

    v_ini = float(v_ini_rows["indice"].iloc[0])
    v_fim = float(v_fim_rows["indice"].iloc[0])

    return {
        "variacao": (v_fim / v_ini) - 1,
        "i_ini": v_ini,
        "i_fim": v_fim,
        "d_ini": r_ini,
        "d_fim": r_fim,
        "metodo": "Divisão de Número-Índice (Série Local)",
        "dados": pd.DataFrame({"data": [r_ini, r_fim], "indice": [v_ini, v_fim]}),
    }


def coletar_sgs_produtorio(serie_codigo, data_inicio, data_fim, timeout=15):
    """Coleta série SGS/BCB e calcula a variação acumulada por produtório de taxas mensais."""
    url = (
        f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?"
        f"formato=json&dataInicial={data_inicio.strftime('%d/%m/%Y')}&dataFinal={data_fim.strftime('%d/%m/%Y')}"
    )
    response = requests.get(url, timeout=timeout)
    df = pd.DataFrame(response.json())
    if df.empty:
        return None

    df["valor_decimal"] = df["valor"].astype(float) / 100
    df["data"] = pd.to_datetime(df["data"], dayfirst=True)

    return {
        "variacao": (1 + df["valor_decimal"]).prod() - 1,
        "metodo": "Produtório de taxas mensais (SGS/BCB)",
        "dados": df[["data", "valor"]],
    }


def _ipeadata_get_json(endpoint, timeout=20):
    ultimo_erro = None
    headers = {
        "User-Agent": "Mozilla/5.0 cl8us-icti",
        "Accept": "application/json",
    }
    for base in ICTI_API_BASES:
        url = f"{base}/{endpoint}"
        try:
            resp = requests.get(url, headers=headers, timeout=timeout)
            resp.raise_for_status()
            return resp.json()
        except Exception as exc:
            ultimo_erro = exc
    raise RuntimeError(f"Não foi possível consultar o Ipeadata. Último erro: {ultimo_erro}")


def carregar_icti_ipeadata(timeout=20):
    """Carrega o ICTI mensal diretamente do Ipeadata.

    A série DIMAC_ICTI2 é retornada pelo Ipeadata como taxa de variação mensal (% a.m.).
    Para uso no cl8us, é criada também uma série de nível sintética, com base 100 acumulada
    a partir da primeira competência disponível. Isso permite exibir índice inicial/final,
    mas o cálculo principal permanece o produtório das taxas mensais.
    """
    dados = _ipeadata_get_json(f"ValoresSerie(SERCODIGO='{ICTI_SERCODIGO}')", timeout=timeout)
    registros = dados.get("value", []) if isinstance(dados, dict) else []
    if not registros:
        raise RuntimeError("A API do Ipeadata retornou a série ICTI vazia.")

    linhas = []
    for item in registros:
        data_raw = item.get("VALDATA") if isinstance(item, dict) else None
        valor_raw = item.get("VALVALOR") if isinstance(item, dict) else None
        data = pd.to_datetime(data_raw, errors="coerce")
        valor = pd.to_numeric(valor_raw, errors="coerce")
        if pd.isna(data) or pd.isna(valor):
            continue
        data = pd.Timestamp(year=int(data.year), month=int(data.month), day=1).normalize()
        linhas.append({
            "data": data,
            "mes_ano": f"{MESES_PT_ABREV[data.month]}/{str(data.year)[-2:]}",
            "taxa_mensal_percentual": float(valor),
            "valor": float(valor),
        })

    df = pd.DataFrame(linhas)
    if df.empty:
        raise RuntimeError("Nenhuma competência válida do ICTI foi identificada no Ipeadata.")

    df = df.sort_values("data").drop_duplicates(subset=["data"], keep="last").reset_index(drop=True)

    nivel_atual = 100.0
    niveis = []
    for _, row in df.iterrows():
        nivel_atual = nivel_atual * (1 + float(row["taxa_mensal_percentual"]) / 100)
        niveis.append(nivel_atual)
    df["indice_nivel_sintetico"] = niveis
    df["fator_mensal"] = 1 + df["taxa_mensal_percentual"] / 100
    return df


def obter_ultima_competencia_icti_ipeadata(timeout=20):
    """Retorna a última competência do ICTI disponível no Ipeadata."""
    df = carregar_icti_ipeadata(timeout=timeout)
    ultima = df.iloc[-1]
    data = pd.Timestamp(ultima["data"])
    return {
        "data": data,
        "mes_ano": ultima["mes_ano"],
        "descricao": f"{MESES_PT_EXTENSO[data.month]}/{data.year}",
        "taxa_mensal_percentual": float(ultima["taxa_mensal_percentual"]),
        "sercodigo": ICTI_SERCODIGO,
        "serie": ICTI_SERCODIGO,
    }


def calcular_icti_ipeadata(data_inicio, data_fim=None, timeout=20):
    """Calcula ICTI automaticamente via Ipeadata.

    Regra transparente adotada para o cl8us:
    - data_inicio representa a data/mês da proposta ou âncora informada;
    - a competência do índice-base utilizada é o mês anterior a data_inicio;
    - data_fim representa a competência final do ciclo;
    - se data_fim não for informada, usa data_inicio + 11 meses;
    - o fator é o produtório das taxas mensais do ICTI entre o mês da proposta/âncora
      e a competência final, inclusive.

    Exemplo: proposta/âncora mar/2023 e final fev/2026 usa índice-base fev/2023
    e acumula mar/2023 até fev/2026.
    """
    if data_inicio is None:
        return None

    data_inicio_ts = pd.Timestamp(data_inicio)
    competencia_proposta = pd.Timestamp(data_inicio_ts.year, data_inicio_ts.month, 1).normalize()
    competencia_base = (competencia_proposta - relativedelta(months=1)).normalize()

    if data_fim is None:
        data_fim_ts = data_inicio_ts + relativedelta(months=11)
    else:
        data_fim_ts = pd.Timestamp(data_fim)
    competencia_final = pd.Timestamp(data_fim_ts.year, data_fim_ts.month, 1).normalize()

    if competencia_final < competencia_proposta:
        return None

    df = carregar_icti_ipeadata(timeout=timeout)
    datas = set(df["data"])

    if competencia_base not in datas or competencia_final not in datas:
        return None

    periodo = df[(df["data"] > competencia_base) & (df["data"] <= competencia_final)].copy()
    if periodo.empty:
        return None

    fator = float(periodo["fator_mensal"].prod())
    variacao = fator - 1

    linha_base = df[df["data"] == competencia_base].iloc[0]
    linha_final = df[df["data"] == competencia_final].iloc[0]

    periodo["fator_acumulado_progressivo"] = periodo["fator_mensal"].cumprod()
    dados = periodo[["data", "taxa_mensal_percentual", "fator_mensal", "fator_acumulado_progressivo"]].copy()
    dados = dados.rename(columns={"taxa_mensal_percentual": "valor"})

    return {
        "variacao": variacao,
        "var": variacao,
        "i_ini": float(linha_base["indice_nivel_sintetico"]),
        "i_fim": float(linha_final["indice_nivel_sintetico"]),
        "d_ini": competencia_base,
        "d_fim": competencia_final,
        "p_ini": competencia_base,
        "p_fim": competencia_final,
        "competencia_proposta": competencia_proposta,
        "competencia_indice_base": competencia_base,
        "competencia_final": competencia_final,
        "d_proposta_ancora": competencia_proposta,
        "d_indice_base": competencia_base,
        "d_final_icti": competencia_final,
        "metodo": "ICTI/Ipeadata: produtório das taxas mensais; índice-base = mês anterior à proposta/âncora",
        "dados": dados,
        "sercodigo": ICTI_SERCODIGO,
        "serie": ICTI_SERCODIGO,
    }
