from dateutil.relativedelta import relativedelta

import pandas as pd
import requests


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



def carregar_icti_ipeadata(timeout=15):
    """Baixa a série ICTI/Ipea no Ipeadata.

    Série: DIMAC_ICTI2.
    Natureza: taxa de variação mensal (% a.m.).
    A função retorna DataFrame mensal com colunas data e taxa_mensal_pct.
    """
    url = "https://www.ipeadata.gov.br/api/odata4/ValoresSerie(SERCODIGO='DIMAC_ICTI2')"
    response = requests.get(url, timeout=timeout, headers={"Accept": "application/json"})
    response.raise_for_status()
    payload = response.json()
    registros = payload.get("value", []) if isinstance(payload, dict) else []
    linhas = []
    for item in registros:
        data = pd.to_datetime(item.get("VALDATA"), errors="coerce")
        valor = pd.to_numeric(item.get("VALVALOR"), errors="coerce")
        if pd.isna(data) or pd.isna(valor):
            continue
        data = pd.Timestamp(data.year, data.month, 1).normalize()
        linhas.append({"data": data, "taxa_mensal_pct": float(valor)})
    df = pd.DataFrame(linhas)
    if df.empty:
        raise ValueError("A série ICTI/Ipeadata retornou vazia ou sem linhas válidas.")
    df = df.sort_values("data").drop_duplicates(subset=["data"], keep="last").reset_index(drop=True)
    return df


def obter_ultima_competencia_icti_ipeadata(timeout=15):
    """Retorna metadados mínimos da última competência disponível do ICTI/Ipeadata."""
    df = carregar_icti_ipeadata(timeout=timeout)
    ultima = df.iloc[-1]
    meses = {
        1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
        7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
    }
    data = pd.Timestamp(ultima["data"])
    return {
        "data": data,
        "descricao": f"{meses[data.month]}/{data.year}",
        "taxa_mensal_pct": float(ultima["taxa_mensal_pct"]),
        "serie": "DIMAC_ICTI2",
        "fonte": "Ipeadata/Ipea",
    }


def calcular_icti_ipeadata(data_inicio, data_fim=None, timeout=15):
    """Calcula ICTI por produtório de taxas mensais da série DIMAC_ICTI2.

    Regra transparente adotada no cl8us:
    - data_inicio representa o mês/data da proposta ou âncora do ciclo;
    - a competência de índice-base é o mês imediatamente anterior;
    - o acumulado inclui as taxas do mês da proposta/âncora até o mês final;
    - isso reproduz a lógica de Índice Final / Índice Base quando a calculadora externa
      mostra, por exemplo, proposta em março/2023 com índice-base em fevereiro/2023.
    """
    inicio = pd.Timestamp(data_inicio.year, data_inicio.month, 1).normalize()
    if data_fim is None:
        fim_dt = data_inicio + relativedelta(months=11)
    else:
        fim_dt = data_fim
    fim = pd.Timestamp(fim_dt.year, fim_dt.month, 1).normalize()
    base_indice = (inicio - relativedelta(months=1)).normalize()

    if fim < inicio:
        return None

    df = carregar_icti_ipeadata(timeout=timeout)
    periodo = df[(df["data"] >= inicio) & (df["data"] <= fim)].copy()

    esperadas = pd.period_range(inicio.to_period("M"), fim.to_period("M"), freq="M")
    encontradas = set(periodo["data"].dt.to_period("M"))
    faltantes = [p for p in esperadas if p not in encontradas]
    if faltantes:
        return None

    if periodo.empty:
        return None

    periodo["valor"] = periodo["taxa_mensal_pct"]
    periodo["valor_decimal"] = periodo["taxa_mensal_pct"] / 100
    periodo["fator_mensal"] = 1 + periodo["valor_decimal"]
    periodo["fator_acumulado"] = periodo["fator_mensal"].cumprod()

    fator = float(periodo["fator_mensal"].prod())
    return {
        "variacao": fator - 1,
        "i_ini": 100.0,
        "i_fim": 100.0 * fator,
        "d_ini": inicio,
        "d_fim": fim,
        "d_indice_base": base_indice,
        "metodo": "Produtório de taxas mensais (ICTI/Ipeadata; base no mês anterior à proposta)",
        "dados": periodo[["data", "valor", "taxa_mensal_pct", "fator_mensal", "fator_acumulado"]].copy(),
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
