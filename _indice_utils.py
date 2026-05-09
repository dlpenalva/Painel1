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
