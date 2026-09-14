import time
import math
from pathlib import Path

from dateutil.relativedelta import relativedelta

import pandas as pd
import requests


# Identidade historica da serie usada pelo projeto e codigo atualmente exposto
# pelo catalogo OData. Em 14/09/2026, o Ipeadata deixou DIMAC_ICTI2 sem valores
# e passou a publicar a mesma serie mensal (% a.m.) como DIMAC12_ICTI2.
ICTI_SERCODIGO = "DIMAC_ICTI2"
ICTI_SERCODIGO_IPEADATA = "DIMAC12_ICTI2"
ICTI_CSV_PADRAO = Path(__file__).resolve().with_name("icti.csv")
ICTI_API_BASE = "https://www.ipeadata.gov.br/api/odata4"

MESES_PT_ABREV = {
    1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
    7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez",
}

MESES_PT_EXTENSO = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
    7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}

ESTADO_INDICE_OK = "OK"
ESTADO_FONTE_INDISPONIVEL = "FONTE_INDISPONIVEL"
ESTADO_COMPETENCIAS_AUSENTES = "COMPETENCIAS_AUSENTES"
ESTADO_FALLBACK_LOCAL = "FALLBACK_LOCAL"


class CompetenciasIndiceAusentes(RuntimeError):
    """Série obtida com sucesso, mas sem todas as competências necessárias."""

    def __init__(self, faltantes, encontradas=()):
        super().__init__("A série consultada não cobre todas as competências necessárias.")
        self.faltantes = list(faltantes)
        self.encontradas = list(encontradas)


class FonteIndiceIndisponivel(RuntimeError):
    """A fonte oficial não pôde ser tecnicamente consultada."""

    def __init__(self, fonte, detalhe=None, contexto=None):
        super().__init__(f"Fonte {fonte} indisponível")
        self.fonte = fonte
        self.detalhe = detalhe
        self.contexto = dict(contexto or {})


def competencias_mensais(data_inicio, data_fim):
    """Competências mm/aaaa entre os meses inicial e final, inclusive."""
    inicio = pd.Timestamp(data_inicio).to_period("M")
    fim = pd.Timestamp(data_fim).to_period("M")
    if fim < inicio:
        return []
    return [p.strftime("%m/%Y") for p in pd.period_range(inicio, fim, freq="M")]


def _competencias_do_dataframe(df):
    if not isinstance(df, pd.DataFrame) or "data" not in df.columns:
        return []
    datas = pd.to_datetime(df["data"], dayfirst=True, errors="coerce").dropna()
    return sorted({data.to_period("M").strftime("%m/%Y") for data in datas})


def _diagnostico_indice(estado, fonte, esperadas, encontradas=(), faltantes=(), detalhe=None):
    return {
        "estado": estado,
        "ok": estado in {ESTADO_INDICE_OK, ESTADO_FALLBACK_LOCAL},
        "fonte": fonte,
        "esperadas": list(esperadas),
        "encontradas": list(encontradas),
        "faltantes": list(faltantes),
        "detalhe_tecnico": detalhe,
    }


def _executar_consulta_diagnosticada(funcao, *, fonte, esperadas):
    try:
        resultado = funcao()
    except CompetenciasIndiceAusentes as exc:
        diagnostico = _diagnostico_indice(
            ESTADO_COMPETENCIAS_AUSENTES,
            fonte,
            esperadas,
            exc.encontradas,
            exc.faltantes,
        )
        return {"resultado": None, "diagnostico": diagnostico}
    except FonteIndiceIndisponivel as exc:
        diagnostico = _diagnostico_indice(
            ESTADO_FONTE_INDISPONIVEL,
            exc.fonte,
            esperadas,
            detalhe=repr(exc.detalhe),
        )
        diagnostico.update(exc.contexto)
        return {"resultado": None, "diagnostico": diagnostico}
    except Exception as exc:
        diagnostico = _diagnostico_indice(
            ESTADO_FONTE_INDISPONIVEL,
            fonte,
            esperadas,
            detalhe=repr(exc),
        )
        return {"resultado": None, "diagnostico": diagnostico}

    if resultado is None:
        diagnostico = _diagnostico_indice(
            ESTADO_COMPETENCIAS_AUSENTES,
            fonte,
            esperadas,
            faltantes=esperadas,
        )
        return {"resultado": None, "diagnostico": diagnostico}

    encontradas = _competencias_do_dataframe(resultado.get("dados"))
    estado = (
        ESTADO_FALLBACK_LOCAL
        if resultado.get("fonte") == "local" and resultado.get("fonte_oficial_indisponivel")
        else ESTADO_INDICE_OK
    )
    diagnostico = _diagnostico_indice(estado, fonte, esperadas, encontradas)
    for chave in (
        "fonte",
        "fonte_original",
        "sercodigo",
        "sercodigo_ipeadata",
        "ultima_competencia_local",
    ):
        if resultado.get(chave) is not None:
            diagnostico[chave] = resultado[chave]
    return {"resultado": resultado, "diagnostico": diagnostico}


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


FONTE_IST_ANATEL = (
    "https://www.gov.br/anatel/pt-br/regulado/competicao/"
    "tarifas-e-precos/valores-do-ist"
)

# Cache compartilhado da serie IST vigente (coerente com o TTL de 1h dos demais
# indices). Serve tanto ao calculo quanto ao aviso da UI, evitando consultar a
# Anatel a cada rerun.
_IST_TTL_SEGUNDOS = 60 * 60
_ist_cache: dict = {"expira_em": 0.0, "df": None, "fonte": None}


def _resetar_cache_ist() -> None:
    """Zera o cache do IST (uso em testes)."""
    _ist_cache.update(expira_em=0.0, df=None, fonte=None)


def carregar_ist_anatel(timeout=15):
    """Baixa a serie oficial do IST na Anatel e devolve DataFrame [data, indice].

    Reutiliza o parser homologado de tools/atualizar_ist_anatel.py (identificacao
    semantica das tabelas: coluna Referencia e coluna IST; a Variacao serve apenas
    de conferencia). Datas no primeiro dia da competencia, numero-indice em float,
    ordenado e sem duplicidades. Lanca excecao em falha de rede/estrutura.
    """
    from tools.atualizar_ist_anatel import baixar_registros_ist

    registros = baixar_registros_ist(timeout=timeout)
    df = pd.DataFrame(
        [{"data": pd.Timestamp(r.competencia), "indice": float(r.indice)} for r in registros]
    )
    df = (
        df.dropna(subset=["data", "indice"])
        .drop_duplicates(subset=["data"], keep="last")
        .sort_values("data")
        .reset_index(drop=True)
    )
    if df.empty:
        raise ValueError("A serie IST/Anatel retornou vazia.")
    return df[["data", "indice"]]


def carregar_ist_atual(caminho="ist.csv", *, ttl=_IST_TTL_SEGUNDOS, timeout=15, _agora=None):
    """Serie IST vigente com fonte: Anatel (primaria, cache TTL) -> ist.csv (fallback).

    Retorna (df, fonte) com fonte em {'anatel', 'local'}. Nunca quebra por rede:
    qualquer falha ao consultar a Anatel (timeout, HTTP, HTML inesperado) cai para
    a base local. Nao inventa nem extrapola competencias.
    """
    agora = time.monotonic() if _agora is None else _agora
    if _ist_cache["df"] is not None and agora < _ist_cache["expira_em"]:
        return _ist_cache["df"], _ist_cache["fonte"]
    try:
        df = carregar_ist_anatel(timeout=timeout)
        fonte = "anatel"
        df.attrs["fonte_oficial_indisponivel"] = False
    except Exception as erro_anatel:
        try:
            df = carregar_ist_local(caminho)
        except Exception as erro_local:
            raise FonteIndiceIndisponivel(
                "Anatel",
                {"erro_anatel": repr(erro_anatel), "erro_fallback_local": repr(erro_local)},
            ) from erro_local
        fonte = "local"
        df.attrs["fonte_oficial_indisponivel"] = True
        df.attrs["erro_fonte_oficial"] = repr(erro_anatel)
    _ist_cache.update(df=df, fonte=fonte, expira_em=agora + ttl)
    return df, fonte


def calcular_ist_numero_indice(data_inicio, caminho="ist.csv", *, _diagnostico=False):
    """Calcula IST por divisão de número-índice entre o mês-base e o mesmo mês 12 meses depois.

    Fonte da série: Anatel (oficial, quando disponível) com fallback para ist.csv.
    A metodologia matemática (v_fim / v_ini) e a memória mensal REAL permanecem
    inalteradas — muda apenas a origem/atualização da série.
    """
    df, fonte = carregar_ist_atual(caminho)

    r_ini = pd.Timestamp(data_inicio.year, data_inicio.month, 1).normalize()
    marco_final = data_inicio + relativedelta(years=1)
    r_fim = pd.Timestamp(marco_final.year, marco_final.month, 1).normalize()

    v_ini_rows = df[df["data"].dt.to_period("M") == r_ini.to_period("M")]
    v_fim_rows = df[df["data"].dt.to_period("M") == r_fim.to_period("M")]

    if v_ini_rows.empty or v_fim_rows.empty:
        if _diagnostico:
            encontradas = _competencias_do_dataframe(df)
            faltantes = [
                data.strftime("%m/%Y")
                for data, linhas in ((r_ini, v_ini_rows), (r_fim, v_fim_rows))
                if linhas.empty
            ]
            if fonte == "local" and df.attrs.get("fonte_oficial_indisponivel"):
                raise FonteIndiceIndisponivel(
                    "Anatel",
                    {
                        "erro_anatel": df.attrs.get("erro_fonte_oficial"),
                        "fallback_local_faltantes": faltantes,
                    },
                )
            raise CompetenciasIndiceAusentes(faltantes, encontradas)
        return None

    v_ini = float(v_ini_rows["indice"].iloc[0])
    v_fim = float(v_fim_rows["indice"].iloc[0])

    # Memoria canonica: preserva a serie mensal REAL do ist.csv no intervalo
    # [mes-base, mes-base + 12 meses], sem interpolar nem fabricar competencias.
    # O RESULTADO continua sendo calculado pelo metodo homologado (v_fim / v_ini);
    # apenas a auditoria/memoria passa a mostrar todas as competencias do periodo.
    serie_periodo = (
        df[(df["data"] >= r_ini) & (df["data"] <= r_fim)][["data", "indice"]]
        .sort_values("data")
        .reset_index(drop=True)
    )

    metodo = (
        "Divisão de Número-Índice (IST/Anatel)" if fonte == "anatel"
        else "Divisão de Número-Índice (IST/base local)"
    )
    return {
        "variacao": (v_fim / v_ini) - 1,
        "i_ini": v_ini,
        "i_fim": v_fim,
        "d_ini": r_ini,
        "d_fim": r_fim,
        "fonte": fonte,
        "fonte_oficial_indisponivel": bool(df.attrs.get("fonte_oficial_indisponivel")),
        "metodo": metodo,
        "dados": serie_periodo,
    }


def coletar_sgs_produtorio(serie_codigo, data_inicio, data_fim, timeout=15, *, _diagnostico=False):
    """Coleta série SGS/BCB e calcula a variação acumulada por produtório de taxas mensais."""
    url = (
        f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?"
        f"formato=json&dataInicial={data_inicio.strftime('%d/%m/%Y')}&dataFinal={data_fim.strftime('%d/%m/%Y')}"
    )
    response = requests.get(url, timeout=timeout)
    response.raise_for_status()
    df = pd.DataFrame(response.json())
    if df.empty:
        if _diagnostico:
            raise FonteIndiceIndisponivel(
                "SGS/BCB", "A fonte respondeu sem uma série utilizável."
            )
        return None

    df["valor_decimal"] = df["valor"].astype(float) / 100
    df["data"] = pd.to_datetime(df["data"], dayfirst=True)

    if _diagnostico:
        esperadas = competencias_mensais(data_inicio, data_fim)
        encontradas = _competencias_do_dataframe(df)
        faltantes = [c for c in esperadas if c not in set(encontradas)]
        if faltantes:
            raise CompetenciasIndiceAusentes(faltantes, encontradas)

    return {
        "variacao": (1 + df["valor_decimal"]).prod() - 1,
        "metodo": "Produtório de taxas mensais (SGS/BCB)",
        "dados": df[["data", "valor"]],
    }


SGS_IPCA = 433
SGS_IGPM = 189
SGS_INPC = 188


def serie_sgs_do_indice(tipo_idx: str) -> str:
    """Codigo da serie SGS/BCB a partir do rotulo do indice selecionado.

    IPCA -> 433, IGP-M -> 189, INPC -> 188. Fonte unica para as calculadoras
    (1 ciclo e multiciclo) — evita o dispatch binario que confundia INPC/IGP-M.
    """
    t = str(tipo_idx or "").upper()
    if "IPCA" in t:
        return str(SGS_IPCA)
    if "INPC" in t:
        return str(SGS_INPC)
    return str(SGS_IGPM)  # IGP-M (padrao dos SGS restantes)


def obter_ultima_competencia_sgs(serie_codigo, timeout=15):
    """Última competência de uma série mensal do SGS/BCB (mesma fonte do cálculo).

    Usada por IPCA (SGS 433) e IGP-M (SGS 189). Consulta o endpoint oficial
    ``.../dados/ultimos/1`` e devolve a competência mais recente sem hard-code de
    mês/ano. Levanta excecao em falha de rede/parse; o chamador decide o fallback.
    """
    url = (
        f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{int(serie_codigo)}/"
        "dados/ultimos/1?formato=json"
    )
    resp = requests.get(url, timeout=timeout)
    resp.raise_for_status()
    dados = resp.json()
    if not dados:
        raise RuntimeError(f"A série SGS {serie_codigo} retornou vazia.")
    data_raw = dados[-1].get("data")
    data = pd.to_datetime(data_raw, dayfirst=True, errors="coerce")
    if pd.isna(data):
        raise RuntimeError(f"Competência inválida na série SGS {serie_codigo}: {data_raw!r}.")
    data = pd.Timestamp(year=int(data.year), month=int(data.month), day=1).normalize()
    return {
        "data": data,
        "mes_ano": f"{data.month:02d}/{data.year}",
        "descricao": f"{data.month:02d}/{data.year}",
        "serie": int(serie_codigo),
    }


def _ipeadata_get_json(endpoint, timeout=20):
    headers = {
        "User-Agent": "Mozilla/5.0 cl8us-icti",
        "Accept": "application/json",
    }
    url = f"{ICTI_API_BASE}/{endpoint}"
    try:
        resp = requests.get(url, headers=headers, timeout=timeout)
        resp.raise_for_status()
        return resp.json()
    except Exception as exc:
        raise RuntimeError(f"Não foi possível consultar o Ipeadata por HTTPS: {exc}") from exc


def _finalizar_serie_icti(df, *, origem):
    """Valida e completa uma serie ICTI sem estimar ou deduplicar dados."""
    obrigatorias = {"data", "taxa_mensal_percentual"}
    if not isinstance(df, pd.DataFrame) or not obrigatorias.issubset(df.columns):
        raise ValueError(f"A série ICTI {origem} possui estrutura inválida.")
    if df.empty:
        raise ValueError(f"A série ICTI {origem} está vazia.")

    serie = df.copy()
    # O OData mistura offsets -02:00 e -03:00 ao longo da serie historica.
    # Normalizar em UTC evita que pandas trate essa variacao de fuso como data
    # invalida; a competencia mensal continua sendo o mesmo primeiro dia.
    datas = pd.to_datetime(serie["data"], errors="coerce", utc=True).dt.tz_convert(None)
    valores = pd.to_numeric(serie["taxa_mensal_percentual"], errors="coerce")
    if datas.isna().any() or valores.isna().any():
        raise ValueError(f"A série ICTI {origem} contém data ou taxa inválida.")

    serie["data"] = datas.apply(
        lambda data: pd.Timestamp(year=int(data.year), month=int(data.month), day=1).normalize()
    )
    serie["taxa_mensal_percentual"] = valores.astype(float)
    if not serie["taxa_mensal_percentual"].map(math.isfinite).all():
        raise ValueError(f"A série ICTI {origem} contém taxa não finita.")
    duplicadas = serie.loc[serie["data"].duplicated(keep=False), "data"]
    if not duplicadas.empty:
        competencia = duplicadas.iloc[0].strftime("%m/%Y")
        raise ValueError(f"A série ICTI {origem} contém competência duplicada: {competencia}.")
    if not serie["data"].is_monotonic_increasing:
        raise ValueError(f"A série ICTI {origem} não está em ordem cronológica.")

    serie = serie.reset_index(drop=True)
    serie["mes_ano"] = serie["data"].apply(
        lambda data: f"{MESES_PT_ABREV[data.month]}/{str(data.year)[-2:]}"
    )
    serie["valor"] = serie["taxa_mensal_percentual"]
    serie["fator_mensal"] = 1 + serie["taxa_mensal_percentual"] / 100
    serie["indice_nivel_sintetico"] = 100 * serie["fator_mensal"].cumprod()
    return serie


def carregar_icti_local(caminho=ICTI_CSV_PADRAO):
    """Carrega a copia local oficial do ICTI (nao e uma serie independente)."""
    df = pd.read_csv(
        caminho,
        sep=";",
        dtype=str,
        keep_default_na=False,
        encoding="utf-8-sig",
    )
    df.columns = [str(coluna).strip().lower() for coluna in df.columns]
    if set(df.columns) != {"competencia", "taxa_mensal_percentual"}:
        raise ValueError(
            "O icti.csv deve conter somente COMPETENCIA;TAXA_MENSAL_PERCENTUAL."
        )
    competencia = df["competencia"].astype(str).str.strip()
    if not competencia.str.fullmatch(r"\d{4}-(0[1-9]|1[0-2])").all():
        raise ValueError("O icti.csv contém competência inválida; use AAAA-MM.")
    taxa_texto = df["taxa_mensal_percentual"].astype(str).str.strip()
    if not taxa_texto.str.fullmatch(r"[+-]?\d+(?:[.,]\d+)?").all():
        raise ValueError("O icti.csv contém taxa mensal não numérica.")
    bruto = pd.DataFrame(
        {
            "data": pd.to_datetime(competencia, format="%Y-%m", errors="coerce"),
            # float() preserva a mesma conversao binaria usada pelo JSON do
            # requests. pd.to_numeric pode arredondar os ultimos digitos de
            # decimais longos e quebrar a igualdade valor a valor do espelho.
            "taxa_mensal_percentual": taxa_texto.map(
                lambda texto: float(texto.replace(",", "."))
            ),
        }
    )
    serie = _finalizar_serie_icti(bruto, origem="local")
    serie.attrs.update(
        fonte="local",
        fonte_original="Ipeadata/Ipea",
        serie=ICTI_SERCODIGO,
        sercodigo_ipeadata=ICTI_SERCODIGO_IPEADATA,
    )
    return serie


def carregar_icti_ipeadata(timeout=20):
    """Carrega o ICTI mensal diretamente do Ipeadata.

    A identidade historica DIMAC_ICTI2 e atualmente exposta pelo Ipeadata sob o
    codigo DIMAC12_ICTI2, como taxa de variacao mensal (% a.m.).
    Para uso no cl8us, é criada também uma série de nível sintética, com base 100 acumulada
    a partir da primeira competência disponível. Isso permite exibir índice inicial/final,
    mas o cálculo principal permanece o produtório das taxas mensais.
    """
    dados = _ipeadata_get_json(
        f"ValoresSerie(SERCODIGO='{ICTI_SERCODIGO_IPEADATA}')", timeout=timeout
    )
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
            raise ValueError("A API do Ipeadata retornou registro ICTI inválido.")
        linhas.append({
            "data": data,
            "taxa_mensal_percentual": float(valor),
        })

    df = pd.DataFrame(linhas)
    df = _finalizar_serie_icti(df, origem="do Ipeadata")
    df.attrs.update(
        fonte="ipeadata",
        fonte_original="Ipeadata/Ipea",
        serie=ICTI_SERCODIGO,
        sercodigo_ipeadata=ICTI_SERCODIGO_IPEADATA,
        fonte_oficial_indisponivel=False,
    )
    return df


def carregar_icti_atual(caminho=ICTI_CSV_PADRAO, *, timeout=20):
    """Serie ICTI vigente: Ipeadata oficial e, em falha tecnica, copia local."""
    try:
        return carregar_icti_ipeadata(timeout=timeout), "ipeadata"
    except Exception as erro_ipeadata:
        try:
            df = carregar_icti_local(caminho)
        except Exception as erro_local:
            raise FonteIndiceIndisponivel(
                "Ipeadata",
                {
                    "erro_ipeadata": repr(erro_ipeadata),
                    "erro_fallback_local": repr(erro_local),
                },
            ) from erro_local
        df.attrs["fonte_oficial_indisponivel"] = True
        df.attrs["erro_fonte_oficial"] = repr(erro_ipeadata)
        return df, "local"


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
        "sercodigo_ipeadata": ICTI_SERCODIGO_IPEADATA,
        "serie": ICTI_SERCODIGO,
        "fonte": "ipeadata",
    }


def obter_ultima_competencia_icti_atual(caminho=ICTI_CSV_PADRAO, timeout=20):
    """Ultima competencia da mesma serie vigente que alimenta o calculo."""
    df, fonte = carregar_icti_atual(caminho, timeout=timeout)
    ultima = df.iloc[-1]
    data = pd.Timestamp(ultima["data"])
    return {
        "data": data,
        "mes_ano": ultima["mes_ano"],
        "descricao": f"{MESES_PT_EXTENSO[data.month]}/{data.year}",
        "taxa_mensal_percentual": float(ultima["taxa_mensal_percentual"]),
        "sercodigo": ICTI_SERCODIGO,
        "sercodigo_ipeadata": ICTI_SERCODIGO_IPEADATA,
        "serie": ICTI_SERCODIGO,
        "fonte": fonte,
        "fonte_original": "Ipeadata/Ipea",
    }


def calcular_icti_ipeadata(
    data_inicio,
    data_fim=None,
    timeout=20,
    *,
    caminho=ICTI_CSV_PADRAO,
    _diagnostico=False,
):
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

    df, fonte = carregar_icti_atual(caminho, timeout=timeout)
    datas = set(df["data"])

    esperadas = competencias_mensais(competencia_proposta, competencia_final)
    encontradas = _competencias_do_dataframe(df)
    faltantes = [c for c in esperadas if c not in set(encontradas)]
    if competencia_base not in datas:
        faltantes = [competencia_base.strftime("%m/%Y"), *faltantes]

    if faltantes:
        if _diagnostico:
            if fonte == "local" and df.attrs.get("fonte_oficial_indisponivel"):
                ultima_local = pd.Timestamp(df["data"].max()).strftime("%m/%Y")
                raise FonteIndiceIndisponivel(
                    "Ipeadata",
                    {
                        "erro_ipeadata": df.attrs.get("erro_fonte_oficial"),
                        "fallback_local_faltantes": faltantes,
                    },
                    contexto={
                        "fallback_local_insuficiente": True,
                        "ultima_competencia_local": ultima_local,
                        "periodo_necessario": (
                            f"{competencia_proposta.strftime('%m/%Y')} a "
                            f"{competencia_final.strftime('%m/%Y')}"
                        ),
                    },
                )
            raise CompetenciasIndiceAusentes(faltantes, encontradas)
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
        "fonte": fonte,
        "fonte_original": "Ipeadata/Ipea",
        "fonte_oficial_indisponivel": bool(df.attrs.get("fonte_oficial_indisponivel")),
        "ultima_competencia_local": (
            pd.Timestamp(df["data"].max()).strftime("%m/%Y") if fonte == "local" else None
        ),
        "sercodigo": ICTI_SERCODIGO,
        "sercodigo_ipeadata": ICTI_SERCODIGO_IPEADATA,
        "serie": ICTI_SERCODIGO,
    }


def consultar_ist_com_diagnostico(data_inicio, caminho="ist.csv"):
    marco_final = pd.Timestamp(data_inicio) + relativedelta(years=1)
    esperadas = competencias_mensais(data_inicio, marco_final)
    return _executar_consulta_diagnosticada(
        lambda: calcular_ist_numero_indice(data_inicio, caminho, _diagnostico=True),
        fonte="Anatel",
        esperadas=esperadas,
    )


def consultar_sgs_com_diagnostico(serie_codigo, data_inicio, data_fim, timeout=15):
    esperadas = competencias_mensais(data_inicio, data_fim)
    return _executar_consulta_diagnosticada(
        lambda: coletar_sgs_produtorio(
            serie_codigo, data_inicio, data_fim, timeout=timeout, _diagnostico=True
        ),
        fonte="SGS/BCB",
        esperadas=esperadas,
    )


def consultar_icti_com_diagnostico(
    data_inicio, data_fim=None, timeout=20, *, caminho=ICTI_CSV_PADRAO
):
    final = pd.Timestamp(data_inicio) + relativedelta(months=11) if data_fim is None else data_fim
    esperadas = competencias_mensais(data_inicio, final)
    return _executar_consulta_diagnosticada(
        lambda: calcular_icti_ipeadata(
            data_inicio,
            data_fim,
            timeout=timeout,
            caminho=caminho,
            _diagnostico=True,
        ),
        fonte="Ipeadata",
        esperadas=esperadas,
    )
