from datetime import date
from pathlib import Path

import pandas as pd
import pytest
import requests

import _indice_utils as iu
import _ui_utils as ui


INICIO = date(2025, 7, 25)
FIM = date(2026, 6, 25)


def _serie_icti(*, remover=()):
    datas = pd.date_range("2025-06-01", "2026-06-01", freq="MS")
    df = pd.DataFrame(
        {
            "data": datas,
            "taxa_mensal_percentual": [0.5] * len(datas),
        }
    )
    df = df[~df["data"].dt.strftime("%m/%Y").isin(remover)].copy()
    df["fator_mensal"] = 1 + df["taxa_mensal_percentual"] / 100
    df["indice_nivel_sintetico"] = 100 * df["fator_mensal"].cumprod()
    return df.reset_index(drop=True)


def _payload_sgs(*, remover=()):
    return [
        {"data": data.strftime("%d/%m/%Y"), "valor": "0.50"}
        for data in pd.date_range("2025-07-01", "2026-06-01", freq="MS")
        if data.strftime("%m/%Y") not in remover
    ]


class _RespostaSGS:
    def __init__(self, payload):
        self.payload = payload

    def raise_for_status(self):
        return None

    def json(self):
        return self.payload


@pytest.fixture(autouse=True)
def _limpar_cache_ist():
    iu._resetar_cache_ist()
    yield
    iu._resetar_cache_ist()


def test_icti_fonte_indisponivel_nao_inventa_doze_competencias(monkeypatch):
    monkeypatch.setattr(
        iu,
        "carregar_icti_ipeadata",
        lambda **kwargs: (_ for _ in ()).throw(requests.ConnectionError("offline")),
    )
    monkeypatch.setattr(
        iu,
        "carregar_icti_local",
        lambda caminho: (_ for _ in ()).throw(FileNotFoundError("sem cópia local")),
    )

    consulta = iu.consultar_icti_com_diagnostico(INICIO, FIM)

    assert consulta["resultado"] is None
    assert consulta["diagnostico"]["estado"] == iu.ESTADO_FONTE_INDISPONIVEL
    assert consulta["diagnostico"]["fonte"] == "Ipeadata"
    assert consulta["diagnostico"]["faltantes"] == []
    assert consulta["diagnostico"]["esperadas"] == [
        "07/2025", "08/2025", "09/2025", "10/2025", "11/2025", "12/2025",
        "01/2026", "02/2026", "03/2026", "04/2026", "05/2026", "06/2026",
    ]


def test_icti_serie_valida_lista_somente_competencia_realmente_ausente(monkeypatch):
    monkeypatch.setattr(iu, "carregar_icti_ipeadata", lambda **kwargs: _serie_icti(remover={"06/2026"}))

    consulta = iu.consultar_icti_com_diagnostico(INICIO, FIM)

    assert consulta["diagnostico"]["estado"] == iu.ESTADO_COMPETENCIAS_AUSENTES
    assert consulta["diagnostico"]["faltantes"] == ["06/2026"]


def test_icti_completo_preserva_periodo_base_e_calculo(monkeypatch):
    monkeypatch.setattr(iu, "carregar_icti_ipeadata", lambda **kwargs: _serie_icti())

    anterior = iu.calcular_icti_ipeadata(INICIO, FIM)
    consulta = iu.consultar_icti_com_diagnostico(INICIO, FIM)
    atual = consulta["resultado"]

    assert consulta["diagnostico"]["estado"] == iu.ESTADO_INDICE_OK
    assert atual["variacao"] == anterior["variacao"] == pytest.approx((1.005**12) - 1)
    assert atual["competencia_proposta"] == pd.Timestamp("2025-07-01")
    assert atual["competencia_indice_base"] == pd.Timestamp("2025-06-01")
    assert atual["competencia_final"] == pd.Timestamp("2026-06-01")


def test_ipca_bcb_indisponivel(monkeypatch):
    monkeypatch.setattr(
        iu.requests,
        "get",
        lambda *args, **kwargs: (_ for _ in ()).throw(requests.Timeout("timeout")),
    )

    consulta = iu.consultar_sgs_com_diagnostico(iu.SGS_IPCA, INICIO, FIM)

    assert consulta["diagnostico"]["estado"] == iu.ESTADO_FONTE_INDISPONIVEL
    assert consulta["diagnostico"]["fonte"] == "SGS/BCB"
    assert consulta["diagnostico"]["faltantes"] == []


def test_sgs_sem_serie_utilizavel_nao_inventa_competencias(monkeypatch):
    monkeypatch.setattr(iu.requests, "get", lambda *args, **kwargs: _RespostaSGS([]))

    consulta = iu.consultar_sgs_com_diagnostico(iu.SGS_IPCA, INICIO, FIM)

    assert consulta["diagnostico"]["estado"] == iu.ESTADO_FONTE_INDISPONIVEL
    assert consulta["diagnostico"]["faltantes"] == []


def test_igpm_serie_valida_lista_somente_competencia_realmente_ausente(monkeypatch):
    monkeypatch.setattr(
        iu.requests, "get", lambda *args, **kwargs: _RespostaSGS(_payload_sgs(remover={"03/2026"}))
    )

    consulta = iu.consultar_sgs_com_diagnostico(iu.SGS_IGPM, INICIO, FIM)

    assert consulta["diagnostico"]["estado"] == iu.ESTADO_COMPETENCIAS_AUSENTES
    assert consulta["diagnostico"]["faltantes"] == ["03/2026"]


def test_inpc_completo_preserva_produtorio(monkeypatch):
    payload = _payload_sgs()
    monkeypatch.setattr(iu.requests, "get", lambda *args, **kwargs: _RespostaSGS(payload))

    anterior = iu.coletar_sgs_produtorio(iu.SGS_INPC, INICIO, FIM)
    consulta = iu.consultar_sgs_com_diagnostico(iu.SGS_INPC, INICIO, FIM)

    assert consulta["diagnostico"]["estado"] == iu.ESTADO_INDICE_OK
    assert consulta["resultado"]["variacao"] == anterior["variacao"]
    assert consulta["resultado"]["variacao"] == pytest.approx((1.005**12) - 1)


def _serie_ist_local(*, incluir_final=True):
    datas = list(pd.date_range("2025-07-01", "2026-07-01", freq="MS"))
    if not incluir_final:
        datas.pop()
    return pd.DataFrame({"data": datas, "indice": [100.0 + i for i in range(len(datas))]})


def test_ist_anatel_indisponivel_com_fallback_local_suficiente(monkeypatch):
    monkeypatch.setattr(
        iu,
        "carregar_ist_anatel",
        lambda **kwargs: (_ for _ in ()).throw(requests.Timeout("offline")),
    )
    monkeypatch.setattr(iu, "carregar_ist_local", lambda caminho: _serie_ist_local())

    consulta = iu.consultar_ist_com_diagnostico(INICIO)

    assert consulta["diagnostico"]["estado"] == iu.ESTADO_FALLBACK_LOCAL
    assert consulta["resultado"]["fonte"] == "local"
    assert consulta["resultado"]["variacao"] == pytest.approx(0.12)


def test_ist_anatel_indisponivel_e_fallback_insuficiente_bloqueia_sem_falsa_lista(monkeypatch):
    monkeypatch.setattr(
        iu,
        "carregar_ist_anatel",
        lambda **kwargs: (_ for _ in ()).throw(requests.Timeout("offline")),
    )
    monkeypatch.setattr(iu, "carregar_ist_local", lambda caminho: _serie_ist_local(incluir_final=False))

    consulta = iu.consultar_ist_com_diagnostico(INICIO)

    assert consulta["resultado"] is None
    assert consulta["diagnostico"]["estado"] == iu.ESTADO_FONTE_INDISPONIVEL
    assert consulta["diagnostico"]["faltantes"] == []


def test_falha_na_consulta_informativa_nao_decide_o_calculo_icti(monkeypatch):
    monkeypatch.setattr(iu, "carregar_icti_ipeadata", lambda **kwargs: _serie_icti())
    monkeypatch.setattr(
        iu,
        "obter_ultima_competencia_icti_ipeadata",
        lambda **kwargs: (_ for _ in ()).throw(requests.Timeout("apenas alerta")),
    )

    consulta = iu.consultar_icti_com_diagnostico(INICIO, FIM)

    assert consulta["diagnostico"]["estado"] == iu.ESTADO_INDICE_OK
    assert consulta["resultado"] is not None


def test_apresentacao_fonte_indisponivel_nao_exibe_competencias(monkeypatch):
    erros = []
    tabelas = []
    monkeypatch.setattr(ui.st, "error", erros.append)
    monkeypatch.setattr(ui.st, "dataframe", lambda *args, **kwargs: tabelas.append(args))
    diagnostico = iu._diagnostico_indice(
        iu.ESTADO_FONTE_INDISPONIVEL,
        "Ipeadata",
        iu.competencias_mensais(INICIO, FIM),
    )

    ui.render_diagnostico_indice(diagnostico)

    assert "Ipeadata" in erros[0]
    assert "Tente novamente" in erros[0]
    assert "Competências faltantes" not in erros[0]
    assert tabelas == []


def test_calculadoras_consumem_mesmo_diagnostico_compartilhado():
    raiz = Path(__file__).resolve().parents[1]
    simples = (raiz / "pages" / "01_Calculo_Simples.py").read_text(encoding="utf-8")
    represados = (raiz / "pages" / "02_Calculo_Represados.py").read_text(encoding="utf-8")

    for codigo in (simples, represados):
        assert "render_diagnostico_indice(validacao_indice" in codigo
        assert '"sem_retorno"' not in codigo
        assert "faltantes\": esperadas" not in codigo
