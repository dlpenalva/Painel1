from datetime import date

import pandas as pd
import pytest
import requests

import _indice_utils as iu


ANCORA = date(2025, 7, 25)
FINAL = date(2026, 6, 25)


def _serie(*, remover=()):
    datas = pd.date_range("2025-06-01", "2026-06-01", freq="MS")
    taxas = [1.01, 0.13, 0.22, 0.36, -0.21, 0.57, 0.22, 0.15, 0.26, 0.46, 0.95, 0.25, 0.63]
    bruto = pd.DataFrame({"data": datas, "taxa_mensal_percentual": taxas})
    bruto = bruto[~bruto["data"].dt.strftime("%m/%Y").isin(remover)].copy()
    return iu._finalizar_serie_icti(bruto, origem="fixture")


def _csv(caminho, serie):
    linhas = ["COMPETENCIA;TAXA_MENSAL_PERCENTUAL"]
    linhas.extend(
        f"{data:%Y-%m};{str(taxa).replace('.', ',')}"
        for data, taxa in zip(serie["data"], serie["taxa_mensal_percentual"])
    )
    caminho.write_text("\n".join(linhas) + "\n", encoding="utf-8")
    return caminho


def test_a_ipeadata_disponivel_usa_online_estado_ok(monkeypatch, tmp_path):
    online = _serie()
    monkeypatch.setattr(iu, "carregar_icti_ipeadata", lambda **kwargs: online)
    consulta = iu.consultar_icti_com_diagnostico(
        ANCORA, FINAL, caminho=tmp_path / "nao_deve_ser_usado.csv"
    )
    assert consulta["diagnostico"]["estado"] == iu.ESTADO_INDICE_OK
    assert consulta["resultado"]["fonte"] == "ipeadata"
    assert consulta["resultado"]["variacao"] == pytest.approx(
        (1 + online.iloc[1:]["taxa_mensal_percentual"] / 100).prod() - 1
    )


def test_b_ipeadata_indisponivel_local_completo_e_identico(monkeypatch, tmp_path):
    serie = _serie()
    caminho = _csv(tmp_path / "icti.csv", serie)
    monkeypatch.setattr(iu, "carregar_icti_ipeadata", lambda **kwargs: serie)
    online = iu.consultar_icti_com_diagnostico(ANCORA, FINAL, caminho=caminho)
    monkeypatch.setattr(
        iu,
        "carregar_icti_ipeadata",
        lambda **kwargs: (_ for _ in ()).throw(requests.Timeout("offline")),
    )
    local = iu.consultar_icti_com_diagnostico(ANCORA, FINAL, caminho=caminho)

    assert local["diagnostico"]["estado"] == iu.ESTADO_FALLBACK_LOCAL
    assert local["resultado"]["fonte"] == "local"
    assert local["resultado"]["variacao"] == online["resultado"]["variacao"]
    pd.testing.assert_frame_equal(local["resultado"]["dados"], online["resultado"]["dados"])


def test_espelho_local_preserva_decimal_longo_valor_a_valor(monkeypatch, tmp_path):
    bruto = pd.DataFrame(
        {
            "data": pd.date_range("2014-04-01", periods=2, freq="MS"),
            "taxa_mensal_percentual": [0.00773474325808099, -0.00843007668022144],
        }
    )
    online = iu._finalizar_serie_icti(bruto, origem="fixture")
    local = iu.carregar_icti_local(_csv(tmp_path / "icti.csv", online))
    assert local["taxa_mensal_percentual"].tolist() == online[
        "taxa_mensal_percentual"
    ].tolist()


def test_c_ipeadata_indisponivel_local_insuficiente_bloqueia(monkeypatch, tmp_path):
    caminho = _csv(tmp_path / "icti.csv", _serie(remover={"06/2026"}))
    monkeypatch.setattr(
        iu,
        "carregar_icti_ipeadata",
        lambda **kwargs: (_ for _ in ()).throw(requests.ConnectionError("offline")),
    )
    consulta = iu.consultar_icti_com_diagnostico(ANCORA, FINAL, caminho=caminho)
    assert consulta["resultado"] is None
    assert consulta["diagnostico"]["estado"] == iu.ESTADO_FONTE_INDISPONIVEL
    assert consulta["diagnostico"]["fallback_local_insuficiente"] is True
    assert consulta["diagnostico"]["ultima_competencia_local"] == "05/2026"
    assert consulta["diagnostico"]["faltantes"] == []


def test_ui_local_insuficiente_informa_limite_e_periodo(monkeypatch):
    import _ui_utils as ui

    mensagens = []
    monkeypatch.setattr(ui.st, "error", mensagens.append)
    ui.render_diagnostico_indice(
        {
            "estado": iu.ESTADO_FONTE_INDISPONIVEL,
            "fonte": "Ipeadata",
            "fallback_local_insuficiente": True,
            "ultima_competencia_local": "05/2026",
            "periodo_necessario": "07/2025 a 06/2026",
        }
    )
    assert "cópia local do ICTI" in mensagens[0]
    assert "05/2026" in mensagens[0]
    assert "07/2025 a 06/2026" in mensagens[0]


def test_d_online_valido_sem_competencia_nao_usa_local(monkeypatch, tmp_path):
    online = _serie(remover={"03/2026"})
    caminho = _csv(tmp_path / "icti.csv", _serie())
    monkeypatch.setattr(iu, "carregar_icti_ipeadata", lambda **kwargs: online)
    consulta = iu.consultar_icti_com_diagnostico(ANCORA, FINAL, caminho=caminho)
    assert consulta["resultado"] is None
    assert consulta["diagnostico"]["estado"] == iu.ESTADO_COMPETENCIAS_AUSENTES
    assert consulta["diagnostico"]["faltantes"] == ["03/2026"]


def test_l_caso_real_nao_pede_julho_ou_agosto_2026(monkeypatch, tmp_path):
    serie = _serie()
    caminho = _csv(tmp_path / "icti.csv", serie)
    monkeypatch.setattr(iu, "carregar_icti_ipeadata", lambda **kwargs: serie)
    consulta = iu.consultar_icti_com_diagnostico(ANCORA, FINAL, caminho=caminho)
    resultado = consulta["resultado"]
    assert consulta["diagnostico"]["esperadas"] == [
        "07/2025", "08/2025", "09/2025", "10/2025", "11/2025", "12/2025",
        "01/2026", "02/2026", "03/2026", "04/2026", "05/2026", "06/2026",
    ]
    assert resultado["competencia_indice_base"] == pd.Timestamp("2025-06-01")
    assert resultado["competencia_final"] == pd.Timestamp("2026-06-01")


def test_m_outros_indices_permanecem_nos_mesmos_dispatches():
    assert iu.serie_sgs_do_indice("IPCA") == str(iu.SGS_IPCA)
    assert iu.serie_sgs_do_indice("IGP-M") == str(iu.SGS_IGPM)
    assert iu.serie_sgs_do_indice("INPC") == str(iu.SGS_INPC)
    assert callable(iu.calcular_ist_numero_indice)


def test_n_paginas_usam_o_mesmo_nucleo_sem_calculo_icti_paralelo():
    raiz = __import__("pathlib").Path(__file__).resolve().parents[1]
    for relativo in ("pages/01_Calculo_Simples.py", "pages/02_Calculo_Represados.py"):
        codigo = (raiz / relativo).read_text(encoding="utf-8")
        assert "consultar_icti_com_diagnostico" in codigo
        assert "taxa_mensal_percentual" not in codigo
