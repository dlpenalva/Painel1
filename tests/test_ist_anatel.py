"""§9 — IST automatico pela fonte oficial da Anatel, com fallback local.

Sem dependencia de internet: a serie oficial e injetada por fixture/mocks.
A metodologia matematica (v_fim / v_ini) permanece inalterada.
"""
from __future__ import annotations

import datetime
from decimal import Decimal

import pandas as pd
import pytest
import requests

import _indice_utils
from tools.atualizar_ist_anatel import extrair_registros_ist


HTML_ANATEL = """
<html><body>
<table><tbody>
<tr><td><strong>Referência</strong></td><td>Variação</td><td>IST</td></tr>
<tr><td><div>Abr/26</div></td><td><p>1,34%</p></td><td><p>358,475</p></td></tr>
<tr><td><p>Mai/26</p></td><td>0,85%</td><td><div>361,530</div></td></tr>
<tr><td><p>Jun/26</p></td><td>0,10%</td><td><div>361,891</div></td></tr>
</tbody></table>
</body></html>
"""


def _df(pares):
    """pares: [(ano, mes, indice), ...] -> DataFrame [data, indice]."""
    return pd.DataFrame(
        [{"data": pd.Timestamp(a, m, 1), "indice": float(v)} for (a, m, v) in pares]
    )


def _serie_com_jun26():
    # 13 competencias jun/25..jun/26 (valores monotonos ficticios), jun/26 = 361,891
    pares = [(2025, 6 + i, 350.0 + i) for i in range(0, 7)]        # jun..dez/25
    pares += [(2026, i, 357.0 + i) for i in range(1, 6)]           # jan..mai/26
    pares += [(2026, 6, 361.891)]                                  # jun/26
    return _df(pares)


def _patch_anatel(monkeypatch, retorno=None, erro=None):
    def fake(*a, **k):
        if erro is not None:
            raise erro
        return retorno
    monkeypatch.setattr(_indice_utils, "carregar_ist_anatel", fake, raising=False)
    _indice_utils._resetar_cache_ist()


# A. HTML Anatel -> ultima competencia junho/2026, indice 361,891 -----------

def test_a_parser_html_ultima_competencia_e_indice():
    registros = extrair_registros_ist(HTML_ANATEL)
    assert registros[-1].competencia == datetime.date(2026, 6, 1)
    assert registros[-1].indice == Decimal("361.891")
    por_comp = {r.competencia: r.indice for r in registros}
    assert por_comp[datetime.date(2026, 5, 1)] == Decimal("361.530")


# B/C. serie online alimenta o calculo, inclusive competencia so na Anatel ----

def test_bc_calculo_usa_serie_online_com_jun26(monkeypatch):
    _patch_anatel(monkeypatch, retorno=_serie_com_jun26())
    df, fonte = _indice_utils.carregar_ist_atual()
    assert fonte == "anatel"
    assert float(df.iloc[-1]["indice"]) == pytest.approx(361.891)

    res = _indice_utils.calcular_ist_numero_indice(datetime.date(2025, 6, 1))
    assert res is not None
    assert res["fonte"] == "anatel"
    assert res["i_fim"] == pytest.approx(361.891)          # jun/26 online
    assert "Anatel" in res["metodo"]


def test_c_competencia_apenas_na_anatel_nao_existe_no_csv(monkeypatch):
    base_local = _indice_utils.carregar_ist_local("ist.csv")
    tem_jun26_local = ((base_local["data"].dt.year == 2026) & (base_local["data"].dt.month == 6)).any()
    assert not tem_jun26_local
    _patch_anatel(monkeypatch, retorno=_serie_com_jun26())
    res = _indice_utils.calcular_ist_numero_indice(datetime.date(2025, 6, 1))
    assert res is not None and res["i_fim"] == pytest.approx(361.891)


# D/E/F. falhas de rede -> fallback local ------------------------------------

@pytest.mark.parametrize("erro", [
    requests.Timeout("timeout"),
    requests.HTTPError("500"),
    ValueError("HTML inesperado / nenhuma tabela valida"),
])
def test_def_fallback_local_em_falha(monkeypatch, erro):
    _patch_anatel(monkeypatch, erro=erro)
    df, fonte = _indice_utils.carregar_ist_atual("ist.csv")
    assert fonte == "local"
    esperado = _indice_utils.carregar_ist_local("ist.csv")
    assert list(df["indice"]) == list(esperado["indice"])


def test_f_fallback_permite_calcular_quando_csv_cobre(monkeypatch):
    _patch_anatel(monkeypatch, erro=requests.Timeout("timeout"))
    res = _indice_utils.calcular_ist_numero_indice(datetime.date(2023, 8, 1))
    assert res is not None
    assert res["fonte"] == "local"
    assert "base local" in res["metodo"]


# G. fallback sem a competencia necessaria -> None (nunca extrapola) ---------

def test_g_fallback_sem_competencia_retorna_none(monkeypatch):
    _patch_anatel(monkeypatch, erro=requests.Timeout("timeout"))
    res = _indice_utils.calcular_ist_numero_indice(datetime.date(2030, 1, 1))
    assert res is None


# H. metodologia (divisao de numero-indice) preservada -----------------------

def test_h_metodologia_divisao_preservada(monkeypatch):
    _patch_anatel(monkeypatch, retorno=_serie_com_jun26())
    res = _indice_utils.calcular_ist_numero_indice(datetime.date(2025, 6, 1))
    assert res["variacao"] == pytest.approx(res["i_fim"] / res["i_ini"] - 1, rel=1e-12)


# I. memoria usa competencias reais, sem interpolacao ------------------------

def test_i_memoria_sem_interpolacao(monkeypatch):
    serie = _serie_com_jun26()
    _patch_anatel(monkeypatch, retorno=serie)
    res = _indice_utils.calcular_ist_numero_indice(datetime.date(2025, 6, 1))
    datas_mem = set(res["dados"]["data"].dt.strftime("%Y-%m"))
    datas_serie = set(serie["data"].dt.strftime("%Y-%m"))
    assert datas_mem.issubset(datas_serie)
    assert res["dados"]["indice"].notna().all()


# J. IPCA, IGP-M e ICTI intactos (nenhuma dependencia do IST) ----------------

def test_j_outros_indices_intactos():
    assert hasattr(_indice_utils, "coletar_sgs_produtorio")     # IPCA(433)/IGP-M(189)
    assert hasattr(_indice_utils, "calcular_icti_ipeadata")     # ICTI/Ipeadata
    doc = _indice_utils.carregar_ist_atual.__doc__ or ""
    assert "SGS" not in doc and "ICTI" not in doc


def test_cache_ttl_reusa_e_expira(monkeypatch):
    chamadas = {"n": 0}

    def fake(*a, **k):
        chamadas["n"] += 1
        return _serie_com_jun26()

    monkeypatch.setattr(_indice_utils, "carregar_ist_anatel", fake, raising=False)
    _indice_utils._resetar_cache_ist()

    _indice_utils.carregar_ist_atual(ttl=100, _agora=0.0)      # consulta
    _indice_utils.carregar_ist_atual(ttl=100, _agora=50.0)     # dentro do TTL -> cache
    assert chamadas["n"] == 1
    _indice_utils.carregar_ist_atual(ttl=100, _agora=200.0)    # apos TTL -> nova consulta
    assert chamadas["n"] == 2
