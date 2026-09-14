from datetime import date
from decimal import Decimal

import pytest

from tools import atualizar_icti_ipeadata as atualizador


def _registro(ano, mes, taxa):
    return atualizador.RegistroICTI(date(ano, mes, 1), Decimal(taxa))


def _gravar(caminho, registros):
    atualizador.escrever_atomico(caminho, registros)
    return caminho.read_bytes()


def test_resposta_vazia_ou_invalida_e_rejeitada():
    for payload in ({}, {"value": []}, {"value": "invalido"}):
        with pytest.raises(atualizador.ErroAtualizacaoICTI):
            atualizador.extrair_registros_ipeadata(payload)


def test_competencia_nova_valida_mantem_historico(tmp_path, monkeypatch):
    caminho = tmp_path / "icti.csv"
    locais = [_registro(2026, 5, "0.25"), _registro(2026, 6, "0.63")]
    oficiais = [*locais, _registro(2026, 7, "-0.01")]
    _gravar(caminho, locais)
    monkeypatch.setattr(atualizador, "baixar_registros_icti", lambda: oficiais)

    novos = atualizador.executar(caminho)

    assert novos == oficiais[-1:]
    assert atualizador.ler_registros_locais(caminho) == oficiais
    assert caminho.read_text(encoding="utf-8").splitlines()[-1] == "2026-07;-0,01"


def test_serie_mais_curta_aborta_e_preserva_bytes(tmp_path, monkeypatch):
    caminho = tmp_path / "icti.csv"
    locais = [_registro(2026, 5, "0.25"), _registro(2026, 6, "0.63")]
    original = _gravar(caminho, locais)
    monkeypatch.setattr(atualizador, "baixar_registros_icti", lambda: locais[:-1])

    with pytest.raises(atualizador.ErroAtualizacaoICTI):
        atualizador.executar(caminho)

    assert caminho.read_bytes() == original


def test_api_indisponivel_nao_modifica_arquivo(tmp_path, monkeypatch):
    caminho = tmp_path / "icti.csv"
    original = _gravar(caminho, [_registro(2026, 6, "0.63")])
    monkeypatch.setattr(
        atualizador,
        "baixar_registros_icti",
        lambda: (_ for _ in ()).throw(atualizador.ErroAtualizacaoICTI("offline")),
    )
    with pytest.raises(atualizador.ErroAtualizacaoICTI, match="offline"):
        atualizador.executar(caminho)
    assert caminho.read_bytes() == original


def test_duplicidade_aborta():
    payload = {
        "value": [
            {"SERCODIGO": atualizador.ICTI_SERCODIGO, "VALDATA": "2026-06-01", "VALVALOR": 0.63},
            {"SERCODIGO": atualizador.ICTI_SERCODIGO, "VALDATA": "2026-06-20", "VALVALOR": 0.63},
        ]
    }
    with pytest.raises(atualizador.ErroAtualizacaoICTI, match="duplicada"):
        atualizador.extrair_registros_ipeadata(payload)


@pytest.mark.parametrize("valor", [None, True, "abc", "NaN", "Infinity"])
def test_valor_invalido_aborta(valor):
    payload = {
        "value": [
            {"SERCODIGO": atualizador.ICTI_SERCODIGO, "VALDATA": "2026-06-01", "VALVALOR": valor}
        ]
    }
    with pytest.raises(atualizador.ErroAtualizacaoICTI):
        atualizador.extrair_registros_ipeadata(payload)


def test_taxas_negativa_e_zero_sao_preservadas():
    payload = {
        "value": [
            {"SERCODIGO": atualizador.ICTI_SERCODIGO, "VALDATA": "2026-06-01", "VALVALOR": 0},
            {"SERCODIGO": atualizador.ICTI_SERCODIGO, "VALDATA": "2026-07-01", "VALVALOR": -0.01},
        ]
    }
    registros = atualizador.extrair_registros_ipeadata(payload)
    assert [r.taxa_mensal_percentual for r in registros] == [Decimal("0"), Decimal("-0.01")]


def test_divergencia_historica_aborta_e_preserva_bytes(tmp_path, monkeypatch):
    caminho = tmp_path / "icti.csv"
    locais = [_registro(2026, 5, "0.25"), _registro(2026, 6, "0.63")]
    original = _gravar(caminho, locais)
    oficiais = [_registro(2026, 5, "9.99"), _registro(2026, 6, "0.63")]
    monkeypatch.setattr(atualizador, "baixar_registros_icti", lambda: oficiais)

    with pytest.raises(atualizador.ErroAtualizacaoICTI, match="divergência histórica"):
        atualizador.executar(caminho)

    assert caminho.read_bytes() == original
