"""Contagem de ciclos efetivamente considerados na apuracao.

A frase do documento nao pode contar linhas do quadro: um ciclo listado como
"Fora da apuracao" permanece no Quadro 1 para rastreabilidade, mas afirmar que
ele foi considerado contradiz o proprio quadro.

Cobre a regra compartilhada (_reajuste_utils) e a redacao resultante no
Despacho Saneador, incluindo a concordancia entre a frase e o Quadro 1.
"""

from __future__ import annotations

import io
from datetime import date

import docx
import pytest

from _reajuste_utils import (
    FRASE_SEM_CICLOS_COMPUTADOS,
    ciclo_computado,
    contar_ciclos_computados,
    expressao_quantidade_ciclos,
)
from _templates_documentos import gerar_despacho_saneador
from test_sumario_executivo import leitura_multiciclo_pc


CICLOS_REAJUSTE = ("C1", "C2", "C3", "C4")


# ---------------------------------------------------------------------------
# Regra compartilhada
# ---------------------------------------------------------------------------

def _por_computar(*marcas):
    return [{"ciclo": f"C{i + 1}", "computar": m} for i, m in enumerate(marcas)]


@pytest.mark.parametrize("marcas,esperado", [
    (("Sim", "Não", "Não", "Não"), 1),
    (("Sim", "Sim", "Não", "Não"), 2),
    (("Sim", "Sim", "Sim", "Sim"), 4),
    (("Não", "Não", "Não", "Não"), 0),
    ((None, None, None, None), 0),
    (("", "", "", ""), 0),
])
def test_contagem_por_marcacao_computar(marcas, esperado):
    assert contar_ciclos_computados(_por_computar(*marcas)) == esperado


def test_marcacao_computar_aceita_variacoes_de_grafia():
    assert ciclo_computado({"computar": "sim"})
    assert ciclo_computado({"computar_nesta_apuracao": "SIM"})
    assert not ciclo_computado({"computar": "Nao"})
    assert not ciclo_computado({"computar": "Não"})


def test_tratamento_financeiro_quando_nao_ha_marcacao_computar():
    assert ciclo_computado({"Tratamento financeiro do ciclo": "Apurar"})
    assert not ciclo_computado({"Tratamento financeiro do ciclo": "Fora da apuração"})


def test_classificacao_exclui_apenas_quando_explicita():
    assert not ciclo_computado({"Situação": "Fora da apuracao"})
    assert not ciclo_computado({"Classificação": "não computado"})
    assert not ciclo_computado({"Situação": "Não aplicável"})
    assert ciclo_computado({"Situação": "TEMPESTIVO"})


def test_marcacao_computar_tem_precedencia_sobre_a_classificacao():
    """COMPUTAR e a fonte canonica; a situacao exibida nao a sobrepoe."""
    assert not ciclo_computado({"computar": "Não", "Situação": "TEMPESTIVO"})


@pytest.mark.parametrize("quantidade,esperado", [
    (1, "1 ciclo"),
    (2, "2 ciclos"),
    (4, "4 ciclos"),
    (0, None),
    (None, None),
])
def test_expressao_com_flexao_correta(quantidade, esperado):
    assert expressao_quantidade_ciclos(quantidade) == esperado


def test_expressao_nunca_usa_a_forma_ciclo_s():
    for quantidade in range(1, 6):
        assert "(s)" not in expressao_quantidade_ciclos(quantidade)


# ---------------------------------------------------------------------------
# Redacao no Despacho Saneador
# ---------------------------------------------------------------------------

def _leitura_com_computados(computados: set[str]) -> dict:
    leitura = leitura_multiciclo_pc()
    parametros = leitura.get("parametros_v10") or {}
    for registro in parametros.get("ciclos") or []:
        nome = str(registro.get("ciclo") or "").upper()
        if nome in CICLOS_REAJUSTE:
            registro["computar_nesta_apuracao"] = "Sim" if nome in computados else "Nao"
    for nome, registro in (parametros.get("por_ciclo") or {}).items():
        if str(nome).upper() in CICLOS_REAJUSTE:
            registro["computar_nesta_apuracao"] = (
                "Sim" if str(nome).upper() in computados else "Nao"
            )
    return leitura


def _documento(leitura: dict):
    return docx.Document(io.BytesIO(gerar_despacho_saneador(leitura)))


def _paragrafo_4(documento) -> str:
    for paragrafo in documento.paragraphs:
        if paragrafo.text.strip().startswith("4."):
            return paragrafo.text
    raise AssertionError("item 4 ausente no Despacho Saneador")


def _quadro_1(documento) -> list[list[str]]:
    tabela = documento.tables[0]
    return [[c.text.strip() for c in linha.cells] for linha in tabela.rows[1:]]


@pytest.mark.parametrize("computados,esperado", [
    ({"C1"}, "considerou 1 ciclo,"),
    ({"C1", "C2"}, "considerou 2 ciclos,"),
    ({"C1", "C2", "C3", "C4"}, "considerou 4 ciclos,"),
])
def test_frase_declara_apenas_os_ciclos_computados(computados, esperado):
    texto = _paragrafo_4(_documento(_leitura_com_computados(computados)))
    assert esperado in texto
    assert "ciclo(s)" not in texto


def test_sem_ciclo_computado_nao_afirma_que_houve_analise_de_ciclos():
    texto = _paragrafo_4(_documento(_leitura_com_computados(set())))
    assert FRASE_SEM_CICLOS_COMPUTADOS in texto
    assert "A análise de reajuste considerou" not in texto
    assert "ciclo(s)" not in texto
    # O valor original continua sendo declarado no mesmo item.
    assert "valor original do contrato" in texto


def test_ciclos_existentes_marcados_como_nao_nao_sao_contados():
    """Os quatro ciclos seguem no Quadro 1, nenhum e declarado como considerado."""
    documento = _documento(_leitura_com_computados(set()))
    assert len(_quadro_1(documento)) == len(CICLOS_REAJUSTE)
    assert FRASE_SEM_CICLOS_COMPUTADOS in _paragrafo_4(documento)


@pytest.mark.parametrize("computados", [
    {"C1"},
    {"C1", "C2"},
    {"C1", "C2", "C3", "C4"},
])
def test_concordancia_entre_a_frase_e_o_quadro_1(computados):
    """O numero da frase nao pode exceder as linhas do quadro nem ignora-las.

    O quadro continua listando todos os ciclos de reajuste; a frase declara
    somente os computados.
    """
    documento = _documento(_leitura_com_computados(computados))
    linhas = _quadro_1(documento)
    assert len(linhas) == len(CICLOS_REAJUSTE)
    esperado = expressao_quantidade_ciclos(len(computados))
    assert esperado in _paragrafo_4(documento)
    assert len(computados) <= len(linhas)


# ---------------------------------------------------------------------------
# Cenario real relatado: C1 computado, C2-C4 fora da apuracao
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def documento_cenario_real():
    from _coleta_oficial import gerar_coleta_oficial_preenchida
    from _coleta_reajuste_documentos import processar_coleta_oficial_runtime

    coleta = gerar_coleta_oficial_preenchida({
        "origem": "Cenario real C1 computado",
        "indice": "IST (Anatel)",
        "data_base_original": "01/02/2025",
        "data_corte": date(2027, 1, 31),
        "ciclos": [{
            "ciclo": "C1",
            "data_inicio": date(2026, 2, 1),
            "data_fim": date(2027, 1, 31),
            "data_pedido": date(2026, 2, 1),
            "financeiro_inicio": date(2026, 4, 1),
            "percentual_aplicado": 0.0307853139440224,
            "situacao": "✅ TEMPESTIVO",
            "objeto_analise_atual": True,
        }],
    })
    payload, _ = processar_coleta_oficial_runtime(coleta)
    return _documento(payload)


def test_cenario_real_declara_um_unico_ciclo(documento_cenario_real):
    texto = _paragrafo_4(documento_cenario_real)
    assert "A análise de reajuste considerou 1 ciclo, com variação acumulada de" in texto
    assert "ciclo(s)" not in texto
    assert "considerou 4" not in texto


def test_cenario_real_mantem_os_demais_ciclos_no_quadro(documento_cenario_real):
    linhas = _quadro_1(documento_cenario_real)
    ciclos = [linha[0] for linha in linhas]
    assert ciclos == list(CICLOS_REAJUSTE)
    fora = [linha for linha in linhas if "Fora da apuracao" in linha[5]]
    assert len(fora) == 3
