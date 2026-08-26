# -*- coding: utf-8 -*-
"""PRECLUSAO-UX1: explicacao visual sem nova regra temporal."""

from datetime import date
from pathlib import Path

import pytest

from _reajuste_utils import (
    classificar_pedido_por_data_exata,
    dias_de_preclusao,
)
from _ui_preclusao import html_indicador_preclusao


ROOT = Path(__file__).resolve().parents[1]
MULTIPLOS = ROOT / "pages" / "02_Calculo_Represados.py"


def _indicadores(at):
    return [
        str(md.value)
        for md in at.markdown
        if 'class="cl8us-preclusao"' in str(md.value)
    ]


def test_precluso_usa_exatamente_a_fronteira_do_classificador():
    referencia = date(2025, 8, 23)
    limite = date(2025, 11, 21)
    pedido = date(2025, 12, 9)
    situacao = classificar_pedido_por_data_exata(pedido, referencia, limite)

    assert situacao == "PRECLUSO"
    assert dias_de_preclusao(situacao, pedido, limite) == 18


@pytest.mark.parametrize(
    ("situacao", "pedido"),
    (
        ("ADIANTADO", date(2025, 8, 22)),
        ("TEMPESTIVO", date(2025, 8, 23)),
        ("TEMPESTIVO*", date(2025, 9, 1)),
        ("OUTRA", date(2025, 12, 9)),
    ),
)
def test_situacao_diferente_de_precluso_nao_produz_dias(situacao, pedido):
    assert dias_de_preclusao(situacao, pedido, date(2025, 11, 21)) is None


def test_sem_pedido_nao_inventa_dias_de_preclusao():
    assert dias_de_preclusao("PRECLUSO", None, date(2025, 11, 21)) is None


def test_html_e_compacto_sem_datas_e_nao_depende_so_de_cor():
    html = html_indicador_preclusao(18, ciclo="C2")

    assert "C2 · 18 dias de preclusão" in html
    assert "role=\"note\"" in html
    assert "aria-label=\"C2 · 18 dias de preclusão\"" in html
    assert "cl8us-preclusao-regular" in html
    assert "cl8us-preclusao-atraso" in html
    assert "grid-template-columns: minmax(5.5rem, 3fr) 1.5rem" in html
    assert "</style><div" in html
    assert "</style>\n<div" not in html
    assert "2025" not in html and "2026" not in html


def test_plural_e_singular_naturais():
    assert "1 dia de preclusão" in html_indicador_preclusao(1)
    assert "2 dias de preclusão" in html_indicador_preclusao(2)


def test_fluxo_real_multiciclo_mostra_somente_preclusos_e_dias_exatos():
    from streamlit.testing.v1 import AppTest

    at = AppTest.from_file(str(MULTIPLOS), default_timeout=180)
    at.run()
    assert not at.exception

    # C1 apto em 10/10/2023; limite canonico em 08/01/2024.
    c1 = at.date_input(key="p1_20231010")

    c1.set_value(date(2023, 10, 9)).run()  # ADIANTADO
    assert not _indicadores(at)

    at.date_input(key="p1_20231010").set_value(date(2023, 10, 10)).run()
    assert not _indicadores(at)  # TEMPESTIVO

    at.date_input(key="p1_20231010").set_value(date(2023, 12, 10)).run()
    assert not _indicadores(at)  # TEMPESTIVO* (apresentacao)

    at.date_input(key="p1_20231010").set_value(date(2024, 1, 26)).run()
    indicadores = _indicadores(at)
    assert len(indicadores) == 1
    assert "C1 · 18 dias de preclusão" in indicadores[0]

    # A preclusao nao move a fronteira de C2: limite 08/01/2025.
    at.date_input(key="p2_20241010").set_value(date(2025, 1, 15)).run()
    indicadores = _indicadores(at)
    assert len(indicadores) == 2
    assert "C1 · 18 dias de preclusão" in indicadores[0]
    assert "C2 · 7 dias de preclusão" in indicadores[1]


def test_fluxo_real_ciclo_unico_mostra_dias_sem_rotulo_redundante():
    from streamlit.testing.v1 import AppTest

    pagina = ROOT / "pages" / "01_Calculo_Simples.py"
    at = AppTest.from_file(str(pagina), default_timeout=180)
    at.run()
    assert not at.exception

    # Apto em 02/08/2024; limite canonico em 31/10/2024.
    at.date_input[0].set_value(date(2023, 8, 2))
    at.date_input[1].set_value(date(2024, 11, 18)).run()
    at.button[0].click().run()

    assert not at.exception
    indicadores = _indicadores(at)
    assert len(indicadores) == 1
    assert "18 dias de preclusão" in indicadores[0]
    assert "C1 ·" not in indicadores[0]


def test_integracao_fica_somente_na_apresentacao_das_calculadoras():
    simples = (ROOT / "pages" / "01_Calculo_Simples.py").read_text(encoding="utf-8")
    multiplos = MULTIPLOS.read_text(encoding="utf-8")

    assert "render_indicador_preclusao(\n        situacao_pedido," in simples
    assert "render_indicador_preclusao(\n        situacao_limpa," in multiplos
    assert "dt_limite" in simples
    assert "d_lim" in multiplos
    assert "gerar_coleta_oficial_preenchida" not in (
        ROOT / "_ui_preclusao.py"
    ).read_text(encoding="utf-8")
