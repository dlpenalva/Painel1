# -*- coding: utf-8 -*-
"""Etapa 32 — régua temporal visual dos marcos (pagina Analisar varios ciclos).

Valida os cenarios A-F do enunciado: a regua e um componente EXCLUSIVAMENTE
visual que reutiliza dados ja calculados pela pagina (input_ciclos e contexto
lateral); nenhuma data pode ser inferida e nenhuma regra de negocio muda.
"""
from datetime import date
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
PAGINA = ROOT / "pages" / "02_Calculo_Represados.py"
FONTE = PAGINA.read_text(encoding="utf-8")


def _app(run=True):
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(PAGINA), default_timeout=120)
    if run:
        at.run()
    return at


def _html_regua(at):
    blocos = [str(md.value) for md in at.markdown if "cl8us-regua" in str(md.value)]
    assert blocos, "regua temporal nao renderizada"
    return blocos[-1]


# ---------------------------------------------------------------- 32.1 rotulo
ROTULO_LATERAL = "Data-base de referência 🔹"


def test_rotulo_lateral_renomeado_sem_mudar_valor():
    at = _app()
    rotulos = [d.label for d in at.date_input]
    assert ROTULO_LATERAL in rotulos
    assert not any("âncora inicial da análise atual" in r for r in rotulos)
    # O valor default do sistema permanece o mesmo (10/10/2022).
    campo = next(d for d in at.date_input if d.label == ROTULO_LATERAL)
    assert campo.value == date(2022, 10, 10)


def test_fonte_nao_criou_campo_novo_nem_texto_explicativo():
    assert FONTE.count("Data-base de referência 🔹") == 1
    for proibido in (
        "Marco anterior ao primeiro ciclo analisado",
        "Primeiro ciclo analisado",
        "Neste caso: C1",
    ):
        assert proibido not in FONTE


# ------------------------------------------------------------ cenarios A/B/C
def test_cenario_a_analise_iniciando_em_c1():
    at = _app()
    html = _html_regua(at)
    assert "Marcos dos ciclos" in html
    assert ">C1<" in html and ">C2<" in html
    # Etapa 43: o pertencimento a apuracao virou indicador global do cabecalho.
    assert html.count("Ciclos em análise: C1 e C2") == 1
    assert html.index("Ciclos em análise: C1 e C2") < html.index(">C1<")
    # Etapa 45: o marco do ciclo e a ELEGIBILIDADE (data-base + 12 meses), a
    # mesma data ja apresentada no bloco do ciclo. A data-base lateral
    # (10/10/2022) nao e mais o marco desenhado.
    assert "Apto em 10/10/2023" in html
    assert "Apto em 10/10/2024" in html


def test_cenario_b_analise_iniciando_em_c2():
    at = _app()
    at.selectbox(key="rep_ciclo_inicial_analise").select("C2")
    at.run()
    html = _html_regua(at)
    # A data lateral e a propria data-base de C2, sem avancar 12 meses: por
    # isso o marco apto de C2 e 10/10/2023, e nao 10/10/2024.
    assert ">C2<" in html and "Apto em 10/10/2023" in html
    assert html.count("Ciclo em análise: C2") == 1
    assert html.index("Ciclo em análise: C2") < html.index(">C2<")


def test_cenario_c_analise_iniciando_em_c4():
    at = _app()
    at.selectbox(key="rep_ciclo_inicial_analise").select("C4")
    at.run()
    html = _html_regua(at)
    assert ">C4<" in html
    assert html.count("Ciclo em análise: C4") == 1


@pytest.mark.parametrize(
    ("ciclo", "data_informada", "apto_esperado"),
    (
        ("C1", date(2023, 2, 1), "01/02/2024"),
        ("C2", date(2023, 2, 1), "01/02/2024"),
        ("C3", date(2024, 2, 1), "01/02/2025"),
        ("C4", date(2025, 10, 10), "10/10/2026"),
    ),
)
def test_data_lateral_e_semente_exata_do_ciclo_inicial(ciclo, data_informada, apto_esperado):
    """A data lateral semeia o ciclo inicial; o marco desenhado e ela + 12 meses.

    Etapa 45: a regua deixou de exibir a data-base e passou a exibir a data
    apta. A semente continua sendo exatamente a data digitada — o que se
    verifica pelo marco resultante, sempre 12 meses adiante dela.
    """
    at = _app()
    campo_base = next(d for d in at.date_input if d.label == ROTULO_LATERAL)
    campo_base.set_value(data_informada)
    at.selectbox(key="rep_ciclo_inicial_analise").select(ciclo)
    at.run()
    html = _html_regua(at)
    assert f">{ciclo}<" in html
    assert f"Apto em {apto_esperado}" in html


# ---------------------------------------------------------------- cenario D
def test_cenario_d_efeito_financeiro_posterior_a_data_base():
    at = _app()
    at.selectbox(key="rep_ciclo_inicial_analise").select("C2")
    at.run()
    pedido = next(
        d for d in at.date_input
        if d.label.startswith("Data do pedido da Contratada — C2")
    )
    # pedido tempestivo em competencia posterior a data-base -> efeito retardado
    pedido.set_value(date(2023, 12, 10))
    at.run()
    html = _html_regua(at)
    assert "Efeitos financeiros" in html
    # camada separada mostra a COMPETENCIA do efeito (01/12/2023), distinta do
    # marco apto do ciclo (10/10/2023), sem alterar o marco exibido.
    assert "01/12/2023" in html
    assert "Apto em 10/10/2023" in html
    # o pedido tardio aparece como marcador secundario, posterior ao marco
    assert "Pedido 10/12/2023" in html
    assert "cl8us-regua-pedido depois" in html


# ---------------------------------------------------------------- cenario E
def test_cenario_e_marco_anterior_ausente_nao_e_inventado():
    at = _app()
    at.selectbox(key="rep_ciclo_inicial_analise").select("C3")
    at.run()
    html = _html_regua(at)
    # sem historico formalizado, C0/C1/C2 nao tem marco disponivel na UI:
    # a regua nao pode inventa-los por inferencia retroativa.
    assert ">C0<" not in html
    assert ">C1<" not in html
    assert ">C2<" not in html
    assert ">C3<" in html


def test_cenario_e2_ciclo_formalizado_aparece_sem_inferencia():
    at = _app()
    campo_base = next(d for d in at.date_input if d.label == ROTULO_LATERAL)
    campo_base.set_value(date(2023, 2, 1))
    at.selectbox(key="rep_ciclo_inicial_analise").select("C2")
    at.run()
    at.radio(key="rep_situacao_anterior_ciclo").set_value(
        "Houve ciclo anterior concedido/formalizado"
    )
    at.run()
    at.date_input(key="rep_marco_temporal_anterior").set_value(date(2022, 2, 1))
    at.run()
    html = _html_regua(at)
    # o marco informado na lateral (dado existente, nao inferido) entra na regua
    assert ">C1<" in html
    assert ">C0<" not in html
    assert "01/02/2022" in html
    assert html.count("Ciclo em análise: C2") == 1
    # o indicador global cobre apenas a ANALISE (C2), nunca o marco historico
    assert "Ciclos em análise: C1" not in html
    # o marco historico nao substitui a data lateral que semeia C2: a lateral
    # (01/02/2023) continua sendo a data-base de C2, cujo marco apto e
    # 01/02/2024. O historico entra sem o rotulo "Apto em", por ser data-base.
    assert html.index(">C1<") < html.index("01/02/2022") < html.index(">C2<")
    assert html.index(">C2<") < html.index("Apto em 01/02/2024")
    assert "Apto em 01/02/2022" not in html


# ---------------------------------------------------------------- cenario F
def test_cenario_f_maximo_de_ciclos_legivel():
    at = _app()
    at.selectbox(key="rep_ciclo_final_analise").select("C4")
    at.run()
    html = _html_regua(at)
    for rotulo in (">C1<", ">C2<", ">C3<", ">C4<"):
        assert rotulo in html
    # legibilidade: colunas com largura minima e overflow contido no componente
    assert "minmax(104px, 1fr)" in html
    assert "overflow-x: auto" in html


# ------------------------------------------------- guardas de nao-regressao
def test_regua_e_somente_visual_sem_novo_motor():
    # A montagem le apenas estruturas ja existentes da pagina.
    assert "_render_regua_temporal_marcos(slot_regua_temporal, marcos_regua)" in FONTE
    ini = FONTE.index(">>> REGUA_TEMPORAL_MARCOS_V1: montagem")
    fim = FONTE.index("<<< REGUA_TEMPORAL_MARCOS_V1", ini)
    bloco = FONTE[ini:fim]
    # nenhuma aritmetica de datas (inferencia) na montagem da regua
    assert "relativedelta" not in bloco
    assert "timedelta" not in bloco
    for origem in ("input_ciclos", "contexto_contratual"):
        assert origem in bloco


def test_regua_fica_antes_do_bloco_de_ciclos():
    assert FONTE.index("slot_regua_temporal = st.container()") < FONTE.index(
        'st.markdown(f"### Ciclo {i}")'
    )
