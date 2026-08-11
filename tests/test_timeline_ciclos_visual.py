# -*- coding: utf-8 -*-
"""Timeline dos ciclos — ajuste exclusivamente visual da pagina "Analisar varios ciclos".

O componente e dinamico (acompanha os ciclos efetivamente computados), reusa
apenas datas que a propria pagina ja calculou e nao altera admissibilidade,
efeitos financeiros, estado ou persistencia.
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
    blocos = [str(md.value) for md in at.markdown if "cl8us-regua-grid" in str(md.value)]
    assert blocos, "timeline dos ciclos nao renderizada"
    return blocos[-1]


def _marcacao_regua(at):
    """Somente a marcacao renderizada, sem o bloco <style> do componente."""
    return _html_regua(at).split("</style>", 1)[-1]


def _notas_ciclos(at):
    return [str(md.value) for md in at.markdown if "cl8us-ciclo-nota" in str(md.value)]


def _captions(at):
    return [str(c.value) for c in at.caption]


def _selecionar_intervalo(at, inicial, final):
    at.selectbox(key="rep_ciclo_inicial_analise").select(inicial)
    at.run()
    at.selectbox(key="rep_ciclo_final_analise").select(final)
    at.run()
    return at


# ------------------------------------------------- componente dinamico (1/2/3)
@pytest.mark.parametrize(
    ("inicial", "final", "esperados"),
    (
        ("C2", "C2", ("C2",)),
        ("C1", "C2", ("C1", "C2")),
        ("C1", "C3", ("C1", "C2", "C3")),
    ),
)
def test_timeline_acompanha_os_ciclos_computados(inicial, final, esperados):
    at = _selecionar_intervalo(_app(), inicial, final)
    html = _html_regua(at)
    for posicao, rotulo in enumerate(esperados, start=1):
        assert f">{rotulo}<" in html
        assert f"{posicao}º ciclo da análise" in html
    # nenhum ciclo excedente e desenhado
    assert f"{len(esperados) + 1}º ciclo da análise" not in html
    assert f"repeat({len(esperados)}, minmax(104px, 1fr))" in html


def test_ciclos_da_analise_nao_sao_rotulados_como_proximo_ciclo():
    at = _selecionar_intervalo(_app(), "C1", "C2")
    html = _html_regua(at)
    assert "2º ciclo da análise" in html
    assert "Próximo ciclo" not in html
    for nota in _notas_ciclos(at):
        assert "Próximo ciclo" not in nota


# --------------------------------------- status global "ciclos em analise"
@pytest.mark.parametrize(
    ("inicial", "final", "frase"),
    (
        ("C2", "C2", "Ciclo em análise: C2"),
        ("C1", "C2", "Ciclos em análise: C1 e C2"),
        ("C1", "C3", "Ciclos em análise: C1, C2 e C3"),
        ("C1", "C4", "Ciclos em análise: C1, C2, C3 e C4"),
    ),
)
def test_status_global_lista_todos_os_ciclos_da_apuracao(inicial, final, frase):
    at = _selecionar_intervalo(_app(), inicial, final)
    html = _html_regua(at)
    marcacao = _marcacao_regua(at)
    # a frase e unica, dinamica e concorda em numero com a quantidade de ciclos
    assert marcacao.count(frase) == 1
    # o selo local desapareceu: nenhum ciclo carrega o status individualmente
    assert "EM ANÁLISE" not in html
    assert "cl8us-regua-tag" not in html
    # e o indicador pertence ao cabecalho, antes da linha temporal
    assert marcacao.index(frase) < marcacao.index("cl8us-regua-grid")
    assert marcacao.index("cl8us-regua-cabecalho") < marcacao.index(frase)


def test_status_global_nao_e_repetido_por_ciclo():
    at = _selecionar_intervalo(_app(), "C1", "C3")
    marcacao = _marcacao_regua(at)
    assert marcacao.count("cl8us-regua-analise-global") == 1
    assert marcacao.count("em análise") == 1
    # a ordem individual de cada ciclo permanece intacta
    for posicao in (1, 2, 3):
        assert f"{posicao}º ciclo da análise" in marcacao
    assert "Próximo ciclo" not in marcacao


def test_status_global_ignora_marco_historico_fora_da_apuracao():
    """Ciclo formalizado aparece na regua, mas nunca na frase da apuracao."""
    at = _app()
    campo_base = next(
        d for d in at.date_input
        if d.label == "Data-base do ciclo em que **esta análise começa** 🔹"
    )
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
    # o historico esta desenhado na regua (com sua propria data)...
    assert ">C1<" in html and "01/02/2022" in html
    assert "Ciclo formalizado" in html
    # ...mas a apuracao corrente e apenas C2.
    assert html.count("Ciclo em análise: C2") == 1
    assert "Ciclos em análise: C1" not in html
    assert "EM ANÁLISE" not in html


def test_frase_do_status_global_e_montada_a_partir_de_em_analise():
    """Guarda de unidade: a lista sai de em_analise, nao da posicao na regua."""
    # a pagina e um script Streamlit; interessa apenas o helper puro de texto.
    fonte_helper = FONTE[
        FONTE.index("def _frase_ciclos_em_analise") : FONTE.index(
            "def _render_regua_temporal_marcos"
        )
    ]
    espaco = {}
    exec(compile(fonte_helper, str(PAGINA), "exec"), espaco)
    frase = espaco["_frase_ciclos_em_analise"]
    historico = {"rotulo": "C1", "em_analise": False}
    assert frase([]) == ""
    assert frase([historico]) == ""
    assert frase([{"rotulo": "C2", "em_analise": True}]) == "Ciclo em análise: C2"
    assert frase(
        [historico, {"rotulo": "C2", "em_analise": True}, {"rotulo": "C3", "em_analise": True}]
    ) == "Ciclos em análise: C2 e C3"
    assert frase(
        [{"rotulo": f"C{n}", "em_analise": True} for n in range(1, 5)]
    ) == "Ciclos em análise: C1, C2, C3 e C4"


# ------------------------------------------------------------- marcos por ciclo
def test_cada_marco_traz_identificacao_data_base_e_status():
    at = _app()
    html = _html_regua(at)
    assert "Marcos dos ciclos" in html
    # C1 semeado pela data lateral (10/10/2022) e C2 pela cadeia ja calculada.
    assert "10/10/2022" in html and "10/10/2023" in html
    assert html.count("TEMPESTIVO") == 2
    # o pertencimento a apuracao e global (cabecalho), nunca um selo por ciclo
    assert html.count("Ciclos em análise: C1 e C2") == 1
    assert "EM ANÁLISE" not in html


def test_status_adiantado_aparece_na_timeline():
    at = _app()
    pedido = next(d for d in at.date_input if d.label.startswith("Data do Pedido C1"))
    pedido.set_value(date(2023, 5, 10))
    at.run()
    html = _html_regua(at)
    assert "ADIANTADO" in html
    # status individual e pertencimento a analise sao conceitos distintos:
    # o ciclo segue anunciado no cabecalho mesmo classificado como ADIANTADO.
    assert html.count("Ciclos em análise: C1 e C2") == 1
    assert "EM ANÁLISE" not in html


def test_efeitos_financeiros_em_linha_propria_quando_pedido_difere_do_inicio():
    at = _app()
    pedido = next(d for d in at.date_input if d.label.startswith("Data do Pedido C1"))
    # tempestivo em competencia posterior a data-base: efeito retardado
    pedido.set_value(date(2023, 12, 10))
    at.run()
    html = _html_regua(at)
    assert "Efeitos financeiros" in html
    # a data-base do ciclo permanece 10/10/2022; o efeito nasce em 01/12/2023
    assert "10/10/2022" in html
    assert "01/12/2023" in html
    assert "TEMPESTIVO*" in html


# ------------------------------------------------------ relacao pedido x ciclo
def test_frase_relaciona_pedido_e_data_base_do_ciclo():
    at = _app()
    notas = _notas_ciclos(at)
    assert len(notas) == 2
    assert "Pedido de 10/10/2023 refere-se ao ciclo C1, iniciado em 10/10/2022." in notas[0]
    assert "Pedido de 10/10/2024 refere-se ao ciclo C2, iniciado em 10/10/2023." in notas[1]
    # destaque azul discreto apenas no ciclo que abre a analise
    assert "cl8us-ciclo-nota ativo" in notas[0]
    assert "cl8us-ciclo-nota ativo" not in notas[1]


def test_frase_acompanha_a_data_informada_pelo_usuario():
    at = _app()
    campo_base = next(
        d for d in at.date_input
        if d.label == "Data-base do ciclo em que **esta análise começa** 🔹"
    )
    campo_base.set_value(date(2023, 9, 7))
    at.run()
    pedido = next(d for d in at.date_input if d.label.startswith("Data do Pedido C1"))
    pedido.set_value(date(2024, 9, 7))
    at.run()
    notas = _notas_ciclos(at)
    assert "Pedido de 07/09/2024 refere-se ao ciclo C1, iniciado em 07/09/2023." in notas[0]


def test_ajuda_sob_o_campo_de_pedido_mostra_a_data_base_de_cada_ciclo():
    at = _app()
    captions = _captions(at)
    assert "Este pedido utiliza a data-base de 10/10/2022." in captions
    assert "Este pedido utiliza a data-base de 10/10/2023." in captions


# --------------------------------------------------------------- responsividade
def test_timeline_e_responsiva_sem_estourar_a_largura():
    at = _selecionar_intervalo(_app(), "C1", "C4")
    html = _html_regua(at)
    for rotulo in (">C1<", ">C2<", ">C3<", ">C4<"):
        assert rotulo in html
    assert "minmax(104px, 1fr)" in html
    assert "overflow-x: auto" in html
    assert "@media (max-width: 640px)" in html


def test_css_restrito_a_esta_secao():
    # nenhum seletor global/agressivo: todas as regras sao prefixadas.
    for bloco in ("_REGUA_TEMPORAL_CSS", "_TIMELINE_CICLOS_CSS"):
        ini = FONTE.index(f"{bloco} = ")
        fim = FONTE.index('"""', FONTE.index('"""', ini) + 3)
        css = FONTE[ini:fim]
        for linha in css.splitlines():
            linha = linha.strip()
            if not linha.startswith("."):
                continue
            assert linha.startswith(".cl8us-"), linha
        assert "div[data-testid" not in css
        assert "stApp" not in css


# ----------------------------------------------- guardas de nao-regressao
def test_timeline_nao_recalcula_nada():
    ini = FONTE.index(">>> REGUA_TEMPORAL_MARCOS_V1: montagem")
    fim = FONTE.index("<<< REGUA_TEMPORAL_MARCOS_V1", ini)
    bloco = FONTE[ini:fim]
    assert "relativedelta" not in bloco
    assert "timedelta" not in bloco
    # status e datas vem de estruturas ja produzidas pela pagina
    assert "sit_emoji" in bloco
    assert "data_semente_exata" in bloco
    assert "inicio_efeito_financeiro" in bloco


def test_resultados_calculados_permanecem_intactos():
    """Datas, admissibilidade e efeitos do cenario padrao seguem os mesmos."""
    at = _app()
    linhas = [
        linha
        for df in at.dataframe
        for linha in df.value.to_dict("records")
        if "Ciclo" in linha and "Data do pedido" in linha
    ]
    assert linhas, "resumo pre-processamento nao encontrado"
    por_ciclo = {linha["Ciclo"]: linha for linha in linhas}
    assert por_ciclo["C1"]["Data do pedido"] == "10/10/2023"
    assert por_ciclo["C1"]["Início financeiro"] == "01/10/2023"
    assert "TEMPESTIVO" in por_ciclo["C1"]["Situação preliminar"]
    assert por_ciclo["C2"]["Data do pedido"] == "10/10/2024"
    assert por_ciclo["C2"]["Início financeiro"] == "01/10/2024"
