# -*- coding: utf-8 -*-
"""VTA-U2 (regra global) — um unico VTA canonico em XLS, web e documentos.

Existe UM VTA oficial por processamento. Toda superficie externa que apresente
VTA / VTA Oficial / Valor Total Atualizado deve usar o VTA canonico calculado
pela metodologia selecionada — nunca a posicao fisica (B10/B11), nunca o
comparativo do contrato integralmente reajustado (B12/B28), nunca a formula
antiga B23.

Estes testes sao de UNIFORMIDADE, nao de igualdade forcada: a convergencia
continua provada em `test_vta_u2_uniformizacao.py`, onde XLS e Python calculam
8.713.820,26 por caminhos independentes.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from _resultado_consolidado import (  # noqa: E402
    ORIGEM_VTA_CANONICA,
    ORIGEM_VTA_INDISPONIVEL,
    montar_resultado_consolidado,
)

# Caso Financeiro real de referencia.
EXECUTADO = 7_300_890.27
AJUSTES = 24_678.92
REMANESCENTE = 1_388_251.07
VTA = 8_713_820.26

# Valores que NUNCA podem ser apresentados como VTA.
POSICAO_FISICA = 7_835_180.34
VTA_FORMULA_ANTIGA = 8_324_859.61
COMPARATIVO_INTEGRAL = 8_499_931.01

# Superficies externas auditadas.
PAGINAS_WEB = sorted((RAIZ / "pages").glob("*.py"))
DOCUMENTOS = (
    RAIZ / "_sumario_executivo.py",
    RAIZ / "_templates_documentos.py",
)

# Fontes alternativas de VTA que nao podem alimentar apresentacao externa.
FONTES_ALTERNATIVAS = (
    "forma1_posicao_atual",
    "forma2_ultima_abertura",
    "forma3_integral_reajustado",
    "VTA_ATUALIZACAO_CHEIA",
)


def _sem_comentarios(fonte: str) -> str:
    """Remove linhas de comentario para nao acusar mencoes documentais."""
    return "\n".join(
        linha for linha in fonte.split("\n")
        if not linha.lstrip().startswith("#")
    )


# ------------------------------------------------------- A. superficies web

@pytest.mark.parametrize("pagina", PAGINAS_WEB, ids=lambda p: p.name)
def test_a1_nenhuma_pagina_web_le_fonte_alternativa_de_vta(pagina):
    """Nenhuma pagina pode ler posicao fisica / ultima abertura / comparativo."""
    fonte = _sem_comentarios(pagina.read_text(encoding="utf-8"))
    for alternativa in FONTES_ALTERNATIVAS:
        assert alternativa not in fonte, (
            f"{pagina.name} le a fonte alternativa {alternativa!r}; "
            "a web deve apresentar apenas o VTA canonico."
        )
    assert "comparativo_VTA" not in fonte


@pytest.mark.parametrize("pagina", PAGINAS_WEB, ids=lambda p: p.name)
def test_a2_nenhuma_pagina_web_rotula_referencia_fisica_como_vta(pagina):
    fonte = pagina.read_text(encoding="utf-8")
    for rotulo in (
        "VTA pela posição atual",
        "VTA pela última posição",
        "VTA por posição física",
        "Utilizado o VTA da última posição",
    ):
        assert rotulo not in fonte, f"{pagina.name} exibe {rotulo!r} como VTA."


def test_a3_card_e_rodape_do_vta_usam_o_consolidado_canonico():
    """Na pagina de resultados, o card e o rodape do VTA leem consolidado.vta."""
    fonte = (RAIZ / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")
    trecho = fonte[fonte.index("def render_resultado_consolidado"):]
    trecho = trecho[:trecho.index("\ndef ")]
    assert "Valor Total Atualizado — VTA" in trecho
    assert 'consolidado.get("vta")' in trecho
    # a primeira ocorrencia do seletor esta no CSS; a renderizacao e a ultima.
    rodape = fonte[fonte.rindex("resultado-rodape-vta"):]
    rodape = rodape[:rodape.index("unsafe_allow_html=True")]
    assert "VALOR TOTAL ATUALIZADO — VTA OFICIAL" in rodape
    assert 'consolidado.get("vta")' in rodape


# ------------------------------------------------------- B. documentos

@pytest.mark.parametrize("modulo", DOCUMENTOS, ids=lambda p: p.name)
def test_b1_documentos_nao_apresentam_fonte_alternativa_de_vta(modulo):
    """Sumario, Saneador e Apostila nao podem imprimir VTA alternativo."""
    fonte = _sem_comentarios(modulo.read_text(encoding="utf-8"))
    for alternativa in FONTES_ALTERNATIVAS:
        assert alternativa not in fonte, (
            f"{modulo.name} imprime a fonte alternativa {alternativa!r}."
        )


@pytest.mark.parametrize("modulo", DOCUMENTOS, ids=lambda p: p.name)
def test_b2_documentos_sem_rotulos_de_vta_concorrente(modulo):
    fonte = _sem_comentarios(modulo.read_text(encoding="utf-8"))
    for rotulo in (
        "VTA pela posição atual do contrato",
        "VTA pela última posição de abertura disponível",
        "Contrato original integralmente reajustado",
        "COMPARATIVO",
    ):
        assert rotulo not in fonte, f"{modulo.name} apresenta {rotulo!r}."


def test_b3_comparativo_interno_permanece_disponivel_para_auditoria():
    """A remocao e de APRESENTACAO: o leitor segue expondo o comparativo."""
    leitor = (RAIZ / "_leitor_masterfile_v10.py").read_text(encoding="utf-8")
    assert "forma3_integral_reajustado" in leitor
    assert "forma1_posicao_atual" in leitor


# ------------------------------------------------- C. fail-closed do VTA

def _resultado(vta_canonico, *, forma1=POSICAO_FISICA, forma2=POSICAO_FISICA):
    return {
        "valor_atualizado_contrato": vta_canonico,
        "valor_represado_a_pagar": AJUSTES,
        "controle": {"modo": "principal", "ciclo_vigente": "C3"},
        "memoria_por_ciclo": {"vta": {"metodo": "financeiro"}},
        "referencias_vta": {
            "posicao_atual_disponivel": forma1 is not None,
            "forma1_posicao_atual": forma1,
            "forma2_ultima_abertura": forma2,
            "forma3_integral_reajustado": COMPARATIVO_INTEGRAL,
        },
        "composicao_vta": {
            "disponivel": True, "bloqueia_formalizacao": False,
            "linhas": [
                {"descricao": "Executado apurado", "valor_atualizado": EXECUTADO},
                {"descricao": "Ajustes ainda devidos", "valor_atualizado": AJUSTES},
                {"descricao": "Remanescente atualizado",
                 "valor_atualizado": REMANESCENTE},
            ],
            "alertas": [],
        },
        "politica_entrega_segura": {
            "status": "PRONTO_PARA_VALIDACAO_FISCAL", "pendencias": [],
            "retroativo": {"metodo": "financeiro"},
        },
        "reconciliacao_xls_python": {"status_geral": "CONCILIADO"},
    }


def test_c1_consolidado_entrega_o_vta_canonico():
    consolidado = montar_resultado_consolidado(_resultado(VTA))
    assert consolidado["vta"] == VTA
    for proibido in (POSICAO_FISICA, VTA_FORMULA_ANTIGA, COMPARATIVO_INTEGRAL):
        assert consolidado["vta"] != proibido


def test_c2_sem_vta_canonico_nenhuma_referencia_assume_o_lugar():
    """Fail-closed: nem posicao atual, nem ultima abertura, nem comparativo."""
    consolidado = montar_resultado_consolidado(
        _resultado(None, forma1=None, forma2=POSICAO_FISICA)
    )
    assert consolidado["vta"] is None
    assert consolidado["vta_origem"] == "indisponivel"
    assert consolidado["vta_usa_ultima_posicao"] is False


def test_c1b_origem_do_vta_e_o_calculo_canonico_em_todo_metodo():
    """VTA-U2.2: `vta_origem` diz de onde o valor VEM. Removido o fallback pela
    posicao fisica, so existe um caminho — o calculo canonico da metodologia.
    Nenhum metodo pode reportar "posicao_atual" nem "ultima_posicao_disponivel".
    """
    for modo, metodo in (("principal", "financeiro"), ("pc", "pc"),
                         ("d", "consumidos")):
        resultado = _resultado(VTA)
        resultado["controle"] = {"modo": modo, "ciclo_vigente": "C3"}
        resultado["memoria_por_ciclo"] = {"vta": {"metodo": metodo}}
        consolidado = montar_resultado_consolidado(resultado)

        assert consolidado["vta"] == VTA, modo
        assert consolidado["vta_origem"] == ORIGEM_VTA_CANONICA, modo
        assert consolidado["vta_origem"] not in (
            "posicao_atual", "ultima_posicao_disponivel"
        ), modo
        assert consolidado["vta_usa_ultima_posicao"] is False, modo


def test_c2b_origem_indisponivel_quando_nao_ha_vta_canonico():
    consolidado = montar_resultado_consolidado(
        _resultado(None, forma1=None, forma2=POSICAO_FISICA)
    )
    assert consolidado["vta_origem"] == ORIGEM_VTA_INDISPONIVEL


def test_c2c_o_codigo_nao_produz_mais_as_origens_antigas():
    """A taxonomia antiga nao pode voltar por outro caminho."""
    fonte = _sem_comentarios(
        (RAIZ / "_resultado_consolidado.py").read_text(encoding="utf-8")
    )
    trecho = fonte[fonte.index("vta_atual = resultado.get"):]
    trecho = trecho[:trecho.index("resultado_incompleto")]
    assert '"posicao_atual"' not in trecho
    assert '"ultima_posicao_disponivel"' not in trecho
    assert "ORIGEM_VTA_CANONICA" in trecho


def test_c3_composicao_visivel_soma_exatamente_o_vta():
    """Item 13: a soma das parcelas apresentadas fecha com o VTA canonico."""
    consolidado = montar_resultado_consolidado(_resultado(VTA))
    linhas = (consolidado["composicao_vta"] or {}).get("linhas") or []
    assert linhas
    soma = round(sum(l["valor_atualizado"] for l in linhas), 2)
    assert soma == consolidado["vta"] == VTA
    assert soma == round(EXECUTADO + AJUSTES + REMANESCENTE, 2)


# --------------------------------------- D. nenhuma superficie hardcoda o VTA

@pytest.mark.parametrize(
    "modulo",
    [*PAGINAS_WEB, *DOCUMENTOS, RAIZ / "_resultado_consolidado.py",
     RAIZ / "_motor_composicao_vta.py"],
    ids=lambda p: p.name,
)
def test_d1_nenhum_valor_do_caso_real_aparece_hardcodado(modulo):
    """A uniformidade e consequencia do calculo, nunca de constante fixa."""
    fonte = modulo.read_text(encoding="utf-8")
    for numero in ("8713820", "8_713_820", "7835180", "7_835_180",
                   "8324859", "8_324_859", "8499931", "8_499_931"):
        assert numero not in fonte, (
            f"{modulo.name} contem o literal {numero!r}; o VTA tem de mudar "
            "quando as entradas mudarem."
        )


def test_d2_o_motor_python_nao_le_o_resultado_do_xls():
    """Independencia real preservada apos a uniformizacao."""
    fonte = (RAIZ / "_motor_composicao_vta.py").read_text(encoding="utf-8")
    for proibido in ("resultados_xls", "VTA_FINAL", "reconciliacao_xls_python"):
        assert proibido not in fonte


# ------------------------------------------ E. PC e Consumido inalterados

def test_e1_pc_e_consumido_seguem_com_metodologia_propria():
    """Cada metodo tem UM VTA canonico proprio; nenhum foi adaptado ao
    Financeiro (o ramo dedicado do Financeiro nao os alcanca)."""
    fonte = (RAIZ / "_motor_composicao_vta.py").read_text(encoding="utf-8")
    assert 'if modo == "pc":' in fonte
    assert 'if modo == "principal":' in fonte
    assert fonte.index('if modo == "pc":') < fonte.index('if modo == "principal":')
    assert "execucao = _execucao_por_ciclo(leitura, por_ciclo, alertas)" in fonte


def test_e2_regra_de_um_unico_vta_vale_para_todo_metodo():
    """A escolha do VTA no consolidado nao depende do metodo: e sempre o
    canonico do proprio metodo, nunca uma referencia estrangeira."""
    fonte = (RAIZ / "_resultado_consolidado.py").read_text(encoding="utf-8")
    # recorte da ESCOLHA do VTA: comeca depois do dicionario de referencias
    # auditaveis (que legitimamente le forma1/forma2 como referencia) e termina
    # antes da classificacao de completude.
    trecho = fonte[fonte.index("# VTA-U2 (achado A)"):]
    trecho = trecho[:trecho.index("resultado_incompleto")]
    assert "vta = vta_atual" in trecho
    assert "forma1_posicao_atual" not in trecho
    assert "forma2_ultima_abertura" not in trecho
    assert not re.search(r"vta\s*=\s*vta_ultima", trecho)
    # e a escolha nao se ramifica por metodo
    assert "metodo_consumidos" not in trecho
    assert "metodo_pc" not in trecho
