"""Foco: aba "5. Texto SIGA" da Adequacao Orcamentaria (apresentacional).

A aba nao introduz calculo novo: consome os valores canonicos ja produzidos
pela aba 4 (retroativo, diferenca futura, complementacao confirmada,
retroativo potencial, cenario com potencial). Estes testes protegem:

  1. a aba "5. Texto SIGA" existe;
  2. o TOTAL A READEQUAR do texto e exatamente a COMPLEMENTACAO CONFIRMADA
     da aba 4 (mesmo numero, mesma fonte);
  3. o retroativo potencial NAO integra o total a readequar;
  4. os meses do texto vem da projecao ja calculada (qtd_meses);
  5. o cenario com potencial usa o valor canonico ja calculado na aba 4;
  6. a vigencia (aba 1) aparece formatada dd/mm/aaaa no texto;
  7. editar a clausula de reajuste (campo manual) reflete no texto;
  8. ausencia de identificacao contratual produz placeholder, nunca
     None/nan;
  9. a aba 4 continua reproduzindo os mesmos resultados de antes;
  10. o box de retroativo potencial da aba 4 recebeu SOMENTE alteracao
      visual (mesmo valor, cor de alerta vermelho suave).
"""
from __future__ import annotations

import re
from datetime import date

import pandas as pd
import pytest
from streamlit.testing.v1 import AppTest

from _adequacao_ui import moeda, parse_moeda_br

CORTE = date(2026, 8, 19)
VIGENCIA = "29/02/2028"
RECONHECIDO = 57701.49
POTENCIAL = 12000.0

_ITENS_PC = [
    {"numero_pc": "PC-1", "data_pc": date(2025, 3, 14), "valor_pc": 12000.0},
    {"numero_pc": "PC-2", "data_pc": date(2025, 11, 5), "valor_pc": 18000.0},
    {"numero_pc": "PC-3", "data_pc": date(2026, 6, 20), "valor_pc": 9000.0},
]


def _sessao_pc(*, com_potencial=True):
    consolidado = {
        "medidas_pc_aplicaveis": True,
        "retroativo_reconhecido": RECONHECIDO,
        "valor_atualizado_em_analise": 41000.0,
        "fora_do_corte": {"aplicavel": True, "data_corte": CORTE},
    }
    if com_potencial:
        consolidado["retroativo_potencial"] = POTENCIAL
    return {
        "itens_pc_v10": {"itens": [dict(i) for i in _ITENS_PC]},
        "valor_represado_a_pagar": RECONHECIDO,
        "variacao_acumulada": 0.0289,
        "modo_apuracao": "Completo",
        "resultado_consolidado": consolidado,
    }


def _run_pc(sessao=None, *, vigencia=VIGENCIA):
    at = AppTest.from_file("pages/12_Adequacao_Orcamentaria.py", default_timeout=180)
    at.session_state["resultado_valor_global"] = sessao if sessao is not None else _sessao_pc()
    at.run()
    at.radio(key="adequacao_v3_origem").set_value("Pedidos de Compra")
    at.text_input(key="adequacao_v3_data_final_vigencia").set_value(vigencia)
    at.run()
    return at


def _blob(at):
    return "\n".join(str(m.value) for m in at.markdown)


def _texto_siga(at):
    return at.text_area(key="adequacao_v3_siga_texto_area").value


def _linha(texto, prefixo):
    achadas = [l for l in texto.splitlines() if l.startswith(prefixo)]
    assert achadas, f"linha {prefixo!r} nao encontrada no texto"
    return achadas[0]


def _valor_do_card(blob, rotulo):
    idx = blob.index(rotulo)
    achado = re.search(r"(-?[\d\.]+,\d{2})", blob[idx + len(rotulo):])
    assert achado, f"valor nao encontrado para {rotulo!r}"
    return parse_moeda_br(achado.group(1))


# --------------------------------------------------------------- 1. aba existe

def test_aba_texto_siga_existe():
    at = _run_pc()
    assert not at.exception, at.exception
    assert len(at.tabs) == 5
    assert at.tabs[4].label == "5. Texto SIGA"


# --------------------------------------------------------- 2. total = confirmada

def test_total_a_readequar_e_exatamente_a_complementacao_confirmada_da_aba4():
    at = _run_pc()
    blob = _blob(at)
    complementacao = _valor_do_card(blob, "COMPLEMENTAÇÃO CONFIRMADA")

    texto = _texto_siga(at)
    total_texto = parse_moeda_br(_linha(texto, "TOTAL A READEQUAR:").split(":", 1)[1])
    assert total_texto == pytest.approx(complementacao)
    assert moeda(complementacao) in _linha(texto, "TOTAL A READEQUAR:")


# --------------------------------------------------- 3. potencial fora do total

def test_retroativo_potencial_nao_integra_o_total_a_readequar():
    at = _run_pc(_sessao_pc(com_potencial=True))
    blob = _blob(at)
    complementacao = _valor_do_card(blob, "COMPLEMENTAÇÃO CONFIRMADA")

    texto = _texto_siga(at)
    total_texto = parse_moeda_br(_linha(texto, "TOTAL A READEQUAR:").split(":", 1)[1])
    # o total bate com a complementacao confirmada (que ja exclui o
    # potencial) e jamais com complementacao + potencial.
    assert total_texto == pytest.approx(complementacao)
    assert total_texto != pytest.approx(complementacao + POTENCIAL)
    # o potencial aparece, mas apenas na secao informativa, nunca somado.
    assert moeda(POTENCIAL) in texto
    linha_potencial = _linha(texto, "Retroativo potencial, ainda em aceitação")
    assert moeda(POTENCIAL) in linha_potencial


# --------------------------------------------------------------- 4. meses da projecao

def test_meses_do_texto_vem_da_projecao_ja_calculada():
    at = _run_pc()
    blob = _blob(at)
    m = re.search(r"(\d+)\s*meses", blob)
    assert m, "indicador de meses nao encontrado na aba 3/4"
    qtd_meses_pagina = m.group(1)

    texto = _texto_siga(at)
    assert f"– {qtd_meses_pagina} meses:" in texto
    assert f"para {qtd_meses_pagina} meses:" in texto


# --------------------------------------------------------------- 5. cenario com potencial

def test_cenario_com_potencial_usa_o_valor_canonico_ja_calculado():
    at = _run_pc(_sessao_pc(com_potencial=True))
    blob = _blob(at)
    cenario_pagina = _valor_do_card(blob, "Cenário com potencial")

    texto = _texto_siga(at)
    linha_cenario = _linha(texto, "Cenário de planejamento considerando")
    cenario_texto = parse_moeda_br(linha_cenario.split(":", 1)[1])
    assert cenario_texto == pytest.approx(cenario_pagina)


def test_sem_potencial_localizado_o_texto_nao_inventa_valor():
    at = _run_pc(_sessao_pc(com_potencial=False))
    texto = _texto_siga(at)
    assert "Não localizado nesta apuração" in texto
    assert "Não aplicável (retroativo potencial não localizado)" in texto
    assert "None" not in texto
    assert "nan" not in texto.lower()


# --------------------------------------------------------------- 6. vigencia formatada

def test_vigencia_exibida_dd_mm_aaaa():
    at = _run_pc(vigencia="29/02/2028")
    texto = _texto_siga(at)
    assert "com vigência até 29/02/2028." in texto


# --------------------------------------------------------------- 7. clausula editavel

def test_alteracao_manual_da_clausula_reflete_no_texto():
    at = _run_pc()
    assert "previsto na Cláusula Oitava." in _texto_siga(at)
    at.text_input(key="adequacao_v3_siga_clausula").set_value("Cláusula Nona")
    at.run()
    texto = _texto_siga(at)
    assert "previsto na Cláusula Nona." in texto
    assert "Cláusula Oitava" not in texto


def test_alteracao_manual_de_contrato_e_contratada_reflete_no_texto():
    at = _run_pc()
    at.text_input(key="adequacao_v3_siga_contrato").set_value("12/2024")
    at.text_input(key="adequacao_v3_siga_contratada").set_value("Empresa XPTO S.A.")
    at.run()
    texto = _texto_siga(at)
    assert "para o Contrato 12/2024, firmado com a Empresa XPTO S.A.," in texto


# --------------------------------------------------- 8. sem identificacao -> placeholder

def test_ausencia_de_identificacao_produz_placeholder_nunca_none_nan():
    at = _run_pc()
    texto = _texto_siga(at)
    assert "[campo a preencher]" in texto
    assert "None" not in texto
    assert "nan" not in texto.lower()
    assert "0/1900" not in texto


# --------------------------------------------------------------- 9. aba 4 preservada

def test_aba_resultado_mantem_os_mesmos_resultados_de_antes():
    at = _run_pc()
    assert not at.exception, at.exception
    blob = _blob(at)
    assert "COMPLEMENTAÇÃO CONFIRMADA" in blob
    assert "Retroativo reconhecido considerado" in blob
    assert "Programação por exercício" in blob
    assert _valor_do_card(blob, "Retroativo reconhecido considerado") == pytest.approx(RECONHECIDO)


# --------------------------------------------- 10. box potencial: so mudanca visual

def test_box_retroativo_potencial_recebe_apenas_alteracao_visual():
    pagina = open("pages/12_Adequacao_Orcamentaria.py", encoding="utf-8").read()
    assert ('render_card_valor("Retroativo potencial (não reconhecido)", '
            'retroativo_potencial, alerta=True)') in pagina
    # a assinatura de render_card_valor ganhou o parametro visual sem
    # alterar nenhuma chamada existente (nenhuma outra invocacao passa
    # alerta=).
    assert pagina.count("alerta=True") == 1
    assert "def render_card_valor(label, valor, nota=\"\", destaque=False, formato=\"moeda\", alerta=False):" in pagina

    at = _run_pc(_sessao_pc(com_potencial=True))
    blob = _blob(at)
    # o valor exibido no card continua sendo exatamente retroativo_potencial.
    assert _valor_do_card(blob, "Retroativo potencial (não reconhecido)") == pytest.approx(POTENCIAL)
