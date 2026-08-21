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
  7. o bloco de campos manuais (Contrato/Contratada/Cláusula de reajuste)
     foi retirado da interface: cláusula usa o mesmo default fixo de antes,
     contrato/contratada vêm direto da fonte automática já existente;
  8. ausencia de identificacao contratual produz placeholder, nunca
     None/nan;
  9. a aba 4 continua reproduzindo os mesmos resultados de antes;
  10. o box de retroativo potencial da aba 4 recebeu SOMENTE alteracao
      visual (mesmo valor, cor de alerta vermelho suave).

ANTI-CONGELAMENTO (bug reproduzido em producao real no PR #83): um
st.text_area com key fixa so usa o parametro value= na primeira vez que o
widget e registrado — em reruns seguintes o Streamlit preserva o valor ja
associado a key, ignorando o novo value= calculado (confirmado em
streamlit/runtime/state/session_state.py, SessionState.register_widget).
O texto ficava congelado no primeiro calculo (antes da vigencia ser
informada) e nunca mais atualizava. O AppTest nao reproduz esse mecanismo
espontaneamente (testado empiricamente), entao os testes abaixo forcam
deliberadamente o cenario congelado — escrevendo um valor obsoleto em
session_state antes do rerun, como aconteceria no navegador real — e
provam que a logica de assinatura (comparar os insumos e resincronizar
st.session_state[key] antes do widget) sempre corrige o texto exibido.
Validado tambem manualmente em navegador real (local e producao): o .txt
baixado e byte a byte identico ao texto exibido em todos os casos.
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


# --------------------------------------------- 1b. bloco superior removido (limpeza)

def test_bloco_de_campos_manuais_removido_sem_espaco_vazio():
    """Prova a remoção cirúrgica pedida: a frase introdutória e os 3 campos
    manuais (com seus ícones de ajuda) somem da interface, sem deixar um
    st.caption/st.columns órfão no lugar; o texto passa a começar direto no
    subheader seguido da área de cópia."""
    pagina = open("pages/12_Adequacao_Orcamentaria.py", encoding="utf-8").read()
    bloco_aba5 = pagina[pagina.index("with tab_siga:"):]

    assert "Confira os dados abaixo e copie o texto para utilização no processo." not in bloco_aba5
    assert 'st.text_input(\n            "Contrato"' not in bloco_aba5
    assert '"Contratada", value=' not in bloco_aba5
    assert '"Cláusula de reajuste"' not in bloco_aba5
    assert "col_id1, col_id2, col_id3" not in bloco_aba5
    assert (
        "Preenchido automaticamente quando disponível na apuração; editável."
        not in bloco_aba5
    )

    at = _run_pc()
    assert not at.exception, at.exception
    assert "adequacao_v3_siga_contrato" not in at.session_state
    assert "adequacao_v3_siga_contratada" not in at.session_state
    assert "adequacao_v3_siga_clausula" not in at.session_state


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


# ------------------------------------------------- 7. clausula fixa (campo removido)

def test_clausula_e_fixa_sem_campo_manual():
    """PR de limpeza da aba 5: o campo manual "Cláusula de reajuste" foi
    retirado da interface. Sem fonte automática para esse dado, o texto
    preserva o MESMO default fixo que o campo já usava antes ("Cláusula
    Oitava"), nunca uma fonte nova."""
    at = _run_pc()
    assert "previsto na Cláusula Oitava." in _texto_siga(at)
    assert "adequacao_v3_siga_clausula" not in {w.key for w in at.text_input}


def test_contrato_e_contratada_vem_automaticos_sem_campo_manual():
    """PR de limpeza da aba 5: os campos manuais "Contrato" e "Contratada"
    foram retirados da interface. O texto passa a usar diretamente a fonte
    automática já existente (_valor_contratual_automatico), sem qualquer
    interação de widget."""
    sessao = _sessao_pc()
    sessao["contexto_contratual_anterior"] = {
        "contrato": "12/2024", "contratada": "Empresa XPTO S.A.",
    }
    at = _run_pc(sessao)
    chaves_widgets = {w.key for w in at.text_input}
    assert "adequacao_v3_siga_contrato" not in chaves_widgets
    assert "adequacao_v3_siga_contratada" not in chaves_widgets
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


# --------------------------------------------- anti-congelamento (regressao)

def test_widget_nao_passa_value_junto_com_key_no_text_area():
    """Causa-raiz do congelamento: value= e key= juntos fazem o Streamlit
    ignorar o novo value em reruns seguintes ao primeiro. A correcao exige
    que o text_area nao receba mais value= (o conteudo vem exclusivamente
    de st.session_state[key], sincronizado explicitamente antes)."""
    pagina = open("pages/12_Adequacao_Orcamentaria.py", encoding="utf-8").read()
    assert 'st.text_area("Texto pronto para copiar", value=texto_siga' not in pagina
    assert 'key="adequacao_v3_siga_texto_area"' in pagina
    assert 'st.session_state["adequacao_v3_siga_texto_area"] = texto_siga' in pagina


def test_texto_area_resincroniza_mesmo_se_session_state_ficar_congelado():
    """Simula exatamente o bug observado em producao: um valor OBSOLETO
    gravado em session_state[chave do text_area] antes de um rerun (e' o
    que acontecia quando o widget ficava preso no primeiro calculo, antes
    da vigencia ser informada). Prova que a logica de assinatura detecta a
    divergencia dos insumos atuais e resincroniza o widget."""
    at = _run_pc()
    texto_correto = _texto_siga(at)
    at.session_state["adequacao_v3_siga_texto_area"] = "TEXTO OBSOLETO DE UM RUN ANTERIOR CONGELADO"
    at.session_state["adequacao_v3_siga_assinatura"] = "assinatura-que-nao-bate-mais"
    # qualquer rerun, mesmo sem nenhuma interacao nova de widget, deve
    # resincronizar o texto a partir da assinatura dos insumos atuais.
    at.run()
    assert not at.exception, at.exception
    assert "TEXTO OBSOLETO" not in _texto_siga(at)
    assert _texto_siga(at) == texto_correto


def test_reruns_sucessivos_mudando_a_vigencia_nao_congelam_o_texto():
    """Reproduz reruns sucessivos alterando a vigencia (aba 1) — unico
    insumo do texto ainda editavel pelo usuario apos a limpeza da aba 5 —
    o mesmo padrao de interacao do navegador real que expos o bug original.
    O texto deve refletir cada mudanca imediatamente, nunca preservando um
    valor de um rerun anterior."""
    at = AppTest.from_file("pages/12_Adequacao_Orcamentaria.py", default_timeout=180)
    at.session_state["resultado_valor_global"] = _sessao_pc(com_potencial=True)
    at.run()
    assert "[campo a preencher]" in _texto_siga(at)
    assert "0 meses" in _texto_siga(at)

    at.radio(key="adequacao_v3_origem").set_value("Pedidos de Compra")
    at.text_input(key="adequacao_v3_data_final_vigencia").set_value(VIGENCIA)
    at.run()
    assert "29/02/2028" in _texto_siga(at)
    assert "18 meses" in _texto_siga(at)
    assert "[campo a preencher]" in _texto_siga(at)  # sem contrato/contratada na apuracao

    at.text_input(key="adequacao_v3_data_final_vigencia").set_value("31/03/2028")
    at.run()
    texto_final = _texto_siga(at)
    assert "31/03/2028" in texto_final
    assert "29/02/2028" not in texto_final  # nao preserva a vigencia do rerun anterior


# --------------------------------------------- 11. programacao por exercicio

def test_programacao_por_exercicio_usa_a_mesma_fonte_da_aba4():
    """A secao 'Programacao por exercicio' do Texto SIGA consome exatamente
    o mesmo DataFrame (cronograma_por_exercicio) ja renderizado na aba 4 —
    mesmos anos, mesmos valores, mesmo TOTAL. Nada e recalculado aqui."""
    at = _run_pc(_sessao_pc(com_potencial=True))
    blob = _blob(at)
    complementacao = _valor_do_card(blob, "COMPLEMENTAÇÃO CONFIRMADA")

    df_cron = next(d.value for d in at.dataframe if "Exercício" in d.value.columns)
    texto = _texto_siga(at)
    assert "Programação por exercício" in texto
    bloco = texto.split("Programação por exercício", 1)[1]

    for _, row in df_cron.iterrows():
        assert f"{row['Exercício']} | {row['Valor']}" in bloco

    linha_total = _linha(texto, "TOTAL | ")
    total_texto = parse_moeda_br(linha_total.split("|", 1)[1])
    assert total_texto == pytest.approx(complementacao)


def test_txt_baixado_usa_a_mesma_variavel_do_texto_exibido_sem_duplicar():
    """NAO duplicar a montagem do texto: o download_button precisa usar a
    MESMA variavel Python texto_siga que sincroniza o widget — nunca uma
    segunda montagem/f-string independente. Prova estatica (uma unica
    definicao de texto_siga no arquivo) + funcional (a mesma variavel
    alimenta session_state[key] e o data= do download_button)."""
    pagina = open("pages/12_Adequacao_Orcamentaria.py", encoding="utf-8").read()
    assert pagina.count("texto_siga = (") == 1
    assert 'st.session_state["adequacao_v3_siga_texto_area"] = texto_siga' in pagina
    assert 'data=texto_siga.encode("utf-8")' in pagina

    at = _run_pc()
    # o valor que alimentaria o download (texto_siga daquele run) e
    # exatamente o que ficou sincronizado em session_state[key].
    assert at.session_state["adequacao_v3_siga_texto_area"] == _texto_siga(at)
