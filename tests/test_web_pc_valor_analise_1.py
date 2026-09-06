# -*- coding: utf-8 -*-
"""WEB-PC-VALOR-ANALISE-1 — redacao do card do valor dos PCs nao pagos.

O card do metodo PC publicava "Valor atualizado em analise", rotulo que nao
dizia de que valor se tratava. Passou a nomear o conceito (PCs realizados, ja
reajustados, ainda nao pagos) e a trazer a explicacao na nota discreta que a
propria pagina ja usa.

Escopo do teste: REDACAO. Ele protege o texto novo e, principalmente, garante
que nada em volta mudou — mesma fonte de valor, mesmo gate, mesma posicao.
"""

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PAGINA = (ROOT / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")

ROTULO_NOVO = "Valor atualizado dos PCs realizados e ainda não pagos"
ROTULO_ANTIGO = "Valor atualizado em análise"


def test_titulo_novo_substituiu_o_antigo_na_pagina():
    assert f'_ROTULO_VALOR_PC_NAO_PAGO = "{ROTULO_NOVO}"' in PAGINA
    assert ROTULO_ANTIGO not in PAGINA


def test_explicacao_acompanha_o_card():
    assert (
        '"Valor dos PCs já realizados, com reajuste aplicado, que permanecem "'
        '\n        "em análise pela área gestora e ainda não foram pagos à contratada."'
        in PAGINA
    )
    # A nota chega pelo mesmo parametro que o card do potencial ja usava, ou
    # seja, pela classe discreta `.resultado-nota-vta` — sem card novo.
    assert "nota=_NOTAS_SEGUNDA_LINHA.get(rotulo)" in PAGINA
    assert '<div class="resultado-nota-vta">' in PAGINA


def test_fonte_do_valor_e_o_gate_do_metodo_pc_seguem_intactos():
    """Nada de calculo mudou: mesma chave, mesmo formatador, mesmo gate."""
    assert (
        '        if consolidado.get("medidas_pc_aplicaveis"):\n'
        "            colunas_segunda_linha.append((\n"
        "                _ROTULO_VALOR_PC_NAO_PAGO,\n"
        '                _moeda_resultado(consolidado.get("valor_atualizado_em_analise")),\n'
        "            ))"
    ) in PAGINA


def test_demais_celulas_da_linha_nao_ganharam_nota():
    """A nota e exclusiva do card dos PCs; as vizinhas continuam sem nota."""
    inicio = PAGINA.index("_NOTAS_SEGUNDA_LINHA = {")
    fim = PAGINA.index("}", PAGINA.index("),", inicio))
    assert PAGINA[inicio:fim].count(":") == 1
    assert "Fora da data de corte" in PAGINA
    assert 'colunas_segunda_linha.append(("Formalização", formalizacao_exibicao))' in PAGINA
