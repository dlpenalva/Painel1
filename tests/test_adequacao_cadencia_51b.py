# -*- coding: utf-8 -*-
"""Etapa 51B — projecao da Adequacao Orcamentaria por CADENCIA real.

Protege a correcao do defeito conceitual: um gasto observado poucas vezes
NAO pode ser tratado como mensalidade. A inferencia separa QUANDO o gasto
ocorre (cadencia/posicao no ciclo) de QUANTO vale a ocorrencia, usando
ciclos preclusos como evidencia HISTORICA (sem gerar retroativo novo) e
mantendo C0 como fallback de baixa confianca. Tambem protege o uso do
FATOR EXATO (nunca o percentual visual arredondado).
"""
from datetime import date

import pandas as pd
import pytest

from _adequacao_orcamentaria import (
    CicloCadencia,
    CADENCIA_IRREGULAR,
    CADENCIA_MENSAL,
    CADENCIA_POR_CICLO,
    CADENCIA_SEMESTRAL,
    CADENCIA_TRIMESTRAL,
    cadencia_por_ciclo_forcada,
    eventos_mensais,
    inferir_cadencia,
    projetar_por_cadencia,
    Pedido,
)
from _adequacao_ui import (
    calcular_projecao,
    ciclos_para_cadencia,
    cronograma_por_exercicio,
    gerar_periodos_projecao,
    montar_base_editor,
    pares_de_financeiro,
    pares_de_pedidos,
)

VALOR_OCORRENCIA = 853961.0
FATOR_EXATO = 1.028899355


def _ciclos_caso_real():
    return [
        CicloCadencia("C0", date(2022, 10, 1), date(2023, 9, 30)),
        CicloCadencia("C1", date(2023, 10, 1), date(2024, 9, 30), precluso=True),
        CicloCadencia("C2", date(2024, 10, 1), date(2025, 9, 30), precluso=True),
        CicloCadencia("C3", date(2025, 10, 1), date(2026, 9, 30)),
    ]


def _pares_caso_real():
    """C0 com investimentos iniciais; C1/C2/C3 com UMA ocorrencia cada."""
    return [
        (2022, 10, VALOR_OCORRENCIA), (2022, 11, 120000.0), (2022, 12, 80000.0),
        (2024, 8, VALOR_OCORRENCIA),   # C1 (precluso) — posicao 10 do ciclo
        (2025, 8, VALOR_OCORRENCIA),   # C2 (precluso) — posicao 10 do ciclo
        (2026, 3, VALOR_OCORRENCIA),   # C3 (corrente) — ja executada
    ]


# ------------------------------------------------------------- 1. CASO REAL

def test_caso_real_uma_ocorrencia_por_ciclo_nao_vira_mensalidade():
    cad = inferir_cadencia(_pares_caso_real(), _ciclos_caso_real(), date(2026, 3, 1))
    assert cad["padrao"] == CADENCIA_POR_CICLO
    assert cad["ocorrencias_por_ciclo"] == 1
    assert cad["posicoes"] == [10]
    assert cad["valor_referencia"] == VALOR_OCORRENCIA
    # C1/C2 preclusos SAO a base historica; C0 (investimento) fica fora.
    assert cad["ciclos_base"] == ["C1", "C2"]
    assert cad["usa_c0"] is False

    base = projetar_por_cadencia(cad, _ciclos_caso_real(),
                                 date(2026, 4, 1), date(2027, 9, 30))
    # Uma UNICA ocorrencia futura no horizonte (C3 ja executou a sua;
    # o ciclo seguinte repete o perfil na posicao 10 -> ago/2027).
    assert base == {date(2027, 8, 1): VALOR_OCORRENCIA}


def test_caso_real_diferenca_futura_com_fator_exato():
    ciclos = _ciclos_caso_real()
    cad = inferir_cadencia(_pares_caso_real(), ciclos, date(2026, 3, 1))
    base = projetar_por_cadencia(cad, ciclos, date(2026, 4, 1), date(2027, 9, 30))
    periodos = gerar_periodos_projecao(pd.Period("2026-03", freq="M"), "30/09/2027")
    editor = montar_base_editor(periodos, 0.0, base_por_competencia=base)
    proj = calcular_projecao(editor, 0.0, FATOR_EXATO, base_por_competencia=base)
    ocorrencias = proj[proj["Valor base considerado"] > 0]
    assert len(ocorrencias) == 1              # NUNCA 18 mensalidades
    assert len(proj) == 18                    # horizonte completo, bases 0
    # R$ 853.961,00 x 1,028899355 = R$ 878.639,92 (fator EXATO, nao 2,89%)
    assert float(ocorrencias["Valor reajustado estimado"].iloc[0]) == 878639.92
    assert float(proj["Diferença futura a adequar"].sum()) == 24678.92


# ------------------------------------------------------------- 2. MENSAL real

def test_mensal_real_continua_mensal():
    pares = [(2025, m, 1000.0 + m) for m in range(1, 13)]
    cad = inferir_cadencia(pares, _ciclos_caso_real(), date(2025, 12, 1))
    assert cad["padrao"] == CADENCIA_MENSAL


# ------------------------------------------------------------- 3/4. SEMESTRAL e TRIMESTRAL

def test_semestral_projeta_duas_ocorrencias_por_ciclo():
    ciclos = _ciclos_caso_real()
    pares = [
        (2023, 12, 500.0), (2024, 6, 600.0),   # C1: posicoes 2 e 8
        (2024, 12, 520.0), (2025, 6, 620.0),   # C2: posicoes 2 e 8
    ]
    cad = inferir_cadencia(pares, ciclos, date(2025, 10, 1))
    assert cad["padrao"] == CADENCIA_SEMESTRAL
    assert cad["ocorrencias_por_ciclo"] == 2
    assert cad["posicoes"] == [2, 8]
    base = projetar_por_cadencia(cad, ciclos, date(2025, 11, 1), date(2026, 9, 30))
    assert set(base) == {date(2025, 12, 1), date(2026, 6, 1)}


def test_trimestral_projeta_quatro_ocorrencias_por_ciclo():
    ciclos = _ciclos_caso_real()
    pares = []
    for ano_ini, ciclo_inicio in ((2023, 10), (2024, 10)):
        for k in range(4):
            mes_ord = (ano_ini * 12 + (ciclo_inicio - 1)) + 3 * k
            pares.append((mes_ord // 12, mes_ord % 12 + 1, 100.0 + k))
    cad = inferir_cadencia(pares, ciclos, date(2025, 10, 1))
    assert cad["padrao"] == CADENCIA_TRIMESTRAL
    assert cad["posicoes"] == [0, 3, 6, 9]
    base = projetar_por_cadencia(cad, ciclos, date(2025, 10, 1), date(2026, 9, 30))
    assert len(base) == 4


# ------------------------------------------------------------- 5. IRREGULAR

def test_irregular_nao_espalha_media_nem_projeta():
    ciclos = _ciclos_caso_real()
    pares = [(2023, 11, 100.0),
             (2024, 11, 50.0), (2024, 12, 70.0), (2025, 2, 90.0), (2025, 7, 30.0)]
    cad = inferir_cadencia(pares, ciclos, date(2025, 10, 1))
    assert cad["padrao"] == CADENCIA_IRREGULAR
    assert projetar_por_cadencia(cad, ciclos, date(2025, 11, 1), date(2026, 12, 31)) == {}
    # Base 0 em todos os meses do editor: nenhuma mensalidade inventada.
    periodos = gerar_periodos_projecao(pd.Period("2025-10", freq="M"), "31/12/2026")
    proj = calcular_projecao(montar_base_editor(periodos, 999.0, base_por_competencia={}),
                             999.0, 1.1, base_por_competencia={})
    assert float(proj["Valor base considerado"].sum()) == 0.0


# ------------------------------------------------------------- 6. C0 extraordinario

def test_c0_com_investimento_inicial_nao_comanda_cadencia():
    ciclos = _ciclos_caso_real()
    pares = [(2022, m, 50000.0) for m in (10, 11, 12)] + [(2023, m, 40000.0) for m in (1, 2, 3)]
    pares += [(2024, 8, 900.0), (2025, 8, 910.0)]   # C1/C2 regulares: 1 por ciclo
    cad = inferir_cadencia(pares, ciclos, date(2025, 10, 1))
    assert cad["ciclos_base"] == ["C1", "C2"]
    assert cad["ocorrencias_por_ciclo"] == 1
    assert cad["usa_c0"] is False


def test_c0_como_fallback_tem_confianca_reduzida():
    ciclos = _ciclos_caso_real()[:2]   # apenas C0 completo e C1 sem eventos
    pares = [(2022, 12, 700.0)]
    cad = inferir_cadencia(pares, ciclos, date(2024, 9, 30))
    assert cad["usa_c0"] is True
    assert cad["confianca"] == "baixa"


# ------------------------------------------------------------- 7. PRECLUSOS

def test_ciclos_preclusos_informam_cadencia_sem_alterar_flag():
    ciclos = _ciclos_caso_real()
    cad = inferir_cadencia(_pares_caso_real(), ciclos, date(2026, 3, 1))
    # A cadencia veio EXCLUSIVAMENTE de ciclos preclusos (C1/C2): preclusao
    # juridica nao apaga o historico de execucao. Nada aqui recalcula
    # retroativo — a funcao nem recebe retroativo.
    assert cad["ciclos_base"] == ["C1", "C2"]
    assert all(c.precluso for c in ciclos if c.nome in cad["ciclos_base"])


# ------------------------------------------------------------- 9. HORIZONTE parcial

def test_ocorrencia_fora_do_horizonte_nao_e_projetada():
    ciclos = _ciclos_caso_real()
    cad = inferir_cadencia(_pares_caso_real(), ciclos, date(2026, 3, 1))
    base = projetar_por_cadencia(cad, ciclos, date(2026, 4, 1), date(2027, 6, 30))
    assert base == {}   # slot ago/2027 cai FORA da vigencia -> nada entra


# ------------------------------------------------------------- 10. PROGRAMACAO por exercicio

def test_programacao_por_exercicio_soma_somente_eventos_projetados():
    ciclos = _ciclos_caso_real()
    cad = inferir_cadencia(_pares_caso_real(), ciclos, date(2026, 3, 1))
    base = projetar_por_cadencia(cad, ciclos, date(2026, 4, 1), date(2027, 9, 30))
    periodos = gerar_periodos_projecao(pd.Period("2026-03", freq="M"), "30/09/2027")
    proj = calcular_projecao(montar_base_editor(periodos, 0.0, base_por_competencia=base),
                             0.0, FATOR_EXATO, base_por_competencia=base)
    retroativo = 25000.0
    cron = cronograma_por_exercicio(proj, retroativo)
    valores = {r["Exercício"]: float(r["Valor"]) for _, r in cron.iterrows()}
    assert valores == {"2026": 25000.0, "2027": 24678.92}
    assert round(sum(valores.values()), 2) == round(retroativo + 24678.92, 2)


# ------------------------------------------------------------- 11/12. FINANCEIRO e PCs

def test_origens_financeiro_e_pcs_alimentam_a_mesma_cadencia():
    fin = pd.DataFrame({
        "_periodo": [pd.Period("2024-08", freq="M"), pd.Period("2025-08", freq="M")],
        "valor": [VALOR_OCORRENCIA, VALOR_OCORRENCIA],
    })
    pares_fin = pares_de_financeiro(fin)
    peds = [Pedido(identificacao="PC-1", data=date(2024, 8, 15), valor=VALOR_OCORRENCIA),
            Pedido(identificacao="PC-2", data=date(2025, 8, 10), valor=VALOR_OCORRENCIA),
            Pedido(identificacao="PC-X", data=date(2025, 8, 10), valor=999.0, considerar=False)]
    pares_pc = pares_de_pedidos(peds)
    assert pares_fin == [(2024, 8, VALOR_OCORRENCIA), (2025, 8, VALOR_OCORRENCIA)]
    assert pares_pc == pares_fin   # PC excluido fica fora
    cad_fin = inferir_cadencia(pares_fin, _ciclos_caso_real(), date(2026, 3, 1))
    cad_pc = inferir_cadencia(pares_pc, _ciclos_caso_real(), date(2026, 3, 1))
    assert cad_fin["padrao"] == cad_pc["padrao"] == CADENCIA_POR_CICLO


# ------------------------------------------------------------- 13. MANUAL

def test_modo_manual_usa_somente_valores_informados():
    periodos = gerar_periodos_projecao(pd.Period("2026-03", freq="M"), "31/08/2026")
    editor = montar_base_editor(periodos, 500.0, base_por_competencia={})
    editor.loc[1, "Valor informado pelo fiscal"] = "1.000,00"
    proj = calcular_projecao(editor, 500.0, 1.1, base_por_competencia={},
                             origem_automatica="Premissa manual")
    assert float(proj["Valor base considerado"].sum()) == 1000.0
    assert (proj["Origem"] == "Valor informado pelo fiscal").sum() == 1
    assert set(proj.loc[proj["Valor base considerado"] == 0, "Origem"]) == {"Premissa manual"}


# ------------------------------------------------------------- premissa POR CICLO forcada

def test_premissa_por_ciclo_forcada_sobre_historico_irregular():
    ciclos = _ciclos_caso_real()
    pares = [(2024, 11, 50.0), (2024, 12, 70.0), (2025, 2, 90.0)]
    cad = cadencia_por_ciclo_forcada(pares, ciclos)
    assert cad["padrao"] == CADENCIA_POR_CICLO
    assert cad["ocorrencias_por_ciclo"] == 1
    assert cad["valor_referencia"] == 70.0   # mediana dos valores


# ------------------------------------------------------------- adaptadores

def test_ciclos_para_cadencia_le_df_ciclos_com_preclusao():
    df = pd.DataFrame({
        "Ciclo": ["C0", "C1", "C2"],
        "Data-base": ["01/10/2022", "01/10/2023", "01/10/2024"],
        "Situação automática": ["Base", "❌ PRECLUSO", "✅ TEMPESTIVO"],
    })
    ciclos = ciclos_para_cadencia(df)
    assert [c.nome for c in ciclos] == ["C0", "C1", "C2"]
    assert ciclos[0].inicio == date(2022, 10, 1)
    assert ciclos[0].fim == date(2023, 9, 1)      # mes anterior ao inicio seguinte
    assert ciclos[1].precluso is True
    assert ciclos[2].precluso is False
    assert ciclos[2].fim == date(2025, 9, 1)      # ultimo ciclo: inicio + 11 meses


def test_eventos_mensais_somam_lancamentos_do_mesmo_mes_e_ignoram_zero():
    eventos = eventos_mensais([(2025, 1, 100.0), (2025, 1, 50.0), (2025, 2, 0.0),
                               (2025, 3, None), (2025, 4, 20.0)])
    assert eventos == [{"mes": date(2025, 1, 1), "valor": 150.0},
                       {"mes": date(2025, 4, 1), "valor": 20.0}]


# ------------------------------------------------------------- 8/14. pagina (estatico)

def _pagina():
    from pathlib import Path
    return (Path(__file__).resolve().parents[1] /
            "pages" / "12_Adequacao_Orcamentaria.py").read_text(encoding="utf-8")


def test_pagina_usa_fator_exato_e_nao_percentual_visual():
    pagina = _pagina()
    assert 'percentual_reajuste = float(ctx["variacao"])' in pagina
    assert "round-trip" in pagina   # comentario da correcao permanece


def test_pagina_tem_premissa_de_projecao_com_estado_isolado():
    pagina = _pagina()
    assert 'key="adequacao_v3_premissa"' in pagina
    for opcao in ("Automática (cadência histórica)", "Mensal (média)", "Por ciclo", "Manual"):
        assert f'"{opcao}"' in pagina
    # Troca de premissa recria o editor (estado nao contamina entre modos).
    assert "adequacao_v3_editor_{origem_hist}_{premissa_proj}_" in pagina
    assert "HISTÓRICO SEM PERIODICIDADE SUFICIENTE" in pagina


# ------------------------------------------------------------- CASO REAL ponta a ponta (AppTest)

def _sessao_caso_real():
    df_fin = pd.DataFrame({
        "Ciclo": ["c0", "c0", "c0", "c1", "c2", "c3"],
        "Competência": ["10/2022", "11/2022", "12/2022", "08/2024", "08/2025", "03/2026"],
        "Valor pago/faturado": [VALOR_OCORRENCIA, 120000.0, 80000.0,
                                VALOR_OCORRENCIA, VALOR_OCORRENCIA, VALOR_OCORRENCIA],
    })
    df_ciclos = pd.DataFrame({
        "Ciclo": ["C0", "C1", "C2", "C3"],
        "Data-base": ["01/10/2022", "01/10/2023", "01/10/2024", "01/10/2025"],
        "Situação automática": ["Base", "❌ PRECLUSO", "❌ PRECLUSO", "✅ TEMPESTIVO"],
    })
    return {
        "df_financeiro_mensal": df_fin,
        "df_ciclos": df_ciclos,
        "valor_represado_a_pagar": 25000.0,
        "variacao_acumulada": FATOR_EXATO - 1,
        "indice": "IST (Anatel)",
        "quantidade_ciclos": "1",
        "modo_apuracao": "Completo",
    }


def test_caso_real_na_web_projeta_por_ciclo_e_usa_fator_exato():
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file("pages/12_Adequacao_Orcamentaria.py", default_timeout=120)
    at.session_state["resultado_valor_global"] = _sessao_caso_real()
    at.run()
    at.text_input(key="adequacao_v3_data_final_vigencia").set_value("30/09/2027")
    at.run()
    assert not at.exception, f"pagina quebrou no caso real: {at.exception}"
    blob = "\n".join(str(m.value) for m in at.markdown)
    # padrao identificado e transparencia da premissa
    assert "1 ocorrência por ciclo" in blob
    # fator EXATO: 853.961,00 x 1,028899355 = 878.639,92 (com 2,89% seria 878.640,47)
    assert "878.639,92" in blob
    assert "878.640,47" not in blob
    # PROJECAO ORCAMENTARIA PROPORCIONAL: o valor recorrente do ciclo
    # (853.961,00 / 12 meses) cobre TODOS os 18 meses restantes (abr/26 a
    # set/27) — diferenca futura = 18 x 2.056,57 = 37.018,26 — em vez de
    # projetar apenas a ocorrencia historica (que deixaria meses sem
    # cobertura orcamentaria). Jamais 18 mensalidades INTEIRAS.
    assert "37.018,26" in blob
    # programacao por exercicio: retroativo + 9 meses proporcionais em 2026,
    # 9 meses proporcionais em 2027
    tabelas = []
    for df_el in at.dataframe:
        df = df_el.value
        if isinstance(df, pd.DataFrame) and "Exercício" in df.columns:
            tabelas.append({str(r["Exercício"]): str(r["Valor"]) for _, r in df.iterrows()})
    assert any(t.get("2026") == "R$ 43.509,13" and t.get("2027") == "R$ 18.509,13"
               and t.get("TOTAL") == "R$ 62.018,26" for t in tabelas), tabelas


# ------------------------------------------ PROJECAO ORCAMENTARIA PROPORCIONAL

def _cadencia_por_ciclo_homolog():
    return {"padrao": CADENCIA_POR_CICLO, "valor_referencia": VALOR_OCORRENCIA,
            "duracao_ciclo_meses": 12}


def test_proporcional_caso_homologacao_21_meses():
    """Base 853.961,00/ciclo de 12 meses, fator exato, 21 meses restantes
    (2026=4, 2027=12, 2028=5). Valores canonicos do motor (arredondamento
    mensal ROUND_HALF_UP), proximos dos alvos de homologacao."""
    from _adequacao_orcamentaria import projetar_por_ciclo_proporcional
    base = projetar_por_ciclo_proporcional(_cadencia_por_ciclo_homolog(),
                                           date(2026, 9, 1), date(2028, 5, 1))
    assert len(base) == 21
    assert all(abs(v - VALOR_OCORRENCIA / 12) < 1e-9 for v in base.values())

    periodos = list(pd.period_range("2026-09", "2028-05", freq="M"))
    editor = montar_base_editor(periodos, 0.0, base_por_competencia=base)
    proj = calcular_projecao(editor, 0.0, FATOR_EXATO, base_por_competencia=base)
    # todos os 21 meses cobertos (nenhum zerado)
    assert (proj["Diferença futura a adequar"] > 0).all()
    por_ano = proj.assign(ano=proj["Competência"].str[-2:]) \
                  .groupby("ano")["Diferença futura a adequar"].sum().round(2)
    assert por_ano["26"] == pytest.approx(8226.28)
    assert por_ano["27"] == pytest.approx(24678.84)
    assert por_ano["28"] == pytest.approx(10282.85)
    # alvos de homologacao (aprox.: arredondamento mensal canonico)
    assert abs(por_ano["26"] - 8226.31) < 0.25
    assert abs(por_ano["27"] - 24678.92) < 0.25
    assert abs(por_ano["28"] - 10282.88) < 0.25
    dif_futura = round(float(proj["Diferença futura a adequar"].sum()), 2)
    assert dif_futura == pytest.approx(43187.97)
    assert abs(dif_futura - 43188.11) < 0.25
    # soma dos exercicios = diferenca futura; + retroativo = complementacao
    retro = 24678.92
    cron = cronograma_por_exercicio(proj, retro)
    soma_exercicios = round(float(cron["Valor"].sum()), 2)
    assert soma_exercicios == pytest.approx(round(dif_futura + retro, 2))
    assert abs(soma_exercicios - 67867.03) < 0.25


def test_proporcional_contrato_terminando_em_julho_cobre_jan_a_jul():
    """Ocorrencia historica em ago/set + vigencia ate julho: jan-jul do ultimo
    exercicio NAO ficam zerados (cobertura proporcional)."""
    from _adequacao_orcamentaria import projetar_por_ciclo_proporcional
    base = projetar_por_ciclo_proporcional(_cadencia_por_ciclo_homolog(),
                                           date(2026, 10, 1), date(2027, 7, 1))
    meses_2027 = [d for d in base if d.year == 2027]
    assert sorted(m.month for m in meses_2027) == [1, 2, 3, 4, 5, 6, 7]
    assert all(base[m] > 0 for m in meses_2027)
    periodos = list(pd.period_range("2026-10", "2027-07", freq="M"))
    proj = calcular_projecao(montar_base_editor(periodos, 0.0, base_por_competencia=base),
                             0.0, FATOR_EXATO, base_por_competencia=base)
    assert (proj["Diferença futura a adequar"] > 0).all()


def test_proporcional_nao_se_aplica_a_mensal_nem_irregular():
    """Mensal continua mensal (media, sem proporcionalizacao); IRREGULAR
    continua fail-closed: a proporcionalizacao devolve {} para ambos."""
    from _adequacao_orcamentaria import projetar_por_ciclo_proporcional
    for padrao in (CADENCIA_MENSAL, CADENCIA_IRREGULAR):
        cad = {"padrao": padrao, "valor_referencia": VALOR_OCORRENCIA,
               "duracao_ciclo_meses": 12}
        assert projetar_por_ciclo_proporcional(cad, date(2026, 1, 1), date(2026, 12, 1)) == {}
    # sem valor de referencia positivo tambem nao inventa cobertura
    cad = {"padrao": CADENCIA_POR_CICLO, "valor_referencia": 0.0, "duracao_ciclo_meses": 12}
    assert projetar_por_ciclo_proporcional(cad, date(2026, 1, 1), date(2026, 12, 1)) == {}
