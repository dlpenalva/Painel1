"""Pagina 12 (Adequacao Orcamentaria) — UX V3 (planilha guiada, 4 abas).

Paginas Streamlit executam ao importar; verificacao estatica (le a fonte), como
nos demais testes de pagina do projeto. A matematica e testada no motor
(test_adequacao_orcamentaria.py) e o golden funcional via AppTest
(test_adequacao_web_golden.py / test_adequacao_ux_v3.py).
"""
from __future__ import annotations

from pathlib import Path

PAGINA = (Path(__file__).resolve().parents[1] / "pages" / "12_Adequacao_Orcamentaria.py").read_text(encoding="utf-8")
UI = (Path(__file__).resolve().parents[1] / "_adequacao_ui.py").read_text(encoding="utf-8")


def test_quatro_abas_planilha_guiada():  # Secao 3, 41.A
    assert 'st.tabs(' in PAGINA
    for aba in ('"1. Base"', '"2. Histórico"', '"3. Projeção"', '"4. Resultado"'):
        assert aba in PAGINA, aba


def test_duas_origens_radio_v3():  # Secao 7 (controle clean testavel)
    assert 'st.radio(' in PAGINA and '"Financeiro"' in PAGINA and '"Pedidos de Compra"' in PAGINA
    assert 'key="adequacao_v3_origem"' in PAGINA


def test_seletor_origem_unico_e_na_base():  # Secao 6, E, L
    assert PAGINA.count('key="adequacao_v3_origem"') == 1          # existe apenas um seletor
    idx_sel = PAGINA.index('key="adequacao_v3_origem"')
    idx_tab2 = PAGINA.index("TAB 2 — HISTORICO")
    assert idx_sel < idx_tab2, "o seletor de origem deve estar na aba Base (antes da TAB 2)"


def test_historico_consome_origem_sem_reofertar_seletor():  # Secao 6, F
    assert "Histórico utilizado:" in PAGINA


def test_pc_separa_elegibilidade_de_exclusao_manual():  # Secao 3/4/G
    assert "atualizar_exclusoes_manuais_pc(" in PAGINA
    # o editor de PC recria por competencia/janela (nao persiste "fora da janela")
    assert 'key=f"adequacao_v3_pc_editor_' in PAGINA


def test_financeiro_situacao_considerada_derivada():  # Secao 7/8
    assert "situacao_financeira_considerada(" in PAGINA
    assert '"Situação considerada"' in PAGINA


def test_usa_o_motor_e_o_view_model_sem_matematica_na_ui():  # Secao 2, 39
    assert "from _adequacao_orcamentaria import" in PAGINA
    assert "from _adequacao_ui import" in PAGINA
    # a pagina nao redefine a matematica do motor nem duplica formulas
    assert "def calcular_adequacao_orcamentaria" not in PAGINA
    assert "def media_pedidos_compra" not in PAGINA
    # o view-model delega ao motor (nao redefine o motor)
    assert "from _adequacao_orcamentaria import" in UI
    assert "def calcular_adequacao_orcamentaria" not in UI
    assert "def media_pedidos_compra" not in UI


def test_pc_usa_checkbox_e_nao_multiselect():  # Secao 14
    assert "carregar_itens_pc_da_sessao" in PAGINA and "pedidos_de_itens_pc(" in PAGINA
    assert 'st.multiselect(' not in PAGINA
    assert 'CheckboxColumn("USAR")' in PAGINA


def test_slider_janela_fica_em_opcoes_avancadas():  # Secao 15, 43
    assert 'st.slider("Janela histórica dos pedidos (meses)", 1, 60' in PAGINA
    assert 'with st.expander("Opções avançadas do histórico"' in PAGINA


def test_data_vigencia_e_o_input_principal():  # Secao 6, 41.C
    assert 'key="adequacao_v3_data_final_vigencia"' in PAGINA
    assert "Data final da vigência contratual" in PAGINA


def test_fator_e_ajustes_fora_do_fluxo_normal():  # Secao 8, 9, 41.D/E
    # "Fator usado" so aparece dentro do expander de Opcoes avancadas.
    idx_exp = PAGINA.index('st.expander("Opções avançadas"')
    idx_fator = PAGINA.index("Fator usado")
    assert idx_exp < idx_fator, "fator tecnico deve estar em Opcoes avancadas"
    assert '"adequacao_v3_retroativo"' in PAGINA   # ajuste do retroativo (avancado)
    assert 'key="adequacao_v3_percentual"' in PAGINA


def test_reset_limpa_somente_v3():  # Secao 26
    assert "Limpar ajustes desta adequação" in PAGINA
    assert 'startswith("adequacao_v3_")' in PAGINA
    # nao pode limpar apuracao/coleta/documentos
    assert 'pop("resultado_valor_global"' not in PAGINA
    assert 'pop("diagnostico_coleta_v2"' not in PAGINA
    assert "st.session_state.clear()" not in PAGINA


def test_resultado_e_complemento():  # Secao 23, N
    assert "referencia_reajustada = _round2(media_ref * fator_reajuste)" in PAGINA
    assert "complementacao = _round2(float(retroativo or 0) + diferenca_futura)" in PAGINA
    assert "COMPLEMENTAÇÃO NECESSÁRIA" in PAGINA


def test_projecao_usa_referencia_da_origem():  # Secao 17/18 — media_ref comum as origens
    assert "montar_base_editor(periodos, media_ref)" in PAGINA
    assert "calcular_projecao(df_editor, media_ref, fator_reajuste)" in PAGINA
    # instrucao direta acima da tabela (Secao 19)
    assert "Deixe vazio para usar a média." in PAGINA


def test_xlsx_only_sem_docx_memorando():  # Secao 25, AA
    assert "gerar_xlsx_projecao(" in PAGINA and '"Baixar XLSX"' in PAGINA
    assert 'st.subheader("Memorando")' not in PAGINA
    assert "gerar_docx_memorando" not in PAGINA
    assert "Contrato para o memorando" not in PAGINA


def test_export_identifica_origem():  # Q
    assert '("Origem histórica", origem_hist_rotulo)' in PAGINA


def test_sem_pagina_linear_antiga_duplicada():  # Secao 48
    # subheaders da arquitetura linear antiga nao devem sobreviver
    for antigo in ('st.subheader("Dados importados da apuração")',
                   'st.subheader("Parâmetros principais")',
                   'st.subheader("Origem histórica")',
                   'st.subheader("Referência mensal")',
                   'st.subheader("Resultado da adequação")'):
        assert antigo not in PAGINA, antigo
