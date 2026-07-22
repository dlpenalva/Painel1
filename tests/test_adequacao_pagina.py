"""Pagina 12 (Adequacao Orcamentaria) alinhada ao metodo da planilha 04.

Paginas Streamlit executam ao importar; verificacao estatica (le a fonte), como
nos demais testes de pagina do projeto. A matematica e testada em
test_adequacao_orcamentaria.py (motor + golden + Excel COM).
"""
from __future__ import annotations

from pathlib import Path

PAGINA = (Path(__file__).resolve().parents[1] / "pages" / "12_Adequacao_Orcamentaria.py").read_text(encoding="utf-8")


def test_duas_origens_sem_terceiro_metodo():  # A, B
    assert 'st.radio(' in PAGINA and '"Financeiro"' in PAGINA and '"Pedidos de compra"' in PAGINA
    assert 'st.subheader("Origem histórica")' in PAGINA


def test_usa_o_motor_unico():  # H — sem matematica duplicada na UI
    assert "from _adequacao_orcamentaria import" in PAGINA
    assert "media_pedidos_compra" in PAGINA and "_round2" in PAGINA
    assert "def calcular_adequacao_orcamentaria" not in PAGINA
    assert "def media_pedidos_compra" not in PAGINA


def test_pc_reutiliza_itens_pc_e_janela():  # D, E, F
    assert "carregar_itens_pc_da_sessao" in PAGINA and "pedidos_de_itens_pc(" in PAGINA
    assert 'st.slider("Janela histórica dos pedidos (meses)", 1, 60' in PAGINA


def test_controles_pc_dentro_do_ramo_pcs():  # C — Financeiro nao mostra controles de PC
    idx_pc = PAGINA.index('if origem_hist == "Pedidos de compra":')
    idx_fin = PAGINA.index('if origem_hist == "Financeiro":')
    idx_slider = PAGINA.index('Janela histórica dos pedidos')
    assert idx_pc < idx_slider < idx_fin  # slider de PC pertence ao ramo PCs


def test_blocos_do_metodo_presentes():  # I, J, P, Q
    for bloco in ("Referência mensal", "Resultado da adequação",
                  "Programação por exercício", "Arquivos da adequação"):
        assert f'st.subheader("{bloco}")' in PAGINA


def test_referencia_reajustada_e_complemento():  # I, N
    assert "referencia_reajustada = _round2(media_ref * fator_reajuste)" in PAGINA
    assert "complementacao = _round2(float(retroativo or 0) + diferenca_futura)" in PAGINA
    assert "COMPLEMENTAÇÃO NECESSÁRIA" in PAGINA


def test_memorando_removido_do_fluxo_principal():  # P
    assert 'st.subheader("Memorando")' not in PAGINA
    assert "Contrato para o memorando" not in PAGINA
    assert "Reajustes/ciclos para o memorando" not in PAGINA


def test_sem_tabela_vazia_de_projecao():  # Q — nao exibe dataframe vazio
    assert "Não há competências futuras a projetar" in PAGINA
    assert "A programação por exercício depende de projeção futura" in PAGINA
