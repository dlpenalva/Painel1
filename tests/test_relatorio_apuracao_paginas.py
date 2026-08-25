# -*- coding: utf-8 -*-
"""NIVEL C — Relatorio de Apuracao padronizado nas duas calculadoras.

Protege a redacao (apresentacao) e a INVARIANCIA das regras/datas:

  * "Janela de Admissibilidade (90 dias)" nas duas paginas;
  * ciclo unico ordinario com UMA linha de variacao (sem "Percentual aplicado
    no acumulado" e sem "Variação acumulada");
  * percentual aplicado aparece SOMENTE quando diverge da variacao apurada
    (ciclo negativo) — e no negocial como "por acordo";
  * preclusao sem efeito nao inventa "Início dos efeitos financeiros";
  * terminologia padronizada ("Pedido realizado em", "Período de apuração do
    índice", "Índice:", "Início dos efeitos financeiros");
  * as linhas que CALCULAM as datas de admissibilidade permanecem identicas
    (a mudanca "(90 dias)" e somente textual).
"""
from __future__ import annotations

from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
SIMPLES = (RAIZ / "pages" / "01_Calculo_Simples.py").read_text(encoding="utf-8")
MULTI = (RAIZ / "pages" / "02_Calculo_Represados.py").read_text(encoding="utf-8")


# ---------------------------------------------------------------------------
# Datas de admissibilidade: regra intocada (somente o rotulo ganhou o sufixo)
# ---------------------------------------------------------------------------

def test_regra_de_datas_do_ciclo_unico_intacta():
    assert "dt_aniv = dt_base + relativedelta(years=1)" in SIMPLES
    assert "dt_limite = dt_aniv + relativedelta(days=90)" in SIMPLES


def test_regra_de_datas_do_multiciclo_intacta():
    # Commit 2928265 (tag HOMOLOGADO_CL8US_MATRIZ_20260810_1728, Etapas 38/45):
    # a elegibilidade do multiciclo passou a vir da CADEIA EXATA — a ancora
    # exata avanca 12 meses e d_aniv espelha essa referencia, em vez do salto
    # mensalizado de 1 ano sobre a data_atual.
    assert "data_referencia_exata = data_atual + relativedelta(months=12)" in MULTI
    assert "d_aniv = data_referencia_exata" in MULTI
    assert "d_lim = d_aniv + relativedelta(days=90)" in MULTI


# ---------------------------------------------------------------------------
# Ciclo unico (pagina 01)
# ---------------------------------------------------------------------------

def test_ciclo_unico_usa_o_novo_padrao():
    assert "Janela de Admissibilidade (90 dias): {janela_adm_str}." in SIMPLES
    assert "**Janela de Admissibilidade (90 dias):** {janela_adm_str}" in SIMPLES
    assert "Pedido realizado em {dt_solic.strftime('%d/%m/%Y')}." in SIMPLES
    assert "Período de apuração do índice: {janela_str}." in SIMPLES
    assert "Variação apurada: {v_fmt}." in SIMPLES
    assert "Índice: {tipo_idx}." in SIMPLES
    assert "Início dos efeitos financeiros: {inicio_efeito_financeiro.strftime('%m/%Y')}." in SIMPLES


def test_ciclo_unico_ordinario_sem_redundancia_de_acumulado():
    assert "Percentual aplicado no acumulado" not in SIMPLES
    assert "Variação acumulada: {percentual_aplicado_fmt}" not in SIMPLES
    assert "Intervalo do {ciclo_label}" not in SIMPLES


def test_ciclo_unico_percentual_aplicado_somente_quando_diverge():
    # regra de apresentacao: linha extra apenas se os percentuais divergem
    assert "if percentual_aplicado_fmt != v_fmt and not preclusao_sem_efeito" in SIMPLES
    assert 'f"Percentual aplicado: {percentual_aplicado_fmt}.' in SIMPLES


def test_ciclo_unico_negocial_preserva_informacoes():
    assert "**Resultado automático:** {situacao_automatica}." in SIMPLES
    assert "**Tratamento aplicado:** (*) ciclo admitido por negociação entre as partes." in SIMPLES
    assert "Percentual aplicado por acordo: {percentual_aplicado_fmt}." in SIMPLES
    assert "Início dos efeitos financeiros por acordo: {inicio_negocial_txt}." in SIMPLES
    assert "O diagnóstico automático de preclusão foi preservado." in SIMPLES


def test_ciclo_unico_preclusao_sem_inicio_de_efeitos():
    assert '"PRECLUSO" in status_ped.upper() and not superacao_negocial' in SIMPLES
    assert '"Efeitos financeiros: não aplicáveis."' in SIMPLES


# ---------------------------------------------------------------------------
# Multiciclo (pagina 02)
# ---------------------------------------------------------------------------

def test_multiciclo_usa_o_novo_padrao():
    assert "Janela de Admissibilidade (90 dias): {h['JanelaAdm']}." in MULTI
    assert "- Janela de Admissibilidade (90 dias): {janela_adm}" in MULTI
    assert "Pedido realizado em {h['Pedido']}." in MULTI
    assert "Período de apuração do índice: {h['Janela']}." in MULTI
    assert "Variação apurada: {h['Variação']}." in MULTI
    assert "Índice: {idx_sel}." in MULTI
    assert "Início dos efeitos financeiros: {h['Início financeiro']}." in MULTI


def test_multiciclo_sem_terminologia_antiga():
    assert "Pedido em {h['Pedido']}" not in MULTI
    assert "Intervalo do C{h['Ciclo']}" not in MULTI
    assert "Resultado: {h['Situação']}. Variação: {h['Variação']}." not in MULTI


def test_multiciclo_percentual_aplicado_somente_quando_diverge():
    assert "h.get('Percentual aplicado') != h['Variação']" in MULTI
    assert "and not eh_precluso" in MULTI


def test_multiciclo_preclusao_sem_inicio_de_efeitos():
    assert '"Efeitos financeiros: não aplicáveis."' in MULTI


def test_multiciclo_negocial_preservado():
    assert "Percentual aplicado por acordo: {h.get('Percentual aplicado', h['Variação'])}." in MULTI
    assert "Início dos efeitos financeiros por acordo: {h.get('Início financeiro', '')}." in MULTI
    assert "O diagnóstico automático de preclusão foi preservado." in MULTI
