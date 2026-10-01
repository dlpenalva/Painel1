"""RESULTADOS-BASELINE (PR 0) — as grandezas economicas, em arquivos reais.

POR QUE ESTA SUITE EXISTE SEPARADA
-----------------------------------
Os 12 cenarios sinteticos de `test_baseline_resultados` nascem do openpyxl e,
portanto, SEM cache de formulas. Isso congela integralmente a logica de
fail-closed, status, ciclos, fatores e mensagens — mas nao produz VTA nem
remanescente, porque o produto (corretamente) recusa-se a inventa-los sem o
recalculo do Excel.

As grandezas ECONOMICAS so existem em arquivo recalculado. Esta suite as
congela a partir de Coletas REAIS ja recalculadas e homologadas em producao.

ONDE FICAM OS GOLDENS
---------------------
Fora do repositorio, em `CL8US_GOLDENS_DIR` (mesma convencao ja adotada por
`test_fail_closed_vta_retroativo_ausente` e `test_adequacao_orcamentaria`).
Ausentes, os testes sao PULADOS — nunca falso-verde. O que fica versionado e a
FOTOGRAFIA (JSON): os numeros entram no diff do PR e ficam auditaveis mesmo em
maquina sem os arquivos.

CONVERGENCIA NATURAL, NAO COPIA
-------------------------------
O teste central aqui nao e "web == celula da RESULTADOS". E o bloco
`reconciliacao_xls_python`, que ja existe em producao: para cada campo ele
guarda o numero do XLS e o numero que o motor Python calculou POR CONTA
PROPRIA, com tolerancia declarada. Conferir esse bloco prova as tres coisas que
o PR 0 precisa provar ao mesmo tempo: XLS correto, web correta, e as duas
cadeias chegando la sozinhas.
"""
from __future__ import annotations

import json
import os
from pathlib import Path

import pytest

from _baseline_fotografia import fotografar_cenario
from _leitor_masterfile_v10 import ler_masterfile_v10
from _templates_documentos import formatar_moeda

DIRETORIO_SNAPSHOTS = Path(__file__).resolve().parent / "baseline_resultados" / "goldens"
REGRAVAR = os.environ.get("RESULTADOS_BASELINE_REGRAVAR") == "1"

DIR_GOLDENS = Path(
    os.environ.get("CL8US_GOLDENS_DIR", r"C:\Users\danie\Downloads\anthropic-skills")
)
DIR_GOLDENS_PC = Path(
    os.environ.get("CL8US_GOLDENS_PC_DIR", r"C:\Users\danie\Downloads")
)

# Cada golden cobre um dos cenarios obrigatorios COM valores economicos reais.
GOLDENS = {
    # Cenarios 1 e 4: Financeiro multiciclo (C1+C2+C3), homologado 26/08/2026.
    "financeiro_multiciclo_validado":
        DIR_GOLDENS / "Coleta_Reajuste_C1_C2_C3_ICTI_26-08-2026.xlsx",
    # Mesma apuracao publicada um dia antes: prova que o VTA nao oscila.
    "financeiro_ciclo_vigente_validado":
        DIR_GOLDENS / "Coleta_Reajuste_C3_ICTI_25agosto2026.xlsx",
    # Cenario 2: metodo PC com retroativo apurado e status pedindo revisao.
    "pc_com_retroativo_revise": DIR_GOLDENS_PC / "Coleta_SMOKE_PR78_79.xlsx",
}

# Valores GRAVADOS NO XLS destes goldens, recalculados pelo Excel com a regra de
# precisao ANTERIOR (C3 com 2,8899...% bruto). Sao o "lado XLS" da conferencia.
# Ficam explicitos no codigo — nao so dentro do JSON — para que qualquer
# alteracao apareca como mudanca de INTENCAO no diff.
VTA_FINANCEIRO_HOMOLOGADO = 8_713_820.26
RETROATIVO_FINANCEIRO_HOMOLOGADO = 24_678.92
EXECUCAO_FINANCEIRA_HOMOLOGADA = 7_300_890.27
REMANESCENTE_FINANCEIRO_HOMOLOGADO = 1_388_251.07
VALOR_RECOMENDADO_METODOLOGIA = 8_689_141.34
VALOR_RECOMENDADO_ANTERIOR = 14_626_459.46

# Compatibilidade retroativa (Etapa 2/03): a Coleta antiga e adaptada em memoria
# e os derivados sao recompostos pela regra VIGENTE (C3 = 2,89%). O lado Python
# passa a publicar estes valores; a diferenca para o XLS (centavos, apenas por
# causa da precisao do percentual) fica exposta como divergencia e mantem a
# formalizacao bloqueada — nunca e escondida nem adotada automaticamente.
VTA_FINANCEIRO_REGRA_VIGENTE = 8_713_821.69
RETROATIVO_FINANCEIRO_REGRA_VIGENTE = 24_679.47
REMANESCENTE_FINANCEIRO_REGRA_VIGENTE = 1_388_251.95
VALOR_RECOMENDADO_REGRA_VIGENTE = 8_689_142.22
CAMPOS_DIVERGENTES_POR_PRECISAO = {
    "RETRO_FIN", "RETRO_OFICIAL", "REM_ATUALIZADO_OFICIAL", "VTA_FINAL",
}


def _caminho(nome: str) -> Path:
    return DIRETORIO_SNAPSHOTS / f"{nome}.json"


def _fotografar(nome: str) -> dict:
    arquivo = GOLDENS[nome]
    if not arquivo.exists():
        pytest.skip(f"golden externo ausente: {arquivo}")
    return fotografar_cenario(arquivo.read_bytes())


@pytest.mark.parametrize("nome", sorted(GOLDENS))
def test_golden_mantem_a_fotografia_registrada(nome):
    """As grandezas economicas do arquivo real continuam as mesmas."""
    from test_baseline_resultados import _diferencas, _gravar

    fotografia = _fotografar(nome)
    caminho = _caminho(nome)
    if REGRAVAR or not caminho.exists():
        _gravar(caminho, fotografia)
        if not REGRAVAR:
            pytest.skip(f"baseline do golden {nome} criado agora; reexecute")
        return
    esperado = json.loads(caminho.read_text(encoding="utf-8"))
    diferencas = _diferencas(esperado, fotografia)
    assert not diferencas, (
        f"O golden {nome} mudou de comportamento em {len(diferencas)} ponto(s).\n  "
        + "\n  ".join(diferencas[:40])
    )


def test_vta_e_retroativo_seguem_a_regra_vigente_com_o_xls_antigo_visivel():
    """O Python publica a regra vigente; o XLS antigo segue visivel e divergente.

    A diferenca para os valores gravados no XLS (R$ 1,43 no VTA) decorre SO da
    precisao do percentual de C3 (2,8899...% x 2,89%) e nunca e adotada sem
    equalizacao: a formalizacao permanece bloqueada.
    """
    web = _fotografar("financeiro_multiciclo_validado")["web"]
    assert web["vta_oficial"] == VTA_FINANCEIRO_REGRA_VIGENTE
    assert web["retroativo_total"] == RETROATIVO_FINANCEIRO_REGRA_VIGENTE
    assert web["vta_origem"] == "vta_canonico"
    assert web["status_apuracao"]["codigo"] == "VALIDADO"
    assert web["status_apuracao"]["status_politica"] == "BLOQUEADO_PARA_FORMALIZACAO"
    # o efeito da precisao e minusculo e esta explicado: < 2 reais sobre 8,7 mi
    assert 0 < web["vta_oficial"] - VTA_FINANCEIRO_HOMOLOGADO < 2.0
    assert 0 < web["retroativo_total"] - RETROATIVO_FINANCEIRO_HOMOLOGADO < 2.0


def test_as_duas_publicacoes_da_mesma_apuracao_convergem():
    """Dois arquivos, dois dias, o mesmo VTA — a apuracao nao oscila."""
    um = _fotografar("financeiro_multiciclo_validado")["web"]
    outro = _fotografar("financeiro_ciclo_vigente_validado")["web"]
    assert um["vta_oficial"] == outro["vta_oficial"] == VTA_FINANCEIRO_REGRA_VIGENTE
    assert um["retroativo_total"] == outro["retroativo_total"]
    assert um["remanescente_atualizado"] == outro["remanescente_atualizado"]


def test_xls_e_python_divergem_so_pela_precisao_do_percentual():
    """Cada cadeia calculou o seu numero; a diferenca e so a precisao de C3.

    Nao se compara a web contra uma celula da RESULTADOS: compara-se, para cada
    grandeza, o que o XLS calculou (regra anterior) e o que o motor Python
    calculou por conta propria (regra vigente, apos a adaptacao em memoria).
    As grandezas que nao dependem do fator continuam CONCILIADAS ao centavo; as
    quatro que dependem dele divergem por menos de R$ 2 — exatamente o
    deslocamento de 2,8899...% para 2,89% — e sao expostas, nunca escondidas.
    """
    convergencia = _fotografar("financeiro_multiciclo_validado")["web"][
        "convergencia_xls_python"
    ]
    assert convergencia["disponivel"] is True
    assert convergencia["sem_cache"] is False
    assert convergencia["status_geral"] == "DIVERGENCIA_RELEVANTE"
    divergentes = {c["campo"] for c in convergencia["campos"]
                   if c.get("status") == "DIVERGENCIA_RELEVANTE"}
    assert divergentes == CAMPOS_DIVERGENTES_POR_PRECISAO
    for campo in convergencia["campos"]:
        if campo["campo"] in CAMPOS_DIVERGENTES_POR_PRECISAO:
            assert 0 < campo["python"] - campo["xls"] < 2.0, campo
    conciliados = {c["campo"] for c in convergencia["campos"]
                   if c.get("status") == "CONCILIADO"}
    assert conciliados >= {"QTD_REM_OFICIAL", "REM_BASE_OFICIAL"}, (
        "as grandezas independentes do fator perderam a conciliacao"
    )
    for campo in convergencia["campos"]:
        if campo.get("status") == "CONCILIADO":
            assert campo["xls"] == campo["python"]


def test_composicao_do_vta_fecha_com_o_vta_oficial():
    """Executado + ajustes devidos + remanescente = VTA, ao centavo."""
    web = _fotografar("financeiro_multiciclo_validado")["web"]
    composicao = web["composicao_vta"]
    assert composicao["disponivel"] and composicao["conciliada"]
    soma = sum(linha["valor_atualizado"] for linha in composicao["linhas"])
    assert round(soma, 2) == web["vta_oficial"], (
        f"as parcelas somam {soma:.2f} e o VTA oficial e {web['vta_oficial']:.2f}"
    )


def test_valor_recomendado_financeiro_mais_remanescente_no_golden():
    leitura = ler_masterfile_v10(
        GOLDENS["financeiro_multiciclo_validado"].read_bytes(),
        exigir_modelo_oficial=True,
    )
    objeto = leitura["objeto_processo"]
    assistente = objeto["consumidores"]["assistente_operacional"]
    dossie = objeto["consumidores"]["dossie_decisao"]
    motor = assistente["motor_metodologias"]
    evidencias = motor["evidencias"]
    resultado = assistente["resultado_recomendado"]

    eventos_financeiros = [
        evento for evento in leitura["event_log_sombra"]["eventos"]
        if evento.get("fonte_parcela") == "Financeiro"
    ]
    assert eventos_financeiros
    assert {evento["ciclo"] for evento in eventos_financeiros} <= {
        "C0", "C1", "C2", "C3", "C4",
    }
    assert not any(evento.get("linha") == 74 for evento in eventos_financeiros)
    assert evidencias["financeiro"]["pago"] == EXECUCAO_FINANCEIRA_HOMOLOGADA
    assert evidencias["financeiro"]["reconhecido"] == RETROATIVO_FINANCEIRO_REGRA_VIGENTE
    assert evidencias["remanescentes"]["valor"] == REMANESCENTE_FINANCEIRO_REGRA_VIGENTE
    assert evidencias["remanescentes"]["fonte"] == (
        "composicao_vta.saldo_remanescente.valor_atualizado"
    )
    assert resultado["metodologia"] == "Financeiro + Remanescentes"
    assert resultado["valor_recomendado"] == VALOR_RECOMENDADO_REGRA_VIGENTE
    assert resultado["valor_recomendado"] != VALOR_RECOMENDADO_ANTERIOR
    assert resultado["vta"] != resultado["valor_recomendado"]

    resumo = dossie["resumo_executivo"]
    assert "Valor recomendado pela metodologia: R$ 8.689.142,22" in resumo
    assert "14.626.459,46" not in resumo


def test_valor_incorreto_desaparece_do_sumario_executivo():
    fotografia = _fotografar("financeiro_multiciclo_validado")
    resumo = fotografia["documentos"]["sumario_executivo"]["sintese"][
        "resumo_executivo"
    ]
    assert "Valor recomendado pela metodologia: R$ 8.689.142,22" in resumo
    assert "14.626.459,46" not in resumo


def test_referencias_auditaveis_nao_sao_o_vta_oficial():
    """As tres referencias da Tabela 1 sao auditoria — jamais o VTA."""
    web = _fotografar("financeiro_multiciclo_validado")["web"]
    referencias = web["referencias_vta_xls"]
    assert referencias["forma1_posicao_atual"] != web["vta_oficial"]
    assert referencias["forma2_ultima_abertura"] != web["vta_oficial"]
    assert referencias["forma3_integral_reajustado"] != web["vta_oficial"]
    assert referencias["reconciliacao_status"] == "RECONCILIADO"


def test_documentos_usam_o_valor_canonico_e_nunca_o_xls_antigo():
    """Com XLS x Python divergentes, nenhum documento publica o valor do XLS.

    O Sumario retem VTA e retroativo divergentes; o Termo de Apostila carrega o
    retroativo canonico (regra vigente), jamais o gravado no XLS antigo. A
    geracao destes documentos segue bloqueada pela politica de entrega (teste
    `test_p1_golden_financeiro_real_antigo_sem_documentos_formalizadores`).
    """
    fotografia = _fotografar("financeiro_multiciclo_validado")
    web, documentos = fotografia["web"], fotografia["documentos"]
    sintese = documentos["sumario_executivo"]["sintese"]
    assert sintese["vta"] is None
    assert sintese["retroativo_total"] is None
    assert sintese["vta_saldo_remanescente_atualizado"] == web["remanescente_atualizado"]

    corpo = " ".join(documentos["termo_apostila"]["linhas_negociais"])
    assert "24.678,92" not in corpo, "o valor do XLS antigo vazou para o Termo de Apostila"
    assert "24.679,47" in corpo, "o retroativo canonico sumiu do Termo de Apostila"


def test_metodo_pc_apura_retroativo_no_arquivo_real():
    """No golden de PC ha retroativo apurado pela cadeia de pedidos."""
    web = _fotografar("pc_com_retroativo_revise")["web"]
    assert web["metodo"]["codigo"] == "pc"
    assert web["retroativo_total"] is not None and web["retroativo_total"] > 0


def test_saneador_pc_golden_exibe_vta_canonico_uma_vez():
    """NOVO-01: o detalhamento PC nunca pode ocultar ou duplicar o VTA."""
    fotografia = _fotografar("pc_com_retroativo_revise")
    vta = fotografia["web"]["vta_oficial"]
    linhas = fotografia["documentos"]["despacho_saneador"]["linhas_negociais"]
    corpo = "\n".join(linhas)
    assert vta == 17_364.98  # regra vigente: execucao 15.585,60 + potencial 1.779,38
    assert corpo.count("Valor Total Atualizado do Contrato") == 1
    assert corpo.count(formatar_moeda(vta)) == 1
