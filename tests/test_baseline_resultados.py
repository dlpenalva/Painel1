"""RESULTADOS-BASELINE (PR 0) — a rede de seguranca da refatoracao.

O QUE ESTES TESTES PROVAM
-------------------------
Que o comportamento de HOJE — antes de qualquer mudanca na aba RESULTADOS —
esta congelado e e reproduzivel:

  * o CONTRATO ESTRUTURAL da aba (formulas, ancoras nomeadas, visibilidade);
  * as GRANDEZAS da cadeia web nos 12 cenarios obrigatorios;
  * o CONTEUDO NEGOCIAL dos documentos que consomem o resultado consolidado;
  * a CONVERGENCIA natural entre a cadeia do XLS e a cadeia Python.

Eles NAO afirmam que os numeros estao certos. Afirmam que sao os MESMOS de
antes. Um vermelho aqui significa "a refatoracao mexeu em algo que deveria ter
ficado parado" — ou, se a mudanca for intencional e aprovada, que o snapshot
precisa ser regravado explicitamente.

COMO REGRAVAR (deliberado, nunca automatico)
--------------------------------------------
    set RESULTADOS_BASELINE_REGRAVAR=1
    python -m pytest tests/test_baseline_resultados.py

A variavel existe para que regravar seja um ATO, visivel no diff do PR. Sem
ela, o teste so compara. Em nenhuma hipotese se regrava para "fazer passar".

CUSTO
-----
Cada cenario carrega o template oficial (~16 MB de XML), processa o upload
completo e gera cinco documentos: ~20-25 s por cenario, ~5 min para os 12. E
uma suite CARA, deliberadamente fora do CI rapido (que roda so os sentinelas).
Rode-a por inteiro antes de abrir o PR da refatoracao; no dia a dia, rode com
`-k` o cenario que voce mexeu. Nenhum marker novo foi introduzido: o repo nao
tem `pytest.ini`, e registrar um marker aqui mudaria a configuracao de toda a
suite — decisao que nao cabe a este PR.
"""
from __future__ import annotations

import json
import os
from pathlib import Path

import pytest

from _baseline_cenarios import DESCRICOES, ORDEM_CENARIOS, bytes_cenario
from _baseline_fotografia import (
    COORDENADAS_LIDAS_EM_RUNTIME,
    fotografar_cenario,
    fotografar_contrato_xls,
)

DIRETORIO_SNAPSHOTS = Path(__file__).resolve().parent / "baseline_resultados"
REGRAVAR = os.environ.get("RESULTADOS_BASELINE_REGRAVAR") == "1"

# CHECKPOINT DE RETORNO desta frente. Se qualquer etapa futura regredir, este
# e o estado para o qual voltar. NAO alterar.
CHECKPOINT_PRE_RESULTADOS = "f8296f7c2962352716edd22044ed9573f5eeee8a"


def _caminho(cenario: str) -> Path:
    return DIRETORIO_SNAPSHOTS / f"{cenario}.json"


def _gravar(caminho: Path, fotografia: dict) -> None:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    caminho.write_text(
        json.dumps(fotografia, ensure_ascii=False, indent=1, sort_keys=True) + "\n",
        encoding="utf-8",
    )


def _diferencas(esperado, obtido, caminho: str = "") -> list[str]:
    """Diferencas com endereco legivel, para o vermelho apontar o culpado."""
    if isinstance(esperado, dict) and isinstance(obtido, dict):
        saida: list[str] = []
        for chave in sorted(set(esperado) | set(obtido)):
            local = f"{caminho}.{chave}" if caminho else str(chave)
            if chave not in esperado:
                saida.append(f"{local}: SURGIU {obtido[chave]!r}")
            elif chave not in obtido:
                saida.append(f"{local}: SUMIU {esperado[chave]!r}")
            else:
                saida.extend(_diferencas(esperado[chave], obtido[chave], local))
        return saida
    if isinstance(esperado, list) and isinstance(obtido, list):
        if len(esperado) != len(obtido):
            return [f"{caminho}: {len(esperado)} itens -> {len(obtido)} itens"]
        saida = []
        for indice, (um, outro) in enumerate(zip(esperado, obtido)):
            saida.extend(_diferencas(um, outro, f"{caminho}[{indice}]"))
        return saida
    if esperado != obtido:
        return [f"{caminho}: {esperado!r} -> {obtido!r}"]
    return []


# --------------------------------------------------------------------------- #
# A — o baseline dos 12 cenarios.
# --------------------------------------------------------------------------- #
@pytest.mark.parametrize("cenario", ORDEM_CENARIOS)
def test_cenario_mantem_a_fotografia_registrada(cenario):
    """Toda grandeza do cenario continua identica a fotografia do PR 0."""
    fotografia = {
        "cenario": cenario,
        "descricao": DESCRICOES[cenario],
        "checkpoint_pre_resultados": CHECKPOINT_PRE_RESULTADOS,
        **fotografar_cenario(bytes_cenario(cenario)),
    }
    caminho = _caminho(cenario)
    if REGRAVAR or not caminho.exists():
        _gravar(caminho, fotografia)
        if not REGRAVAR:
            pytest.skip(f"baseline de {cenario} criado agora; reexecute para comparar")
        return
    esperado = json.loads(caminho.read_text(encoding="utf-8"))
    diferencas = _diferencas(esperado, fotografia)
    assert not diferencas, (
        f"O cenario {cenario} mudou de comportamento em {len(diferencas)} ponto(s). "
        "Se a mudanca for intencional e aprovada, regrave com "
        "RESULTADOS_BASELINE_REGRAVAR=1 e justifique cada linha no PR.\n  "
        + "\n  ".join(diferencas[:40])
    )


# --------------------------------------------------------------------------- #
# B — o contrato de leitura da aba RESULTADOS.
# --------------------------------------------------------------------------- #
def test_coordenadas_lidas_em_runtime_continuam_existindo():
    """As nove celulas que o runtime le nao podem sumir da aba.

    Este e o contrato MINIMO que a refatoracao do PR 1 precisa preservar: fora
    destas coordenadas (e dos intervalos nomeados), a aba RESULTADOS e
    apresentacao. Se o PR 1 mover alguma delas, tem de mover tambem o leitor —
    e este teste e o lembrete.
    """
    contrato = fotografar_contrato_xls(bytes_cenario("01_financeiro_normal"))
    assert contrato["aba_presente"]
    assert contrato["visibilidade"] == "visible"
    assert contrato["e_ultima_aba"], "a aba RESULTADOS deve ser a ultima do arquivo"
    ausentes = [
        coordenada for coordenada in COORDENADAS_LIDAS_EM_RUNTIME
        if coordenada not in contrato["formulas"]
    ]
    assert not ausentes, (
        "coordenadas lidas pelo runtime ficaram vazias na aba RESULTADOS: "
        + ", ".join(ausentes)
    )


def test_ancoras_nomeadas_da_aba_resultados_estao_estaveis():
    """Os intervalos nomeados ancorados na aba sao API, nao decoracao."""
    contrato = fotografar_contrato_xls(bytes_cenario("01_financeiro_normal"))
    assert contrato["nomes_ancorados_na_aba"] == [
        "AJUSTES_DEVIDOS",
        "AUDITORIA_ABERTURA_STATUS",
        "AUDITORIA_SITUACAO_ATUAL_STATUS",
        "CONFERENCIA_FORMACAO_VTA",
        "EXECUCAO_ATUALIZADA_CICLO",
        "EXECUTADO_APURADO",
        "OPCOES_APLICAR_MANUAL",
        "SALDO_REMANESCENTE_ATUAL",
        "STATUS_RESULTADOS",
        "VTA_ATUALIZACAO_CHEIA",
    ]
    nomes = contrato["nomes_definidos"]
    assert nomes["STATUS_RESULTADOS"] == "RESULTADOS!$B$3"
    assert nomes["VTA_FINAL"] == "MEMORIA_RESULTADOS!$B$26"
    assert nomes["RETRO_OFICIAL"] == "MEMORIA_RESULTADOS!$B$16"


def test_as_entradas_de_ajuste_manual_sao_contrato_dentro_do_excel():
    """As linhas 43-50 sao lidas por formulas de OUTRAS abas do workbook.

    Este acoplamento e invisivel para qualquer teste de Python: quem consome
    RESULTADOS!C43:G50 e a MEMORIA_RESULTADOS, dentro do proprio Excel, para
    compor VTA e retroativo. Renumerar ou mover o bloco "5. AJUSTES MANUAIS"
    quebraria o calculo sem produzir um unico vermelho na suite — por isso o
    endereco entra no contrato, ao lado das nove coordenadas do runtime.
    """
    contrato = fotografar_contrato_xls(bytes_cenario("01_financeiro_normal"))
    referencias = contrato["coordenadas_lidas_por_outras_abas"]
    esperadas = {
        f"{coluna}{linha}"
        for linha in range(43, 51)
        for coluna in ("C", "D", "G")
        # D so existe nas tres primeiras linhas (retroativo manual, ajuste do
        # VTA e VTA substitutivo); os complementos historicos nao a usam.
        if not (coluna == "D" and linha >= 46)
    }
    # Contrato vigente em origin/main: comparativo_VTA também consome o fator
    # histórico exibido em RESULTADOS!H5. PC-UX-1 preserva essa referência.
    esperadas.add("H5")
    assert set(referencias) == esperadas, (
        "mudou o conjunto de celulas da RESULTADOS consumidas por formulas de "
        "outras abas"
    )
    assert referencias["H5"] == ["comparativo_VTA"]
    assert all(
        abas == ["MEMORIA_RESULTADOS"]
        for coordenada, abas in referencias.items()
        if coordenada != "H5"
    )


def test_o_titulo_da_aba_e_o_gate_de_integridade():
    """`_coleta_reajuste` rejeita o arquivo se A1 nao for exatamente isto."""
    contrato = fotografar_contrato_xls(bytes_cenario("01_financeiro_normal"))
    assert contrato["titulo_a1"] == "RESULTADOS CONSOLIDADOS — REAJUSTE CONTRATUAL"


# --------------------------------------------------------------------------- #
# C — o que cada cenario existe para exercitar continua sendo exercitado.
# --------------------------------------------------------------------------- #
def _web(cenario: str) -> dict:
    return fotografar_cenario(bytes_cenario(cenario), com_documentos=False)["web"]


def test_cada_metodo_chega_ao_consolidado_com_o_proprio_codigo():
    """Financeiro, PC e Consumido nao podem se confundir no consolidado."""
    assert _web("01_financeiro_normal")["metodo"]["codigo"] == "financeiro"
    assert _web("02_pc")["metodo"]["codigo"] == "pc"
    assert _web("03_itens_consumidos")["metodo"]["codigo"] == "consumidos"


def test_multiciclo_computa_tres_ciclos():
    web = _web("04_multiciclo")
    assert web["ciclos_considerados"] == ["C1", "C2", "C3"]
    assert web["ciclo_vigente"] == "C3"


def test_reajuste_negativo_aplicado_e_neutralizado_divergem_no_fator():
    """A decisao sobre a variacao negativa tem de chegar ao fator do ciclo."""
    aplicado = _web("05_reajuste_negativo_aplicado")["ciclos"][1]
    neutralizado = _web("06_reajuste_negativo_neutralizado")["ciclos"][1]
    assert aplicado["percentual_efetivamente_aplicado"] != \
        neutralizado["percentual_efetivamente_aplicado"], (
        "aplicar e neutralizar a variacao negativa nao podem produzir o mesmo "
        "percentual efetivo"
    )
    assert neutralizado["percentual_efetivamente_aplicado"] == 0.0


def test_ausencia_de_ciclo_em_execucao_nao_vira_zero():
    """Fail-closed: sem a fotografia fisica, o VTA e INDISPONIVEL, nunca 0,00."""
    web = _web("07_sem_ciclo_em_execucao")
    assert web["vta_oficial"] is None
    assert web["vta_origem"] == "indisponivel"


def test_arquivo_sem_recalculo_declara_o_cache_ausente():
    """O produto pede recalculo em vez de inventar numero — e o que se congela."""
    fotografia = fotografar_cenario(
        bytes_cenario("10_sem_recalculo_do_excel"), com_documentos=False
    )
    assert fotografia["valores_xls"]["cache_ausente"] is True
    assert fotografia["web"]["vta_oficial"] is None
    assert fotografia["web"]["status_apuracao"]["disponivel"] is False


def test_pcs_com_e_sem_efeito_financeiro_produzem_retroativos_distintos():
    """PC sem efeito financeiro nao gera retroativo reconhecido."""
    com_efeito = _web("02_pc")["retroativo_total"]
    sem_efeito = _web("11_pcs_sem_efeito_financeiro")["retroativo_total"]
    assert com_efeito is not None and com_efeito > 0
    assert sem_efeito == 0.0, (
        "PCs anteriores ao inicio do efeito financeiro nao podem gerar "
        f"retroativo reconhecido (obtido: {sem_efeito})"
    )


# --------------------------------------------------------------------------- #
# D — a convergencia e natural, nunca copia.
# --------------------------------------------------------------------------- #
def test_convergencia_declara_a_ausencia_de_cache_em_vez_de_conciliar_vazio():
    """Sem cache, o bloco de convergencia diz `sem_cache`, nao CONCILIADO."""
    convergencia = _web("01_financeiro_normal")["convergencia_xls_python"]
    assert convergencia["sem_cache"] is True
    assert convergencia["status_geral"] != "CONCILIADO"


def test_o_bloco_de_convergencia_preserva_as_duas_colunas():
    """Cada campo guarda `xls` E `python` — a igualdade tem de ser observavel.

    Se um dia a web deixar de calcular e passar a copiar o XLS, as duas colunas
    continuariam iguais; o que denuncia a troca e a existencia das duas colunas
    e da origem, e por isso o baseline as fotografa em vez de so comparar.
    """
    convergencia = _web("01_financeiro_normal")["convergencia_xls_python"]
    campos = convergencia["campos"] or []
    assert campos, "o bloco de convergencia nao pode chegar vazio"
    for campo in campos:
        assert "xls" in campo and "python" in campo, (
            f"campo {campo.get('campo')} perdeu uma das colunas da conferencia"
        )
