"""P0-ROBUSTEZ-VALORES-1: ausencia de VTA/retroativo nunca vira zero.

Contrato fixado aqui (fail-closed, alinhado ao VTA-U2):

* AUSENTE / NAO CALCULAVEL / SEM CACHE  ->  ``None``  ->  UI "INDISPONIVEL";
* ZERO ECONOMICO REALMENTE APURADO      ->  ``0.0``   ->  UI "R$ 0,00".

O defeito corrigido: ``_numero`` tem padrao ``0.0``, entao
``_numero(capacidades["calculos"]["vta"]["valor"])`` devolvia ``0.0`` quando a
capacidade estava indisponivel (``valor=None``). Esse zero fabricado descia por
``valor_atualizado_contrato``/``valor_represado_a_pagar`` ate
``montar_resultado_consolidado``, que ainda o carimbava como ``vta_canonico``.
Os dois callsites passaram a usar ``_numero_opcional``, que preserva a ausencia.

Cobertura complementar (nao duplicada): as demais suites de ``vta_origem``
partem de payloads em que o VTA ja chega ausente — VTA-U2 (fim do fallback pela
posicao fisica), VTA-C2 (fail-closed do Consumido) e ultima posicao de abertura.
Aqui a ausencia nasce de um XLSX real sem cache de formulas, atravessando o
adaptador que continha o defeito.
"""

from __future__ import annotations

import io
import os
from datetime import date
from pathlib import Path

import pytest

from _coleta_oficial import gerar_coleta_oficial_preenchida
from _coleta_reajuste_documentos import (
    _numero,
    _numero_opcional,
    processar_coleta_oficial_runtime,
)
from _resultado_consolidado import (
    ORIGEM_VTA_CANONICA,
    ORIGEM_VTA_INDISPONIVEL,
    montar_resultado_consolidado,
)


ROOT = Path(__file__).resolve().parents[1]

# Goldens reais homologados em producao (26/08/2026, app em 458dbfb). Ficam
# fora do repositorio: quando ausentes o TESTE D e pulado, nunca falso-verde.
DIR_GOLDENS = Path(
    os.environ.get("CL8US_GOLDENS_DIR", r"C:\Users\danie\Downloads\anthropic-skills")
)
GOLDEN_25 = DIR_GOLDENS / "Coleta_Reajuste_C3_ICTI_25agosto2026.xlsx"
GOLDEN_26 = DIR_GOLDENS / "Coleta_Reajuste_C1_C2_C3_ICTI_26-08-2026.xlsx"

VTA_ESPERADO = 8_713_820.26
RETROATIVO_ESPERADO = 24_678.92


def _coleta_sem_cache_de_formulas() -> bytes:
    """Coleta gerada por openpyxl: toda formula fica sem valor gravado.

    E o estado exato do arquivo que o fiscal reenvia sem ter aberto no Excel —
    nao ha VTA canonico nem valor apurado nas capacidades.
    """
    conteudo = gerar_coleta_oficial_preenchida({
        "origem": "P0-ROBUSTEZ-VALORES-1 — ausencia sem cache",
        "indice": "IST (Anatel)",
        "data_base_original": "01/02/2025",
        "data_corte": date(2027, 1, 31),
        "ciclos": [{
            "ciclo": "C1",
            "data_inicio": date(2026, 2, 1),
            "data_fim": date(2027, 1, 31),
            "data_pedido": date(2026, 2, 1),
            "financeiro_inicio": date(2026, 4, 1),
            "percentual_aplicado": 0.0307853139440224,
            "situacao": "✅ TEMPESTIVO",
            "objeto_analise_atual": True,
        }],
    })
    return conteudo.getvalue() if isinstance(conteudo, io.BytesIO) else bytes(conteudo)


@pytest.fixture(scope="module")
def apuracao_sem_calculo():
    resultado, diagnostico = processar_coleta_oficial_runtime(
        _coleta_sem_cache_de_formulas()
    )
    return resultado, diagnostico, montar_resultado_consolidado(resultado, diagnostico)


# ---------------------------------------------------------------------------
# TESTE A — ausencia real de VTA
# ---------------------------------------------------------------------------

def test_a_vta_ausente_permanece_none(apuracao_sem_calculo):
    resultado, _diagnostico, consolidado = apuracao_sem_calculo

    assert resultado["valor_atualizado_contrato"] is None
    assert consolidado["vta"] is None
    assert consolidado["vta_origem"] == ORIGEM_VTA_INDISPONIVEL


def test_a_vta_ausente_nunca_e_zero_fabricado(apuracao_sem_calculo):
    resultado, _diagnostico, consolidado = apuracao_sem_calculo

    # O ponto do defeito: 0.0 e um valor economico legitimo, entao um zero
    # fabricado aqui seria indistinguivel de "o contrato vale zero".
    assert resultado["valor_atualizado_contrato"] != 0.0
    assert consolidado["vta"] != 0.0
    assert consolidado["vta_origem"] != ORIGEM_VTA_CANONICA
    # Os espelhos do mesmo numero acompanham a ausencia.
    assert resultado["valor_calculado_sem_aditivos"] is None
    assert resultado["valor_global_financeiro"] is None


# ---------------------------------------------------------------------------
# TESTE B — ausencia real de retroativo
# ---------------------------------------------------------------------------

def test_b_retroativo_ausente_permanece_none(apuracao_sem_calculo):
    resultado, _diagnostico, consolidado = apuracao_sem_calculo

    assert resultado["valor_represado_a_pagar"] is None
    assert consolidado["retroativo_reconhecido"] is None


def test_b_retroativo_ausente_nunca_e_zero_fabricado(apuracao_sem_calculo):
    resultado, _diagnostico, consolidado = apuracao_sem_calculo

    assert resultado["valor_represado_a_pagar"] != 0.0
    assert consolidado["retroativo_reconhecido"] != 0.0
    assert resultado["delta_total"] is None
    assert resultado["delta_acumulado"] is None
    # Flag derivada nao pode explodir nem afirmar disponibilidade sobre None.
    assert resultado["retroativo_estimado_itens_estoque_disponivel"] is False


# ---------------------------------------------------------------------------
# TESTE C — zero economico legitimo continua zero
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "entrada, esperado",
    [
        (0, 0.0),          # zero apurado (int)
        (0.0, 0.0),        # zero apurado (float)
        ("0", 0.0),        # zero apurado vindo de celula textual
        (24_678.92, 24_678.92),
        (-125.5, -125.5),
        (None, None),      # ausencia
        ("", None),        # celula vazia
        ("   ", None),     # texto nao numerico
        (True, None),      # booleano nunca vira 1.0
    ],
)
def test_c_numero_opcional_separa_ausencia_de_zero(entrada, esperado):
    assert _numero_opcional(entrada) == esperado


def test_c_numero_mantem_fallback_zero_historico():
    """Os demais callsites de ``_numero`` nao podem ter mudado de semantica."""
    assert _numero(None) == 0.0
    assert _numero("") == 0.0
    assert _numero(None, 1.0) == 1.0
    assert _numero(0.0) == 0.0


def test_c_zero_apurado_sobrevive_ate_o_consolidado():
    """Zero real no payload permanece 0.0 e segue sendo VTA canonico."""
    resultado = {
        "valor_atualizado_contrato": 0.0,
        "valor_represado_a_pagar": 0.0,
        "controle": {"modo": "principal", "ciclo_vigente": "C1"},
        "memoria_por_ciclo": {"vta": {"metodo": "financeiro"}},
    }
    consolidado = montar_resultado_consolidado(resultado, {})

    assert consolidado["vta"] == 0.0
    assert consolidado["vta"] is not None
    assert consolidado["vta_origem"] == ORIGEM_VTA_CANONICA
    assert consolidado["retroativo_reconhecido"] == 0.0
    assert consolidado["retroativo_reconhecido"] is not None


def test_c_ausencia_no_payload_nao_vira_zero_no_consolidado():
    resultado = {
        "valor_atualizado_contrato": None,
        "valor_represado_a_pagar": None,
        "controle": {"modo": "principal", "ciclo_vigente": "C1"},
        "memoria_por_ciclo": {"vta": {"metodo": "financeiro"}},
    }
    consolidado = montar_resultado_consolidado(resultado, {})

    assert consolidado["vta"] is None
    assert consolidado["vta_origem"] == ORIGEM_VTA_INDISPONIVEL
    assert consolidado["retroativo_reconhecido"] is None


# ---------------------------------------------------------------------------
# TESTE D — goldens reais homologados em producao
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "caminho, ciclos_esperados",
    [
        (GOLDEN_25, ["C3"]),
        (GOLDEN_26, ["C1", "C2", "C3"]),
    ],
    ids=["golden-25-08-C3", "golden-26-08-C1-C2-C3"],
)
def test_d_goldens_reais_inalterados(caminho, ciclos_esperados):
    if not caminho.exists():
        pytest.skip(f"Golden real indisponivel neste ambiente: {caminho.name}")

    resultado, diagnostico = processar_coleta_oficial_runtime(caminho.read_bytes())
    consolidado = montar_resultado_consolidado(resultado, diagnostico)

    assert consolidado["metodo"]["rotulo"] == "Financeiro"
    assert consolidado["vta"] == pytest.approx(VTA_ESPERADO, abs=0.005)
    assert consolidado["retroativo_reconhecido"] == pytest.approx(
        RETROATIVO_ESPERADO, abs=0.005
    )
    assert consolidado["vta_origem"] == ORIGEM_VTA_CANONICA
    # STATUS-CANON-1: vocabulario do painel = vocabulario da aba RESULTADOS.
    assert consolidado["status_confiabilidade"] == "VALIDADO"
    assert consolidado["status_apuracao"]["codigo"] == "VALIDADO"
    assert consolidado["status_apuracao"]["origem"] == "resultados_xls"
    assert consolidado["formalizacao"]["bloqueada"] is False
    assert diagnostico["metadados"]["ciclos_em_analise"] == ciclos_esperados


def test_d_golden_26_preserva_execucao_dos_ciclos_preclusos():
    """PRECLUSO retira efeito financeiro, nunca a execucao ja paga."""
    if not GOLDEN_26.exists():
        pytest.skip(f"Golden real indisponivel neste ambiente: {GOLDEN_26.name}")

    resultado, _diagnostico = processar_coleta_oficial_runtime(GOLDEN_26.read_bytes())

    # C0 + C1 + C2 + C3 permanecem integralmente no executado apurado.
    assert resultado["valor_pago_efetivo"] == pytest.approx(7_300_890.27, abs=0.005)
    # Somente C3 (TEMPESTIVO) gera delta; C1/C2 preclusos entram pelo valor-base.
    assert resultado["valor_represado_a_pagar"] == pytest.approx(
        RETROATIVO_ESPERADO, abs=0.005
    )
    assert resultado["remanescente_reajustado"] == pytest.approx(
        1_388_251.07, abs=0.005
    )
    assert resultado["valor_atualizado_contrato"] == pytest.approx(
        VTA_ESPERADO, abs=0.005
    )


# ---------------------------------------------------------------------------
# Ressalva de Pedidos de Compra so descreve casos apurados POR PC
# ---------------------------------------------------------------------------

def test_ressalva_de_pc_e_condicionada_ao_metodo_pc():
    """Um caso Financeiro bloqueado nao pode ser rotulado como divergencia PC."""
    pagina = (ROOT / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")

    assert (
        'if status_em_conferencia and consolidado.get("medidas_pc_aplicaveis")'
        in pagina
    )
    # STATUS-CANON-1: a politica de bloqueio nao mudou de conteudo, mudou de
    # eixo — passou a ser lida da FORMALIZACAO, e nao do status da apuracao.
    assert 'status_em_conferencia = bool(formalizacao.get("bloqueada"))' in pagina
    assert 'status_em_conferencia = status == "BLOQUEADO"' not in pagina
    # A mensagem especifica de PC acompanha o motivo da formalizacao.
    assert "motivo_formalizacao = (" in pagina
