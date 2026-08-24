from copy import deepcopy
from pathlib import Path

from _resultado_consolidado import montar_resultado_consolidado


ROOT = Path(__file__).resolve().parents[1]
PAGINA = (ROOT / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")
RUNTIME = (ROOT / "_coleta_reajuste_documentos.py").read_text(encoding="utf-8")


def _resultado(*, posicao_atual_valida, vta_atual, vta_ultima):
    return {
        "valor_atualizado_contrato": vta_atual,
        "valor_represado_a_pagar": 125.0,
        "controle": {"modo": "pc", "ciclo_vigente": "C2"},
        "memoria_por_ciclo": {"vta": {"metodo": "pc"}},
        "referencias_vta": {
            "posicao_atual_disponivel": posicao_atual_valida,
            "forma1_posicao_atual": vta_atual,
            "forma2_ultima_abertura": vta_ultima,
        },
        "totais_canonicos_pc": {
            "ate_o_corte": {
                "retroativo": 125.0,
                "valor_atualizado_em_analise": 50.0,
                "delta_potencial": 25.0,
            },
            "posterior_ao_corte": {"quantidade": 0, "valor_pc": 0.0},
        },
        "composicao_vta": {
            "disponivel": True,
            "bloqueia_formalizacao": False,
            "linhas": [{"descricao": "Parcela canônica", "valor_atualizado": 1.0}],
            "alertas": [],
        },
        "politica_entrega_segura": {
            "status": "PRONTO_PARA_VALIDACAO_FISCAL",
            "pendencias": [],
            "retroativo": {"metodo": "pc"},
        },
        "reconciliacao_xls_python": {"status_geral": "CONCILIADO"},
    }


def test_caso_a_posicao_atual_valida_tem_precedencia_exclusiva():
    resultado = _resultado(
        posicao_atual_valida=True,
        vta_atual=1_000.0,
        vta_ultima=900.0,
    )
    consolidado = montar_resultado_consolidado(resultado)

    assert consolidado["vta"] == 1_000.0
    assert consolidado["vta_origem"] == "posicao_atual"
    assert consolidado["vta_usa_ultima_posicao"] is False


def test_caso_b_sem_posicao_atual_nao_adota_a_referencia_fisica():
    """VTA-U2 achado A / CASO 3: sem VTA canonico, a referencia fisica da
    ultima abertura NAO vira VTA — o resultado fica indisponivel."""
    resultado = _resultado(
        posicao_atual_valida=False,
        vta_atual=None,
        vta_ultima=137_235_360.64,
    )
    antes = deepcopy(resultado)
    consolidado = montar_resultado_consolidado(resultado)

    assert consolidado["vta"] is None
    assert consolidado["vta_origem"] == "indisponivel"
    assert consolidado["vta_usa_ultima_posicao"] is False
    assert resultado == antes
    assert "#9A5B00" in PAGINA
    assert "resultado-valor-ultima-posicao" in PAGINA
    assert (
        "As referências auditáveis não substituem o VTA." in PAGINA
    )


def test_caso_b2_posicao_atual_ausente_mas_vta_canonico_presente():
    """VTA-U2 achado A / CASO 2: com VTA canonico calculado, a ausencia da
    posicao fisica atual nao desloca o VTA para a ultima abertura."""
    consolidado = montar_resultado_consolidado(
        _resultado(posicao_atual_valida=False, vta_atual=1_000.0, vta_ultima=900.0)
    )

    assert consolidado["vta"] == 1_000.0
    assert consolidado["vta_origem"] == "posicao_atual"
    assert consolidado["vta_usa_ultima_posicao"] is False


def test_caso_b3_referencias_fisicas_seguem_expostas_como_referencia():
    """VTA-U2 achado A / CASO 4: B10/B11 continuam disponiveis, rotuladas
    como referencia auditavel, sem nunca serem chamadas de VTA."""
    consolidado = montar_resultado_consolidado(
        _resultado(posicao_atual_valida=False, vta_atual=None, vta_ultima=900.0)
    )
    ref = consolidado["referencias_auditaveis"]

    assert ref["ultima_abertura_disponivel"] == 900.0
    assert ref["rotulos"]["posicao_fisica_atual"].startswith("Referência auditável")
    assert ref["rotulos"]["ultima_abertura_disponivel"].startswith(
        "Referência auditável"
    )
    for rotulo in ref["rotulos"].values():
        assert "VTA" not in rotulo


def test_caso_c_sem_nenhuma_base_valida_mantem_vta_indisponivel():
    consolidado = montar_resultado_consolidado(
        _resultado(posicao_atual_valida=False, vta_atual=None, vta_ultima=None)
    )

    assert consolidado["vta"] is None
    assert consolidado["vta_origem"] == "indisponivel"
    assert consolidado["vta_usa_ultima_posicao"] is False


def test_caso_d_pagina_comeca_no_uploader_sem_download_redundante():
    for texto in (
        "Baixar arquivo de trabalho",
        "Use o modelo oficial com fórmulas para registrar a apuração contratual.",
        "Baixar Arquivo Coleta Oficial",
    ):
        assert texto not in PAGINA
    assert "Enviar Arquivo Coleta Oficial preenchido" in PAGINA
    assert 'key="upload_coleta_documentos"' in PAGINA
    assert "processar_coleta_oficial_runtime" in PAGINA


def test_caso_e_invariancia_das_grandezas_e_da_fonte_canonica():
    atual = _resultado(posicao_atual_valida=True, vta_atual=1_000.0, vta_ultima=900.0)
    ultima = _resultado(posicao_atual_valida=False, vta_atual=None, vta_ultima=900.0)

    consolidado_atual = montar_resultado_consolidado(atual)
    consolidado_ultima = montar_resultado_consolidado(ultima)

    assert consolidado_atual["vta"] == atual["valor_atualizado_contrato"]
    # VTA-U2: sem VTA canonico o resultado e indisponivel, nunca a referencia.
    assert consolidado_ultima["vta"] is None
    assert (
        consolidado_ultima["referencias_auditaveis"]["ultima_abertura_disponivel"]
        == ultima["referencias_vta"]["forma2_ultima_abertura"]
    )
    for consolidado in (consolidado_atual, consolidado_ultima):
        assert consolidado["retroativo_reconhecido"] == 125.0
        assert consolidado["retroativo_potencial"] == 25.0
    assert '"referencias_vta": leitura.get("referencias_vta") or {}' in RUNTIME
    assert "forma2_ultima_abertura" not in RUNTIME
