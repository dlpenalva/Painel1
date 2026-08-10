"""Contratos permanentes da invalidacao escopada ao trocar o XLS."""
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from _estado_apuracao_upload import (  # noqa: E402
    assinatura_conteudo_upload,
    invalidar_estado_caso,
)


def _estado_caso_a() -> dict:
    return {
        "assinatura_upload_docs": assinatura_conteudo_upload(b"conteudo-a"),
        "assinatura_processada_upload_docs": "processado-a",
        "resultado_valor_global": {"caso": "A"},
        "diagnostico_coleta_v2": {"caso": "A"},
        "dados_admissibilidade": {"contrato": "A"},
        "arquivo_sumario_executivo_pdf": b"sumario-a",
        "arquivo_despacho_saneador_docx": b"saneador-a",
        "arquivo_termo_apostila_docx": b"apostila-a",
        "arquivo_dou_docx": b"dou-a",
        "previa_dou": "texto do caso A",
        "saneador_texto": "saneador do caso A",
        "adequacao_v3_retroativo": 100.0,
        "adequacao_v3_editor_Financeiro_06/2026": {"edited_rows": {}},
        "tema_interface": "claro",
        "navegacao_global": "central",
    }


def test_a_para_b_remove_resultado_e_diagnostico_anteriores() -> None:
    estado = _estado_caso_a()

    houve_troca = invalidar_estado_caso(
        estado, assinatura_conteudo_upload(b"conteudo-b")
    )

    assert houve_troca is True
    assert "resultado_valor_global" not in estado
    assert "diagnostico_coleta_v2" not in estado
    assert "assinatura_processada_upload_docs" not in estado


def test_a_para_b_remove_documentos_textos_e_estado_dinamico_da_adequacao() -> None:
    estado = _estado_caso_a()

    invalidar_estado_caso(estado, assinatura_conteudo_upload(b"conteudo-b"))

    for chave in (
        "arquivo_sumario_executivo_pdf",
        "arquivo_despacho_saneador_docx",
        "arquivo_termo_apostila_docx",
        "arquivo_dou_docx",
        "previa_dou",
        "saneador_texto",
        "adequacao_v3_retroativo",
        "adequacao_v3_editor_Financeiro_06/2026",
    ):
        assert chave not in estado


def test_a_para_b_remove_admissibilidade_especifica_do_caso() -> None:
    estado = _estado_caso_a()

    invalidar_estado_caso(estado, assinatura_conteudo_upload(b"conteudo-b"))

    assert "dados_admissibilidade" not in estado


def test_rerun_do_mesmo_conteudo_nao_limpa_estado() -> None:
    estado = _estado_caso_a()
    estado_antes = estado.copy()

    houve_troca = invalidar_estado_caso(
        estado, assinatura_conteudo_upload(b"conteudo-a")
    )

    assert houve_troca is False
    assert estado == estado_antes


def test_mesmo_nome_com_bytes_diferentes_e_novo_caso() -> None:
    nome_a = nome_b = "Coleta_Reajuste.xlsx"
    assinatura_a = assinatura_conteudo_upload(b"versao-a")
    assinatura_b = assinatura_conteudo_upload(b"versao-b")
    estado = {
        "assinatura_upload_docs": assinatura_a,
        "resultado_valor_global": {"caso": "A"},
    }

    houve_troca = invalidar_estado_caso(estado, assinatura_b)

    assert nome_a == nome_b
    assert assinatura_a != assinatura_b
    assert houve_troca is True
    assert "resultado_valor_global" not in estado
    assert estado["assinatura_upload_docs"] == assinatura_b


def test_troca_preserva_estado_global_e_de_interface() -> None:
    estado = _estado_caso_a()

    invalidar_estado_caso(estado, assinatura_conteudo_upload(b"conteudo-b"))

    assert estado["tema_interface"] == "claro"
    assert estado["navegacao_global"] == "central"
