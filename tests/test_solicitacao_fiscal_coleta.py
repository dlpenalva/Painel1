"""ADM-UX-COM1: solicitacao interna opcional apos a admissibilidade."""

from pathlib import Path

from _solicitacao_fiscal import (
    COMANDO_SOLICITACAO_FISCAL,
    TEXTO_SOLICITACAO_FISCAL,
    sincronizar_texto_solicitacao_fiscal,
)


ROOT = Path(__file__).resolve().parents[1]
SIMPLES = ROOT / "pages" / "01_Calculo_Simples.py"
MULTIPLOS = ROOT / "pages" / "02_Calculo_Represados.py"
UPLOAD = ROOT / "pages" / "03_Valor_Global.py"


def _fonte(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def test_texto_base_aprovado_e_generico():
    texto = TEXTO_SOLICITACAO_FISCAL
    assert texto.startswith(
        "Assunto: Solicitação de informações para continuidade da análise de reajuste contratual"
    )
    for trecho in (
        "itens_Remanesc",
        "financeiro",
        "itens_PC",
        "CICLO_EM_EXECUCAO",
        "VTA — Valor Total Atualizado",
    ):
        assert trecho in texto
    for proibido in (
        "[CONTRATO]",
        "[CONTRATADA]",
        "[NOME]",
        "[PROCESSO]",
        "CT-",
    ):
        assert proibido not in texto


def test_edicao_sobrevive_rerun_e_nova_analise_restaura_base():
    estado = {}
    assinatura_a = ("simples", ("C1", "2023-08-02", "IST"))
    assinatura_b = ("simples", ("C2", "2024-08-02", "IST"))

    assert sincronizar_texto_solicitacao_fiscal(estado, assinatura_a) == TEXTO_SOLICITACAO_FISCAL
    estado["solicitacao_fiscal_coleta_texto"] = "Edição específica desta análise"
    assert sincronizar_texto_solicitacao_fiscal(estado, assinatura_a) == "Edição específica desta análise"
    assert sincronizar_texto_solicitacao_fiscal(estado, assinatura_b) == TEXTO_SOLICITACAO_FISCAL


def test_acao_opcional_nos_dois_fluxos_sem_gate():
    assert COMANDO_SOLICITACAO_FISCAL == (
        "Solicitar ao fiscal o preenchimento da planilha Coleta"
    )
    for pagina in (SIMPLES, MULTIPLOS):
        fonte = _fonte(pagina)
        assert "render_solicitacao_fiscal_coleta(" in fonte
        chamada = fonte[fonte.rindex("render_solicitacao_fiscal_coleta(") :]
        assert "st.stop()" not in chamada
        assert "if st.button" not in chamada


def test_comunicacao_contratada_permanece_antes_e_independente():
    simples = _fonte(SIMPLES)
    multiplos = _fonte(MULTIPLOS)
    assert simples.rindex("render_email_contratada(") < simples.rindex(
        "render_solicitacao_fiscal_coleta("
    )
    assert multiplos.rindex("render_email_contratada(") < multiplos.rindex(
        "render_solicitacao_fiscal_coleta("
    )
    assert "ASSUNTO_EMAIL_CONTRATADA" in _fonte(ROOT / "_email_contratada.py")


def test_upload_nao_recebe_a_nova_acao():
    assert "render_solicitacao_fiscal_coleta" not in _fonte(UPLOAD)


def test_helper_nao_importa_motores_calculos_ou_xlsx():
    fonte = _fonte(ROOT / "_solicitacao_fiscal.py")
    for proibido in (
        "_reajuste_utils",
        "_indice_utils",
        "_coleta_oficial",
        "openpyxl",
        "xlsxwriter",
        "pandas",
    ):
        assert proibido not in fonte
