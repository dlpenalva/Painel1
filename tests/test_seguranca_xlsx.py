from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from openpyxl import Workbook

import _seguranca_xlsx as seguranca
from _coleta_oficial import gerar_coleta_oficial_preenchida
from _coleta_reajuste import ler_coleta_reajuste
from _coleta_reajuste_documentos import processar_coleta_oficial_runtime


ROOT = Path(__file__).resolve().parents[1]
COMPONENTES = {
    "[Content_Types].xml": "<Types/>",
    "_rels/.rels": "<Relationships/>",
    "xl/workbook.xml": "<workbook/>",
}


def _zip(membros: list[tuple[str, bytes | str]]) -> bytes:
    saida = BytesIO()
    with ZipFile(saida, "w", ZIP_DEFLATED) as pacote:
        for nome, conteudo in membros:
            pacote.writestr(nome, conteudo)
    return saida.getvalue()


def _xlsx_minimo_valido() -> bytes:
    saida = BytesIO()
    Workbook().save(saida)
    return saida.getvalue()


@pytest.mark.parametrize("conteudo", [b"", b"texto renomeado para xlsx"])
def test_rejeita_conteudo_que_nao_e_zip(conteudo: bytes) -> None:
    with pytest.raises(seguranca.XlsxInvalidoError, match="não é um XLSX válido"):
        seguranca.validar_xlsx_antes_do_parser(conteudo)


def test_rejeita_zip_generico_e_mensagem_nao_vaza_detalhe_interno() -> None:
    with pytest.raises(seguranca.XlsxEstruturaError) as erro:
        seguranca.validar_xlsx_antes_do_parser(_zip([("nota.txt", "seguro")]))
    assert str(erro.value) == seguranca.MENSAGEM_ESTRUTURA_XLSX
    assert "Content_Types" not in str(erro.value)


def test_rejeita_zip_truncado() -> None:
    truncado = _xlsx_minimo_valido()[:-20]
    with pytest.raises(seguranca.XlsxInvalidoError):
        seguranca.validar_xlsx_antes_do_parser(truncado)


def test_aceita_xlsx_minimo_real() -> None:
    conteudo = _xlsx_minimo_valido()
    assert bytes(seguranca.validar_xlsx_antes_do_parser(conteudo)) == conteudo


@pytest.mark.parametrize(
    "arquivo",
    ["Coleta_Reajuste.xlsx", "COLETA_REAJUSTE_OFICIAL.xlsx"],
)
def test_aceita_modelos_oficiais(arquivo: str) -> None:
    conteudo = (ROOT / "templates" / arquivo).read_bytes()
    assert bytes(seguranca.validar_xlsx_antes_do_parser(conteudo)) == conteudo


def test_aceita_modelo_oficial_preenchido_gerado() -> None:
    conteudo = gerar_coleta_oficial_preenchida({})
    assert bytes(seguranca.validar_xlsx_antes_do_parser(conteudo)) == conteudo


def test_rejeita_tamanho_compactado_excessivo(monkeypatch) -> None:
    conteudo = _xlsx_minimo_valido()
    monkeypatch.setattr(seguranca, "TAMANHO_MAXIMO_ARQUIVO", len(conteudo) - 1)
    with pytest.raises(seguranca.XlsxLimiteError):
        seguranca.validar_xlsx_antes_do_parser(conteudo)


def test_rejeita_total_descompactado_excessivo(monkeypatch) -> None:
    conteudo = _zip(list(COMPONENTES.items()) + [("xl/dados.xml", b"A" * 100)])
    monkeypatch.setattr(seguranca, "TAMANHO_MAXIMO_DESCOMPACTADO", 99)
    with pytest.raises(seguranca.XlsxLimiteError):
        seguranca.validar_xlsx_antes_do_parser(conteudo)


def test_rejeita_maior_membro_excessivo(monkeypatch) -> None:
    conteudo = _zip(list(COMPONENTES.items()) + [("xl/dados.xml", b"A" * 100)])
    monkeypatch.setattr(seguranca, "TAMANHO_MAXIMO_MEMBRO", 99)
    with pytest.raises(seguranca.XlsxLimiteError):
        seguranca.validar_xlsx_antes_do_parser(conteudo)


def test_rejeita_quantidade_excessiva_de_membros(monkeypatch) -> None:
    conteudo = _zip(list(COMPONENTES.items()) + [("xl/extra.xml", b"x")])
    monkeypatch.setattr(seguranca, "QUANTIDADE_MAXIMA_MEMBROS", 3)
    with pytest.raises(seguranca.XlsxLimiteError):
        seguranca.validar_xlsx_antes_do_parser(conteudo)


def test_rejeita_razao_de_compressao_excessiva(monkeypatch) -> None:
    conteudo = _zip(list(COMPONENTES.items()) + [("xl/dados.xml", b"A" * 2_000)])
    monkeypatch.setattr(seguranca, "RAZAO_MAXIMA_COMPRESSAO_MEMBRO", 2)
    with pytest.raises(seguranca.XlsxLimiteError):
        seguranca.validar_xlsx_antes_do_parser(conteudo)


def test_rejeita_entrada_duplicada() -> None:
    membros = list(COMPONENTES.items()) + [("xl/workbook.xml", "duplicado")]
    with pytest.warns(UserWarning, match="Duplicate name"):
        conteudo = _zip(membros)
    with pytest.raises(seguranca.XlsxEstruturaError):
        seguranca.validar_xlsx_antes_do_parser(conteudo)


@pytest.mark.parametrize("nome", ["../segredo.xml", "/absoluto.xml", "C:/absoluto.xml"])
def test_rejeita_traversal_e_caminhos_absolutos(nome: str) -> None:
    conteudo = _zip(list(COMPONENTES.items()) + [(nome, "x")])
    with pytest.raises(seguranca.XlsxEstruturaError):
        seguranca.validar_xlsx_antes_do_parser(conteudo)


@pytest.mark.parametrize("ausente", ["[Content_Types].xml", "xl/workbook.xml"])
def test_rejeita_componente_ooxml_minimo_ausente(ausente: str) -> None:
    conteudo = _zip([(n, c) for n, c in COMPONENTES.items() if n != ausente])
    with pytest.raises(seguranca.XlsxEstruturaError):
        seguranca.validar_xlsx_antes_do_parser(conteudo)


def test_leitor_rejeita_antes_de_openpyxl(monkeypatch) -> None:
    def nao_pode_abrir(*_args, **_kwargs):
        raise AssertionError("openpyxl.load_workbook não deveria ser chamado")

    monkeypatch.setattr("_coleta_reajuste.load_workbook", nao_pode_abrir)
    with pytest.raises(seguranca.XlsxInvalidoError):
        ler_coleta_reajuste(b"nao e XLSX")


def test_runtime_rejeita_antes_de_acionar_leitores(monkeypatch) -> None:
    def nao_pode_ler(*_args, **_kwargs):
        raise AssertionError("leitor não deveria ser chamado")

    monkeypatch.setattr("_coleta_reajuste_documentos.ler_masterfile_v10", nao_pode_ler)
    with pytest.raises(seguranca.XlsxInvalidoError):
        processar_coleta_oficial_runtime(b"nao e XLSX")


def test_runtime_nao_propaga_erro_tecnico_do_parser(monkeypatch) -> None:
    monkeypatch.setattr(
        "_coleta_reajuste_documentos.ler_masterfile_v10", lambda *_a, **_k: {"ok": True}
    )
    monkeypatch.setattr(
        "_coleta_reajuste_documentos.ler_coleta_reajuste",
        lambda *_a, **_k: {"valido": True},
    )

    def falha_tecnica(*_args, **_kwargs):
        raise ValueError("detalhe interno do parser")

    monkeypatch.setattr(
        "_coleta_reajuste_documentos.adaptar_coleta_reajuste_para_documentos",
        falha_tecnica,
    )
    with pytest.raises(seguranca.XlsxInvalidoError) as erro:
        processar_coleta_oficial_runtime(_xlsx_minimo_valido())
    assert str(erro.value) == seguranca.MENSAGEM_XLSX_INVALIDO
    assert "detalhe interno" not in str(erro.value)


def test_marcador_evitaria_cinco_inspecoes_zip(monkeypatch) -> None:
    validado = seguranca.validar_xlsx_antes_do_parser(_xlsx_minimo_valido())

    def nao_pode_reabrir(*_args, **_kwargs):
        raise AssertionError("ZIP já validado não deve ser reinspecionado")

    monkeypatch.setattr(seguranca, "ZipFile", nao_pode_reabrir)
    for _ in range(5):
        assert seguranca.garantir_xlsx_validado(validado) is validado
