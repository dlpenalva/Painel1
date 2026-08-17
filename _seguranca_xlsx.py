"""Fronteira de segurança para XLSX recebidos e gerados pelo cl8us.

Esta camada valida somente o contêiner ZIP/OOXML antes do primeiro parser.
Regras funcionais da Coleta continuam nos leitores existentes.
"""
from __future__ import annotations

from collections import Counter
from io import BytesIO
from pathlib import PurePosixPath
from zipfile import ZipFile


TAMANHO_MAXIMO_ARQUIVO = 20 * 1024 * 1024
TAMANHO_MAXIMO_DESCOMPACTADO = 100 * 1024 * 1024
TAMANHO_MAXIMO_MEMBRO = 50 * 1024 * 1024
QUANTIDADE_MAXIMA_MEMBROS = 250
RAZAO_MAXIMA_COMPRESSAO_MEMBRO = 100

MENSAGEM_XLSX_INVALIDO = "O arquivo enviado não é um XLSX válido ou está corrompido."
MENSAGEM_LIMITE_XLSX = "O arquivo excede os limites de segurança permitidos."
MENSAGEM_ESTRUTURA_XLSX = "O arquivo XLSX possui uma estrutura interna não permitida."

_ASSINATURA_ZIP_LOCAL = b"PK\x03\x04"
_COMPONENTES_OOXML_MINIMOS = frozenset(
    {"[Content_Types].xml", "_rels/.rels", "xl/workbook.xml"}
)


class ErroSegurancaXlsx(ValueError):
    """Erro controlado e seguro para exibição ao usuário."""


class XlsxInvalidoError(ErroSegurancaXlsx):
    """O conteúdo não forma um pacote XLSX íntegro."""


class XlsxLimiteError(ErroSegurancaXlsx):
    """Um limite de recursos do pacote foi excedido."""


class XlsxEstruturaError(ErroSegurancaXlsx):
    """A estrutura interna do ZIP não é permitida."""


class _ConteudoXlsxValidado(bytes):
    """Marca interna que evita reinspecionar o mesmo upload nos leitores."""


def _nome_membro_permitido(nome: str) -> bool:
    if not nome or "\x00" in nome:
        return False
    normalizado = nome.replace("\\", "/")
    caminho = PurePosixPath(normalizado)
    if caminho.is_absolute() or ".." in caminho.parts:
        return False
    return not (caminho.parts and caminho.parts[0].endswith(":"))


def validar_xlsx_antes_do_parser(conteudo: bytes) -> bytes:
    """Valida limites ZIP/OOXML e devolve os mesmos bytes marcados como seguros.

    O diretório central é inspecionado antes de ``testzip()``, de modo que
    nenhum membro seja descompactado antes da aprovação dos limites.
    """
    if isinstance(conteudo, _ConteudoXlsxValidado):
        return conteudo
    if not isinstance(conteudo, (bytes, bytearray, memoryview)):
        raise XlsxInvalidoError(MENSAGEM_XLSX_INVALIDO)

    bytes_conteudo = bytes(conteudo)
    if not bytes_conteudo or not bytes_conteudo.startswith(_ASSINATURA_ZIP_LOCAL):
        raise XlsxInvalidoError(MENSAGEM_XLSX_INVALIDO)
    if len(bytes_conteudo) > TAMANHO_MAXIMO_ARQUIVO:
        raise XlsxLimiteError(MENSAGEM_LIMITE_XLSX)

    try:
        with ZipFile(BytesIO(bytes_conteudo)) as pacote:
            membros = pacote.infolist()
            if not membros:
                raise XlsxInvalidoError(MENSAGEM_XLSX_INVALIDO)
            if len(membros) > QUANTIDADE_MAXIMA_MEMBROS:
                raise XlsxLimiteError(MENSAGEM_LIMITE_XLSX)

            nomes = [membro.filename for membro in membros]
            if any(contagem > 1 for contagem in Counter(nomes).values()):
                raise XlsxEstruturaError(MENSAGEM_ESTRUTURA_XLSX)

            total_descompactado = 0
            for membro in membros:
                if not _nome_membro_permitido(membro.filename):
                    raise XlsxEstruturaError(MENSAGEM_ESTRUTURA_XLSX)
                if membro.file_size > TAMANHO_MAXIMO_MEMBRO:
                    raise XlsxLimiteError(MENSAGEM_LIMITE_XLSX)

                total_descompactado += membro.file_size
                if total_descompactado > TAMANHO_MAXIMO_DESCOMPACTADO:
                    raise XlsxLimiteError(MENSAGEM_LIMITE_XLSX)

                if not membro.is_dir() and membro.file_size:
                    if membro.compress_size == 0:
                        raise XlsxLimiteError(MENSAGEM_LIMITE_XLSX)
                    razao = membro.file_size / membro.compress_size
                    if razao > RAZAO_MAXIMA_COMPRESSAO_MEMBRO:
                        raise XlsxLimiteError(MENSAGEM_LIMITE_XLSX)

            if pacote.testzip() is not None:
                raise XlsxInvalidoError(MENSAGEM_XLSX_INVALIDO)

            if not _COMPONENTES_OOXML_MINIMOS.issubset(nomes):
                raise XlsxEstruturaError(MENSAGEM_ESTRUTURA_XLSX)
    except ErroSegurancaXlsx:
        raise
    except Exception as exc:
        raise XlsxInvalidoError(MENSAGEM_XLSX_INVALIDO) from exc

    return _ConteudoXlsxValidado(bytes_conteudo)


def garantir_xlsx_validado(conteudo: bytes) -> bytes:
    """Valida chamadas isoladas e vira operação sem custo após a fronteira."""
    return validar_xlsx_antes_do_parser(conteudo)


def opcoes_excel_writer_seguro() -> dict[str, dict[str, bool]]:
    """Opções XlsxWriter que preservam texto não confiável como texto."""
    return {
        "options": {
            "strings_to_formulas": False,
            "strings_to_urls": False,
        }
    }
