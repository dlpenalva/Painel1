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

# Orçamento de geometria do workbook. A dimensão de uma aba é declarada pelo
# arquivo, não pela aplicação: uma única célula remota infla o retângulo sem
# inflar o pacote, e qualquer varredura posterior materializa esse retângulo
# inteiro. Os tetos vêm do corpus legítimo medido (máximos observados:
# 5.001 linhas, 61 colunas, 145.029 células na maior aba, 204.278 no workbook
# e 16 abas), com margem de 1,6x a 2,5x.
MAX_LINHAS_POR_ABA = 10_000
MAX_COLUNAS_POR_ABA = 100
MAX_AREA_POR_ABA = 300_000
MAX_AREA_TOTAL_WORKBOOK = 500_000
MAX_ABAS_WORKBOOK = 16

# Allowlist por NOME, deliberadamente indiferente ao estado da aba: no corpus
# real, financeiro, itens_Consumidos, itens_PC e aditivos aparecem ora visible,
# ora hidden. Este gate rejeita apenas nome não permitido; a obrigatoriedade
# das abas continua sendo tratada pela lógica de negócio existente.
ABAS_PERMITIDAS = frozenset(
    {
        "CONTROLE",
        "parametros",
        "financeiro",
        "itens_Remanesc",
        "itens_Consumidos",
        "itens_PC",
        "aditivos",
        "posicao_contratual",
        "itens_RC",
        "historico_VU",
        "RESULTADOS",
        "comparativo_VTA",
        "posicao_referencia",
        "cobertura_temporal",
        "MEMORIA_RESULTADOS",
        "CICLO_EM_EXECUCAO",
    }
)

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


def validar_geometria_workbook(wb) -> None:
    """Aprova o orçamento de varredura do workbook logo após o ``load_workbook``.

    Chamado ANTES de qualquer percurso de células, para que as varreduras
    existentes (``_formulas`` e afins) operem sempre dentro de um retângulo já
    aprovado. Lê apenas ``sheetnames``, ``max_row`` e ``max_column`` — nunca
    ``iter_rows``, ``iter_cols``, ``values``, ``rows`` ou ``columns``, que
    materializariam justamente o retângulo sob suspeita.

    A ordem é a mais barata primeiro: contagem de abas, depois nomes, depois a
    geometria de cada aba e, por último, a área acumulada.
    """
    abas = wb.sheetnames
    if len(abas) > MAX_ABAS_WORKBOOK:
        raise XlsxLimiteError(MENSAGEM_LIMITE_XLSX)

    nao_permitidas = [aba for aba in abas if aba not in ABAS_PERMITIDAS]
    if nao_permitidas:
        raise XlsxEstruturaError(MENSAGEM_ESTRUTURA_XLSX)

    area_total = 0
    for ws in wb.worksheets:
        linhas = ws.max_row or 0
        colunas = ws.max_column or 0
        if linhas > MAX_LINHAS_POR_ABA or colunas > MAX_COLUNAS_POR_ABA:
            raise XlsxLimiteError(MENSAGEM_LIMITE_XLSX)

        area = linhas * colunas
        if area > MAX_AREA_POR_ABA:
            raise XlsxLimiteError(MENSAGEM_LIMITE_XLSX)

        area_total += area
        if area_total > MAX_AREA_TOTAL_WORKBOOK:
            raise XlsxLimiteError(MENSAGEM_LIMITE_XLSX)


def opcoes_excel_writer_seguro() -> dict[str, dict[str, bool]]:
    """Opções XlsxWriter que preservam texto não confiável como texto."""
    return {
        "options": {
            "strings_to_formulas": False,
            "strings_to_urls": False,
        }
    }
