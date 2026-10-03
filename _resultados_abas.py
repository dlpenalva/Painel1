"""Resolucao das abas de RESULTADOS por versao do arquivo de Coleta.

A partir da Coleta 11.2 (RESULTADOS-EXECUTIVO-V3) o XLS tem DUAS abas:

* ``RESULTADOS``         — pagina executiva; so espelha fontes canonicas e NAO
                           e lida por coordenada por nenhum consumidor Python;
* ``RESULTADOS_DETALHE`` — a antiga aba RESULTADOS, renomeada pelo Excel com
                           todas as formulas, nomes definidos e ajustes manuais
                           nas MESMAS coordenadas.

Arquivos anteriores (11.0, 11.1 e as linhagens PRE_11) so tem ``RESULTADOS``, e
la ela ainda e a aba tecnica. Toda leitura por coordenada da camada tecnica
(status B3, referencias B10:B13/H10:H13, A1, ajustes C43:G50...) deve passar
por :func:`aba_resultados_tecnica`, para nunca ler a pagina executiva de um
arquivo novo achando que e a aba antiga.
"""

from __future__ import annotations

from typing import Any

ABA_RESULTADOS = "RESULTADOS"
ABA_RESULTADOS_DETALHE = "RESULTADOS_DETALHE"

# Titulo literal A1 de cada camada tecnica (gate de integridade da aba).
TITULO_RESULTADOS_LEGADO = "RESULTADOS CONSOLIDADOS — REAJUSTE CONTRATUAL"
TITULO_RESULTADOS_DETALHE = "RESULTADOS — DETALHE TÉCNICO DA APURAÇÃO"
TITULOS_TECNICOS = (TITULO_RESULTADOS_LEGADO, TITULO_RESULTADOS_DETALHE)
# A pagina executiva tem o titulo em B2 (a coluna A e margem).
TITULO_RESULTADOS_EXECUTIVO = "RESULTADO DA APURAÇÃO"
CELULA_TITULO_EXECUTIVO = "B2"


def _nomes(wb_ou_nomes: Any) -> list[str]:
    nomes = getattr(wb_ou_nomes, "sheetnames", wb_ou_nomes)
    return list(nomes or [])


def tem_resultados_executivo(wb_ou_nomes: Any) -> bool:
    """True quando o arquivo segue a arquitetura executiva + detalhe."""
    nomes = _nomes(wb_ou_nomes)
    return ABA_RESULTADOS_DETALHE in nomes and ABA_RESULTADOS in nomes


def aba_resultados_tecnica(wb_ou_nomes: Any) -> str | None:
    """Nome da aba TECNICA de RESULTADOS (coordenadas homologadas).

    Arquivo novo -> ``RESULTADOS_DETALHE``; arquivo anterior -> ``RESULTADOS``;
    sem nenhuma das duas -> ``None``.

    Fail-closed: workbook com marcador publico 11.2+ SEM ``RESULTADOS_DETALHE``
    devolve ``None`` — a RESULTADOS dele e a pagina executiva e nunca pode ser
    lida como camada tecnica. A versao vem do marcador canonico
    (`_compatibilidade_coleta.exige_resultados_detalhe`); uma lista de nomes
    de abas nao carrega marcador e segue a regra estrutural.
    """
    nomes = _nomes(wb_ou_nomes)
    if ABA_RESULTADOS_DETALHE in nomes:
        return ABA_RESULTADOS_DETALHE
    if ABA_RESULTADOS in nomes:
        if hasattr(wb_ou_nomes, "sheetnames"):
            from _compatibilidade_coleta import exige_resultados_detalhe

            if exige_resultados_detalhe(wb_ou_nomes):
                return None
        return ABA_RESULTADOS
    return None
