"""Fabrica de artefatos caros da suite: gera uma vez, entrega copias baratas.

POR QUE ISSO EXISTE
-------------------
``gerar_coleta_oficial_preenchida`` custa ~13 s por chamada. Nao e desperdicio
do gerador: o template oficial tem cerca de 16 MB de XML em uma unica aba, e
cada geracao carrega o workbook inteiro, varre a matriz de formulas duas vezes,
salva e reabre para conferir. Como cada teste chamava o gerador de novo com a
MESMA entrada, o mesmo arquivo era reconstruido dezenas de vezes por sessao.

O QUE E SEGURO CACHEAR
----------------------
Somente BYTES, que sao imutaveis. O cache guarda os bytes de cada entrada
distinta; quem precisa mutar recebe um ``Workbook`` novo, carregado na hora a
partir desses bytes. Nenhum objeto openpyxl e compartilhado entre testes, nao
ha estado global mutavel e nenhum teste depende da ordem de execucao — rodar um
teste sozinho apenas gera o artefato dele na primeira vez.

QUANDO NAO USAR
---------------
Se o teste precisa observar o ATO de gerar (regeracao, idempotencia, efeito de
um monkeypatch no gerador ou no template), chame ``gerar_coleta_oficial_preenchida``
diretamente: o cache mascararia justamente a propriedade sob teste.
"""
from __future__ import annotations

import json
from io import BytesIO
from typing import Any

from openpyxl import load_workbook

_CACHE: dict[str, bytes] = {}


def _chave(dados: dict[str, Any] | None) -> str:
    return json.dumps(dados, sort_keys=True, default=str)


def bytes_coleta_oficial(dados: dict[str, Any] | None = None) -> bytes:
    """Bytes da Coleta oficial preenchida com ``dados``, gerados uma unica vez."""
    from _coleta_oficial import gerar_coleta_oficial_preenchida

    chave = _chave(dados)
    if chave not in _CACHE:
        _CACHE[chave] = gerar_coleta_oficial_preenchida(dados)
    return _CACHE[chave]


def workbook_coleta_oficial(dados: dict[str, Any] | None = None, *, data_only: bool = False):
    """Workbook NOVO e independente, montado a partir dos bytes em cache."""
    return load_workbook(BytesIO(bytes_coleta_oficial(dados)), data_only=data_only)


def estatisticas_cache() -> dict[str, int]:
    """Usado pelos testes da propria fabrica."""
    return {"entradas": len(_CACHE), "bytes": sum(len(v) for v in _CACHE.values())}
