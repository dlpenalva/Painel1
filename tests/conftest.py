"""Configuracao comum de testes.

§9 (IST/Anatel): os testes NAO dependem da internet. Por padrao, a consulta a
Anatel fica "indisponivel", forcando o fallback local do IST — comportamento
deterministico e identico ao anterior para todos os testes existentes. Os testes
especificos de IST/Anatel re-patcham `carregar_ist_anatel` para injetar fixtures.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


@pytest.fixture(autouse=True)
def _ist_offline_por_padrao(monkeypatch):
    import _indice_utils

    _indice_utils._resetar_cache_ist()

    def _sem_rede(*args, **kwargs):
        raise RuntimeError("Anatel indisponivel nos testes (rede desabilitada por padrao).")

    monkeypatch.setattr(_indice_utils, "carregar_ist_anatel", _sem_rede, raising=False)
    yield
    _indice_utils._resetar_cache_ist()
