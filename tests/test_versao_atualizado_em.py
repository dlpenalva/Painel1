from __future__ import annotations

import _versao


def test_atualizado_em_prioriza_git_quando_disponivel(monkeypatch):
    monkeypatch.setattr(_versao, "_data_ultimo_commit", lambda: "08/10/2026 19:50")
    monkeypatch.setattr(_versao, "ATUALIZADO_EM_FALLBACK", "08/10/2026 21:30")

    assert _versao.atualizado_em() == "08/10/2026 19:50"


def test_atualizado_em_usa_fallback_somente_sem_git(monkeypatch):
    monkeypatch.setattr(_versao, "_data_ultimo_commit", lambda: None)
    monkeypatch.setattr(_versao, "ATUALIZADO_EM_FALLBACK", "08/10/2026 19:50")

    assert _versao.atualizado_em() == "08/10/2026 19:50"
