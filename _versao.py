"""Carimbo da versao publicada do Master 2.0.

O commit mais recente e a fonte primaria no Streamlit Cloud. O fallback deve
ser atualizado em toda entrega para manter o marcador mesmo quando o Git nao
estiver disponivel no ambiente de execucao.
"""

from __future__ import annotations

import subprocess
from datetime import datetime
from pathlib import Path


ATUALIZADO_EM_FALLBACK = "14/07/2026 17:12"


def _data_ultimo_commit() -> str | None:
    try:
        resultado = subprocess.run(
            ["git", "log", "-1", "--format=%cI"],
            cwd=Path(__file__).resolve().parent,
            capture_output=True,
            text=True,
            timeout=5,
        )
        valor = (resultado.stdout or "").strip()
        if not valor:
            return None
        return datetime.fromisoformat(valor).strftime("%d/%m/%Y %H:%M")
    except Exception:
        return None


def atualizado_em() -> str:
    """Retorna o carimbo visivel no formato brasileiro."""
    data_commit = _data_ultimo_commit()
    if not data_commit:
        return ATUALIZADO_EM_FALLBACK
    formato = "%d/%m/%Y %H:%M"
    return max((data_commit, ATUALIZADO_EM_FALLBACK), key=lambda valor: datetime.strptime(valor, formato))
