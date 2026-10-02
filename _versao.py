"""Versoes publicas e carimbo da versao publicada do Cl8us.

O commit mais recente e a fonte primaria no Streamlit Cloud. O fallback deve
ser atualizado em toda entrega para manter o marcador mesmo quando o Git nao
estiver disponivel no ambiente de execucao.

Politica de versionamento publico (sempre no formato ``XX.X``):

* mudanca apenas visual, textual ou de interface pode alterar
  ``CL8US_VERSION`` sem alterar ``COLETA_VERSION``;
* mudanca estrutural relevante do XLS altera ``COLETA_VERSION``.

``MASTERFILE_VERSION`` continua sendo um marcador tecnico legado onde ainda
for necessario; nao deve ser substituido automaticamente por estas versoes.
"""

from __future__ import annotations

import subprocess
from datetime import datetime
from pathlib import Path


# 11.4: aba RESULTADOS ganha o quadro informativo "Execucao sem efeito financeiro"
# (Financeiro, PCs e Itens) — mudanca estrutural do XLS: COLETA_VERSION 11.1.
# 11.3: modelos em branco do Despacho Saneador e do Termo de Apostila com a mesma
# estrutura dos documentos gerados. 11.2: memoria de calculo da garantia em XLSX.
CL8US_VERSION = "11.4"
COLETA_VERSION = "11.1"
# Modelos de Coleta da familia 11.x aceitos SEM adaptacao: o 11.1 so acrescentou
# um quadro informativo na RESULTADOS/MEMORIA_RESULTADOS (formulas); a Coleta 11.0
# nao o possui e continua valida. Nao remover versoes desta lista sem decisao
# expressa — um arquivo 11.0 nunca e bloqueado por nao ter o quadro novo.
COLETA_VERSOES_ACEITAS = ("11.0", "11.1")
# Janela de compatibilidade retroativa: a versao atual e DUAS linhagens anteriores
# homologadas (PRE_11_L1 e PRE_11_L2, ver _compatibilidade_coleta). Estrutura fora
# desta janela e rejeitada. A formalizacao de Coleta compatibilizada e decidida
# em _formalizacao_compatibilidade, por evidencia tecnica e nunca por versao.
COLETA_COMPATIBILIDADE_ANTERIORES = 2

ATUALIZADO_EM_FALLBACK = "02/10/2026 19:15"


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
