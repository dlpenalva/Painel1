"""Log apartado de decisoes manuais da GCC (RFC §7).

Fonte de verdade: JSONL append-only por contrato, chaveado pelo hash sha256
(16 hex) do arquivo de entrada. O Excel de saida apenas espelha as decisoes
em coluna somente-leitura; regenerar o consolidado nunca perde decisao.
"""
from __future__ import annotations

import hashlib
import json
from datetime import datetime
from pathlib import Path
from typing import Any

DIRETORIO_PADRAO = Path(__file__).resolve().parent / "_decisoes_gcc"


def hash_entrada(conteudo: bytes | None) -> str:
    if not isinstance(conteudo, (bytes, bytearray)) or not conteudo:
        return ""
    return hashlib.sha256(conteudo).hexdigest()[:16]


def _caminho(hash_arquivo: str, diretorio: Path | None = None) -> Path:
    base = Path(diretorio) if diretorio else DIRETORIO_PADRAO
    return base / f"{hash_arquivo}.jsonl"


def registrar_decisao(
    hash_arquivo: str,
    registro_id: str,
    decisao: str,
    justificativa: str = "",
    analista: str = "",
    status_anterior: str = "",
    diretorio: Path | None = None,
    **dados: Any,
) -> dict[str, Any]:
    """Acrescenta uma decisao ao log (append-only) e a devolve."""
    if not hash_arquivo:
        raise ValueError("hash_arquivo vazio: decisao precisa de vinculo com a entrada.")
    if not registro_id or not decisao:
        raise ValueError("registro_id e decisao sao obrigatorios.")
    if not str(justificativa or "").strip():
        raise ValueError("justificativa obrigatoria para registrar decisao GCC.")
    evento = {
        "timestamp": datetime.now().astimezone().isoformat(timespec="seconds"),
        "hash_entrada": hash_arquivo,
        "registro_id": registro_id,
        "decisao": decisao,
        "justificativa": justificativa or "",
        "analista": analista or "",
        "status_anterior": status_anterior or "",
        "status": dados.pop("status", "vigente"),
        "versao_motor": dados.pop("versao_motor", "reconciliacao_evidencias_v10.5.4"),
        **dados,
    }
    caminho = _caminho(hash_arquivo, diretorio)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    with caminho.open("a", encoding="utf-8") as arquivo:
        arquivo.write(json.dumps(evento, ensure_ascii=False) + "\n")
    return evento


def carregar_decisoes(
    hash_arquivo: str,
    diretorio: Path | None = None,
) -> dict[str, dict[str, Any]]:
    """Le o log e devolve a ultima decisao por registro_id.

    O historico integral permanece no arquivo (append-only); aqui so a
    decisao vigente interessa para reconciliar.
    """
    if not hash_arquivo:
        return {}
    caminho = _caminho(hash_arquivo, diretorio)
    if not caminho.exists():
        return {}
    vigentes: dict[str, dict[str, Any]] = {}
    with caminho.open("r", encoding="utf-8") as arquivo:
        for linha in arquivo:
            linha = linha.strip()
            if not linha:
                continue
            try:
                evento = json.loads(linha)
            except json.JSONDecodeError:
                continue
            registro_id = str(evento.get("registro_id") or "")
            if registro_id:
                vigentes[registro_id] = evento
    return vigentes


def historico_decisoes(
    hash_arquivo: str,
    diretorio: Path | None = None,
) -> list[dict[str, Any]]:
    """Historico integral (auditoria), na ordem de registro."""
    if not hash_arquivo:
        return []
    caminho = _caminho(hash_arquivo, diretorio)
    if not caminho.exists():
        return []
    eventos: list[dict[str, Any]] = []
    with caminho.open("r", encoding="utf-8") as arquivo:
        for linha in arquivo:
            linha = linha.strip()
            if not linha:
                continue
            try:
                eventos.append(json.loads(linha))
            except json.JSONDecodeError:
                continue
    return eventos
