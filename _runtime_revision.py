"""Coerencia de revisao do codigo local dentro de um processo vivo.

O Streamlit Community Cloud atualiza os arquivos do checkout sem reiniciar o
processo Python. As paginas voltam a ser compiladas a partir do disco (o
``ScriptCache`` e limpo quando as sessoes se desconectam), mas os modulos
locais ja importados permanecem em ``sys.modules`` com os objetos da revisao
anterior. O resultado e uma arvore hibrida: arquivo novo executando com modulo
antigo — ``ImportError`` no melhor caso, calculo com regra velha no pior.

Este modulo detecta a troca de revisao pelo SHA do checkout e, apenas nesse
momento, remove de ``sys.modules`` os modulos do proprio projeto para que os
imports seguintes leiam o disco. Em rerun normal (mesma revisao) o custo e a
leitura de dois arquivos pequenos e nada e recarregado.
"""

from __future__ import annotations

import sys
import threading
from pathlib import Path


RAIZ = Path(__file__).resolve().parent

# A revisao ja carregada no processo vive fora do proprio modulo: a purga
# remove este arquivo de sys.modules e um sentinela interno se perderia.
# Somente metadado tecnico (o SHA do checkout) e guardado aqui.
_ATRIBUTO_SENTINELA = "_painel1_revisao_carregada"

_TRAVA = threading.Lock()

_DIRETORIOS_IGNORADOS = ("site-packages", "dist-packages", ".venv", "venv")


class RevisaoIncoerenteError(RuntimeError):
    """A troca de revisao foi detectada mas a purga nao pode ser concluida."""


def _diretorio_comum(git: Path) -> Path:
    """Onde ficam refs/ e packed-refs (numa worktree, o diretorio principal)."""
    ponteiro = git / "commondir"
    if not ponteiro.is_file():
        return git
    destino = Path(ponteiro.read_text(encoding="utf-8", errors="replace").strip())
    return destino if destino.is_absolute() else (git / destino).resolve()


def _resolver_sha(git: Path) -> str | None:
    cabeca = git / "HEAD"
    if not cabeca.is_file():
        return None
    conteudo = cabeca.read_text(encoding="utf-8", errors="replace").strip()
    if not conteudo.startswith("ref:"):
        return conteudo or None
    referencia = conteudo[4:].strip()
    comum = _diretorio_comum(git)
    for base in dict.fromkeys((git, comum)):
        solta = base / referencia
        if solta.is_file():
            return solta.read_text(encoding="utf-8", errors="replace").strip() or None
    for base in dict.fromkeys((git, comum)):
        empacotadas = base / "packed-refs"
        if not empacotadas.is_file():
            continue
        for linha in empacotadas.read_text(encoding="utf-8", errors="replace").splitlines():
            if linha.startswith(("#", "^")):
                continue
            partes = linha.split()
            if len(partes) == 2 and partes[1] == referencia:
                return partes[0]
    return None


def revisao_atual(raiz: Path | None = None) -> str | None:
    """SHA do checkout, ou ``None`` quando a identidade nao pode ser lida.

    Leitura direta do plumbing do Git (dois arquivos pequenos), sem subprocesso,
    para que o custo em rerun normal seja irrelevante.
    """
    base = Path(raiz) if raiz is not None else RAIZ
    try:
        ponteiro = base / ".git"
        if ponteiro.is_file():  # worktree: ".git" e um arquivo "gitdir: <caminho>"
            destino = ponteiro.read_text(encoding="utf-8", errors="replace").strip()
            if not destino.startswith("gitdir:"):
                return None
            git = Path(destino[7:].strip())
            if not git.is_absolute():
                git = (base / git).resolve()
        elif ponteiro.is_dir():
            git = ponteiro
        else:
            return None
        return _resolver_sha(git)
    except OSError:
        return None


def _eh_modulo_local(modulo: object, raiz: Path) -> bool:
    arquivo = getattr(modulo, "__file__", None)
    if not arquivo:
        return False
    try:
        caminho = Path(arquivo).resolve()
    except OSError:
        return False
    if any(parte in _DIRETORIOS_IGNORADOS for parte in caminho.parts):
        return False
    return raiz in caminho.parents


def modulos_locais_carregados(raiz: Path | None = None) -> list[str]:
    """Nomes em ``sys.modules`` cujo arquivo pertence ao proprio checkout."""
    base = (Path(raiz) if raiz is not None else RAIZ).resolve()
    return [
        nome
        for nome, modulo in list(sys.modules.items())
        if nome != "__main__" and _eh_modulo_local(modulo, base)
    ]


def _purgar(raiz: Path) -> list[str]:
    nomes = modulos_locais_carregados(raiz)
    for nome in nomes:
        sys.modules.pop(nome, None)
    return nomes


def garantir_revisao_coerente(raiz: Path | None = None) -> list[str]:
    """Alinha ``sys.modules`` a revisao presente no disco.

    Devolve os modulos removidos. Lista vazia significa que nada foi feito:
    primeiro boot, mesma revisao ou identidade indisponivel.
    """
    base = (Path(raiz) if raiz is not None else RAIZ).resolve()
    revisao = revisao_atual(base)
    if revisao is None:
        return []
    with _TRAVA:
        anterior = getattr(sys, _ATRIBUTO_SENTINELA, None)
        if anterior == revisao:
            return []
        # A revisao e registrada antes da purga: uma sessao concorrente que
        # entre neste ponto no mesmo instante nao repete a remocao.
        setattr(sys, _ATRIBUTO_SENTINELA, revisao)
        if anterior is None:
            return []  # primeiro boot: nada foi importado da revisao anterior
        try:
            return _purgar(base)
        except Exception as erro:  # pragma: no cover - fail-closed
            raise RevisaoIncoerenteError(
                "A revisao do codigo mudou e os modulos da revisao anterior nao "
                "puderam ser descarregados. Reinicie o aplicativo (Manage app -> Reboot)."
            ) from erro
