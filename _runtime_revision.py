"""Coerencia de revisao do codigo local dentro de um processo vivo.

O Streamlit Community Cloud atualiza os arquivos do checkout sem reiniciar o
processo Python. As paginas voltam a ser compiladas a partir do disco (o
``ScriptCache`` e limpo quando as sessoes se desconectam), mas os modulos
locais ja importados permanecem em ``sys.modules`` com os objetos da revisao
anterior. O resultado e uma arvore hibrida: arquivo novo executando com modulo
antigo — ``ImportError`` no melhor caso, calculo com regra velha no pior.

Este modulo detecta a troca de revisao pelo SHA do checkout e, apenas nesse
momento, executa a transicao: remove de ``sys.modules`` os modulos do proprio
projeto e esvazia ``st.cache_data``. Em rerun normal (mesma revisao) o custo e
a leitura de dois arquivos pequenos e nada e recarregado.

A revisao nova so e declarada coerente DEPOIS que a transicao inteira terminou
com sucesso. Se qualquer etapa falhar, o processo fica marcado como incoerente
e toda execucao seguinte falha ate o reboot — nunca um hibrido silencioso.

A coordenacao (trava e estado) vive em ``sys``, e nao neste modulo, porque a
propria transicao remove este arquivo de ``sys.modules``: uma sessao que
reimporte o helper no meio da troca recebe uma instancia nova do modulo e
precisa encontrar a MESMA trava do processo.
"""

from __future__ import annotations

import sys
import threading
from pathlib import Path


RAIZ = Path(__file__).resolve().parent

# Somente metadado tecnico do processo: a trava, a situacao e o SHA. Nunca
# dado de usuario.
_ATRIBUTO_ESTADO = "_painel1_revisao_estado"

NOVO = "NOVO"            # nada registrado ainda neste processo
PRONTO = "PRONTO"        # a revisao registrada e coerente
TRANSICAO = "TRANSICAO"  # troca em curso, sob a trava
FALHA = "FALHA"          # transicao incompleta: so o reboot recupera

_DIRETORIOS_IGNORADOS = ("site-packages", "dist-packages", ".venv", "venv")

_ORIENTACAO_REBOOT = (
    "Reinicie o aplicativo (Manage app -> Reboot app) antes de prosseguir."
)


class RevisaoIncoerenteError(RuntimeError):
    """Nao ha garantia de que a execucao usa uma unica revisao do codigo."""


def _estado_do_processo() -> dict:
    """A trava e o estado da revisao, compartilhados por todo o processo.

    ``dict.setdefault`` e atomico: duas sessoes que cheguem juntas no primeiro
    boot obtem o MESMO dicionario e, portanto, a MESMA trava.
    """
    estado = getattr(sys, _ATRIBUTO_ESTADO, None)
    if estado is None:
        estado = sys.__dict__.setdefault(
            _ATRIBUTO_ESTADO,
            {"trava": threading.RLock(), "situacao": NOVO, "revisao": None},
        )
    return estado


# --------------------------------------------------------------------------
# identidade da revisao
# --------------------------------------------------------------------------

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


# --------------------------------------------------------------------------
# acoes da transicao
# --------------------------------------------------------------------------

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


def _modulos_de_revisao_anterior(raiz: Path) -> list[str]:
    """Modulos locais que podem pertencer a outra revisao.

    O proprio helper nao conta: num processo recem-iniciado ele e o unico
    modulo local carregado quando ``app.py`` chama a blindagem, e purga-lo ai
    seria trabalho a toa.
    """
    return [nome for nome in modulos_locais_carregados(raiz) if nome != __name__]


def _purgar(raiz: Path) -> list[str]:
    nomes = modulos_locais_carregados(raiz)
    for nome in nomes:
        sys.modules.pop(nome, None)
    return nomes


def _limpar_caches_de_dados() -> None:
    """Esvazia ``st.cache_data`` uma unica vez, dentro da transicao.

    A chave de cache do Streamlit hasheia o fonte da funcao decorada, mas nao o
    dos helpers que ela chama: uma funcao inalterada que chama um helper local
    alterado devolveria o valor da revisao anterior ate o TTL. ``sys.modules``
    coerente nao implica artefato cacheado coerente.

    ``st.session_state`` e ``st.cache_resource`` nao sao tocados.
    """
    import streamlit as st

    st.cache_data.clear()


# --------------------------------------------------------------------------
# maquina de estados da revisao
# --------------------------------------------------------------------------

def _falhar(estado: dict, motivo: str, causa: BaseException | None = None) -> None:
    estado["situacao"] = FALHA
    erro = RevisaoIncoerenteError(motivo + " " + _ORIENTACAO_REBOOT)
    if causa is not None:
        raise erro from causa
    raise erro


def _transicionar(estado: dict, raiz: Path, revisao: str | None) -> list[str]:
    """Executa a troca inteira e so entao declara a revisao coerente."""
    estado["situacao"] = TRANSICAO
    try:
        removidos = _purgar(raiz)
        _limpar_caches_de_dados()
    except Exception as erro:  # noqa: BLE001 - qualquer falha aqui e fail-closed
        _falhar(
            estado,
            "A revisao do codigo mudou e a troca nao pode ser concluida.",
            erro,
        )
    except BaseException:
        # Interrupcao dura no meio da troca: o processo nao pode voltar a
        # considerar esta revisao coerente so porque a excecao subiu.
        estado["situacao"] = FALHA
        raise
    estado["situacao"] = PRONTO
    estado["revisao"] = revisao
    return removidos


def _sem_identidade(estado: dict, raiz: Path) -> list[str]:
    """Nao foi possivel ler o SHA do checkout."""
    if estado["situacao"] == PRONTO and estado["revisao"] is None:
        return []  # ambiente sem Git desde o inicio: coerente por construcao
    if estado["situacao"] == NOVO and not _modulos_de_revisao_anterior(raiz):
        estado["situacao"] = PRONTO
        estado["revisao"] = None
        return []
    _falhar(
        estado,
        "A revisao do codigo deixou de ser identificavel e nao ha como garantir "
        "que esta execucao usa uma unica revisao.",
    )
    return []  # inalcancavel: _falhar sempre levanta


def garantir_revisao_coerente(raiz: Path | None = None) -> list[str]:
    """Alinha o runtime a revisao presente no disco.

    Devolve os modulos removidos. Lista vazia significa que nada foi feito:
    primeiro boot limpo ou mesma revisao. Levanta ``RevisaoIncoerenteError``
    quando a coerencia nao pode ser garantida.
    """
    base = (Path(raiz) if raiz is not None else RAIZ).resolve()
    estado = _estado_do_processo()
    with estado["trava"]:
        if estado["situacao"] == FALHA:
            raise RevisaoIncoerenteError(
                "A troca de revisao anterior falhou neste processo. " + _ORIENTACAO_REBOOT
            )

        revisao = revisao_atual(base)
        if revisao is None:
            return _sem_identidade(estado, base)

        if estado["situacao"] == PRONTO and estado["revisao"] == revisao:
            return []

        if estado["situacao"] == NOVO and not _modulos_de_revisao_anterior(base):
            # Processo recem-iniciado: nada da revisao anterior foi importado.
            estado["situacao"] = PRONTO
            estado["revisao"] = revisao
            return []

        return _transicionar(estado, base, revisao)
