"""DEPLOY-STALE-1: uma execucao usa uma revisao coerente do codigo local.

O Streamlit Community Cloud atualiza os arquivos do checkout sem reiniciar o
processo. As paginas voltam a ser compiladas do disco, mas os modulos locais
seguem em ``sys.modules`` com os objetos da revisao anterior.

Estes testes provam o defeito, a deteccao da troca de revisao e o alinhamento,
sem depender do Streamlit Cloud e sem depender de um modulo especifico.
"""
from __future__ import annotations

import subprocess
import sys
import threading
from pathlib import Path

import pytest

import _runtime_revision as rr


RAIZ = Path(__file__).resolve().parents[1]

SENTINELA = "_painel1_revisao_carregada"


# --------------------------------------------------------------------------
# infraestrutura: um checkout sintetico com duas revisoes
# --------------------------------------------------------------------------

MODULO_A = "FUNCAO_EXISTENTE = 'A'\n"
MODULO_B = "FUNCAO_EXISTENTE = 'B'\nFUNCAO_NOVA = 'nova'\n"

CONSUMIDOR_A = "from modulo_auxiliar import FUNCAO_EXISTENTE\nRESULTADO = FUNCAO_EXISTENTE\n"
CONSUMIDOR_B = "from modulo_auxiliar import FUNCAO_NOVA\nRESULTADO = FUNCAO_NOVA\n"


def _gravar_sha(raiz: Path, sha: str) -> None:
    git = raiz / ".git"
    (git / "refs" / "heads").mkdir(parents=True, exist_ok=True)
    (git / "HEAD").write_text("ref: refs/heads/main\n", encoding="utf-8")
    (git / "refs" / "heads" / "main").write_text(sha + "\n", encoding="utf-8")


def _escrever_revisao(raiz: Path, sha: str, *, modulo: str, consumidor: str) -> None:
    """Materializa uma revisao: arquivos no disco + SHA no plumbing do Git."""
    (raiz / "modulo_auxiliar.py").write_text(modulo, encoding="utf-8")
    (raiz / "consumidor.py").write_text(consumidor, encoding="utf-8")
    _gravar_sha(raiz, sha)


def _executar_consumidor(raiz: Path) -> dict:
    """Equivale ao que o Streamlit faz: recompila o arquivo do disco e executa."""
    fonte = (raiz / "consumidor.py").read_text(encoding="utf-8")
    espaco: dict = {}
    exec(compile(fonte, str(raiz / "consumidor.py"), "exec"), espaco)
    return espaco


def _limpar_sentinela() -> str | None:
    anterior = getattr(sys, SENTINELA, None)
    if hasattr(sys, SENTINELA):
        delattr(sys, SENTINELA)
    return anterior


def _restaurar_sentinela(anterior: str | None) -> None:
    if anterior is None:
        if hasattr(sys, SENTINELA):
            delattr(sys, SENTINELA)
    else:
        setattr(sys, SENTINELA, anterior)


@pytest.fixture
def checkout(tmp_path, monkeypatch):
    raiz = tmp_path / "checkout"
    raiz.mkdir()
    _escrever_revisao(raiz, "a" * 40, modulo=MODULO_A, consumidor=CONSUMIDOR_A)
    monkeypatch.syspath_prepend(str(raiz))
    sinteticos = ("modulo_auxiliar", "consumidor", "outro_modulo", "pacote_local")
    for nome in sinteticos:
        sys.modules.pop(nome, None)
    anterior = _limpar_sentinela()
    yield raiz
    for nome in sinteticos:
        sys.modules.pop(nome, None)
    _restaurar_sentinela(anterior)


def _promover_para_b(raiz: Path) -> None:
    _escrever_revisao(raiz, "b" * 40, modulo=MODULO_B, consumidor=CONSUMIDOR_B)


# --------------------------------------------------------------------------
# 1. o defeito existe (contrato-base da frente)
# --------------------------------------------------------------------------

def test_sem_blindagem_o_processo_produz_arvore_hibrida(checkout):
    rr.garantir_revisao_coerente(checkout)          # primeiro boot na revisao A
    assert _executar_consumidor(checkout)["RESULTADO"] == "A"

    _promover_para_b(checkout)                      # deploy, processo vivo

    with pytest.raises(ImportError) as erro:
        _executar_consumidor(checkout)              # SEM chamar a blindagem
    assert "FUNCAO_NOVA" in str(erro.value)


# --------------------------------------------------------------------------
# 2. teste principal de regressao: A -> B com a blindagem
# --------------------------------------------------------------------------

def test_troca_de_revisao_entrega_o_modulo_novo(checkout):
    rr.garantir_revisao_coerente(checkout)
    assert _executar_consumidor(checkout)["RESULTADO"] == "A"

    _promover_para_b(checkout)

    removidos = rr.garantir_revisao_coerente(checkout)
    assert "modulo_auxiliar" in removidos

    espaco = _executar_consumidor(checkout)         # nenhum ImportError
    assert espaco["RESULTADO"] == "nova"
    assert sys.modules["modulo_auxiliar"].FUNCAO_EXISTENTE == "B"


def test_apos_a_troca_as_execucoes_seguintes_ficam_estaveis(checkout):
    rr.garantir_revisao_coerente(checkout)
    _executar_consumidor(checkout)
    _promover_para_b(checkout)
    assert rr.garantir_revisao_coerente(checkout)   # dispara uma vez

    _executar_consumidor(checkout)
    referencia = sys.modules["modulo_auxiliar"]

    for _ in range(3):
        assert rr.garantir_revisao_coerente(checkout) == []
        assert _executar_consumidor(checkout)["RESULTADO"] == "nova"
        assert sys.modules["modulo_auxiliar"] is referencia


# --------------------------------------------------------------------------
# 3. mesma revisao e primeiro boot nao mexem em nada
# --------------------------------------------------------------------------

def test_mesma_revisao_nao_recarrega_nada(checkout):
    rr.garantir_revisao_coerente(checkout)
    _executar_consumidor(checkout)
    referencia = sys.modules["modulo_auxiliar"]

    for _ in range(5):
        assert rr.garantir_revisao_coerente(checkout) == []
        assert sys.modules["modulo_auxiliar"] is referencia


def test_primeiro_boot_nao_purga(checkout):
    import modulo_auxiliar  # noqa: F401  — ja carregado antes do primeiro boot

    assert rr.garantir_revisao_coerente(checkout) == []
    assert "modulo_auxiliar" in sys.modules
    assert getattr(sys, SENTINELA) == "a" * 40


def test_sem_identidade_de_revisao_nada_acontece(tmp_path):
    raiz = tmp_path / "sem_git"
    raiz.mkdir()
    (raiz / "modulo_auxiliar.py").write_text(MODULO_A, encoding="utf-8")
    assert rr.revisao_atual(raiz) is None
    assert rr.garantir_revisao_coerente(raiz) == []


# --------------------------------------------------------------------------
# 4. escopo da purga
# --------------------------------------------------------------------------

def test_terceiros_nunca_sao_removidos(checkout):
    import openpyxl
    import pandas
    import streamlit

    rr.garantir_revisao_coerente(checkout)
    _executar_consumidor(checkout)
    _promover_para_b(checkout)

    removidos = rr.garantir_revisao_coerente(checkout)

    externos = {"streamlit", "pandas", "openpyxl", "sys", "os", "pytest"}
    assert not [n for n in removidos if n.split(".")[0] in externos]
    assert sys.modules["streamlit"] is streamlit
    assert sys.modules["pandas"] is pandas
    assert sys.modules["openpyxl"] is openpyxl


def test_purga_alcanca_varios_modulos_locais(checkout):
    (checkout / "outro_modulo.py").write_text("VALOR = 1\n", encoding="utf-8")
    (checkout / "pacote_local").mkdir()
    (checkout / "pacote_local" / "__init__.py").write_text("VALOR = 1\n", encoding="utf-8")

    rr.garantir_revisao_coerente(checkout)
    import outro_modulo  # noqa: F401
    import pacote_local  # noqa: F401
    _executar_consumidor(checkout)

    _promover_para_b(checkout)
    removidos = set(rr.garantir_revisao_coerente(checkout))

    assert {"modulo_auxiliar", "outro_modulo", "pacote_local"} <= removidos


def test_modulos_de_venv_dentro_do_checkout_ficam_de_fora(checkout):
    venv = checkout / ".venv" / "Lib" / "site-packages"
    venv.mkdir(parents=True)
    (venv / "biblioteca_de_terceiro.py").write_text("VALOR = 1\n", encoding="utf-8")
    sys.path.insert(0, str(venv))
    try:
        import biblioteca_de_terceiro  # noqa: F401
        assert "biblioteca_de_terceiro" not in rr.modulos_locais_carregados(checkout)
    finally:
        sys.path.remove(str(venv))
        sys.modules.pop("biblioteca_de_terceiro", None)


# --------------------------------------------------------------------------
# 5. identidade da revisao
# --------------------------------------------------------------------------

def test_revisao_e_o_sha_do_checkout_e_nao_um_timestamp():
    sha = rr.revisao_atual(RAIZ)
    assert sha is not None
    assert len(sha) == 40 and all(c in "0123456789abcdef" for c in sha)

    esperado = subprocess.run(
        ["git", "rev-parse", "HEAD"],
        cwd=RAIZ, capture_output=True, text=True, check=True,
    ).stdout.strip()
    assert sha == esperado


def test_revisao_le_ref_empacotada(tmp_path):
    raiz = tmp_path / "empacotado"
    (raiz / ".git").mkdir(parents=True)
    (raiz / ".git" / "HEAD").write_text("ref: refs/heads/main\n", encoding="utf-8")
    (raiz / ".git" / "packed-refs").write_text(
        "# pack-refs with: peeled fully-peeled sorted\n"
        + "c" * 40 + " refs/heads/main\n",
        encoding="utf-8",
    )
    assert rr.revisao_atual(raiz) == "c" * 40


def test_revisao_le_head_destacado(tmp_path):
    raiz = tmp_path / "destacado"
    (raiz / ".git").mkdir(parents=True)
    (raiz / ".git" / "HEAD").write_text("d" * 40 + "\n", encoding="utf-8")
    assert rr.revisao_atual(raiz) == "d" * 40


# --------------------------------------------------------------------------
# 6. sentinela: metadado tecnico, fora do modulo purgado
# --------------------------------------------------------------------------

def test_sentinela_sobrevive_a_purga_do_proprio_modulo(checkout):
    rr.garantir_revisao_coerente(checkout)
    _promover_para_b(checkout)
    rr.garantir_revisao_coerente(checkout)
    assert getattr(sys, SENTINELA) == "b" * 40


def test_sentinela_guarda_somente_o_sha(checkout):
    rr.garantir_revisao_coerente(checkout)
    valor = getattr(sys, SENTINELA)
    assert isinstance(valor, str) and len(valor) == 40


# --------------------------------------------------------------------------
# 7. concorrencia: a purga acontece uma unica vez por troca de revisao
# --------------------------------------------------------------------------

def test_sessoes_concorrentes_purgam_uma_unica_vez(checkout):
    rr.garantir_revisao_coerente(checkout)
    _executar_consumidor(checkout)
    _promover_para_b(checkout)

    resultados: list[list[str]] = []
    trava = threading.Lock()

    def sessao():
        removidos = rr.garantir_revisao_coerente(checkout)
        with trava:
            resultados.append(removidos)

    fios = [threading.Thread(target=sessao) for _ in range(8)]
    for f in fios:
        f.start()
    for f in fios:
        f.join()

    assert sum(1 for r in resultados if r) == 1


# --------------------------------------------------------------------------
# 8. o incidente real do PR #130, com os bytes das duas revisoes
# --------------------------------------------------------------------------

SIMBOLOS_130 = (
    "APLICAR_VARIACAO_NEGATIVA",
    "NEUTRALIZAR_VARIACAO_NEGATIVA",
    "resolver_tratamento_variacao_negativa",
    "situacao_com_tratamento_variacao_negativa",
)

BLOCO_130 = (
    "from _reajuste_utils import (\n"
    + "".join("    " + s + ",\n" for s in SIMBOLOS_130)
    + ")\n"
)

REVISAO_ANTES_130 = "e72dc04"
REVISAO_DEPOIS_130 = "161bda9"


def _bytes_de(revisao: str, caminho: str) -> str:
    return subprocess.run(
        ["git", "show", revisao + ":" + caminho],
        cwd=RAIZ, capture_output=True, text=True, check=True,
    ).stdout


def _historico_disponivel() -> bool:
    for revisao in (REVISAO_ANTES_130, REVISAO_DEPOIS_130):
        concluido = subprocess.run(
            ["git", "cat-file", "-e", revisao + "^{commit}"],
            cwd=RAIZ, capture_output=True, text=True,
        )
        if concluido.returncode != 0:
            return False
    return True


historico_real = pytest.mark.skipif(
    not _historico_disponivel(),
    reason="clone raso: as revisoes e72dc04/161bda9 do incidente #130 nao estao presentes",
)


@pytest.fixture
def checkout_130(tmp_path, monkeypatch):
    raiz = tmp_path / "painel1"
    raiz.mkdir()
    _gravar_sha(raiz, "a" * 40)
    monkeypatch.syspath_prepend(str(raiz))
    salvo = sys.modules.pop("_reajuste_utils", None)
    anterior = _limpar_sentinela()
    yield raiz
    sys.modules.pop("_reajuste_utils", None)
    if salvo is not None:
        sys.modules["_reajuste_utils"] = salvo
    _restaurar_sentinela(anterior)


def _implantar_revisao_antes(raiz: Path) -> None:
    (raiz / "_reajuste_utils.py").write_text(
        _bytes_de(REVISAO_ANTES_130, "_reajuste_utils.py"), encoding="utf-8")


def _implantar_revisao_depois(raiz: Path) -> None:
    """Promove o checkout sintetico de e72dc04 para 161bda9."""
    _gravar_sha(raiz, "b" * 40)
    (raiz / "_reajuste_utils.py").write_text(
        _bytes_de(REVISAO_DEPOIS_130, "_reajuste_utils.py"), encoding="utf-8")


@historico_real
def test_incidente_130_reproduz_sem_a_blindagem(checkout_130):
    _implantar_revisao_antes(checkout_130)
    rr.garantir_revisao_coerente(checkout_130)
    import _reajuste_utils  # noqa: F401

    _implantar_revisao_depois(checkout_130)

    with pytest.raises(ImportError) as erro:
        exec(compile(BLOCO_130, "pages/01_Calculo_Simples.py", "exec"), {})
    assert "APLICAR_VARIACAO_NEGATIVA" in str(erro.value)


@historico_real
def test_incidente_130_nao_ocorre_com_a_blindagem(checkout_130):
    _implantar_revisao_antes(checkout_130)
    rr.garantir_revisao_coerente(checkout_130)
    import _reajuste_utils  # noqa: F401

    _implantar_revisao_depois(checkout_130)

    assert "_reajuste_utils" in rr.garantir_revisao_coerente(checkout_130)

    espaco: dict = {}
    exec(compile(BLOCO_130, "pages/01_Calculo_Simples.py", "exec"), espaco)
    for simbolo in SIMBOLOS_130:
        assert simbolo in espaco


@historico_real
def test_import_com_alias_nao_resolveria_o_incidente(checkout_130):
    """Alternativa E: `import modulo as alias` continua entregando o objeto stale."""
    _implantar_revisao_antes(checkout_130)
    rr.garantir_revisao_coerente(checkout_130)
    import _reajuste_utils as antigo

    _implantar_revisao_depois(checkout_130)

    import _reajuste_utils as depois
    assert depois is antigo
    assert not hasattr(depois, "APLICAR_VARIACAO_NEGATIVA")


# --------------------------------------------------------------------------
# 9. o app entra pela blindagem antes de qualquer import local
# --------------------------------------------------------------------------

def test_app_chama_a_blindagem_antes_dos_imports_locais():
    linhas = (RAIZ / "app.py").read_text(encoding="utf-8").splitlines()
    guarda = next(i for i, l in enumerate(linhas) if "garantir_revisao_coerente()" in l)
    locais = [
        i for i, l in enumerate(linhas)
        if l.startswith(("from _", "import _")) and "_runtime_revision" not in l
    ]
    assert locais, "app.py deveria importar modulos locais"
    assert guarda < min(locais)


def test_blindagem_nao_toca_session_state_nem_caches():
    fonte = (RAIZ / "_runtime_revision.py").read_text(encoding="utf-8")
    for proibido in ("session_state", "cache_data", "cache_resource", "importlib", "streamlit"):
        assert proibido not in fonte
