"""Guarda de versionamento obrigatorio (tools/verificar_versionamento.py).

Cenarios A-H da regra (docs/VERSIONAMENTO.md), excecao exata por AST e uma
prova ponta a ponta num repositorio git temporario (o mesmo comando do CI).
"""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "tools"))

import verificar_versionamento as vv  # noqa: E402

BASE = vv.Versoes("11.7", "11.3", ("11.0", "11.1", "11.2", "11.3"), "05/10/2026 16:53")
SO_CL8US = vv.Versoes("11.8", "11.3", BASE.aceitas, "06/10/2026 10:00")
CL8US_E_COLETA = vv.Versoes(
    "11.8", "11.4", BASE.aceitas + ("11.4",), "06/10/2026 10:00"
)


def _tem(erros: list[str], mensagem: str) -> bool:
    return any(e.startswith(mensagem) for e in erros)


# --------------------------------------------------------------- cenarios A-H
def test_a_somente_testes_nao_exige_bump():
    assert vv.avaliar(["tests/test_x.py", "tests/fixtures/a.xlsx"], BASE, BASE) == []


def test_b_pagina_exige_cl8us():
    erros = vv.avaliar(["pages/02_Calculo_Represados.py"], BASE, BASE)
    assert _tem(erros, vv.MSG_CL8US)
    assert not _tem(erros, vv.MSG_COLETA)
    assert vv.avaliar(["pages/02_Calculo_Represados.py"], BASE, SO_CL8US) == []


@pytest.mark.parametrize("arquivo", [
    "templates/COLETA_REAJUSTE_OFICIAL.xlsx",   # C
    "_memoria_calculo.py",                      # D
    "_gerador_masterfile.py",                   # E
    "_coleta_oficial.py",
    "_ciclo_em_execucao.py",
    "_apresentacao_pc_xls.py",
])
def test_c_d_e_superficie_da_coleta_exige_as_duas_versoes(arquivo):
    sem_bump = vv.avaliar([arquivo], BASE, BASE)
    assert _tem(sem_bump, vv.MSG_CL8US) and _tem(sem_bump, vv.MSG_COLETA)
    so_cl8us = vv.avaliar([arquivo], BASE, SO_CL8US)
    assert _tem(so_cl8us, vv.MSG_COLETA) and not _tem(so_cl8us, vv.MSG_CL8US)
    assert vv.avaliar([arquivo], BASE, CL8US_E_COLETA) == []


def test_f_versoes_incrementadas_passam():
    alterados = ["app.py", "pages/05_Garantia.py", "_memoria_calculo.py",
                 "templates/COLETA_REAJUSTE_OFICIAL.xlsx", "tests/test_y.py"]
    assert vv.avaliar(alterados, BASE, CL8US_E_COLETA) == []


def test_g_coleta_incrementada_fora_das_aceitas_falha():
    atual = vv.Versoes("11.8", "11.4", BASE.aceitas, "06/10/2026 10:00")
    erros = vv.avaliar(["templates/COLETA_REAJUSTE_OFICIAL.xlsx"], BASE, atual)
    assert erros == [vv.MSG_ACEITAS]


@pytest.mark.parametrize("arquivo", [
    "tools/aplicar_ajustes_xls_ux_pos174.py",
    "tools/ab_ajustes_xls_ux_pos174.py",
    "docs/VERSIONAMENTO.md",
    ".github/workflows/ci-pr.yml",
    "README.txt",
    "teste_numero_pc_oficial.py",
    "atualizar_homologacao_timeline_data_base.bat",
    "icti.csv",
])
def test_h_ferramenta_interna_ou_documentacao_sem_falso_positivo(arquivo):
    assert vv.avaliar([arquivo], BASE, BASE) == []


# ------------------------------------------------------------ regras extras
def test_dependencias_de_producao_exigem_cl8us():
    assert _tem(vv.avaliar(["requirements.txt"], BASE, BASE), vv.MSG_CL8US)


def test_versao_que_retrocede_ou_fora_do_formato_falha():
    volta = vv.Versoes("11.6", "11.3", BASE.aceitas, "06/10/2026 10:00")
    assert any("nao avanca" in e for e in vv.avaliar(["app.py"], BASE, volta))
    torto = vv.Versoes("11.8.1", "11.3", BASE.aceitas, "06/10/2026 10:00")
    assert any("formato XX.X" in e for e in vv.avaliar(["app.py"], BASE, torto))


@pytest.mark.parametrize("fallback", [
    "05/10/2026 16:53",   # igual ao da base
    "01/10/2026 08:00",   # anterior ao da base
    "2026-10-06 10:00",   # fora do formato
])
def test_bump_sem_atualizar_fallback_falha(fallback):
    atual = vv.Versoes("11.8", "11.3", BASE.aceitas, fallback)
    assert vv.avaliar(["app.py"], BASE, atual) == [vv.MSG_FALLBACK]


def test_alteracao_so_de_comentario_nao_exige_bump():
    antes = 'def f(x):\n    """Doc."""\n    return x + 1\n'
    depois = '# comentario novo\ndef f(x):\n    """Outra doc."""\n\n    return  x + 1  # ok\n'
    assert vv.mesma_semantica_python(antes, depois)
    assert vv.avaliar(["_reajuste_utils.py"], BASE, BASE, {"_reajuste_utils.py"}) == []
    assert not vv.mesma_semantica_python(antes, depois.replace("+ 1", "+ 2"))
    assert not vv.mesma_semantica_python(None, depois)  # arquivo novo


def test_versao_vigente_do_repositorio_obedece_a_politica():
    atual = vv.ler_versoes((RAIZ / "_versao.py").read_text(encoding="utf-8"))
    assert vv.FORMATO_VERSAO.fullmatch(atual.cl8us)
    assert vv.FORMATO_VERSAO.fullmatch(atual.coleta)
    assert atual.coleta in atual.aceitas
    assert vv.avaliar([], atual, atual) == []


# ----------------------------------------------------- ponta a ponta (git)
def _git(repo: Path, *args: str) -> None:
    subprocess.run(["git", *args], cwd=repo, check=True, capture_output=True)


def _versao_py(cl8us: str, coleta: str, aceitas: tuple[str, ...], fallback: str) -> str:
    return (
        f'CL8US_VERSION = "{cl8us}"\nCOLETA_VERSION = "{coleta}"\n'
        f"COLETA_VERSOES_ACEITAS = {aceitas!r}\n"
        f'ATUALIZADO_EM_FALLBACK = "{fallback}"\n'
    )


def _rodar(repo: Path) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(RAIZ / "tools" / "verificar_versionamento.py"),
         "--base", "HEAD^1"],
        cwd=repo, capture_output=True, text=True, encoding="utf-8",
    )


@pytest.mark.skipif(shutil.which("git") is None, reason="git indisponivel")
def test_comando_do_ci_falha_sem_bump_e_passa_com_bump(tmp_path):
    repo = tmp_path / "repo"
    (repo / "pages").mkdir(parents=True)
    _git(repo, "init", "-q")
    _git(repo, "config", "user.email", "ci@exemplo")
    _git(repo, "config", "user.name", "ci")
    (repo / "_versao.py").write_text(_versao_py(*BASE.__dict__.values()), encoding="utf-8")
    (repo / "pages" / "p.py").write_text("x = 1\n", encoding="utf-8")
    _git(repo, "add", ".")
    _git(repo, "commit", "-q", "-m", "base")

    # Sem bump: falha com a mensagem da regra.
    (repo / "pages" / "p.py").write_text("x = 2\n", encoding="utf-8")
    _git(repo, "commit", "-qam", "muda pagina")
    sem_bump = _rodar(repo)
    assert sem_bump.returncode == 1
    assert vv.MSG_CL8US in sem_bump.stdout

    # Com bump deliberado no mesmo commit: passa.
    _git(repo, "reset", "-q", "--soft", "HEAD^1")
    (repo / "_versao.py").write_text(
        _versao_py(*SO_CL8US.__dict__.values()), encoding="utf-8"
    )
    _git(repo, "commit", "-qam", "muda pagina + bump")
    com_bump = _rodar(repo)
    assert com_bump.returncode == 0, com_bump.stdout
    assert "Guarda de versionamento: OK" in com_bump.stdout
