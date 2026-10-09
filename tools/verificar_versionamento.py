# -*- coding: utf-8 -*-
"""Guarda de versionamento obrigatorio (regra em docs/VERSIONAMENTO.md).

DETECTA e FALHA; nunca incrementa versao sozinha. Compara a base do PR com o
HEAD:

* arquivo de superficie USER-FACING alterado e CL8US_VERSION igual a base
  -> "Alteracao user-facing detectada sem incremento de CL8US_VERSION.";
* arquivo de superficie da COLETA alterado e COLETA_VERSION igual a base
  -> "Alteracao da Coleta detectada sem incremento de COLETA_VERSION.";
* COLETA_VERSION fora de COLETA_VERSOES_ACEITAS
  -> "COLETA_VERSION atual nao consta em COLETA_VERSOES_ACEITAS.";
* versao alterada que nao avanca, ou fora do formato XX.X;
* CL8US_VERSION incrementada sem atualizar ATUALIZADO_EM_FALLBACK, ou com
  fallback fora do formato dd/mm/aaaa HH:MM.

O fallback deve mudar em toda entrega, mas pode ser corrigido para um horario
anterior quando o valor anterior estava errado. A monotonicidade do fallback
nao e usada como prova de publicacao; quando o Git esta disponivel, o timestamp
do commit e a fonte primaria.

As superficies sao LISTAS DECLARATIVAS (abaixo), revisaveis no PR. Um .py
alterado so em comentarios/formatacao (AST identica, docstrings ignoradas) nao
conta como alteracao: e a unica excecao automatica, e e exata.

Uso no CI:  python tools/verificar_versionamento.py --base HEAD^1
"""
from __future__ import annotations

import argparse
import ast
import fnmatch
import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import PurePosixPath

ARQUIVO_VERSAO = "_versao.py"

SUPERFICIE_COLETA = (
    "templates/COLETA_REAJUSTE_OFICIAL.xlsx",
    "_coleta_oficial.py",
    "_gerador_masterfile.py",
    "_memoria_calculo.py",
    "_ciclo_em_execucao.py",
    "_apresentacao_pc_xls.py",
)

SUPERFICIE_USUARIO = (
    "app.py",
    "pages/*",
    "_*.py",
    "templates/*",
    "assets/*",
    ".streamlit/*",
) + SUPERFICIE_COLETA

FORA_DE_PRODUCAO = (
    "tests/*",
    "docs/*",
    "tools/*",
    ".github/*",
    "*.md",
    "*.txt",
    "*.bat",
    "teste_*.py",
)
SEMPRE_PRODUCAO = (
    "requirements.txt",
    "tools/atualizar_ist_anatel.py",
)

MSG_CL8US = "Alteração user-facing detectada sem incremento de CL8US_VERSION."
MSG_COLETA = "Alteração da Coleta detectada sem incremento de COLETA_VERSION."
MSG_ACEITAS = "COLETA_VERSION atual não consta em COLETA_VERSOES_ACEITAS."
MSG_FALLBACK = (
    "CL8US_VERSION incrementada sem atualizar ATUALIZADO_EM_FALLBACK "
    "(formato dd/mm/aaaa HH:MM)."
)

FORMATO_VERSAO = re.compile(r"\d{2}\.\d")
FORMATO_FALLBACK = "%d/%m/%Y %H:%M"


@dataclass(frozen=True)
class Versoes:
    cl8us: str
    coleta: str
    aceitas: tuple[str, ...]
    fallback: str


def ler_versoes(fonte: str) -> Versoes:
    valores: dict[str, object] = {}
    for no in ast.parse(fonte).body:
        if isinstance(no, ast.Assign) and len(no.targets) == 1:
            alvo = no.targets[0]
            if isinstance(alvo, ast.Name):
                try:
                    valores[alvo.id] = ast.literal_eval(no.value)
                except ValueError:
                    continue
    return Versoes(
        cl8us=str(valores["CL8US_VERSION"]),
        coleta=str(valores["COLETA_VERSION"]),
        aceitas=tuple(valores["COLETA_VERSOES_ACEITAS"]),
        fallback=str(valores["ATUALIZADO_EM_FALLBACK"]),
    )


def _casa(caminho: str, padroes: tuple[str, ...]) -> bool:
    caminho = PurePosixPath(caminho.replace("\\", "/")).as_posix()
    for padrao in padroes:
        if padrao.endswith("/*"):
            if caminho.startswith(padrao[:-1]):
                return True
        elif "/" in padrao:
            if caminho == padrao:
                return True
        elif "/" not in caminho and fnmatch.fnmatchcase(caminho, padrao):
            return True
    return False


def eh_user_facing(caminho: str) -> bool:
    if _casa(caminho, SEMPRE_PRODUCAO):
        return True
    if _casa(caminho, FORA_DE_PRODUCAO):
        return False
    return _casa(caminho, SUPERFICIE_USUARIO)


def eh_coleta(caminho: str) -> bool:
    return _casa(caminho, SUPERFICIE_COLETA)


def _sem_docstrings(arvore: ast.AST) -> ast.AST:
    for no in ast.walk(arvore):
        corpo = getattr(no, "body", None)
        if (
            isinstance(no, (ast.Module, ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef))
            and corpo
            and isinstance(corpo[0], ast.Expr)
            and isinstance(getattr(corpo[0], "value", None), ast.Constant)
            and isinstance(corpo[0].value.value, str)
        ):
            no.body = corpo[1:] or [ast.Pass()]
    return arvore


def mesma_semantica_python(antes: str | None, depois: str | None) -> bool:
    if antes is None or depois is None:
        return False
    try:
        a = ast.dump(_sem_docstrings(ast.parse(antes.lstrip("﻿"))))
        b = ast.dump(_sem_docstrings(ast.parse(depois.lstrip("﻿"))))
    except SyntaxError:
        return False
    return a == b


def _avanca(base: str, atual: str) -> bool:
    return tuple(map(int, atual.split("."))) > tuple(map(int, base.split(".")))


def avaliar(
    alterados: list[str],
    base: Versoes,
    atual: Versoes,
    equivalentes: set[str] | frozenset[str] = frozenset(),
) -> list[str]:
    relevantes = [c for c in alterados if c not in equivalentes]
    user_facing = sorted(c for c in relevantes if eh_user_facing(c))
    coleta = sorted(c for c in relevantes if eh_coleta(c))
    erros: list[str] = []

    for nome, valor in (("CL8US_VERSION", atual.cl8us), ("COLETA_VERSION", atual.coleta)):
        if not FORMATO_VERSAO.fullmatch(valor):
            erros.append(f"{nome} fora do formato XX.X: {valor!r}.")
    if erros:
        return erros

    if atual.coleta not in atual.aceitas:
        erros.append(MSG_ACEITAS)

    cl8us_mudou = atual.cl8us != base.cl8us
    coleta_mudou = atual.coleta != base.coleta
    if cl8us_mudou and not _avanca(base.cl8us, atual.cl8us):
        erros.append(f"CL8US_VERSION nao avanca: {base.cl8us} -> {atual.cl8us}.")
    if coleta_mudou and not _avanca(base.coleta, atual.coleta):
        erros.append(f"COLETA_VERSION nao avanca: {base.coleta} -> {atual.coleta}.")

    if user_facing and not cl8us_mudou:
        erros.append(f"{MSG_CL8US} Arquivos: {', '.join(user_facing)}")
    if coleta and not coleta_mudou:
        erros.append(f"{MSG_COLETA} Arquivos: {', '.join(coleta)}")

    if cl8us_mudou:
        try:
            datetime.strptime(atual.fallback, FORMATO_FALLBACK)
            datetime.strptime(base.fallback, FORMATO_FALLBACK)
            if atual.fallback == base.fallback:
                erros.append(MSG_FALLBACK)
        except ValueError:
            erros.append(MSG_FALLBACK)
    return erros


def _git(*args: str) -> str:
    return subprocess.run(
        ["git", *args], check=True, capture_output=True, text=True, encoding="utf-8"
    ).stdout


def _conteudo(ref: str, caminho: str) -> str | None:
    try:
        return _git("show", f"{ref}:{caminho}")
    except subprocess.CalledProcessError:
        return None


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    p.add_argument("--base", required=True, help="ref da base (ex.: HEAD^1 no CI)")
    p.add_argument("--head", default="HEAD")
    args = p.parse_args()

    alterados = [
        linha.strip() for linha in
        _git("diff", "--name-only", args.base, args.head).splitlines() if linha.strip()
    ]
    base = ler_versoes(_conteudo(args.base, ARQUIVO_VERSAO) or "")
    atual = ler_versoes(_conteudo(args.head, ARQUIVO_VERSAO) or "")
    equivalentes = {
        c for c in alterados
        if c.endswith(".py")
        and mesma_semantica_python(_conteudo(args.base, c), _conteudo(args.head, c))
    }
    erros = avaliar(alterados, base, atual, equivalentes)

    print(f"Base {args.base}: Cl8us {base.cl8us} / Coleta {base.coleta}")
    print(f"HEAD {args.head}: Cl8us {atual.cl8us} / Coleta {atual.coleta}")
    print(f"Arquivos alterados: {len(alterados)} "
          f"({len(equivalentes)} so com comentarios/formatacao)")
    if erros:
        for erro in erros:
            print(f"::error::{erro}")
        print("Regra: docs/VERSIONAMENTO.md (o bump e decidido pelo desenvolvedor, "
              "no mesmo PR).")
        return 1
    print("Guarda de versionamento: OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
