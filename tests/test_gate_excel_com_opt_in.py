# -*- coding: utf-8 -*-
"""TEST-P2B1 — trava estrutural do opt-in de Excel COM (RUN_EXCEL_INTEGRATION).

Garante, por leitura estatica do CODIGO-FONTE (AST, sem abrir Excel), que nos
9 arquivos que a TEST-P1 identificou como "COM incondicional" toda funcao que
chama Dispatch/DispatchEx so pode ser alcancada por um caminho protegido por
`RUN_EXCEL_INTEGRATION`:

  - a propria funcao (ou uma ancestral no grafo de chamadas/dependencia de
    fixture) tem `@pytest.mark.skipif(...RUN_EXCEL_INTEGRATION...)`; ou
  - a propria funcao (ou uma ancestral) faz `pytest.skip(...)` condicionado a
    `RUN_EXCEL_INTEGRATION` como guarda incondicional.

Duas direcoes de alcancabilidade sao consideradas, porque os 9 arquivos usam
ambos os padroes:
  (a) DEPENDENCIA DE FIXTURE — uma funcao que recebe outra fixture do mesmo
      arquivo como parametro esta protegida se essa fixture estiver protegida
      (ex.: `cenario(real_temporal)` herda o skip de `real_temporal`);
  (b) CHAMADA DE HELPER — uma funcao/fixture auxiliar chamada literalmente
      (ex.: `_b8_com(...)`, `_abrir(excel, ...)`) esta protegida se TODOS os
      pontos que a consomem (chamada literal ou injecao via parametro de
      fixture) estiverem protegidos.

Protege contra a regressao futura de alguem remover o gate e o Excel voltar a
abrir automaticamente sob a suite local.
"""
from __future__ import annotations

import ast
from pathlib import Path

import pytest

TESTS_DIR = Path(__file__).resolve().parent

ARQUIVOS_COM_GATE_OBRIGATORIO = (
    "test_441_base_fisica_c0_temporal.py",
    "test_ciclo_em_execucao_protecao.py",
    "test_cobertura_temporal_ciclo.py",
    "test_hotfix_resultados_retro_vta.py",
    "test_integridade_template_xlsx.py",
    "test_temporalidade_aditivos_data_efeito.py",
    "test_vta_c2_consumido_canonico.py",
    "test_vta_m2_financeiro_condicional.py",
    "test_vta_posicoes_resultados.py",
)


def _func_defs(tree: ast.AST) -> dict[str, ast.FunctionDef]:
    return {n.name: n for n in ast.walk(tree) if isinstance(n, ast.FunctionDef)}


def _params(node: ast.FunctionDef) -> list[str]:
    return [a.arg for a in node.args.args]


def _node_source(node: ast.FunctionDef, lines: list[str]) -> str:
    return "\n".join(lines[node.lineno - 1: node.end_lineno])


def _calls_dispatch_directly(node: ast.FunctionDef) -> bool:
    for sub in ast.walk(node):
        if isinstance(sub, ast.Call):
            f = sub.func
            name = f.attr if isinstance(f, ast.Attribute) else (
                f.id if isinstance(f, ast.Name) else None
            )
            if name in ("Dispatch", "DispatchEx"):
                return True
    return False


def _is_skipif_gated(node: ast.FunctionDef) -> bool:
    for dec in node.decorator_list:
        try:
            texto = ast.unparse(dec)
        except Exception:  # pragma: no cover - so em Python muito antigo
            texto = ""
        if "skipif" in texto and "RUN_EXCEL_INTEGRATION" in texto:
            return True
    return False


def _has_own_unconditional_skip(node: ast.FunctionDef, lines: list[str]) -> bool:
    """Guarda `if ...RUN_EXCEL_INTEGRATION...: pytest.skip(...)` no proprio
    corpo, ANTES de qualquer outra coisa relevante — cobre fixtures-raiz que
    delegam o Excel COM a um `tools.*` externo (nao literalmente `Dispatch(`
    neste arquivo)."""
    corpo = _node_source(node, lines)
    pos_gate = corpo.find("RUN_EXCEL_INTEGRATION")
    pos_skip = corpo.find("pytest.skip(")
    return pos_gate != -1 and pos_skip != -1 and pos_gate < pos_skip


def _literal_callers(nome: str, func_defs: dict, lines: list[str]) -> set[str]:
    alvo = f"{nome}("
    achados = set()
    for fname, node in func_defs.items():
        if fname == nome:
            continue
        if alvo in _node_source(node, lines):
            achados.add(fname)
    return achados


def _param_consumers(nome: str, func_defs: dict) -> set[str]:
    return {
        fname for fname, node in func_defs.items()
        if fname != nome and nome in _params(node)
    }


def _is_protected(fname: str, func_defs: dict, lines: list[str], seen: set) -> bool:
    if fname in seen:
        return False  # corta ciclo — nao alcancavel por essa via
    seen = seen | {fname}
    node = func_defs.get(fname)
    if node is None:
        return False  # nome externo ao arquivo (ex.: fixture do conftest)

    if _is_skipif_gated(node) or _has_own_unconditional_skip(node, lines):
        return True

    # (a) dependencia de fixture: qualquer parametro que seja outra funcao/
    # fixture deste arquivo e ESSA fixture estiver protegida, basta.
    for p in _params(node):
        dep = func_defs.get(p)
        if dep is not None and _is_protected(p, func_defs, lines, seen):
            return True

    # (b) consumidores (chamada literal OU injecao via parametro): todos
    # precisam estar protegidos, senao existe um caminho desprotegido.
    consumidores = (
        _literal_callers(fname, func_defs, lines)
        | _param_consumers(fname, func_defs)
    )
    if consumidores and all(
        _is_protected(c, func_defs, lines, seen) for c in consumidores
    ):
        return True

    return False


@pytest.mark.parametrize("nome_arquivo", ARQUIVOS_COM_GATE_OBRIGATORIO)
def test_gate_run_excel_integration_protege_todo_dispatch(nome_arquivo):
    caminho = TESTS_DIR / nome_arquivo
    fonte = caminho.read_text(encoding="utf-8")
    lines = fonte.splitlines()
    tree = ast.parse(fonte)
    func_defs = _func_defs(tree)

    assert "RUN_EXCEL_INTEGRATION" in fonte, (
        f"{nome_arquivo}: gate RUN_EXCEL_INTEGRATION ausente do arquivo inteiro"
    )

    diretos = [
        fname for fname, node in func_defs.items()
        if _calls_dispatch_directly(node)
    ]
    assert diretos, (
        f"{nome_arquivo}: nenhuma chamada Dispatch/DispatchEx encontrada — "
        f"lista de arquivos desatualizada ou refatoracao mudou o padrao de COM"
    )

    desprotegidos = [
        fname for fname in diretos
        if not _is_protected(fname, func_defs, lines, seen=set())
    ]
    assert not desprotegidos, (
        f"{nome_arquivo}: Dispatch/DispatchEx alcancavel sem gate "
        f"RUN_EXCEL_INTEGRATION em: {desprotegidos}"
    )


def test_lista_de_arquivos_cobre_exatamente_os_9_da_test_p1():
    """Trava a propria lista: se um arquivo for renomeado/removido, o teste
    acima para de cobri-lo silenciosamente — este teste acusa a divergencia."""
    existentes = {
        p.name for p in TESTS_DIR.glob("test_*.py")
        if p.name != "test_gate_excel_com_opt_in.py"
    }
    faltando = set(ARQUIVOS_COM_GATE_OBRIGATORIO) - existentes
    assert not faltando, f"arquivos da lista nao encontrados em tests/: {faltando}"
    assert len(ARQUIVOS_COM_GATE_OBRIGATORIO) == 9
