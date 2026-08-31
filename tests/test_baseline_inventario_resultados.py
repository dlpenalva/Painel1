"""RESULTADOS-BASELINE (PR 0) — inventario vivo dos testes acoplados a aba.

O QUE ESTE ARQUIVO FAZ
----------------------
Varre `tests/` inteiro, encontra toda referencia as abas RESULTADOS e
MEMORIA_RESULTADOS, classifica cada ocorrencia por REGRA EXPLICITA e compara o
resultado com o inventario gravado. Ele nao corrige nem remove teste nenhum: o
PR 0 existe para INVENTARIAR e provar o comportamento atual.

POR QUE UM INVENTARIO EXECUTAVEL, E NAO UMA TABELA ESCRITA A MAO
-----------------------------------------------------------------
Tabela a mao envelhece em uma semana. Este inventario e recalculado a cada
execucao: se alguem adicionar um teste que toca a aba, ou mover uma coordenada,
o vermelho aparece aqui antes de aparecer no PR 1. E o mapa que diz, para cada
acoplamento, se ele PRECISA sobreviver a refatoracao (contrato) ou se e
apresentacao que sera reescrita.

AS CINCO CLASSES (definicao operacional, aplicada por regra)
------------------------------------------------------------
A. CONTRATO LEGITIMO — cita intervalo nomeado do contrato vivo, uma das nove
   coordenadas lidas pelo runtime, ou uma API Python do consolidado.
B. TESTE DE LEIAUTE  — verifica apresentacao (cor, fonte, borda, largura,
   merge, formato, visibilidade, quebra de texto).
C. TESTE LEGADO      — endereca coordenada FORA do leiaute atual (linha > 87)
   ou um aplicador historico de template.
D. SUSPEITO          — endereca celula dentro do leiaute atual que nao e
   contrato vivo nem verificacao de apresentacao: passa hoje, mas pode estar
   conferindo celula sem consumidor.
E. OUTRO             — toca apenas MEMORIA_RESULTADOS ou cita a aba em texto,
   sem endereco.

A classificacao e HEURISTICA e assumida como tal: ela orienta a triagem do
PR 2, nao autoriza a apagar nada. Nenhum teste foi alterado por este PR.
"""
from __future__ import annotations

import ast
import json
import os
import re
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
DIRETORIO_TESTES = RAIZ / "tests"
INVENTARIO = (RAIZ / "tests" / "baseline_resultados"
              / "inventario_testes_resultados.json")
REGRAVAR = os.environ.get("RESULTADOS_BASELINE_REGRAVAR") == "1"

# Este arquivo e o harness nao entram no inventario: sao o instrumento de
# medida, nao o objeto medido.
ARQUIVOS_DO_HARNESS = {
    "test_baseline_inventario_resultados.py",
    "test_baseline_resultados.py",
    "test_baseline_resultados_goldens.py",
    "test_resultados_contract_pr1.py",
    "test_retroativo_potencial_pc.py",
    "_baseline_cenarios.py",
    "_baseline_fotografia.py",
}

# Contrato vivo: as nove coordenadas que o runtime consulta na aba executiva.
COORDENADAS_DE_CONTRATO = {
    "A1", "B3", "B10", "B11", "B12", "B13", "H10", "H11", "H13",
}

NOMES_DE_CONTRATO = {
    "METODO_RETROATIVO", "TOLERANCIA_DIVERGENCIA", "VALOR_MANUAL_RETRO",
    "JUSTIFICATIVA_RETRO", "RETRO_FIN", "RETRO_PC", "RETRO_ITENS",
    "RETRO_OFICIAL", "VTA_CALCULADO", "AJUSTE_MANUAL_VTA", "VTA_MANUAL_OFICIAL",
    "VTA_FINAL", "QTD_REM_OFICIAL", "REM_BASE_OFICIAL", "REM_ATUALIZADO_OFICIAL",
    "STATUS_RESULTADOS", "VTA_ATUALIZACAO_CHEIA", "EXECUCAO_ATUALIZADA_CICLO",
    "SALDO_REMANESCENTE_ATUAL", "OPCOES_APLICAR_MANUAL",
}

APIS_PYTHON = (
    "status_resultados", "referencias_vta", "resultados_xls",
    "reconciliacao_xls_python", "resultado_consolidado",
    "montar_resultado_consolidado",
)

MARCADORES_DE_LEIAUTE = (
    "fill", "font", "border", "alignment", "width", "height", "merge",
    "number_format", "numfmt", "hidden", "sheet_state", "color", "rgb", "dxf",
    "conditional", "freeze", "wrap_text", "shrink", "indent", "tabcolor",
    "column_dimensions", "row_dimensions", "visib", "oculta", "largura",
)

MARCADORES_DE_APLICADOR = ("tools/aplicar", "tools\\aplicar", "aplicar_")

_RE_MEMORIA = re.compile(r"MEMORIA_RESULTADOS")
_RE_ABA = re.compile(r"(?<![A-Z_])RESULTADOS(?![A-Z_])")
_RE_COORD = re.compile(r"\b([A-Z]{1,2})(\d{1,4})\b")
_ULTIMA_LINHA_DO_LEIAUTE = 87


def _funcao_da_linha(funcoes, numero: int) -> str:
    escolhida = None
    for inicio, fim, nome in funcoes:
        if inicio <= numero <= fim and (escolhida is None or inicio > escolhida[0]):
            escolhida = (inicio, nome)
    return escolhida[1] if escolhida else "<modulo>"


def _classificar(janela: str, coordenadas: list[str],
                 so_memoria: bool) -> tuple[str, str]:
    """Aplica as regras na ordem em que a especificacao do PR 0 as define."""
    if any(nome in janela for nome in NOMES_DE_CONTRATO) or \
            any(api in janela for api in APIS_PYTHON):
        return "A", "cita intervalo nomeado do contrato vivo ou API do consolidado"
    if not so_memoria and any(coord in COORDENADAS_DE_CONTRATO
                              for coord in coordenadas):
        return "A", "endereca coordenada lida pelo runtime de producao"
    if any(marcador in janela.lower() for marcador in MARCADORES_DE_LEIAUTE):
        return "B", "verifica apresentacao (formato, cor, geometria ou visibilidade)"
    if any(marcador in janela for marcador in MARCADORES_DE_APLICADOR):
        return "C", "referencia aplicador historico de template"
    fora_do_leiaute = [
        coord for coord in coordenadas
        if int(_RE_COORD.match(coord).group(2)) > _ULTIMA_LINHA_DO_LEIAUTE
    ]
    if fora_do_leiaute and not so_memoria:
        return "C", ("endereca coordenada fora do leiaute atual (A1:J87): "
                     + ", ".join(sorted(fora_do_leiaute)))
    if so_memoria:
        return "E", "toca apenas MEMORIA_RESULTADOS (aba tecnica, nao a executiva)"
    if not coordenadas:
        return "E", "cita a aba em texto, sem endereco"
    return "D", ("endereca celula do leiaute atual que nao e contrato vivo "
                 "nem verificacao de apresentacao")


def levantar_inventario() -> list[dict]:
    """Recalcula o inventario a partir do codigo, sempre do zero."""
    registros: list[dict] = []
    for caminho in sorted(DIRETORIO_TESTES.glob("*.py")):
        if caminho.name in ARQUIVOS_DO_HARNESS:
            continue
        fonte = caminho.read_text(encoding="utf-8")
        if "RESULTADOS" not in fonte:
            continue
        linhas = fonte.splitlines()
        funcoes = [
            (no.lineno, getattr(no, "end_lineno", no.lineno), no.name)
            for no in ast.walk(ast.parse(fonte))
            if isinstance(no, (ast.FunctionDef, ast.AsyncFunctionDef))
        ]
        for numero, linha in enumerate(linhas, start=1):
            tem_memoria = bool(_RE_MEMORIA.search(linha))
            tem_aba = bool(_RE_ABA.search(_RE_MEMORIA.sub("", linha)))
            if not (tem_memoria or tem_aba):
                continue
            janela = " ".join(linhas[max(0, numero - 2):numero + 2])
            coordenadas = sorted({m.group(0) for m in _RE_COORD.finditer(janela)})
            so_memoria = tem_memoria and not tem_aba
            classe, justificativa = _classificar(janela, coordenadas, so_memoria)
            registros.append({
                "arquivo": f"tests/{caminho.name}",
                "linha": numero,
                "teste": _funcao_da_linha(funcoes, numero),
                "referencia": linha.strip()[:160],
                "coordenadas": coordenadas,
                "aba": ("MEMORIA_RESULTADOS" if so_memoria
                        else "ambas" if tem_memoria else "RESULTADOS"),
                "classificacao": classe,
                "sobrevive_ao_pr1": classe == "A",
                "atualizar_no_pr2": classe in {"B", "C", "D"},
                "justificativa": justificativa,
            })
    return registros


def resumo(registros: list[dict]) -> dict:
    contagem: dict[str, int] = {}
    for registro in registros:
        classe = registro["classificacao"]
        contagem[classe] = contagem.get(classe, 0) + 1
    tocam_a_executiva = {
        registro["arquivo"] for registro in registros
        if registro["aba"] in {"RESULTADOS", "ambas"}
    }
    return {
        "arquivos": len({registro["arquivo"] for registro in registros}),
        "ocorrencias": len(registros),
        "por_classificacao": dict(sorted(contagem.items())),
        "arquivos_que_tocam_a_aba_executiva": sorted(tocam_a_executiva),
        "arquivos_so_de_memoria": sorted(
            {registro["arquivo"] for registro in registros} - tocam_a_executiva
        ),
    }


def test_inventario_dos_testes_acoplados_esta_registrado():
    """O mapa dos acoplamentos continua o mesmo — ou o diff explica por que."""
    registros = levantar_inventario()
    atual = {"resumo": resumo(registros), "ocorrencias": registros}
    if REGRAVAR or not INVENTARIO.exists():
        INVENTARIO.parent.mkdir(parents=True, exist_ok=True)
        INVENTARIO.write_text(
            json.dumps(atual, ensure_ascii=False, indent=1) + "\n", encoding="utf-8"
        )
        if not REGRAVAR:
            pytest.skip("inventario criado agora; reexecute para comparar")
        return
    esperado = json.loads(INVENTARIO.read_text(encoding="utf-8"))
    assert atual["resumo"] == esperado["resumo"], (
        "o conjunto de testes acoplados a aba RESULTADOS mudou; regrave o "
        "inventario e explique a mudanca no PR"
    )
    assert atual["ocorrencias"] == esperado["ocorrencias"]


def test_nenhum_acoplamento_de_contrato_passa_despercebido():
    """Todo teste classe A precisa sobreviver ao PR 1 — e estar listado."""
    registros = levantar_inventario()
    contrato = [registro for registro in registros
                if registro["classificacao"] == "A"]
    assert contrato, "o inventario nao encontrou nenhum acoplamento de contrato"
    assert all(registro["sobrevive_ao_pr1"] for registro in contrato)


def test_a_triagem_cobre_todas_as_ocorrencias():
    """Nenhuma ocorrencia fica sem classe: E existe justamente para isso."""
    registros = levantar_inventario()
    assert all(
        registro["classificacao"] in {"A", "B", "C", "D", "E"}
        for registro in registros
    )
