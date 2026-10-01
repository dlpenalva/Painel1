"""Compatibilidade retroativa da Coleta (Etapa 2/03).

O Cl8us 11.0 processa tres linhagens de Coleta:

* ``COLETA_11`` — marcador publico ``Modelo de Coleta = 11.0`` (estrutura atual);
* ``PRE_11_L1`` — 15/16 abas, ``parametros`` expandida, ``NUMERO_PC``, sem marcador;
* ``PRE_11_L2`` — 11 abas, ``parametros`` enxuta, ``NUMERO_PC``, sem marcador.

Nenhum numero de versao e inventado para arquivo antigo. O que se prova aqui:

1. a deteccao e deterministica e nao confunde as linhagens (nem usa nome/data);
2. entradas do fiscal nao sao alteradas pela adaptacao;
3. so DERIVADOS sao recompostos, e a formula reproduzida coincide com o cache do
   Excel do proprio arquivo (senao o bloco nao e adaptado);
4. o arquivo antigo adaptado e EQUIVALENTE a mesma Coleta regerada no Excel;
5. L1 == 11.0 nos tres metodos (Financeiro, PC, Itens Consumidos); L2 == 11.0
   onde a linhagem tem a informacao, e fail-closed onde nao tem (PC sem
   INICIO_EFEITO_FINANCEIRO);
6. nenhum gate documental foi removido: Coleta com precisao antiga segue
   bloqueada para formalizacao, e a divergencia XLS x Python fica exposta.

As fixtures sao Coletas SINTETICAS recalculadas no Excel real
(`tools/construir_fixtures_compat_coleta.py`); o MANIFESTO guarda o SHA-256.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import io
import json
import math
import re
import warnings
from pathlib import Path

import pytest
from openpyxl import load_workbook

import _compatibilidade_coleta as cc
from _compatibilidade_valores import (
    ATRIBUTO_AUDITORIA,
    MENSAGEM_L2_ADITIVO,
    aplicar_compatibilidade_valores,
    restricoes_compatibilidade,
)
from _leitor_masterfile_v10 import ler_masterfile_v10
from _politica_entrega_segura import MENSAGEM_COLETA_PRECISAO_ANTERIOR

warnings.filterwarnings("ignore")

ROOT = Path(__file__).resolve().parents[1]
PASTA = Path(__file__).resolve().parent / "fixtures" / "coletas_compat"
MANIFESTO = json.loads((PASTA / "MANIFESTO.json").read_text(encoding="utf-8"))

METODOS = ("financeiro", "pc", "consumidos")
LINHAGEM_DE = {"coleta_11": cc.LINHAGEM_COLETA_11, "pre11_l1": cc.LINHAGEM_PRE_11_L1,
               "pre11_l2": cc.LINHAGEM_PRE_11_L2}


def _caminho(nome: str) -> Path:
    return PASTA / f"{nome}.xlsx"


def _bytes(nome: str) -> bytes:
    return _caminho(nome).read_bytes()


# --------------------------------------------------------------------------- #
# 0. Integridade das fixtures.
# --------------------------------------------------------------------------- #
@pytest.mark.parametrize("nome", sorted(MANIFESTO))
def test_fixture_corresponde_ao_manifesto(nome):
    info = MANIFESTO[nome]
    conteudo = _bytes(nome)
    assert hashlib.sha256(conteudo).hexdigest() == info["sha256"]
    assert len(conteudo) == info["bytes"]


def test_fixtures_nao_carregam_identidade_nem_dado_contratual():
    """So codigos sinteticos (ITEM-00n / PC-2024-00n) e nenhum metadado pessoal."""
    import zipfile

    for nome in MANIFESTO:
        with zipfile.ZipFile(_caminho(nome)) as z:
            nucleo = z.read("docProps/core.xml").decode("utf-8", "ignore")
            app = z.read("docProps/app.xml").decode("utf-8", "ignore")
        assert "dpenalva" not in nucleo + app and "danie" not in nucleo + app, nome
    wb = load_workbook(_caminho("pre11_l2_financeiro"), data_only=True)
    itens = {wb["itens_Remanesc"].cell(r, 1).value for r in range(2, 8)} - {None}
    assert itens == {"ITEM-001", "ITEM-002", "ITEM-003"}


# --------------------------------------------------------------------------- #
# 1. Deteccao deterministica, sem falso positivo.
# --------------------------------------------------------------------------- #
@pytest.mark.parametrize("nome", sorted(MANIFESTO))
def test_cada_fixture_e_detectada_na_sua_linhagem(nome):
    wb = load_workbook(_caminho(nome))
    deteccao = cc.detectar_linhagem_coleta(wb)
    esperado = LINHAGEM_DE[MANIFESTO[nome]["linhagem_esperada"]]
    assert deteccao["codigo"] == esperado
    assert deteccao["suportada"] is True
    assert deteccao["modelo_canonico"] == "11.0"
    assert deteccao["compatibilidade_aplicada"] is (esperado != cc.LINHAGEM_COLETA_11)
    assert deteccao["marcador_publico"] == ("11.0" if esperado == cc.LINHAGEM_COLETA_11 else None)


def test_fingerprints_das_tres_linhagens():
    """Assinaturas estruturais: historicas CONGELADAS (L1/L2) e 11.0 deterministica.

    O leiaute 11.0 atual e o L1 historico sao o mesmo desenho, mas o modelo
    evoluiu em cabecalhos entre as duas datas; por isso o fingerprint nao e
    igual e o que faz um arquivo ser COLETA_11 e o marcador publico. Os
    fingerprints de L1 e L2 abaixo sao fatos historicos: se mudarem, a
    deteccao mudou de criterio.
    """
    def fp(nome):
        return cc.detectar_linhagem_coleta(load_workbook(_caminho(nome)))["fingerprint"]

    f11, fl1, fl2 = fp("coleta_11_financeiro"), fp("pre11_l1_financeiro"), fp("pre11_l2_financeiro")
    assert fl1.startswith("3b39c8f677492a95")
    assert fl2.startswith("c7c9b400a5ff")
    assert len({f11, fl1, fl2}) == 3
    assert all(len(f) == 64 for f in (f11, fl1, fl2))
    # estavel entre arquivos da mesma linhagem: nao depende do conteudo preenchido
    assert fp("pre11_l2_pc") == fp("pre11_l2_consumidos") == fl2
    assert fp("pre11_l1_consumidos") == fp("pre11_l1_financeiro_regerada") == fl1
    assert fp("coleta_11_pc") == fp("coleta_11_consumidos") == f11


def test_marcador_e_o_unico_diferenciador_entre_11_e_l1():
    wb = load_workbook(_caminho("coleta_11_financeiro"))
    assert cc.detectar_linhagem_coleta(wb)["codigo"] == cc.LINHAGEM_COLETA_11
    for coordenada in ("A24", "B24", "A25", "B25"):
        wb["CONTROLE"][coordenada].value = None
    assert cc.detectar_linhagem_coleta(wb)["codigo"] == cc.LINHAGEM_PRE_11_L1


def test_marcador_inventado_ou_desconhecido_nao_e_homologado():
    wb = load_workbook(_caminho("pre11_l1_financeiro"))
    wb["CONTROLE"]["A24"], wb["CONTROLE"]["B24"] = "Modelo de Coleta", "10.9"
    deteccao = cc.detectar_linhagem_coleta(wb)
    assert deteccao["codigo"] == cc.LINHAGEM_NAO_HOMOLOGADA and not deteccao["suportada"]
    assert "Nenhuma versão foi inferida" in cc.mensagem_linhagem_nao_homologada(deteccao)


def test_marcador_11_em_estrutura_l2_nao_e_promovido():
    wb = load_workbook(_caminho("pre11_l2_financeiro"))
    wb["CONTROLE"]["A24"], wb["CONTROLE"]["B24"] = "Modelo de Coleta", "11.0"
    assert cc.detectar_linhagem_coleta(wb)["codigo"] == cc.LINHAGEM_NAO_HOMOLOGADA


@pytest.mark.parametrize("mutacao", ["sem_numero_pc", "sem_aba", "cabecalho_trocado"])
def test_estrutura_desconhecida_e_fail_closed(mutacao):
    wb = load_workbook(_caminho("pre11_l2_financeiro"))
    if mutacao == "sem_numero_pc":
        wb["itens_PC"]["A1"] = "PC"
    elif mutacao == "sem_aba":
        del wb["itens_RC"]
    else:
        wb["parametros"]["E1"] = "PERCENTUAL"
    deteccao = cc.detectar_linhagem_coleta(wb)
    assert deteccao["codigo"] == cc.LINHAGEM_NAO_HOMOLOGADA
    assert deteccao["modelo_canonico"] is None and deteccao["compatibilidade_aplicada"] is False


def test_abas_opcionais_do_l1_nao_decidem_a_linhagem():
    """Sem `cobertura_temporal` (opcional no runtime) o arquivo segue L1; sem as
    abas estruturais do L1 (ex.: MEMORIA_RESULTADOS) deixa de ser homologado."""
    wb = load_workbook(_caminho("pre11_l1_financeiro"))
    del wb["cobertura_temporal"]
    assert cc.detectar_linhagem_coleta(wb)["codigo"] == cc.LINHAGEM_PRE_11_L1
    del wb["MEMORIA_RESULTADOS"]
    assert cc.detectar_linhagem_coleta(wb)["codigo"] == cc.LINHAGEM_NAO_HOMOLOGADA


def test_modelo_anterior_sem_numero_pc_nao_e_homologado():
    """`templates/Coleta_Reajuste.xlsx` (11 abas, sem NUMERO_PC) e um negativo real."""
    wb = load_workbook(ROOT / "templates" / "Coleta_Reajuste.xlsx")
    assert cc.detectar_linhagem_coleta(wb)["codigo"] == cc.LINHAGEM_NAO_HOMOLOGADA


def test_leitor_recusa_linhagem_nao_homologada_sem_inferir_versao():
    wb = load_workbook(_caminho("pre11_l2_financeiro"))
    wb["itens_PC"]["A1"] = "PC"
    saida = io.BytesIO()
    wb.save(saida)
    leitura = ler_masterfile_v10(saida.getvalue(), exigir_modelo_oficial=True)
    assert leitura["ok"] is False
    assert "não homologado" in leitura["erro"] and "Nenhuma versão foi inferida" in leitura["erro"]
    assert leitura["coleta_linhagem"]["codigo"] == cc.LINHAGEM_NAO_HOMOLOGADA


def test_deteccao_nao_usa_nome_nem_data_do_arquivo():
    """A assinatura depende so do workbook: os mesmos bytes, qualquer origem."""
    import inspect

    conteudo = _bytes("pre11_l2_financeiro")
    a = cc.detectar_linhagem_coleta(load_workbook(io.BytesIO(conteudo)))
    b = cc.detectar_linhagem_coleta(load_workbook(io.BytesIO(conteudo)))
    assert a == b
    assert list(inspect.signature(cc.detectar_linhagem_coleta).parameters) == ["wb"]


# --------------------------------------------------------------------------- #
# 2-4. Preservacao, reproducao do cache e equivalencia celula a celula.
# --------------------------------------------------------------------------- #
ABAS_ENTRADA_E_DERIVADOS = (
    "parametros", "financeiro", "historico_VU", "itens_Remanesc", "itens_RC",
    "itens_Consumidos", "itens_PC", "aditivos", "CICLO_EM_EXECUCAO",
)
ABAS_VISAO_XLS = {"RESULTADOS", "MEMORIA_RESULTADOS", "comparativo_VTA",
                  "posicao_referencia", "cobertura_temporal", "CONTROLE"}


def _eh_formula(valor) -> bool:
    return isinstance(valor, str) and valor.startswith("=")


def _valores(nome: str):
    return load_workbook(_caminho(nome), data_only=True)


@pytest.mark.parametrize("nome", ["pre11_l1_financeiro", "pre11_l2_financeiro"])
def test_adaptacao_preserva_entradas_e_altera_so_derivados(nome):
    formulas = load_workbook(_caminho(nome))
    original = _valores(nome)
    adaptado = _valores(nome)
    auditoria = aplicar_compatibilidade_valores(adaptado)
    assert auditoria["aplicada"] is True and auditoria["celulas_alteradas"] > 50

    constantes = 0
    alteradas = []
    for aba in formulas.sheetnames:
        if aba in ABAS_VISAO_XLS:
            continue
        wf, wo, wa = formulas[aba], original[aba], adaptado[aba]
        for linha in wf.iter_rows(min_row=1, max_row=min(wf.max_row, 5001)):
            for celula in linha:
                valor = celula.value
                if valor in (None, ""):
                    continue
                antes = wo.cell(celula.row, celula.column).value
                depois = wa.cell(celula.row, celula.column).value
                if not _eh_formula(valor):
                    constantes += 1
                    assert depois == antes, (
                        f"ENTRADA alterada: {aba}!{celula.coordinate} {antes!r} -> {depois!r}"
                    )
                elif depois != antes:
                    alteradas.append((aba, celula.coordinate))
    assert constantes > 400
    assert alteradas, "a adaptacao nao alterou nenhum derivado"
    assert {aba for aba, _ in alteradas} <= set(ABAS_ENTRADA_E_DERIVADOS)


@pytest.mark.parametrize("nome", ["pre11_l1_financeiro", "pre11_l2_financeiro"])
def test_entradas_do_fiscal_por_campo(nome):
    """Campo a campo nas abas de entrada (datas, pagos, PCs, itens, consumo)."""
    original, adaptado = _valores(nome), _valores(nome)
    aplicar_compatibilidade_valores(adaptado)
    campos = {
        "parametros": "ACDEG", "financeiro": "ACG", "itens_PC": "ABDG",
        "itens_Consumidos": "ABCEG", "itens_Remanesc": "ABC", "aditivos": "ABDEHK",
    }
    comparados = 0
    for aba, colunas in campos.items():
        ultima = 6 if aba == "parametros" else 80     # A11:E15 e memoria derivada
        for coluna in colunas:
            for linha in range(2, ultima + 1):
                antes = original[aba][f"{coluna}{linha}"].value
                assert adaptado[aba][f"{coluna}{linha}"].value == antes, (aba, coluna, linha)
                comparados += antes is not None
    assert comparados > 100        # 136-137 campos de entrada neste cenario
    # o percentual BRUTO do fiscal/calculadora continua o que foi informado
    assert adaptado["parametros"]["E3"].value == 0.05123816


@pytest.mark.parametrize("lin", ["l1", "l2"])
def test_adaptado_bruto_equivale_a_coleta_regerada_celula_a_celula(lin):
    """O arquivo ANTIGO adaptado == a MESMA Coleta regerada no Excel real.

    Unicas excecoes, todas explicadas: a entrada `parametros!E3` (percentual
    bruto informado), e celulas que dependem da VISAO do XLS
    (MEMORIA_RESULTADOS), que por desenho continua como o Excel gravou.
    """
    adaptado = _valores(f"pre11_{lin}_financeiro")
    regerada = _valores(f"pre11_{lin}_financeiro_regerada")
    auditoria = aplicar_compatibilidade_valores(adaptado)
    assert auditoria["blocos_nao_reproduziveis"] == {}
    assert {"financeiro", "historico_VU", "itens_PC", "itens_Consumidos",
            "itens_Remanesc", "itens_RC", "aditivos"} <= set(auditoria["blocos_adaptados"])

    def excecao(aba, celula):
        if aba == "parametros" and celula.coordinate == "E3":
            return True
        if aba == "itens_Remanesc" and celula.column >= 43:          # AQ:BH (parcelas A/B)
            return True
        return aba == "itens_PC" and celula.coordinate in {"M19", "S19"}

    comparadas = 0
    for aba in ABAS_ENTRADA_E_DERIVADOS:
        if aba not in adaptado.sheetnames:
            continue
        wa, wr = adaptado[aba], regerada[aba]
        for linha in wa.iter_rows(min_row=1, max_row=min(wa.max_row, 260)):
            for celula in linha:
                esperado = wr.cell(celula.row, celula.column).value
                comparadas += 1
                a, e = celula.value, esperado
                if isinstance(a, (int, float)) and isinstance(e, (int, float)):
                    igual = abs(a - e) <= 0.005
                else:
                    igual = (a in (None, "") and e in (None, "")) or a == e
                if not igual and not excecao(aba, celula):
                    pytest.fail(f"{aba}!{celula.coordinate}: adaptado={a!r} regerada={e!r}")
    assert comparadas > 10000


@pytest.mark.parametrize("nome", [
    "pre11_l1_financeiro", "pre11_l1_pc", "pre11_l1_consumidos",
    "pre11_l2_financeiro", "pre11_l2_pc", "pre11_l2_consumidos",
])
def test_formulas_reproduzidas_coincidem_com_o_cache_de_cada_arquivo(nome):
    """Protecao contra formula diferente: nenhum bloco deixa de reproduzir o cache."""
    wb = _valores(nome)
    auditoria = aplicar_compatibilidade_valores(wb)
    assert auditoria["aplicada"] is True
    assert auditoria["blocos_nao_reproduziveis"] == {}
    assert auditoria["ciclos_precisao_bruta"] == ["C1"]


def test_bloco_que_nao_reproduz_o_cache_nao_e_adaptado():
    """Se a formula do arquivo difere da reproduzida, o valor do XLS e mantido."""
    wb = _valores("pre11_l2_financeiro")
    original = wb["financeiro"]["E3"].value
    wb["financeiro"]["E3"].value = original + 7.77        # cache "adulterado"
    auditoria = aplicar_compatibilidade_valores(wb)
    assert "financeiro" in auditoria["blocos_nao_reproduziveis"]
    assert "financeiro" not in auditoria["blocos_adaptados"]
    assert wb["financeiro"]["E3"].value == original + 7.77      # nada foi "corrigido"
    assert "historico_VU" in auditoria["blocos_adaptados"]       # os demais seguem


def test_adaptacao_e_idempotente_e_ignora_arquivo_ja_oficial():
    wb = _valores("pre11_l1_financeiro")
    primeira = aplicar_compatibilidade_valores(wb)
    instantaneo = [c.value for c in wb["historico_VU"]["D"]][:10]
    assert aplicar_compatibilidade_valores(wb) is primeira
    assert getattr(wb, ATRIBUTO_AUDITORIA) is primeira
    assert [c.value for c in wb["historico_VU"]["D"]][:10] == instantaneo

    regerada = aplicar_compatibilidade_valores(_valores("pre11_l1_financeiro_regerada"))
    assert regerada["aplicada"] is False and "oficiais" in regerada["motivo"]
    onze = aplicar_compatibilidade_valores(_valores("coleta_11_financeiro"))
    assert onze["aplicada"] is False and onze["linhagem"] == cc.LINHAGEM_COLETA_11


def test_arquivo_fisico_nunca_e_alterado():
    conteudo = _bytes("pre11_l2_financeiro")
    antes = hashlib.sha256(conteudo).hexdigest()
    ler_masterfile_v10(conteudo, exigir_modelo_oficial=True)
    assert hashlib.sha256(conteudo).hexdigest() == antes == MANIFESTO["pre11_l2_financeiro"]["sha256"]


# --------------------------------------------------------------------------- #
# 5. Equivalencia no runtime (as cadeias de producao).
# --------------------------------------------------------------------------- #
_CACHE_RUNTIME: dict = {}


def _runtime(nome: str):
    if nome not in _CACHE_RUNTIME:
        from _baseline_fotografia import fotografar_web
        from _coleta_reajuste_documentos import processar_coleta_oficial_runtime

        resultado, diagnostico = processar_coleta_oficial_runtime(_bytes(nome))
        _CACHE_RUNTIME[nome] = (resultado, diagnostico, fotografar_web(resultado, diagnostico))
    return _CACHE_RUNTIME[nome]


# Campos de ORIGEM XLS (visao do arquivo), identidade e gates: nao sao "calculo
# Python" e por isso nao entram na comparacao de equivalencia.
_IGNORADAS = {
    "data_processamento", "_resultado_lido_do_excel", "contrato_xls", "valores_xls",
    "coleta_linhagem", "coleta_modelo_canonico", "compatibilidade_aplicada",
    "derivados_recalculados", "compatibilidade_valores", "versao_detectada",
    "origem_coleta", "diagnostico_coleta", "hash_entrada", "processo_ref",
    "chave_canonica", "avisos", "alertas", "painel_executivo",
    # texto dos cabecalhos de itens_PC (renomeados entre modelos): metadado de
    # estrutura, nao calculo
    "campos_detectados",
}
_FILTRO_XLS_E_GATES = re.compile(
    r"capacidades\.|resultados_xls|status_resultados|resultados_progressivos|"
    r"referencias_vta|reconciliacao_xls_python|bloqueios|formalizacao|"
    r"politica_entrega|ressalvas|campos_nao_confiaveis|limitacoes|"
    r"resultado_consolidado|ciclos_precisao_bruta"
)


def _normalizar(valor):
    import pandas as pd

    if isinstance(valor, pd.DataFrame):
        return [_normalizar(r) for r in valor.to_dict("records")]
    if isinstance(valor, dict):
        return {str(k): _normalizar(v) for k, v in valor.items() if str(k) not in _IGNORADAS}
    if isinstance(valor, (list, tuple)):
        return [_normalizar(v) for v in valor]
    if isinstance(valor, float) and math.isnan(valor):
        return None
    if isinstance(valor, (dt.datetime, dt.date, pd.Timestamp)):
        return str(valor)[:10]
    return valor


def _comparar(a, b, caminho="", saida=None):
    saida = [] if saida is None else saida
    if isinstance(a, dict) and isinstance(b, dict):
        for k in sorted(set(a) | set(b)):
            if k not in a or k not in b:
                saida.append(f"{caminho}.{k}: presente em um lado so")
            else:
                _comparar(a[k], b[k], f"{caminho}.{k}", saida)
    elif isinstance(a, list) and isinstance(b, list):
        if len(a) != len(b):
            saida.append(f"{caminho}: tamanho {len(a)} != {len(b)}")
        for i, (x, y) in enumerate(zip(a, b)):
            _comparar(x, y, f"{caminho}[{i}]", saida)
    elif (isinstance(a, (int, float)) and isinstance(b, (int, float))
          and not isinstance(a, bool) and not isinstance(b, bool)):
        if abs(a - b) > 0.005:
            saida.append(f"{caminho}: {a} != {b}")
    elif a != b:
        saida.append(f"{caminho}: {str(a)[:60]} != {str(b)[:60]}")
    return saida


def _diferencas_python(referencia: str, outro: str) -> list[str]:
    ref = _normalizar(_runtime(referencia)[0])
    alvo = _normalizar(_runtime(outro)[0])
    return [d for d in _comparar(ref, alvo) if not _FILTRO_XLS_E_GATES.search(d)]


@pytest.mark.parametrize("metodo", METODOS)
def test_l1_antigo_equivale_ao_11_no_lado_python_nos_tres_metodos(metodo):
    """Com precisao ANTIGA no XLS, o Python publica exatamente o que o 11.0 publica."""
    diferencas = _diferencas_python(f"coleta_11_{metodo}", f"pre11_l1_{metodo}")
    assert diferencas == [], "\n".join(diferencas[:30])


def test_l1_regerada_equivale_ao_11_e_ao_l1_antigo_adaptado():
    # regerada (percentual oficial) x 11.0
    assert _diferencas_python("coleta_11_financeiro", "pre11_l1_financeiro_regerada") == []
    # antigo adaptado x regerada da mesma linhagem
    assert _diferencas_python("pre11_l1_financeiro_regerada", "pre11_l1_financeiro") == []


@pytest.mark.parametrize("lin,metodo", [("l1", "financeiro"), ("l1", "consumidos"),
                                        ("l2", "consumidos")])
def test_numeros_canonicos_iguais_ao_11(lin, metodo):
    """VTA, retroativo, remanescente, fatores, financeiro, valores unitarios."""
    r_ref, _, ref = _runtime(f"coleta_11_{metodo}")
    r_lin, _, web = _runtime(f"pre11_{lin}_{metodo}")
    for campo in ("vta_oficial", "retroativo_total", "remanescente_atualizado",
                  "execucao_atualizada_do_ciclo", "valor_atualizado_contrato"):
        assert web.get(campo) == ref.get(campo), (lin, metodo, campo)
    assert [c["fator_acumulado"] for c in web["ciclos"] if c.get("fator_acumulado")] == [
        c["fator_acumulado"] for c in ref["ciclos"] if c.get("fator_acumulado")
    ]
    for df in ("df_financeiro_mensal", "df_valores_unitarios_ciclo", "df_delta_por_ciclo"):
        a, b = _normalizar(r_ref[df]), _normalizar(r_lin[df])
        colunas = [k for k in (a[0] if a else {}) if isinstance(a[0][k], (int, float))]
        numeros_a = [[x[k] for k in colunas] for x in a]
        numeros_b = [[y[k] for k in colunas] for y in b]
        assert _comparar(numeros_a, numeros_b) == [], df
    assert r_lin["total_devido_reajustado"] == r_ref["total_devido_reajustado"]


def test_valores_do_caso_equivalente_sao_os_esperados():
    """Ancora numerica (11.0): o caso e pequeno o bastante para conferir a mao."""
    _, _, w = _runtime("coleta_11_financeiro")
    assert w["vta_oficial"] == 1_340_973.60
    assert w["retroativo_total"] == 26_112.00
    # (120+10 un x 250 + 80 x 1.500 + 40 x 3.200) x 1,0512; o +10 e o aditivo
    assert w["remanescente_atualizado"] == 294_861.60
    _, _, p = _runtime("coleta_11_pc")
    assert p["retroativo_total"] == 14_156.80                  # (180.000 + 96.500) x 5,12%
    assert p["retroativo_potencial"] == 7_436.80               # 145.250 x 5,12% (PC nao pago)
    _, _, c = _runtime("coleta_11_consumidos")
    assert c["vta_oficial"] == 289_514.88


def test_pc_l1_recompoe_retroativo_reconhecido_e_potencial():
    """H/I/J dos PCs: o derivado bruto do cache nao vaza para o retroativo."""
    resultado, _, web = _runtime("pre11_l1_pc")
    itens = {i["numero_pc"]: i for i in resultado["itens_pc_v10"]["itens"]}
    assert itens["PC-2024-003"]["delta_potencial"] == 7_436.80     # nao 7.442,34 (bruto)
    assert web["retroativo_total"] == 14_156.80


def test_l2_pc_e_fail_closed_sem_inicio_do_efeito_financeiro():
    """L2 nao tem parametros!H: o efeito do PC e indeterminavel => nada e inventado."""
    resultado, _, web = _runtime("pre11_l2_pc")
    bloqueios = " ".join(resultado["bloqueios_formalizacao"])
    assert "inicio do efeito financeiro ausente ou inconsistente" in bloqueios
    assert resultado["formalizacao_bloqueada"] is True
    assert web["status_apuracao"]["status_politica"] == "BLOQUEADO_PARA_FORMALIZACAO"


@pytest.mark.parametrize("nome", ["pre11_l2_financeiro", "pre11_l2_financeiro_regerada"])
def test_l2_com_aditivo_e_fail_closed_e_a_diferenca_e_so_o_aditivo(nome):
    """No L2 o aditivo nao entra na quantidade remanescente ajustada.

    L2: QTD_REM_AJUSTADA_Cn = QTD_REM_BASE_Cn. 11.0: BASE + DELTA_Cn. A diferenca
    de VTA e remanescente para o 11.0 e exatamente o aditivo (10 un x R$ 250 x
    1,0512 = R$ 2.628,00) — regra de modelo, nao precisao de percentual. Como o
    aditivo esta no arquivo e a posicao nao, o resultado fica bloqueado com o
    motivo tecnico, em vez de publicar um VTA incompleto em silencio.
    """
    resultado, _, web = _runtime(nome)
    assert MENSAGEM_L2_ADITIVO in resultado["bloqueios_formalizacao"]
    assert resultado["formalizacao_bloqueada"] is True
    leitura = ler_masterfile_v10(_bytes(nome), exigir_modelo_oficial=True)
    assert [r["codigo"] for r in leitura["compatibilidade_restricoes"]] == [
        "L2_ADITIVO_SEM_EFEITO_NA_POSICAO"
    ]
    _, _, ref = _runtime("coleta_11_financeiro")
    assert round(ref["vta_oficial"] - web["vta_oficial"], 2) == 2_628.00
    assert round(ref["remanescente_atualizado"] - web["remanescente_atualizado"], 2) == 2_628.00
    # o que nao depende da posicao remanescente e identico ao 11.0
    assert web["retroativo_total"] == ref["retroativo_total"] == 26_112.00


def test_restricao_do_l2_so_existe_com_aditivo_em_ciclo_de_reajuste():
    wb = _valores("pre11_l2_financeiro")
    assert [r["codigo"] for r in restricoes_compatibilidade(wb)] == [
        "L2_ADITIVO_SEM_EFEITO_NA_POSICAO"
    ]
    wb["aditivos"]["L2"].value = 0.0                      # sem delta de quantidade
    assert restricoes_compatibilidade(wb) == []
    # L1 e 11.0 ja incorporam o aditivo na posicao: nenhuma restricao
    assert restricoes_compatibilidade(_valores("pre11_l1_financeiro")) == []
    assert restricoes_compatibilidade(_valores("coleta_11_financeiro")) == []


# --------------------------------------------------------------------------- #
# 6. Nenhum gate foi removido; a divergencia fica exposta.
# --------------------------------------------------------------------------- #
@pytest.mark.parametrize("nome", ["pre11_l1_financeiro", "pre11_l1_pc", "pre11_l1_consumidos",
                                  "pre11_l2_financeiro"])
def test_precisao_antiga_segue_bloqueada_para_formalizacao(nome):
    resultado = _runtime(nome)[0]
    assert resultado["formalizacao_bloqueada"] is True
    assert resultado["politica_entrega_segura"]["pode_confirmar"] is False
    documentos = resultado["capacidades"]["documentos"]
    if MENSAGEM_COLETA_PRECISAO_ANTERIOR in resultado["bloqueios_formalizacao"]:
        for chave in ("sumario_executivo", "despacho_saneador", "termo_apostila"):
            assert documentos[chave]["habilitado"] is False, chave
            assert documentos[chave]["motivo"] == MENSAGEM_COLETA_PRECISAO_ANTERIOR


@pytest.mark.parametrize("nome", ["pre11_l1_financeiro", "pre11_l2_financeiro"])
def test_bloqueio_de_precisao_anterior_permanece_no_metodo_financeiro(nome):
    """Politica existente (inalterada): a mensagem de precisao vale para o Financeiro."""
    resultado, diagnostico, _ = _runtime(nome)
    assert any("precisao bruta" in a for a in diagnostico.get("avisos") or [])
    assert MENSAGEM_COLETA_PRECISAO_ANTERIOR in resultado["bloqueios_formalizacao"]


@pytest.mark.parametrize("nome", ["pre11_l1_pc", "pre11_l1_consumidos", "pre11_l2_consumidos"])
def test_pc_e_consumidos_com_precisao_antiga_sao_bloqueados_pela_divergencia(nome):
    """PC/Consumidos: a protecao e a reconciliacao XLS x Python (regra ja existente)."""
    resultado, diagnostico, _ = _runtime(nome)
    assert any("precisao bruta" in a for a in diagnostico.get("avisos") or [])
    assert any("Divergência relevante XLS × Python" in b
               for b in resultado["bloqueios_formalizacao"]), nome
    assert resultado["formalizacao_bloqueada"] is True


def test_divergencia_xls_python_fica_exposta_e_nao_e_adotada():
    """O XLS antigo diz 26.131,44; o Python, 26.112,00: ambos ficam visiveis."""
    _, _, web = _runtime("pre11_l1_financeiro")
    convergencia = web["convergencia_xls_python"]
    assert convergencia["status_geral"] == "DIVERGENCIA_RELEVANTE"
    campos = {c["campo"]: c for c in convergencia["campos"]}
    assert campos["RETRO_FIN"]["xls"] == 26_131.44 and campos["RETRO_FIN"]["python"] == 26_112.00
    assert campos["VTA_FINAL"]["xls"] == 1_341_003.94 and campos["VTA_FINAL"]["python"] == 1_340_973.60
    assert web["vta_oficial"] == 1_340_973.60            # a web usa o canonico, nao o XLS
    assert campos["QTD_REM_OFICIAL"]["status"] == "CONCILIADO"


def test_coleta_11_e_l1_regerada_continuam_sem_bloqueio():
    for nome in ("coleta_11_financeiro", "coleta_11_pc", "pre11_l1_financeiro_regerada"):
        resultado = _runtime(nome)[0]
        assert resultado["bloqueios_formalizacao"] == [], nome
        assert resultado["formalizacao_bloqueada"] is False, nome


def test_auditoria_da_compatibilidade_no_resultado_da_leitura():
    leitura = ler_masterfile_v10(_bytes("pre11_l1_financeiro"), exigir_modelo_oficial=True)
    assert leitura["ok"] is True
    assert leitura["coleta_linhagem"]["codigo"] == cc.LINHAGEM_PRE_11_L1
    assert leitura["coleta_modelo_canonico"] == "11.0"
    assert leitura["compatibilidade_aplicada"] is True
    assert "financeiro" in leitura["derivados_recalculados"]
    assert leitura["compatibilidade_valores"]["blocos_nao_reproduziveis"] == {}
    assert any("anterior ao modelo 11.0" in a for a in leitura["avisos"])

    onze = ler_masterfile_v10(_bytes("coleta_11_financeiro"), exigir_modelo_oficial=True)
    assert onze["coleta_linhagem"]["codigo"] == cc.LINHAGEM_COLETA_11
    assert onze["compatibilidade_aplicada"] is False and onze["derivados_recalculados"] == []


def test_contexto_de_producao_adapta_uma_vez_antes_dos_leitores():
    """O caminho do upload (ContextoColeta) entrega o workbook ja adaptado."""
    from _contexto_coleta import ContextoColeta

    with ContextoColeta(_bytes("pre11_l1_financeiro")) as contexto:
        wb = contexto.workbook_valores
        assert getattr(wb, ATRIBUTO_AUDITORIA)["aplicada"] is True
        assert wb["parametros"]["F3"].value == 1.0512
        assert wb["historico_VU"]["D2"].value == 262.80
        assert contexto.workbook_formulas["parametros"]["F3"].value.startswith("=")
