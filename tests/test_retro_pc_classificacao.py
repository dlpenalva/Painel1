"""Regressao PR #200 — classificacao do retroativo no metodo Pedidos de Compra.

Defeito (caso real telt_exempl_coleta.xlsx): RESULTADOS!C9 mostrava R$ 36.397,14
de retroativo reconhecido (itens_PC!H, PCs pagos) e D9 R$ 119,20 potencial
(itens_PC!J, PCs nao pagos), mas o Sumario Executivo dizia "Em analise
R$ 36.397,14 / Reconhecido R$ 0,00". Causa: _motor_temporal._classificar_pc (Q8)
so tirava um PC pago de "potencial" se TIPO_FINANCEIRO contivesse "reconhec",
e o default e "Impacto Potencial" em todas as linhas.

Conceitos separados:
1. reconhecido no calculo = retroativo sobre pagamento definitivo comprovado;
2. potencial = PCs nao pagos (nao entra no retroativo do assistente);
3. reconhecimento administrativo = nao comprovado nesta apuracao (nunca
   inferido de TIPO_FINANCEIRO, campo tecnico do leitor, nem do pagamento);
4. diferenca ja paga = so com evidencia (valor pago = valor atualizado).
"""
from __future__ import annotations

import copy
import os
import sys
from datetime import date
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from _assistente_fiscal import _classificar_retroativo  # noqa: E402
from _motor_temporal import (  # noqa: E402
    DELTA_JA_PAGO,
    DELTA_POTENCIAL,
    DELTA_RECONHECIDO,
    montar_motor_temporal,
)
from _sumario_executivo import (  # noqa: E402
    NAO_APLICAVEL,
    ROTULO_DIFERENCA_PAGA,
    ROTULO_EM_ANALISE,
    ROTULO_RECONHECIDO_CALCULO,
    ROTULO_RECONHECIMENTO_ADM,
    SEM_RECONHECIMENTO_ADM,
    gerar_sumario_executivo,
    montar_dados_sumario_executivo,
)
from test_sumario_executivo import (  # noqa: E402
    leitura_multiciclo_pc,
    leitura_simples_financeiro,
)

DATA_COM_EFEITO = date(2024, 5, 10)    # C1, apos o inicio dos efeitos
DATA_SEM_EFEITO = date(2024, 4, 10)    # C1, antes do inicio dos efeitos


def _por_ciclo():
    limites = {
        "C0": (date(2022, 10, 1), date(2023, 9, 30)),
        "C1": (date(2023, 10, 1), date(2024, 9, 30)),
        "C2": (date(2024, 10, 1), date(2025, 9, 30)),
        "C3": (date(2025, 10, 1), date(2026, 9, 30)),
        "C4": (date(2026, 10, 1), date(2027, 9, 30)),
    }
    return {
        c: {
            "ciclo": c, "data_inicio": i, "data_fim": f,
            "computar_nesta_apuracao": "Sim" if c == "C1" else "Nao",
            "percentual_reajuste": 0.10 if c == "C1" else 0.0,
            "fator_acumulado": 1.10 if c == "C1" else 1.0,
            "inicio_efeito_financeiro": date(2024, 4, 18) if c == "C1" else None,
        }
        for c, (i, f) in limites.items()
    }


def _pc(numero, pago="Sim", data_pc=DATA_COM_EFEITO, valor_pago=None,
        tipo="Impacto Potencial"):
    campos = {"tipo_financeiro": tipo}
    if pago is not None:
        campos["pc_pago_a_contratada"] = pago
        campos["elegivel_retroativo_pc"] = pago == "Sim"
    if valor_pago is not None:
        campos["valor_pago"] = valor_pago
    item = {"numero_pc": numero, "data_pc": data_pc, "ciclo": "C1",
            "valor_pc": 100.0, "campos_vta": campos}
    if pago is not None:
        item["pc_pago_a_contratada"] = pago
    return item


CENARIOS = {
    # nome: (item, natureza, retroativo)
    "pago_com_retroativo": (_pc("P1"), DELTA_RECONHECIDO, 10.0),
    "nao_pago_potencial": (_pc("P2", pago="Nao"), DELTA_POTENCIAL, 0.0),
    "pago_sem_diferenca": (_pc("P3", data_pc=DATA_SEM_EFEITO), DELTA_JA_PAGO, 0.0),
    "pago_com_glosa": (_pc("P4", valor_pago=80.0), DELTA_RECONHECIDO, 8.0),
    "pagamento_nao_confirmado": (_pc("P5", pago=None), DELTA_POTENCIAL, 0.0),
    # TIPO_FINANCEIRO e tecnico: "Retroativo Reconhecido" nao muda nada.
    "tipo_reconhecido_pago": (
        _pc("P6", tipo="Retroativo Reconhecido"), DELTA_RECONHECIDO, 10.0),
    "tipo_reconhecido_nao_pago": (
        _pc("P8", pago="Nao", tipo="Retroativo Reconhecido"), DELTA_POTENCIAL, 0.0),
    "diferenca_ja_paga": (_pc("P7", valor_pago=110.0), DELTA_JA_PAGO, 11.0),
}


def _motor(itens):
    return montar_motor_temporal({
        "ciclos": {"por_ciclo": _por_ciclo()},
        "itens_pc": {"itens": list(itens)},
        "resumo": {},
    })


@pytest.mark.parametrize("nome", list(CENARIOS))
def test_motor_classifica_cada_cenario(nome):
    item, natureza, retroativo = CENARIOS[nome]
    pc = _motor([item]).pcs[0]
    assert pc.natureza_delta == natureza, nome
    assert pc.retroativo == pytest.approx(retroativo), nome
    # Nenhum campo de reconhecimento administrativo e inferido pelo motor.
    assert not hasattr(pc, "reconhecimento_formal")
    assert "reconhecimento_formal" not in pc.to_dict()


def test_pc_pago_nao_e_mais_potencial_pelo_texto_padrao():
    # O texto "Impacto Potencial" (default do leitor e do XLS) nao decide mais
    # a natureza de um PC com pagamento definitivo.
    pc = _motor([_pc("P1", tipo="Impacto Potencial")]).pcs[0]
    assert pc.natureza_delta == DELTA_RECONHECIDO   # reconhecido NO CALCULO


def test_pagamento_nao_confirmado_nao_gera_retroativo_nem_reconhecimento():
    pc = _motor([_pc("P5", pago=None)]).pcs[0]
    assert pc.retroativo == 0.0
    assert any(a["codigo"] == "PAGO_INDETERMINADO" for a in pc.alertas)


def test_assistente_separa_categorias_sem_duplicar_nem_omitir():
    motor = _motor(item for item, *_ in CENARIOS.values())
    pcs = [p.to_dict() for p in motor.pcs]
    retro = _classificar_retroativo(
        {"situacao_pcs": {"pcs": pcs}}, {"retro_reconhecido_financeiro": 0.0}
    )
    por_estado = retro["por_estado"]
    # reconhecido no calculo: P1 (10) + P4 glosa (8) + P6 (10, pago)
    assert por_estado["reconhecido"] == pytest.approx(28.0)
    # diferenca ja paga: so P7, com evidencia (pago = valor atualizado)
    assert por_estado["pago"] == pytest.approx(11.0)
    # potencial de PC nao pago nao vira retroativo do assistente
    assert por_estado["em análise"] == pytest.approx(0.0)
    # nenhuma parcela duplicada ou omitida
    soma_pcs = sum(p["retroativo"] or 0.0 for p in pcs if (p["retroativo"] or 0.0) > 0)
    assert retro["valor"] == pytest.approx(soma_pcs) == pytest.approx(39.0)
    assert len(retro["origens"]) == 4
    # nenhum reconhecimento administrativo inferido de TIPO_FINANCEIRO
    assert "reconhecido_formalmente" not in retro
    assert all("reconhecimento_formal" not in o for o in retro["origens"])


def _com_metodo(leitura, metodo):
    leitura = copy.deepcopy(leitura)
    leitura["resultados_xls"] = {"disponivel": True,
                                 "valores": {"METODO_RETROATIVO": metodo}}
    return leitura


def _quadro(pdf: bytes) -> dict:
    fitz = pytest.importorskip("fitz")
    rotulos = {ROTULO_RECONHECIDO_CALCULO, ROTULO_EM_ANALISE,
               ROTULO_DIFERENCA_PAGA, ROTULO_RECONHECIMENTO_ADM}
    quadro = {}
    for page in fitz.open(stream=pdf, filetype="pdf"):
        linhas: dict = {}
        for bloco in page.get_text("dict")["blocks"]:
            for linha in bloco.get("lines", []):
                y = round((linha["bbox"][1] + linha["bbox"][3]) / 2)
                texto = "".join(s["text"] for s in linha["spans"]).strip()
                if texto:
                    linhas.setdefault(y, []).append((linha["bbox"][0], texto))
        for y in sorted(linhas):
            partes = [t for _, t in sorted(linhas[y])]
            if partes[0] in rotulos and len(partes) == 2:
                quadro[partes[0]] = partes[1]
    return quadro


def test_sumario_pc_pago_aparece_como_reconhecido_no_calculo():
    leitura = _com_metodo(leitura_multiciclo_pc(), "PCs")
    dados = montar_dados_sumario_executivo(leitura)
    fin = dados["financeiro"]
    total = fin["retroativo_total"]
    assert total and total > 0
    assert fin["retroativo_por_estado"]["reconhecido"] == pytest.approx(total)
    assert fin["retroativo_por_estado"]["em análise"] == 0.0
    quadro = _quadro(gerar_sumario_executivo(leitura))
    assert quadro[ROTULO_RECONHECIDO_CALCULO] != "R$ 0,00"
    assert quadro[ROTULO_RECONHECIMENTO_ADM] == SEM_RECONHECIMENTO_ADM
    assert quadro[ROTULO_DIFERENCA_PAGA] == "R$ 0,00"


def test_sumario_pc_mostra_potencial_da_composicao():
    leitura = _com_metodo(leitura_multiciclo_pc(), "PCs")
    dados = montar_dados_sumario_executivo(leitura)
    dados["financeiro"]["retroativo_potencial_pc"] = 119.2
    from _sumario_executivo import gerar_sumario_executivo_pdf
    quadro = _quadro(gerar_sumario_executivo_pdf(dados))
    assert quadro[ROTULO_EM_ANALISE] == "R$ 119,20"


@pytest.mark.parametrize("metodo", ["Financeiro", "Itens"])
def test_metodos_financeiro_e_itens_sem_regressao(metodo):
    sem = montar_dados_sumario_executivo(leitura_simples_financeiro())
    com = montar_dados_sumario_executivo(
        _com_metodo(leitura_simples_financeiro(), metodo))
    assert com["sintese"] == sem["sintese"]
    assert (com["financeiro"]["retroativo_por_estado"]
            == sem["financeiro"]["retroativo_por_estado"])
    quadro = _quadro(gerar_sumario_executivo(
        _com_metodo(leitura_simples_financeiro(), metodo)))
    assert quadro[ROTULO_EM_ANALISE] == NAO_APLICAVEL
    assert quadro[ROTULO_RECONHECIMENTO_ADM] == SEM_RECONHECIMENTO_ADM


EXEMPLO_REAL = os.environ.get("CL8US_EXEMPLO_RETRO_PC")


@pytest.mark.skipif(not EXEMPLO_REAL, reason="defina CL8US_EXEMPLO_RETRO_PC com o xlsx real")
def test_exemplo_real_coerente_com_resultados():
    import openpyxl
    from _coleta_reajuste_documentos import processar_coleta_oficial_runtime
    conteudo = Path(EXEMPLO_REAL).read_bytes()
    xls = openpyxl.load_workbook(EXEMPLO_REAL, data_only=True)["RESULTADOS"]
    resultado, _ = processar_coleta_oficial_runtime(conteudo)
    dados = montar_dados_sumario_executivo(resultado)
    fin = dados["financeiro"]
    assert fin["retroativo_por_estado"]["reconhecido"] == pytest.approx(xls["C9"].value)
    assert fin["retroativo_potencial_pc"] == pytest.approx(xls["D9"].value)
    assert fin["retroativo_por_estado"]["em análise"] == 0.0
    assert dados["sintese"]["vta"] == pytest.approx(xls["B9"].value)


# =========================================================================== #
# Ajustes finais do PR #200
# (1) recomendacao: calculo nao comprova aprovacao administrativa;
# (2) RESULTADOS!C10 (= MEMORIA_RESULTADOS!AG9) distingue os dois conceitos.
# =========================================================================== #
from _assistente_fiscal import (  # noqa: E402
    ACAO_GERAR_APOSTILAMENTO,
    ACAO_SUBMETER_RECONHECIMENTO,
    _proxima_acao,
)
from _dossie_decisao import (  # noqa: E402
    ATO_ANALISE_RECONHECIMENTO,
    ATO_APOSTILAMENTO,
    STATUS_APTO,
    _decidir_status,
)


def _retro(reconhecido, em_analise=0.0, **extra):
    return {
        "existe": reconhecido + em_analise > 0,
        "valor": reconhecido + em_analise,
        "por_estado": {"em análise": em_analise, "reconhecido": reconhecido, "pago": 0.0},
        **extra,
    }


@pytest.mark.parametrize("extra", [{}, {"reconhecido_formalmente": 36397.14}])
def test_recomendacao_sempre_condiciona_formalizacao_a_analise(extra):
    # Nem um campo legado de "reconhecimento" promove a apostilamento.
    r = _proxima_acao({"escolhida": "Pedidos de Compra (PC)"}, _retro(36397.14, **extra), [])
    assert r["acao"] == ACAO_SUBMETER_RECONHECIMENTO
    assert r["acao"] != ACAO_GERAR_APOSTILAMENTO
    assert "não comprova a aprovação" in r["motivo"]


def _assistente(retro):
    return {"evidencia": {"suficiente": True},
            "resultado_recomendado": {"valor_recomendado": 1.0},
            "retroativo": retro, "inconsistencias": []}


@pytest.mark.parametrize("extra", [{}, {"reconhecido_formalmente": 36397.14}])
def test_dossie_preserva_status_e_condiciona_a_analise(extra):
    status, ato, motivo = _decidir_status(_assistente(_retro(36397.14, **extra)))
    assert status == STATUS_APTO          # status e confirmacao humana preservados
    assert ato == ATO_ANALISE_RECONHECIMENTO
    assert ato != ATO_APOSTILAMENTO
    assert "nao comprova aprovacao administrativa" in motivo


def test_legenda_c10_distingue_calculo_de_aprovacao():
    import io
    from openpyxl import load_workbook
    from _coleta_oficial import TEMPLATE_COLETA_OFICIAL, obter_coleta_oficial_bytes
    from _conferencias_execucao import TEXTOS_MEMORIA_RESULTADOS
    texto = TEXTOS_MEMORIA_RESULTADOS["AG9"]
    assert "no cálculo" in texto and "reconhecimento administrativo" in texto
    assert "a pagar à contratada" not in texto
    for wb in (load_workbook(TEMPLATE_COLETA_OFICIAL),
               load_workbook(io.BytesIO(obter_coleta_oficial_bytes()))):
        assert wb["MEMORIA_RESULTADOS"]["AG9"].value == texto
        assert wb["RESULTADOS"]["C10"].value == "=MEMORIA_RESULTADOS!$AG$9"
        assert wb["RESULTADOS"]["C8"].value == "RETROATIVO RECONHECIDO"


# =========================================================================== #
# Integracao: leitor oficial -> Motor Temporal -> assistente -> dossie.
# TIPO_FINANCEIRO (itens_PC!Y) e tecnico e sobrescrito pelo leitor; mesmo
# preenchido com "Retroativo Reconhecido", nao pode gerar reconhecimento
# administrativo nem recomendar apostilamento.
# =========================================================================== #
def _coleta_pcs_tipo_reconhecido() -> bytes:
    import io as _io
    from openpyxl import load_workbook
    from _coleta_oficial import gerar_coleta_oficial_preenchida
    from test_vta_pc_independente_27b import _dados_corte_c2
    wb = load_workbook(_io.BytesIO(gerar_coleta_oficial_preenchida(_dados_corte_c2())))
    wb["CONTROLE"]["B1"] = "Pedidos de Compras"
    # Sem recalculo do Excel a formula do fator nao tem valor em cache:
    # fixa o fator acumulado do C1 (10%) como literal no fixture.
    wb["parametros"]["F3"] = 1.10
    ws = wb["itens_PC"]
    for linha, (numero, pago) in enumerate((("PC-PAGO", "Sim"), ("PC-ABERTO", "Nao")), 2):
        ws[f"A{linha}"] = numero
        ws[f"B{linha}"] = date(2024, 1, 15)     # C1, com efeito (fator 1,10)
        ws[f"D{linha}"] = 600.0
        ws[f"G{linha}"] = pago
        ws[f"Y{linha}"] = "Retroativo Reconhecido"
    rem = wb["itens_Remanesc"]
    rem["A2"], rem["B2"], rem["C2"], rem["E2"], rem["G2"] = "ITEM-1", 10.0, 100.0, 4.0, 4.0
    buf = _io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_integrado_tipo_financeiro_nao_gera_reconhecimento_administrativo():
    from _assistente_fiscal import montar_assistente_fiscal
    from _dossie_decisao import montar_dossie_decisao
    from _leitor_masterfile_v10 import ler_masterfile_v10
    from _painel_executivo import montar_painel_executivo

    leitura = ler_masterfile_v10(_coleta_pcs_tipo_reconhecido())
    itens = {i["numero_pc"]: i for i in (leitura.get("itens_pc_v10") or {}).get("itens") or []}
    assert set(itens) == {"PC-PAGO", "PC-ABERTO"}
    # O leitor sobrescreve o campo tecnico: o texto digitado nao sobrevive.
    for item in itens.values():
        assert (item.get("campos_vta") or {}).get("tipo_financeiro") != "Retroativo Reconhecido"

    painel = montar_painel_executivo(leitura)
    pcs = {p["numero_pc"]: p for p in painel["situacao_pcs"]["pcs"]}
    assert pcs["PC-PAGO"]["natureza_delta"] == DELTA_RECONHECIDO      # no calculo
    assert pcs["PC-PAGO"]["retroativo"] == pytest.approx(60.0)
    assert pcs["PC-ABERTO"]["natureza_delta"] == DELTA_POTENCIAL
    assert pcs["PC-ABERTO"]["retroativo"] == 0.0
    assert all("reconhecimento_formal" not in p for p in pcs.values())

    assistente = montar_assistente_fiscal(leitura, painel)
    retro = assistente["retroativo"]
    assert retro["por_estado"]["reconhecido"] == pytest.approx(60.0)
    assert "reconhecido_formalmente" not in retro
    assert assistente["proxima_acao"]["acao"] == ACAO_SUBMETER_RECONHECIMENTO
    assert assistente["proxima_acao"]["acao"] != ACAO_GERAR_APOSTILAMENTO

    dossie = montar_dossie_decisao(assistente, painel)
    assert dossie["ato_ou_providencia"] != ATO_APOSTILAMENTO
    if dossie["status"] == STATUS_APTO:
        assert dossie["ato_ou_providencia"] == ATO_ANALISE_RECONHECIMENTO

    dados = montar_dados_sumario_executivo(leitura)
    quadro = _quadro(gerar_sumario_executivo(leitura))
    assert quadro[ROTULO_RECONHECIMENTO_ADM] == SEM_RECONHECIMENTO_ADM
    assert dados["financeiro"]["retroativo_por_estado"]["reconhecido"] == pytest.approx(60.0)
