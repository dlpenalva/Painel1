"""Testes do Sumario Executivo em PDF (Etapa 5).

Mapeamento dos 13 requisitos do documento:
  R1  Indice contratual .................. test_r1_indice_contratual_da_memoria
  R2  Datas dos pedidos .................. test_r2_datas_pedido_por_ciclo
  R3  Ciclos analisados .................. test_r3_r4_ciclos_e_periodos
  R4  Periodo de cada ciclo .............. test_r3_r4_ciclos_e_periodos
  R5  Meses sem efeito por ciclo ......... test_r5_meses_sem_efeito_*
  R6  Variacao por ciclo e acumulada ..... test_r6_variacao_por_ciclo_e_acumulada
  R7  Valor pago por ciclo ............... test_r7_r8_r9_metodo_financeiro
  R8  Delta por ciclo e total ............ test_r7_r8_r9_metodo_financeiro
  R9  Retroativo reconhecido a pagar ..... test_r7_r8_r9_metodo_financeiro
  R10 Retroativo em analise (PC) ......... test_r10_pc_em_analise_separado
  R11 Itens e valores por ciclo .......... test_r11_itens_vu_e_totais
  R12 Memoria de calculo por ciclo ....... test_r12_memoria_*
  R13 Aditivos aplicaveis ................ test_r13_aditivos_*
"""
from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from _sumario_executivo import (  # noqa: E402
    NAO_APLICAVEL,
    NAO_INFORMADO,
    formatar_moeda,
    formatar_percentual,
    gerar_sumario_executivo,
    gerar_sumario_executivo_pdf,
    montar_dados_sumario_executivo,
)


# ---------------------------------------------------------------------------
# Leituras sinteticas canonicas (formato do leitor v10)
# ---------------------------------------------------------------------------

def _ciclo_param(nome, inicio, fim, pct=None, fator=None, computar=None,
                 inicio_efeito=None, situacao=None):
    return {
        "ciclo": nome,
        "computar_nesta_apuracao": computar,
        "data_inicio": inicio,
        "data_fim": fim,
        "percentual_reajuste": pct,
        "fator_acumulado": fator,
        "situacao": situacao,
        "inicio_efeito_financeiro": inicio_efeito,
    }


def leitura_simples_financeiro():
    """Cenario 1: calculo simples (C1) pelo metodo Financeiro, com aditivo."""
    por_ciclo = {
        "C0": _ciclo_param("C0", date(2024, 2, 1), date(2025, 1, 31),
                           fator=1.0, situacao="Base"),
        "C1": _ciclo_param("C1", date(2025, 2, 1), date(2026, 1, 31),
                           pct=0.0525, fator=1.0525, computar="Sim",
                           inicio_efeito=date(2025, 4, 1), situacao="Computado"),
        "C2": _ciclo_param("C2", date(2026, 2, 1), date(2027, 1, 31)),
        "C3": _ciclo_param("C3", date(2027, 2, 1), date(2028, 1, 31)),
        "C4": _ciclo_param("C4", date(2028, 2, 1), date(2029, 1, 31)),
    }
    return {
        "ok": True,
        "controle": {"modo": "SIMPLES", "versao": "V10.4.8",
                     "ciclo_vigente": "C1", "data_corte": date(2026, 6, 30)},
        "resumo": {},
        "parametros_v10": {"ok": True, "por_ciclo": por_ciclo,
                           "ciclos": list(por_ciclo.values())},
        "memoria_calculo": {
            "C1": [
                {"tipo": "MES", "ordem": 1, "competencia": "2025-02-15",
                 "valor_indice": 0.0045, "fator_mensal": 1.0045,
                 "fator_acumulado": 1.0045},
                {"tipo": "MES", "ordem": 2, "competencia": "2025-03-15",
                 "valor_indice": 0.0032, "fator_mensal": 1.0032,
                 "fator_acumulado": 1.007714},
                {"tipo": "RESULTADO", "ordem": 3, "fator_acumulado": 1.0525,
                 "variacao_final": 0.0525, "metodo_fonte": "ICTI [SGS-433]"},
            ],
        },
        "vta_sombra": {"parcelas_computadas": [
            {"identificador": "financeiro:C1:base:2",
             "fonte_parcela": "Financeiro", "ciclo": "C1",
             "valor": 1000.0, "valor_atualizado": 1052.5, "linha": 2},
            {"identificador": "financeiro:C1:base:3",
             "fonte_parcela": "Financeiro", "ciclo": "C1",
             "valor": 2000.0, "valor_atualizado": 2105.0, "linha": 3},
            {"identificador": "aditivos:C1:4", "fonte_parcela": "Aditivo",
             "ciclo": "C1", "valor": 500.0, "linha": 4,
             "justificativa_vta": "Aditivo computavel (K=Sim)."},
        ]},
        "historico_vu": {"itens": [
            {"item": "1", "descricao": "Servico A", "vu_original": 10.0,
             "vu_ciclos": {"C1": 10.525}, "vu_vigente": 10.525,
             "fator_acumulado": 1.0525},
        ]},
        "itens_contrato": {"itens": [
            {"item": "1", "descricao": "Servico A", "qtd_contratada": 100.0,
             "vu_original": 10.0, "valor_total_original": 1000.0},
        ]},
    }


def leitura_multiciclo_pc():
    """Cenario 2: multiciclo (C1+C2) pelo metodo Pedido de Compras (IST)."""
    por_ciclo = {
        "C0": _ciclo_param("C0", date(2023, 5, 1), date(2024, 4, 30),
                           fator=1.0, situacao="Base"),
        "C1": _ciclo_param("C1", date(2024, 5, 1), date(2025, 4, 30),
                           pct=0.031, fator=1.031, computar="Sim",
                           inicio_efeito=date(2024, 5, 1), situacao="Computado"),
        "C2": _ciclo_param("C2", date(2025, 5, 1), date(2026, 4, 30),
                           pct=0.028, fator=1.059868, computar="Sim",
                           inicio_efeito=date(2025, 8, 1), situacao="Computado"),
        "C3": _ciclo_param("C3", date(2026, 5, 1), date(2027, 4, 30)),
        "C4": _ciclo_param("C4", date(2027, 5, 1), date(2028, 4, 30)),
    }
    itens_pc = [
        {"numero_pc": "4500001", "data_pc": date(2024, 7, 10),
         "valor_pc": 800.0, "valor_pago": 800.0, "linha": 2,
         "pc_pago_a_contratada": "Sim",
         "elegivel_retroativo_pc": True, "efeito_financeiro_pc": "Sim",
         "status_pagamento_pc": "PAGO"},
        {"numero_pc": "4500002", "data_pc": date(2025, 9, 20),
         "valor_pc": 1200.0, "valor_pago": 1200.0, "linha": 3,
         "pc_pago_a_contratada": "Sim",
         "elegivel_retroativo_pc": True, "efeito_financeiro_pc": "Sim",
         "status_pagamento_pc": "PAGO"},
    ]
    memoria_ist = [
        {"tipo": "INDICE", "ordem": 1, "competencia": "2024-05-01",
         "valor_indice": 104.37},
        {"tipo": "INDICE", "ordem": 2, "competencia": "2025-04-01",
         "valor_indice": 107.60},
        {"tipo": "RESULTADO", "ordem": 3, "fator_acumulado": 1.031,
         "variacao_final": 0.031, "metodo_fonte": "IST [ist.csv]"},
    ]
    memoria_c2 = [
        {"tipo": "INDICE", "ordem": 1, "competencia": "2025-05-01",
         "valor_indice": 107.60},
        {"tipo": "INDICE", "ordem": 2, "competencia": "2026-04-01",
         "valor_indice": 110.61},
        {"tipo": "RESULTADO", "ordem": 3, "fator_acumulado": 1.059868,
         "variacao_final": 0.028, "metodo_fonte": "IST [ist.csv]"},
    ]
    return {
        "ok": True,
        "controle": {"modo": "PC", "versao": "V10.4.8",
                     "ciclo_vigente": "C2", "data_corte": date(2026, 6, 30)},
        "resumo": {},
        "parametros_v10": {"ok": True, "por_ciclo": por_ciclo,
                           "ciclos": list(por_ciclo.values())},
        "memoria_calculo": {"C1": memoria_ist, "C2": memoria_c2},
        "itens_pc_v10": {"ok": True, "itens": itens_pc},
        "historico_vu": {"itens": []},
    }


def leitura_ausencias():
    """Cenario 3: dados parcialmente ausentes; nada vira zero."""
    por_ciclo = {
        "C0": _ciclo_param("C0", date(2024, 2, 1), date(2025, 1, 31),
                           fator=1.0),
        "C1": _ciclo_param("C1", date(2025, 2, 1), date(2026, 1, 31),
                           pct=None, fator=None, computar="Sim"),
    }
    return {
        "ok": True,
        "controle": {"modo": "SIMPLES", "versao": "V10.4.8",
                     "ciclo_vigente": "C1"},
        "resumo": {},
        "parametros_v10": {"ok": True, "por_ciclo": por_ciclo,
                           "ciclos": list(por_ciclo.values())},
        "historico_vu": {"itens": []},
    }


def _dados(leitura, **kwargs):
    dados = montar_dados_sumario_executivo(leitura, **kwargs)
    assert dados["disponivel"], dados.get("motivo")
    return dados


def _ciclo(dados, nome):
    return next(c for c in dados["ciclos"] if c["ciclo"] == nome)


def _texto_pdf(pdf: bytes) -> str:
    fitz = pytest.importorskip("fitz")
    doc = fitz.open(stream=pdf, filetype="pdf")
    return "\n".join(page.get_text() for page in doc)


# ---------------------------------------------------------------------------
# R1 - Indice contratual
# ---------------------------------------------------------------------------

def test_r1_indice_contratual_da_memoria():
    dados = _dados(leitura_simples_financeiro())
    assert dados["identificacao"]["indice"] == "ICTI [SGS-433]"


def test_r1_indice_ausente_nao_informado():
    dados = _dados(leitura_ausencias())
    assert dados["identificacao"]["indice"] == NAO_INFORMADO


# ---------------------------------------------------------------------------
# R2 - Datas dos pedidos de reajuste
# ---------------------------------------------------------------------------

def test_r2_datas_pedido_por_ciclo():
    dados = _dados(
        leitura_simples_financeiro(),
        identificacao={"datas_pedido": {"C1": date(2025, 3, 20)}},
    )
    assert _ciclo(dados, "C1")["data_pedido"] == "20/03/2025"
    assert _ciclo(dados, "C0")["data_pedido"] == NAO_INFORMADO


def test_r2_data_pedido_ausente_nao_informado():
    dados = _dados(leitura_simples_financeiro())
    assert _ciclo(dados, "C1")["data_pedido"] == NAO_INFORMADO


# ---------------------------------------------------------------------------
# R3/R4 - Ciclos analisados e periodos (ordem cronologica, C0 distinto)
# ---------------------------------------------------------------------------

def test_r3_r4_ciclos_e_periodos():
    dados = _dados(leitura_multiciclo_pc())
    nomes = [c["ciclo"] for c in dados["ciclos"]]
    assert nomes == ["C0", "C1", "C2", "C3", "C4"]  # ordem cronologica
    c0 = _ciclo(dados, "C0")
    assert c0["eh_base"] is True
    assert c0["percentual_reajuste"] == NAO_APLICAVEL  # C0 distinto dos ciclos
    c1 = _ciclo(dados, "C1")
    assert c1["data_inicio"] == "01/05/2024"
    assert c1["data_fim"] == "30/04/2025"


# ---------------------------------------------------------------------------
# R5 - Meses sem efeito financeiro, por ciclo
# ---------------------------------------------------------------------------

def test_r5_meses_sem_efeito_identificados():
    dados = _dados(leitura_simples_financeiro())
    bloco = _ciclo(dados, "C1")["meses_sem_efeito"]
    assert bloco["status"] == "ok"
    assert bloco["quantidade"] == 2
    assert bloco["competencias"] == ["02/2025", "03/2025"]


def test_r5_sem_gap_zero_meses():
    dados = _dados(leitura_multiciclo_pc())
    bloco = _ciclo(dados, "C1")["meses_sem_efeito"]
    assert bloco["quantidade"] == 0 and bloco["competencias"] == []
    bloco_c2 = _ciclo(dados, "C2")["meses_sem_efeito"]
    assert bloco_c2["competencias"] == ["05/2025", "06/2025", "07/2025"]


def test_r5_c0_nao_aplicavel_e_ausente_nao_informado():
    dados = _dados(leitura_ausencias())
    assert _ciclo(dados, "C0")["meses_sem_efeito"]["status"] == NAO_APLICAVEL
    bloco = _ciclo(dados, "C1")["meses_sem_efeito"]
    assert bloco["status"] == NAO_INFORMADO
    assert bloco["quantidade"] is None  # ausencia nao vira zero


# ---------------------------------------------------------------------------
# R6 - Variacao percentual por ciclo e acumulada
# ---------------------------------------------------------------------------

def test_r6_variacao_por_ciclo_e_acumulada():
    dados = _dados(leitura_multiciclo_pc())
    assert _ciclo(dados, "C1")["percentual_reajuste"] == pytest.approx(0.031)
    assert _ciclo(dados, "C2")["percentual_reajuste"] == pytest.approx(0.028)
    sintese = dados["sintese"]
    assert sintese["variacao_acumulada"] == pytest.approx(0.059868)
    assert sintese["ciclo_referencia_acumulado"] == "C2"


def test_r6_variacao_ausente_nao_vira_zero():
    dados = _dados(leitura_ausencias())
    assert _ciclo(dados, "C1")["percentual_reajuste"] is None
    assert dados["sintese"]["variacao_acumulada"] is None


# ---------------------------------------------------------------------------
# R7/R8/R9 - Valor pago, delta e retroativo reconhecido (Financeiro)
# ---------------------------------------------------------------------------

def test_r7_r8_r9_metodo_financeiro():
    dados = _dados(leitura_simples_financeiro())
    fin = dados["financeiro"]
    linhas = fin["financeiro_por_ciclo"]
    assert len(linhas) == 1
    linha = linhas[0]
    assert linha["ciclo"] == "C1"
    assert linha["valor_pago"] == pytest.approx(3000.0)      # R7
    assert linha["valor_atualizado"] == pytest.approx(3157.5)
    assert linha["delta"] == pytest.approx(157.5)            # R8 por ciclo
    assert fin["delta_total_financeiro"] == pytest.approx(157.5)  # R8 total


# ---------------------------------------------------------------------------
# R10 - Retroativo em analise no metodo PC, separado do reconhecido
# ---------------------------------------------------------------------------

def test_r10_pc_em_analise_separado():
    dados = _dados(leitura_multiciclo_pc())
    fin = dados["financeiro"]
    assert fin["financeiro_por_ciclo"] == []  # nao mistura metodos
    pc = {linha["ciclo"]: linha for linha in fin["pc_por_ciclo"]}
    assert set(pc) == {"C1", "C2"}
    assert pc["C1"]["valor_pago"] == pytest.approx(800.0)
    assert pc["C1"]["delta"] == pytest.approx(round(800.0 * 1.031 - 800.0, 2))
    assert fin["delta_total_pc"] == pytest.approx(
        pc["C1"]["delta"] + pc["C2"]["delta"]
    )


def test_r10_ausencia_de_metodo_nao_vira_zero():
    dados = _dados(leitura_ausencias())
    fin = dados["financeiro"]
    assert fin["financeiro_por_ciclo"] == []
    assert fin["pc_por_ciclo"] == []
    assert fin["delta_total_financeiro"] is None
    assert fin["delta_total_pc"] is None


# ---------------------------------------------------------------------------
# R11 - Itens: VU e total em C0 e por ciclo
# ---------------------------------------------------------------------------

def test_r11_itens_vu_e_totais():
    dados = _dados(leitura_simples_financeiro())
    itens = dados["itens"]
    assert len(itens) == 1
    item = itens[0]
    assert item["vu_c0"] == pytest.approx(10.0)
    assert item["total_c0"] == pytest.approx(1000.0)
    assert item["vu_ciclos"]["C1"] == pytest.approx(10.525)
    assert item["total_ciclos"]["C1"] == pytest.approx(1052.5)


def test_r11_sem_itens_lista_vazia():
    dados = _dados(leitura_ausencias())
    assert dados["itens"] == []


# ---------------------------------------------------------------------------
# R12 - Memoria de calculo por ciclo (ICTI/SGS/IST, simples e multiciclo)
# ---------------------------------------------------------------------------

def test_r12_memoria_icti_simples():
    dados = _dados(leitura_simples_financeiro())
    memoria = dados["memoria_calculo"]
    assert [b["ciclo"] for b in memoria] == ["C1"]
    registros = memoria[0]["registros"]
    assert registros[0]["tipo"] == "MES"
    assert registros[0]["competencia"] == "02/2025"
    resultado = registros[-1]
    assert resultado["tipo"] == "RESULTADO"
    assert resultado["variacao_final"] == pytest.approx(0.0525)
    assert resultado["metodo_fonte"] == "ICTI [SGS-433]"


def test_r12_memoria_ist_multiciclo():
    dados = _dados(leitura_multiciclo_pc())
    memoria = {b["ciclo"]: b["registros"] for b in dados["memoria_calculo"]}
    assert set(memoria) == {"C1", "C2"}
    assert memoria["C1"][0]["tipo"] == "INDICE"
    assert memoria["C1"][0]["valor_indice"] == pytest.approx(104.37)
    assert memoria["C2"][-1]["metodo_fonte"] == "IST [ist.csv]"


def test_r12_legado_sem_memoria():
    dados = _dados(leitura_ausencias())
    assert dados["memoria_calculo"] == []  # tabela vazia nao e exibida


# ---------------------------------------------------------------------------
# R13 - Aditivos aplicaveis
# ---------------------------------------------------------------------------

def test_r13_aditivo_computavel_exibido():
    dados = _dados(leitura_simples_financeiro())
    itens = dados["aditivos"]["itens"]
    assert len(itens) == 1
    assert itens[0]["ciclo"] == "C1"
    assert itens[0]["valor_atualizado"] == pytest.approx(500.0)
    assert itens[0]["anterior_formalizacao"] == "Sim"


def test_r13_sem_aditivos_nao_aplicavel():
    dados = _dados(leitura_ausencias())
    assert dados["aditivos"]["itens"] == []


def test_r13_aditivo_sem_inicio_efeito_nao_informado():
    leitura = leitura_simples_financeiro()
    leitura["parametros_v10"]["por_ciclo"]["C1"]["inicio_efeito_financeiro"] = None
    dados = _dados(leitura)
    assert dados["aditivos"]["itens"][0]["anterior_formalizacao"] == NAO_INFORMADO


# ---------------------------------------------------------------------------
# Cabecalho e formatadores
# ---------------------------------------------------------------------------

def test_cabecalho_identificacao_campos_permitidos():
    """Somente campos nao-proibidos existem na identificacao canonica."""
    dados = _dados(leitura_simples_financeiro())
    ident = dados["identificacao"]
    # Campos que devem existir
    assert ident["indice"] == "ICTI [SGS-433]"
    assert ident["metodo"] != ""
    assert ident["data_corte"] == "30/06/2026"
    assert "ciclo_vigente" in ident
    assert "gerado_em" in ident
    # Campos proibidos nao devem existir no dict de dados
    for campo in ("empresa", "contrato", "processo",
                  "versao_masterfile", "objeto_processo_id"):
        assert campo not in ident, f"Campo proibido presente: {campo}"


def test_formatadores_exibicao():
    assert formatar_moeda(1234567.891) == "R$ 1.234.567,89"
    assert formatar_moeda(None) == NAO_INFORMADO  # ausencia nao vira R$ 0,00
    assert formatar_percentual(0.0525) == "5,2500%"
    assert formatar_percentual(None) == NAO_INFORMADO


def test_leitura_invalida_indisponivel():
    dados = montar_dados_sumario_executivo({"ok": False})
    assert dados["disponivel"] is False
    assert dados["motivo"]
    pdf = gerar_sumario_executivo_pdf(dados)
    assert pdf.startswith(b"%PDF-")


# ---------------------------------------------------------------------------
# PDF: validade, pesquisabilidade, multiplas paginas, cabecalhos repetidos
# ---------------------------------------------------------------------------

def test_pdf_valido_e_pesquisavel():
    pdf = gerar_sumario_executivo(leitura_simples_financeiro())
    assert pdf.startswith(b"%PDF-")
    assert b"%%EOF" in pdf[-1024:]
    texto = _texto_pdf(pdf)
    assert "Sumário Executivo" in texto
    for secao in ("1. Identificação", "2. Síntese da apuração",
                  "3. Ciclos e efeitos financeiros", "4. Valores financeiros",
                  "5. Itens e valores atualizados", "6. Memória de cálculo",
                  "7. Aditivos aplicáveis"):
        assert secao in texto, secao


def test_pdf_multiplas_paginas_com_cabecalho_repetido():
    fitz = pytest.importorskip("fitz")
    leitura = leitura_multiciclo_pc()
    # Infla a memoria (cenario valido: ate 79 linhas no bloco J2:R80) para
    # forcar quebra de pagina na tabela da memoria de calculo.
    leitura["memoria_calculo"]["C1"] = [
        {"tipo": "MES", "ordem": i, "competencia": f"2024-{(i % 12) + 1:02d}-15",
         "valor_indice": 0.003}
        for i in range(1, 60)
    ] + [{"tipo": "RESULTADO", "ordem": 60, "fator_acumulado": 1.031,
          "variacao_final": 0.031, "metodo_fonte": "IST [ist.csv]"}]
    pdf = gerar_sumario_executivo(leitura)
    doc = fitz.open(stream=pdf, filetype="pdf")
    assert doc.page_count >= 3
    # rodape com somente "Gerado em dd/mm/aaaa." em todas as paginas
    for page in doc:
        texto = page.get_text()
        assert "Gerado em" in texto
        assert "Página" not in texto  # numero de pagina removido
    # cabecalho da tabela de memoria repetido na pagina seguinte a quebra
    paginas_com_cabecalho = [
        n for n, page in enumerate(doc, start=1)
        if "Fatoracumulado" in page.get_text().replace("\n", "").replace(" ", "")
        and "Competência" in page.get_text().replace("\n", "")
    ]
    assert len(paginas_com_cabecalho) >= 2


def test_pdf_ausencias_sem_tabelas_vazias():
    pdf = gerar_sumario_executivo(leitura_ausencias())
    texto = _texto_pdf(pdf)
    assert "Não aplicável" in texto
    assert "Não informado" in texto
    # nada de zero inventado nos quadros financeiros ausentes
    assert "R$ 0,00" not in texto


# ---------------------------------------------------------------------------
# Novos testes: campos proibidos, secao 8, rodape, deltas (Etapa 5 v2)
# ---------------------------------------------------------------------------

def test_pdf_campos_proibidos_ausentes():
    """Nenhum dos cinco campos proibidos deve aparecer no texto do PDF."""
    fitz = pytest.importorskip("fitz")
    pdf = gerar_sumario_executivo(
        leitura_simples_financeiro(),
        identificacao={
            "empresa": "ACME Telecom S.A.",
            "contrato": "CT-001/2024",
            "processo": "PROC-2026/0001",
        },
    )
    texto = _texto_pdf(pdf)
    # Rotulos e valores dos campos proibidos
    for termo in ("Empresa", "Contrato", "Processo", "Versão do MasterFile",
                  "Objeto do processo", "ACME Telecom S.A.", "CT-001/2024",
                  "PROC-2026/0001"):
        assert termo not in texto, f"Campo proibido encontrado no PDF: {termo!r}"


def test_pdf_sem_secao_observacoes():
    """Secao '8. Observacoes de consistencia' nao deve aparecer no PDF."""
    fitz = pytest.importorskip("fitz")
    pdf = gerar_sumario_executivo(leitura_simples_financeiro())
    texto = _texto_pdf(pdf)
    assert "Observações de consistência" not in texto
    assert "8. Observações" not in texto
    # O documento termina na secao 7
    assert "7. Aditivos aplicáveis" in texto


def test_pdf_rodape_somente_data():
    """Rodape deve conter exclusivamente 'Gerado em dd/mm/aaaa.'"""
    fitz = pytest.importorskip("fitz")
    import re
    pdf = gerar_sumario_executivo(leitura_simples_financeiro())
    doc = fitz.open(stream=pdf, filetype="pdf")
    padrao_data = re.compile(r"Gerado em \d{2}/\d{2}/\d{4}\.")
    for page in doc:
        texto = page.get_text()
        assert padrao_data.search(texto), (
            f"Rodape sem 'Gerado em dd/mm/aaaa.': {texto!r}"
        )
        # Nenhum outro conteudo no rodape
        assert "Página" not in texto
        assert "Cl8us" not in texto


def test_deltas_preservados_no_pdf():
    """Delta por ciclo e delta total devem estar presentes no PDF."""
    fitz = pytest.importorskip("fitz")
    pdf = gerar_sumario_executivo(leitura_simples_financeiro())
    texto = _texto_pdf(pdf)
    # Delta R$ 157,50 deve aparecer (metodo financeiro C1)
    assert "157,50" in texto
    # Total (delta) deve aparecer como rotulo de linha
    assert "Total (delta)" in texto


def test_deltas_multiciclo_pc_preservados():
    """Deltas do metodo PC por ciclo preservados no PDF multiciclo."""
    fitz = pytest.importorskip("fitz")
    pdf = gerar_sumario_executivo(leitura_multiciclo_pc())
    texto = _texto_pdf(pdf)
    # Delta C1 PC: 800 * 1.031 - 800 = 24.80
    assert "24,80" in texto
    assert "Total (delta)" in texto
