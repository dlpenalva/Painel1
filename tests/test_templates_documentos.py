"""Testes dos modelos canonicos de Termo de Apostila (§6/§10.2) e
Despacho Saneador (§7/§10.3).
"""
from __future__ import annotations

import sys
from io import BytesIO
from pathlib import Path

import pytest
from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from _templates_documentos import (  # noqa: E402
    diagnosticar_campos_manuais,
    gerar_despacho_saneador,
    gerar_termo_apostila,
    _extrair_dados,
    _ta_secao4_composicao_vta,
    _fmt_pct_doc,
)
from _sanitizacao_documental import contem_emoji  # noqa: E402
from test_sumario_executivo import (  # noqa: E402
    leitura_ausencias,
    leitura_multiciclo_pc,
    leitura_simples_financeiro,
)

# ---------------------------------------------------------------------------
# Campos manuais
# ---------------------------------------------------------------------------

CAMPOS_TERMO = {
    "contrato": "TLB-CTR-2025/00001",
    "empresa_contratada": "Empresa XPTO S.A., CNPJ 00.000.000/0001-00",
    "representante_telebras_1_nome": "Fulano de Tal",
    "representante_telebras_1_matricula": "12345",
    "representante_telebras_2_cargo": "Diretor Financeiro",
    "representante_telebras_2_matricula": "67890",
    "memoria_calculo_ref": "TLB-AUT-2026/00700",
    "concordancia_ref": "TLB-AUT-2026/00500",
    "regularidade_ref": "TLB-AUT-2026/00400",
    "adequacao_orcamentaria_ref": "TLB-DES-2026/00300",
    "processo_ref": "TLB-PRO-2026/01100",
    "valor_original_contrato": 1000000.0,
    "local_data": "20/07/2026",
}

CAMPOS_SANEADOR = {
    "contrato": "TLB-CTR-2025/00001",
    "processo_pleito": "TLB-AUT-2025/00100",
    "data_proposta": "02/08/2023",
    "referencia_analise": "TLB-AUT-2026/00200",
    "data_corte_descricao": "abril de 2026, último mês com liquidação",
    "adequacao_orcamentaria_ref": "TLB-DES-2026/00300",
    "adequacao_orcamentaria_valor": 123456.78,
    "regularidade_ref": "TLB-AUT-2026/00400",
    "concordancia_ref": "TLB-AUT-2026/00500",
    "valor_original_contrato": 1000000.0,
}

# Vocabulario de implementacao proibido nos documentos (§5).
TERMOS_TECNICOS_PROIBIDOS = [
    "Base executada do financeiro", "G não exclui base", "EFEITO_FINANCEIRO",
    "RESULTADOS!B23", "RESULTADOS!B26", "coluna G", "EFEITO_FINANCEIRO_PC",
    "QTD_REM_AJUSTADA", "fonte dupla", "sheet12.xml", "motor Python",
    "XLS × Python", "delta financeiro", "valor nominal", "base financeira",
    "parcelas_computadas", "vta_sombra", "fonte_parcela",
]


def _texto_docx(docx_bytes: bytes) -> str:
    doc = Document(BytesIO(docx_bytes))
    partes = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                partes.append(cell.text)
    return "\n".join(partes)


def _titulos_quadros(docx_bytes: bytes) -> list[str]:
    doc = Document(BytesIO(docx_bytes))
    return [" | ".join(c.text for c in t.rows[0].cells) for t in doc.tables]


# ---------------------------------------------------------------------------
# Validade basica
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("gerador,leitura", [
    (gerar_despacho_saneador, leitura_simples_financeiro),
    (gerar_termo_apostila, leitura_simples_financeiro),
    (gerar_despacho_saneador, leitura_multiciclo_pc),
    (gerar_termo_apostila, leitura_multiciclo_pc),
    (gerar_despacho_saneador, leitura_ausencias),
    (gerar_termo_apostila, leitura_ausencias),
])
def test_docx_valido(gerador, leitura):
    b = gerador(leitura())
    assert isinstance(b, bytes) and b[:2] == b"PK" and len(b) > 100
    Document(BytesIO(b))  # abre sem erro


# ---------------------------------------------------------------------------
# §10.2 — APOSTILA
# ---------------------------------------------------------------------------

def test_apostila_titulo_exato():
    texto = _texto_docx(gerar_termo_apostila(leitura_simples_financeiro(), campos_manuais=CAMPOS_TERMO))
    assert "MINUTA DE TERMO DE APOSTILAMENTO" in texto


def test_apostila_qualificacao_canonica():
    texto = _texto_docx(gerar_termo_apostila(leitura_multiciclo_pc(), campos_manuais=CAMPOS_TERMO))
    assert "sociedade de economia mista" in texto
    assert "00.336.701/0001-04" in texto
    assert "SIG, Quadra 04, Bloco A" in texto
    assert "70.610-440" in texto
    assert "parágrafo 7º do art. 81 da Lei nº 13.303" in texto
    assert "Diretriz nº 229/2018" in texto
    # Nao pode conter a qualificacao antiga divergente
    assert "33.200.056/0001-41" not in texto
    assert "empresa pública federal" not in texto
    assert "SAS Quadra 05" not in texto


def test_apostila_sete_considerandos():
    texto = "\n" + _texto_docx(gerar_termo_apostila(leitura_multiciclo_pc(), campos_manuais=CAMPOS_TERMO))
    assert "CONSIDERANDO:" in texto
    for n in range(1, 8):
        assert f"\n{n}. " in texto, f"considerando {n} ausente"


def test_apostila_secoes_1_a_9():
    texto = _texto_docx(gerar_termo_apostila(leitura_multiciclo_pc(), campos_manuais=CAMPOS_TERMO))
    assert "FORMALIZA-SE O PRESENTE TERMO DE APOSTILA:" in texto
    assert "1. Dos reajustes concedidos" in texto
    assert "2. Da apuração financeira do retroativo" in texto
    assert "3. Da memória fiscal do Valor Total Atualizado" in texto
    assert "4. Da composição sintética do Valor Total Atualizado" in texto
    assert "5. Dos valores unitários" in texto
    assert "6. Dos aditivos e supressões considerados" in texto
    assert "Termo de Apostila." in texto              # §7
    assert "atualizar a garantia contratual" in texto  # §8
    assert "vincula-se, para todos os fins" in texto   # §9


def test_apostila_quadros_1_a_4_presentes():
    quadros = _titulos_quadros(gerar_termo_apostila(leitura_multiciclo_pc(), campos_manuais=CAMPOS_TERMO))
    assert "Ref. | Ciclo | Percentual aplicado | Efeitos financeiros | Situação" in quadros  # Q1
    assert "Ciclo | Valor pago efetivo | Valor teórico calculado | Diferença/retroativo" in quadros  # Q2
    assert "Ref. | Descrição | Valor" in quadros  # Q3
    assert "Ref. | Parcela | Valor" in quadros  # Q4


def test_apostila_vu_ate_ultimo_ciclo_sem_futuro():
    b = gerar_termo_apostila(leitura_simples_financeiro(), campos_manuais=CAMPOS_TERMO)
    quadros = _titulos_quadros(b)
    vu = [q for q in quadros if "VU C0" in q]
    assert vu, "tabela de VU ausente"
    assert "VU C1" in vu[0]
    assert "VU C2" not in vu[0]  # ciclo futuro nao entra
    assert "VU C3" not in vu[0]


def test_apostila_duas_assinaturas_telebras_sem_contratada():
    texto = _texto_docx(gerar_termo_apostila(leitura_simples_financeiro(), campos_manuais=CAMPOS_TERMO))
    assert texto.count("TELECOMUNICAÇÕES BRASILEIRAS S.A. - TELEBRAS") >= 3  # qualificacao + 2 assinaturas
    linhas = texto.splitlines()
    assert not any(l.strip() == "CONTRATADA" for l in linhas)


def test_apostila_equacao_sintetica_nao_soma_vta_a_si_mesmo():
    dados = _extrair_dados(leitura_simples_financeiro(), None)
    dados["parcelas_vta"] = [
        {"fonte_parcela": "Financeiro", "ciclo": "C0", "valor_atualizado": 1000.0},
        {"fonte_parcela": "Financeiro", "ciclo": "C1", "valor_atualizado": 2000.0},
        {"fonte_parcela": "Aditivo", "ciclo": "C1", "valor": 500.0},
    ]
    dados["vta"] = 3500.0
    doc = Document()
    _ta_secao4_composicao_vta(doc, dados)
    texto = "\n".join(p.text for p in doc.paragraphs)
    assert "A + B + C = D = R$ 3.500,00." in texto
    tabela = doc.tables[0]
    valores = [row.cells[-1].text for row in tabela.rows]
    assert valores.count("R$ 3.500,00") == 2       # ref_vta + Total
    assert "R$ 7.000,00" not in "\n".join(valores)  # nunca dobra o VTA


def test_apostila_sem_termos_tecnicos_e_sem_emoji():
    for leit in (leitura_simples_financeiro, leitura_multiciclo_pc):
        texto = _texto_docx(gerar_termo_apostila(leit(), campos_manuais=CAMPOS_TERMO))
        assert not contem_emoji(texto)
        for termo in TERMOS_TECNICOS_PROIBIDOS:
            assert termo not in texto, f"termo tecnico proibido: {termo}"


# ---------------------------------------------------------------------------
# §10.3 — SANEADOR
# ---------------------------------------------------------------------------

def test_saneador_assunto():
    texto = _texto_docx(gerar_despacho_saneador(leitura_simples_financeiro(), campos_manuais=CAMPOS_SANEADOR))
    assert "DESPACHO SANEADOR" in texto
    assert "Saneamento para formalização de Termo de Apostila de Reajuste - TLB-CTR-2025/00001" in texto


def test_saneador_sequencia_itens():
    texto = "\n" + _texto_docx(gerar_despacho_saneador(leitura_multiciclo_pc(), campos_manuais=CAMPOS_SANEADOR))
    for n in range(1, 13):
        assert f"\n{n}. " in texto, f"item {n} ausente"


def test_saneador_quadros_e_composicao():
    b = gerar_despacho_saneador(leitura_multiciclo_pc(), campos_manuais=CAMPOS_SANEADOR)
    texto = _texto_docx(b)
    quadros = _titulos_quadros(b)
    assert "Quadro 1 - Síntese dos ciclos de reajuste" in texto
    assert "Ciclo | Data-base | Data do pedido | Início financeiro | Fim financeiro | Situação | Percentual aplicado" in quadros
    assert "Ciclo | Valor pago efetivo | Valor teórico calculado | Diferença/retroativo" in quadros
    assert "Descrição | Valor" in quadros  # Quadro 3 (sem coluna Ref.)
    assert "Quadro 4 - Síntese dos principais valores" in texto
    assert "De forma didática" in texto  # composicao didatica (par. 7)


def test_saneador_inicio_financeiro_usa_efeito_real():
    b = gerar_despacho_saneador(leitura_multiciclo_pc(), campos_manuais=CAMPOS_SANEADOR)
    tabela = Document(BytesIO(b)).tables[0]  # Quadro 1
    linha_c2 = next(r for r in tabela.rows if r.cells[0].text == "C2")
    assert linha_c2.cells[1].text == "01/05/2025"   # Data-base
    assert linha_c2.cells[3].text == "01/08/2025"   # Inicio financeiro (efeito real)


def test_saneador_sem_tabela_de_valores_unitarios():
    texto = _texto_docx(gerar_despacho_saneador(leitura_simples_financeiro(), campos_manuais=CAMPOS_SANEADOR))
    assert "VU C0" not in texto
    assert "Valores Unitários" not in texto


def test_saneador_itens_administrativos_presentes():
    texto = _texto_docx(gerar_despacho_saneador(leitura_multiciclo_pc(), campos_manuais=CAMPOS_SANEADOR))
    assert "adequação orçamentária" in texto
    assert "certidões de regularidade" in texto
    assert "manifestou concordância" in texto
    assert "garantia contratual" in texto
    assert "aditivos e supressões" in texto


def test_saneador_conclusao_normal_afirma_inexistencia():
    texto = _texto_docx(gerar_despacho_saneador(leitura_multiciclo_pc(), campos_manuais=CAMPOS_SANEADOR))
    assert "inexistindo pendência crítica" in texto


def test_saneador_softblock_nao_afirma_inexistencia():
    cm = dict(CAMPOS_SANEADOR, pendencia_critica=True)
    texto = _texto_docx(gerar_despacho_saneador(leitura_multiciclo_pc(), campos_manuais=cm))
    assert "inexistindo pendência crítica" not in texto
    assert "permanecendo pendentes" in texto
    assert "consolidados para análise" in texto


def test_saneador_conclusao_numerada_13_sem_docs_desatualizados():
    texto = _texto_docx(gerar_despacho_saneador(leitura_multiciclo_pc(), campos_manuais=CAMPOS_SANEADOR))
    # Sem docs_desatualizados: 12 (garantia) -> 13 (conclusao) -> Quadro 4
    assert "13. Diante do exposto" in texto
    assert "14. " not in texto
    assert "mostram-se desatualizados" not in texto  # item 13 de docs ausente


def test_saneador_conclusao_numerada_14_com_docs_desatualizados():
    cm = dict(CAMPOS_SANEADOR, docs_desatualizados=["SEI 999/2026", "SEI 888/2026"])
    texto = _texto_docx(gerar_despacho_saneador(leitura_multiciclo_pc(), campos_manuais=cm))
    # Com docs_desatualizados: 13 (documentos) -> 14 (conclusao)
    assert "13. Após atualizações" in texto
    assert "mostram-se desatualizados, devendo ser desconsiderados" in texto
    assert "14. Diante do exposto" in texto
    assert "13. Diante do exposto" not in texto


def test_saneador_conclusao_numerada_respeita_softblock():
    # Numeracao correta mesmo sob soft-block, sem afirmar inexistencia de pendencia.
    cm_sem = dict(CAMPOS_SANEADOR, pendencia_critica=True)
    texto_sem = _texto_docx(gerar_despacho_saneador(leitura_multiciclo_pc(), campos_manuais=cm_sem))
    assert "13. Diante do exposto" in texto_sem
    assert "permanecendo pendentes" in texto_sem
    assert "inexistindo pendência crítica" not in texto_sem

    cm_com = dict(CAMPOS_SANEADOR, pendencia_critica=True, docs_desatualizados=["SEI 999/2026"])
    texto_com = _texto_docx(gerar_despacho_saneador(leitura_multiciclo_pc(), campos_manuais=cm_com))
    assert "14. Diante do exposto" in texto_com
    assert "permanecendo pendentes" in texto_com
    assert "inexistindo pendência crítica" not in texto_com


def test_saneador_sem_termos_tecnicos_e_sem_emoji():
    for leit in (leitura_simples_financeiro, leitura_multiciclo_pc):
        texto = _texto_docx(gerar_despacho_saneador(leit(), campos_manuais=CAMPOS_SANEADOR))
        assert not contem_emoji(texto)
        for termo in TERMOS_TECNICOS_PROIBIDOS:
            assert termo not in texto, f"termo tecnico proibido: {termo}"


# ---------------------------------------------------------------------------
# Robustez comum
# ---------------------------------------------------------------------------

def test_ausencias_nao_viram_zero():
    for gerador in (gerar_termo_apostila, gerar_despacho_saneador):
        texto = _texto_docx(gerador(leitura_ausencias(), campos_manuais={}))
        assert "R$ 0,00" not in texto or "[PREENCHER:" in texto


def test_campos_ausentes_geram_preencher():
    for gerador in (gerar_termo_apostila, gerar_despacho_saneador):
        texto = _texto_docx(gerador(leitura_simples_financeiro(), campos_manuais={}))
        assert "[PREENCHER:" in texto


def test_diagnostico_pendencias():
    pend = diagnosticar_campos_manuais(leitura_simples_financeiro(), campos_manuais=None)
    assert isinstance(pend, list) and pend
    for item in pend:
        assert {"campo", "descricao", "documento"} <= set(item)
    campos = dict(CAMPOS_TERMO, **CAMPOS_SANEADOR)
    pend2 = [p["campo"] for p in diagnosticar_campos_manuais(leitura_simples_financeiro(), campos_manuais=campos)]
    for chave in campos:
        assert chave not in pend2


def test_sem_dados_padtec_hardcoded():
    for gerador in (gerar_termo_apostila, gerar_despacho_saneador):
        texto = _texto_docx(gerador(leitura_multiciclo_pc(), campos_manuais={}))
        for proibida in ("TLB-CTR-2022/00067", "PADTEC", "158.292.598"):
            assert proibida not in texto


def test_fmt_pct_doc():
    assert _fmt_pct_doc(0.04) == "4,00%"
    assert _fmt_pct_doc(0.106231) == "10,62%"
    assert _fmt_pct_doc(-0.02) == "-2,00%"
