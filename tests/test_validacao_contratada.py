"""Testes focais da funcionalidade aditiva "Validação com a contratada".

Cobre a tarefa: comunicado pós-Coleta para conferência da apuração pela
contratada antes da formalização da Apostila. A funcionalidade é somente
aditiva — os testes de regressão que já cobrem os 4 cards, o box "Resultado
da apuração" e os 6 cards de documentos (test_upload_docs_flow.py,
test_status_confiabilidade_ui.py, test_resultado_consolidado.py) continuam
passando sem alteração; este arquivo cobre apenas o bloco novo.
"""
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

PAGINA = (ROOT / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")

_DOCS_SEIS_CARDS = {
    chave: {
        "nome": titulo, "estado": "completo", "rotulo": "Disponível",
        "classificacao": "", "motivo": "", "habilitado": True,
    }
    for chave, titulo in (
        ("sumario_executivo", "Sumário Executivo"),
        ("adequacao_orcamentaria", "Adequação Orçamentária"),
        ("despacho_saneador", "Despacho Saneador"),
        ("termo_apostila", "Termo de Apostila"),
        ("garantia_contratual", "Garantia Contratual"),
        ("dou", "DOU"),
    )
}


class TestEstruturaNaPagina(unittest.TestCase):
    """Assertivas estruturais sobre o código-fonte da página (mesmo padrão
    já usado por test_upload_docs_flow.py e test_status_confiabilidade_ui.py
    nesta base para este arquivo)."""

    def test_bloco_aparece_depois_dos_seis_cards_e_antes_do_stop(self):
        guarda = PAGINA.index('if st.session_state.get("assinatura_processada_upload_docs")')
        render_docs = PAGINA.index("render_documentos_funcionais_upload(resultado)", guarda)
        render_validacao = PAGINA.index("render_validacao_contratada(resultado, diagnostico_coleta)", render_docs)
        parada = PAGINA.index("st.stop()", render_validacao)
        self.assertLess(guarda, render_docs)
        self.assertLess(render_docs, render_validacao)
        self.assertLess(render_validacao, parada)

    def test_quatro_cards_de_resumo_preservados(self):
        trecho = PAGINA[PAGINA.index('resumo_indice, resumo_ciclos, resumo_retro, resumo_acum = st.columns(4)'):]
        trecho = trecho[:trecho.index("render_resultado_consolidado(resultado, diagnostico_coleta)")]
        self.assertIn('resumo_indice.metric("Índice"', trecho)
        self.assertIn('resumo_ciclos.metric("Ciclos analisados"', trecho)
        self.assertIn('resumo_retro.metric("Retroativo reconhecido"', trecho)
        self.assertIn('resumo_acum.metric("Percentual acumulado"', trecho)

    @staticmethod
    def _bloco_render():
        inicio = PAGINA.index("def render_validacao_contratada")
        fim = PAGINA.index("# <<< VALIDACAO_CONTRATADA_POS_COLETA_V1", inicio)
        return PAGINA[inicio:fim]

    def test_titulo_e_texto_do_bloco(self):
        bloco = self._bloco_render()
        self.assertIn('st.markdown("### Validação com a contratada")', bloco)
        self.assertIn(
            "Comunicado para conferência da apuração antes da formalização da Apostila.",
            bloco,
        )

    def test_acoes_visualizar_e_baixar_txt(self):
        bloco = self._bloco_render()
        self.assertIn('with st.expander("Visualizar comunicado")', bloco)
        self.assertIn("st.code(texto_comunicado, language=None)", bloco)
        self.assertIn('"Baixar TXT"', bloco)
        self.assertIn("data=texto_comunicado.encode(", bloco)

    def test_txt_deriva_da_mesma_string_exibida(self):
        # A garantia de "TXT identico ao texto exibido" e estrutural: st.code e
        # st.download_button leem a MESMA variavel texto_comunicado, calculada
        # uma unica vez por rerun -- sem widget com key propria, portanto sem a
        # classe de bug de congelamento ja vista na aba 5 (Texto SIGA).
        bloco = self._bloco_render()
        self.assertEqual(bloco.count("texto_comunicado = gerar_texto_validacao_contratada"), 1)
        self.assertNotIn("st.text_area", bloco)
        self.assertNotIn("st.session_state[", bloco)

    def test_nao_toca_nos_seis_cards_de_documentos(self):
        # render_documentos_funcionais_upload permanece exatamente como estava.
        assinatura = PAGINA[
            PAGINA.index("def render_documentos_funcionais_upload"):
            PAGINA.index("def _invalidar_caso_antes_do_rerun_upload")
        ]
        # UX-CARDS: as duas linhas de tres colunas passaram a ser geradas pelo
        # laco sobre GRUPOS_CARDS_UPLOAD; o que este teste guarda continua sendo
        # a grade 3+3 e a ausencia do bloco de validacao dentro dos cards.
        self.assertIn("for grupo, chaves_do_grupo in GRUPOS_CARDS_UPLOAD:", assinatura)
        self.assertIn("st.columns(3)", assinatura)
        self.assertNotIn("render_validacao_contratada", assinatura)


class TestAppTestValidacaoContratada(unittest.TestCase):
    """Prova end-to-end no runtime Streamlit (AppTest), mesmo padrão de
    tests/test_pacote_pos5casos_focais.py::TestAppTestStatusEDocumentos."""

    CAMINHO_PAGINA = str(ROOT / "pages" / "03_Valor_Global.py")

    def test_metodo_pc_com_potencial_e_perda_de_competencia(self):
        import pandas as pd
        from streamlit.testing.v1 import AppTest

        consolidado = {
            "vta": 500000.0,
            "vta_origem": "posicao_atual",
            "vta_usa_ultima_posicao": False,
            "retroativo_reconhecido": 12345.67,
            "valor_atualizado_em_analise": 1000.0,
            "retroativo_potencial": 890.12,
            "medidas_pc_aplicaveis": True,
            "fora_do_corte": {"aplicavel": True, "quantidade": None, "valor_informado": None, "data_corte": None},
            "metodo": {"codigo": "pc", "rotulo": "Pedidos de Compra"},
            "ciclo_vigente": "C2",
            "status_confiabilidade": "CONFIÁVEL",
            "mensagem_status": "Resultado apurado sem ressalvas materiais identificadas.",
            "formalizacao": {"bloqueada": False, "status": "SEM BLOQUEIO", "mensagem": "Sem bloqueio explícito à formalização."},
            "bloqueios": [],
            "ressalvas": [],
            "campos_nao_confiaveis": [],
            "composicao_vta": {"disponivel": False, "conciliada": None, "exibivel": False, "linhas": [], "mensagem": "", "ha_aditivos_nao_computados": False},
        }
        diagnostico = {
            "status_base": "OK",
            "pronto_para_consolidar": True,
            "metadados": {"indice": "IST", "ciclos_em_analise": ["C1", "C2"]},
        }
        resultado = {
            "capacidades": {"documentos": _DOCS_SEIS_CARDS},
            "diagnostico_coleta": diagnostico,
            "variacao_acumulada": 0.075,
            "indice": "IST",
            "resultado_consolidado": consolidado,
            "objeto_processo": {
                "contrato": {
                    "dados_basicos": {},
                    "datas_relevantes": {
                        "ciclos": [
                            {"ciclo": "C1", "data_inicio_ddmmaaaa": "01/01/2025", "data_fim_ddmmaaaa": "31/12/2025"},
                            {"ciclo": "C2", "data_inicio_ddmmaaaa": "01/01/2026", "data_fim_ddmmaaaa": "31/12/2026"},
                        ]
                    },
                },
                "dados_operacionais": {
                    "parametros_v10": {
                        "por_ciclo": {
                            "C1": {"data_inicio": "2025-01-01", "inicio_efeito_financeiro": "2025-02-01"},
                            "C2": {"data_inicio": "2026-01-01", "inicio_efeito_financeiro": "2026-01-01"},
                        }
                    }
                },
            },
            "df_ciclos": pd.DataFrame([
                {"Ciclo": "C1", "Variação": 0.045},
                {"Ciclo": "C2", "Variação": 0.030},
            ]),
            "df_financeiro_mensal": pd.DataFrame([
                {"Ciclo": "C1", "Competência": "01/2025", "Valor pago/faturado": 1500.0, "Efeito financeiro": "Nao"},
                {"Ciclo": "C1", "Competência": "02/2025", "Valor pago/faturado": 2000.0, "Efeito financeiro": "Sim"},
                {"Ciclo": "C2", "Competência": "01/2026", "Valor pago/faturado": 3000.0, "Efeito financeiro": "Sim"},
            ]),
            "totais_canonicos_pc": {
                "ate_o_corte": {"retroativo": 12345.67},
                "por_ciclo": {
                    "C1": {"retroativo": 5000.0, "delta_potencial": 890.12, "quantidade": 2},
                    "C2": {"retroativo": 7345.67, "delta_potencial": 0.0, "quantidade": 0},
                }
            },
        }

        at = AppTest.from_file(self.CAMINHO_PAGINA, default_timeout=90)
        at.session_state["assinatura_processada_upload_docs"] = "sig"
        at.session_state["assinatura_upload_docs"] = "sig"
        at.session_state["resultado_valor_global"] = resultado
        at.session_state["diagnostico_coleta_v2"] = diagnostico
        at.run()
        self.assertFalse(at.exception, msg=str(getattr(at, "exception", "")))

        # Os 4 cards de resumo continuam presentes e com os valores canônicos.
        metricas = {m.label: m.value for m in at.metric}
        self.assertEqual(metricas.get("Índice"), "IST")
        self.assertEqual(metricas.get("Ciclos analisados"), "C1, C2")
        self.assertEqual(metricas.get("Retroativo reconhecido"), "R$ 12.345,67")
        self.assertEqual(metricas.get("Percentual acumulado"), "7,50%")

        markdowns = " \n ".join(m.value for m in at.markdown)
        self.assertIn("Validação com a contratada", markdowns)
        # Ordem: os 6 cards de documentos vêm antes do novo bloco.
        self.assertLess(markdowns.index("Termo de Apostila"), markdowns.index("Validação com a contratada"))

        self.assertIn("Baixar TXT", [d.label for d in at.download_button])
        self.assertIn("Visualizar comunicado", [e.label for e in at.expander])

        texto = next(c.value for c in at.code)
        self.assertIn("[a preencher]", texto)  # contrato sem fonte automática hoje
        self.assertNotIn("firmado com a", texto)  # item 1: trecho excluído
        self.assertIn("TOTAL RETROATIVO RECONHECIDO: R$ 12.345,67", texto)
        self.assertIn("TOTAL RETROATIVO POTENCIAL: R$ 890,12", texto)
        # PERÍODO/ANO tambem em mm/aaaa (hotfix pos-PR#87, item 4).
        self.assertIn("C1 | 01/2025 a 12/2025 | R$ 5.000,00", texto)
        # EFEITO FINANCEIRO: fonte canonica parametros_v10.por_ciclo, nao df_financeiro_mensal.
        # PERIODO e EFEITO FINANCEIRO em mm/aaaa (hotfix pos-PR#87, item 2).
        self.assertIn("C1 | 01/2025 a 12/2025 | IST | 4,50% | A partir de 02/2025", texto)
        # COMPETÊNCIAS SEM EFEITO: mesma regra homologada do Sumário Executivo (Requisito 5).
        # item 3: sem a coluna VALOR CORRESPONDENTE.
        self.assertIn("CICLO | COMPETÊNCIAS SEM EFEITO", texto)
        self.assertNotIn("VALOR CORRESPONDENTE", texto)
        self.assertIn("C1 | 01/2025", texto)
        self.assertNotIn("C1 | 01/2025 | —", texto)
        self.assertNotIn("Não houve perda de competências neste ciclo.", texto)
        self.assertNotIn("Para facilitar o registro, a concordância poderá ser manifestada", texto)
        self.assertNotIn("De acordo com os ciclos, índices, efeitos financeiros e valores apresentados", texto)
        # item 7: parágrafo final removido.
        self.assertNotIn("favor indicar o ciclo, competência, Pedido de Compra", texto)

    def test_metodo_financeiro_sem_potencial_e_sem_perda(self):
        import pandas as pd
        from streamlit.testing.v1 import AppTest

        consolidado = {
            "vta": 300000.0,
            "vta_origem": "posicao_atual",
            "vta_usa_ultima_posicao": False,
            "retroativo_reconhecido": 4321.0,
            "valor_atualizado_em_analise": None,
            "retroativo_potencial": None,
            "medidas_pc_aplicaveis": False,
            "fora_do_corte": {"aplicavel": False, "quantidade": None, "valor_informado": None, "data_corte": None},
            "metodo": {"codigo": "financeiro", "rotulo": "Financeiro"},
            "ciclo_vigente": "C1",
            "status_confiabilidade": "CONFIÁVEL",
            "mensagem_status": "Resultado apurado sem ressalvas materiais identificadas.",
            "formalizacao": {"bloqueada": False, "status": "SEM BLOQUEIO", "mensagem": "Sem bloqueio explícito à formalização."},
            "bloqueios": [],
            "ressalvas": [],
            "campos_nao_confiaveis": [],
            "composicao_vta": {"disponivel": False, "conciliada": None, "exibivel": False, "linhas": [], "mensagem": "", "ha_aditivos_nao_computados": False},
        }
        diagnostico = {
            "status_base": "OK",
            "pronto_para_consolidar": True,
            "metadados": {"indice": "IPCA", "ciclos_em_analise": ["C1"]},
        }
        resultado = {
            "capacidades": {"documentos": _DOCS_SEIS_CARDS},
            "diagnostico_coleta": diagnostico,
            "variacao_acumulada": 0.02,
            "indice": "IPCA",
            "resultado_consolidado": consolidado,
            "objeto_processo": {},
            "df_ciclos": pd.DataFrame([{"Ciclo": "C1", "Variação": 0.02}]),
            "df_financeiro_mensal": pd.DataFrame([
                {"Ciclo": "C1", "Competência": "01/2025", "Valor pago/faturado": 1000.0, "Efeito financeiro": "Sim"},
            ]),
            "df_financeiro_por_ciclo": pd.DataFrame([
                {"Ciclo": "C1", "Delta do ciclo": 4321.0},
            ]),
        }

        at = AppTest.from_file(self.CAMINHO_PAGINA, default_timeout=90)
        at.session_state["assinatura_processada_upload_docs"] = "sig"
        at.session_state["assinatura_upload_docs"] = "sig"
        at.session_state["resultado_valor_global"] = resultado
        at.session_state["diagnostico_coleta_v2"] = diagnostico
        at.run()
        self.assertFalse(at.exception, msg=str(getattr(at, "exception", "")))

        texto = next(c.value for c in at.code)
        self.assertIn("TOTAL RETROATIVO RECONHECIDO: R$ 4.321,00", texto)
        self.assertNotIn("TOTAL RETROATIVO POTENCIAL", texto)
        self.assertIn("Não houve perda de competências neste ciclo.", texto)
        self.assertNotIn("firmado com a", texto)  # item 1
        self.assertNotIn("VALOR CORRESPONDENTE", texto)  # item 3
        self.assertNotIn("favor indicar o ciclo, competência, Pedido de Compra", texto)  # item 7


if __name__ == "__main__":
    unittest.main()
