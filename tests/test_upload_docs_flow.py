import sys
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

PAGINA = (ROOT / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")

from _capacidades_apuracao import SEIS_DOCUMENTOS_CANONICOS


class TestFluxoUploadDocs(unittest.TestCase):
    def test_antes_do_upload_exibe_somente_o_uploader(self):
        self.assertNotIn('"Baixar Arquivo Coleta Oficial"', PAGINA)
        self.assertNotIn("Baixar arquivo de trabalho", PAGINA)
        self.assertNotIn(
            "Use o modelo oficial com fórmulas para registrar a apuração contratual.",
            PAGINA,
        )
        self.assertIn("Enviar Arquivo Coleta Oficial preenchido", PAGINA)
        self.assertIn('key="upload_coleta_documentos"', PAGINA)
        trecho = PAGINA[PAGINA.index("if arquivo is None:"):PAGINA.index("conteudo_upload =")]
        self.assertIn("st.stop()", trecho)
        self.assertNotIn("processar_coleta_oficial_runtime", trecho)
        self.assertNotIn("render_documentos_funcionais_upload", trecho)

    def test_apos_upload_exige_processamento_explicito(self):
        inicio = PAGINA.index("conteudo_upload =")
        botao = PAGINA.index('if st.button("Processar"', inicio)
        trecho = PAGINA[inicio:botao]
        self.assertIn('st.caption(f"Arquivo enviado: {arquivo.name}")', trecho)
        self.assertIn("assinatura_conteudo_upload(conteudo_upload)", trecho)
        self.assertIn(
            "invalidar_estado_caso(st.session_state, assinatura_upload)", trecho
        )
        self.assertNotIn("processar_coleta_oficial_runtime", trecho)
        self.assertNotIn("render_documentos_funcionais_upload", trecho)

        bloco_processar = PAGINA[botao:PAGINA.index('if st.session_state.get("assinatura_processada_upload_docs")', botao)]
        self.assertIn("processar_coleta_oficial_runtime(conteudo_upload)", bloco_processar)
        self.assertIn('st.session_state["assinatura_processada_upload_docs"] = assinatura_upload', bloco_processar)

    def test_troca_invalida_no_callback_antes_do_rerun(self):
        callback = PAGINA[PAGINA.index("def _invalidar_caso_antes_do_rerun_upload"):]
        uploader = callback[:callback.index("if arquivo is None:")]
        self.assertIn("getattr(novo_arquivo, \"size\", 0) > TAMANHO_MAXIMO_ARQUIVO", callback)
        self.assertIn("assinatura_conteudo_upload(novo_arquivo.getvalue())", callback)
        self.assertIn(
            "invalidar_estado_caso(st.session_state, nova_assinatura)", callback
        )
        self.assertIn(
            "on_change=_invalidar_caso_antes_do_rerun_upload", uploader
        )

    def test_paineis_redundantes_nao_podem_ser_reintroduzidos(self):
        self.assertNotIn("render_status_apuracao", PAGINA)
        self.assertNotIn("render_status_documentos", PAGINA)
        self.assertNotIn("Status da Apuração", PAGINA)
        self.assertNotIn("Documentos da Apuração", PAGINA)

    def test_apos_processar_exibe_exatamente_seis_cards_funcionais(self):
        self.assertEqual(len(SEIS_DOCUMENTOS_CANONICOS), 6)
        self.assertEqual(
            tuple(nome for _, nome in SEIS_DOCUMENTOS_CANONICOS),
            (
                "Sumário Executivo",
                "Adequação Orçamentária",
                "Despacho Saneador",
                "Termo de Apostila",
                "Garantia Contratual",
                "DOU",
            ),
        )
        # DOCUMENTOS_FUNCIONAIS_UPLOAD deve referenciar o registro canônico
        self.assertIn("DOCUMENTOS_FUNCIONAIS_UPLOAD = SEIS_DOCUMENTOS_CANONICOS", PAGINA)
        guarda = PAGINA.index('if st.session_state.get("assinatura_processada_upload_docs")')
        render = PAGINA.index("render_documentos_funcionais_upload(resultado)", guarda)
        parada = PAGINA.index("st.stop()", render)
        self.assertLess(guarda, render)
        self.assertLess(render, parada)

    def test_ausencia_dos_cinco_acessos_antigos_na_interface(self):
        for titulo in (
            '"Planilha Executiva"',
            '"Itens por Ciclo"',
            '"Relatório Executivo"',
            '"Memória de Cálculo e Marcos"',
            '"Checklist Processual"',
        ):
            self.assertNotIn(titulo, PAGINA)

    def test_apos_processar_exibe_cards_antes_do_resultado_consolidado(self):
        inicio = PAGINA.index("if resultado:", PAGINA.index("diagnostico_coleta ="))
        render = PAGINA.index("render_documentos_funcionais_upload(resultado)", inicio)
        trecho = PAGINA[inicio:render]
        self.assertEqual(trecho.count(".metric("), 4)
        pos_cards = [
            trecho.index('resumo_indice.metric("Índice"'),
            trecho.index('resumo_ciclos.metric("Ciclos analisados"'),
            trecho.index('resumo_retro.metric("Retroativo reconhecido"'),
            trecho.index('resumo_acum.metric("Percentual acumulado"'),
        ]
        pos_resultado = trecho.index(
            "render_resultado_consolidado(resultado, diagnostico_coleta)"
        )
        self.assertEqual(pos_cards, sorted(pos_cards))
        self.assertLess(pos_cards[-1], pos_resultado)
        self.assertIn('st.caption(f"ID da apuração:', trecho)

    def test_pendencias_multiplas_usam_chaves_semanticas_unicas(self):
        inicio = PAGINA.index("def _render_pendencia_documento")
        fim = PAGINA.index("def _render_acao_documento_upload", inicio)
        helper = PAGINA[inicio:fim]
        self.assertIn('key=f"upload_docs_{chave}_pendencia"', helper)

        render_inicio = PAGINA.index("def render_documentos_funcionais_upload")
        render_fim = PAGINA.index("# Interface", render_inicio)
        render = PAGINA[render_inicio:render_fim]
        self.assertIn("_render_acao_documento_upload(chave, documento, resultado)", render)
        self.assertIn("except Exception as exc:", render)

        labels_links = (
            "Abrir Adequação Orçamentária",
            "Abrir Garantia Contratual",
            "Abrir DOU",
        )
        self.assertEqual(len(labels_links), len(set(labels_links)))
        for label in labels_links:
            self.assertIn(label, PAGINA)

    def test_widgets_dinamicos_com_key_explicita_nao_repetem_chave(self):
        inicio = PAGINA.index("def _render_pendencia_documento")
        fim = PAGINA.index("def render_documentos_funcionais_upload", inicio)
        trecho = PAGINA[inicio:fim]
        chaves = (
            'key=f"upload_docs_{chave}_pendencia"',
            'key="upload_docs_sumario_executivo"',
            'key="upload_docs_despacho_saneador"',
            'key="upload_docs_termo_apostila"',
        )
        for chave in chaves:
            self.assertEqual(trecho.count(chave), 1)

    def test_grade_tres_colunas_na_renderizacao(self):
        render_inicio = PAGINA.index("def render_documentos_funcionais_upload")
        render_fim = PAGINA.index("# Interface", render_inicio)
        render = PAGINA[render_inicio:render_fim]
        self.assertIn("st.columns(3)", render)
        self.assertNotIn("st.columns(4)", render)

    def test_novos_geradores_importados(self):
        self.assertIn("from _sumario_executivo import gerar_sumario_executivo", PAGINA)
        self.assertIn("from _templates_documentos import gerar_despacho_saneador, gerar_termo_apostila", PAGINA)
        self.assertIn("from _capacidades_apuracao import SEIS_DOCUMENTOS_CANONICOS", PAGINA)

    def test_backend_oficial_permanece_no_fluxo_sem_gerador_de_download(self):
        self.assertNotIn("TEMPLATE_COLETA_OFICIAL", PAGINA)
        self.assertNotIn("assinatura_template_coleta", PAGINA)
        self.assertNotIn("_coleta_oficial_cacheada", PAGINA)
        self.assertIn("processar_coleta_oficial_runtime", PAGINA)
        self.assertIn("resultado_valor_global", PAGINA)
        self.assertIn("diagnostico_coleta_v2", PAGINA)


if __name__ == "__main__":
    unittest.main()
