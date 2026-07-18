import ast
from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]
PAGINA = (ROOT / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")


def _documentos_funcionais() -> tuple[tuple[str, str], ...]:
    arvore = ast.parse(PAGINA)
    for no in arvore.body:
        if isinstance(no, ast.Assign):
            if any(isinstance(alvo, ast.Name) and alvo.id == "DOCUMENTOS_FUNCIONAIS_UPLOAD" for alvo in no.targets):
                return ast.literal_eval(no.value)
    raise AssertionError("DOCUMENTOS_FUNCIONAIS_UPLOAD não foi encontrado")


class TestFluxoUploadDocs(unittest.TestCase):
    def test_antes_do_upload_exibe_somente_download_e_upload(self):
        self.assertIn('"Baixar Arquivo Coleta Oficial"', PAGINA)
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
        self.assertIn('"assinatura_upload_docs"', trecho)
        self.assertNotIn("processar_coleta_oficial_runtime", trecho)
        self.assertNotIn("render_documentos_funcionais_upload", trecho)

        bloco_processar = PAGINA[botao:PAGINA.index('if st.session_state.get("assinatura_processada_upload_docs")', botao)]
        self.assertIn("processar_coleta_oficial_runtime(conteudo_upload)", bloco_processar)
        self.assertIn('st.session_state["assinatura_processada_upload_docs"] = assinatura_upload', bloco_processar)

    def test_paineis_redundantes_nao_podem_ser_reintroduzidos(self):
        self.assertNotIn("render_status_apuracao", PAGINA)
        self.assertNotIn("render_status_documentos", PAGINA)
        self.assertNotIn("Status da Apuração", PAGINA)
        self.assertNotIn("Documentos da Apuração", PAGINA)

    def test_apos_processar_exibe_exatamente_oito_cards_funcionais(self):
        documentos = _documentos_funcionais()
        self.assertEqual(len(documentos), 8)
        self.assertEqual(
            tuple(nome for _, nome in documentos),
            (
                "Planilha Executiva",
                "Itens por Ciclo",
                "Relatório Executivo",
                "Memória de Cálculo e Marcos",
                "Termo de Apostila",
                "Garantia Contratual",
                "DOU",
                "Checklist Processual",
            ),
        )
        guarda = PAGINA.index('if st.session_state.get("assinatura_processada_upload_docs")')
        render = PAGINA.index("render_documentos_funcionais_upload(resultado)", guarda)
        parada = PAGINA.index("st.stop()", render)
        self.assertLess(guarda, render)
        self.assertLess(render, parada)

    def test_apos_processar_restaura_exatamente_quatro_resumos_antes_dos_cards(self):
        inicio = PAGINA.index("if resultado:", PAGINA.index("diagnostico_coleta ="))
        render = PAGINA.index("render_documentos_funcionais_upload(resultado)", inicio)
        trecho = PAGINA[inicio:render]
        self.assertEqual(trecho.count(".metric("), 4)
        rotulos = (
            'metric("Índice"',
            'metric("Ciclo vigente"',
            'metric("Meses com valor"',
            'metric("Itens remanescentes"',
        )
        posicoes = [trecho.index(rotulo) for rotulo in rotulos]
        self.assertEqual(posicoes, sorted(posicoes))
        self.assertIn('contagens.get("competencias_com_valor", 0)', trecho)

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
            "Abrir Relatório Executivo",
            "Abrir Termo de Apostila",
            "Abrir Garantia Contratual",
            "Abrir DOU",
            "Abrir Checklist Processual",
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
            'key="upload_docs_planilha_executiva"',
            'key="upload_docs_itens_ciclo"',
            'key="upload_docs_memoria_marcos"',
        )
        for chave in chaves:
            self.assertEqual(trecho.count(chave), 1)

    def test_novo_arquivo_coleta_e_backend_oficial_permanecem_no_fluxo(self):
        self.assertIn("TEMPLATE_COLETA_OFICIAL", PAGINA)
        self.assertIn("assinatura_template_coleta", PAGINA)
        self.assertIn("processar_coleta_oficial_runtime", PAGINA)
        self.assertIn("resultado_valor_global", PAGINA)
        self.assertIn("diagnostico_coleta_v2", PAGINA)


if __name__ == "__main__":
    unittest.main()
