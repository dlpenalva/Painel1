"""Regressão: preservar a apuração ao voltar da Adequação para o Valor Global.

Cobre a demanda original de que a navegação
    Valor Global -> Adequação Orçamentária -> "Voltar para Valor Global"
não pode exigir novo upload nem reprocessamento: a fonte de verdade após o
processamento é o ``st.session_state``, não o widget ``file_uploader``.
"""
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from _estado_apuracao_upload import (
    CHAVE_ASSINATURA_PROCESSADA,
    CHAVE_RESULTADO,
    apuracao_persistida_valida,
)

PAGINA = (ROOT / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")
ADEQUACAO = (ROOT / "pages" / "12_Adequacao_Orcamentaria.py").read_text(encoding="utf-8")


class TestPredicadoReidratacao(unittest.TestCase):
    """Comportamento puro da decisão de reutilizar a apuração persistida."""

    def test_b_e_estado_valido_permite_reidratar_sem_novo_upload(self):
        estado = {
            CHAVE_ASSINATURA_PROCESSADA: "abc123",
            CHAVE_RESULTADO: {"variacao_acumulada": 0.0402, "capacidades": {}},
        }
        self.assertTrue(apuracao_persistida_valida(estado))

    def test_g_resultado_ausente_nao_reidrata(self):
        self.assertFalse(apuracao_persistida_valida({CHAVE_ASSINATURA_PROCESSADA: "abc123"}))
        self.assertFalse(
            apuracao_persistida_valida(
                {CHAVE_ASSINATURA_PROCESSADA: "abc123", CHAVE_RESULTADO: None}
            )
        )

    def test_f_apos_invalidacao_sem_assinatura_processada_nao_reidrata(self):
        # Após carregar arquivo diferente/incompatível, a assinatura processada é
        # removida (mesmo que um resultado antigo ainda esteja em memória): não pode
        # haver reutilização indevida entre arquivos distintos.
        estado_pos_invalidacao = {CHAVE_RESULTADO: {"variacao_acumulada": 0.0402}}
        self.assertFalse(apuracao_persistida_valida(estado_pos_invalidacao))

    def test_estado_vazio_nao_reidrata(self):
        self.assertFalse(apuracao_persistida_valida({}))


class TestFluxoReidratacaoNaPagina(unittest.TestCase):
    """Assertivas estruturais sobre o guard do file_uploader na página."""

    def test_a_processamento_persiste_estado_canonico(self):
        botao = PAGINA.index('if st.button("Processar"')
        guarda = PAGINA.index(
            'if st.session_state.get("assinatura_processada_upload_docs")', botao
        )
        bloco = PAGINA[botao:guarda]
        self.assertIn("processar_coleta_oficial_runtime(conteudo_upload)", bloco)
        self.assertIn('st.session_state["diagnostico_coleta_v2"] = diagnostico_processado', bloco)
        self.assertIn('st.session_state["resultado_valor_global"] = resultado_processado', bloco)
        self.assertIn(
            'st.session_state["assinatura_processada_upload_docs"] = assinatura_upload', bloco
        )

    def test_b_e_guard_reidrata_em_vez_de_parar_quando_ha_apuracao(self):
        inicio = PAGINA.index("if arquivo is None:")
        # O ramo "sem arquivo" vai até o início do ramo com arquivo (else:/conteudo_upload).
        fim = PAGINA.index("conteudo_upload = arquivo.getvalue()", inicio)
        ramo_sem_arquivo = PAGINA[inicio:fim]
        # Sem arquivo, consulta o session_state antes de parar.
        self.assertIn("apuracao_persistida_valida(st.session_state)", ramo_sem_arquivo)
        self.assertIn(
            'assinatura_upload = st.session_state["assinatura_processada_upload_docs"]',
            ramo_sem_arquivo,
        )
        # st.stop() só ocorre no else interno da reidratação (sem apuração válida).
        self.assertIn("st.stop()", ramo_sem_arquivo)

    def test_e_reidratacao_nao_reprocessa_nem_exige_arquivo(self):
        inicio = PAGINA.index("if arquivo is None:")
        fim = PAGINA.index("conteudo_upload = arquivo.getvalue()", inicio)
        ramo_sem_arquivo = PAGINA[inicio:fim]
        self.assertNotIn("processar_coleta_oficial_runtime", ramo_sem_arquivo)
        self.assertNotIn("arquivo.getvalue()", ramo_sem_arquivo)

    def test_c_d_render_ocorre_apos_guard_unico(self):
        # As quatro métricas e os seis documentos são renderizados a partir do
        # resultado do session_state, após o guard de assinatura processada.
        guarda = PAGINA.index('if st.session_state.get("assinatura_processada_upload_docs")')
        resultado_lido = PAGINA.index('resultado = st.session_state.get("resultado_valor_global")', guarda)
        render = PAGINA.index("render_documentos_funcionais_upload(resultado)", resultado_lido)
        trecho_render = PAGINA[resultado_lido:render]
        self.assertEqual(trecho_render.count(".metric("), 4)
        self.assertLess(guarda, resultado_lido)
        self.assertLess(resultado_lido, render)

    def test_f_invalidacao_por_arquivo_diferente_preservada(self):
        # O ramo com arquivo mantém a limpeza dos estados derivados quando a
        # assinatura do novo arquivo difere da última vista pelo uploader.
        else_idx = PAGINA.index("else:", PAGINA.index("if arquivo is None:"))
        bloco_com_arquivo = PAGINA[else_idx:PAGINA.index('if st.button("Processar"', else_idx)]
        self.assertIn('if st.session_state.get("assinatura_upload_docs") != assinatura_upload:', bloco_com_arquivo)
        self.assertIn('st.session_state.pop("resultado_valor_global", None)', bloco_com_arquivo)
        self.assertIn('st.session_state.pop("diagnostico_coleta_v2", None)', bloco_com_arquivo)
        self.assertIn('st.session_state.pop("assinatura_processada_upload_docs", None)', bloco_com_arquivo)

    def test_adequacao_nao_limpa_apuracao(self):
        # A página de Adequação apenas lê o resultado e navega de volta; jamais
        # remove ou limpa as chaves da apuração.
        self.assertIn('st.switch_page("pages/03_Valor_Global.py")', ADEQUACAO)
        self.assertNotIn('pop("resultado_valor_global"', ADEQUACAO)
        self.assertNotIn('pop("diagnostico_coleta_v2"', ADEQUACAO)
        self.assertNotIn('pop("assinatura_processada_upload_docs"', ADEQUACAO)
        self.assertNotIn("st.session_state.clear()", ADEQUACAO)


class TestReidratacaoComportamentalAppTest(unittest.TestCase):
    """Executa a página no runtime do Streamlit (AppTest) para reproduzir o retorno
    da navegação sem novo upload — prova end-to-end de que a apuração é reidratada."""

    CAMINHO_PAGINA = str(ROOT / "pages" / "03_Valor_Global.py")

    @staticmethod
    def _semear_apuracao(at):
        at.session_state["assinatura_processada_upload_docs"] = "sig-teste"
        at.session_state["assinatura_upload_docs"] = "sig-teste"
        at.session_state["resultado_valor_global"] = {
            "capacidades": {"documentos": {}},
            "variacao_acumulada": 0.0402,
        }
        at.session_state["diagnostico_coleta_v2"] = {
            "metadados": {"indice": "IST", "ciclos_em_analise": ["C1"]},
            "pronto_para_consolidar": True,
        }

    def test_b_c_retorno_sem_arquivo_reidrata_quatro_metricas(self):
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file(self.CAMINHO_PAGINA, default_timeout=90)
        self._semear_apuracao(at)
        at.run()  # nenhum arquivo no uploader: simula o retorno da Adequação
        self.assertFalse(at.exception, msg=str(getattr(at, "exception", "")))
        rotulos = [m.label for m in at.metric]
        self.assertEqual(len(rotulos), 4)
        self.assertIn("Índice", rotulos)
        self.assertIn("Ciclos analisados", rotulos)
        self.assertIn("Retroativo reconhecido", rotulos)
        self.assertIn("Percentual acumulado", rotulos)

    def test_d_retorno_sem_arquivo_renderiza_os_seis_documentos(self):
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file(self.CAMINHO_PAGINA, default_timeout=90)
        self._semear_apuracao(at)
        at.run()
        self.assertFalse(at.exception, msg=str(getattr(at, "exception", "")))
        markdowns = " \n ".join(m.value for m in at.markdown)
        for titulo in (
            "Sumário Executivo",
            "Adequação Orçamentária",
            "Despacho Saneador",
            "Termo de Apostila",
            "Garantia Contratual",
            "DOU",
        ):
            self.assertIn(titulo, markdowns)

    def test_estado_vazio_sem_arquivo_nao_reidrata_nem_renderiza_metricas(self):
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file(self.CAMINHO_PAGINA, default_timeout=90)
        at.run()  # sessão limpa, sem arquivo: deve parar no bloco de upload
        self.assertFalse(at.exception, msg=str(getattr(at, "exception", "")))
        self.assertEqual(len(at.metric), 0)


if __name__ == "__main__":
    unittest.main()
