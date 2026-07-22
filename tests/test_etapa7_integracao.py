"""Testes de integração da Etapa 7 — seis documentos canônicos na interface."""
import sys
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

UPLOAD = (ROOT / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")
CENTRAL = (ROOT / "pages" / "06_Central_Arquivos.py").read_text(encoding="utf-8")

from _capacidades_apuracao import SEIS_DOCUMENTOS_CANONICOS


CHAVES_CANONICAS = tuple(c for c, _ in SEIS_DOCUMENTOS_CANONICOS)
NOMES_CANONICOS = tuple(n for _, n in SEIS_DOCUMENTOS_CANONICOS)

DOCS_SEMPRE_ACESSIVEIS = ("adequacao_orcamentaria", "garantia_contratual", "dou")
DOCS_DEPENDENTES_XLS = ("sumario_executivo", "despacho_saneador", "termo_apostila")

DOCS_ANTIGOS_REMOVIDOS = (
    "Planilha Executiva",
    "Itens por Ciclo",
    "Relatório Executivo",
    "Memória de Cálculo e Marcos",
    "Checklist Processual",
)


class TestRegistroCanonico(unittest.TestCase):
    def test_exatamente_seis_documentos(self):
        self.assertEqual(len(SEIS_DOCUMENTOS_CANONICOS), 6)

    def test_ordem_e_nomes_exatos(self):
        self.assertEqual(
            NOMES_CANONICOS,
            (
                "Sumário Executivo",
                "Adequação Orçamentária",
                "Despacho Saneador",
                "Termo de Apostila",
                "Garantia Contratual",
                "DOU",
            ),
        )

    def test_chaves_unicas(self):
        self.assertEqual(len(CHAVES_CANONICAS), len(set(CHAVES_CANONICAS)))

    def test_tres_sempre_acessiveis_presentes(self):
        for chave in DOCS_SEMPRE_ACESSIVEIS:
            self.assertIn(chave, CHAVES_CANONICAS)

    def test_tres_dependentes_xls_presentes(self):
        for chave in DOCS_DEPENDENTES_XLS:
            self.assertIn(chave, CHAVES_CANONICAS)


class TestUploadConvergente(unittest.TestCase):
    def test_referencia_registro_canonico(self):
        self.assertIn("DOCUMENTOS_FUNCIONAIS_UPLOAD = SEIS_DOCUMENTOS_CANONICOS", UPLOAD)

    def test_importa_registro_canonico(self):
        self.assertIn("from _capacidades_apuracao import SEIS_DOCUMENTOS_CANONICOS", UPLOAD)

    def test_importa_geradores_novos(self):
        self.assertIn("from _sumario_executivo import gerar_sumario_executivo", UPLOAD)
        self.assertIn("from _templates_documentos import gerar_despacho_saneador, gerar_termo_apostila", UPLOAD)

    def test_ausencia_docs_antigos_na_interface(self):
        for nome in DOCS_ANTIGOS_REMOVIDOS:
            self.assertNotIn(f'"{nome}"', UPLOAD)

    def test_grade_tres_colunas(self):
        render_ini = UPLOAD.index("def render_documentos_funcionais_upload")
        render_fim = UPLOAD.index("# Interface", render_ini)
        render = UPLOAD[render_ini:render_fim]
        self.assertIn("st.columns(3)", render)
        self.assertNotIn("st.columns(4)", render)

    def test_sumario_usa_download_button_pdf(self):
        acao_ini = UPLOAD.index("def _render_acao_documento_upload")
        acao_fim = UPLOAD.index("def render_documentos_funcionais_upload", acao_ini)
        acao = UPLOAD[acao_ini:acao_fim]
        self.assertIn('key="upload_docs_sumario_executivo"', acao)
        self.assertIn("gerar_sumario_executivo(resultado)", acao)
        self.assertIn('"Baixar PDF"', acao)

    def test_despacho_usa_download_button_docx(self):
        acao_ini = UPLOAD.index("def _render_acao_documento_upload")
        acao_fim = UPLOAD.index("def render_documentos_funcionais_upload", acao_ini)
        acao = UPLOAD[acao_ini:acao_fim]
        self.assertIn('key="upload_docs_despacho_saneador"', acao)
        self.assertIn("gerar_despacho_saneador(resultado)", acao)
        self.assertIn('"Baixar DOCX"', acao)
        self.assertIn("Despacho_Saneador_Instrucao_Processual.docx", acao)

    def test_apostila_usa_download_button_docx(self):
        acao_ini = UPLOAD.index("def _render_acao_documento_upload")
        acao_fim = UPLOAD.index("def render_documentos_funcionais_upload", acao_ini)
        acao = UPLOAD[acao_ini:acao_fim]
        self.assertIn('key="upload_docs_termo_apostila"', acao)
        self.assertIn("gerar_termo_apostila(resultado)", acao)
        self.assertIn("Termo_de_Apostila_Reajuste_Contratual.docx", acao)

    def test_sempre_acessiveis_usam_page_link(self):
        acao_ini = UPLOAD.index("def _render_acao_documento_upload")
        acao_fim = UPLOAD.index("def render_documentos_funcionais_upload", acao_ini)
        acao = UPLOAD[acao_ini:acao_fim]
        self.assertIn("Abrir Adequação Orçamentária", acao)
        self.assertIn("Abrir Garantia Contratual", acao)
        self.assertIn("Abrir DOU", acao)
        self.assertIn("pages/12_Adequacao_Orcamentaria.py", acao)
        self.assertIn("pages/05_Garantia.py", acao)
        self.assertIn("pages/13_DOU.py", acao)


class TestCentralConvergente(unittest.TestCase):
    def test_exatamente_seis_entradas_no_catalogo(self):
        catalogo = CENTRAL[: CENTRAL.index("def aplicar_css_central")]
        self.assertEqual(catalogo.count('"nome":'), 6)

    def test_nomes_canonicos_presentes(self):
        for nome in NOMES_CANONICOS:
            self.assertIn(nome, CENTRAL)

    def test_grade_tres_colunas(self):
        self.assertIn("st.columns(3)", CENTRAL)
        self.assertNotIn("st.columns(4)", CENTRAL)

    def test_sempre_acessivel_field_presente(self):
        self.assertIn('"sempre_acessivel"', CENTRAL)
        self.assertIn('"sempre_acessivel": True', CENTRAL)
        self.assertIn('"sempre_acessivel": False', CENTRAL)

    def test_bypass_bloqueio_para_sempre_acessiveis(self):
        self.assertIn("not sempre_acessivel", CENTRAL)

    def test_intro_menciona_seis_documentos(self):
        self.assertIn("seis documentos oficiais", CENTRAL)

    def test_session_keys_canonicos(self):
        for key in (
            "arquivo_sumario_executivo_pdf",
            "arquivo_adequacao_orcamentaria_xlsx",
            "arquivo_despacho_saneador_docx",
            "arquivo_termo_apostila_docx",
            "arquivo_garantia_pdf",
            "arquivo_dou_docx",
        ):
            self.assertIn(key, CENTRAL)

    def test_mime_pdf_e_docx_presentes(self):
        self.assertIn('"application/pdf"', CENTRAL)
        self.assertIn('"application/vnd.openxmlformats-officedocument.wordprocessingml.document"', CENTRAL)


class TestUploadECentralConvergentes(unittest.TestCase):
    def test_nomes_canonicos_presentes_em_central(self):
        # Central tem os nomes como strings literais nos dicts
        for nome in NOMES_CANONICOS:
            self.assertIn(nome, CENTRAL, msg=f'"{nome}" ausente em 06_Central_Arquivos.py')

    def test_upload_tem_marcadores_por_documento(self):
        # Upload itera nomes dinamicamente; verifica marcadores concretos por chave
        marcadores = (
            "upload_docs_sumario_executivo",
            "Abrir Adequação Orçamentária",
            "upload_docs_despacho_saneador",
            "upload_docs_termo_apostila",
            "Abrir Garantia Contratual",
            "Abrir DOU",
        )
        for marcador in marcadores:
            self.assertIn(marcador, UPLOAD, msg=f'"{marcador}" ausente em 03_Valor_Global.py')

    def test_docs_antigos_ausentes_em_ambas_superficies(self):
        for nome in DOCS_ANTIGOS_REMOVIDOS:
            self.assertNotIn(f'"{nome}"', UPLOAD, msg=f'"{nome}" ainda presente em 03_Valor_Global.py')
            self.assertNotIn(f'"{nome}"', CENTRAL, msg=f'"{nome}" ainda presente em 06_Central_Arquivos.py')


if __name__ == "__main__":
    unittest.main()
