"""UX-CARDS — agrupamento visual dos seis cards pos-upload.

Protege apenas a ESTRUTURA (presenca dos seis cards, grupamento em duas linhas
de tres e ordem dentro de cada linha). Nao testa cor nem qualquer decisao
decorativa: a paleta pode ser reajustada sem quebrar este arquivo.
"""
import ast
import sys
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

PAGINA = (ROOT / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")

from _capacidades_apuracao import SEIS_DOCUMENTOS_CANONICOS


def _constante(nome):
    """Le o literal de uma constante de modulo sem executar a pagina."""
    arvore = ast.parse(PAGINA)
    for no in arvore.body:
        alvos = no.targets if isinstance(no, ast.Assign) else []
        for alvo in alvos:
            if isinstance(alvo, ast.Name) and alvo.id == nome:
                return ast.literal_eval(no.value)
    raise AssertionError(f"constante {nome} nao encontrada em 03_Valor_Global.py")


GRUPOS = _constante("GRUPOS_CARDS_UPLOAD")
CHAVES_CANONICAS = {chave for chave, _ in SEIS_DOCUMENTOS_CANONICOS}


class TestAgrupamentoCardsUpload(unittest.TestCase):
    def test_duas_linhas_de_tres(self):
        self.assertEqual(len(GRUPOS), 2)
        for _grupo, chaves in GRUPOS:
            self.assertEqual(len(chaves), 3)

    def test_cobre_exatamente_os_seis_canonicos_sem_repetir(self):
        renderizadas = [chave for _grupo, chaves in GRUPOS for chave in chaves]
        self.assertEqual(len(renderizadas), 6)
        self.assertEqual(len(set(renderizadas)), 6)
        self.assertEqual(set(renderizadas), CHAVES_CANONICAS)

    def test_linha_1_documentos_gerados_na_ordem(self):
        grupo, chaves = GRUPOS[0]
        self.assertEqual(grupo, "documento")
        self.assertEqual(
            chaves, ("despacho_saneador", "termo_apostila", "sumario_executivo")
        )

    def test_linha_2_acoes_na_ordem(self):
        grupo, chaves = GRUPOS[1]
        self.assertEqual(grupo, "acao")
        self.assertEqual(
            chaves, ("adequacao_orcamentaria", "garantia_contratual", "dou")
        )

    def test_titulos_continuam_vindo_da_tupla_canonica(self):
        # A constante de agrupamento carrega so posicao: nenhum titulo literal.
        self.assertIn("titulos = dict(DOCUMENTOS_FUNCIONAIS_UPLOAD)", PAGINA)
        self.assertIn('st.markdown(f"#### {titulo}")', PAGINA)

    def test_render_itera_os_grupos_em_colunas_de_tres(self):
        inicio = PAGINA.index("def render_documentos_funcionais_upload")
        fim = PAGINA.index("# Interface", inicio)
        render = PAGINA[inicio:fim]
        self.assertIn("for grupo, chaves_do_grupo in GRUPOS_CARDS_UPLOAD:", render)
        self.assertIn("st.columns(3)", render)
        self.assertNotIn("st.columns(4)", render)
        self.assertIn("_render_acao_documento_upload(chave, documento, resultado)", render)

    def test_destinos_e_downloads_intactos(self):
        for marcador in (
            'key="upload_docs_sumario_executivo"',
            'key="upload_docs_despacho_saneador"',
            'key="upload_docs_termo_apostila"',
            "Sumario_Executivo_Reajuste_Contratual.pdf",
            "Despacho_Saneador_Instrucao_Processual.docx",
            "Termo_de_Apostila_Reajuste_Contratual.docx",
            "pages/12_Adequacao_Orcamentaria.py",
            "pages/05_Garantia.py",
            "pages/13_DOU.py",
        ):
            self.assertIn(marcador, PAGINA)


if __name__ == "__main__":
    unittest.main()
