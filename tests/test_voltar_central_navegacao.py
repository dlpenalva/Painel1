"""Etapa 6: botao 'Voltar para Central' nas paginas Adequacao, Garantia e DOU.

Requisito: renomear o antigo botao funcional da Adequacao para
'← Voltar para Central' e replicar o MESMO destino/comportamento em Garantia e
DOU. O destino e a pagina pos-upload/Valor Global ja homologada
(pages/03_Valor_Global.py), que reidrata a apuracao (resultado_valor_global,
diagnostico_coleta_v2, assinatura_processada_upload_docs) sem novo upload,
reexibindo as quatro metricas e os seis documentos.

As paginas Streamlit executam ao serem importadas; a verificacao e estatica
(le o codigo-fonte), como nos demais testes de pagina do projeto.
"""
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PAGES = ROOT / "pages"

DESTINO_COMUM = "pages/03_Valor_Global.py"
PAGINAS = (
    "12_Adequacao_Orcamentaria.py",
    "05_Garantia.py",
    "13_DOU.py",
)


class TestVoltarParaCentral(unittest.TestCase):
    def test_cada_pagina_tem_botao_com_destino_comum(self):
        for nome in PAGINAS:
            fonte = (PAGES / nome).read_text(encoding="utf-8")
            self.assertIn("← Voltar para Central", fonte, f"{nome}: botao ausente")
            self.assertIn(
                f'st.switch_page("{DESTINO_COMUM}")', fonte,
                f"{nome}: destino nao aponta para a pagina pos-upload",
            )

    def test_adequacao_tem_apenas_um_botao_de_retorno(self):
        fonte = (PAGES / "12_Adequacao_Orcamentaria.py").read_text(encoding="utf-8")
        # o antigo rotulo foi substituido, nao duplicado
        self.assertNotIn("← Voltar para Valor Global", fonte)
        self.assertEqual(fonte.count("← Voltar para Central"), 1)

    def test_nenhum_botao_referencia_central_de_arquivos(self):
        for nome in PAGINAS:
            fonte = (PAGES / nome).read_text(encoding="utf-8")
            self.assertNotIn(
                'st.switch_page("pages/06_Central_Arquivos.py")', fonte,
                f"{nome}: destino incorreto (Central de Arquivos)",
            )


if __name__ == "__main__":
    unittest.main()
