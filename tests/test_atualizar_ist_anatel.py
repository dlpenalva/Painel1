from datetime import date
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from tools.atualizar_ist_anatel import (
    ErroAtualizacaoIST,
    RegistroIST,
    anexar_registros,
    extrair_registros_ist,
    ler_registros_locais,
    planejar_atualizacao,
)


HTML_EXEMPLO = """
<html><body>
<table><tbody>
<tr><td><strong>Referência</strong></td><td>Variação</td><td>IST</td></tr>
<tr><td><div>Abr/26</div></td><td><p>1,34%</p></td><td><p>358,475</p></td></tr>
<tr><td><p>Mai/26</p></td><td>0,85%</td><td><div>361,530</div></td></tr>
</tbody></table>
</body></html>
"""


class TestAtualizadorISTAnatel(unittest.TestCase):
    def test_extrai_tabela_html_atual_da_anatel(self):
        registros = extrair_registros_ist(HTML_EXEMPLO)
        self.assertEqual([r.competencia for r in registros], [date(2026, 4, 1), date(2026, 5, 1)])
        self.assertEqual(registros[-1].indice, Decimal("361.530"))

    def test_aceita_numero_portugues_com_milhar(self):
        html = HTML_EXEMPLO.replace("361,530", "1.234,567")
        self.assertEqual(extrair_registros_ist(html)[-1].indice, Decimal("1234.567"))

    def test_bloqueia_valores_conflitantes_na_mesma_competencia(self):
        html = HTML_EXEMPLO + HTML_EXEMPLO.replace("361,530", "999,999")
        with self.assertRaises(ErroAtualizacaoIST):
            extrair_registros_ist(html)

    def test_planeja_somente_competencia_nova_e_continua(self):
        locais = [
            RegistroIST(date(2026, 3, 1), Decimal("353.721")),
            RegistroIST(date(2026, 4, 1), Decimal("358.475")),
        ]
        oficiais = locais + [RegistroIST(date(2026, 5, 1), Decimal("361.530"))]
        self.assertEqual(planejar_atualizacao(locais, oficiais), oficiais[-1:])

    def test_bloqueia_divergencia_historica(self):
        locais = [RegistroIST(date(2026, 4, 1), Decimal("358.475"))]
        oficiais = [RegistroIST(date(2026, 4, 1), Decimal("358.999"))]
        with self.assertRaises(ErroAtualizacaoIST):
            planejar_atualizacao(locais, oficiais)

    def test_bloqueia_lacuna_entre_local_e_novo_mes(self):
        locais = [RegistroIST(date(2026, 3, 1), Decimal("353.721"))]
        oficiais = locais + [RegistroIST(date(2026, 5, 1), Decimal("361.530"))]
        with self.assertRaises(ErroAtualizacaoIST):
            planejar_atualizacao(locais, oficiais)

    def test_anexa_sem_reescrever_registros_existentes(self):
        with TemporaryDirectory() as pasta:
            caminho = Path(pasta) / "ist.csv"
            original = "MES_ANO;INDICE_NIVEL\nmar/26;353,721\nabr/26;358,475\n"
            caminho.write_text(original, encoding="utf-8")
            anexar_registros(caminho, [RegistroIST(date(2026, 5, 1), Decimal("361.530"))])
            self.assertEqual(
                caminho.read_text(encoding="utf-8"),
                original + "mai/26;361,530\n",
            )
            self.assertEqual(ler_registros_locais(caminho)[-1].competencia, date(2026, 5, 1))

    def test_preserva_bom_utf8_do_arquivo_oficial(self):
        with TemporaryDirectory() as pasta:
            caminho = Path(pasta) / "ist.csv"
            caminho.write_text(
                "MES_ANO;INDICE_NIVEL\nabr/26;358,475\n",
                encoding="utf-8-sig",
            )
            anexar_registros(caminho, [RegistroIST(date(2026, 5, 1), Decimal("361.530"))])
            self.assertTrue(caminho.read_bytes().startswith(b"\xef\xbb\xbf"))


if __name__ == "__main__":
    unittest.main()
