"""C0.3 — O documento deve registrar a ORIGEM REAL da Coleta.

Prova que o rótulo documental de origem reflete o arquivo efetivamente
processado: Arquivo Oficial (Modelo B) recebe o nome oficial; layout legado
(Modelo A) é identificado explicitamente como legado — nunca o nome oficial
indiscriminadamente.
"""

import unittest
from pathlib import Path

from _coleta_oficial import NOME_ARQUIVO_COLETA_OFICIAL, obter_coleta_oficial_bytes
from _coleta_reajuste_documentos import _rotulo_origem_coleta

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_LEGADO = ROOT / "templates" / "Coleta_Reajuste.xlsx"


class OrigemDocumentalColeta(unittest.TestCase):
    def test_arquivo_oficial_recebe_nome_oficial(self):
        rotulo = _rotulo_origem_coleta(obter_coleta_oficial_bytes())
        self.assertEqual(rotulo, NOME_ARQUIVO_COLETA_OFICIAL)

    def test_arquivo_legado_e_identificado_como_legado(self):
        rotulo = _rotulo_origem_coleta(TEMPLATE_LEGADO.read_bytes())
        self.assertIn("legado", rotulo.lower())
        self.assertIn("Modelo A", rotulo)
        self.assertNotEqual(rotulo, NOME_ARQUIVO_COLETA_OFICIAL)


if __name__ == "__main__":
    unittest.main()
