"""Etapa 5a: divergencia XLS x Python nao bloqueia os 3 documentos nominais.

Separacao de responsabilidades exigida:
  1. bloqueio da apuracao/formalizacao (pode_formalizar) -> inalterado;
  2. disponibilidade documental -> Sumario Executivo, Despacho Saneador e
     Termo de Apostila permanecem habilitados quando o UNICO motivo e a
     divergencia relevante XLS x Python;
  3. confiabilidade de campos individuais -> mascaramento (Etapa 5b).

Aqui provamos apenas (2), na funcao pura aplicar_bloqueio_documental.
"""
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from _coleta_reajuste_documentos import (
    DOCS_LIBERADOS_APESAR_DIVERGENCIA,
    aplicar_bloqueio_documental,
)

BLOQUEIO_DIVERGENCIA = (
    "Divergência relevante XLS × Python em REM_ATUALIZADO_OFICIAL "
    "(Valor remanescente atualizado): XLS=201090.75 vs Python=201090.94. "
    "Nenhum dos resultados é adotado automaticamente; equalizar antes de formalizar."
)
BLOQUEIO_OUTRO = "Índice ou fator acumulado ausente em: C2."


def _capacidades():
    docs = {}
    for chave in (
        "sumario_executivo", "despacho_saneador", "termo_apostila",
        "garantia_contratual", "dou", "minuta_apostilamento", "relatorio_executivo",
    ):
        docs[chave] = {
            "nome": chave, "habilitado": True, "estado": "completo",
            "rotulo": "Disponível", "classificacao": "", "motivo": "",
        }
    return {"documentos": docs}


class TestDisponibilidadeDocumental(unittest.TestCase):
    def test_divergencia_nao_bloqueia_os_tres_documentos(self):
        cap = _capacidades()
        aplicar_bloqueio_documental(cap, [BLOQUEIO_DIVERGENCIA])
        for chave in DOCS_LIBERADOS_APESAR_DIVERGENCIA:
            self.assertTrue(
                cap["documentos"][chave]["habilitado"],
                f"{chave} deveria seguir disponivel apesar da divergencia",
            )

    def test_divergencia_bloqueia_os_demais_documentos(self):
        cap = _capacidades()
        aplicar_bloqueio_documental(cap, [BLOQUEIO_DIVERGENCIA])
        for chave in ("garantia_contratual", "dou", "relatorio_executivo"):
            self.assertFalse(
                cap["documentos"][chave]["habilitado"],
                f"{chave} deveria continuar bloqueado pela divergencia (regra atual)",
            )

    def test_bloqueio_nao_divergencia_bloqueia_ate_os_tres(self):
        cap = _capacidades()
        aplicar_bloqueio_documental(cap, [BLOQUEIO_OUTRO])
        for chave in DOCS_LIBERADOS_APESAR_DIVERGENCIA:
            self.assertFalse(
                cap["documentos"][chave]["habilitado"],
                f"{chave} deveria ser bloqueado por motivo que nao e divergencia",
            )

    def test_divergencia_mais_outro_bloqueio_tambem_bloqueia_os_tres(self):
        cap = _capacidades()
        aplicar_bloqueio_documental(cap, [BLOQUEIO_DIVERGENCIA, BLOQUEIO_OUTRO])
        for chave in DOCS_LIBERADOS_APESAR_DIVERGENCIA:
            self.assertFalse(
                cap["documentos"][chave]["habilitado"],
                f"{chave}: havendo outro bloqueio alem da divergencia, bloqueia",
            )
            self.assertEqual(cap["documentos"][chave]["motivo"], BLOQUEIO_OUTRO)

    def test_sem_bloqueios_preserva_habilitados(self):
        cap = _capacidades()
        aplicar_bloqueio_documental(cap, [])
        self.assertTrue(all(d["habilitado"] for d in cap["documentos"].values()))


if __name__ == "__main__":
    unittest.main()
