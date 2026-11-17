"""Coleta 12.8: cor de ressalva na faixa RESULTADOS!B12:G13 (requer lxml).

Com STATUS_RESULTADOS = CONFERIR ou BASE INSUFICIENTE, a faixa de situacao
(C12 = "APURACAO CONCLUIDA") recebia o verde de VALIDADO (C6EFCE/006100), cor
que em RESULTADOS e reservada a VALIDADO/RECONCILIADO (G6, C27:C34). Esta
ferramenta insere UMA regra de formatacao condicional, logo apos a de REVISE,
que pinta a faixa com o ambar de ressalva ja usado por ESTIMADO em G6
(FFEB9C/9C5700). Reaproveita o dxf existente: styles.xml nao muda.

Patch XML direto (sem reserializar com openpyxl): preserva extensoes x14 e
as demais partes do pacote. Idempotente.
"""
from __future__ import annotations

from io import BytesIO
from pathlib import Path
import sys
import tempfile
import zipfile

from lxml import etree as ET

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
T = "{" + NS + "}"

FAIXA = "B12:G13"
FORMULA_REVISE = 'LEFT($C$12,4)="REVI"'
FORMULA_RESSALVA = (
    'AND(LEFT($C$12,4)="APUR",'
    'OR(STATUS_RESULTADOS="CONFERIR",STATUS_RESULTADOS="BASE INSUFICIENTE"))'
)
# Regra de G6 cujo dxf (ambar de ressalva) e reaproveitado.
FORMULA_FONTE_DXF = 'ISNUMBER(SEARCH("ESTIMAD",$G$6))'


def _alvo_resultados(z: zipfile.ZipFile) -> str:
    book = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    alvos = {r.get("Id"): r.get("Target") for r in rels}
    for sheet in book.find(T + "sheets"):
        if sheet.get("name") == "RESULTADOS":
            alvo = alvos[sheet.get("{" + NS_R + "}id")]
            return alvo.lstrip("/") if alvo.startswith("/") else "xl/" + alvo
    raise ValueError("Aba RESULTADOS ausente.")


def _regras(root, sqref: str):
    for cf in root.findall(T + "conditionalFormatting"):
        if cf.get("sqref") == sqref:
            return cf
    return None


def patch_planilha(xml: bytes) -> bytes | None:
    """Devolve o XML da aba com a regra inserida, ou None se ja estiver."""
    root = ET.fromstring(xml)
    faixa = _regras(root, FAIXA)
    g6 = _regras(root, "G6")
    if faixa is None or g6 is None:
        raise ValueError("Formatacao condicional de B12:G13/G6 nao encontrada.")
    formulas = [r.findtext(T + "formula") for r in faixa.findall(T + "cfRule")]
    if FORMULA_RESSALVA in formulas:
        return None
    dxf = next(
        (r.get("dxfId") for r in g6.findall(T + "cfRule")
         if r.findtext(T + "formula") == FORMULA_FONTE_DXF),
        None,
    )
    # O Excel reordena os <cfRule> ao salvar e renumera dxfId: a precedencia e
    # o atributo priority, e REVISE/dxf sao localizados pela formula.
    revise = next(
        (r for r in faixa.findall(T + "cfRule")
         if r.findtext(T + "formula") == FORMULA_REVISE),
        None,
    )
    if dxf is None or revise is None:
        raise ValueError("Estrutura inesperada nas regras de B12:G13/G6.")
    # Nova regra logo apos REVISE: prioridade = REVISE + 1; as seguintes da
    # mesma faixa (APUR com/sem potencial) e quaisquer posteriores descem 1.
    p_revise = int(revise.get("priority"))
    for regra in root.iter(T + "cfRule"):
        p = int(regra.get("priority"))
        if p > p_revise:
            regra.set("priority", str(p + 1))
    # Sem stopIfTrue (regra da aba executiva: regras sobrepostas convivem); a
    # prioridade maior ja faz o ambar vencer o verde/oliva nas mesmas propriedades.
    nova = ET.Element(T + "cfRule", type="expression", dxfId=dxf,
                      priority=str(p_revise + 1))
    ET.SubElement(nova, T + "formula").text = FORMULA_RESSALVA
    faixa.insert(list(faixa).index(revise) + 1, nova)
    # Mesmo formato de serializacao de tools/aplicar_conferencias_execucao.py.
    return ET.tostring(root)


def aplicar(path: Path) -> bool:
    """Aplica no pacote em `path` (atomicamente). Retorna True se alterou."""
    path = Path(path)
    original = path.read_bytes()
    with zipfile.ZipFile(BytesIO(original)) as z:
        alvo = _alvo_resultados(z)
        novo = patch_planilha(z.read(alvo))
        if novo is None:
            return False
        out = BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dest:
            for entry in z.infolist():
                dest.writestr(entry, novo if entry.filename == alvo else z.read(entry.filename))
    with tempfile.NamedTemporaryFile(dir=path.parent, suffix=".tmp", delete=False) as tmp:
        temporario = Path(tmp.name)
        tmp.write(out.getvalue())
    try:
        temporario.replace(path)
    finally:
        temporario.unlink(missing_ok=True)
    return True


if __name__ == "__main__":
    for arg in sys.argv[1:]:
        print(arg, "alterado" if aplicar(Path(arg)) else "ja aplicado")
