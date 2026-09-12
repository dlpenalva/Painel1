"""Atualiza somente o XML de RESULTADOS, preservando os demais membros XLSX."""
from io import BytesIO
from pathlib import Path
import sys
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from _apresentacao_pc_xls import ALTURA_LINHA_86, valores_apresentacao_pc


def aplicar(path):
    path = Path(path)
    original = path.read_bytes()
    wb = load_workbook(BytesIO(original), read_only=True)
    valores = valores_apresentacao_pc(wb["RESULTADOS"])
    wb.close()
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with ZipFile(BytesIO(original)) as z:
        workbook = etree.fromstring(z.read("xl/workbook.xml"))
        sheet = workbook.xpath('//m:sheet[@name="RESULTADOS"]', namespaces=ns)[0]
        rid = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = next(r.get("Target") for r in rels if r.get("Id") == rid)
        member = target.lstrip("/") if target.startswith("/") else "xl/" + target
        root = etree.fromstring(z.read(member))
        for address, value in valores.items():
            cell = root.xpath(f'//m:c[@r="{address}"]', namespaces=ns)[0]
            cell.attrib.pop("t", None)
            for child in list(cell):
                if etree.QName(child).localname in ("f", "v", "is"):
                    cell.remove(child)
            etree.SubElement(cell, "{" + ns["m"] + "}f").text = value[1:]
        row = root.xpath('//m:row[@r="25"]', namespaces=ns)[0]
        row.set("ht", str(max(float(row.get("ht", 0)), 48)))
        row.set("customHeight", "1")
        row86 = root.xpath('//m:row[@r="86"]', namespaces=ns)[0]
        row86.set("ht", str(max(float(row86.get("ht", 0)), ALTURA_LINHA_86)))
        row86.set("customHeight", "1")
        output = BytesIO()
        with ZipFile(output, "w", ZIP_DEFLATED) as dest:
            for info in z.infolist():
                dest.writestr(info, etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True) if info.filename == member else z.read(info.filename))
    path.write_bytes(output.getvalue())


if __name__ == "__main__":
    aplicar(sys.argv[1] if len(sys.argv) > 1 else ROOT / "templates/COLETA_REAJUSTE_OFICIAL.xlsx")
