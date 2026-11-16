"""Atualiza somente celulas de conferencia no pacote oficial (requer lxml).

Nao reserializa o workbook com openpyxl: preserva extensoes x14 e demais partes.
"""
from pathlib import Path
from io import BytesIO
import sys
import zipfile

from lxml import etree as ET
from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from _conferencias_execucao import aplicar_conferencias


def aplicar(path):
    original = path.read_bytes()
    wb = load_workbook(BytesIO(original))
    before = {s.title:{c.coordinate:c.value for row in s for c in row} for s in wb}
    aplicar_conferencias(wb)
    changes = {s.title:{c.coordinate:c.value for row in s for c in row
                       if c.value != before[s.title].get(c.coordinate)} for s in wb}
    ns = {'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    tag = '{'+ns['m']+'}'
    with zipfile.ZipFile(BytesIO(original)) as z:
        book = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        targets = {r.get('Id'):r.get('Target') for r in rels}
        patches = {}
        for sheet in book.find(tag+'sheets'):
            edits = changes.get(sheet.get('name'))
            if not edits: continue
            target = targets[sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')]
            target = target.lstrip('/') if target.startswith('/') else 'xl/'+target
            root = ET.fromstring(z.read(target)); data = root.find(tag+'sheetData')
            for addr,value in edits.items():
                rnum = int(''.join(filter(str.isdigit,addr)))
                row = data.find(f'{tag}row[@r="{rnum}"]')
                if row is None:
                    row = ET.Element(tag+'row',r=str(rnum))
                    i = next((i for i,r in enumerate(data) if int(r.get('r'))>rnum),len(data))
                    data.insert(i,row)
                cell = row.find(f'{tag}c[@r="{addr}"]')
                if cell is None:
                    cell = ET.Element(tag+'c',r=addr)
                    col = column_index_from_string(''.join(filter(str.isalpha,addr)))
                    i = next((i for i,c in enumerate(row) if column_index_from_string(''.join(filter(str.isalpha,c.get('r'))))>col),len(row))
                    row.insert(i,cell)
                for child in list(cell): cell.remove(child)
                cell.attrib.pop('t',None)
                if isinstance(value,str) and value.startswith('='):
                    ET.SubElement(cell,tag+'f').text = value[1:]
                else:
                    cell.set('t','inlineStr')
                    ET.SubElement(ET.SubElement(cell,tag+'is'),tag+'t').text = str(value)
            patches[target] = ET.tostring(root)
        out = BytesIO()
        with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as dest:
            for entry in z.infolist(): dest.writestr(entry,patches.get(entry.filename,z.read(entry.filename)))
    path.write_bytes(out.getvalue())

if __name__ == '__main__':
    aplicar(Path(sys.argv[1]))
