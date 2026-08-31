# -*- coding: utf-8 -*-
"""Publica os 14 defined names do RESULTADOS-CONTRACT-1 via Excel COM.

O aplicador nao escreve celulas, formulas, estilos ou validacoes. Trabalha em
copia temporaria, adiciona somente os names autorizados, recalcula, salva e
normaliza a reserializacao do Excel contra o pacote homologado e reabre o
resultado final no Excel real. O destino so e substituido depois das travas.
"""
from __future__ import annotations

import argparse
import gc
import hashlib
import json
import shutil
import tempfile
import zipfile
from pathlib import Path

import pythoncom
import win32com.client


SHA256_BASE_HOMOLOGADA = (
    "415301c0b2b8cf8ba2eb7a6d5a6b82169aa387ac4d3f73c2ffcc7185aa166255"
)

NOMES_NOVOS = {
    "EXECUTADO_APURADO": "=RESULTADOS!$B$83",
    "AJUSTES_DEVIDOS": "=RESULTADOS!$B$84",
    "CONFERENCIA_FORMACAO_VTA": "=RESULTADOS!$B$87",
    "PC_TOTAL_CADASTRADO": "=MEMORIA_RESULTADOS!$T$33",
    "PC_TOTAL_ATE_CORTE": "=MEMORIA_RESULTADOS!$T$34",
    "PC_TOTAL_COM_EFEITO": "=MEMORIA_RESULTADOS!$T$36",
    "PC_TOTAL_SEM_EFEITO": "=MEMORIA_RESULTADOS!$T$37",
    "AUDITORIA_SITUACAO_ATUAL_CONTRATO": "=MEMORIA_RESULTADOS!$W$50",
    "AUDITORIA_ULTIMA_REFERENCIA_ABERTURA": "=MEMORIA_RESULTADOS!$W$48",
    "AUDITORIA_COMPARATIVO_INTEGRAL": "=comparativo_VTA!$B$208",
    "AUDITORIA_DIFERENCA_REFERENCIAS": "=MEMORIA_RESULTADOS!$W$51",
    "AUDITORIA_SITUACAO_ATUAL_STATUS": "=RESULTADOS!$H$10",
    "AUDITORIA_ABERTURA_STATUS": "=RESULTADOS!$H$11",
    "AUDITORIA_CONFERENCIA_STATUS": "=MEMORIA_RESULTADOS!$W$52",
}

XL_CALC_AUTOMATIC = -4105
XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16
XL_CELLTYPE_ALL_VALIDATION = -4174


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _names(wb) -> dict[str, tuple[str, bool]]:
    return {
        str(item.Name): (str(item.RefersTo), bool(item.Visible))
        for item in wb.Names
    }


def _sheet_state(wb) -> list[tuple[str, int, object, bool]]:
    state = []
    for ws in wb.Worksheets:
        try:
            color = ws.Tab.Color
        except Exception:
            color = None
        state.append(
            (str(ws.Name), int(ws.Visible), color, bool(ws.ProtectContents))
        )
    return state


def _formula_counts(wb) -> dict[str, int]:
    counts = {}
    for ws in wb.Worksheets:
        try:
            counts[str(ws.Name)] = int(
                ws.UsedRange.SpecialCells(XL_CELLTYPE_FORMULAS).Count
            )
        except Exception:
            counts[str(ws.Name)] = 0
    return counts


def _resultados_snapshot(wb) -> dict[str, str]:
    ws = wb.Worksheets("RESULTADOS")
    snapshot = {}
    for row in range(1, 88):
        for column in range(1, 15):
            cell = ws.Cells(row, column)
            formula = cell.Formula
            if formula not in (None, ""):
                snapshot[str(cell.Address).replace("$", "")] = str(formula)
    return snapshot


def _validation_snapshot(wb) -> dict[str, tuple[int, str, str]]:
    snapshot = {}
    for ws in wb.Worksheets:
        try:
            cells = ws.UsedRange.SpecialCells(XL_CELLTYPE_ALL_VALIDATION)
        except Exception:
            continue
        for cell in cells.Cells:
            try:
                address = str(cell.Address).replace("$", "")
                snapshot[f"{ws.Name}!{address}"] = (
                    int(cell.Validation.Type),
                    str(cell.Validation.Formula1 or ""),
                    str(cell.Validation.Formula2 or ""),
                )
            except Exception:
                continue
    return snapshot


def _snapshot(wb) -> dict:
    return {
        "names": _names(wb),
        "sheets": _sheet_state(wb),
        "formula_counts": _formula_counts(wb),
        "resultados": _resultados_snapshot(wb),
        "validations": _validation_snapshot(wb),
    }


def _assert_no_errors(wb) -> None:
    problems = []
    for ws in wb.Worksheets:
        for kind in (XL_CELLTYPE_FORMULAS, XL_CELLTYPE_CONSTANTS):
            try:
                cells = ws.UsedRange.SpecialCells(kind, XL_ERRORS)
                problems.append(f"{ws.Name}!{cells.Address}")
            except Exception:
                continue
    if problems:
        raise RuntimeError(f"Erros estruturais apos recalculo: {problems}")


def _assert_unchanged(before: dict, after: dict) -> None:
    for key in ("sheets", "formula_counts", "resultados", "validations"):
        if before[key] != after[key]:
            raise RuntimeError(f"Trava violada; estrutura mudou em {key}.")
    changed = {
        name: (definition, after["names"].get(name))
        for name, definition in before["names"].items()
        if after["names"].get(name) != definition
    }
    if changed:
        raise RuntimeError(f"Defined names preexistentes mudaram: {changed}")


def _assert_new_names(wb) -> None:
    names = _names(wb)
    for name, target in NOMES_NOVOS.items():
        actual = names.get(name)
        if actual is None or actual[0].upper() != target.upper():
            raise RuntimeError(
                f"Defined name {name}: esperado {target}, encontrado {actual}."
            )
    if any(
        definition[0].upper() == "=MEMORIA_RESULTADOS!$T$35"
        for definition in names.values()
    ):
        raise RuntimeError("MEMORIA_RESULTADOS!T35 recebeu name proibido.")
    if names.get("VTA_FINAL", ("", False))[0].upper() != \
            "=MEMORIA_RESULTADOS!$B$26":
        raise RuntimeError("VTA_FINAL mudou de fonte.")


def _normalizar_pacote(base_path: Path, excel_path: Path) -> None:
    """Preserva todos os parts homologados e incorpora so os names do Excel."""
    normalized = excel_path.with_name(f"{excel_path.stem}.normalizado.xlsx")
    with zipfile.ZipFile(base_path, "r") as base_zip, zipfile.ZipFile(
        excel_path, "r"
    ) as excel_zip, zipfile.ZipFile(normalized, "w") as output_zip:
        base_workbook = base_zip.read("xl/workbook.xml").decode("utf-8")
        excel_workbook = excel_zip.read("xl/workbook.xml").decode("utf-8")

        imported = []
        for name in NOMES_NOVOS:
            marker = f'<definedName name="{name}">'
            start = excel_workbook.find(marker)
            if start < 0:
                raise RuntimeError(f"Name ausente no workbook salvo pelo Excel: {name}")
            end = excel_workbook.find("</definedName>", start)
            if end < 0:
                raise RuntimeError(f"Name malformado no workbook salvo pelo Excel: {name}")
            imported.append(excel_workbook[start : end + len("</definedName>")])

        closing = "</definedNames>"
        if closing not in base_workbook:
            raise RuntimeError("Pacote homologado nao possui bloco definedNames.")
        final_workbook = base_workbook.replace(
            closing, "".join(imported) + closing, 1
        ).encode("utf-8")

        for info in base_zip.infolist():
            payload = (
                final_workbook
                if info.filename == "xl/workbook.xml"
                else base_zip.read(info.filename)
            )
            output_zip.writestr(info, payload)

    with zipfile.ZipFile(normalized, "r") as check_zip, zipfile.ZipFile(
        base_path, "r"
    ) as base_zip:
        changed = sorted(
            info.filename
            for info in base_zip.infolist()
            if base_zip.read(info.filename) != check_zip.read(info.filename)
        )
    if changed != ["xl/workbook.xml"]:
        raise RuntimeError(f"Normalizacao alterou parts nao autorizados: {changed}")
    normalized.replace(excel_path)


def aplicar(path: Path) -> dict[str, object]:
    path = path.resolve()
    if not path.is_file():
        raise FileNotFoundError(path)
    before_hash = _sha256(path)
    if before_hash != SHA256_BASE_HOMOLOGADA:
        raise RuntimeError(
            "Template nao corresponde a base homologada: "
            f"esperado {SHA256_BASE_HOMOLOGADA}, encontrado {before_hash}."
        )

    temp_dir = Path(tempfile.mkdtemp(prefix="resultados_contract_pr1_"))
    temp_file = temp_dir / path.name
    shutil.copyfile(path, temp_file)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(
            str(temp_file), UpdateLinks=0, ReadOnly=False, CorruptLoad=0
        )
        before = _snapshot(workbook)
        collisions = sorted(set(NOMES_NOVOS) & set(before["names"]))
        if collisions:
            raise RuntimeError(f"Names novos ja existem: {collisions}")
        for name, target in NOMES_NOVOS.items():
            workbook.Names.Add(Name=name, RefersTo=target, Visible=True)
        after_add = _snapshot(workbook)
        _assert_unchanged(before, after_add)
        _assert_new_names(workbook)
        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()
        _assert_no_errors(workbook)
        workbook.Save()
        workbook.Close(SaveChanges=False)
        workbook = None

        _normalizar_pacote(path, temp_file)

        workbook = excel.Workbooks.Open(
            str(temp_file), UpdateLinks=0, ReadOnly=True, CorruptLoad=0
        )
        reopened = _snapshot(workbook)
        _assert_unchanged(before, reopened)
        _assert_new_names(workbook)
        _assert_no_errors(workbook)
        workbook.Close(SaveChanges=False)
        workbook = None
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()
        del workbook
        del excel
        gc.collect()
        pythoncom.CoUninitialize()

    shutil.copyfile(temp_file, path)
    after_hash = _sha256(path)
    shutil.rmtree(temp_dir, ignore_errors=True)
    return {
        "path": str(path),
        "sha256_before": before_hash,
        "sha256_after": after_hash,
        "names_added": NOMES_NOVOS,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("template", type=Path)
    args = parser.parse_args()
    print(json.dumps(aplicar(args.template), indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
