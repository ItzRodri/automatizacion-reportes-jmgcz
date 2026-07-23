"""
Reinyecta las validaciones de lista (desplegables) de tipo x14 que openpyxl
elimina al guardar. Estas validaciones son las que usan como origen un rango
de OTRA hoja (Subalcaldes!A2:A16 y Directorio!A2:A11), una característica de
Excel que openpyxl no soporta y descarta silenciosamente (con warning) al
volver a guardar el archivo.

Uso:
    from xlsx_dropdown_fix import restore_dropdowns
    restore_dropdowns("archivo.xlsx", sheet_name="Matriz")
"""
import re
import shutil
import zipfile
import os

EXT_BLOCK_TEMPLATE = (
    '<extLst><ext uri="{{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}}" '
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
    '<x14:dataValidations count="{count}" '
    'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">{validations}'
    '</x14:dataValidations></ext></extLst>'
)

VALIDATION_TEMPLATE = (
    '<x14:dataValidation type="list" allowBlank="1">'
    '<x14:formula1><xm:f>{formula}</xm:f></x14:formula1>'
    '<x14:formula2><xm:f>0</xm:f></x14:formula2>'
    '<xm:sqref>{sqref}</xm:sqref></x14:dataValidation>'
)


def _find_sheet_xml_path(tmpdir, sheet_name):
    wb_xml = open(os.path.join(tmpdir, "xl", "workbook.xml"), encoding="utf-8").read()
    rels_xml = open(os.path.join(tmpdir, "xl", "_rels", "workbook.xml.rels"), encoding="utf-8").read()

    m = re.search(r'<sheet[^>]*name="%s"[^>]*r:id="(rId\d+)"' % re.escape(sheet_name), wb_xml)
    if not m:
        raise ValueError(f"No se encontró la hoja '{sheet_name}' en workbook.xml")
    rid = m.group(1)

    m2 = None
    for rel_match in re.finditer(r'<Relationship\b[^>]*/>', rels_xml):
        rel_tag = rel_match.group(0)
        if f'Id="{rid}"' in rel_tag:
            m2 = re.search(r'Target="([^"]+)"', rel_tag)
            break
    if not m2:
        raise ValueError(f"No se encontró el relationship '{rid}'")
    target = m2.group(1).replace("worksheets/", "").lstrip("/")
    target = target.split("/")[-1]
    return os.path.join(tmpdir, "xl", "worksheets", target)


def restore_dropdowns(filepath, sheet_name="Matriz",
                       validations=(
                           ("Subalcaldes!$A$2:$A$16", "C2:C101"),
                           ("Directorio!$A$2:$A$11", "F2:F101"),
                       )):
    """Reinyecta desplegables x14 (listas que apuntan a otra hoja) en `sheet_name`."""
    tmpdir = filepath + "__tmp_extract"
    if os.path.exists(tmpdir):
        shutil.rmtree(tmpdir)
    os.makedirs(tmpdir)

    with zipfile.ZipFile(filepath, "r") as z:
        z.extractall(tmpdir)

    sheet_path = _find_sheet_xml_path(tmpdir, sheet_name)
    xml = open(sheet_path, encoding="utf-8").read()

    if "x14:dataValidations" in xml:
        shutil.rmtree(tmpdir)
        return  # ya tiene los desplegables, no hacer nada

    val_xml = "".join(
        VALIDATION_TEMPLATE.format(formula=formula, sqref=sqref)
        for formula, sqref in validations
    )
    ext_block = EXT_BLOCK_TEMPLATE.format(count=len(validations), validations=val_xml)

    if "<extLst>" in xml:
        xml = xml.replace("</extLst>", ext_block.replace("<extLst>", "").replace("</extLst>", "") + "</extLst>", 1)
    elif "</worksheet>" in xml:
        xml = xml.replace("</worksheet>", ext_block + "</worksheet>")
    else:
        raise ValueError("No se encontró </worksheet> para insertar los desplegables")

    with open(sheet_path, "w", encoding="utf-8") as f:
        f.write(xml)

    tmp_zip = filepath + ".tmp"
    with zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as zout:
        for root, _, files in os.walk(tmpdir):
            for fname in files:
                full = os.path.join(root, fname)
                arcname = os.path.relpath(full, tmpdir)
                zout.write(full, arcname)

    shutil.rmtree(tmpdir)
    os.replace(tmp_zip, filepath)


if __name__ == "__main__":
    import sys
    restore_dropdowns(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "Matriz")
    print("Desplegables reinyectados en", sys.argv[1])
