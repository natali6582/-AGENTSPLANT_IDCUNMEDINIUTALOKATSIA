from argparse import ArgumentParser
from copy import deepcopy
from difflib import unified_diff
from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZipFile
import hashlib
import json
import posixpath
import re
import xml.etree.ElementTree as ET


def digest(data):
    return hashlib.sha256(data).hexdigest()


def letters(number):
    value = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        value = chr(65 + remainder) + value
    return value


def workbook_model(data):
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    cells, references, locations, sheets = {}, {}, {}, []
    with ZipFile(BytesIO(data)) as archive:
        assert archive.testzip() is None
        shared = ["".join(item.itertext()) for item in
                  ET.fromstring(archive.read("xl/sharedStrings.xml")).findall("s:si", ns)]
        rels = {item.attrib["Id"]: item.attrib["Target"] for item in
                ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))}
        for sheet in ET.fromstring(archive.read("xl/workbook.xml")).findall("s:sheets/s:sheet", ns):
            name = sheet.attrib["name"]
            sheets.append(name)
            target = rels[sheet.attrib["{" + ns["r"] + "}id"]]
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            last_row = 0
            for row_index, row in enumerate(ET.fromstring(archive.read(part)).findall("s:sheetData/s:row", ns)):
                row_number = int(row.attrib.get("r", last_row + 1))
                last_row, column = row_number, 0
                for cell_index, cell in enumerate(row.findall("s:c", ns)):
                    column += 1
                    address = cell.attrib.get("r", letters(column) + str(row_number))
                    if "r" in cell.attrib:
                        column = 0
                        for character in re.match(r"[A-Z]+", address)[0]:
                            column = column * 26 + ord(character) - 64
                    key = (name, address)
                    value_node, inline = cell.find("s:v", ns), cell.find("s:is", ns)
                    if cell.attrib.get("t") == "s" and value_node is not None:
                        index = int(value_node.text)
                        value = shared[index]
                        references.setdefault(index, []).append(key)
                    else:
                        value = "".join(inline.itertext()) if inline is not None else value_node.text if value_node is not None else ""
                    cells[key] = value
                    locations[key] = (part, row_index, cell_index)
    return cells, references, locations, sheets


def patch_cell(data, key, value):
    cells, references, locations, sheets = workbook_model(data)
    assert cells[("מידע כללי", "F1")] == "תוספת לסוכן זה"
    changed_parts = {}
    with ZipFile(BytesIO(data)) as source:
        if cells[key]:
            indices = [index for index, keys in references.items() if key in keys]
            assert len(indices) == 1 and references[indices[0]] == [key]
            part = "xl/sharedStrings.xml"
            original = source.read(part)
            strings = list(re.finditer(rb"<(?:\w+:)?si\b[^>]*>.*?</(?:\w+:)?si>", original, re.S))
            match = strings[indices[0]]
            text_pattern = rb"(<(?:\w+:)?t\b[^>]*>)(.*?)(</(?:\w+:)?t>)"
            assert len(list(re.finditer(text_pattern, match.group(), re.S))) == 1
            replacement = re.sub(text_pattern, lambda m: m[1] + escape(value).encode("utf-8") + m[3],
                                 match.group(), flags=re.S)
            changed_parts[part] = original[:match.start()] + replacement + original[match.end():]
        else:
            part, row_index, cell_index = locations[key]
            original = source.read(part)
            rows = list(re.finditer(rb"<row\b[^>]*>.*?</row>", original, re.S))
            row_match = rows[row_index]
            row_cells = list(re.finditer(rb"<c\b[^>]*(?:/>|>.*?</c>)", row_match.group(), re.S))
            cell_match = row_cells[cell_index]
            old = cell_match.group()
            assert re.fullmatch(rb'<c\s+s="[^"]+"\s*/>', old)
            shared_part = "xl/sharedStrings.xml"
            shared_xml = source.read(shared_part)
            shared_nodes = re.findall(rb"<si\b[^>]*>.*?</si>", shared_xml, re.S)
            shared_index = len(shared_nodes)
            replacement = old[:-2].rstrip() + b' t="s"><v>' + str(shared_index).encode() + b"</v></c>"
            new_row = row_match.group()[:cell_match.start()] + replacement + row_match.group()[cell_match.end():]
            changed_parts[part] = original[:row_match.start()] + new_row + original[row_match.end():]
            new_shared = re.sub(rb'(<sst\b[^>]*\bcount=")(\d+)(")',
                                lambda m: m[1] + str(int(m[2]) + 1).encode() + m[3], shared_xml, count=1)
            new_shared = re.sub(rb'(<sst\b[^>]*\buniqueCount=")(\d+)(")',
                                lambda m: m[1] + str(int(m[2]) + 1).encode() + m[3], new_shared, count=1)
            assert new_shared.count(b"</sst>") == 1
            new_shared = new_shared.replace(
                b"</sst>", b'<si><t xml:space="preserve">' + escape(value).encode("utf-8") + b"</t></si></sst>")
            assert re.findall(rb"<si\b[^>]*>.*?</si>", new_shared, re.S)[:-1] == shared_nodes
            changed_parts[shared_part] = new_shared
        output = BytesIO()
        with ZipFile(output, "w") as destination:
            destination.comment = source.comment
            for info in source.infolist():
                destination.writestr(info, changed_parts.get(info.filename, source.read(info.filename)))
    result = output.getvalue()
    after, _, _, after_sheets = workbook_model(result)
    assert sheets == after_sheets and len(sheets) == 60
    assert cells.keys() == after.keys()
    assert [cell for cell in cells if cells[cell] != after[cell]] == [key]
    assert after[key] == value
    with ZipFile(BytesIO(data)) as before_zip, ZipFile(BytesIO(result)) as after_zip:
        assert before_zip.namelist() == after_zip.namelist()
        assert set(part for part in before_zip.namelist() if before_zip.read(part) != after_zip.read(part)) == set(changed_parts)
        for before_info, after_info in zip(before_zip.infolist(), after_zip.infolist()):
            for attribute in ["filename", "date_time", "compress_type", "comment", "extra", "create_system",
                              "create_version", "extract_version", "flag_bits", "volume", "internal_attr", "external_attr"]:
                assert getattr(before_info, attribute) == getattr(after_info, attribute)
        if not cells[key]:
            # The only worksheet change is the one blank cell's value and necessary type.
            part, row_index, cell_index = locations[key]
            ns_uri = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            before_tree = ET.fromstring(before_zip.read(part))
            after_tree = ET.fromstring(after_zip.read(part))
            before_row = before_tree.findall(".//{" + ns_uri + "}row")[row_index]
            after_row = after_tree.findall(".//{" + ns_uri + "}row")[row_index]
            assert before_row[cell_index].get("s") == after_row[cell_index].get("s")
            after_row.remove(after_row[cell_index])
            after_row.insert(cell_index, deepcopy(before_row[cell_index]))
            assert ET.tostring(before_tree) == ET.tostring(after_tree)
    return result, {"changed_cells": [list(key)], "sheet_count": len(sheets),
                    "changed_package_parts": list(changed_parts), "non_target_cells_and_properties_unchanged": True}


def replace_once(text, old, new):
    assert text.count(old) == 1, old[:100]
    return text.replace(old, new, 1)


def build_prompt(original, config):
    bom = original.startswith(b"\xef\xbb\xbf")
    decoded = original.decode("utf-8-sig")
    newline = "\r\n" if "\r\n" in decoded else "\n"
    text = decoded.replace("\r\n", "\n")
    initial = text
    start = text.index("באיתור בקש מועמדים העומדים בתנאים.")
    end = text.index("\n\nמקורות המחקר", start)
    text = text[:start] + (config / "research_paragraph.txt").read_text(encoding="utf-8").rstrip() + text[end:]
    start = text.index("סדר העבודה ותנאי המשך\n")
    end = text.index("\n\nכיוון שוק מבוסס מחקר", start)
    text = text[:start] + (config / "workflow_section.txt").read_text(encoding="utf-8").rstrip() + text[end:]
    text = replace_once(text,
        "חשב כל פרופיל בנפרד לפי מיפוי הטווחים וכללי ההחרגה שבתוספת GetAllocationPolicy. המיפוי משמש לחישוב בלבד ואינו משנה את הגדרות רכיבי המדיניות או החשיפות.",
        "בחן כל פרופיל בנפרד, גם אם הפילוח הקיים בו חורג, לפי מיפוי הטווחים וכללי הבחינה שבתוספת GetAllocationPolicy. המיפוי משמש לחישוב בלבד ואינו משנה את הגדרות רכיבי המדיניות או החשיפות. חריגה קיימת אינה כיוון שוק ואינה עילה לשינוי ללא המחקר הנדרש. חשב מהערכים המקוריים באמצעות אותם עוגנים, העברות וספים; אין מסלול תיקון אוטומטי נפרד.")
    text = replace_once(text,
        "- סכום הפילוח בדיוק 100, וכל ערך נמצא בטווח הממופה אליו.",
        "- סכום הפילוח בדיוק 100, וכל ארבעת ערכי הפילוח הסופיים נמצאים בטווחים הממופים, לרבות ערכים שלא שונו. שיפור שמשאיר חריגה כלשהי בפרופיל אינו מספיק.")
    text = replace_once(text,
        "פרופיל שהוחרג מראש בגלל חריגה מטווח נשאר בערכיו המקוריים ומשתתף בבדיקת הסדר; אין לדרוש תיקון שלו כדי לאשר חבילות תקינות אחרות.",
        "פרופיל שלא חושבה עבורו חבילה תקפה נשאר בערכיו המקוריים ומשתתף בבדיקת הסדר; אין לפסול אותו מראש רק בגלל חריגה קיימת. אפשר לאשר חבילות תקינות של פרופילים אחרים אם הן עומדות גם בבדיקת הסדר. הצג במפורש חריגה שנותרה בפרופיל שלא השתנה, ואל תציג את המדיניות כולה כאילו כל פרופיליה עומדים בגבולות.")
    text = replace_once(text,
        "אין לחשב הצעת שינוי כאשר חסרה מדיניות תקינה או שהאימות הנדרש לא הושלם.",
        "אין לחשב הצעת שינוי כאשר חסרים נתוני מדיניות תקינים מבנית או שהאימות הנדרש לא הושלם.")
    text = replace_once(text,
        "הגעה לעוגנים, חריגה קיימת מהטווחים או פסילת חבילות בבדיקת ההצעה.",
        "הגעה לעוגנים, חריגה שלא ניתן היה להביא לתוך הגבולות באמצעות השינויים המבוססים שחושבו, או פסילת חבילות בבדיקת ההצעה.")
    assert "כללי ההחרגה" not in text and "שהוחרג מראש" not in text
    assert "אל תוסיף חיפוש שלישי" not in text
    for invariant in ["0.75 × (Max - Min)", "0.25 × (Max - Min)", "לפחות שני מוסדות שונים",
                      "לפחות 2 נקודות אחוז", "אין להשתמש ברכיב ללא כיוון מוכח כמקור מאזן",
                      "equity אינו יורד", "cash אינו עולה"]:
        assert initial.count(invariant) == text.count(invariant)
    output = text.replace("\n", newline).encode("utf-8")
    if bom:
        output = b"\xef\xbb\xbf" + output
    difference = "".join(unified_diff(initial.splitlines(True), text.splitlines(True),
                                     fromfile="original_prompt", tofile="updated_prompt"))
    return output, difference


def run(inputs, config, output):
    assert not output.exists(), "Each build requires a fresh output directory"
    output.mkdir()
    baseline = json.loads((config / "baseline.json").read_text(encoding="utf-8-sig"))
    reports = []
    for entry in baseline["targets"]:
        source = (inputs / entry["name"]).read_bytes()
        assert digest(source) == entry["sha256"].lower()
        if entry["name"].endswith(".txt"):
            result, difference = build_prompt(source, config)
            (output / "prompt.diff").write_text(difference, encoding="utf-8")
            checks = {"encoding_and_line_endings_preserved": True}
        else:
            config_file = "policy_addition.txt" if entry["name"].startswith("GetAllocationPolicy") else "search_addition.txt"
            value = (config / config_file).read_text(encoding="utf-8").rstrip()
            result, checks = patch_cell(source, ("מידע כללי", "F2"), value)
        (output / entry["name"]).write_bytes(result)
        assert (output / entry["name"]).read_bytes() == result
        reports.append({"file": entry["name"], "sha256": digest(result), **checks})
    (output / "verification.json").write_text(json.dumps(reports, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(reports, ensure_ascii=False))


def main():
    parser = ArgumentParser()
    parser.add_argument("--inputs", type=Path, required=True)
    parser.add_argument("--config", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    run(args.inputs.resolve(), args.config.resolve(), args.output.resolve())


if __name__ == "__main__":
    main()
