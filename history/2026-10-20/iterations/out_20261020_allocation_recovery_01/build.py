from argparse import ArgumentParser
from difflib import unified_diff
from pathlib import Path
import hashlib
import json
from xlsx_helpers import patch_cell, workbook_model


def digest(data):
    return hashlib.sha256(data).hexdigest()


def update_prompt(source, work):
    decoded = source.decode("utf-8-sig")
    newline = "\r\n" if "\r\n" in decoded else "\n"
    before = decoded.replace("\r\n", "\n")
    start = before.index("סדר העבודה ותנאי המשך\n")
    end = before.index("\n\nכיוון שוק מבוסס מחקר", start)
    section = (work / "workflow_section.txt").read_text(encoding="utf-8").rstrip()
    after = before[:start] + section + before[end:]
    assert before[:start] == after[:start]
    assert after.endswith(before[end:])
    assert "בנוסף מותרת קריאה אחת בלבד להשלמת אימות" not in after
    assert "קריאת השלמה רביעית" in after and "ניסיון התאוששות אחד בלבד המשותף לכל המחזור" in after
    result = after.replace("\n", newline).encode("utf-8")
    if source.startswith(b"\xef\xbb\xbf"):
        result = b"\xef\xbb\xbf" + result
    difference = "".join(unified_diff(before.splitlines(True), after.splitlines(True),
                                     fromfile="before_prompt", tofile="after_prompt"))
    return result, difference


def run(work, output):
    assert not output.exists(), "Use a fresh build directory"
    output.mkdir()
    baseline = json.loads((work / "baseline.json").read_text(encoding="utf-8-sig"))
    reports = []
    for item in baseline["targets"]:
        original = (work / "inputs" / item["name"]).read_bytes()
        assert digest(original) == item["sha256"].lower()
        if item["name"].endswith(".txt"):
            result, difference = update_prompt(original, work)
            (output / "prompt.diff").write_text(difference, encoding="utf-8")
            checks = {"changed_section": "סדר העבודה ותנאי המשך", "all_other_sections_unchanged": True}
        elif item["name"].startswith("WebSearch"):
            value = (work / "search_addition.txt").read_text(encoding="utf-8").rstrip()
            result, checks = patch_cell(original, ("מידע כללי", "F2"), value)
            assert checks["changed_package_parts"] == ["xl/sharedStrings.xml"]
        else:
            result = original
            checks = {"changed_cells": [], "sheet_count": len(workbook_model(original)[3]),
                      "byte_identical_to_source": True}
        (output / item["name"]).write_bytes(result)
        assert (output / item["name"]).read_bytes() == result
        reports.append({"file": item["name"], "sha256": digest(result), **checks})
    (output / "verification.json").write_text(json.dumps(reports, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(reports, ensure_ascii=False))


def main():
    parser = ArgumentParser()
    parser.add_argument("--work", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    run(args.work.resolve(), args.output.resolve())


if __name__ == "__main__":
    main()
