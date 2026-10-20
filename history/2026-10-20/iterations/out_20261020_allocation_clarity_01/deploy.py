from argparse import ArgumentParser
from pathlib import Path
import hashlib
import json


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run(work):
    baseline = json.loads((work / "baseline.json").read_text(encoding="utf-8-sig"))
    checks = json.loads((work / "build_final_a" / "verification.json").read_text(encoding="utf-8"))
    expected = {entry["file"]: entry["sha256"] for entry in checks}
    assert set(expected) == {"פרומפט עדכון 12_57 20_10.txt",
                             "GetAllocationPolicy V2.xlsx", "WebSearch (1) V2.xlsx"}
    assert {entry["name"] for entry in baseline["targets"]} == set(expected)
    for path in (work / "build_final_a").iterdir():
        assert path.read_bytes() == (work / "build_final_b" / path.name).read_bytes()
    assert (work / "cases_a.json").read_bytes() == (work / "cases_b.json").read_bytes()
    for entry in baseline["protected"]:
        assert digest(Path(entry["path"])) == entry["sha256"].lower()
    payloads = []
    for entry in baseline["targets"]:
        destination = Path(entry["path"])
        source = work / "build_final_a" / entry["name"]
        assert destination.name == entry["name"] and destination.is_absolute()
        assert digest(source) == expected[entry["name"]]
        assert digest(Path(baseline["backup"]) / entry["name"]) == entry["sha256"].lower()
        assert digest(destination) in {entry["sha256"].lower(), expected[entry["name"]]}, "Concurrent source change"
        payloads.append((entry, destination, source.read_bytes()))
    for entry, destination, payload in payloads:
        if destination.read_bytes() != payload:
            destination.write_bytes(payload)
        assert digest(destination) == expected[entry["name"]]
    for entry in baseline["protected"]:
        assert digest(Path(entry["path"])) == entry["sha256"].lower()
    report = {
        "files": [{"path": entry["path"], "sha256": digest(destination),
                   "changed_from_baseline": expected[entry["name"]] != entry["sha256"].lower()}
                  for entry, destination, _ in payloads],
        "backup": baseline["backup"],
        "protected_files_unchanged": len(baseline["protected"]),
        "repeat_build_and_case_reports_byte_identical": True,
        "static_scenarios_passed": json.loads((work / "cases_a.json").read_text(encoding="utf-8"))["case_count"],
        "changed_workbook_cells": {"WebSearch (1) V2.xlsx": [["מידע כללי", "F2"]]},
        "workbook_sheets_styles_and_other_cells_unchanged": True,
        "visual_preview": "Existing import-format Artifact Tool limitation; package preservation verified",
        "live_agent_tested": False,
        "imported_into_live_system": False,
    }
    encoded = json.dumps(report, ensure_ascii=False, sort_keys=True, indent=2)
    report_path = work / "deployment.json"
    if report_path.exists():
        assert report_path.read_text(encoding="utf-8") == encoded
    report_path.write_text(encoded, encoding="utf-8")
    print(encoded)


def main():
    parser = ArgumentParser()
    parser.add_argument("--work", type=Path, required=True)
    args = parser.parse_args()
    run(args.work.resolve())


if __name__ == "__main__":
    main()
