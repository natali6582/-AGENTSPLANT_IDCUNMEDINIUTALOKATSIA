from argparse import ArgumentParser
from pathlib import Path
import hashlib
import json


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run(work):
    baseline = json.loads((work / "baseline.json").read_text(encoding="utf-8-sig"))
    verification = json.loads((work / "build_final_a" / "verification.json").read_text(encoding="utf-8"))
    expected = {item["file"]: item["sha256"] for item in verification}
    expected_names = {"פרומפט מלא- תוספת להנחיות הכלליות 20_10.txt",
                      "GetAllocationPolicy V2.xlsx", "WebSearch (1) V2.xlsx"}
    assert set(expected) == expected_names
    assert {item["name"] for item in baseline["targets"]} == expected_names
    case_bytes = (work / "cases_a.json").read_bytes()
    assert case_bytes == (work / "cases_b.json").read_bytes()
    for item in (work / "build_final_a").iterdir():
        assert item.read_bytes() == (work / "build_final_b" / item.name).read_bytes()
    for item in baseline["protected"]:
        assert digest(Path(item["path"])) == item["sha256"].lower()
    payloads = []
    for item in baseline["targets"]:
        destination = Path(item["path"]).resolve()
        source = work / "build_final_a" / item["name"]
        assert destination.name == item["name"]
        assert digest(source) == expected[item["name"]]
        assert digest(Path(baseline["backup"]) / item["name"]) == item["sha256"].lower()
        assert digest(destination) in {item["sha256"].lower(), expected[item["name"]]}, "Destination changed since snapshot"
        payloads.append((destination, source.read_bytes()))
    for destination, payload in payloads:
        if destination.read_bytes() != payload:
            destination.write_bytes(payload)
        assert destination.read_bytes() == payload
    for item in baseline["protected"]:
        assert digest(Path(item["path"])) == item["sha256"].lower()
    report = {
        "deployed": [{"path": str(destination), "sha256": digest(destination)} for destination, _ in payloads],
        "protected_files_unchanged": len(baseline["protected"]),
        "backup": baseline["backup"],
        "repeat_build_byte_identical": True,
        "case_reports_byte_identical": True,
        "static_case_count": len(json.loads(case_bytes)["cases_passed"]),
        "workbook_changes": [{"file": item["file"], "cells": item["changed_cells"], "sheets": item["sheet_count"]}
                             for item in verification if "changed_cells" in item],
        "preview": "Original import-format workbook cannot be loaded by Artifact Tool: Arg_NullReferenceException",
        "found_not_fixed": "Existing duplicated AllocationPolicyUpdateAgent@v version prefix is unchanged; runtime effect not established",
        "imported_into_live_system": False,
        "live_agent_tested": False,
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
