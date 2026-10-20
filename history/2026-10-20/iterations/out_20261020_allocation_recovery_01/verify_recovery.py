from argparse import ArgumentParser
from pathlib import Path
import json
from xlsx_helpers import workbook_model


def trace(events, prior_calls=0, recovery_used=False):
    """Independent reference cases for the written contract, not an agent runner."""
    calls, stage, actions = prior_calls, "discovery", []
    for event in events:
        assert calls < 3, "Fourth search call is forbidden"
        actions.append(stage)
        calls += 1
        if stage in {"discovery", "discovery_recovery"}:
            if event == "empty":
                if stage == "discovery" and not recovery_used and 3 - calls >= 2:
                    stage, recovery_used = "discovery_recovery", True
                    continue
                return {"actions": actions, "status": "incomplete", "calls": calls}
            if event in {"candidates", "valid_no_candidates"}:
                stage = "verification"
                continue
        elif stage in {"verification", "verification_recovery"}:
            if event == "verified":
                return {"actions": actions, "status": "verified", "calls": calls}
            if event == "rejected_or_no_candidates":
                return {"actions": actions, "status": "completed_without_evidence", "calls": calls}
            if event in {"empty", "source_gap"}:
                if not recovery_used and calls < 3:
                    stage, recovery_used = "verification_recovery", True
                    continue
                return {"actions": actions, "status": "incomplete", "calls": calls}
        if event == "explicit_failure":
            return {"actions": actions, "status": "incomplete", "calls": calls}
        raise AssertionError((stage, event))
    return {"actions": actions, "status": "pending_" + stage, "calls": calls}


def run(work, build):
    cases = [
        ("empty_discovery_recovers_then_verifies", ["empty", "candidates", "verified"], 0, False,
         ["discovery", "discovery_recovery", "verification"], "verified"),
        ("second_empty_discovery_stops", ["empty", "empty"], 0, False,
         ["discovery", "discovery_recovery"], "incomplete"),
        ("recovered_discovery_is_not_verification", ["empty", "candidates"], 0, False,
         ["discovery", "discovery_recovery"], "pending_verification"),
        ("verification_gap_after_discovery_recovery_cannot_retry", ["empty", "candidates", "source_gap"], 0, False,
         ["discovery", "discovery_recovery", "verification"], "incomplete"),
        ("verification_recovery_still_available_after_good_discovery", ["candidates", "source_gap", "verified"], 0, False,
         ["discovery", "verification", "verification_recovery"], "verified"),
        ("empty_verification_uses_shared_recovery", ["candidates", "empty", "verified"], 0, False,
         ["discovery", "verification", "verification_recovery"], "verified"),
        ("valid_empty_candidate_list_does_not_restart_discovery", ["valid_no_candidates", "rejected_or_no_candidates"], 0, False,
         ["discovery", "verification"], "completed_without_evidence"),
        ("recovery_does_not_repeat", ["candidates", "source_gap", "source_gap"], 0, False,
         ["discovery", "verification", "verification_recovery"], "incomplete"),
        ("explicit_failure_does_not_gain_new_retry_permission", ["explicit_failure"], 0, False,
         ["discovery"], "incomplete"),
        ("budget_must_reserve_verification", ["empty"], 1, False,
         ["discovery"], "incomplete"),
        ("used_recovery_is_not_reset", ["empty"], 0, True,
         ["discovery"], "incomplete"),
    ]
    results = []
    for name, events, prior_calls, used, actions, status in cases:
        result = trace(events, prior_calls, used)
        assert result["actions"] == actions and result["status"] == status, (name, result)
        assert result["calls"] <= 3
        if status == "verified":
            assert result["actions"][-1] in {"verification", "verification_recovery"}
        results.append({"name": name, **result})
    prompt_name = "פרומפט מלא- תוספת להנחיות הכלליות 20_10.txt"
    prompt = (build / prompt_name).read_text(encoding="utf-8-sig")
    for clause in ["ניסיון התאוששות אחד בלבד המשותף לכל המחזור", "לפחות", "שמור את הקריאה הנותרת לאימות",
                   "אם גם האיתור החוזר ריק", "כולל אימות נפרד גם לאחר איתור חוזר",
                   "אין ניסיון נוסף להשלמת אימות", "קריאת השלמה רביעית",
                   "תוצאה תקינה שלא הותירה מועמדים אינה עילה לאיתור חוזר"]:
        assert clause in prompt
    search = (build / "WebSearch (1) V2.xlsx").read_bytes()
    search_text = workbook_model(search)[0][("מידע כללי", "F2")]
    assert "שלוש קריאות WebSearch" in search_text and "ניסיון ההתאוששות היחיד משותף" in search_text
    assert "איתור חוזר אינו אימות" in search_text
    policy_name = "GetAllocationPolicy V2.xlsx"
    assert (build / policy_name).read_bytes() == (work / "inputs" / policy_name).read_bytes()
    actual_log = (work / "inputs" / "conversation.txt").read_text(encoding="utf-8-sig").replace("\r\n", "\n")
    assert "הפלט שהתקבל:\n(ריק)" in actual_log
    assert actual_log.count("שם הכלי: WebSearch") == 1
    return {"mode": "Static prompt checks and independent reference scenarios; no live agent execution",
            "cases": results, "case_count": len(results), "latest_log_contains_empty_first_search": True,
            "allocation_policy_v2_unchanged": True, "live_agent_tested": False}


def main():
    parser = ArgumentParser()
    parser.add_argument("--work", type=Path, required=True)
    parser.add_argument("--build", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    report = run(args.work.resolve(), args.build.resolve())
    args.output.write_text(json.dumps(report, ensure_ascii=False, sort_keys=True, indent=2), encoding="utf-8")
    print(json.dumps({"case_count": report["case_count"], "all_passed": True, "live_agent_tested": False}))


if __name__ == "__main__":
    main()
