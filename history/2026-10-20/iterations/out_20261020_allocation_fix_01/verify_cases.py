from argparse import ArgumentParser
from copy import deepcopy
from decimal import Decimal
from pathlib import Path
import json


def policy_from_log(path):
    text = path.read_text(encoding="utf-8-sig")
    return json.JSONDecoder().raw_decode(text.split("הפלט שהתקבל:", 1)[1].lstrip())[0]


def eligible_profiles(policy):
    assert sorted(p["ProfileCode"] for p in policy["Profiles"]) == [1, 2, 3, 4, 5]
    for profile in policy["Profiles"]:
        allocation = profile["Allocation"]
        assert sorted(item["Component"] for item in allocation) == ["alternative", "cash", "equity", "realEstate"]
        assert all(isinstance(item["Value"], (int, float)) and not isinstance(item["Value"], bool)
                   and 0 <= item["Value"] <= 100 for item in allocation)
        assert sum(item["Value"] for item in allocation) == 100
        for component in ["equity", "alternative", "realEstate", "cashExp"]:
            bounds = [item for item in profile["Exposures"] if item["Component"] == component]
            assert len(bounds) == 1 and 0 <= bounds[0]["Min"] <= bounds[0]["Max"] <= 100
    return [p["ProfileCode"] for p in policy["Profiles"]]


def proposal_errors(policy, code, values, evidence):
    profile = next(p for p in policy["Profiles"] if p["ProfileCode"] == code)
    existing = {v["Component"]: Decimal(str(v["Value"])) for v in profile["Allocation"]}
    bounds = {v["Component"]: v for v in profile["Exposures"]}
    proposed = {key: Decimal(str(value)) for key, value in values.items()}
    errors = []
    if set(proposed) != set(existing) or sum(proposed.values()) != 100:
        return ["allocation_sum_or_components"]
    for component, value in proposed.items():
        limit = bounds["cashExp" if component == "cash" else component]
        if not Decimal(str(limit["Min"])) <= value <= Decimal(str(limit["Max"])):
            errors.append("remaining_violation")
        difference = value - existing[component]
        if difference:
            if abs(difference) < 2:
                errors.append("below_transfer_threshold")
            direction = "increase" if difference > 0 else "decrease"
            if len(set(evidence.get((component, direction), []))) < 2:
                errors.append("insufficient_verified_research")
    combined = {}
    for original in policy["Profiles"]:
        combined[original["ProfileCode"]] = proposed if original["ProfileCode"] == code else {
            v["Component"]: Decimal(str(v["Value"])) for v in original["Allocation"]}
    for first in range(1, 5):
        if combined[first]["equity"] > combined[first + 1]["equity"] or combined[first]["cash"] < combined[first + 1]["cash"]:
            errors.append("profile_order")
    return sorted(set(errors))


def run(first_log, second_log, prompt_path):
    policy = policy_from_log(first_log)
    assert policy == policy_from_log(second_log)
    assert eligible_profiles(policy) == [1, 2, 3, 4, 5]
    incomplete = deepcopy(policy)
    incomplete["Profiles"][0]["Allocation"][0]["Value"] = 44
    try:
        eligible_profiles(incomplete)
    except AssertionError:
        pass
    else:
        raise AssertionError("Malformed allocation was accepted")
    # Synthetic research labels and proposal fixtures; no real market research is asserted.
    evidence = {("equity", "increase"): ["fixture_A", "fixture_B"],
                ("alternative", "decrease"): ["fixture_C", "fixture_D"]}
    valid = {"equity": 53.75, "alternative": 11.25, "realEstate": 20, "cash": 15}
    assert proposal_errors(policy, 1, valid, evidence) == []
    partial = {"equity": 47, "alternative": 18, "realEstate": 20, "cash": 15}
    assert "remaining_violation" in proposal_errors(policy, 1, partial, evidence)
    assert "insufficient_verified_research" in proposal_errors(policy, 1, valid, {})
    one_institution = {key: ["fixture_A", "fixture_A"] for key in evidence}
    assert "insufficient_verified_research" in proposal_errors(policy, 1, valid, one_institution)
    order_violation = {"equity": 57.5, "alternative": 7.5, "realEstate": 20, "cash": 15}
    assert proposal_errors(policy, 1, order_violation, evidence) == ["profile_order"]
    bad_sum = dict(valid, cash=16)
    assert proposal_errors(policy, 1, bad_sum, evidence) == ["allocation_sum_or_components"]
    prompt = prompt_path.read_text(encoding="utf-8-sig")
    for clause in ["גם אם כל חמשת הפרופילים חורגים", "לכל היותר שלוש קריאות WebSearch",
                   "אין קריאת השלמה נוספת", "שימוש חוזר בתוצאות אינו מאפס",
                   "לרבות ערכים שלא שונו", "כבר הוכח שאינו עומד בתנאים"]:
        assert clause in prompt, clause
    for obsolete in ["כללי ההחרגה", "שהוחרג מראש", "אל תוסיף חיפוש שלישי"]:
        assert obsolete not in prompt
    # Evaluate the provided failed verification against the written source requirement.
    log = second_log.read_text(encoding="utf-8-sig")
    assert "יש לאתר את המסמך הספציפי באתר" in log
    assert "הנחיה לאתר את המסמך" in prompt
    return {
        "mode": "Static contract and arithmetic fixtures; not an agent execution",
        "same_policy_in_both_logs": True,
        "existing_out_of_bounds_profiles_eligible_for_examination": [1, 2, 3, 4, 5],
        "cases_passed": [
            "valid_in_bounds_fixture_for_previously_out_of_bounds_profile",
            "partial_correction_rejected",
            "unverified_research_rejected",
            "single_distinct_institution_rejected",
            "profile_order_violation_rejected",
            "incorrect_allocation_sum_rejected",
            "malformed_source_data_rejected",
            "bounded_verification_retry_and_no_automatic_budget_reset",
            "homepage_only_UBS_result_requires_verification_completion",
        ],
        "research_fixture_is_synthetic": True,
        "live_agent_tested": False,
    }


def main():
    parser = ArgumentParser()
    parser.add_argument("--first-log", type=Path, required=True)
    parser.add_argument("--second-log", type=Path, required=True)
    parser.add_argument("--prompt", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    report = run(args.first_log, args.second_log, args.prompt)
    encoded = json.dumps(report, ensure_ascii=False, sort_keys=True, indent=2)
    args.output.write_text(encoded, encoding="utf-8")
    print(encoded)


if __name__ == "__main__":
    main()
