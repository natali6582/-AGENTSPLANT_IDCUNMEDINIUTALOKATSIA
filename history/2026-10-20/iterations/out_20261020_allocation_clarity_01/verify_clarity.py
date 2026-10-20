from argparse import ArgumentParser
from pathlib import Path
import json
import re
from xlsx_helpers import workbook_model


def recovery_decision(count, used, trigger, stage_known=True, linked=True):
    if type(count) is not int or not 0 <= count <= 3 or type(used) is not bool or not stage_known or not linked:
        return 'no_calls_unknown_state'
    if trigger in {'valid_empty_candidates', 'supported_rejection', 'below_evidence_threshold', 'explicit_failure'}:
        return 'no_recovery_trigger'
    if used:
        return 'recovery_already_used'
    required = 2 if trigger == 'discovery_empty' else 1 if trigger == 'verification_gap' else None
    assert required is not None
    return 'recovery_required' if 3 - count >= required else 'insufficient_budget'


def workflow(text):
    section = text.split('סדר העבודה ותנאי המשך\n', 1)[1].split('\n\nכיוון שוק מבוסס מחקר', 1)[0]
    return {int(line.split('.', 1)[0]): line for line in section.splitlines() if re.match(r'^\d+\. ', line)}


def run(work, build):
    prompt_name = 'פרומפט עדכון 12_57 20_10.txt'
    before = (work / 'inputs' / prompt_name).read_text(encoding='utf-8-sig')
    after = (build / prompt_name).read_text(encoding='utf-8-sig')
    before_rules, after_rules = workflow(before), workflow(after)
    assert before_rules.keys() == after_rules.keys() == set(range(1, 10))
    assert [key for key in before_rules if before_rules[key] != after_rules[key]] == [2, 6]
    assert 'ניתן לנצל אותו להשלמת אימות' not in after
    for clause in ['חובה לבצע ניסיון אחד להשלמת האימות', 'אין לוותר עליו לפי שיקול דעת',
                   'אין להניח אפס קריאות או ניסיון פנוי', 'אפשר לשחזר נתונים אלה מתיעוד מספק בשיחה',
                   'אל תבצע קריאות WebSearch נוספות באותו מחזור',
                   'עצם קיומם של מסמכים מאומתים אינו מוכיח שהרצף הושלם',
                   'נותרו שתי קריאות לפחות', 'לכל היותר שלוש קריאות WebSearch',
                   'כולל אימות נפרד גם לאחר איתור חוזר', 'אין לבצע אחריו קריאת השלמה רביעית']:
        assert clause in after, clause
    assert before.split('\n\nכיוון שוק מבוסס מחקר', 1)[1] == after.split('\n\nכיוון שוק מבוסס מחקר', 1)[1]
    cases = [
        ('empty_discovery_requires_retry_and_verification', (1, False, 'discovery_empty'), 'recovery_required'),
        ('verification_gap_requires_recovery', (2, False, 'verification_gap'), 'recovery_required'),
        ('one_slot_cannot_recover_discovery', (2, False, 'discovery_empty'), 'insufficient_budget'),
        ('no_fourth_call', (3, False, 'verification_gap'), 'insufficient_budget'),
        ('shared_recovery_cannot_be_reused', (2, True, 'verification_gap'), 'recovery_already_used'),
        ('unknown_count_is_not_zero', (None, False, 'discovery_empty'), 'no_calls_unknown_state'),
        ('unknown_recovery_is_not_unused', (2, None, 'verification_gap'), 'no_calls_unknown_state'),
        ('unknown_stage_blocks_new_calls', (2, False, 'verification_gap', False), 'no_calls_unknown_state'),
        ('unlinked_results_block_new_calls', (2, False, 'verification_gap', True, False), 'no_calls_unknown_state'),
        ('conflicting_counter_blocks_new_calls', (-1, False, 'verification_gap'), 'no_calls_unknown_state'),
        ('valid_empty_list_does_not_trigger_recovery', (2, False, 'valid_empty_candidates'), 'no_recovery_trigger'),
        ('supported_rejection_does_not_trigger_recovery', (2, False, 'supported_rejection'), 'no_recovery_trigger'),
        ('low_support_does_not_trigger_recovery', (2, False, 'below_evidence_threshold'), 'no_recovery_trigger'),
        ('explicit_failure_does_not_add_retry', (2, False, 'explicit_failure'), 'no_recovery_trigger'),
    ]
    report_cases = []
    for name, inputs, expected in cases:
        result = recovery_decision(*inputs)
        assert result == expected, (name, result)
        report_cases.append({'name': name, 'result': result})
    search_text = workbook_model((build / 'WebSearch (1) V2.xlsx').read_bytes())[0][('מידע כללי', 'F2')]
    for clause in ['חובה לבצע את הניסיון', 'זו אינה רשות', 'אין להניח אפס או ניסיון פנוי',
                   'לכל היותר שלוש קריאות WebSearch', 'לאיתור חוזר חייב להישאר תקציב גם לאימות נפרד']:
        assert clause in search_text, clause
    assert (build / 'GetAllocationPolicy V2.xlsx').read_bytes() == (work / 'inputs' / 'GetAllocationPolicy V2.xlsx').read_bytes()
    return {'case_count': len(cases), 'cases': report_cases, 'workflow_items_changed': [2, 6],
            'research_acceptance_and_allocation_algorithm_unchanged': True,
            'scope': 'Static wording, package checks, and independent reference cases; no live agent execution',
            'live_agent_tested': False}


def main():
    parser = ArgumentParser()
    parser.add_argument('--work', type=Path, required=True)
    parser.add_argument('--build', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    report = run(args.work.resolve(), args.build.resolve())
    args.output.write_text(json.dumps(report, ensure_ascii=False, sort_keys=True, indent=2), encoding='utf-8')
    print(json.dumps({'case_count': report['case_count'], 'all_passed': True, 'live_agent_tested': False}))


if __name__ == '__main__':
    main()
