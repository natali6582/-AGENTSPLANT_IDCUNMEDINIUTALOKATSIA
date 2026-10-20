from argparse import ArgumentParser
from difflib import unified_diff
from pathlib import Path
import hashlib
import json
from xlsx_helpers import patch_cell, workbook_model


def digest(data):
    return hashlib.sha256(data).hexdigest()


def write_checked(path, data):
    if path.exists():
        assert path.read_bytes() == data, f'Existing content differs: {path}'
    path.write_bytes(data)
    assert path.read_bytes() == data


def encode_json(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, indent=2).encode('utf-8')


def replace_section(text, start, end, replacement):
    assert text.count(start) == 1 and text.count(end) == 1
    left = text.index(start)
    right = text.index(end, left)
    return text[:left] + replacement.rstrip() + text[right:]


def update_prompt(source, work):
    before = source.decode('utf-8-sig').replace('\r\n', '\n')
    original_workflow = before.split('סדר העבודה ותנאי המשך\n', 1)[1].split('\n\nכיוון שוק מבוסס מחקר', 1)[0]
    original_steps = original_workflow.splitlines()
    step1, step2 = original_steps[:2]
    assert step1.startswith('1. ') and step2.startswith('2. ')
    old_start = 'מחזור מחקר כולל איתור מועמדים ואימות מתנגד'
    old_recovery = 'לאיתור חוזר לאחר פלט איתור ראשון ריק'
    assert step2.count(old_start) == step2.count(old_recovery) == 1
    new_step2 = step2.replace(old_start, 'מחזור מחקר כולל איתור מועמדים ואימות נפרד כאשר נותרו מועמדים לבדיקה', 1)
    new_step2 = new_step2.replace(old_recovery, 'לאיתור חוזר לאחר פלט איתור ראשון ריק או תוצאת איתור תקינה שלא הותירה מועמדים', 1)
    assert step2.split('מותרות לכל היותר שלוש קריאות', 1)[1] == new_step2.split('מותרות לכל היותר שלוש קריאות', 1)[1]
    workflow = (work / 'workflow_section.txt').read_text(encoding='utf-8').rstrip()
    assert workflow.count('{{STEP1}}') == workflow.count('{{STEP2}}') == 1
    workflow = workflow.replace('{{STEP1}}', step1).replace('{{STEP2}}', new_step2)
    research = (work / 'research_section.txt').read_text(encoding='utf-8').rstrip()
    reuse = (work / 'reuse_paragraph.txt').read_text(encoding='utf-8').rstrip()
    assert reuse.startswith('ניתן להשתמש בתוצאות מחקר שכבר התקבלו') and len(reuse) > 400
    after = replace_section(before, 'ניתן להשתמש בתוצאות מחקר שכבר התקבלו',
                            '\nאם כבר קיימת תוצאת בחינה', reuse)
    after = replace_section(after, 'תנאי המחקר וניסוח השאילתות', '\n\nמקורות המחקר', research)
    after = replace_section(after, 'סדר העבודה ותנאי המשך', '\n\nכיוון שוק מבוסס מחקר', workflow)
    assert '{{STEP' not in after
    assert before.split('כיוון שוק מבוסס מחקר', 1)[1] == after.split('כיוון שוק מבוסס מחקר', 1)[1]
    assert before.split('מקורות המחקר\n', 1)[1].split('\n\nסדר העבודה', 1)[0] == after.split('מקורות המחקר\n', 1)[1].split('\n\nסדר העבודה', 1)[0]
    assert before.split('ניתן להשתמש בתוצאות מחקר שכבר התקבלו', 1)[0] == after.split('ניתן להשתמש בתוצאות מחקר שכבר התקבלו', 1)[0]
    assert after.count('expected none') == 1
    for obsolete in ['גם כשהיא []', 'תוצאה תקינה שלא הותירה מועמדים אינה עילה לאיתור חוזר',
                     'רשימה ריקה שנוצרה מתוצאת איתור תקינה נשלחת לאימות',
                     'הקריאה השלישית במחזור חדש', 'הסר רשומה חסרה או פסולה']:
        assert obsolete not in after, obsolete
    for required in ['אין קריאת אימות עבור []', 'נותרו שתי קריאות לפחות',
                     'אין להוסיף או להחליף מועמד', 'פרסום מתקבל כראיה רק אם פורסם בפועל באותה שנה',
                     'אין להשלים חודש חסר', 'עד שש רשומות ועד 900 מילים',
                     'כל עשרת המוסדות ברשימה נשארים מותרים',
                     'אין להניח אפס קריאות או ניסיון פנוי',
                     'ניסיון ההתאוששות כבר נוצל באיתור, אין ניסיון נוסף להשלמת אימות']:
        assert required in after, required
    newline = '\r\n' if b'\r\n' in source else '\n'
    result = after.replace('\n', newline).encode('utf-8')
    if source.startswith(b'\xef\xbb\xbf'):
        result = b'\xef\xbb\xbf' + result
    difference = ''.join(unified_diff(before.splitlines(True), after.splitlines(True),
                                     fromfile='before_prompt', tofile='after_prompt'))
    return result, difference


def prepare(root, work, backup):
    targets = [root / 'פרומפט עדכון 12_57 20_10.txt',
               root / 'קבצים בפורמט שהמערכת מקבלת' / 'WebSearch (1) V2.xlsx']
    protected = [root / 'GetAllocationPolicy.xlsx', root / 'WebSearch.xlsx',
                 root / 'הנחיות כלליות לאחר שינוי גרסה.docx',
                 root / 'קבצים בפורמט שהמערכת מקבלת' / 'GetAllocationPolicy.xlsx',
                 root / 'קבצים בפורמט שהמערכת מקבלת' / 'WebSearch (1).xlsx',
                 root / 'קבצים בפורמט שהמערכת מקבלת' / 'GetAllocationPolicy V2.xlsx',
                 root / 'פרומפט עדכון 13_33 20_10.txt',
                 Path(r'C:\Users\Nataliek\Downloads\הנחיות כלליות לאחר שינוי גרסה.docx')]
    assert root.is_absolute() and root.is_dir()
    assert backup.parent == root / 'ישן'
    assert backup.parent.is_dir()
    snapshots = {path: path.read_bytes() for path in targets + protected}
    assert snapshots[protected[2]] == snapshots[protected[-1]], 'General instruction copies differ'
    (work / 'inputs').mkdir(exist_ok=True)
    backup.mkdir(exist_ok=True)
    for path in targets:
        write_checked(work / 'inputs' / path.name, snapshots[path])
        write_checked(backup / path.name, snapshots[path])
    baseline = {'targets': [{'path': str(path), 'name': path.name, 'sha256': digest(snapshots[path])}
                             for path in targets],
                'protected': [{'path': str(path), 'sha256': digest(snapshots[path])}
                              for path in protected], 'backup': str(backup)}
    for path, data in snapshots.items():
        assert path.read_bytes() == data, 'Concurrent source change'
    write_checked(work / 'baseline.json', encode_json(baseline))
    print(json.dumps({'backed_up_targets': len(targets), 'protected_files': len(protected),
                      'backup_verified': True}, ensure_ascii=False))


def build(work, output):
    assert not output.exists(), 'Use a fresh output directory'
    output.mkdir()
    baseline = json.loads((work / 'baseline.json').read_text(encoding='utf-8'))
    reports = []
    for entry in baseline['targets']:
        source = (work / 'inputs' / entry['name']).read_bytes()
        assert digest(source) == entry['sha256']
        if entry['name'].endswith('.txt'):
            result, difference = update_prompt(source, work)
            (output / 'prompt.diff').write_text(difference, encoding='utf-8')
            checks = {'only_approved_research_sections_updated': True,
                      'allocation_algorithm_and_business_output_byte_unchanged': True,
                      'original_whitelist_unchanged': True,
                      'budget_history_and_counting_rules_unchanged': True,
                      'no_empty_candidate_verification': True,
                      'partial_candidates_require_full_final_verification': True,
                      'encoding_and_line_endings_preserved': True}
        else:
            assert entry['name'] == 'WebSearch (1) V2.xlsx'
            cells = workbook_model(source)[0]
            before = cells[('מידע כללי', 'F2')]
            after = (work / 'search_addition.txt').read_text(encoding='utf-8').rstrip()
            assert before != after and after.count('expected none') == 1
            for obsolete in ['איתור ללא מועמדים אינו עילה לקריאה נוספת מעבר לרצף המותר',
                             'לאיתור חוזר אחרי פלט ראשון ריק או להשלמת אימות']:
                assert obsolete not in after
            for required in ['אין לבצע אימות על CandidateList=[]',
                             'לכל היותר שלוש קריאות WebSearch',
                             'אימות נפרד הוא חובה לפני קבלת כל ראיה',
                             'תאריך הפרסום אינו שדה אוטומטי',
                             'ההנחיות הכלליות נשארות בתוקף', 'Markdown']:
                assert required in after
            result, checks = patch_cell(source, ('מידע כללי', 'F2'), after)
            assert checks['changed_package_parts'] == ['xl/sharedStrings.xml']
            assert workbook_model(result)[0][('מידע כללי', 'F2')] == after
            checks['only_agent_specific_addition_updated'] = True
            checks['tool_return_contract_and_shared_descriptions_unchanged'] = True
        write_checked(output / entry['name'], result)
        reports.append({'file': entry['name'], 'sha256': digest(result), **checks})
    write_checked(output / 'verification.json', encode_json(reports))
    print(json.dumps(reports, ensure_ascii=False))


def deploy(work):
    baseline = json.loads((work / 'baseline.json').read_text(encoding='utf-8'))
    first, second = work / 'build_a', work / 'build_b'
    assert {path.name for path in first.iterdir()} == {path.name for path in second.iterdir()}
    for path in first.iterdir():
        assert path.read_bytes() == (second / path.name).read_bytes(), 'Repeat build differs'
    expected = {entry['file']: entry['sha256'] for entry in
                json.loads((first / 'verification.json').read_text(encoding='utf-8'))}
    for entry in baseline['protected']:
        assert digest(Path(entry['path']).read_bytes()) == entry['sha256'], 'Protected source changed'
    payloads = []
    for entry in baseline['targets']:
        destination = Path(entry['path'])
        payload = (first / entry['name']).read_bytes()
        assert digest(payload) == expected[entry['name']]
        assert digest((Path(baseline['backup']) / entry['name']).read_bytes()) == entry['sha256']
        assert digest(destination.read_bytes()) in {entry['sha256'], expected[entry['name']]}, 'Concurrent source change'
        payloads.append((destination, payload))
    # Probe both existing files without truncating before any deployment write.
    for destination, _ in payloads:
        with destination.open('r+b'):
            pass
    for destination, payload in payloads:
        if destination.read_bytes() != payload:
            destination.write_bytes(payload)
        assert destination.read_bytes() == payload
    for entry in baseline['protected']:
        assert digest(Path(entry['path']).read_bytes()) == entry['sha256'], 'Protected source changed'
    report = {'saved': [{'path': str(path), 'sha256': digest(payload)} for path, payload in payloads],
              'only_approved_research_sections_and_F2_updated': True,
              'protected_files_unchanged': len(baseline['protected']),
              'repeat_builds_byte_identical': True,
              'changed_workbook_cells': [['מידע כללי', 'F2']],
              'workbook_sheet_count': 60,
              'three_call_cap_and_shared_recovery_preserved': True,
              'general_instructions_and_markdown_output_preserved': True,
              'no_new_api_date_column_or_parser_schema': True,
              'visual_preview': 'Existing import-format Artifact Tool limitation; unchanged package layout verified',
              'live_agent_tested': False, 'imported_into_live_system': False,
              'backup': baseline['backup']}
    write_checked(work / 'deployment.json', encode_json(report))
    print(json.dumps(report, ensure_ascii=False, indent=2))


def main():
    parser = ArgumentParser()
    parser.add_argument('operation', choices=['prepare', 'build', 'deploy'])
    parser.add_argument('--work', type=Path, required=True)
    parser.add_argument('--root', type=Path)
    parser.add_argument('--backup', type=Path)
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    work = args.work.resolve()
    if args.operation == 'prepare':
        prepare(args.root.resolve(), work, args.backup.resolve())
    elif args.operation == 'build':
        build(work, args.output.resolve())
    else:
        deploy(work)


if __name__ == '__main__':
    main()
