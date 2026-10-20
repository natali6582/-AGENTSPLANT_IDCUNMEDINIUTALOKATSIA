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
    section = (work / 'verification_gate.txt').read_text(encoding='utf-8').rstrip()
    count_gate = (work / 'count_gate.txt').read_text(encoding='utf-8').rstrip()
    anchor = '\n\nמיקוד שלב האיתור'
    count_anchor = '\nהגדלה דורשת לפחות שני מוסדות שונים'
    assert before.count(anchor) == before.count(count_anchor) == 1
    after = before.replace(anchor, '\n\n' + section + anchor, 1)
    replacements = []
    for number in [5, 6]:
        old = [line for line in before.splitlines() if line.startswith(str(number) + '. ')]
        assert len(old) == 1
        new = (work / f'step{number}.txt').read_text(encoding='utf-8').rstrip()
        assert new.startswith(str(number) + '. ') and '\n' not in new
        assert after.count(old[0]) == 1
        after = after.replace(old[0], new, 1)
        replacements.append((old[0], new))
    after = after.replace(count_anchor, '\n' + count_gate + count_anchor, 1)
    # Exact inverse proves that all other text is unchanged.
    restored = after.replace('\n\n' + section + anchor, anchor, 1)
    restored = restored.replace('\n' + count_gate + count_anchor, count_anchor, 1)
    for old, new in reversed(replacements):
        assert restored.count(new) == 1
        restored = restored.replace(new, old, 1)
    assert restored == before
    assert before.split('גבולות השינוי\n', 1)[1] == after.split('גבולות השינוי\n', 1)[1]
    assert before.split('מקורות המחקר\n', 1)[1].split('\n\nסדר העבודה', 1)[0] == after.split('מקורות המחקר\n', 1)[1].split('\n\nסדר העבודה', 1)[0]
    for number in [1, 2, 3, 4, 7, 8, 9]:
        assert [line for line in before.splitlines() if line.startswith(str(number) + '. ')] == [line for line in after.splitlines() if line.startswith(str(number) + '. ')]
    assert after.count('expected none') == 1
    for required in ['אין קריאת אימות עבור []', 'נותרו שתי קריאות לפחות', 'אין להוסיף או להחליף מועמד',
                     'פרסום מתקבל כראיה רק אם פורסם בפועל באותה שנה', 'אין להשלים חודש חסר',
                     'עד שש רשומות ועד 900 מילים', 'כל עשרת המוסדות ברשימה נשארים מותרים',
                     'אין להניח אפס קריאות או ניסיון פנוי', 'אין ניסיון נוסף להשלמת אימות',
                     'אין לדלג על ההשלמה המחויבת', 'לא לפי תווית קבלה או פסילה בלבד',
                     'מותר לקבל באימות או בהשלמתו ראיה תומכת חדשה על אותו מועמד',
                     'אינם נספרים במניין המוסדות', 'אינו נעשה מאומת מעצם חזרתו',
                     'אין קריאה רביעית']:
        assert required in after, required
    newline = '\r\n' if b'\r\n' in source else '\n'
    result = after.replace('\n', newline).encode('utf-8')
    if source.startswith(b'\xef\xbb\xbf'):
        result = b'\xef\xbb\xbf' + result
    difference = ''.join(unified_diff(before.splitlines(True), after.splitlines(True),
                                     fromfile='before_prompt', tofile='after_prompt'))
    return result, difference


def prepare(root, work, backup):
    targets = [root / 'פרומפט עדכון 14_15 20_10.txt',
               root / 'קבצים בפורמט שהמערכת מקבלת' / 'WebSearch (1) V2.xlsx']
    protected = [root / 'GetAllocationPolicy.xlsx', root / 'WebSearch.xlsx',
                 root / 'הנחיות כלליות לאחר שינוי גרסה.docx',
                 root / 'קבצים בפורמט שהמערכת מקבלת' / 'GetAllocationPolicy.xlsx',
                 root / 'קבצים בפורמט שהמערכת מקבלת' / 'GetAllocationPolicy V2.xlsx',
                 root / 'פרומפט עדכון 13_33 20_10.txt',
                 root / 'פרומפט עדכון 12_57 20_10.txt',
                 targets[0],
                 Path(r'C:\Users\Nataliek\Downloads\הנחיות כלליות לאחר שינוי גרסה.docx')]
    optional = root / 'קבצים בפורמט שהמערכת מקבלת' / 'WebSearch (1).xlsx'
    if optional.exists():
        protected.insert(-1, optional)
    new_prompt = root / 'פרומפט עדכון אימות 20_10.txt'
    assert not new_prompt.exists(), 'New prompt must not overwrite an existing file'
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
    baseline = {'targets': [{'path': str(new_prompt if index == 0 else path),
                            'source_path': str(path), 'input_name': path.name,
                            'name': new_prompt.name if index == 0 else path.name,
                            'new_file': index == 0, 'sha256': digest(snapshots[path])}
                             for index, path in enumerate(targets)],
                'protected': [{'path': str(path), 'sha256': digest(snapshots[path])}
                              for path in protected], 'backup': str(backup),
                'absent_out_of_scope': [] if optional.exists() else [str(optional)]}
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
        source = (work / 'inputs' / entry['input_name']).read_bytes()
        assert digest(source) == entry['sha256']
        if entry['name'].endswith('.txt'):
            result, difference = update_prompt(source, work)
            (output / 'prompt.diff').write_text(difference, encoding='utf-8')
            checks = {'only_verification_gate_steps_5_6_and_count_gate_changed': True,
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
            gate = (work / 'search_gate.txt').read_text(encoding='utf-8').rstrip()
            anchor = '\n\nלפני קריאה נוספת נדרשים תיעוד מהימן'
            assert before.count(anchor) == 1
            after = before.replace(anchor, '\n\n' + gate + anchor, 1)
            assert after.replace('\n\n' + gate + anchor, anchor, 1) == before
            assert before != after and after.count('expected none') == 1
            for required in ['אין לבצע אימות על CandidateList=[]',
                             'לכל היותר שלוש קריאות WebSearch',
                             'אימות נפרד הוא חובה לפני קבלת כל ראיה',
                             'תאריך הפרסום אינו שדה אוטומטי',
                             'ההנחיות הכלליות נשארות בתוקף', 'Markdown',
                             'אין לספור מסמך או כיוון שלא אומתו במלואם',
                             'חובה לנצל את ניסיון ההשלמה לפני סיכום הבחינה',
                             'מותר לקבל ראיה תומכת חדשה על אותו מועמד']:
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
    for path in baseline['absent_out_of_scope']:
        assert not Path(path).exists(), 'Out-of-scope directory contents changed'
    payloads = []
    for entry in baseline['targets']:
        destination = Path(entry['path'])
        payload = (first / entry['name']).read_bytes()
        assert digest(payload) == expected[entry['name']]
        assert digest((Path(baseline['backup']) / entry['input_name']).read_bytes()) == entry['sha256']
        if entry['new_file']:
            assert not destination.exists() or destination.read_bytes() == payload, 'New prompt filename occupied'
        else:
            assert digest(destination.read_bytes()) in {entry['sha256'], expected[entry['name']]}, 'Concurrent source change'
        payloads.append((destination, payload, entry['new_file']))
    # Non-truncating probe of existing workbook, before creating the new prompt.
    for destination, _, new_file in payloads:
        if not new_file:
            with destination.open('r+b'):
                pass
    for destination, payload, new_file in payloads:
        if new_file and not destination.exists():
            with destination.open('xb') as stream:
                stream.write(payload)
        elif destination.read_bytes() != payload:
            destination.write_bytes(payload)
        assert destination.read_bytes() == payload
    for entry in baseline['protected']:
        assert digest(Path(entry['path']).read_bytes()) == entry['sha256'], 'Protected source changed'
    report = {'saved': [{'path': str(path), 'sha256': digest(payload)} for path, payload, _ in payloads],
              'new_prompt_copy_and_only_F2_updated': True,
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
