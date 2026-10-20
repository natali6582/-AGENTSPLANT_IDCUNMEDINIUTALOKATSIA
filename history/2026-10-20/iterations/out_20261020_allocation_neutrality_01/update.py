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


def approved_text():
    return ('אל תציין בשאילתת האימות תוצאה צפויה או מסקנה שנקבעה מראש, כגון expected none. '
            'כאשר CandidateList=[] תאר זאת כנתון הקלט, בלי להסיק שאין פרסומים מתאימים מחוץ לרשימה.')


def prepare(root, work, backup):
    targets = [root / 'פרומפט עדכון 12_57 20_10.txt',
               root / 'קבצים בפורמט שהמערכת מקבלת' / 'WebSearch (1) V2.xlsx']
    protected = [root / 'GetAllocationPolicy.xlsx', root / 'WebSearch.xlsx',
                 root / 'הנחיות כלליות לאחר שינוי גרסה.docx',
                 root / 'קבצים בפורמט שהמערכת מקבלת' / 'GetAllocationPolicy.xlsx',
                 root / 'קבצים בפורמט שהמערכת מקבלת' / 'WebSearch (1).xlsx',
                 root / 'קבצים בפורמט שהמערכת מקבלת' / 'GetAllocationPolicy V2.xlsx']
    assert root.is_absolute() and root.is_dir()
    assert backup.parent == root / 'ישן'
    assert backup.parent.is_dir()
    snapshots = {path: path.read_bytes() for path in targets + protected}
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
    addition = approved_text()
    reports = []
    for entry in baseline['targets']:
        source = (work / 'inputs' / entry['name']).read_bytes()
        assert digest(source) == entry['sha256']
        if entry['name'].endswith('.txt'):
            before = source.decode('utf-8-sig')
            anchor = 'בקש לברר באופן ניטרלי אם המסמך קיים באתר הרשמי ואם תוכנו ומועד פרסומו עומדים בתנאים, תוך ניסיון לפסול כל מועמד.'
            assert before.count(anchor) == 1 and addition not in before
            after = before.replace(anchor, anchor + ' ' + addition, 1)
            assert after.replace(' ' + addition, '', 1) == before
            assert after.count(addition) == 1
            assert before.split('סדר העבודה ותנאי המשך', 1)[1] == after.split('סדר העבודה ותנאי המשך', 1)[1]
            result = after.encode('utf-8')
            if source.startswith(b'\xef\xbb\xbf'):
                result = b'\xef\xbb\xbf' + result
            assert result.replace((' ' + addition).encode('utf-8'), b'', 1) == source
            difference = ''.join(unified_diff(before.splitlines(True), after.splitlines(True),
                                              fromfile='before_prompt', tofile='after_prompt'))
            (output / 'prompt.diff').write_text(difference, encoding='utf-8')
            checks = {'only_approved_text_inserted': True, 'approved_text_occurrences': 1,
                      'budget_recovery_verification_calculation_sections_byte_unchanged': True,
                      'encoding_and_line_endings_preserved': True}
        else:
            assert entry['name'] == 'WebSearch (1) V2.xlsx'
            cells = workbook_model(source)[0]
            before = cells[('מידע כללי', 'F2')]
            anchor = 'באימות מחקר של סוכן זה, פרטי המועמד והציטוט שנשלחו בשאילתה הם טענות לבדיקה, ולא ראיות שכבר אומתו. חזרה עליהם בתשובה אינה אישור לקבלת המסמך.'
            assert before.count(anchor) == 1 and addition not in before
            after = before.replace(anchor, anchor + ' ' + addition, 1)
            assert after.replace(' ' + addition, '', 1) == before
            result, checks = patch_cell(source, ('מידע כללי', 'F2'), after)
            assert checks['changed_package_parts'] == ['xl/sharedStrings.xml']
            assert workbook_model(result)[0][('מידע כללי', 'F2')].count(addition) == 1
            checks['only_approved_text_inserted'] = True
            checks['approved_text_occurrences'] = 1
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
              'only_approved_addition_in_each_target': True,
              'protected_files_unchanged': len(baseline['protected']),
              'repeat_builds_byte_identical': True,
              'changed_workbook_cells': [['מידע כללי', 'F2']],
              'workbook_sheet_count': 60,
              'budget_recovery_and_verification_rules_unchanged': True,
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
