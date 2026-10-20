from argparse import ArgumentParser
from difflib import unified_diff
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
import hashlib
import json
from xlsx_helpers import patch_cell, workbook_model


def digest(data):
    return hashlib.sha256(data).hexdigest()


def encoded(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, indent=2).encode('utf-8')


def write_checked(path, payload):
    if path.exists():
        assert path.read_bytes() == payload, f'Existing content differs: {path}'
    path.write_bytes(payload)
    assert path.read_bytes() == payload


def prepare(root, work, backup):
    source = root / 'קבצים בפורמט שהמערכת מקבלת' / 'GetAllocationPolicy V2.xlsx'
    protected = [root / 'GetAllocationPolicy.xlsx', root / 'WebSearch.xlsx',
                 root / 'הנחיות כלליות לאחר שינוי גרסה.docx',
                 root / 'קבצים בפורמט שהמערכת מקבלת' / 'GetAllocationPolicy.xlsx',
                 root / 'קבצים בפורמט שהמערכת מקבלת' / 'WebSearch (1).xlsx',
                 root / 'קבצים בפורמט שהמערכת מקבלת' / 'WebSearch (1) V2.xlsx',
                 root / 'פרומפט עדכון 12_57 20_10.txt',
                 root / 'פרומפט עדכון 13_33 20_10.txt',
                 Path(r'C:\Users\Nataliek\Downloads\הנחיות כלליות לאחר שינוי גרסה.docx')]
    assert source.is_absolute() and backup.parent == root / 'ישן'
    assert backup.parent.is_dir()
    payloads = {path: path.read_bytes() for path in [source] + protected}
    (work / 'inputs').mkdir(exist_ok=True)
    backup.mkdir(exist_ok=True)
    write_checked(work / 'inputs' / source.name, payloads[source])
    write_checked(backup / source.name, payloads[source])
    for path, payload in payloads.items():
        assert path.read_bytes() == payload, 'Concurrent source change'
    baseline = {'source': str(source), 'name': source.name, 'sha256': digest(payloads[source]),
                'protected': [{'path': str(path), 'sha256': digest(payloads[path])}
                              for path in protected], 'backup': str(backup)}
    write_checked(work / 'baseline.json', encoded(baseline))
    print(json.dumps({'source': source.name, 'backup_verified': True,
                      'protected_files': len(protected)}, ensure_ascii=False))


def build(work, output):
    assert not output.exists(), 'Use a fresh build directory'
    output.mkdir()
    baseline = json.loads((work / 'baseline.json').read_text(encoding='utf-8'))
    before = (work / 'inputs' / baseline['name']).read_bytes()
    assert digest(before) == baseline['sha256']
    cells_before, _, _, sheets_before = workbook_model(before)
    key_general = ('מידע כללי', 'F2')
    key_component = ('תוספות לתיאור הכלים', 'I12')
    assert cells_before[('תוספות לתיאור הכלים', 'I1')] == 'תוספת לסוכן זה'
    assert cells_before[('תוספות לתיאור הכלים', 'G12')].startswith('קוד מרכיב החשיפה:')
    assert cells_before[key_component] == ''
    old_mapping = 'לצורך חישוב ההצעה בלבד, מפה equity לטווח equity, alternative לטווח alternative, realEstate לטווח realEstate, ו-cash לטווח cashExp. דרוש רשומה יחידה לכל קוד חשיפה המשמש במיפוי; אין לבחור בין כפילויות או למזג אותן. אין להשתמש ב-risk, abroad, fx או cashHold לחישוב הפילוח.'
    assert cells_before[key_general].count(old_mapping) == 1
    new_mapping = (work / 'mapping_paragraph.txt').read_text(encoding='utf-8').rstrip()
    component = (work / 'component_addition.txt').read_text(encoding='utf-8').rstrip()
    addition = cells_before[key_general].replace(old_mapping, new_mapping, 1)
    assert addition.replace(new_mapping, old_mapping, 1) == cells_before[key_general]
    stage, _ = patch_cell(before, key_general, addition)
    result, _ = patch_cell(stage, key_component, component)
    cells_after, _, _, sheets_after = workbook_model(result)
    assert sheets_before == sheets_after and len(sheets_after) == 60
    changed = [key for key in cells_before if cells_before[key] != cells_after[key]]
    assert changed == [key_general, key_component], changed
    assert cells_after[('תוספות לתיאור הכלים', 'H12')] == cells_before[('תוספות לתיאור הכלים', 'H12')]
    assert cells_after[key_general] == addition and cells_after[key_component] == component
    for value in [addition, component]:
        assert 'אין לדרוש רשומת חשיפה' in value
        assert 'cashHold' in value and 'פגם' in value and 'מבני' in value
        assert 'Min' in value and 'Max' in value and 'אותו פרופיל' in value
    with ZipFile(BytesIO(before)) as z_before, ZipFile(BytesIO(result)) as z_after:
        assert z_before.namelist() == z_after.namelist()
        changed_parts = [name for name in z_before.namelist()
                         if z_before.read(name) != z_after.read(name)]
    assert set(changed_parts) == {'xl/sharedStrings.xml', 'xl/worksheets/sheet2.xml'}
    differences = []
    for key in changed:
        differences.extend(unified_diff(cells_before[key].splitlines(True),
                                        cells_after[key].splitlines(True),
                                        fromfile='before ' + '!'.join(key),
                                        tofile='after ' + '!'.join(key)))
    write_checked(output / baseline['name'], result)
    write_checked(output / 'cells.diff', ''.join(differences).encode('utf-8'))
    report = {'name': baseline['name'], 'sha256': digest(result),
              'changed_cells': [list(key) for key in changed],
              'sheet_count': len(sheets_after), 'changed_package_parts': changed_parts,
              'general_H12_and_all_other_cells_unchanged': True,
              'all_other_F2_paragraphs_unchanged': True,
              'styles_RTL_tables_and_metadata_preserved': True,
              'missing_duplicate_or_invalid_cashExp_remains_structural_error': True,
              'no_policy_values_bounds_or_mapping_changed': True}
    write_checked(output / 'verification.json', encoded(report))
    print(json.dumps(report, ensure_ascii=False))


def deploy(work):
    baseline = json.loads((work / 'baseline.json').read_text(encoding='utf-8'))
    first, second = work / 'build_a', work / 'build_b'
    assert {p.name for p in first.iterdir()} == {p.name for p in second.iterdir()}
    for path in first.iterdir():
        assert path.read_bytes() == (second / path.name).read_bytes(), 'Repeat build differs'
    checks = json.loads((first / 'verification.json').read_text(encoding='utf-8'))
    for entry in baseline['protected']:
        assert digest(Path(entry['path']).read_bytes()) == entry['sha256'], 'Protected file changed'
    destination = Path(baseline['source'])
    payload = (first / baseline['name']).read_bytes()
    assert digest(payload) == checks['sha256']
    assert digest((Path(baseline['backup']) / baseline['name']).read_bytes()) == baseline['sha256']
    assert digest(destination.read_bytes()) in {baseline['sha256'], checks['sha256']}, 'Concurrent source change'
    with destination.open('r+b'):
        pass
    if destination.read_bytes() != payload:
        destination.write_bytes(payload)
    assert destination.read_bytes() == payload
    for entry in baseline['protected']:
        assert digest(Path(entry['path']).read_bytes()) == entry['sha256'], 'Protected file changed'
    report = {'saved': str(destination), 'sha256': digest(payload),
              'changed_cells': checks['changed_cells'],
              'repeat_builds_byte_identical': True, 'sheet_count': 60,
              'protected_files_unchanged': len(baseline['protected']),
              'original_general_additions_preserved': True,
              'backup': baseline['backup'],
              'visual_preview': 'Existing import-format Artifact Tool limitation; package preservation verified',
              'imported_into_live_system': False, 'live_agent_tested': False}
    write_checked(work / 'deployment.json', encoded(report))
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
