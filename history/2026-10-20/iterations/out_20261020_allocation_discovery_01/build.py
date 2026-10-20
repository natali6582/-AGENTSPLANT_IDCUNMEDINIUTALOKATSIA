from argparse import ArgumentParser
from difflib import unified_diff
from pathlib import Path
import hashlib
import json
from xlsx_helpers import patch_cell, workbook_model


def digest(data):
    return hashlib.sha256(data).hexdigest()


def update_prompt(source, work):
    decoded = source.decode('utf-8-sig')
    newline = '\r\n' if '\r\n' in decoded else '\n'
    before = decoded.replace('\r\n', '\n')
    anchor = '\n\nמקורות המחקר\n'
    assert before.count(anchor) == 1
    addition = (work / 'discovery_guidance.txt').read_text(encoding='utf-8').rstrip()
    assert 'מיקוד שלב האיתור' not in before
    after = before.replace(anchor, '\n\n' + addition + anchor, 1)
    assert after.replace('\n\n' + addition, '', 1) == before
    assert before.split('סדר העבודה ותנאי המשך\n', 1)[1] == after.split('סדר העבודה ותנאי המשך\n', 1)[1]
    output = after.replace('\n', newline).encode('utf-8')
    if source.startswith(b'\xef\xbb\xbf'):
        output = b'\xef\xbb\xbf' + output
    difference = ''.join(unified_diff(before.splitlines(True), after.splitlines(True),
                                     fromfile='before_prompt', tofile='after_prompt'))
    return output, difference


def run(work, output):
    assert not output.exists(), 'Use a fresh build directory'
    output.mkdir()
    baseline = json.loads((work / 'baseline.json').read_text(encoding='utf-8-sig'))
    reports = []
    for entry in baseline['targets']:
        source = (work / 'inputs' / entry['name']).read_bytes()
        assert digest(source) == entry['sha256'].lower()
        if entry['name'].endswith('.txt'):
            result, difference = update_prompt(source, work)
            (output / 'prompt.diff').write_text(difference, encoding='utf-8')
            checks = {'added_section': 'מיקוד שלב האיתור', 'all_existing_prompt_text_unchanged': True,
                      'call_budget_recovery_verification_and_calculation_rules_unchanged': True}
        elif entry['name'].startswith('WebSearch'):
            before = workbook_model(source)[0][('מידע כללי', 'F2')]
            insertion = (work / 'search_discovery_addition.txt').read_text(encoding='utf-8').rstrip()
            paragraphs = before.split('\n\n', 1)
            assert len(paragraphs) == 2 and 'בעת איתור מועמדים למחקר זה' not in before
            value = paragraphs[0] + '\n\n' + insertion + '\n\n' + paragraphs[1]
            assert value.replace('\n\n' + insertion, '', 1) == before
            result, checks = patch_cell(source, ('מידע כללי', 'F2'), value)
            assert checks['changed_package_parts'] == ['xl/sharedStrings.xml']
        else:
            result = source
            checks = {'changed_cells': [], 'sheet_count': len(workbook_model(source)[3]),
                      'byte_identical_to_source': True}
        (output / entry['name']).write_bytes(result)
        assert (output / entry['name']).read_bytes() == result
        reports.append({'file': entry['name'], 'sha256': digest(result), **checks})
    (output / 'verification.json').write_text(json.dumps(reports, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(reports, ensure_ascii=False))


def main():
    parser = ArgumentParser()
    parser.add_argument('--work', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    run(args.work.resolve(), args.output.resolve())


if __name__ == '__main__':
    main()
