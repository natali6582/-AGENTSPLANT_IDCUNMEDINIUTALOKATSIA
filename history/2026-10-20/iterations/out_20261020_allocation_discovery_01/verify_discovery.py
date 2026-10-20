from argparse import ArgumentParser
from pathlib import Path
import json
from xlsx_helpers import workbook_model


def admissible_candidate(record, institutions, document_types, year):
    fields = {'institution', 'document', 'publicationDate', 'officialDomain', 'documentType', 'component', 'direction', 'evidence'}
    content = record['content']
    return (set(content) == fields and all(content.values())
            and content['institution'] in institutions and content['documentType'] in document_types
            and content['component'] in {'equity', 'cash', 'alternative', 'realEstate'}
            and content['direction'] in {'increase', 'decrease'}
            and content['publicationDate'].startswith(str(year) + '-')
            and record['official_specific_source'] and record['supported_publication_month']
            and record['supported_domain'] and record['supported_document']
            and record['explicit_strategic_allocation'] and not record['inferred_direction'])


def run(work, build):
    name = 'פרומפט עדכון 12_57 20_10.txt'
    before = (work / 'inputs' / name).read_text(encoding='utf-8-sig')
    after = (build / name).read_text(encoding='utf-8-sig')
    addition = (work / 'discovery_guidance.txt').read_text(encoding='utf-8').rstrip()
    assert after.replace('\n\n' + addition, '', 1) == before
    assert after.split('סדר העבודה ותנאי המשך\n', 1)[1] == before.split('סדר העבודה ותנאי המשך\n', 1)[1]
    references = before.split('\nמקורות המחקר\n', 1)[1].split('\n\nסדר העבודה', 1)[0]
    institutions = {line.split(': ', 1)[0] for line in references.splitlines() if ': https://' in line}
    assert institutions == {'Goldman Sachs', 'JPMorgan', 'BlackRock', 'Vanguard', 'Morgan Stanley', 'UBS', 'Citi', 'IMF', 'World Bank', 'OECD'}
    document_types = {'Capital Market Assumptions', 'Investment Outlook', 'Market Outlook', 'Strategic Asset Allocation',
                      'Asset Allocation Outlook', 'House View', 'CIO research equivalent'}
    for clause in ['מסמך ומקור קודמים לרשומה', 'אין לשלוח קריאה נפרדת לכל מוסד',
                   'כל עשרת המוסדות נשארים מותרים', 'ללא הבטחה לסינון טכני',
                   'ללא מספר מינימלי', 'אין להסיק חודש או שנת פרסום',
                   'כל רשומה שנותרה היא מועמדת בלבד', 'אינו מתיר חיפוש נוסף']:
        assert clause in after, clause
    for invariant in ['עד שש רשומות ועד 900 מילים', 'לכל היותר שלוש קריאות WebSearch',
                      'כולל אימות נפרד גם לאחר איתור חוזר', 'לפחות שני מוסדות שונים',
                      'חובה לבצע ניסיון אחד להשלמת האימות', 'אין להניח אפס קריאות או ניסיון פנוי']:
        assert after.count(invariant) == before.count(invariant) > 0
    valid = {'content': {'institution': 'JPMorgan', 'document': 'Synthetic strategic publication fixture',
                         'publicationDate': '2026-06', 'officialDomain': 'example.invalid',
                         'documentType': 'Strategic Asset Allocation', 'component': 'equity',
                         'direction': 'increase', 'evidence': 'Synthetic explicit strategic allocation statement'},
             'official_specific_source': True, 'supported_publication_month': True,
             'supported_domain': True, 'supported_document': True,
             'explicit_strategic_allocation': True, 'inferred_direction': False}
    # Independent acceptance examples; synthetic sources are never claimed to be real research.
    cases = [('supported_candidate', {}, {}, True),
             ('outside_whitelist_LPL', {'institution': 'LPL Financial'}, {}, False),
             ('outside_whitelist_Neuberger', {'institution': 'Neuberger Berman'}, {}, False),
             ('outside_whitelist_Kempen', {'institution': 'Van Lanschot Kempen'}, {}, False),
             ('inferred_month', {}, {'supported_publication_month': False}, False),
             ('forecast_title_is_not_publication_year', {'publicationDate': '2025-12'}, {}, False),
             ('inferred_direction', {}, {'inferred_direction': True}, False),
             ('macro_only', {}, {'explicit_strategic_allocation': False}, False),
             ('investor_behavior_is_not_recommendation', {}, {'explicit_strategic_allocation': False}, False),
             ('homepage_only', {}, {'official_specific_source': False}, False),
             ('inferred_domain', {}, {'supported_domain': False}, False),
             ('invented_document', {}, {'supported_document': False}, False),
             ('missing_evidence', {'evidence': ''}, {}, False),
             ('unsupported_document_type', {'documentType': 'Press Release'}, {}, False)]
    results = []
    for label, content_updates, evidence_updates, expected in cases:
        candidate = {**valid, **evidence_updates, 'content': {**valid['content'], **content_updates}}
        actual = bool(admissible_candidate(candidate, institutions, document_types, 2026))
        assert actual == expected, label
        results.append({'name': label, 'eligible': actual})
    for institution in institutions:
        assert admissible_candidate({**valid, 'content': {**valid['content'], 'institution': institution}}, institutions, document_types, 2026)
    for document_type in document_types:
        assert admissible_candidate({**valid, 'content': {**valid['content'], 'documentType': document_type}}, institutions, document_types, 2026)
    search_name = 'WebSearch (1) V2.xlsx'
    old_cells = workbook_model((work / 'inputs' / search_name).read_bytes())[0]
    new_cells = workbook_model((build / search_name).read_bytes())[0]
    assert [key for key in old_cells if old_cells[key] != new_cells[key]] == [('מידע כללי', 'F2')]
    tool_addition = (work / 'search_discovery_addition.txt').read_text(encoding='utf-8').rstrip()
    assert new_cells[('מידע כללי', 'F2')].replace('\n\n' + tool_addition, '', 1) == old_cells[('מידע כללי', 'F2')]
    for clause in ['פרסום רשמי מסוים של מוסד מותר', 'אינם מסנן טכני מובטח', 'אין להשלים פרטי מועמד בהשערה', 'ואימות נפרד']:
        assert clause in new_cells[('מידע כללי', 'F2')]
    assert (build / 'GetAllocationPolicy V2.xlsx').read_bytes() == (work / 'inputs' / 'GetAllocationPolicy V2.xlsx').read_bytes()
    return {'case_count': len(cases), 'cases': results, 'all_ten_institutions_remain_permitted': True,
            'original_document_types_remain_permitted': True, 'all_original_prompt_rules_byte_equivalent': True,
            'call_budget_recovery_verification_rules_unchanged': True, 'source_and_other_tool_cells_unchanged': True,
            'mode': 'Static content and independent synthetic acceptance cases; no search-provider execution',
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
