from argparse import ArgumentParser
from pathlib import Path
import hashlib
import json


def digest(data):
    return hashlib.sha256(data).hexdigest()


def run(root, work, backup):
    targets = [root / 'פרומפט עדכון 12_57 20_10.txt',
               root / 'קבצים בפורמט שהמערכת מקבלת' / 'GetAllocationPolicy V2.xlsx',
               root / 'קבצים בפורמט שהמערכת מקבלת' / 'WebSearch (1) V2.xlsx']
    protected = [root / 'GetAllocationPolicy.xlsx', root / 'WebSearch.xlsx',
                 root / 'הנחיות כלליות לאחר שינוי גרסה.docx',
                 root / 'קבצים בפורמט שהמערכת מקבלת' / 'GetAllocationPolicy.xlsx',
                 root / 'קבצים בפורמט שהמערכת מקבלת' / 'WebSearch (1).xlsx', targets[1]]
    sources = {path: path.read_bytes() for path in set(targets + protected)}
    records = []
    for path in targets:
        payload = sources[path]
        for destination in [work / 'inputs' / path.name, backup / path.name]:
            if destination.exists():
                assert destination.read_bytes() == payload
            destination.write_bytes(payload)
            assert destination.read_bytes() == payload
        records.append({'path': str(path), 'name': path.name, 'sha256': digest(payload)})
    for path, payload in sources.items():
        assert path.read_bytes() == payload, 'Concurrent source change'
    report = {'targets': records, 'protected': [{'path': str(path), 'sha256': digest(sources[path])}
                                               for path in protected], 'backup': str(backup)}
    (work / 'baseline.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'backup_verified': len(records), 'protected_files': len(protected),
                      'prompt': targets[0].name, 'source_hashes': [item['sha256'] for item in records]}, ensure_ascii=False))


def main():
    parser = ArgumentParser()
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--work', type=Path, required=True)
    parser.add_argument('--backup', type=Path, required=True)
    args = parser.parse_args()
    run(args.root, args.work, args.backup)


if __name__ == '__main__':
    main()
