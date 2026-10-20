from argparse import ArgumentParser
from pathlib import Path
import hashlib
import json


def sha(data):
    return hashlib.sha256(data).hexdigest()


def main():
    parser = ArgumentParser()
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--approved', type=Path, required=True)
    parser.add_argument('--approved-policy', type=Path, required=True)
    parser.add_argument('--report', type=Path, required=True)
    args = parser.parse_args()
    root = args.root.resolve()
    tool_dir = root / 'קבצים בפורמט שהמערכת מקבלת'
    target = tool_dir / 'WebSearch (1) V2.xlsx'
    policy = tool_dir / 'GetAllocationPolicy V2.xlsx'
    payload = args.approved.read_bytes()
    expected = '76d3e1c66671e39e5cfa9be28d141d66238878dbe53fc6f383d7def74901ff44'
    expected_policy = 'b33b00c061ac77f4c6ce3685905cefe586862e5c43cb2f0f655aeff6d70e5af5'
    assert sha(payload) == expected
    assert sha(args.approved_policy.read_bytes()) == expected_policy
    assert target.read_bytes() == payload, 'Current target differs; investigate instead of overwriting'
    assert policy.read_bytes() == args.approved_policy.read_bytes()
    protected = list(root.glob('*.txt')) + list(root.glob('*.xlsx')) + list(root.glob('*.docx'))
    protected += [item for item in tool_dir.glob('*.xlsx') if item != target]
    snapshots = {item: item.read_bytes() for item in protected}
    root_entries = {item.name for item in root.iterdir()}
    tool_entries = {item.name for item in tool_dir.iterdir()}
    backup = root / 'ישן' / 'גיבוי_20261020_שמירה_חוזרת_WebSearch_V2_01'
    assert backup.parent.is_dir()
    backup.mkdir(exist_ok=True)
    backup_file = backup / target.name
    if backup_file.exists():
        assert backup_file.read_bytes() == payload
    else:
        with backup_file.open('xb') as stream:
            stream.write(payload)
    assert backup_file.read_bytes() == payload
    # Open without truncation first; a sharing lock stops the task safely.
    with target.open('r+b') as stream:
        assert stream.read() == payload
        stream.seek(0)
        written = stream.write(payload)
        assert written == len(payload)
        stream.truncate()
        stream.flush()
    assert target.read_bytes() == payload
    for item, original in snapshots.items():
        assert item.read_bytes() == original, 'Unrelated file changed'
    assert root_entries == {item.name for item in root.iterdir()}
    assert tool_entries == {item.name for item in tool_dir.iterdir()}
    report = {'websearch_saved': str(target), 'websearch_sha256': sha(target.read_bytes()),
              'content_already_matched_approved_version': True,
              'policy_checked_unchanged': str(policy), 'policy_sha256': sha(policy.read_bytes()),
              'protected_files_unchanged': len(snapshots),
              'root_and_tool_directory_entries_unchanged': True,
              'backup': str(backup_file), 'live_system_tested': False}
    data = json.dumps(report, ensure_ascii=False, sort_keys=True, indent=2).encode('utf-8')
    if args.report.exists():
        assert args.report.read_bytes() == data
    args.report.write_bytes(data)
    assert args.report.read_bytes() == data
    print(json.dumps(report, ensure_ascii=False))


if __name__ == '__main__':
    main()
