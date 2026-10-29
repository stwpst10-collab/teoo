"""Export the committed teoo tree as a ZIP for Claude web app uploads (no Git there).

Only files tracked at the chosen commit are included, so ignored personal state
(.local, .env, workspace.json, credentials) never enters the snapshot.
"""

import argparse
from datetime import datetime, timezone
import io
import json
from pathlib import Path
import subprocess
import sys
import tarfile
import zipfile


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / '.agents/skills/teoo-wiki/scripts'))
import wiki  # noqa: E402  (shared origin check and manifest name)

FOLDER = 'teoo'


def git(*args, binary=False):
    result = subprocess.run(['git', '-C', str(ROOT), *args], capture_output=True, check=True)
    return result.stdout if binary else result.stdout.decode('utf-8').strip()


def export(output, ref='HEAD'):
    origins = git('config', '--get-all', 'remote.origin.url').splitlines()
    if len(origins) != 1 or not wiki.correct_origin(origins[0]):
        raise SystemExit('origin is not github.com/stwpst10-collab/teoo; refusing to export.')
    commit = git('rev-parse', '--short', ref)
    manifest = {
        'remote': wiki.REMOTE,
        'commit': commit,
        'branch': git('rev-parse', '--abbrev-ref', ref) if ref == 'HEAD' else ref,
        'exported': datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
        'note': '커밋된 파일만 포함. 내보낸 뒤의 로컬 미커밋 변경은 없음.',
    }
    archive = tarfile.open(fileobj=io.BytesIO(git('archive', '--format=tar', commit, binary=True)))
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as bundle:
        for member in archive.getmembers():
            if member.isfile():
                bundle.writestr(f'{FOLDER}/{member.name}', archive.extractfile(member).read())
        bundle.writestr(f'{FOLDER}/{wiki.SNAPSHOT}', json.dumps(manifest, ensure_ascii=False, indent=2))
    return manifest


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--ref', default='HEAD', help='Commit or branch to export (default: HEAD)')
    parser.add_argument('--out', type=Path, help='Output ZIP path (default: .local/스냅샷/...)')
    args = parser.parse_args()
    day = datetime.now().strftime('%Y-%m-%d')
    commit = git('rev-parse', '--short', args.ref)
    output = args.out or ROOT / '.local' / '스냅샷' / f'teoo 자료 스냅샷 {day} {commit}.zip'
    manifest = export(output, args.ref)
    if args.ref == 'HEAD' and git('status', '--porcelain'):
        print('주의: 미커밋 변경은 스냅샷에 포함되지 않았습니다.', file=sys.stderr)
    print(json.dumps({'output': str(output), 'bytes': output.stat().st_size, **manifest},
                     ensure_ascii=False, indent=2))
    return 0


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    raise SystemExit(main())
