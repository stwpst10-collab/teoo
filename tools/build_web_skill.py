"""Build the Claude web app skill ZIP from the shared teoo-wiki sources.

The ZIP is deterministic, so `--check` can confirm the shared copy is current.
"""

import argparse
import io
from pathlib import Path
import re
import sys
import zipfile


ROOT = Path(__file__).resolve().parents[1]
SHARED = ROOT / '.agents/skills/teoo-wiki'
OUTPUT = ROOT / '05-협업/배포/teoo 위키 Claude 웹앱 스킬.zip'
# Published path inside the ZIP -> source. Codex-only metadata (agents/) is left out.
FILES = {
    'teoo-wiki/SKILL.md': ROOT / 'tools/claude_web_skill/SKILL.md',
    'teoo-wiki/references/rules.md': SHARED / 'references/rules.md',
    'teoo-wiki/references/workflows.md': SHARED / 'references/workflows.md',
    'teoo-wiki/scripts/wiki.py': SHARED / 'scripts/wiki.py',
}
SECRET = re.compile(r'gh[pousr]_[A-Za-z0-9]{20,}|github_pat_|-----BEGIN [A-Z ]*PRIVATE KEY|://[^/\s:@]+:[^/\s@]+@')


def build():
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w') as bundle:
        for name, source in sorted(FILES.items()):
            data = source.read_bytes().replace(b'\r\n', b'\n')
            if SECRET.search(data.decode('utf-8')):
                raise SystemExit(f'Credential-like text in {source.relative_to(ROOT)}; not packaging.')
            info = zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_STORED  # byte-identical on every OS
            info.create_system = 3  # ZipInfo defaults differ between Windows and Unix
            info.external_attr = 0o644 << 16
            bundle.writestr(info, data)
    return buffer.getvalue()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true', help='Fail if the shared ZIP is missing or stale')
    parser.add_argument('--out', type=Path, default=OUTPUT)
    args = parser.parse_args()
    data = build()
    if args.check:
        current = args.out.is_file() and args.out.read_bytes() == data
        print('Web skill ZIP is current' if current else 'Web skill ZIP is stale; run python tools/build_web_skill.py')
        return 0 if current else 1
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_bytes(data)
    print(f'Wrote {args.out} ({len(data)} bytes, {len(FILES)} files)')
    return 0


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
