"""Read-only validation for the shared Markdown vault (Python standard library)."""

import argparse
import json
import os
from pathlib import Path
import re
from urllib.parse import unquote


IGNORED_DIRS = {'.git', '.obsidian', '.omx', '.local', '.tmp', 'tmp', 'node_modules', '__pycache__'}
HISTORICAL_DIRS = {'90 작업자료', '80-산출물'}
STRUCTURED_DIRS = {'10-지식', '20-리서치', '30-근거', '40-결정', '50-검증'}
CONFLICT = re.compile(r'^(?:<{7}(?: .*)?|={7}|>{7}(?: .*)?|\|{7}(?: .*)?)$', re.M)
WIKI = re.compile(r'\[\[([^\]\n]+)\]\]')
MARKDOWN = re.compile(r'(?<!!)\[[^\]\n]*\]\(\s*(<[^>\n]+>|[^\s\n]+(?:\([^\n]*?\)[^\s\n]*)?)(?:\s+["\'][^\n]*?["\'])?\s*\)|!\[[^\]\n]*\]\(\s*(<[^>\n]+>|[^\s\n]+)\s*\)')


def files_in(root):
    """Prune caches, but retain archived files as possible link targets."""
    result = []
    for folder, dirs, names in os.walk(root):
        dirs[:] = sorted(d for d in dirs if d not in IGNORED_DIRS and not Path(folder, d).is_symlink())
        result.extend(Path(folder, name) for name in sorted(names))
    return result


def prose_only(text):
    """Ignore fenced and inline code examples when checking links."""
    lines = []
    fence = None
    for line in text.splitlines():
        marker = re.match(r'^\s{0,3}(`{3,}|~{3,})', line)
        if fence:
            if marker and marker[1][0] == fence[0] and len(marker[1]) >= len(fence):
                fence = None
            lines.append('')
        elif marker:
            fence = marker[1]
            lines.append('')
        else:
            lines.append(re.sub(r'(`+).*?\1', '', line))
    return '\n'.join(lines)


def frontmatter(text):
    """Read the scalar fields required by this vault, without a YAML dependency."""
    match = re.match(r'\A---\s*\n(.*?)\n---(?:\s*\n|$)', text, re.S)
    if not match:
        return {}
    fields = {}
    for line in match[1].splitlines():
        item = re.match(r'^([A-Za-z_][A-Za-z0-9_-]*):\s*(.*?)\s*$', line)
        if item:
            fields[item[1]] = item[2].strip('"\'')
    return fields


def is_absolute_local(target):
    return bool(re.match(r'^(?:[A-Za-z]:[/\\]|file:|[/\\])', target, re.I))


def validate(root):
    root = Path(root).resolve()
    errors = []
    if not root.is_dir():
        return [f'Vault directory does not exist: {root}']
    files = files_in(root)
    file_set = {p.resolve() for p in files}
    names = {}
    for path in files:
        for key in {path.name, path.stem}:
            names.setdefault(key, set()).add(path.resolve())
    ids = {}

    def fail(path, message):
        errors.append(f'{path.relative_to(root).as_posix()}: {message}')

    def check_relative(path, target, kind):
        if is_absolute_local(target):
            fail(path, f'Absolute local {kind} link: {target}')
            return
        resolved = (path.parent / target).resolve()
        if not resolved.is_relative_to(root):
            fail(path, f'{kind} link leaves vault: {target}')
        elif not resolved.exists():
            fail(path, f'Broken {kind} link: {target}')

    for path in files:
        relative = path.relative_to(root)
        output_index = relative.as_posix() == '80-산출물/README.md'
        if (relative.parts[0] in HISTORICAL_DIRS and not output_index) or '첨부' in relative.parts:
            continue
        if path.suffix not in {'.md', '.canvas'}:
            continue
        try:
            text = path.read_text(encoding='utf-8-sig')
        except (UnicodeError, OSError) as exc:
            fail(path, f'Cannot read UTF-8: {exc}')
            continue
        if CONFLICT.search(text):
            fail(path, 'Unresolved merge conflict marker')
        if path.suffix == '.canvas':
            try:
                canvas = json.loads(text)
                for node in canvas.get('nodes', []):
                    if node.get('type') == 'file':
                        target = node.get('file', '')
                        if not target or is_absolute_local(target) or (root / target).resolve() not in file_set:
                            fail(path, f'Broken Canvas file: {target}')
            except (json.JSONDecodeError, AttributeError, TypeError) as exc:
                fail(path, f'Invalid Canvas: {exc}')
            continue

        fields = frontmatter(text)
        required = set()
        if relative.parts[0] in STRUCTURED_DIRS and path.name != 'README.md':
            required = {'id', 'type', 'status', 'project'}
            if relative.parts[0] == '20-리서치':
                required |= {'author', 'question', 'review', 'source_status', 'created'}
        missing = sorted(key for key in required if not fields.get(key))
        if missing:
            fail(path, f'Missing metadata: {", ".join(missing)}')
        if relative.parts[0] == '20-리서치' and required:
            for key, allowed in (
                ('review', {'pending', 'reviewed'}),
                ('source_status', {'primary', 'secondary', 'mixed', 'unverified'}),
            ):
                if fields.get(key) and fields[key] not in allowed:
                    fail(path, f'Invalid {key}: {fields[key]} (allowed: {", ".join(sorted(allowed))})')
        note_id = fields.get('id')
        if note_id:
            if note_id in ids:
                fail(path, f'Duplicate id {note_id}: {ids[note_id]}')
            else:
                ids[note_id] = relative.as_posix()

        prose = prose_only(text)
        for match in WIKI.finditer(prose):
            target = unquote(match[1].split('|', 1)[0].split('#', 1)[0]).strip()
            if not target:
                continue
            if is_absolute_local(target):
                fail(path, f'Absolute local wiki link: {target}')
                continue
            if '/' in target or '\\' in target:
                candidates = {(base / target).resolve() for base in (root, path.parent)}
                candidates |= {p.with_suffix('.md') for p in candidates if not p.suffix}
                found = candidates & file_set
            else:
                found = names.get(target, set())
            if not found:
                fail(path, f'Broken wiki link: {target}')
            elif len(found) > 1:
                fail(path, f'Ambiguous wiki link: {target}')
        for match in MARKDOWN.finditer(prose):
            target = (match[1] or match[2]).strip('<>')
            target = unquote(target).split('#', 1)[0]
            if not target:
                continue
            if is_absolute_local(target):
                fail(path, f'Absolute local Markdown link: {target}')
            elif re.match(r'^[A-Za-z][A-Za-z0-9+.-]*:', target):
                continue
            else:
                check_relative(path, target, 'Markdown')
    return errors


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, default=Path(__file__).resolve().parents[1])
    args = parser.parse_args()
    errors = validate(args.root)
    for error in errors:
        print(f'ERROR: {error}')
    print(f'Validation {"FAILED" if errors else "PASSED"}: {len(errors)} error(s)')
    return bool(errors)


if __name__ == '__main__':
    raise SystemExit(main())
