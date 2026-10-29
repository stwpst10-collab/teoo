#!/usr/bin/env python3
"""Read-only Markdown access to a verified local teoo Git checkout."""
import argparse
import json
import os
from pathlib import Path, PureWindowsPath
import re
import subprocess
import sys
from urllib.parse import urlsplit


REMOTE = 'https://github.com/stwpst10-collab/teoo.git'
CORPUS = {'00-운영', '10-지식', '20-리서치', '30-근거', '40-결정', '50-검증', '80-산출물'}
EXCLUDED = {'첨부', '90 작업자료', 'node_modules', 'tmp', '__pycache__'}
ALIASES = {'ovira': ('ovira', '오비라'), '오비라': ('ovira', '오비라')}


def git(repo, *args, optional=False):
    result = subprocess.run(
        ['git', '-C', str(repo), *args], capture_output=True,
        encoding='utf-8', errors='replace', timeout=20,
        env={**os.environ, 'GIT_OPTIONAL_LOCKS': '0', 'GIT_TERMINAL_PROMPT': '0'},
    )
    if result.returncode and not optional:
        raise ValueError('Local Git inspection failed; provide a valid teoo checkout with --repo.')
    return result.stdout.strip() if not result.returncode else ''


def correct_origin(value):
    if value.startswith('git@github.com:'):
        path = value[len('git@github.com:'):]
    else:
        try:
            parsed = urlsplit(value)
            valid_host = parsed.scheme == 'https' and parsed.hostname == 'github.com' and parsed.port in (None, 443)
        except ValueError:
            return False
        if not valid_host:
            return False
        if parsed.query or parsed.fragment:
            return False
        path = parsed.path.lstrip('/')
    return path in {'stwpst10-collab/teoo', 'stwpst10-collab/teoo.git'}


def repository(explicit):
    candidate = Path(explicit).expanduser() if explicit else None
    if candidate is None:
        cwd = Path.cwd().resolve()
        for parent in (cwd, *cwd.parents):
            for possible in (parent, parent / 'venture'):
                if (possible / '.git').exists():
                    candidate = possible
                    break
            if candidate is not None:
                break
    if candidate is None:
        # Installed packages may retain a private pointer next to scripts/.
        script = Path(__file__).resolve()
        for config in (script.parent.parent / 'workspace.json', script.parent / 'workspace.json'):
            if config.is_file():
                settings = json.loads(config.read_text(encoding='utf-8-sig'))
                value = settings.get('repository_root') if isinstance(settings, dict) else None
                if not isinstance(value, str) or not value.strip():
                    raise ValueError('workspace.json must contain repository_root; pass --repo explicitly.')
                candidate = Path(value).expanduser()
                if not candidate.is_absolute():
                    candidate = config.parent / candidate
                break
    if candidate is None or not candidate.is_dir():
        raise ValueError('No local teoo checkout found. Pass --repo PATH for an existing teoo clone.')
    repo = Path(git(candidate.resolve(), 'rev-parse', '--show-toplevel')).resolve()
    origins = git(repo, 'config', '--get-all', 'remote.origin.url', optional=True).splitlines()
    if len(origins) != 1 or not correct_origin(origins[0]):
        raise ValueError('Origin is missing or is not github.com/stwpst10-collab/teoo. Pass a valid teoo clone with --repo.')
    return repo


def allowed(parts):
    return bool(parts) and parts[0] in CORPUS and all(
        part not in EXCLUDED and not part.startswith('.') for part in parts)


def read_path(repo, relative):
    path = Path(relative)
    if path.is_absolute() or PureWindowsPath(relative).is_absolute() or '..' in path.parts or '\\' in relative:
        raise ValueError('Use a repository-relative Markdown path without traversal.')
    if not allowed(path.parts) or path.suffix.lower() != '.md':
        raise ValueError('Only Markdown in the allowed wiki corpus can be read.')
    full = repo / path
    resolved = full.resolve()
    if not resolved.is_relative_to(repo) or not allowed(resolved.relative_to(repo).parts) or not full.is_file():
        raise ValueError('The requested document is missing or resolves outside the allowed wiki corpus.')
    return full


def documents(repo):
    for folder in sorted(CORPUS):
        start = repo / folder
        if not start.is_dir() or start.is_symlink() or not start.resolve().is_relative_to(repo):
            continue
        for root, dirs, files in os.walk(start, followlinks=False):
            dirs[:] = sorted(d for d in dirs if d not in EXCLUDED and not d.startswith('.')
                             and not Path(root, d).is_symlink())
            for name in sorted(files):
                path = Path(root, name)
                if path.suffix.lower() == '.md':
                    try:
                        yield read_path(repo, path.relative_to(repo).as_posix())
                    except ValueError:
                        continue


def metadata(text):
    front = re.match(r'\A---\s*\n(.*?)\n---(?:\s*\n|$)', text, re.S)
    fields = dict(re.findall(r'^([\w-]+):\s*([^\n]*)$', front[1], re.M)) if front else {}
    body = text[front.end():] if front else text
    title = re.search(r'^#\s+(.+)$', body, re.M)
    return {'title': title[1].strip() if title else '',
            **{key: fields.get(key, '').strip(' \"\'') for key in ('id', 'status', 'review', 'aliases')}}


def search(repo, query, limit):
    query = query.strip().casefold()
    if not query:
        raise ValueError('Search query cannot be empty.')
    terms = ALIASES.get(query, (query,))
    hits = []
    for path in documents(repo):
        text = path.read_text(encoding='utf-8-sig')
        matches = [(n, line) for n, line in enumerate(text.splitlines(), 1)
                   if any(term in line.casefold() for term in terms)]
        if not matches:
            continue
        info = metadata(text)
        line, snippet = matches[0]
        score = 0 if any(term in (info['title'] + ' ' + info['aliases']).casefold() for term in terms) else 1
        hits.append((score, {'path': path.relative_to(repo).as_posix(), **info,
                             'line': line, 'snippet': snippet.strip()[:240]}))
    hits.sort(key=lambda hit: (hit[0], hit[1]['path']))
    return {'results': [item for _, item in hits[:limit]], 'total': len(hits), 'truncated': len(hits) > limit}


def positive(value):
    number = int(value)
    if number < 1:
        raise argparse.ArgumentTypeError('Must be positive.')
    return number


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo', help='Existing local teoo checkout')
    commands = parser.add_subparsers(dest='command', required=True)
    commands.add_parser('status')
    find = commands.add_parser('search')
    find.add_argument('query')
    find.add_argument('--limit', type=positive, default=10)
    read = commands.add_parser('read')
    read.add_argument('path')
    read.add_argument('--max-chars', type=positive, default=12000)
    args = parser.parse_args()
    try:
        repo = repository(args.repo)
        if args.command == 'status':
            data = {'repository': str(repo), 'remote': REMOTE,
                    'branch': git(repo, 'symbolic-ref', '--short', 'HEAD', optional=True) or 'detached',
                    'dirty': bool(git(repo, 'status', '--porcelain'))}
        elif args.command == 'search':
            data = search(repo, args.query, args.limit)
        else:
            path = read_path(repo, args.path)
            text = path.read_text(encoding='utf-8-sig')
            data = {'path': path.relative_to(repo).as_posix(), **metadata(text),
                    'content': text[:args.max_chars], 'truncated': len(text) > args.max_chars}
        print(json.dumps(data, ensure_ascii=False, indent=2))
        return 0
    except (ValueError, OSError, subprocess.SubprocessError) as exc:
        # Never echo Git stderr, remote URLs, or credential-bearing config values.
        message = str(exc) if isinstance(exc, ValueError) and not isinstance(exc, json.JSONDecodeError) else 'Local wiki access failed; verify the checkout and workspace.json.'
        print(json.dumps({'error': message}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    raise SystemExit(main())
