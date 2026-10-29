"""Read GitHub PRs into local Markdown; never merge, push, or edit review notes.

Requires authenticated GitHub CLI (gh). --watch runs only while invoked.
"""

import argparse
from datetime import datetime, timezone
import html
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import time


class SyncError(Exception):
    """A failed retrieval must not replace the last successful dashboard."""


def repository_name(value):
    value = value.strip()
    for prefix in ('https://github.com/', 'ssh://git@github.com/', 'git@github.com:'):
        if value.startswith(prefix):
            value = value[len(prefix):]
            break
    value = value.removesuffix('/').removesuffix('.git')
    if not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9_.-]*/[A-Za-z0-9][A-Za-z0-9_.-]*', value):
        raise SyncError('저장소를 owner/name 또는 GitHub origin 주소로 지정하세요.')
    return value


def command(args, root):
    try:
        result = subprocess.run(args, cwd=root, capture_output=True, text=True,
                                encoding='utf-8', timeout=120, check=False)
    except FileNotFoundError as exc:
        raise SyncError(f'{args[0]} 실행 도구가 없습니다. 설치·연결 후 다시 조회하세요.') from exc
    except subprocess.TimeoutExpired as exc:
        raise SyncError('조회 시간이 초과됐습니다. 마지막 성공본을 유지합니다.') from exc
    if result.returncode:
        # Do not print CLI stderr: it can contain private repository/auth details.
        raise SyncError(f'{args[0]} 조회 실패 (종료 코드 {result.returncode}). 저장소 접근 권한과 연결 상태를 확인하세요.')
    return result.stdout


def flatten_pages(data):
    if not isinstance(data, list) or any(not isinstance(page, list) for page in data):
        raise SyncError('GitHub 응답이 예상한 페이지 목록이 아닙니다.')
    rows = [row for page in data for row in page]
    if any(not isinstance(row, dict) for row in rows):
        raise SyncError('GitHub 응답의 PR 항목이 잘못됐습니다.')
    return rows


def fetch_pulls(root, repo):
    endpoint = f'repos/{repo}/pulls?state=all&per_page=100'
    raw = command(['gh', 'api', '--hostname', 'github.com', '--method', 'GET',
                   '--paginate', '--slurp', endpoint], root)
    try:
        return flatten_pages(json.loads(raw))
    except json.JSONDecodeError as exc:
        raise SyncError('GitHub 응답 JSON을 읽지 못했습니다.') from exc


def normalize(row):
    try:
        number = row['number']
        sha = row['head']['sha']
        if type(number) is not int or number < 1 or not re.fullmatch(r'[0-9a-fA-F]{40,64}', sha):
            raise ValueError('number/sha')
        if row['state'] not in ('open', 'closed'):
            raise ValueError('state')
        result = {key: row.get(key) for key in (
            'title', 'state', 'draft', 'created_at', 'updated_at', 'closed_at',
            'merged_at', 'merge_commit_sha')}
        result.update(number=number, head_sha=sha,
                      author=(row.get('user') or {}).get('login', '(deleted)'),
                      base=row['base']['ref'])
        if not result['merged_at']:
            result['merge_commit_sha'] = None  # GitHub can return a test-merge SHA for open PRs.
        return result
    except (KeyError, TypeError, ValueError) as exc:
        raise SyncError('필수 PR 필드가 없거나 잘못됐습니다. 기존 현황판을 유지합니다.') from exc


def atomic_write(path, text, create_only=False):
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix='pr-sync-', suffix='.tmp', dir=path.parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8', newline='\n') as stream:
            stream.write(text)
        if create_only:
            try:
                os.link(temporary, path)
            except FileExistsError:
                pass  # A reviewer created it after planning; keep their work.
        else:
            os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def read_json(path, default):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding='utf-8'))
    except (ValueError, OSError) as exc:
        raise SyncError(f'기존 기록을 읽지 못했습니다: {path.name}') from exc


def history_records(area, repo):
    """Shared logs preserve identity and snapshots even on a fresh local checkout."""
    records = {}
    for folder, pattern in (('요청', r'^- 저장소: (\S+)\s*$'),
                            ('검토', r'^repository:\s*(\S+)\s*$')):
        for path in (area / folder).glob('PR-*.md'):
            match = re.search(pattern, path.read_text(encoding='utf-8'), re.M)
            if not match or match[1].lower() != repo.lower():
                raise SyncError(f'기존 문서의 저장소가 다르거나 확인되지 않습니다: {path.name}')
    for path in (area / '이력').glob('PR-*.jsonl'):
        try:
            lines = path.read_text(encoding='utf-8').splitlines()
            events = [json.loads(line) for line in lines if line.strip()]
            if not events or any(e['repository'].lower() != repo.lower() for e in events):
                raise ValueError('repository')
            last = events[-1]['snapshot']
            if path.stem != f"PR-{last['number']:04d}":
                raise ValueError('number')
            records[str(last['number'])] = last
        except (ValueError, KeyError, TypeError) as exc:
            raise SyncError(f'기존 이력의 저장소/내용을 확인하지 못했습니다: {path.name}') from exc
    return records


def label(value):
    """Remote titles remain text, never executable Markdown/HTML links."""
    text = html.escape(str(value or '—')).replace('\n', ' ').replace('\r', ' ')
    for char in '[]|`*#':
        text = text.replace(char, f'&#{ord(char)};')
    return text


def state_label(pull):
    if pull.get('merged_at'):
        return '병합됨'
    if pull['state'] == 'closed':
        return '닫힘·미병합'
    return '초안' if pull.get('draft') else '열림·검토대기'


def review_stub(repo, pull):
    return f'''---
repository: {repo}
pr: {pull['number']}
review_status: pending
initial_head: {pull['head_sha']}
reviewed_head: none
---

# PR #{pull['number']} 검토 기록

자동으로 최초 생성한 양식입니다. initial_head는 수집 당시 버전이며 검토 완료를 뜻하지 않습니다.
이 파일은 동기화 도구가 덮어쓰지 않습니다. 실제 검토 후 reviewed_head에 검토한 전체 SHA를 적으세요.

- 검토자 / 검토일:
- 채택한 내용과 근거:
- 보류·반영하지 않은 내용과 이유:
- 영향받는 지식·결정 ID 및 링크:
- 실제 온톨로지 반영 커밋:
- 검증 결과:

GitHub 병합 사실과 온톨로지 내용 반영 완료는 별도로 확인합니다.
'''


def review_state(text, pull):
    match = re.search(r'^reviewed_head:\s*([0-9a-fA-F]{40,64})\s*$', text, re.M)
    if match:
        return '검토 버전 일치' if match[1].lower() == pull['head_sha'].lower() else '검토 후 변경됨·재검토 필요'
    initial = re.search(r'^initial_head:\s*([0-9a-fA-F]{40,64})\s*$', text, re.M)
    if initial and initial[1].lower() != pull['head_sha'].lower():
        return '미검토·최초 수집 후 변경됨'
    return '미검토'


def render_card(repo, pull, checked_at, seen, review):
    key = f"PR-{pull['number']:04d}"
    availability = '이번 조회에서 확인' if seen else '이번 조회에서 누락·이전 기록 보존 (종료로 판단하지 않음)'
    return f'''# PR #{pull['number']} · {label(pull['title'])}

자동 수집 문서입니다. 판단은 [검토 기록](../검토/{key}.md)에 작성하세요.

- 저장소: {repo}
- 작성자: {label(pull['author'])}
- GitHub 요청: https://github.com/{repo}/pull/{pull['number']}
- 원격 상태: {state_label(pull)}
- 조회 상태: {availability}
- 목록 조회 성공 시각 (UTC): {checked_at}
- 현재 제출 버전: `{pull['head_sha']}`
- 대상 브랜치: {label(pull['base'])}
- 생성 / 변경: {label(pull['created_at'])} / {label(pull['updated_at'])}
- 닫힘 / 병합: {label(pull['closed_at'])} / {label(pull['merged_at'])}
- GitHub 병합 커밋: {label(pull['merge_commit_sha']) if pull.get('merged_at') else '미병합'}
- 로컬 검토 버전: {review}

[수집 이력](../이력/{key}.jsonl). 이력은 조회 시 관측한 변경이며 GitHub의 모든 중간 이벤트를 복제하지 않습니다.
PR 본문·댓글·파일 diff는 이 목록 수집 범위에 포함되지 않습니다. 실제 검토 시 제출 SHA의 변경 내용을 별도로 읽어야 합니다.
'''


def render_dashboard(repo, pulls, checked_at, seen_numbers, review_states):
    lines = ['# PR 현황', '', f'저장소: https://github.com/{repo}', '',
             f'마지막 목록 조회 성공 (UTC): **{checked_at}**', '',
             '현재 GitHub 화면이 아니라 마지막 성공 조회의 사본입니다. 자동 실행 중인지 여부는 이 문서만으로 보장하지 않습니다.', '',
             '병합됨은 GitHub 상태입니다. 실제 온톨로지 반영 내용·이유·커밋은 요청별 검토 기록에서 확인하세요.', '']
    for group in ('열림·검토대기', '초안', '병합됨', '닫힘·미병합'):
        lines += [f'## {group}', '', '| 요청 | 제목 | 작성자 | 검토 버전 | 조회 |', '|---|---|---|---|---|']
        members = [p for p in pulls if state_label(p) == group]
        for pull in members:
            n = pull['number']
            availability = '확인' if n in seen_numbers else '이번 조회 누락·이전 상태'
            lines.append(f"| [#{n}](요청/PR-{n:04d}.md) | {label(pull['title'])} | {label(pull['author'])} | {review_states[n]} | {availability} |")
        if not members:
            lines.append('| — | 마지막 조회 및 보존 기록에 해당 항목 없음 | — | — | — |')
        lines.append('')
    return '\n'.join(lines)


def sync_snapshot(root, repo, rows, checked_at):
    """Apply one fully fetched snapshot. Review notes are created once only."""
    root = Path(root)
    repo = repository_name(repo)
    incoming = [normalize(row) for row in rows]  # Validate all before writing.
    numbers = [p['number'] for p in incoming]
    if len(set(numbers)) != len(numbers):
        raise SyncError('중복 PR 번호가 응답에 있습니다. 조회를 다시 실행하세요.')
    state_dir = root / '.local' / 'pr-sync'
    binding_path = state_dir / 'repository.json'
    binding = read_json(binding_path, {'repository': repo})
    if binding.get('repository', '').lower() != repo.lower():
        raise SyncError('이 폴더의 PR 기록은 다른 저장소에 연결되어 있습니다. 별도 작업 폴더를 사용하세요.')
    cache_path = state_dir / repo.replace('/', '__') / 'snapshot.json'
    previous = read_json(cache_path, {'pulls': {}})
    area = root / '05-협업'
    records = dict(previous['pulls'])
    records.update(history_records(area, repo))  # Shared history can be newer after git pull.
    for pull in incoming:
        records[str(pull['number'])] = pull
    pulls = sorted(records.values(), key=lambda p: p['number'], reverse=True)
    seen = set(numbers)
    planned = []
    review_states = {}
    # Plan all output before committing files; corrupt existing logs fail closed.
    for pull in pulls:
        key = f"PR-{pull['number']:04d}"
        review_path = area / '검토' / f'{key}.md'
        review = review_path.read_text(encoding='utf-8') if review_path.exists() else review_stub(repo, pull)
        if not review_path.exists():
            planned.append((review_path, review))
        review_states[pull['number']] = review_state(review, pull)
        history = area / '이력' / f'{key}.jsonl'
        log = history.read_text(encoding='utf-8') if history.exists() else ''
        try:
            last = json.loads(log.splitlines()[-1]) if log.strip() else None
        except ValueError as exc:
            raise SyncError(f'기존 이력이 손상됐습니다: {key}') from exc
        if pull['number'] in seen and (not last or last.get('snapshot') != pull):
            event = {'observed_at': checked_at, 'repository': repo, 'snapshot': pull}
            planned.append((history, log.rstrip('\n') + ('\n' if log else '') + json.dumps(event, ensure_ascii=False) + '\n'))
        planned.append((area / '요청' / f'{key}.md', render_card(repo, pull, checked_at, pull['number'] in seen, review_states[pull['number']])))
    for path, content in planned:
        atomic_write(path, content, create_only=path.parent == area / '검토')
    atomic_write(binding_path, json.dumps({'repository': repo}) + '\n')
    atomic_write(cache_path, json.dumps({'repository': repo, 'checked_at': checked_at,
                                       'pulls': records, 'raw_response': rows}, ensure_ascii=False, indent=2) + '\n')
    # Dashboard is written last, so it never advertises an unfinished update.
    atomic_write(area / 'PR-현황.md', render_dashboard(repo, pulls, checked_at, seen, review_states))
    return len(incoming)


def sync_once(root, repo=None):
    repo = repository_name(repo or command(['git', 'remote', 'get-url', 'origin'], root))
    rows = fetch_pulls(root, repo)
    checked_at = datetime.now(timezone.utc).isoformat(timespec='seconds')
    return sync_snapshot(root, repo, rows, checked_at)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument('--repo', help='GitHub owner/name; defaults to origin')
    parser.add_argument('--watch', action='store_true', help='Poll until stopped; does not install a background service')
    parser.add_argument('--interval', type=int, default=300, help='Watch interval in seconds, minimum 60')
    args = parser.parse_args(argv)
    if args.interval < 60:
        parser.error('--interval must be at least 60 seconds')
    try:
        while True:
            try:
                count = sync_once(args.root, args.repo)
                print(f'PR 조회 성공: {count}건. 05-협업/PR-현황.md 갱신.', flush=True)
            except (SyncError, OSError, ValueError, KeyError) as exc:
                print(f'PR 조회 실패: {exc} 마지막 성공 현황의 시각을 확인하세요.', file=sys.stderr, flush=True)
                if not args.watch:
                    return 1
            if not args.watch:
                return 0
            time.sleep(args.interval)
    except KeyboardInterrupt:
        return 0


if __name__ == '__main__':
    raise SystemExit(main())
