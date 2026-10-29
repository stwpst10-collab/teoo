"""Temporary fixtures only; no GitHub access or real vault changes."""

import json
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch

from sync_prs import (SyncError, atomic_write, flatten_pages, main, repository_name,
                     sync_snapshot)


def pull(number=1):
    return {'number': number, 'title': '경쟁사 조사', 'state': 'open', 'draft': False,
            'user': {'login': 'researcher'}, 'head': {'sha': 'a' * 40},
            'base': {'ref': 'main'}, 'created_at': '2026-10-29T00:00:00Z',
            'updated_at': '2026-10-29T00:00:00Z', 'closed_at': None,
            'merged_at': None, 'merge_commit_sha': None}


class SyncTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.area = self.root / '05-협업'

    def sync(self, rows):
        return sync_snapshot(self.root, 'owner/venture', rows, '2026-10-29T01:00:00+00:00')

    def test_repository_and_page_validation(self):
        for value in ('owner/venture', 'https://github.com/owner/venture.git',
                      'git@github.com:owner/venture.git', 'ssh://git@github.com/owner/venture.git'):
            self.assertEqual('owner/venture', repository_name(value))
        for value in ('owner/../repo', 'https://other.example/o/r', 'o/r?bad=1'):
            with self.assertRaises(SyncError):
                repository_name(value)
        self.assertEqual([pull(), pull(2)], flatten_pages([[pull()], [pull(2)]]))
        for data in ({}, [pull()], [[None]]):
            with self.assertRaises(SyncError):
                flatten_pages(data)

    def test_idempotence_and_review_preservation(self):
        self.sync([pull()])
        history = self.area / '이력/PR-0001.jsonl'
        before = history.read_bytes()
        review = self.area / '검토/PR-0001.md'
        review.write_text('repository: owner/venture\nmy actual reasoning\nreviewed_head: ' + 'a' * 40, encoding='utf-8')
        self.sync([pull()])
        self.assertEqual(before, history.read_bytes())
        self.assertIn('my actual reasoning', review.read_text(encoding='utf-8'))
        self.assertIn('검토 버전 일치', (self.area / 'PR-현황.md').read_text(encoding='utf-8'))

    def test_new_revision_marks_review_stale(self):
        self.sync([pull()])
        review = self.area / '검토/PR-0001.md'
        review.write_text('repository: owner/venture\nreviewed_head: ' + 'a' * 40, encoding='utf-8')
        changed = pull()
        changed['head']['sha'] = 'b' * 40
        self.sync([changed])
        self.assertIn('재검토 필요', (self.area / 'PR-현황.md').read_text(encoding='utf-8'))
        events = (self.area / '이력/PR-0001.jsonl').read_text(encoding='utf-8').splitlines()
        self.assertEqual(2, len(events))
        self.assertEqual('a' * 40, json.loads(events[0])['snapshot']['head_sha'])

    def test_missing_cached_item_is_not_closed(self):
        self.sync([pull()])
        self.sync([])
        text = (self.area / '요청/PR-0001.md').read_text(encoding='utf-8')
        self.assertIn('열림·검토대기', text)
        self.assertIn('이번 조회에서 누락', text)
        self.assertTrue((self.area / '검토/PR-0001.md').exists())

    def test_group_states_do_not_claim_content_adoption(self):
        rows = [pull(n) for n in range(1, 5)]
        rows[1]['draft'] = True
        rows[2].update(state='closed', merged_at='2026-10-29T01:00:00Z')
        rows[3]['state'] = 'closed'
        self.sync(rows)
        text = (self.area / 'PR-현황.md').read_text(encoding='utf-8')
        for group in ('열림·검토대기', '초안', '병합됨', '닫힘·미병합'):
            self.assertIn('## ' + group, text)
        self.assertIn('review_status: pending', (self.area / '검토/PR-0003.md').read_text(encoding='utf-8'))

    def test_open_pr_test_merge_sha_is_not_actual_merge(self):
        row = pull()
        row['merge_commit_sha'] = 'c' * 40
        self.sync([row])
        card = (self.area / '요청/PR-0001.md').read_text(encoding='utf-8')
        self.assertNotIn('c' * 40, card)
        self.assertIn('GitHub 병합 커밋: 미병합', card)

    def test_invalid_or_duplicate_response_writes_nothing(self):
        bad = pull()
        bad['head']['sha'] = '../bad'
        for rows in ([pull(), bad], [pull(), pull()]):
            with self.assertRaises(SyncError):
                self.sync(rows)
        self.assertEqual([], list(self.root.iterdir()))

    def test_api_failure_preserves_last_success(self):
        self.sync([pull()])
        before = {p.relative_to(self.root): p.read_bytes() for p in self.root.rglob('*') if p.is_file()}
        with patch('sync_prs.fetch_pulls', side_effect=SyncError('offline')):
            self.assertEqual(1, main(['--root', str(self.root), '--repo', 'owner/venture']))
        after = {p.relative_to(self.root): p.read_bytes() for p in self.root.rglob('*') if p.is_file()}
        self.assertEqual(before, after)

    def test_repository_switch_fails_closed(self):
        self.sync([pull()])
        with self.assertRaises(SyncError):
            sync_snapshot(self.root, 'different/venture', [pull()], 'later')

    def test_shared_history_survives_fresh_checkout_without_local_cache(self):
        self.sync([pull()])
        shutil.rmtree(self.root / '.local')
        with self.assertRaises(SyncError):
            sync_snapshot(self.root, 'different/venture', [pull()], 'later')
        self.sync([])
        self.assertIn('이번 조회에서 누락', (self.area / '요청/PR-0001.md').read_text(encoding='utf-8'))

    def test_create_only_keeps_a_concurrent_review(self):
        path = self.root / 'review.md'
        atomic_write(path, 'human', create_only=True)
        atomic_write(path, 'automatic', create_only=True)
        self.assertEqual('human', path.read_text(encoding='utf-8'))

    def test_shared_history_takes_precedence_over_stale_local_cache(self):
        self.sync([pull()])
        cache = self.root / '.local/pr-sync/owner__venture/snapshot.json'
        stale = cache.read_bytes()
        row = pull()
        row['head']['sha'] = 'b' * 40
        self.sync([row])
        cache.write_bytes(stale)
        self.sync([])
        card = (self.area / '요청/PR-0001.md').read_text(encoding='utf-8')
        self.assertIn('b' * 40, card)
        self.assertNotIn('`' + 'a' * 40 + '`', card)

    def test_remote_title_is_literal_and_generated_links_validate(self):
        row = pull()
        row['title'] = '[[missing]] | <script>\n[bad](file:///C:/secret)'
        self.sync([row])
        from validate import validate
        self.assertEqual([], validate(self.root))

    def test_fetch_is_read_only_paginated_gh_command(self):
        from sync_prs import fetch_pulls
        with patch('sync_prs.command', return_value=json.dumps([[pull()]])) as cmd:
            self.assertEqual([pull()], fetch_pulls(self.root, 'owner/venture'))
        args = cmd.call_args.args[0]
        self.assertIn('--paginate', args)
        self.assertIn('--slurp', args)
        self.assertIn('GET', args)
        self.assertIn('repos/owner/venture/pulls?state=all&per_page=100', args)


if __name__ == '__main__':
    unittest.main()
