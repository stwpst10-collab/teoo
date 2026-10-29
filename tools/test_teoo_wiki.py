"""Exercise the read-only helper through its public CLI and temporary Git repos."""
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest


SCRIPT = Path(__file__).resolve().parents[1] / '.agents/skills/teoo-wiki/scripts/wiki.py'
ORIGIN = 'https://github.com/stwpst10-collab/teoo.git'


class WikiCliTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)
        self.repo = self.base / 'venture'
        self.repo.mkdir()
        self.git('init', '-q')
        self.git('remote', 'add', 'origin', ORIGIN)
        self.write('10-지식/제품/오비라.md', '---\nid: ALT-002\nstatus: 판단유보\nreview: pending\naliases: ["Ovira"]\n---\n# 오비라\n베타 사용자 80명. 유료 여부 미확인.\n')

    def git(self, *args):
        return subprocess.run(['git', '-C', str(self.repo), *args], capture_output=True, check=True)

    def write(self, path, text):
        p = self.repo / path
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(text, encoding='utf-8')
        return p

    def cli(self, *args, explicit=True, cwd=None, script=SCRIPT):
        command = [sys.executable, str(script)]
        if explicit:
            command += ['--repo', str(self.repo)]
        return subprocess.run(command + list(args), cwd=cwd or self.base,
                              capture_output=True, encoding='utf-8',
                              env={**os.environ, 'PYTHONIOENCODING': 'utf-8'})

    def result(self, *args, **kwargs):
        run = self.cli(*args, **kwargs)
        self.assertEqual(0, run.returncode, run.stderr)
        return json.loads(run.stdout)

    def test_status_and_supported_origins(self):
        for origin in (ORIGIN, 'https://github.com/stwpst10-collab/teoo',
                       'git@github.com:stwpst10-collab/teoo.git'):
            self.git('remote', 'set-url', 'origin', origin)
            data = self.result('status')
            self.assertEqual(ORIGIN, data['remote'])
            self.assertTrue(data['dirty'])
            self.assertTrue(data['branch'])

    def test_mismatched_or_missing_origin_fails_without_content(self):
        for origin in ('https://secret:token@github.com/elsewhere/private.git',
                       'https://github.com.evil.test/stwpst10-collab/teoo.git',
                       'https://github.com:secret/stwpst10-collab/teoo.git'):
            self.git('remote', 'set-url', 'origin', origin)
            run = self.cli('read', '10-지식/제품/오비라.md')
            self.assertNotEqual(0, run.returncode)
            self.assertNotIn('80명', run.stdout + run.stderr)
            self.assertNotIn('secret', run.stdout + run.stderr)
            self.assertNotIn('token', run.stdout + run.stderr)
        self.git('remote', 'remove', 'origin')
        self.assertNotEqual(0, self.cli('status').returncode)

    def test_credentials_are_never_printed(self):
        self.git('remote', 'set-url', 'origin', 'https://secret:token@github.com/stwpst10-collab/teoo.git')
        data = self.result('status')
        self.assertEqual(ORIGIN, data['remote'])
        self.assertNotIn('secret', json.dumps(data))

    def test_search_metadata_bilingual_and_exclusions(self):
        self.write('10-지식/제품/다른 이름.md', '# Ovira 다른 사례\nOvira')
        for p in ('90 작업자료/old.md', '30-근거/첨부/old.md',
                  '10-지식/.local/hidden.md', '10-지식/node_modules/old.md',
                  '05-협업/private.md'):
            self.write(p, '# Ovira\nOvira')
        data = self.result('search', 'oViRa')
        self.assertEqual(2, len(data['results']))
        item = next(r for r in data['results'] if r['id'] == 'ALT-002')
        self.assertEqual('판단유보', item['status'])
        self.assertEqual('pending', item['review'])
        self.assertTrue(item['line'] >= 1)
        self.assertEqual(2, len(self.result('search', '오비라')['results']))
        self.assertEqual(1, len(self.result('search', 'Ovira', '--limit', '1')['results']))

    def test_read_and_truncation(self):
        data = self.result('read', '10-지식/제품/오비라.md', '--max-chars', '20')
        self.assertTrue(data['truncated'])
        self.assertEqual(20, len(data['content']))
        self.assertEqual('ALT-002', data['id'])

    def test_read_rejects_traversal_excluded_and_non_markdown(self):
        self.write('10-지식/data.txt', 'hidden')
        self.write('30-근거/첨부/raw.md', 'hidden')
        for path in ('../secret.md', '10-지식/../10-지식/제품/오비라.md',
                     str(self.repo / '10-지식/제품/오비라.md'),
                     '30-근거/첨부/raw.md', '10-지식/data.txt'):
            with self.subTest(path=path):
                self.assertNotEqual(0, self.cli('read', path).returncode)

    def test_external_symlink_is_rejected_and_not_searched(self):
        outside = self.base / 'outside.md'
        outside.write_text('# SECRET Ovira', encoding='utf-8')
        link = self.repo / '10-지식/linked.md'
        try:
            link.symlink_to(outside)
        except OSError:
            self.skipTest('Creating symlinks is not supported for this user')
        self.assertNotEqual(0, self.cli('read', '10-지식/linked.md').returncode)
        self.assertEqual([], self.result('search', 'SECRET')['results'])

    def test_discovery_ancestors_workspace_and_venture_child(self):
        self.assertEqual(ORIGIN, self.result('status', explicit=False)['remote'])
        self.assertEqual(ORIGIN, self.result('status', explicit=False, cwd=self.repo / '10-지식')['remote'])
        installed = self.base / 'installed/scripts/wiki.py'
        installed.parent.mkdir(parents=True)
        shutil.copyfile(SCRIPT, installed)
        (installed.parent.parent / 'workspace.json').write_text(
            json.dumps({'repository_root': str(self.repo)}), encoding='utf-8')
        unrelated = self.base / 'unrelated'
        unrelated.mkdir()
        self.assertEqual(ORIGIN, self.result('status', explicit=False, cwd=unrelated, script=installed)['remote'])

    def snapshot(self, remote=ORIGIN):
        folder = self.base / 'snapshot'
        shutil.copytree(self.repo, folder, ignore=shutil.ignore_patterns('.git'))
        (folder / 'teoo-snapshot.json').write_text(json.dumps(
            {'remote': remote, 'commit': 'abc1234', 'branch': 'main', 'exported': '2026-10-29T00:00:00Z'}),
            encoding='utf-8')
        return folder

    def test_snapshot_without_git_reports_manifest_and_reads(self):
        folder = self.snapshot()
        run = self.cli('status', explicit=False, cwd=self.base)
        data = json.loads(self.cli('--repo', str(folder), 'status', explicit=False).stdout)
        self.assertEqual(('snapshot', 'abc1234', '2026-10-29T00:00:00Z'),
                         (data['source'], data['commit'], data['exported']))
        self.assertEqual(0, run.returncode)  # sibling Git checkout is still discovered
        found = json.loads(self.cli('--repo', str(folder), 'search', '오비라', explicit=False).stdout)
        self.assertEqual('ALT-002', found['results'][0]['id'])
        self.assertEqual('pending', found['results'][0]['review'])

    def test_snapshot_for_other_repository_is_rejected(self):
        folder = self.snapshot('https://github.com/elsewhere/private.git')
        run = self.cli('--repo', str(folder), 'read', '10-지식/제품/오비라.md', explicit=False)
        self.assertNotEqual(0, run.returncode)
        self.assertNotIn('80명', run.stdout + run.stderr)
        (folder / 'teoo-snapshot.json').unlink()
        self.assertNotEqual(0, self.cli('--repo', str(folder), 'status', explicit=False).returncode)

    def test_git_status_reports_commit(self):
        self.assertEqual('none', self.result('status')['commit'])
        self.git('add', '-A')
        self.git('-c', 'user.name=t', '-c', 'user.email=t@example.com', 'commit', '-qm', 'init')
        data = self.result('status')
        self.assertEqual(('git', False), (data['source'], data['dirty']))
        self.assertRegex(data['commit'], r'^[0-9a-f]{7,}$')


if __name__ == '__main__':
    unittest.main()
