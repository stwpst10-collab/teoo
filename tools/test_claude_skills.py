"""Check the Claude Code skill, the Claude web app ZIP, and snapshot export."""
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import unittest
import zipfile


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'tools'))
import build_web_skill  # noqa: E402


def frontmatter(path):
    text = path.read_text(encoding='utf-8')
    block = re.match(r'\A---\n(.*?)\n---\n', text, re.S)[1]
    return dict(re.findall(r'^([\w-]+):\s*(.*)$', block, re.M)), text


class ClaudeSkillTests(unittest.TestCase):
    def test_all_skill_entries_share_name_and_rules(self):
        for path in (ROOT / '.agents/skills/teoo-wiki/SKILL.md', ROOT / '.claude/skills/teoo-wiki/SKILL.md',
                     ROOT / 'tools/claude_web_skill/SKILL.md'):
            with self.subTest(path=path.relative_to(ROOT).as_posix()):
                fields, text = frontmatter(path)
                self.assertEqual('teoo-wiki', fields['name'])
                self.assertTrue(0 < len(fields['description']) <= 200)
                self.assertIn('rules.md', text)

    def test_claude_code_skill_points_at_existing_shared_files(self):
        skill = ROOT / '.claude/skills/teoo-wiki/SKILL.md'
        _, text = frontmatter(skill)
        for link in re.findall(r'\]\(([^)]+)\)', text):
            if not link.startswith('http'):
                self.assertTrue((skill.parent / link).resolve().is_file(), link)
        for command in re.findall(r'python ([\w./-]+\.py)', text):
            self.assertTrue((ROOT / command).is_file(), command)

    def test_web_zip_is_current_and_complete(self):
        data = build_web_skill.build()
        self.assertEqual(data, build_web_skill.OUTPUT.read_bytes(),
                         'Run python tools/build_web_skill.py and commit the ZIP.')
        with zipfile.ZipFile(build_web_skill.OUTPUT) as bundle:
            names = set(bundle.namelist())
            self.assertEqual(set(build_web_skill.FILES), names)
            self.assertIn('teoo-wiki/SKILL.md', names)
            self.assertTrue(all(name.startswith('teoo-wiki/') for name in names))
            text = '\n'.join(bundle.read(name).decode('utf-8') for name in names)
        for forbidden in ('workspace.json"', 'ghp_', 'github_pat_', 'warrenbff.git'):
            self.assertNotIn(forbidden, text)

    def test_snapshot_export_runs_with_web_skill_helper(self):
        with tempfile.TemporaryDirectory() as temp:
            out = Path(temp) / 'snapshot.zip'
            run = subprocess.run([sys.executable, str(ROOT / 'tools/export_snapshot.py'), '--out', str(out)],
                                 capture_output=True, encoding='utf-8',
                                 env={**os.environ, 'PYTHONIOENCODING': 'utf-8'})
            self.assertEqual(0, run.returncode, run.stderr)
            with zipfile.ZipFile(out) as bundle:
                names = bundle.namelist()
                self.assertFalse([n for n in names if '/.local/' in n or n.endswith('workspace.json')])
                bundle.extractall(Path(temp) / 'x')
            web = Path(temp) / 'skill'
            with zipfile.ZipFile(build_web_skill.OUTPUT) as bundle:
                bundle.extractall(web)
            helper = [sys.executable, str(web / 'teoo-wiki/scripts/wiki.py'), '--repo', str(Path(temp) / 'x/teoo')]
            status = subprocess.run(helper + ['status'], capture_output=True, encoding='utf-8', cwd=temp,
                                    env={**os.environ, 'PYTHONIOENCODING': 'utf-8'})
            self.assertEqual(0, status.returncode, status.stderr)
            self.assertEqual('snapshot', json.loads(status.stdout)['source'])


if __name__ == '__main__':
    unittest.main()
