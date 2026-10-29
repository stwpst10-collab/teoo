"""Regression checks for vault validation; all fixtures are temporary."""

import json
from pathlib import Path
import tempfile
import unittest

from validate import files_in, validate


class VaultValidationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def write(self, name, content=''):
        path = self.root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding='utf-8')
        return path

    def note(self, name='10-지식/SOL-001-제품.md', note_id='SOL-001', body=''):
        return self.write(name, f'---\nid: {note_id}\ntype: solution\nstatus: 가설\nproject: venture\n---\n{body}')

    def test_valid_links_and_archive_target(self):
        self.note(body='[[자료|별칭]]\n![[이미지.png]]\n[자료](<../90 작업자료/자료.md#제목>)')
        self.write('90 작업자료/자료.md', '[[historical-broken-link]]')
        self.write('30-근거/첨부/이미지.png')
        self.assertEqual([], validate(self.root))

    def test_broken_and_ambiguous_wiki_links(self):
        self.write('README.md', '[[missing]] [[same]]')
        self.write('a/same.md')
        self.write('b/same.md')
        errors = validate(self.root)
        self.assertTrue(any('Broken wiki' in e for e in errors))
        self.assertTrue(any('Ambiguous wiki' in e for e in errors))

    def test_duplicate_id_is_rejected_with_korean_filenames(self):
        self.note('10-지식/발목형 웰니스 디바이스.md')
        self.note('10-지식/다른 제품 설명.md')
        errors = validate(self.root)
        self.assertEqual(1, len(errors))
        self.assertIn('Duplicate id SOL-001', errors[0])

    def test_korean_filename_with_spaces_and_stable_metadata_id(self):
        self.note('10-지식/발목형 웰니스 디바이스.md')
        self.write('README.md', '[[발목형 웰니스 디바이스]]\n[제품](<10-지식/발목형 웰니스 디바이스.md>)')
        self.write('사업 관계도.canvas', json.dumps({
            'nodes': [{'type': 'file', 'file': '10-지식/발목형 웰니스 디바이스.md'}],
        }))
        self.assertEqual([], validate(self.root))

    def test_korean_filename_does_not_disable_link_validation(self):
        self.note('10-지식/발목형 웰니스 디바이스.md', body='[[없는 고객 문제]]\n[근거](<없는 근거.md>)')
        errors = validate(self.root)
        self.assertEqual(2, len(errors))
        self.assertTrue(any('Broken wiki link: 없는 고객 문제' in e for e in errors))
        self.assertTrue(any('Broken Markdown link: 없는 근거.md' in e for e in errors))

    def test_url_encoded_and_angle_space_links(self):
        self.write('README.md', '[x](some%20file.txt#anchor) [y](<some file.txt>) [z](https://example.com/a) [m](mailto:a@b.com)')
        self.write('some file.txt')
        self.assertEqual([], validate(self.root))

    def test_conflict_marker(self):
        self.write('README.md', '<<<<<<< HEAD\ntext\n=======\nother\n>>>>>>> branch\n')
        self.assertTrue(any('conflict marker' in e for e in validate(self.root)))

    def test_canvas(self):
        self.write('map.canvas', json.dumps({'nodes': [{'type': 'file', 'file': 'missing.md'}]}))
        self.assertTrue(any('Broken Canvas' in e for e in validate(self.root)))
        self.write('missing.md')
        self.assertEqual([], validate(self.root))

    def test_fenced_and_inline_examples_are_not_links(self):
        self.write('README.md', '```md\n[[missing]] [x](missing.md)\n```\n`[[missing]]`\n~~~\n[[missing]]\n~~~')
        self.assertEqual([], validate(self.root))

    def test_metadata_and_research_requirements(self):
        self.note('20-리서치/R-001-조사.md', 'R-001')
        self.write('20-리서치/README.md', 'Index')
        errors = validate(self.root)
        self.assertEqual(1, len(errors))
        self.assertIn('author', errors[0])
        self.assertIn('source_status', errors[0])

    def test_local_absolute_and_outside_vault_links(self):
        self.write('README.md', '[x](C:/private/file.md) [y](../outside.md) [z](file:///C:/x)')
        errors = validate(self.root)
        self.assertEqual(3, len(errors))
        self.assertTrue(any('leaves vault' in e for e in errors))

    def test_ignored_caches_and_outputs(self):
        self.write('.obsidian/bad.md', '[[missing]]')
        self.write('nested/node_modules/bad.md', '[[missing]]')
        self.write('80-산출물/bad.md', '[[missing]]')
        self.assertEqual([], validate(self.root))

    def test_output_index_is_validated_but_nested_outputs_are_not(self):
        self.write('80-산출물/README.md', '[missing](missing.png)')
        self.write('80-산출물/run/README.md', '[[missing]]')
        errors = validate(self.root)
        self.assertEqual(1, len(errors))
        self.assertIn('80-산출물/README.md: Broken Markdown link', errors[0])

    def test_local_and_temporary_folders_are_not_indexed(self):
        for folder in ('.local', '.tmp', 'tmp', 'nested/tmp'):
            self.write(f'{folder}/secret.md', '[[missing]]')
        self.assertEqual([], files_in(self.root))
        self.assertEqual([], validate(self.root))
        self.write('README.md', '[[secret]]')
        self.assertTrue(any('Broken wiki link: secret' in e for e in validate(self.root)))

    def test_research_status_enums(self):
        path = self.note('20-리서치/R-001-조사.md', 'R-001')
        base = path.read_text(encoding='utf-8').replace(
            'project: venture\n',
            'project: venture\nauthor: researcher\nquestion: Q01\ncreated: 2026-10-29\n',
        )
        for review in ('pending', 'reviewed'):
            for source_status in ('primary', 'secondary', 'mixed', 'unverified'):
                with self.subTest(review=review, source_status=source_status):
                    path.write_text(base.replace('project: venture\n', f'project: venture\nreview: {review}\nsource_status: {source_status}\n'), encoding='utf-8')
                    self.assertEqual([], validate(self.root))
        path.write_text(base.replace('project: venture\n', 'project: venture\nreview: approved\nsource_status: unknown\n'), encoding='utf-8')
        errors = validate(self.root)
        self.assertEqual(2, len(errors))
        self.assertTrue(any('Invalid review' in e for e in errors))
        self.assertTrue(any('Invalid source_status' in e for e in errors))

    def test_template_status_placeholders_are_not_enum_checked(self):
        self.write('00-운영/리서치-양식.md', '---\nreview: pending | reviewed\nsource_status: primary | secondary\n---\n')
        self.assertEqual([], validate(self.root))

    def test_checker_does_not_mutate_files(self):
        path = self.note(body='[[missing]]')
        before = path.read_bytes()
        validate(self.root)
        self.assertEqual(before, path.read_bytes())


if __name__ == '__main__':
    unittest.main()
