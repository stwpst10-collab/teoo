"""Checks for the web report builder (tools/report/build_html.py). No matplotlib needed."""

import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / 'report'))
import build_html as bh  # noqa: E402

MINI = '''# ================= 1. Investment Highlights =================
h1('Investment Highlights')
s1('시장')
s2('국내 성인 여성 41.6%가 심한 월경통(A).', '바닥 인구 약 470만 명(D).')
h1('시장 분석')
s1('규모')
table(cap('표 하나'), ['항목', '값'], [['수입', '3.49억 달러']], '자료: 관세청 (근거 등급 A)')
doc.save(OUT)
'''


class BuildHtmlTests(unittest.TestCase):
    def current(self):
        return bh.load_report(bh.read(bh.CONTENT), bh.read(bh.BUILD), bh.read(bh.CHART_DATA))

    def test_current_report_matches_docx_structure(self):
        rep = self.current()
        titles = [c['title'] for c in rep['chapters']]
        self.assertEqual(titles, ['Investment Highlights', '시장 분석', '사업 분석 및 투자포인트', '경쟁 분석',
                                  '재무계획', '매출추정 및 사업가치', '리스크 및 대응', '결론'])
        self.assertEqual(len(rep['meta']['metrics']), 8)
        self.assertIn('시장 분석', rep['meta']['summary'])
        self.assertTrue(rep['meta']['thesis'].startswith('그날의 류는'))
        blocks = [b for c in rep['chapters'] for s in c['sections'] for b in s['blocks']]
        figures = [b for b in blocks if b['kind'] == 'figure']
        self.assertTrue(figures and all(f['chart'] in rep['charts'] for f in figures), 'every figure has chart data')
        self.assertTrue(all(b['src'].startswith('자료') for b in blocks if b['kind'] in ('table', 'figure')))

    def test_changes_summary_newest_first_and_bold(self):
        chg = self.current()['changes']
        self.assertTrue(chg['base'].startswith('v8'), 'base edition fixed at v8')
        self.assertTrue(chg['rows'] and all(len(r) == 6 for r in chg['rows']))
        latest = chg['rows'][0]['ver']
        self.assertEqual(chg['latest'], latest)
        self.assertIn(f'({latest})', bh.read(bh.BUILD), 'docx header/cover carry the latest edition')
        vers = [r['ver'] for r in chg['rows']]
        self.assertEqual(vers, sorted(vers, key=lambda v: [int(x) for x in v[1:].split('.')], reverse=True))
        page = bh.render(self.current(), built='t')
        self.assertIn('id="changes"', page)
        self.assertEqual(page.count('<tr class="latest">'), sum(r['latest'] for r in chg['rows']))
        self.assertLess(page.index('id="changes"'), page.index('id="update"'))

    def test_first_share_page_is_self_contained(self):
        rep = self.current()
        page = bh.render(rep, built='test')
        self.assertIn('첫 공유', page)
        self.assertNotRegex(page, r'<script|<link|src="http|@import')
        self.assertIn('name="viewport"', page)
        tables = sum(1 for c in rep['chapters'] for s in c['sections'] for b in s['blocks'] if b['kind'] == 'table')
        self.assertEqual(page.count('[표 '), tables)
        self.assertGreaterEqual(page.count('<svg'), 8)
        self.assertIn('반영 보류', page)  # SOM stays held while finance results are not reflected

    def test_grade_badges_only_on_grades(self):
        self.assertIn('class="g gA"', bh.fmt('수입 3.49억 달러 (A)'))
        self.assertIn('class="g gD"', bh.fmt('약 406만 명(19~49세, D)'))
        self.assertIn('class="g gB"', bh.fmt('근거 등급 B'))
        for text in ('A83060.01', 'B2B 채널', 'Class I', 'A안 대신 B안'):
            self.assertNotIn('class="g', bh.fmt(text), text)
        self.assertIn('<strong>굵게</strong>', bh.fmt('**굵게**'))
        self.assertNotIn('<b>', bh.fmt('<b>x</b>'))

    def test_section_diff_and_number_changes(self):
        old = bh.load_report(MINI)
        new_src = MINI.replace('약 470만 명', '약 406만 명').replace(
            "s1('규모')", "s1('규모')\ns2('새 문장')").replace("doc.save", "s1('신규 절')\ns2('추가')\ndoc.save")
        diff = bh.diff_reports(old, bh.load_report(new_src))
        self.assertEqual(diff['new'], [('시장 분석', '신규 절')])
        mods = {m['key']: m['numbers'] for m in diff['modified']}
        self.assertIn(('Investment Highlights', '시장'), mods)
        self.assertEqual(mods[('Investment Highlights', '시장')][0]['change'], '470 → 406')
        self.assertEqual(diff['status'][('시장 분석', '규모')], '수정')
        self.assertIsNone(bh.diff_reports(None, old))

    def test_ledger_diff(self):
        head = '| ID | 지표 | 값 | 단위·통화 | 기준연도 | 정의·범위 | 출처 | 확인일 | 등급 | 메모 |\n'
        old = bh.load_ledger(head + '| N017 | 심한 월경통 여성 | 약 470~500 | 만 명 | 2022 | x | y | z | D | m |')
        new = bh.load_ledger(head + '| N017 | 심한 월경통 여성 | 약 406 | 만 명 | 2026 | x | y | z | D | m |\n'
                             '| N049 | 신규 | 1 | 명 | 2026 | x | y | z | A | m |')
        d = bh.diff_ledger(old, new)
        self.assertEqual([k for k, _ in d['new']], ['N049'])
        self.assertEqual(d['changed'][0][1]['value'], '약 470~500')

    def test_notes_and_tag_choice(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / 'notes.md'
            p.write_text('# 요약\n## share-2026-11-04\n- 첫 줄\n- 둘째 줄\n## share-2026-12-01\n- 다음\n', encoding='utf-8')
            notes = bh.load_notes(p)
        self.assertEqual(notes['share-2026-11-04'], ['첫 줄', '둘째 줄'])
        tags = ['share-2026-11-04', 'share-2026-12-01']
        self.assertEqual(bh.pick_compare(tags, {'share-2026-12-01'}), ('share-2026-12-01', 'share-2026-11-04'))
        self.assertEqual(bh.pick_compare(tags, set()), (None, 'share-2026-12-01'))
        self.assertEqual(bh.pick_compare([], set()), (None, None))

    @unittest.skipUnless(shutil.which('git'), 'git not available')
    def test_git_show_at_share_tag(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            run = lambda *a: subprocess.run(['git', *a], cwd=root, check=True, capture_output=True)  # noqa: E731
            run('init', '-q'); run('config', 'user.email', 't@example.com'); run('config', 'user.name', 't')
            (root / 'a.py').write_text(MINI, encoding='utf-8')
            run('add', '.'); run('commit', '-qm', 'one'); run('tag', 'share-2026-11-04')
            (root / 'a.py').write_text(MINI.replace('470', '406'), encoding='utf-8')
            run('commit', '-qam', 'two')
            self.assertEqual(bh.share_tags(root), ['share-2026-11-04'])
            self.assertIn('470', bh.git_show('share-2026-11-04', 'a.py', root))
            self.assertRegex(bh.tag_date('share-2026-11-04', root), r'^\d{4}-\d{2}-\d{2}$')


if __name__ == '__main__':
    unittest.main()
