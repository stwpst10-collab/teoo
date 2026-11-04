"""'v8 대비 변경 요약' data loader shared by build_fullreport.py (docx page) and build_html.py (web table).

The data lives at the top of report_content.py (before the Investment Highlights marker):
  CHANGES_BASE    fixed base edition, e.g. 'v8 (2026-11-03)'
  CHANGES         list of (판, 날짜, 장, 바뀐 점, 근거 등급), newest edition first
  CHANGES_COMMON  short lines that apply to every edition (shown under the table)
Read with ast so neither builder has to execute anything outside the manuscript body.
"""
import ast

FIELDS = ('ver', 'date', 'chapter', 'what', 'grade')


def load_changes(content_src):
    out = {'base': '', 'rows': [], 'common': [], 'latest': '', 'questions': [], 'questions_src': ''}
    if not content_src:
        return out
    for node in ast.parse(content_src).body:
        if not isinstance(node, ast.Assign):
            continue
        names = {getattr(t, 'id', '') for t in node.targets}
        try:
            val = ast.literal_eval(node.value)
        except ValueError:
            continue
        if 'CHANGES_BASE' in names:
            out['base'] = val
        elif 'CHANGES' in names:
            out['rows'] = [dict(zip(FIELDS, r)) for r in val]
        elif 'CHANGES_COMMON' in names:
            out['common'] = list(val)
        elif 'OPEN_QUESTIONS' in names:
            out['questions'] = [dict(zip(('chapter', 'question', 'status'), r)) for r in val]
        elif 'OPEN_QUESTIONS_SRC' in names:
            out['questions_src'] = val
    if out['rows']:
        out['latest'] = out['rows'][0]['ver']
        for r in out['rows']:
            r['latest'] = r['ver'] == out['latest']
    return out
