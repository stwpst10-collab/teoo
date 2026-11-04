"""Build the single-file web report (GitHub Pages) from the same manuscript as the docx.

Inputs
  report_content.py   chapters, bullets, tables, figures (same blocks the docx uses)
  build_fullreport.py cover thesis, key metrics, chapter summaries (read with ast, never executed)
  chart_data.py       chart numbers -> inline SVG (no matplotlib, no fonts needed)
  web_figures.py      web-only figures (TAM/SAM/SOM funnel, price comparison, measurement checklist)
  share_notes.md      human summary per share-* tag
  git share-* tags    previous shared state for the "이번 업데이트" box and section badges

Usage: python tools/report/build_html.py [--out _site/index.html]
"""
import argparse
import ast
import datetime as dt
import difflib
import html
import re
import subprocess
import sys
from collections import Counter
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
CONTENT = 'tools/report/report_content.py'
BUILD = 'tools/report/build_fullreport.py'
CHART_DATA = 'tools/report/chart_data.py'
LEDGER = '10-지식/시장/시장 수치 대장.md'
NOTES = HERE / 'share_notes.md'
MARK_START, MARK_END = '# ================= 1. Investment Highlights', 'doc.save(OUT)'

sys.path.insert(0, str(HERE))


# ---------------------------------------------------------------- git helpers
def git(*args, root=ROOT):
    try:
        r = subprocess.run(['git', *args], cwd=root, capture_output=True, text=True, encoding='utf-8')
    except FileNotFoundError:
        return None
    return r.stdout if r.returncode == 0 else None


def git_show(ref, path, root=ROOT):
    return git('show', f'{ref}:{path}', root=root)


def share_tags(root=ROOT):
    out = git('tag', '-l', 'share-*', '--sort=creatordate', root=root) or ''
    return [t for t in out.split() if t]


def tag_date(tag, root=ROOT):
    return (git('log', '-1', '--format=%cs', tag, root=root) or '').strip()


def pick_compare(tags, head_tags):
    """Return (current_label, previous_tag). On a share tag build, compare with the tag before it."""
    current = next((t for t in reversed(tags) if t in head_tags), None)
    if current:
        i = tags.index(current)
        return current, (tags[i - 1] if i > 0 else None)
    return None, (tags[-1] if tags else None)


# ---------------------------------------------------------------- manuscript loading
class _ChartRef(str):
    def __truediv__(self, name):
        return name


def load_report(content_src, build_src=None, chart_src=None):
    """Execute the manuscript body with recording stand-ins and return a report dict."""
    from changes import load_changes
    rep = {'chapters': [], 'meta': _build_meta(build_src) if build_src else {}, 'charts': _chart_data(chart_src),
           'changes': load_changes(content_src)}
    cur = {'ch': None, 'sec': None}

    def section(title):
        ch = cur['ch']
        sec = {'title': title, 'blocks': []}
        ch['sections'].append(sec); cur['sec'] = sec
        return sec

    def h1(t):
        ch = {'title': t, 'sections': []}
        rep['chapters'].append(ch); cur['ch'] = ch
        section('')

    def s1(t):
        section(t)

    def push(block):
        if cur['ch'] is None:
            h1('')
        cur['sec']['blocks'].append(block)

    def s2(*ts):
        for t in ts:
            push({'kind': 'b0', 'text': t})

    def s3(*ts):
        for t in ts:
            push({'kind': 'b1', 'text': t})

    def table(title, header, rows, src, widths=None):
        push({'kind': 'table', 'title': title, 'header': list(header), 'rows': [list(r) for r in rows], 'src': src})

    def figure(title, png, src, width_cm=15.0):
        push({'kind': 'figure', 'title': title, 'chart': str(png), 'src': src})

    env = {'h1': h1, 'h2': s1, 's1': s1, 's2': s2, 's3': s3, 'blank': lambda: None,
           'add': lambda kind, text='': {'s1': s1, 's2': s2, 's3': s3, 'h1': h1, 'h2': s1}[kind](text),
           'cap': lambda t: t, 'table': table, 'figure': figure, 'CHARTS': _ChartRef('charts'),
           'doc': None, 'OUT': None}
    start = content_src.find(MARK_START)
    end = content_src.find(MARK_END, max(start, 0))
    body = content_src[start if start >= 0 else 0:end if end >= 0 else None]
    exec(compile(body, 'report_content', 'exec'), env)  # repository manuscript, same as the docx build
    return rep


def _literal(node):
    try:
        return ast.literal_eval(node)
    except Exception:
        return None


def _build_meta(src):
    """Read cover/summary literals from build_fullreport.py without running it."""
    meta = {'summary': {}, 'metrics': [], 'thesis': '', 'notice': '', 'subtitle': '', 'date': ''}
    tree = ast.parse(src)
    for node in ast.walk(tree):
        if isinstance(node, ast.Assign) and any(getattr(t, 'id', '') == 'SUMMARY' for t in node.targets):
            meta['summary'] = _literal(node.value) or {}
        if isinstance(node, ast.FunctionDef) and node.name == 'overview':
            for c in ast.walk(node):
                if isinstance(c, ast.Call) and getattr(c.func, 'id', '') == 'kv_table' and c.args:
                    meta['metrics'] = _literal(c.args[0]) or []
        if isinstance(node, ast.Call):
            name = getattr(node.func, 'id', '')
            first = _literal(node.args[0]) if node.args else None
            if name == 'callout' and any(k.arg == 'label' and _literal(k.value) == '한 문장 논지' for k in node.keywords):
                meta['thesis'] = first or ''
            if name == 'runs' and len(node.args) > 1:
                s = _literal(node.args[1])
                if isinstance(s, str) and s.startswith('본 자료는'):
                    meta['notice'] = s
                if isinstance(s, str) and s.startswith('손을 멈추지 않는'):
                    meta['subtitle'] = s
            if name == 'kv_table' and isinstance(first, list) and first and first[0][0] == '작성일':
                meta['date'] = first[0][1]
    return meta


def _chart_data(src):
    if not src:
        return {}
    env = {}
    exec(compile(src, 'chart_data', 'exec'), env)
    return env.get('CHARTS', {})


def load_ledger(src):
    rows = {}
    for line in (src or '').splitlines():
        m = re.match(r'\|\s*(N\d{3})\s*\|', line)
        if m:
            cells = [c.strip() for c in line.strip().strip('|').split('|')]
            if len(cells) >= 9:
                rows[m.group(1)] = {'name': cells[1], 'value': cells[2], 'unit': cells[3], 'year': cells[4], 'grade': cells[8]}
    return rows


def load_notes(path=NOTES):
    notes, cur = {}, None
    if not Path(path).exists():
        return notes
    for line in Path(path).read_text(encoding='utf-8').splitlines():
        m = re.match(r'##\s+(\S+)', line)
        if m:
            cur = m.group(1); notes[cur] = []
        elif cur and line.strip().startswith('- '):
            notes[cur].append(line.strip()[2:])
    return notes


# ---------------------------------------------------------------- diff
NUM = re.compile(r'\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?')


def section_units(rep):
    """{(chapter, section): [text units]} including summaries, key metrics and chart data."""
    out = {}
    meta = rep.get('meta', {})
    for ch in rep['chapters']:
        for sec in ch['sections']:
            units = []
            if sec['title'] == '' and ch['title'] in meta.get('summary', {}):
                units.append('요약: ' + meta['summary'][ch['title']])
            if sec['title'] == '' and ch['title'] == 'Investment Highlights' and meta.get('metrics'):
                units += [f'{k}: {v}' for k, v in meta['metrics']]
            for b in sec['blocks']:
                if b['kind'] in ('b0', 'b1'):
                    units.append(b['text'])
                elif b['kind'] == 'table':
                    units.append('표: ' + b['title'])
                    units += [' | '.join(map(str, r)) for r in b['rows']]
                    units.append(b['src'])
                elif b['kind'] == 'figure':
                    d = rep.get('charts', {}).get(b['chart'], {})
                    units.append('그림: ' + b['title'] + ' ' + repr(d.get('values') or d.get('series')))
                    units.append(b['src'])
            out[(ch['title'], sec['title'])] = units
    return out


def number_changes(old_units, new_units, limit=8):
    changes = []
    sm = difflib.SequenceMatcher(a=old_units, b=new_units, autojunk=False)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op != 'replace':
            continue
        for a, b in zip(old_units[i1:i2], new_units[j1:j2]):
            na, nb = Counter(NUM.findall(a)), Counter(NUM.findall(b))
            gone, came = list((na - nb).elements())[:5], list((nb - na).elements())[:5]
            if not gone and not came:
                continue
            if gone:
                change = f'{", ".join(gone)} → {", ".join(came) or "삭제"}'
            else:
                change = '추가 ' + ', '.join(came)
            changes.append({'context': re.sub(r'\*\*', '', b)[:40], 'change': change})
    return changes[:limit]


def diff_reports(old, new):
    """Section-level comparison: {'new': [...], 'modified': [...], 'removed': [...], 'status': {key: '신규'|'수정'}}."""
    if old is None:
        return None
    a, b = section_units(old), section_units(new)
    res = {'new': [], 'modified': [], 'removed': [k for k in a if k not in b], 'status': {}}
    for k, units in b.items():
        if k not in a:
            res['new'].append(k); res['status'][k] = '신규'
        elif a[k] != units:
            res['modified'].append({'key': k, 'numbers': number_changes(a[k], units)})
            res['status'][k] = '수정'
    return res


def diff_ledger(old, new):
    if old is None:
        return None
    out = {'new': [], 'changed': []}
    for k, r in new.items():
        if k not in old:
            out['new'].append((k, r))
        elif (old[k]['value'], old[k]['grade']) != (r['value'], r['grade']):
            out['changed'].append((k, old[k], r))
    return out


# ---------------------------------------------------------------- rendering helpers
GRADE = re.compile(r'(?<=[(\s,·/])([A-D])(?=[),·/]|\s*\)|$)|(?<=등급 )([A-D])(?![A-Za-z0-9])')


def fmt(text):
    """Escape, apply **bold**, and turn grade letters into colored badges."""
    s = html.escape(str(text))
    s = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', s)
    return GRADE.sub(lambda m: f'<span class="g g{(m.group(1) or m.group(2))}">{m.group(1) or m.group(2)}</span>', s)


def slug(*parts):
    return 's-' + re.sub(r'[^0-9A-Za-z가-힣]+', '-', '-'.join(p for p in parts if p)).strip('-')


BLUE, ORANGE, INK, MUTED, GRID = '#2a78d6', '#eb6834', '#1a1a1a', '#5f5f5c', '#e4e4e0'


def _t(x, y, s, anchor='start', size=11, color=INK, weight='normal'):
    return (f'<text x="{x:.1f}" y="{y:.1f}" text-anchor="{anchor}" font-size="{size}" fill="{color}" '
            f'font-weight="{weight}">{html.escape(str(s))}</text>')


def _num(v, unit='', decimals=1):
    return (f'{v:,.{decimals}f}' if v % 1 else f'{v:,.0f}') + unit


def svg_chart(d):
    kind, W = d['kind'], 480
    lines = wrap(d.get('note', ''), 40)
    note = ''.join(_t(0, 14 + 15 * i, ln, size=11, color=MUTED) for i, ln in enumerate(lines))
    head = 15 * len(lines) + 4
    series = d.get('series') or [('', d.get('values'), None)]
    legend = []
    if len(series) > 1:
        for j, (name, _, _) in enumerate(series):
            legend += [f'<rect x="{j * 110}" y="{head + 1}" width="10" height="10" fill="{(BLUE, ORANGE)[j]}"/>',
                       _t(14 + j * 110, head + 10, name, size=10.5)]
        head += 18
    if kind == 'hbar':
        n, lab_w, top, row = len(d['labels']), 150, head + 6, 30
        H = top + n * row + 6
        bw = W - lab_w - 50
        parts = [note]
        for i, (lab, v) in enumerate(zip(d['labels'], d['values'])):
            y = top + i * row
            w = bw * v / d['vmax']
            parts += [_t(lab_w - 8, y + 15, lab, 'end'),
                      f'<rect x="{lab_w}" y="{y + 4}" width="{w:.1f}" height="16" fill="{BLUE}" rx="2"/>',
                      _t(lab_w + w + 6, y + 16, f'{v:g}{d["unit"]}')]
        return _svg(W, H, parts, d.get('note', ''))
    labels = d['labels']
    n = len(labels)
    top, base_h, bottom = head + 16, 170, 40
    lo, hi = d.get('vmin', 0), d['vmax']
    y0 = top + base_h * hi / (hi - lo)
    sy = lambda v: top + base_h * (hi - v) / (hi - lo)  # noqa: E731
    slot = (W - 20) / n
    parts = [note, *legend, f'<line x1="0" y1="{y0:.1f}" x2="{W}" y2="{y0:.1f}" stroke="{MUTED}" stroke-width="1"/>']
    colors = [BLUE, ORANGE]
    k = len(series)
    bw = min(slot * 0.5 / k * (1.3 if k > 1 else 1), 46)
    for i, lab in enumerate(labels):
        cx = 10 + slot * (i + 0.5)
        for j, (name, vals, texts) in enumerate(series):
            v = vals[i]
            x = cx - k * bw / 2 + j * bw
            y1, y2 = sorted((sy(v), y0))
            parts.append(f'<rect x="{x:.1f}" y="{y1:.1f}" width="{bw - 2:.1f}" height="{max(y2 - y1, 0.5):.1f}" fill="{colors[j % 2]}" rx="1.5"/>')
            txt = texts[i] if texts else (_num(v, d.get('unit', ''), d.get('decimals', 1)) if k == 1 else f'{v:,}')
            parts.append(_t(x + bw / 2 - 1, (y1 - 4) if v >= 0 else (y2 + 12), txt, 'middle', 10))
        for li, line in enumerate(str(lab).split('\n')):
            parts.append(_t(cx, top + base_h + 16 + li * 13, line, 'middle', 10.5))
    return _svg(W, top + base_h + bottom, parts, d.get('note', ''))


def wrap(text, width):
    """Split a chart note into lines of about `width` characters, breaking at spaces."""
    lines, cur = [], ''
    for word in str(text).split(' '):
        if cur and len(cur) + 1 + len(word) > width:
            lines.append(cur); cur = word
        else:
            cur = f'{cur} {word}' if cur else word
    return lines + ([cur] if cur else [])


def _svg(w, h, parts, title):
    return (f'<svg viewBox="0 0 {w} {h:.0f}" role="img" aria-label="{html.escape(title)}" '
            f'xmlns="http://www.w3.org/2000/svg" font-family="inherit">{"".join(parts)}</svg>')


def html_funnel(f, show_som):
    """Responsive funnel built from divs so text stays readable on phones."""
    rows = []
    for st in f['steps']:
        held = st.get('held') and not show_som
        val = '반영 보류 (재무 결과 미반영)' if held else st['value']
        rows.append(f'<div class="fn{" held" if held else ""}" style="width:{st["width"] * 100:.0f}%">'
                    f'<b>{html.escape(st["name"])} · {fmt(val)}</b><span>{fmt(st["label"])}</span></div>')
    return f'<div class="funnel" role="img" aria-label="{html.escape(f["title"])}">{"".join(rows)}</div>'


def html_prices(p):
    """Price comparison as CSS bars; each group scaled to its own maximum."""
    out = []
    for gname, items in p['groups']:
        gmax = max((hi or v) for _, v, hi, _, _ in items) * 1.08
        out.append(f'<div class="pg">{html.escape(gname)}</div>')
        for lab, v, hi, grade, ours in items:
            cls = 'pb ours' if ours else 'pb'
            if hi:
                bar = (f'<i class="{cls} rng" style="margin-left:{v / gmax * 100:.1f}%;width:{(hi - v) / gmax * 100:.1f}%"></i>')
                txt = f'{v:,}~{hi:,}원'
            else:
                bar = f'<i class="{cls}" style="width:{v / gmax * 100:.1f}%"></i>'
                txt = f'{v:,}원'
            out.append(f'<div class="pr"><div class="pl">{fmt(lab)}</div><div class="pt">{bar}</div>'
                       f'<div class="pv">{txt} <span class="legend">({fmt(grade)})</span></div></div>')
    return f'<div class="prices" role="img" aria-label="{html.escape(p["title"])}">{"".join(out)}</div>'


# ---------------------------------------------------------------- page
CSS = """
:root{--navy:#1f3a5f;--accent:#2a78d6;--ink:#1a1a1a;--muted:#5f6670;--rule:#c9d2de;--panel:#eef3f9;--zebra:#f5f7fa}
*{box-sizing:border-box}html{-webkit-text-size-adjust:100%}
body{margin:0;font-family:"Malgun Gothic","Apple SD Gothic Neo","Noto Sans KR",system-ui,sans-serif;color:var(--ink);line-height:1.65;background:#fafbfc}
main{max-width:880px;margin:0 auto;padding:16px 16px 64px;background:#fff}
header.cover{border-bottom:3px solid var(--navy);padding:18px 0 14px;margin-bottom:12px}
.kicker{color:var(--accent);font-weight:700;font-size:13px;letter-spacing:.02em}
h1.title{font-size:30px;margin:4px 0;color:var(--navy)}
.sub{font-size:16px;margin:0 0 8px}
.meta{color:var(--muted);font-size:13px}
.box{background:var(--panel);border-left:4px solid var(--accent);padding:12px 14px;margin:14px 0;border-radius:4px}
.box h2{margin:0 0 6px;font-size:16px;color:var(--navy)}
.upd{border-left-color:#d9822b;background:#fff7ee}
.chg table{min-width:480px;font-size:13px}.chg tr.latest td{background:#e3ecf7}
h2.ch{font-size:22px;color:var(--navy);border-bottom:2px solid var(--navy);padding-bottom:4px;margin-top:40px}
h3{font-size:16px;margin:22px 0 6px}h3::before{content:"■ ";color:var(--accent)}
ul{padding-left:20px;margin:6px 0}li{margin:3px 0}li.l1{list-style:"– ";margin-left:18px;color:#333}
.tw{overflow-x:auto;-webkit-overflow-scrolling:touch;margin:8px 0}
table{border-collapse:collapse;width:100%;font-size:13.5px;min-width:520px}
th{background:var(--navy);color:#fff;padding:6px 8px;text-align:center;font-weight:600}
td{border-bottom:1px solid var(--rule);padding:6px 8px;vertical-align:top}tr:nth-child(even) td{background:var(--zebra)}
.cap{font-weight:700;color:var(--navy);font-size:14px;margin-top:14px}.cap span{color:var(--accent)}
.src{color:var(--muted);font-size:12.5px;margin:2px 0 10px}
figure{margin:10px 0}figure svg{width:100%;height:auto;background:#fcfcfb;border-radius:4px}
.g{display:inline-block;min-width:1.4em;padding:0 .35em;border-radius:9px;font-size:.78em;font-weight:700;line-height:1.5;text-align:center;vertical-align:.08em}
.gA{background:#1e7d4f;color:#fff}.gB{background:#2a5fb0;color:#fff}.gC{background:#b26b00;color:#fff}.gD{background:#6b6f7a;color:#fff}
.badge{font-size:11px;font-weight:700;padding:1px 7px;border-radius:9px;margin-left:6px;vertical-align:middle}
.b-new{background:#d9822b;color:#fff}.b-mod{background:#fde7cf;color:#8a4b00;border:1px solid #e8b27a}
.kv{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:8px;margin:8px 0}
.kv div{background:var(--panel);padding:8px 10px;border-radius:4px;font-size:13.5px}.kv b{display:block;color:var(--navy);font-size:12.5px}
nav.toc a{color:var(--accent);text-decoration:none}nav.toc ol{list-style:none;padding-left:4px}nav.toc li{margin:2px 0}
.chk li{list-style:none;margin-left:-18px}.chk li::before{content:"☐ ";color:var(--accent)}
.legend{font-size:12.5px;color:var(--muted)}

.funnel{display:flex;flex-direction:column;align-items:center;gap:6px;margin:8px 0}.fn{background:#2a78d6;color:#fff;border-radius:6px;padding:8px 10px;text-align:center;min-width:210px}.fn b{display:block;font-size:15px}.fn span{font-size:12.5px}.fn.held{background:#c9d2de;color:#1f3a5f}.fn .g{background:#fff;color:#1f3a5f}.prices{margin:8px 0}.pg{font-weight:700;margin:8px 0 2px}.pr{display:grid;grid-template-columns:minmax(120px,38%) 1fr auto;gap:8px;align-items:center;font-size:13px;margin:3px 0}.pt{background:#f1f3f6;height:14px;border-radius:3px;overflow:hidden}.pb{display:block;height:14px;background:#2a78d6;border-radius:3px}.pb.ours{background:#eb6834}.pb.rng{opacity:.55}.pv{white-space:nowrap}@media (max-width:560px){.pr{grid-template-columns:1fr auto}.pl{grid-column:1 / -1}}
footer{color:var(--muted);font-size:12px;margin-top:40px;border-top:1px solid var(--rule);padding-top:8px}
@media (max-width:560px){h1.title{font-size:24px}h2.ch{font-size:19px}body{font-size:15px}main{padding:10px 12px 48px}}
"""


def render(rep, diff=None, ledger_diff=None, notes=None, current=None, previous=None, history=(), built=None,
           show_som=None):
    import web_figures as wf
    show_som = wf.SHOW_SOM if show_som is None else show_som
    meta, out = rep.get('meta', {}), []
    status = (diff or {}).get('status', {})
    counters = {'table': 0, 'figure': 0}
    w = out.append
    w('<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">'
      '<meta name="viewport" content="width=device-width,initial-scale=1">'
      f'<title>그날의 류 보고서 (웹)</title><style>{CSS}</style></head><body><main>')
    w('<header class="cover"><div class="kicker">INDUSTRY &amp; COMPANY REPORT | 사내벤처 분석 보고서</div>'
      f'<h1 class="title">그날의 류</h1><p class="sub">{fmt(meta.get("subtitle", ""))}</p>'
      f'<div class="meta">{fmt(meta.get("date", ""))} · 웹판 생성 {html.escape(built or "")}</div></header>')
    if meta.get('thesis'):
        w(f'<div class="box"><h2>한 문장 논지</h2>{fmt(meta["thesis"])}</div>')

    # ---- v8 대비 변경 요약 (누적, 판 단위) — same data as the docx page after the cover
    chg = rep.get('changes') or {}
    if chg.get('rows'):
        w(html_changes(chg))

    # ---- 이번 업데이트 (지난 공유 이후, share 태그 단위) — 판별 내용은 위 표에 있으므로 여기서는 반복하지 않음
    w('<section class="box upd" id="update"><h2>이번 업데이트 <span class="legend">지난 공유 이후</span></h2>')
    label = current or '공유 전 최신본'
    if previous is None:
        w(f'<p><b>첫 공유</b> ({html.escape(label)}). 이전 share 태그가 없어 섹션 비교는 없음. '
          '판별 변경은 위 <a href="#changes">변경 요약</a> 표 참조.</p>')
    else:
        w(f'<p><b>{html.escape(label)}</b> · 비교 기준 <b>{html.escape(previous)}</b>. 판별 변경은 위 <a href="#changes">변경 요약</a> 표 참조.</p>')
    note_key = current or next((k for k in (notes or {}) if k not in history_tags(history)), None)
    if notes and note_key in notes and notes[note_key]:
        w('<ul>' + ''.join(f'<li>{fmt(x)}</li>' for x in notes[note_key]) + '</ul>')
    if diff:
        def link(k):
            return f'<a href="#{slug(*k)}">{html.escape(" › ".join(p for p in k if p))}</a>'
        if diff['new']:
            w('<p><b>신규 섹션</b>: ' + ', '.join(link(k) for k in diff['new']) + '</p>')
        if diff['modified']:
            w('<p><b>수정 섹션</b></p><ul>')
            for m in diff['modified']:
                nums = '; '.join(f'{html.escape(c["context"])}… <b>{html.escape(c["change"])}</b>' for c in m['numbers'])
                w(f'<li>{link(m["key"])}' + (f'<br><span class="legend">바뀐 수치: {nums}</span>' if nums else '') + '</li>')
            w('</ul>')
        if diff['removed']:
            w('<p><b>삭제 섹션</b>: ' + ', '.join(html.escape(" › ".join(p for p in k if p)) for k in diff['removed']) + '</p>')
        if not (diff['new'] or diff['modified'] or diff['removed']):
            w('<p>본문 변경 없음.</p>')
    if ledger_diff and (ledger_diff['new'] or ledger_diff['changed']):
        w('<p><b>시장 수치 대장</b></p><ul>')
        for k, r in ledger_diff['changed']:
            w(f'<li>{k} {html.escape(r["name"])}: {html.escape(r_old_new(r, ledger_diff, k))}</li>')
        for k, r in ledger_diff['new']:
            w(f'<li>{k} {html.escape(r["name"])} (신규): {html.escape(r["value"])} {html.escape(r["unit"])} ({fmt(r["grade"])})</li>')
        w('</ul>')
    w('</section>')

    # ---- toc
    w('<nav class="toc"><b>목차</b><ol>' + ('<li><a href="#changes">변경 요약</a></li>' if chg.get('rows') else '')
      + '<li><a href="#figs">그림으로 보기 (웹 전용)</a></li>')
    for i, ch in enumerate(rep['chapters'], 1):
        w(f'<li><a href="#{slug(ch["title"])}">{i}. {html.escape(ch["title"])}</a></li>')
    w('<li><a href="#history">변경 이력</a></li></ol></nav>')
    w('<p class="legend">근거 등급: <span class="g gA">A</span> 공시·정부 통계·규제 문서·논문 원문 '
      '<span class="g gB">B</span> 리포트·언론 원문 인용 <span class="g gC">C</span> 업체 주장·판매용 요약 '
      '<span class="g gD">D</span> 팀 가설·내부 추정·계획 가정</p>')

    # ---- web-only figures
    w('<section id="figs"><h2 class="ch">그림으로 보기 (웹 전용)</h2>'
      '<p class="legend">docx 보고서에는 없는 요약 그림. 숫자와 등급은 아래 자료 줄과 시장 수치 대장을 따른다.</p>')
    f = wf.FUNNEL
    w(f'<div class="cap">{fmt(f["title"])}</div>{html_funnel(f, show_som)}<ul>')
    for st in f['steps']:
        if st.get('held') and not show_som:
            w(f'<li><b>{st["name"]}</b>: 반영 보류. 재무·시뮬레이션 결과를 보고서에 반영하지 않기로 함(2026-11-04). {fmt(st["note"])}</li>')
        else:
            w(f'<li><b>{st["name"]} {fmt(st["value"])}</b> <span class="g g{st["grade"]}">{st["grade"]}</span> = {fmt(st["calc"])}. {fmt(st["note"])}</li>')
    w(f'</ul><p class="src">{fmt(f["src"])}</p>')
    p = wf.PRICES
    w(f'<div class="cap">{fmt(p["title"])}</div>{html_prices(p)}'
      '<p class="legend">주황 = 그날의 류 계획 가격, 파랑 = 시장 관측값, 옅은 막대 = 범위</p>'
      f'<p class="src">{fmt(p["src"])}</p>')
    c = wf.CHECKLIST
    w(f'<div class="cap">{fmt(c["title"])}</div><ul class="chk">')
    for item, affects, how in c['items']:
        w(f'<li><b>{fmt(item)}</b> — 영향: {fmt(affects)} · 측정: {fmt(how)}</li>')
    w(f'</ul><p class="src">{fmt(c["src"])}</p></section>')

    # ---- chapters
    for i, ch in enumerate(rep['chapters'], 1):
        w(f'<section id="{slug(ch["title"])}"><h2 class="ch">{i}. {html.escape(ch["title"])}'
          f'{badge(status.get((ch["title"], "")))}</h2>')
        if ch['title'] in meta.get('summary', {}):
            w(f'<div class="box"><h2>핵심 요약</h2>{fmt(meta["summary"][ch["title"]])}</div>')
        for sec in ch['sections']:
            if sec['title']:
                w(f'<h3 id="{slug(ch["title"], sec["title"])}">{fmt(sec["title"])}{badge(status.get((ch["title"], sec["title"])))}</h3>')
            if ch['title'] == 'Investment Highlights' and sec['title'] == '' and meta.get('metrics'):
                w('<h3>핵심 지표</h3><div class="kv">' + ''.join(f'<div><b>{fmt(k)}</b>{fmt(v)}</div>' for k, v in meta['metrics']) + '</div>')
            render_blocks(sec['blocks'], rep.get('charts', {}), counters, w)
        w('</section>')

    # ---- history
    w('<section id="history"><h2 class="ch">변경 이력</h2>')
    if not history:
        w('<p>아직 share 태그가 없음. 이 페이지가 첫 공유본이 된다.</p>')
    else:
        w('<div class="tw"><table><tr><th>태그</th><th>날짜</th><th>요약</th><th>섹션 변경</th></tr>')
        for h in reversed(list(history)):
            summ = '<br>'.join(fmt(x) for x in (notes or {}).get(h['tag'], [])) or '—'
            w(f'<tr><td>{html.escape(h["tag"])}</td><td>{html.escape(h["date"])}</td><td>{summ}</td><td>{html.escape(h["changes"])}</td></tr>')
        w('</table></div>')
    w('</section>')
    w(f'<footer>{fmt(meta.get("notice", ""))}<br>원본: tools/report/report_content.py · 생성 {html.escape(built or "")}</footer>')
    w('</main></body></html>')
    return ''.join(out)


def html_changes(chg):
    base = chg['base'].split()[0] if chg['base'] else ''
    out = [f'<section class="box chg" id="changes"><h2>{html.escape(base)} 대비 변경 요약</h2>'
           f'<p class="legend">기준판 {html.escape(chg["base"])} → 최신판 <b>{html.escape(chg["latest"])}</b>. 최신판 변경은 맨 위·굵게.</p>'
           '<div class="tw"><table><tr><th>판</th><th>장</th><th>바뀐 점</th><th>근거 등급</th></tr>']
    for r in chg['rows']:
        cls = ' class="latest"' if r['latest'] else ''
        cell = (lambda x: f'<b>{fmt(x)}</b>') if r['latest'] else fmt
        out.append(f'<tr{cls}><td>{cell(r["ver"])}<br><span class="legend">{html.escape(r["date"])}</span></td>'
                   f'<td>{cell(r["chapter"])}</td><td>{cell(r["what"])}</td><td>{cell(r["grade"])}</td></tr>')
    out.append('</table></div>')
    if chg['common']:
        out.append('<ul class="legend">' + ''.join(f'<li>{fmt(x)}</li>' for x in chg['common']) + '</ul>')
    out.append('</section>')
    if chg.get('questions'):
        out.append('<section class="box chg" id="questions"><h2>검토 중 의문사항</h2>'
                   f'<p class="legend">논리 구조를 다지기 위해 판마다 남기는 열린 질문({html.escape(chg.get("questions_src", ""))} 등). 해결되면 상태 갱신.</p>'
                   '<div class="tw"><table><tr><th>장</th><th>의문</th><th>현재 상태 · 다음 조치</th></tr>')
        for q in chg['questions']:
            out.append(f'<tr><td>{fmt(q["chapter"])}</td><td>{fmt(q["question"])}</td><td>{fmt(q["status"])}</td></tr>')
        out.append('</table></div></section>')
    return ''.join(out)


def r_old_new(r, ld, k):
    old = next(o for kk, o, _ in ld['changed'] if kk == k)
    s = f'{old["value"]} → {r["value"]} {r["unit"]}'
    if old['grade'] != r['grade']:
        s += f' (등급 {old["grade"]} → {r["grade"]})'
    return s


def history_tags(history):
    return {h['tag'] for h in history}


def badge(s):
    if s == '신규':
        return '<span class="badge b-new">신규</span>'
    if s == '수정':
        return '<span class="badge b-mod">수정</span>'
    return ''


def render_blocks(blocks, charts, counters, w):
    open_ul = False
    for b in blocks:
        if b['kind'] in ('b0', 'b1'):
            if not open_ul:
                w('<ul>'); open_ul = True
            w(f'<li class="{"l1" if b["kind"] == "b1" else ""}">{fmt(b["text"])}</li>')
            continue
        if open_ul:
            w('</ul>'); open_ul = False
        if b['kind'] == 'table':
            counters['table'] += 1
            w(f'<div class="cap"><span>[표 {counters["table"]}]</span> {fmt(b["title"])}</div><div class="tw"><table><tr>'
              + ''.join(f'<th>{fmt(h)}</th>' for h in b['header']) + '</tr>')
            for r in b['rows']:
                w('<tr>' + ''.join(f'<td>{fmt(c)}</td>' for c in r) + '</tr>')
            w(f'</table></div><p class="src">{fmt(b["src"])}</p>')
        elif b['kind'] == 'figure':
            counters['figure'] += 1
            d = charts.get(b['chart'])
            body = svg_chart(d) if d else f'<p class="legend">(그림 데이터 없음: {html.escape(b["chart"])})</p>'
            w(f'<div class="cap"><span>[그림 {counters["figure"]}]</span> {fmt(b["title"])}</div>'
              f'<figure>{body}</figure><p class="src">{fmt(b["src"])}</p>')
    if open_ul:
        w('</ul>')


# ---------------------------------------------------------------- assemble
def read(path):
    return (ROOT / path).read_text(encoding='utf-8')


def report_at(ref):
    src = git_show(ref, CONTENT)
    if src is None:
        return None
    return load_report(src, git_show(ref, BUILD), git_show(ref, CHART_DATA))


def summarize(d):
    if d is None:
        return '첫 공유'
    return f'신규 {len(d["new"])} · 수정 {len(d["modified"])} · 삭제 {len(d["removed"])}'


def build(out_path):
    rep = load_report(read(CONTENT), read(BUILD), read(CHART_DATA))
    tags = share_tags()
    head_tags = set((git('tag', '--points-at', 'HEAD') or '').split())
    current, previous = pick_compare(tags, head_tags)
    old = report_at(previous) if previous else None
    diff = diff_reports(old, rep) if old else None
    ledger_diff = diff_ledger(load_ledger(git_show(previous, LEDGER)), load_ledger(read(LEDGER))) if previous else None
    history, prev_rep = [], None
    for t in tags:
        r = report_at(t)
        history.append({'tag': t, 'date': tag_date(t), 'changes': summarize(diff_reports(prev_rep, r) if prev_rep and r else None)})
        prev_rep = r
    sha = (git('rev-parse', '--short', 'HEAD') or '').strip()
    built = f'{dt.datetime.now().strftime("%Y-%m-%d %H:%M")} · {sha}' if sha else dt.datetime.now().strftime('%Y-%m-%d %H:%M')
    page = render(rep, diff, ledger_diff, load_notes(), current, previous, history, built)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(page, encoding='utf-8')
    return out_path


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--out', default=str(ROOT / '_site' / 'index.html'))
    p = build(ap.parse_args().out)
    print('saved', p)
