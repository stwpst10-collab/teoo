"""해커톤 투자 보고서(워드) 도구.

  python tools/pitch/doc.py build            초안 생성(이미 있으면 --force 필요)
  python tools/pitch/doc.py sync             엑셀 최신본 값으로 연동 차트·표·출처 문구 갱신
  python tools/pitch/doc.py add <차트ID> --after "문구"   해당 문구가 있는 문단 뒤에 차트 추가
  python tools/pitch/doc.py list             문서에 들어간 차트·연동 표 목록

차트는 그림 '대체 텍스트'에 chart:<ID>, 연동 표는 표 제목(대체 텍스트)에 excel:<ID>로 표시한다.
사용자가 워드에서 글을 고치고 저장해도 이 표시는 유지되므로, sync는 사용자가 쓴 글을 건드리지 않고
표시된 차트·표·'자료: 시뮬레이션 모델링' 문단만 바꾼다. 워드가 파일을 열고 있으면 저장하지 않는다.
"""
import argparse
import re
import sys
from datetime import date
from io import BytesIO
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.text.paragraph import Paragraph

sys.path.insert(0, str(Path(__file__).resolve().parent))
import charts  # noqa: E402
import excel_link  # noqa: E402

ROOT = excel_link.ROOT
OUT_DIR = ROOT / '80-산출물' / '해커톤'
DOC = Path(__import__('os').environ.get('PITCH_DOC') or OUT_DIR / '그날의 류 투자 보고서 해커톤용.docx')  # 시험용 경로 지정 가능
SRC_MD = ROOT / '80-산출물' / '투자자 보고서' / '그날의 류 투자자 보고서 2026-11-07.md'
FONT = 'Malgun Gothic'
INK, MUTED, ACCENT = RGBColor(0x0B, 0x0B, 0x0B), RGBColor(0x5F, 0x5F, 0x5C), RGBColor(0x2A, 0x78, 0xD6)
IMG_W = Cm(16)
XL_CAPTION = '자료: 시뮬레이션 모델링'

# 초안에서 차트를 넣을 위치: (차트ID, 이 문구로 시작하는 문단 뒤, 출처 문구)
PLACEMENTS = [
    ('pms_gap', '질병관리청 조사 기준', '자료: 질병관리청(2025), 건강보험심사평가원(2025)'),
    ('market_funnel', '429만~458만명', '자료: 질병관리청, 행정안전부 주민등록인구(2026.10), 당사 추정'),
    ('excel_pl_3y', '최종 가격은', None),
]
PL_TABLE_AFTER = 'excel_pl_3y'


# ---------- 공통 ----------
def set_font(run, size=None, bold=None, color=None):
    run.font.name = FONT
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts'); rpr.insert(0, rfonts)
    for a in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rfonts.set(qn(a), FONT)
    if size: run.font.size = Pt(size)
    if bold is not None: run.font.bold = bold
    if color is not None: run.font.color.rgb = color


def add_inline(par, text, size=None, color=None, bold_all=False):
    """**굵게** 표기를 살려 문단에 글을 넣는다."""
    for i, part in enumerate(re.split(r'\*\*(.+?)\*\*', text)):
        if part:
            set_font(par.add_run(part), size=size, bold=bold_all or i % 2 == 1, color=color)


def insert_after(par):
    p = OxmlElement('w:p')
    par._p.addnext(p)
    return Paragraph(p, par._parent)


def shade(cell, hex_fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hex_fill)
    tcpr.append(shd)


def table_borders(tbl, color='D9D9D6'):
    tblpr = tbl._tbl.tblPr
    b = OxmlElement('w:tblBorders')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement(f'w:{side}')
        e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), '4'); e.set(qn('w:color'), color)
        b.append(e)
    tblpr.append(b)


def tag_table(tbl, tag):
    tblpr = tbl._tbl.tblPr
    for old in tblpr.findall(qn('w:tblCaption')):
        tblpr.remove(old)
    cap = OxmlElement('w:tblCaption'); cap.set(qn('w:val'), tag)
    desc = tblpr.find(qn('w:tblDescription'))
    (desc.addprevious(cap) if desc is not None else tblpr.append(cap))


def table_tag(tbl):
    cap = tbl._tbl.tblPr.find(qn('w:tblCaption'))
    return cap.get(qn('w:val')) if cap is not None else None


def write_cell(cell, text, bold=False, size=9.5, align=None):
    par = cell.paragraphs[0]
    for r in list(par.runs):
        r._element.getparent().remove(r._element)
    add_inline(par, text, size=size, bold_all=bold)
    if align: par.alignment = align


def make_table(after_par, rows, header=True, widths=None):
    tbl = after_par.part.document.add_table(rows=len(rows), cols=len(rows[0]))
    # 문서 끝에 생긴 표를 원하는 위치로 옮긴다
    after_par._p.addnext(tbl._tbl)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_borders(tbl)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = tbl.cell(i, j)
            write_cell(c, str(val), bold=header and i == 0)
            if header and i == 0:
                shade(c, 'EEF4FC')
            if widths:
                c.width = widths[j]
    keep_table_together(tbl)
    return tbl


def keep_table_together(tbl):
    """표가 쪽 사이에서 갈라지지 않게: 행 나눔 금지 + 마지막 행 전까지 다음 단락과 함께."""
    for i, row in enumerate(tbl.rows):
        trpr = row._tr.get_or_add_trPr()
        if trpr.find(qn('w:cantSplit')) is None:
            trpr.append(OxmlElement('w:cantSplit'))
        for c in row.cells:
            for p in c.paragraphs:
                p.paragraph_format.keep_with_next = i < len(tbl.rows) - 1
                p.paragraph_format.space_after = Pt(1)


def caption(par, text):
    par.paragraph_format.space_after = Pt(10)
    add_inline(par, text, size=8.5, color=MUTED)
    return par


def picture(par, png, cid):
    run = par.add_run()
    shape = run.add_picture(BytesIO(png), width=IMG_W)
    docpr = shape._inline.docPr
    docpr.set('descr', f'chart:{cid}')
    docpr.set('title', charts.CHARTS[cid]['title'])
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.keep_with_next = True
    return shape


def xl_caption_text(d):
    r = d['rows']
    # 파일명이 '시뮬레이션 모델링 …'으로 시작해 XL_CAPTION과 겹치므로 '자료: ' 뒤에 파일명만 쓴다
    return (f"자료: {d['file'].name} (엑셀 연동, {date.today():%Y.%m.%d} 갱신). "
            f"가정: 본체 {r['본체'][0]:,.0f}원, 패치 {r['패치'][0]:,.0f}원, 고정비 연 {r['고정비'][0] / 1e8:,.1f}억원")


def pl_rows(d):
    r = d['rows']
    out = [['연차', '본체 판매(대)', '패치 판매(개)', '매출', '매출이익률', '영업이익']]
    for i, y in enumerate(d['years']):
        out.append([f'{y}년차', f"{r['본체 수량'][i]:,.0f}", f"{r['패치 수량'][i]:,.0f}",
                    f"{r['매출 총액'][i] / 1e8:,.2f}억원", f"{r['매익률'][i] * 100:.1f}%", f"{r['영업이익'][i] / 1e8:,.2f}억원"])
    return out


def ensure_closed(path):
    lock = path.with_name('~$' + path.name[2:])
    try:
        with open(path, 'r+b'):
            pass
    except PermissionError:
        sys.exit(f'워드가 "{path.name}"을(를) 열고 있습니다. 저장 후 워드를 닫고 다시 실행하세요.')
    if lock.exists():
        print(f'참고: 잠금 파일 {lock.name} 이 남아 있음(워드가 비정상 종료됐을 수 있음).')


# ---------- build ----------
def styles(doc):
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = Cm(2.2)
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    st = doc.styles['Normal']
    st.font.name = FONT; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.line_spacing = 1.35
    for name, size in (('Title', 26), ('Heading 1', 16), ('Heading 2', 12.5)):
        s = doc.styles[name]
        s.font.name = FONT; s.font.size = Pt(size); s.font.bold = True; s.font.color.rgb = INK
        rpr = s.element.get_or_add_rPr()
        rf = rpr.find(qn('w:rFonts'))
        if rf is None:
            rf = OxmlElement('w:rFonts'); rpr.insert(0, rf)
        for a in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
            rf.set(qn(a), FONT)
        for a in ('w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme'):
            if rf.get(qn(a)) is not None:
                del rf.attrib[qn(a)]
        s.paragraph_format.space_before = Pt(4 if name == 'Heading 1' else 0)
        s.paragraph_format.space_after = Pt(8)
        s.paragraph_format.keep_with_next = True


def build(force=False):
    if DOC.exists() and not force:
        sys.exit(f'이미 있음: {DOC}  (덮어쓰려면 --force, 사용자 수정분이 사라짐)')
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc = Document()
    styles(doc)
    body_first = doc.paragraphs[0] if doc.paragraphs else None
    lines = SRC_MD.read_text(encoding='utf-8').splitlines()
    # 맨 위 안내 문구(">")와 첫 구분선 건너뜀
    start = next(i for i, l in enumerate(lines) if l.startswith('**사내벤처'))
    lines = lines[start:]
    i, first_section = 0, True
    while i < len(lines):
        l = lines[i].rstrip()
        if not l or l == '---':
            i += 1; continue
        if l.startswith('|'):
            block = []
            while i < len(lines) and lines[i].startswith('|'):
                if not re.match(r'^\|[-|\s]+\|$', lines[i]):
                    block.append([c.strip() for c in lines[i].strip('|').split('|')])
                i += 1
            anchor = doc.add_paragraph()
            make_table(anchor, block)
            anchor._p.getparent().remove(anchor._p)
            doc.add_paragraph().paragraph_format.space_after = Pt(2)
            continue
        if l.startswith('# '):
            p = doc.add_paragraph(style='Title'); add_inline(p, l[2:], size=26, bold_all=True)
        elif l.startswith('## '):
            p = doc.add_paragraph(style='Heading 1'); add_inline(p, l[3:], size=16, bold_all=True, color=INK)
        elif re.fullmatch(r'\*\*[^*]+\*\*', l):  # 장 라벨(투자포인트 요약, 문제 …)
            cover = l.startswith('**사내벤처')
            p = doc.add_paragraph()
            if not cover and not first_section:
                p.paragraph_format.space_before = Pt(18)
            add_inline(p, l.strip('*'), size=9.5, color=ACCENT, bold_all=True)
            p.paragraph_format.space_after = Pt(0); p.paragraph_format.keep_with_next = True
            if not cover and first_section:
                # 표지 뒤 첫 장 앞에서 쪽 나눔
                p.paragraph_format.page_break_before = True
                first_section = False
        elif l.startswith('- '):
            p = doc.add_paragraph(style='List Bullet'); add_inline(p, l[2:])
        elif l.startswith('자료:'):
            caption(doc.add_paragraph(), l)
        elif re.match(r'^\d+\. ', l):
            p = doc.add_paragraph(); add_inline(p, l, size=8.5, color=MUTED); p.paragraph_format.space_after = Pt(2)
        else:
            p = doc.add_paragraph(); add_inline(p, l)
        i += 1
    if body_first is not None and not body_first.text:
        body_first._p.getparent().remove(body_first._p)
    xl = excel_link.load()
    for cid, starts, src in PLACEMENTS:
        par = next(p for p in doc.paragraphs if p.text.startswith(starts))
        pic = insert_after(par)
        picture(pic, charts.render(cid, xl), cid)
        nxt = pic._p.getnext()
        if nxt is not None and nxt.tag == qn('w:p') and Paragraph(nxt, pic._parent).text.startswith('자료:'):
            continue  # 원문에 이미 출처 문단이 바로 뒤에 있으면 그대로 쓴다
        cap = caption(insert_after(pic), xl_caption_text(xl) if charts.CHARTS[cid]['excel'] else src)
        if cid == PL_TABLE_AFTER:
            tbl = make_table(cap, pl_rows(xl))
            tag_table(tbl, 'excel:pl_table')
    doc.save(DOC)
    print(f'생성: {DOC}\n엑셀: {xl["file"].name}')


# ---------- 각주: 본문 [n] → 쪽 아래 각주, 맨 끝 번호 목록 삭제 ----------
FN_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'
FN_RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes'
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
SOURCES_HEAD = '자료 및 유의사항'
CITE = re.compile(r'\s?\[(\d+)\]')


def footnotes_part(doc):
    from docx.opc.packuri import PackURI
    from docx.opc.part import XmlPart
    from docx.oxml import parse_xml
    for rel in doc.part.rels.values():
        if rel.reltype == FN_RT:
            return rel.target_part
    xml = (f'<w:footnotes xmlns:w="{W_NS}">'
           '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>'
           '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
           '</w:footnotes>')
    part = XmlPart(PackURI('/word/footnotes.xml'), FN_CT, parse_xml(xml), doc.part.package)
    doc.part.relate_to(part, FN_RT)
    return part


def _fn_run(text=None, ref=None, size=None, mark=False):
    r = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    rf = OxmlElement('w:rFonts')
    for a in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rf.set(qn(a), FONT)
    rpr.append(rf)
    if not mark and ref is None:
        col = OxmlElement('w:color'); col.set(qn('w:val'), '5F5F5C'); rpr.append(col)
    if size:
        sz = OxmlElement('w:sz'); sz.set(qn('w:val'), str(int(size * 2))); rpr.append(sz)
    if ref is not None or mark:
        va = OxmlElement('w:vertAlign'); va.set(qn('w:val'), 'superscript'); rpr.append(va)
    r.append(rpr)
    if mark:
        r.append(OxmlElement('w:footnoteRef'))
    elif ref is not None:
        e = OxmlElement('w:footnoteReference'); e.set(qn('w:id'), str(ref)); r.append(e)
    else:
        t = OxmlElement('w:t'); t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve'); t.text = text; r.append(t)
    return r


def _all_paragraphs(doc):
    yield from doc.paragraphs
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                yield from c.paragraphs


def footnotes():
    ensure_closed(DOC)
    doc = Document(DOC)
    paras = doc.paragraphs
    head = next((i for i, p in enumerate(paras) if p.text.strip() == SOURCES_HEAD), None)
    sources = {}
    if head is not None:
        for p in paras[head + 1:]:
            m = re.match(r'^(\d+)\.\s*(.+)', p.text.strip())
            if m:
                sources[int(m.group(1))] = m.group(2)
                p._p.getparent().remove(p._p)
        hp = paras[head]
        if sources and hp.runs:
            hp.runs[0].text = '유의사항'
            for r in hp.runs[1:]:
                r.text = ''
    store = Path(__file__).with_name('sources.json')
    import json
    if sources:  # 목록을 지운 뒤에도 새 [n]을 각주로 바꿀 수 있게 보관
        store.write_text(json.dumps(sources, ensure_ascii=False, indent=1), encoding='utf-8')
    elif store.exists():
        sources = {int(k): v for k, v in json.loads(store.read_text(encoding='utf-8')).items()}
    if not sources:
        sys.exit('출처 목록을 찾지 못함')
    part = footnotes_part(doc)
    root = part.element
    ids = [int(f.get(qn('w:id'))) for f in root.findall(qn('w:footnote'))]
    nid = max(ids + [0]) + 1
    made, missing = 0, set()
    for p in _all_paragraphs(doc):
        for r in list(p.runs):
            t = r.text
            if not CITE.search(t):
                continue
            pieces = CITE.split(t)  # [글, 번호, 글, 번호, 글…]
            anchor = r._element
            for k, piece in enumerate(pieces):
                if k % 2 == 0:
                    if not piece:
                        continue
                    nr = __import__('copy').deepcopy(r._element)
                    for child in list(nr):
                        if child.tag != qn('w:rPr'):
                            nr.remove(child)
                    tt = OxmlElement('w:t'); tt.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve'); tt.text = piece
                    nr.append(tt)
                else:
                    n = int(piece)
                    if n not in sources:
                        missing.add(n); continue
                    fn = OxmlElement('w:footnote'); fn.set(qn('w:id'), str(nid))
                    fp = OxmlElement('w:p')
                    ppr = OxmlElement('w:pPr'); sp = OxmlElement('w:spacing')
                    sp.set(qn('w:after'), '0'); sp.set(qn('w:line'), '240'); sp.set(qn('w:lineRule'), 'auto')
                    ppr.append(sp); fp.append(ppr)
                    fp.append(_fn_run(mark=True, size=8))
                    fp.append(_fn_run(' ' + sources[n], size=8))
                    fn.append(fp); root.append(fn)
                    nr = _fn_run(ref=nid)
                    nid += 1; made += 1
                anchor.addnext(nr); anchor = nr
            r._element.getparent().remove(r._element)
    doc.save(DOC)
    print(f'각주 {made}개 생성, 맨 끝 출처 목록 {len(sources) if head is not None else 0}줄 정리'
          + (f', 목록에 없는 번호: {sorted(missing)}' if missing else ''))


# ---------- sync / add / list ----------
def iter_charts(doc):
    for shape in doc.inline_shapes:
        descr = shape._inline.docPr.get('descr') or ''
        if descr.startswith('chart:'):
            yield descr[6:], shape


def sync():
    ensure_closed(DOC)
    doc = Document(DOC)
    xl = excel_link.load()
    done = []
    for cid, shape in iter_charts(doc):
        if cid not in charts.CHARTS:
            print(f'경고: 등록되지 않은 차트 {cid} (건너뜀)'); continue
        rid = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
        doc.part.related_parts[rid]._blob = charts.render(cid, xl)
        done.append(cid)
    for tbl in doc.tables:
        if table_tag(tbl) == 'excel:pl_table':
            rows = pl_rows(xl)
            while len(tbl.rows) < len(rows):
                tbl.add_row()
            while len(tbl.rows) > len(rows):
                tbl._tbl.remove(tbl.rows[-1]._tr)
            for i, row in enumerate(rows):
                for j, val in enumerate(row):
                    write_cell(tbl.cell(i, j), val, bold=i == 0)
            done.append('표:pl_table')
    for p in doc.paragraphs:
        if p.text.startswith(XL_CAPTION):
            for r in list(p.runs):
                r._element.getparent().remove(r._element)
            add_inline(p, xl_caption_text(xl), size=8.5, color=MUTED)
    doc.save(DOC)
    print(f'갱신: {", ".join(done) or "없음"}  (엑셀: {xl["file"].name})')


def add(cid, after, source=None):
    if cid not in charts.CHARTS:
        sys.exit(f'없는 차트 ID: {cid}. 등록된 ID: {", ".join(charts.CHARTS)}')
    ensure_closed(DOC)
    doc = Document(DOC)
    hits = [p for p in doc.paragraphs if after in p.text]
    if not hits:
        sys.exit(f'"{after}" 문구가 있는 문단을 찾지 못함')
    xl = excel_link.load() if charts.CHARTS[cid]['excel'] else None
    pic = insert_after(hits[0])
    picture(pic, charts.render(cid, xl), cid)
    caption(insert_after(pic), xl_caption_text(xl) if xl else (source or '자료: 당사'))
    doc.save(DOC)
    print(f'추가: {cid} → "{hits[0].text[:30]}…" 뒤')


def list_items():
    doc = Document(DOC)
    for cid, s in iter_charts(doc):
        print(f'차트 {cid:16s} 엑셀연동={charts.CHARTS.get(cid, {}).get("excel")}  {charts.CHARTS.get(cid, {}).get("title")}')
    for t in doc.tables:
        if table_tag(t):
            print(f'표   {table_tag(t)}')


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest='cmd', required=True)
    b = sub.add_parser('build'); b.add_argument('--force', action='store_true')
    sub.add_parser('sync'); sub.add_parser('list'); sub.add_parser('footnotes')
    a = sub.add_parser('add'); a.add_argument('cid'); a.add_argument('--after', required=True); a.add_argument('--source')
    args = ap.parse_args()
    {'build': lambda: build(args.force), 'sync': sync, 'list': list_items, 'footnotes': footnotes,
     'add': lambda: add(args.cid, args.after, args.source)}[args.cmd]()
