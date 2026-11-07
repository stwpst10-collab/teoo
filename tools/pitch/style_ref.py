"""참고 양식(기업분석보고서 '바이오뱅크힐링.docx', 2026-11-07 사용자 지정)의 줄간격·여백·글머리기호를 해커톤 문서에 적용.

  python tools/pitch/style_ref.py        열린 문서(또는 PITCH_DOC)에 적용, 실행 취소 1회로 되돌림 가능

스타일 '글머리 1단계'(▪)·'글머리 2단계'(-)를 만들어 둔다. 워드에서 스타일만 고르면 같은 모양이 된다.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import live  # noqa: E402

CM = live.CM
FONT = live.d.FONT
AUTO = -16777216  # 자동 색(검정)
MULT = 12.9  # 줄간격 배수 1.075 (LineSpacingRule=5일 때 포인트 환산)
SPEC = {
    'margins': (1.2, 1.2, 2.2, 1.5),  # 좌 우 위 아래 (cm)
    'body': dict(size=9, before=0, after=4, left=0.5),
    'h1': dict(size=10, before=12, after=4),
    'h2': dict(size=9.5, before=6, after=2),
    'b1': dict(char=chr(0xF0A7), font='Wingdings', size=9, before=4, after=4, num=0.5, text=1.0),
    'b2': dict(char='-', font=FONT, size=9, before=0, after=0, num=1.1, text=1.6),
    'caption': dict(size=7.5, left=0.5),
    'table': dict(size=8),
    'foot': dict(size=8),
}


def _pf(f, before, after, rule=5, spacing=MULT, align=3):
    f.SpaceBefore = before; f.SpaceAfter = after
    f.LineSpacingRule = rule
    if rule == 5:
        f.LineSpacing = spacing
    f.Alignment = align


def _bullet_style(doc, name, b):
    try:
        st = doc.Styles(name)
    except Exception:
        st = doc.Styles.Add(name, 1)
    st.BaseStyle = doc.Styles(-1).NameLocal
    st.Font.Name = FONT; st.Font.Size = b['size']; st.Font.Bold = False; st.Font.Color = AUTO  # Styles.Add가 선택 서식을 물려받으므로 명시
    pf = st.ParagraphFormat
    _pf(pf, b['before'], b['after'])
    pf.LeftIndent = b['text'] * CM; pf.FirstLineIndent = -(b['text'] - b['num']) * CM
    lt = doc.ListTemplates.Add(False)
    lv = lt.ListLevels(1)
    lv.NumberFormat = b['char']; lv.NumberStyle = 23
    lv.Font.Name = b['font']; lv.Font.Size = b['size']
    lv.NumberPosition = b['num'] * CM; lv.TextPosition = b['text'] * CM; lv.TabPosition = b['text'] * CM
    lv.TrailingCharacter = 0  # 탭
    lv.LinkedStyle = name
    return st, lt


def run():
    app, doc = live.word_doc(open_if_needed=False)
    with live.Undo(app, '참고 양식 적용'):
        ps = doc.PageSetup
        ps.LeftMargin, ps.RightMargin, ps.TopMargin, ps.BottomMargin = [x * CM for x in SPEC['margins']]
        normal = doc.Styles(-1)
        normal.Font.Name = FONT; normal.Font.Size = SPEC['body']['size']
        _pf(normal.ParagraphFormat, 0, SPEC['body']['after'])
        for sid, key in ((-2, 'h1'), (-3, 'h2')):
            st = doc.Styles(sid); h = SPEC[key]
            st.Font.Name = FONT; st.Font.Size = h['size']; st.Font.Bold = True; st.Font.Color = 0
            _pf(st.ParagraphFormat, h['before'], h['after'])
        for sid in (-20, -21):
            st = doc.Styles(sid); st.Font.Size = 9
            _pf(st.ParagraphFormat, 0, 2, rule=0, align=0)
        fs = doc.Styles(-30); fs.Font.Size = SPEC['foot']['size']
        _pf(fs.ParagraphFormat, 0, 0, rule=0, align=3)
        st1, lt1 = _bullet_style(doc, '글머리 1단계', SPEC['b1'])
        st2, lt2 = _bullet_style(doc, '글머리 2단계', SPEC['b2'])

        counts = dict(h1=0, h2=0, bullet=0, body=0, caption=0, table=0)
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            r = p.Range
            if r.Information(12):
                continue
            t = r.Text.strip()
            if p.OutlineLevel == 1:
                r.Font.Size = SPEC['h1']['size']; counts['h1'] += 1
            elif p.OutlineLevel == 2:
                r.Font.Size = SPEC['h2']['size']; counts['h2'] += 1
            elif r.ListFormat.ListType != 0:
                p.Style = st1
                r.ListFormat.ApplyListTemplate(lt1, False, 0)
                r.Font.Size = SPEC['b1']['size']
                r.Font.Color = AUTO
                if r.Text.strip() in ('유의사항',):  # ▪ 소제목은 굵게
                    r.Font.Bold = True
                p.LeftIndent = SPEC['b1']['text'] * CM; p.FirstLineIndent = -(SPEC['b1']['text'] - SPEC['b1']['num']) * CM
                counts['bullet'] += 1
            elif t.startswith('자료:'):
                r.Font.Size = SPEC['caption']['size']
                p.LeftIndent = SPEC['caption']['left'] * CM; p.FirstLineIndent = 0
                _pf(p.Format, 0, 8, rule=0, align=0)
                counts['caption'] += 1
            elif r.InlineShapes.Count:
                p.LeftIndent = 0; p.FirstLineIndent = 0; p.Alignment = 1
            elif t and chr(12) not in r.Text and p.Style.NameLocal == normal.NameLocal and r.Font.Size and r.Font.Size < 12:
                r.Font.Size = SPEC['body']['size']
                p.LeftIndent = SPEC['body']['left'] * CM; p.FirstLineIndent = 0
                _pf(p.Format, 0, SPEC['body']['after'])
                counts['body'] += 1
        for ti in range(1, doc.Tables.Count + 1):
            tb = doc.Tables(ti)
            tb.Range.Font.Size = SPEC['table']['size']
            _pf(tb.Range.ParagraphFormat, 1, 1, rule=0, align=0)
            tb.Rows.LeftIndent = SPEC['body']['left'] * CM
            counts['table'] += 1
        # 작업 기준 6장: '다음 문단과 함께'·'현재 문단 나누지 않음'·'앞에서 쪽 나누기' 금지(편집 기호 검은 네모)
        for sid in (-1, -2, -3, -4, -30):
            pf = doc.Styles(sid).ParagraphFormat
            pf.KeepWithNext = False; pf.KeepTogether = False; pf.PageBreakBefore = False
        for i in range(1, doc.Paragraphs.Count + 1):
            f = doc.Paragraphs(i).Format
            if f.KeepWithNext or f.KeepTogether or f.PageBreakBefore:
                f.KeepWithNext = False; f.KeepTogether = False; f.PageBreakBefore = False
        for t in range(1, doc.TablesOfContents.Count + 1):
            doc.TablesOfContents(t).Update()
    doc.Save()
    print('적용:', counts, '| 쪽수', doc.ComputeStatistics(2))


if __name__ == '__main__':
    run()
