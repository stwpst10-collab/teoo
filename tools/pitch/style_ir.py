"""열린 해커톤 문서에 IR 리포트 서식 적용(작업 기준 6장). 실행 취소 1회로 되돌릴 수 있다.

- 1단계 소제목(구 파란 라벨): ▪ + 굵은 검정 11pt
- 2단계 목록(구 • 목록): - + 들여쓰기, 양쪽 맞춤
- 본문 문단: 양쪽 맞춤
- 목차 줄간격 축소
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import live  # noqa: E402

CM = live.CM
LABEL_MAX = 20  # 라벨 판정: 짧은 한 줄, 9.5pt 이하 굵은 글씨(초안의 파란 라벨)


def is_label(p):
    t = p.Range.Text.strip()
    if not t or len(t) > LABEL_MAX or p.OutlineLevel != 10 or p.Range.Information(12):
        return False
    if p.Range.ListFormat.ListType != 0:
        return p.Range.ListFormat.ListTemplate.ListLevels(1).NumberFormat == '▪'
    f = p.Range.Font
    return f.Bold == -1 and f.Size and f.Size <= 9.5 and t not in ('목차',) and not t.startswith('사내벤처')


def run():
    app, doc = live.word_doc(open_if_needed=False)
    paras = [doc.Paragraphs(i) for i in range(1, doc.Paragraphs.Count + 1)]
    labels = [p for p in paras if is_label(p)]
    bullets = [p for p in paras if p.Range.ListFormat.ListType != 0 and p not in labels
               and p.Range.ListFormat.ListTemplate.ListLevels(1).NumberFormat != '▪']
    with live.Undo(app, 'IR 서식'):
        # 1단계 ▪ 소제목
        if labels:
            lt1 = doc.ListTemplates.Add(False)
            lv = lt1.ListLevels(1)
            lv.NumberFormat = '▪'; lv.NumberStyle = 23  # 글머리기호
            lv.TrailingCharacter = 1; lv.Font.Name = live.d.FONT; lv.Font.Size = 9
            lv.NumberPosition = 0; lv.TextPosition = 0.45 * CM; lv.TabPosition = 0
            for p in labels:
                r = p.Range
                r.Font.Color = 0; r.Font.Size = 11; r.Font.Bold = True
                p.Format.SpaceBefore = 14; p.Format.SpaceAfter = 4
                r.ListFormat.ApplyListTemplate(lt1, False, 0)
        # 2단계 - 목록
        if bullets:
            lt2 = doc.ListTemplates.Add(False)
            lv = lt2.ListLevels(1)
            lv.NumberFormat = '-'; lv.NumberStyle = 23
            lv.TrailingCharacter = 1; lv.Font.Name = live.d.FONT; lv.Font.Size = 10
            lv.NumberPosition = 0.45 * CM; lv.TextPosition = 0.85 * CM; lv.TabPosition = 0
            for p in bullets:
                p.Range.ListFormat.ApplyListTemplate(lt2, False, 0)
                p.Alignment = 3
                p.Format.SpaceAfter = 3
        # 본문 문단 양쪽 맞춤
        for p in paras:
            if p.Range.Information(12) or p.OutlineLevel != 10 or p.Range.ListFormat.ListType != 0:
                continue
            t = p.Range.Text.strip()
            if len(t) > 40 and not t.startswith('자료:') and p.Range.InlineShapes.Count == 0:
                p.Alignment = 3
        # 목차 줄간격
        for sid in (-20, -21):
            f = doc.Styles(sid).ParagraphFormat
            f.SpaceBefore = 0; f.SpaceAfter = 2; f.LineSpacingRule = 0
        for t in range(1, doc.TablesOfContents.Count + 1):
            doc.TablesOfContents(t).Update()
    doc.Save()
    print(f'▪ 소제목 {len(labels)}개, - 목록 {len(bullets)}개, 쪽수 {doc.ComputeStatistics(2)}')


def merge_labels():
    """절 제목 바로 위 소제목(▪ 라벨)을 제목 앞에 '주제: ' 로 합치고 라벨 문단은 지운다(최종 보고서 v10.12 제목 형식).
    목록 들여쓰기도 통일한다."""
    import re
    app, doc = live.word_doc(open_if_needed=False)
    merged = 0
    with live.Undo(app, '소제목을 절 제목에 합침'):
        i = doc.Paragraphs.Count
        while i > 1:
            p = doc.Paragraphs(i)
            prev = doc.Paragraphs(i - 1)
            if p.OutlineLevel == 2 and is_label(prev):
                label = prev.Range.Text.strip()
                r = p.Range
                m = re.match(r'^\d+\.\d+\s+', r.Text)
                pos = r.Start + (len(m.group(0)) if m else 0)
                doc.Range(pos, pos).InsertAfter(f'{label}: ')
                prev.Range.Delete()
                merged += 1
            i -= 1
        for k in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(k)
            lf = p.Range.ListFormat
            if lf.ListType != 0 and lf.ListTemplate.ListLevels(1).NumberFormat == '-':
                p.LeftIndent = 0.85 * CM
                p.FirstLineIndent = -0.4 * CM
        for t in range(1, doc.TablesOfContents.Count + 1):
            doc.TablesOfContents(t).Update()
    doc.Save()
    print(f'소제목 {merged}개를 절 제목에 합침, 쪽수 {doc.ComputeStatistics(2)}')


if __name__ == '__main__':
    {'run': run, 'merge': merge_labels}[sys.argv[1] if len(sys.argv) > 1 else 'run']()
