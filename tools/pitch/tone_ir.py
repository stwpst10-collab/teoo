"""문장 끝을 증권사 기업분석 리포트식 명사형 종결(개조식)로 바꾼다. 열린 워드 문서에 바로, 실행 취소 1회로 되돌림 가능.

  python tools/pitch/tone_ir.py          적용
  python tools/pitch/tone_ir.py --list   바꿀 후보만 출력(새 표현이 MAP에 없으면 '미등록'으로 표시)

제목(개요 수준 1·2)과 목차는 건드리지 않는다. 굵은 글씨·각주 번호는 문장 끝 글자만 바꾸므로 유지된다.
새 표현이 나오면 MAP에 추가한다(자동 활용 변환은 '선례다→선롐'처럼 틀릴 수 있어 쓰지 않는다).
"""
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import live  # noqa: E402

MAP = {
    '있다': '있음', '없다': '없음', '없었다': '없었음', '않는다': '않음', '않다': '않음',
    '계획이다': '계획', '예정이다': '예정', '전망이다': '전망', '상황이다': '상황', '수준이다': '수준',
    '판단한다': '판단', '판단된다': '판단', '예상된다': '예상', '파악된다': '파악', '추정된다': '추정',
    '요약된다': '요약', '확인된다': '확인', '관측된다': '관측', '기대된다': '기대',
    '공략한다': '공략', '확보한다': '확보', '확보했다': '확보', '차별화한다': '차별화', '자극한다': '자극',
    '택했다': '택함', '겪었다': '겪었음', '엇갈렸다': '엇갈렸음', '줄었다': '줄었음', '늘었다': '늘었음',
    '우호적이다': '우호적', '여성이다': '여성', '구조다': '구조', '경우다': '경우', '반대다': '반대',
    '선례다': '선례', '목표다': '목표', '다르다': '다름', '어렵다': '어려움', '약하다': '약함',
    '필요하다': '필요', '둔다': '둠', '했다': '함', '한다': '함', '된다': '됨', '이다': '임',
}
# '다' 뒤에 마침표·각주 번호(\x02)·닫는 따옴표·괄호가 어떤 순서로 섞여도 잡는다(예: 34,900원이다.¹⁹ → '다.\x02').
# 2026-11-08: 이전 식은 '다\x02.'만 잡고 '다.\x02'를 놓쳐 한 문서에 '~있음'과 '~다.'가 섞였음.
# '존재한다(검색 범위 기준).'처럼 '다(' 뒤 괄호 설명이 붙은 문장은 --list에서 '괄호' 후보로 따로 보여 준다.
PAT = re.compile(r'(\S+다)(?=[.\x02」"\'”’)]+(?:\s|\r|\x07|$)|[\r\x07]|$)')  # 문장 중간 '유지하다 점유율'은 제외
PAREN = re.compile(r'(\S+다)(?=\([^)]*\)[.\x02]*(?:\s|\r|\x07|$))')


def run(list_only=False):
    app, doc = live.word_doc(open_if_needed=False)
    toc = doc.TablesOfContents(1).Range if doc.TablesOfContents.Count else None
    changed, unknown = 0, {}
    ctx = live.Undo(app, '문장 끝 명사형 종결') if not list_only else None
    if ctx:
        ctx.__enter__()
    try:
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            r = p.Range
            if toc is not None and r.Start >= toc.Start and r.End <= toc.End:
                continue
            if p.OutlineLevel in (1, 2):
                continue
            t = r.Text
            for m in PAREN.finditer(t):  # 괄호 설명이 붙은 문장 끝: 자동으로 안 바꾸고 알려만 줌
                unknown['괄호:' + m.group(1)] = t[max(0, m.start() - 20):m.end() + 12].replace('\r', '')
            for m in reversed(list(PAT.finditer(t))):
                w = m.group(1)
                new = MAP.get(w)
                if new is None:  # 긴 어절은 끝부분으로 다시 찾기(예: 관리받지않는다 X, 확장했다 → '했다')
                    for k in sorted(MAP, key=len, reverse=True):
                        if w.endswith(k) and len(k) >= 2 and k in ('한다', '된다', '있다', '없다', '이다', '했다'):
                            new = w[:-len(k)] + MAP[k]
                            break
                if new is None:
                    unknown[w] = t[max(0, m.start() - 20):m.end()].replace('\r', '')
                    continue
                if list_only:
                    print(f'{w} → {new}')
                    continue
                s = r.Start + m.start()
                doc.Range(s, s + len(w)).Text = new
                changed += 1
    finally:
        if ctx:
            ctx.__exit__(None, None, None)
    if not list_only:
        doc.Save()
    print(f'바꿈 {changed}곳 | 미등록 {len(unknown)}개: {unknown}')


if __name__ == '__main__':
    run('--list' in sys.argv)
