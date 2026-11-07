"""모바일 목록: 폰에서 저장소 파일을 한 번 눌러 열 수 있는 목록(모바일 목록.md)을 만든다.

- md·이미지: GitHub 화면으로 연다(앱·브라우저 모두 표시됨)
- docx·xlsx·pptx: Microsoft Office 웹 뷰어로 연다(저장소가 공개라 로그인 없이 미리보기)
- pdf·영상·html 등: GitHub 화면(브라우저에서 미리보기, 또는 다운로드)
폴더 안은 이름 역순(날짜·버전이 앞에 붙어 최신이 위로)으로 정렬한다.

  python tools/mobile_index.py          목록 갱신(바뀌었으면 파일을 다시 씀)
커밋 때 자동 갱신: .git/hooks/pre-commit 에서 이 스크립트를 실행한다.
"""
import subprocess
from collections import defaultdict
from pathlib import Path
from urllib.parse import quote

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / '모바일 목록.md'
REPO = 'stwpst10-collab/teoo'
BRANCH = 'main'
SKIP_TOP = {'.agents', '.claude', '.github', 'tools'}
SKIP_NAMES = {'.gitkeep', '.gitignore', '.gitattributes', OUT.name}
OFFICE = {'.docx', '.xlsx', '.pptx', '.doc', '.xls', '.ppt'}
ICON = {'.md': '📝', '.pdf': '📕', '.docx': '📘', '.xlsx': '📗', '.pptx': '📙', '.png': '🖼', '.jpg': '🖼',
        '.jpeg': '🖼', '.mp4': '🎬', '.html': '🌐'}


def files():
    out = subprocess.run(['git', '-c', 'core.quotepath=off', 'ls-files', '-z'], cwd=ROOT,
                         capture_output=True, check=True).stdout.decode('utf-8')
    for f in out.split('\0'):
        p = Path(f)
        if not f or p.parts[0] in SKIP_TOP or p.name in SKIP_NAMES or len(p.parts) == 1 and p.suffix != '.md':
            continue
        yield f


def link(f):
    ext = Path(f).suffix.lower()
    if ext in OFFICE:
        raw = f'https://raw.githubusercontent.com/{REPO}/{BRANCH}/{quote(f)}'
        return 'https://view.officeapps.live.com/op/view.aspx?src=' + quote(raw, safe='')
    return f'https://github.com/{REPO}/blob/{BRANCH}/{quote(f)}'


def build():
    groups = defaultdict(list)
    for f in files():
        groups[str(Path(f).parent).replace('\\', '/')].append(f)
    lines = ['# 모바일 목록', '',
             '폰에서 파일을 바로 열기 위한 목록이다. 커밋할 때 자동으로 갱신된다(`tools/mobile_index.py`).',
             '워드·엑셀(📘📗)은 Office 웹 뷰어로 열리고, 나머지는 GitHub 화면으로 열린다. 폴더 안은 최신이 위.', '']
    tops = sorted({g.split('/')[0] for g in groups})
    lines.append('바로가기: ' + ' · '.join(f'[{t}](#{quote(t.replace(" ", "-"))})' for t in tops if t != '.'))
    lines.append('')
    for top in tops:
        if top != '.':
            lines += [f'## {top}', '']
        for g in sorted((g for g in groups if g.split('/')[0] == top), key=lambda g: (g.count('/'), g)):
            if g != top and g != '.':
                lines += [f'**{g.split("/", 1)[1]}**', '']
            for f in sorted(groups[g], key=lambda f: Path(f).name, reverse=True):
                p = Path(f)
                lines.append(f'- {ICON.get(p.suffix.lower(), "📄")} [{p.name}]({link(f)})')
            lines.append('')
    return '\n'.join(lines).rstrip() + '\n'


if __name__ == '__main__':
    text = build()
    if not OUT.exists() or OUT.read_text(encoding='utf-8') != text:
        OUT.write_text(text, encoding='utf-8', newline='\n')
        print(f'갱신: {OUT.name}')
    else:
        print('변경 없음')
