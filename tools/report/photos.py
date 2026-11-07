"""경쟁 제품 사진 묶음 그림(그리드 PNG)을 만든다.

    python tools/report/photos.py

원본 사진은 30-근거/첨부/ 아래 출처 노트와 함께 보관하고, 이 스크립트는 그 사진을 읽어
tools/report/photos/<그림>.png 를 만든다. 보고서 원고(report_content.py)는 PHOTOS / '<그림>.png'로 넣는다.
원본이 없는 칸은 건너뛰고 경고만 낸다.
"""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[2]
OUT = Path(__file__).resolve().parent / 'photos'
PMS = '30-근거/첨부/2026-11-07 PMS 표방 기기 사진/'
COMP = '30-근거/첨부/2026-11-07 경쟁 제품 사진/'

# 그림 이름 → [(원본 경로, 제품명, 한 줄 설명, 잘라낼 영역(원본 비율 0~1: 왼, 위, 오, 아래) 또는 None)]
GRIDS = {
    'p1_pms_devices.png': [
        (PMS + 'Lutea 헤드밴드 (Samphire 공식 출시 글).png', 'Lutea·Nettle (영국 Samphire)', '머리띠형 뇌 미세전류 자극 · $589', None),
        (PMS + 'OhmBody 귀 부착 기기 (Dallas Innovates).png', 'OhmBody (미국 Spark Biomedical)', '귀 부착 신경자극 · $719 · 1회용 패드', None),
        (PMS + '발목 지압 스트랩 RCT 그림1 (Frontiers 2024).png', '발목 지압 스트랩 (중국, 연구용)', '삼음교 압박 돌기 밴드 · 생리통 임상 78명', (0.335, 0.0, 0.67, 1.0)),
    ],
    'p2_overseas_period.png': [
        (COMP + 'Beurer EM 50.jpg', 'Beurer EM 50 (독일)', 'TENS + 43°C 온열 · 약 7.5만~12.6만원', None),
        (COMP + 'Livia 제품 사진.png', 'Livia (미국·이스라엘)', 'TENS · $129 · FDA 510(k)', None),
        (COMP + 'Myoovi 2.0.jpg', 'Myoovi 2.0 (영국)', 'TENS · £79.99 · 교체 패드', None),
        (COMP + 'TensCare Ova+.jpg', 'TensCare Ova+ (영국)', 'TENS · £37.50 · FDA 510(k)', None),
        (COMP + 'Stoov Wrap.png', 'Stoov Wrap (네덜란드)', '무선 온열 벨트 · €109.95~', None),
        (COMP + 'ThermaCare Menstrual.webp', 'ThermaCare (미국)', '일회용 생리 온열 패치', None),
        (COMP + 'Sea-Band 제품 사진.jpg', 'Sea-Band (영국)', '손목 지압 밴드 · $15.40', None),
        (COMP + 'Hyperice Venom Go.png', 'Hyperice Venom Go (미국)', '부착형 온열·진동 · $129', None),
    ],
    'p3_domestic.png': [
        (COMP + '올리브영 케어플러스 아랫배 찜질 패드.jpg', '케어플러스 아랫배 찜질 패드', '올리브영 PB · 1매 1,200원', None),
        (COMP + '바로랩 릴리프 벨트.jpg', '바로랩 릴리프 벨트', '생리통 검색 최상단 · 3.5만·7.4만원', None),
        (COMP + '오아 허그워머.jpg', '오아 허그워머', '무선 찜질기 · 34,900원', None),
        (COMP + '슬룸 골반 케어.jpg', '슬룸 골반 케어', '하복부 온열 · 19.9만원', None),
        (COMP + '풀리오 종아리 마사지기 V3.jpg', '풀리오 종아리 V3', '공기압 마사지기', None),
        (COMP + '발란스핏 EMS 슬림 벨트 PRO.png', '발란스핏 EMS 슬림 벨트 PRO', 'EMS 벨트', None),
        (COMP + 'Therabody RecoveryTherm Cube.png', 'Therabody RecoveryTherm Cube', '온냉 착용 기기 · $159.99 (해외)', None),
    ],
}

CELL_W, IMG_H, TEXT_H, PAD, COLS = 420, 320, 78, 18, 4
FONT = 'C:/Windows/Fonts/malgun.ttf'
FONT_B = 'C:/Windows/Fonts/malgunbd.ttf'


def _font(path, size):
    try:
        return ImageFont.truetype(path, size)
    except OSError:
        return ImageFont.load_default()


def _fit(im, w, h):
    im = im.convert('RGBA')
    bg = Image.new('RGBA', im.size, (255, 255, 255, 255))
    im = Image.alpha_composite(bg, im).convert('RGB')
    # 흰 여백 자동 제거(제품이 작게 보이는 렌더 이미지 대응)
    from PIL import ImageChops
    box = ImageChops.difference(im, Image.new('RGB', im.size, 'white')).convert('L').point(lambda v: 255 if v > 18 else 0).getbbox()
    if box:
        m = 12
        im = im.crop((max(0, box[0] - m), max(0, box[1] - m), min(im.width, box[2] + m), min(im.height, box[3] + m)))
    im.thumbnail((w, h), Image.LANCZOS)
    canvas = Image.new('RGB', (w, h), 'white')
    canvas.paste(im, ((w - im.width) // 2, (h - im.height) // 2))
    return canvas


def build(name, items):
    cells = []
    for rel, title, sub, crop in items:
        path = ROOT / rel
        if not path.exists():
            print(f'  건너뜀(원본 없음): {rel}')
            continue
        im = Image.open(path)
        if crop:
            w, h = im.size
            im = im.crop((int(crop[0] * w), int(crop[1] * h), int(crop[2] * w), int(crop[3] * h)))
        cells.append((_fit(im, CELL_W, IMG_H), title, sub))
    if not cells:
        print(f'  {name}: 사진이 없어 만들지 않음')
        return
    cols = min(COLS, len(cells))
    rows = (len(cells) + cols - 1) // cols
    W = cols * CELL_W + (cols + 1) * PAD
    H = rows * (IMG_H + TEXT_H) + (rows + 1) * PAD
    sheet = Image.new('RGB', (W, H), 'white')
    d = ImageDraw.Draw(sheet)
    f_t, f_s = _font(FONT_B, 21), _font(FONT, 18)
    for i, (img, title, sub) in enumerate(cells):
        r, c = divmod(i, cols)
        x = PAD + c * (CELL_W + PAD)
        y = PAD + r * (IMG_H + TEXT_H + PAD)
        sheet.paste(img, (x, y))
        d.rectangle([x, y, x + CELL_W - 1, y + IMG_H - 1], outline=(220, 220, 220))
        d.text((x + CELL_W // 2, y + IMG_H + 12), title, font=f_t, fill=(20, 20, 20), anchor='mt')
        d.text((x + CELL_W // 2, y + IMG_H + 44), sub, font=f_s, fill=(90, 90, 90), anchor='mt')
    OUT.mkdir(exist_ok=True)
    sheet.save(OUT / name, optimize=True)
    print(f'  {name}: {len(cells)}칸 저장')


if __name__ == '__main__':
    for name, items in GRIDS.items():
        build(name, items)
