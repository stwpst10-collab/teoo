# 보고서 생성 도구

`80-산출물/최종 보고서`의 산업·기업 분석 보고서(v3 이후)를 만드는 스크립트다. 본문은 `report_content.py`, 레이아웃은 `build_fullreport.py`, 차트는 `charts.py`가 담당한다.

## 만드는 순서

```bash
python tools/report/charts.py <차트폴더>
python tools/report/build_fullreport.py tools/report/report_content.py <차트폴더> <출력.docx>
```

사진 묶음 그림(경쟁 제품)은 `python tools/report/photos.py`로 `tools/report/photos/`에 만든다(원본은 `30-근거/첨부`).

워드 마무리(목차·필드 갱신, 표 제목·머리말 행만 쪽 끝에 남은 곳 정리, docx·pdf 저장)는 `finish_word.py`가 한다. 사용자 워드와 분리된 새 인스턴스(DispatchEx)를 숨긴 채 쓰고 그 인스턴스만 종료하며, 빈 쪽이 있으면 출력에 표시한다.

```bash
python tools/report/finish_word.py <빌드.docx> "<최종 보고서 폴더>/그날의 류 보고서 YYYY-MM-DD vX.Y.docx" "<…>.pdf"
```

## 양식·표기 (v10.13부터, 기준: `00-운영/보고서 워드 작업 기준.md` 2·4·6장)

- 양식 값(여백 좌우 1.2/위 2.2/아래 1.5cm, 맑은 고딕 9pt·줄간격 1.075·양쪽 맞춤, 장 제목 10pt·절 제목 9.5pt, 글머리 「글머리 1단계」▪·「글머리 2단계」-, 각주 8pt 1줄, 표 8pt, 목차 2수준)은 `build_fullreport.py` 앞부분에 있다.
- 출처는 원고 문장 안에 `[^출처(시점). 근거 성격.]`으로 쓰면 그 쪽 아래 워드 각주가 된다. 근거 등급 A~D 글자는 쓰지 않고 말(공식 자료·학술지 논문·언론 보도·업체 자료·직접 관측·당사 추정)로 적는다. 웹 보고서에서는 문장 뒤 작은 회색 글씨로 보인다.
- 표·그림 아래 「자료: …」 한 줄은 유지한다. 표 제목은 표의 반복 머리글 행, 그림 제목은 그림과 같은 문단에 넣어 제목만 쪽 끝에 남지 않게 한다.
- '다음 문단과 함께'·'현재 문단을 나누지 않음'·'앞에서 쪽 나누기'는 쓰지 않는다. 장 앞 쪽 나누기는 앞 장 마지막 문장 끝에 붙인다(`page_break()`).
- 절(h2)이 하나뿐인 장은 절 번호를 붙이지 않는다. 절 제목은 "주제: 결론", `s1`은 목차에 들어가지 않는 굵은 소제목 줄이다.

필요 패키지: `python-docx`, `matplotlib`. 글꼴은 맑은 고딕.

## 원칙

- 본문의 숫자·주장은 `10-지식`, `20-리서치`, `30-근거`, `10-지식/시장/시장 수치 대장.md`에서 확인할 수 있어야 한다.
- 분석 보고서에는 추진조건·추진 현황·팀 구성·회사 개요(진행 단계·요청 규모)·발표 전 과제·개정 이력을 넣지 않는다. 기준: [최종 보고서 설계](<../../80-산출물/최종 보고서/최종 보고서 설계.md>)
- 표·그림은 `table()`·`figure()`로 넣고 「자료:」 줄에 출처를 적는다(등급 글자 대신 근거 성격을 말로).
- 새 판은 기존 파일을 덮어쓰지 않고 날짜·버전을 올려 저장한다.

## 웹 보고서 (GitHub Pages, 팀원 공유용)

`build_html.py`가 docx와 같은 원고(`report_content.py`)로 단일 HTML을 만든다. 인라인 CSS·SVG만 쓰고 외부 파일을 불러오지 않으며, 휴대폰에서도 읽힌다.

- **읽는 것:** 본문 `report_content.py`, 표지·핵심 지표·장별 요약 `build_fullreport.py`(실행하지 않고 값만 읽음), 그림 숫자 `chart_data.py`(docx 차트 `charts.py`도 같은 파일을 씀), 웹 전용 그림 `web_figures.py`(TAM·SAM·SOM 깔때기, 가격 비교, 측정 필요 체크리스트), 사람이 쓰는 요약 `share_notes.md`.
- **이번 업데이트 박스:** 직전 `share-*` 태그 시점의 `report_content.py`·`10-지식/시장/시장 수치 대장.md`를 `git show`로 꺼내 섹션 단위로 비교한다. 신규·수정·삭제 섹션과 바뀐 수치(이전값 → 새값), 수치 대장의 변경 행을 자동으로 적고 본문 섹션에 신규·수정 배지를 단다. 태그가 없으면 "첫 공유"로 표시한다.
- **변경 이력:** 모든 `share-*` 태그의 날짜, `share_notes.md` 요약, 섹션 변경 수를 누적해 보여 준다.
- **SOM 표시:** 재무·시뮬레이션 결과를 보고서에 반영하지 않는 동안 깔때기의 SOM은 "반영 보류"로 보인다. 반영이 결정되면 `web_figures.py`의 `SHOW_SOM = True`로 바꾼다.

로컬에서 만들기:

```bash
python tools/report/build_html.py            # _site/index.html (git 무시 대상)
python -m unittest discover -s tools -p 'test_build_html.py'
```

### 공유 절차

1. `tools/report/share_notes.md`에 `## share-YYYY-MM-DD` 제목을 만들고 이번 공유 요약 3~5줄을 쓴다. 태그를 붙이기 전에는 이 제목이 다음 공유본의 초안으로 쓰인다.
2. 보고서 원고·요약 변경을 main에 병합한다. main push만으로는 배포하지 않는다(공개 범위 확인 전, 상위 편집자 결정).
3. 공유할 커밋에 `share-YYYY-MM-DD` 태그를 만들고 push하면 `.github/workflows/pages.yml`이 웹 보고서를 배포한다(`git tag share-2026-11-04 && git push origin share-2026-11-04`). 직전 share 태그와 비교한 변경이 표시되고, 다음 공유본은 이 태그와 비교된다. 수동 배포는 Actions의 workflow_dispatch로 한다.

처음 한 번은 저장소 Settings → Pages → Source를 "GitHub Actions"로 바꿔야 배포된다. 병합·태그·push는 최종 편집자(상위 세션)가 한다.
