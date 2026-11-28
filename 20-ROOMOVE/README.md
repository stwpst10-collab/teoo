# ROOMOVE (사내벤처 최종 선발 제안 아이템, 2026-11-25~)

- `근거/` 팀 초안 원본(ROOMOVE 초안 2026-11-25.pdf), 팀 통합 AI 리포트(CJ_ONF3기_5조_루무브_통합AI리포트_20261112, docx·pdf)
- `리서치/` 외부 리서치 기록(경쟁사·시장 숫자·수익모델·규제·고객 문제·딥리서치 반박 검증), `시뮬레이션/` 사업 아이디어 시뮬레이션 1~10단계, `확장/` 확장 방향 피드백 검증
- `산출물/` 보고서(해커톤 보고서 양식, docx·pdf)와 차트
  - `ROOMOVE 보고서 2026-11-25 v1` 전체 보고서. 열린 워드에서 직접 고친 내용이 있어 **docx가 원본**(빌더로 다시 만들면 그 수정이 사라짐)
  - `ROOMOVE 확장·경쟁 분석 2026-11-26 v1` 경쟁 분석(v1 4장 문구 + 오늘의집 확장 경로)과 확장 전략·고객 생애가치만 떼어 낸 별도 보고서. 본문 `tools/report_expansion.py`, 차트 `tools/charts_expansion.py`. 2026-11-26 밤 워드에서 직접 고친 판(표 10-1·플라이휠 그림 추가, 대장→목록, B2B→제휴사 공급, 결합률→함께 쓰는 비율)이 **docx 원본**이라 빌더로 다시 만들면 그 수정이 사라짐. pdf는 그 이전 판
  - `ROOMOVE 확장·경쟁 분석 2026-11-27 v2` 최신본. 확장 단계를 이사 → 입주 직후 홈서비스 → 거주 중 반복 서비스·자산 기록 → 주거 보장·IoT → 다음 주거 이벤트로 재배열(내부 GPT 논의 반영), 미소·직방 경쟁 추가, 4.4 1년 단위경제(중개 기준), 검증 항목 10개. **docx가 원본**(Word에서 직접 고침). 발표 PPT·확장 논리 1장도 v2(`tools/build_deck.js`)
  - `ROOMOVE 시장조사 팀 톤 2026-11-27 v1.pptx` 산업·시장 장표는 이 파일 하나에서만 관리(흐름, 시계열 지수, 하위 항목, 산업 변화, 포지셔닝 맵). 장표를 만드는 스크립트는 `tools/build_slide_*.js`, `tools/build_slides_timeseries.js`, 숫자 근거는 `리서치/확장/주거 산업 시계열·주거 산업 변화 근거`
- `tools/` 보고서 빌더. 본문은 `tools/report_content.py`를 고치고 아래로 다시 만든다

```
python 20-ROOMOVE/tools/charts.py 20-ROOMOVE/산출물/charts
python 20-ROOMOVE/tools/build_report.py 20-ROOMOVE/tools/report_content.py 20-ROOMOVE/산출물/charts <임시.docx>
python tools/report/finish_word.py <임시.docx> <산출물 docx 절대경로> <산출물 pdf 절대경로>

# 확장·경쟁 분석 보고서
python 20-ROOMOVE/tools/charts_expansion.py 20-ROOMOVE/산출물/charts
python 20-ROOMOVE/tools/build_report.py 20-ROOMOVE/tools/report_expansion.py 20-ROOMOVE/산출물/charts <임시.docx>
```

finish_word.py의 출력 경로에 공백이 있으면 PDF 파일명이 %20으로 저장될 수 있어, 임시 폴더로 내보낸 뒤 산출물로 복사한다.

서식·표현은 `00-운영/보고서 워드 작업 기준.md`를 따른다. 리서치 결과는 보고서에 넣기 전 사용자에게 설명하고 합의한 부분만 반영한다.
