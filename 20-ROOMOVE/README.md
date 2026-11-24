# ROOMOVE (사내벤처 최종 선발 제안 아이템, 2026-11-25~)

- `근거/` 팀 초안 원본(ROOMOVE 초안 2026-11-25.pdf)
- `리서치/` 외부 리서치 기록(경쟁사·시장 숫자·수익모델·규제·고객 문제·딥리서치 반박 검증), `시뮬레이션/` 사업 아이디어 시뮬레이션 1~10단계
- `산출물/` 보고서(해커톤 보고서 양식, docx·pdf)와 차트
- `tools/` 보고서 빌더. 본문은 `tools/report_content.py`를 고치고 아래로 다시 만든다

```
python 20-ROOMOVE/tools/charts.py 20-ROOMOVE/산출물/charts
python 20-ROOMOVE/tools/build_report.py 20-ROOMOVE/tools/report_content.py 20-ROOMOVE/산출물/charts <임시.docx>
python tools/report/finish_word.py <임시.docx> <산출물 docx 절대경로> <산출물 pdf 절대경로>
```

서식·표현은 `00-운영/보고서 워드 작업 기준.md`를 따른다. 리서치 결과는 보고서에 넣기 전 사용자에게 설명하고 합의한 부분만 반영한다.
