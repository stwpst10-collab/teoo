# 보고서 생성 도구

`80-산출물/최종 보고서`의 산업·기업 분석 보고서(v3 이후)를 만드는 스크립트다. 본문은 `report_content.py`, 레이아웃은 `build_fullreport.py`, 차트는 `charts.py`가 담당한다.

## 만드는 순서

```bash
python tools/report/charts.py <차트폴더>
python tools/report/build_fullreport.py tools/report/report_content.py <차트폴더> <출력.docx>
```

Word로 목차·쪽 번호를 갱신하고 PDF를 만든다(Windows PowerShell).

```powershell
$w = New-Object -ComObject Word.Application; $d = $w.Documents.Open("<출력.docx>")
$d.Fields.Update() | Out-Null; $d.TablesOfContents.Item(1).Update()
$d.SaveAs2("<최종.docx>", 16); $d.ExportAsFixedFormat("<최종.pdf>", 17); $d.Close($false); $w.Quit()
```

필요 패키지: `python-docx`, `matplotlib`. 글꼴은 맑은 고딕.

## 원칙

- 본문의 숫자·주장은 `10-지식`, `20-리서치`, `30-근거`, `10-지식/시장/시장 수치 대장.md`에서 확인할 수 있어야 한다.
- 분석 보고서에는 추진조건·추진 현황·팀 구성·회사 개요(진행 단계·요청 규모)·발표 전 과제·개정 이력을 넣지 않는다. 기준: [최종 보고서 설계](<../../80-산출물/최종 보고서/최종 보고서 설계.md>)
- 표·그림은 `table()`·`figure()`로 넣고 출처 줄에 근거 등급을 적는다.
- 새 판은 기존 파일을 덮어쓰지 않고 날짜·버전을 올려 저장한다.
