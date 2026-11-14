@echo off
chcp 65001 >nul
title 엑셀 → 워드 보고서 자동 반영
echo 이 창을 켜 두면 시뮬레이션 엑셀을 저장할 때마다 해커톤 워드 보고서의 차트·표가 자동으로 바뀝니다.
echo 끄려면 이 창을 닫으세요.
echo.
python "%~dp0..\..\tools\pitch\live.py" watch
pause
