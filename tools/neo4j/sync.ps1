# 그날의 류 venture 위키 → Neo4j 동기화 (scope: venture-wiki)
# 서버·Python 환경은 투자 온톨로지(stock-investing-skills)에 설치된 것을 함께 쓴다. 데이터는 scope로 분리된다.
$ErrorActionPreference = 'Stop'
$invest = Join-Path $env:USERPROFILE 'Desktop\코딩\stock-investing-skills'
if (-not (Get-NetTCPConnection -LocalPort 7687 -State Listen -ErrorAction SilentlyContinue)) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $invest 'neo4j\start.ps1')
}
& (Join-Path $invest '.cache\neo4j-venv\Scripts\python.exe') (Join-Path $PSScriptRoot 'sync_neo4j.py') @args
if ($LASTEXITCODE -ne 0) { throw 'Neo4j sync failed.' }
