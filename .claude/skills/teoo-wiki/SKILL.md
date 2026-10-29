---
name: teoo-wiki
description: 사내벤처 그날의 류의 teoo 온톨로지를 검색하고 출처를 읽어 답하며, 조사·지식·산출물을 반영하고 요청 시 PR을 제출한다. 온톨로지 개념·관계, 오비라 등 경쟁사·수요 검증 근거, 최신 시뮬레이션 모델링·프로토타입 시안 질문에 사용한다. 외부 가치투자 온톨로지와 구분한다.
argument-hint: "[질문 또는 작업 요청]"
---

# teoo 팀 위키 (Claude Code)

요청: $ARGUMENTS

기준 저장소는 **https://github.com/stwpst10-collab/teoo**다. Codex용 스킬과 같은 도구·규칙을 그대로 쓴다. 별도 위키 사본을 만들지 않는다.

## 시작

1. teoo 저장소 루트(`AGENTS.md`와 `.agents/skills/teoo-wiki`가 있는 폴더)를 찾는다. 보통 현재 작업 폴더다. 이하 명령은 그 루트에서 실행한다.
2. 공통 운영 규칙 [rules.md](../../../.agents/skills/teoo-wiki/references/rules.md)를 읽는다. 답변 형식·반영·제출 규칙은 그 파일을 따른다. 자료 위치와 PR 절차는 [workflows.md](../../../.agents/skills/teoo-wiki/references/workflows.md)에 있다.
3. 기준 상태를 확인한다. 결과의 `branch`·`commit`·`dirty`를 답변의 기준 시점으로 적는다.

   ```sh
   python .agents/skills/teoo-wiki/scripts/wiki.py status
   ```

   origin이 teoo가 아니라는 오류가 나면 멈추고 사용자에게 알린다. 다른 저장소로 바꾸지 않는다.

## 조회 (읽기 전용)

```sh
python .agents/skills/teoo-wiki/scripts/wiki.py search "오비라"
python .agents/skills/teoo-wiki/scripts/wiki.py read "20-리서치/오비라 커뮤니티 수요 검증.md"
```

- 문장 전체가 아니라 핵심어로 검색한다. 예: "최신 시뮬레이션 모델링 찾아줘" → `search "시뮬레이션"` 후 `80-산출물/README.md`와 해당 산출물 노트를 읽어 최신 버전을 고른다.
- 개념·관계는 `10-지식` 노트의 링크와 `40-결정`, `50-검증/리서치 질문 목록.md`를 따라간다. 검색 결과 줄은 증거가 아니므로 연결된 리서치·근거 본문까지 읽는다.
- 조회 요청에서는 Write·Edit 등 파일 변경을 하지 않는다.

## 반영

사용자가 반영·작성을 요청했을 때만 파일을 만든다. `git status --short`로 다른 사람의 변경을 먼저 확인하고, rules.md의 반영 규칙(한국어 파일명, 메타데이터 ID, `review: pending`, 결정과 다른 의견은 제안으로)을 따른다. 끝나면 실행한다.

```sh
python tools/validate.py
python -m unittest discover -s tools -p "test_*.py"   # 코드나 스킬 파일을 바꾼 경우
```

`.agents/skills/teoo-wiki` 아래 공통 파일이나 `tools/claude_web_skill/SKILL.md`를 바꿨다면 `python tools/build_web_skill.py`로 웹앱 ZIP도 다시 만든다.

## 제출 (요청 시에만)

workflows.md의 PR 절차를 따른다. 요약하면:

1. `git remote get-url origin`이 stwpst10-collab/teoo인지 확인한다.
2. `gh auth status`와 `gh api repos/stwpst10-collab/teoo --jq .permissions`로 실제 push 권한을 확인한다.
3. `research/작성자/주제` 브랜치에 요청된 파일만 stage·커밋하고 push한 뒤 `gh pr create --repo stwpst10-collab/teoo --base main`으로 PR을 만든다. 본문은 `.github/pull_request_template.md`를 따른다.
4. main 직접 push, 자동 병합, 강제 push를 하지 않는다.
5. `gh`가 없거나 인증·권한이 없으면 로컬 커밋 또는 변경과 검증까지 마치고 "업로드되지 않음"과 필요한 인증 작업을 정확히 알린다. 토큰을 채팅으로 받지 않는다.
