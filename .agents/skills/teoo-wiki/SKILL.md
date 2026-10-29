---
name: teoo-wiki
description: 사내벤처 그날의 류의 teoo 온톨로지를 검색하고, 출처를 읽어 답하며, 조사·지식·산출물을 반영하고 요청된 PR을 제출한다. teoo 위키, 그날의 류 자료, 오비라 사례, 팀 온톨로지 질문에 사용한다. 외부 가치투자 온톨로지와 구분한다.
---

# teoo 팀 위키 (Codex)

기준 저장소는 **https://github.com/stwpst10-collab/teoo**다. 기존 한글 Markdown 온톨로지를 직접 사용한다. `omx_wiki` 같은 별도 사본을 만들지 않는다. OMX나 특정 플러그인 설치가 없어도 Git·Python과 파일 도구로 수행할 수 있다.

**작업 전에 [references/rules.md](references/rules.md)를 읽는다.** 답변 형식(사실·회사 주장·추론·미확인 구분, 근거 파일·날짜·검토 상태 표시), 반영·제출 규칙은 그 파일 하나에서 관리한다. 이 파일은 Codex에서의 저장소 찾기와 도구 사용만 다룬다.

## 저장소 찾기

스킬 디렉터리를 `SKILL_ROOT`, 로컬 teoo 체크아웃을 `REPO`로 삼는다. 아래 표기는 설명용이며 실제 실행에서는 확인한 경로를 인자로 전달한다.

1. 현재 폴더와 상위 폴더, 현재 폴더의 `venture`, 설치된 스킬의 `workspace.json`을 후보로 확인한다. `python SKILL_ROOT/scripts/wiki.py --repo REPO status`로 저장소와 브랜치·커밋·미커밋 상태를 읽는다. `--repo`를 생략하면 도우미가 후보를 찾는다.
2. 원격이 `stwpst10-collab/teoo`인지 확인한다. 이름이 비슷한 다른 저장소나 외부 가치투자 온톨로지를 대신 쓰지 않는다. 사용자 요청이 다른 저장소라면 이 스킬의 대상을 임의로 바꾸지 않는다.
3. 체크아웃이 없으면 사용자가 지정한 빈 경로 또는 아직 존재하지 않는 `~/teoo`에 `git clone https://github.com/stwpst10-collab/teoo.git`으로 내려받는다. 기존 폴더를 덮어쓰지 않는다. 저장소 공개 여부와 네트워크·파일 권한 제한을 따른다.
4. `REPO/AGENTS.md`, `README.md`와 관련 질문·현재 결정을 읽는다. 설치된 스킬보다 저장소의 최신 주제 내용과 운영 규칙을 따른다.

자료 경로와 PR 절차는 [references/workflows.md](references/workflows.md)를 필요할 때 읽는다. 이미 내려받은 저장소가 있으면 검색 질문마다 전체를 다시 복제하거나 pull하지 않는다. 동기화 요청은 미커밋 작업과 브랜치를 확인한 뒤 fast-forward 가능한 경우에만 진행한다.

## 도구

```sh
python SKILL_ROOT/scripts/wiki.py --repo REPO status
python SKILL_ROOT/scripts/wiki.py --repo REPO search "오비라"
python SKILL_ROOT/scripts/wiki.py --repo REPO read "10-지식/제품/오비라.md"
```

검색은 문장 전체보다 핵심어(예: `시뮬레이션`, `수요 검증`)로 한다. 검색 결과는 증거 본문이 아니므로 주장과 연결된 리서치·원문까지 읽는다. 결과가 잘렸으면 필요한 부분을 추가로 읽는다. 이 도우미는 읽기 전용이며 파일을 수정하거나 PR을 보내지 않는다.

## 요청별 처리

- **찾기·설명**: 위 도구로 찾고 읽은 뒤 rules.md의 답변 형식으로 답한다. 파일을 수정하지 않는다.
- **조사·반영**: rules.md의 반영 규칙대로 기존 주제·질문을 먼저 찾고 양식에 맞춰 작성한 뒤 `python tools/validate.py`를 실행한다.
- **PR 제출·검토**: 사용자가 요청한 경우에만 workflows.md의 PR 절차를 따른다. 대상은 stwpst10-collab/teoo의 main이며 자동 병합하지 않는다.
