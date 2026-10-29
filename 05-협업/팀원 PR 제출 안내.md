---
type: workflow
project: venture
status: 로컬 준비·origin 연결·공개 승인 완료 — 소유자 인증 대기
updated: 2026-10-29
---

# 팀원 PR 제출 안내

이 저장소에서 자료를 고친 뒤 PR(변경 검토 요청)을 제출한다. 사용자 계정 `stwpst10-collab`이 최종 내용을 검토한다. 같은 저장소의 개인별 작업 브랜치를 사용한다.

## 처음 한 번

1. 관리자가 보낸 저장소 초대를 GitHub에서 수락한다.
2. 관리자가 공유한 정확한 저장소 페이지를 연다. 계정 프로필 주소와 저장소 주소는 다르다.
3. [파일명 규칙](../00-운영/파일명과-메타데이터.md)과 [리서치 양식](../00-운영/양식/리서치.md)을 읽는다.

## GitHub 화면에서 간단히 제출하기

1. 저장소의 최신 `main`에서 브랜치 선택 메뉴를 열고 `research/내아이디/조사주제` 이름의 작업 브랜치를 만든다.
2. 작업 브랜치에서 조사 노트를 추가하거나 수정한다. 파일명은 `생리통 기기 구매 동기 조사.md`처럼 쓰고 기존 양식의 속성을 채운다.
3. 변경을 해당 작업 브랜치에 커밋한다. `main`에 직접 저장하지 않는다.
4. **Pull requests → New pull request**에서 base를 `main`, compare를 내 작업 브랜치로 선택한다.
5. 변경 파일 목록을 확인하고 **Create pull request**를 누른다. 자동으로 표시되는 양식에 출처·결론·영향·미확인 사항을 적는다.
6. `ontology-check` 결과를 확인한다. 오류가 있으면 같은 작업 브랜치에서 수정하면 PR도 갱신된다.
7. 관리자의 의견에 답하고 검토를 기다린다. 제출자가 자신의 PR을 임의로 병합하지 않는다.

긴 조사·여러 파일·엑셀 작업은 저장소를 PC에 내려받아 편집한다. [참여 안내](../CONTRIBUTING.md)의 LLM 요청 문장을 사용할 수 있다. 엑셀·이미지는 원본을 덮어쓰지 않고 변경 이유와 검증 결과를 함께 제출한다.

## 관리자가 확인할 것

- 정확한 저장소와 권한을 확인한 뒤 팀원을 초대한다. 소유자 계정의 비밀번호나 토큰을 공유하지 않는다.
- `.github/CODEOWNERS`는 모든 변경의 검토자로 `stwpst10-collab`을 지정한다. 이 파일만으로 검토가 강제되지는 않는다.
- 지원되는 플랜에서 `main`에 PR 필수, 코드 소유자 검토 필수, 승인 후 새 커밋이 올라오면 재검토, `ontology-check` 성공 필수, 대화 해결 필수를 설정한다.
- main 강제 push·삭제는 허용하지 않는다. 관리자 예외 범위는 실제 운영 권한에 맞춰 확인한다.
- [PR 현황](PR-현황.md)은 실제 GitHub 조회 성공 후 갱신한다. 조회 실패를 요청 없음으로 표시하지 않는다.

비공개 개인 저장소의 보호 기능은 GitHub Pro, 비공개 조직 저장소는 Team 이상에서 지원된다. 무료 비공개 저장소의 협업 가능 여부와 승인 강제 가능 여부는 구분한다. 자신의 PR은 승인할 수 없으므로 관리자 본인 변경의 검토 방식도 정해야 한다. 소유자 승인 필수와 소유자만 병합 가능은 다른 설정이다. [브랜치 보호](https://docs.github.com/en/repositories/configuring-branches-and-merges-in-your-repository/managing-protected-branches/about-protected-branches), [코드 소유자](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/about-code-owners), [PR 검토](https://docs.github.com/en/pull-requests/how-tos/review-pull-requests/reviewing-proposed-changes-in-a-pull-request).

## 연결 상태

2026-10-29 확인: [stwpst10-collab/teoo](https://github.com/stwpst10-collab/teoo)는 공개·빈 저장소다. 사용자가 공개 유지를 승인했고 로컬 origin 연결을 완료했다. 현재 기존 PC 인증 계정 `warrenbff`는 읽기 권한만 있어 별도 소유자 인증을 진행 중이다. 업로드 대상은 teoo 하나이며 warrenbff 소유 저장소에는 업로드하지 않는다. 실제 업로드·팀원 초대·브랜치 보호는 아직 수행하지 않았다.
