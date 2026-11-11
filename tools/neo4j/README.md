# venture 위키 → Neo4j

venture 폴더의 Markdown 노트와 `[[링크]]`를 Neo4j 그래프로 올린다. 원본은 이 폴더의 노트이고, Neo4j는 원본을 보여 주는 사본이다.

## 실행

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/neo4j/sync.ps1
```

서버가 꺼져 있으면 먼저 켠다(30~60초). 노트를 고친 뒤 다시 실행하면 갱신된다. 원본에서 지운 노트는 `active=false`로 남는다.

## 구성

- 서버·Java·Python 환경은 투자 온톨로지(`Desktop/코딩/stock-investing-skills`)에 설치된 Neo4j Community 5.26을 함께 쓴다. Community 판은 사용자 DB가 하나뿐이어서 `scope` 값으로 나눈다.
  - 그날의 류: `scope='venture-wiki'`
  - 투자 온톨로지: `scope='stock-investing-ontology'` (이 스크립트는 건드리지 않는다)
- 접속: <http://127.0.0.1:7474/browser/>, `bolt://127.0.0.1:7687`, 사용자 `neo4j`. 비밀번호는 `stock-investing-skills/neo4j/local.json`에 있고 이 저장소에는 넣지 않는다.
- 제외 폴더: `.git`, `_site`, `tools`, `.obsidian`, `.local`, `90 작업자료`. 첨부는 `.json`·`.canvas`만 올린다.
- 링크 점검 결과는 `.local/neo4j/sync-report.json`(Git 제외)에 남는다. frontmatter가 YAML로 안 읽히는 노트는 원본을 고치지 않고 `frontmatter_raw`로 보존한다.

## 예시 쿼리

```cypher
MATCH p=(n:Note {scope:'venture-wiki', active:true})-[:LINKS_TO]->(m)
RETURN p LIMIT 150;
```

`LINKS_TO`는 문서 간 참조일 뿐 근거·인과를 뜻하지 않는다.
