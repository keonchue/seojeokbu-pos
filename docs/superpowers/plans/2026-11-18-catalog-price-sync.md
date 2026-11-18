# 카탈로그 공개 + 앱 가격·할인 연동 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 카탈로그를 `keonchue.github.io/seojeokbu-catalog/` 로 공개하고, POS 앱의 정가·책별 할인율이 바뀌면 카탈로그에 곧바로 보이게 한다.

**Architecture:** 앱이 연결된 책(`catalogNo`)들로 공개 문서 `public/catalog` 를 써 두고(바뀔 때만), 보안 규칙은 그 문서 하나만 누구나 읽게 연다. 카탈로그는 시작할 때 Firestore REST 로 그 문서를 읽어 CSV 값을 덮어쓰고, 3초 안에 못 읽으면 CSV 값 그대로 그린다.

**Tech Stack:** 정적 HTML + 바닐라 JS (두 저장소 모두 빌드 없음), Firebase Firestore v10.12.2 (앱), Firestore REST v1 (카탈로그), GitHub Pages, Node 18+ 시험 스크립트 (`node:vm` 으로 index.html 의 함수를 꺼내 시험 — `tests/sheet-contract.mjs` 와 같은 방식).

**Spec:** `docs/superpowers/specs/2026-11-18-catalog-price-sync-design.md`

## Global Constraints

- 재고·제목·원가·매출은 공개 문서에 넣지 않는다. `public/catalog` 는 `{ items: { "<번호>": { price, salePct } }, updatedAt }` 뿐.
- 책 문서의 연결 필드 이름은 `catalogNo` (양의 정수). 번호 하나 ↔ 앱 책 하나. 택배비 책(`isShipping`)은 연결하지 않는다.
- 할인가 식은 앱과 같게: `Math.round(price * (100 - pct) / 100)`. `salePct` 는 앱의 `promoPct(b)` 값(0~100 정수).
- 카탈로그는 공개 문서를 **3000ms** 안에 못 읽으면 CSV 값으로 그린다. 오류는 화면에 띄우지 않는다.
- 공개 문서 주소: `https://firestore.googleapis.com/v1/projects/seojeokbu-pos/databases/(default)/documents/public/catalog`
- 카탈로그 CSV 주소(앱이 연결 화면에서 읽음): `https://keonchue.github.io/seojeokbu-catalog/books.csv`
- 코드 주석·화면 글은 지금 파일들처럼 한국어로.
- POS 저장소 코드 커밋 전에는 Codex 리뷰(`/codex:review` 또는 codex-rescue, 리뷰만). 리뷰→수정은 최대 2번.
- `editBookModal` 을 고치므로 POS 커밋 전에 시트 계약 시험을 돌린다: `node tests/sheet-contract.mjs` (주소·토큰은 사용자 환경변수 `SHEET_TEST_URL` / `SHEET_TEST_TOKEN` 에 있다). `apps-script/Code.gs` 는 고치지 않는다.
- `seojeokbu-catalog-main/` (zip 사본) 은 POS 저장소에 커밋하지 않는다.
- 바깥으로 나가는 일(새 저장소 만들기·Pages 켜기·push·채림 님 초대)은 그 직전에 사용자에게 확인한다. 새 저장소 만들기·Pages 는 이 계획 승인으로 허락된 것으로 보고, 초대만 따로 묻는다.

## Review Focus

1. 보안 규칙을 아직 콘솔에 안 올렸을 때 (REST 가 403) → 카탈로그는 아무 말 없이 CSV 값으로 보여야 한다. → Task 2 `readLive` 시험.
2. 오프라인이었던 기기가 옛 캐시로 공개 문서를 덮어쓰는 경우 → 서버에서 온 스냅샷이고 보낼 쓰기가 없을 때만 쓴다. → Task 3 `shouldPublish` 시험.
3. `catalogNo` 가 이상한 값(`"12 "`, `0`, `-3`, `"abc"`, 빈칸)이거나 두 책이 같은 번호 → 이상한 값은 연결 없음, 같은 번호는 먼저 만든 책 하나만. → Task 3 `catalogItems` 시험.
4. 앱 가격이 0이거나 할인 100% → 0원 정가는 '가격 문의', 100% 할인은 '0원'으로 보여야 한다('가격 문의' 아님). → Task 2 `pr` 시험.
5. 카탈로그 번호가 CSV 에 없는 줄인데 공개 문서에만 있음 → 카탈로그에 새 책이 생기지 않는다(무시). → Task 2 `applyLive` 시험.

---

## File Structure

카탈로그 저장소 (새로 만들어 `C:\Users\Keon\Desktop\Moim\seojeokbu-catalog` 에 받는다):
- `index.html` — 수정: 공개 문서 읽기(`readLive`, `applyLive`), 할인율 배지(`pr`, `sale`).
- `tests/live-prices.mjs` — 새로: index.html 의 `<script>` 에서 함수를 꺼내 시험.
- `README.md` — 수정: 새 주소, 가격·할인은 앱에서 바꾼다는 안내.

POS 저장소:
- `tests/grab-index.mjs` — 새로: `tests/sheet-contract.mjs` 의 `grab()` 을 옮겨 `export function grabFrom(src, name)` 으로 내보낸다. sheet-contract 는 이걸 import.
- `tests/catalog-sync.mjs` — 새로: 앱 쪽 순수 함수 시험.
- `index.html` — 수정: 순수 함수 4개, 공개 문서 쓰기, 책 수정 화면 칸, 카탈로그 연결 창, 설정 버튼.
- `firestore.rules` — 수정: `public/catalog` 규칙.

---

### Task 1: 카탈로그 공개 저장소 + Pages

**Files:** 카탈로그 저장소 전체 (내용 변경 없음)

**Interfaces:**
- Produces: `https://keonchue.github.io/seojeokbu-catalog/` (index.html, books.csv 가 200), 로컬 경로 `C:\Users\Keon\Desktop\Moim\seojeokbu-catalog` (git, origin = keonchue/seojeokbu-catalog)

- [ ] **Step 1:** `gh repo clone chaelym/seojeokbu-catalog "C:/Users/Keon/Desktop/Moim/seojeokbu-catalog"` 로 최신본을 받고, `diff -rq` 로 zip 사본(`Book Store/seojeokbu-catalog-main/seojeokbu-catalog-main`)과 다른 점이 있으면 사용자에게 알린다 (원본이 이긴다).
- [ ] **Step 2:** `gh repo create keonchue/seojeokbu-catalog --public`, origin 을 그 저장소로 바꾸고 원래 것은 `upstream` 으로 남긴 뒤 `git push -u origin main`.
- [ ] **Step 3:** `gh api -X POST repos/keonchue/seojeokbu-catalog/pages -f "source[branch]=main" -f "source[path]=/"`.
- [ ] **Step 4: 확인** — 빌드가 `built` 가 될 때까지 지켜본 뒤 `curl -s -o /dev/null -w "%{http_code}" https://keonchue.github.io/seojeokbu-catalog/books.csv` → `200`, 페이지 `<title>` 이 카탈로그 것.

### Task 2: 카탈로그가 공개 가격표를 읽는다

**Files:**
- Modify: `seojeokbu-catalog/index.html` (`sale`/`pr` 정의, `applyCsv`, 끝의 `fetch("books.csv")…finally(start)` 묶음, 배지 문구 `50% 할인` 세 곳)
- Create: `seojeokbu-catalog/tests/live-prices.mjs`
- Modify: `seojeokbu-catalog/README.md`

**Interfaces:**
- Consumes: Task 3 이 쓰는 문서 모양 `items: { "<no>": { price: int, salePct: int } }` (REST 응답에서는 `fields.items.mapValue.fields["<no>"].mapValue.fields.price.integerValue` 처럼 문자열 숫자)
- Produces (index.html 최상위):
  - `function readLive(json) -> { [no: number]: { price: number, salePct: number } } | null` — REST 응답 JSON 을 평평하게. 모양이 틀리면 `null`.
  - `function applyLive(items) -> void` — `all`/`DATA` 의 항목 중 번호가 있는 것만 `i.price`, `i.pct` 를 덮어쓴다. 없는 번호는 무시.
  - 항목 필드 `i.pct` (0~100). CSV 에서는 `반값할인`이 O 이면 50, 아니면 0. `sale(i)` 는 `i.pct > 0`.
  - `pr(i)` — `pct>0`: `<s class="was">정가</s><span class="now">할인가</span>`; 할인가 0 이면 `0원`. `i.dc` 면 `할인가 : <최종가>`. 정가 0 이면 `가격 문의`.
  - 배지 문구는 `${i.pct}% 할인`.

- [ ] **Step 1: 시험 쓰기** — `tests/live-prices.mjs`: index.html 의 `<script>` 를 읽어 `won`, `sale`, `pr`, `readLive`, `applyLive` 를 `vm` 에서 돌린다 (최상위 선언 잘라 오기는 POS 의 `tests/grab-index.mjs` 와 같은 방식으로, 이 저장소에 작은 사본을 둔다). 단언:
  - `readLive({ fields: { items: { mapValue: { fields: { "12": { mapValue: { fields: { price: { integerValue: "155000" }, salePct: { integerValue: "50" } } } } } } } } })` → `{ 12: { price: 155000, salePct: 50 } }`
  - `readLive({ error: { code: 403 } })` → `null`; `readLive(null)` → `null`
  - `applyLive({ 12: { price: 150000, salePct: 30 }, 999: { price: 1, salePct: 0 } })` 뒤 `DATA` 의 12번 `price === 150000`, `pct === 30`, `DATA` 어느 칸에도 999번이 생기지 않음
  - `pr({ price: 100000, pct: 30 })` 에 `70,000원` 과 `<s class="was">100,000원</s>`; `pr({ price: 5000, pct: 100 })` 에 `0원`, `가격 문의` 없음; `pr({ price: 0, pct: 0 })` → `가격 문의`; `pr({ price: 20000, pct: 10, dc: true })` 에 `할인가 : 18,000원`
- [ ] **Step 2:** `node tests/live-prices.mjs` → 함수가 없어서 실패.
- [ ] **Step 3: 만들기** — 위 함수들. 시작은 `Promise.all([CSV 읽기, Promise.race([REST fetch → readLive, 3000ms 뒤 null])])` → CSV 적용 → live 있으면 `applyLive` → `start()`. REST 는 `cache: "no-store"`. `applyLive` 는 `start()` 가 `all` 을 채우기 전에 부르므로 `DATA.sections` 를 돈다.
- [ ] **Step 4:** `node tests/live-prices.mjs` → 모두 통과.
- [ ] **Step 5: 브라우저 확인** — `npx serve` 등으로 띄워 (규칙을 아직 안 올렸으므로 403) CSV 값 그대로, 콘솔 오류 없음, 첫 화면이 3초 넘게 비지 않음.
- [ ] **Step 6:** README 고치기 (새 주소, 가격·할인은 POS 앱 → 카탈로그로 자동, `반값할인` 칸은 앱에 연결 안 된 책에만 쓰임).
- [ ] **Step 7:** 작업 트리 리뷰를 Codex 에 맡기고(리뷰만) 맞는 지적은 고친다 → 커밋 `카탈로그가 앱의 가격·할인을 읽는다` → 사용자 확인 후 push, Pages 에서 다시 200 확인.

### Task 3: 앱 — 연결 번호와 공개 가격표 계산 (순수 함수)

**Files:**
- Create: `tests/grab-index.mjs`, `tests/catalog-sync.mjs`
- Modify: `tests/sheet-contract.mjs` (자기 `grab` 대신 `grabFrom` import — 동작 그대로)
- Modify: `index.html` (`promoPct`/`salePrice` 바로 아래에 새 함수들)

**Interfaces:**
- Produces (index.html 최상위, 이후 Task 4·5 가 씀):
  - `function catNoOf(b) -> number | 0` — `b.catalogNo` 를 정수로. 양의 정수가 아니면 0. 문자열 `"12 "` → 12.
  - `function catalogItems(list) -> { [no: string]: { price: number, salePct: number } }` — 택배비 빼고 `catNoOf>0` 인 책만. `list` 는 `createdAt` 오름차순이라 같은 번호면 먼저 온 책. `salePct = promoPct(b)`.
  - `function catalogKey(items) -> string` — 번호 숫자순으로 정렬한 JSON. 같은 내용이면 같은 문자열.
  - `function shouldPublish(meta, lastKey, key) -> boolean` — `!meta.fromCache && !meta.hasPendingWrites && lastKey !== null && key !== lastKey`. (`lastKey === null` 은 공개 문서를 아직 못 읽음)
  - `function parseCatalogRows(text) -> Array<{ no: number, name: string, color: string, cat: string }>` — 카탈로그 `books.csv` (따옴표·BOM 처리, 헤더 `번호,분류,이름,색상`). 번호 없는 줄은 뺀다.
  - `function suggestCatalogMatch(row, list) -> string | null` — 가장 맞는 앱 책 id. 이름·색상을 띄어쓰기 없이 비교해, 앱 제목(`b.title`)이 카탈로그 이름을 품고 색상도 품는 책을 고른다. 여럿이면 제목이 가장 짧은 것. 하나도 없으면 null. 택배비 제외.

- [ ] **Step 1:** `tests/grab-index.mjs` 로 `grab` 을 옮기고 sheet-contract 가 import 하게 바꾼 뒤 `SHEET_TEST_DRY=1 node tests/sheet-contract.mjs` → 전과 같은 출력.
- [ ] **Step 2: 시험 쓰기** — `tests/catalog-sync.mjs` 에서 `isShipping` 과 그 의존, `promoPct`, 위 6개를 꺼내 단언:
  - `catNoOf({catalogNo:"12 "})===12`; `0`, `-3`, `"abc"`, `undefined`, `2.5` → 0
  - `catalogItems([{id:'a',title:'가',price:45000,catalogNo:1}, {id:'b',title:'나',price:30000,catalogNo:1}, {id:'c',title:'다',price:155000,catalogNo:12,salePct:50}, {id:'d',title:'택배비',price:0,catalogNo:3}, {id:'e',title:'라',price:1000}])` → `{ "1": {price:45000,salePct:0}, "12": {price:155000,salePct:50} }` (택배비 판정은 실제 `isShipping` 규칙에 맞는 제목으로)
  - `catalogKey({"12":…,"1":…}) === catalogKey({"1":…,"12":…})`, 값이 다르면 다른 문자열
  - `shouldPublish({fromCache:true,hasPendingWrites:false},'x','y')===false`; pending true → false; `lastKey null` → false; 같은 key → false; 서버·변경 → true
  - `parseCatalogRows('\uFEFF번호,분류,이름,색상,번역,가격\n1,개역개정 성경,지갑식 미니 단본,그레이,,45000\n50,일반 도서,"하루 한 장 성경필사, 시편",,,12800\n')` → 두 줄, 두 번째 `name === '하루 한 장 성경필사, 시편'`
  - `suggestCatalogMatch({name:'지갑식 미니 단본',color:'그레이'}, [{id:'x',title:'개역개정 지갑식 미니 단본 블랙'},{id:'y',title:'개역개정 지갑식 미니단본 그레이'}]) === 'y'`; 맞는 게 없으면 `null`
- [ ] **Step 3:** `node tests/catalog-sync.mjs` → 함수가 없어서 실패.
- [ ] **Step 4:** index.html 에 함수 6개 넣기.
- [ ] **Step 5:** `node tests/catalog-sync.mjs` → 통과, `SHEET_TEST_DRY=1 node tests/sheet-contract.mjs` → 전과 같음.
- [ ] **Step 6:** 커밋은 Task 4 와 같이 한다 (Codex 리뷰를 한 번에).

### Task 4: 앱 — 공개 가격표 쓰기, 책 수정 칸, 연결 창, 규칙

**Files:**
- Modify: `index.html` — import 에 필요한 것(`setDoc`, `serverTimestamp` 가 없으면 추가), `startListeners()`, `editBookModal`, 설정 창(할인 책 관리 칸 아래), 새 `catalogLinkModal()`, `renderAll` 의 다시 그리기
- Modify: `firestore.rules`
- Test: `tests/catalog-sync.mjs` (Task 3 그대로), 미리보기 `node preview/build.mjs && node preview/serve.mjs`

**Interfaces:**
- Consumes: Task 3 의 `catNoOf`, `catalogItems`, `catalogKey`, `shouldPublish`, `parseCatalogRows`, `suggestCatalogMatch`
- Produces: 문서 `public/catalog` = `{ items: catalogItems(books), updatedAt: serverTimestamp() }`; 책 필드 `catalogNo` (숫자 또는 `null`); 원장 줄 `{ reason: '카탈로그 연결', field: 'catalogNo', note: '<옛> → <새>' }`

- [ ] **Step 1: 공개 문서 쓰기** — `let catalogLastKey = null;`. `startListeners()` 에 `onSnapshot(doc(db,'public','catalog'), …)` → 있으면 `catalogKey(data.items||{})`, 없으면 `''`. 책 목록 스냅샷 콜백에서 `const key = catalogKey(catalogItems(books)); if (shouldPublish(snap.metadata, catalogLastKey, key)) setDoc(doc(db,'public','catalog'), { items, updatedAt: serverTimestamp() }).catch(console.error)`. 실패해도 토스트는 띄우지 않는다 (규칙을 올리기 전엔 실패가 정상).
- [ ] **Step 2: 규칙** — `firestore.rules` 의 `meta` 규칙 아래에:
  ```
  // 카탈로그 웹사이트(손님용)가 읽는 가격표 — 번호별 정가·할인율만 들어 있다
  match /public/catalog {
    allow read: if true;
    allow write: if isStore();
  }
  ```
- [ ] **Step 3: 책 수정 칸** — `editBookModal` 의 거래처 칸 아래 `<input id="mCatNo" inputmode="numeric">` (라벨 "카탈로그 번호 (선택)"; 택배비면 숨김). 저장할 때 `catNoOf({catalogNo: 입력})||null`; 다른 책이 같은 번호면 `toast('카탈로그 ${n}번은 이미 「${제목}」에 이어져 있습니다.', 'bad')` 후 저장 중단. `updateDoc` 필드에 `catalogNo` 추가, 바뀌면 원장 줄. 시트로 가는 코드(putQ·queueForSheet·adjustRecord)는 건드리지 않는다.
- [ ] **Step 4: 연결 창** — `catalogLinkModal()` (생김새는 `saleModal` 을 따른다): 카탈로그 CSV 를 `fetch(CATALOG_CSV, {cache:'no-store'})` → `parseCatalogRows`. 줄마다 `번호 · 이름 · 색상` 과 앱 책 `<select>` (첫 칸 "연결 안 함"; 기본값은 이미 그 번호인 책, 없으면 `suggestCatalogMatch`, 추천은 "추천" 표시). 같은 책을 두 줄에 고르면 저장 버튼에서 막는다. "연결 저장" → 바뀐 책만 `runBatched` 로 `catalogNo` 갱신(빠진 책은 `null`) + 원장 줄 + `toast('n권을 카탈로그에 이었습니다.','good')`. CSV 를 못 읽으면 창 안에 "카탈로그 목록을 못 읽었습니다" 한 줄. 설정 창 '할인 책' 칸 아래에 `카탈로그 · 연결 n종` 라벨과 버튼(`id="sCatalog"`, 아이콘 `ic('link')` — 없으면 묶음에 있는 아무 선 아이콘).
- [ ] **Step 5: 미리보기 확인** — `node preview/build.mjs && node preview/serve.mjs` 로 띄워: 책 수정에서 번호를 넣으면 가짜 저장소의 `public/catalog` 에 그 번호가 생기고, 같은 내용으로 다시 저장해도 `setDoc` 이 더 불리지 않는다 (콘솔/가짜 저장소로 확인). 연결 창이 열리고(미리보기에선 CSV 를 실제 주소에서 읽는다), 중복 선택이 막힌다.
- [ ] **Step 6: 시험** — `node tests/catalog-sync.mjs` 통과, `node tests/sheet-contract.mjs` (시험용 시트, 환경변수) 통과.
- [ ] **Step 7: 리뷰** — 작업 트리를 Codex 에 리뷰만 맡긴다. 넘길 원자료: 설계 문서, `firestore.rules` 변경, 공개 문서 모양. 맞는 지적은 고치고, 동의 안 하는 건 이유를 적는다 (최대 2번).
- [ ] **Step 8: 커밋** — `카탈로그 연동 — 공개 가격표(public/catalog), 카탈로그 번호 칸 · 연결 창`. push 는 사용자 확인 후.

### Task 5: 실제로 켜기 (사용자와 함께)

- [ ] **Step 1:** 사용자에게 Firebase 콘솔 → Firestore → 규칙에 새 `firestore.rules` 를 붙여 넣고 게시해 달라고 부탁한다.
- [ ] **Step 2:** POS push 뒤 앱에서 설정 → 카탈로그 연결로 117줄을 잇는다 (사용자가 확인).
- [ ] **Step 3: 확인** — `curl` 로 공개 문서 REST 주소가 200 이고 `items` 에 번호들이 있는지, 카탈로그 사이트에서 앱에서 바꾼 책 하나의 가격이 바뀌어 보이는지.
- [ ] **Step 4:** 채림 님(`chaelym`)을 `keonchue/seojeokbu-catalog` 협업자로 초대할지 사용자에게 묻고, 좋다면 `gh api -X PUT repos/keonchue/seojeokbu-catalog/collaborators/chaelym -f permission=push`.
- [ ] **Step 5:** 사용자에게 남은 일 알리기: 기존 artifact 링크 대신 새 주소 안내, `seojeokbu-catalog-main/` 사본 지울지.
