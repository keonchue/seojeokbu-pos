# CLAUDE.md

## 커밋 전 Codex 리뷰

- 코드를 바꾼 뒤 커밋하기 전에는 항상 Codex에게 리뷰를 맡긴다 (`/codex:review`, 커밋 전 작업 트리 대상).
- `/codex:review`를 직접 실행할 수 없으면 codex-rescue에게 같은 범위로 리뷰만 맡긴다 (코드는 수정하지 말라고 명시한다).
- Codex 지적 중 맞는 건 고치고, 동의하지 않는 건 이유를 적어서 사용자에게 보고한다.
- 리뷰 → 수정은 최대 2번까지만 반복한다. 그래도 남는 문제는 커밋하기 전에 사용자에게 물어본다.
- 문서만 바꾼 커밋(예: `*.md`, `docs/`)은 리뷰를 생략한다.

## 시트로 가는 코드를 고칠 때

2026-10-17 ~ 10-26 에 시트 드롭다운 밖의 값이 가서 원장에 반쪽 줄이 970개 쌓였다. 코드 리뷰만으로는
시트 쪽 규칙(드롭다운 · 재고관리 이름 목록 · 수식)을 못 본다. 그래서:

- `apps-script/Code.gs`, 또는 `index.html` 에서 시트로 가는 줄을 만드는 코드(sheetRecords · sheetMethod ·
  stockRecord · adjustRecord · flushSheet 등)를 고치면, 커밋 전에 **시험용 시트**로 계약 시험을 통과시킨다:
  `SHEET_TEST_URL=… SHEET_TEST_TOKEN=… node tests/sheet-contract.mjs` (주소·토큰은 사용자에게 받는다, 저장소에 안 넣는다).
- `Code.gs` 를 고쳤으면 시험용 시트의 Apps Script 에도 같은 코드를 새 버전으로 배포해 달라고 사용자에게 먼저 부탁한다
  (버전이 다르면 시험이 멈춘다). 시험이 통과한 뒤에 실제 시트 배포를 부탁한다.
- 앱에 새 결제방식·구분 값이 생기면 `tests/sheet-contract.mjs` 의 거래 목록과 `Code.gs` 의 `APP_PAY_VALUES` ·
  `APP_KIND_VALUES` 도 같이 고친다.
- 원인을 말할 때는 증거로 확인한 것만 사실로 말하고, 확인 전이면 가설이라고 밝힌다.
- 리뷰를 맡길 때는 내 결론 말고 원자료(시트 규칙 원문, 실제 줄, 오류 메시지)를 넘긴다.
