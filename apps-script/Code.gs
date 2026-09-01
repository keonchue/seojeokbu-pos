/**
 * 서적부 POS → 기존 스프레드시트 미러
 * ---------------------------------------------------------------------------
 * 이 시트의 원본은 '입출고기록' 탭 하나다. 계좌이체모음 / 현금모음 / 입고모음 /
 * 택배비모음 은 데이터가 아니라 A2 의 FILTER('입출고기록'!A:M, 조건) 결과를
 * 보여 주는 화면이고, 재고관리의 현재재고와 월별정리 피벗도 원장만 읽는다.
 *
 *   계좌이체모음!A2 = FILTER('입출고기록'!A:M, REGEXMATCH('입출고기록'!J:J, "계좌이체"))
 *   현금모음!A2     = FILTER('입출고기록'!A:M, REGEXMATCH('입출고기록'!J:J, "현금"))
 *   입고모음!A2     = FILTER('입출고기록'!A:M, REGEXMATCH('입출고기록'!D:D, "입고"))
 *   택배비모음!A2   = FILTER('입출고기록'!A:M, REGEXMATCH('입출고기록'!C:C, "\[기타\] 택배비"))
 *   재고관리!D2     = SUMIFS('입출고기록'!E:E, C:C, A2, D:D, "입고")
 *                   - SUMIFS('입출고기록'!E:E, C:C, A2, D:D, "출고*")
 *   월별정리 피벗    = 입출고기록!B:I  (열 전체 참조라 원장이 길어져도 알아서 따라온다)
 *
 * (xlsx 로 내보내면 이 피벗 범위가 B1:I1930 처럼 그 시점 시트 크기로 굳어 보인다.
 *  엑셀 피벗 캐시가 무한 열 참조를 못 담아서 그럴 뿐, 시트 쪽 설정은 B:I 다.
 *  2026-08-31 에 피벗 편집기로 직접 확인함 — 내보낸 파일만 보고 판단하지 말 것)
 *
 * 그래서 앱은 거래를 **원장에만** 적는다. 화면 탭에 직접 쓰면 그 줄은 원장에 없어서
 * 재고·집계 어디에도 안 잡히고(고아 줄), FILTER 가 자랄 자리를 막아 탭 전체를
 * #REF! 로 날려 버린다. (2026-08-31 이전 코드가 실제로 그러고 있었다)
 *
 * 예외는 **새 책 등록**뿐이다(2026-09-01). `재고관리` 와 `안전재고관리` 는 FILTER 화면이
 * 아니라 사람이 관리하는 기준정보 목록이고, 원장에 거래가 쌓여도 새 책 줄이 저절로
 * 생기지는 않는다. 줄이 없으면 그 책은 현재재고·상태 판정에서 통째로 빠진다.
 * 두 탭 다 가나다 순이라 맨 밑이 아니라 순서 맞는 자리에 끼워 넣는다 (`addBooks`).
 *
 * 원장의 표기 관습 — 이걸 따라야 위 수식이 맞는다:
 *   판매      구분 출고   수량 +n   출고단가·출고금액 양수   결제방식 현금 / 계좌이체(지역할인)
 *   입고      구분 입고   수량 +n   입고단가·입고금액        결제방식 비움, 거래처는 고객명 칸
 *   환불·반품 구분 입고   수량 +n   출고단가·출고금액 음수   결제방식 현금 환불
 *   재고조정  구분 입고/출고 수량 절대값  금액 비움           결제방식 재고조정
 *   택배비    구분 택배비 수량 +n   출고단가·출고금액 양수   결제방식 결제수단
 * (반품을 구분 '반품' 으로 적으면 "입고" 도 "출고*" 도 아니라 재고가 안 되돌아간다)
 *
 * 이 표기대로 적으면 화면 탭도 저절로 맞는다: 환불 줄은 결제방식에 '현금'이 들어 있어
 * 현금모음에 뜨고 구분이 '입고'라 입고모음에도 뜬다(손으로 적어 온 방식 그대로).
 * 재고조정 줄은 결제방식이 결제수단이 아니라 결제수단 탭에는 안 뜨고,
 * 택배비 줄은 제품명이 '[기타] 택배비'라 택배비모음이 알아서 집어 간다.
 *
 * 앱에서 기록을 지우면 시트의 그 줄도 지운다 (숨김 `기록ID` 열로 찾는다).
 * 반대로 단가·안전재고는 앱이 재고관리 탭에서 읽어 갈 수 있고, 손으로 적은 줄은
 * 앱이 `rows` 로 읽어 가져간 뒤 `mark` 로 기록ID 를 찍어 두 번 가져오지 않게 한다.
 *
 * 설치 순서
 *   1) 스프레드시트 → 확장 프로그램 → Apps Script → 이 파일 내용을 붙여넣기
 *   2) `토큰_만들기` 실행 → 로그에 찍힌 토큰을 복사 (앱 설정에 넣을 값)
 *   3) `설치_확인` 실행 → 원장의 머리글을 제대로 찾는지 확인
 *   4) 배포 → 새 배포 → 유형 '웹 앱'
 *        - 실행 계정: 나
 *        - 액세스 권한: 링크가 있는 모든 사용자
 *      → 생성된 /exec URL 을 앱 설정에 토큰과 함께 입력
 *
 * 코드를 고친 뒤에는 반드시 '배포 관리 → 기존 배포의 ✏️ → 버전: 새 버전' 으로
 * 다시 배포해야 /exec URL 에 반영된다. ('새 배포'를 만들면 주소가 바뀐다)
 *
 * 토큰은 이 파일이 아니라 스크립트 속성에 저장된다. (저장소가 public 이라
 * 코드에 적어두면 아무나 시트에 데이터를 밀어넣을 수 있다)
 */

/* 이 파일을 고칠 때마다 이 값을 올린다 (그리고 index.html 의 SHEET_MIRROR_VERSION 도 같이).
   앱이 두 값을 비교해서 '새 배포를 안 했다'를 스스로 알려준다.
   주소창에 /exec 를 그대로 열어봐도 지금 배포된 버전이 보인다. */
var MIRROR_VERSION = '2026-09-01a';

// ── 원본 원장. 앱이 거래를 쓰는 곳은 여기 하나뿐이다 (새 책 줄만 재고관리에도) ──
var LEDGER_TAB = '입출고기록';

/* 예전(2026-08-31 이전) 코드가 줄을 붙이던 화면 탭들.
   이제 여기에 쓰지는 않지만, 그때 붙은 줄을 지울 수 있어야 하고
   `앱줄_원장으로_옮기기` 가 훑을 대상이기도 해서 이름은 남겨 둔다. */
var LEGACY_TABS = ['계좌이체모음', '현금모음', '입고모음', '택배비모음'];

// 머리글 이름은 탭마다 조금씩 다르다. (입고 탭은 '입출고구분' 대신 '입고',
// '고객명' 대신 '거래처' 를 쓴다) 그래서 위치가 아니라 이름으로 찾는다.
var HEADER_ALIASES = {
  날짜:     ['날짜'],
  제품명:   ['제품명', '상품명'],
  구분:     ['입출고구분', '입고', '구분'],
  수량:     ['수량'],
  입고단가: ['입고단가'],
  출고단가: ['출고단가'],
  입고금액: ['입고금액'],
  출고금액: ['출고금액'],
  결제:     ['결제방식', '결제'],
  고객:     ['고객명', '거래처'],
  연락처:   ['연락처'],
  주소:     ['주소'],
};

var HEADER_SCAN_ROWS = 8;      // 머리글이 첫 줄이 아닐 수 있어 위쪽 몇 줄을 훑는다
var SENT_SHEET = '_전송기록';   // 같은 거래를 두 번 적지 않기 위한 숨김 시트
var TOKEN_KEY = 'MIRROR_TOKEN';

/* 앱에서 거래를 지웠을 때 시트 줄도 같이 지우려면 "이 줄이 어느 거래인지"를 알아야 한다.
   그래서 원장 맨 끝에 이 이름의 열을 만들어 두고 거래 키를 적는다. 열은 숨겨 두므로
   평소에는 보이지 않고, 손으로 적은 줄은 이 칸이 비어 있어 영향을 받지 않는다.
   화면 탭들의 FILTER 가 A:M 만 읽으므로 주소(M) 오른쪽에 두면 화면에 새지 않는다. */
var KEY_HEADER = '기록ID';

/* ===================== 웹앱 진입점 ===================== */

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (!body.token || body.token !== scriptToken()) {
      return json({ ok: false, error: '토큰이 맞지 않습니다.' });
    }
    // 앱이 '지금 배포된 버전이 뭐냐'만 물어볼 때 (버전은 json() 이 붙여 준다)
    if (body.action === 'ping') return json({ ok: true });

    // 앱이 기준정보(단가·안전재고)를 읽어 갈 때
    if (body.action === 'books') {
      return json(readStockTab(SpreadsheetApp.getActiveSpreadsheet()));
    }

    // 앱이 '손으로 적은 줄'을 가져가려고 읽을 때 / 가져간 뒤 표시를 남길 때
    if (body.action === 'rows') {
      return json(readUnlinkedRows(SpreadsheetApp.getActiveSpreadsheet(), body.from, body.to));
    }
    if (body.action === 'mark') {
      return withLock(function (ss) { return markRows(ss, body.marks || []); });
    }

    var items = body.items || [];
    var deletes = body.deletes || [];
    var marks = body.marks || [];
    var books = body.books || [];
    if (!items.length && !deletes.length && !marks.length && !books.length) {
      return json({ ok: true, written: 0, skipped: 0, removed: 0, marked: 0, added: 0 });
    }

    return withLock(function (ss) {
      var out = writeItems(ss, items);
      out.removed = deleteKeys(ss, deletes);
      var m = markRows(ss, marks);
      out.marked = m.marked;
      if (m.mismatched.length) out.mismatched = m.mismatched;
      if (m.busy.length) out.busy = m.busy;
      // 새 책은 기준정보 탭에도 한 줄 (원장 기록과 달리 실패해도 나머지는 그대로 간다)
      var bk = addBooks(ss, books);
      out.added = bk.added;
      out.bookSkipped = bk.skipped;
      if (bk.failed.length) out.bookErrors = bk.failed;
      return out;
    });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

// 시트를 동시에 건드리면 줄이 겹쳐 쓰일 수 있어 잠근다
function withLock(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return json(fn(SpreadsheetApp.getActiveSpreadsheet()));
  } finally {
    lock.releaseLock();
  }
}

// 브라우저로 열어보면 살아 있는지와 배포된 버전만 알려준다 (토큰·데이터는 노출하지 않는다)
function doGet() {
  return json({ ok: true, service: '서적부 POS 시트 미러' });
}

// 모든 응답에 배포된 버전을 실어 보낸다 — 앱이 이걸 보고 재배포가 필요한지 판단한다
function json(obj) {
  obj.version = MIRROR_VERSION;
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ===================== 실제 기록 ===================== */

function writeItems(ss, items) {
  var sent = sentKeys(ss);
  var entries = [];
  var newKeys = [];
  var skipped = 0;

  items.forEach(function (item) {
    if (!item || !item.key) return;
    if (sent[item.key]) { skipped++; return; }   // 이미 보낸 거래 — 다시 안 적는다
    (item.records || []).forEach(function (rec) {
      entries.push({ key: item.key, rec: rec });
    });
    newKeys.push(item.key);
    sent[item.key] = true;
  });

  var written = entries.length
    ? appendRecords(sheetByName(ss, LEDGER_TAB), entries)
    : 0;
  if (newKeys.length) rememberKeys(ss, newKeys);

  return { ok: true, written: written, skipped: skipped };
}

function sheetByName(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('"' + name + '" 탭을 찾을 수 없습니다. 탭 이름을 확인하세요.');
  return sh;
}

/**
 * 머리글 줄을 찾아 이름 → 열 번호 표를 만든다.
 * A열이 비어 있고 B열부터 시작하는 지금 구조도, 나중에 열을 옮겨도 그대로 동작한다.
 */
function findHeader(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (!lastRow || !lastCol) throw new Error('"' + sheet.getName() + '" 탭이 비어 있습니다.');

  var rows = Math.min(HEADER_SCAN_ROWS, lastRow);
  var grid = sheet.getRange(1, 1, rows, lastCol).getDisplayValues();

  for (var r = 0; r < rows; r++) {
    var cells = grid[r].map(function (v) { return String(v).trim(); });
    if (cells.indexOf('날짜') === -1 || cells.indexOf('제품명') === -1) continue;
    var col = {};
    Object.keys(HEADER_ALIASES).forEach(function (key) {
      var names = HEADER_ALIASES[key];
      for (var i = 0; i < names.length; i++) {
        var at = cells.indexOf(names[i]);
        if (at !== -1) { col[key] = at + 1; return; }
      }
    });
    var keyAt = cells.indexOf(KEY_HEADER);
    if (keyAt !== -1) col.기록ID = keyAt + 1;
    return { row: r + 1, col: col, width: lastCol, cells: cells };
  }
  throw new Error('"' + sheet.getName() + '" 탭에서 머리글(날짜·제품명) 줄을 찾지 못했습니다.');
}

function appendRecords(sheet, entries) {
  if (!entries.length) return 0;
  var h = findHeader(sheet);
  ensureKeyColumn(sheet, h);

  // 머리글에 이름이 붙은 열까지만 쓴다. 시트 폭(getLastColumn) 만큼 쓰면
  // 오른쪽의 빈 서식 칸까지 훑고 지나가 남의 수식을 지울 수 있다.
  var width = 1;
  Object.keys(h.col).forEach(function (k) { width = Math.max(width, h.col[k]); });

  var values = entries.map(function (entry) {
    var rec = entry.rec;
    var row = [];
    for (var i = 0; i < width; i++) row.push('');
    Object.keys(rec).forEach(function (key) {
      var c = h.col[key];
      if (!c || c > width) return;          // 이 탭에 없는 열은 버린다
      var v = rec[key];
      if (v === null || v === undefined || v === '') return;
      row[c - 1] = (key === '날짜') ? toDate(v) : v;
    });
    row[h.col.기록ID - 1] = entry.key;       // 나중에 이 줄을 되찾기 위한 표시
    return row;
  });

  var start = Math.max(lastFilledRow(sheet, h) + 1, h.row + 1);
  sheet.getRange(start, 1, values.length, width).setValues(values);
  return values.length;
}

/* 실제 기록이 끝나는 줄. getLastRow() 를 쓰면 안 된다 — 원장은 아래쪽 수백 줄에
   출고단가 자동계산 수식이 미리 깔려 있어서(IF(D="출고", VLOOKUP…)) 빈 줄인데도
   '마지막 줄'로 잡힌다. 그 뒤에 붙이면 원장 한가운데에 빈 줄 수백 개가 남아
   사람이 장부를 훑을 수 없다. 그래서 제품명이 실제로 적혀 있는 마지막 줄을 기준으로 삼는다.
   (수식·피벗은 열 전체를 보므로 어느 쪽이든 집계는 맞다 — 이건 읽기 쉬우라고 하는 것) */
function lastFilledRow(sheet, h) {
  var last = sheet.getLastRow();
  var col = h.col.제품명;
  if (!col || last <= h.row) return h.row;
  var vals = sheet.getRange(h.row + 1, col, last - h.row, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim()) return h.row + 1 + i;
  }
  return h.row;
}

/* 기록ID 열이 없으면 만들고 숨긴다. 이미 있으면 아무것도 하지 않는다.

   자리는 '머리글에 이름이 붙은 마지막 열 바로 오른쪽'(원장에서는 주소 M 다음의 N).
   getLastColumn()+1 을 쓰면 안 된다 — 원장은 빈 서식 때문에 폭이 32라 AG 열에 박히고,
   그러면 사람이 눈으로 찾을 수 없다. 그 자리가 이미 다른 값으로 차 있으면 물러선다. */
function ensureKeyColumn(sheet, h) {
  if (h.col.기록ID) return h.col.기록ID;

  var want = 1;
  Object.keys(h.col).forEach(function (k) { want = Math.max(want, h.col[k] + 1); });
  var occupied = h.cells && h.cells[want - 1];
  var c = occupied ? sheet.getLastColumn() + 1 : want;

  sheet.getRange(h.row, c).setValue(KEY_HEADER);
  try { sheet.hideColumns(c); } catch (e) {}   // 숨기기에 실패해도 기록 자체는 되게 둔다
  h.col.기록ID = c;
  h.width = Math.max(h.width, c);
  return c;
}

/* 앱에서 지운 거래의 줄을 시트에서도 없앤다.
   기록ID 가 비어 있는 줄(손으로 적은 것)은 절대 건드리지 않는다.
   예전 코드가 화면 탭에 붙여 둔 줄도 지울 수 있게 그쪽도 같이 훑는다. */
function deleteKeys(ss, keys) {
  var want = {};
  (keys || []).forEach(function (k) { if (k) want[String(k)] = true; });
  if (!Object.keys(want).length) return 0;

  var removed = 0;
  [LEDGER_TAB].concat(LEGACY_TABS).forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var h;
    try { h = findHeader(sheet); } catch (e) { return; }
    if (!h.col.기록ID) return;               // 이 탭엔 표시가 없다 = 지울 것도 없다

    var last = sheet.getLastRow();
    if (last <= h.row) return;
    var vals = sheet.getRange(h.row + 1, h.col.기록ID, last - h.row, 1).getValues();
    // 아래에서 위로 지워야 남은 줄의 행 번호가 밀리지 않는다
    for (var i = vals.length - 1; i >= 0; i--) {
      if (want[String(vals[i][0])]) { sheet.deleteRow(h.row + 1 + i); removed++; }
    }
  });
  return removed;
}

// 월별 피벗이 '날짜 - 월'로 묶으려면 글자가 아니라 진짜 날짜여야 한다
function toDate(s) {
  var m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : s;
}

function ymd(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(v == null ? '' : v).trim();
  // '2026. 8. 1' / '2026. 7 .26' 처럼 손으로 적은 글자 날짜도 받아 준다
  var m = s.match(/^(\d{4})\s*[.\-\/]\s*(\d{1,2})\s*[.\-\/]?\s*(\d{1,2})?/);
  if (!m) return s;
  var mm = ('0' + m[2]).slice(-2), dd = ('0' + (m[3] || '1')).slice(-2);
  return m[1] + '-' + mm + '-' + dd;
}

/* ===================== 시트 → 앱 (손으로 적은 줄 가져가기) =====================
   앱을 만든 뒤에도 며칠은 시트에 손으로 적어서 앱에 그 매출이 없다.
   기록ID 칸이 빈 줄이 곧 '손으로 적은 줄'이므로, 기간을 받아 그것만 돌려준다. */

function readUnlinkedRows(ss, from, to) {
  var sheet = ss.getSheetByName(LEDGER_TAB);
  if (!sheet) return { ok: false, error: '"' + LEDGER_TAB + '" 탭을 찾을 수 없습니다.' };

  var h = findHeader(sheet);
  var last = sheet.getLastRow();
  if (last <= h.row) return { ok: true, rows: [] };

  var width = Math.max(h.width, h.col.기록ID || 0);
  var grid = sheet.getRange(h.row + 1, 1, last - h.row, width).getValues();
  var lo = from ? String(from) : '', hi = to ? String(to) : '';

  var pick = function (row, key) {
    var c = h.col[key];
    if (!c) return '';
    var v = row[c - 1];
    return (v === null || v === undefined) ? '' : v;
  };

  var out = [];
  for (var i = 0; i < grid.length; i++) {
    var row = grid[i];
    var name = String(pick(row, '제품명')).trim();
    if (!name) continue;
    if (h.col.기록ID && String(row[h.col.기록ID - 1] || '').trim()) continue;   // 이미 앱이 아는 줄
    var date = ymd(pick(row, '날짜'));
    if (lo && date < lo) continue;
    if (hi && date > hi) continue;
    out.push({
      row: h.row + 1 + i,
      날짜: date,
      제품명: name,
      구분: String(pick(row, '구분')).trim(),
      수량: pick(row, '수량'),
      입고단가: pick(row, '입고단가'),
      출고단가: pick(row, '출고단가'),
      입고금액: pick(row, '입고금액'),
      출고금액: pick(row, '출고금액'),
      결제: String(pick(row, '결제')).trim(),
      고객: String(pick(row, '고객')).trim(),
      연락처: String(pick(row, '연락처')).trim(),
      주소: String(pick(row, '주소')).trim(),
    });
  }
  return { ok: true, rows: out };
}

/* 앱이 가져간 줄에 기록ID 를 찍는다 — 두 번 가져오지 않기 위한 표시.
   행 번호만 믿으면 안 된다(그 사이에 사람이 행을 넣고 뺐을 수 있다). 그래서
   날짜·제품명·수량이 앱이 본 것과 같을 때만 쓰고, 다르면 건너뛰고 알려준다. */
function markRows(ss, marks) {
  if (!marks.length) return { ok: true, marked: 0, mismatched: [], busy: [] };

  var sheet = sheetByName(ss, LEDGER_TAB);
  var h = findHeader(sheet);
  ensureKeyColumn(sheet, h);
  var last = sheet.getLastRow();

  var marked = 0, mismatched = [], busy = [];
  marks.forEach(function (m) {
    var r = Number(m && m.row);
    if (!r || r <= h.row || r > last) { mismatched.push(m && m.row); return; }

    var width = Math.max(h.width, h.col.기록ID);
    var row = sheet.getRange(r, 1, 1, width).getValues()[0];
    var at = function (key) { return h.col[key] ? row[h.col[key] - 1] : ''; };

    var same = ymd(at('날짜')) === String(m.날짜) &&
               String(at('제품명')).trim() === String(m.제품명).trim() &&
               Number(at('수량') || 0) === Number(m.수량 || 0);
    if (!same) { mismatched.push(r); return; }

    var already = String(row[h.col.기록ID - 1] || '').trim();
    if (already && already !== String(m.key)) { busy.push(r); return; }

    sheet.getRange(r, h.col.기록ID).setValue(m.key);
    marked++;
  });
  return { ok: true, marked: marked, mismatched: mismatched, busy: busy };
}

/* ===================== 시트 → 앱 (기준정보 읽기) =====================
   앱으로 가져오는 건 사람이 직접 적는 값(단가·안전재고)뿐이다.
   '현재재고'는 시트가 원장을 집계해 만든 수식이라, 그걸 앱에 되돌리면
   앱 → 시트 → 앱 으로 도는 순환이 된다. 참고용으로 같이 보내되 앱에서 기본은 꺼 둔다. */

var STOCK_TAB = '재고관리';

/* 재고관리 탭의 머리글 줄과 열 자리를 찾는다 (읽기·쓰기가 같이 쓴다).
   이 탭은 제품명 칸에 머리글이 없다. 그래서 '단가' 열을 찾고 그 왼쪽을 제품명으로 본다.
   찾은 자리는 0부터 세는 번호다 — 시트에 넘길 때는 +1 할 것. */
function stockHeader(sheet) {
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (!lastRow || !lastCol) return null;
  var grid = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  for (var r = 0; r < Math.min(HEADER_SCAN_ROWS, grid.length); r++) {
    var cells = grid[r].map(function (v) { return String(v).trim(); });
    var price = cells.indexOf('단가');
    if (price < 1 || cells.indexOf('안전재고') === -1) continue;
    return {
      row: r, grid: grid, lastRow: lastRow, lastCol: lastCol,
      at: { 제품명: price - 1, 단가: price,
            안전재고: cells.indexOf('안전재고'), 현재재고: cells.indexOf('현재재고'),
            상태: cells.indexOf('상태') },
    };
  }
  return null;
}

function readStockTab(ss) {
  var sheet = ss.getSheetByName(STOCK_TAB);
  if (!sheet) return { ok: false, error: '"' + STOCK_TAB + '" 탭을 찾을 수 없습니다.' };

  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (!lastRow || !lastCol) return { ok: false, error: '"' + STOCK_TAB + '" 탭이 비어 있습니다.' };

  var h = stockHeader(sheet);
  if (!h) {
    return { ok: false, error: '"' + STOCK_TAB + '" 탭에서 단가·안전재고 머리글을 찾지 못했습니다.' };
  }
  var grid = h.grid, head = h.row, at = h.at;

  // 수식이 만든 #N/A 같은 오류값은 빈칸으로 넘긴다 (앱에서 '건드리지 않음'으로 해석된다)
  var books = [];
  for (var i = head + 1; i < grid.length; i++) {
    var row = grid[i];
    var name = String(row[at.제품명] || '').trim();
    if (!name) continue;
    var pick = function (c) {
      var v = (c >= 0) ? String(row[c] || '').trim() : '';
      return (v.charAt(0) === '#') ? '' : v;
    };
    books.push({ 제품명: name, 단가: pick(at.단가),
                 안전재고: pick(at.안전재고), 현재재고: pick(at.현재재고) });
  }
  // 화면 탭에 아직 옛 미러가 붙여 둔 줄이 남아 있으면 현재재고는 그만큼 모자란 값이다.
  // 앱이 이걸 보고 '재고 맞추기'를 잠근다 — 안 그러면 옮기기 전에 맞췄다가 두 번 보정된다.
  return { ok: true, books: books, stray: countStrayRows(ss) };
}

/* 화면 탭(계좌이체모음 등)에 남아 있는 '앱이 직접 쓴 줄' 수.
   0이 아니면 `앱줄_원장으로_옮기기` 를 아직 안 돌린 것이다. */
function countStrayRows(ss) {
  var n = 0;
  LEGACY_TABS.forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var h;
    try { h = findHeader(sheet); } catch (e) { return; }
    if (!h.col.기록ID || sheet.getLastRow() <= h.row) return;
    sheet.getRange(h.row + 1, h.col.기록ID, sheet.getLastRow() - h.row, 1)
         .getValues().forEach(function (r) { if (String(r[0]).trim()) n++; });
  });
  return n;
}

/* ===================== 앱 → 시트 (새 책을 기준정보에 추가) =====================
   원장과 달리 `재고관리` 는 사람이 관리하는 목록이고 **가나다 순으로 정렬돼 있다.**
   그래서 맨 밑에 붙이지 않고 순서에 맞는 자리에 끼워 넣는다.
   안전재고·현재재고·상태는 줄마다 같은 수식이고 A열만 상대참조로 다르다. 그래서 수식을
   여기서 지어내지 않고 **이웃 줄에서 복사해 온다** — 나중에 시트에서 수식을 고쳐도
   이 코드를 같이 고칠 필요가 없다. */
var SAFETY_TAB = '안전재고관리';

/* 시트가 쓰는 가나다 정렬을 그대로 흉내 낸다.
   2026-09-01 에 재고관리 212종·안전재고관리 213종 전부와 맞춰 본 규칙:
   **공백·기호·숫자 < 한글 < 영문**. (그래서 '할인성경 (비닐 x)' 가 '(비닐ㅇ)' 보다 앞이고,
   'ESV 영어성경' 은 '히브리어 원어성경' 뒤에 온다. 코드포인트 순으로 하면 둘 다 뒤집힌다) */
function charRank(ch) {
  var o = ch.charCodeAt(0);
  if ((o >= 0xAC00 && o <= 0xD7A3) ||     // 가~힣
      (o >= 0x3131 && o <= 0x318E) ||     // ㄱ ㅇ 같은 낱자
      (o >= 0x1100 && o <= 0x11FF)) return 1;
  if ((o >= 65 && o <= 90) || (o >= 97 && o <= 122) || o > 0x7F) return 2;   // 영문·그 밖의 글자
  return 0;                                                                  // 공백·기호·숫자
}
function koCompare(a, b) {
  a = String(a); b = String(b);
  var n = Math.min(a.length, b.length);
  for (var i = 0; i < n; i++) {
    var ra = charRank(a.charAt(i)), rb = charRank(b.charAt(i));
    if (ra !== rb) return ra - rb;
    var ca = a.charCodeAt(i), cb = b.charCodeAt(i);
    if (ca !== cb) return ca - cb;
  }
  return a.length - b.length;
}

/* 이름 목록에서 새 이름이 들어갈 자리(1부터 세는 줄 번호)를 찾는다.
   이미 같은 이름이 있으면 -1 — 사람이 적어 둔 단가·안전재고를 덮어쓰면 안 되기 때문이다. */
function placeFor(names, name, first) {
  for (var i = 0; i < names.length; i++) {
    if (names[i] === name) return -1;
  }
  for (var j = 0; j < names.length; j++) {
    if (koCompare(names[j], name) > 0) return first + j;
  }
  return first + names.length;
}

/* 한 줄 끼워 넣고, 이웃 줄에서 수식·서식을 복사해 온다. 새 줄 번호를 돌려준다. */
function insertSorted(sheet, headRow, nameCol, width, names, name) {
  var at = placeFor(names, name, headRow + 1);
  if (at < 0) return -1;

  var appended = at > headRow + names.length;
  if (!appended) sheet.insertRowBefore(at);
  if (names.length) {
    // 복사해 올 이웃: 윗줄이 있으면 윗줄, 새 줄이 맨 앞이면 방금 밀려난 아랫줄
    var src = (at > headRow + 1) ? at - 1 : at + 1;
    sheet.getRange(src, 1, 1, width).copyTo(sheet.getRange(at, 1, 1, width));
  }
  sheet.getRange(at, nameCol).setValue(name);
  return at;
}

/* `안전재고관리` 는 제품명·안전재고 두 칸짜리 단순한 표다. 재고관리의 안전재고 칸이
   이 탭을 VLOOKUP 하므로, 여기 줄이 없으면 새 책의 안전재고가 #N/A 로 뜬다. */
function addSafetyRow(ss, name, safety) {
  var sheet = ss.getSheetByName(SAFETY_TAB);
  if (!sheet) return false;
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (!lastRow || !lastCol) return false;

  var grid = sheet.getRange(1, 1, Math.min(HEADER_SCAN_ROWS, lastRow), lastCol).getDisplayValues();
  var head = -1, nameAt = -1, safeAt = -1;
  for (var r = 0; r < grid.length; r++) {
    var cells = grid[r].map(function (v) { return String(v).trim(); });
    var s = cells.indexOf('안전재고');
    if (s === -1) continue;
    head = r; safeAt = s;
    nameAt = cells.indexOf('제품명');
    if (nameAt === -1) nameAt = (s > 0) ? s - 1 : 0;
    break;
  }
  if (head === -1) return false;

  var names = readNames(sheet, head + 1, nameAt + 1, lastRow);
  var at = insertSorted(sheet, head + 1, nameAt + 1,
                        Math.max(nameAt, safeAt) + 1, names, name);
  if (at < 0) return false;
  sheet.getRange(at, safeAt + 1).setValue(Number(safety) || 0);
  return true;
}

/* 제품명 열에 실제로 적혀 있는 이름들 (머리글 아래부터, 빈 줄에서 끊는다) */
function readNames(sheet, headRow, nameCol, lastRow) {
  if (lastRow <= headRow) return [];
  var vals = sheet.getRange(headRow + 1, nameCol, lastRow - headRow, 1).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var v = String(vals[i][0]).trim();
    if (!v) break;              // 목록은 중간에 비지 않는다. 빈 줄부터는 여백이다
    out.push(v);
  }
  return out;
}

/* 앱에서 새로 등록한 책을 재고관리(+안전재고관리)에 한 줄씩 만든다.
   이미 있는 이름은 건너뛴다 — 사람이 손본 단가를 앱 값으로 덮어쓰지 않기 위해서다. */
function addBooks(ss, books) {
  var added = 0, skipped = 0, failed = [];
  if (!books || !books.length) return { added: 0, skipped: 0, failed: failed };

  var sheet = sheetByName(ss, STOCK_TAB);
  books.forEach(function (b) {
    var name = String((b && b.제품명) || '').trim();
    if (!name) return;
    try {
      var h = stockHeader(sheet);
      if (!h) throw new Error('단가·안전재고 머리글을 찾지 못했습니다.');
      var width = Math.max(h.at.제품명, h.at.단가, h.at.안전재고,
                           h.at.현재재고, h.at.상태) + 1;
      var names = readNames(sheet, h.row + 1, h.at.제품명 + 1, sheet.getLastRow());
      var at = insertSorted(sheet, h.row + 1, h.at.제품명 + 1, width, names, name);
      if (at < 0) { skipped++; return; }
      var price = Number(b.단가);
      sheet.getRange(at, h.at.단가 + 1).setValue(isNaN(price) ? '' : price);
      addSafetyRow(ss, name, b.안전재고);
      added++;
    } catch (e) {
      failed.push(name + ' — ' + (e.message || e));
    }
  });
  return { added: added, skipped: skipped, failed: failed };
}

/* ===================== 중복 방지 ===================== */

function sentSheet(ss) {
  var sh = ss.getSheetByName(SENT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SENT_SHEET);
    sh.getRange(1, 1, 1, 2).setValues([['키', '보낸 시각']]);
    sh.hideSheet();
  }
  return sh;
}

function sentKeys(ss) {
  var sh = sentSheet(ss);
  var n = sh.getLastRow() - 1;
  var map = {};
  if (n > 0) {
    sh.getRange(2, 1, n, 1).getValues().forEach(function (r) {
      if (r[0]) map[String(r[0])] = true;
    });
  }
  return map;
}

function rememberKeys(ss, keys) {
  var sh = sentSheet(ss);
  var now = new Date();
  var rows = keys.map(function (k) { return [k, now]; });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 2).setValues(rows);
}

/* ===================== 설치 도우미 (직접 실행) ===================== */

function 탭이름_확인() {
  var names = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function (s, i) {
    return (i + 1) + '. ' + s.getName() + '  (' + s.getLastRow() + '줄)';
  });
  Logger.log('이 스프레드시트의 탭 목록:\n' + names.join('\n'));
}

function 토큰_만들기() {
  var token = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty(TOKEN_KEY, token);
  Logger.log('토큰이 만들어졌습니다. 앱 설정에 이 값을 넣으세요:\n\n' + token);
}

function scriptToken() {
  return PropertiesService.getScriptProperties().getProperty(TOKEN_KEY);
}

function 설치_확인() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = [];
  out.push('이 코드의 버전: ' + MIRROR_VERSION +
           '  (배포된 버전과 다르면 배포 관리 → ✏️ → 새 버전)');
  out.push(scriptToken() ? '토큰: 설정됨' : '토큰: 없음 → 토큰_만들기 를 먼저 실행하세요');

  var sh = ss.getSheetByName(LEDGER_TAB);
  if (!sh) {
    out.push('✗ 원장 "' + LEDGER_TAB + '" 탭이 없습니다.');
  } else {
    try {
      var h = findHeader(sh);
      var found = Object.keys(h.col).map(function (k) {
        return k + '=' + columnLetter(h.col[k]);
      });
      out.push('✓ 원장 "' + LEDGER_TAB + '" (머리글 ' + h.row + '행, 마지막 ' +
               sh.getLastRow() + '행)\n    ' + found.join(', '));
      if (!h.col.기록ID) out.push('    기록ID 열은 첫 기록 때 자동으로 만들어집니다.');
    } catch (e) {
      out.push('✗ 원장 "' + LEDGER_TAB + '" : ' + e.message);
    }
  }

  out.push('');
  out.push('화면 탭(수식이 채우는 곳 — 앱은 여기에 쓰지 않습니다):');
  LEGACY_TABS.forEach(function (name) {
    var s = ss.getSheetByName(name);
    if (!s) { out.push('  · ' + name + ' — 없음'); return; }
    var a1 = String(s.getRange(1, 1).getFormula() || '');
    var stray = 0;
    try {
      var hh = findHeader(s);
      if (hh.col.기록ID && s.getLastRow() > hh.row) {
        s.getRange(hh.row + 1, hh.col.기록ID, s.getLastRow() - hh.row, 1)
         .getValues().forEach(function (r) { if (String(r[0]).trim()) stray++; });
      }
    } catch (e) {}
    out.push('  · ' + name + (a1.indexOf('FILTER') !== -1 ? ' — FILTER 수식 ✓' : ' — A1에 수식 없음 ⚠') +
             (stray ? '\n      ⚠ 앱이 직접 쓴 줄 ' + stray + '개가 남아 있습니다 → 앱줄_원장으로_옮기기 실행' : ''));
  });
  Logger.log(out.join('\n'));
}

function columnLetter(n) {
  var s = '';
  while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; }
  return s;
}

/* ===================== 일회성 정리 (직접 실행) =====================
   2026-08-31 이전의 미러는 원장이 아니라 화면 탭(계좌이체모음 등) 맨 밑에 줄을 붙였다.
   그 줄들은 FILTER 결과 바로 아래에 떠 있는 고아라서
     · 원장에 없으니 재고관리·월별정리·시트16 어디에도 안 잡히고
     · FILTER 가 자랄 자리를 막아, 손으로 한 줄만 더 적어도 탭이 #REF! 로 통째로 사라진다.
   이 함수가 그 줄들을 원장으로 옮기고, 화면 탭을 원래의 순수한 수식 화면으로 되돌린다.
   기록ID 가 있는 줄만 건드리므로 손으로 적은 줄과 FILTER 결과는 그대로다. 한 번만 실행하면 된다. */
function 앱줄_원장으로_옮기기() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var out = [];
  try {
    var ledger = sheetByName(ss, LEDGER_TAB);
    var lh = findHeader(ledger);
    ensureKeyColumn(ledger, lh);

    var all = [];
    LEGACY_TABS.forEach(function (name) {
      var sheet = ss.getSheetByName(name);
      if (!sheet) { out.push('· ' + name + ' — 탭 없음'); return; }

      var h;
      try { h = findHeader(sheet); }
      catch (e) { out.push('· ' + name + ' — 머리글 못 찾음: ' + e.message); return; }
      if (!h.col.기록ID) { out.push('· ' + name + ' — 기록ID 열 없음 (옮길 것 없음)'); return; }

      var last = sheet.getLastRow();
      if (last <= h.row) { out.push('· ' + name + ' — 비어 있음'); return; }

      var width = Math.max(h.width, h.col.기록ID);
      var grid = sheet.getRange(h.row + 1, 1, last - h.row, width).getValues();

      var entries = [], rowsToDelete = [];
      for (var i = 0; i < grid.length; i++) {
        var row = grid[i];
        var key = String(row[h.col.기록ID - 1] || '').trim();
        if (!key) continue;                      // 손으로 적은 줄 · FILTER 결과 — 안 건드린다
        var rec = {};
        Object.keys(HEADER_ALIASES).forEach(function (k) {
          var c = h.col[k];
          if (!c) return;
          var v = row[c - 1];
          if (v === null || v === undefined || v === '') return;
          rec[k] = v;
        });
        if (!rec.제품명) continue;
        rec.구분 = ledgerKind(rec.구분, rec.결제, rec.출고금액);
        entries.push({ key: key, rec: rec });
        rowsToDelete.push(h.row + 1 + i);
      }

      if (entries.length) {
        all = all.concat(entries);
        // 아래에서 위로 지워야 남은 줄의 행 번호가 밀리지 않는다
        for (var d = rowsToDelete.length - 1; d >= 0; d--) sheet.deleteRow(rowsToDelete[d]);
      }
      /* 앱이 만들었던 기록ID 열은 '지우지' 말고 '비운다'.
         이 탭들의 FILTER 는 A:M 열을 채우는데, 입고모음은 기록ID 가 하필 그 안쪽 L열에 있다.
         열을 통째로 지우면 배열이 자리를 못 잡아 탭이 통째로 흔들릴 수 있다.
         내용만 지우면 빈 열 하나가 남을 뿐이고 화면은 원래대로 돌아간다. */
      sheet.getRange(h.row, h.col.기록ID, Math.max(sheet.getLastRow() - h.row + 1, 1), 1).clearContent();
      out.push('· ' + name + ' — ' + entries.length + '줄을 원장으로 옮기고 기록ID 열을 비웠습니다');
    });

    // 탭별로 모은 것을 그대로 붙이면 원장 끝이 계좌이체 → 현금 → 입고 순으로 엉킨다.
    // 날짜순으로 정렬해 붙여야 원장을 눈으로 훑을 수 있다.
    all.sort(function (a, b) {
      var x = ymd(a.rec.날짜), y = ymd(b.rec.날짜);
      return x < y ? -1 : (x > y ? 1 : 0);
    });
    if (all.length) appendRecords(ledger, all);

    out.unshift('원장 "' + LEDGER_TAB + '" 으로 모두 ' + all.length + '줄을 옮겼습니다.');
    out.push('');
    out.push('이제 네 화면 탭이 #REF! 없이 보이는지, 옮긴 줄이 수식 결과에 섞여 들어왔는지 확인하세요.');
  } finally {
    lock.releaseLock();
  }
  Logger.log(out.join('\n'));
}

/* 옛 줄의 '입출고구분'을 원장 관습으로 맞춘다.
   앱이 반품을 '반품'으로 적어 왔는데, 원장에서는 "입고"도 "출고*"도 아니라
   재고관리 수식이 재고를 되돌리지 않는다. 시트의 손기록 관습대로 '입고'로 바꾼다.
   (금액은 이미 음수라 매출도 그만큼 줄어든다 — 손으로 적어 온 환불 줄과 같은 모양) */
function ledgerKind(kind, pay, amount) {
  var k = String(kind || '').trim();
  if (k === '반품') return '입고';
  if (k === '') return (Number(amount || 0) < 0) ? '입고' : '출고';
  return k;
}
