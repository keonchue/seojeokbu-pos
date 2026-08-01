/**
 * 서적부 POS → 기존 스프레드시트 미러
 * ---------------------------------------------------------------------------
 * 앱에서 판매·반품·입고가 저장되면 이 웹앱으로 보내고, 여기서 기존 시트의
 * 계좌이체 / 현금 / 입고 탭 맨 아래에 줄을 덧붙인다.
 *
 * 설치 순서
 *   1) 스프레드시트 → 확장 프로그램 → Apps Script → 이 파일 내용을 붙여넣기
 *   2) 함수 목록에서 `탭이름_확인` 실행 → 실행 로그에 나온 실제 탭 이름을
 *      아래 TABS 에 그대로 옮겨 적기
 *   3) `토큰_만들기` 실행 → 로그에 찍힌 토큰을 복사 (앱 설정에 넣을 값)
 *   4) `설치_확인` 실행 → 세 탭의 머리글을 제대로 찾는지 확인
 *   5) 배포 → 새 배포 → 유형 '웹 앱'
 *        - 실행 계정: 나
 *        - 액세스 권한: 링크가 있는 모든 사용자
 *      → 생성된 /exec URL 을 앱 설정에 토큰과 함께 입력
 *
 * 코드를 고친 뒤에는 반드시 '배포 관리 → 수정 → 새 버전'으로 다시 배포해야
 * /exec URL 에 반영된다.
 *
 * 토큰은 이 파일이 아니라 스크립트 속성에 저장된다. (저장소가 public 이라
 * 코드에 적어두면 아무나 시트에 데이터를 밀어넣을 수 있다)
 */

// ── 실제 탭 이름으로 바꿔 주세요 (2단계) ────────────────────────────────
var TABS = {
  계좌이체: '계좌이체',
  현금: '현금',
  입고: '입고',
};

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

/* ===================== 웹앱 진입점 ===================== */

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (!body.token || body.token !== scriptToken()) {
      return json({ ok: false, error: '토큰이 맞지 않습니다.' });
    }
    var items = body.items || [];
    if (!items.length) return json({ ok: true, written: 0, skipped: 0 });

    // 시트를 동시에 건드리면 줄이 겹쳐 쓰일 수 있어 잠근다
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      return json(writeItems(items));
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

// 브라우저로 열어보면 살아 있는지만 알려준다 (토큰·데이터는 노출하지 않는다)
function doGet() {
  return json({ ok: true, service: '서적부 POS 시트 미러' });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ===================== 실제 기록 ===================== */

function writeItems(items) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sent = sentKeys(ss);
  var buckets = {};   // 탭 이름 → 줄 목록. 탭별로 모아서 한 번에 쓴다
  var newKeys = [];
  var skipped = 0;

  items.forEach(function (item) {
    if (!item || !item.key) return;
    if (sent[item.key]) { skipped++; return; }   // 이미 보낸 거래 — 다시 안 적는다
    var tab = tabFor(item);
    if (!tab) return;
    (buckets[tab] = buckets[tab] || []).push.apply(buckets[tab], item.records || []);
    newKeys.push(item.key);
    sent[item.key] = true;
  });

  var written = 0;
  Object.keys(buckets).forEach(function (name) {
    written += appendRecords(sheetByName(ss, name), buckets[name]);
  });
  if (newKeys.length) rememberKeys(ss, newKeys);

  return { ok: true, written: written, skipped: skipped };
}

// 판매·반품은 결제수단에 따라 계좌이체/현금 탭으로, 입고는 입고 탭으로 간다
function tabFor(item) {
  if (item.type === 'stock') return TABS.입고;
  var pay = String((item.records && item.records[0] && item.records[0].결제) || '');
  if (pay.indexOf('현금') === 0) return TABS.현금;
  if (pay.indexOf('계좌이체') === 0) return TABS.계좌이체;
  return TABS.계좌이체;   // 결제수단이 비어 있으면 계좌이체 탭으로 모은다
}

function sheetByName(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('"' + name + '" 탭을 찾을 수 없습니다. TABS 설정을 확인하세요.');
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
    return { row: r + 1, col: col, width: lastCol };
  }
  throw new Error('"' + sheet.getName() + '" 탭에서 머리글(날짜·제품명) 줄을 찾지 못했습니다.');
}

function appendRecords(sheet, records) {
  if (!records.length) return 0;
  var h = findHeader(sheet);

  var width = h.width;
  Object.keys(h.col).forEach(function (k) { width = Math.max(width, h.col[k]); });

  var values = records.map(function (rec) {
    var row = [];
    for (var i = 0; i < width; i++) row.push('');
    Object.keys(rec).forEach(function (key) {
      var c = h.col[key];
      if (!c) return;                       // 이 탭에 없는 열(예: 입고 탭의 주소)은 버린다
      var v = rec[key];
      if (v === null || v === undefined || v === '') return;
      row[c - 1] = (key === '날짜') ? toDate(v) : v;
    });
    return row;
  });

  // 머리글 바로 아래를 침범하지 않도록 마지막 줄 다음부터 쓴다
  var start = Math.max(sheet.getLastRow() + 1, h.row + 1);
  sheet.getRange(start, 1, values.length, width).setValues(values);
  return values.length;
}

// 월별 피벗이 '날짜 - 월'로 묶으려면 글자가 아니라 진짜 날짜여야 한다
function toDate(s) {
  var m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : s;
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
  out.push(scriptToken() ? '토큰: 설정됨' : '토큰: 없음 → 토큰_만들기 를 먼저 실행하세요');

  Object.keys(TABS).forEach(function (kind) {
    var name = TABS[kind];
    var sh = ss.getSheetByName(name);
    if (!sh) { out.push('✗ ' + kind + ' → "' + name + '" 탭 없음'); return; }
    try {
      var h = findHeader(sh);
      var found = Object.keys(h.col).map(function (k) {
        return k + '=' + columnLetter(h.col[k]);
      });
      out.push('✓ ' + kind + ' → "' + name + '" (머리글 ' + h.row + '행, 마지막 ' +
               sh.getLastRow() + '행)\n    ' + found.join(', '));
    } catch (e) {
      out.push('✗ ' + kind + ' → "' + name + '" : ' + e.message);
    }
  });
  Logger.log(out.join('\n'));
}

function columnLetter(n) {
  var s = '';
  while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; }
  return s;
}
