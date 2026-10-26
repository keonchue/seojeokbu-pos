#!/usr/bin/env node
/* 시트 계약 시험 — 앱이 시트로 보내는 줄을 **시험용 시트**에 실제로 보내 본다.

   왜: 2026-10-17 ~ 10-26 에 시트의 결제방식 드롭다운 밖의 값('계좌이체(김인숙)')이 가서
   시트가 줄을 반쯤 쓰고 거절했고, 앱이 같은 묶음을 계속 다시 보내 반쪽 줄이 970개 쌓였다.
   코드만 읽는 리뷰로는 시트 쪽 규칙(드롭다운 · 재고관리 이름 목록)을 못 본다. 그래서
   index.html 의 진짜 줄 만들기 함수(sheetRecords 등)를 그대로 꺼내 와서 줄을 만들고,
   시트가 받는지 · 거절하는지 · 반쪽 줄이 남는지 · 다시 보내도 안 겹치는지를 확인한 뒤 지운다.

   실행 (시트 관련 코드를 고치면 커밋 전에 반드시):
     SHEET_TEST_URL=<시험용 /exec 주소> SHEET_TEST_TOKEN=<토큰> node tests/sheet-contract.mjs
   시험용 시트의 스크립트는 apps-script/Code.gs 와 같은 버전으로 배포돼 있어야 한다
   (버전이 다르면 멈춘다). 주소와 토큰은 저장소에 넣지 않는다. 실제 장부 주소로는 돌리지 않는다.
   보내지 않고 만들어질 줄만 보려면: SHEET_TEST_DRY=1 node tests/sheet-contract.mjs */
import fs from 'fs';
import vm from 'vm';
import path from 'path';
import { fileURLToPath } from 'url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const URL_ = process.env.SHEET_TEST_URL, TOKEN = process.env.SHEET_TEST_TOKEN;
const DRY = !!process.env.SHEET_TEST_DRY;
if (!DRY && (!URL_ || !TOKEN)) { console.error('SHEET_TEST_URL · SHEET_TEST_TOKEN 을 넣어 주세요 (시험용 시트 것).'); process.exit(2); }

/* ---------- index.html 에서 줄 만들기 함수 꺼내기 ---------- */
const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
const src = html.match(/<script type="module">([\s\S]*?)<\/script>/)[1];

// 이름으로 최상위 선언 하나를 통째로 잘라 온다 (괄호 · 문자열 · 템플릿을 세어 끝을 찾는다)
function grab(name) {
  const m = new RegExp(`^(?:const|let|function)\\s+${name.replace('$', '\\$')}\\b`, 'm').exec(src);
  if (!m) throw new Error(`index.html 에서 ${name} 을(를) 못 찾았습니다 — 이름이 바뀌었으면 이 시험도 고치세요`);
  let i = m.index, depth = 0, seenBody = false;
  const isFn = m[0].startsWith('function');
  const stack = [];
  for (; i < src.length; i++) {
    const c = src[i], top = stack[stack.length - 1];
    if (top === '"' || top === "'") { if (c === '\\') i++; else if (c === top) stack.pop(); continue; }
    if (top === '`') { if (c === '\\') i++; else if (c === '`') stack.pop(); else if (c === '$' && src[i + 1] === '{') { stack.push('${'); i++; depth++; } continue; }
    if (top === '/') { if (c === '\\') i++; else if (c === '/') stack.pop(); continue; }
    if (c === '/' && src[i + 1] === '/') { i = src.indexOf('\n', i); continue; }
    if (c === '/' && src[i + 1] === '*') { i = src.indexOf('*/', i) + 1; continue; }
    if (c === '"' || c === "'" || c === '`') { stack.push(c); continue; }
    if (c === '/' && /[=(,:!&|?{;]\s*$/.test(src.slice(Math.max(0, i - 20), i))) { stack.push('/'); continue; }
    // 함수는 인자 괄호 밖(깊이 0)에서 여는 { 가 본문이다 — function f({ a, b }) 의 { 와 헷갈리지 않게
    if (c === '{' || c === '(' || c === '[') { if (c === '{' && depth === 0) seenBody = true; depth++; continue; }
    if (c === '}' || c === ')' || c === ']') {
      depth--;
      if (top === '${' && c === '}') { stack.pop(); continue; }
      if (isFn && seenBody && depth === 0 && c === '}') return src.slice(m.index, i + 1);
      continue;
    }
    if (!isFn && depth === 0 && c === ';') return src.slice(m.index, i + 1);
  }
  throw new Error(`${name} 의 끝을 못 찾았습니다`);
}

const NAMES = ['CATS', 'BIBLE_SUBS', 'SHEET_CATS', 'knownCat', 'bibleSub', 'catOf', 'sheetCat', 'sheetTitle',
  'catForTitle', 'isSplitTx', 'splitParts', '_splitCache', 'payParts', 'isFreeTx', 'ymd',
  'SHEET_DISC_LABELS', 'discLabelOf', 'sheetMethod', 'sheetCustomer', 'SHEET_PAY_RE', 'sheetSafeRecord',
  'spreadCut', 'netAmounts', 'sheetRecords', 'stockRecord', 'adjustRecord'];
const ctx = vm.createContext({ books: [], console });
vm.runInContext(NAMES.map(grab).join('\n') +
  '\n;globalThis.__api = { sheetRecords, stockRecord, adjustRecord, sheetSafeRecord, sheetTitle };', ctx);
const api = ctx.__api;

/* ---------- 시험용 시트 부르기 ---------- */
async function call(body) {
  const r = await fetch(URL_, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ token: TOKEN, ...body }), redirect: 'follow' });
  const t = await r.text();
  try { return JSON.parse(t); } catch { throw new Error('시트 응답이 JSON 이 아님: ' + t.slice(0, 200)); }
}

let fails = 0;
const check = (ok, msg) => { console.log(`${ok ? '  ✓' : '  ✗'} ${msg}`); if (!ok) fails++; };

const gsVersion = (fs.readFileSync(path.join(ROOT, 'apps-script', 'Code.gs'), 'utf8').match(/MIRROR_VERSION = '([^']+)'/) || [])[1];
if (!DRY) {
  const ping = await call({ action: 'ping' });
  console.log(`시험용 시트 스크립트 ${ping.version} · 저장소 Code.gs ${gsVersion} · 시트 "${ping.sheet}"`);
  // 실제 장부에 대고 돌리면 거래·책 줄을 쓰고 지운다 — 시험용 시트 이름이 아니면 아무것도 안 하고 멈춘다
  if (!String(ping.sheet || '').includes('[시험용]')) {
    console.error('✗ 시험용 시트가 아닙니다 (시트 이름에 "[시험용]" 이 없음). 실제 장부 주소로는 돌리지 않습니다.');
    process.exit(1);
  }
  if (ping.version !== gsVersion) {
    console.error('✗ 버전이 다릅니다. 시험용 시트의 Apps Script 에 지금 Code.gs 를 붙여넣고 새 버전으로 배포한 뒤 다시 돌리세요.');
    process.exit(1);
  }
}

/* ---------- 앱이 실제로 만드는 거래들 ---------- */
// 2099-12-31 로 적는다 — 시험 줄이 진짜 장부 날짜와 섞이지 않고, 끝나고 날짜로 찾아 지울 수 있게
const T = '2099-12-31T03:00:00.000Z', DAY = '2099-12-31';
const run = Date.now().toString(36);
const BOOK = { title: '천로역정', category: '일반' }, BIBLE = { title: '대조 성경', category: '성경 - 개역한글' };
const SHIP = { title: '택배비', category: '기타' };
const it = (b, qty, price) => ({ ...b, qty, price, amount: qty * price });
const tx = o => ({ time: T, customer: null, phone: null, address: null, discount: 0, discountLabel: null, split: null, ...o });
const sale = (items, o = {}) => {
  const total = items.reduce((s, i) => s + i.amount, 0);
  const discount = o.discount || 0;
  return tx({ items, total, payable: total - discount, received: total - discount, method: '계좌이체', ...o });
};

const good = {
  '현금 판매': sale([it(BOOK, 1, 15000)], { method: '현금' }),
  '현금 판매 · 연락처 주소': sale([it(BOOK, 1, 15000)], { method: '현금', customer: '시험 손님', phone: '010-0000-0000', address: '어딘가 1' }),
  '계좌이체 + 택배비': sale([it(BOOK, 2, 15000), it(SHIP, 1, 3100)], { customer: '시험' }),
  '지역할인': sale([it(BOOK, 1, 15000)], { discount: 1500, discountLabel: '지역할인' }),
  '직원할인': sale([it(BIBLE, 1, 50000)], { method: '현금', discount: 5000, discountLabel: '직원할인' }),
  '할인 이름 비움': sale([it(BOOK, 1, 15000)], { discount: 1000 }),
  '할인 칸에 손님 이름 (사고 재현)': sale([it(BOOK, 2, 15000)], { discount: 3000, discountLabel: '김시험' }),
  '할인 칸에 메모 + 고객명': sale([it(BOOK, 1, 15000)], { discount: 1000, discountLabel: '차액 현금으로 드림', customer: '박시험' }),
  '분할결제 (책 + 택배비)': sale([it(BOOK, 1, 15000), it(SHIP, 1, 3100)], { method: '분할', split: { cash: 3100, transfer: 15000 } }),
  '분할결제 + 할인': sale([it(BOOK, 3, 15000)], { method: '분할', discount: 4500, discountLabel: '지역할인', split: { cash: 10000, transfer: 30500 } }),
  '0원 출고 (열람용)': sale([it(BOOK, 1, 15000)], { method: '현금', discount: 15000, discountLabel: '열람용' }),
  '현금 반품': tx({ type: 'return', items: [it(BOOK, -1, 15000)], total: -15000, payable: -15000, received: -15000, method: '현금' }),
  '계좌이체 반품 + 택배비': tx({ type: 'return', items: [it(BOOK, -1, 15000), it(SHIP, -1, 3100)], total: -18100, payable: -18100, received: -18100, method: '계좌이체' }),
  '분할 건 0원 반품': tx({ type: 'return', items: [it(BOOK, -1, 15000)], total: 0, payable: 0, received: 0, method: '분할' }),
  '교환 (차액 받음)': tx({ type: 'exchange', items: [it(BOOK, -1, 15000), it(BIBLE, 1, 50000)], total: 35000, payable: 35000, received: 35000, method: '계좌이체' }),
  '교환 (차액 없음 · 원래 분할)': tx({ type: 'exchange', items: [it(BOOK, -1, 15000), it(BOOK, 1, 15000)], total: 0, payable: 0, received: 0, method: '분할' }),
};
const items = [];
let n = 0;
for (const [label, t] of Object.entries(good)) items.push({ key: `test:${run}:${++n}`, label, records: api.sheetRecords(t) });
items.push({ key: `test:${run}:${++n}`, label: '입고', records: [api.stockRecord({ ...BOOK, qty: 10, cost: 13500, supplier: '시험 거래처', time: T })] });
items.push({ key: `test:${run}:${++n}`, label: '재고조정 +', records: [api.adjustRecord({ ...BOOK, delta: 2, note: '시험', time: T })] });
items.push({ key: `test:${run}:${++n}`, label: '재고조정 −', records: [api.adjustRecord({ ...BOOK, delta: -2, note: '시험', time: T })] });
// 이름 바꾸기 — 새 이름 줄(book) + 재고 옮기기 한 쌍
const NEW = `[일반] 시험용 새이름 ${run}`;
items.push({ key: `test:${run}:${++n}`, label: '이름 바꾸기 (재고 옮김)', records: [
  api.adjustRecord({ ...BOOK, delta: -3, note: '시험 이름 바꿈', time: T }),
  api.adjustRecord({ title: NEW.replace('[일반] ', ''), category: '일반', delta: 3, note: '시험 이름 바꿈', time: T }),
] });
// 일부러 틀린 것 — 거절돼야 하고, 반쪽 줄을 남기면 안 되고, 옆 거래를 막으면 안 된다
const bad = [
  { key: `test:${run}:bad1`, label: '(거절돼야 함) 결제방식 목록 밖 — 보내기 직전 고치기 없이',
    records: [{ 날짜: DAY, 제품명: '[일반] 천로역정', 구분: '출고', 수량: 1, 출고단가: 15000, 출고금액: 15000, 결제: '계좌이체(목록에없음)' }] },
  { key: `test:${run}:bad2`, label: '(거절돼야 함) 재고관리에 없는 제품명',
    records: [{ 날짜: DAY, 제품명: '[일반] 없는 책 시험', 구분: '출고', 수량: 1, 출고단가: 1000, 출고금액: 1000, 결제: '현금' }] },
];

if (DRY) {
  for (const x of [...items, ...bad]) {
    console.log(x.label);
    for (const r of x.records.map(api.sheetSafeRecord))
      console.log(`    ${r.제품명} | ${r.구분} | ${r.수량} | ${r.출고금액 ?? r.입고금액 ?? ''} | ${r.결제 ?? ''} | ${r.고객 ?? ''}`);
  }
  process.exit(0);
}

/* ---------- 앱이 보내는 모양 그대로 (보내기 직전 고치기 sheetSafeRecord 포함) ---------- */
const send = list => list.map(x => ({ key: x.key, records: x.records.map(api.sheetSafeRecord) }));
const keys = [...items, ...bad].map(x => x.key);
const before = await call({ action: 'books' });
const stockOf = (res, name) => { const r = (res.books || []).find(b => String(b.제품명).trim() === name); return r ? String(r.현재재고) : null; };

let ok = true;
try {
  console.log('\n1) 처음 보내기');
  const r1 = await call({ books: [{ 제품명: NEW, 단가: 15000, 안전재고: 0 }],
    items: [...send(items.slice(0, 5)), { key: bad[0].key, records: bad[0].records }, ...send(items.slice(5)), { key: bad[1].key, records: bad[1].records }] });
  check(r1.ok, `시트 응답 ok (${r1.error || ''})`);
  const rejected = new Set((r1.rejected || []).map(r => r.key));
  for (const x of items) check(!rejected.has(x.key), `받음 — ${x.label}`);
  for (const x of bad) check(rejected.has(x.key), `거절됨 — ${x.label}`);
  const want = items.reduce((s, x) => s + x.records.length, 0);
  check(r1.written === want, `적힌 줄 ${r1.written} = 보낸 정상 줄 ${want}`);
  // 기록ID 가 있든 없든 그 날짜의 줄을 다 센다 — 거절된 줄이 조금이라도 남았으면 여기서 걸린다
  const c1 = await call({ action: 'countDate', date: DAY });
  check(c1.count === want, `시트에 실제로 있는 시험 날짜 줄 ${c1.count} = ${want} (거절된 줄은 한 줄도 안 남음)`);

  console.log('\n2) 반쪽 줄이 남았나');
  const rows = await call({ action: 'rows', from: DAY, to: DAY });
  check((rows.rows || []).length === 0, `시험 날짜의 기록ID 없는 줄 ${(rows.rows || []).length}개 (0이어야 함)`);

  console.log('\n3) 같은 묶음을 다시 보내도 안 겹치나 (앱이 재시도하는 상황)');
  const r2 = await call({ items: [...send(items), ...bad.map(x => ({ key: x.key, records: x.records }))] });
  check(r2.written === 0, `새로 적힌 줄 ${r2.written} (0이어야 함)`);
  check(r2.skipped === items.length, `이미 있음으로 건너뜀 ${r2.skipped} = ${items.length}`);
  const rows2 = await call({ action: 'rows', from: DAY, to: DAY });
  check((rows2.rows || []).length === 0, '여전히 반쪽 줄 없음');
  const c2 = await call({ action: 'countDate', date: DAY });
  check(c2.count === want, `다시 보낸 뒤에도 시험 날짜 줄 ${c2.count} = ${want}`);

  console.log('\n4) 매일 점검이 거절을 알아채나');
  const hz = await call({ action: 'health' });
  check(hz.ok && (hz.problems || []).some(p => p.includes('거절')), '점검 결과에 "거절한 기록" 이 나옴');
  check(!(hz.problems || []).some(p => p.includes('반쪽 줄') || p.includes('두 번 적힌')), '점검 결과에 반쪽 줄 · 중복 줄 없음');
  for (const p of hz.problems || []) console.log('     · ' + p.slice(0, 160));
} catch (e) {
  ok = false;
  console.error('시험 중 오류:', e.message);
} finally {
  console.log('\n5) 치우기');
  // 치우다 실패하면 몇 번 더 해 보고, 그래도 남으면 무엇이 남았는지 알린다
  for (let tryN = 1; tryN <= 3; tryN++) {
    try {
      const c = await call({ deletes: keys, unbooks: [{ 제품명: NEW, bookId: 'test' }] });
      const left = await call({ action: 'countDate', date: DAY });
      const after = await call({ action: 'books' });
      check(c.ok, `시험 줄 ${c.removed}줄 지움 · 시험 책 줄 ${c.dropped}개 뺌`);
      check(stockOf(after, '[일반] 천로역정') === stockOf(before, '[일반] 천로역정'), `천로역정 재고가 시험 전과 같음 (${stockOf(after, '[일반] 천로역정')})`);
      check(stockOf(after, NEW) === null, '시험용 새 이름 줄이 재고관리에 안 남음');
      check(left.count === 0, `시험 날짜(${DAY}) 줄이 하나도 안 남음 (${left.count})`);
      break;
    } catch (e) {
      console.error(`  치우기 ${tryN}번째 실패: ${e.message}`);
      if (tryN === 3) { fails++; console.error(`  ✗ 시험용 시트에 ${DAY} 날짜 줄과 "${NEW}" 가 남았을 수 있습니다. 손으로 지워 주세요.`); }
    }
  }
}

console.log(fails || !ok ? `\n✗ 실패 ${fails}건 — 커밋하지 마세요` : '\n✓ 모두 통과');
process.exit(fails || !ok ? 1 : 0);
