#!/usr/bin/env node
/* 카탈로그 연동 시험 — index.html 의 진짜 함수를 꺼내 돌려 본다.
   실행: node tests/catalog-sync.mjs

   앱은 책에 카탈로그 번호(catalogNo)를 붙여 두고, 이어진 책들의 정가·할인율만 Firestore
   public/catalog 에 써 둔다. 손님용 카탈로그 사이트(keonchue.github.io/seojeokbu-catalog)가 그걸 읽는다.
   재고·원가·매출은 이 문서에 들어가면 안 된다. */
import vm from 'vm';
import { grabFrom, indexSrc } from './grab-index.mjs';

const src = indexSrc();
const NAMES = ['isShipping', 'promoPct', 'catNoOf', 'catalogItems', 'catalogKey', 'shouldPublish',
  'parseCatalogRows', 'suggestCatalogMatch'];
const ctx = vm.createContext({ console });
vm.runInContext(NAMES.map(n => grabFrom(src, n)).join('\n') +
  `\n;globalThis.__api = { ${NAMES.join(', ')} };`, ctx);
const { catNoOf, catalogItems, catalogKey, shouldPublish, parseCatalogRows, suggestCatalogMatch } = ctx.__api;

let fails = 0;
const eq = (name, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) fails++;
  console.log(`${ok ? '통과' : '실패'}  ${name}${ok ? '' : `\n      받은 값: ${JSON.stringify(got)}\n      바란 값: ${JSON.stringify(want)}`}`);
};

// catNoOf — 이상한 값은 '연결 없음'
eq('catNoOf: 문자열 "12 " → 12', catNoOf({ catalogNo: '12 ' }), 12);
eq('catNoOf: 숫자 그대로', catNoOf({ catalogNo: 7 }), 7);
for (const bad of [0, -3, 'abc', undefined, null, 2.5, '']) eq(`catNoOf: ${JSON.stringify(bad)} → 0`, catNoOf({ catalogNo: bad }), 0);

// catalogItems — 택배비 · 번호 없는 책 빼고, 같은 번호면 먼저 온(먼저 만든) 책
const list = [
  { id: 'a', title: '가', price: 45000, catalogNo: 1, stock: 3, cost: 30000 },
  { id: 'b', title: '나', price: 30000, catalogNo: 1 },
  { id: 'c', title: '다', price: 155000, catalogNo: 12, salePct: 50 },
  { id: 'd', title: '택배비', price: 0, catalogNo: 3 },
  { id: 'e', title: '라', price: 1000 },
];
eq('catalogItems: 번호별 정가·할인율만', catalogItems(list), { 1: { price: 45000, salePct: 0 }, 12: { price: 155000, salePct: 50 } });

// catalogKey — 순서가 달라도 같은 내용이면 같은 문자열
const k1 = catalogKey({ 12: { price: 1, salePct: 0 }, 1: { price: 2, salePct: 0 } });
const k2 = catalogKey(JSON.parse('{"1":{"salePct":0,"price":2},"12":{"price":1,"salePct":0}}'));
eq('catalogKey: 순서 무관', k1 === k2, true);
eq('catalogKey: 값이 다르면 다름', k1 === catalogKey({ 12: { price: 1, salePct: 10 }, 1: { price: 2, salePct: 0 } }), false);

// shouldPublish — 서버에서 온 스냅샷이고 보낼 쓰기가 없을 때, 바뀐 경우만
const server = { fromCache: false, hasPendingWrites: false };
eq('shouldPublish: 캐시 스냅샷이면 안 씀', shouldPublish({ fromCache: true, hasPendingWrites: false }, 'x', 'y'), false);
eq('shouldPublish: 보낼 쓰기가 남았으면 안 씀', shouldPublish({ fromCache: false, hasPendingWrites: true }, 'x', 'y'), false);
eq('shouldPublish: 공개 문서를 아직 못 읽었으면 안 씀', shouldPublish(server, null, 'y'), false);
eq('shouldPublish: 같으면 안 씀', shouldPublish(server, 'y', 'y'), false);
eq('shouldPublish: 서버 · 바뀜 → 씀', shouldPublish(server, 'x', 'y'), true);
eq('shouldPublish: 공개 문서가 없음("") · 바뀜 → 씀', shouldPublish(server, '', 'y'), true);

// parseCatalogRows — 카탈로그 books.csv (BOM · 따옴표 속 쉼표)
const rows = parseCatalogRows('﻿번호,분류,이름,색상,번역,가격\r\n1,개역개정 성경,지갑식 미니 단본,그레이,,45000\r\n50,일반 도서,"하루 한 장 성경필사, 시편",,,12800\r\n,빈 번호,무시,,,0\r\n');
eq('parseCatalogRows: 두 줄', rows.length, 2);
eq('parseCatalogRows: 첫 줄', rows[0], { no: 1, name: '지갑식 미니 단본', color: '그레이', cat: '개역개정 성경' });
eq('parseCatalogRows: 따옴표 속 쉼표', rows[1] && rows[1].name, '하루 한 장 성경필사, 시편');

// suggestCatalogMatch — 이름 + 색상을 띄어쓰기 없이 비교
const app = [
  { id: 'x', title: '개역개정 지갑식 미니 단본 블랙' },
  { id: 'y', title: '개역개정 지갑식 미니단본 그레이' },
  { id: 'z', title: '택배비 지갑식 미니 단본 그레이' },
];
eq('suggestCatalogMatch: 색상까지 맞는 책', suggestCatalogMatch({ name: '지갑식 미니 단본', color: '그레이' }, app), 'y');
eq('suggestCatalogMatch: 없으면 null', suggestCatalogMatch({ name: '없는 성경', color: '' }, app), null);
eq('suggestCatalogMatch: 색상 없는 줄은 이름만, 여럿이면 제목 짧은 것',
  suggestCatalogMatch({ name: '지갑식 미니 단본', color: '' }, app), 'x');

console.log(fails ? `\n${fails}개 실패` : '\n모두 통과');
process.exit(fails ? 1 : 0);
