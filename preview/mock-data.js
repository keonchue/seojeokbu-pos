/* 미리보기용 데이터 — 실제 가게 데이터가 아니다. 날짜는 오늘을 기준으로 매번 새로 만든다. */
const BOOKS = [
  ['개역개정 큰글 성경 (가죽 · 색인)', '성경 - 개역개정', 68000, 42000, 14, 5],
  ['개역개정 슬림 성경 (비닐)', '성경 - 개역개정', 32000, 19000, 23, 8],
  ['개역개정 쉬운 해설 성경', '성경 - 개역개정', 54000, 33000, 3, 5],
  ['개역한글 단본 성경', '성경 - 개역한글', 28000, 17000, 11, 4],
  ['개역한글 소형 합본 (찬송 포함)', '성경 - 개역한글', 36000, 22000, 7, 3],
  ['NIV 한영 해설 성경', '성경 - 한영/영어', 72000, 46000, 5, 2],
  ['ESV 영어 성경 (페이퍼백)', '성경 - 한영/영어', 24000, 15000, 9, 3],
  ['어린이 그림 성경', '성경 - 어린이', 26000, 15500, 12, 4],
  ['새찬송가 (대 · 색인)', '찬송가', 22000, 13000, 18, 6],
  ['새찬송가 (소)', '찬송가', 14000, 8000, 0, 4],
  ['노아의 방주 팝업북', '어린이', 18000, 10500, 6, 2],
  ['다윗과 골리앗 색칠공부', '어린이', 6500, 3500, 30, 10],
  ['예수님이 들려주신 이야기 12권 세트', '어린이', 84000, 55000, 2, 1],
  ['순전한 기독교', '일반', 15800, 10200, 8, 3],
  ['고백록', '일반', 19800, 12800, 4, 2],
  ['매일 묵상 365 (2027)', '일반', 13000, 7800, 40, 10],
  ['성경 책갈피 (가죽)', '기타', 5000, 2200, 25, 5],
  ['택배비', '기타', 4000, null, 0, 0],
];

const NAMES = ['김은혜', '이요셉', '박하늘', '최믿음', '정소망', '한사랑', '윤다니엘', '오에스더', '강바울', '서룻'];

function rng(seed) { let s = seed >>> 0; return () => ((s = (s * 1664525 + 1013904223) >>> 0) / 2 ** 32); }

export function seed() {
  const out = {};
  const now = new Date();
  const t0 = now.getTime();
  BOOKS.forEach(([title, category, price, cost, stock, safety], i) => {
    out['books/b' + i] = {
      title, category, price, cost, stock, safety, sold: 0, supplier: i < 10 ? '생명의말씀사' : '두란노',
      createdAt: t0 - (90 - i) * 864e5,
    };
  });

  const r = rng(20261112);
  const pick = a => a[Math.floor(r() * a.length)];
  const sellable = BOOKS.map((b, i) => i).filter(i => BOOKS[i][0] !== '택배비');
  let n = 0;
  for (let day = 29; day >= 0; day--) {
    const date = new Date(now); date.setDate(now.getDate() - day);
    if (date.getDay() === 1) continue;                 // 월요일은 쉰다
    const count = day === 0 ? 4 : 2 + Math.floor(r() * 6) + (date.getDay() === 0 ? 5 : 0);
    for (let k = 0; k < count; k++) {
      const at = new Date(date); at.setHours(10 + Math.floor(r() * 8), Math.floor(r() * 60), 0, 0);
      if (day === 0 && at > now) at.setTime(now.getTime() - (k + 1) * 23 * 6e4);
      const lines = 1 + (r() < 0.3 ? 1 : 0);
      const items = [];
      for (let j = 0; j < lines; j++) {
        const i = pick(sellable); const [title, category, price, cost] = BOOKS[i];
        const qty = r() < 0.85 ? 1 : 2;
        items.push({ title, category, qty, price, amount: price * qty, cost });
        out['books/b' + i].sold += qty;
      }
      const total = items.reduce((s, x) => s + x.amount, 0);
      const disc = r() < 0.12 ? Math.round(total * 0.1 / 100) * 100 : 0;
      const payable = total - disc;
      const m = r(); const method = m < 0.62 ? '계좌이체' : m < 0.92 ? '현금' : '분할';
      const split = method === '분할' ? { cash: Math.round(payable * 0.4 / 1000) * 1000, transfer: 0 } : null;
      if (split) split.transfer = payable - split.cash;
      const named = method !== '현금' || r() < 0.2;
      out['transactions/t' + (n++)] = {
        time: at.toISOString(), items, qty: items.reduce((s, x) => s + x.qty, 0),
        total, discount: disc, discountPct: disc ? 10 : null, payable, received: payable, diff: 0,
        method, split, discountLabel: disc ? '지역할인' : null,
        customer: named ? pick(NAMES) : null, phone: null, address: null,
      };
    }
  }

  // 입고 기록 몇 줄
  [[0, 10, 6], [1, 20, 9], [8, 15, 13], [15, 30, 20]].forEach(([i, qty, ago], k) => {
    const [title, category, , cost] = BOOKS[i];
    out['stockLogs/l' + k] = {
      time: new Date(t0 - ago * 864e5).toISOString(), bookId: 'b' + i, title, category,
      reason: '입고', field: 'stock', before: 0, after: qty, delta: qty, qty, cost,
      amount: cost * qty, supplier: out['books/b' + i].supplier, txId: null, note: null, by: null,
    };
  });

  const open = new Date(now); open.setHours(10, 2, 0, 0);
  const ymd = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  out['meta/session'] = { open: true, openedAt: open.toISOString(), openedBy: null, bizDate: ymd(now), seq: 120 };
  out['meta/init'] = { seeded: true };
  out['meta/closings'] = { list: [] };
  return out;
}
