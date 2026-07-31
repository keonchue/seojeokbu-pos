/* ─────────────────────────────────────────────────────────────
   서적부 POS 서비스워커

   목적: GitHub Pages나 Firebase CDN(gstatic)이 잠깐 죽어도,
        와이파이가 끊겨도 앱 화면이 그대로 뜨게 만든다.
        (데이터 동기화는 Firestore의 오프라인 캐시가 따로 담당)

   ※ 앱을 수정해서 배포할 때는 아래 VERSION 숫자를 꼭 올릴 것.
      올려야 사용자 기기에 새 버전이 내려간다.
   ───────────────────────────────────────────────────────────── */

const VERSION = 'v2';
const CACHE = `seojeokbu-pos-${VERSION}`;

// 앱 껍데기 — 이것만 있으면 오프라인에서도 화면이 뜬다
const APP_SHELL = [
  './',
  './index.html',
  './firebase-config.js',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png',
  './apple-touch-icon.png',
];

// Firebase SDK (외부 CDN) — 캐시해두지 않으면 CDN 장애 시 앱이 아예 안 뜬다
const FIREBASE_VER = '10.12.2';
const VENDOR = [
  'firebase-app.js',
  'firebase-auth.js',
  'firebase-firestore.js',
].map(f => `https://www.gstatic.com/firebasejs/${FIREBASE_VER}/${f}`);

// 실시간 통신 경로 — 절대 캐시하면 안 되는 도메인
const NEVER_CACHE = [
  'firestore.googleapis.com',
  'identitytoolkit.googleapis.com',
  'securetoken.googleapis.com',
  'firebaseinstallations.googleapis.com',
];

self.addEventListener('install', event => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE);
    // 앱 껍데기는 반드시 받아야 함
    await cache.addAll(APP_SHELL);
    // 외부 CDN은 실패해도 설치를 막지 않는다 (다음 온라인 접속 때 다시 시도)
    await Promise.all(VENDOR.map(url =>
      cache.add(new Request(url, { mode: 'cors' })).catch(() => {})
    ));
    self.skipWaiting();
  })());
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

// 페이지가 "지금 바로 새 버전으로 교체" 요청을 보낼 때
self.addEventListener('message', e => {
  if (e.data === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);
  if (NEVER_CACHE.some(h => url.hostname.endsWith(h))) return;  // 네트워크로 직행

  // 페이지 이동: 네트워크 우선, 실패하면 캐시된 앱 껍데기
  if (req.mode === 'navigate') {
    event.respondWith((async () => {
      try {
        const fresh = await fetch(req);
        (await caches.open(CACHE)).put('./index.html', fresh.clone());
        return fresh;
      } catch {
        const cache = await caches.open(CACHE);
        return (await cache.match('./index.html')) || (await cache.match('./')) ||
          new Response('오프라인 상태이고 저장된 앱도 없습니다.', {
            status: 503, headers: { 'Content-Type': 'text/plain; charset=utf-8' }
          });
      }
    })());
    return;
  }

  const isVendor = url.hostname === 'www.gstatic.com';
  const isOwn = url.origin === self.location.origin;
  if (!isVendor && !isOwn) return;

  // 나머지 정적 자원: 캐시 먼저 주고(빠름), 뒤에서 조용히 갱신
  event.respondWith((async () => {
    const cache = await caches.open(CACHE);
    const hit = await cache.match(req);
    const network = fetch(req).then(res => {
      if (res && (res.ok || res.type === 'opaque')) cache.put(req, res.clone());
      return res;
    }).catch(() => null);

    if (hit) { event.waitUntil(network); return hit; }
    const res = await network;
    return res || new Response('', { status: 504 });
  })());
});
