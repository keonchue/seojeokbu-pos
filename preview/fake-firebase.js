/* 디자인 미리보기 전용 가짜 Firebase.
   index.html 의 Firebase import 를 이 파일로 바꿔 끼우면 로그인 없이 앱이 뜬다.
   모든 읽기·쓰기는 이 탭의 메모리에서만 돌고, 진짜 Firestore · 시트로는 아무것도 가지 않는다.
   새로고침하면 처음 데이터로 돌아간다. */
import { STORE_EMAIL } from '../firebase-config.js';
import { seed } from './mock-data.js';

/* ---------- app / auth ---------- */
export const initializeApp = cfg => ({ cfg });
const user = { email: STORE_EMAIL, uid: 'preview' };
export const getAuth = () => ({ currentUser: user });
export const onAuthStateChanged = (auth, cb) => { setTimeout(() => cb(user), 0); return () => {}; };
export const signInWithEmailAndPassword = async () => ({ user });
export const signOut = async () => {};

/* ---------- firestore ---------- */
const store = new Map();          // 'books/abc' → data
for (const [path, data] of Object.entries(seed())) store.set(path, structuredClone(data));
const listeners = new Set();
let autoId = 0;
const newId = () => 'p' + Date.now().toString(36) + (autoId++).toString(36);

export const persistentLocalCache = () => ({});
export const persistentMultipleTabManager = () => ({});
export const initializeFirestore = () => ({ db: true });

export const collection = (db, name) => ({ kind: 'col', path: name });
export function doc(base, ...segs) {
  if (base && base.kind === 'col') {
    const id = segs.length ? segs.join('/') : newId();
    return ref(base.path + '/' + id);
  }
  return ref(segs.join('/'));
}
const ref = path => ({ kind: 'doc', path, id: path.split('/').pop() });

export const orderBy = (field, dir = 'asc') => ({ t: 'order', field, dir });
export const limit = n => ({ t: 'limit', n });
export const query = (col, ...cs) => ({ kind: 'query', path: col.path, cs });

export const increment = n => ({ __op: 'inc', n });
export const arrayUnion = (...v) => ({ __op: 'union', v });

function applyOps(prev, patch) {
  const out = { ...(prev || {}) };
  for (const [k, v] of Object.entries(patch)) {
    if (v && v.__op === 'inc') out[k] = (Number(out[k]) || 0) + v.n;
    else if (v && v.__op === 'union') out[k] = [...(out[k] || []), ...v.v.filter(x => !(out[k] || []).includes(x))];
    else out[k] = v;
  }
  return out;
}

const docSnap = (path, data) => ({
  id: path.split('/').pop(), ref: ref(path),
  exists: () => data !== undefined, data: () => (data === undefined ? undefined : structuredClone(data)),
  metadata: { hasPendingWrites: false, fromCache: false },
});

function runQuery(q) {
  const prefix = q.path + '/';
  let rows = [...store.entries()]
    .filter(([p]) => p.startsWith(prefix) && !p.slice(prefix.length).includes('/'))
    .map(([p, d]) => docSnap(p, d));
  for (const c of q.cs || []) {
    if (c.t === 'order') {
      const s = c.dir === 'desc' ? -1 : 1;
      rows.sort((a, b) => { const x = a.data()[c.field], y = b.data()[c.field]; return x < y ? -s : x > y ? s : 0; });
    }
  }
  const lim = (q.cs || []).find(c => c.t === 'limit');
  if (lim) rows = rows.slice(0, lim.n);
  return { docs: rows, size: rows.length, empty: !rows.length, forEach: f => rows.forEach(f),
           metadata: { fromCache: false, hasPendingWrites: false } };
}
const snapOf = target => target.kind === 'doc'
  ? docSnap(target.path, store.get(target.path))
  : runQuery(target.kind === 'col' ? { path: target.path, cs: [] } : target);

let notifyQueued = false;
function notify() {
  if (notifyQueued) return;
  notifyQueued = true;
  setTimeout(() => { notifyQueued = false; listeners.forEach(l => { try { l.next(snapOf(l.target)); } catch (e) { console.error(e); } }); }, 0);
}

export function onSnapshot(target, a, b, c) {
  const next = typeof a === 'function' ? a : b;
  const l = { target, next };
  listeners.add(l);
  setTimeout(() => next(snapOf(target)), 0);
  return () => listeners.delete(l);
}

export const getDoc = async r => snapOf(r);
export const getDocs = async q => snapOf(q);
export async function setDoc(r, data, opts) {
  store.set(r.path, opts && opts.merge ? applyOps(store.get(r.path), data) : applyOps({}, data)); notify();
}
export async function addDoc(col, data) { const r = doc(col); await setDoc(r, data); return r; }
export async function updateDoc(r, data) { store.set(r.path, applyOps(store.get(r.path), data)); notify(); }
export async function deleteDoc(r) { store.delete(r.path); notify(); }

export function writeBatch() {
  const ops = [];
  return {
    set(r, d, o) { ops.push(() => setDoc(r, d, o)); return this; },
    update(r, d) { ops.push(() => updateDoc(r, d)); return this; },
    delete(r) { ops.push(() => deleteDoc(r)); return this; },
    async commit() { for (const op of ops) await op(); },
  };
}
export async function runTransaction(db, fn) {
  const tr = {
    get: async r => snapOf(r),
    set(r, d, o) { setDoc(r, d, o); return tr; },
    update(r, d) { updateDoc(r, d); return tr; },
    delete(r) { deleteDoc(r); return tr; },
  };
  return fn(tr);
}
