/* 디자인 미리보기 만들기 — 진짜 index.html 을 그대로 쓰되 Firebase 만 가짜로 바꿔 끼운다.
     node preview/build.mjs            → preview/after.html  (지금 작업 트리의 index.html)
     node preview/build.mjs main       → preview/before.html (main 브랜치의 index.html)
   만든 파일은 .gitignore 에 걸려 있다. 보려면 node preview/serve.mjs */
import { readFileSync, writeFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const ref = process.argv[2];
const src = ref
  ? execFileSync('git', ['show', `${ref}:index.html`], { cwd: here, encoding: 'utf8', maxBuffer: 64 << 20 })
  : readFileSync(join(here, '..', 'index.html'), 'utf8');

let html = src;
const n = (html.match(/https:\/\/www\.gstatic\.com\/firebasejs\/[^"']+/g) || []).length;
if (n !== 3) throw new Error(`Firebase import 가 3개가 아니라 ${n}개다 — build.mjs 를 index.html 에 맞춰 고칠 것`);
html = html.replace(/https:\/\/www\.gstatic\.com\/firebasejs\/[^"']+/g, './preview/fake-firebase.js');
// 상대 경로(firebase-config.js · 아이콘 · manifest)가 저장소 맨 위를 가리키게
html = html.replace(/<head>/, '<head>\n<base href="../">');
// 미리보기가 서비스워커를 깔아 localhost 를 캐시하지 않게
html = html.replace(/if \('serviceWorker' in navigator\)/, "if (false && 'serviceWorker' in navigator)");
html = html.replace(/<title>([^<]*)<\/title>/, `<title>$1 · 미리보기 ${ref ? '(전)' : '(후)'}</title>`);

const out = join(here, ref ? 'before.html' : 'after.html');
writeFileSync(out, html);
console.log('wrote', out);
