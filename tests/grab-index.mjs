/* index.html 의 <script type="module"> 에서 최상위 선언 하나를 이름으로 통째로 잘라 온다
   (괄호 · 문자열 · 템플릿을 세어 끝을 찾는다). tests/sheet-contract.mjs · tests/catalog-sync.mjs 가 같이 쓴다. */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
export const indexSrc = () => fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8')
  .match(/<script type="module">([\s\S]*?)<\/script>/)[1];

export function grabFrom(src, name) {
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
