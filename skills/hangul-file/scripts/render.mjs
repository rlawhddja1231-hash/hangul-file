// render.mjs — rhwp(@rhwp/core, MIT, https://github.com/edwardkim/rhwp) 로 한글 파일을
// 실제 쪽 모양 그대로 그린다. 크로미움(playwright) 안에서 WASM 을 돌린다.
//
//   node render.mjs png 문서.hwp 출력폴더 [최대쪽수]
//   node render.mjs pdf 문서.hwp 출력.pdf
//   node render.mjs svg 문서.hwp 출력폴더 [최대쪽수]
//
// 환경변수 RHWP_CORE 에 @rhwp/core 폴더 경로를 준다(없으면 이 파일 옆 node_modules 에서 찾는다).
import { createRequire } from 'module';
import fs from 'fs';
import path from 'path';

const require = createRequire(import.meta.url);
const { chromium } = require('playwright');

const [, , mode, src, out, maxArg] = process.argv;
if (!['png', 'pdf', 'svg'].includes(mode) || !src || !out) {
  console.error('쓰는 법: node render.mjs png|pdf|svg 문서.hwp 출력 [최대쪽수]');
  process.exit(2);
}
const here = path.dirname(new URL(import.meta.url).pathname);
const core = process.env.RHWP_CORE || path.join(here, 'node_modules/@rhwp/core');
if (!fs.existsSync(path.join(core, 'rhwp_bg.wasm'))) {
  console.error(JSON.stringify({ error: 'rhwp-core-missing', core }));
  process.exit(3);
}
const maxPages = parseInt(maxArg || '9999', 10);
const data = fs.readFileSync(src).toString('base64');

const browser = await chromium.launch();
try {
  const page = await browser.newPage({ viewport: { width: 1000, height: 1400 } });
  await page.route('https://rhwp.local/**', async (route) => {
    const p = new URL(route.request().url()).pathname;
    if (p === '/') {
      return route.fulfill({
        contentType: 'text/html',
        body: '<!doctype html><meta charset="utf-8"><style>html,body{margin:0;background:#fff}' +
          '.pg{break-after:page;page-break-after:always}.pg:last-child{break-after:auto}' +
          '.pg svg{display:block}</style><body><div id=v></div></body>',
      });
    }
    const f = path.join(core, path.basename(p));
    return route.fulfill({ contentType: f.endsWith('.wasm') ? 'application/wasm' : 'text/javascript', body: fs.readFileSync(f) });
  });
  await page.goto('https://rhwp.local/');
  const info = await page.evaluate(async (b64) => {
    const mod = await import('./rhwp.js'); // 위 page.route 가 이 파일 옆의 @rhwp/core 에서 내어 줌
    const c = document.createElement('canvas').getContext('2d');
    globalThis.measureTextWidth = (font, text) => { c.font = font; return c.measureText(text).width; };
    await mod.default({ module_or_path: './rhwp_bg.wasm' });
    const bytes = Uint8Array.from(atob(b64), (x) => x.charCodeAt(0));
    globalThis.__doc = new mod.HwpDocument(bytes);
    return { pages: globalThis.__doc.pageCount() };
  }, data);
  const n = Math.min(info.pages, maxPages);

  if (mode === 'pdf') {
    const size = await page.evaluate((n) => {
      const v = document.getElementById('v');
      let w = 0, h = 0;
      for (let i = 0; i < n; i++) {
        const d = document.createElement('div');
        d.className = 'pg';
        d.innerHTML = globalThis.__doc.renderPageSvg(i);
        v.appendChild(d);
        const s = d.querySelector('svg');
        const bb = s.getBoundingClientRect();
        w = Math.max(w, bb.width); h = Math.max(h, bb.height);
      }
      return { w, h };
    }, n);
    await page.pdf({ path: out, width: `${size.w}px`, height: `${size.h}px`, printBackground: true,
      margin: { top: 0, right: 0, bottom: 0, left: 0 } });
    console.log(JSON.stringify({ pages: info.pages, written: n, pdf: out }));
  } else {
    fs.mkdirSync(out, { recursive: true });
    const made = [];
    for (let i = 0; i < n; i++) {
      const svg = await page.evaluate((i) => {
        const s = globalThis.__doc.renderPageSvg(i);
        document.getElementById('v').innerHTML = s;
        return s;
      }, i);
      const name = `쪽${String(i + 1).padStart(2, '0')}`;
      if (mode === 'svg') {
        const f = path.join(out, name + '.svg');
        fs.writeFileSync(f, svg);
        made.push(f);
      } else {
        const f = path.join(out, name + '.png');
        await (await page.$('#v svg')).screenshot({ path: f });
        made.push(f);
      }
    }
    console.log(JSON.stringify({ pages: info.pages, made }));
  }
} finally {
  await browser.close();
}
