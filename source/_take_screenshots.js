#!/usr/bin/env node
/**
 * Takes light-mode screenshots of Lumber Loader pages using Edge CDP.
 * Usage: node source/_take_screenshots.js
 * Requires: Edge running at --remote-debugging-port=9222 (script starts it)
 */
const http = require('http');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

const EDGE = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
const OUT = path.join(__dirname, '..');
const BASE = 'http://localhost:8123';

const sleep = (ms) => new Promise(r => setTimeout(r, ms));

async function getTargets() {
  return new Promise((resolve, reject) => {
    const req = http.get('http://localhost:9222/json', (res) => {
      let data = '';
      res.on('data', d => (data += d));
      res.on('end', () => { try { resolve(JSON.parse(data)); } catch (e) { reject(e); } });
    });
    req.on('error', reject);
  });
}

async function openSession(wsUrl) {
  return new Promise((resolve) => {
    const ws = new WebSocket(wsUrl);
    let id = 1;
    const pending = {};
    ws.onmessage = (ev) => {
      const msg = JSON.parse(ev.data);
      if (msg.id && pending[msg.id]) {
        const { res, rej } = pending[msg.id];
        if (msg.error) rej(new Error(msg.error.message));
        else res(msg.result);
        delete pending[msg.id];
      }
    };
    const send = (method, params = {}) =>
      new Promise((res, rej) => {
        const myId = id++;
        pending[myId] = { res, rej };
        ws.send(JSON.stringify({ id: myId, method, params }));
      });
    ws.onopen = () => resolve({ send, close: () => ws.close() });
    ws.onerror = (e) => { throw e; };
  });
}

async function capture(send, filename) {
  const result = await send('Page.captureScreenshot', { format: 'png', fromSurface: true });
  const outPath = path.join(OUT, filename);
  fs.writeFileSync(outPath, Buffer.from(result.data, 'base64'));
  console.log('  saved:', filename);
}

async function nav(send, url, waitMs = 3000) {
  await send('Page.navigate', { url });
  await sleep(waitMs);
}

async function js(send, expr) {
  await send('Runtime.evaluate', { expression: `(async()=>{${expr}})()`, awaitPromise: true });
  await sleep(400);
}

async function main() {
  const edge = spawn(EDGE, [
    '--headless',
    '--remote-debugging-port=9222',
    '--window-size=960,820',
    '--no-sandbox',
    '--disable-gpu',
    'about:blank',
  ]);
  edge.stderr.on('data', () => {});

  // Wait for CDP to come up
  let targets;
  for (let i = 0; i < 30; i++) {
    try {
      targets = await getTargets();
      if (targets.length) break;
    } catch (_) {}
    await sleep(300);
  }
  if (!targets || !targets.length) throw new Error('Edge CDP never came up');

  const target = targets.find((t) => t.type === 'page') || targets[0];
  const { send, close } = await openSession(target.webSocketDebuggerUrl);

  try {
    // ── 1. Landing page ──────────────────────────────────────────────────────
    console.log('1. Landing page');
    await nav(send, `${BASE}/index.html`);
    await js(send, `document.documentElement.setAttribute('data-theme','light'); localStorage.setItem('cb-theme','light');`);
    await sleep(300);
    await capture(send, 'landing_screenshot.png');

    // ── 2. Tally mode — length palette + recommendations ────────────────────
    console.log('2. Tally mode');
    await nav(send, `${BASE}/lumber_loader.html`);
    await js(send, `
      document.documentElement.setAttribute('data-theme','light');
      localStorage.setItem('cb-theme','light');
      await new Promise(r=>setTimeout(r,800));
      // select 8,10,12,14,16,20 ft
      ['8','10','12','14','16','20'].forEach(l=>{ if(!state.selected.includes(+l)) toggleLen(l); });
      runRecommend();
      await new Promise(r=>setTimeout(r,300));
      switchTab('tallies');
      await new Promise(r=>setTimeout(r,300));
      window.scrollTo(0,0);
    `);
    await sleep(500);
    await capture(send, 'tally_screenshot.png');

    // ── 3. Visual Layout ─────────────────────────────────────────────────────
    console.log('3. Visual Layout');
    await js(send, `
      loadTally('S',0);
      await new Promise(r=>setTimeout(r,600));
      setMode('build');
      await new Promise(r=>setTimeout(r,400));
      switchTab('visual');
      await new Promise(r=>setTimeout(r,600));
      window.scrollTo(0,600);
    `);
    await sleep(600);
    await capture(send, 'html_app_screenshot.png');

    // ── 4. Manifest ──────────────────────────────────────────────────────────
    console.log('4. Manifest');
    await js(send, `
      switchTab('manifest');
      await new Promise(r=>setTimeout(r,800));
      window.scrollTo(0,600);
    `);
    await sleep(600);
    await capture(send, 'manifest_html_screenshot.png');

  } finally {
    close();
    edge.kill();
  }

  console.log('Done!');
}

main().catch((e) => { console.error(e); process.exit(1); });
