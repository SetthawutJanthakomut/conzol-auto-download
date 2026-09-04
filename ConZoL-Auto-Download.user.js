// ==UserScript==
// @name         GULF ConZoL - Auto Download + Rename + Sort
// @namespace    gmtp.conzol
// @version      4.6
// @description  Download PDFs and native attachments from GULF ConZoL EDMS automatically - names each file and sorts it into the folder ConZoL assigns.
// @match        https://edms.gulf.co.th/dms/drawing.asp*
// @match        http://edms.gulf.co.th/dms/drawing.asp*
// @updateURL    https://raw.githubusercontent.com/SetthawutJanthakomut/conzol-auto-download/main/ConZoL-Auto-Download.user.js
// @downloadURL  https://raw.githubusercontent.com/SetthawutJanthakomut/conzol-auto-download/main/ConZoL-Auto-Download.user.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @run-at       document-idle
// @grant        none
// ==/UserScript==


(function () {
  'use strict';

  // Guard against a double install (userscript + extension, or two scripts)
  // If a panel already exists the later copy stops here - otherwise ids collide and buttons stop responding
  if (document.getElementById('edmsdl')) return;

  const VERSION = '4.6';   // kept in sync with @version at build time
  const UPDATE_URL = 'https://raw.githubusercontent.com/SetthawutJanthakomut/conzol-auto-download/main/ConZoL-Auto-Download.user.js';   // filled in per language at build time

  // ---------------- Settings ----------------
  const CFG = {
    delayMs: 500,          // pause between files
    searchDelayMs: 400,    // pause between search pages
    retry: 1,
    maxNameLen: 180,
    supersededDir: '_Superseded',
    unsortedDir: '_Unsorted',
    // Folders never touched - not scanned, not re-sorted (cancelled MDR documents live here)
    ignoreDirs: ['_Deleted', '_Archive', '_Cancelled']
  };

  // Optional labels for area codes - add as you learn them (unlisted codes use the bare code)
  const AREA_LABELS = {
    '0500': '0500 Seawater Intake-Outfall',
    '1400': '1400 Berth 1'
  };
  // ------------------------------------------------


  const XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  let xlsxLoading = null;
  function ensureXLSX() {
    if (typeof window.XLSX !== 'undefined') return Promise.resolve(window.XLSX);
    if (xlsxLoading) return xlsxLoading;
    xlsxLoading = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = XLSX_URL;
      s.onload = () => (typeof window.XLSX !== 'undefined') ? res(window.XLSX) : rej(new Error('Spreadsheet library failed to load'));
      s.onerror = () => rej(new Error('Cannot reach cdnjs.cloudflare.com'));
      document.head.appendChild(s);
      setTimeout(() => rej(new Error('Timed out loading spreadsheet library')), 20000);
    });
    return xlsxLoading;
  }

  const $ = (t, p) => Object.assign(document.createElement(t), p || {});
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const norm = (s) => String(s || '').toUpperCase().replace(/\s+/g, '').trim();
  const FSA = typeof window.showDirectoryPicker === 'function';

  let stopFlag = false, running = false;
  let mdrList = [], mdrExcluded = 0, lastReport = [];
  let rootDir = null;          // FileSystemDirectoryHandle
  let existing = new Map();    // DOCNO -> [{name, rev, rank, parent}]

  // ============ File and folder names ============
  const FN_RE  = /^(GMTP-[A-Z0-9\-]+?)-([TR])(\d+)_/i;
  const DOC_RE = /\bGMTP-[A-Z0-9]+(?:-[A-Z0-9]+)*-\d{3,4}\b/i;
  // No extension whitelist - ConZoL can hand back any attachment type (zip, dwg, dgn, msg, ...)
  // Accept any file whose name starts with a document number, except part-downloads
  const SKIP_EXT = /\.(tmp|crdownload|part|partial|download|!ut)$/i;

  function safeName(r, ext) {
    let base = r.doc + (r.rev ? '-' + r.rev : '') + (r.title ? '_' + r.title : '');
    base = base.replace(/[\\/:*?"<>|\r\n\t]/g, '-').replace(/\s+/g, ' ').replace(/[. ]+$/g, '').trim();
    if (base.length > CFG.maxNameLen) base = base.slice(0, CFG.maxNameLen).trim();
    return base + '.' + String(ext || 'pdf').toLowerCase();
  }

  const sanitizeFolder = (s) => String(s).replace(/[\\/:*?"<>|\r\n\t]/g, '-')
    .replace(/\s+/g, ' ').replace(/[. ]+$/g, '').trim().slice(0, 100) || '_Unsorted';

  // ============ Document group - taken from ConZoL itself ============
  // Group header rows look like:  "MA-CAL:MARINE Calculation"
  // Cached in localStorage so already-downloaded files can be sorted later
  const GC_KEY = 'edms_groupcache_v1';
  let groupCache = (() => {
    try { return JSON.parse(localStorage.getItem(GC_KEY) || '{}'); } catch (e) { return {}; }
  })();
  const saveGroupCache = () => { try { localStorage.setItem(GC_KEY, JSON.stringify(groupCache)); } catch (e) {} };
  function rememberGroup(doc, code, name) {
    if (!doc || !code) return;
    groupCache[doc.toUpperCase()] = { c: code, n: name };
  }

  // Derive the area code from the document number using ConZoL's group code
  //   doc GMTP-1400-MA-DWG-401  group MA-DWG  -> area 1400
  //   doc GMTP-CAZ-COJ-MS-007   group COJ-MS  -> area CAZ
  //   doc GMTP-MA-RPT-001       group MA-RPT  -> no area
  function areaOf(doc, groupCode) {
    if (!groupCode) return '';
    const toks = String(doc).toUpperCase().split('-');
    const i = toks.indexOf(String(groupCode).toUpperCase().split('-')[0]);
    return i > 1 ? toks.slice(1, i).join('-') : '';
  }

  function targetPath(doc, groupCode, groupName) {
    const g = groupCache[String(doc).toUpperCase()];
    const code = groupCode || (g && g.c) || '';
    const name = groupName || (g && g.n) || '';
    if (!code) return [CFG.unsortedDir];
    const parts = [sanitizeFolder(name ? code + ' ' + name : code)];
    const cb = el('edl-area');
    if (!cb || cb.checked) {
      const area = areaOf(doc, code);
      if (area) parts.push(sanitizeFolder(AREA_LABELS[area] || area));
    }
    return parts;
  }

  const revRank = (series, num) => (String(series).toUpperCase() === 'T' ? 1000 : 0) + parseInt(num, 10);

  // ============ IndexedDB - remember the chosen folder ============
  function idb() {
    return new Promise((res, rej) => {
      const r = indexedDB.open('edms_autodl', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('kv');
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
  }
  async function idbPut(k, v) {
    const db = await idb();
    return new Promise((res, rej) => {
      const t = db.transaction('kv', 'readwrite');
      t.objectStore('kv').put(v, k);
      t.oncomplete = () => res(); t.onerror = () => rej(t.error);
    });
  }
  async function idbGet(k) {
    const db = await idb();
    return new Promise((res, rej) => {
      const t = db.transaction('kv', 'readonly');
      const q = t.objectStore('kv').get(k);
      q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error);
    });
  }

  // ============ File System Access ============
  async function ensureDir(parts) {
    let d = rootDir;
    for (const p of parts) d = await d.getDirectoryHandle(p, { create: true });
    return d;
  }

  async function scanFolder(dir, out, path) {
    for await (const entry of dir.values()) {
      if (entry.kind === 'directory') {
        if (CFG.ignoreDirs.some((d) => d.toLowerCase() === entry.name.toLowerCase())) continue;
        await scanFolder(entry, out, path.concat(entry.name));
        continue;
      }
      if (SKIP_EXT.test(entry.name)) continue;
      if (entry.name.lastIndexOf('.') <= 0) continue;   // must have an extension
      const m = FN_RE.exec(entry.name);
      if (!m) continue;
      const doc = m[1].toUpperCase();
      if (!out.has(doc)) out.set(doc, []);
      out.get(doc).push({
        name: entry.name,
        rev: (m[2] + m[3]).toUpperCase(),
        rank: revRank(m[2], m[3]),
        ext: (entry.name.split('.').pop() || 'pdf').toLowerCase(),
        parent: dir,
        path: path,
        handle: entry
      });
    }
    return out;
  }

  async function refreshExisting() {
    existing = new Map();
    if (!rootDir) return existing;
    await scanFolder(rootDir, existing, []);
    return existing;
  }

  async function writeInto(dir, name, blob) {
    const fh = await dir.getFileHandle(name, { create: true });
    const w = await fh.createWritable();
    await w.write(blob);
    await w.close();
  }

  // Superseded revisions go to a _Superseded folder inside the document's own folder
  //   MA-DWG MARINE Drawing\1400 Berth 1\_Superseded\
  async function moveToSuperseded(item, destParts) {
    const sup = await ensureDir((destParts || []).concat(CFG.supersededDir));
    try {
      if (item.handle.move) { await item.handle.move(sup, item.name); return true; }
    } catch (e) { /* fall through to copy+delete */ }
    try {
      const file = await item.handle.getFile();
      await writeInto(sup, item.name, file);
      await item.parent.removeEntry(item.name);
      return true;
    } catch (e) { return false; }
  }

  // ============ Download ============
  async function fetchDoc(url) {
    const res = await fetch(url, { credentials: 'same-origin' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const ct = (res.headers.get('content-type') || '').toLowerCase();
    const blob = await res.blob();
    if (ct.includes('text/html') || blob.size < 1024) throw new Error('No file returned - ConZoL session may have expired');
    return blob;
  }

  function browserDownload(name, blob) {
    const url = URL.createObjectURL(blob);
    const a = $('a', { href: url, download: name });
    a.style.display = 'none';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  // ============ Search via fetch ============
  // Read the result table - keep group headers and rows in ConZoL's own order
  function parseRows(root) {
    const rows = [];
    let gCode = '', gName = '';
    for (const tr of root.querySelectorAll('tr')) {
      const cells = tr.cells;
      if (!cells) continue;

      // Group header row: single wide colspan cell, text "CODE:Group name"
      if (cells.length === 1 && (cells[0].colSpan || 1) > 1) {
        const t = String(cells[0].innerText || '').replace(/\s+/g, ' ').trim();
        const m = /^([A-Z0-9][A-Z0-9\-]{1,20})\s*:\s*(.{1,80})$/i.exec(t);
        if (m) { gCode = m[1].toUpperCase(); gName = m[2].trim(); }
        continue;
      }

      if (cells.length < 11) continue;
      const cb = tr.querySelector('input[type=checkbox][name^="p"]');
      if (!cb || !/^p\d+$/.test(cb.name)) continue;
      const doc = String(cells[1].innerText || '').replace(/\s+/g, ' ').trim();
      if (!doc) continue;
      rememberGroup(doc, gCode, gName);
      // FILE+ column: the native attachment (normally a .ZIP)
      const fa = cells[11] ? cells[11].querySelector('a[href*="type=c"]') : null;
      const fhref = fa ? fa.getAttribute('href') : '';
      const fext = fhref ? String((/[?&]ext=([^&]+)/i.exec(fhref) || [])[1] || 'zip').toLowerCase() : '';
      rows.push({
        fileid: cb.name.slice(1), doc,
        rev: String(cells[2].innerText || '').replace(/\s+/g, ' ').trim(),
        title: String(cells[3].innerText || '').replace(/\s+/g, ' ').trim(),
        groupCode: gCode, groupName: gName,
        fileHref: fhref, fileExt: fext, cb
      });
    }
    if (rows.length) saveGroupCache();
    return rows;
  }

  function formBase() {
    const f = document.myform, o = {};
    if (f) [...f.elements].forEach((e) => { if (e.name && e.type !== 'submit') o[e.name] = e.value; });
    o.postback = '1'; o.searchmode = '0'; o.orderby = 'S';
    ['search', 'searchid', 'searchrev', 'searchname', 'searchspec', 'searchremark',
     'searchcomment', 'searchkeyword', 'equipmenttag', 'wbsid', 'ownerref',
     'doctype', 'searchgid', 'searchscope'].forEach((k) => { o[k] = ''; });
    return o;
  }

  async function searchPage(extra) {
    const body = new URLSearchParams(Object.assign(formBase(), extra || {}));
    const res = await fetch('drawing.asp', {
      method: 'POST', credentials: 'same-origin',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body
    });
    if (!res.ok) throw new Error('Search failed - HTTP ' + res.status);
    const d = new DOMParser().parseFromString(await res.text(), 'text/html');
    return { rows: parseRows(d), pages: [...d.querySelectorAll('select[name=page] option')].map((o) => o.value) };
  }

  function disciplineCodes() {
    const sel = document.myform && document.myform.worktype;
    if (!sel) return new Set();
    return new Set([...sel.options].map((o) => o.value).filter(Boolean).map((v) => v.toUpperCase()));
  }
  function disciplineOf(docNo, codes) {
    for (const tok of String(docNo).toUpperCase().split('-')) if (codes.has(tok)) return tok;
    return null;
  }

  // ============ Read the MDR file ============
  function readWorkbook(wb, sheetName) {
    const keep = new Map(), dropped = new Set();
    const sheets = sheetName === '*' ? wb.SheetNames : [sheetName];
    for (const sn of sheets) {
      const ws = wb.Sheets[sn];
      if (!ws || !ws['!ref']) continue;
      const rng = XLSX.utils.decode_range(ws['!ref']);
      for (let R = rng.s.r; R <= rng.e.r; R++) {
        let docVal = null, titleVal = '', isDeleted = false;
        for (let C = rng.s.c; C <= rng.e.c; C++) {
          const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
          if (!cell) continue;
          const txt = String(cell.w != null ? cell.w : (cell.v != null ? cell.v : '')).trim();
          if (!txt) continue;
          const low = txt.toLowerCase();
          if (low === 'delete' || low === 'deleted') isDeleted = true;
          if (cell.s && cell.s.font && cell.s.font.strike) isDeleted = true;
          if (!docVal) {
            const m = txt.match(DOC_RE);
            if (m && norm(m[0]) === norm(txt)) {
              docVal = m[0].toUpperCase();
              const nb = ws[XLSX.utils.encode_cell({ r: R, c: C + 1 })];
              const nt = nb ? String(nb.w != null ? nb.w : (nb.v != null ? nb.v : '')).trim() : '';
              if (nt.length > 3 && !DOC_RE.test(nt)) titleVal = nt;
            }
          }
        }
        if (!docVal) continue;
        if (isDeleted) { dropped.add(norm(docVal)); continue; }
        if (!keep.has(norm(docVal))) keep.set(norm(docVal), { doc: docVal, title: titleVal, sheet: sn });
      }
    }
    for (const k of dropped) keep.delete(k);
    mdrExcluded = dropped.size;
    return [...keep.values()];
  }

  function readTextList(text) {
    const out = new Map();
    for (const line of String(text).split(/[\r\n,;\t]+/)) {
      const m = line.match(DOC_RE);
      if (m) { const d = m[0].toUpperCase(); if (!out.has(norm(d))) out.set(norm(d), { doc: d, title: '', sheet: '' }); }
    }
    return [...out.values()];
  }

  function applyManualExclusions(list) {
    const ta = el('edl-exclude');
    const txt = ta ? ta.value : '';
    if (!txt || !txt.trim()) return list;
    const ex = new Set(readTextList(txt).map((m) => norm(m.doc)));
    if (!ex.size) return list;
    return list.filter((m) => !ex.has(norm(m.doc)));
  }

  // ============ UI ============
  const box = $('div');
  box.id = 'edmsdl';
  box.innerHTML = `
    <style>
      #edmsdl{position:fixed;top:8px;right:8px;z-index:999999;width:378px;
        font:11px Tahoma,Verdana,sans-serif;background:#fff;color:#222;
        border:1px solid #b9a;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.28)}
      #edmsdl .hd{background:#7a6a52;color:#fff;padding:6px 9px;border-radius:7px 7px 0 0;
        font-weight:bold;cursor:move;display:flex;justify-content:space-between;align-items:center}
      #edmsdl .bd{padding:8px 9px}
      #edmsdl fieldset{border:1px solid #ddd;border-radius:5px;margin:0 0 7px;padding:6px 8px}
      #edmsdl legend{font-weight:bold;color:#6b5b43;padding:0 4px}
      #edmsdl label{display:block;margin:3px 0}
      #edmsdl button{font:11px Tahoma;padding:4px 9px;margin:4px 4px 0 0;cursor:pointer;
        border:1px solid #999;border-radius:4px;background:#f3f0ea}
      #edmsdl button.go{background:#3d7a3d;color:#fff;border-color:#2c5c2c}
      #edmsdl button.go2{background:#2f5d8a;color:#fff;border-color:#254a6e}
      #edmsdl button.pick{background:#8a5a2f;color:#fff;border-color:#6e4522}
      #edmsdl button.stop{background:#a33;color:#fff;border-color:#822}
      #edmsdl select,#edmsdl input[type=file],#edmsdl textarea{font:11px Tahoma;max-width:100%}
      #edmsdl .log{max-height:200px;overflow:auto;margin-top:6px;border-top:1px solid #ddd;
        padding-top:5px;font:10px Consolas,monospace;white-space:pre-wrap}
      #edmsdl .ok{color:#2c7a2c}#edmsdl .er{color:#b02020}#edmsdl .sk{color:#888}#edmsdl .wn{color:#a06000}
      #edmsdl .st{font-weight:bold;margin-top:5px}
      #edmsdl .x{cursor:pointer;font-weight:normal}
      #edmsdl .ver{font:9px Tahoma;font-weight:normal;opacity:.65;margin-left:5px;letter-spacing:.3px}
      #edmsdl .ver.new{opacity:1;background:#d9822b;color:#fff;padding:1px 6px;border-radius:9px;cursor:pointer}
    </style>
    <div class="hd"><span>⬇ ConZoL Auto Download <span class="ver" id="edl-ver">v${VERSION}</span></span><span class="x" id="edl-min">–</span></div>
    <div class="bd" id="edl-body">

      <fieldset><legend>Destination folder</legend>
        <div id="edl-dirinfo" class="wn">No folder chosen yet</div>
        <button class="pick" id="edl-pick">Choose folder…</button>
        <button id="edl-rescan">Rescan</button>
        <button id="edl-tidy">Re-sort folders</button>
      </fieldset>

      <fieldset><legend>1) From the current search results</legend>
        <label><input type="checkbox" id="edl-checked"> Only ticked rows</label>
        <button class="go" id="edl-run">Download this page</button>
      </fieldset>

      <fieldset><legend>2) From an MDR file (automatic)</legend>
        <input type="file" id="edl-file" accept=".xlsx,.xlsm,.csv,.txt">
        <div style="margin-top:4px">Sheet: <select id="edl-sheet"><option value="*">— choose a file first —</option></select></div>
        <div id="edl-mdrinfo" class="sk" style="margin-top:3px">No list loaded</div>
        <div style="margin-top:4px">Exclude these documents:</div>
        <textarea id="edl-exclude" rows="2" style="width:100%;font:10px Consolas,monospace" placeholder="e.g. GMTP-MA-RPT-012"></textarea>
        <button class="go2" id="edl-runmdr">Search ConZoL + download all</button>
      </fieldset>

      <fieldset style="margin-top:2px"><legend>What to download</legend>
        <label><input type="checkbox" id="edl-getpdf" checked> PDF (PDF+ column)</label>
        <label><input type="checkbox" id="edl-getfile"> Native attachment (FILE+ column · .zip)</label>
      </fieldset>
      <label><input type="checkbox" id="edl-skip" checked> Skip files already in the folder</label>
      <label><input type="checkbox" id="edl-sup" checked> Move superseded revisions to _Superseded</label>
      <label><input type="checkbox" id="edl-area" checked> Sub-folder per area code (1400 / 0500 / PCC …)</label>
      <label><input type="checkbox" id="edl-inactive"> Include non-active documents in search</label>
      <div>
        <button class="stop" id="edl-stop">Stop</button>
        <button id="edl-csv">Save CSV report</button>
      </div>
      <div class="st" id="edl-status">Ready</div>
      <div class="log" id="edl-log"></div>
    </div>`;
  document.body.appendChild(box);

  const el = (id) => box.querySelector('#' + id);
  const logEl = el('edl-log'), statusEl = el('edl-status');
  const log = (m, c) => { logEl.appendChild($('div', { className: c || '', textContent: m })); logEl.scrollTop = logEl.scrollHeight; };

  el('edl-min').onclick = () => { const b = el('edl-body'); b.style.display = b.style.display === 'none' ? '' : 'none'; };

  (() => {
    const hd = box.querySelector('.hd');
    let sx, sy, ox, oy, drag = false;
    hd.onmousedown = (e) => {
      if (e.target.id === 'edl-min' || e.target.id === 'edl-ver') return;
      drag = true; sx = e.clientX; sy = e.clientY;
      const r = box.getBoundingClientRect(); ox = r.left; oy = r.top; e.preventDefault();
    };
    document.addEventListener('mousemove', (e) => {
      if (!drag) return;
      box.style.left = (ox + e.clientX - sx) + 'px';
      box.style.top = (oy + e.clientY - sy) + 'px';
      box.style.right = 'auto';
    });
    document.addEventListener('mouseup', () => { drag = false; });
  })();

  // ---- destination folder ----
  function showDirInfo(extra) {
    const d = el('edl-dirinfo');
    if (!rootDir) {
      d.className = 'wn';
      d.textContent = FSA ? 'No folder chosen - files will go to Downloads'
                          : 'Browser not supported - files will go to Downloads';
      return;
    }
    d.className = 'ok';
    let n = 0; existing.forEach((v) => { n += v.length; });
    d.textContent = `📁 ${rootDir.name} · ${n} file(s) already here` + (extra ? ' · ' + extra : '');
  }

  async function useDir(handle, quiet) {
    rootDir = handle;
    await refreshExisting();
    showDirInfo();
    if (!quiet) log(`· Using folder "${handle.name}" - found ${[...existing.values()].reduce((a, b) => a + b.length, 0)} existing file(s)`, 'ok');
  }

  el('edl-pick').onclick = async () => {
    if (!FSA) { log('This browser cannot pick a folder - use a recent Chrome or Edge', 'er'); return; }
    try {
      const h = await window.showDirectoryPicker({ mode: 'readwrite', id: 'edmsdl' });
      await idbPut('dir', h);
      await useDir(h);
    } catch (e) { if (e.name !== 'AbortError') log('Could not open that folder: ' + e.message, 'er'); }
  };

  el('edl-rescan').onclick = async () => {
    if (!rootDir) return;
    await refreshExisting(); showDirInfo();
    log('· Folder rescanned', 'sk');
  };

  // ---- re-sort every file into its current target folder ----
  async function moveTo(item, dirParts) {
    if ((item.path || []).join('\\') === dirParts.join('\\')) return 'same';
    const dir = await ensureDir(dirParts);
    try {
      if (item.handle.move) { await item.handle.move(dir, item.name); return 'moved'; }
    } catch (e) { /* fall through to copy+delete */ }
    const file = await item.handle.getFile();
    await writeInto(dir, item.name, file);
    await item.parent.removeEntry(item.name);
    return 'moved';
  }

  async function pruneEmpty(dir, path) {
    let empty = true;
    const subs = [];
    for await (const e of dir.values()) {
      if (e.kind !== 'directory') { empty = false; continue; }
      if (CFG.ignoreDirs.some((d) => d.toLowerCase() === e.name.toLowerCase())) { empty = false; continue; }
      subs.push(e);
    }
    for (const s of subs) {
      const gone = await pruneEmpty(s, path.concat(s.name));
      if (!gone) empty = false;
    }
    if (empty && path.length) {
      try {
        let parent = rootDir;
        for (const p of path.slice(0, -1)) parent = await parent.getDirectoryHandle(p);
        await parent.removeEntry(path[path.length - 1]);
        return true;
      } catch (e) { return false; }
    }
    return false;
  }

  el('edl-tidy').onclick = async () => {
    if (running) return;
    if (!rootDir) { log('· Choose a destination folder first', 'er'); return; }
    if (!(await ensurePermission())) { log('· Write permission denied', 'er'); return; }
    running = true; logEl.innerHTML = ''; lastReport = [];
    log('Scanning folder…');
    await refreshExisting();

    // Resolve unknown document groups - ask ConZoL one by one
    const unknown = [...existing.keys()].filter((d) => !groupCache[d]);
    if (unknown.length) {
      log(`Asking ConZoL for the group of ${unknown.length} document(s)…`);
      for (const d of unknown) {
        if (stopFlag) break;
        try {
          const { rows } = await searchPage({ search: d, worktype: '', dstatus: '' });
          const hit = rows.find((r) => norm(r.doc) === norm(d));
          if (hit && hit.groupCode) log(`   ${d} → ${hit.groupCode} ${hit.groupName}`, 'sk');
          else log(`   ${d} → not found in ConZoL`, 'wn');
        } catch (e) { log(`   ${d} → ${e.message}`, 'er'); }
        await sleep(CFG.searchDelayMs);
      }
      saveGroupCache();
    }

    let moved = 0, sup = 0, same = 0, fail = 0;
    const doSup = el('edl-sup').checked;

    for (const [doc, list] of [...existing.entries()]) {
      const parts = targetPath(doc, '', '');

      // Group by extension first - the PDF and the attachment are two file types of
      // the same document, not older revisions of each other. Without this split one
      const byExt = new Map();
      for (const it of list) {
        const e = it.ext || 'pdf';
        if (!byExt.has(e)) byExt.set(e, []);
        byExt.get(e).push(it);
      }

      for (const group of byExt.values()) {
      const sorted = [...group].sort((a, b) => b.rank - a.rank);
      for (let i = 0; i < sorted.length; i++) {
        const it = sorted[i];
        const dest = (i === 0 || !doSup) ? parts : parts.concat(CFG.supersededDir);
        try {
          const r = await moveTo(it, dest);
          if (r === 'moved') {
            if (i === 0 || !doSup) { moved++; log(`✓ ${it.name} → ${dest.join('\\')}`, 'ok'); }
            else { sup++; log(`↳ ${it.name} → _Superseded`, 'wn'); }
          } else same++;
        } catch (e) {
          fail++;
          log(`✗ ${it.name} → ${e.name === 'NotAllowedError' ? 'permission denied' : e.message} (file may be open)`, 'er');
        }
      }
      }
    }

    try { await pruneEmpty(rootDir, []); } catch (e) {}
    await refreshExisting(); showDirInfo();
    statusEl.textContent = `Sorted: moved ${moved} · already correct ${same} · superseded ${sup} · failed ${fail}`;
    log('— Sorting finished —');
    running = false;
  };

  async function ensurePermission() {
    if (!rootDir) return false;
    const opt = { mode: 'readwrite' };
    if ((await rootDir.queryPermission(opt)) === 'granted') return true;
    return (await rootDir.requestPermission(opt)) === 'granted';
  }

  // Restore the previously chosen folder (permission confirmed again on first use)
  (async () => {
    if (!FSA) { showDirInfo(); return; }
    try {
      const h = await idbGet('dir');
      if (!h) { showDirInfo(); return; }
      if ((await h.queryPermission({ mode: 'readwrite' })) === 'granted') {
        await useDir(h, true);
        log(`· Remembered folder: ${h.name}`, 'ok');
      } else {
        rootDir = h;
        el('edl-dirinfo').className = 'wn';
        el('edl-dirinfo').textContent = `📁 ${h.name} — press a download button and allow access again`;
      }
    } catch (e) { showDirInfo(); }
  })();

  // Quietly check GitHub for a newer version - stays silent if unreachable
  (async () => {
    if (!/^https?:/.test(UPDATE_URL)) return;
    try {
      const r = await fetch(UPDATE_URL + '?t=' + Date.now(), { cache: 'no-store' });
      if (!r.ok) return;
      const m = /@version\s+(\S+)/.exec((await r.text()).slice(0, 2000));
      if (!m || m[1] === VERSION) return;
      const b = el('edl-ver');
      if (!b) return;
      b.textContent = 'v' + VERSION + ' → ' + m[1];
      b.className = 'ver new';
      b.title = 'A newer version is available - click to install';
      b.onclick = (e) => { e.stopPropagation(); window.open(UPDATE_URL, '_blank'); };
      log('Version ' + m[1] + ' is available (running ' + VERSION + ') - click the orange tag in the title bar to install', 'wn');
    } catch (e) { /* ignore */ }
  })();

  el('edl-stop').onclick = () => { stopFlag = true; statusEl.textContent = 'Stopping…'; };

  // ---- MDR file picker ----
  let workbook = null;
  function showMdrInfo(fname) {
    const i = el('edl-mdrinfo');
    i.className = 'ok';
    i.textContent = `Using ${mdrList.length} document number(s)` +
      (mdrExcluded ? ` · ${mdrExcluded} cancelled in MDR excluded` : '') + (fname ? ` · ${fname}` : '');
  }
  const refreshMdr = () => {
    if (!workbook) return;
    mdrList = applyManualExclusions(readWorkbook(workbook, el('edl-sheet').value));
    showMdrInfo('');
  };
  el('edl-sheet').onchange = refreshMdr;
  el('edl-exclude').onchange = refreshMdr;

  el('edl-file').onchange = async (ev) => {
    const f = ev.target.files && ev.target.files[0];
    if (!f) return;
    try {
      mdrExcluded = 0;
      if (/\.(csv|txt)$/i.test(f.name)) {
        workbook = null;
        mdrList = readTextList(await f.text());
        el('edl-sheet').innerHTML = '<option value="*">(text file)</option>';
      } else {
        await ensureXLSX();
        workbook = XLSX.read(await f.arrayBuffer(), { type: 'array', cellStyles: true });
        el('edl-sheet').innerHTML = '<option value="*">— all sheets —</option>' +
          workbook.SheetNames.map((n) => `<option value="${n.replace(/"/g, '&quot;')}">${n}</option>`).join('');
        mdrList = readWorkbook(workbook, '*');
      }
      mdrList = applyManualExclusions(mdrList);
      showMdrInfo(f.name);
      log(`· Read ${f.name}: using ${mdrList.length} document(s)` + (mdrExcluded ? ` (${mdrExcluded} cancelled excluded)` : ''), 'ok');
    } catch (e) {
      el('edl-mdrinfo').className = 'er';
      el('edl-mdrinfo').textContent = 'Cannot read file: ' + e.message;
    }
  };

  // ---- CSV report ----
  el('edl-csv').onclick = async () => {
    if (!lastReport.length) { log('· No report yet - run a download first', 'sk'); return; }
    const esc = (s) => '"' + String(s == null ? '' : s).replace(/"/g, '""') + '"';
    const head = ['Document No.', 'Rev', 'Title', 'Result', 'File', 'Folder', 'Sheet'];
    const csv = '﻿' + [head, ...lastReport.map((r) => [r.doc, r.rev, r.title, r.result, r.file, r.folder || '', r.sheet || ''])]
      .map((r) => r.map(esc).join(',')).join('\r\n');
    const d = new Date();
    const stamp = d.getFullYear() + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0');
    const name = `ConZoL_download_report_${stamp}.csv`;
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    if (rootDir && await ensurePermission()) { await writeInto(rootDir, name, blob); log('· Saved ' + name + ' into the folder', 'ok'); }
    else browserDownload(name, blob);
  };

  // ============ Download runner ============
  async function runDownload(items) {
    let ok = 0, skipped = 0, fail = 0, sup = 0, i = 0;
    const skip = el('edl-skip').checked;
    const doSup = el('edl-sup').checked;
    const wantPdf = el('edl-getpdf').checked;
    const wantFile = el('edl-getfile').checked;
    const useFS = !!rootDir;

    if (!wantPdf && !wantFile) {
      log('· Nothing selected - tick PDF and/or Native attachment first', 'er');
      return { ok, skipped, fail, sup };
    }

    for (const r of items) {
      if (stopFlag) { log('■ Stopped by user', 'er'); break; }
      i++;
      const doc = r.doc.toUpperCase();
      const rm = /^([TR])(\d+)$/i.exec(r.rev);
      const rank = rm ? revRank(rm[1], rm[2]) : -1;
      if (!existing.has(doc)) existing.set(doc, []);
      const have = existing.get(doc);
      const parts = targetPath(doc, r.groupCode, r.groupName);

      const kinds = [];
      if (wantPdf) kinds.push({ ext: 'pdf', url: 'getfile.asp?type=p&fileid=' + r.fileid + '&docid=0' });
      if (wantFile && r.fileHref) kinds.push({ ext: r.fileExt || 'zip', url: r.fileHref });

      if (!kinds.length) {
        skipped++;
        lastReport.push({ ...r, result: 'No file of the selected type', file: '', folder: '' });
        log(`[${i}/${items.length}] skipped ${doc}-${r.rev} - no attachment`, 'sk');
        continue;
      }

      for (const k of kinds) {
        if (stopFlag) break;
        const name = safeName(r, k.ext);

        if (skip && have.some((h) => h.rev === r.rev.toUpperCase() && (h.ext || 'pdf') === k.ext)) {
          skipped++;
          lastReport.push({ ...r, result: 'Skipped (already present)', file: name, folder: '' });
          log(`[${i}/${items.length}] skipped ${name}`, 'sk');
          continue;
        }

        statusEl.textContent = `Downloading ${i}/${items.length} …`;
        let done = false;
        for (let a = 0; a <= CFG.retry && !done; a++) {
          try {
            const blob = await fetchDoc(k.url);
            if (useFS) {
              const dir = await ensureDir(parts);
              await writeInto(dir, name, blob);
              if (doSup) {
                for (const h of have.slice()) {
                  if ((h.ext || 'pdf') === k.ext && h.rank < rank && h.name !== name) {
                    if (await moveToSuperseded(h, parts)) {
                      sup++; log(`      ↳ ${h.name} → _Superseded`, 'wn');
                      const ix = have.indexOf(h); if (ix >= 0) have.splice(ix, 1);
                    }
                  }
                }
              }
              try {
                have.push({ name, rev: r.rev.toUpperCase(), rank, ext: k.ext,
                            parent: dir, path: parts, handle: await dir.getFileHandle(name) });
              } catch (e) {}
            } else {
              browserDownload(name, blob);
            }
            ok++; done = true;
            lastReport.push({ ...r, result: 'Downloaded', file: name, folder: parts.join('\\') });
            log(`[${i}/${items.length}] ✓ ${name} (${(blob.size / 1048576).toFixed(2)} MB)`, 'ok');
          } catch (e) {
            if (a === CFG.retry) {
              fail++;
              lastReport.push({ ...r, result: 'Error: ' + e.message, file: name, folder: '' });
              log(`[${i}/${items.length}] ✗ ${name} → ${e.message}`, 'er');
            } else await sleep(1200);
          }
        }
        await sleep(CFG.delayMs);
      }
    }
    return { ok, skipped, fail, sup };
  }

  async function preflight() {
    if (rootDir) {
      if (!(await ensurePermission())) { log('Write permission denied - saving to Downloads instead', 'wn'); rootDir = null; }
      else if (!existing.size) await refreshExisting();
    } else if (FSA) {
      log('⚠ No destination folder chosen - files go to Downloads with no sorting', 'wn');
    }
  }

  // ============ Mode 1 ============
  el('edl-run').onclick = async () => {
    if (running) return;
    running = true; stopFlag = false; logEl.innerHTML = ''; lastReport = [];
    await preflight();
    let rows = parseRows(document);
    if (el('edl-checked').checked) rows = rows.filter((r) => r.cb && r.cb.checked);
    if (!rows.length) {
      statusEl.textContent = 'No documents on this page';
      log('Run a SEARCH first so the list has results', 'er'); running = false; return;
    }
    log(`Found ${rows.length} document(s) on this page`);
    const s = await runDownload(rows.map((r) => ({ ...r, sheet: '' })));
    statusEl.textContent = `Done: downloaded ${s.ok} · skipped ${s.skipped} · failed ${s.fail}` + (s.sup ? ` · superseded ${s.sup}` : '');
    showDirInfo(); log('— Finished —');
    running = false;
  };

  // ============ Mode 2 ============
  el('edl-runmdr').onclick = async () => {
    if (running) return;
    if (!mdrList.length) { log('· Choose an MDR file first', 'er'); return; }
    running = true; stopFlag = false; logEl.innerHTML = ''; lastReport = [];
    await preflight();

    const dstatus = el('edl-inactive').checked ? '' : 'A';
    const codes = disciplineCodes();
    const wanted = new Map(mdrList.map((m) => [norm(m.doc), m]));

    const byDisc = new Map(); const noDisc = [];
    for (const m of mdrList) {
      const d = disciplineOf(m.doc, codes);
      if (d) { if (!byDisc.has(d)) byDisc.set(d, []); byDisc.get(d).push(m); } else noDisc.push(m);
    }
    log(`MDR ${mdrList.length} document(s) · disciplines: ${[...byDisc.keys()].join(', ') || '-'}`);

    const index = new Map();
    for (const [disc] of byDisc) {
      if (stopFlag) break;
      statusEl.textContent = `Searching ConZoL: ${disc} …`;
      let page = 1, total = 1;
      do {
        const extra = { worktype: disc, dstatus, page: String(page) };
        if (page > 1) extra.pagechange = '1';
        const { rows, pages } = await searchPage(extra);
        total = Math.max(pages.length || 1, total);
        rows.forEach((r) => { const k = norm(r.doc); if (!index.has(k)) index.set(k, r); });
        log(`  · ${disc} page ${page}/${total} → ${rows.length}`, 'sk');
        page++; await sleep(CFG.searchDelayMs);
      } while (page <= total && !stopFlag);
    }

    const probe = [...wanted.keys()].filter((k) => !index.has(k));
    if (probe.length && !stopFlag) {
      log(`Looking up ${probe.length} document(s) individually…`, 'wn');
      for (const k of probe) {
        if (stopFlag) break;
        try {
          const { rows } = await searchPage({ search: wanted.get(k).doc, worktype: '', dstatus });
          const hit = rows.find((r) => norm(r.doc) === k);
          if (hit) index.set(k, hit);
        } catch (e) {}
        await sleep(CFG.searchDelayMs);
      }
    }

    const items = [], notFound = [];
    for (const [k, m] of wanted) {
      const hit = index.get(k);
      if (hit) items.push({ ...hit, sheet: m.sheet || '' }); else notFound.push(m);
    }
    log(`Matched ${items.length} of ${wanted.size}`, items.length ? 'ok' : 'er');
    if (notFound.length) {
      log(`Not found in ConZoL - ${notFound.length} document(s):`, 'wn');
      notFound.forEach((m) => log('   – ' + m.doc, 'wn'));
      notFound.forEach((m) => lastReport.push({ doc: m.doc, rev: '', title: m.title || '', result: 'Not found in ConZoL', file: '', folder: '', sheet: m.sheet || '' }));
    }

    if (items.length) {
      const s = await runDownload(items);
      statusEl.textContent = `Done: downloaded ${s.ok} · skipped ${s.skipped} · failed ${s.fail} · not found ${notFound.length}` + (s.sup ? ` · superseded ${s.sup}` : '');
    } else statusEl.textContent = `Nothing to download (${notFound.length} not found)`;
    showDirInfo(); log('— Finished — press "Save CSV report" to keep the results');
    running = false;
  };
})();
