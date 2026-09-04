// ConZoL Auto Download - Chrome Extension build


(function () {
  'use strict';

  // กันกรณีติดตั้งซ้อนกัน (userscript + extension หรือสคริปต์ 2 ชุด)
  // ถ้ามีกล่องอยู่แล้ว ให้ชุดที่มาทีหลังหยุดทำงาน ไม่งั้น id จะซ้ำและปุ่มจะกดไม่ติด
  if (document.getElementById('edmsdl')) return;

  const VERSION = '4.8';   // ซิงก์อัตโนมัติจาก @version ตอน build
  const UPDATE_URL = 'https://raw.githubusercontent.com/SetthawutJanthakomut/conzol-auto-download/main/ConZoL-Auto-Download.th.user.js';   // build.py ใส่ให้ตามภาษา

  // ---------------- ตั้งค่าได้ตรงนี้ ----------------
  const CFG = {
    delayMs: 500,          // เว้นระยะระหว่างไฟล์
    searchDelayMs: 400,    // เว้นระยะระหว่างการค้นหน้า
    retry: 1,
    maxNameLen: 180,
    supersededDir: '_Superseded',
    unsortedDir: '_Unsorted',
    // โฟลเดอร์ที่ห้ามแตะ — ไม่สแกน ไม่จัดใหม่ (เอกสารที่ถูกยกเลิกใน MDR อยู่ในนี้)
    ignoreDirs: ['_Deleted', '_Archive', '_เก็บ']
  };

  // ชื่อกำกับเลขพื้นที่ — เพิ่มได้เรื่อย ๆ ตามที่รู้ (เลขที่ไม่มีในนี้จะใช้เลขเปล่า)
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
      s.onload = () => (typeof window.XLSX !== 'undefined') ? res(window.XLSX) : rej(new Error('โหลดไลบรารีไม่สำเร็จ'));
      s.onerror = () => rej(new Error('เข้าถึง cdnjs.cloudflare.com ไม่ได้'));
      document.head.appendChild(s);
      setTimeout(() => rej(new Error('หมดเวลารอไลบรารี')), 20000);
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

  // ============ ชื่อไฟล์ / โฟลเดอร์ ============
  const FN_RE  = /^(GMTP-[A-Z0-9\-]+?)-([TR])(\d+)_/i;
  const DOC_RE = /\bGMTP-[A-Z0-9]+(?:-[A-Z0-9]+)*-\d{3,4}\b/i;
  // ไม่จำกัดนามสกุล — ConZoL ส่งไฟล์แนบมาเป็นอะไรก็ได้ (zip, dwg, dgn, msg, ...)
  // รับทุกไฟล์ที่ชื่อขึ้นต้นด้วยเลขเอกสาร ยกเว้นไฟล์ที่ยังโหลดไม่จบ
  const SKIP_EXT = /\.(tmp|crdownload|part|partial|download|!ut)$/i;

  function safeName(r, ext) {
    let base = r.doc + (r.rev ? '-' + r.rev : '') + (r.title ? '_' + r.title : '');
    base = base.replace(/[\\/:*?"<>|\r\n\t]/g, '-').replace(/\s+/g, ' ').replace(/[. ]+$/g, '').trim();
    if (base.length > CFG.maxNameLen) base = base.slice(0, CFG.maxNameLen).trim();
    return base + '.' + String(ext || 'pdf').toLowerCase();
  }

  const sanitizeFolder = (s) => String(s).replace(/[\\/:*?"<>|\r\n\t]/g, '-')
    .replace(/\s+/g, ' ').replace(/[. ]+$/g, '').trim().slice(0, 100) || '_Unsorted';

  // ============ กลุ่มเอกสาร — เอาจาก ConZoL ล้วน ๆ ============
  // หัวกลุ่มในผลค้นหาหน้าตาแบบนี้:  "MA-CAL:MARINE Calculation"
  // เก็บไว้ใน localStorage เพื่อใช้ตอนจัดโฟลเดอร์ไฟล์ที่โหลดไปแล้ว
  const GC_KEY = 'edms_groupcache_v1';
  let groupCache = (() => {
    try { return JSON.parse(localStorage.getItem(GC_KEY) || '{}'); } catch (e) { return {}; }
  })();
  const saveGroupCache = () => { try { localStorage.setItem(GC_KEY, JSON.stringify(groupCache)); } catch (e) {} };
  function rememberGroup(doc, code, name) {
    if (!doc || !code) return;
    groupCache[doc.toUpperCase()] = { c: code, n: name };
  }

  // แยกเลขพื้นที่ออกจากเลขเอกสาร โดยอ้างจากรหัสกลุ่มที่ ConZoL บอก
  //   doc GMTP-1400-MA-DWG-401  group MA-DWG  -> area 1400
  //   doc GMTP-CAZ-COJ-MS-007   group COJ-MS  -> area CAZ
  //   doc GMTP-MA-RPT-001       group MA-RPT  -> area (ไม่มี)
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

  // ============ IndexedDB — จำโฟลเดอร์ที่เลือกไว้ ============
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
      if (entry.name.lastIndexOf('.') <= 0) continue;   // ต้องมีนามสกุล
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

  // Rev เก่าไปอยู่ _Superseded ที่ซ้อนอยู่ในโฟลเดอร์ของเอกสารนั้นเอง
  //   MA-DWG MARINE Drawing\1400 Berth 1\_Superseded\
  async function moveToSuperseded(item, destParts) {
    const sup = await ensureDir((destParts || []).concat(CFG.supersededDir));
    try {
      if (item.handle.move) { await item.handle.move(sup, item.name); return true; }
    } catch (e) { /* ลองวิธีสำรอง */ }
    try {
      const file = await item.handle.getFile();
      await writeInto(sup, item.name, file);
      await item.parent.removeEntry(item.name);
      return true;
    } catch (e) { return false; }
  }

  // ============ ดาวน์โหลด ============
  async function fetchDoc(url) {
    const res = await fetch(url, { credentials: 'same-origin' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const ct = (res.headers.get('content-type') || '').toLowerCase();
    const blob = await res.blob();
    if (ct.includes('text/html') || blob.size < 1024) throw new Error('ไม่ได้ไฟล์ (session หมดอายุ?)');
    return blob;
  }

  function browserDownload(name, blob) {
    const url = URL.createObjectURL(blob);
    const a = $('a', { href: url, download: name });
    a.style.display = 'none';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  // ============ ค้นหาผ่าน fetch ============
  // อ่านตารางผลค้นหา — เก็บทั้งหัวกลุ่มและแถวเอกสาร ตามลำดับที่ ConZoL จัดมา
  function parseRows(root) {
    const rows = [];
    let gCode = '', gName = '';
    for (const tr of root.querySelectorAll('tr')) {
      const cells = tr.cells;
      if (!cells) continue;

      // แถวหัวกลุ่ม: ช่องเดียว colspan กว้าง ข้อความ "CODE:ชื่อกลุ่ม"
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
      // คอลัมน์ FILE+ : ไฟล์แนบต้นฉบับ (ปกติเป็น .ZIP)
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
    if (!res.ok) throw new Error('ค้นหาไม่สำเร็จ HTTP ' + res.status);
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

  // ============ อ่านไฟล์ MDR ============
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
      #edmsdl button.chk{background:#e8e2d5;color:#5a4a33;border-color:#b3a58c}
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

      <fieldset><legend>โฟลเดอร์ปลายทาง</legend>
        <div id="edl-dirinfo" class="wn">ยังไม่ได้เลือกโฟลเดอร์</div>
        <button class="pick" id="edl-pick">เลือกโฟลเดอร์…</button>
        <button id="edl-rescan">สแกนใหม่</button>
        <button id="edl-tidy">จัดโฟลเดอร์ใหม่</button>
      </fieldset>

      <fieldset><legend>1) จากหน้าผลค้นหาปัจจุบัน</legend>
        <label><input type="checkbox" id="edl-checked"> เฉพาะแถวที่ติ๊ก checkbox</label>
        <button class="chk" id="edl-check">เช็คก่อน (ไม่โหลดจริง)</button>
        <button class="go" id="edl-run">โหลดจากหน้านี้</button>
      </fieldset>

      <fieldset><legend>2) จากไฟล์ MDR (อัตโนมัติ)</legend>
        <input type="file" id="edl-file" accept=".xlsx,.xlsm,.csv,.txt">
        <div style="margin-top:4px">ชีต: <select id="edl-sheet"><option value="*">— เลือกไฟล์ก่อน —</option></select></div>
        <div id="edl-mdrinfo" class="sk" style="margin-top:3px">ยังไม่ได้โหลดรายการ</div>
        <div style="margin-top:4px">ไม่เอาเอกสารเหล่านี้:</div>
        <textarea id="edl-exclude" rows="2" style="width:100%;font:10px Consolas,monospace" placeholder="เช่น GMTP-MA-RPT-012"></textarea>
        <button class="chk" id="edl-checkmdr">เช็คก่อน (ไม่โหลดจริง)</button>
        <button class="go2" id="edl-runmdr">ค้น ConZoL + โหลดทั้งหมด</button>
      </fieldset>

      <fieldset><legend>3) ทำรายการเอกสารเป็น Excel</legend>
        <div>discipline: <input id="edl-listdisc" style="width:210px;font:10px Consolas,monospace" placeholder="เว้นว่าง = ทุกอย่าง · เช่น MA-DWG, MA-CAL"></div>
        <button class="chk" id="edl-list">สร้างไฟล์รายการ (.xlsx)</button>
      </fieldset>

      <fieldset style="margin-top:2px"><legend>ชนิดไฟล์ที่จะโหลด</legend>
        <label><input type="checkbox" id="edl-getpdf" checked> PDF (คอลัมน์ PDF+)</label>
        <label><input type="checkbox" id="edl-getfile"> ไฟล์แนบต้นฉบับ (คอลัมน์ FILE+ · .zip)</label>
      </fieldset>
      <label><input type="checkbox" id="edl-skip" checked> ข้ามไฟล์ที่มีอยู่ในโฟลเดอร์แล้ว</label>
      <label><input type="checkbox" id="edl-sup" checked> ย้าย Rev เก่าเข้า _Superseded</label>
      <label><input type="checkbox" id="edl-area" checked> แยกโฟลเดอร์ย่อยตามพื้นที่ (1400 / 0500 / PCC …)</label>
      <label><input type="checkbox" id="edl-inactive"> ค้นรวมเอกสารที่ไม่ Active</label>
      <div>
        <button class="stop" id="edl-stop">หยุด</button>
        <button id="edl-csv">บันทึกรายงาน CSV</button>
      </div>
      <div class="st" id="edl-status">พร้อมทำงาน</div>
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

  // ---- โฟลเดอร์ปลายทาง ----
  function showDirInfo(extra) {
    const d = el('edl-dirinfo');
    if (!rootDir) {
      d.className = 'wn';
      d.textContent = FSA ? 'ยังไม่ได้เลือกโฟลเดอร์ (จะเซฟลง Downloads แทน)'
                          : 'เบราว์เซอร์นี้ไม่รองรับ — จะเซฟลง Downloads';
      return;
    }
    d.className = 'ok';
    let n = 0; existing.forEach((v) => { n += v.length; });
    d.textContent = `📁 ${rootDir.name} · มีไฟล์อยู่แล้ว ${n}` + (extra ? ' · ' + extra : '');
  }

  async function useDir(handle, quiet) {
    rootDir = handle;
    await refreshExisting();
    showDirInfo();
    if (!quiet) log(`· ใช้โฟลเดอร์ "${handle.name}" — สแกนเจอไฟล์เดิม ${[...existing.values()].reduce((a, b) => a + b.length, 0)}`, 'ok');
  }

  el('edl-pick').onclick = async () => {
    if (!FSA) { log('เบราว์เซอร์นี้ไม่รองรับการเลือกโฟลเดอร์ ใช้ Chrome/Edge เวอร์ชันใหม่', 'er'); return; }
    try {
      const h = await window.showDirectoryPicker({ mode: 'readwrite', id: 'edmsdl' });
      await idbPut('dir', h);
      await useDir(h);
    } catch (e) { if (e.name !== 'AbortError') log('เลือกโฟลเดอร์ไม่สำเร็จ: ' + e.message, 'er'); }
  };

  el('edl-rescan').onclick = async () => {
    if (!rootDir) return;
    await refreshExisting(); showDirInfo();
    log('· สแกนโฟลเดอร์ใหม่แล้ว', 'sk');
  };

  // ---- จัดโฟลเดอร์ใหม่ทั้งหมดตามผังปัจจุบัน ----
  async function moveTo(item, dirParts) {
    if ((item.path || []).join('\\') === dirParts.join('\\')) return 'same';
    const dir = await ensureDir(dirParts);
    try {
      if (item.handle.move) { await item.handle.move(dir, item.name); return 'moved'; }
    } catch (e) { /* ลองวิธีสำรอง */ }
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
    if (!rootDir) { log('· เลือกโฟลเดอร์ปลายทางก่อน', 'er'); return; }
    if (!(await ensurePermission())) { log('· ไม่ได้รับสิทธิ์เขียนโฟลเดอร์', 'er'); return; }
    running = true; logEl.innerHTML = ''; lastReport = [];
    log('กำลังสแกนโฟลเดอร์…');
    await refreshExisting();

    // หากลุ่มของเอกสารที่ยังไม่รู้ — ถาม ConZoL ทีละตัว
    const unknown = [...existing.keys()].filter((d) => !groupCache[d]);
    if (unknown.length) {
      log(`ถาม ConZoL หากลุ่มของ ${unknown.length} เอกสาร…`);
      for (const d of unknown) {
        if (stopFlag) break;
        try {
          const { rows } = await searchPage({ search: d, worktype: '', dstatus: '' });
          const hit = rows.find((r) => norm(r.doc) === norm(d));
          if (hit && hit.groupCode) log(`   ${d} → ${hit.groupCode} ${hit.groupName}`, 'sk');
          else log(`   ${d} → ไม่พบใน ConZoL`, 'wn');
        } catch (e) { log(`   ${d} → ${e.message}`, 'er'); }
        await sleep(CFG.searchDelayMs);
      }
      saveGroupCache();
    }

    let moved = 0, sup = 0, same = 0, fail = 0;
    const doSup = el('edl-sup').checked;

    for (const [doc, list] of [...existing.entries()]) {
      const parts = targetPath(doc, '', '');

      // แยกตามนามสกุลก่อน — PDF กับไฟล์แนบเป็นไฟล์คนละชนิดของเอกสารเดียวกัน
      // ไม่ใช่ Rev เก่าของกันและกัน ถ้าไม่แยกจะมีตัวหนึ่งโดนย้ายเข้า _Superseded
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
          log(`✗ ${it.name} → ${e.name === 'NotAllowedError' ? 'ไม่มีสิทธิ์' : e.message} (ไฟล์อาจเปิดค้างอยู่)`, 'er');
        }
      }
      }
    }

    try { await pruneEmpty(rootDir, []); } catch (e) {}
    await refreshExisting(); showDirInfo();
    statusEl.textContent = `จัดเสร็จ: ย้าย ${moved} · เข้าที่แล้ว ${same} · _Superseded ${sup} · ผิดพลาด ${fail}`;
    log('— จบการจัดโฟลเดอร์ —');
    running = false;
  };

  async function ensurePermission() {
    if (!rootDir) return false;
    const opt = { mode: 'readwrite' };
    if ((await rootDir.queryPermission(opt)) === 'granted') return true;
    return (await rootDir.requestPermission(opt)) === 'granted';
  }

  // คืนค่าโฟลเดอร์ที่เคยเลือกไว้ (ต้องกดอนุญาตอีกครั้งตอนใช้งานจริง)
  (async () => {
    if (!FSA) { showDirInfo(); return; }
    try {
      const h = await idbGet('dir');
      if (!h) { showDirInfo(); return; }
      if ((await h.queryPermission({ mode: 'readwrite' })) === 'granted') {
        await useDir(h, true);
        log(`· จำโฟลเดอร์ไว้: ${h.name}`, 'ok');
      } else {
        rootDir = h;
        el('edl-dirinfo').className = 'wn';
        el('edl-dirinfo').textContent = `📁 ${h.name} — กดปุ่มโหลดแล้วอนุญาตสิทธิ์อีกครั้ง`;
      }
    } catch (e) { showDirInfo(); }
  })();

  // เช็คเงียบ ๆ ว่ามีเวอร์ชันใหม่บน GitHub ไหม — เข้าไม่ถึงก็ข้ามไป ไม่รบกวน
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
      b.title = 'มีเวอร์ชันใหม่ กดเพื่อติดตั้ง';
      b.onclick = (e) => { e.stopPropagation(); window.open(UPDATE_URL, '_blank'); };
      log('มีเวอร์ชันใหม่ ' + m[1] + ' (ตอนนี้ ' + VERSION + ') — กดป้ายส้มบนหัวกล่องเพื่อติดตั้ง', 'wn');
    } catch (e) { /* เงียบ */ }
  })();

  el('edl-stop').onclick = () => { stopFlag = true; statusEl.textContent = 'กำลังหยุด…'; };

  // ---- เลือกไฟล์ MDR ----
  let workbook = null;
  function showMdrInfo(fname) {
    const i = el('edl-mdrinfo');
    i.className = 'ok';
    i.textContent = `ใช้ ${mdrList.length} เลขเอกสาร` +
      (mdrExcluded ? ` · ตัดที่ถูกลบใน MDR ${mdrExcluded}` : '') + (fname ? ` · ${fname}` : '');
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
        el('edl-sheet').innerHTML = '<option value="*">(ไฟล์ข้อความ)</option>';
      } else {
        await ensureXLSX();
        workbook = XLSX.read(await f.arrayBuffer(), { type: 'array', cellStyles: true });
        el('edl-sheet').innerHTML = '<option value="*">— ทุกชีต —</option>' +
          workbook.SheetNames.map((n) => `<option value="${n.replace(/"/g, '&quot;')}">${n}</option>`).join('');
        mdrList = readWorkbook(workbook, '*');
      }
      mdrList = applyManualExclusions(mdrList);
      showMdrInfo(f.name);
      log(`· อ่าน ${f.name}: ใช้ ${mdrList.length} รายการ` + (mdrExcluded ? ` (ตัดที่ถูกลบ ${mdrExcluded})` : ''), 'ok');
    } catch (e) {
      el('edl-mdrinfo').className = 'er';
      el('edl-mdrinfo').textContent = 'อ่านไฟล์ไม่ได้: ' + e.message;
    }
  };

  // ---- รายงาน CSV ----
  el('edl-csv').onclick = async () => {
    if (!lastReport.length) { log('· ยังไม่มีรายงาน', 'sk'); return; }
    const esc = (s) => '"' + String(s == null ? '' : s).replace(/"/g, '""') + '"';
    const head = ['Document No.', 'Rev', 'Title', 'Result', 'File', 'Folder', 'Sheet'];
    const csv = '﻿' + [head, ...lastReport.map((r) => [r.doc, r.rev, r.title, r.result, r.file, r.folder || '', r.sheet || ''])]
      .map((r) => r.map(esc).join(',')).join('\r\n');
    const d = new Date();
    const stamp = d.getFullYear() + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0');
    const name = `ConZoL_download_report_${stamp}.csv`;
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    if (rootDir && await ensurePermission()) { await writeInto(rootDir, name, blob); log('· บันทึก ' + name + ' ลงโฟลเดอร์แล้ว', 'ok'); }
    else browserDownload(name, blob);
  };

  // ============ ตัวรันดาวน์โหลด ============
  async function runDownload(items, dry) {
    let ok = 0, skipped = 0, fail = 0, sup = 0, i = 0;
    const skip = el('edl-skip').checked;
    const doSup = el('edl-sup').checked;
    const wantPdf = el('edl-getpdf').checked;
    const wantFile = el('edl-getfile').checked;
    const useFS = !!rootDir;

    if (!wantPdf && !wantFile) {
      log('· ยังไม่ได้เลือกว่าจะโหลดอะไร — ติ๊ก PDF หรือ ไฟล์แนบ อย่างน้อยหนึ่งอย่าง', 'er');
      return { ok, skipped, fail, sup };
    }

    for (const r of items) {
      if (stopFlag) { log('■ หยุดโดยผู้ใช้', 'er'); break; }
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
        lastReport.push({ ...r, result: 'ไม่มีไฟล์ตามชนิดที่เลือก', file: '', folder: '' });
        log(`[${i}/${items.length}] ข้าม ${doc}-${r.rev} — ไม่มีไฟล์แนบ`, 'sk');
        continue;
      }

      for (const k of kinds) {
        if (stopFlag) break;
        const name = safeName(r, k.ext);

        if (skip && have.some((h) => h.rev === r.rev.toUpperCase() && (h.ext || 'pdf') === k.ext)) {
          skipped++;
          lastReport.push({ ...r, result: 'ข้าม (มีอยู่แล้ว)', file: name, folder: '' });
          log(`[${i}/${items.length}] ข้าม ${name}`, 'sk');
          continue;
        }

        if (dry) {
          ok++;
          const willSup = [];
          if (doSup) for (const h of have) {
            if ((h.ext || 'pdf') === k.ext && h.rank < rank && h.name !== name) willSup.push(h);
          }
          sup += willSup.length;
          lastReport.push({ ...r, result: 'จะโหลด', file: name, folder: parts.join('\\') });
          log(`[${i}/${items.length}] + ${name}  →  ${parts.join('\\')}`, 'ok');
          for (const h of willSup) {
            lastReport.push({ ...r, rev: h.rev, result: 'จะย้ายเข้า _Superseded', file: h.name,
                              folder: (h.path || parts).concat(CFG.supersededDir).join('\\') });
            log(`      ↳ ${h.name} → _Superseded`, 'wn');
          }
          continue;
        }

        statusEl.textContent = `กำลังโหลด ${i}/${items.length} …`;
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
            lastReport.push({ ...r, result: 'สำเร็จ', file: name, folder: parts.join('\\') });
            log(`[${i}/${items.length}] ✓ ${name} (${(blob.size / 1048576).toFixed(2)} MB)`, 'ok');
          } catch (e) {
            if (a === CFG.retry) {
              fail++;
              lastReport.push({ ...r, result: 'ผิดพลาด: ' + e.message, file: name, folder: '' });
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
      if (!(await ensurePermission())) { log('ไม่ได้รับสิทธิ์เขียนโฟลเดอร์ — จะเซฟลง Downloads แทน', 'wn'); rootDir = null; }
      else if (!existing.size) await refreshExisting();
    } else if (FSA) {
      log('⚠ ยังไม่ได้เลือกโฟลเดอร์ปลายทาง — ไฟล์จะลง Downloads และไม่แยกโฟลเดอร์', 'wn');
    }
  }

  // ============ โหมด 1 ============
  async function runPage(dry) {
    if (running) return;
    running = true; stopFlag = false; logEl.innerHTML = ''; lastReport = [];
    if (dry) log('โหมดเช็คก่อน — ไม่มีการเขียนหรือย้ายไฟล์จริง', 'wn');
    await preflight();
    let rows = parseRows(document);
    if (el('edl-checked').checked) rows = rows.filter((r) => r.cb && r.cb.checked);
    if (!rows.length) {
      statusEl.textContent = 'ไม่พบเอกสารในหน้านี้';
      log('กด SEARCH ให้มีผลลัพธ์ก่อน', 'er'); running = false; return;
    }
    log(`พบ ${rows.length} รายการในหน้านี้`);
    const s = await runDownload(rows.map((r) => ({ ...r, sheet: '' })), dry);
    statusEl.textContent = dry
      ? `ผลการเช็ค: จะโหลด ${s.ok} · ข้าม ${s.skipped}` + (s.sup ? ` · เก่า→_Superseded ${s.sup}` : '')
      : `เสร็จ: สำเร็จ ${s.ok} · ข้าม ${s.skipped} · ผิดพลาด ${s.fail}` + (s.sup ? ` · เก่า→_Superseded ${s.sup}` : '');
    showDirInfo();
    log(dry ? '— จบการเช็ค (ยังไม่ได้โหลด) — กด "บันทึกรายงาน CSV" เพื่อเก็บผล' : '— จบการทำงาน —');
    running = false;
  }
  el('edl-run').onclick = () => runPage(false);
  el('edl-check').onclick = () => runPage(true);

  // ============ โหมด 2 ============
  async function runMdr(dry) {
    if (running) return;
    if (!mdrList.length) { log('· ยังไม่ได้เลือกไฟล์ MDR', 'er'); return; }
    running = true; stopFlag = false; logEl.innerHTML = ''; lastReport = [];
    if (dry) log('โหมดเช็คก่อน — ค้น ConZoL อย่างเดียว ไม่โหลดไฟล์จริง', 'wn');
    await preflight();

    const dstatus = el('edl-inactive').checked ? '' : 'A';
    const codes = disciplineCodes();
    const wanted = new Map(mdrList.map((m) => [norm(m.doc), m]));

    const byDisc = new Map(); const noDisc = [];
    for (const m of mdrList) {
      const d = disciplineOf(m.doc, codes);
      if (d) { if (!byDisc.has(d)) byDisc.set(d, []); byDisc.get(d).push(m); } else noDisc.push(m);
    }
    log(`MDR ${mdrList.length} รายการ · discipline: ${[...byDisc.keys()].join(', ') || '-'}`);

    const index = new Map();
    for (const [disc] of byDisc) {
      if (stopFlag) break;
      statusEl.textContent = `กำลังค้น ConZoL: ${disc} …`;
      let page = 1, total = 1;
      do {
        const extra = { worktype: disc, dstatus, page: String(page) };
        if (page > 1) extra.pagechange = '1';
        const { rows, pages } = await searchPage(extra);
        total = Math.max(pages.length || 1, total);
        rows.forEach((r) => { const k = norm(r.doc); if (!index.has(k)) index.set(k, r); });
        log(`  · ${disc} หน้า ${page}/${total} → ${rows.length}`, 'sk');
        page++; await sleep(CFG.searchDelayMs);
      } while (page <= total && !stopFlag);
    }

    const probe = [...wanted.keys()].filter((k) => !index.has(k));
    if (probe.length && !stopFlag) {
      log(`ค้นเพิ่มทีละรายการ ${probe.length} ตัว…`, 'wn');
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
    log(`จับคู่ได้ ${items.length} / ${wanted.size}`, items.length ? 'ok' : 'er');
    if (notFound.length) {
      log(`ไม่พบใน ConZoL ${notFound.length} รายการ:`, 'wn');
      notFound.forEach((m) => log('   – ' + m.doc, 'wn'));
      notFound.forEach((m) => lastReport.push({ doc: m.doc, rev: '', title: m.title || '', result: 'ไม่พบใน ConZoL', file: '', folder: '', sheet: m.sheet || '' }));
    }

    if (items.length) {
      const s = await runDownload(items, dry);
      statusEl.textContent = dry
        ? `ผลการเช็ค: จะโหลด ${s.ok} · ข้าม ${s.skipped} · ไม่พบ ${notFound.length}` + (s.sup ? ` · เก่า→_Superseded ${s.sup}` : '')
        : `เสร็จ: สำเร็จ ${s.ok} · ข้าม ${s.skipped} · ผิดพลาด ${s.fail} · ไม่พบ ${notFound.length}` + (s.sup ? ` · เก่า→_Superseded ${s.sup}` : '');
    } else statusEl.textContent = `ไม่มีรายการให้โหลด (ไม่พบ ${notFound.length})`;
    showDirInfo();
    log(dry ? '— จบการเช็ค (ยังไม่ได้โหลด) — กด "บันทึกรายงาน CSV" เพื่อเก็บผล'
            : '— จบการทำงาน — กด "บันทึกรายงาน CSV" เพื่อเก็บผล');
    running = false;
  }
  el('edl-runmdr').onclick = () => runMdr(false);
  el('edl-checkmdr').onclick = () => runMdr(true);

  // ============ โหมด 3 — ทำรายการเอกสารเป็น Excel ============
  // ชื่อเรื่องอยู่ในชื่อไฟล์ หลัง "<DocNo>-<Rev>_" จนถึงนามสกุล
  function titleFromFile(name) {
    const m = FN_RE.exec(name);
    if (!m) return '';
    let t = name.slice(m[0].length);
    const dot = t.lastIndexOf('.');
    if (dot > 0) t = t.slice(0, dot);
    return t.trim();
  }

  async function collectFolderRows() {
    const out = [];
    if (!rootDir) return out;
    if (!existing.size) await refreshExisting();
    for (const doc of [...existing.keys()].sort()) {
      const list = existing.get(doc) || [];
      const top = new Map();   // นามสกุล -> rank สูงสุด = ตัวล่าสุด
      for (const it of list) {
        const e = it.ext || 'pdf';
        if (!top.has(e) || it.rank > top.get(e)) top.set(e, it.rank);
      }
      const sorted = [...list].sort((a, b) => (b.rank - a.rank) ||
        String(a.ext || '').localeCompare(String(b.ext || '')));
      for (const it of sorted) {
        let size = '', mtime = '';
        try {
          const f = await it.handle.getFile();
          size = Math.round(f.size / 10485.76) / 100;
          const d = new Date(f.lastModified);
          mtime = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') +
                  '-' + String(d.getDate()).padStart(2, '0');
        } catch (e) {}
        out.push({
          doc, rev: it.rev, rank: it.rank,
          title: titleFromFile(it.name),
          file: it.name,
          ext: String(it.ext || '').toUpperCase(),
          folder: (it.path || []).join('\\') || '.',
          status: it.rank === top.get(it.ext || 'pdf') ? 'Current' : 'Superseded',
          size, mtime
        });
      }
    }
    return out;
  }

  async function collectConzolRows(discList) {
    const dstatus = el('edl-inactive').checked ? '' : 'A';
    const index = new Map();
    for (const disc of discList) {
      if (stopFlag) break;
      let page = 1, total = 1;
      do {
        statusEl.textContent = `กำลังค้น ConZoL: ${disc || 'ทั้งหมด'} หน้า ${page} …`;
        const extra = { worktype: disc, dstatus, page: String(page) };
        if (page > 1) extra.pagechange = '1';
        const { rows, pages } = await searchPage(extra);
        total = Math.max(pages.length || 1, total);
        rows.forEach((r) => { const k = norm(r.doc); if (!index.has(k)) index.set(k, { ...r, disc }); });
        log(`  · ${disc || 'ทั้งหมด'} หน้า ${page}/${total} → ${rows.length}`, 'sk');
        page++; await sleep(CFG.searchDelayMs);
      } while (page <= total && !stopFlag);
    }
    return [...index.values()];
  }

  function mkSheet(aoa, widths) {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = widths.map((w) => ({ wch: w }));
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 },
      e: { r: Math.max(aoa.length - 1, 1), c: widths.length - 1 } }) };
    return ws;
  }

  async function makeLists() {
    if (running) return;
    running = true; stopFlag = false; logEl.innerHTML = '';
    try {
      await ensureXLSX();

      // ---- ชีต 1: ไฟล์ที่มีอยู่ในโฟลเดอร์ ----
      let fRows = [];
      if (rootDir && await ensurePermission()) {
        statusEl.textContent = 'กำลังอ่านโฟลเดอร์ …';
        log('อ่านไฟล์ในโฟลเดอร์ปลายทาง …');
        await refreshExisting();
        fRows = await collectFolderRows();
        log(`  พบไฟล์ ${fRows.length} ไฟล์ · ${existing.size} เลขเอกสาร`, 'ok');
      } else {
        log('· ยังไม่ได้เลือกโฟลเดอร์ปลายทาง — ชีต Folder จะว่าง', 'wn');
      }

      // ---- ชีต 2: เอกสารใน ConZoL ----
      const codes = disciplineCodes();
      const typed = String(el('edl-listdisc').value || '')
        .split(/[,\s]+/).map((s) => s.trim().toUpperCase()).filter(Boolean);
      const bad = typed.filter((d) => !codes.has(d));
      if (bad.length) log('· ไม่มี discipline นี้ใน ConZoL: ' + bad.join(', '), 'wn');
      const discList = typed.filter((d) => codes.has(d));
      if (!discList.length) discList.push(...[...codes].sort());
      log(`ค้น ConZoL ${discList.length} discipline: ${discList.join(', ')}`);
      const cRows = await collectConzolRows(discList);
      log(`  พบเอกสาร ${cRows.length} รายการ`, cRows.length ? 'ok' : 'wn');

      // ---- ชีต 3: เทียบกัน ----
      const byDocFolder = new Map();
      for (const r of fRows) {
        const k = norm(r.doc);
        if (!byDocFolder.has(k) || r.rank > byDocFolder.get(k).rank) byDocFolder.set(k, r);
      }
      const byDocConzol = new Map();
      for (const r of cRows) byDocConzol.set(norm(r.doc), r);

      const rankOf = (rev) => { const m = /^([TR])(\d+)$/i.exec(String(rev || '')); return m ? revRank(m[1], m[2]) : -1; };
      const cmp = [];
      for (const k of new Set([...byDocConzol.keys(), ...byDocFolder.keys()])) {
        const c = byDocConzol.get(k), f = byDocFolder.get(k);
        let result;
        if (c && !f) result = 'ยังไม่มีในโฟลเดอร์ — ต้องโหลด';
        else if (f && !c) result = 'ไม่พบใน ConZoL';
        else {
          const rc = rankOf(c.rev), rf = rankOf(f.rev);
          if (rc === rf) result = 'ตรงกัน';
          else if (rc > rf) result = 'ConZoL ใหม่กว่า — ต้องโหลด';
          else result = 'ในโฟลเดอร์ใหม่กว่า ConZoL';
        }
        cmp.push({
          doc: (c && c.doc) || (f && f.doc) || k,
          title: (c && c.title) || (f && f.title) || '',
          crev: c ? c.rev : '', frev: f ? f.rev : '',
          folder: f ? f.folder : (c ? targetPath(c.doc, c.groupCode, c.groupName).join('\\') : ''),
          result
        });
      }
      cmp.sort((a, b) => a.doc.localeCompare(b.doc));

      // ---- เขียนไฟล์ ----
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, mkSheet(
        [['S/N', 'Document No.', 'Rev', 'Title', 'File Name', 'Type', 'Folder', 'Status', 'Size (MB)', 'Modified']]
          .concat(fRows.map((r, i) => [i + 1, r.doc, r.rev, r.title, r.file, r.ext, r.folder, r.status, r.size, r.mtime])),
        [6, 26, 6, 46, 52, 7, 34, 12, 10, 12]), 'Folder');

      XLSX.utils.book_append_sheet(wb, mkSheet(
        [['S/N', 'Document No.', 'Rev', 'Title', 'Discipline', 'Group Code', 'Group Name', 'Area', 'PDF', 'Native File', 'Target Folder']]
          .concat(cRows.map((r, i) => [i + 1, r.doc, r.rev, r.title, r.disc || '', r.groupCode || '', r.groupName || '',
            areaOf(r.doc, r.groupCode), 'Y', r.fileHref ? String(r.fileExt || 'zip').toUpperCase() : '',
            targetPath(r.doc, r.groupCode, r.groupName).join('\\')])),
        [6, 26, 6, 46, 11, 12, 30, 10, 5, 11, 34]), 'ConZoL');

      XLSX.utils.book_append_sheet(wb, mkSheet(
        [['S/N', 'Document No.', 'Title', 'ConZoL Rev', 'Folder Rev', 'Folder', 'Result']]
          .concat(cmp.map((r, i) => [i + 1, r.doc, r.title, r.crev, r.frev, r.folder, r.result])),
        [6, 26, 46, 11, 11, 34, 30]), 'Compare');

      const d = new Date();
      const stamp = d.getFullYear() + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0');
      const name = `ConZoL_document_list_${stamp}.xlsx`;
      const blob = new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })],
        { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      if (rootDir && await ensurePermission()) { await writeInto(rootDir, name, blob); log('· บันทึก ' + name + ' ลงโฟลเดอร์แล้ว', 'ok'); }
      else browserDownload(name, blob);

      const need = cmp.filter((r) => /ต้องโหลด/.test(r.result)).length;
      statusEl.textContent = `รายการเสร็จ: โฟลเดอร์ ${fRows.length} ไฟล์ · ConZoL ${cRows.length} รายการ · ยังไม่ครบ ${need}`;
      log('— จบการทำรายการ —');
    } catch (e) {
      log('· ทำรายการไม่สำเร็จ: ' + e.message, 'er');
      statusEl.textContent = 'ทำรายการไม่สำเร็จ';
    }
    showDirInfo();
    running = false;
  }
  el('edl-list').onclick = () => makeLists();
})();
