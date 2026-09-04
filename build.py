import re, json, sys

SRC = 'EDMS-AutoDownload.user.js'
s = open(SRC, encoding='utf-8').read()

VER = re.search(r'^// @version\s+(\S+)', s, re.M).group(1)
# ซิงก์ค่า VERSION ในโค้ดให้ตรงกับ @version เสมอ
s = re.sub(r"const VERSION = '[^']*';", "const VERSION = '%s';" % VER, s, count=1)
open(SRC, 'w', encoding='utf-8').write(s)

RAW = 'https://raw.githubusercontent.com/SetthawutJanthakomut/conzol-auto-download/main/'

def to_ext(src):
    src = re.sub(r'^// ==UserScript==.*?// ==/UserScript==\n', '', src, flags=re.S)
    loader = """
  const XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  let xlsxLoading = null;
  function ensureXLSX() {
    if (typeof window.XLSX !== 'undefined') return Promise.resolve(window.XLSX);
    if (xlsxLoading) return xlsxLoading;
    xlsxLoading = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = XLSX_URL;
      s.onload = () => (typeof window.XLSX !== 'undefined') ? res(window.XLSX) : rej(new Error('LIBFAIL'));
      s.onerror = () => rej(new Error('LIBNET'));
      document.head.appendChild(s);
      setTimeout(() => rej(new Error('LIBTIMEOUT')), 20000);
    });
    return xlsxLoading;
  }
"""
    src = src.replace("  const $ = (t, p) => Object.assign", loader + "\n  const $ = (t, p) => Object.assign", 1)
    src = src.replace("        if (typeof XLSX === 'undefined') throw new Error('โหลดไลบรารีอ่าน .xlsx ไม่ได้');",
                      "        await ensureXLSX();")
    return src

# ---------- ไทย ----------
ext_th = to_ext(s)
for a, b in [("'LIBFAIL'", "'โหลดไลบรารีไม่สำเร็จ'"),
             ("'LIBNET'", "'เข้าถึง cdnjs.cloudflare.com ไม่ได้'"),
             ("'LIBTIMEOUT'", "'หมดเวลารอไลบรารี'")]:
    ext_th = ext_th.replace(a, b)
open('build-th/conzol.js', 'w', encoding='utf-8').write('// ConZoL Auto Download - Chrome Extension build\n\n' + ext_th)

th_us = s
if '@updateURL' not in th_us:
    u = RAW + 'ConZoL-Auto-Download.th.user.js'
    th_us = th_us.replace('// @run-at       document-idle',
                          '// @updateURL    %s\n// @downloadURL  %s\n// @run-at       document-idle' % (u, u), 1)
open('build-th/ConZoL-Auto-Download.user.js', 'w', encoding='utf-8').write(th_us)

# ---------- อังกฤษ ----------
M = json.load(open('i18n-en.json', encoding='utf-8'))
en = ext_th
miss = []
for a, b in M:
    if a not in en:
        miss.append(a[:60])
    en = en.replace(a, b)
open('build-en/conzol.js', 'w', encoding='utf-8').write(en)

hdr = """// ==UserScript==
// @name         GULF ConZoL - Auto Download + Rename + Sort
// @namespace    gmtp.conzol
// @version      %s
// @description  Download PDFs and native attachments from GULF ConZoL EDMS automatically - names each file and sorts it into the folder ConZoL assigns.
// @match        https://edms.gulf.co.th/dms/drawing.asp*
// @match        http://edms.gulf.co.th/dms/drawing.asp*
// @updateURL    %sConZoL-Auto-Download.user.js
// @downloadURL  %sConZoL-Auto-Download.user.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @run-at       document-idle
// @grant        none
// ==/UserScript==

""" % (VER, RAW, RAW)
open('build-en/ConZoL-Auto-Download.user.js', 'w', encoding='utf-8').write(
    hdr + en.replace('// ConZoL Auto Download - Chrome Extension build\n\n', ''))

# ---------- manifest ----------
for d, desc in [('build-th', 'ดาวน์โหลด PDF และไฟล์แนบจาก GULF ConZoL EDMS อัตโนมัติ ตั้งชื่อไฟล์และแยกโฟลเดอร์ตามกลุ่มเอกสารให้เอง'),
                ('build-en', 'Download PDFs and native attachments from GULF ConZoL EDMS automatically - names each file and sorts it into the right folder.')]:
    json.dump({"manifest_version": 3, "name": "ConZoL Auto Download", "version": VER, "description": desc,
               "content_scripts": [{"matches": ["https://edms.gulf.co.th/dms/drawing.asp*",
                                                "http://edms.gulf.co.th/dms/drawing.asp*"],
                                    "js": ["conzol.js"], "run_at": "document_idle", "world": "MAIN"}],
               "icons": {"16": "icon16.png", "32": "icon32.png", "48": "icon48.png", "128": "icon128.png"},
               "minimum_chrome_version": "111"},
              open(d + '/manifest.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=2)

left = [(i, l.strip()) for i, l in enumerate(en.split('\n'), 1) if re.search(r'[฀-๿]', l)]
print('version:', VER)
print('missing translations:', len(miss))
for m in miss: print('   !', m)
print('thai left in EN build:', len(left))
for i, l in left: print('  ', i, l[:100])
if miss or left: sys.exit(1)
