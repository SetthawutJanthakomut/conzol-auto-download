# ConZoL Auto Download

A browser tool for the GULF ConZoL EDMS (`edms.gulf.co.th`).
It downloads documents, names them properly, and files them into folders — automatically.

ConZoL gives you `GMTP-1400-MA-DWG-401-T3.pdf`.
This gives you `GMTP-1400-MA-DWG-401-T3_Berth # 1 Mooring Dolphin Deck Plan.pdf`, already in
`MA-DWG MARINE Drawing\1400 Berth 1\`.

Available in **English** and **ไทย** — same features, same version, built from one source.

---

## Install

**Install one language only.** Two copies at once means two panels on the page.

### With Tampermonkey — recommended, updates itself

**1. Install [Tampermonkey](https://chromewebstore.google.com/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo)**

**2. Turn on *Allow User Scripts*** — required, and easy to miss

Chrome will not run any userscript until this is enabled. Open

```
chrome://extensions/?id=dhdgffkkebhmkfjojejmpbldmpobfkfo
```

and switch **Allow User Scripts** on. If the switch is not there, turn on
**Developer mode** (top right of `chrome://extensions`) first and it will appear.

Skip this and the script installs fine but the panel never shows up.

**3. Click the version you want, then press *Install***

| Language | Install link |
|---|---|
| English | [ConZoL-Auto-Download.user.js](../../raw/main/ConZoL-Auto-Download.user.js) |
| ไทย | [ConZoL-Auto-Download.th.user.js](../../raw/main/ConZoL-Auto-Download.th.user.js) |

New versions arrive on their own — nothing to re-send, nobody to chase.

### As a Chrome extension — no Tampermonkey needed

1. **Code → Download ZIP**, unzip somewhere permanent
2. `chrome://extensions` → **Developer mode** on → **Load unpacked**
3. Select `extension-en` (English) or `extension-th` (ไทย)

No auto-update on this route.

---

## Use

1. Open `https://edms.gulf.co.th/dms/drawing.asp` and log in — a panel appears top right
2. **Choose folder…** → pick where documents go (asked once, remembered)
3. Tick what you want: **PDF** (PDF+ column) and/or **Native attachment** (FILE+ column, `.zip`)
4. Then either
   - run a SEARCH and press **Download this page**, or
   - pick your MDR `.xlsx` and press **Search ConZoL + download all**

**Check first (no download)** does a full dry run: it works out exactly what would be downloaded,
what would be skipped and which revisions would move to `_Superseded`, and lists it in the log —
without writing a single file. **Save CSV report** works after a check too, so you can review the
plan in Excel before committing to it.

**Build list (.xlsx)** produces one workbook with three sheets, in the destination folder:

| Sheet | What's in it |
|---|---|
| **Folder** | every document file on disk — document no., revision, title, file name, type, folder, Current/Superseded, size, date |
| **ConZoL** | every document ConZoL holds for the disciplines you asked for — revision, title, group, area, whether a native attachment exists, and the folder it belongs in |
| **Compare** | the two sides side by side — *Up to date*, *ConZoL is newer - to download*, *Not in the folder yet - to download*, *Not found in ConZoL* |
| **MDR** | your own MDR, refreshed — only when an MDR file is loaded in step 2 |

The MDR sheet keeps the register's own columns (S/N, Document No., Title, Activity ID, Budget, Class,
the latest Rev., Issue Status, Owner's Reply, Progress) and adds ConZoL's current revision beside them,
plus **MDR vs ConZoL** (*Up to date* / *ConZoL is newer than the MDR* / *MDR is newer than ConZoL* /
*Different revision series (T/R)*) and **Result**. Column positions are found from the header text, so a
reordered MDR still reads correctly. Rows struck through or marked *Delete* are left out and named in the log.

Leave the discipline box empty for every discipline in ConZoL's dropdown, or type the ones you want
(`MA-DWG, MA-CAL`). It reads like an MDR and opens straight in Excel.

**Re-sort folders** tidies files already on disk.

---

## What it handles

| | |
|---|---|
| **Naming** | `<DocNo>-<Rev>_<Title>.<ext>` — title comes from ConZoL |
| **Folders** | Read from ConZoL's own group headers (`MA-CAL:MARINE Calculation`), never hard-coded — new disciplines and document types work with no changes |
| **Areas** | `1400`, `0500`, `PCC`, `CAZ` … become sub-folders; add friendly labels in `AREA_LABELS` |
| **Revisions** | Latest revision stays in place, older ones move to a `_Superseded\` folder inside the same document folder |
| **Repeat runs** | Files already on disk are skipped — only new documents and changed revisions download |
| **Cancelled docs** | Rows struck through or marked *Delete* in the MDR are ignored |
| **Reporting** | **Save CSV report** lists what downloaded, what was skipped, what ConZoL does not have |

Files are written straight to your folder through the File System Access API — not through Chrome's
download system. Nothing to configure, no save-as prompts, no leftover `.tmp` files.

Folders named `_Deleted`, `_Archive`, `_Cancelled` are never scanned, moved or removed.

**Requires** Chrome or Edge 111+. Reading an `.xlsx` loads a spreadsheet library from
`cdnjs.cloudflare.com` on first use.

---

## Repository layout

```
ConZoL-Auto-Download.user.js      English userscript  ← install link, do not move
ConZoL-Auto-Download.th.user.js   Thai userscript     ← install link, do not move
extension-en/                     English Chrome extension
extension-th/                     Thai Chrome extension
build.py  i18n-en.json  build.sh  build both languages from one source
```

`ConZoL-Auto-Download.th.user.js` is the source of truth. The English build is generated
from it through `i18n-en.json`, so a change only has to be made once.

### Releasing a change

```bash
# 1. edit ConZoL-Auto-Download.th.user.js
# 2. raise // @version
bash build.sh          # rebuilds all four outputs, fails if a string is untranslated
git add -A && git commit -m "..." && git push
```

Tampermonkey installs the new version for everyone on its own.
See [PUBLISH.txt](PUBLISH.txt) for the update-URL setup and Chrome Web Store notes.

## Licence

MIT — see [LICENSE](LICENSE).

---
---

# ConZoL Auto Download (ภาษาไทย)

เครื่องมือช่วยงานเอกสารบน GULF ConZoL EDMS (`edms.gulf.co.th`)
ดาวน์โหลดเอกสาร ตั้งชื่อไฟล์ให้ครบ แล้วแยกเข้าโฟลเดอร์ให้เองอัตโนมัติ

ConZoL ให้มาแค่ `GMTP-1400-MA-DWG-401-T3.pdf`
ตัวนี้ให้ `GMTP-1400-MA-DWG-401-T3_Berth # 1 Mooring Dolphin Deck Plan.pdf`
พร้อมวางไว้ใน `MA-DWG MARINE Drawing\1400 Berth 1\` ให้เรียบร้อย

---

## วิธีติดตั้ง

**เลือกภาษาเดียวเท่านั้น** ถ้าลงทั้งคู่จะมีกล่องซ้อนกัน 2 อัน

### ผ่าน Tampermonkey — แนะนำ อัปเดตให้เอง

**1. ติดตั้ง [Tampermonkey](https://chromewebstore.google.com/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo)**

**2. เปิด *Allow User Scripts*** — ข้อนี้ห้ามข้าม คนลืมกันบ่อย

Chrome จะไม่ยอมรันสคริปต์เลยถ้าไม่เปิดข้อนี้ วางลิงก์นี้ในช่อง URL

```
chrome://extensions/?id=dhdgffkkebhmkfjojejmpbldmpobfkfo
```

แล้วเปิดสวิตช์ **Allow User Scripts** ถ้าไม่เห็นสวิตช์ ให้เปิด
**Developer mode** (มุมขวาบนของหน้า `chrome://extensions`) ก่อน แล้วมันจะโผล่มา

ถ้าข้ามข้อนี้ สคริปต์จะติดตั้งสำเร็จ แต่กล่องจะไม่ขึ้นบนหน้า ConZoL

**3. กดลิงก์ภาษาที่ต้องการ แล้วกด *Install***

| ภาษา | ลิงก์ติดตั้ง |
|---|---|
| ไทย | [ConZoL-Auto-Download.th.user.js](../../raw/main/ConZoL-Auto-Download.th.user.js) |
| English | [ConZoL-Auto-Download.user.js](../../raw/main/ConZoL-Auto-Download.user.js) |

มีเวอร์ชันใหม่เมื่อไหร่ เครื่องทุกคนอัปเดตเอง ไม่ต้องไล่ส่งไฟล์

### แบบ Chrome extension — ไม่ต้องใช้ Tampermonkey

1. **Code → Download ZIP** แล้วแตกไฟล์ลงโฟลเดอร์ถาวร
2. `chrome://extensions` → เปิด **Developer mode** → กด **Load unpacked**
3. เลือกโฟลเดอร์ `extension-th` (ไทย) หรือ `extension-en` (อังกฤษ)

ทางนี้ไม่มีอัปเดตอัตโนมัติ

---

## วิธีใช้

1. เปิด `https://edms.gulf.co.th/dms/drawing.asp` แล้วล็อกอิน — กล่องจะโผล่มุมขวาบน
2. กด **เลือกโฟลเดอร์…** ชี้ที่เก็บเอกสาร (ถามครั้งเดียว จำไว้ให้)
3. ติ๊กชนิดไฟล์: **PDF** (คอลัมน์ PDF+) และ/หรือ **ไฟล์แนบต้นฉบับ** (คอลัมน์ FILE+ · `.zip`)
4. เลือกวิธีโหลด
   - กด SEARCH ตามปกติ แล้วกด **โหลดจากหน้านี้** หรือ
   - เลือกไฟล์ MDR `.xlsx` แล้วกด **ค้น ConZoL + โหลดทั้งหมด**

ปุ่ม **เช็คก่อน (ไม่โหลดจริง)** จะไล่ดูให้ครบว่าจะโหลดอะไรบ้าง ข้ามอะไร และตัวไหนจะถูกย้ายเข้า
`_Superseded` แล้วแสดงในบันทึกการทำงาน โดยไม่เขียนไฟล์จริงสักไฟล์ กด **บันทึกรายงาน CSV**
หลังเช็คได้เลย จะได้เอาไปดูใน Excel ก่อนตัดสินใจโหลดจริง

ปุ่ม **สร้างไฟล์รายการ (.xlsx)** จะได้ไฟล์เดียว 3 ชีต เก็บลงโฟลเดอร์ปลายทางให้เลย

| ชีต | มีอะไร |
|---|---|
| **Folder** | ไฟล์ทุกไฟล์ที่มีอยู่ในโฟลเดอร์ — เลขเอกสาร, Rev, ชื่อเรื่อง, ชื่อไฟล์, ชนิด, โฟลเดอร์, Current/Superseded, ขนาด, วันที่ |
| **ConZoL** | เอกสารทุกตัวที่มีใน ConZoL ตาม discipline ที่เลือก — Rev, ชื่อเรื่อง, กลุ่ม, พื้นที่, มีไฟล์แนบไหม และโฟลเดอร์ที่ควรอยู่ |
| **Compare** | เทียบสองฝั่ง — *Up to date*, *ConZoL ใหม่กว่า — ต้องโหลด*, *ยังไม่มีในโฟลเดอร์ — ต้องโหลด*, *ไม่พบใน ConZoL* |
| **MDR** | MDR ของคุณเอง อัปเดตสถานะให้ — ขึ้นเฉพาะตอนโหลดไฟล์ MDR ในข้อ 2 แล้ว |

ชีต MDR เก็บคอลัมน์เดิมของ MDR ไว้ (S/N, Document No., Title, Activity ID, Budget, Class, Rev. ล่าสุด,
Issue Status, Owner's Reply, Progress) แล้วต่อท้ายด้วย Rev ปัจจุบันใน ConZoL, **MDR vs ConZoL**
(*ตรงกัน* / *ConZoL ใหม่กว่า MDR* / *MDR ใหม่กว่า ConZoL* / *คนละชุด Rev (T/R)*) และ **Result**
ตำแหน่งคอลัมน์อ่านจากข้อความหัวตาราง สลับคอลัมน์ใน MDR ก็ยังอ่านได้ แถวที่ขีดฆ่าหรือเขียน *Delete*
จะถูกตัดออกและแจ้งเลขเอกสารไว้ในบันทึกการทำงาน

ช่อง discipline เว้นว่าง = ไล่ทุกอย่างที่มีใน dropdown ของ ConZoL หรือพิมพ์เฉพาะที่ต้องการก็ได้
(`MA-DWG, MA-CAL`) หน้าตาอ่านง่ายคล้าย MDR เปิดใน Excel ได้เลย

ปุ่ม **จัดโฟลเดอร์ใหม่** ใช้จัดไฟล์เก่าที่กองอยู่ให้เข้าที่

---

## สิ่งที่ทำให้อัตโนมัติ

| | |
|---|---|
| **ชื่อไฟล์** | `<เลขเอกสาร>-<Rev>_<ชื่อเรื่อง>.<นามสกุล>` — ชื่อเรื่องดึงจาก ConZoL |
| **โฟลเดอร์** | อ่านจากหัวกลุ่มที่ ConZoL พิมพ์มาเอง (`MA-CAL:MARINE Calculation`) ไม่ได้ฮาร์ดโค้ด มีกลุ่มใหม่ก็รองรับทันที |
| **พื้นที่** | `1400`, `0500`, `PCC`, `CAZ` แยกเป็นโฟลเดอร์ย่อย ใส่ชื่อกำกับเพิ่มได้ที่ `AREA_LABELS` |
| **Rev** | เก็บ Rev ล่าสุดไว้ ตัวเก่าย้ายเข้า `_Superseded\` ที่อยู่ในโฟลเดอร์เดียวกับเอกสารนั้น |
| **รันซ้ำ** | ข้ามไฟล์ที่มีอยู่แล้ว โหลดเฉพาะเอกสารใหม่กับ Rev ที่เปลี่ยน |
| **เอกสารยกเลิก** | แถวที่ขีดฆ่าหรือเขียน *Delete* ใน MDR จะถูกข้าม |
| **รายงาน** | ปุ่ม **บันทึกรายงาน CSV** สรุปว่าโหลดอะไร ข้ามอะไร อะไรไม่มีใน ConZoL |

เขียนไฟล์ลงโฟลเดอร์โดยตรง ไม่ผ่านระบบ Download ของ Chrome — ไม่ต้องตั้งค่าอะไร
ไม่มีหน้าต่างถามที่เก็บ ไม่มีไฟล์ `.tmp` ค้าง

โฟลเดอร์ชื่อ `_Deleted`, `_Archive`, `_Cancelled` ระบบจะไม่แตะเลย

**ต้องใช้** Chrome หรือ Edge 111 ขึ้นไป · การอ่าน `.xlsx` ครั้งแรกต้องต่อเน็ตได้
เพื่อโหลดไลบรารีจาก `cdnjs.cloudflare.com`

---

## โครงสร้างไฟล์ใน repo

```
ConZoL-Auto-Download.user.js      userscript อังกฤษ  ← เป็นลิงก์ติดตั้ง ห้ามย้าย
ConZoL-Auto-Download.th.user.js   userscript ไทย     ← เป็นลิงก์ติดตั้ง ห้ามย้าย
extension-en/                     Chrome extension อังกฤษ
extension-th/                     Chrome extension ไทย
build.py  i18n-en.json  build.sh  ตัว build ทั้งสองภาษาจากต้นฉบับเดียว
```

ต้นฉบับคือ `ConZoL-Auto-Download.th.user.js` ตัวอังกฤษ generate จากไฟล์นี้ผ่าน `i18n-en.json`
แก้ที่เดียวได้ครบทั้งคู่

### ขั้นตอนออกเวอร์ชันใหม่

```bash
# 1. แก้ ConZoL-Auto-Download.th.user.js
# 2. เพิ่มเลข // @version
bash build.sh          # build ใหม่ทั้ง 4 ไฟล์ ถ้าแปลตกจะฟ้อง error
git add -A && git commit -m "..." && git push
```


---

## Troubleshooting / ปัญหาที่พบบ่อย

**No panel on the ConZoL page / ไม่เห็นกล่องบนหน้า ConZoL**

1. *Allow User Scripts* is off — the most common cause. See install step 2.
   → เปิด *Allow User Scripts* ตามขั้นตอนที่ 2
2. You are not on `drawing.asp`, or not logged in.
   → ต้องอยู่หน้า `drawing.asp` และล็อกอินแล้ว
3. Press F5.

**Buttons do nothing / กดปุ่มแล้วไม่มีอะไรเกิดขึ้น**

Two copies are installed. Check with F12 → Console:

```js
document.querySelectorAll('#edmsdl').length
```

More than 1 means a duplicate — keep one language, and do not run the
userscript and the extension together.
→ ได้มากกว่า 1 แปลว่าลงซ้ำ ให้เหลือภาษาเดียว และอย่าใช้ userscript คู่กับ extension

**Choose folder does nothing / กดเลือกโฟลเดอร์แล้วเงียบ**

Usually the duplicate above. Otherwise check the panel's log line for the
error name — `NotAllowedError` normally means an IT policy blocks websites
from writing files.
→ ส่วนใหญ่มาจากการลงซ้ำข้างบน ถ้าไม่ใช่ ให้ดูบรรทัด log ในกล่องว่าขึ้น error อะไร
`NotAllowedError` มักแปลว่า policy ขององค์กรห้ามเว็บเขียนไฟล์ลงเครื่อง
