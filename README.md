# ConZoL Auto Download

A browser tool for the GULF ConZoL EDMS (`edms.gulf.co.th`).
It downloads documents, names them properly, and files them into folders — automatically.

ConZoL gives you `GMTP-1400-MA-DWG-401-T3.pdf`.
This gives you `GMTP-1400-MA-DWG-401-T3_Berth # 1 Mooring Dolphin Deck Plan.pdf`, already in
`MA-DWG MARINE Drawing\1400 Berth 1\`.

---

## Install (choose one)

### A. Tampermonkey — recommended, updates itself

1. Install [Tampermonkey](https://chromewebstore.google.com/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo)
2. Click **[ConZoL-Auto-Download.user.js](../../raw/main/ConZoL-Auto-Download.user.js)** → Install

New versions arrive on their own. Thai UI: [ConZoL-Auto-Download.th.user.js](../../raw/main/ConZoL-Auto-Download.th.user.js)

### B. Chrome extension — no Tampermonkey needed

1. Download this repository (**Code → Download ZIP**) and unzip somewhere permanent
2. `chrome://extensions` → **Developer mode** on → **Load unpacked**
3. Select the `extension` folder

No auto-update on this route.

---

## Use

1. Open `https://edms.gulf.co.th/dms/drawing.asp` and log in — a panel appears top right
2. **Choose folder…** → pick where documents go (asked once, remembered)
3. Tick what you want: **PDF** (PDF+ column) and/or **Native attachment** (FILE+ column, `.zip`)
4. Then either
   - run a SEARCH and press **Download this page**, or
   - pick your MDR `.xlsx` and press **Search ConZoL + download all**

**Re-sort folders** tidies files already on disk.

---

## What it handles

| | |
|---|---|
| **Naming** | `<DocNo>-<Rev>_<Title>.<ext>` — title comes from ConZoL |
| **Folders** | Read from ConZoL's own group headers (`MA-CAL:MARINE Calculation`), never hard-coded — new disciplines and document types work with no changes |
| **Areas** | `1400`, `0500`, `PCC`, `CAZ` … become sub-folders; add friendly labels in `AREA_LABELS` |
| **Revisions** | Latest revision stays in place, older ones move to `_Superseded\` |
| **Repeat runs** | Files already on disk are skipped — only new documents and changed revisions download |
| **Cancelled docs** | Rows struck through or marked *Delete* in the MDR are ignored |
| **Reporting** | **Save CSV report** lists what downloaded, what was skipped, what ConZoL does not have |

Files are written straight to your folder through the File System Access API — not through Chrome's
download system. Nothing to configure, no save-as prompts, no leftover `.tmp` files.

Folders named `_Deleted`, `_Archive`, `_Cancelled` are never scanned, moved or removed.

---

## Requirements

- Chrome or Edge 111+
- Reading an `.xlsx` loads a spreadsheet library from `cdnjs.cloudflare.com` on first use

## Releasing an update

1. Edit the `.user.js`, raise `// @version`
2. Commit and push

Tampermonkey installs it for everyone on its own. See [PUBLISH.txt](PUBLISH.txt) for the
`@updateURL` setup and the Chrome Web Store notes.

## Licence

MIT — see [LICENSE](LICENSE).
