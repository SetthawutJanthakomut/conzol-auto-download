ConZoL Auto Download   v4.1
===========================================================
Downloads documents from GULF ConZoL EDMS automatically.

  - Names every file    <DocNo>-<Rev>_<Title>.pdf
  - Downloads the PDF, the native attachment (FILE+), or both
  - Sorts each file into the folder ConZoL itself assigns
  - Keeps the latest revision, moves older ones to _Superseded
  - Can work straight from your MDR spreadsheet


-----------------------------------------------------------
INSTALL   (once, about a minute)
-----------------------------------------------------------
1. Unzip to a permanent folder,
   for example   D:\Tools\ConZoL Auto Download

   *** Do not leave it in Downloads and delete it later. ***
   *** Chrome reads the files from that folder forever;  ***
   *** delete the folder and the extension disappears.   ***

2. In Chrome, open:   chrome://extensions

3. Turn on   Developer mode   (top right)

4. Click     Load unpacked   (top left)

5. Select the folder from step 1
   (select the folder itself, not a file inside it)

6. Done - "ConZoL Auto Download" appears in the list


-----------------------------------------------------------
HOW TO USE
-----------------------------------------------------------
1. Open  https://edms.gulf.co.th/dms/drawing.asp  and log in.
   A panel titled "ConZoL Auto Download" appears top right.

2. Click  [Choose folder...]  and pick where documents
   should be stored. Chrome asks for permission once -
   click Allow. The choice is remembered.

3. Under "What to download", tick what you need:
       PDF                  - the PDF+ column
       Native attachment    - the FILE+ column (.zip)
   Both can be ticked at the same time.

4. Pick how to download:

   A - From the current search results
       Run SEARCH as usual, then click [Download this page].

   B - From an MDR file  (automatic)
       Click Choose File and select your MDR (.xlsx),
       then click [Search ConZoL + download all].
       It searches every group and every result page by
       itself and downloads whatever is missing.

5. [Re-sort folders] tidies files already on disk into the
   correct folders.


-----------------------------------------------------------
WHAT IT DOES FOR YOU
-----------------------------------------------------------
- Adds the document title to the file name.
  ConZoL alone gives you only the number, e.g.
      ConZoL      GMTP-1400-MA-DWG-401-T3.pdf
      this tool   GMTP-1400-MA-DWG-401-T3_Berth # 1 Mooring
                  Dolphin Deck Plan.pdf

- Native attachments get the same name with the original
  extension, so the PDF and the source files sit together:
      GMTP-PEC-MA-TQ-002-R0_<title>.pdf
      GMTP-PEC-MA-TQ-002-R0_<title>.zip

- Sorts by the group ConZoL assigns, for example
      MA-DWG MARINE Drawing\1400 Berth 1\
      MA-CAL MARINE Calculation\
      CO-MA CONSTRUCTION Marine (Jetty)\
  Group names are read from ConZoL at run time, so new
  disciplines and document types work with no changes here.

- Skips files you already have. Only new documents and
  changed revisions are downloaded. PDF and attachment are
  tracked separately.

- Moves superseded revisions into  _Superseded\  so the
  working folders always show the current revision only.

- Ignores documents cancelled in the MDR
  (struck through, or marked Delete).

- [Save CSV report] writes a list of what was downloaded,
  what was skipped, and what could not be found in ConZoL.


-----------------------------------------------------------
GOOD TO KNOW
-----------------------------------------------------------
- Files are written straight to your folder, not through
  Chrome's download system. Nothing to configure, no
  save-as prompts, no leftover .tmp files.

- Not every document has a native attachment. Documents
  with nothing in the FILE+ column are reported as
  "No file of the selected type" and skipped.

- Attachments are large. Ticking both roughly triples the
  volume - a single .zip can be several hundred MB.

- Folders named  _Deleted, _Archive, _Cancelled  are never
  scanned, moved or removed.

- Needs a recent Chrome or Edge (File System Access API).

- Reading an .xlsx needs internet access the first time,
  to load a spreadsheet library from cdnjs.cloudflare.com.


-----------------------------------------------------------
Using Tampermonkey instead?  Turn on Allow User Scripts
-----------------------------------------------------------
Not needed for the extension route (Load unpacked), but
Tampermonkey will not run any script until this is enabled.

1. Open in Chrome:
   chrome://extensions/?id=dhdgffkkebhmkfjojejmpbldmpobfkfo

2. Switch  Allow User Scripts  on

3. If the switch is not there, turn on Developer mode
   (top right of chrome://extensions) first and it appears

Symptom when missed: the script installs fine, but no panel
ever shows up on the ConZoL page.

-----------------------------------------------------------
TROUBLESHOOTING
-----------------------------------------------------------
No panel in the top right corner
   - You must be on drawing.asp and logged in
   - Check the extension is enabled in chrome://extensions
   - Press F5 to reload the page

Load unpacked is blocked
   - Your IT policy forbids self-installed extensions.
     Use the Tampermonkey version instead: install the
     Tampermonkey extension, then import the file
     ConZoL-Auto-Download.user.js

"No file returned - session may have expired"
   - Log in to ConZoL again and retry

A file will not move during Re-sort
   - It is open in a PDF viewer, or a sync client is
     holding it. Close it and run Re-sort again.
