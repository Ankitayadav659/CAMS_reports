# CAMS Commission Automation System

Two independent steps in one Windows desktop app (Python + Tkinter):

| | Input | Output |
|---|---|---|
| **Step 1 – Discrepancy Checker** | original CAMS commission Excel | `Discrepancy_Result.xlsx` → sheets *Raw Data (Highlighted)* + *Discrepancy Report* |
| **Step 2 – Investor-wise Automation** | `Discrepancy_Result.xlsx` **+** the large original transaction Excel | `Investor_Wise_Final.xlsx` → *Investor Sheet Index* + one sheet per investor/combination |

You can run Step 1, close the app, and run Step 2 later.

---------------------------------------------------------------------------------------------------

## SETUP — Hinglish, step by step (VS Code beginner)

1. **Folder:** ZIP ko extract karein, jaise `C:\CAMS_Commission_Automation`. (Folder ke andar `main.py` dikhna chahiye.)
2. **Python install:** https://www.python.org/downloads/ se **64-bit Python 3.12** download karein. Installer ki *pehli screen* par **"Add python.exe to PATH"** ko tick karein, phir *Install Now*.
3. **VS Code:** Start menu → *Visual Studio Code* → **File ▸ Open Folder…** → `CAMS_Commission_Automation` folder select karein. Extensions panel (Ctrl+Shift+X) me **Python** (Microsoft) install karein.
4. **Terminal kholein:** VS Code me **Terminal ▸ New Terminal**.
5. **Virtual environment banayein** (terminal me):
   ```
   python -m venv .venv
   ```
6. **Interpreter select karein:** `Ctrl+Shift+P` → *Python: Select Interpreter* → `.\.venv\Scripts\python.exe` chunein.
   Agar PowerShell activation block kare to ek baar chalayein: `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`, phir terminal band karke naya kholein.
7. **Libraries install:**
   ```
   .\.venv\Scripts\python -m pip install -r requirements.txt
   ```
8. **App chalayein:** `.\.venv\Scripts\python main.py` (ya VS Code me `main.py` kholkar ▶ Run). App ka window khulega.
9. **Step 1:** tab *Step 1* → **Browse** se original CAMS Excel → output name auto-fill ho jayega (`Discrepancy_Result.xlsx`) → **Generate Discrepancy Report** → progress bar/log dekhein → *Done* message aane par **Open Output Folder**.
10. **Step 2:** tab *Step 2* → pehle box me **Discrepancy_Result.xlsx**, doosre box me **badi original transaction Excel** → **Generate Investor-wise Sheets**.
11. **Verify:** `Investor_Wise_Final.xlsx` kholein → pehli sheet *Investor Sheet Index* (har investor ka count, negative-brokerage count, status) → investor sheet me **Brokerage** column `BRKAGE_RAT` ke turant baad; negative rows `TRADE_DATE` se `Brokerage` tak light red.
12. **Aage se VS Code ke bina:** folder me **`Start_Application.bat` par double-click**. (Pehli baar khud `.venv` bana kar libraries install karta hai; baad me seedha app kholta hai. Error dekhna ho to `Start_Application_Debug.bat`.) Desktop shortcut: `Start_Application.bat` par right-click ▸ *Send to ▸ Desktop (create shortcut)*.
13. **Desktop icon (recommended):** `Create_Desktop_Shortcut.bat` par double-click karein → desktop par app ke apne icon wala "CAMS Commission Automation" shortcut ban jayega (pehle ek baar `Start_Application.bat` chala lein).

Command line (optional):
```
.\.venv\Scripts\python main.py step1 "cams comm.xlsx" "Discrepancy_Result.xlsx"
.\.venv\Scripts\python main.py step2 "Discrepancy_Result.xlsx" "c(1).xlsx" "Investor_Wise_Final.xlsx"
```
Tests: `.\.venv\Scripts\python -m unittest discover -s tests -t .`

---------------------------------------------------------------------------------------------------

## ↻ Update Existing Sheets (for files you already have)
Tab **Update Existing Sheets**: select a **folder** that contains your final investor Excel files (with or without the Brokerage column) → **Update All Sheets**. New copies are written to the output folder (default `<your folder>\Updated_Sheets`) with the recalculated Brokerage, the **Rate Change** column and the **Rate Summary** sheet. **Originals are never changed**; the output folder must be a different folder. Sheets are recognised by the columns `TRADE_DATE` and `BRKAGE_RAT`; other files/sheets (index, old summaries) are skipped, a broken file does not stop the rest, and running it again on the updated copies is safe. Only `.xlsx` / `.xlsm` files are supported (not PDF).
Command line: `python main.py update "C:\FinalSheets" "C:\FinalSheets_Updated"`

## 📊 Rate Change column + Rate Summary sheet (per client)
Every investor file now explains *how the rate moved*:
* **Rate Change** column right after Brokerage: `▲ Increased` (Brokerage negative), `▼ Decreased` (positive), `Same` (0). A hover note on the Brokerage header explains: **minus (−) = rate increased**.
* **Rate Summary** sheet (first sheet of the file; one per Scheme+Folio+Txn if the investor has several — named `Summary <sheet>`): a simple one-page report, no charts. **How to read** (minus = rate went UP, plus = rate went DOWN) → **Summary** (total transactions, first–last date, how many times the rate went up / down, current rate) → **Rate periods table** with only 5 columns: *No. | From Date | To Date | Rate | What happened* (e.g. "▲ Rate went UP by 0.3961 (was 0.4200)"). Red row = went up, blue row = went down, grey = first rate.
* The summary uses fixed values (not formulas) and skips rows without a valid date or rate. It is only created in separate-file mode.
* Switch off in `config.py`: `ADD_RATE_CHANGE_COLUMN = False`, `ADD_RATE_SUMMARY = False`.
* The red highlight is now applied (static fill **and** live rule) only from TRADE_DATE through Brokerage, as specified; earlier the live rule coloured the whole row in Excel.

## 📁 Separate Excel file per investor (default)
Step 2 now creates **one Excel file for every investor** inside the folder `Investor_Excels` (next to the output file you choose). The file you choose (`Investor_Wise_Final.xlsx`) becomes the **index**: every investor/combination, counts, negative-brokerage counts, status and a **clickable link to that investor's file**.
* File name = investor name (Windows-invalid characters removed, reserved names like CON avoided, unique, path kept short).
* An investor with several Scheme + Folio + Transaction Type combinations gets **one file with one sheet per combination**, sheet named `Scheme-Folio-Txn` (e.g. `A1-00123-P`). One combination → sheet named after the investor.
* Name variants of the same person (case/space only) share one file; genuinely different names (e.g. RAHUL SHAH / RAHUL SHAHA) get different files.
* Inside each file everything is as before: all source columns, oldest → newest, **Brokerage** after BRKAGE_RAT, negative rows light red from TRADE_DATE to Brokerage.
* Untick the checkbox on the Step 2 / One-Click page (or set `SEPARATE_FILES_PER_INVESTOR = False` in `config.py`) to get the old single-workbook layout.
* Re-running replaces files with the same name in `Investor_Excels` (a warning is logged); old files of other names are not deleted.

## ⚡ One-Click mode
Tab **One-Click (Both Steps)**: select the CAMS Excel + an output folder (source file box is optional — blank means the same file is used) → **Run Everything**. Step 1 runs, then Step 2 starts by itself; both `Discrepancy_Result.xlsx` and `Investor_Wise_Final.xlsx` are saved in the folder. It calls the same two engines as the separate tabs. If Step 1 finds no temporary discrepancies, it stops with a clear message (Step 1's file is kept). Step 1 and Step 2 tabs still work independently.
Command line: `python main.py both "cams comm.xlsx" "C:\Output" ["source.xlsx"]`

## Front end (professional desktop UI)
* **Dashboard** (opens first): one card per task (One-Click is marked *Recommended*) and **Recent activity** — the last runs with status, time, duration and an *Open* button for their output folder.
* **Sidebar** grouped into *Workflow / Tools / Support*; a blue dot shows which task is still running. **Light / Dark / System** theme switch at the bottom.
* **Every task page has two columns, sized for a 1366 × 768 laptop** (the window opens maximised):
  * left — the files, the options and a big **Run** button with the progress bar. Under every file box a live check line says whether the file exists, its size and date, how many Excel files a folder holds, or that an existing output file will be **replaced**;
  * right — **Results** (green "Completed in …" banner, number tiles, *Open output folder / file*) and the colour-coded **Activity log** with *Copy* and *Clear*.
* Status badge on every page: *Ready → Running… → Completed / Failed*. When a task finishes, a small notification slides in at the bottom-right and the taskbar button flashes if the window is in the background; errors still show a dialog.
* Step 1's output is filled into Step 2 automatically (or press *Use Step 1 output*).
* Remembers your **theme, the last folder per task and the run history** in `%APPDATA%\CAMS Commission Automation\settings.json` (safe to delete).
* Asks before closing while a task is still running.
* **Keyboard:** `Ctrl+1` … `Ctrl+6` switch pages, `Ctrl+Enter` runs the open task.
* Own app icon (`assets/app.ico`, drawn by `app_icon.py`) on the window and taskbar. **`Create_Desktop_Shortcut.bat`** puts a "CAMS Commission Automation" icon on the desktop that opens the app directly (run `Start_Application.bat` once first).

## Project layout
```
main.py               entry point (GUI, or  step1 / step2  on the command line)
gui.py                CustomTkinter front end (dashboard, two-column task pages, notifications, themes)
app_settings.py       remembers theme / last folders / recent runs (per Windows user)
app_icon.py           draws the app icon -> assets/app.ico
gui_classic.py        plain ttk fallback, used automatically if customtkinter is missing
config.py             every threshold / limit / name in one place
discrepancy_engine.py STEP 1 (v5.0 logic, vectorised)
investor_engine.py    STEP 2 (matching, sorting, Brokerage, planning)
excel_reader.py       streaming Excel reader (openpyxl read-only or calamine) + identifier normalisation
database.py           SQLite temp store for matched rows (auto-deleted)
rate_summary.py       Rate Summary sheet: periods table + key figures
sheet_updater.py      Update Existing Sheets tool (adds Rate Change + Rate Summary to ready files)
excel_writer.py       sheet naming, workbook planning, XlsxWriter streaming writers, formats
reporting.py          thread-local log/progress hooks + UserError
tests/                86 automated tests, reference_v5.py (your original logic), benchmark.py, gui_smoke.py
```

## Step 1 — business rules (unchanged from v5.0)
Group = AMC_CODE + SCHEME_COD + FOLIO_NO + INV_NAME + TRXN_TYPE (types never mixed). Quarterly mode of BRKAGE_RAT (2 dp), quarters with < 2 valid rates ignored. **TEMPORARY**: drop < −0.2 vs previous eligible quarter **and** any later quarter recovers ≥ 80 % of the drop (permanent drops not flagged). **INTRA**: row differs from quarter mode (4 dp) by ≥ 0.2 **and** its rate is not in the group's "normal pool" (rate present in ≥ 80 % of the group's quarters or in ≥ 2 quarters). TEMPORARY wins over INTRA. Sort TRADE_DATE → AMC_CODE → TRXN_TYPE → INV_NAME. Layout, light-red `#FFCCCC`, blue headers, widths, date format unchanged.

### ⚠ Things found in the v5.0 source (please read)
1. **"Daily split" rule is only a comment.** The header says same-day rates that sum to a normal total are valid splits. That logic does **not exist** in the code; the code uses the *normal pool* rule above. I kept the code's behaviour and did **not** add a daily-sum rule. Tell me if you want it added (that would be a business-rule change).
2. **Date bug (real).** v5.0 turns every cell into text and then calls `to_datetime(dayfirst=True)`. For **real Excel date cells** this swaps day and month when day ≤ 12 (5-Jan-2024 → 1-May-2024) and can silently drop rows whose day > 12. I reproduced this with your original code. If your TRADE_DATE column holds *text* like `05-01-2024`, v5.0 was fine and this app gives identical results (tested). **Default now parses dates correctly** (real dates stay real; ISO text year-first; other text day-first). To reproduce the old behaviour exactly set `LEGACY_V5_DATE_PARSING = True` in `config.py`. Since this can change which quarter a row falls in, compare one old and one new report before relying on it.
3. **Index bug.** When rows with bad dates were dropped, v5.0 could attach INTRA flags to the wrong rows. Fixed.

## Step 2 — how it works
* Combinations come from the **Discrepancy Report** sheet of Excel A (deduplicated). **Every transaction comes from Excel B**, never from A's raw sheet; all dates, not just the discrepancy quarter.
* **Match = SCHEME_COD + FOLIO_NO + TRXN_TYPE, exact.** Identifiers are only normalised by trimming spaces / NBSP and turning `123.0` into `123`; **leading zeros are kept** (`00123` ≠ `123`). No fuzzy matching.
* **Safety separation:** rows with the same triple but a different AMC_CODE or INV_NAME (case/space-insensitive) are **not mixed in**; they are listed in the *Excluded Rows Log* sheet. A blank AMC/investor in the source cannot contradict the report, so it is assigned if only one combination fits (noted in the index). Switch with `STRICT_INVESTOR_VALIDATION` in `config.py`.
* **Sort:** real parsed dates, oldest → newest; ties keep the original source row order; rows with an invalid/missing date go to the **end** of their group with no Brokerage. Each group is calculated on its own — nothing carries across groups.
* **Brokerage = previous BRKAGE_RAT − current BRKAGE_RAT**, inserted right after BRKAGE_RAT. First row blank. A missing/invalid rate leaves that row and the next one blank (counted in the index notes). Format `+0.1361 / 0.0000 / -0.2661`. Real formulas: `=IF(AND(ISNUMBER(G2),ISNUMBER(G3)),ROUND(G2-G3,10),"")`. (`ROUND(…,10)` only removes floating-point dust.) The calculated result is also stored in the file, so the numbers and red rows show even before Excel recalculates; Excel recalculates on open.
* **Highlight:** negative Brokerage ⇒ whole range from **TRADE_DATE through Brokerage** `#FFCCCC`, as static fill **and** a conditional-format rule (`=AND(ISNUMBER($H2),$H2<0)`). If you later edit a rate by hand the rule updates automatically, but the static fill of that row does not; clear the fill in that case.
* **Worksheet names:** investor name with invalid characters `[]:*?/\` replaced, ≤ 31 chars, unique case-insensitively (`Name`, `Name (2)`, …); one investor with several combinations ⇒ separate sheets.
* **Index:** *Investor Sheet Index* (all requested columns + *Workbook File* + *Notes*; "No matching transactions found" for unmatched groups). Extra sheets only when needed: *Excluded Rows Log*, *File Index*.
* **Limits:** a group > 1,048,575 rows continues on `Name p2`, `Name p3` … (first Brokerage of a continuation sheet is a static number). Workbooks are split deterministically (groups packed in sorted order, never split across files) at `MAX_SHEETS_PER_WORKBOOK = 200` or `MAX_CELLS_PER_WORKBOOK = 30,000,000` → `Investor_Wise_Final.xlsx`, `_Part2.xlsx`, … Each file carries the full index.
* **Errors handled with readable messages:** missing columns/sheets, empty or corrupt file, output open in Excel, same input/output, low disk space. Temporary files (SQLite DB, XlsxWriter temp files) are removed after success **and** failure.

## Memory / speed design
Streaming read (openpyxl read-only, or calamine for files ≤ 60 MB — `READER` / `CALAMINE_MAX_FILE_MB` in `config.py`), only needed columns for Step 1, single pass over the source for Step 2, matched rows to SQLite on disk, SQL-indexed sorting, XlsxWriter constant-memory output. No pandas for Step 2.

## Test status (what was actually run)
Run in my sandbox (Linux, Python 3.12, pandas 3.0.2, numpy 2.4.4, XlsxWriter 3.2.9, openpyxl 3.1.5, calamine):
* **86 automated tests pass.** Step 1 equals your original v5.0 code cell-for-cell on random data (legacy-date mode) and, with default settings, on text-date data. Step 2 tests cover every case in your list 1–19.
* Brokerage formulas were **recalculated by LibreOffice** (cached results removed first) and match.
* GUI smoke test (both tabs, success + error path) passed under a virtual display. The redesigned UI's smoke test (all four task pages + error path) and the 86 tests were re-run on Windows 10 / Python 3.14 / customtkinter 6.0.
* Large file: 1.05 M rows × 9 cols — Step 1 ≈ 96 s, Step 2 ≈ 69 s (fast reader, ≈ 1.0 GB peak). 300 k rows × 27 cols — streaming reader: Step 1 112 s / Step 2 113 s, **269 MB peak**; fast reader: 37 s / 28 s, ≈ 780 MB peak.

**Not tested:** real Windows, Microsoft Excel itself (only LibreOffice), pandas 2.x, files with 100+ columns × 1 M rows, Excel files larger than a few hundred MB. Timings will differ on your PC.
