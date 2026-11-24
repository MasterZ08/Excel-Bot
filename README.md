# Random Excel Data Generator

A Windows 10/11 desktop app (Python 3.11+, CustomTkinter) that **continuously generates synthetic Excel (.xlsx) and CSV files** at a set interval until you press STOP or close the window. All data is fictional (e-mails use reserved `example.com/.org/.net` domains; phone numbers use fictional ranges where Faker is not used).

## Install & run
1. Install Python 3.11+ from python.org (tick **Add python.exe to PATH**).
2. Double-click **`run.bat`** — it creates a `.venv`, installs `requirements.txt` on first run, and starts the app.

Manual: `pip install -r requirements.txt` then `python main.py`.

## Using it
| Control | What it does |
|---|---|
| **START** | Generates a file immediately, saves it, waits *Delay* seconds, repeats **with no file limit**. (If *Continuous generation* is off, START makes exactly one file.) |
| **PAUSE / RESUME** | Pause freezes the countdown; a file already being written finishes first. |
| **STOP** | Finishes the current file (if any), cancels the countdown instantly, no further cycles. Click again while it says *FORCE STOP* to abandon an in-progress huge file (no partial file is left). |
| **Generate One File** | Exactly one file, then idle. |
| **Randomize Again** | New dataset right now with current settings (while running: skips the countdown). |

Defaults: continuous ON, 100 records, 5 s delay (allowed 1–3600 s), unlimited files, XLSX.

**Settings:** output folder, filename prefix, category, XLSX/CSV/Both, records (1–1,000,000), delay, ID style (random/sequential) + prefix/suffix, fixed seed, duplicate records (%), date range, amount min/max (blank = category default; applies to salary / unit price / transaction amount), age range, Faker locale, currency symbol, summary sheets, dark/light/system theme. Settings, history and sequential-ID counters persist between launches (`configs/`).

**Fixed seed:** each START is reproducible, but every file within a run still differs (seed + file number).

### Categories
Personal Information · Employees · Business/Sales · Financial (running balances, never overdrawn) · Inventory · Education · **Custom Dataset**.

### Custom Dataset tab
Add / duplicate / remove / reorder / rename columns; types: text, integer, decimal, name, email, phone, date, ID, boolean, address, currency, custom list; per-column ranges, decimals, ID prefix/suffix/mode, %-true, list values. **Save config… / Load config…** store settings + columns as JSON (loading keeps your current output folder and theme). See `configs/example_support_tickets.json`.

### Output files
`<prefix>_<seq 5 digits>_<YYYYMMDD_HHMMSS>.xlsx|csv` — files are created with exclusive-create, so **nothing is ever overwritten**; the sequence continues after restarts.
XLSX: sheets **Data** (Excel table, filters, frozen bold coloured header, banded rows, auto widths, date/currency/number formats), **Summary** (live `SUM/AVERAGE/MIN/MAX/COUNTA` formulas) and **Breakdown** (counts per category). CSV is UTF-8 with BOM (opens correctly in Excel).

### Preview / History / Log tabs
Preview shows 20 records (refresh = new values, never saved). History lists every file (double-click opens it), can be exported to CSV or cleared, and survives restarts. Errors are logged to `logs/app.log`; the run continues after recoverable errors, and pauses itself after 10 failed files in a row or on an invalid setting.

## Layout
```
main.py  gui.py  generators.py  excel_exporter.py  csv_exporter.py
config_manager.py  history_manager.py  worker.py  utils.py
requirements.txt  run.bat  build_exe.bat  make_examples.py
output/  configs/  logs/  examples/  tests/
```

## Tests
`python -m unittest discover -s tests -v` — 31 headless tests: every category, differing data per cycle, seeds, IDs, ranges, XLSX/CSV validity, no-overwrite, continuous/pause/resume/stop timing, no file limit, error recovery, abort without partial files, 100,000-record export (~7 s for XLSX+CSV).

## Package as an EXE (PyInstaller)
Double-click **`build_exe.bat`**, or:
```
pip install -r requirements-build.txt
pyinstaller --noconfirm --clean --onefile --windowed --name RandomExcelDataGenerator --collect-all customtkinter --collect-all faker main.py
```
Result: `dist\RandomExcelDataGenerator.exe`. `output\`, `configs\`, `logs\` are created beside the EXE.

## Notes
* Without Faker the app falls back to a built-in English generator (shown in the sidebar); install Faker for locale-specific names/addresses.
* Practical size: 100k–250k rows is quick; 1,000,000 rows needs several GB of RAM.
* `examples/` was produced by `make_examples.py`.
"# Excel-Bot" 
