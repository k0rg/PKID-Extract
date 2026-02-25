# Copilot Instructions for PKID-Extract

## Big picture
- Single-script **tkinter GUI application** in `extract_PKID.py`, **Windows-only**.
- `oa3tool.exe` is **bundled in the repo root** — no user download required.
- User browses for an input CSV or XLSX; the app decodes HW hashes via the bundled tool and writes an output CSV with `ProductKeyID`.
- External boundary: `oa3tool.exe` is the only decoding mechanism — no internal HW hash logic.

## Architecture – `extract_PKID.py`
- **`PKID_Extract.bat`** – Windows launcher; checks for Python via `where`, offers `winget install` if missing, then runs `extract_PKID.py`.
- **Prerequisite checks** (`_check_python_version`, `_check_tkinter`) run before any tkinter import; exit with actionable messages on failure.
- **`_BUNDLED_OA3TOOL`** – path resolved relative to the script's own directory (`os.path.dirname(os.path.abspath(__file__))`).
- **`PKIDExtractApp`** (tkinter.Tk subclass) – file browse dialog (CSV/XLSX), column-mapping dropdowns, progress bar, log pane, "Open File" button.
- **`_ensure_openpyxl()`** – prompts the user to pip-install `openpyxl` on first XLSX use if not already installed.
- **`_read_xlsx_rows()` / `_read_csv_rows()`** – normalise both formats into `(headers, list[dict])`. XLSX files that are actually HTML, old .xls, or plain CSV are detected via `_sniff_real_format()` and handled gracefully.
- **`auto_map_columns()`** – matches file headers against known aliases (`SERIAL_ALIASES`, `HWHASH_ALIASES`) so the user doesn't need exact column names.
- **`run_extraction()`** – pure logic: reads CSV/XLSX, calls `subprocess.run([tool, "/decodehwhash:<hash>"])`, parses stdout with regex `<p n="ProductKeyID" v="(\d+)" />`, writes output CSV. Accepts `on_progress` / `on_log` callbacks.
- Processing runs in a **daemon thread**; GUI updates are marshalled via `self.after()`.

## Data flow
1. User selects input CSV or XLSX → headers loaded → `auto_map_columns()` pre-selects dropdowns.
2. Output path auto-suggested as `<input_base>_output.csv`; log file as `<output_base>_log.txt`.
3. Per-row: `subprocess.run` → regex match → append dict → progress callback.
4. Output schema is always `SerialNumber, HWHash, ProductKeyID` regardless of input column names.

## Project-specific conventions
- **Single file** – keep all logic in `extract_PKID.py`; no package/framework structure.
- **`utf-8-sig` encoding** for every CSV and log read/write (BOM compatibility with Excel).
- **Column names are canonical**: output always uses `SerialNumber`, `HWHash`, `ProductKeyID`.
- **Subprocess safety**: argument list, never shell strings.
- **Logging**: append-mode plain-text lines; also echoed to the GUI log pane.
- **Thread safety**: worker thread must only update the GUI through `self.after()`.

## Key alias lists (column auto-mapping)
- `SERIAL_ALIASES`: `serialnumber`, `serial_number`, `sn`, `device serial number`, …
- `HWHASH_ALIASES`: `hwhash`, `hw_hash`, `hardware hash`, `hash`, …
- To add new aliases, append to the list at the top of the file.

## Environment notes
- **Windows-only** — uses `os.startfile()`, assumes `.exe` tool availability. No Linux/macOS compat needed.
- `oa3tool.exe` must sit beside `extract_PKID.py` in the repo root.
- Only stdlib is used (`tkinter`, `csv`, `subprocess`, `re`, `threading`). `openpyxl` is the sole optional dependency, auto-installed via pip when a user first opens an XLSX file.

## Primary references
- `PKID_Extract.bat` – recommended entry point; handles Python installation
- `extract_PKID.py` – all runtime logic and GUI
- `oa3tool.exe` – bundled decoding tool (do not remove)
- `README.md` – user-facing usage guide and input/output examples
