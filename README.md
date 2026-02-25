
# PKID Extract Tool

A Windows GUI application that extracts Product Key IDs from hardware hashes using Microsoft's `oa3tool.exe`. Browse for your CSV or Excel file, let the tool auto-detect columns, and get results in one click.

## Prerequisites

### Option A – Standalone executable (recommended)

Download **`PKID_Extract.exe`** from the [latest GitHub Release](../../releases/latest). No Python installation required — just double-click the `.exe` to launch the GUI. `oa3tool.exe` is already bundled inside.

### Option B – Run from source

If you prefer to run the Python script directly:

| Requirement | Details |
|---|---|
| **Python 3.8+** | If Python is not installed, the included `PKID_Extract.bat` launcher will detect this and offer to **install it automatically via `winget`**. You can also install manually from [python.org/downloads](https://www.python.org/downloads/) – ensure **"Add Python to PATH"** is checked. |
| **tkinter** | Included with the standard Python installer on Windows. If missing, the script exits with instructions to re-run the installer and enable "tcl/tk and IDLE". |

> **Note:** `oa3tool.exe` is **bundled in this repository** — no separate download or ADK installation is required. If you open an `.xlsx` file, the tool will prompt to install `openpyxl` automatically via pip (one-time).

## Preparing Your Input File

The input file can be a **CSV** (`.csv`) or **Excel** (`.xlsx`) file. It must contain at least two columns representing a serial number and a hardware hash. The column names do **not** need to match exactly — the tool automatically recognises common variations:

| Canonical Name | Recognised Column Headers |
|---|---|
| SerialNumber | `SerialNumber`, `Serial_Number`, `Serial Number`, `SN`, `Device Serial Number`, … |
| HWHash | `HWHash`, `HW_Hash`, `Hardware Hash`, `Hash`, … |

If your column names are not auto-detected, you can select them manually from the dropdowns.

Example `input.csv`:

```csv
Device Serial Number,Hardware Hash
0F3XXXXXXXGT,ABC123XYZ456...
A12B34C56D78E9,DEF456UVW123...
```

## How to Use

1. **Launch the tool**

   **Standalone `.exe`** — Double-click **`PKID_Extract.exe`** downloaded from the [Releases](../../releases/latest) page. That's it.

   **From source** — Double-click **`PKID_Extract.bat`**, which will:
   - Check if Python is installed.
   - If not, offer to install it automatically via `winget`.
   - Launch the GUI once Python is available.

   Or run directly if Python is already installed:

   ```bash
   python extract_PKID.py
   ```

2. **Select Input File** — click **Browse…** and choose your `.csv` or `.xlsx` file. The tool will:
   - If an `.xlsx` file is selected and `openpyxl` is not installed, offer to install it automatically.
   - Read the file headers.
   - Auto-map serial number and hardware hash columns (shown in the Column Mapping section).
   - If auto-detection fails, select the correct columns manually from the dropdowns.

3. **Click ▶ Process** — a progress bar tracks each row. Warnings and errors appear in the log pane at the bottom of the window.

4. **View results** — when processing completes, the output file path is displayed with an **Open File** button to launch it directly.

### Output Files

Both files are created automatically in the **same directory** as the input file:

| File | Naming Convention | Contents |
|---|---|---|
| Output CSV | `<input_name>_output.csv` | `SerialNumber`, `HWHash`, `ProductKeyID` |
| Log file | `<input_name>_output_log.txt` | Warnings and errors (only created if issues occur) |

Example output CSV:

```csv
SerialNumber,HWHash,ProductKeyID
0F3XXXXXXXXXGT,ABC123XYZ456,3XXXXXXXXXXX6
A12B34C56D78E9,DEF456UVW123,1234567890123
```

Example log:

```
Warning: ProductKeyID not found for Serial Number: 0F3XXXXXXXXXGT
Error: Failed to run oa3tool for Serial Number: A12B34C56D78E9 - <error details>
```

## Troubleshooting

| Problem | Solution |
|---|---|
| **"oa3tool.exe Missing"** | Ensure `oa3tool.exe` is in the same folder as `extract_PKID.py`. Re-download the repository if needed. |
| **Columns not auto-detected** | Your file headers don't match any known alias. Select the correct columns manually from the dropdowns. |
| **Empty or missing output** | Check the log pane / log file for per-row errors. Ensure hardware hashes are valid and properly formatted. |
| **`openpyxl` install fails** | Install manually with `pip install openpyxl`, then re-launch the tool. |
| **`tkinter` not available** | Re-run the Python installer, click "Modify", and ensure "tcl/tk and IDLE" is checked. |

## License

This project is provided under the [MIT License](LICENSE).
