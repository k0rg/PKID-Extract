
# PKID Extract Tool

A Windows GUI application that extracts Product Key IDs from hardware hashes using Microsoft's `oa3tool.exe`. Browse for your CSV, let the tool auto-detect columns, and get results in one click.

## Prerequisites

| Requirement | Details |
|---|---|
| **Python 3.8+** | [python.org/downloads](https://www.python.org/downloads/) – ensure **"Add Python to PATH"** is checked during install. The script checks your version on launch and will tell you if an upgrade is needed. |
| **tkinter** | Included with the standard Python installer on Windows. If missing, the script exits with instructions to re-run the installer and enable "tcl/tk and IDLE". |

> **Note:** `oa3tool.exe` is **bundled in this repository** — no separate download or ADK installation is required. The script automatically uses the copy located next to `extract_PKID.py`. No third-party pip packages are needed.

## Preparing Your Input CSV

The input CSV must contain at least two columns representing a serial number and a hardware hash. The column names do **not** need to match exactly — the tool automatically recognises common variations:

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

   Double-click `extract_PKID.py`, or run from a terminal:

   ```bash
   python extract_PKID.py
   ```

2. **Select Input CSV** — click **Browse…** and choose your CSV file. The tool will:
   - Read the CSV headers.
   - Auto-map serial number and hardware hash columns (shown in the Column Mapping section).
   - If auto-detection fails, select the correct columns manually from the dropdowns.

3. **Click ▶ Process** — a progress bar tracks each row. Warnings and errors appear in the log pane at the bottom of the window.

4. **View results** — when processing completes, the output file path is displayed with an **Open File** button to launch it directly.

### Output Files

Both files are created automatically in the **same directory** as the input CSV:

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
| **Columns not auto-detected** | Your CSV headers don't match any known alias. Select the correct columns manually from the dropdowns. |
| **Empty or missing output** | Check the log pane / log file for per-row errors. Ensure hardware hashes are valid and properly formatted. |
| **`tkinter` not available** | Re-run the Python installer, click "Modify", and ensure "tcl/tk and IDLE" is checked. |

## License

This project is provided under the [MIT License](LICENSE).
