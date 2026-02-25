"""
PKID Extract Tool – GUI Application
Extracts Product Key IDs from hardware hashes using oa3tool.exe.
Provides a Windows-friendly GUI: browse for files, auto-map CSV columns,
view progress, and inspect logs without touching the command line.
"""

import csv
import subprocess
import re
import os
import sys
import threading
import tempfile

# ---------------------------------------------------------------------------
# Prerequisite checks (run before any tkinter import)
# ---------------------------------------------------------------------------

MIN_PYTHON = (3, 8)


def _check_python_version():
    if sys.version_info < MIN_PYTHON:
        print(
            f"Error: Python {MIN_PYTHON[0]}.{MIN_PYTHON[1]}+ is required "
            f"(you have {sys.version_info.major}.{sys.version_info.minor}).\n"
            "Download the latest version from https://www.python.org/downloads/"
        )
        sys.exit(1)


def _check_tkinter():
    """Verify tkinter is available and give actionable install guidance if not."""
    try:
        import tkinter  # noqa: F401
    except ImportError:
        print(
            "Error: tkinter is not installed.\n\n"
            "tkinter is required for the graphical interface.\n"
            "Re-run the Python installer, click 'Modify', and ensure\n"
            "'tcl/tk and IDLE' is checked.\n"
            "https://www.python.org/downloads/"
        )
        sys.exit(1)


_check_python_version()
_check_tkinter()

import tkinter as tk  # noqa: E402  (deferred until after check)
from tkinter import ttk, filedialog, messagebox  # noqa: E402


# ---------------------------------------------------------------------------
# openpyxl – lazy import with auto-install prompt
# ---------------------------------------------------------------------------

def _ensure_openpyxl():
    """Return the openpyxl module, installing it first if necessary.

    Prompts the user for permission via a messagebox before installing.
    Returns None if the user declines or installation fails.
    """
    try:
        import openpyxl
        return openpyxl
    except ImportError:
        pass

    install = messagebox.askyesno(
        "Install Required Package",
        "Reading .xlsx files requires the 'openpyxl' package,\n"
        "which is not currently installed.\n\n"
        "Would you like to install it now?\n"
        "(runs: pip install openpyxl)",
    )
    if not install:
        return None

    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "openpyxl"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except Exception as exc:
        messagebox.showerror(
            "Installation Failed",
            f"Could not install openpyxl:\n{exc}\n\n"
            "Install it manually with:  pip install openpyxl",
        )
        return None

    try:
        import openpyxl
        return openpyxl
    except ImportError:
        messagebox.showerror(
            "Import Failed",
            "openpyxl was installed but could not be imported.\n"
            "Try restarting the application.",
        )
        return None


def _read_xlsx_rows(path: str) -> tuple[list[str], list[dict[str, str]]]:
    """Read an .xlsx file and return (headers, rows_as_dicts).

    All cell values are converted to strings to match CSV behaviour.
    Raises ImportError if openpyxl is not available.
    """
    import openpyxl  # caller must ensure this is installed
    from zipfile import BadZipFile

    # -- Detect files with .xlsx extension that aren't real OOXML ------------
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except BadZipFile:
        # The file is not a valid zip/xlsx.  Try to figure out what it
        # actually is so we can give a helpful message (or just read it).
        actual = _sniff_real_format(path)
        if actual == "csv":
            # It's really a CSV/TSV disguised as .xlsx – fall back silently
            return _read_csv_rows(path)
        elif actual == "html":
            raise ValueError(
                "This .xlsx file is actually an HTML document (some export tools do this).\n"
                "Open it in Excel, then use File \u2192 Save As \u2192 'CSV (Comma delimited) (*.csv)'."
            )
        elif actual == "xls":
            raise ValueError(
                "This file appears to be an older .xls (Excel 97-2003) format.\n"
                "Open it in Excel, then use File \u2192 Save As \u2192 'CSV (Comma delimited) (*.csv)'."
            )
        else:
            raise ValueError(
                "This .xlsx file could not be opened (it is not a valid Excel workbook).\n"
                "Open it in Excel, then use File \u2192 Save As \u2192 'CSV (Comma delimited) (*.csv)'."
            )

    ws = wb.active
    row_iter = ws.iter_rows(values_only=True)

    # First row = headers
    raw_headers = next(row_iter, None)
    if not raw_headers:
        wb.close()
        return [], []

    headers = [str(h).strip() if h is not None else "" for h in raw_headers]
    rows: list[dict[str, str]] = []
    for raw in row_iter:
        row_dict = {}
        for col_name, value in zip(headers, raw):
            row_dict[col_name] = str(value).strip() if value is not None else ""
        rows.append(row_dict)

    wb.close()
    return headers, rows


def _sniff_real_format(path: str) -> str:
    """Peek at the first bytes of a file to guess its actual format."""
    try:
        with open(path, "rb") as f:
            head = f.read(512)
    except Exception:
        return "unknown"

    # Old-style .xls (Compound File Binary / OLE2)
    if head[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "xls"

    # HTML (some tools export a table as .xlsx)
    text = head.decode("utf-8", errors="ignore").strip().lower()
    if text.startswith("<!doctype") or text.startswith("<html") or "<table" in text[:256]:
        return "html"

    # Likely plain-text / CSV / TSV
    if all(b < 128 for b in head[:256]) or head[:3] == b"\xef\xbb\xbf":
        return "csv"

    return "unknown"


def _read_csv_rows(path: str) -> tuple[list[str], list[dict[str, str]]]:
    """Fallback: read a plain-text CSV/TSV and return (headers, rows_as_dicts)."""
    with open(path, mode="r", newline="", encoding="utf-8-sig") as f:
        # Sniff the dialect to handle tabs, semicolons, etc.
        sample = f.read(8192)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        except csv.Error:
            dialect = "excel"  # fall back to default comma-separated
        reader = csv.DictReader(f, dialect=dialect)
        headers = list(reader.fieldnames or [])
        rows = [row for row in reader]
    return headers, rows


def _is_xlsx(path: str) -> bool:
    return path.lower().endswith(".xlsx")

# ---------------------------------------------------------------------------
# Column auto-mapping: common alternative names → canonical column names
# ---------------------------------------------------------------------------
SERIAL_ALIASES = [
    "serialnumber", "serial_number", "serial number", "sn",
    "serial", "device serial", "device serial number", "device_serial_number",
]
HWHASH_ALIASES = [
    "hwhash", "hw_hash", "hw hash", "hardwarehash", "hardware_hash",
    "hardware hash", "hash",
]


def _normalise(name: str) -> str:
    """Lower-case, strip whitespace/BOM for comparison."""
    return name.strip().strip("\ufeff").lower()


def auto_map_columns(headers: list[str]) -> dict[str, str | None]:
    """Return {'SerialNumber': <matched_header>, 'HWHash': <matched_header>}.

    Values are None when no match is found.
    """
    mapping: dict[str, str | None] = {"SerialNumber": None, "HWHash": None}
    for h in headers:
        norm = _normalise(h)
        if norm in SERIAL_ALIASES:
            mapping["SerialNumber"] = h
        elif norm in HWHASH_ALIASES:
            mapping["HWHash"] = h
    return mapping


# ---------------------------------------------------------------------------
# Core extraction logic (unchanged subprocess + regex approach)
# ---------------------------------------------------------------------------


def run_extraction(
    tool_path: str,
    input_file: str,
    output_file: str,
    log_file: str,
    serial_col: str,
    hash_col: str,
    on_progress=None,
    on_log=None,
):
    """Process the input CSV and write results.  Callbacks update the GUI."""
    output_data: list[dict[str, str]] = []

    with open(input_file, mode="r", newline="", encoding="utf-8-sig") as infile:
        reader = csv.DictReader(infile)
        rows = list(reader)

    total = len(rows)
    if total == 0:
        if on_log:
            on_log("Error: Input CSV contains no data rows.")
        return False

    for i, row in enumerate(rows):
        serial_number = row[serial_col]
        hw_hash = row[hash_col]

        product_key_id = "Not Found"

        try:
            result = subprocess.run(
                [tool_path, f"/decodehwhash:{hw_hash}"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )

            if result.returncode == 0:
                matches = re.findall(
                    r'<p n="ProductKeyID" v="(\d+)" />', result.stdout
                )
                if matches:
                    product_key_id = matches[0]
                else:
                    msg = f"Warning: ProductKeyID not found for Serial Number: {serial_number}"
                    _append_log(log_file, msg)
                    if on_log:
                        on_log(msg)
            else:
                msg = f"Error: Failed to run oa3tool for Serial Number: {serial_number} - {result.stderr}"
                _append_log(log_file, msg)
                if on_log:
                    on_log(msg)
        except FileNotFoundError:
            msg = f"Error: oa3tool.exe not found at '{tool_path}'. Aborting."
            _append_log(log_file, msg)
            if on_log:
                on_log(msg)
            return False
        except Exception as exc:
            msg = f"Error: Unexpected failure for Serial Number: {serial_number} - {exc}"
            _append_log(log_file, msg)
            if on_log:
                on_log(msg)

        output_data.append(
            {
                "SerialNumber": serial_number,
                "HWHash": hw_hash,
                "ProductKeyID": product_key_id,
            }
        )

        if on_progress:
            on_progress(i + 1, total)

    # Write results
    with open(output_file, mode="w", newline="", encoding="utf-8-sig") as outfile:
        writer = csv.DictWriter(
            outfile, fieldnames=["SerialNumber", "HWHash", "ProductKeyID"]
        )
        writer.writeheader()
        writer.writerows(output_data)

    if on_log:
        on_log(f"Processing complete – {total} rows written to {output_file}")
    return True


def _append_log(log_path: str, message: str):
    with open(log_path, "a", encoding="utf-8-sig") as f:
        f.write(message + "\n")


# ---------------------------------------------------------------------------
# Bundled oa3tool.exe path
# ---------------------------------------------------------------------------

# Resolve relative to the script's own directory so it works regardless of cwd.
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BUNDLED_OA3TOOL = os.path.join(_SCRIPT_DIR, "oa3tool.exe")


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------


class PKIDExtractApp(tk.Tk):
    """Main application window."""

    def __init__(self):
        super().__init__()
        self.title("PKID Extract Tool")
        self.resizable(False, False)
        self._build_ui()
        self._csv_headers: list[str] = []
        self._processing = False
        self._detect_oa3tool()

    # ---- UI construction ---------------------------------------------------

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # ── Input file frame ─────────────────────────────────────────
        file_frame = ttk.LabelFrame(self, text="Input", padding=8)
        file_frame.grid(row=0, column=0, sticky="ew", **pad)

        ttk.Label(file_frame, text="File:").grid(row=0, column=0, sticky="w")
        self.input_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.input_var, width=60).grid(
            row=0, column=1, sticky="ew", padx=(4, 0)
        )
        ttk.Button(file_frame, text="Browse…", command=self._browse_input).grid(
            row=0, column=2, padx=(4, 0)
        )

        file_frame.columnconfigure(1, weight=1)

        # ── Column mapping frame ────────────────────────────────────────
        col_frame = ttk.LabelFrame(self, text="Column Mapping", padding=8)
        col_frame.grid(row=1, column=0, sticky="ew", **pad)

        ttk.Label(col_frame, text="SerialNumber column:").grid(
            row=0, column=0, sticky="w"
        )
        self.serial_combo = ttk.Combobox(col_frame, state="readonly", width=30)
        self.serial_combo.grid(row=0, column=1, sticky="w", padx=(4, 0))

        ttk.Label(col_frame, text="HWHash column:").grid(
            row=1, column=0, sticky="w"
        )
        self.hash_combo = ttk.Combobox(col_frame, state="readonly", width=30)
        self.hash_combo.grid(row=1, column=1, sticky="w", padx=(4, 0))

        self.map_label = ttk.Label(col_frame, text="Load an input CSV to detect columns.", foreground="gray")
        self.map_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 0))

        # ── Progress & action ────────────────────────────────────────────
        action_frame = ttk.Frame(self, padding=8)
        action_frame.grid(row=2, column=0, sticky="ew", **pad)

        self.progress = ttk.Progressbar(action_frame, length=400, mode="determinate")
        self.progress.grid(row=0, column=0, sticky="ew")
        self.progress_label = ttk.Label(action_frame, text="0 / 0")
        self.progress_label.grid(row=0, column=1, padx=(8, 0))

        self.run_btn = ttk.Button(
            action_frame, text="▶  Process", command=self._start_processing
        )
        self.run_btn.grid(row=1, column=0, columnspan=2, pady=(8, 0))

        # ── Output result row (hidden until processing completes) ────────
        self.output_frame = ttk.Frame(action_frame)
        # not gridded yet – shown by _processing_done

        self.output_path_var = tk.StringVar()
        ttk.Label(self.output_frame, text="Output:").grid(row=0, column=0, sticky="w")
        ttk.Entry(
            self.output_frame, textvariable=self.output_path_var,
            state="readonly", width=52,
        ).grid(row=0, column=1, sticky="ew", padx=(4, 0))
        self.open_btn = ttk.Button(
            self.output_frame, text="Open File", command=self._open_output_file
        )
        self.open_btn.grid(row=0, column=2, padx=(4, 0))
        self.output_frame.columnconfigure(1, weight=1)

        action_frame.columnconfigure(0, weight=1)

        # ── Log display ──────────────────────────────────────────────────
        log_frame = ttk.LabelFrame(self, text="Log", padding=8)
        log_frame.grid(row=3, column=0, sticky="nsew", **pad)

        self.log_text = tk.Text(log_frame, height=10, width=80, state="disabled", wrap="word")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

    # ---- oa3tool detection --------------------------------------------------

    def _detect_oa3tool(self):
        """Verify the bundled oa3tool.exe is present."""
        if os.path.isfile(_BUNDLED_OA3TOOL):
            self._log(f"oa3tool.exe found: {_BUNDLED_OA3TOOL}")
        else:
            self._log(
                "WARNING: oa3tool.exe not found next to this script.\n"
                f"Expected location: {_BUNDLED_OA3TOOL}\n"
                "Please place oa3tool.exe in the same folder as extract_PKID.py."
            )

    # ---- Browse helpers ----------------------------------------------------

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Select Input File",
            filetypes=[
                ("Supported files", "*.csv *.xlsx"),
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx"),
                ("All files", "*.*"),
            ],
        )
        if path:
            if _is_xlsx(path):
                openpyxl = _ensure_openpyxl()
                if openpyxl is None:
                    self._log("Cancelled: openpyxl is required to read .xlsx files.")
                    return
            self.input_var.set(path)
            self._load_headers(path)

    # ---- Column detection & mapping ----------------------------------------

    def _load_headers(self, path: str):
        """Read headers from the chosen file, populate dropdowns, auto-map."""
        try:
            if _is_xlsx(path):
                headers, _rows = _read_xlsx_rows(path)
                self._csv_headers = headers
            else:
                with open(path, mode="r", newline="", encoding="utf-8-sig") as f:
                    reader = csv.DictReader(f)
                    self._csv_headers = list(reader.fieldnames or [])
        except Exception as exc:
            self._log(f"Error reading file headers: {exc}")
            return

        if not self._csv_headers:
            self._log("Warning: CSV appears to have no header row.")
            return

        self.serial_combo["values"] = self._csv_headers
        self.hash_combo["values"] = self._csv_headers

        mapping = auto_map_columns(self._csv_headers)
        auto_mapped: list[str] = []

        if mapping["SerialNumber"]:
            self.serial_combo.set(mapping["SerialNumber"])
            auto_mapped.append(f'SerialNumber → "{mapping["SerialNumber"]}"')
        else:
            self.serial_combo.set("")

        if mapping["HWHash"]:
            self.hash_combo.set(mapping["HWHash"])
            auto_mapped.append(f'HWHash → "{mapping["HWHash"]}"')
        else:
            self.hash_combo.set("")

        if auto_mapped:
            info = "Auto-mapped: " + ", ".join(auto_mapped)
            self.map_label.config(text=info, foreground="green")
            self._log(info)
        else:
            self.map_label.config(
                text="Could not auto-detect columns – please select manually.",
                foreground="orange",
            )
            self._log("Columns not auto-detected. Please map them manually above.")

    # ---- Processing --------------------------------------------------------

    def _validate(self) -> bool:
        if not os.path.isfile(_BUNDLED_OA3TOOL):
            messagebox.showerror(
                "oa3tool.exe Missing",
                f"oa3tool.exe was not found at:\n{_BUNDLED_OA3TOOL}\n\n"
                "Place oa3tool.exe in the same folder as this script.",
            )
            return False
        inp = self.input_var.get().strip()
        if not inp:
            messagebox.showwarning("Missing Path", "Please select an input file.")
            return False
        if not os.path.isfile(inp):
            messagebox.showerror(
                "File Not Found",
                f"The selected input file does not exist:\n{inp}",
            )
            return False
        if not self.serial_combo.get():
            messagebox.showwarning("Column Mapping", "Please select the SerialNumber column.")
            return False
        if not self.hash_combo.get():
            messagebox.showwarning("Column Mapping", "Please select the HWHash column.")
            return False
        return True

    def _start_processing(self):
        if self._processing:
            return
        if not self._validate():
            return

        self._processing = True
        self.run_btn.config(state="disabled")
        self.output_frame.grid_forget()  # hide previous output row
        self.progress["value"] = 0
        self.progress_label.config(text="0 / 0")

        # Auto-derive output & log paths next to the input file
        inp = self.input_var.get().strip()
        directory = os.path.dirname(inp)
        base = os.path.splitext(os.path.basename(inp))[0]
        out = os.path.join(directory, f"{base}_output.csv")
        log_path = os.path.join(directory, f"{base}_output_log.txt")

        thread = threading.Thread(
            target=self._run_in_thread,
            args=(
                _BUNDLED_OA3TOOL,
                inp,
                out,
                log_path,
                self.serial_combo.get(),
                self.hash_combo.get(),
            ),
            daemon=True,
        )
        thread.start()

    def _run_in_thread(self, tool, inp, out, log_path, serial_col, hash_col):
        try:
            success = run_extraction(
                tool_path=tool,
                input_file=inp,
                output_file=out,
                log_file=log_path,
                serial_col=serial_col,
                hash_col=hash_col,
                on_progress=self._on_progress,
                on_log=self._log_threadsafe,
            )
        except Exception as exc:
            self._log_threadsafe(f"Error: Unexpected failure – {exc}")
            success = False
        self.after(0, self._processing_done, success, out)

    def _on_progress(self, current: int, total: int):
        """Called from the worker thread; schedules GUI update."""
        self.after(0, self._update_progress, current, total)

    def _update_progress(self, current: int, total: int):
        pct = current / total * 100 if total else 0
        self.progress["value"] = pct
        self.progress_label.config(text=f"{current} / {total}")

    def _processing_done(self, success: bool, output_path: str):
        self._processing = False
        self.run_btn.config(state="normal")
        if success:
            self.output_path_var.set(output_path)
            self.output_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
            self._log(f"Results saved to: {output_path}")
        else:
            messagebox.showerror(
                "Failed",
                "Processing stopped due to errors. Check the log for details.",
            )

    def _open_output_file(self):
        """Open the output CSV with the default Windows application."""
        path = self.output_path_var.get()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("File Not Found", "Output file does not exist.")
            return
        try:
            os.startfile(path)
        except OSError as exc:
            messagebox.showerror("Open Failed", f"Could not open file:\n{exc}")

    # ---- Logging -----------------------------------------------------------

    def _log(self, message: str):
        self.log_text.config(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _log_threadsafe(self, message: str):
        self.after(0, self._log, message)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = PKIDExtractApp()
    app.mainloop()
