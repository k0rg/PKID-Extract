
# Product Key ID Extraction Script

This script extracts Product Key IDs for hardware hashes using the `oa3tool.exe`. It processes a CSV file containing serial numbers and hardware hashes, decodes the hardware hashes, and writes the extracted Product Key IDs to a new CSV file. If any errors or missing ProductKeyID values are encountered, they will be logged into a separate (txt) log file.

## Features
- Reads serial numbers and hardware hashes from an input CSV.
- Runs the [oa3tool](https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/oa3-command-line-config-file-syntax?view=windows-11) to decode the hardware hash and extract Product Key IDs.
- Displays a progress bar while processing the entries.
- Logs any errors or warnings to a separate log file.
- Saves the output in a new CSV file.

## Requirements
- Python 3.x
- [oa3tool](https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/oa3-command-line-config-file-syntax?view=windows-11) (OEM Activation Tool)
- Dependencies:
  - `re` (Regex for product key extraction)
  - `subprocess` (for executing oa3tool)

## Script Usage

1. **Download and Install oa3tool**: 
   - Ensure `oa3tool.exe` is available on your local machine.
   - More information about oa3tool.exe can be found here: [oa3tool](https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/oa3-command-line-config-file-syntax?view=windows-11)
   - The OA 3.0 tool (OA3Tool.exe) is part of the Windows Assessment and Deployment Kit (Windows ADK). For installation instructions, see [Installing the Windows ADK](https://learn.microsoft.com/en-us/previous-versions/windows/hh825494(v=win.10))

2. **Prepare the Input CSV**:
   - The input CSV must include two columns: `SerialNumber` and `HWHash`. 
   - Example of `input.csv` format:
   ```csv
   SerialNumber,HWHash
   0F3XXXXXXXGT,ABC123XYZ456...
   A12B34C56D78E9,DEF456UVW123...
   ```

3. **Script Configuration**:
   - Modify the script variables to point to the correct file paths for:
     - `tool_path` (Path to `oa3tool.exe`)
     - `input_file` (Path to the input CSV)
     - `output_file` (Path for the output CSV)
     - `log_file` (Path for the error log file)
   
   Example paths:
   ```python
   tool_path = r"C:\path\to\oa3tool.exe"
   input_file = r'C:\path\to\input.csv'
   output_file = r'C:\path\to\decoded_product_keys.csv'
   log_file = r'C:\path\to\error_log.txt'
   ```

4. **Run the Script**:
   - Run the script using the following command:

   ```bash
   python3 script_name.py
   ```

5. **Output**:
   - The decoded Product Key IDs will be saved to the specified output file (`output_file`).
   - Any errors or missing ProductKeyID values will be logged in the specified log file (`log_file`).
   - A progress bar will be displayed while processing the entries.

## Example Output

### Output CSV (`decoded_product_keys.csv`):

```csv
SerialNumber,HWHash,ProductKeyID
0F3XXXXXXXXXGT,ABC123XYZ456,3XXXXXXXXXXX6
A12B34C56D78E9,DEF456UVW123,1234567890123
```

### Error Log (`error_log.txt`):

```txt
Warning: ProductKeyID not found for Serial Number: 0F3XXXXXXXXXGT
Error: Failed to run oa3tool for Serial Number: A12B34C56D78E9 - Some error message here...
```

## Troubleshooting

- **Errors with `oa3tool`**:
  - Ensure the path to `oa3tool.exe` is correct and accessible.

- **Empty Output**:
  - Ensure that the hardware hashes and serial numbers in the input CSV are valid and properly formatted.

## License

This script is provided under the [MIT License](LICENSE).
```

This README includes sections on setup, usage, expected input/output, and some common troubleshooting tips. You can modify it as needed!