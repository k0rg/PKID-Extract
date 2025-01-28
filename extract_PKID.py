import csv
import subprocess
import re
import sys

# Define paths
tool_path = r"C:\path\to\oa3tool.exe"  # Update this with the correct path to oa3tool.exe
input_file = r'C:\path\to\input.csv'
output_file = r'C:\path\to\decoded_product_keys.csv'
log_file = r'C:\path\to\error_log.txt'

output_data = []

# Read input CSV
with open(input_file, mode='r', newline='', encoding='utf-8-sig') as infile:
    csv_reader = csv.DictReader(infile)
    total_rows = sum(1 for _ in csv_reader)  # Count total rows
    infile.seek(0)  # Reset the file pointer

    # Processing each row
    for i, row in enumerate(csv_reader):
        serial_number = row['SerialNumber']
        hw_hash = row['HWHash']

        # Run oa3tool and capture the output
        result = subprocess.run([tool_path, f"/decodehwhash:{hw_hash}"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

        product_key_id = 'Not Found'

        if result.returncode == 0:
            matches = re.findall(r'<p n="ProductKeyID" v="(\d+)" />', result.stdout)
            if matches:
                product_key_id = matches[0]
            else:
                # Log if ProductKeyID is missing
                with open(log_file, 'a', encoding='utf-8-sig') as log:
                    log.write(f"Warning: ProductKeyID not found for Serial Number: {serial_number}\n")
        else:
            # Log error if oa3tool failed to run
            with open(log_file, 'a', encoding='utf-8-sig') as log:
                log.write(f"Error: Failed to run oa3tool for Serial Number: {serial_number} - {result.stderr}\n")
        
        output_data.append({
            'SerialNumber': serial_number,
            'HWHash': hw_hash,
            'ProductKeyID': product_key_id
        })

        # Simple progress indicator
        progress = (i + 1) / total_rows * 100
        sys.stdout.write(f'\rProgress: {i + 1}/{total_rows} ({progress:.2f}%)')
        sys.stdout.flush()

# Finish and output progress
print("\nProcessing Complete")

# Write results to CSV
with open(output_file, mode='w', newline='', encoding='utf-8-sig') as outfile:
    fieldnames = ['SerialNumber', 'HWHash', 'ProductKeyID']
    writer = csv.DictWriter(outfile, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(output_data)

print(f"Results have been saved to {output_file}")
