# ServiceNow Excel to Table Uploader

Small script to read the first 3 sheets from an Excel workbook and push rows to a ServiceNow table. It also demonstrates a lightweight read from the same table.

## Requirements
- Windows
- Python 3.8+
- Packages:
  - requests
  - pandas
  - python-dotenv
  - openpyxl

Install dependencies:
```powershell
pip install requests pandas python-dotenv openpyxl

Files
test.py — main script (reads Excel, reads from ServiceNow, posts rows)
Book1.xlsx — input workbook (placed at C:\Users\mshafei\Desktop\code\Book1.xlsx)
variables.env or .env — environment variables file (kept out of source control)
```
Notes:

SN_VERIFY_CERT=false disables SSL verification (INSECURE — only for testing).
If your org uses a custom CA, point verification to a CA bundle or enable verification and provide the CA file as needed.
Excel input expectations
Script reads up to the first 3 sheets.
Each sheet is converted to records where column headers become dictionary keys.
Typical columns used in the script: name, email, phone (edit script mapping as required).
Usage
Open PowerShell and run:



Usage
```
Open PowerShell and run:

# run
python "C:\Users\mshafei\Desktop\code\test.py"
```
Make sure variables.env is present and credentials are correct.

# Behaviour
Reads sheet rows, maps fields, posts each row to the ServiceNow table configured by SN_TABLE.
Performs a GET to validate connection before posting.
Prints success/errors for each row.
Troubleshooting
SSL errors (CERTIFICATE_VERIFY_FAILED): set SN_VERIFY_CERT=false (for quick testing) or install certifi and configure verification properly. Prefer pointing to your org CA bundle.
Unauthorized (401): check SN_USER / SN_PASS.
Forbidden (403): account lacks permissions for the target table.
Changes

Edit test.py to change field mappings, target table, or to add more validation before posting.
