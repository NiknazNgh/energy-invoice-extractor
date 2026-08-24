# Energy Invoice PDF Extractor & Reconciliation Tool

A Python project for extracting structured electricity-invoice data from PDF files and exporting the results to Excel and CSV with built-in reconciliation and review diagnostics.

The parser is currently tailored to the **Texas GLO / State Power Program invoice layout** used by this project. It uses rule-based PDF text extraction and validation rather than OCR or AI-generated values.

## What it does

- Reads one PDF invoice or recursively processes a folder of PDFs.
- Extracts billing period, demand, usage, 4CP, energy, ERCOT, transmission, taxes, and bill-total fields.
- Writes the invoice data in a fixed 27-column schema.
- Appends `Review Status` and `Review Notes`.
- Flags invoices as `NEEDS REVIEW` when parsed values do not reconcile.
- Explains the dollar discrepancy instead of returning a generic validation warning.
- Searches source PDF text for a charge line whose amount matches the discrepancy.
- Avoids double-counting nested charges during reconciliation.
- Creates an Excel workbook with:
  - `All Invoices`
  - `Review Log`
  - one worksheet per detected account
- Also creates a machine-friendly CSV export.

## Exact output column order

1. Production Month
2. From
3. To
4. Invoice Date
5. Power Factor
6. Load Factor
7. Actual Demand (KW)
8. Billing Demand (KW)
9. 4CP Charges Qty (KW)
10. 4CP Charges Rate ($/KW)
11. 4CP Charges ($)
12. Usage - Actual KWH
13. UOM
14. Energy Charge
15. Nodal Congestion Charge
16. Market Securitization (Debt) Financing - Default Charge
17. Prior Period Pass Through Charge
18. ERCOT Cont Reserve Serv (ECRS)
19. Firm Fuel Supply Service
20. Firm Fuel Supply Service - Backbill
21. Market Securitization - Uplift Charge
22. TX-ERCOT Admin Fees - CIL
23. Transmission Charges
24. Taxes & PUC Assessment Charge
25. Ancilliary Service Obligation Adjustment
26. Other Taxes
27. Bill Total

Then:

- Review Status
- Review Notes

See [`docs/OUTPUT_COLUMNS.md`](docs/OUTPUT_COLUMNS.md) for validation details.

## Review logic

The reconciliation uses non-overlapping summary buckets:

```text
Parsed Summary =
    Energy Charge
  + Transmission Charges
  + Taxes & PUC Assessment Charge
  + Prior Period Pass Through Charge
  + Other Taxes
```

If the prior-period total is missing, parsed prior-period component charges can be used as a fallback.

The tool intentionally does **not** add these twice:

- `4CP Charges ($)` is not added separately because it is part of Transmission Charges.
- `Nodal Congestion Charge` and the Default Charge are not added again because they are included in Energy Charge.
- Prior-period detail components are not added when the prior-period summary total exists.

### Example review message

```text
NEEDS REVIEW
Bill total reconciliation mismatch: Bill Total is higher than the parsed summary by $186,208.06.
Bill Total=$500,000.00; Parsed Summary=$313,791.94.
Possible source line explaining discrepancy: Special Capacity Charge 186,208.06
```

If no single known field or PDF source line explains the difference, the tool reports the unreconciled amount and asks for manual review of omitted charges, credits, or subtotals.

## Project structure

```text
energy-invoice-extractor/
├── invoice_extractor.py
├── README.md
├── requirements.txt
├── requirements-dev.txt
├── pyproject.toml
├── run_extractor.ps1
├── LICENSE
├── CHANGELOG.md
├── .gitignore
├── .github/
│   └── workflows/
│       └── tests.yml
├── data/
│   └── input/
│       └── .gitkeep
├── output/
│   └── .gitkeep
├── docs/
│   └── OUTPUT_COLUMNS.md
└── tests/
    └── test_invoice_extractor.py
```

## Requirements

- Python 3.10+
- pandas
- PyMuPDF
- openpyxl

## Installation

### Windows PowerShell

```powershell
git clone https://github.com/YOUR-USERNAME/energy-invoice-extractor.git
cd energy-invoice-extractor

python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

For development/testing:

```powershell
pip install -r requirements-dev.txt
pytest
```

## Usage

### Option 1: default repository folders

Put invoice PDFs inside:

```text
data/input/
```

Then run:

```powershell
python .\invoice_extractor.py
```

Results are written to:

```text
output/invoice_extracted_table.xlsx
output/invoice_extracted_table.csv
```

### Option 2: specify your own input and output folders

```powershell
python .\invoice_extractor.py "C:\Invoices\2026" "C:\Invoices\Results"
```

You can also use the included PowerShell wrapper:

```powershell
.\run_extractor.ps1 -InputPath "C:\Invoices\2026" -OutputPath "C:\Invoices\Results"
```

## Excel output

The workbook contains:

### All Invoices

The complete invoice table in the required column order.

### Review Log

A compact audit table containing:

- Account
- Provider
- Production Month
- Source File
- Review Status
- Review Notes

### Account worksheets

A separate worksheet is generated for each detected account.

## Validation states

### `OK`

No validation issue was detected and the bill total reconciles within the configured tolerance.

### `NEEDS REVIEW`

Used when one or more of the following occurs:

- Bill Total does not reconcile.
- A required summary field cannot be parsed.
- A field label appears in the PDF but its amount was not extracted.
- 4CP quantity × rate differs from the parsed 4CP charge.
- A PDF cannot be parsed.

`Review Notes` contains the actual reason and dollar discrepancy.

## Testing

Run:

```powershell
pytest
```

The included tests verify:

- exact 27-column output order;
- production-month calculation;
- percentage conversion;
- successful bill reconciliation;
- discrepancy source-line identification;
- prevention of 4CP double-counting.

GitHub Actions automatically runs the tests on pushes and pull requests.

## Data protection

Real invoices may contain account numbers, billing details, addresses, or operational information. The repository `.gitignore` excludes:

- PDF files;
- files inside `data/input/`;
- generated Excel files;
- generated CSV files.

Do not commit real invoices or sensitive production data to a public repository unless you are authorized to publish them.

## Current limitations

- The parser depends on text embedded in the PDF. A scanned/image-only invoice requires OCR before this script can extract values.
- Regex patterns are tailored to the current Texas GLO / State Power Program invoice wording.
- Invoice layout or label changes may require parser updates.
- A `NEEDS REVIEW` result is an audit flag, not proof that the invoice itself is incorrect.

## Roadmap

Potential next improvements:

- provider-specific parser modules;
- OCR fallback for scanned invoices;
- configuration-driven charge definitions;
- unit/integration tests using sanitized sample invoices;
- summary dashboard for monthly usage and cost trends;
- logging and structured error reports;
- command-line flags for tolerance and provider selection.

## License

MIT License. See [`LICENSE`](LICENSE).
