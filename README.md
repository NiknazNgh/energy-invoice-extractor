# Energy Invoice PDF Extractors — Electricity + Natural Gas

A Python automation project for converting utility invoice PDFs into structured, reviewable Excel/CSV datasets.

The repository contains two invoice workflows:

- **Electricity** — Texas GLO / State Power Program invoice extraction, 27-field standardized output, reconciliation, and discrepancy diagnostics.
- **Natural Gas** — Atmos Energy invoice extraction, multi-year line-item parsing, summary calculations, incremental workbook updates, and unexpected-charge review flags.

Both workflows use rule-based PDF parsing. They do **not** invent values when a field cannot be found.

## Project structure

```text
energy-invoice-extractor/
├── invoice_extractor.py              # electricity extractor
├── gas_invoice_extractor.py          # Atmos natural-gas extractor
├── run_extractor.ps1                 # electricity PowerShell runner
├── run_gas_extractor.ps1             # gas PowerShell runner
├── README.md
├── requirements.txt
├── requirements-dev.txt
├── pyproject.toml
├── LICENSE
├── CHANGELOG.md
├── .gitignore
├── .github/
│   └── workflows/
│       └── tests.yml
├── data/
│   ├── input/                        # electricity PDFs (ignored by Git)
│   └── gas/
│       └── input/
│           ├── 2019/
│           ├── 2020/
│           ├── 2021/
│           ├── 2022/
│           ├── 2023/
│           ├── 2024/
│           ├── 2025/
│           └── 2026/                # gas PDFs (ignored by Git)
├── output/
│   ├── .gitkeep
│   └── gas/
│       └── .gitkeep
├── docs/
│   ├── OUTPUT_COLUMNS.md             # electricity schema
│   └── GAS_OUTPUT_COLUMNS.md         # gas schema
└── tests/
    ├── conftest.py
    ├── test_invoice_extractor.py
    └── test_gas_invoice_extractor.py
```

## Requirements

- Python 3.10+
- pandas
- PyMuPDF
- openpyxl

## Installation

```powershell
git clone https://github.com/NiknazNgh/energy-invoice-extractor.git
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

# Electricity workflow

## Default folders

Put electricity PDFs in:

```text
data/input/
```

Run:

```powershell
python .\invoice_extractor.py
```

or:

```powershell
.\run_extractor.ps1
```

Default outputs:

```text
output/invoice_extracted_table.xlsx
output/invoice_extracted_table.csv
```

You can also supply your own paths:

```powershell
python .\invoice_extractor.py "C:\Invoices\Electricity\2026" "C:\Invoices\Electricity\Results"
```

## Electricity features

- Exact 27-column invoice schema.
- Power factor, load factor, demand, 4CP, usage, ERCOT, transmission, taxes, and bill total extraction.
- `Review Status` and detailed `Review Notes`.
- Bill-total reconciliation using non-overlapping summary buckets.
- Source-line matching when a dollar discrepancy can be tied to a specific PDF line.
- Excel `All Invoices`, `Review Log`, and per-account worksheets.
- CSV export.

See [`docs/OUTPUT_COLUMNS.md`](docs/OUTPUT_COLUMNS.md).

# Natural-gas workflow

The gas extractor is tailored to **Atmos Energy invoice PDFs** and supports historical year folders plus incremental Excel updates.

## Default folders

Place gas invoices in their corresponding year folder:

```text
data/gas/input/2019/
data/gas/input/2020/
...
data/gas/input/2026/
```

Only PDFs whose filename contains `Invoice` are processed. This helps avoid accidentally parsing GCR filing/rate documents as invoices.

Run:

```powershell
python .\gas_invoice_extractor.py
```

or:

```powershell
.\run_gas_extractor.ps1
```

Default output:

```text
output/gas/atmos_energy_invoices.xlsx
```

## Custom gas paths and years

```powershell
python .\gas_invoice_extractor.py `
  "C:\Invoices\NaturalGas" `
  "C:\Invoices\Results\atmos_energy_invoices.xlsx" `
  --years 2024 2025 2026
```

Force a complete rebuild instead of incremental append:

```powershell
python .\gas_invoice_extractor.py `
  ".\data\gas\input" `
  ".\output\gas\atmos_energy_invoices.xlsx" `
  --years 2019 2020 2021 2022 2023 2024 2025 2026 `
  --full-rebuild
```

## Natural-gas features

- Multi-year invoice discovery.
- Text-row, word-coordinate, and token fallback parsing.
- Service account, customer, facility, address, MCF, MMBtu, rates, and charge extraction.
- Tiered sales-service calculations.
- Customer charge, plant-protection fee, MGRT, street/alley fee, FIN 48 tax refund, pipeline-safety fee, and GCR extraction.
- Raw line-item retention for auditability.
- Validation of summary current charges against parsed line-item totals.
- Unexpected/new charge types are retained and marked `NEEDS REVIEW` instead of being silently ignored.
- Incremental update mode skips already-recorded PDFs and old/same billing months.
- Failed-file worksheet for parsing errors.

See [`docs/GAS_OUTPUT_COLUMNS.md`](docs/GAS_OUTPUT_COLUMNS.md).

# Testing

Run all electricity and gas unit tests:

```powershell
pytest
```

GitHub Actions runs the same test suite on pushes and pull requests.

# Data and privacy

Do **not** commit real invoice PDFs, generated workbooks, account details, customer information, or private network paths to a public repository.

The included `.gitignore` excludes PDFs and generated Excel/CSV files. The public version uses repository-local defaults rather than organization-specific network paths.

# Scope

These parsers are designed for the invoice layouts represented by this project. Utility providers can change PDF formatting. A `NEEDS REVIEW`, `CHECK`, or parsing failure should be treated as a prompt for manual verification rather than automatically accepted as correct.
