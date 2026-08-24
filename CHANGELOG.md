# Changelog

## 1.1.0 - 2026-08-24

- Added Atmos Energy natural-gas invoice extraction to the existing electricity project.
- Added multi-year gas invoice discovery and incremental Excel update support.
- Added gas summary, raw line item, validation, and failed-file worksheets.
- Added unexpected-charge detection and `NEEDS REVIEW` flags for gas invoices.
- Added portable gas input/output paths and command-line year selection.
- Removed organization-specific network paths and customer-name fallback from the public gas version.
- Added gas documentation, PowerShell runner, and tests.

## 1.0.0 - 2026-08-24

- Added PDF electricity invoice extraction to Excel and CSV.
- Standardized the 27-field electricity output column order.
- Added Firm Fuel Supply Service - Backbill.
- Renamed final total to Bill Total.
- Added 4CP quantity/rate/charge validation.
- Added bill-total reconciliation without double-counting nested charges.
- Added specific discrepancy diagnostics and source-line matching.
- Added Review Status and Review Notes.
- Added Excel Review Log sheet and per-account worksheets.
- Removed machine-specific default network paths for portable GitHub use.
