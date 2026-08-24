param(
    [string]$InputPath = ".\data\gas\input",
    [string]$OutputFile = ".\output\gas\atmos_energy_invoices.xlsx",
    [string[]]$Years = @("2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026"),
    [switch]$FullRebuild
)

$ErrorActionPreference = "Stop"

$arguments = @(
    ".\gas_invoice_extractor.py",
    $InputPath,
    $OutputFile,
    "--years"
) + $Years

if ($FullRebuild) {
    $arguments += "--full-rebuild"
}

python @arguments
