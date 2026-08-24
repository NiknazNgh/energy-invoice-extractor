param(
    [string]$InputPath = ".\data\input",
    [string]$OutputPath = ".\output"
)

$ErrorActionPreference = "Stop"
python .\invoice_extractor.py $InputPath $OutputPath
