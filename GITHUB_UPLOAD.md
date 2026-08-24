# Publish this project to GitHub

## Option A — GitHub website + Git

1. Create a new empty repository on GitHub named `energy-invoice-extractor`.
2. Do **not** add a README, `.gitignore`, or license on GitHub because they are already included here.
3. Open PowerShell in this project folder and run:

```powershell
git init
git add .
git commit -m "Initial release: energy invoice extractor"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/energy-invoice-extractor.git
git push -u origin main
```

## Option B — GitHub CLI

If GitHub CLI (`gh`) is installed and authenticated:

```powershell
git init
git add .
git commit -m "Initial release: energy invoice extractor"
git branch -M main
gh repo create energy-invoice-extractor --public --source=. --remote=origin --push
```

Use `--private` instead of `--public` if the repository should not be public.

## Before publishing

Run:

```powershell
pytest
```

Then verify that no real invoices, Excel/CSV outputs, account data, or private network paths are staged:

```powershell
git status
git diff --cached
```

The provided `.gitignore` is designed to exclude PDFs, invoice inputs, and generated spreadsheet outputs.
