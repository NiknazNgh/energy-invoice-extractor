# Publish or update this project on GitHub

Repository:

```text
https://github.com/NiknazNgh/energy-invoice-extractor
```

## Existing repository: add the gas workflow

After copying the updated project files into your local repository, run:

```powershell
git status
git add .
git commit -m "Add natural gas invoice extraction workflow"
git push origin main
```

Before committing, confirm that no invoice PDFs, Excel/CSV outputs, customer data, or private network paths are staged:

```powershell
git status
git diff --cached
```

## Fresh clone

```powershell
git clone https://github.com/NiknazNgh/energy-invoice-extractor.git
cd energy-invoice-extractor
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements-dev.txt
pytest
```

## Optional repository rename

Because the repository now handles both electricity and natural gas, a broader name such as `utility-invoice-extractor` would be reasonable later. Renaming is optional; keeping `energy-invoice-extractor` is also accurate.
