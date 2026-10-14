# Automated Testing Scheduling

This repository automates submitting rows from an Excel schedule into a Microsoft Form. It can be run locally or from GitHub Actions and supports:

- downloading the source Excel workbook from an HTTP(S) URL
- reading a specific worksheet and row range
- simulating or live-submitting form entries via Playwright
- writing a CSV summary to `processed.csv`

## Repository layout

- `main_msforms.py` - canonical entry point for local use
- `main_msforms_Version4.py` - implementation used by the automation workflow
- `config.yml` - canonical repo configuration
- `config_Version5.yml` - legacy compatibility copy
- `requirements.txt` - canonical dependency list
- `requirements_Version5.txt` - legacy compatibility copy
- `github_workflows_submit-msforms_Version2.yml` - GitHub Actions automation

## Local setup

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python -m playwright install --with-deps chromium
```

## Configuration

Edit `config.yml` (or set `CONFIG_PATH`) with the Excel URL, sheet name, form URL, and field mappings used by the form.

Example values:

- `excel_path`: SharePoint or local `.xlsx` path
- `sheet_name`: worksheet to read
- `single_row`: optional 1-based row override
- `simulate`: `true` to print what would be submitted without launching the browser
- `headless`: `true` in CI, `false` for interactive debugging
- `form_url`: target Microsoft Form URL
- `mapping`: Excel column name to the visible form question text

## Running the script

```bash
python main_msforms.py
```

Optional environment overrides:

```bash
export CONFIG_PATH=config.yml
export EXCEL_BEARER_TOKEN="<token if the file requires authentication>"
```

## Outputs

The script writes a `processed.csv` file summarizing each processed row with its index and final status.
