# ynab-import

Imports bank transaction exports (XLSX) into YNAB via the API.

Reads `Transaction*.xlsx` files from `~/Downloads`, transforms them into YNAB transactions, uploads in batches, and deletes the source files on success.

## Repo boundary

This repo owns downloaded bank/card statement imports and the `~/Downloads` launchd watcher. `/Users/aamin/dev/projects/ynab-tools` owns YNAB reporting/export utilities only. Do not add bank-import watchers or upload logic there.

## Supported file formats

| Format | Detection | Columns |
|--------|-----------|---------|
| Credit card statement | `Card Number:` in row 1 | Date, Details, Amount, Currency, Debit/Credit, Status |
| Bank account statement | `Account Number:` in row 1 | Date, Details, Description, Amount, Currency, Balance, Debit/Credit, Status |

Imported statuses: `SETTLED`, `REFUNDED` (cards) · `SETTLED`, `SUCCESS`, `REFUNDED` (bank accounts)

## Installation

```bash
git clone https://github.com/your-username/ynab-import.git
cd ynab-import
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Configuration

**1. YNAB credentials** — add to `~/.ynab.env`:

```
YNAB_ACCESS_TOKEN=your-token-here
YNAB_BUDGET_ID=your-budget-id-here
```

Get your token at [app.ynab.com/settings/developer](https://app.ynab.com/settings/developer).

**2. Account mappings** — copy `.env.example` to `.env` and fill in your YNAB account UUIDs:

```bash
cp .env.example .env
```

Map the last digits of each card/account number to its YNAB account UUID:

```
YNAB_ACCOUNT_2887=uuid-for-your-card
YNAB_ACCOUNT_02=uuid-for-your-bank-account
YNAB_ACCOUNT_03=uuid-for-saver-plus
```

Current local mapping includes Saver Plus account `097XXXXXXXX03` as `YNAB_ACCOUNT_03`.

## Usage

```bash
# Run directly
venv/bin/python3 ynab_import.py

# Or via the Alfred wrapper
./run_import.sh
```

Logs are written to `~/logs/ynab_upload.log` (overwritten each run).

## Alfred integration

Point an Alfred **Run Script** workflow at `run_import.sh`. It will run the import and display a macOS notification with the result.

## Requirements

Python 3.9+, `openpyxl`
