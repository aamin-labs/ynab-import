#!/usr/bin/env python3
"""
YNAB Transaction Importer
Reads Transaction*.xlsx files from Downloads and uploads to YNAB via API.
"""

import os
import sys
import json
import glob
import logging
import time
import re
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError

try:
    import openpyxl
except ImportError:
    print("Error: openpyxl not installed. Run: pip install -r requirements.txt", file=sys.stderr)
    sys.exit(1)

DOWNLOADS_DIR = Path.home() / "Downloads"
LOG_FILE = Path.home() / "logs" / "ynab_upload.log"
API_BASE = "https://api.ynab.com/v1"
BATCH_SIZE = 200
SETTLED_STATUSES_CARD = {"SETTLED", "REFUNDED"}
SETTLED_STATUSES_BANK = {"SETTLED", "SUCCESS", "REFUNDED"}


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

def setup_logging():
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    fmt = logging.Formatter("%(asctime)s  [%(levelname)s]  %(message)s",
                            datefmt="%Y-%m-%d %H:%M:%S,%f"[:-3])
    logger = logging.getLogger("ynab_import")
    logger.setLevel(logging.INFO)

    fh = logging.FileHandler(LOG_FILE, mode="w", encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    return logger


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

def _parse_env_file(path: Path) -> dict:
    result = {}
    if not path.exists():
        return result
    with open(path) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                key, _, val = line.partition("=")
                result[key.strip()] = val.strip()
    return result


def load_config(logger) -> dict:
    """Load token, budget_id, and account mappings."""
    # Read both env files; env vars take highest priority
    home_env = _parse_env_file(Path.home() / ".ynab.env")
    project_env = _parse_env_file(Path(__file__).parent / ".env")

    def get(key):
        return os.environ.get(key) or project_env.get(key) or home_env.get(key)

    token = get("YNAB_ACCESS_TOKEN")
    if not token:
        logger.error("YNAB_ACCESS_TOKEN not found. Set it in ~/.ynab.env or .env")
        sys.exit(1)

    budget_id = get("YNAB_BUDGET_ID") or "last-used"

    # Collect all YNAB_ACCOUNT_XXXX keys from all sources
    account_map = {}
    for source in (home_env, project_env, dict(os.environ)):
        for key, val in source.items():
            if key.startswith("YNAB_ACCOUNT_"):
                digits = key[len("YNAB_ACCOUNT_"):]
                account_map[digits] = val

    return {"token": token, "budget_id": budget_id, "account_map": account_map}


# ---------------------------------------------------------------------------
# File discovery
# ---------------------------------------------------------------------------

def find_transaction_files() -> list:
    pattern = str(DOWNLOADS_DIR / "Transaction*.xlsx")
    files = sorted(glob.glob(pattern))
    return [Path(f) for f in files]


# ---------------------------------------------------------------------------
# XLSX parsing
# ---------------------------------------------------------------------------

def _extract_identifier(header_cell: str) -> tuple:
    """
    Returns (identifier, format_type) from the A1 cell text.
    format_type is 'card' or 'bank'.
    """
    first_line = header_cell.strip().splitlines()[0]
    if first_line.startswith("Card Number:"):
        # Extract last 4 digits: "Card Number: 522873******2887"
        m = re.search(r"(\d{4})\s*$", first_line)
        identifier = m.group(1) if m else None
        return identifier, "card"
    elif first_line.startswith("Account Number:"):
        # Extract trailing digits: "Account Number: 101XXXXXXXX02"
        m = re.search(r"(\d+)\s*$", first_line)
        identifier = m.group(1) if m else None
        return identifier, "bank"
    return None, "unknown"


def parse_xlsx(filepath: Path, logger) -> tuple:
    """
    Parse a transaction XLSX file.
    Returns (identifier, format_type, transactions_list).
    Each transaction dict has: date_str, payee, memo, amount_str, debit_credit, status
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        wb.close()
        return None, "unknown", []

    # Row 1 (index 0): identifier header
    a1 = str(rows[0][0] or "").strip()
    identifier, fmt = _extract_identifier(a1)

    # Row 4 (index 3): column headers — skip it
    # Rows 5+ (index 4+): data
    transactions = []

    if fmt == "card":
        # Columns: Date | Details | Amount | Currency | Debit/Credit | Status
        valid_statuses = SETTLED_STATUSES_CARD
        for row in rows[4:]:
            if len(row) < 6:
                continue
            date_val, details, amount, currency, debit_credit, status = row[:6]
            if not date_val or not status:
                continue
            status = str(status).strip().upper()
            if status not in valid_statuses:
                continue
            transactions.append({
                "date_str": str(date_val).strip(),
                "payee": str(details or "").strip(),
                "memo": "",
                "amount_str": str(amount or "0").strip(),
                "debit_credit": str(debit_credit or "").strip(),
                "status": status,
            })

    elif fmt == "bank":
        # Columns: Date | Details | Description | Amount | Currency | Balance | Debit/Credit | Status
        valid_statuses = SETTLED_STATUSES_BANK
        for row in rows[4:]:
            if len(row) < 8:
                continue
            date_val, details, description, amount, currency, balance, debit_credit, status = row[:8]
            if not date_val or not status:
                continue
            status = str(status).strip().upper()
            if status not in valid_statuses:
                continue
            transactions.append({
                "date_str": str(date_val).strip(),
                "payee": str(description or "").strip(),
                "memo": str(details or "").strip(),
                "amount_str": str(amount or "0").strip(),
                "debit_credit": str(debit_credit or "").strip(),
                "status": status,
            })

    wb.close()
    return identifier, fmt, transactions


# ---------------------------------------------------------------------------
# Transformation
# ---------------------------------------------------------------------------

def parse_date(date_str: str) -> str:
    """Convert 'Mar 11, 2026' -> '2026-03-11'"""
    return datetime.strptime(date_str.strip(), "%b %d, %Y").strftime("%Y-%m-%d")


def to_milliunits(amount_str: str, debit_credit: str) -> int:
    clean = amount_str.replace(",", "").strip()
    millis = int(round(float(clean) * 1000))
    if debit_credit.strip().lower() == "debit":
        millis = -abs(millis)
    else:
        millis = abs(millis)
    return millis


def build_ynab_transactions(parsed_txns: list, account_id: str) -> list:
    """Convert parsed rows to YNAB API transaction objects."""
    # Count occurrences of (amount, date) pairs for import_id uniqueness
    occurrence_counter = defaultdict(int)
    ynab_txns = []

    for txn in parsed_txns:
        try:
            date = parse_date(txn["date_str"])
        except ValueError:
            continue

        millis = to_milliunits(txn["amount_str"], txn["debit_credit"])
        key = (millis, date)
        occurrence_counter[key] += 1
        occurrence = occurrence_counter[key]

        import_id = f"YNAB:{millis}:{date}:{occurrence}"

        payee = txn["payee"][:50] if txn["payee"] else None
        memo = txn["memo"][:200] if txn["memo"] else None

        obj = {
            "account_id": account_id,
            "date": date,
            "amount": millis,
            "cleared": "cleared",
            "approved": False,
            "import_id": import_id,
        }
        if payee:
            obj["payee_name"] = payee
        if memo:
            obj["memo"] = memo

        ynab_txns.append(obj)

    return ynab_txns


# ---------------------------------------------------------------------------
# YNAB API upload
# ---------------------------------------------------------------------------

def upload_transactions(transactions: list, budget_id: str, token: str, logger) -> tuple:
    """
    Upload transactions in batches.
    Returns (imported_count, duplicate_count).
    Raises SystemExit on auth failure, RuntimeError on other HTTP errors.
    """
    total_imported = 0
    total_dupes = 0
    url = f"{API_BASE}/budgets/{budget_id}/transactions"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    for batch_num, start in enumerate(range(0, len(transactions), BATCH_SIZE), 1):
        batch = transactions[start:start + BATCH_SIZE]
        body = json.dumps({"transactions": batch}).encode("utf-8")
        req = Request(url, data=body, headers=headers, method="POST")

        try:
            with urlopen(req) as resp:
                data = json.loads(resp.read())
                created = data["data"].get("transactions") or []
                dupes = data["data"].get("duplicate_import_ids") or []
                imported = len(created)
                total_imported += imported
                total_dupes += len(dupes)
                dupe_note = f" ({len(dupes)} duplicates skipped)" if dupes else ""
                logger.info(f"Batch {batch_num}: \u2714 Uploaded {imported} transactions{dupe_note}")
        except HTTPError as e:
            if e.code == 401:
                logger.error("Authentication failed — check YNAB_ACCESS_TOKEN")
                sys.exit(1)
            body_text = e.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"HTTP {e.code}: {body_text}")
        except URLError as e:
            raise RuntimeError(f"Network error: {e.reason}")

    return total_imported, total_dupes


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    logger = setup_logging()
    logger.info("\u2500" * 5 + " Import started " + "\u2500" * 5)
    start_time = time.time()

    config = load_config(logger)

    files = find_transaction_files()
    if not files:
        logger.info("No Transaction*.xlsx files found in Downloads")
        return

    files_processed = 0
    files_skipped = 0
    total_imported = 0

    for filepath in files:
        identifier, fmt, parsed_txns = parse_xlsx(filepath, logger)

        if not identifier:
            logger.warning(f"Could not detect account identifier in {filepath.name}, skipping")
            files_skipped += 1
            continue

        type_label = "card" if fmt == "card" else "account"
        logger.info(f"Detected {type_label} ending in {identifier} from {filepath.name}")

        account_id = config["account_map"].get(identifier)
        if not account_id:
            logger.warning(f"No YNAB account mapped for identifier '{identifier}' — add YNAB_ACCOUNT_{identifier}=<uuid> to .env")
            files_skipped += 1
            continue

        if not parsed_txns:
            logger.info(f"No importable transactions in {filepath.name}")
            files_skipped += 1
            continue

        ynab_txns = build_ynab_transactions(parsed_txns, account_id)
        logger.info(f"Prepared {len(ynab_txns)} transactions for import")

        try:
            imported, dupes = upload_transactions(ynab_txns, config["budget_id"], config["token"], logger)
            total_imported += imported
            files_processed += 1
            filepath.unlink()
            logger.info(f"Deleted {filepath.name}")
        except RuntimeError as e:
            logger.error(f"Upload failed for {filepath.name}: {e}")
            files_skipped += 1

    elapsed = time.time() - start_time
    logger.info(
        f"\U0001f389 Batch complete! Processed {files_processed} file(s), "
        f"skipped {files_skipped}, imported {total_imported} transactions in {elapsed:.1f}s"
    )


if __name__ == "__main__":
    main()
