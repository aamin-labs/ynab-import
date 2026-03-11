#!/bin/bash
# YNAB Transaction Importer — Alfred wrapper

DIR="$(cd "$(dirname "$0")" && pwd)"

"$DIR/venv/bin/python3" "$DIR/ynab_import.py"

EXIT_CODE=$?

if [ $EXIT_CODE -eq 0 ]; then
    SUMMARY=$(tail -1 /Users/aamin/logs/ynab_upload.log | sed 's/.*\]  //')
    osascript -e "display notification \"$SUMMARY\" with title \"YNAB Import\" sound name \"Glass\""
else
    osascript -e "display notification \"Import failed — check log for details\" with title \"YNAB Import\" sound name \"Basso\""
fi
