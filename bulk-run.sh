#!/bin/bash

# Parse arguments
YEAR=""
MONTH=""
AFTER_DAY=""
SPREADSHEET_ID=""
STORE=""

while [[ $# -gt 0 ]]; do
  case $1 in
    --year)
      YEAR="$2"
      shift 2
      ;;
    --month)
      MONTH="$2"
      shift 2
      ;;
    --after-day)
      AFTER_DAY="$2"
      shift 2
      ;;
    --spreadsheet-id)
      SPREADSHEET_ID="$2"
      shift 2
      ;;
    --store)
      STORE="$2"
      shift 2
      ;;
    *)
      echo "Unknown option: $1"
      exit 1
      ;;
  esac
done

# Validate required arguments
if [[ -z "$YEAR" || -z "$MONTH" || -z "$AFTER_DAY" || -z "$SPREADSHEET_ID" || -z "$STORE" ]]; then
  echo "Usage: $0 --year YYYY --month MM --after-day DD --spreadsheet-id ID --store STORE_NAME"
  echo "Example: $0 --year 2026 --month 04 --after-day 07 --spreadsheet-id 1azIe-VlxovmI8CrRMUj6LeJeprJMPoCQvI-tTAPwHxA --store ICA"
  exit 1
fi

for f in bills/ICA\ Supermarket\ Brommaplan\ ${YEAR}-${MONTH}-*.pdf; do
  # -6 means start 6 characters from the end of the filename, 2 means extract 2 characters, thus the day (DD) is extracted from a filename of form <*DD.pdf>
  dd="${f: -6:2}" 
  if [[ "$dd" > "$AFTER_DAY" ]]; then
    uv run python receipt_processor.py \
      --spreadsheet-id "$SPREADSHEET_ID" \
      "$f" --store="$STORE"
  fi
done
