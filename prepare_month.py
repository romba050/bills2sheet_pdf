#!/usr/bin/env python3
"""
Prepare a new month's spreadsheet for receipt import.

The monthly spreadsheet is created by copying last month's in Google Drive
(which carries over the Totals tab and its bound Apps Script). This script
then deletes every leftover receipt tab, keeping only the tabs listed in
classifier_config.json's "skip_tabs" (e.g. "Totals").

Usage:
    uv run python prepare_month.py --spreadsheet-id <ID> [--dry-run]

The spreadsheet must be shared with the service account as Editor.
"""

import argparse
import sys

from buyer_classifier import execute_with_retry, get_sheets_service, load_config


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--spreadsheet-id', required=True,
                        help='Spreadsheet to clean (the long string in the sheet URL)')
    parser.add_argument('--dry-run', action='store_true',
                        help='Only print which tabs would be deleted')
    args = parser.parse_args()

    config = load_config()
    service = get_sheets_service()
    meta = execute_with_retry(service.spreadsheets().get(spreadsheetId=args.spreadsheet_id))
    title = meta['properties']['title']
    keep = set(config['skip_tabs'])
    to_delete = [s['properties'] for s in meta['sheets'] if s['properties']['title'] not in keep]
    kept = [s['properties']['title'] for s in meta['sheets'] if s['properties']['title'] in keep]

    if not kept:
        print(f"Error: '{title}' has none of the tabs to keep ({', '.join(sorted(keep))}).")
        print("Refusing to delete everything - is this the right spreadsheet?")
        sys.exit(1)
    if not to_delete:
        print(f"'{title}' already contains only: {', '.join(kept)}. Nothing to do.")
        return

    print(f"Spreadsheet: {title}")
    for props in to_delete:
        print(f"  {'would delete' if args.dry_run else 'deleting'}: {props['title']}")
    if args.dry_run:
        return

    requests = [{'deleteSheet': {'sheetId': props['sheetId']}} for props in to_delete]
    execute_with_retry(service.spreadsheets().batchUpdate(
        spreadsheetId=args.spreadsheet_id, body={'requests': requests}))
    print(f"Done. Remaining tabs: {', '.join(kept)}")


if __name__ == '__main__':
    main()
