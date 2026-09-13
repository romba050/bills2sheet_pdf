#!/usr/bin/env python3
"""
Grocery expense buyer classifier.

Pre-fills the buyer column (Shared/My/Jessica expenses) of receipt sheets with a
first guess, learned from historically corrected months. Transparent
frequency-count rules only, no ML. The rule table lives in a version-controlled
CSV that can be edited by hand; manually edited rules are never overwritten.

Commands:
    uv run python buyer_classifier.py learn
    uv run python buyer_classifier.py classify --spreadsheet-id ID [--sheet TAB] [--apply]
    uv run python buyer_classifier.py classify --spreadsheet-id ID --evaluate
    uv run python buyer_classifier.py rules

All thresholds and spreadsheet IDs are in classifier_config.json.
"""

import argparse
import csv
import json
import re
import sys
import time
from datetime import date, datetime
from pathlib import Path

try:
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    from google.oauth2 import service_account
    from rapidfuzz import fuzz, process as fuzz_process
except ImportError:
    print("Error: dependencies not installed. Run: uv sync")
    sys.exit(1)


def execute_with_retry(request, attempts=5):
    """Execute a Sheets API request, backing off on rate limits (60 reads/min)."""
    for attempt in range(attempts - 1):
        try:
            return request.execute()
        except HttpError as e:
            if e.resp.status not in (429, 500, 503):
                raise
            wait = 20 * (attempt + 1)
            print(f"Rate limited by Sheets API, retrying in {wait}s...")
            time.sleep(wait)
    return request.execute()


SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CONFIG_FILE = 'classifier_config.json'

BUYERS = ['shared', 'me', 'jessica']
# Buyer -> 0-based column index in the A:D item table (B, C, D)
BUYER_COL = {'shared': 1, 'me': 2, 'jessica': 3}
COL_BUYER = {v: k for k, v in BUYER_COL.items()}

ITEM_HEADER = ['Item', 'Shared expenses', 'My expenses', 'Jessica expenses']


def is_item_header(cells: list) -> bool:
    """Tolerant header detection: older sheets name the columns slightly
    differently, but always as item / shared / my-ish / Jessica."""
    a, b, c, d = (str(cell).strip().lower() for cell in cells[:4])
    return (a.startswith('item')
            and 'shared' in b
            and ('my' in c or 'mine' in c or 'basile' in c)
            and 'jessica' in d)

RULE_FIELDS = ['item_normalized', 'assigned_buyer', 'source',
               'shared_count', 'me_count', 'jessica_count',
               'confidence', 'last_seen']

MONTH_NAMES = {name.lower(): i for i, name in enumerate(
    ['january', 'february', 'march', 'april', 'may', 'june',
     'july', 'august', 'september', 'october', 'november', 'december'], start=1)}


def load_config() -> dict:
    config_path = Path(__file__).parent / CONFIG_FILE
    if not config_path.exists():
        print(f"Error: {CONFIG_FILE} not found next to this script.")
        sys.exit(1)
    with open(config_path) as f:
        return json.load(f)


def get_sheets_service(credentials_file: str = 'service_account.json'):
    """Authenticate with the same service account used by receipt_processor."""
    if not Path(credentials_file).exists():
        print(f"Error: {credentials_file} not found.")
        sys.exit(1)
    creds = service_account.Credentials.from_service_account_file(credentials_file, scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)


# --- Item normalization -----------------------------------------------------

def normalize(item: str) -> str:
    """Reduce a receipt item string to a canonical rule key.

    Lowercases, strips punctuation and any token containing a digit
    (quantities, sizes, discount codes like "2f45kr" or "69kr/st"), so that
    "Jordgubbar 69kr/st" and "Jordgubbar" share one key. Falls back to the
    bare lowercased string if stripping removes everything.
    """
    s = item.lower().strip()
    s = re.sub(r'[*"\'!?:;()\[\]]', ' ', s)
    tokens = [t for t in s.split() if not re.search(r'\d', t)]
    result = ' '.join(tokens)
    if not result:
        result = ' '.join(s.split())
    return result


# --- Rule table -------------------------------------------------------------

def assign_buyer(shared: float, me: float, jessica: float, shared_threshold: float) -> str:
    """Majority label, except shared wins already at shared_threshold share."""
    total = shared + me + jessica
    if total <= 0:
        return 'shared'
    if shared / total >= shared_threshold:
        return 'shared'
    if me > jessica:
        return 'me'
    if jessica > me:
        return 'jessica'
    return 'shared'


def load_rules(config: dict) -> tuple[dict, list]:
    """Load the rule CSV.

    Returns (rules dict keyed by item_normalized, ordered key list).
    Detects manual edits: an 'auto' row whose assigned_buyer disagrees with
    its own counts was edited by the user and is flipped to 'manual'.
    Detection uses the shared_threshold recorded when the table was written
    (rules/.learn_meta.json), so changing the threshold in the config is not
    mistaken for a user edit.
    """
    rules_path = Path(config['rules_csv'])
    rules, order = {}, []
    if not rules_path.exists():
        return rules, order
    meta_path = rules_path.parent / '.learn_meta.json'
    detection_threshold = config['shared_threshold']
    if meta_path.exists():
        with open(meta_path) as f:
            detection_threshold = json.load(f).get('shared_threshold', detection_threshold)
    with open(rules_path, newline='') as f:
        for row in csv.DictReader(f):
            key = row['item_normalized']
            row['shared_count'] = float(row['shared_count'] or 0)
            row['me_count'] = float(row['me_count'] or 0)
            row['jessica_count'] = float(row['jessica_count'] or 0)
            if row['source'] == 'auto':
                expected = assign_buyer(row['shared_count'], row['me_count'],
                                        row['jessica_count'], detection_threshold)
                if row['assigned_buyer'] != expected:
                    print(f"Detected manual edit of '{key}' "
                          f"({expected} -> {row['assigned_buyer']}), marking as manual.")
                    row['source'] = 'manual'
            rules[key] = row
            order.append(key)
    return rules, order


def save_rules(config: dict, rules: dict, order: list) -> None:
    """Write the rule CSV, preserving existing row order; new keys go last, sorted."""
    rules_path = Path(config['rules_csv'])
    rules_path.parent.mkdir(parents=True, exist_ok=True)
    with open(rules_path.parent / '.learn_meta.json', 'w') as f:
        json.dump({'shared_threshold': config['shared_threshold']}, f)
    new_keys = sorted(k for k in rules if k not in set(order))
    with open(rules_path, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=RULE_FIELDS)
        writer.writeheader()
        for key in list(order) + new_keys:
            if key not in rules:
                continue
            row = dict(rules[key])
            for field in ('shared_count', 'me_count', 'jessica_count'):
                row[field] = f"{float(row[field]):g}"
            writer.writerow({field: row.get(field, '') for field in RULE_FIELDS})


# --- Sheet reading ----------------------------------------------------------

def list_item_tabs(service, spreadsheet_id: str, skip_tabs: list) -> list:
    meta = execute_with_retry(service.spreadsheets().get(spreadsheetId=spreadsheet_id))
    return [s['properties']['title'] for s in meta['sheets']
            if s['properties']['title'] not in skip_tabs]


def fetch_tabs(service, spreadsheet_id: str, tabs: list) -> dict:
    """Fetch the A:D range of many tabs in a single batchGet request."""
    if not tabs:
        return {}
    result = execute_with_retry(service.spreadsheets().values().batchGet(
        spreadsheetId=spreadsheet_id, ranges=[f"'{t}'!A1:D400" for t in tabs]))
    return {tab: vr.get('values', [])
            for tab, vr in zip(tabs, result.get('valueRanges', []))}


def parse_item_rows(values: list) -> tuple[list, bool]:
    """Parse one receipt tab's A:D values into (item rows, header_found).
    Rows are (row_number, item_name, amount, buyer_or_None, n_filled_columns).

    buyer is None when no expense column is filled; n_filled_columns > 1
    signals an invalid row (amount in several columns).
    """
    rows = []
    header_seen = False
    for i, row in enumerate(values, start=1):
        row = row + [''] * (4 - len(row))
        if not header_seen:
            if is_item_header(row):
                header_seen = True
            continue
        name = str(row[0]).strip()
        if not name:
            continue
        filled = []
        for col in (1, 2, 3):
            cell = str(row[col]).strip().replace(',', '.')
            if cell:
                try:
                    filled.append((col, float(cell)))
                except ValueError:
                    pass
        if len(filled) == 1:
            col, amount = filled[0]
            rows.append((i, name, amount, COL_BUYER[col], 1))
        elif len(filled) > 1:
            rows.append((i, name, None, None, len(filled)))
        else:
            rows.append((i, name, None, None, 0))
    return rows, header_seen


def tab_date(tab: str, spreadsheet_label: str) -> date | None:
    """Date of a receipt tab, from its name (YYYY-MM-DD-...) or else from the
    spreadsheet label (e.g. 'June 2026')."""
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})', tab)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    m = re.search(r'([A-Za-z]+)\s+(\d{4})', spreadsheet_label)
    if m and m.group(1).lower() in MONTH_NAMES:
        return date(int(m.group(2)), MONTH_NAMES[m.group(1).lower()], 15)
    return None


# --- learn ------------------------------------------------------------------

def cmd_learn(config: dict, args) -> None:
    service = get_sheets_service()
    history = dict(config['history_spreadsheets'])
    if args.spreadsheet_id:
        history = {f'cli-{i}': sid for i, sid in enumerate(args.spreadsheet_id)}
    if args.only:
        wanted = {label.lower() for label in args.only}
        unknown = wanted - {label.lower() for label in history}
        if unknown:
            print(f"Error: --only labels not in the config's history_spreadsheets: "
                  f"{', '.join(sorted(unknown))}")
            print(f"Available: {', '.join(history)}")
            sys.exit(1)
        history = {label: sid for label, sid in history.items()
                   if label.lower() in wanted}
    if args.exclude:
        excluded = {e.lower() for e in args.exclude}
        history = {label: sid for label, sid in history.items()
                   if sid.lower() not in excluded and label.lower() not in excluded}
    print(f"Learning from: {', '.join(history)}\n")

    half_life = config['recency_half_life_months']
    use_recency = config['recency_weighting']
    today = date.today()

    counts = {}      # key -> {'shared': w, 'me': w, 'jessica': w}
    last_seen = {}   # key -> date
    warnings = []

    for label, sid in history.items():
        try:
            tabs = list_item_tabs(service, sid, config['skip_tabs'])
        except Exception as e:
            warnings.append(f"Cannot read '{label}' ({sid}): {e}")
            continue
        tab_values = fetch_tabs(service, sid, tabs)
        for tab in tabs:
            rows, header_found = parse_item_rows(tab_values[tab])
            if not header_found:
                warnings.append(f"Skipped '{label}' / '{tab}': no item table header found.")
                continue
            item_rows = [r for r in rows if r[3] is not None]
            if not item_rows:
                continue
            if (config['skip_all_shared_tabs_when_learning']
                    and all(r[3] == 'shared' for r in item_rows)):
                warnings.append(f"Skipped '{label}' / '{tab}': every amount is in the "
                                "Shared column (looks like a raw, uncorrected import).")
                continue
            when = tab_date(tab, label)
            weight = 1.0
            if use_recency and when:
                age_months = (today - when).days / 30.44
                weight = 0.5 ** (age_months / half_life)
            for row_num, name, amount, buyer, n_filled in rows:
                if n_filled > 1:
                    warnings.append(f"'{label}' / '{tab}' row {row_num} ('{name}'): "
                                    f"{n_filled} expense columns filled, skipped.")
                    continue
                if buyer is None:
                    continue
                key = normalize(name)
                counts.setdefault(key, {b: 0.0 for b in BUYERS})[buyer] += weight
                if when and (key not in last_seen or when > last_seen[key]):
                    last_seen[key] = when
        print(f"Learned from '{label}'")

    if not counts:
        print("No historical data could be read. Are the spreadsheets shared with the "
              "service account?")
        for w in warnings:
            print(f"  warning: {w}")
        sys.exit(1)

    rules, order = load_rules(config)
    # Auto rows always reflect exactly the months scanned in this run; rules
    # for items not seen in any scanned month are dropped (manual rows stay).
    stale = [k for k, r in rules.items() if r['source'] != 'manual' and k not in counts]
    for key in stale:
        del rules[key]
    updated, added, kept_manual = 0, 0, 0
    for key, c in sorted(counts.items()):
        if key in rules and rules[key]['source'] == 'manual':
            kept_manual += 1
            continue
        total = sum(c.values())
        buyer = assign_buyer(c['shared'], c['me'], c['jessica'], config['shared_threshold'])
        confidence = c[buyer] / total if total else 0.0
        if buyer != 'shared':
            confidence = max(c.values()) / total if total else 0.0
        row = {
            'item_normalized': key,
            'assigned_buyer': buyer,
            'source': 'auto',
            'shared_count': round(c['shared'], 2),
            'me_count': round(c['me'], 2),
            'jessica_count': round(c['jessica'], 2),
            'confidence': f"{confidence:.2f}",
            'last_seen': last_seen.get(key, '').isoformat() if key in last_seen else '',
        }
        if key in rules:
            updated += 1
        else:
            added += 1
        rules[key] = row
    save_rules(config, rules, order)

    print(f"\nRule table written to {config['rules_csv']}: "
          f"{added} added, {updated} updated, {len(stale)} stale removed, "
          f"{kept_manual} manual rules untouched.")
    if warnings:
        print(f"\n{len(warnings)} warning(s):")
        for w in warnings:
            print(f"  {w}")
    if added and not order:
        print("\nTip: commit the rule table to git so rule changes show up in history.")


# --- classify ---------------------------------------------------------------

def predict(key: str, rules: dict, config: dict):
    """Return (buyer, match_kind, matched_key). match_kind: exact/fuzzy/default."""
    if key in rules:
        return rules[key]['assigned_buyer'], 'exact', key
    if rules:
        match = fuzz_process.extractOne(
            key, rules.keys(), scorer=fuzz.ratio, score_cutoff=config['fuzzy_threshold'])
        if match:
            return rules[match[0]]['assigned_buyer'], 'fuzzy', match[0]
    return config['default_buyer'], 'default', None


def cmd_classify(config: dict, args) -> None:
    service = get_sheets_service()
    rules, _ = load_rules(config)
    if not rules:
        print(f"Error: rule table {config['rules_csv']} is missing or empty. "
              "Run 'learn' first.")
        sys.exit(1)

    spreadsheet_id = args.spreadsheet_id
    tabs = [args.sheet] if args.sheet else list_item_tabs(service, spreadsheet_id,
                                                          config['skip_tabs'])
    dry_run = not args.apply
    stats = {'classified': 0, 'exact': 0, 'fuzzy': 0, 'default': 0,
             'skipped_filled': 0, 'moved': 0}
    flagged, eval_rows = [], []

    tab_values = fetch_tabs(service, spreadsheet_id, tabs)
    for tab in tabs:
        rows, header_found = parse_item_rows(tab_values[tab])
        if not header_found:
            print(f"Skipped '{tab}': no item table header found.")
            continue
        updates = []
        for row_num, name, amount, buyer, n_filled in rows:
            if n_filled != 1:
                continue
            key = normalize(name)
            if args.evaluate:
                predicted, kind, _ = predict(key, rules, config)
                eval_rows.append((tab, name, buyer, predicted, kind))
                continue
            if buyer != 'shared':
                # Amount already moved to My/Jessica: user has handled this row.
                stats['skipped_filled'] += 1
                continue
            predicted, kind, matched = predict(key, rules, config)
            stats['classified'] += 1
            stats[kind] += 1
            if kind == 'fuzzy':
                flagged.append(f"'{tab}' row {row_num}: '{name}' ~ rule '{matched}' "
                               f"-> {predicted}")
            elif kind == 'default':
                flagged.append(f"'{tab}' row {row_num}: '{name}' unmatched "
                               f"-> default {predicted}")
            if predicted != 'shared':
                stats['moved'] += 1
                col_letter = chr(ord('A') + BUYER_COL[predicted])
                prefix = '[dry-run] ' if dry_run else ''
                print(f"{prefix}'{tab}' row {row_num}: '{name}' {amount} -> {predicted}")
                if not dry_run:
                    updates.append({'range': f"'{tab}'!B{row_num}:D{row_num}",
                                    'values': [['', '', '']]})
                    updates.append({'range': f"'{tab}'!{col_letter}{row_num}",
                                    'values': [[amount]]})
        if updates:
            execute_with_retry(service.spreadsheets().values().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={'valueInputOption': 'USER_ENTERED', 'data': updates}))

    if args.evaluate:
        total = len(eval_rows)
        correct = sum(1 for _, _, actual, predicted, _ in eval_rows if actual == predicted)
        print(f"\nEvaluation on {total} already-assigned rows: "
              f"{correct} correct ({100 * correct / total:.1f}%)" if total else "No rows.")
        for tab, name, actual, predicted, kind in eval_rows:
            if actual != predicted:
                print(f"  wrong ({kind}): '{tab}' '{name}': predicted {predicted}, "
                      f"actual {actual}")
        return

    mode = 'DRY-RUN (no cells written; use --apply to write)' if dry_run else 'applied'
    print(f"\nSummary ({mode}):")
    print(f"  {stats['classified']} rows classified "
          f"({stats['exact']} exact, {stats['fuzzy']} fuzzy, {stats['default']} defaulted)")
    print(f"  {stats['moved']} amounts {'would move' if dry_run else 'moved'} "
          f"out of the Shared column")
    print(f"  {stats['skipped_filled']} rows already assigned by you, untouched")
    if flagged:
        print("\nCheck these:")
        for line in flagged:
            print(f"  {line}")


# --- rules ------------------------------------------------------------------

def cmd_rules(config: dict, args) -> None:
    rules, order = load_rules(config)
    if not rules:
        print(f"Rule table {config['rules_csv']} is missing or empty. Run 'learn' first.")
        return
    rows = sorted(rules.values(),
                  key=lambda r: -(r['shared_count'] + r['me_count'] + r['jessica_count']))
    print(f"{'item':40} {'buyer':8} {'source':7} {'shared':>7} {'me':>7} "
          f"{'jessica':>7} {'conf':>5} last_seen")
    for r in rows:
        print(f"{r['item_normalized'][:40]:40} {r['assigned_buyer']:8} {r['source']:7} "
              f"{r['shared_count']:7.2f} {r['me_count']:7.2f} {r['jessica_count']:7.2f} "
              f"{r['confidence']:>5} {r['last_seen']}")
    print(f"\n{len(rules)} rules "
          f"({sum(1 for r in rules.values() if r['source'] == 'manual')} manual)")


def main():
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest='command', required=True)

    p_learn = sub.add_parser('learn', help='Rebuild the rule table from historical sheets')
    p_learn_src = p_learn.add_mutually_exclusive_group()
    p_learn_src.add_argument('--only', action='append',
                             help='Learn only from these config months, by label, '
                                  'e.g. --only "July 2026" (repeatable, case-insensitive)')
    p_learn_src.add_argument('--spreadsheet-id', action='append',
                             help='Learn only from these raw spreadsheet IDs (repeatable)')
    p_learn.add_argument('--exclude', action='append', default=[],
                         help='Skip a config history entry (label or ID, repeatable)')

    p_cls = sub.add_parser('classify', help='Pre-fill buyer columns of the current month')
    p_cls.add_argument('--spreadsheet-id', required=True,
                       help='Target spreadsheet ID (the long string in the sheet URL)')
    p_cls.add_argument('--sheet', help='Only this tab (default: all item tabs)')
    p_cls.add_argument('--apply', action='store_true',
                       help='Write to the sheet (default is dry-run)')
    p_cls.add_argument('--evaluate', action='store_true',
                       help='Score predictions against already-assigned rows; writes nothing')

    sub.add_parser('rules', help='Show the rule table')

    args = parser.parse_args()
    config = load_config()
    {'learn': cmd_learn, 'classify': cmd_classify, 'rules': cmd_rules}[args.command](config, args)


if __name__ == '__main__':
    main()
