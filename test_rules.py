"""Tests for the rule-table logic of buyer_classifier: buyer assignment
thresholds, manual-edit detection, order-preserving CSV round-trips, and
prediction matching. Run with: uv run python test_rules.py
Writes results to debugging/rules_test_results.txt.
"""

import csv
import tempfile
from pathlib import Path

from buyer_classifier import assign_buyer, load_rules, save_rules, predict

RESULTS = []
FAILED = 0


def check(name, condition, detail=''):
    global FAILED
    FAILED += not condition
    RESULTS.append(f"{'PASS' if condition else 'FAIL'} {name}"
                   + (f" ({detail})" if detail and not condition else ''))


def make_config(tmpdir):
    return {
        'rules_csv': str(Path(tmpdir) / 'rules.csv'),
        'shared_threshold': 0.3333,
        'fuzzy_threshold': 90,
        'default_buyer': 'shared',
    }


def main():
    # assign_buyer: shared wins at >= 1/3 share
    check('majority me', assign_buyer(0, 5, 1, 0.3333) == 'me')
    check('majority jessica', assign_buyer(0, 1, 5, 0.3333) == 'jessica')
    check('shared wins at exactly 1/3', assign_buyer(2, 4, 0, 0.3333) == 'shared')
    check('shared below 1/3 loses', assign_buyer(1, 5, 0, 0.3333) == 'me')
    check('me/jessica tie falls back to shared', assign_buyer(0, 3, 3, 0.3333) == 'shared')
    check('no counts falls back to shared', assign_buyer(0, 0, 0, 0.3333) == 'shared')

    with tempfile.TemporaryDirectory() as tmpdir:
        config = make_config(tmpdir)

        # round-trip preserves row order and values
        rules = {
            'zucchini': {'item_normalized': 'zucchini', 'assigned_buyer': 'me',
                         'source': 'auto', 'shared_count': 0.0, 'me_count': 4.0,
                         'jessica_count': 0.0, 'confidence': '1.00', 'last_seen': '2026-06-01'},
            'avocado': {'item_normalized': 'avocado', 'assigned_buyer': 'shared',
                        'source': 'auto', 'shared_count': 3.0, 'me_count': 1.0,
                        'jessica_count': 0.0, 'confidence': '0.75', 'last_seen': '2026-06-01'},
        }
        save_rules(config, rules, ['zucchini', 'avocado'])
        loaded, order = load_rules(config)
        check('round-trip keeps order', order == ['zucchini', 'avocado'], str(order))
        check('round-trip keeps values',
              loaded['avocado']['assigned_buyer'] == 'shared'
              and loaded['zucchini']['me_count'] == 4.0)

        # new keys are appended after existing ones, sorted
        rules['banan'] = dict(rules['avocado'], item_normalized='banan')
        rules['aubergine'] = dict(rules['avocado'], item_normalized='aubergine')
        save_rules(config, rules, order)
        _, order2 = load_rules(config)
        check('new keys appended sorted',
              order2 == ['zucchini', 'avocado', 'aubergine', 'banan'], str(order2))

        # manual-edit detection: user flips assigned_buyer against the counts
        with open(config['rules_csv'], newline='') as f:
            data = list(csv.DictReader(f))
        for row in data:
            if row['item_normalized'] == 'zucchini':
                row['assigned_buyer'] = 'jessica'  # disagrees with me_count=4
        with open(config['rules_csv'], 'w', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=data[0].keys())
            writer.writeheader()
            writer.writerows(data)
        loaded, _ = load_rules(config)
        check('user edit detected as manual', loaded['zucchini']['source'] == 'manual')
        check('user-edited buyer preserved', loaded['zucchini']['assigned_buyer'] == 'jessica')
        check('untouched auto rows stay auto', loaded['avocado']['source'] == 'auto')

        # prediction: exact, fuzzy, default
        buyer, kind, _ = predict('avocado', loaded, config)
        check('exact match', (buyer, kind) == ('shared', 'exact'))
        buyer, kind, matched = predict('avocadoo', loaded, config)
        check('fuzzy match', kind == 'fuzzy' and matched == 'avocado')
        buyer, kind, _ = predict('helt ny vara', loaded, config)
        check('unknown item defaults to shared', (buyer, kind) == ('shared', 'default'))

        # missing rule file loads empty
        empty, empty_order = load_rules(make_config(tmpdir + '/nope'))
        check('missing csv loads empty', empty == {} and empty_order == [])

    summary = f"{len(RESULTS) - FAILED}/{len(RESULTS)} passed"
    RESULTS.append(summary)
    out = '\n'.join(RESULTS)
    Path('debugging').mkdir(exist_ok=True)
    Path('debugging/rules_test_results.txt').write_text(out + '\n')
    print(out)
    raise SystemExit(1 if FAILED else 0)


if __name__ == '__main__':
    main()
