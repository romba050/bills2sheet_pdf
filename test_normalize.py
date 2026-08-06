"""Unit tests for buyer_classifier.normalize, using real item strings from the
receipt sheets. Run with: uv run python test_normalize.py
Writes results to debugging/normalize_test_results.txt.
"""

from pathlib import Path

from buyer_classifier import normalize

# (input, expected) pairs sampled from actual Food Expenses sheets
CASES = [
    # plain items
    ('Avocado', 'avocado'),
    ('Banan Eko', 'banan eko'),
    ('Fish & crisp', 'fish & crisp'),
    ('Kvarg Mango/Pass', 'kvarg mango/pass'),
    # leading receipt markers
    ('*Avfallspåse 30L', 'avfallspåse'),
    ('*Vispgrädde 36%', 'vispgrädde'),
    # sizes, quantities, percentages
    ('Avfallspåse 30L', 'avfallspåse'),
    ('Näsdukar 10p', 'näsdukar'),
    ('Pepsi Max 10p', 'pepsi max'),
    ('Laxfilé 4x100', 'laxfilé'),
    ('Kvarg natur 0,3%', 'kvarg natur'),
    ('Gräddfil 12%', 'gräddfil'),
    ('Carlsberg 0,5% EKO', 'carlsberg eko'),
    ('Roslagsägg EKO 12p', 'roslagsägg eko'),
    ('Ekologiska 12-pack', 'ekologiska'),
    ('Gaständare 1-pack', 'gaständare'),
    ('Gaffel trä 25p', 'gaffel trä'),
    ('Jordgubbar SE 500g', 'jordgubbar se'),
    ('Blandfärs 50/50', 'blandfärs'),
    ('Heineken 3,5% B 6p', 'heineken b'),
    # discount lines must share the key of the discounted item
    ('Jordgubbar 69kr/st', 'jordgubbar'),
    ('Jordgubbar', 'jordgubbar'),
    ('Ekologiska 35kr/st', 'ekologiska'),
    ('Avfallspåse 2f45kr', 'avfallspåse'),
    ('Halloumi 2f57kr', 'halloumi'),
    ('Wasa Sand 2f18.93kr', 'wasa sand'),
    ('Kokosmjölk 2f25kr', 'kokosmjölk'),
    # weight-priced items with composite tokens
    ('Lök röd 0,145kg*18,88Kr/kg', 'lök röd'),
    # coupons, deposits
    ('Kuponger 100 100kr', 'kuponger'),
    ('Pant', 'pant'),
    ('Pantretur, låg moms', 'pantretur, låg moms'),
    ('Rabatt: Kupong', 'rabatt kupong'),
    # whitespace and case
    ('  StjärnMix   ', 'stjärnmix'),
    ('Solsken ', 'solsken'),
    # all-digit fallback keeps something usable
    ('6 cocktail drinks Midsommar', 'cocktail drinks midsommar'),
]

# pairs that must NOT collapse to the same key
DISTINCT = [
    ('Kvarg blåbär utan', 'Kvarg Mango/Pass'),
    ('Pesto Genovese', 'Pesto Rosso'),
    ('Paprika Röd', 'Peppar Röd'),
]


def main():
    lines, failed = [], 0
    for raw, expected in CASES:
        got = normalize(raw)
        ok = got == expected
        failed += not ok
        lines.append(f"{'PASS' if ok else 'FAIL'} normalize({raw!r}) = {got!r}"
                     + ('' if ok else f" (expected {expected!r})"))
    for a, b in DISTINCT:
        ok = normalize(a) != normalize(b)
        failed += not ok
        lines.append(f"{'PASS' if ok else 'FAIL'} distinct: {a!r} vs {b!r}")
    summary = f"{len(CASES) + len(DISTINCT) - failed}/{len(CASES) + len(DISTINCT)} passed"
    lines.append(summary)
    out = '\n'.join(lines)
    Path('debugging').mkdir(exist_ok=True)
    Path('debugging/normalize_test_results.txt').write_text(out + '\n')
    print(out)
    raise SystemExit(1 if failed else 0)


if __name__ == '__main__':
    main()
