#!/usr/bin/env python3
"""Every open Flywheel position with its numbers and a TradingView chart link. Read-only.

Holdings come from the Bloom cloud copy (the last Robinhood sync plus trades logged since),
prices and indicators from the app's own price feed.

    /usr/bin/python3 tools/flywheel_check.py
"""
import json, os, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from robinhood_sync import api

MIN_VALUE = 1.00          # leftover dust below this is not a position


def main():
    raw = api('action=load_dca_data')
    cloud = json.loads(raw['data'])
    feed = {e['ticker']: e for e in api('action=dca_prices')['etfs']}
    synced = max((h['date'] for h in cloud['history'] if h.get('type') == 'reconcile'), default='never')
    rows = []
    for t, h in cloud['portfolios']['stocks']['holdings'].items():
        f = feed.get(t, {})
        value = h.get('shares', 0) * f.get('price', 0)
        if value < MIN_VALUE:
            continue
        cost = h.get('costBasis', 0)
        rows.append((t, value, (value - cost) / cost * 100 if cost else 0, f))
    rows.sort(key=lambda r: r[2])
    print('Holdings as of Bloom save %s (last Robinhood sync %s)\n' % (raw['lastSaved'][:16], synced[:10]))
    print('%-6s %9s %7s %6s %6s %8s  %s' % ('', 'value', 'return', 'RSI', 'DSS3D', 'vs 200d', 'chart'))
    for t, value, ret, f in rows:
        sma = f.get('sma200') or 0
        vs200 = '%+.0f%%' % ((f['price'] / sma - 1) * 100) if sma else '--'
        print('%-6s %9.2f %+6.1f%% %6s %6s %8s  https://www.tradingview.com/chart/?symbol=%s' % (
            t, value, ret, '%.0f' % f['rsi14'] if f.get('rsi14') is not None else '--',
            '%.0f' % f['dss3d'] if f.get('dss3d') is not None else '--', vs200, t.replace('-', '.')))
    print('\n%d positions, $%.2f' % (len(rows), sum(r[1] for r in rows)))


if __name__ == '__main__':
    main()
