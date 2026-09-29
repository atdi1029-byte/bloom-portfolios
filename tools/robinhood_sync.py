#!/usr/bin/python3
"""Set Bloom's holdings to match Robinhood ("bloom sync").

Input is a positions file transcribed from Robinhood's positions list: one row per holding
with shares, total return ($) and total return (%). Cost basis is derived from the return
(or from value minus return when the % is too small to be precise).

    /usr/bin/python3 tools/robinhood_sync.py sync-data/positions-2026-10-12.json              # preview only
    /usr/bin/python3 tools/robinhood_sync.py sync-data/positions-2026-10-12.json --apply      # write to the cloud
    ... --sold TSLA IBIT     record full sells (at today's price) for positions that were sold

positions file:  {"positions": [["SCHG", 25.82, 53.22, 6.06], ["GLD", 1.9, -72.61, -9.15], ...]}

What it does:
  * every held ticker is set to Robinhood's shares and cost; anything Bloom holds that isn't
    in the list is set to 0 (and gets a sell entry if named in --sold)
  * SCHD sits in two portfolios; the Robinhood total is split in Bloom's current Cushion:Growth ratio
  * a 'reconcile' history entry is added: the app treats trades logged before it as already
    counted, so a device that syncs later can't add them twice
  * today's value snapshot is redone with the new holdings
  * the cloud copy is backed up to sync-data/ first, uploaded (dry run, then real), and read back
The IRA isn't touched: it isn't in the Robinhood account these screenshots come from.
"""
import argparse, base64, copy, gzip, json, os, re, sys, time, urllib.parse, urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
DATA_DIR = os.path.join(ROOT, 'sync-data')          # git-ignored: holdings stay off GitHub
API = re.search(r"const API_BASE = '([^']+)'", open(os.path.join(ROOT, 'index.html'), encoding='utf-8').read()).group(1)
SYNC_PORTFOLIOS = ('cushion', 'growth', 'stocks')


def api(query):
    with urllib.request.urlopen(API + '?' + query, timeout=90) as r:
        return json.loads(r.read().decode())


def order_lists():
    html = open(os.path.join(ROOT, 'index.html'), encoding='utf-8').read()
    found = {k: re.findall(r"'([A-Z.\-]+)'", m) for k, m in
             re.findall(r"(\w+): \{\n        name: '[^']*',[\s\S]*?order: \[([^\]]*)\]", html)}
    assert set(found) == {'cushion', 'growth', 'ira', 'stocks'}, 'could not read portfolio order lists'
    return found


def iso(ms):
    return time.strftime('%Y-%m-%dT%H:%M:%S', time.gmtime(ms / 1000)) + '.%03dZ' % (ms % 1000)


def build(cloud, positions, prices, sold):
    order = order_lists()
    d = copy.deepcopy(cloud)
    d.pop('cash', None)                               # cash tracking was removed from the app
    before = {pk: {t: dict(h) for t, h in p['holdings'].items()} for pk, p in cloud['portfolios'].items()}
    now = int(time.time() * 1000)

    target = {}
    for t, shares, ret, pct in positions:
        if t not in prices:
            sys.exit('No price for %s — is it in DCA_TICKERS and a Bloom order list?' % t)
        value = shares * prices[t]
        cost = ret / (pct / 100) if abs(pct) >= 1.5 else value - ret
        target[t] = (shares, round(cost, 2))

    # Sells for positions that were closed, at today's price
    sells = []
    for i, t in enumerate(sold):
        pk = next((k for k in SYNC_PORTFOLIOS if before[k].get(t, {}).get('shares', 0) > 0), None)
        if not pk:
            print('  (skip sell %s: Bloom has no shares)' % t); continue
        h = before[pk][t]
        amount = round(h['shares'] * prices[t], 2)
        sell = {'id': 'sell_%d' % (now - 3000 + i * 100), 'date': iso(now - 3000 + i * 100), 'portfolio': pk,
                'totalAmount': amount, 'type': 'sell', 'ticker': t, 'shares': h['shares'], 'price': prices[t],
                'profit': round(amount - h.get('costBasis', 0), 2)}
        if pk == 'stocks':   # Growth prices at the sale, for the Flywheel optimizer's shadow
            bench = {g: prices[g] for g in order['growth'] if prices.get(g)}
            if bench:
                sell['bench'] = bench
        sells.append(sell)
    d['history'] += sells

    # Where each ticker lives
    home = {}
    for t in target:
        owners = [pk for pk in SYNC_PORTFOLIOS if t in order[pk]]
        if t == 'SCHD':
            owners = ['cushion', 'growth']
        if not owners:
            sys.exit('%s is not in any Bloom portfolio list — add it to the Flywheel first' % t)
        if len(owners) > 1 and t != 'SCHD':
            sys.exit('%s is in several portfolios %s — decide the split first' % (t, owners))
        home[t] = owners

    for pk in SYNC_PORTFOLIOS:
        for t, h in d['portfolios'][pk]['holdings'].items():
            if h.get('shares', 0) > 0 and t not in target:
                h['shares'] = 0; h['costBasis'] = 0
    for t, (shares, cost) in target.items():
        if home[t] == ['cushion', 'growth']:
            c0 = before['cushion'].get(t, {}).get('shares', 0); g0 = before['growth'].get(t, {}).get('shares', 0)
            fc = c0 / (c0 + g0) if c0 + g0 > 0 else 1.0
            cs, cc = round(shares * fc, 4), round(cost * fc, 2)
            parts = {'cushion': (cs, cc), 'growth': (round(shares - cs, 4), round(cost - cc, 2))}
        else:
            parts = {home[t][0]: (shares, cost)}
        for pk, (s_, c_) in parts.items():
            d['portfolios'][pk]['holdings'][t] = {'shares': s_, 'costBasis': c_}

    changes = []
    for pk in SYNC_PORTFOLIOS:
        for t in sorted(set(before[pk]) | set(d['portfolios'][pk]['holdings'])):
            a = before[pk].get(t, {}).get('shares', 0); b = d['portfolios'][pk]['holdings'].get(t, {}).get('shares', 0)
            if abs(a - b) > 1e-6:
                changes.append([pk, t, round(a, 6), b])
    note = 'Holdings set to match Robinhood (%d changes)' % len(changes)
    if sells:
        note += '; sold all ' + ' and '.join(x['ticker'] for x in sells)
    d['history'].append({'id': 'reconcile_%d' % now, 'date': iso(now), 'portfolio': 'all', 'type': 'reconcile',
                         'totalAmount': 0, 'note': note, 'changes': changes})

    today = iso(now)[:10]
    snap = {'date': today, 'portfolios': {}}
    for pk, p in d['portfolios'].items():
        snap['portfolios'][pk] = {'value': round(sum(h.get('shares', 0) * prices.get(t, 0) for t, h in p['holdings'].items()), 2),
                                  'cost': round(sum(h.get('costBasis', 0) for h in p['holdings'].values()), 2)}
    d['valueSnapshots'] = [s for s in d.get('valueSnapshots', []) if s['date'] != today] + [snap]
    d['lastModified'] = now
    d.setdefault('deletedIds', [])
    return d, before, changes, sells


def summarize(before, after, prices, changes, sells):
    for pk in SYNC_PORTFOLIOS:
        def tot(h):
            v = sum(x.get('shares', 0) * prices.get(t, 0) for t, x in h.items()); c = sum(x.get('costBasis', 0) for x in h.values())
            return v, c
        v0, c0 = tot(before[pk]); v1, c1 = tot(after['portfolios'][pk]['holdings'])
        print('  %-8s value %10.2f -> %10.2f   cost %10.2f -> %10.2f   return %+9.2f -> %+9.2f' % (pk, v0, v1, c0, c1, v0 - c0, v1 - c1))
    for x in sells:
        print('  sell %-5s %.6f sh @ %.2f = $%.2f' % (x['ticker'], x['shares'], x['price'], x['totalAmount']))
    print('  %d holdings change:' % len(changes))
    for pk, t, a, b in changes:
        print('    %-8s %-5s %12.6f -> %12.6f' % (pk, t, a, b))


def upload(blob, dry):
    b64 = base64.b64encode(gzip.compress(blob.encode())).decode()
    chunks = [b64[i:i + 5000] for i in range(0, len(b64), 5000)]
    u = ('rhdry' if dry else 'rhsync') + str(int(time.time()))
    for i, c in enumerate(chunks):
        r = api('action=dca_save_chunk&u=%s&i=%d&cd=%s' % (u, i, urllib.parse.quote(c, safe='')))
        if r.get('status') != 'ok':
            sys.exit('chunk %d failed: %s' % (i, r))
    for _ in range(3):
        r = api('action=dca_save_done&u=%s&n=%d&z=1%s' % (u, len(chunks), '&dry=1' if dry else ''))
        if r.get('status') == 'ok':
            return r
        time.sleep(2)
    sys.exit('save failed: %s' % r)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('positions')
    ap.add_argument('--sold', nargs='*', default=[])
    ap.add_argument('--apply', action='store_true')
    a = ap.parse_args()
    positions = json.load(open(a.positions))['positions']
    raw = api('action=load_dca_data')
    cloud = json.loads(raw['data'])
    prices = {e['ticker']: e['price'] for e in api('action=dca_prices')['etfs'] if e.get('price')}
    after, before, changes, sells = build(cloud, positions, prices, a.sold)
    print('Cloud copy last saved %s' % raw['lastSaved'])
    summarize(before, after, prices, changes, sells)
    if not a.apply:
        print('\nPreview only. Re-run with --apply to write it.'); return
    os.makedirs(DATA_DIR, exist_ok=True)
    stamp = time.strftime('%Y-%m-%d_%H%M')
    json.dump(raw, open(os.path.join(DATA_DIR, 'cloud-before-sync-%s.json' % stamp), 'w'))
    blob = json.dumps(after, separators=(',', ':'))
    print('dry run:', upload(blob, True))
    print('write:  ', upload(blob, False))
    back = api('action=load_dca_data')
    print('read back matches:', back['data'] == blob, '| lastSaved', back['lastSaved'])


if __name__ == '__main__':
    main()
