#!/usr/bin/python3
"""Build Bloom's record: for each portfolio, the money that went in, what it's worth, and what
the same money would be worth in the S&P 500 (VOO) or in the Growth mix.

    /usr/bin/python3 tools/statements.py              # preview: tables only
    /usr/bin/python3 tools/statements.py --tickers    # also every ticker against the S&P 500
    /usr/bin/python3 tools/statements.py --apply      # store the record in Bloom's synced data

Where the trades come from, best source first:
  * Robinhood monthly statements (PDFs in sync-data/statements/): every Buy/Sell in "Account
    Activity", exact. Each month's trades must turn last month's holdings into this month's,
    for every ticker, or it stops. Text is pulled out with macOS PDFKit (tools/pdf2txt.js; the
    Homebrew pdftotext here is an Intel build that won't run).
  * After the last statement, up to a Robinhood match (sync-data/positions-YYYY-MM-DD.json):
    the change in shares is real, the days are Bloom's own entries for that ticker (the middle
    of the gap if it has none), priced at that day's close. Estimated.
  * After the last match: the trades logged in Bloom.
A ticker counts for the portfolio it sits in today (SCHD is split in Bloom's current
Cushion:Growth ratio). Prices are Yahoo daily closes; dividends are left out of the gains on
every side and reported separately (estimated from shares held on each ex-date).

"The same money" = each day's net deposits put into VOO (or the Growth mix) at that day's
close, and each day's net sale proceeds taken back out.

--apply stores the record on the latest Robinhood match entry in Bloom's history (an entry
every copy of the app already keeps and syncs), after backing up the cloud copy. The app
draws its return charts and the Flywheel optimizer from it and carries on from its own daily
snapshots after the record's last day. Run it after every bloom sync and whenever a new
statement is added. Statements and the record hold account details, so they stay in the
git-ignored sync-data/ folder and in the synced data, never in this repo.
"""
import argparse, calendar, collections, datetime, glob, json, os, re, subprocess, sys, time, urllib.request, zoneinfo

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
DIR = os.path.join(ROOT, 'sync-data', 'statements')
sys.path.insert(0, HERE)
import robinhood_sync as rs

num = lambda s: float(s.replace(',', ''))
TRADE = re.compile(r'(?:^|\s)([A-Z][A-Z.\-]{0,5}) (Margin|Cash) (Buy|Sell) (\d\d)/(\d\d)/(\d{4}) ([\d.,]+) \$([\d,]+\.\d+) \$([\d,]+\.\d+)\s*$')
HOLD = re.compile(r'(?:^|\s)([A-Z][A-Z.\-]{0,5}) (Margin|Cash) ([\d.,]+) \$([\d,]+\.\d+) \$([\d,]+\.\d+)')
SPL = re.compile(r'CUSIP: (\S+) (?:([A-Z][A-Z.\-]{0,5}) )?(Margin|Cash) SPL (\d\d)/(\d\d)/(\d{4}) ([\d.,]+)')
CUSIPSYM = re.compile(r'CUSIP: (\S+) ([A-Z][A-Z.\-]{0,5}) (Margin|Cash) ')
PERIOD = re.compile(r'(\d\d)/01/(\d{4}) to \d\d/\d\d/\d{4}')
PORTFOLIOS = ('cushion', 'growth', 'stocks')
NAMES = {'cushion': '4-Year Cushion', 'growth': 'Lifetime Growth', 'stocks': 'Flywheel'}
BENCH = 'VOO'


def to_text():
    """PDF -> sync-data/statements/txt/YYYY-MM.txt, once per PDF"""
    txt = os.path.join(DIR, 'txt')
    os.makedirs(txt, exist_ok=True)
    seen_path = os.path.join(txt, '_done.json')
    seen = json.load(open(seen_path)) if os.path.exists(seen_path) else {}
    for pdf in sorted(glob.glob(os.path.join(DIR, '*.pdf'))):
        key = '%s:%d' % (os.path.basename(pdf), os.path.getmtime(pdf))
        if key in seen and os.path.exists(os.path.join(txt, seen[key])): continue
        tmp = os.path.join(txt, '_tmp.txt')
        subprocess.run(['osascript', '-l', 'JavaScript', os.path.join(HERE, 'pdf2txt.js'), pdf, tmp], check=True, capture_output=True)
        m = PERIOD.search(open(tmp, encoding='utf-8').read())
        if not m: sys.exit('No statement period found in %s' % pdf)
        seen[key] = '%s-%s.txt' % (m.group(2), m.group(1))
        os.replace(tmp, os.path.join(txt, seen[key]))
    json.dump(seen, open(seen_path, 'w'))


def parse():
    files = sorted(glob.glob(os.path.join(DIR, 'txt', '20??-??.txt')))
    cusip = {}
    for f in files:
        for line in open(f, encoding='utf-8'):
            m = CUSIPSYM.search(line)
            if m: cusip[m.group(1)] = m.group(2)
    months = {}
    for f in files:
        state, trades, hold, splits = None, [], collections.defaultdict(float), []
        for line in open(f, encoding='utf-8'):
            line = line.rstrip('\n')
            if 'Portfolio Summary' in line: state = 'hold'
            elif 'Account Activity' in line: state = 'act'
            elif 'Executed Trades Pending Settlement' in line: state = 'pending'   # shows up again next month
            if state == 'hold':
                m = HOLD.search(line)
                if m: hold[m.group(1)] += num(m.group(3))
            elif state == 'act':
                m = TRADE.search(line)
                if m:
                    sym, _, side, mo, dd, yy, q, p, amt = m.groups()
                    trades.append({'sym': sym, 'side': side, 'date': '%s-%s-%s' % (yy, mo, dd), 'qty': num(q), 'price': num(p), 'amount': num(amt)})
                    continue
                m = SPL.search(line)
                if m:
                    splits.append({'sym': m.group(2) or cusip.get(m.group(1)), 'date': '%s-%s-%s' % (m.group(6), m.group(4), m.group(5)), 'qty': num(m.group(7))})
                elif re.search(r'(Margin|Cash) (Buy|Sell) \d\d/', line):
                    sys.exit('Could not read this trade line in %s:\n  %s' % (f, line))
        months[os.path.basename(f)[:7]] = {'trades': trades, 'hold': dict(hold), 'splits': splits}
    prev = collections.defaultdict(float)
    for key in sorted(months):
        mo = months[key]
        calc = collections.defaultdict(float, prev)
        for t in mo['trades']: calc[t['sym']] += t['qty'] if t['side'] == 'Buy' else -t['qty']
        for s in mo['splits']: calc[s['sym']] += s['qty']
        bad = {t: (round(calc.get(t, 0), 6), mo['hold'].get(t, 0)) for t in set(calc) | set(mo['hold'])
               if abs(calc.get(t, 0) - mo['hold'].get(t, 0)) > 2e-6}
        print('%s: %d trades, month-end shares %s' % (key, len(mo['trades']), 'all match' if not bad else 'DO NOT MATCH %s' % bad))
        if bad: sys.exit('Stopping: a trade is missing or misread.')
        prev = collections.defaultdict(float, mo['hold'])
    return months


def yahoo(tickers, start):
    """Daily closes (split-adjusted), dividends per share by ex-date, and splits, per ticker"""
    p1 = int(time.mktime(time.strptime(start, '%Y-%m-%d'))) - 86400 * 7
    out = {}
    for t in tickers:
        url = 'https://query1.finance.yahoo.com/v8/finance/chart/%s?period1=%d&period2=%d&interval=1d&events=div,splits' % (t, p1, int(time.time()))
        r = json.load(urllib.request.urlopen(urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'}), timeout=30))['chart']['result'][0]
        off = r['meta'].get('gmtoffset', 0)
        day = lambda ts: time.strftime('%Y-%m-%d', time.gmtime(int(ts) + off))
        ev = r.get('events') or {}
        out[t] = {'close': {day(ts): c for ts, c in zip(r['timestamp'], r['indicators']['quote'][0]['close']) if c},
                  'div': {day(k): v['amount'] for k, v in (ev.get('dividends') or {}).items()},
                  'split': {day(k): v['numerator'] / v['denominator'] for k, v in (ev.get('splits') or {}).items()}}
        time.sleep(0.1)
    return out


def build(cloud, a):
    months = parse()
    order = rs.order_lists()
    gw, gbase = rs.growth_mix()
    hold_now = {pk: {t: h for t, h in cloud['portfolios'][pk]['holdings'].items() if h.get('shares', 0) > 0} for pk in PORTFOLIOS}
    sc, sg = (hold_now[pk].get('SCHD', {}).get('shares', 0) for pk in ('cushion', 'growth'))
    schd_c = sc / (sc + sg) if sc + sg > 0 else 0.5
    unfiled = set()
    def parts(t):
        if t == 'SCHD': return [('cushion', schd_c), ('growth', 1 - schd_c)]
        pk = next((k for k in PORTFOLIOS if t in order[k]), None)
        if not pk: unfiled.add(t)
        return [(pk or 'stocks', 1.0)]

    last = max(months)
    stmt_end = '%s-%02d' % (last, calendar.monthrange(int(last[:4]), int(last[5:]))[1])
    history = sorted(cloud.get('history') or [], key=lambda h: h.get('date', ''))
    recons = [h for h in history if h.get('type') == 'reconcile']
    pos_files = sorted(f for f in glob.glob(os.path.join(rs.DATA_DIR, 'positions-*.json')) if os.path.basename(f)[10:20] > stmt_end)

    tickers = {t['sym'] for m in months.values() for t in m['trades']} | set(gw) | {BENCH}
    for f in pos_files: tickers |= {p[0] for p in json.load(open(f))['positions']}
    for pk in PORTFOLIOS: tickers |= set(hold_now[pk])
    start = min(t['date'] for m in months.values() for t in m['trades'])
    Y = yahoo(sorted(tickers), start)

    # Trading days, up to the last finished one
    ny = datetime.datetime.now(zoneinfo.ZoneInfo('America/New_York'))
    days = sorted(d for d in Y[BENCH]['close'] if d >= start)
    if days[-1] == ny.strftime('%Y-%m-%d') and (ny.hour, ny.minute) < (16, 15): days.pop()
    if a.end: days = [d for d in days if d <= a.end]
    through = days[-1]
    tday = lambda d: next((x for x in days if x >= d), None)        # None: after the record's last day
    def close(t, d):
        c = Y[t]['close']
        if d in c: return c[d]
        before = [x for x in c if x <= d]
        if not before: sys.exit('No price for %s on or before %s' % (t, d))
        return c[max(before)]
    # Shares in today's units: Yahoo's closes are already adjusted for later splits
    def unit(t, d):
        f = 1.0
        for sd, r in Y[t]['split'].items():
            if sd > d: f *= r
        return f

    # ---- every trade: (day, portfolio, ticker, shares, dollars), buys positive
    ev = []
    for key in sorted(months):
        for t in months[key]['trades']:
            s = 1 if t['side'] == 'Buy' else -1
            for pk, f in parts(t['sym']):
                ev.append((tday(t['date']), pk, t['sym'], s * t['qty'] * unit(t['sym'], t['date']) * f, s * t['amount'] * f))
    held = {t: q * unit(t, stmt_end) for t, q in months[last]['hold'].items()}
    after, guessed = stmt_end + 'T23:59:59.999Z', []
    for f in pos_files:
        fdate = os.path.basename(f)[10:20]
        pos = {p[0]: p[1] for p in json.load(open(f))['positions']}
        near = [h for h in recons if abs((datetime.date.fromisoformat(h['date'][:10]) - datetime.date.fromisoformat(fdate)).days) <= 1]
        until = near[-1]['date'] if near else fdate + 'T23:59:59.999Z'
        gap = [h for h in history if after < h.get('date', '') < until]
        gap_days = [d for d in days if after[:10] < d <= until[:10]]
        if not gap_days: sys.exit('No trading days between %s and the match on %s' % (after[:10], fdate))
        mid = gap_days[len(gap_days) // 2]
        for t in sorted(set(held) | set(pos)):
            delta = pos.get(t, 0) - held.get(t, 0)
            if abs(delta) < 1e-6: continue
            w = collections.defaultdict(float)
            for h in gap:
                d = tday(h['date'][:10])
                if not d or d > gap_days[-1]: d = gap_days[-1]
                if delta > 0 and h.get('type') == 'dca':
                    for x in h.get('allocations') or []:
                        if x['ticker'] == t: w[d] += x.get('shares') or 0
                if delta < 0 and h.get('type') == 'sell' and h.get('ticker') == t: w[d] += h.get('shares') or 0
            if not sum(w.values()):
                w = {mid: 1.0}
                guessed.append('%s %+.4f (%+.2f)' % (t, delta, delta * close(t, mid)))
            tot = sum(w.values())
            for d, x in w.items():
                q = delta * x / tot
                for pk, fr in parts(t): ev.append((d, pk, t, q * fr, q * close(t, d) * fr))
        held, after = pos, until
    for h in history:                                   # logged in Bloom since
        if not h.get('date', '') > after or h.get('portfolio') not in PORTFOLIOS: continue
        d = tday(h['date'][:10])
        if h.get('type') == 'sell':
            ev.append((d, h['portfolio'], h['ticker'], -(h.get('shares') or 0), -(h.get('totalAmount') or 0)))
        elif h.get('type') == 'dca':
            for x in h.get('allocations') or []:
                ev.append((d, h['portfolio'], x['ticker'], x.get('shares') or 0, x.get('amount') or 0))
    if unfiled: print('Not in any Bloom list any more, counted as Flywheel: %s' % ', '.join(sorted(unfiled)))
    if guessed: print('No Bloom entry for these moves between %s and the match, so they sit on the middle day: %s' % (stmt_end, '; '.join(guessed)))

    nav = lambda d: sum(w * close(g, d) / gbase[g] for g, w in gw.items())
    spx = [round(close(BENCH, d), 2) for d in days]
    gro = [round(nav(d), 5) for d in days]
    on = collections.defaultdict(list)
    for e in ev: on[e[0]].append(e)

    def replay(pick):
        """Daily worth and net money in for the trades pick(pk, ticker) selects, plus both shadows"""
        sh = collections.defaultdict(float)
        cf = div = 0.0
        S, units_div = {'spx': 0.0, 'gro': 0.0}, {'spx': 0.0}
        v, ins, shadow = [], [], {'spx': [], 'gro': []}
        for i, d in enumerate(days):
            div += sum(q * Y[t]['div'][d] for t, q in sh.items() if q > 1e-9 and d in Y[t]['div'])
            flow = 0.0
            for (_, pk, t, q, amt) in on.get(d, []):
                if pick(pk, t): sh[t] += q; flow += amt
            for k, px in (('spx', spx), ('gro', gro)):
                S[k] = (S[k] * px[i] / px[i - 1] if i else 0.0) + flow
                shadow[k].append(S[k])
            cf += flow
            v.append(sum(q * close(t, d) for t, q in sh.items() if q > 1e-9)); ins.append(cf)
        late = collections.defaultdict(float)               # logged for a day after the record ends
        for (_, pk, t, q, amt) in on.get(None, []):
            if pick(pk, t): late[t] += q
        return {'v': v, 'in': ins, 'spx': shadow['spx'], 'gro': shadow['gro'], 'div': div, 'shares': dict(sh), 'late': dict(late)}

    out = {pk: replay(lambda p, t, pk=pk: p == pk) for pk in PORTFOLIOS}
    # Bloom's holdings now against the rebuilt ones; k ties the app's cost basis to money in
    problems = []
    for pk in PORTFOLIOS:
        r = out[pk]
        mine = collections.defaultdict(float, r['shares'])
        for t, q in r['late'].items(): mine[t] += q
        for t in sorted(set(mine) | set(hold_now[pk])):
            x, y = mine.get(t, 0), hold_now[pk].get(t, {}).get('shares', 0)
            if abs(x - y) > 6e-3:      # Robinhood's positions list shows bigger holdings to 2 decimals
                problems.append('%s %s: rebuilt %.4f, Bloom %.4f' % (pk, t, x, y))
        v_app = sum(h['shares'] * close(t, through) for t, h in hold_now[pk].items())
        cost = sum(h.get('costBasis', 0) for h in hold_now[pk].values())
        r['delta'] = v_app - r['v'][-1]
        r['k'] = cost - r['in'][-1] - r['delta']
        r['cost'] = cost
    print('Bloom holdings vs rebuilt: %s' % ('all match' if not problems else '; '.join(problems)))
    return {'days': days, 'through': through, 'exact': stmt_end, 'spx': spx, 'gro': gro, 'p': out,
            'replay': replay, 'events': ev, 'matches': [os.path.basename(f)[10:20] for f in pos_files], 'close': close}


def report(b, tickers):
    days = b['days']
    span = '%s to %s' % (days[0], b['through'])
    print('\n%s. Exact through %s (statements)%s.' % (span, b['exact'],
          ', then the %s match and Bloom\'s log' % ' / '.join(b['matches']) if b['matches'] else ', then Bloom\'s log'))
    print('\n%-16s %10s %10s %9s %7s | %12s %9s | %12s %9s | %9s' %
          ('', 'money in', 'worth', 'gain', '', 'in S&P 500', 'vs S&P', 'in Growth', 'vs Growth', 'dividends'))
    rows = [(NAMES[pk], b['p'][pk]) for pk in PORTFOLIOS] + [('Everything', b['replay'](lambda p, t: True))]
    for name, r in rows:
        v, cf = r['v'][-1], r['in'][-1]
        print('%-16s %10.2f %10.2f %+9.2f %+6.1f%% | %12.2f %+9.2f | %12.2f %+9.2f | %9.2f' %
              (name, cf, v, v - cf, 100 * (v - cf) / cf if cf else 0, r['spx'][-1], v - r['spx'][-1], r['gro'][-1], v - r['gro'][-1], r['div']))
    f = b['p']['stocks']
    print('\nFlywheel: open positions %+.2f, banked from sales %+.2f.' % (f['v'][-1] + f['delta'] - f['cost'], f['k']))
    if tickers:
        for pk in PORTFOLIOS:
            print('\n%s by ticker: net money in, worth, gain, same money in the S&P 500, difference' % NAMES[pk])
            rows = []
            for t in sorted({e[2] for e in b['events'] if e[1] == pk}):
                r = b['replay'](lambda p, x, t=t, pk=pk: p == pk and x == t)
                rows.append((t, r['in'][-1], r['v'][-1], r['v'][-1] - r['in'][-1], r['spx'][-1] - r['in'][-1], r['v'][-1] - r['spx'][-1]))
            for r in sorted(rows, key=lambda r: r[5]): print('  %-6s %9.2f %9.2f %+9.2f %+9.2f %+9.2f' % r)
            print('  ahead of the S&P 500: %d of %d' % (sum(1 for r in rows if r[5] > 0), len(rows)))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--end', default=None, help='last day of the record (default: the last finished trading day)')
    ap.add_argument('--tickers', action='store_true', help='also print every ticker against the S&P 500')
    ap.add_argument('--apply', action='store_true', help='store the record in Bloom\'s synced data')
    a = ap.parse_args()
    to_text()
    raw = rs.api('action=load_dca_data')
    cloud = json.loads(raw['data'])
    b = build(cloud, a)
    report(b, a.tickers)
    r2 = lambda xs: [round(x, 2) for x in xs]
    now = int(time.time() * 1000)
    record = {'at': rs.iso(now), 'through': b['through'], 'exact': b['exact'], 'days': b['days'], 'spx': b['spx'], 'gro': b['gro'],
              'p': {pk: {'v': r2(r['v']), 'in': r2(r['in']), 'k': round(r['k'], 2), 'div': round(r['div'], 2)} for pk, r in b['p'].items()}}
    json.dump(record, open(os.path.join(DIR, 'record.json'), 'w'))
    if not a.apply:
        print('\nPreview only. Re-run with --apply to store it in Bloom.'); return
    entry = next((h for h in sorted(cloud.get('history') or [], key=lambda h: h.get('date', ''), reverse=True) if h.get('type') == 'reconcile'), None)
    if not entry: sys.exit('Bloom has no Robinhood match yet — run a bloom sync first.')
    for h in cloud['history']: h.pop('record', None)
    entry['record'] = record
    cloud['lastModified'] = now
    stamp = time.strftime('%Y-%m-%d_%H%M')
    json.dump(raw, open(os.path.join(rs.DATA_DIR, 'cloud-before-record-%s.json' % stamp), 'w'))
    blob = json.dumps(cloud, separators=(',', ':'))
    print('\ndry run:', rs.upload(blob, True))
    print('write:  ', rs.upload(blob, False))
    back = rs.api('action=load_dca_data')
    print('read back matches:', back['data'] == blob, '| lastSaved', back['lastSaved'])


if __name__ == '__main__':
    main()
