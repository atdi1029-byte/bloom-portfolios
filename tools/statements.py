#!/usr/bin/python3
"""Rebuild the Flywheel's history from Robinhood monthly statements (PDF).

    /usr/bin/python3 tools/statements.py                 # read every PDF in sync-data/statements/
    /usr/bin/python3 tools/statements.py --end 2026-09-28

Steps:
  * text is pulled out of each PDF with macOS PDFKit (tools/pdf2txt.js; the Homebrew
    pdftotext here is an Intel build that won't run), into sync-data/statements/txt/YYYY-MM.txt
  * every Buy/Sell in "Account Activity" is read ("Executed Trades Pending Settlement" is
    skipped: those trades show up again in the next month's activity with their trade date)
  * each month's trades must turn last month's holdings into this month's, for every ticker,
    or it stops
  * Flywheel trades (tickers in the app's Flywheel list) are replayed day by day against the
    same dollars in the Growth mix, on Yahoo closing prices, prices only (dividends left out)
Output: a month table, and sync-data/statements/flywheel_rebuild.json.
Statements hold account details, so they stay in the git-ignored sync-data/ folder.
"""
import argparse, collections, glob, json, math, os, re, subprocess, sys, time, urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
DIR = os.path.join(ROOT, 'sync-data', 'statements')
num = lambda s: float(s.replace(',', ''))
TRADE = re.compile(r'(?:^|\s)([A-Z][A-Z.\-]{0,5}) (Margin|Cash) (Buy|Sell) (\d\d)/(\d\d)/(\d{4}) ([\d.,]+) \$([\d,]+\.\d+) \$([\d,]+\.\d+)\s*$')
HOLD = re.compile(r'(?:^|\s)([A-Z][A-Z.\-]{0,5}) (Margin|Cash) ([\d.,]+) \$([\d,]+\.\d+) \$([\d,]+\.\d+)')
SPL = re.compile(r'CUSIP: (\S+) (?:([A-Z][A-Z.\-]{0,5}) )?(Margin|Cash) SPL (\d\d)/(\d\d)/(\d{4}) ([\d.,]+)')
CUSIPSYM = re.compile(r'CUSIP: (\S+) ([A-Z][A-Z.\-]{0,5}) (Margin|Cash) ')
PERIOD = re.compile(r'(\d\d)/01/(\d{4}) to \d\d/\d\d/\d{4}')


def to_text():
    os.makedirs(os.path.join(DIR, 'txt'), exist_ok=True)
    for pdf in glob.glob(os.path.join(DIR, '*.pdf')):
        tmp = os.path.join(DIR, 'txt', '_tmp.txt')
        subprocess.run(['osascript', '-l', 'JavaScript', os.path.join(HERE, 'pdf2txt.js'), pdf, tmp], check=True, capture_output=True)
        m = PERIOD.search(open(tmp, encoding='utf-8').read())
        if not m: sys.exit('No statement period found in %s' % pdf)
        os.replace(tmp, os.path.join(DIR, 'txt', '%s-%s.txt' % (m.group(2), m.group(1))))


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
            elif 'Executed Trades Pending Settlement' in line: state = 'pending'
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


def closes(tickers, start):
    p1 = int(time.mktime(time.strptime(start, '%Y-%m-%d'))) - 86400 * 5
    out = {}
    for t in tickers:
        url = 'https://query1.finance.yahoo.com/v8/finance/chart/%s?period1=%d&period2=%d&interval=1d' % (t, p1, int(time.time()))
        r = json.load(urllib.request.urlopen(urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'}), timeout=30))['chart']['result'][0]
        out[t] = {time.strftime('%Y-%m-%d', time.gmtime(ts)): c for ts, c in zip(r['timestamp'], r['indicators']['quote'][0]['close']) if c}
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--end', default=None, help='last day to rebuild (default: last statement month end)')
    a = ap.parse_args()
    to_text()
    months = parse()
    html = open(os.path.join(ROOT, 'index.html'), encoding='utf-8').read()
    lists = {k: re.findall(r"'([A-Z.\-]+)'", m) for k, m in re.findall(r"(\w+): \{\n        name: '[^']*',[\s\S]*?order: \[([^\]]*)\]", html)}
    fly = set(lists['stocks'])
    growth = {t: int(p) for t, p in re.findall(r"(\w+):\s*\{ pct: (\d+)", html[html.index("growth: {"):html.index("ira: {")])}
    trades = sorted([t for m in months.values() for t in m['trades'] if t['sym'] in fly], key=lambda t: t['date'])
    start = trades[0]['date']
    end = a.end or '%s-31' % max(months)
    px = closes(sorted({t['sym'] for t in trades} | set(growth)), start)
    days = sorted(d for d in px['VTI'] if start <= d <= end)
    on = collections.defaultdict(list)
    for t in trades:
        if t['date'] <= end: on[next(d for d in days if d >= t['date'])].append(t)
    shares, units, rows = collections.defaultdict(float), collections.defaultdict(float), []
    cf, prev, L, sx2, n, peak, worst = 0.0, None, 0.0, 0.0, 0, 0.0, (0.0, None)
    tot = sum(growth.values())
    for d in days:
        flow = 0.0
        for t in on.get(d, []):
            if t['side'] == 'Buy':
                shares[t['sym']] += t['qty']; flow += t['amount']
                for g, w in growth.items(): units[g] += t['amount'] * w / tot / px[g][d]
            else:
                shares[t['sym']] -= t['qty']; flow -= t['amount']
                sv = sum(u * px[g][d] for g, u in units.items())
                for g in units: units[g] *= max(0.0, 1 - t['amount'] / sv)
        F = sum(q * px[s][d] for s, q in shares.items() if q > 1e-9)
        Sv = sum(u * px[g][d] for g, u in units.items())
        cf += flow
        if prev:
            x = math.log(((F - flow) / prev[0]) / ((Sv - flow) / prev[1]))
            L += x; sx2 += x * x; n += 1; peak = max(peak, L)
            if 1 - math.exp(L - peak) > worst[0]: worst = (1 - math.exp(L - peak), d)
        rows.append({'d': d, 'f': round(F, 2), 's': round(Sv, 2), 'cf': round(cf, 2)})
        prev = (F, Sv)
    print('\n%-8s %9s %10s %12s %9s %7s' % ('month', 'money in', 'Flywheel', 'same in Grw', 'diff $', 'diff'))
    last_cf, Lrun = 0.0, 0.0
    for mo in sorted({r['d'][:7] for r in rows}):
        r = [x for x in rows if x['d'].startswith(mo)][-1]
        print('%-8s %9.2f %10.2f %12.2f %+9.2f %+6.1f%%' % (mo, r['cf'] - last_cf, r['f'], r['s'], r['f'] - r['s'], 100 * (r['f'] / r['s'] - 1)))
        last_cf = r['cf']
    print('\nPer dollar, %s to %s: Flywheel %+.1f%% vs Growth. Swing vs Growth %.0f%% a year. Worst stretch %.1f%% behind (%s).'
          % (days[0], days[-1], 100 * (math.exp(L) - 1), 100 * math.sqrt(sx2 / n * 252), 100 * worst[0], worst[1]))
    json.dump({'rows': rows, 'trades': trades, 'shares_end': {s: q for s, q in shares.items() if q > 1e-9}},
              open(os.path.join(DIR, 'flywheel_rebuild.json'), 'w'))


if __name__ == '__main__':
    main()
