#!/usr/bin/env python3
"""Weekly Flywheel check: which open positions have left the buy zone. Read-only.

Holdings come from the Bloom cloud copy (the last Robinhood sync plus trades logged since),
prices and indicators from the app's own price feed. The zone is the app's own Cheap / Fair /
Expensive label (calcValuationScore in index.html, mirrored below), the same score the
splitter uses to pick the Flywheel buys. Cheap (green) = still a buy, nothing to check.

Writes sync-data/flywheel_check.html (git-ignored, holdings stay off GitHub) and remembers
each ticker's zone in sync-data/flywheel_zones.json, so next week's page can say which
ones just left green.

    /usr/bin/python3 tools/flywheel_check.py          # print + write the page
    /usr/bin/python3 tools/flywheel_check.py --open   # ... and open it
"""
import html, json, os, re, subprocess, sys, time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from robinhood_sync import api, ROOT, DATA_DIR

MIN_VALUE = 1.00          # leftover dust below this is not a position
PAGE = os.path.join(DATA_DIR, 'flywheel_check.html')
ZONES = os.path.join(DATA_DIR, 'flywheel_zones.json')


def valuation_score(p):
    """calcValuationScore from index.html: higher = cheaper; 0 = no data"""
    if not p.get('price') or (not p.get('sma50') and not p.get('sma200')):
        return 0
    rsi = p['rsi14'] if p.get('rsi14') is not None else 50
    rsi_score = max(0, min(1, (70 - rsi) / 40))
    ma_score = 0.5
    if p.get('sma50', 0) > 0 and p.get('sma200', 0) > 0:
        avg = ((p['sma50'] - p['price']) / p['sma50'] + (p['sma200'] - p['price']) / p['sma200']) / 2
        ma_score = max(0, min(1, 0.5 + avg * 10))
    dd_score = max(0, min(1, (p['high52'] - p['price']) / p['high52'] * 5)) if p.get('high52', 0) > 0 else 0.5
    score = rsi_score * 0.35 + ma_score * 0.35 + dd_score * 0.30
    if (p['dss3d'] if p.get('dss3d') is not None else 50) < 15:
        score += 0.15
    return score


def zone_of(vs):
    if vs == 0:
        return 'No data', None
    shown = max(0, round((1 - vs) * 100))   # the number on the app's Flywheel card
    return ('Cheap' if shown <= 30 else 'Fair' if shown <= 70 else 'Expensive'), shown


def flywheel_targets():
    src = open(os.path.join(ROOT, 'index.html'), encoding='utf-8').read()
    block = src[src.index('    stocks: {'):]
    block = block[:block.index('order: [')]
    return {t.strip("'"): int(p) for t, p in re.findall(r"('?[\w-]+'?):\s*\{ pct: (\d+)", block)}


def take_profit_pct():
    src = open(os.path.join(ROOT, 'index.html'), encoding='utf-8').read()
    return int(re.search(r'const TAKE_PROFIT_PCT = (\d+)', src).group(1))


def collect():
    raw = api('action=load_dca_data')
    cloud = json.loads(raw['data'])
    feed = {e['ticker']: e for e in api('action=dca_prices')['etfs']}
    targets = flywheel_targets()
    recs = [h for h in cloud['history'] if h.get('type') == 'reconcile']
    synced = recs[-1]['date'] if recs else 'never'
    # Tickers the record keeps apart (the gold hedge): held on the Flywheel page but never picked
    aside = ((recs[-1].get('record') or {}).get('aside') or {}).get('tickers', []) if recs else []
    rows = []
    for t, h in cloud['portfolios']['stocks']['holdings'].items():
        f = feed.get(t, {})
        value = h.get('shares', 0) * f.get('price', 0)
        if value < MIN_VALUE:
            continue
        cost = h.get('costBasis', 0)
        vs = valuation_score(f)
        zone, shown = zone_of(vs)
        rows.append({'t': t, 'value': value, 'ret': (value - cost) / cost * 100 if cost else 0,
                     'vs': vs, 'zone': zone, 'shown': shown, 'rsi': f.get('rsi14'), 'dss': f.get('dss3d'),
                     'target': targets.get(t, 0), 'aside': t in aside})
    total = sum(r['value'] for r in rows)
    tp = take_profit_pct()
    for r in rows:
        r['pct'] = r['value'] / total * 100 if total else 0
        # The app's two sell badges on the Flywheel cards
        r['flags'] = []
        if r['aside']:
            continue
        if r['ret'] >= tp and r['vs'] < 0.40:
            r['flags'].append('Take profit (%.0f%%)' % r['ret'])
        if r['target'] and r['pct'] >= r['target'] * 2 and total > 1000:
            r['flags'].append('Trim $%.0f (%.0f%% of Flywheel, target %d%%)' % (
                r['value'] - total * r['target'] / 100, r['pct'], r['target']))
    return rows, total, raw['lastSaved'], synced


def tv(t):
    return 'https://www.tradingview.com/chart/?symbol=' + t.replace('-', '.')


def render(rows, total, saved, synced, last, last_at):
    check = to_check(rows)
    skip = sorted([r for r in rows if r not in check and not r['aside']], key=lambda r: r['t'])
    hedge = [r for r in rows if r['aside']]

    def n(v, f='%.0f'):
        return f % v if v is not None else '--'

    def row(r):
        was = last.get(r['t'])
        news = '<span class="new">left green</span>' if was == 'Cheap' and r['zone'] != 'Cheap' else ''
        flags = ''.join('<span class="flag">%s</span>' % html.escape(f) for f in r['flags'])
        return f'''<a class="row" href="{tv(r['t'])}" target="_blank" rel="noopener">
  <span class="tk">{r['t']}</span>
  <span class="pill {r['zone'].split()[0].lower()}">{r['zone']}{' ' + str(r['shown']) if r['shown'] is not None else ''}</span>
  <span class="num {'up' if r['ret'] >= 0 else 'down'}">{r['ret']:+.1f}%</span>
  <span class="num">${r['value']:,.2f}</span>
  <span class="num dim">RSI {n(r['rsi'])} · DSS {n(r['dss'])}</span>
  <span class="tags">{news}{flags}</span>
</a>'''

    def chips(rs):
        return ''.join(f'<a class="chip" href="{tv(r["t"])}" target="_blank" rel="noopener">{r["t"]}'
                       f' <span class="{"up" if r["ret"] >= 0 else "down"}">{r["ret"]:+.0f}%</span></a>' for r in rs)
    since = f'Compared with the check on {last_at[:10]}.' if last_at else 'First check: next week it will mark the ones that just left green.'
    return f'''<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Flywheel Check</title>
<style>
:root {{ --bg:#faf6f4; --card:#fff; --text:#2b2024; --dim:#8a7b80; --line:#eadfe2; --rose:#b5476b;
  --green:#3f7d55; --green-bg:#e3f1e7; --amber:#9a6a12; --amber-bg:#f8eed6; --red:#b23a3a; --red-bg:#f7e1e1; }}
@media (prefers-color-scheme: dark) {{ :root {{ --bg:#17121a; --card:#221b25; --text:#efe6ea; --dim:#9d8e95;
  --line:#352b38; --rose:#e48aa8; --green:#7fc796; --green-bg:#1f3326; --amber:#e2b55a; --amber-bg:#3a2e14;
  --red:#ef8a8a; --red-bg:#3d1f22; }} }}
* {{ box-sizing:border-box; }}
body {{ margin:0; background:var(--bg); color:var(--text); font:15px/1.4 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif; }}
main {{ max-width:760px; margin:0 auto; padding:24px 16px 48px; }}
h1 {{ font-size:1.4rem; margin:0 0 4px; color:var(--rose); }}
h2 {{ font-size:1rem; margin:28px 0 10px; }}
.sub {{ color:var(--dim); font-size:.85rem; margin:0; }}
.row {{ display:flex; flex-wrap:wrap; gap:6px 12px; align-items:center;
  background:var(--card); border:1px solid var(--line); border-radius:10px; padding:10px 12px; margin-bottom:8px;
  color:inherit; text-decoration:none; }}
.row:hover {{ border-color:var(--rose); }}
.row:visited .tk {{ color:var(--dim); }}
.tk {{ font-weight:700; font-size:1.05rem; color:var(--rose); width:52px; }}
.num {{ font-variant-numeric:tabular-nums; }}
.dim {{ color:var(--dim); font-size:.85rem; margin-left:auto; }}
.tags {{ flex-basis:100%; display:flex; flex-wrap:wrap; gap:6px; }}
.tags:empty {{ display:none; }}
.pill {{ font-size:.78rem; font-weight:600; padding:2px 8px; border-radius:99px; justify-self:start; }}
.pill.cheap {{ background:var(--green-bg); color:var(--green); }}
.pill.fair {{ background:var(--amber-bg); color:var(--amber); }}
.pill.expensive {{ background:var(--red-bg); color:var(--red); }}
.pill.no {{ background:var(--line); color:var(--dim); }}
.flag, .new {{ font-size:.75rem; font-weight:600; padding:2px 8px; border-radius:6px; }}
.flag {{ background:var(--red-bg); color:var(--red); }}
.new {{ background:var(--amber-bg); color:var(--amber); }}
.up {{ color:var(--green); }} .down {{ color:var(--red); }}
.chips {{ display:flex; flex-wrap:wrap; gap:6px; }}
.chip {{ background:var(--card); border:1px solid var(--line); border-radius:8px; padding:5px 9px; color:var(--text);
  text-decoration:none; font-weight:600; font-size:.88rem; }}
.chip span {{ font-weight:500; font-size:.8rem; }}
.note {{ color:var(--dim); font-size:.8rem; margin-top:28px; }}
</style></head><body><main>
<h1>Flywheel check</h1>
<p class="sub">{len(check)} to check · {len(skip)} still in the green buy zone{' · gold hedge apart' if hedge else ''} · {len(rows)} open, ${total:,.2f} · {time.strftime('%a %b %-d, %-I:%M %p')}</p>
<h2>Check these: out of the green zone or flagged</h2>
{''.join(row(r) for r in check) or '<p class="sub">Nothing to check: every position is still Cheap.</p>'}
<h2>Still in the green buy zone: skip</h2>
<div class="chips">{chips(skip)}</div>
{f'<h2>Gold hedge: kept apart, not a pick</h2><div class="chips">{chips(hedge)}</div>' if hedge else ''}
<p class="note">Zone = Bloom's own Cheap / Fair / Expensive score (RSI, distance from the 50 and 200-day averages,
drop from the 52-week high, DSS 3D bonus), the same score the splitter picks buys with. Flags are the app's sell
badges: Take profit at +{take_profit_pct()}% when no longer cheap, Trim when a position is twice its target weight.
Holdings from Bloom (saved {saved[:10]}, last Robinhood sync {synced[:10]}). {since}
Every ticker opens its TradingView chart.</p>
</main></body></html>'''


def to_check(rows):
    return sorted([r for r in rows if not r['aside'] and (r['zone'] != 'Cheap' or r['flags'])],
                  key=lambda r: (not r['flags'], r['vs']))


def main():
    rows, total, saved, synced = collect()
    prev = json.load(open(ZONES)) if os.path.exists(ZONES) else {}
    page = render(rows, total, saved, synced, prev.get('zones', {}), prev.get('at'))
    os.makedirs(DATA_DIR, exist_ok=True)
    open(PAGE, 'w', encoding='utf-8').write(page)
    json.dump({'at': time.strftime('%Y-%m-%dT%H:%M'), 'zones': {r['t']: r['zone'] for r in rows}}, open(ZONES, 'w'))

    check = to_check(rows)
    cheap = sum(1 for r in rows if not r['aside']) - len(check)
    print('%d open, $%.2f · %d to check, %d still Cheap\n' % (len(rows), total, len(check), cheap))
    for r in check:
        print('%-6s %-14s %+6.1f%%  %s  %s' % (r['t'], r['zone'] + (' %d' % r['shown'] if r['shown'] is not None else ''),
                                            r['ret'], tv(r['t']), '; '.join(r['flags'])))
    print('\nPage: ' + PAGE)
    if '--open' in sys.argv:
        subprocess.run(['open', PAGE])


if __name__ == '__main__':
    main()
