# Bloom sync (match Bloom to Robinhood)

Trigger: the user says **"bloom sync"**, usually with screenshots of Robinhood's positions list.

1. **Transcribe every row**: ticker, shares, total return in dollars, total return %. Robinhood sorts
   the list by dollar return, so check the screenshots join up (the last row of one should be
   followed by the first row of the next) and that the last one ends at the "Lists" section.
   Ask for a missing screenshot rather than guess.
2. Save the rows as `sync-data/positions-YYYY-MM-DD.json`:
   `{"positions": [["SCHG", 25.82, 53.22, 6.06], ...]}`. The `sync-data/` folder is git-ignored,
   so holdings stay off GitHub.
3. **Preview**: `/usr/bin/python3 tools/robinhood_sync.py sync-data/positions-YYYY-MM-DD.json`.
   Add `--sold TICKER ...` for positions the user says they sold, which records a sell at today's price.
   Anything Bloom holds that isn't in the list is set to 0 either way.
4. **Apply**: re-run with `--apply`. The script backs up the cloud copy to `sync-data/` first.
   It uploads a dry run, then the real copy, and reads it back.
5. Tell the user what changed (the script prints it). Bloom's header should then say
   "Matched to Robinhood today" on both devices. It turns amber after 14 days without a sync.

Notes
- SCHD is held in both Cushion and Growth. The Robinhood total is split in Bloom's current ratio.
  The user agreed to this on 2026-09-28.
- The IRA isn't part of this account and is never touched.
- A ticker the script can't place (not in any Bloom order list) stops the run. Add it to the
  Flywheel first. It also needs to go into `DCA_TICKERS` in the Apps Script (see memory: Bloom = SPX script).
- In the app, a sync is a `type: 'reconcile'` history entry. Trades logged before it count as
  already included when another device syncs later, so they are never added twice.
