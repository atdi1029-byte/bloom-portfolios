// ============================================================
// Stock Trade DCA Auto-Logger — Google Apps Script
// Receives TradingView webhook JSON, logs to Google Sheets,
// and provides custom functions for the Dashboard tab.
// ============================================================

// === SHEET NAME CONSTANTS ===
var SIGNAL_LOG  = 'Signal Log';
var POSITIONS   = 'Positions';
var DASHBOARD   = 'Dashboard';
var CONFIG      = 'Config';
var PERFORMANCE = 'Performance';
var SPX_SETTINGS = 'SPX Settings';
var WEEKLY_TASKS = 'Weekly Tasks';

// ---------------------------------------------------------------
// doPost — Webhook receiver (FAST PATH)
// Only logs to Signal Log + marks as unprocessed. Returns immediately.
// Heavy work (Positions dedup, Dashboard auto-add) is handled by
// processSignalQueue() which runs on a 1-minute trigger.
//
// TradingView sends JSON like:
//   {"signal":"buy","symbol":"AAPL","price":185.50,
//    "timeframe":"4h","rsi":42.5,"indicator":"buy_dca_v2"}
// ---------------------------------------------------------------
function doPost_orig(e) {
  try {
    var raw  = e.postData.contents;
    var data = JSON.parse(raw);

    // --- Dashboard update actions (non-webhook, handled inline) ---
    if (data.action === 'update_position') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      return updatePosition_(ss, data);
    }
    if (data.action === 'update_ticker') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      return updateTicker_(ss, data);
    }
    if (data.action === 'save_analytics_csv') {
      return saveAnalyticsCSV_(data.data || '');
    }

    // --- Normal webhook: cache and return INSTANTLY ---
    // CacheService is ~50ms vs SpreadsheetApp ~3-15s on cold start.
    // processSignalQueue (1-min trigger) drains the cache to the sheet.
    var ticker = (data.symbol || '').toUpperCase();
    var signal = (data.signal || '').toLowerCase();

    var cache = CacheService.getScriptCache();
    var queue = JSON.parse(cache.get('webhookQueue') || '[]');
    queue.push({ raw: raw, ts: new Date().toISOString() });
    // Cache expires after 6 min — processSignalQueue runs every 1 min
    cache.put('webhookQueue', JSON.stringify(queue), 360);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', ticker: ticker, signal: signal }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---------------------------------------------------------------
// processSignalQueue — Runs on a 1-minute time trigger.
// Picks up PENDING rows from Signal Log, deduplicates against
// Positions, writes to Positions, and auto-adds tickers.
// Install: Triggers > Add Trigger > processSignalQueue >
//          Time-driven > Minutes timer > Every 1 minute
// ---------------------------------------------------------------
function processSignalQueue() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(SIGNAL_LOG);

  // --- Drain CacheService queue first (written by doPost fast path) ---
  var cache = CacheService.getScriptCache();
  var queueRaw = cache.get('webhookQueue');
  if (queueRaw) {
    cache.remove('webhookQueue');
    var queued = JSON.parse(queueRaw);
    for (var q = 0; q < queued.length; q++) {
      var item = queued[q];
      var data = JSON.parse(item.raw);
      var timestamp = new Date(item.ts);
      var ticker    = (data.symbol || '').toUpperCase();
      var signal    = (data.signal || '').toLowerCase();
      var price     = data.price || data.entry || '';
      var timeframe = data.timeframe || '';
      var rsiVal    = data.rsi || '';
      var indicator = data.indicator || '';
      var count     = data.add_count || data.reduce_count || '';
      var source    = 'webhook';
      logSheet.appendRow([
        timestamp, ticker, signal, price, count,
        timeframe, rsiVal, indicator, source, item.raw, 'PENDING'
      ]);
    }
  }

  var logData = logSheet.getDataRange().getValues();

  // Find all PENDING rows (column K = index 10)
  var pending = [];
  for (var i = logData.length - 1; i >= 1; i--) {
    if (String(logData[i][10] || '').trim() === 'PENDING') {
      pending.push({ row: i + 1, data: logData[i] });
    }
  }

  if (pending.length === 0) return;

  // Load Positions once for dedup
  var posSheet = ss.getSheetByName(POSITIONS);
  var posData = posSheet.getDataRange().getValues();
  var now = new Date();
  var fourHoursAgo = new Date(now.getTime() - 4 * 60 * 60 * 1000);

  // Build set of recent ticker+signal combos for fast dedup
  var recentKeys = {};
  for (var d = posData.length - 1; d >= 1; d--) {
    var dTs = posData[d][0];
    if (!(dTs instanceof Date) || dTs < fourHoursAgo) break;
    var key = String(posData[d][1]).toUpperCase() + '|' + String(posData[d][2]).toLowerCase();
    recentKeys[key] = true;
  }

  // Process each pending signal (oldest first)
  pending.reverse();
  var tickersToAdd = {};

  for (var p = 0; p < pending.length; p++) {
    var row = pending[p].data;
    var timestamp = row[0];
    var ticker = String(row[1] || '').toUpperCase();
    var signal = String(row[2] || '').toLowerCase();
    var price  = row[3] || '';
    var dedupKey = ticker + '|' + signal;

    if (!recentKeys[dedupKey]) {
      posSheet.appendRow([
        timestamp, ticker, signal, price,
        '',  // Action
        '',  // Outcome
        '',  // Entry Size $
        '',  // Profit Locked $
        ''   // Notes
      ]);
      // Mark as recent so subsequent pending signals for same ticker+signal dedup
      recentKeys[dedupKey] = true;
    }

    tickersToAdd[ticker] = true;

    // Clear the PENDING flag
    logSheet.getRange(pending[p].row, 11).setValue('DONE');
  }

  // Auto-add any new tickers to Dashboard (batched)
  var tickers = Object.keys(tickersToAdd);
  for (var t = 0; t < tickers.length; t++) {
    autoAddTicker_(ss, tickers[t]);
  }
}

// ---------------------------------------------------------------
// updatePosition_ — Mark action (Entered/Skipped) or outcome
// (TP Hit/Stopped Out) on a Positions row from the dashboard.
// Matches row by ticker + signal + price (most recent match).
// ---------------------------------------------------------------
function updatePosition_(ss, data) {
  var ticker = (data.ticker || '').toUpperCase();
  var signal = (data.signal || '').toLowerCase();
  var price  = data.price || '';
  var field  = (data.field || '').toLowerCase();   // 'action' or 'outcome'
  var value  = data.value || '';                    // 'Entered','Skipped','TP Hit','Stopped Out','Open'
  var entrySize = data.entrySize || '';
  var profitLocked = data.profitLocked || '';

  var posSheet = ss.getSheetByName(POSITIONS);
  var posData  = posSheet.getDataRange().getValues();

  // Find matching rows based on field type
  var matchRows = [];
  for (var i = posData.length - 1; i >= 1; i--) {
    var rowTicker = String(posData[i][1]).toUpperCase();
    var rowSignal = String(posData[i][2]).toLowerCase();
    var rowAction = String(posData[i][4] || '').trim().toLowerCase();
    var rowOutcome = String(posData[i][5] || '').trim().toLowerCase();

    if (rowTicker !== ticker) continue;

    // Match same side: buy+add are buy-side, sell+reduce are sell-side
    var buySide = ['buy', 'add'];
    var sellSide = ['sell', 'reduce'];
    var signalSide = buySide.indexOf(signal) >= 0 ? 'buy' : 'sell';
    var rowSide = buySide.indexOf(rowSignal) >= 0 ? 'buy' : 'sell';

    if (field === 'action') {
      // For action marking, match exact signal
      if (rowSignal !== signal && rowAction === '') {
        continue;
      }
      if (rowAction === '') matchRows.push(i + 1);
    } else if ((field === 'outcome' || field === 'partial_profit') && rowSide === signalSide && rowAction === 'entered' && (rowOutcome === '' || rowOutcome === 'open' || rowOutcome === '0x0')) {
      // For outcome: find ALL entered open rows on the same side (buy+add or sell+reduce)
      matchRows.push(i + 1);
    } else if (field === 'reopen' && rowSide === signalSide && rowAction === 'entered' && (rowOutcome === 'tp hit' || rowOutcome === 'stopped out')) {
      // For reopen: find ALL closed rows on this side to reset back to Open
      matchRows.push(i + 1);
    }
  }

  if (matchRows.length === 0) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'No matching row found for ' + ticker + ' ' + signal + ' field=' + field }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (field === 'action') {
    for (var m = 0; m < matchRows.length; m++) {
      posSheet.getRange(matchRows[m], 5).setValue(value);  // Column E = Action
    }
    if (entrySize) posSheet.getRange(matchRows[0], 7).setValue(Number(entrySize)); // Column G
    // Auto-update base size on Dashboard when entering a trade
    // Only update base on initial buy/sell — DCA add/reduce should NOT overwrite base
    var isInitialEntry = (signal === 'buy' || signal === 'sell');
    if (value.toLowerCase() === 'entered' && entrySize && isInitialEntry) {
      var dash = ss.getSheetByName(DASHBOARD);
      if (dash) {
        var dashData = dash.getDataRange().getValues();
        for (var d = 1; d < dashData.length; d++) {
          if (String(dashData[d][0]).toUpperCase() === ticker) {
            if (signal === 'buy') {
              dash.getRange(d + 1, 2).setValue(Number(entrySize)); // Column B = Buy Base
            } else {
              dash.getRange(d + 1, 11).setValue(Number(entrySize)); // Column K = Sell Base
            }
            break;
          }
        }
      }
    }
  } else if (field === 'outcome') {
    // Close ALL open rows on the same side (buy+add or sell+reduce)
    var newRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['TP Hit', 'Stopped Out', 'Open', '0x0'], true)
      .setAllowInvalid(false)
      .build();
    var closeDate = new Date();
    for (var m = 0; m < matchRows.length; m++) {
      var outcomeCell = posSheet.getRange(matchRows[m], 6);
      outcomeCell.setDataValidation(newRule);
      outcomeCell.setValue(value);  // Column F = Outcome
      // Update timestamp to close date so stats filter includes this trade
      posSheet.getRange(matchRows[m], 1).setValue(closeDate); // Column A = Timestamp
    }
    // Put profit on the most recent row (first in matchRows since we iterate reverse)
    if (profitLocked) posSheet.getRange(matchRows[0], 8).setValue(Number(profitLocked)); // Column H
  } else if (field === 'partial_profit') {
    // Won button — save running profit total without closing the trade
    posSheet.getRange(matchRows[0], 8).setValue(Number(value)); // Column H = Profit Locked
  } else if (field === 'reopen') {
    // Reopen: reset closed rows back to Open, clear profit locked
    var reopenRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['TP Hit', 'Stopped Out', 'Open', '0x0'], true)
      .setAllowInvalid(false)
      .build();
    for (var m = 0; m < matchRows.length; m++) {
      posSheet.getRange(matchRows[m], 6).setDataValidation(reopenRule);
      posSheet.getRange(matchRows[m], 6).setValue('Open'); // Column F = Outcome
      posSheet.getRange(matchRows[m], 8).setValue('');     // Column H = clear Profit Locked
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', updated: field, value: value, rows: matchRows.length }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------
// updateTicker_ — Edit a ticker's base size or other Dashboard
// tab values directly from the landing page.
// ---------------------------------------------------------------
function updateTicker_(ss, data) {
  var ticker = (data.ticker || '').toUpperCase();
  var field  = (data.field || '').toLowerCase();   // 'base'
  var value  = data.value;

  var dash = ss.getSheetByName(DASHBOARD);
  var dashData = dash.getDataRange().getValues();

  var matchRow = -1;
  for (var i = 1; i < dashData.length; i++) {
    if (String(dashData[i][0]).toUpperCase() === ticker) {
      matchRow = i + 1;
      break;
    }
  }

  if (matchRow === -1) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Ticker not found on Dashboard' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (field === 'base') {
    dash.getRange(matchRow, 2).setValue(Number(value)); // Column B = Buy Base $
  } else if (field === 'sellbase') {
    dash.getRange(matchRow, 11).setValue(Number(value)); // Column K = Sell Base $
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', updated: ticker + '.' + field, value: value }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------
// patchTradeProfit_ — Fix profit column on most recent row for ticker
// Params: ticker, profit
// ---------------------------------------------------------------
function patchTradeProfit_(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ticker = (params.ticker || '').toUpperCase();
  var profit = Number(params.profit || 0);
  if (!ticker) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Missing ticker' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var posSheet = ss.getSheetByName(POSITIONS);
  var posData = posSheet.getDataRange().getValues();
  for (var i = posData.length - 1; i >= 1; i--) {
    if (String(posData[i][1]).toUpperCase() === ticker) {
      posSheet.getRange(i + 1, 8).setValue(profit);
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', row: i + 1, profit: profit }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'error', message: 'Ticker not found' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------
// addManualTrade_ — Insert a completed trade directly into Positions
// For trades that happened outside the app (missed signals, etc.)
// Params: ticker, signal, price, outcome, profit, entrySize
// ---------------------------------------------------------------
function addManualTrade_(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ticker = (params.ticker || '').toUpperCase();
  var signal = (params.signal || 'buy').toLowerCase();
  var price = params.price || 0;
  var outcome = params.outcome || '';
  var profit = params.profit || 0;
  var entrySize = params.entrySize || '';
  // pending=true adds as unacted signal (shows in Action Needed)
  var isPending = (params.pending === 'true' || params.pending === '1');

  if (!ticker) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Missing ticker' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var posSheet = ss.getSheetByName(POSITIONS);

  if (isPending) {
    // Unacted signal — empty Action + Outcome → shows in Action Needed
    posSheet.appendRow([
      new Date(), ticker, signal, price,
      '', '', '', '', 'Manual signal'
    ]);
  } else {
    // Completed trade
    posSheet.appendRow([
      new Date(),   // A: Timestamp
      ticker,       // B: Ticker
      signal,       // C: Signal
      price,        // D: Price
      'Entered',    // E: Action
      outcome || 'Stopped Out', // F: Outcome
      entrySize,    // G: Entry Size
      Number(profit), // H: Profit Locked
      'Manual entry'  // I: Notes
    ]);
  }

  // Auto-add ticker to Dashboard if not already there
  autoAddTicker_(ss, ticker);

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', ticker: ticker, signal: signal, pending: isPending }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------
// doGet — Health check + Dashboard JSON API
// If called with ?action=dashboard, returns full dashboard data.
// Otherwise returns a simple health check.
// ---------------------------------------------------------------
// JSONP wrapper — supports &callback= param for cross-origin JSONP calls
function jsonpWrap_(json, callback) {
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet_orig(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  var callback = (e && e.parameter && e.parameter.callback) || '';

  if (action === 'ping') {
    return jsonpWrap_(JSON.stringify({ status: 'ok', pong: true }), callback);
  }

  if (action === 'dashboard') {
    return serveDashboardJSON_();
  }

  if (action === 'tasks') {
    return serveWeeklyTasks_();
  }

  if (action === 'add_portfolio') {
    return addPortfolio_(e.parameter.amount || '0');
  }

  if (action === 'update_position') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    return updatePosition_(ss, {
      ticker: e.parameter.ticker || '',
      signal: e.parameter.signal || '',
      price:  e.parameter.price || '',
      field:  e.parameter.field || '',
      value:  e.parameter.value || '',
      entrySize: e.parameter.entrySize || '',
      profitLocked: e.parameter.profitLocked || ''
    });
  }

  if (action === 'add_trade') {
    return addManualTrade_(e.parameter);
  }

  if (action === 'patch_trade') {
    return patchTradeProfit_(e.parameter);
  }

  if (action === 'liquidity') {
    return serveLiquidityJSON_();
  }

  if (action === 'net_liquidity') {
    return serveNetLiquidityJSON_();
  }

  if (action === 'global_liquidity') {
    return serveGlobalLiquidityJSON_();
  }

  if (action === 'dca_prices') {
    return serveDcaPricesJSON_(e.parameter.extra || '');
  }

  if (action === 'toggle_dca') {
    var mode  = e.parameter.mode  || 'false';
    var total = e.parameter.total || '0';
    var ss2 = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss2.getSheetByName(SPX_SETTINGS);
    if (!settingsSheet) { settingsSheet = ss2.insertSheet(SPX_SETTINGS); }
    settingsSheet.getRange('A1').setValue(mode);
    settingsSheet.getRange('A2').setValue(total);
    return jsonpWrap_(JSON.stringify({ ok: true }), callback);
  }

  // NOTE: save_dca_data, load_dca_data, dca_save_chunk and dca_save_done are now
  // handled by BLOOM_handle() in doGet() (see the BLOOM STORAGE section at the end
  // of this file). They never reach doGet_orig.

  if (action === 'set_stats_start') {
    var dateStr = e.parameter.date || '';
    if (dateStr === 'clear') {
      setConfig_('stats_start_date', '');
    } else {
      setConfig_('stats_start_date', dateStr);
    }
    return jsonpWrap_(JSON.stringify({ status: 'ok', stats_start_date: dateStr }), callback);
  }

  if (action === 'save_analytics_csv') {
    return saveAnalyticsCSV_(e.parameter.data || '');
  }

  if (action === 'load_analytics_csv') {
    return loadAnalyticsCSV_();
  }

  // --- Magic Formula ---
  if (action === 'magic_formula') {
    return serveMagicFormulaJSON_(callback);
  }

  if (action === 'quote') {
    return serveQuoteJSON_(e.parameter.ticker || '', callback);
  }

  // --- Tide DSS Cycle Allocator ---
  if (action === 'tide_data') {
    return serveTideDataJSON_(callback);
  }
  if (action === 'tide_universe') {
    return serveTideUniverseJSON_(callback);
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      message: 'Trade Tracker webhook is live',
      timestamp: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------
// serveWeeklyTasks_ — Returns tasks from Weekly Tasks tab
// ---------------------------------------------------------------
function serveWeeklyTasks_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(WEEKLY_TASKS);
  var tasks = [];
  if (sheet) {
    var data = sheet.getRange('A2:A').getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim()) tasks.push(String(data[i][0]).trim());
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', tasks: tasks }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------
// serveDashboardJSON_ — Builds JSON payload for the web dashboard
// Returns: { tickers, recentSignals, stats, actionNeeded }
// ---------------------------------------------------------------
function serveDashboardJSON_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Load Dashboard tab for ticker list ---
  var dash = ss.getSheetByName(DASHBOARD);
  var dashData = dash ? dash.getDataRange().getValues() : [];

  // --- Load Positions tab ---
  var posSheet = ss.getSheetByName(POSITIONS);
  var posData = posSheet ? posSheet.getDataRange().getValues() : [];

  // --- Load Signal Log tab ---
  var logSheet = ss.getSheetByName(SIGNAL_LOG);
  var logData = logSheet ? logSheet.getDataRange().getValues() : [];

  // --- Build ticker objects from Dashboard tab ---
  var tickers = [];
  for (var i = 1; i < dashData.length; i++) {
    var row = dashData[i];
    var ticker = String(row[0] || '').toUpperCase().trim();
    if (!ticker) continue;

    var buyBase = Number(row[1]) || 30;
    var sellBase = (row[10] !== '' && row[10] !== undefined) ? Number(row[10]) : 0;

    // Get all position rows for this ticker (loaded once above)
    var tickerRows = parsePositionRows_(posData, ticker);

    tickers.push({
      ticker:     ticker,
      base:       buyBase,
      phase:      String(row[2] || ''),
      lastSignal: String(row[3] || '').toLowerCase(),
      price:      row[4] || 0,
      nextSize:   row[5] || 0,
      locked:     row[6] || 0,
      stopouts:   row[7] || 0,
      nextAction: String(row[8] || ''),
      status:     String(row[9] || ''),

      // Per-side data
      buyBase:       buyBase,
      sellBase:      sellBase,
      buyPhase:      phaseInfoBySide_(tickerRows, 'buy'),
      buyNextSize:   nextSizeBySide_(tickerRows, 'buy', buyBase),
      buyLocked:     sumLockedBySide_(tickerRows, 'buy'),
      buyStopouts:   consecutiveStopoutsBySide_(tickerRows, 'buy'),
      buyStatus:     tickerStatusBySide_(tickerRows, 'buy'),
      buyNextAction: nextActionBySide_(tickerRows, 'buy'),
      buyLastSignal: lastSignalBySide_(tickerRows, 'buy'),
      buyPrice:      lastPriceBySide_(tickerRows, 'buy'),
      sellPhase:      phaseInfoBySide_(tickerRows, 'sell'),
      sellNextSize:   nextSizeBySide_(tickerRows, 'sell', sellBase),
      sellLocked:     sumLockedBySide_(tickerRows, 'sell'),
      sellStopouts:   consecutiveStopoutsBySide_(tickerRows, 'sell'),
      sellStatus:     tickerStatusBySide_(tickerRows, 'sell'),
      sellNextAction: nextActionBySide_(tickerRows, 'sell'),
      sellLastSignal: lastSignalBySide_(tickerRows, 'sell'),
      sellPrice:      lastPriceBySide_(tickerRows, 'sell'),
      buyOutcome:     lastOutcomeBySide_(tickerRows, 'buy'),
      buyProfitLocked: lastProfitBySide_(tickerRows, 'buy'),
      sellOutcome:     lastOutcomeBySide_(tickerRows, 'sell'),
      sellProfitLocked: lastProfitBySide_(tickerRows, 'sell'),

      // Badge system
      buyBadge:  calcBadgeGrade_(tickerRows, 'buy'),
      sellBadge: calcBadgeGrade_(tickerRows, 'sell'),
      buyEffectiveSize:  (function() {
        var b = calcBadgeGrade_(tickerRows, 'buy');
        var ns = nextSizeBySide_(tickerRows, 'buy', buyBase);
        return (typeof ns === 'number') ? ns * b.multiplier : ns;
      })(),
      sellEffectiveSize: (function() {
        var b = calcBadgeGrade_(tickerRows, 'sell');
        var ss2 = (row[10] !== '' && row[10] !== undefined) ? Number(row[10]) : 0;
        var ns = nextSizeBySide_(tickerRows, 'sell', ss2);
        return (typeof ns === 'number') ? ns * b.multiplier : ns;
      })()
    });
  }

  // --- Recent signals (last 20 from Signal Log) ---
  var recentSignals = [];
  var startIdx = Math.max(1, logData.length - 20);
  for (var j = logData.length - 1; j >= startIdx; j--) {
    var lr = logData[j];
    var ts = lr[0];
    var tsStr = '';
    if (ts instanceof Date) {
      tsStr = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    } else {
      tsStr = String(ts);
    }
    recentSignals.push({
      timestamp: tsStr,
      ticker:    String(lr[1] || '').toUpperCase(),
      signal:    String(lr[2] || '').toLowerCase(),
      price:     lr[3] || 0,
      timeframe: String(lr[5] || ''),
      rsi:       lr[6] || null
    });
  }

  // --- Action needed (positions with no Action or no Outcome marked) ---
  var actionNeeded = [];
  for (var k = 1; k < posData.length; k++) {
    var pr = posData[k];
    var pTicker = String(pr[1] || '').toUpperCase();
    var pSignal = String(pr[2] || '').toLowerCase();
    var pAction = String(pr[4] || '').trim();
    var pOutcome = String(pr[5] || '').trim();

    if (!pTicker) continue;

    if (pAction === '') {
      // Needs Entered/Skipped
      var pts = pr[0];
      var ptsStr = '';
      if (pts instanceof Date) {
        ptsStr = Utilities.formatDate(pts, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      } else {
        ptsStr = String(pts);
      }
      actionNeeded.push({
        ticker:    pTicker,
        signal:    pSignal,
        price:     pr[3] || 0,
        timestamp: ptsStr,
        type:      'mark_action',
        message:   'Mark ' + pSignal.toUpperCase() + ' as Entered/Skipped'
      });
    } else if (pAction.toLowerCase() === 'entered' && (pOutcome === '' || pOutcome.toLowerCase() === 'open')) {
      actionNeeded.push({
        ticker:    pTicker,
        signal:    pSignal,
        price:     pr[3] || 0,
        timestamp: '',
        type:      'mark_outcome',
        message:   'Mark outcome for ' + pTicker + ' ' + pSignal.toUpperCase()
      });
    }
  }

  // --- Stats start date filter ---
  var statsStartDate = null;
  var statsStartStr = '';
  var ssd = getConfig_('stats_start_date');
  if (ssd) {
    var parsed = new Date(ssd);
    if (!isNaN(parsed.getTime())) {
      statsStartDate = parsed;
      statsStartStr = Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  }

  // --- Overall stats ---
  var allTrades = getAllCompletedTrades_();
  var trades = allTrades;
  if (statsStartDate) {
    trades = allTrades.filter(function(t) {
      var ts = t.timestamp;
      if (ts instanceof Date) return ts >= statsStartDate;
      return true;
    });
  }
  var wins = trades.filter(function(t) { return t.win; });
  var losses = trades.filter(function(t) { return !t.win; });
  var totalProfit = 0, totalLost = 0;
  wins.forEach(function(t) { totalProfit += t.profitLocked; });
  losses.forEach(function(t) { totalLost += Math.abs(t.profitLocked) || 0; });
  var netPnl = totalProfit - totalLost;
  var winRate = trades.length > 0 ? (wins.length / trades.length * 100).toFixed(1) : '0';
  var streaks = trades.length > 0 ? calcStreaks_(trades) : { bestWin: 0, worstLoss: 0, current: 0 };

  // Today's signals count
  var today = new Date();
  var todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var todaySignals = 0;
  for (var m = 1; m < logData.length; m++) {
    var logTs = logData[m][0];
    if (logTs instanceof Date) {
      var logDay = Utilities.formatDate(logTs, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (logDay === todayStr) todaySignals++;
    }
  }

  // Open positions count (use per-side statuses)
  var openCount = 0;
  tickers.forEach(function(t) {
    var bs = t.buyStatus || '';
    var ss = t.sellStatus || '';
    if (bs === 'OPEN') openCount++;
    if (ss === 'OPEN') openCount++;
  });

  // Best / worst tickers
  var tickerPerf = {};
  trades.forEach(function(t) {
    if (!tickerPerf[t.ticker]) tickerPerf[t.ticker] = { wins: 0, total: 0 };
    tickerPerf[t.ticker].total++;
    if (t.win) tickerPerf[t.ticker].wins++;
  });
  var bestTicker = '', worstTicker = '', bestWR = -1, worstWR = 101;
  for (var tp in tickerPerf) {
    var wr = tickerPerf[tp].wins / tickerPerf[tp].total * 100;
    if (wr > bestWR) { bestWR = wr; bestTicker = tp; }
    if (wr < worstWR) { worstWR = wr; worstTicker = tp; }
  }

  // Load DCA pokemon state
  var settingsSheet2 = ss.getSheetByName(SPX_SETTINGS);
  var savedDcaMode  = settingsSheet2 ? String(settingsSheet2.getRange('A1').getValue()) : 'false';
  var savedDcaTotal = settingsSheet2 ? Number(settingsSheet2.getRange('A2').getValue()) || 0 : 0;

  var stats = {
    totalTrades:  trades.length,
    wins:         wins.length,
    losses:       losses.length,
    winRate:      winRate + '%',
    netPnl:       '$' + netPnl.toFixed(2),
    totalProfit:  '$' + totalProfit.toFixed(2),
    totalLost:    '$' + totalLost.toFixed(2),
    openPositions: openCount,
    todaySignals:  todaySignals,
    bestWinStreak: streaks.bestWin,
    worstLossStreak: streaks.worstLoss,
    currentStreak: streaks.current > 0 ? streaks.current + 'W' : (streaks.current < 0 ? Math.abs(streaks.current) + 'L' : '0'),
    bestTicker:   bestTicker ? bestTicker + ' (' + bestWR.toFixed(0) + '%)' : 'N/A',
    worstTicker:  worstTicker ? worstTicker + ' (' + worstWR.toFixed(0) + '%)' : 'N/A',
    dcaMode:      savedDcaMode === 'true',
    dcaTotal:     savedDcaTotal,
    statsStartDate: statsStartStr || null
  };

  // --- Closed winning trades for Graveyard ---
  var closedTrades = wins.map(function(t) {
    var ts = t.timestamp;
    var dateStr = '';
    if (ts instanceof Date) {
      dateStr = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      dateStr = String(ts).split(' ')[0];
    }
    return { ticker: t.ticker, profit: t.profitLocked, date: dateStr, signal: t.signal, price: t.price };
  });

  // --- Attach upcoming earnings dates to tickers ---
  var earningsMap = getEarningsMap_();
  tickers.forEach(function(t) {
    var e = earningsMap[t.ticker];
    t.earningsDate = e ? e.date : null;
    t.earningsDaysAway = e ? e.daysAway : null;
  });

  var payload = {
    status:        'ok',
    timestamp:     new Date().toISOString(),
    tickers:       tickers,
    recentSignals: recentSignals,
    actionNeeded:  actionNeeded,
    closedTrades:  closedTrades,
    stats:         stats
  };

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------
// Helper: auto-add ticker to Dashboard tab if not already there
// Creates row with ticker, default base size, and all formulas
// ---------------------------------------------------------------
function autoAddTicker_(ss, ticker) {
  if (!ticker) return;
  var dash = ss.getSheetByName(DASHBOARD);
  if (!dash) return;
  var data = dash.getDataRange().getValues();
  // Check if ticker already exists
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase() === ticker.toUpperCase()) return;
  }
  // Add new ticker with formulas
  var row = dash.getLastRow() + 1;
  var baseSize = Number(getConfig_('Base Position Size')) || 30;
  dash.getRange(row, 1).setValue(ticker);
  dash.getRange(row, 2).setValue(baseSize);
  dash.getRange(row, 3).setFormula('=PHASE_INFO(A' + row + ')');
  dash.getRange(row, 4).setFormula('=IFERROR(INDEX(FILTER(\'Positions\'!C:C,\'Positions\'!B:B=A' + row + '),COUNTA(FILTER(\'Positions\'!C:C,\'Positions\'!B:B=A' + row + '))),"")');
  dash.getRange(row, 5).setFormula('=IFERROR(INDEX(FILTER(\'Positions\'!D:D,\'Positions\'!B:B=A' + row + '),COUNTA(FILTER(\'Positions\'!D:D,\'Positions\'!B:B=A' + row + '))),"")');
  dash.getRange(row, 6).setFormula('=NEXT_SIZE(A' + row + ',B' + row + ')');
  dash.getRange(row, 7).setFormula('=IFERROR(SUMIFS(\'Positions\'!H:H,\'Positions\'!B:B,A' + row + '),0)');
  dash.getRange(row, 8).setFormula('=CONSECUTIVE_STOPOUTS(A' + row + ')');
  dash.getRange(row, 9).setFormula('=NEXT_ACTION(A' + row + ')');
  dash.getRange(row, 10).setFormula('=TICKER_STATUS(A' + row + ')');
  dash.getRange(row, 11).setValue(0); // Column K = Sell Base (0 until first sell trade)
}

// ---------------------------------------------------------------
// Helper: check if outcome counts as a win (TP Hit or Closed with profit)
// ---------------------------------------------------------------
function isWinOutcome_(outcome, profitLocked) {
  if (outcome === 'tp hit') return true;
  if (outcome === 'closed' && profitLocked && Number(profitLocked) > 0) return true;
  return false;
}

// ---------------------------------------------------------------
// Helper: get config value by label
// Config tab layout: Column A = label, Column B = value
// ---------------------------------------------------------------
function getConfig_(label) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG);
  var data  = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === label) return data[i][1];
  }
  return null;
}

function setConfig_(label, value) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG);
  var data  = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === label) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  // Not found — append new row
  sheet.appendRow([label, value]);
}

// ---------------------------------------------------------------
// Helper: get all Positions rows for a ticker
// Returns array of row objects sorted by timestamp ascending
// ---------------------------------------------------------------
function getPositionRows_(ticker) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(POSITIONS);
  var data  = sheet.getDataRange().getValues();
  return parsePositionRows_(data, ticker);
}

// Parse position rows from pre-loaded data (avoids re-reading sheet)
function parsePositionRows_(posData, ticker) {
  var rows = [];
  for (var i = 1; i < posData.length; i++) {
    if (String(posData[i][1]).toUpperCase() === ticker.toUpperCase()) {
      rows.push({
        timestamp: posData[i][0],
        ticker:    posData[i][1],
        signal:    String(posData[i][2]).toLowerCase(),
        price:     posData[i][3],
        action:    String(posData[i][4]).toLowerCase(),
        outcome:   String(posData[i][5]).toLowerCase(),
        entrySize: posData[i][6],
        profitLocked: posData[i][7],
        notes:     posData[i][8]
      });
    }
  }
  return rows;
}

// Filter position rows by side (buy = buy/add, sell = sell/reduce)
function filterRowsBySide_(rows, side) {
  var buySignals = ['buy', 'add'];
  var sellSignals = ['sell', 'reduce'];
  var allowed = (side === 'sell') ? sellSignals : buySignals;
  return rows.filter(function(r) { return allowed.indexOf(r.signal) >= 0; });
}

// ---------------------------------------------------------------
// Badge grade table and calculation
// ---------------------------------------------------------------
var BADGE_GRADES = [
  { grade: 1,  name: 'Boulder',     multiplier: 1.0, weight: 0,  winsReq: 0,  maxStopRate: null, maxAlloc: 0 },
  { grade: 2,  name: 'Cascade',     multiplier: 1.0, weight: 0,  winsReq: 1,  maxStopRate: null, maxAlloc: 0 },
  { grade: 3,  name: 'Thunder',     multiplier: 1.0, weight: 5,  winsReq: 2,  maxStopRate: null, maxAlloc: 25 },
  { grade: 4,  name: 'Rainbow',     multiplier: 1.3, weight: 10, winsReq: 3,  maxStopRate: null, maxAlloc: 50 },
  { grade: 5,  name: 'Soul',        multiplier: 1.5, weight: 15, winsReq: 4,  maxStopRate: 0.30, maxAlloc: 75 },
  { grade: 6,  name: 'Marsh',       multiplier: 1.7, weight: 20, winsReq: 6,  maxStopRate: 0.25, maxAlloc: 100 },
  { grade: 7,  name: 'Volcano',     multiplier: 2.0, weight: 25, winsReq: 8,  maxStopRate: 0.25, maxAlloc: 150 },
  { grade: 8,  name: 'Earth',       multiplier: 2.3, weight: 30, winsReq: 10, maxStopRate: 0.20, maxAlloc: 200 },
  { grade: 9,  name: 'Elite Four',  multiplier: 2.5, weight: 35, winsReq: 15, maxStopRate: 0.20, maxAlloc: 300 },
  { grade: 10, name: 'Master',      multiplier: 3.0, weight: 40, winsReq: 20, maxStopRate: 0.15, maxAlloc: 500 }
];

function calcBadgeGrade_(rows, side) {
  var sideRows = filterRowsBySide_(rows, side);
  var entered = sideRows.filter(function(r) { return r.action === 'entered'; });
  var completed = entered.filter(function(r) { return r.outcome === 'tp hit' || r.outcome === 'stopped out' || r.outcome === 'closed'; });
  var wins = completed.filter(function(r) { return isWinOutcome_(r.outcome, r.profitLocked); });
  var losses = completed.filter(function(r) { return !isWinOutcome_(r.outcome, r.profitLocked) && (r.outcome === 'stopped out' || r.outcome === 'closed'); });
  var stopRate = completed.length > 0 ? losses.length / completed.length : 0;

  // Find highest eligible grade
  var grade = BADGE_GRADES[0]; // default Boulder
  for (var i = BADGE_GRADES.length - 1; i >= 0; i--) {
    var bg = BADGE_GRADES[i];
    if (wins.length >= bg.winsReq) {
      if (bg.maxStopRate === null || stopRate < bg.maxStopRate) {
        grade = bg;
        break;
      }
    }
  }

  // Demotion: consecutive stop-outs
  var consStops = consecutiveStopoutsBySide_(rows, side);
  var demotion = 0;
  if (consStops >= 3) demotion = 2;
  else if (consStops >= 2) demotion = 1;

  if (demotion > 0) {
    var newGradeIdx = Math.max(0, BADGE_GRADES.indexOf(grade) - demotion);
    grade = BADGE_GRADES[newGradeIdx];
  }

  return {
    grade: grade.grade,
    name: grade.name,
    multiplier: grade.multiplier,
    weight: grade.weight,
    wins: wins.length,
    losses: losses.length,
    stopRate: completed.length > 0 ? (stopRate * 100).toFixed(1) + '%' : '0%'
  };
}

// ---------------------------------------------------------------
// addPortfolio_ — Allocate paycheck money to buy-side Grade 3+ tickers
// Called via GET: ?action=add_portfolio&amount=300
// ---------------------------------------------------------------
function addPortfolio_(amountStr) {
  var amount = Number(amountStr) || 0;
  if (amount <= 0) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Amount must be > 0' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName(DASHBOARD);
  var dashData = dash.getDataRange().getValues();
  var posSheet = ss.getSheetByName(POSITIONS);
  var posData = posSheet.getDataRange().getValues();

  // Calculate badge for each ticker's buy side
  var eligible = [];
  for (var i = 1; i < dashData.length; i++) {
    var ticker = String(dashData[i][0] || '').toUpperCase().trim();
    if (!ticker) continue;
    var tickerRows = parsePositionRows_(posData, ticker);
    var badge = calcBadgeGrade_(tickerRows, 'buy');
    if (badge.grade >= 3) {
      eligible.push({
        row: i + 1,
        ticker: ticker,
        badge: badge.name,
        grade: badge.grade,
        weight: badge.weight,
        currentBase: Number(dashData[i][1]) || 30
      });
    }
  }

  if (eligible.length === 0) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'No Grade 3+ tickers found', allocations: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Diminishing returns: effective_weight = badge_weight × absorb_rate
  // absorb_rate = 1000 / (1000 + currentBase)
  // High-base tickers absorb less, low-base tickers absorb more
  var totalEffWeight = 0;
  eligible.forEach(function(e) {
    e.absorbRate = 1000 / (1000 + e.currentBase);
    e.effWeight = e.weight * e.absorbRate;
    totalEffWeight += e.effWeight;
    // Look up maxAlloc for this grade
    var bg = BADGE_GRADES.filter(function(b) { return b.grade === e.grade; })[0];
    e.maxAlloc = bg ? bg.maxAlloc : 9999;
  });

  // Initial proportional allocation
  eligible.forEach(function(e) {
    var pct = totalEffWeight > 0 ? e.effWeight / totalEffWeight : 0;
    e.alloc = Math.round(amount * pct * 100) / 100;
    e.capped = false;
  });

  // Iterative cap + redistribute loop
  for (var iter = 0; iter < 20; iter++) {
    var overflow = 0;
    eligible.forEach(function(e) {
      if (!e.capped && e.alloc > e.maxAlloc) {
        overflow += e.alloc - e.maxAlloc;
        e.alloc = e.maxAlloc;
        e.capped = true;
      }
    });
    if (overflow <= 0.01) break;
    // Redistribute overflow to uncapped tickers proportionally
    var uncappedWeight = 0;
    eligible.forEach(function(e) { if (!e.capped) uncappedWeight += e.effWeight; });
    if (uncappedWeight <= 0) break; // all capped
    eligible.forEach(function(e) {
      if (!e.capped) {
        e.alloc += Math.round(overflow * (e.effWeight / uncappedWeight) * 100) / 100;
      }
    });
  }

  // Calculate totals
  var totalAllocated = 0;
  var allocations = [];
  eligible.forEach(function(e) {
    totalAllocated += e.alloc;
    var newBase = e.currentBase + e.alloc;
    // Update Dashboard column B (Buy Base $)
    dash.getRange(e.row, 2).setValue(newBase);
    allocations.push({
      ticker: e.ticker,
      badge: e.badge,
      grade: e.grade,
      weight: e.weight,
      maxAlloc: e.maxAlloc,
      capped: e.capped,
      percent: (totalEffWeight > 0 ? (e.effWeight / totalEffWeight) * 100 : 0).toFixed(1) + '%',
      amount: e.alloc,
      newBase: newBase
    });
  });

  totalAllocated = Math.round(totalAllocated * 100) / 100;
  var unallocated = Math.round((amount - totalAllocated) * 100) / 100;

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', totalAllocated: totalAllocated, unallocated: unallocated, allocations: allocations }))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- Side-aware helpers (used by serveDashboardJSON_) ---

function consecutiveStopoutsBySide_(rows, side) {
  var sideRows = filterRowsBySide_(rows, side);
  var entered = sideRows.filter(function(r) { return r.action === 'entered'; });
  var count = 0;
  for (var i = entered.length - 1; i >= 0; i--) {
    if (entered[i].outcome === 'stopped out') count++;
    else break;
  }
  return count;
}

function phaseInfoBySide_(rows, side) {
  var sideRows = filterRowsBySide_(rows, side);
  var entered = sideRows.filter(function(r) { return r.action === 'entered'; });

  if (entered.length === 0) {
    var pending = sideRows.filter(function(r) { return r.action === '' || r.action === 'skipped'; });
    if (pending.length > 0) return 'Phase 1 — Initial Entry';
    return 'No trades';
  }

  var last = entered[entered.length - 1];
  var consStops = consecutiveStopoutsBySide_(rows, side);
  var hasProfit = entered.some(function(r) { return isWinOutcome_(r.outcome, r.profitLocked); });
  var totalLocked = 0;
  entered.forEach(function(r) {
    if (r.profitLocked && !isNaN(r.profitLocked)) totalLocked += Number(r.profitLocked);
  });

  if (last.outcome === 'stopped out') {
    if (consStops >= 3) return 'Phase 4 — Sit Out (3+ stops)';
    return 'Phase 4 — Stopped Out (' + consStops + 'x)';
  }
  if (isWinOutcome_(last.outcome, last.profitLocked)) return 'Phase 2 — Profit Taken, Waiting';
  if (last.outcome === 'closed' && Number(last.profitLocked || 0) <= 0) return 'Phase 4 — Stopped Out (closed)';
  if (last.signal === 'add' || last.signal === 'reduce') return 'Phase 3 — DCA Add';
  if ((last.signal === 'buy' || last.signal === 'sell') && hasProfit && totalLocked > 0) return 'Phase 5 — Profit Reinvest';
  if (last.signal === 'buy' || last.signal === 'sell') return 'Phase 1 — Initial Entry';
  if (last.outcome === 'open') return 'Open Position';
  return 'Unknown';
}

function nextSizeBySide_(rows, side, baseSize) {
  var sideRows = filterRowsBySide_(rows, side);
  var entered = sideRows.filter(function(r) { return r.action === 'entered'; });
  var addSize = baseSize / 2;
  var stop1Size = baseSize / 2;
  var stop2Size = baseSize / 4;

  if (entered.length === 0) return baseSize;

  var last = entered[entered.length - 1];
  var consStops = consecutiveStopoutsBySide_(rows, side);

  if (last.outcome === 'stopped out') {
    if (consStops >= 3) return 'SIT OUT';
    if (consStops === 2) return stop2Size;
    return stop1Size;
  }
  if (last.signal === 'add' || last.signal === 'reduce') return addSize;

  var hasProfit = entered.some(function(r) { return isWinOutcome_(r.outcome, r.profitLocked); });
  var totalLocked = 0;
  entered.forEach(function(r) {
    if (isWinOutcome_(r.outcome, r.profitLocked) && r.profitLocked && !isNaN(r.profitLocked)) totalLocked += Number(r.profitLocked);
  });
  if (hasProfit && totalLocked > 0 && (last.signal === 'buy' || last.signal === 'sell')) return baseSize + totalLocked;

  return baseSize;
}

function tickerStatusBySide_(rows, side) {
  var sideRows = filterRowsBySide_(rows, side);
  var entered = sideRows.filter(function(r) { return r.action === 'entered'; });

  if (entered.length === 0) {
    var unacted = sideRows.filter(function(r) { return r.action === ''; });
    if (unacted.length > 0) return 'NEW SIGNAL';
    return 'IDLE';
  }

  var last = entered[entered.length - 1];
  if (last.outcome === 'open' || last.outcome === '') return 'OPEN';
  if (isWinOutcome_(last.outcome, last.profitLocked)) return 'PROFIT';
  if (last.outcome === 'stopped out' || (last.outcome === 'closed' && Number(last.profitLocked || 0) <= 0)) {
    var stops = consecutiveStopoutsBySide_(rows, side);
    if (stops >= 3) return 'BLOCKED';
    return 'STOPPED';
  }
  return 'IDLE';
}

function nextActionBySide_(rows, side) {
  var sideRows = filterRowsBySide_(rows, side);
  var entered = sideRows.filter(function(r) { return r.action === 'entered'; });

  var unacted = sideRows.filter(function(r) { return r.action === '' && r.signal !== ''; });
  if (unacted.length > 0) {
    var last = unacted[unacted.length - 1];
    return 'Mark ' + last.signal.toUpperCase() + ' as Entered/Skipped';
  }

  if (entered.length === 0) return 'Wait for signal';

  var last = entered[entered.length - 1];
  var consStops = consecutiveStopoutsBySide_(rows, side);

  if (last.outcome === 'open' || last.outcome === '' || last.outcome === '0x0') return 'Mark outcome (TP Hit / Stopped Out)';
  if (last.outcome === 'stopped out') {
    if (consStops >= 3) return 'Sit out — wait for reset';
    return 'Wait for next signal (reduced size)';
  }
  if (isWinOutcome_(last.outcome, last.profitLocked)) return 'Wait for next signal (profit reinvest)';
  if (last.outcome === 'closed' && Number(last.profitLocked || 0) <= 0) return 'Wait for next signal (reduced size)';
  return 'Wait for signal';
}

function sumLockedBySide_(rows, side) {
  var sideRows = filterRowsBySide_(rows, side);
  var entered = sideRows.filter(function(r) { return r.action === 'entered'; });
  var total = 0;
  entered.forEach(function(r) {
    var o = (r.outcome || '').toLowerCase();
    if (o && o !== 'open' && r.profitLocked && !isNaN(r.profitLocked)) total += Number(r.profitLocked);
  });
  return total;
}

function lastSignalBySide_(rows, side) {
  var sideRows = filterRowsBySide_(rows, side);
  if (sideRows.length === 0) return '';
  return sideRows[sideRows.length - 1].signal;
}

function lastPriceBySide_(rows, side) {
  var sideRows = filterRowsBySide_(rows, side);
  if (sideRows.length === 0) return 0;
  return sideRows[sideRows.length - 1].price;
}

function lastOutcomeBySide_(rows, side) {
  var sideRows = filterRowsBySide_(rows, side);
  var entered = sideRows.filter(function(r) { return r.action === 'entered'; });
  if (entered.length === 0) return '';
  return entered[entered.length - 1].outcome;
}

function lastProfitBySide_(rows, side) {
  var sideRows = filterRowsBySide_(rows, side);
  var entered = sideRows.filter(function(r) { return r.action === 'entered'; });
  if (entered.length === 0) return 0;
  var last = entered[entered.length - 1];
  return (last.profitLocked && !isNaN(last.profitLocked)) ? Number(last.profitLocked) : 0;
}

// ---------------------------------------------------------------
// CONSECUTIVE_STOPOUTS(ticker)
// Custom function — counts consecutive "Stopped Out" outcomes
// from most recent backwards, stopping at first non-stop-out.
// ---------------------------------------------------------------
function CONSECUTIVE_STOPOUTS(ticker) {
  if (!ticker) return 0;
  var rows = getPositionRows_(ticker);
  // Filter to entered trades only
  var entered = rows.filter(function(r) { return r.action === 'entered'; });
  var count = 0;
  // Walk backwards from most recent
  for (var i = entered.length - 1; i >= 0; i--) {
    if (entered[i].outcome === 'stopped out') {
      count++;
    } else {
      break;
    }
  }
  return count;
}

// ---------------------------------------------------------------
// PHASE_INFO(ticker)
// Custom function — returns current phase description string
//
// Phases:
//   1 — Initial Entry: last signal = buy, no prior profit → $30
//   2 — First Profit: last outcome = TP Hit → waiting for next signal
//   3 — DCA Add: last signal = add → half base ($15)
//   4 — Stopped Out: 1 stop=$15, 2 stops=$7.50, 3+ stops=sit out
//   5 — Profit Reinvest: buy signal + has prior profit → $30 + locked
// ---------------------------------------------------------------
function PHASE_INFO(ticker) {
  if (!ticker) return '';
  var rows    = getPositionRows_(ticker);
  var entered = rows.filter(function(r) { return r.action === 'entered'; });

  if (entered.length === 0) {
    // Check if there's a pending signal (logged but not yet acted on)
    var pending = rows.filter(function(r) { return r.action === '' || r.action === 'skipped'; });
    if (pending.length > 0) return 'Phase 1 — Initial Entry';
    return 'No data';
  }

  var last = entered[entered.length - 1];
  var consStops = CONSECUTIVE_STOPOUTS(ticker);
  var hasProfit = entered.some(function(r) { return isWinOutcome_(r.outcome, r.profitLocked); });
  var totalLocked = 0;
  entered.forEach(function(r) {
    if (r.profitLocked && !isNaN(r.profitLocked)) totalLocked += Number(r.profitLocked);
  });

  // Phase 4 — Stopped out
  if (last.outcome === 'stopped out') {
    if (consStops >= 3) return 'Phase 4 — Sit Out (3+ stops)';
    return 'Phase 4 — Stopped Out (' + consStops + 'x)';
  }

  // Phase 2 — Just took profit, waiting
  if (isWinOutcome_(last.outcome, last.profitLocked)) {
    return 'Phase 2 — Profit Taken, Waiting';
  }

  // Closed with loss
  if (last.outcome === 'closed' && Number(last.profitLocked || 0) <= 0) {
    return 'Phase 4 — Stopped Out (closed)';
  }

  // Phase 3 — DCA add
  if (last.signal === 'add') {
    return 'Phase 3 — DCA Add';
  }

  // Phase 5 — New buy with prior profit
  if (last.signal === 'buy' && hasProfit && totalLocked > 0) {
    return 'Phase 5 — Profit Reinvest';
  }

  // Phase 1 — Initial entry
  if (last.signal === 'buy' || last.signal === 'sell') {
    return 'Phase 1 — Initial Entry';
  }

  // Open position
  if (last.outcome === 'open' || last.outcome === '') {
    return 'Open Position';
  }

  return 'Unknown';
}

// ---------------------------------------------------------------
// NEXT_SIZE(ticker)
// Custom function — returns the dollar amount for next position
//
// Uses Config tab values:
//   Base Position Size, DCA Add Size, Stop-out 1 Size,
//   Stop-out 2 Size, Max Consecutive Stops
// ---------------------------------------------------------------
function NEXT_SIZE(ticker, tickerBase) {
  if (!ticker) return '';
  var rows    = getPositionRows_(ticker);
  var entered = rows.filter(function(r) { return r.action === 'entered'; });

  // Use per-ticker base if provided, otherwise fall back to global Config
  var globalBase = Number(getConfig_('Base Position Size')) || 30;
  var baseSize   = (tickerBase && !isNaN(tickerBase)) ? Number(tickerBase) : globalBase;
  var maxStops   = Number(getConfig_('Max Consecutive Stops')) || 3;

  // All other sizes scale from this ticker's base
  var addSize    = baseSize / 2;       // DCA add = half base
  var stop1Size  = baseSize / 2;       // 1 stop-out = half base
  var stop2Size  = baseSize / 4;       // 2 stop-outs = quarter base

  if (entered.length === 0) return baseSize;

  var last = entered[entered.length - 1];
  var consStops = CONSECUTIVE_STOPOUTS(ticker);

  // Stopped out — reduce size progressively
  if (last.outcome === 'stopped out') {
    if (consStops >= maxStops) return 'SIT OUT';
    if (consStops === 2) return stop2Size;
    if (consStops === 1) return stop1Size;
    return stop1Size;
  }

  // Last signal was add → use DCA add size
  if (last.signal === 'add' || last.signal === 'reduce') {
    return addSize;
  }

  // Has prior profit → base + all locked profit
  var hasProfit = entered.some(function(r) { return isWinOutcome_(r.outcome, r.profitLocked); });
  var totalLocked = 0;
  entered.forEach(function(r) {
    if (isWinOutcome_(r.outcome, r.profitLocked) && r.profitLocked && !isNaN(r.profitLocked)) totalLocked += Number(r.profitLocked);
  });

  if (hasProfit && totalLocked > 0 && (last.signal === 'buy' || last.signal === 'sell')) {
    return baseSize + totalLocked;
  }

  return baseSize;
}

// ---------------------------------------------------------------
// NEXT_ACTION(ticker)
// Custom function — returns what the user should do next
// ---------------------------------------------------------------
function NEXT_ACTION(ticker) {
  if (!ticker) return '';
  var rows    = getPositionRows_(ticker);
  var entered = rows.filter(function(r) { return r.action === 'entered'; });

  // Check for unacted signals
  var unacted = rows.filter(function(r) {
    return r.action === '' && r.signal !== '';
  });
  if (unacted.length > 0) {
    var last = unacted[unacted.length - 1];
    return 'Mark ' + last.signal.toUpperCase() + ' as Entered/Skipped';
  }

  if (entered.length === 0) return 'Wait for signal';

  var last = entered[entered.length - 1];
  var consStops = CONSECUTIVE_STOPOUTS(ticker);

  // Open position — mark outcome
  if (last.outcome === 'open' || last.outcome === '') {
    return 'Mark outcome (TP Hit / Stopped Out)';
  }

  // Stopped out
  if (last.outcome === 'stopped out') {
    if (consStops >= 3) return 'Sit out — wait for reset';
    return 'Wait for next signal (reduced size)';
  }

  // TP hit or Closed with profit
  if (isWinOutcome_(last.outcome, last.profitLocked)) {
    return 'Wait for next signal (profit reinvest)';
  }

  // Closed with loss
  if (last.outcome === 'closed' && Number(last.profitLocked || 0) <= 0) {
    return 'Wait for next signal (reduced size)';
  }

  return 'Wait for signal';
}

// ---------------------------------------------------------------
// TICKER_STATUS(ticker)
// Custom function — returns a status emoji/label
// ---------------------------------------------------------------
function TICKER_STATUS(ticker) {
  if (!ticker) return '';
  var rows    = getPositionRows_(ticker);
  var entered = rows.filter(function(r) { return r.action === 'entered'; });

  if (entered.length === 0) {
    var unacted = rows.filter(function(r) { return r.action === ''; });
    if (unacted.length > 0) return 'NEW SIGNAL';
    return 'IDLE';
  }

  var last = entered[entered.length - 1];
  if (last.outcome === 'open' || last.outcome === '') return 'OPEN';
  if (isWinOutcome_(last.outcome, last.profitLocked)) return 'PROFIT';
  if (last.outcome === 'stopped out' || (last.outcome === 'closed' && Number(last.profitLocked || 0) <= 0)) {
    var stops = CONSECUTIVE_STOPOUTS(ticker);
    if (stops >= 3) return 'BLOCKED';
    return 'STOPPED';
  }
  return 'IDLE';
}

// ---------------------------------------------------------------
// Helper: get ALL completed trades across all tickers
// Joins Positions data with Signal Log to get RSI + timeframe
// Returns array of enriched trade objects
// ---------------------------------------------------------------
function getAllCompletedTrades_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Load Signal Log for RSI/timeframe lookup
  var logSheet = ss.getSheetByName(SIGNAL_LOG);
  var logData  = logSheet.getDataRange().getValues();
  // Build lookup: key = "TICKER|signal|price" → {rsi, timeframe}
  // Use closest timestamp match as tiebreaker
  var logLookup = {};
  for (var i = 1; i < logData.length; i++) {
    var key = String(logData[i][1]).toUpperCase() + '|' +
              String(logData[i][2]).toLowerCase() + '|' +
              String(logData[i][3]);
    if (!logLookup[key]) logLookup[key] = [];
    logLookup[key].push({
      timestamp: logData[i][0],
      rsi:       logData[i][6],
      timeframe: logData[i][5]
    });
  }

  // Load Positions
  var posSheet = ss.getSheetByName(POSITIONS);
  var posData  = posSheet.getDataRange().getValues();
  var trades = [];

  for (var j = 1; j < posData.length; j++) {
    var action  = String(posData[j][4]).toLowerCase();
    var outcome = String(posData[j][5]).toLowerCase();

    // Only include entered trades with a resolved outcome
    if (action !== 'entered') continue;
    if (outcome !== 'tp hit' && outcome !== 'stopped out' && outcome !== 'closed') continue;

    var ticker = String(posData[j][1]).toUpperCase();
    var signal = String(posData[j][2]).toLowerCase();
    var price  = posData[j][3];
    var entrySize = posData[j][6];
    var profitLocked = posData[j][7];

    // Cross-reference Signal Log for RSI + timeframe
    var lookupKey = ticker + '|' + signal + '|' + String(price);
    var rsi = '';
    var timeframe = '';
    if (logLookup[lookupKey] && logLookup[lookupKey].length > 0) {
      // Take the first match (closest in time for same signal)
      rsi = logLookup[lookupKey][0].rsi;
      timeframe = logLookup[lookupKey][0].timeframe;
      logLookup[lookupKey].shift(); // consume so next duplicate gets next match
    }

    trades.push({
      timestamp:    posData[j][0],
      ticker:       ticker,
      signal:       signal,
      price:        price,
      outcome:      outcome,
      win:          outcome === 'tp hit' || (outcome === 'closed' && (profitLocked ? Number(profitLocked) : 0) > 0),
      entrySize:    entrySize ? Number(entrySize) : 0,
      profitLocked: profitLocked ? Number(profitLocked) : 0,
      rsi:          rsi !== '' ? Number(rsi) : null,
      timeframe:    timeframe || 'unknown'
    });
  }

  return trades;
}

// ---------------------------------------------------------------
// Helper: calculate streak (consecutive wins or losses) for a
// sorted array of trades. Returns {bestWin, worstLoss, current}.
// ---------------------------------------------------------------
function calcStreaks_(trades) {
  var bestWin = 0, worstLoss = 0, current = 0;
  var runWin = 0, runLoss = 0;
  for (var i = 0; i < trades.length; i++) {
    if (trades[i].win) {
      runWin++;
      runLoss = 0;
      if (runWin > bestWin) bestWin = runWin;
    } else {
      runLoss++;
      runWin = 0;
      if (runLoss > worstLoss) worstLoss = runLoss;
    }
  }
  // Current streak (from end)
  if (trades.length > 0) {
    var lastWin = trades[trades.length - 1].win;
    current = 0;
    for (var k = trades.length - 1; k >= 0; k--) {
      if (trades[k].win === lastWin) current++;
      else break;
    }
    if (!lastWin) current = -current; // negative = loss streak
  }
  return { bestWin: bestWin, worstLoss: worstLoss, current: current };
}

// ---------------------------------------------------------------
// buildPerformanceReport()
// Rebuilds the Performance tab from scratch with full analytics.
// Call from menu: Trading Tools → Refresh Performance
// ---------------------------------------------------------------
function buildPerformanceReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PERFORMANCE) || ss.insertSheet(PERFORMANCE);
  sheet.clearContents();

  var trades = getAllCompletedTrades_();

  if (trades.length === 0) {
    sheet.getRange('A1').setValue('No completed trades yet. Mark outcomes in Positions tab first.');
    sheet.getRange('A1').setFontWeight('bold');
    return;
  }

  var row = 1;

  // ========== SECTION 1: OVERALL SUMMARY ==========
  var wins   = trades.filter(function(t) { return t.win; });
  var losses = trades.filter(function(t) { return !t.win; });
  var totalProfit = 0, totalInvested = 0;
  trades.forEach(function(t) {
    totalInvested += t.entrySize;
    if (t.win) totalProfit += t.profitLocked;
  });
  var totalLost = 0;
  losses.forEach(function(t) { totalLost += t.entrySize; });
  var streaks = calcStreaks_(trades);
  var winRate = trades.length > 0 ? (wins.length / trades.length * 100).toFixed(1) : 0;

  sheet.getRange(row, 1).setValue('OVERALL PERFORMANCE');
  sheet.getRange(row, 1).setFontWeight('bold').setFontSize(12);
  row++;
  var summaryData = [
    ['Total Trades Entered', trades.length],
    ['Wins (TP Hit)',        wins.length],
    ['Losses (Stopped Out)', losses.length],
    ['Win Rate',             winRate + '%'],
    ['Total Profit Locked',  '$' + totalProfit.toFixed(2)],
    ['Total Invested',       '$' + totalInvested.toFixed(2)],
    ['Total Lost (Stops)',   '$' + totalLost.toFixed(2)],
    ['Net P&L',              '$' + (totalProfit - totalLost).toFixed(2)],
    ['Best Win Streak',      streaks.bestWin],
    ['Worst Loss Streak',    streaks.worstLoss],
    ['Current Streak',       streaks.current > 0 ? streaks.current + 'W' : Math.abs(streaks.current) + 'L']
  ];
  sheet.getRange(row, 1, summaryData.length, 2).setValues(summaryData);
  sheet.getRange(row, 1, summaryData.length, 1).setFontWeight('bold');
  row += summaryData.length + 2;

  // ========== SECTION 2: BY TICKER ==========
  sheet.getRange(row, 1).setValue('PERFORMANCE BY TICKER');
  sheet.getRange(row, 1).setFontWeight('bold').setFontSize(12);
  row++;
  var tickerHeader = ['Ticker', 'Trades', 'Wins', 'Losses', 'Win %', 'Profit', 'Lost', 'Net', 'Best W', 'Worst L', 'Current'];
  sheet.getRange(row, 1, 1, tickerHeader.length).setValues([tickerHeader]);
  sheet.getRange(row, 1, 1, tickerHeader.length).setFontWeight('bold');
  row++;

  // Group by ticker
  var tickers = {};
  trades.forEach(function(t) {
    if (!tickers[t.ticker]) tickers[t.ticker] = [];
    tickers[t.ticker].push(t);
  });

  // Sort tickers by win rate descending
  var tickerNames = Object.keys(tickers).sort(function(a, b) {
    var wrA = tickers[a].filter(function(t) { return t.win; }).length / tickers[a].length;
    var wrB = tickers[b].filter(function(t) { return t.win; }).length / tickers[b].length;
    return wrB - wrA;
  });

  var tickerRows = [];
  tickerNames.forEach(function(name) {
    var tt = tickers[name];
    var tw = tt.filter(function(t) { return t.win; });
    var tl = tt.filter(function(t) { return !t.win; });
    var tp = 0, tlost = 0;
    tw.forEach(function(t) { tp += t.profitLocked; });
    tl.forEach(function(t) { tlost += t.entrySize; });
    var ts = calcStreaks_(tt);
    tickerRows.push([
      name,
      tt.length,
      tw.length,
      tl.length,
      (tw.length / tt.length * 100).toFixed(0) + '%',
      '$' + tp.toFixed(2),
      '$' + tlost.toFixed(2),
      '$' + (tp - tlost).toFixed(2),
      ts.bestWin,
      ts.worstLoss,
      ts.current > 0 ? ts.current + 'W' : Math.abs(ts.current) + 'L'
    ]);
  });
  if (tickerRows.length > 0) {
    sheet.getRange(row, 1, tickerRows.length, tickerHeader.length).setValues(tickerRows);
    row += tickerRows.length;
  }
  row += 2;

  // ========== SECTION 3: BY SIGNAL TYPE ==========
  sheet.getRange(row, 1).setValue('PERFORMANCE BY SIGNAL TYPE');
  sheet.getRange(row, 1).setFontWeight('bold').setFontSize(12);
  row++;
  var sigHeader = ['Signal', 'Trades', 'Wins', 'Losses', 'Win %', 'Avg Entry $'];
  sheet.getRange(row, 1, 1, sigHeader.length).setValues([sigHeader]);
  sheet.getRange(row, 1, 1, sigHeader.length).setFontWeight('bold');
  row++;

  var signals = {};
  trades.forEach(function(t) {
    if (!signals[t.signal]) signals[t.signal] = [];
    signals[t.signal].push(t);
  });

  var sigRows = [];
  ['buy', 'sell', 'add', 'reduce'].forEach(function(sig) {
    var st = signals[sig] || [];
    if (st.length === 0) return;
    var sw = st.filter(function(t) { return t.win; });
    var avgEntry = 0;
    st.forEach(function(t) { avgEntry += t.entrySize; });
    avgEntry = st.length > 0 ? avgEntry / st.length : 0;
    sigRows.push([
      sig.toUpperCase(),
      st.length,
      sw.length,
      st.length - sw.length,
      (sw.length / st.length * 100).toFixed(0) + '%',
      '$' + avgEntry.toFixed(2)
    ]);
  });
  if (sigRows.length > 0) {
    sheet.getRange(row, 1, sigRows.length, sigHeader.length).setValues(sigRows);
    row += sigRows.length;
  }
  row += 2;

  // ========== SECTION 4: BY RSI RANGE ==========
  sheet.getRange(row, 1).setValue('PERFORMANCE BY RSI AT ENTRY');
  sheet.getRange(row, 1).setFontWeight('bold').setFontSize(12);
  row++;
  var rsiHeader = ['RSI Range', 'Trades', 'Wins', 'Losses', 'Win %'];
  sheet.getRange(row, 1, 1, rsiHeader.length).setValues([rsiHeader]);
  sheet.getRange(row, 1, 1, rsiHeader.length).setFontWeight('bold');
  row++;

  var rsiBuckets = {
    '20-30': [], '30-40': [], '40-50': [], '50-60': [], '60-70': [], '70-80': []
  };
  trades.forEach(function(t) {
    if (t.rsi === null) return;
    var r = t.rsi;
    if (r >= 20 && r < 30) rsiBuckets['20-30'].push(t);
    else if (r >= 30 && r < 40) rsiBuckets['30-40'].push(t);
    else if (r >= 40 && r < 50) rsiBuckets['40-50'].push(t);
    else if (r >= 50 && r < 60) rsiBuckets['50-60'].push(t);
    else if (r >= 60 && r < 70) rsiBuckets['60-70'].push(t);
    else if (r >= 70 && r < 80) rsiBuckets['70-80'].push(t);
  });

  var rsiRows = [];
  ['20-30', '30-40', '40-50', '50-60', '60-70', '70-80'].forEach(function(range) {
    var rt = rsiBuckets[range];
    if (rt.length === 0) return;
    var rw = rt.filter(function(t) { return t.win; });
    rsiRows.push([
      range,
      rt.length,
      rw.length,
      rt.length - rw.length,
      (rw.length / rt.length * 100).toFixed(0) + '%'
    ]);
  });
  if (rsiRows.length > 0) {
    sheet.getRange(row, 1, rsiRows.length, rsiHeader.length).setValues(rsiRows);
    row += rsiRows.length;
  }
  row += 2;

  // ========== SECTION 5: BY TIMEFRAME ==========
  sheet.getRange(row, 1).setValue('PERFORMANCE BY TIMEFRAME');
  sheet.getRange(row, 1).setFontWeight('bold').setFontSize(12);
  row++;
  var tfHeader = ['Timeframe', 'Trades', 'Wins', 'Losses', 'Win %'];
  sheet.getRange(row, 1, 1, tfHeader.length).setValues([tfHeader]);
  sheet.getRange(row, 1, 1, tfHeader.length).setFontWeight('bold');
  row++;

  var timeframes = {};
  trades.forEach(function(t) {
    var tf = t.timeframe || 'unknown';
    if (!timeframes[tf]) timeframes[tf] = [];
    timeframes[tf].push(t);
  });

  var tfRows = [];
  Object.keys(timeframes).sort().forEach(function(tf) {
    var ft = timeframes[tf];
    var fw = ft.filter(function(t) { return t.win; });
    tfRows.push([
      tf,
      ft.length,
      fw.length,
      ft.length - fw.length,
      (fw.length / ft.length * 100).toFixed(0) + '%'
    ]);
  });
  if (tfRows.length > 0) {
    sheet.getRange(row, 1, tfRows.length, tfHeader.length).setValues(tfRows);
    row += tfRows.length;
  }
  row += 2;

  // ========== SECTION 6: BY INDICATOR (buy_dca_v2 vs sell_dca_v5) ==========
  sheet.getRange(row, 1).setValue('PERFORMANCE BY INDICATOR');
  sheet.getRange(row, 1).setFontWeight('bold').setFontSize(12);
  row++;
  sheet.getRange(row, 1).setValue('(Which script generated the signal)');
  sheet.getRange(row, 1).setFontColor('#888888');
  row++;
  // We don't have indicator in Positions directly, but we can derive it:
  // buy_dca_v2 = buy/add signals, sell_dca_v5 = sell/reduce/buy(failed breakdown)
  // Actually, buy signals come from BOTH indicators. We'll group by signal instead.
  // This section is covered by "By Signal Type" above — skip duplication.
  // Instead, add a note:
  sheet.getRange(row, 1).setValue('Buy indicator signals: buy, add  |  Sell indicator signals: sell, reduce, buy (failed breakdown)');
  row += 3;

  // ========== SECTION 7: COMPLETED TRADE LOG ==========
  sheet.getRange(row, 1).setValue('COMPLETED TRADE LOG');
  sheet.getRange(row, 1).setFontWeight('bold').setFontSize(12);
  row++;
  var logHeader = ['Date', 'Ticker', 'Signal', 'Price', 'Entry $', 'Outcome', 'Profit', 'RSI', 'Timeframe'];
  sheet.getRange(row, 1, 1, logHeader.length).setValues([logHeader]);
  sheet.getRange(row, 1, 1, logHeader.length).setFontWeight('bold');
  row++;

  // Most recent first
  var sortedTrades = trades.slice().reverse();
  var logRows = [];
  sortedTrades.forEach(function(t) {
    var dateStr = '';
    if (t.timestamp instanceof Date) {
      dateStr = Utilities.formatDate(t.timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    } else {
      dateStr = String(t.timestamp);
    }
    logRows.push([
      dateStr,
      t.ticker,
      t.signal.toUpperCase(),
      t.price,
      t.entrySize > 0 ? '$' + t.entrySize.toFixed(2) : '',
      t.win ? 'TP Hit' : 'Stopped Out',
      t.win && t.profitLocked > 0 ? '$' + t.profitLocked.toFixed(2) : '',
      t.rsi !== null ? t.rsi.toFixed(1) : '',
      t.timeframe
    ]);
  });
  if (logRows.length > 0) {
    sheet.getRange(row, 1, logRows.length, logHeader.length).setValues(logRows);
  }

  // Auto-resize columns
  for (var c = 1; c <= 11; c++) {
    sheet.autoResizeColumn(c);
  }

  SpreadsheetApp.getUi().alert('Performance report updated! ' + trades.length + ' completed trades analyzed.');
}

// ---------------------------------------------------------------
// onOpen — Adds Trading Tools menu
// ---------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Trading Tools')
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addItem('Refresh Performance', 'buildPerformanceReport')
    .addSeparator()
    .addItem('Add Ticker to Dashboard', 'addTickerPrompt')
    .addItem('Test Webhook (POST sample)', 'testWebhook')
    .addToUi();
}

// ---------------------------------------------------------------
// refreshDashboard — Force-recalculate all custom functions
// ---------------------------------------------------------------
function refreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName(DASHBOARD);
  if (!dash) {
    SpreadsheetApp.getUi().alert('Dashboard tab not found.');
    return;
  }
  // Touch a dummy cell to force recalculation
  var cell = dash.getRange('Z1');
  cell.setValue(new Date().getTime());
  SpreadsheetApp.flush();
  cell.clearContent();
  SpreadsheetApp.getUi().alert('Dashboard refreshed.');
}

// ---------------------------------------------------------------
// addTickerPrompt — Add a ticker row to Dashboard
// ---------------------------------------------------------------
function addTickerPrompt() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Add Ticker', 'Enter ticker symbol (e.g. AAPL):', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;

  var ticker = result.getResponseText().toUpperCase().trim();
  if (!ticker) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName(DASHBOARD);
  var lastRow = dash.getLastRow() + 1;

  // Ask for base size
  var sizeResult = ui.prompt('Base Size', 'Enter base position size for ' + ticker + ' (e.g. 30, 10, 5):', ui.ButtonSet.OK_CANCEL);
  if (sizeResult.getSelectedButton() !== ui.Button.OK) return;
  var baseSize = Number(sizeResult.getResponseText().trim()) || 30;

  // Write ticker, base size, and formulas
  dash.getRange(lastRow, 1).setValue(ticker);
  dash.getRange(lastRow, 2).setValue(baseSize);
  dash.getRange(lastRow, 3).setFormula('=PHASE_INFO(A' + lastRow + ')');
  dash.getRange(lastRow, 4).setFormula('=IFERROR(INDEX(FILTER(\'Positions\'!C:C,\'Positions\'!B:B=A' + lastRow + '),COUNTA(FILTER(\'Positions\'!C:C,\'Positions\'!B:B=A' + lastRow + '))),"")');
  dash.getRange(lastRow, 5).setFormula('=IFERROR(INDEX(FILTER(\'Positions\'!D:D,\'Positions\'!B:B=A' + lastRow + '),COUNTA(FILTER(\'Positions\'!D:D,\'Positions\'!B:B=A' + lastRow + '))),"")');
  dash.getRange(lastRow, 6).setFormula('=NEXT_SIZE(A' + lastRow + ',B' + lastRow + ')');
  dash.getRange(lastRow, 7).setFormula('=IFERROR(SUMIFS(\'Positions\'!H:H,\'Positions\'!B:B,A' + lastRow + '),0)');
  dash.getRange(lastRow, 8).setFormula('=CONSECUTIVE_STOPOUTS(A' + lastRow + ')');
  dash.getRange(lastRow, 9).setFormula('=NEXT_ACTION(A' + lastRow + ')');
  dash.getRange(lastRow, 10).setFormula('=TICKER_STATUS(A' + lastRow + ')');

  ui.alert('Added ' + ticker + ' (base $' + baseSize + ') to Dashboard row ' + lastRow);
}

// ---------------------------------------------------------------
// testWebhook — Simulate a POST for testing
// ---------------------------------------------------------------
function testWebhook() {
  var ui = SpreadsheetApp.getUi();
  var testData = {
    signal: 'buy',
    symbol: 'TEST',
    price: 100.00,
    timeframe: '4h',
    rsi: 42.5,
    indicator: 'test'
  };

  // Simulate the doPost logic directly
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var timestamp = new Date();
    var raw = JSON.stringify(testData);

    var logSheet = ss.getSheetByName(SIGNAL_LOG);
    logSheet.appendRow([
      timestamp, testData.symbol, testData.signal, testData.price, '',
      testData.timeframe, testData.rsi, testData.indicator, 'test', raw
    ]);

    var posSheet = ss.getSheetByName(POSITIONS);
    posSheet.appendRow([
      timestamp, testData.symbol, testData.signal, testData.price,
      '', '', '', '', 'Test entry'
    ]);

    ui.alert('Test signal logged! Check Signal Log and Positions tabs.');
  } catch (err) {
    ui.alert('Error: ' + err.message);
  }
}

// ---------------------------------------------------------------
// setupSheet — One-time setup: creates all tabs with headers
// Run this once after pasting the script into Apps Script.
// ---------------------------------------------------------------
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Signal Log ---
  var log = ss.getSheetByName(SIGNAL_LOG) || ss.insertSheet(SIGNAL_LOG);
  if (log.getLastRow() === 0 || log.getRange('A1').getValue() === '') {
    log.getRange('A1:J1').setValues([[
      'Timestamp', 'Ticker', 'Signal', 'Price', 'Count',
      'Timeframe', 'RSI', 'Indicator', 'Source', 'Raw JSON'
    ]]);
    log.getRange('A1:J1').setFontWeight('bold');
    log.setFrozenRows(1);
  }

  // --- Positions ---
  var pos = ss.getSheetByName(POSITIONS) || ss.insertSheet(POSITIONS);
  if (pos.getLastRow() === 0 || pos.getRange('A1').getValue() === '') {
    pos.getRange('A1:I1').setValues([[
      'Timestamp', 'Ticker', 'Signal', 'Price',
      'Action', 'Outcome', 'Entry Size $', 'Profit Locked $', 'Notes'
    ]]);
    pos.getRange('A1:I1').setFontWeight('bold');
    pos.setFrozenRows(1);

    // Add data validation dropdowns for Action and Outcome columns
    var actionRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Entered', 'Skipped'], true)
      .setAllowInvalid(false)
      .build();
    var outcomeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['TP Hit', 'Stopped Out', 'Open', '0x0'], true)
      .setAllowInvalid(false)
      .build();
    // Apply to rows 2-500
    pos.getRange('E2:E500').setDataValidation(actionRule);
    pos.getRange('F2:F500').setDataValidation(outcomeRule);
  }

  // --- Dashboard ---
  var dash = ss.getSheetByName(DASHBOARD) || ss.insertSheet(DASHBOARD);
  if (dash.getLastRow() === 0 || dash.getRange('A1').getValue() === '') {
    dash.getRange('A1:J1').setValues([[
      'Ticker', 'Base $', 'Phase', 'Last Signal', 'Price',
      'Next Size $', 'Locked Profit', 'Stop-outs', 'Next Action', 'Status'
    ]]);
    dash.getRange('A1:J1').setFontWeight('bold');
    dash.setFrozenRows(1);
  }

  // --- Config ---
  var cfg = ss.getSheetByName(CONFIG) || ss.insertSheet(CONFIG);
  if (cfg.getLastRow() === 0 || cfg.getRange('A1').getValue() === '') {
    cfg.getRange('A1:B5').setValues([
      ['Base Position Size', 30],
      ['DCA Add Size', 15],
      ['Stop-out 1 Size', 15],
      ['Stop-out 2 Size', 7.50],
      ['Max Consecutive Stops', 3]
    ]);
    cfg.getRange('A1:A5').setFontWeight('bold');
  }

  // --- Performance ---
  var perf = ss.getSheetByName(PERFORMANCE) || ss.insertSheet(PERFORMANCE);
  if (perf.getLastRow() === 0 || perf.getRange('A1').getValue() === '') {
    perf.getRange('A1').setValue('Run Trading Tools → Refresh Performance after completing some trades.');
    perf.getRange('A1').setFontWeight('bold');
  }

  SpreadsheetApp.getUi().alert('Setup complete! All tabs created with headers.');
}

// ---------------------------------------------------------------
// Keep-alive — prevents cold starts that cause webhook timeouts
// Set up: Triggers → Add Trigger → keepAlive → Time-driven → Every 5 minutes
// ---------------------------------------------------------------
function keepAlive() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
}

// ===============================================================
// TGA LIQUIDITY INDICATOR
// Fetches Treasury General Account data from fiscaldata.treasury.gov
// and serves computed metrics to the dashboard.
// ===============================================================

var LIQUIDITY = 'Liquidity';
var NET_LIQUIDITY = 'Net Liquidity';
var TGA_API = 'https://api.fiscaldata.treasury.gov/services/api/fiscal_service/v1/accounting/dts/operating_cash_balance';
var FRED_API_KEY = '342c63c345b28c2216c193896b30df87';
var FRED_BASE = 'https://api.stlouisfed.org/fred/series/observations';

// ---------------------------------------------------------------
// refreshTgaData — Fetches last 200 days of TGA closing balances
// from the Treasury API and caches them in the Liquidity sheet.
// Set up: Triggers → Add Trigger → refreshTgaData → Time-driven → Day timer → 6am-7am
// ---------------------------------------------------------------
function refreshTgaData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LIQUIDITY);
  if (!sheet) {
    sheet = ss.insertSheet(LIQUIDITY);
    sheet.getRange('A1:D1').setValues([['Date', 'Close Balance', 'Deposits', 'Withdrawals']]);
    sheet.getRange('A1:D1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Calculate date 200 days ago (extra history so 30d SMA covers full chart)
  var now = new Date();
  var startDate = new Date(now.getTime() - 200 * 24 * 60 * 60 * 1000);
  var startStr = Utilities.formatDate(startDate, 'America/New_York', 'yyyy-MM-dd');

  // Fetch closing balances
  var closingUrl = TGA_API +
    '?filter=account_type:eq:Treasury General Account (TGA) Closing Balance,' +
    'record_date:gte:' + startStr +
    '&sort=-record_date&page[size]=200&fields=record_date,open_today_bal,account_type';

  // Also fetch deposits and withdrawals
  var depositsUrl = TGA_API +
    '?filter=account_type:eq:Total TGA Deposits (Table II),' +
    'record_date:gte:' + startStr +
    '&sort=-record_date&page[size]=200&fields=record_date,open_today_bal';

  var withdrawalsUrl = TGA_API +
    '?filter=account_type:eq:Total TGA Withdrawals (Table II) (-),' +
    'record_date:gte:' + startStr +
    '&sort=-record_date&page[size]=200&fields=record_date,open_today_bal';

  try {
    var closingResp = UrlFetchApp.fetch(closingUrl, { muteHttpExceptions: true });
    var depositsResp = UrlFetchApp.fetch(depositsUrl, { muteHttpExceptions: true });
    var withdrawalsResp = UrlFetchApp.fetch(withdrawalsUrl, { muteHttpExceptions: true });

    var closingData = JSON.parse(closingResp.getContentText()).data || [];
    var depositsData = JSON.parse(depositsResp.getContentText()).data || [];
    var withdrawalsData = JSON.parse(withdrawalsResp.getContentText()).data || [];

    // Build lookup maps for deposits and withdrawals by date
    var depMap = {};
    depositsData.forEach(function(d) { depMap[d.record_date] = Number(d.open_today_bal) || 0; });
    var wdMap = {};
    withdrawalsData.forEach(function(d) { wdMap[d.record_date] = Number(d.open_today_bal) || 0; });

    // Clear old data (keep header)
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clearContent();
    }

    // Write rows (most recent first)
    var rows = [];
    for (var i = 0; i < closingData.length; i++) {
      var d = closingData[i];
      var date = d.record_date;
      var balance = Number(d.open_today_bal) || 0;
      var deposits = depMap[date] || 0;
      var withdrawals = wdMap[date] || 0;
      rows.push([date, balance, deposits, withdrawals]);
    }

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    }

  } catch (err) {
    Logger.log('TGA fetch error: ' + err.message);
  }
}

// ---------------------------------------------------------------
// serveLiquidityJSON_ — Returns TGA time series + computed metrics
// Called via ?action=liquidity
// ---------------------------------------------------------------
function serveLiquidityJSON_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LIQUIDITY);

  // If no data or data is stale (>1 day old), refresh
  var needsRefresh = !sheet || sheet.getLastRow() <= 1;
  if (!needsRefresh && sheet.getLastRow() > 1) {
    var latestDate = sheet.getRange(2, 1).getValue();
    if (latestDate) {
      var latest = new Date(latestDate);
      var age = (new Date().getTime() - latest.getTime()) / (1000 * 60 * 60);
      if (age > 36) needsRefresh = true; // stale if >36 hours (covers weekends gracefully)
    }
  }
  if (needsRefresh) {
    refreshTgaData();
    sheet = ss.getSheetByName(LIQUIDITY);
  }

  if (!sheet || sheet.getLastRow() <= 1) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'No TGA data available yet', series: [], metrics: {} }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();

  // Build series (data is most-recent-first from sheet)
  var series = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    var dateVal = data[i][0];
    var dateStr = (dateVal instanceof Date)
      ? Utilities.formatDate(dateVal, 'America/New_York', 'yyyy-MM-dd')
      : String(dateVal).slice(0, 10);
    series.push({
      date: dateStr,
      balance: Number(data[i][1]) || 0,
      deposits: Number(data[i][2]) || 0,
      withdrawals: Number(data[i][3]) || 0
    });
  }

  // Sort ascending by date for calculations
  series.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

  // Current values
  var current = series.length > 0 ? series[series.length - 1] : { balance: 0, date: '' };
  var currentBal = current.balance; // in millions

  // 7-day and 30-day changes
  var d7 = series.length >= 6 ? series[Math.max(0, series.length - 6)] : series[0];
  var d30 = series.length >= 22 ? series[Math.max(0, series.length - 22)] : series[0];
  var d90 = series[0];

  var change7d = currentBal - (d7 ? d7.balance : currentBal);
  var change30d = currentBal - (d30 ? d30.balance : currentBal);
  var change90d = currentBal - (d90 ? d90.balance : currentBal);

  var pct7d = d7 && d7.balance > 0 ? (change7d / d7.balance * 100) : 0;
  var pct30d = d30 && d30.balance > 0 ? (change30d / d30.balance * 100) : 0;

  // 7-day moving average
  var sma7 = 0;
  var sma7Count = Math.min(5, series.length);
  for (var s = series.length - 1; s >= series.length - sma7Count && s >= 0; s--) {
    sma7 += series[s].balance;
  }
  sma7 = sma7Count > 0 ? sma7 / sma7Count : currentBal;

  // 30-day moving average
  var sma30 = 0;
  var sma30Count = Math.min(22, series.length);
  for (var s2 = series.length - 1; s2 >= series.length - sma30Count && s2 >= 0; s2--) {
    sma30 += series[s2].balance;
  }
  sma30 = sma30Count > 0 ? sma30 / sma30Count : currentBal;

  // Weekly flow (avg daily net over last 5 business days)
  var weeklyFlow = 0;
  if (series.length >= 6) {
    for (var w = series.length - 1; w >= series.length - 5 && w >= 0; w--) {
      weeklyFlow += series[w].deposits - series[w].withdrawals;
    }
  }

  // Zone classification (in millions)
  var zone = 'NORMAL';
  var zoneColor = 'gray';
  if (currentBal < 200000) { zone = 'FLOOD'; zoneColor = 'green'; }
  else if (currentBal < 400000) { zone = 'LOW'; zoneColor = 'green'; }
  else if (currentBal > 800000) { zone = 'DRAIN'; zoneColor = 'red'; }
  else if (currentBal > 600000) { zone = 'HIGH'; zoneColor = 'orange'; }

  // Trend: is 7d SMA above or below 30d SMA?
  var trend = sma7 >= sma30 ? 'rising' : 'falling';
  // Rising TGA = draining liquidity = bearish for stocks
  // Falling TGA = injecting liquidity = bullish for stocks
  var signal = trend === 'falling' ? 'bullish' : 'bearish';

  // Generate insight text
  var insight = '';
  var weeklyRate = Math.round(weeklyFlow);
  if (signal === 'bullish') {
    insight = 'TGA draining $' + formatMillions_(Math.abs(change7d)) + ' over 7d — liquidity tailwind for equities';
  } else {
    insight = 'TGA building $' + formatMillions_(Math.abs(change7d)) + ' over 7d — liquidity headwind for equities';
  }

  var metrics = {
    currentBalance: currentBal,
    currentBalanceFormatted: '$' + formatMillions_(currentBal),
    currentDate: current.date,
    change7d: change7d,
    change7dFormatted: (change7d >= 0 ? '+' : '-') + '$' + formatMillions_(Math.abs(change7d)),
    pct7d: pct7d.toFixed(1) + '%',
    change30d: change30d,
    change30dFormatted: (change30d >= 0 ? '+' : '-') + '$' + formatMillions_(Math.abs(change30d)),
    pct30d: pct30d.toFixed(1) + '%',
    change90d: change90d,
    sma7: sma7,
    sma30: sma30,
    weeklyFlow: weeklyFlow,
    zone: zone,
    zoneColor: zoneColor,
    trend: trend,
    signal: signal,
    insight: insight
  };

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', series: series, metrics: metrics }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Helper: format millions to human-readable (e.g. 798943 → "798.9B")
function formatMillions_(millions) {
  if (Math.abs(millions) >= 1000000) {
    return (millions / 1000000).toFixed(2) + 'T';
  } else if (Math.abs(millions) >= 1000) {
    return (millions / 1000).toFixed(1) + 'B';
  } else {
    return millions.toFixed(0) + 'M';
  }
}

// ---------------------------------------------------------------
// refreshNetLiquidity — Fetches WALCL, RRP, SP500 from FRED API,
// combines with TGA, calculates Net Liquidity (WALCL - TGA - RRP)
// Set up: Triggers → Add Trigger → refreshNetLiquidity → Day timer → 7am-8am
// ---------------------------------------------------------------
function refreshNetLiquidity() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NET_LIQUIDITY);
  if (!sheet) {
    sheet = ss.insertSheet(NET_LIQUIDITY);
    sheet.getRange('A1:F1').setValues([['Date', 'WALCL', 'TGA', 'RRP', 'Net Liquidity', 'SP500']]);
    sheet.getRange('A1:F1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  var now = new Date();
  var startDate = new Date(now.getTime() - 400 * 24 * 60 * 60 * 1000);
  var startStr = Utilities.formatDate(startDate, 'America/New_York', 'yyyy-MM-dd');

  try {
    var walclResp = UrlFetchApp.fetch(
      FRED_BASE + '?series_id=WALCL&api_key=' + FRED_API_KEY +
      '&file_type=json&sort_order=asc&observation_start=' + startStr,
      { muteHttpExceptions: true }
    );
    var walclData = JSON.parse(walclResp.getContentText()).observations || [];

    var rrpResp = UrlFetchApp.fetch(
      FRED_BASE + '?series_id=RRPONTSYD&api_key=' + FRED_API_KEY +
      '&file_type=json&sort_order=asc&observation_start=' + startStr,
      { muteHttpExceptions: true }
    );
    var rrpData = JSON.parse(rrpResp.getContentText()).observations || [];

    var spxResp = UrlFetchApp.fetch(
      FRED_BASE + '?series_id=SP500&api_key=' + FRED_API_KEY +
      '&file_type=json&sort_order=asc&observation_start=' + startStr,
      { muteHttpExceptions: true }
    );
    var spxData = JSON.parse(spxResp.getContentText()).observations || [];

    // Read TGA from existing Liquidity sheet
    var liqSheet = ss.getSheetByName(LIQUIDITY);
    var tgaMap = {};
    if (liqSheet && liqSheet.getLastRow() > 1) {
      var tgaRows = liqSheet.getRange(2, 1, liqSheet.getLastRow() - 1, 2).getValues();
      tgaRows.forEach(function(r) {
        var d = (r[0] instanceof Date)
          ? Utilities.formatDate(r[0], 'America/New_York', 'yyyy-MM-dd')
          : String(r[0]).slice(0, 10);
        tgaMap[d] = Number(r[1]) || 0;
      });
    }

    // Build lookup maps (FRED uses "." for missing values)
    var walclMap = {};
    walclData.forEach(function(o) {
      if (o.value !== '.') walclMap[o.date] = Number(o.value) || 0;
    });
    var rrpMap = {};
    rrpData.forEach(function(o) {
      if (o.value !== '.') rrpMap[o.date] = (Number(o.value) || 0) * 1000;
    });
    var spxMap = {};
    spxData.forEach(function(o) {
      if (o.value !== '.') spxMap[o.date] = Number(o.value) || 0;
    });

    // Collect all unique dates, sort ascending
    var dateSet = {};
    Object.keys(rrpMap).forEach(function(d) { dateSet[d] = true; });
    Object.keys(spxMap).forEach(function(d) { dateSet[d] = true; });
    Object.keys(tgaMap).forEach(function(d) { dateSet[d] = true; });
    var allDates = Object.keys(dateSet).sort();

    // Forward-fill WALCL across daily dates
    var ffWalcl = {};
    var lastWalcl = null;
    for (var i = 0; i < allDates.length; i++) {
      var d = allDates[i];
      if (walclMap[d] !== undefined) lastWalcl = walclMap[d];
      if (lastWalcl !== null) ffWalcl[d] = lastWalcl;
    }

    // Build aligned rows (only dates where all 4 components exist)
    var rows = [];
    for (var i = allDates.length - 1; i >= 0; i--) {
      var d = allDates[i];
      if (ffWalcl[d] === undefined || tgaMap[d] === undefined ||
          rrpMap[d] === undefined || spxMap[d] === undefined) continue;
      var netLiq = ffWalcl[d] - tgaMap[d] - rrpMap[d];
      rows.push([d, ffWalcl[d], tgaMap[d], rrpMap[d], netLiq, spxMap[d]]);
    }

    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).clearContent();
    }
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    }
  } catch (err) {
    Logger.log('Net Liquidity fetch error: ' + err.message);
  }
}

// ---------------------------------------------------------------
// serveNetLiquidityJSON_ — Returns Net Liquidity time series + metrics
// Called via ?action=net_liquidity
// ---------------------------------------------------------------
function serveNetLiquidityJSON_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NET_LIQUIDITY);

  var needsRefresh = !sheet || sheet.getLastRow() <= 1;
  if (!needsRefresh && sheet.getLastRow() > 1) {
    var latestDate = sheet.getRange(2, 1).getValue();
    if (latestDate) {
      var latest = new Date(latestDate);
      var age = (new Date().getTime() - latest.getTime()) / (1000 * 60 * 60);
      if (age > 36) needsRefresh = true;
    }
  }
  if (needsRefresh) {
    refreshTgaData();
    refreshNetLiquidity();
    sheet = ss.getSheetByName(NET_LIQUIDITY);
  }

  if (!sheet || sheet.getLastRow() <= 1) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'No net liquidity data yet', series: [], metrics: {} }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  var series = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    var dateStr = (data[i][0] instanceof Date)
      ? Utilities.formatDate(data[i][0], 'America/New_York', 'yyyy-MM-dd')
      : String(data[i][0]).slice(0, 10);
    series.push({
      date: dateStr,
      walcl: Number(data[i][1]) || 0,
      tga: Number(data[i][2]) || 0,
      rrp: Number(data[i][3]) || 0,
      netLiq: Number(data[i][4]) || 0,
      spx: Number(data[i][5]) || 0
    });
  }
  series.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

  if (series.length === 0) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', series: [], metrics: {} }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var current = series[series.length - 1];
  var d30 = series[Math.max(0, series.length - 22)];
  var d90 = series[Math.max(0, series.length - 64)];

  var change30d = current.netLiq - d30.netLiq;
  var change90d = current.netLiq - d90.netLiq;
  var spxChange30d = current.spx - d30.spx;

  var correlation = (change30d > 0 && spxChange30d > 0) || (change30d < 0 && spxChange30d < 0)
    ? 'aligned' : 'divergent';

  var metrics = {
    currentNetLiq: current.netLiq,
    currentNetLiqFormatted: '$' + formatMillions_(current.netLiq),
    currentSpx: current.spx,
    currentDate: current.date,
    change30d: change30d,
    change30dFormatted: (change30d >= 0 ? '+' : '-') + '$' + formatMillions_(Math.abs(change30d)),
    change90d: change90d,
    change90dFormatted: (change90d >= 0 ? '+' : '-') + '$' + formatMillions_(Math.abs(change90d)),
    correlation: correlation,
    walcl: current.walcl,
    tga: current.tga,
    rrp: current.rrp
  };

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', series: series, metrics: metrics }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===============================================================
// GLOBAL LIQUIDITY INDEX — Fed + ECB + BOJ balance sheets in USD
// Inspired by Michael Howell's Capital Wars framework.
// FRED series: WALCL (Fed), ECBASSETSW (ECB, EUR), JPNASSETS (BOJ, JPY)
// Forex: DEXUSEU (USD per EUR), DEXJPUS (JPY per USD)
// Set up: Triggers → Add Trigger → refreshGlobalLiquidity → Day timer → 8am-9am
// ===============================================================

var GLOBAL_LIQUIDITY = 'Global Liquidity';

function refreshGlobalLiquidity() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(GLOBAL_LIQUIDITY);
  if (!sheet) {
    sheet = ss.insertSheet(GLOBAL_LIQUIDITY);
    sheet.getRange('A1:H1').setValues([[
      'Date', 'Fed (USD M)', 'ECB (USD M)', 'BOJ (USD M)',
      'Global Liq (USD M)', 'SP500', 'EUR/USD', 'JPY/USD'
    ]]);
    sheet.getRange('A1:H1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  var now = new Date();
  var startDate = new Date(now.getTime() - 400 * 24 * 60 * 60 * 1000);
  var startStr = Utilities.formatDate(startDate, 'America/New_York', 'yyyy-MM-dd');

  try {
    // Fetch all 6 FRED series
    var fetchFred = function(seriesId) {
      var resp = UrlFetchApp.fetch(
        FRED_BASE + '?series_id=' + seriesId + '&api_key=' + FRED_API_KEY +
        '&file_type=json&sort_order=asc&observation_start=' + startStr,
        { muteHttpExceptions: true }
      );
      return JSON.parse(resp.getContentText()).observations || [];
    };

    var walclData = fetchFred('WALCL');
    var ecbData = fetchFred('ECBASSETSW');
    var bojData = fetchFred('JPNASSETS');
    var eurUsdData = fetchFred('DEXUSEU');
    var jpyUsdData = fetchFred('DEXJPUS');
    var spxData = fetchFred('SP500');

    // Build lookup maps
    var buildMap = function(obs, multiplier) {
      var map = {};
      obs.forEach(function(o) {
        if (o.value !== '.') map[o.date] = (Number(o.value) || 0) * (multiplier || 1);
      });
      return map;
    };

    var walclMap = buildMap(walclData, 1);       // already millions USD
    var ecbMap = buildMap(ecbData, 1);            // millions EUR
    var bojMap = buildMap(bojData, 100);           // 100M JPY → millions JPY
    var eurMap = buildMap(eurUsdData, 1);          // USD per EUR
    var jpyMap = buildMap(jpyUsdData, 1);          // JPY per USD
    var spxMap = buildMap(spxData, 1);

    // Collect all dates from daily series (SPX has most dates)
    var dateSet = {};
    Object.keys(spxMap).forEach(function(d) { dateSet[d] = true; });
    Object.keys(eurMap).forEach(function(d) { dateSet[d] = true; });
    var allDates = Object.keys(dateSet).sort();

    // Forward-fill weekly series (WALCL, ECB, BOJ) and forex
    var ffWalcl = {}, ffEcb = {}, ffBoj = {}, ffEur = {}, ffJpy = {};
    var lastW = null, lastE = null, lastB = null, lastEur = null, lastJpy = null;

    for (var i = 0; i < allDates.length; i++) {
      var d = allDates[i];
      if (walclMap[d] !== undefined) lastW = walclMap[d];
      if (ecbMap[d] !== undefined) lastE = ecbMap[d];
      if (bojMap[d] !== undefined) lastB = bojMap[d];
      if (eurMap[d] !== undefined) lastEur = eurMap[d];
      if (jpyMap[d] !== undefined) lastJpy = jpyMap[d];
      if (lastW !== null) ffWalcl[d] = lastW;
      if (lastE !== null) ffEcb[d] = lastE;
      if (lastB !== null) ffBoj[d] = lastB;
      if (lastEur !== null) ffEur[d] = lastEur;
      if (lastJpy !== null) ffJpy[d] = lastJpy;
    }

    // Build aligned rows: convert ECB and BOJ to USD
    var rows = [];
    for (var i = allDates.length - 1; i >= 0; i--) {
      var d = allDates[i];
      if (ffWalcl[d] === undefined || ffEcb[d] === undefined ||
          ffBoj[d] === undefined || ffEur[d] === undefined ||
          ffJpy[d] === undefined || spxMap[d] === undefined) continue;

      var fedUsd = ffWalcl[d];                          // millions USD
      var ecbUsd = ffEcb[d] * ffEur[d];                 // millions EUR * USD/EUR = millions USD
      var bojUsd = ffBoj[d] / ffJpy[d];                 // millions JPY / (JPY/USD) = millions USD
      var globalLiq = fedUsd + ecbUsd + bojUsd;

      rows.push([d, fedUsd, ecbUsd, bojUsd, globalLiq, spxMap[d], ffEur[d], ffJpy[d]]);
    }

    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).clearContent();
    }
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 8).setValues(rows);
    }
  } catch (err) {
    Logger.log('Global Liquidity fetch error: ' + err.message);
  }
}

// ---------------------------------------------------------------
// serveGlobalLiquidityJSON_ — Returns Global Liquidity time series + metrics
// Called via ?action=global_liquidity
// ---------------------------------------------------------------
function serveGlobalLiquidityJSON_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(GLOBAL_LIQUIDITY);

  var needsRefresh = !sheet || sheet.getLastRow() <= 1;
  if (!needsRefresh && sheet.getLastRow() > 1) {
    var latestDate = sheet.getRange(2, 1).getValue();
    if (latestDate) {
      var latest = new Date(latestDate);
      var age = (new Date().getTime() - latest.getTime()) / (1000 * 60 * 60);
      if (age > 36) needsRefresh = true;
    }
  }
  if (needsRefresh) {
    refreshGlobalLiquidity();
    sheet = ss.getSheetByName(GLOBAL_LIQUIDITY);
  }

  if (!sheet || sheet.getLastRow() <= 1) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'No global liquidity data yet', series: [], metrics: {} }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  var series = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    var dateStr = (data[i][0] instanceof Date)
      ? Utilities.formatDate(data[i][0], 'America/New_York', 'yyyy-MM-dd')
      : String(data[i][0]).slice(0, 10);
    series.push({
      date: dateStr,
      fed: Number(data[i][1]) || 0,
      ecb: Number(data[i][2]) || 0,
      boj: Number(data[i][3]) || 0,
      globalLiq: Number(data[i][4]) || 0,
      spx: Number(data[i][5]) || 0
    });
  }
  series.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

  if (series.length === 0) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', series: [], metrics: {} }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var current = series[series.length - 1];
  var d30 = series[Math.max(0, series.length - 22)];
  var d90 = series[Math.max(0, series.length - 64)];
  var d180 = series[Math.max(0, series.length - 130)];

  var change30d = current.globalLiq - d30.globalLiq;
  var change90d = current.globalLiq - d90.globalLiq;
  var change180d = current.globalLiq - d180.globalLiq;
  var spxChange30d = current.spx - d30.spx;

  var pct30d = d30.globalLiq > 0 ? (change30d / d30.globalLiq * 100) : 0;
  var pct90d = d90.globalLiq > 0 ? (change90d / d90.globalLiq * 100) : 0;

  var correlation = (change30d > 0 && spxChange30d > 0) || (change30d < 0 && spxChange30d < 0)
    ? 'aligned' : 'divergent';

  // Regime classification based on 90-day momentum
  var regime = 'STABLE';
  if (pct90d > 3) regime = 'EXPANSION';
  else if (pct90d > 1) regime = 'GROWING';
  else if (pct90d < -3) regime = 'CONTRACTION';
  else if (pct90d < -1) regime = 'TIGHTENING';

  var regimeColor = 'gray';
  if (regime === 'EXPANSION') regimeColor = 'green';
  else if (regime === 'GROWING') regimeColor = 'green';
  else if (regime === 'CONTRACTION') regimeColor = 'red';
  else if (regime === 'TIGHTENING') regimeColor = 'orange';

  var metrics = {
    currentGlobalLiq: current.globalLiq,
    currentGlobalLiqFormatted: '$' + formatMillions_(current.globalLiq),
    currentSpx: current.spx,
    currentDate: current.date,
    fed: current.fed,
    ecb: current.ecb,
    boj: current.boj,
    change30d: change30d,
    change30dFormatted: (change30d >= 0 ? '+' : '-') + '$' + formatMillions_(Math.abs(change30d)),
    change90d: change90d,
    change90dFormatted: (change90d >= 0 ? '+' : '-') + '$' + formatMillions_(Math.abs(change90d)),
    pct30d: (pct30d >= 0 ? '+' : '') + pct30d.toFixed(1) + '%',
    pct90d: (pct90d >= 0 ? '+' : '') + pct90d.toFixed(1) + '%',
    regime: regime,
    regimeColor: regimeColor,
    correlation: correlation
  };

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', series: series, metrics: metrics }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===============================================================
// DCA PORTFOLIO — ETF PRICE + TECHNICAL DATA
// Fetches live prices from "DCA Portfolio" sheet tab (GOOGLEFINANCE)
// and calculates RSI(14), SMA(50), SMA(200) from Yahoo Finance.
// ===============================================================

var DCA_SHEET = 'DCA Portfolio';
var DCA_TICKERS = [
  // Cushion (7)
  'JPST', 'VTIP', 'SPSB', 'DGRO', 'SCHD', 'SGOV', 'VOO',
  // Growth (6)
  'VTI', 'SCHG', 'VXUS', 'AVUV', 'SCHD', 'VO',
  // IRA (7)
  'JEPI', 'JEPQ', 'VNQ', 'GLDM', 'BND',
  // Flywheel (54)
  'NVDA', 'MSFT', 'AAPL', 'AMZN', 'GOOG', 'META', 'BRK-B', 'JPM',
  'GS', 'MS', 'BAC',
  'UNH', 'LLY', 'COST', 'NFLX', 'PG', 'GLD', 'GDX', 'SMH', 'EWZ',
  'IBIT', 'SLV', 'SIL', 'TSLA', 'CPER', 'TSM', 'AVGO', 'XOM', 'V', 'CAT',
  'INTC', 'HD', 'BA', 'DIS', 'WMT', 'SLX', 'DELL', 'IBM', 'KO',
  'JNJ', 'WM', 'O', 'LOW', 'TGT', 'MA', 'GDXJ', 'TLT', 'CVX', 'DE', 'UNP', 'MCD',
  'SBUX', 'PEP', 'EQIX', 'HPQ',
  'NVO', 'CB', 'ODFL', 'AMGN', 'ETN', 'PWR', 'EMR', 'AME', 'FDX',
  'NSC', 'DECK', 'CL', 'MDLZ', 'SYY',
  'ROST', 'DLTR', 'DUK', 'SO',
  'WELL', 'DLR', 'DHR',
  'CPAY', 'GLW', 'APH', 'KEYS', 'JBL', 'ARW', 'INCY'
];

// ---------------------------------------------------------------
// serveDcaPricesJSON_ — Returns live ETF prices + technicals
// Called via ?action=dca_prices
// ---------------------------------------------------------------
function serveDcaPricesJSON_(extraParam) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DCA_SHEET);

  // Read prices from sheet (row 2+, columns: A=Ticker, B=Price, C=52wHigh, D=52wLow, E=ChangePct)
  var priceMap = {};
  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    for (var i = 0; i < data.length; i++) {
      var ticker = String(data[i][0]).toUpperCase().trim();
      if (ticker) {
        priceMap[ticker] = {
          price: Number(data[i][1]) || 0,
          high52: Number(data[i][2]) || 0,
          low52: Number(data[i][3]) || 0,
          changePct: Number(data[i][4]) || 0
        };
      }
    }
  }

  // Read cached technicals from DCA Technicals sheet (pre-computed by refreshDcaTechnicals)
  var techMap = {};
  var techSheet = ss.getSheetByName('DCA Technicals');
  if (techSheet && techSheet.getLastRow() > 1) {
    var techData = techSheet.getRange(2, 1, techSheet.getLastRow() - 1, 7).getValues();
    for (var j = 0; j < techData.length; j++) {
      var tk = String(techData[j][0]).toUpperCase().trim();
      if (tk) {
        techMap[tk] = {
          rsi14: Number(techData[j][1]) || 50,
          sma50: Number(techData[j][2]) || 0,
          sma200: Number(techData[j][3]) || 0,
          price: Number(techData[j][4]) || 0,
          high52: Number(techData[j][5]) || 0,
          dss3d: Number(techData[j][6]) || 50
        };
      }
    }
  }

  // Build response — sheet prices + cached technicals (no Yahoo calls here)
  var etfs = [];
  for (var t = 0; t < DCA_TICKERS.length; t++) {
    var ticker = DCA_TICKERS[t];
    var info = priceMap[ticker] || { price: 0, high52: 0, low52: 0, changePct: 0 };
    var tech = techMap[ticker] || { rsi14: 50, sma50: 0, sma200: 0, price: 0, high52: 0, dss3d: 50 };

    etfs.push({
      ticker: ticker,
      price: info.price || tech.price || 0,
      high52: info.high52 || tech.high52 || 0,
      low52: info.low52 || 0,
      changePct: info.changePct || 0,
      rsi14: tech.rsi14,
      sma50: tech.sma50,
      sma200: tech.sma200,
      dss3d: tech.dss3d
    });
  }

  // Extra MF tickers — served from "MF Technicals" sheet (pre-computed by refreshMFTechnicals)
  if (extraParam) {
    var dcaSet = {};
    DCA_TICKERS.forEach(function(t) { dcaSet[t] = true; });
    var extras = extraParam.split(',')
      .map(function(t) { return t.trim().toUpperCase(); })
      .filter(function(t) { return t && !dcaSet[t]; });
    var mfTechMap = {};
    var mfSheet = ss.getSheetByName('MF Technicals');
    if (mfSheet && mfSheet.getLastRow() > 1) {
      var mfRows = mfSheet.getRange(2, 1, mfSheet.getLastRow() - 1, 6).getValues();
      for (var m = 0; m < mfRows.length; m++) {
        var mt = String(mfRows[m][0]).toUpperCase().trim();
        if (mt) {
          mfTechMap[mt] = {
            price:  Number(mfRows[m][1]) || 0,
            high52: Number(mfRows[m][2]) || 0,
            rsi14:  Number(mfRows[m][3]) || 50,
            sma50:  Number(mfRows[m][4]) || 0,
            sma200: Number(mfRows[m][5]) || 0
          };
        }
      }
    }
    extras.forEach(function(ticker) {
      var mf = mfTechMap[ticker];
      if (mf && mf.price > 0) {
        etfs.push({
          ticker: ticker,
          price: mf.price,
          high52: mf.high52,
          low52: 0,
          changePct: 0,
          rsi14: mf.rsi14,
          sma50: mf.sma50,
          sma200: mf.sma200
        });
      }
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', etfs: etfs }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------
// refreshMFTechnicals — Pre-compute price/RSI/SMA/high52 for MF
// top-30 tickers not already in DCA_TICKERS. Caches to "MF Technicals".
// Trigger: Day timer → 7am-8am
// ---------------------------------------------------------------
function refreshMFTechnicals() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resSheet = ss.getSheetByName(MF_RESULTS_SHEET);
  if (!resSheet || resSheet.getLastRow() < 2) return;

  var dcaSet = {};
  DCA_TICKERS.forEach(function(t) { dcaSet[t] = true; });

  var mfData = resSheet.getRange(2, 1, resSheet.getLastRow() - 1, 2).getValues();
  var mfTickers = [];
  for (var i = 0; i < mfData.length; i++) {
    var tk = String(mfData[i][1]).toUpperCase().trim();
    if (tk && !dcaSet[tk]) mfTickers.push(tk);
  }
  if (mfTickers.length === 0) return;

  var sheet = ss.getSheetByName('MF Technicals');
  if (!sheet) {
    sheet = ss.insertSheet('MF Technicals');
    sheet.getRange('A1:F1').setValues([['Ticker','Price','High52','RSI14','SMA50','SMA200']]);
    sheet.getRange('A1:F1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  var rows = [];
  for (var j = 0; j < mfTickers.length; j++) {
    var tech = calcTechnicals_(mfTickers[j]);
    rows.push([
      mfTickers[j],
      Math.round(tech.price  * 100) / 100,
      Math.round(tech.high52 * 100) / 100,
      Math.round(tech.rsi14  * 100) / 100,
      Math.round(tech.sma50  * 100) / 100,
      Math.round(tech.sma200 * 100) / 100
    ]);
    if (j < mfTickers.length - 1) Utilities.sleep(500);
  }

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).clear();
  }
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }
  Logger.log('MF Technicals: refreshed ' + rows.length + ' non-portfolio MF tickers');
}

// ---------------------------------------------------------------
// refreshDcaTechnicals — Pre-compute RSI/SMA for all DCA tickers
// and cache in "DCA Technicals" sheet. Run on a daily timer trigger.
// ---------------------------------------------------------------
function refreshDcaTechnicals() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('DCA Technicals');
  if (!sheet) {
    sheet = ss.insertSheet('DCA Technicals');
  }
  // Header with price + high52 columns
  sheet.getRange('A1:G1').setValues([['Ticker', 'RSI14', 'SMA50', 'SMA200', 'Price', 'High52', 'DSS3D']]);
  sheet.getRange('A1:G1').setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Deduplicate tickers
  var seen = {};
  var unique = [];
  for (var i = 0; i < DCA_TICKERS.length; i++) {
    if (!seen[DCA_TICKERS[i]]) { seen[DCA_TICKERS[i]] = true; unique.push(DCA_TICKERS[i]); }
  }

  var rows = [];
  for (var i = 0; i < unique.length; i++) {
    var ticker = unique[i];
    var tech = calcTechnicals_(ticker);
    rows.push([ticker, Math.round(tech.rsi14 * 100) / 100, Math.round(tech.sma50 * 100) / 100, Math.round(tech.sma200 * 100) / 100, Math.round(tech.price * 100) / 100, Math.round(tech.high52 * 100) / 100, Math.round(tech.dss3d * 100) / 100]);
  }

  // Clear old data and write new
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).clear();
  }
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }
}

// ---------------------------------------------------------------
// calcTechnicals_ — Fetch Yahoo Finance historical data and compute
// RSI(14), SMA(50), SMA(200) for a given ticker.
// ---------------------------------------------------------------
function calcTechnicals_(ticker) {
  var result = { rsi14: 50, sma50: 0, sma200: 0, price: 0, high52: 0, dss3d: 50 };

  try {
    var url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + ticker +
              '?range=1y&interval=1d&includePrePost=false';
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(resp.getContentText());

    var quote = json.chart.result[0].indicators.quote[0];
    var closes = quote.close;
    var highs  = quote.high || [];
    var prices = [];
    for (var i = 0; i < closes.length; i++) {
      if (closes[i] !== null && closes[i] !== undefined) prices.push(closes[i]);
    }

    if (prices.length < 15) return result;

    // Last close price + 52-week high
    result.price = prices[prices.length - 1];
    var h52 = 0;
    for (var h = 0; h < highs.length; h++) {
      if (highs[h] && highs[h] > h52) h52 = highs[h];
    }
    result.high52 = h52;

    // RSI(14) — Wilder's smoothed method (full history warmup)
    var gains = 0, losses = 0;
    // Seed with first 14 changes
    for (var r2 = 1; r2 <= 14; r2++) {
      var d2 = prices[r2] - prices[r2 - 1];
      if (d2 > 0) gains += d2;
      else losses += Math.abs(d2);
    }
    var avgGain = gains / 14;
    var avgLoss = losses / 14;
    // Smooth through ALL remaining bars
    for (var r3 = 15; r3 < prices.length; r3++) {
      var d3 = prices[r3] - prices[r3 - 1];
      avgGain = (avgGain * 13 + (d3 > 0 ? d3 : 0)) / 14;
      avgLoss = (avgLoss * 13 + (d3 < 0 ? Math.abs(d3) : 0)) / 14;
    }
    var rs = avgLoss > 0 ? avgGain / avgLoss : 100;
    result.rsi14 = 100 - (100 / (1 + rs));

    // SMA(50)
    if (prices.length >= 50) {
      var sum50 = 0;
      for (var s = prices.length - 50; s < prices.length; s++) sum50 += prices[s];
      result.sma50 = sum50 / 50;
    }

    // SMA(200)
    if (prices.length >= 200) {
      var sum200 = 0;
      for (var s2 = prices.length - 200; s2 < prices.length; s2++) sum200 += prices[s2];
      result.sma200 = sum200 / 200;
    }

    // DSS Bressert on 3-day bars (stochLen=10, emaLen=5)
    // Step 1: Build 3-day close bars (last close of each 3-day group)
    var bars3d = [];
    for (var b = 0; b < prices.length; b += 3) {
      var end = Math.min(b + 2, prices.length - 1);
      bars3d.push(prices[end]);
    }
    if (bars3d.length >= 12) {
      var stochLen = 10, emaLen = 5;
      var emaMult = 2.0 / (emaLen + 1);
      // First stochastic
      var raw1 = [];
      for (var d = 0; d < bars3d.length; d++) {
        var lookStart = Math.max(0, d - stochLen + 1);
        var hh = bars3d[lookStart], ll = bars3d[lookStart];
        for (var k = lookStart; k <= d; k++) { if (bars3d[k] > hh) hh = bars3d[k]; if (bars3d[k] < ll) ll = bars3d[k]; }
        raw1.push(hh !== ll ? (bars3d[d] - ll) / (hh - ll) * 100 : 50);
      }
      // EMA of raw1
      var smooth1 = [raw1[0]];
      for (var e1 = 1; e1 < raw1.length; e1++) smooth1.push(raw1[e1] * emaMult + smooth1[e1-1] * (1 - emaMult));
      // Second stochastic on smooth1
      var raw2 = [];
      for (var d2 = 0; d2 < smooth1.length; d2++) {
        var ls2 = Math.max(0, d2 - stochLen + 1);
        var hh2 = smooth1[ls2], ll2 = smooth1[ls2];
        for (var k2 = ls2; k2 <= d2; k2++) { if (smooth1[k2] > hh2) hh2 = smooth1[k2]; if (smooth1[k2] < ll2) ll2 = smooth1[k2]; }
        raw2.push(hh2 !== ll2 ? (smooth1[d2] - ll2) / (hh2 - ll2) * 100 : 50);
      }
      // EMA of raw2 = final DSS line
      var dssLine = [raw2[0]];
      for (var e2 = 1; e2 < raw2.length; e2++) dssLine.push(raw2[e2] * emaMult + dssLine[e2-1] * (1 - emaMult));
      result.dss3d = Math.round(dssLine[dssLine.length - 1] * 100) / 100;
    }

  } catch (err) {
    Logger.log('calcTechnicals_ error for ' + ticker + ': ' + err.message);
  }

  return result;
}

// ---------------------------------------------------------------
// setupDcaSheet — One-time setup: creates "DCA Portfolio" tab
// with GOOGLEFINANCE formulas for live prices.
// Run from Trading Tools menu or manually.
// ---------------------------------------------------------------
function setupDcaSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DCA_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DCA_SHEET);
  }

  // Clear existing data
  if (sheet.getLastRow() > 0) sheet.clear();

  // Build all rows in one batch
  var rows = [['Ticker', 'Price', '52w High', '52w Low', 'Change %']];
  var formulas = [];
  for (var i = 0; i < DCA_TICKERS.length; i++) {
    var t = DCA_TICKERS[i];
    rows.push([t, '', '', '', '']);
    formulas.push([
      '=GOOGLEFINANCE("' + t + '","price")',
      '=GOOGLEFINANCE("' + t + '","high52")',
      '=GOOGLEFINANCE("' + t + '","low52")',
      '=GOOGLEFINANCE("' + t + '","changepct")'
    ]);
  }

  // Single batch write for tickers
  sheet.getRange(1, 1, rows.length, 5).setValues(rows);
  // Single batch write for formulas
  sheet.getRange(2, 2, formulas.length, 4).setFormulas(formulas);

  sheet.getRange('A1:E1').setFontWeight('bold');
  sheet.setFrozenRows(1);
  for (var c = 1; c <= 5; c++) sheet.autoResizeColumn(c);
  SpreadsheetApp.getUi().alert('DCA Portfolio sheet created with GOOGLEFINANCE formulas for ' + DCA_TICKERS.length + ' ETFs.');
}

// ---------------------------------------------------------------
// DCA Portfolio Cloud Backup — LEGACY LAYOUT (read-only now)
// The old single-cell layout ("DCA Backup" tab: A1 = data, D1 =
// snapshots, B1 = timestamp) hit Sheets' 50,000-char cell limit.
// It is read ONCE by BLOOM_migrateLegacy() and then left untouched.
// All reads/writes now go through the BLOOM STORAGE section at the
// end of this file (saveDcaData_, loadDcaData_ and the old
// dailyDcaSnapshot were removed from here).
// ---------------------------------------------------------------
var DCA_BACKUP_SHEET = 'DCA Backup';

// ---------------------------------------------------------------
// EARNINGS CALENDAR
// Fetches upcoming earnings dates from FMP API, caches in Earnings sheet.
// Set up: Triggers → Add Trigger → refreshEarningsCalendar → Day timer → 7am-8am
// ---------------------------------------------------------------
var EARNINGS_SHEET = 'Earnings';

function refreshEarningsCalendar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(EARNINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(EARNINGS_SHEET);
    sheet.getRange('A1:C1').setValues([['Ticker', 'Date', 'Time']]);
    sheet.getRange('A1:C1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Get tracked tickers from Dashboard
  var dash = ss.getSheetByName(DASHBOARD);
  if (!dash) return;
  var dashData = dash.getDataRange().getValues();
  var trackedTickers = {};
  for (var i = 1; i < dashData.length; i++) {
    var t = String(dashData[i][0] || '').toUpperCase().trim();
    if (t) trackedTickers[t] = true;
  }

  // Fetch next 3 months of earnings from Alpha Vantage (reuses FRED_API_KEY)
  try {
    var url = 'https://www.alphavantage.co/query?function=EARNINGS_CALENDAR&horizon=3month&apikey=' + FRED_API_KEY;
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var csv = resp.getContentText();
    var lines = csv.split('\n');

    // CSV: symbol,name,reportDate,fiscalDateEnding,estimate,currency,timeOfTheDay
    var rows = [];
    for (var j = 1; j < lines.length; j++) {
      var cols = lines[j].split(',');
      if (cols.length < 7) continue;
      var sym = String(cols[0] || '').toUpperCase().trim();
      var reportDate = String(cols[2] || '').trim();
      var timeOfDay = String(cols[6] || '').trim();
      if (trackedTickers[sym] && reportDate) {
        rows.push([sym, reportDate, timeOfDay]);
      }
    }

    // Clear old data
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
    }

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    }

    Logger.log('Earnings calendar refreshed: ' + rows.length + ' upcoming for tracked tickers');
  } catch (err) {
    Logger.log('Earnings fetch error: ' + err.message);
  }
}

// getEarningsMap_ — Returns { TICKER: { date: 'YYYY-MM-DD', daysAway: N } }
function getEarningsMap_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(EARNINGS_SHEET);
  if (!sheet || sheet.getLastRow() <= 1) return {};

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var map = {};

  for (var i = 0; i < data.length; i++) {
    var ticker = String(data[i][0] || '').toUpperCase();
    var dateStr = String(data[i][1] || '');
    if (!ticker || !dateStr) continue;

    var earnDate = new Date(dateStr + 'T00:00:00');
    var daysAway = Math.round((earnDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
    if (daysAway < 0) continue; // skip past dates

    // Keep the closest upcoming earnings date per ticker
    if (!map[ticker] || daysAway < map[ticker].daysAway) {
      map[ticker] = { date: dateStr, daysAway: daysAway };
    }
  }
  return map;
}

// ====== ANALYTICS CSV SYNC ======
function saveAnalyticsCSV_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Analytics CSV');
  if (!sheet) {
    sheet = ss.insertSheet('Analytics CSV');
  }
  sheet.getRange('A1').setValue(data);
  sheet.getRange('B1').setValue(new Date().toISOString());
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function loadAnalyticsCSV_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Analytics CSV');
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', data: null }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var raw = sheet.getRange('A1').getValue();
  var ts = sheet.getRange('B1').getValue();
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', data: raw || null, lastSaved: ts || null }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===============================================================
// MAGIC FORMULA — Greenblatt's stock ranking system
// Scans US stocks, ranks by Earnings Yield + Return on
// Tangible Capital, serves top 30 to Bloom.
// Uses Yahoo Finance (no API key needed) + SEC ticker list.
// ===============================================================

var MF_UNIVERSE_SHEET = 'MF Universe';
var MF_RAW_SHEET = 'MF Raw Data';
var MF_PROGRESS_SHEET = 'MF Progress';
var MF_RESULTS_SHEET = 'Magic Formula';
var MF_BATCH_SIZE = 25; // Alpha Vantage — 5 calls/min, 6-min timeout = ~25 safe

// ---------------------------------------------------------------
// testMagicFormula — Quick test: fetch AAPL fundamentals via Alpha Vantage.
// ---------------------------------------------------------------
function testMagicFormula() {
  var result = getMagicFormulaMetrics_('AAPL');
  if (result) {
    Logger.log('AAPL — Earnings Yield: ' + (result.earningsYield * 100).toFixed(2) + '%' +
      ', ROTC: ' + (result.rotc * 100).toFixed(2) + '%' +
      ', MktCap: $' + (result.marketCap / 1e9).toFixed(1) + 'B' +
      ', PE: ' + result.pe.toFixed(1) +
      ', Sector: ' + result.sector);
  } else {
    Logger.log('AAPL — Failed to get metrics');
  }
}

// ---------------------------------------------------------------
// refreshMagicFormulaUniverse — Monthly trigger.
// Fetches NYSE + NASDAQ tickers from GitHub (rreichel3/US-Stock-Symbols),
// filters out financials, utilities, micro-caps (<$50M), warrants, etc.
// GitHub raw files never block server-side requests.
// ---------------------------------------------------------------
function refreshMagicFormulaUniverse() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MF_UNIVERSE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(MF_UNIVERSE_SHEET);
  }

  var sources = [
    'https://raw.githubusercontent.com/rreichel3/US-Stock-Symbols/main/nyse/nyse_full_tickers.json',
    'https://raw.githubusercontent.com/rreichel3/US-Stock-Symbols/main/nasdaq/nasdaq_full_tickers.json'
  ];

  var tickers = [];
  var seen = {};

  for (var s = 0; s < sources.length; s++) {
    try {
      var resp = UrlFetchApp.fetch(sources[s], { muteHttpExceptions: true });
      var code = resp.getResponseCode();
      if (code !== 200) {
        Logger.log('MF Universe: Source ' + s + ' returned ' + code);
        continue;
      }
      var data = JSON.parse(resp.getContentText());
      Logger.log('MF Universe: Source ' + s + ' returned ' + data.length + ' entries');

      for (var i = 0; i < data.length; i++) {
        var entry = data[i];
        var ticker = String(entry.symbol || '').toUpperCase().trim();
        var name = String(entry.name || '').trim();
        var sector = String(entry.sector || '').trim();
        var mktCapStr = String(entry.marketCap || '0');
        var mktCap = parseFloat(mktCapStr) || 0;

        // Skip: no ticker, already seen, special chars, too long
        if (!ticker || seen[ticker]) continue;
        if (ticker.indexOf('/') !== -1 || ticker.indexOf('^') !== -1 || ticker.indexOf('.') !== -1) continue;
        if (ticker.length > 5) continue;

        var nameUpper = name.toUpperCase();

        // Skip non-operating-company securities
        if (name.indexOf('Warrant') !== -1 || name.indexOf('Right') !== -1 || name.indexOf('Unit') !== -1) continue;
        if (nameUpper.indexOf('PREFERRED') !== -1 || nameUpper.indexOf('PFD') !== -1) continue;
        if (nameUpper.indexOf('NOTES ') !== -1 || nameUpper.indexOf('BOND') !== -1 || nameUpper.indexOf('SUBORDINATED') !== -1) continue;
        if (nameUpper.indexOf('DEPOSITARY') !== -1 || nameUpper.indexOf(' ADR') !== -1 || nameUpper.indexOf(' ADS') !== -1) continue;
        if (nameUpper.indexOf('ACQUISITION') !== -1 || nameUpper.indexOf('SPAC') !== -1 || nameUpper.indexOf('BLANK CHECK') !== -1) continue;
        if (nameUpper.indexOf('REALTY') !== -1 || nameUpper.indexOf('REAL ESTATE') !== -1 || nameUpper.indexOf(' REIT') !== -1) continue;
        if (nameUpper.indexOf('ETF') !== -1 || nameUpper.indexOf('FUND') !== -1 || nameUpper.indexOf('TRUST') !== -1) continue;
        if (nameUpper.indexOf('(IRELAND)') !== -1 || nameUpper.indexOf('(CANADA)') !== -1 || nameUpper.indexOf('(CAYMAN') !== -1) continue;
        if (nameUpper.indexOf('LP') !== -1 && nameUpper.indexOf('PARTNERSHIP') === -1) continue; // MLPs

        // Skip financials and utilities
        if (sector === 'Finance' || sector === 'Financial' || sector === 'Financial Services' || sector === 'Utilities') continue;
        // Skip real estate sector
        if (sector === 'Real Estate') continue;

        // Skip small-caps — only $500M+ for quality data & real businesses
        if (mktCap > 0 && mktCap < 500000000) continue;

        seen[ticker] = true;
        tickers.push([ticker, name]);
      }
    } catch (err) {
      Logger.log('MF Universe: Error on source ' + s + ': ' + err.message);
    }
  }

  if (tickers.length === 0) {
    Logger.log('MF Universe: No tickers found from any source');
    return;
  }

  // Write to sheet
  var rows = [['Ticker', 'Name']];
  for (var j = 0; j < tickers.length; j++) {
    rows.push(tickers[j]);
  }
  sheet.clear();
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange('A1:B1').setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Reset progress for new cycle
  var progSheet = ss.getSheetByName(MF_PROGRESS_SHEET);
  if (!progSheet) { progSheet = ss.insertSheet(MF_PROGRESS_SHEET); }
  progSheet.getRange('A1').setValue(0);
  var totalStocks = tickers.length;
  var totalBatches = Math.ceil(totalStocks / MF_BATCH_SIZE);
  progSheet.getRange('B1').setValue(totalBatches);
  progSheet.getRange('C1').setValue(new Date().toISOString());

  Logger.log('MF Universe: ' + totalStocks + ' tickers, ' + totalBatches + ' batches — ready to scan');
}

// ---------------------------------------------------------------
// getMagicFormulaMetrics_ — Fetch EBIT, EV, working capital, and
// fixed assets from Yahoo Finance for a single ticker.
// Returns { earningsYield, rotc, marketCap, sector, pe } or null.
// ---------------------------------------------------------------
function getMagicFormulaMetrics_(ticker) {
  // Uses Alpha Vantage OVERVIEW endpoint — 1 call per ticker.
  // Derives Magic Formula metrics:
  //   Earnings Yield ≈ EBITDA / EV (via 1/EVToEBITDA)
  //   Return on Capital ≈ ReturnOnAssetsTTM (proxy for ROIC)
  try {
    var url = 'https://www.alphavantage.co/query?function=OVERVIEW&symbol=' +
      ticker + '&apikey=' + FRED_API_KEY;
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var code = resp.getResponseCode();
    if (code !== 200) return null;

    var data = JSON.parse(resp.getContentText());

    // Alpha Vantage returns {} or {"Note": "..."} on rate limit / bad ticker
    if (!data || !data.Symbol || data.Note) return null;

    var mktCap = parseFloat(data.MarketCapitalization) || 0;
    if (mktCap < 500000000) return null; // skip under $500M

    var evToEbitda = parseFloat(data.EVToEBITDA) || 0;
    var ebitda = parseFloat(data.EBITDA) || 0;
    var roa = parseFloat(data.ReturnOnAssetsTTM) || 0;
    var pe = parseFloat(data.PERatio) || 0;
    var sector = String(data.Sector || '');

    // Skip financials and utilities (double check — universe should already filter)
    if (sector === 'FINANCE' || sector === 'Financial Services' ||
        sector === 'UTILITIES' || sector === 'Utilities') return null;

    // Need both metrics to rank
    if (evToEbitda <= 0 || roa <= 0) return null;

    // Earnings Yield = EBITDA / EV = 1 / EVToEBITDA
    var earningsYield = 1.0 / evToEbitda;

    // Return on Capital proxy = ROA (higher = better quality business)
    var rotc = roa;

    return {
      earningsYield: earningsYield,
      rotc: rotc,
      marketCap: mktCap,
      sector: sector,
      pe: pe
    };
  } catch (err) {
    return null;
  }
}

// Quick diagnostic — run this to verify Alpha Vantage works
function testMagicFormulaDebug() {
  var url = 'https://www.alphavantage.co/query?function=OVERVIEW&symbol=AAPL&apikey=' + FRED_API_KEY;
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  Logger.log('HTTP ' + resp.getResponseCode());
  var text = resp.getContentText();
  Logger.log('Response (first 500 chars): ' + text.substring(0, 500));

  // Log the key fields to see what's available
  var data = JSON.parse(text);
  Logger.log('Key fields — MktCap: ' + data.MarketCapitalization +
    ', EBITDA: ' + data.EBITDA +
    ', EVToEBITDA: ' + data.EVToEBITDA +
    ', ROA: ' + data.ReturnOnAssetsTTM +
    ', ROE: ' + data.ReturnOnEquityTTM +
    ', PE: ' + data.PERatio +
    ', Sector: ' + data.Sector +
    ', ProfitMargin: ' + data.ProfitMargin +
    ', OperatingMargin: ' + data.OperatingMarginTTM);

  Utilities.sleep(13000); // rate limit
  var result = getMagicFormulaMetrics_('AAPL');
  if (result) {
    Logger.log('AAPL — EY: ' + (result.earningsYield * 100).toFixed(2) +
      '%, ROA: ' + (result.rotc * 100).toFixed(2) +
      '%, MktCap: $' + (result.marketCap / 1e9).toFixed(1) +
      'B, PE: ' + result.pe.toFixed(1) +
      ', Sector: ' + result.sector);
  } else {
    Logger.log('AAPL — Failed (check fields above for None/0 values)');
  }
}

// ---------------------------------------------------------------
// refreshMagicFormulaBatch — Daily trigger (6 AM).
// Fetches fundamentals for one batch of tickers via Yahoo Finance.
// When all batches done, runs ranking computation.
// ---------------------------------------------------------------
function refreshMagicFormulaBatch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var progSheet = ss.getSheetByName(MF_PROGRESS_SHEET);
  if (!progSheet) {
    Logger.log('MF Batch: No progress sheet — run refreshMagicFormulaUniverse first');
    return;
  }
  var batchIndex = Number(progSheet.getRange('A1').getValue()) || 0;
  var totalBatches = Number(progSheet.getRange('B1').getValue()) || 0;

  if (totalBatches === 0) {
    Logger.log('MF Batch: No batches configured — run refreshMagicFormulaUniverse first');
    return;
  }

  // If all batches done, compute rankings and reset
  if (batchIndex >= totalBatches) {
    computeMagicFormulaRankings_();
    progSheet.getRange('A1').setValue(0);
    progSheet.getRange('C1').setValue(new Date().toISOString());
    Logger.log('MF Batch: All batches complete — rankings computed, cycle reset');
    return;
  }

  // Read universe
  var uniSheet = ss.getSheetByName(MF_UNIVERSE_SHEET);
  if (!uniSheet || uniSheet.getLastRow() < 2) {
    Logger.log('MF Batch: Empty universe');
    return;
  }
  var uniData = uniSheet.getRange(2, 1, uniSheet.getLastRow() - 1, 2).getValues();

  var startIdx = batchIndex * MF_BATCH_SIZE;
  var endIdx = Math.min(startIdx + MF_BATCH_SIZE, uniData.length);

  if (startIdx >= uniData.length) {
    computeMagicFormulaRankings_();
    progSheet.getRange('A1').setValue(0);
    return;
  }

  // Get or create raw data sheet
  var rawSheet = ss.getSheetByName(MF_RAW_SHEET);
  if (!rawSheet) {
    rawSheet = ss.insertSheet(MF_RAW_SHEET);
    rawSheet.getRange('A1:H1').setValues([[
      'Ticker', 'Name', 'Sector', 'MarketCap',
      'EarningsYield', 'ROTC', 'PE', 'DateFetched'
    ]]);
    rawSheet.getRange('A1:H1').setFontWeight('bold');
    rawSheet.setFrozenRows(1);
  }

  // If batch 0, clear old raw data for fresh cycle
  if (batchIndex === 0) {
    if (rawSheet.getLastRow() > 1) {
      rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 8).clear();
    }
  }

  var newRows = [];
  var now = new Date().toISOString();

  for (var i = startIdx; i < endIdx; i++) {
    var ticker = String(uniData[i][0]).trim();
    var name = String(uniData[i][1]).trim();
    if (!ticker) continue;

    var metrics = getMagicFormulaMetrics_(ticker);
    if (metrics && metrics.earningsYield > 0 && metrics.rotc > 0) {
      newRows.push([
        ticker, name, metrics.sector || '', metrics.marketCap,
        metrics.earningsYield, metrics.rotc, metrics.pe, now
      ]);
    }

    // Delay between Yahoo Finance calls
    // Alpha Vantage: 5 calls/min = 12s between calls
    if (i < endIdx - 1) Utilities.sleep(12500);
  }

  // Append new rows
  if (newRows.length > 0) {
    var nextRow = rawSheet.getLastRow() + 1;
    rawSheet.getRange(nextRow, 1, newRows.length, 8).setValues(newRows);
  }

  progSheet.getRange('A1').setValue(batchIndex + 1);

  Logger.log('MF Batch ' + (batchIndex + 1) + '/' + totalBatches +
    ': processed ' + (endIdx - startIdx) + ' tickers, ' + newRows.length + ' valid');
}

// ---------------------------------------------------------------
// computeMagicFormulaRankings_ — Ranks all stocks by both metrics,
// computes combined rank, writes top 30 to results sheet.
// ---------------------------------------------------------------
function computeMagicFormulaRankings_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawSheet = ss.getSheetByName(MF_RAW_SHEET);

  if (!rawSheet || rawSheet.getLastRow() < 2) {
    Logger.log('MF Rankings: No raw data');
    return;
  }

  var data = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 8).getValues();

  // Build array of objects
  var stocks = [];
  for (var i = 0; i < data.length; i++) {
    var ey = Number(data[i][4]) || 0;
    var rotc = Number(data[i][5]) || 0;
    if (ey <= 0 || rotc <= 0) continue;
    stocks.push({
      ticker: String(data[i][0]),
      name: String(data[i][1]),
      sector: String(data[i][2]),
      marketCap: Number(data[i][3]) || 0,
      earningsYield: ey,
      rotc: rotc,
      pe: Number(data[i][6]) || 0
    });
  }

  if (stocks.length === 0) {
    Logger.log('MF Rankings: No valid stocks to rank');
    return;
  }

  // Rank by earnings yield (descending — highest EY = rank 1 = cheapest)
  stocks.sort(function(a, b) { return b.earningsYield - a.earningsYield; });
  for (var j = 0; j < stocks.length; j++) {
    stocks[j].eyRank = j + 1;
  }

  // Rank by ROTC (descending — highest ROTC = rank 1 = best business)
  stocks.sort(function(a, b) { return b.rotc - a.rotc; });
  for (var k = 0; k < stocks.length; k++) {
    stocks[k].rotcRank = k + 1;
  }

  // Combined rank (lower = better)
  for (var m = 0; m < stocks.length; m++) {
    stocks[m].combinedRank = stocks[m].eyRank + stocks[m].rotcRank;
  }

  // Sort by combined rank ascending
  stocks.sort(function(a, b) { return a.combinedRank - b.combinedRank; });

  // Take top 30
  var top30 = stocks.slice(0, 30);

  // Write to results sheet
  var resSheet = ss.getSheetByName(MF_RESULTS_SHEET);
  if (!resSheet) {
    resSheet = ss.insertSheet(MF_RESULTS_SHEET);
  }
  resSheet.clear();

  var header = [['Rank', 'Ticker', 'Name', 'Sector', 'MarketCap',
    'EarningsYield', 'ROTC', 'EY Rank', 'ROTC Rank',
    'CombinedRank', 'PE', 'LastUpdated']];
  resSheet.getRange(1, 1, 1, 12).setValues(header);
  resSheet.getRange('A1:L1').setFontWeight('bold');
  resSheet.setFrozenRows(1);

  var now = new Date().toISOString();
  var rows = [];
  for (var n = 0; n < top30.length; n++) {
    var s = top30[n];
    rows.push([
      n + 1, s.ticker, s.name, s.sector, s.marketCap,
      s.earningsYield, s.rotc, s.eyRank, s.rotcRank,
      s.combinedRank, s.pe, now
    ]);
  }

  if (rows.length > 0) {
    resSheet.getRange(2, 1, rows.length, 12).setValues(rows);
  }

  Logger.log('MF Rankings: Top 30 computed from ' + stocks.length + ' stocks');
}

// ---------------------------------------------------------------
// serveMagicFormulaJSON_ — Returns top 30 Magic Formula stocks
// as JSONP for Bloom frontend.
// ---------------------------------------------------------------
function serveMagicFormulaJSON_(callback) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MF_RESULTS_SHEET);

  if (!sheet || sheet.getLastRow() < 2) {
    return jsonpWrap_(JSON.stringify({
      status: 'ok', stocks: [], lastUpdated: null
    }), callback);
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
  var stocks = [];
  for (var i = 0; i < data.length; i++) {
    stocks.push({
      rank: Number(data[i][0]),
      ticker: String(data[i][1]),
      name: String(data[i][2]),
      sector: String(data[i][3]),
      marketCap: Number(data[i][4]),
      earningsYield: Number(data[i][5]),
      rotc: Number(data[i][6]),
      eyRank: Number(data[i][7]),
      rotcRank: Number(data[i][8]),
      combinedRank: Number(data[i][9]),
      pe: Number(data[i][10]),
      lastUpdated: String(data[i][11])
    });
  }

  return jsonpWrap_(JSON.stringify({
    status: 'ok',
    stocks: stocks,
    lastUpdated: stocks.length > 0 ? stocks[0].lastUpdated : null
  }), callback);
}

// ---------------------------------------------------------------
// serveQuoteJSON_ — Get live price for any ticker via Yahoo Finance.
// Used by Bloom to price Magic Formula stocks for buy flow.
// ---------------------------------------------------------------
function serveQuoteJSON_(ticker, callback) {
  ticker = String(ticker).toUpperCase().trim();
  if (!ticker) {
    return jsonpWrap_(JSON.stringify({ status: 'error', message: 'no ticker' }), callback);
  }

  try {
    var url = 'https://query1.finance.yahoo.com/v8/finance/chart/' +
      encodeURIComponent(ticker) + '?interval=1d&range=1d';
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(resp.getContentText());
    var meta = (json.chart && json.chart.result && json.chart.result[0] && json.chart.result[0].meta) || {};
    var price = meta.regularMarketPrice || meta.previousClose || 0;

    return jsonpWrap_(JSON.stringify({
      status: 'ok', ticker: ticker, price: price
    }), callback);
  } catch (err) {
    return jsonpWrap_(JSON.stringify({
      status: 'error', ticker: ticker, message: err.message
    }), callback);
  }
}
// ===============================================================
// TIDE — DSS Cycle Allocator (Apps Script backend)
// Add this code to your existing Bloom/SPX Apps Script.
// Then add the action handler and set up a daily trigger.
// ===============================================================

// --- CONSTANTS ---
var TIDE_SHEET = 'Tide DSS';
var TIDE_UNIVERSE_SHEET = 'Tide Universe';
var TIDE_PROGRESS_SHEET = 'Tide Progress';
var TIDE_BATCH_SIZE = 25; // tickers per batch

// --- DSS PARAMETERS (match Pine Script defaults) ---
var TIDE_STOCH_LEN = 10;
var TIDE_EMA_LEN = 5;
var TIDE_WT_1D = 20;
var TIDE_WT_3D = 50;
var TIDE_WT_1W = 30;
var TIDE_COMP_SMOOTH = 8;
var TIDE_SIG_SMOOTH = 5;
var TIDE_BUY_BELOW = 30;
var TIDE_SELL_ABOVE = 70;


// ===============================================================
// ADD THIS TO YOUR doGet() ACTION ROUTER:
// ===============================================================
//
//   if (action === 'tide_data') {
//     return serveTideDataJSON_(callback);
//   }
//
//   if (action === 'tide_universe') {
//     return serveTideUniverseJSON_(callback);
//   }
//


// ===============================================================
// STEP 1: Initialize the universe
// Run this ONCE to populate the Tide Universe sheet.
// Uses S&P 500 + Magic Formula stocks.
// ===============================================================
function initTideUniverse() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TIDE_UNIVERSE_SHEET);
  if (!sheet) sheet = ss.insertSheet(TIDE_UNIVERSE_SHEET);

  // S&P 500 core — ~200 most liquid names
  var tickers = [
    'AAPL','MSFT','GOOGL','AMZN','NVDA','META',
    'TSLA','BRK-B','JPM','V','UNH','JNJ','XOM',
    'PG','MA','HD','CVX','MRK','ABBV','PEP',
    'KO','COST','LLY','AVGO','WMT','MCD','CSCO',
    'CRM','AMD','ADBE','ACN','TMO','DHR','TXN',
    'ABT','NEE','PM','UNP','RTX','HON','AMGN',
    'IBM','LOW','QCOM','CAT','BA','GE','DE',
    'AMAT','GS','BLK','SPGI','AXP','MDLZ',
    'ISRG','GILD','ADI','LRCX','PLD','VRTX',
    'REGN','SYK','BKNG','PANW','ZTS','MMC',
    'CME','KLAC','SNPS','CDNS','NFLX','NXPI',
    'MRVL','ON','FTNT','CRWD','DDOG','SNOW',
    'CI','ELV','MCK','PSX','VLO','MPC','HES',
    'SLB','EOG','COP','OXY','DVN','FANG','HAL',
    'APD','SHW','LIN','ECL','DD','EMR','ETN',
    'ITW','ROK','AME','FTV','DOV','GWW','SWK',
    'PH','CMI','IR','PCAR','WAB','GD','LMT',
    'NOC','TDG','HWM','HII','TXT','LHX',
    'ADP','PAYX','CTAS','WM','RSG','VRSK',
    'ICE','MCO','MSCI','FIS','FISV','GPN',
    'PYPL','INTU','ANSS','KEYS','NOW','WDAY',
    'ZS','OKTA','MDB','TEAM','DOCU','HUBS',
    'WFC','BAC','C','USB','PNC','TFC','CFG',
    'KEY','RF','HBAN','FITB','MTB','ZION',
    'CL','EL','CHD','CLX','HRL','SJM','MKC',
    'HSY','K','GIS','CAG','CPB','TSN',
    'SBUX','YUM','DPZ','CMG','DKNG',
    'DIS','CMCSA','T','VZ','TMUS','CHTR',
    'EA','TTWO','WBD','PARA','NWSA',
    'NKE','TJX','ROST','BURL','DG','DLTR',
    'ORLY','AZO','AAP','BBY',
    'F','GM','RIVN','LCID',
    'PLTR','COIN','SQ','UBER','SHOP','ARM',
    'NET','TTD','RBLX','MELI','SE','NU',
    'SOFI','SMCI','AFRM','ROKU','SNAP',
    'PINS','HOOD','UPST','IONQ','RGTI',
    'MU','WDC','STX','HPQ','HPE','DELL',
    'VICI','O','SPG','AMT','CCI','EQIX',
    'DLR','PSA','EXR','AVB','MAA',
    'VTR','WELL','OHI','NNN',
  ];

  // Also pull Magic Formula top tickers if available
  var mfSheet = ss.getSheetByName('Magic Formula');
  if (mfSheet && mfSheet.getLastRow() > 1) {
    var mfData = mfSheet.getRange(2, 2,
      Math.min(mfSheet.getLastRow() - 1, 50), 1)
      .getValues();
    var tickerSet = {};
    tickers.forEach(function(t) { tickerSet[t] = true; });
    mfData.forEach(function(row) {
      var tk = String(row[0]).toUpperCase().trim();
      if (tk && !tickerSet[tk]) {
        tickers.push(tk);
        tickerSet[tk] = true;
      }
    });
  }

  // Write to sheet
  sheet.clear();
  var rows = [['Ticker']];
  tickers.forEach(function(t) { rows.push([t]); });
  sheet.getRange(1, 1, rows.length, 1).setValues(rows);
  sheet.getRange('A1').setFontWeight('bold');

  // Init progress
  var prog = ss.getSheetByName(TIDE_PROGRESS_SHEET);
  if (!prog) prog = ss.insertSheet(TIDE_PROGRESS_SHEET);
  prog.getRange('A1').setValue(0); // batch index
  prog.getRange('B1').setValue(
    Math.ceil(tickers.length / TIDE_BATCH_SIZE));
  prog.getRange('C1').setValue(new Date().toISOString());

  // Init results sheet
  var results = ss.getSheetByName(TIDE_SHEET);
  if (!results) results = ss.insertSheet(TIDE_SHEET);
  results.clear();
  results.getRange('A1:L1').setValues([[
    'Ticker','Price','Change%','Composite',
    'Signal','DSS_1D','DSS_3D','DSS_1W',
    'Zone','AboveSig','Updated','History'
  ]]);
  results.getRange('A1:L1').setFontWeight('bold');
  results.setFrozenRows(1);

  Logger.log('Tide: Universe initialized with ' +
    tickers.length + ' tickers');
}


// ===============================================================
// STEP 2: refreshTideDSSBatch — Run on a time-driven trigger
// (every 15 minutes, or daily). Processes one batch of tickers.
// When all batches done, resets for next cycle.
// ===============================================================
function refreshTideDSSBatch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var prog = ss.getSheetByName(TIDE_PROGRESS_SHEET);
  if (!prog) { Logger.log('Tide: No progress sheet'); return; }

  var batchIdx = Number(prog.getRange('A1').getValue()) || 0;
  var totalBatches = Number(prog.getRange('B1').getValue()) || 1;

  // Get universe
  var uSheet = ss.getSheetByName(TIDE_UNIVERSE_SHEET);
  if (!uSheet || uSheet.getLastRow() < 2) {
    Logger.log('Tide: No universe — run initTideUniverse()');
    return;
  }
  var allTickers = uSheet.getRange(2, 1,
    uSheet.getLastRow() - 1, 1).getValues()
    .map(function(r) { return String(r[0]).trim(); })
    .filter(function(t) { return t; });

  var start = batchIdx * TIDE_BATCH_SIZE;
  var batch = allTickers.slice(start,
    start + TIDE_BATCH_SIZE);

  if (batch.length === 0) {
    // All done — reset for next cycle
    prog.getRange('A1').setValue(0);
    prog.getRange('C1').setValue(
      new Date().toISOString());
    Logger.log('Tide: Full cycle complete. ' +
      allTickers.length + ' tickers processed.');
    return;
  }

  Logger.log('Tide: Batch ' + (batchIdx + 1) +
    '/' + totalBatches + ' — ' +
    batch.join(','));

  // Results sheet
  var results = ss.getSheetByName(TIDE_SHEET);
  if (!results) {
    results = ss.insertSheet(TIDE_SHEET);
    results.getRange('A1:L1').setValues([[
      'Ticker','Price','Change%','Composite',
      'Signal','DSS_1D','DSS_3D','DSS_1W',
      'Zone','AboveSig','Updated','History'
    ]]);
  }

  // Build map of existing rows for upsert
  var existingRows = {};
  if (results.getLastRow() > 1) {
    var existing = results.getRange(2, 1,
      results.getLastRow() - 1, 1).getValues();
    for (var e = 0; e < existing.length; e++) {
      existingRows[String(existing[e][0])] = e + 2;
    }
  }

  // Process each ticker in batch
  for (var i = 0; i < batch.length; i++) {
    var ticker = batch[i];
    try {
      var result = computeTideDSS_(ticker);
      if (!result) continue;

      var row = [
        ticker,
        result.price,
        result.changePct,
        result.composite,
        result.sigLine,
        result.dss1d,
        result.dss3d,
        result.dss1w,
        result.zone,
        result.aboveSig ? 'Y' : 'N',
        new Date().toISOString(),
        result.history.join(',')
      ];

      // Upsert
      if (existingRows[ticker]) {
        results.getRange(existingRows[ticker], 1, 1, 12)
          .setValues([row]);
      } else {
        results.appendRow(row);
        existingRows[ticker] = results.getLastRow();
      }

    } catch (err) {
      Logger.log('Tide: Error on ' + ticker +
        ': ' + err.message);
    }

    // Rate limit — Yahoo needs ~200ms between calls
    if (i < batch.length - 1) Utilities.sleep(300);
  }

  // Advance batch
  prog.getRange('A1').setValue(batchIdx + 1);
  Logger.log('Tide: Batch ' + (batchIdx + 1) +
    ' complete');
}


// ===============================================================
// computeTideDSS_ — Fetch 2y daily closes from Yahoo Finance,
// compute DSS Bressert on daily/3D/weekly, blend into composite.
// ===============================================================
function computeTideDSS_(ticker) {
  // Fetch 2 years of daily data (enough for DSS warmup)
  var url = 'https://query1.finance.yahoo.com/v8/finance/chart/' +
    encodeURIComponent(ticker) +
    '?range=2y&interval=1d&includePrePost=false';

  var resp = UrlFetchApp.fetch(url,
    { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return null;

  var json = JSON.parse(resp.getContentText());
  if (!json.chart || !json.chart.result ||
      !json.chart.result[0]) return null;

  var result = json.chart.result[0];
  var quote = result.indicators.quote[0];
  var rawCloses = quote.close || [];

  // Filter nulls
  var closes = [];
  for (var i = 0; i < rawCloses.length; i++) {
    if (rawCloses[i] !== null &&
        rawCloses[i] !== undefined) {
      closes.push(rawCloses[i]);
    }
  }

  if (closes.length < 60) return null; // need enough data

  var n = closes.length;
  var price = closes[n - 1];
  var prevPrice = closes[n - 2];
  var changePct = ((price - prevPrice) /
    prevPrice * 100);

  // Compute DSS for each timeframe
  var dss1d = dssBressert_(closes,
    TIDE_STOCH_LEN, TIDE_EMA_LEN);

  var closes3d = groupCandles_(closes, 3);
  var dss3dRaw = dssBressert_(closes3d,
    TIDE_STOCH_LEN, TIDE_EMA_LEN);
  var dss3d = expandToDaily_(dss3dRaw, 3, n);

  var closes1w = groupCandles_(closes, 5);
  var dss1wRaw = dssBressert_(closes1w,
    TIDE_STOCH_LEN, TIDE_EMA_LEN);
  var dss1w = expandToDaily_(dss1wRaw, 5, n);

  // Weighted composite
  var totalW = TIDE_WT_1D + TIDE_WT_3D + TIDE_WT_1W;
  var rawComp = [];
  for (var j = 0; j < n; j++) {
    rawComp.push(totalW > 0 ?
      (dss1d[j] * TIDE_WT_1D +
       dss3d[j] * TIDE_WT_3D +
       dss1w[j] * TIDE_WT_1W) / totalW : 50);
  }

  // Smooth
  var comp = TIDE_COMP_SMOOTH > 1 ?
    emaCalc_(rawComp, TIDE_COMP_SMOOTH) : rawComp;
  var sig = emaCalc_(comp, TIDE_SIG_SMOOTH);

  var last = n - 1;
  var c = comp[last];
  var s = sig[last];

  // Cycle check: must have been overbought in last
  // 120 bars to qualify for buy zone
  var lookback = Math.min(120, comp.length);
  var hadCycle = false;
  for (var lk = comp.length - lookback;
    lk < comp.length; lk++) {
    if (comp[lk] > TIDE_SELL_ABOVE) {
      hadCycle = true;
      break;
    }
  }

  // Zone — only BUY if it completed a cycle
  var zone = c < TIDE_BUY_BELOW && hadCycle ? 'BUY' :
    c > TIDE_SELL_ABOVE ? 'SELL' :
    c < TIDE_BUY_BELOW && !hadCycle ? 'WEAK' : 'HOLD';

  // Last 50 bars of composite for sparkline
  var history = [];
  for (var h = Math.max(0, comp.length - 50);
    h < comp.length; h++) {
    history.push(Math.round(comp[h] * 10) / 10);
  }

  return {
    price: Math.round(price * 100) / 100,
    changePct: Math.round(changePct * 100) / 100,
    composite: Math.round(c * 10) / 10,
    sigLine: Math.round(s * 10) / 10,
    dss1d: Math.round(dss1d[last] * 10) / 10,
    dss3d: Math.round(dss3d[last] * 10) / 10,
    dss1w: Math.round(dss1w[last] * 10) / 10,
    zone: zone,
    aboveSig: c > s,
    history: history
  };
}


// ===============================================================
// DSS Bressert — GAS implementation
// ===============================================================
function dssBressert_(closes, stochLen, emaLen) {
  var n = closes.length;

  // Pass 1: stochastic of close
  var raw1 = [];
  for (var i = 0; i < n; i++) {
    var start = Math.max(0, i - stochLen + 1);
    var hi = -Infinity, lo = Infinity;
    for (var k = start; k <= i; k++) {
      if (closes[k] > hi) hi = closes[k];
      if (closes[k] < lo) lo = closes[k];
    }
    raw1.push(hi !== lo ?
      (closes[i] - lo) / (hi - lo) * 100 : 50);
  }

  // EMA of raw1
  var sm1 = emaCalc_(raw1, emaLen);

  // Pass 2: stochastic of smoothed
  var raw2 = [];
  for (var j = 0; j < n; j++) {
    var s2 = Math.max(0, j - stochLen + 1);
    var h2 = -Infinity, l2 = Infinity;
    for (var m = s2; m <= j; m++) {
      if (sm1[m] > h2) h2 = sm1[m];
      if (sm1[m] < l2) l2 = sm1[m];
    }
    raw2.push(h2 !== l2 ?
      (sm1[j] - l2) / (h2 - l2) * 100 : 50);
  }

  return emaCalc_(raw2, emaLen);
}

function emaCalc_(arr, period) {
  if (!arr.length) return [];
  var k = 2.0 / (period + 1);
  var out = [arr[0]];
  for (var i = 1; i < arr.length; i++) {
    out.push(arr[i] * k + out[i - 1] * (1 - k));
  }
  return out;
}

function groupCandles_(daily, size) {
  var out = [];
  for (var i = 0; i < daily.length; i += size) {
    var end = Math.min(i + size, daily.length);
    out.push(daily[end - 1]);
  }
  return out;
}

// Expand HTF DSS to daily — no look-ahead.
// Days within a period use the PREVIOUS
// completed period's value.
function expandToDaily_(tfDss, groupSize, dailyLen) {
  var out = [];
  for (var i = 0; i < tfDss.length; i++) {
    var cnt = Math.min(groupSize,
      dailyLen - i * groupSize);
    var val = i > 0 ? tfDss[i-1] : tfDss[0];
    for (var j = 0; j < cnt - 1; j++) {
      out.push(val);
    }
    out.push(tfDss[i]); // last day = current
  }
  while (out.length < dailyLen) {
    out.push(out[out.length - 1] || 50);
  }
  return out.slice(0, dailyLen);
}


// ===============================================================
// SERVE tide_data — Returns all pre-computed DSS values
// ===============================================================
function serveTideDataJSON_(callback) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TIDE_SHEET);

  if (!sheet || sheet.getLastRow() < 2) {
    return jsonpWrap_(JSON.stringify({
      status: 'ok', stocks: [],
      lastUpdated: null
    }), callback);
  }

  var data = sheet.getRange(2, 1,
    sheet.getLastRow() - 1, 12).getValues();
  var stocks = [];
  for (var i = 0; i < data.length; i++) {
    var histStr = String(data[i][11]);
    var history = histStr ? histStr.split(',')
      .map(function(v) { return parseFloat(v) || 50; })
      : [];

    stocks.push({
      ticker: String(data[i][0]),
      price: Number(data[i][1]),
      changePct: Number(data[i][2]),
      composite: Number(data[i][3]),
      sigLine: Number(data[i][4]),
      dss1d: Number(data[i][5]),
      dss3d: Number(data[i][6]),
      dss1w: Number(data[i][7]),
      zone: String(data[i][8]),
      aboveSig: String(data[i][9]) === 'Y',
      updated: String(data[i][10]),
      history: history
    });
  }

  // Get MF rankings if available (for bonus weighting)
  var mfRanks = {};
  var mfSheet = ss.getSheetByName('Magic Formula');
  if (mfSheet && mfSheet.getLastRow() > 1) {
    var mfData = mfSheet.getRange(2, 1,
      Math.min(mfSheet.getLastRow() - 1, 30), 7)
      .getValues();
    for (var m = 0; m < mfData.length; m++) {
      var tk = String(mfData[m][1]).toUpperCase();
      mfRanks[tk] = {
        rank: Number(mfData[m][0]),
        earningsYield: Number(mfData[m][5]),
        rotc: Number(mfData[m][6])
      };
    }
  }

  return jsonpWrap_(JSON.stringify({
    status: 'ok',
    stocks: stocks,
    mfRanks: mfRanks,
    lastUpdated: new Date().toISOString()
  }), callback);
}


// ===============================================================
// SERVE tide_universe — Returns ticker list
// ===============================================================
function serveTideUniverseJSON_(callback) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TIDE_UNIVERSE_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    return jsonpWrap_(JSON.stringify({
      status: 'ok', tickers: []
    }), callback);
  }
  var data = sheet.getRange(2, 1,
    sheet.getLastRow() - 1, 1).getValues();
  var tickers = data.map(function(r) {
    return String(r[0]).trim();
  }).filter(function(t) { return t; });

  return jsonpWrap_(JSON.stringify({
    status: 'ok', tickers: tickers
  }), callback);
}


// ===============================================================
// SETUP INSTRUCTIONS:
// ===============================================================
//
// 1. Paste this code into your existing Apps Script project
//
// 2. Add these lines to your doGet() function's action router:
//
//    if (action === 'tide_data') {
//      return serveTideDataJSON_(callback);
//    }
//    if (action === 'tide_universe') {
//      return serveTideUniverseJSON_(callback);
//    }
//
// 3. Run initTideUniverse() once to create the sheets
//
// 4. Set up a time-driven trigger:
//    - Function: refreshTideDSSBatch
//    - Trigger: Every 15 minutes (processes 25 tickers per run)
//    - Full cycle for 200 tickers = ~2 hours
//    - Or run it every 5 min for faster refresh
//
// 5. Redeploy the Apps Script web app
//
// 6. Update Tide's index.html API_BASE to your endpoint
//

// ============================================================================
// Dashboard Additions (v2026-09-14)
// Stats engine, dashboard cache, close handler, state sync, API key gate.
// These wrap doGet_orig / doPost_orig above.
// ============================================================================

var DASHBOARD_ADDITIONS_VERSION = 'v2026-09-14-additions';

var DASH_STATS_ENGINE = true;             // false = leave the original script's numbers untouched
var DASH_POSITIONS_SHEET_ = 'Positions';  // tab name; found by its headers (Ticker + Outcome) if renamed
var DASH_LEG_OUTCOME_ = 'Closed';         // word written on the extra DCA legs of a closed position

var DASH_CACHE_TTL_SEC = 600;          // 10 min ceiling; writes and webhooks invalidate sooner
var DASH_CACHE_KEY_ = 'spx_dash_v1';
var DASH_CACHE_CHUNK_ = 90000;         // CacheService caps a value at 100 KB

// GET actions that change the sheet. Gated by API_KEY once it is set.
var DASHBOARD_WRITE_ACTIONS_ = [
  'update_position', 'add_trade', 'set_stats_start', 'set_review', 'toggle_dca',
  'set_state', 'add_portfolio', 'save_dca_data', 'save_analytics_csv'
];
// Subset that changes what ?action=dashboard returns
var DASH_INVALIDATING_ACTIONS_ = ['update_position', 'add_trade', 'set_stats_start', 'toggle_dca', 'add_portfolio'];

var DASHBOARD_STATE_SHEET_ = 'Dashboard State'; // hidden tab, A1 holds the JSON blob

// ---------------------------------------------------------------------------
// Entry points (wrappers around the originals above)
// ---------------------------------------------------------------------------

function doGet(e) {
  var extra = handleDashboardExtras_(e);
  if (extra) return extra;

  // Bloom cloud storage (save_dca_data / load_dca_data / dca_save_chunk / dca_save_done)
  // — chunked multi-cell layout, see the BLOOM STORAGE section at the end of this file.
  var bloom = BLOOM_handle(e);
  if (bloom) return bloom;

  var p = (e && e.parameter) || {};
  var action = String(p.action || '');

  if (action === 'dashboard') return serveDashboard_(e, String(p.nocache || '') === '1');
  if (action === 'diag') return dashJson_(diagPositions_(p));
  if (DASH_STATS_ENGINE && action === 'update_position' && String(p.field || '') === 'outcome') {
    var closed = closePositionRows_(p);
    invalidateDashboardCache_();
    return dashJson_(closed);
  }
  if (DASH_STATS_ENGINE && action === 'update_position' && String(p.field || '') === 'reopen') {
    var reopened = reopenPositionRows_(p);
    invalidateDashboardCache_();
    return dashJson_(reopened);
  }

  var out = doGet_orig(e);
  if (DASH_INVALIDATING_ACTIONS_.indexOf(action) !== -1) invalidateDashboardCache_();
  return out;
}

function doPost(e) {
  var out = doPost_orig(e);
  invalidateDashboardCache_(); // a new signal changes Action Needed / Recent Signals
  return out;
}

// ---------------------------------------------------------------------------
// Dashboard cache
// ---------------------------------------------------------------------------

function serveDashboard_(e, bypass) {
  var t0 = Date.now();
  if (!bypass) {
    var hit = readDashCache_();
    if (hit) return dashJson_(annotateDash_(hit.json, { hit: true, builtAt: hit.builtAt, servedMs: Date.now() - t0 }));
  }
  var out = doGet_orig(e);
  var json = (out && typeof out.getContent === 'function') ? out.getContent() : null;
  if (!json) return out; // not a TextOutput — hand it back untouched
  if (DASH_STATS_ENGINE) json = applyStatsEngine_(json);
  var builtAt = new Date().toISOString();
  var buildMs = Date.now() - t0;
  try { writeDashCache_(json, builtAt); } catch (err) { /* cache is best-effort */ }
  return dashJson_(annotateDash_(json, { hit: false, builtAt: builtAt, buildMs: buildMs }));
}

function annotateDash_(json, info) {
  try {
    var obj = JSON.parse(json);
    if (obj && typeof obj === 'object' && !Array.isArray(obj)) {
      obj._cache = info;
      return JSON.stringify(obj);
    }
  } catch (err) { /* fall through */ }
  return json;
}

function writeDashCache_(jsonStr, builtAt) {
  var b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(jsonStr)).getBytes());
  var chunks = [];
  for (var i = 0; i < b64.length; i += DASH_CACHE_CHUNK_) chunks.push(b64.slice(i, i + DASH_CACHE_CHUNK_));
  var entries = {};
  chunks.forEach(function (c, idx) { entries[DASH_CACHE_KEY_ + '_' + idx] = c; });
  var cache = CacheService.getScriptCache();
  cache.putAll(entries, DASH_CACHE_TTL_SEC);
  cache.put(DASH_CACHE_KEY_ + '_meta', JSON.stringify({ n: chunks.length, builtAt: builtAt }), DASH_CACHE_TTL_SEC);
}

function readDashCache_() {
  var cache = CacheService.getScriptCache();
  var metaRaw = cache.get(DASH_CACHE_KEY_ + '_meta');
  if (!metaRaw) return null;
  var meta;
  try { meta = JSON.parse(metaRaw); } catch (err) { return null; }
  var keys = [];
  for (var i = 0; i < meta.n; i++) keys.push(DASH_CACHE_KEY_ + '_' + i);
  var parts = cache.getAll(keys);
  var b64 = '';
  for (i = 0; i < meta.n; i++) {
    if (!parts[keys[i]]) return null;
    b64 += parts[keys[i]];
  }
  try {
    var json = Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip')).getDataAsString();
    return { json: json, builtAt: meta.builtAt };
  } catch (err) {
    return null;
  }
}

function invalidateDashboardCache_() {
  try { CacheService.getScriptCache().remove(DASH_CACHE_KEY_ + '_meta'); } catch (err) { /* ignore */ }
}

function warmDashboardCache_() {
  serveDashboard_({ parameter: { action: 'dashboard' }, parameters: { action: ['dashboard'] }, queryString: 'action=dashboard' }, true);
}
function installDashboardCacheWarmer() {
  removeDashboardCacheWarmer();
  ScriptApp.newTrigger('warmDashboardCache_').timeBased().everyMinutes(5).create();
  Logger.log('Dashboard cache warmer installed (every 5 min).');
}
function removeDashboardCacheWarmer() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'warmDashboardCache_') ScriptApp.deleteTrigger(t);
  });
}

// ---------------------------------------------------------------------------
// Stats engine: reads the Positions tab, decides open/closed and win/loss
// ---------------------------------------------------------------------------

function dashNorm_(v) { return String(v == null ? '' : v).trim().toLowerCase(); }
function dashNum_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var n = parseFloat(String(v == null ? '' : v).replace(/[$,\s]/g, ''));
  return isFinite(n) ? n : 0;
}
function dashSide_(signal) {
  var s = dashNorm_(signal);
  return (s === 'sell' || s === 'reduce' || s === 'short' || s === 'cover') ? 'sell' : 'buy';
}
function dashIsClosedOutcome_(o) {
  var s = dashNorm_(o);
  return s !== '' && s !== '0x0' && s !== 'oxo' && s !== 'open';
}
function dashClassify_(r) {
  if (r.profit > 0) return 'win';
  if (r.profit < 0) return 'loss';
  var o = dashNorm_(r.outcome);
  if (o.indexOf('tp') !== -1 || o.indexOf('win') !== -1 || o.indexOf('profit') !== -1) return 'win';
  if (o.indexOf('stop') !== -1 || o.indexOf('loss') !== -1 || o.indexOf('lost') !== -1) return 'loss';
  return 'flat';
}
function dashFmtMoney_(n) { return (n < 0 ? '-$' : '$') + Math.abs(n).toFixed(2); }
function dashFmtDate_(v) {
  var d = (v instanceof Date) ? v : new Date(v);
  if (isNaN(d.getTime())) return String(v || '');
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function dashPositions_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;
  var sh = ss.getSheetByName(DASH_POSITIONS_SHEET_);
  if (!sh) {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var lc = sheets[i].getLastColumn();
      if (lc < 2) continue;
      var h = sheets[i].getRange(1, 1, 1, Math.min(lc, 26)).getValues()[0].map(dashNorm_);
      if (h.some(function (x) { return x.indexOf('ticker') !== -1; }) && h.some(function (x) { return x.indexOf('outcome') !== -1; })) { sh = sheets[i]; break; }
    }
  }
  if (!sh) return null;
  var values = sh.getDataRange().getValues();
  var hdr = (values[0] || []).map(dashNorm_);
  function col(keys, fallback) {
    for (var k = 0; k < keys.length; k++) {
      for (var c = 0; c < hdr.length; c++) if (hdr[c] && hdr[c].indexOf(keys[k]) !== -1) return c;
    }
    return fallback;
  }
  var cols = {
    ts:      col(['timestamp', 'time', 'date'], 0),
    ticker:  col(['ticker', 'symbol'], 1),
    signal:  col(['signal'], 2),
    price:   col(['price', 'entry'], 3),
    action:  col(['action'], 4),
    outcome: col(['outcome', 'result'], 5),
    size:    col(['size'], 6),
    profit:  col(['profit', 'p&l', 'pnl', 'locked'], 7),
    notes:   col(['note', 'comment'], 8)
  };
  var headerLooksLikeHeader = hdr.some(function (x) { return x.indexOf('ticker') !== -1 || x.indexOf('outcome') !== -1; });
  var start = headerLooksLikeHeader ? 1 : 0;
  var rows = [];
  for (var r = start; r < values.length; r++) {
    var v = values[r];
    var ticker = String(v[cols.ticker] == null ? '' : v[cols.ticker]).trim();
    if (!ticker) continue;
    rows.push({
      row: r + 1,
      ts: v[cols.ts],
      ticker: ticker,
      signal: dashNorm_(v[cols.signal]),
      side: dashSide_(v[cols.signal]),
      price: dashNum_(v[cols.price]),
      action: String(v[cols.action] == null ? '' : v[cols.action]).trim(),
      outcome: String(v[cols.outcome] == null ? '' : v[cols.outcome]).trim(),
      size: dashNum_(v[cols.size]),
      profit: dashNum_(v[cols.profit]),
      notes: String(v[cols.notes] == null ? '' : v[cols.notes])
    });
  }
  return { sheet: sh, name: sh.getName(), cols: cols, header: values[0] || [], rows: rows };
}

function dashCloseTime_(r) {
  var m = /closed (\d{4}-\d{2}-\d{2} \d{2}:\d{2})/.exec(r.notes || '');
  var d = m ? new Date(m[1].replace(' ', 'T')) : ((r.ts instanceof Date) ? r.ts : new Date(r.ts));
  return isNaN(d.getTime()) ? 0 : d.getTime();
}
function dashByCloseTime_(a, b) { return (dashCloseTime_(a) - dashCloseTime_(b)) || (a.row - b.row); }
function dashIsEntered_(r) { return dashNorm_(r.action) === 'entered'; }
function dashIsOpen_(r) { return dashIsEntered_(r) && !dashIsClosedOutcome_(r.outcome); }
function dashIsClosed_(r) { return dashIsEntered_(r) && dashIsClosedOutcome_(r.outcome); }

function dashComputeStats_(pos, statsStartDate) {
  var start = statsStartDate ? new Date(statsStartDate) : null;
  if (start && isNaN(start.getTime())) start = null;
  var closed = pos.rows.filter(dashIsClosed_).filter(function (r) {
    if (!start) return true;
    var d = (r.ts instanceof Date) ? r.ts : new Date(r.ts);
    return isNaN(d.getTime()) ? true : d >= start;
  }).sort(dashByCloseTime_);
  var wins = 0, losses = 0, net = 0, gw = 0, gl = 0;
  var cur = 0, curType = '', bestW = 0, worstL = 0;
  var byTicker = {};
  closed.forEach(function (r) {
    var k = dashClassify_(r);
    net += r.profit;
    byTicker[r.ticker] = (byTicker[r.ticker] || 0) + r.profit;
    if (k === 'win') { wins++; gw += r.profit; }
    else if (k === 'loss') { losses++; gl += r.profit; }
    if (k === 'flat') return;
    if (k === curType) cur++; else { curType = k; cur = 1; }
    if (k === 'win' && cur > bestW) bestW = cur;
    if (k === 'loss' && cur > worstL) worstL = cur;
  });
  var best = '', worst = '', bestV = 0, worstV = 0;
  Object.keys(byTicker).forEach(function (t) {
    if (byTicker[t] > bestV) { bestV = byTicker[t]; best = t; }
    if (byTicker[t] < worstV) { worstV = byTicker[t]; worst = t; }
  });
  var openKeys = {};
  pos.rows.filter(dashIsOpen_).forEach(function (r) { openKeys[r.ticker.toUpperCase() + '_' + r.side] = true; });
  var decided = wins + losses;
  return {
    totalTrades: decided,
    wins: wins,
    losses: losses,
    winRate: (decided ? (wins / decided * 100) : 0).toFixed(1) + '%',
    netPnl: dashFmtMoney_(net),
    totalProfit: dashFmtMoney_(gw),
    totalLost: dashFmtMoney_(gl),
    bestWinStreak: bestW,
    worstLossStreak: worstL,
    currentStreak: cur ? (cur + (curType === 'win' ? 'W' : 'L')) : '0',
    bestTicker: best,
    worstTicker: worst,
    openPositions: Object.keys(openKeys).length,
    _closedCount: closed.length
  };
}

function applyStatsEngine_(json) {
  var obj;
  try { obj = JSON.parse(json); } catch (e) { return json; }
  if (!obj || typeof obj !== 'object') return json;
  try {
    var pos = dashPositions_();
    if (!pos) { obj._stats_engine = 'error: Positions tab not found'; return JSON.stringify(obj); }
    obj.stats = obj.stats || {};
    var st = dashComputeStats_(pos, obj.stats.statsStartDate);
    Object.keys(st).forEach(function (k) { if (k.charAt(0) !== '_') obj.stats[k] = st[k]; });

    var openRows = {};
    pos.rows.filter(dashIsOpen_).forEach(function (r) {
      var k = r.ticker.toUpperCase() + '_' + r.side;
      (openRows[k] = openRows[k] || []).push(r);
    });
    var lastClosed = {};
    pos.rows.filter(dashIsClosed_).forEach(function (r) { lastClosed[r.ticker.toUpperCase() + '_' + r.side] = r; });
    var seen = {};
    obj.tickers = Array.isArray(obj.tickers) ? obj.tickers : [];
    obj.tickers.forEach(function (t) {
      if (!t || t.ticker == null) return;
      var tk = String(t.ticker).toUpperCase();
      seen[tk] = true;
      ['buy', 'sell'].forEach(function (side) {
        var k = tk + '_' + side, statusKey = side + 'Status';
        var open = openRows[k];
        if (open && open.length) {
          t[statusKey] = 'OPEN';
          if (!t[side + 'Price']) t[side + 'Price'] = open[0].price;
          t[side + 'LastSignal'] = open[open.length - 1].signal || t[side + 'LastSignal'] || side;
          t[side + 'ProfitLocked'] = open.reduce(function (a, r) { return a + r.profit; }, 0);
          t[side + 'Outcome'] = open[open.length - 1].outcome || '';
        } else if (String(t[statusKey] || '').toUpperCase() === 'OPEN') {
          var lc = lastClosed[k];
          var k2 = lc ? dashClassify_(lc) : 'flat';
          t[statusKey] = k2 === 'win' ? 'PROFIT' : k2 === 'loss' ? 'STOPPED' : 'IDLE';
        }
      });
    });
    Object.keys(openRows).forEach(function (k) {
      var tk = k.split('_')[0], side = k.split('_')[1];
      if (seen[tk]) return;
      var open = openRows[k];
      var t = { ticker: open[0].ticker };
      t[side + 'Status'] = 'OPEN';
      t[side + 'Price'] = open[0].price;
      t[side + 'LastSignal'] = open[open.length - 1].signal || side;
      t[side + 'ProfitLocked'] = open.reduce(function (a, r) { return a + r.profit; }, 0);
      t[side + 'Phase'] = 'Manual';
      t[side + 'NextSize'] = 0;
      t[(side === 'buy' ? 'sell' : 'buy') + 'Status'] = 'IDLE';
      obj.tickers.push(t);
      seen[tk] = true;
    });

    var wins = pos.rows.filter(dashIsClosed_).filter(function (r) { return dashClassify_(r) === 'win'; }).sort(dashByCloseTime_);
    obj.closedTrades = wins.reverse().slice(0, 60).map(function (r) {
      return { ticker: r.ticker, profit: r.profit, date: dashFmtDate_(r.ts), signal: r.signal };
    });
    obj._stats_engine = 'ok: ' + st._closedCount + ' closed rows, ' + Object.keys(openRows).length + ' open positions, sheet "' + pos.name + '"';
  } catch (err) {
    obj._stats_engine = 'error: ' + (err && err.message ? err.message : err);
  }
  return JSON.stringify(obj);
}

function closePositionRows_(p) {
  var pos = dashPositions_();
  if (!pos) return { status: 'error', message: 'Positions tab not found' };
  var ticker = String(p.ticker || '').trim(), side = dashSide_(p.signal), outcome = String(p.value || 'Closed').trim();
  var pnl = dashNum_(p.profitLocked);
  var open = pos.rows.filter(function (r) { return dashIsOpen_(r) && r.ticker.toUpperCase() === ticker.toUpperCase() && r.side === side; });
  if (!open.length) return { status: 'error', message: 'No open Entered row for ' + ticker + ' ' + side.toUpperCase() + ' - nothing to close' };
  var sh = pos.sheet, c = pos.cols, main = open[open.length - 1];
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var lock = LockService.getScriptLock();
  try { lock.waitLock(5000); } catch (e) { return { status: 'error', message: 'sheet is busy, try again' }; }
  try {
    open.forEach(function (r) {
      var isMain = r.row === main.row;
      sh.getRange(r.row, c.outcome + 1).setValue(isMain ? outcome : DASH_LEG_OUTCOME_);
      sh.getRange(r.row, c.profit + 1).setValue(isMain ? pnl : 0);
      // Update timestamp to close date so stats filter includes this trade
      sh.getRange(r.row, c.ts + 1).setValue(new Date());
      var note = (r.notes ? r.notes + ' | ' : '') + (isMain ? 'closed ' + stamp + (open.length > 1 ? ' (' + open.length + ' legs)' : '') : 'leg of ' + ticker + ' close ' + stamp + ' - P&L on row ' + main.row);
      sh.getRange(r.row, c.notes + 1).setValue(note);
    });
  } finally { lock.releaseLock(); }
  return { status: 'ok', ticker: ticker, side: side, outcome: outcome, profitLocked: pnl, row: main.row, rows: open.map(function (r) { return r.row; }), closedAt: stamp };
}

function reopenPositionRows_(p) {
  var pos = dashPositions_();
  if (!pos) return { status: 'error', message: 'Positions tab not found' };
  var ticker = String(p.ticker || '').trim(), side = dashSide_(p.signal);
  var closed = pos.rows.filter(function (r) { return dashIsClosed_(r) && r.ticker.toUpperCase() === ticker.toUpperCase() && r.side === side; });
  if (!closed.length) return { status: 'error', message: 'No closed row for ' + ticker + ' ' + side.toUpperCase() };
  var main = closed[closed.length - 1];
  var legs = pos.rows.filter(function (r) { return r.notes.indexOf('P&L on row ' + main.row) !== -1; });
  var sh = pos.sheet, c = pos.cols;
  [main].concat(legs).forEach(function (r) {
    sh.getRange(r.row, c.outcome + 1).setValue('');
    sh.getRange(r.row, c.profit + 1).setValue('');
    sh.getRange(r.row, c.notes + 1).setValue((r.notes ? r.notes + ' | ' : '') + 'reopened');
  });
  return { status: 'ok', ticker: ticker, side: side, rows: [main.row].concat(legs.map(function (r) { return r.row; })) };
}

function diagPositions_(p) {
  var pos = dashPositions_();
  if (!pos) return { status: 'error', message: 'Positions tab not found (looked for "' + DASH_POSITIONS_SHEET_ + '" and any tab with Ticker + Outcome headers)' };
  var tk = String(p.ticker || '').trim().toUpperCase();
  var rows = tk ? pos.rows.filter(function (r) { return r.ticker.toUpperCase() === tk; }) : pos.rows.slice(-15);
  return {
    status: 'ok',
    engine: DASH_STATS_ENGINE,
    sheet: pos.name,
    header: pos.header,
    columns: pos.cols,
    totalRows: pos.rows.length,
    stats: dashComputeStats_(pos, p.statsStartDate || ''),
    rows: rows.slice(-25).map(function (r) {
      return { row: r.row, ts: dashFmtDate_(r.ts), ticker: r.ticker, signal: r.signal, side: r.side, price: r.price, action: r.action, outcome: r.outcome, profit: r.profit,
               open: dashIsOpen_(r), counts: dashIsClosed_(r) ? dashClassify_(r) : '-' };
    })
  };
}

// ---------------------------------------------------------------------------
// version / key gate / state sync
// ---------------------------------------------------------------------------

function handleDashboardExtras_(e) {
  var p = (e && e.parameter) || {};
  var action = String(p.action || '');

  if (action === 'version') {
    return dashJson_({
      status: 'ok',
      script_version: DASHBOARD_ADDITIONS_VERSION,
      auth: !!dashApiKey_(),
      cache_ttl_sec: DASH_CACHE_TTL_SEC,
      stats_engine: DASH_STATS_ENGINE,
      server_time: new Date().toISOString()
    });
  }

  var key = dashApiKey_();
  if (key && DASHBOARD_WRITE_ACTIONS_.indexOf(action) !== -1 && String(p.key || '') !== key) {
    return dashJson_({ status: 'error', message: 'Unauthorized: missing or wrong key for ' + action });
  }

  if (action === 'get_state') {
    return dashJson_({ status: 'ok', state: dashReadState_() });
  }

  if (action === 'set_state') {
    var incoming;
    try { incoming = JSON.parse(p.data || '{}'); }
    catch (err) { return dashJson_({ status: 'error', message: 'set_state: data is not valid JSON' }); }
    if (!incoming || typeof incoming !== 'object') return dashJson_({ status: 'error', message: 'set_state: expected an object' });

    var lock = LockService.getScriptLock();
    try { lock.waitLock(5000); }
    catch (err) { return dashJson_({ status: 'error', message: 'set_state: sheet is busy, try again' }); }
    try {
      var current = dashReadState_();
      Object.keys(incoming).forEach(function (k) {
        var inc = incoming[k];
        if (!inc || typeof inc.v !== 'string') return;
        var incTs = Number(inc.ts || 0);
        if (!current[k] || incTs > Number(current[k].ts || 0)) current[k] = { v: inc.v, ts: incTs };
      });
      dashWriteState_(current);
      return dashJson_({ status: 'ok', state: current });
    } finally {
      lock.releaseLock();
    }
  }

  return null;
}

function dashJson_(objOrString) {
  var text = (typeof objOrString === 'string') ? objOrString : JSON.stringify(objOrString);
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

function dashApiKey_() {
  return PropertiesService.getScriptProperties().getProperty('API_KEY') || '';
}

function dashStateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;
  var sh = ss.getSheetByName(DASHBOARD_STATE_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(DASHBOARD_STATE_SHEET_);
    sh.getRange('A1').setValue('{}');
    sh.getRange('B1').setValue('JSON blob written by the dashboard (get_state / set_state). Do not edit by hand.');
    sh.hideSheet();
  }
  return sh;
}

function dashReadState_() {
  try {
    var sh = dashStateSheet_();
    var raw = sh ? String(sh.getRange('A1').getValue() || '{}')
                 : (PropertiesService.getScriptProperties().getProperty('CLIENT_STATE') || '{}');
    var obj = JSON.parse(raw);
    return (obj && typeof obj === 'object') ? obj : {};
  } catch (e) {
    return {};
  }
}

function dashWriteState_(obj) {
  var raw = JSON.stringify(obj);
  var sh = dashStateSheet_();
  if (sh) sh.getRange('A1').setValue(raw);
  else PropertiesService.getScriptProperties().setProperty('CLIENT_STATE', raw);
}

// Run ONCE from the editor to enable API key gating on writes.
function setDashboardApiKey() {
  var key = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('API_KEY', key);
  Logger.log('API_KEY set. Paste this into app.js →  const API_KEY = \'' + key + '\';');
}

// ===============================================================
// BLOOM STORAGE — chunked cloud storage for Bloom Portfolios
// ===============================================================
// Replaces the old single-cell "DCA Backup" layout, which failed once the
// blob grew past Google Sheets' 50,000-character-per-cell limit:
//   "Your input contains more than the maximum of 50000 characters in a single cell."
//
// The Bloom state blob (portfolios + history + dividends + snapshots + opps flags)
// is now stored as 40,000-char chunks, one per row, in a "BloomData" tab. The
// previous save is kept as a one-deep backup, the old layout is migrated
// automatically on first use, and the daily snapshot trigger runs on top of it.
//
// The client (index.html) needs NO changes — same actions, params and responses.
//
// Wiring (already done in this file):
//   • doGet() calls BLOOM_handle(e) right after handleDashboardExtras_(e)
//   • the existing time-based trigger keeps calling dailyDcaSnapshot()
//
// Editor helpers:
//   BLOOM_status()              – what's stored (chunks, length, snapshots…)
//   BLOOM_restorePreviousSave() – swap the previous save back in (one-deep undo)
//   dailyDcaSnapshot()          – run today's snapshot now
//
// Layout of the BloomData tab:
//   A1 chunks     B1 <count>      C1 prevChunks     D1 <count>
//   A2 lastSaved  B2 <ISO time>   C2 prevLastSaved  D2 <ISO time>
//   A3 length     B3 <chars>      C3 prevLength     D3 <chars>
//   A4 note       B4 <text>
//   A6.. current chunks (prefixed "~" so Sheets never parses them as numbers/formulas)
//   C6.. previous save's chunks (backup)
// ===============================================================

var BLOOM_SHEET_NAME     = 'BloomData'; // created automatically
var BLOOM_LOG_SHEET_NAME = 'BloomLog';  // created automatically; keeps the last 200 events
var BLOOM_CELL_CHUNK     = 40000;       // chars per cell (Sheets limit is 50,000)
var BLOOM_FIRST_ROW      = 6;           // chunk rows start here (rows 1–4 are metadata)
var BLOOM_MAX_SNAPSHOTS  = 365;         // mirrors the client's cap
var BLOOM_CHUNK_TTL_SEC  = 1800;        // how long client upload chunks live in cache

// ---------------------------------------------------------------
// Router — called from doGet(). Returns null for actions it doesn't own.
// ---------------------------------------------------------------
function BLOOM_handle(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || '';
  try {
    switch (action) {
      case 'load_dca_data': {
        var blob = BLOOM_readBlob();
        if (!blob) return BLOOM_respond(e, { status: 'empty' });
        return BLOOM_respond(e, { status: 'ok', data: blob, lastSaved: BLOOM_getLastSaved() });
      }

      case 'save_dca_data': {
        var json = p.data || '';
        var check = BLOOM_validateBlob(json);
        if (check) return BLOOM_respond(e, { status: 'error', message: check });
        BLOOM_ensureMigrated_();
        var info = BLOOM_writeBlob(json);
        return BLOOM_respond(e, { status: 'ok', length: info.length, chunks: info.chunks, lastSaved: info.lastSaved });
      }

      case 'dca_save_chunk': {
        var idx = parseInt(p.i, 10);
        if (isNaN(idx) || idx < 0) return BLOOM_respond(e, { status: 'error', message: 'bad chunk index' });
        CacheService.getScriptCache().put('bloom_chunk_' + idx, String(p.cd || ''), BLOOM_CHUNK_TTL_SEC);
        return BLOOM_respond(e, { status: 'ok', i: idx });
      }

      case 'dca_save_done': {
        var n = parseInt(p.n, 10);
        if (isNaN(n) || n <= 0) return BLOOM_respond(e, { status: 'error', message: 'bad chunk count' });
        var assembled = BLOOM_assembleChunks(n);
        if (assembled.error) return BLOOM_respond(e, { status: 'error', message: assembled.error });
        var check2 = BLOOM_validateBlob(assembled.json);
        if (check2) return BLOOM_respond(e, { status: 'error', message: check2 });
        BLOOM_ensureMigrated_();
        var info2 = BLOOM_writeBlob(assembled.json);
        BLOOM_clearChunkCache(n);
        return BLOOM_respond(e, { status: 'ok', length: info2.length, chunks: info2.chunks, lastSaved: info2.lastSaved });
      }

      default:
        return null; // not ours — the rest of doGet handles it
    }
  } catch (err) {
    BLOOM_log('ERROR in ' + action + ': ' + (err && err.message ? err.message : err));
    return BLOOM_respond(e, { status: 'error', message: String(err && err.message ? err.message : err) });
  }
}

function BLOOM_respond(e, obj) {
  var json = JSON.stringify(obj);
  var cb = e && e.parameter && e.parameter.callback;
  if (cb && /^[\w$.]+$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------
// Chunked upload (client sends 1200-char pieces, then dca_save_done)
// ---------------------------------------------------------------
function BLOOM_assembleChunks(n) {
  var cache = CacheService.getScriptCache();
  var keys = [];
  for (var i = 0; i < n; i++) keys.push('bloom_chunk_' + i);
  // The page-hide "beacon" path fires all chunks at once, so a couple may still be
  // landing when dca_save_done arrives. Retry briefly before giving up.
  var parts = null, missing = [];
  for (var attempt = 0; attempt < 4; attempt++) {
    var got = cache.getAll(keys);
    missing = [];
    parts = [];
    for (var j = 0; j < n; j++) {
      var v = got['bloom_chunk_' + j];
      if (v === null || v === undefined) missing.push(j); else parts.push(v);
    }
    if (missing.length === 0) break;
    Utilities.sleep(700);
  }
  if (missing.length) return { error: 'missing chunk(s): ' + missing.slice(0, 10).join(',') };
  return { json: parts.join('') };
}

function BLOOM_clearChunkCache(n) {
  var keys = [];
  for (var i = 0; i < n; i++) keys.push('bloom_chunk_' + i);
  try { CacheService.getScriptCache().removeAll(keys); } catch (e) {}
}

function BLOOM_validateBlob(json) {
  if (!json || typeof json !== 'string') return 'empty payload';
  var data;
  try { data = JSON.parse(json); } catch (e) { return 'payload is not valid JSON'; }
  if (!data || typeof data !== 'object' || !data.portfolios) return 'payload has no portfolios';
  return null;
}

// ---------------------------------------------------------------
// Blob read / write
// ---------------------------------------------------------------
function BLOOM_getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BLOOM_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(BLOOM_SHEET_NAME);
    sh.getRange('A1:A4').setValues([['chunks'], ['lastSaved'], ['length'], ['note']]);
    sh.getRange('C1:C3').setValues([['prevChunks'], ['prevLastSaved'], ['prevLength']]);
    sh.getRange('B1').setValue(0);
    sh.getRange('D1').setValue(0);
    // Force plain-text format on the chunk columns so nothing is ever auto-parsed.
    sh.getRange('A' + BLOOM_FIRST_ROW + ':A').setNumberFormat('@');
    sh.getRange('C' + BLOOM_FIRST_ROW + ':C').setNumberFormat('@');
    sh.setColumnWidth(1, 120);
    sh.setColumnWidth(3, 120);
  }
  return sh;
}

function BLOOM_getLastSaved() {
  var v = BLOOM_getSheet().getRange('B2').getValue();
  if (!v) return null;
  if (v instanceof Date) return v.toISOString();
  return String(v);
}

/** Runs the one-time migration if the new layout is still empty (no-op otherwise). */
function BLOOM_ensureMigrated_() {
  var sh = BLOOM_getSheet();
  var n = parseInt(sh.getRange('B1').getValue(), 10) || 0;
  if (n <= 0) BLOOM_migrateLegacy(sh);
}

/** Returns the current blob as a JSON string, or '' if nothing is stored. */
function BLOOM_readBlob(opts) {
  var sh = BLOOM_getSheet();
  var n = parseInt(sh.getRange('B1').getValue(), 10) || 0;
  if (n <= 0) {
    var migrated = BLOOM_migrateLegacy(sh, opts);
    if (!migrated) return '';
    n = parseInt(sh.getRange('B1').getValue(), 10) || 0;
    if (n <= 0) return '';
  }
  var vals = sh.getRange(BLOOM_FIRST_ROW, 1, n, 1).getValues();
  var parts = [];
  for (var i = 0; i < vals.length; i++) parts.push(BLOOM_unprefix(vals[i][0]));
  return parts.join('');
}

/**
 * Writes the blob (JSON string) into chunk rows. Keeps the previous save in column C.
 * opts.keepLastSaved = true leaves the lastSaved timestamp alone (used by the daily
 * snapshot so a server-side write never outranks the user's own device in the client's
 * "which copy wins" comparison). opts.noLock = true when the caller already holds the
 * script lock.
 */
function BLOOM_writeBlob(json, opts) {
  opts = opts || {};
  var lock = opts.noLock ? null : LockService.getScriptLock();
  if (lock) lock.waitLock(30000);
  try {
    var sh = BLOOM_getSheet();
    var lastRow = Math.max(sh.getLastRow(), BLOOM_FIRST_ROW);
    var span = lastRow - BLOOM_FIRST_ROW + 1;

    // 1. Move current save to the backup column
    var curN = parseInt(sh.getRange('B1').getValue(), 10) || 0;
    if (curN > 0) {
      var cur = sh.getRange(BLOOM_FIRST_ROW, 1, curN, 1).getValues();
      sh.getRange(BLOOM_FIRST_ROW, 3, span, 1).clearContent();
      sh.getRange(BLOOM_FIRST_ROW, 3, curN, 1).setValues(cur);
      sh.getRange('D1:D3').setValues([[curN], [sh.getRange('B2').getValue()], [sh.getRange('B3').getValue()]]);
    }

    // 2. Write the new chunks
    var chunks = [];
    for (var i = 0; i < json.length; i += BLOOM_CELL_CHUNK) chunks.push(['~' + json.substring(i, i + BLOOM_CELL_CHUNK)]);
    sh.getRange(BLOOM_FIRST_ROW, 1, span, 1).clearContent();
    sh.getRange(BLOOM_FIRST_ROW, 1, chunks.length, 1).setValues(chunks);

    var lastSaved = opts.keepLastSaved ? (sh.getRange('B2').getValue() || new Date().toISOString()) : new Date().toISOString();
    if (lastSaved instanceof Date) lastSaved = lastSaved.toISOString();
    sh.getRange('B1:B3').setValues([[chunks.length], [lastSaved], [json.length]]);
    SpreadsheetApp.flush();
    return { chunks: chunks.length, length: json.length, lastSaved: lastSaved };
  } finally {
    if (lock) lock.releaseLock();
  }
}

function BLOOM_unprefix(v) {
  var s = (v === null || v === undefined) ? '' : String(v);
  return s.charAt(0) === '~' ? s.substring(1) : s;
}

// ---------------------------------------------------------------
// One-time migration from the old "DCA Backup" layout
//   A1 = blob without snapshots, D1 = snapshots array, B1 = last-saved timestamp
// (falls back to scanning every tab for a JSON cell containing "portfolios")
// The old cells are left untouched.
// ---------------------------------------------------------------
function BLOOM_migrateLegacy(sh, opts) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var found = null;

  // Preferred: the known DCA Backup layout
  var legacy = ss.getSheetByName(DCA_BACKUP_SHEET);
  if (legacy) {
    var a1 = legacy.getRange('A1').getValue();
    if (typeof a1 === 'string' && a1.charAt(0) === '{') {
      try {
        var data = JSON.parse(a1);
        if (data && data.portfolios) {
          var snaps = [];
          var d1 = legacy.getRange('D1').getValue();
          if (typeof d1 === 'string' && d1.charAt(0) === '[') {
            try { snaps = JSON.parse(d1); } catch (e1) { snaps = []; }
          }
          if (!Array.isArray(snaps)) snaps = [];
          // Very old layout kept snapshots inside A1 — merge both, D1 wins on same date
          var byDate = {};
          (Array.isArray(data.valueSnapshots) ? data.valueSnapshots : []).forEach(function (s) { if (s && s.date) byDate[s.date] = s; });
          snaps.forEach(function (s) { if (s && s.date) byDate[s.date] = s; });
          var merged = Object.keys(byDate).sort().map(function (k) { return byDate[k]; });
          if (merged.length > BLOOM_MAX_SNAPSHOTS) merged = merged.slice(-BLOOM_MAX_SNAPSHOTS);
          data.valueSnapshots = merged;

          var b1 = legacy.getRange('B1').getValue();
          var lastSaved = b1 ? (b1 instanceof Date ? b1.toISOString() : String(b1))
                             : (data.lastModified ? new Date(data.lastModified).toISOString() : new Date().toISOString());
          found = { json: JSON.stringify(data), lastSaved: lastSaved, where: DCA_BACKUP_SHEET + '!A1 + D1' };
        }
      } catch (e2) { /* fall through to the generic scan */ }
    }
  }

  // Fallback: any cell anywhere that holds a Bloom blob
  if (!found) {
    var best = null;
    var sheets = ss.getSheets();
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      if (sheet.getName() === BLOOM_SHEET_NAME || sheet.getName() === BLOOM_LOG_SHEET_NAME) continue;
      if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) continue;
      var vals = sheet.getDataRange().getValues();
      for (var r = 0; r < vals.length; r++) {
        for (var c = 0; c < vals[r].length; c++) {
          var v = vals[r][c];
          if (typeof v !== 'string' || v.length < 20) continue;
          if (v.indexOf('"portfolios"') === -1 || v.charAt(0) !== '{') continue;
          var parsed;
          try { parsed = JSON.parse(v); } catch (e3) { continue; }
          if (!parsed || !parsed.portfolios) continue;
          if (!best || v.length > best.json.length) {
            best = { json: v, data: parsed, where: sheet.getName() + '!' + sheet.getRange(r + 1, c + 1).getA1Notation() };
          }
        }
      }
    }
    if (best) {
      found = {
        json: best.json,
        lastSaved: best.data.lastModified ? new Date(best.data.lastModified).toISOString() : new Date().toISOString(),
        where: best.where
      };
    }
  }

  if (!found) { BLOOM_log('Migration: no legacy blob found (fresh install)'); return false; }

  var info = BLOOM_writeBlob(found.json, { noLock: !!(opts && opts.noLock) });
  sh.getRange('B2').setValue(found.lastSaved);
  sh.getRange('B4').setValue('migrated from ' + found.where + ' on ' + new Date().toISOString());
  BLOOM_log('Migrated legacy blob from ' + found.where + ' (' + info.length + ' chars → ' + info.chunks + ' chunks)');
  return true;
}

// ---------------------------------------------------------------
// One-deep undo
// ---------------------------------------------------------------
function BLOOM_restorePreviousSave() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = BLOOM_getSheet();
    var prevN = parseInt(sh.getRange('D1').getValue(), 10) || 0;
    if (prevN <= 0) { Logger.log('No previous save to restore.'); return; }
    var prev = sh.getRange(BLOOM_FIRST_ROW, 3, prevN, 1).getValues();
    var curN = parseInt(sh.getRange('B1').getValue(), 10) || 0;
    var cur = curN > 0 ? sh.getRange(BLOOM_FIRST_ROW, 1, curN, 1).getValues() : [];
    var lastRow = Math.max(sh.getLastRow(), BLOOM_FIRST_ROW);
    var span = lastRow - BLOOM_FIRST_ROW + 1;
    var metaB = sh.getRange('B1:B3').getValues();
    var metaD = sh.getRange('D1:D3').getValues();
    sh.getRange(BLOOM_FIRST_ROW, 1, span, 1).clearContent();
    sh.getRange(BLOOM_FIRST_ROW, 3, span, 1).clearContent();
    sh.getRange(BLOOM_FIRST_ROW, 1, prevN, 1).setValues(prev);
    if (cur.length) sh.getRange(BLOOM_FIRST_ROW, 3, cur.length, 1).setValues(cur);
    sh.getRange('B1:B3').setValues(metaD);
    sh.getRange('D1:D3').setValues(metaB);
    SpreadsheetApp.flush();
    BLOOM_log('Restored previous save (' + metaD[2][0] + ' chars). The replaced save is now the backup.');
    Logger.log('Restored previous save. Open Bloom and tap ↓ Restore to pull it onto your device.');
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------
// dailyDcaSnapshot — time-based trigger (Day timer, 4pm–5pm)
// Records each portfolio's value/cost for today into valueSnapshots.
// ---------------------------------------------------------------
function dailyDcaSnapshot() {
  try {
    // Pass 1 (no lock): find out which tickers are held and fetch their prices. Price
    // lookups can take a while, and we don't want to block client saves during them.
    var json = BLOOM_readBlob();
    if (!json) { BLOOM_log('Snapshot skipped: no data stored yet'); return; }
    var tickers = BLOOM_heldTickers(JSON.parse(json));
    if (!tickers.length) { BLOOM_log('Snapshot skipped: no holdings'); return; }

    var prices = BLOOM_getPrices(tickers);
    var priced = tickers.filter(function (t) { return prices[t] > 0; });
    if (!priced.length) { BLOOM_log('Snapshot ABORTED: no prices available for ' + tickers.length + ' tickers'); return; }
    var missing = tickers.filter(function (t) { return !(prices[t] > 0); });

    // Pass 2 (locked, short): re-read the latest blob, add today's snapshot, write it back.
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(60000)) { BLOOM_log('Snapshot skipped: storage busy'); return; }
    try {
      json = BLOOM_readBlob({ noLock: true });
      var data = JSON.parse(json);
      if (!data.portfolios) { BLOOM_log('Snapshot skipped: blob has no portfolios'); return; }

      // Same UTC-date convention as the client (new Date().toISOString().slice(0,10))
      var today = new Date().toISOString().slice(0, 10);
      var snap = { date: today, portfolios: {} };
      Object.keys(data.portfolios).forEach(function (pk) {
        var h = (data.portfolios[pk] && data.portfolios[pk].holdings) || {};
        var value = 0, cost = 0;
        Object.keys(h).forEach(function (t) {
          var shares = Number(h[t] && h[t].shares) || 0;
          value += shares * (prices[t] || 0);
          cost += Number(h[t] && h[t].costBasis) || 0;
        });
        snap.portfolios[pk] = { value: Math.round(value * 100) / 100, cost: Math.round(cost * 100) / 100 };
      });

      // Replace any same-day entry (the post-close value is the better one), then cap.
      var snaps = (data.valueSnapshots || []).filter(function (s) { return s && s.date && s.date !== today; });
      snaps.push(snap);
      snaps.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
      if (snaps.length > BLOOM_MAX_SNAPSHOTS) snaps = snaps.slice(-BLOOM_MAX_SNAPSHOTS);
      data.valueSnapshots = snaps;

      // Do NOT touch data.lastModified — that stamp belongs to the user's devices.
      var info = BLOOM_writeBlob(JSON.stringify(data), { keepLastSaved: true, noLock: true });

      var totals = Object.keys(snap.portfolios).map(function (pk) { return pk + '=$' + snap.portfolios[pk].value; }).join(' ');
      BLOOM_log('Snapshot ' + today + ' saved (' + info.length + ' chars, ' + info.chunks + ' chunks, ' + snaps.length + ' snapshots). ' + totals +
        (missing.length ? ' | NO PRICE for: ' + missing.join(',') : ''));
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    BLOOM_log('Snapshot ERROR: ' + (err && err.stack ? err.stack : err));
    throw err; // keep Google's failure email as a backstop
  }
}

/** Tickers with shares > 0 across all portfolios (avoids 60+ quote calls for empty Flywheel slots). */
function BLOOM_heldTickers(data) {
  var needed = {};
  Object.keys((data && data.portfolios) || {}).forEach(function (pk) {
    var h = (data.portfolios[pk] && data.portfolios[pk].holdings) || {};
    Object.keys(h).forEach(function (t) { if ((Number(h[t] && h[t].shares) || 0) > 0) needed[t] = true; });
  });
  return Object.keys(needed);
}

// ---------------------------------------------------------------
// Prices for the snapshot
//   1. serveDcaPricesJSON_ — the same numbers the app shows (GOOGLEFINANCE + cached technicals)
//   2. serveQuoteJSON_ per missing ticker (Yahoo Finance), with BRK.B → BRK-B style aliasing
// ---------------------------------------------------------------
function BLOOM_getPrices(tickers) {
  var out = {};
  try {
    var resp = JSON.parse(serveDcaPricesJSON_('').getContent());
    if (resp && resp.etfs && resp.etfs.length) {
      resp.etfs.forEach(function (x) {
        if (x && x.ticker && Number(x.price) > 0) out[x.ticker] = Number(x.price);
      });
    }
  } catch (e) { BLOOM_log('dca_prices lookup failed: ' + (e && e.message ? e.message : e)); }

  var started = Date.now();
  tickers.forEach(function (t) {
    if (out[t] > 0) return;
    var alias = String(t).replace(/\./g, '-');        // holdings may say BRK.B, prices say BRK-B
    if (alias !== t && out[alias] > 0) { out[t] = out[alias]; return; }
    if (Date.now() - started > 200000) return;       // stay well inside the 6-minute trigger limit
    try {
      var q = JSON.parse(serveQuoteJSON_(alias, '').getContent());
      if (q && Number(q.price) > 0) out[t] = Number(q.price);
    } catch (e2) {}
  });
  return out;
}

// ---------------------------------------------------------------
// Logging — BloomLog tab, last 200 events
// ---------------------------------------------------------------
function BLOOM_log(msg) {
  Logger.log(msg);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(BLOOM_LOG_SHEET_NAME) || ss.insertSheet(BLOOM_LOG_SHEET_NAME);
    sh.appendRow([new Date(), String(msg)]);
    var rows = sh.getLastRow();
    if (rows > 200) sh.deleteRows(1, rows - 200);
  } catch (e) {}
}

// ---------------------------------------------------------------
// Status — run from the editor
// ---------------------------------------------------------------
function BLOOM_status() {
  var json = BLOOM_readBlob();
  if (!json) { Logger.log('Nothing stored yet.'); return; }
  var sh = BLOOM_getSheet();
  var data = JSON.parse(json);
  var holdings = 0;
  Object.keys(data.portfolios || {}).forEach(function (pk) {
    var h = (data.portfolios[pk] && data.portfolios[pk].holdings) || {};
    Object.keys(h).forEach(function (t) { if ((Number(h[t] && h[t].shares) || 0) > 0) holdings++; });
  });
  var snaps = data.valueSnapshots || [];
  Logger.log([
    'Blob length:   ' + json.length + ' chars in ' + sh.getRange('B1').getValue() + ' chunk(s)',
    'lastSaved:     ' + BLOOM_getLastSaved(),
    'lastModified:  ' + (data.lastModified ? new Date(data.lastModified).toISOString() : '(none)'),
    'Note:          ' + (sh.getRange('B4').getValue() || '(none)'),
    'Portfolios:    ' + Object.keys(data.portfolios || {}).join(', '),
    'Held tickers:  ' + holdings,
    'History:       ' + (data.history || []).length + ' entries',
    'Dividends:     ' + (data.dividendHistory || []).length + ' entries',
    'Snapshots:     ' + snaps.length + (snaps.length ? ' (' + snaps[0].date + ' → ' + snaps[snaps.length - 1].date + ')' : ''),
    'Backup:        ' + (parseInt(sh.getRange('D1').getValue(), 10) || 0) + ' chunk(s) from ' + (sh.getRange('D2').getValue() || '(none)')
  ].join('\n'));
}
