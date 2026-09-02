/*******************************************************
 * GTPro Signal DB — 00_Config.gs
 * ระบบเก็บสัญญาณ GTPro + Paper Trade + แจ้งเตือน Telegram
 * Backend: Google Apps Script + Google Sheets (ฟรี 100%)
 *******************************************************/

// ชื่อชีตทั้งหมด
var SHEETS = {
  CONFIG:  'Config',
  MAP:     'Symbol_Map',
  RAW:     'Signals_Raw',
  TRADES:  'Trades',
  DAILY:   'Daily_PL',
  STATS:   'Stats'
};

// หัวคอลัมน์ (ห้ามสลับลำดับ — โค้ดอ้างอิงตำแหน่งนี้)
var HEADERS = {
  RAW: ['signal_id','received_at','source','symbol','tf','action','price','tp','sl','note','status','raw'],
  TRADES: ['trade_id','symbol','tf','side','lots','entry_time','entry_price','tp','sl',
           'exit_time','exit_price','exit_reason','pnl_pts','pnl_pct','pnl_usd',
           'r_multiple','duration_min','status','signal_open','signal_close',
           'mode','chart_id','note'],
  DAILY: ['date','mode','trades','wins','losses','winrate_pct','pnl_usd','cum_pnl_usd',
          'best_usd','worst_usd'],
  STATS: ['scope','trades','wins','losses','winrate_pct','avg_win_usd','avg_loss_usd',
          'profit_factor','expectancy_usd','max_dd_usd','total_pnl_usd','updated_at'],
  MAP:  ['symbol','provider','api_symbol','value_per_point','digits','note']
};

// ค่าเริ่มต้นของชีต Config (แก้ค่าได้ในชีต ไม่ต้องแก้โค้ด)
var DEFAULT_CONFIG = [
  ['TELEGRAM_TOKEN',      '',                  'Token จาก @BotFather'],
  ['TELEGRAM_CHAT_ID',    '',                  'Chat ID ปลายทาง (รันฟังก์ชัน findMyChatId)'],
  ['WEBHOOK_SECRET',      'CHANGE_ME_123',     'รหัสลับกันคนยิง Webhook มั่ว'],
  ['EMAIL_FROM',          'noreply@tradingview.com', 'ผู้ส่งอีเมล Alert (Path B)'],
  ['EMAIL_LOOKBACK_MIN',  '15',                'ย้อนอ่านเมลกี่นาที'],
  ['DEFAULT_LOTS',        '1',                 'ขนาดไม้เริ่มต้น'],
  ['ALLOW_PYRAMID',       'FALSE',             'TRUE = เปิดไม้ซ้อนทิศเดียวกันได้'],
  ['CLOSE_ON_OPPOSITE',   'TRUE',              'TRUE = สัญญาณตรงข้ามปิดไม้เดิมทันที'],
  ['NOTIFY_OPEN',         'TRUE',              'แจ้งเตือนตอนเปิดไม้'],
  ['NOTIFY_CLOSE',        'TRUE',              'แจ้งเตือนตอนปิดไม้'],
  ['DAILY_SUMMARY_HOUR',  '23',                'ชั่วโมงที่ส่งสรุปรายวัน (0-23)'],
  ['TZ',                  'Asia/Bangkok',      'โซนเวลา'],
  ['DEFAULT_MODE',        'LIVE',              'LIVE = เทรดจริง / TEST = backtest (เปลี่ยนด้วย /mode)']
];

// ตำแหน่งคอลัมน์ในชีต Trades ที่โค้ดอ้างอิงบ่อย (นับจาก 1)
var COL_STATUS = 18;
var COL_MODE   = 21;
var COL_CHART  = 22;
var COL_NOTE   = 23;

var MODE_LIVE = 'LIVE';
var MODE_TEST = 'TEST';

var CACHE = {};

/** อ่านค่าจากชีต Config (cache ในรอบการรันเดียว) */
function cfg(key, fallback) {
  if (!CACHE.config) {
    CACHE.config = {};
    var sh = ss().getSheetByName(SHEETS.CONFIG);
    if (sh) {
      var rows = sh.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0]) CACHE.config[String(rows[i][0]).trim()] = rows[i][1];
      }
    }
  }
  var v = CACHE.config[key];
  if (v === undefined || v === '') return (fallback !== undefined ? fallback : '');
  return v;
}

function cfgBool(key, fallback) {
  var v = String(cfg(key, fallback ? 'TRUE' : 'FALSE')).trim().toUpperCase();
  return v === 'TRUE' || v === 'YES' || v === '1';
}

function cfgNum(key, fallback) {
  var v = parseFloat(cfg(key, fallback));
  return isNaN(v) ? fallback : v;
}

function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function tz() {
  return cfg('TZ', 'Asia/Bangkok');
}

/** จัดรูปแบบตัวเลขพร้อม comma separator */
function fmt(n, digits) {
  if (n === null || n === undefined || isNaN(n)) return '-';
  digits = (digits === undefined) ? 2 : digits;
  var s = Number(n).toFixed(digits);
  var parts = s.split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return parts.join('.');
}

function fmtSigned(n, digits) {
  if (n === null || n === undefined || isNaN(n)) return '-';
  return (n >= 0 ? '+' : '') + fmt(n, digits);
}

function nowStr() {
  return Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd HH:mm:ss');
}

function dateKey(d) {
  return Utilities.formatDate(d, tz(), 'yyyy-MM-dd');
}

/** เติมศูนย์หน้าเลขให้ครบ 2 หลัก — ใช้ทั้งปฏิทินและตัวอ่านวันที่ */
function pad2_(n) {
  var s = String(n);
  return s.length < 2 ? '0' + s : s;
}

/** ทำให้ค่าโหมดเป็นมาตรฐานเสมอ — คืน 'LIVE' หรือ 'TEST' */
function normMode(v) {
  var s = String(v || '').trim().toUpperCase();
  if (s === 'TEST' || s === 'BT' || s === 'BACKTEST') return MODE_TEST;
  return MODE_LIVE;
}

/** โหมดเริ่มต้นที่ตั้งไว้ในชีต Config */
function defaultMode() {
  return normMode(cfg('DEFAULT_MODE', MODE_LIVE));
}

/*******************************************************
 * 01_Setup.gs — สร้างชีต ตั้ง Trigger เมนู
 * ▶ รันฟังก์ชัน setup() ครั้งเดียวตอนติดตั้ง
 *******************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ GTPro')
    .addItem('1. ติดตั้งระบบ (setup)', 'setup')
    .addItem('2. หา Telegram Chat ID', 'findMyChatId')
    .addItem('3. ทดสอบส่ง Telegram', 'testTelegram')
    .addItem('4. เปิดรับสัญญาณทาง Telegram', 'installTelegramTrigger')
    .addItem('5. สร้าง API key สำหรับ Dashboard', 'generateApiKey')
    .addItem('6. อัปเกรดชีตให้รองรับรูป + โหมด Test', 'migrateAddModeAndChart')
    .addItem('7. เปิดเตือนข่าวเศรษฐกิจ', 'installNewsTrigger')
    .addSeparator()
    .addItem('อ่านข้อความ Telegram เดี๋ยวนี้', 'pollTelegram')
    .addItem('ตรวจ TP/SL เดี๋ยวนี้', 'monitorOpenTrades')
    .addItem('อ่านอีเมล Alert เดี๋ยวนี้', 'pollGmail')
    .addItem('คำนวณสถิติใหม่ทั้งหมด', 'rebuildAll')
    .addItem('สร้างปฏิทินใหม่', 'buildCalendar')
    .addItem('ส่งสรุปรายวันเดี๋ยวนี้', 'sendDailySummary')
    .addItem('เก็บกวาดรูปที่ไม่ได้ใช้', 'cleanupOrphanCharts')
    .addItem('ทดสอบดึงข่าวเดี๋ยวนี้', 'testNews')
    .addSeparator()
    .addItem('ใส่ข้อมูลตัวอย่าง (ทดสอบ)', 'seedDemoData')
    .addToUi();
}

function setup() {
  var s = ss();

  // ---- Config ----
  var cf = ensureSheet(SHEETS.CONFIG, ['key','value','คำอธิบาย']);
  if (cf.getLastRow() < 2) {
    cf.getRange(2, 1, DEFAULT_CONFIG.length, 3).setValues(DEFAULT_CONFIG);
  }
  cf.setColumnWidth(1, 190); cf.setColumnWidth(2, 260); cf.setColumnWidth(3, 320);

  // ---- Symbol_Map ----
  var mp = ensureSheet(SHEETS.MAP, HEADERS.MAP);
  if (mp.getLastRow() < 2) {
    mp.getRange(2, 1, 6, 6).setValues([
      ['XAUUSD',  'yahoo',   'GC=F',     100, 2, 'ทองคำ 1 lot = 100 oz → 1 จุด = 100 USD'],
      ['BTCUSDT', 'binance', 'BTCUSDT',    1, 2, 'Binance spot'],
      ['ETHUSDT', 'binance', 'ETHUSDT',    1, 2, 'Binance spot'],
      ['EURUSD',  'yahoo',   'EURUSD=X', 100000, 5, '1 lot = 100,000 → 1 pip = 10 USD'],
      ['SET',     'manual',  '',           1, 2, 'ดัชนี SET — Yahoo ไม่มีข้อมูล ต้องส่ง exit เอง'],
      ['*',       'yahoo',   '',           1, 2, 'ค่า default สำหรับคู่ที่ไม่ได้ระบุ']
    ]);
  }

  // ---- ตารางข้อมูล ----
  ensureSheet(SHEETS.RAW,    HEADERS.RAW);
  ensureSheet(SHEETS.TRADES, HEADERS.TRADES);
  ensureSheet(SHEETS.DAILY,  HEADERS.DAILY);
  ensureSheet(SHEETS.STATS,  HEADERS.STATS);

  installTriggers_();
  try { buildCalendar(); } catch (e) { Logger.log(e); }

  SpreadsheetApp.getActive().toast('ติดตั้งเสร็จแล้ว — ไปกรอก TELEGRAM_TOKEN ในชีต Config ต่อ', 'GTPro', 10);
}

function ensureSheet(name, headers) {
  var s = ss();
  var sh = s.getSheetByName(name);
  if (!sh) sh = s.insertSheet(name);
  if (headers && headers.length) {
    // เขียนหัวคอลัมน์ใหม่ถ้ามีอะไรไม่ตรง — รองรับกรณีเพิ่มคอลัมน์ทีหลัง
    var cur = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    var same = true;
    for (var i = 0; i < headers.length; i++) {
      if (String(cur[i] || '') !== headers[i]) { same = false; break; }
    }
    if (!same) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function installTriggers_() {
  // ลบ trigger เดิมกันซ้ำ
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (['pollGmail','monitorOpenTrades','sendDailySummary','rebuildAll'].indexOf(f) >= 0) {
      ScriptApp.deleteTrigger(t);
    }
  });
  // หมายเหตุ: pollTelegram ไม่ถูกลบตรงนี้ — ติดตั้งแยกด้วย installTelegramTrigger

  ScriptApp.newTrigger('pollGmail').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('monitorOpenTrades').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('sendDailySummary').timeBased()
    .atHour(cfgNum('DAILY_SUMMARY_HOUR', 23)).everyDays(1).inTimezone(tz()).create();
}

/**
 * อัปเกรดชีตเดิมให้มีคอลัมน์ mode / chart_id / note
 * ปลอดภัยกับข้อมูลเดิม เพราะเพิ่มต่อท้าย ไม่ขยับคอลัมน์เก่า
 */
function migrateAddModeAndChart() {
  // 1) Trades — เพิ่มคอลัมน์ mode / chart_id / note แล้วเติมไม้เดิมเป็น LIVE
  var sh = ensureSheet(SHEETS.TRADES, HEADERS.TRADES);
  var last = sh.getLastRow();
  if (last >= 2) {
    var rng = sh.getRange(2, COL_MODE, last - 1, 1);
    var v = rng.getValues();
    var filled = 0;
    for (var i = 0; i < v.length; i++) {
      if (!v[i][0]) { v[i][0] = MODE_LIVE; filled++; }
    }
    if (filled) rng.setValues(v);
  }

  // 2) Config — เพิ่มแถว DEFAULT_MODE
  ensureSheet(SHEETS.CONFIG, ['key','value','คำอธิบาย']);
  if (!cfg('DEFAULT_MODE', '')) {
    ss().getSheetByName(SHEETS.CONFIG)
      .appendRow(['DEFAULT_MODE', 'LIVE', 'LIVE = เทรดจริง / TEST = backtest (เปลี่ยนด้วย /mode)']);
    CACHE.config = null;
  }

  // 3) Daily_PL — โครงเปลี่ยน (แทรกคอลัมน์ mode) ล้างแล้วสร้างใหม่จาก Trades
  //    ปลอดภัย เพราะชีตนี้เป็นข้อมูลที่คำนวณได้ทั้งหมด ไม่ใช่ต้นฉบับ
  var dsh = ss().getSheetByName(SHEETS.DAILY);
  if (dsh && dsh.getLastRow() > 1) {
    dsh.getRange(2, 1, dsh.getLastRow() - 1, Math.max(dsh.getLastColumn(), 10)).clearContent();
  }
  ensureSheet(SHEETS.DAILY, HEADERS.DAILY);

  rebuildAll();
  SpreadsheetApp.getActive().toast(
    'อัปเกรดเรียบร้อย — ไม้เดิมทั้งหมดเป็น LIVE, พร้อมรับรูปและโหมด TEST แล้ว', 'GTPro', 10);
}

/** ล้างข้อมูลทดสอบทั้งหมด (ไม่แตะ Config / Symbol_Map) */
function clearData() {
  [SHEETS.RAW, SHEETS.TRADES, SHEETS.DAILY, SHEETS.STATS].forEach(function (n) {
    var sh = ss().getSheetByName(n);
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    }
  });
  SpreadsheetApp.getActive().toast('ล้างข้อมูลแล้ว', 'GTPro', 5);
}

/*******************************************************
 * 02_Parser.gs — แปลงข้อความ Alert เป็น object สัญญาณ
 * ฟังก์ชันในไฟล์นี้เป็น pure function ทั้งหมด (ทดสอบแยกได้)
 *******************************************************/

var ACTION_CLOSE = ['CLOSE','EXIT','FLAT','CLOSED','ปิด','ปิดไม้','ปิดออเดอร์'];
var ACTION_BUY   = ['BUY','LONG','BULL','BULLISH','ซื้อ','ขึ้น'];
var ACTION_SELL  = ['SELL','SHORT','BEAR','BEARISH','ขาย','ลง'];

/** แปลงสตริงตัวเลขที่อาจมี comma → number */
function toNum(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  var s = String(v).replace(/,/g, '').replace(/[^0-9.\-]/g, '');
  if (s === '' || s === '-' || s === '.') return null;
  var n = parseFloat(s);
  return isFinite(n) ? n : null;
}

/** หา action จากข้อความ — ตรวจ CLOSE ก่อนเสมอ */
function detectAction(text) {
  var t = ' ' + String(text).toUpperCase() + ' ';
  function hit(list) {
    for (var i = 0; i < list.length; i++) {
      var w = list[i].toUpperCase();
      // ใช้ขอบเขตคำสำหรับอักษรละติน, ใช้ indexOf สำหรับภาษาไทย
      if (/^[A-Z]+$/.test(w)) {
        if (new RegExp('(^|[^A-Z])' + w + '([^A-Z]|$)').test(t)) return true;
      } else if (t.indexOf(w) >= 0) return true;
    }
    return false;
  }
  if (hit(ACTION_CLOSE)) return 'CLOSE';
  if (hit(ACTION_BUY))   return 'BUY';
  if (hit(ACTION_SELL))  return 'SELL';
  return null;
}

/**
 * ดึงค่าตัวเลขจาก key ที่มีได้หลายชื่อ เช่น tp / take profit / เป้า
 * รองรับทั้ง "tp: 1600", "tp=1600" และ "tp 1600" (เว้นวรรคเฉย ๆ)
 * มีขอบเขตคำกันจับผิด เช่น "sl" ใน "slippage"
 */
function grab(text, names) {
  for (var i = 0; i < names.length; i++) {
    var re = new RegExp('(?:^|[^A-Za-z0-9_])(?:' + names[i] +
                        ')\\s*[:=]?\\s*"?([\\-0-9][0-9,\\.]*)', 'i');
    var m = String(text).match(re);
    if (m) return toNum(m[1]);
  }
  return null;
}

function grabStr(text, names) {
  for (var i = 0; i < names.length; i++) {
    var re = new RegExp('(?:^|[^A-Za-z0-9_])(?:' + names[i] +
                        ')\\s*[:=]?\\s*"?([A-Za-z0-9_\\.\\:\\-]+)', 'i');
    var m = String(text).match(re);
    if (m) return m[1];
  }
  return null;
}

// คำที่หน้าตาเหมือนชื่อสัญลักษณ์แต่ไม่ใช่ — กันเดาผิด เช่น "GTPRO BUY signal on BTCUSDT"
var SYMBOL_STOPWORDS = [
  'GTPRO','GTPROEMA','ALERT','ALERTS','TRADINGVIEW','SIGNAL','SIGNALS','NOTIFY',
  'ENTRY','PRICE','PROFIT','STOP','LOSS','TAKE','TARGET','CLOSE','OPEN','EXIT',
  'BUY','SELL','LONG','SHORT','ORDER','ACTION','SYMBOL','TICKER','TIMEFRAME',
  'INTERVAL','CHART','TREND','SUPPLY','DEMAND','BULL','BEAR','SCORE','INDICATOR',
  'THIS','THAT','WAS','SENT','FROM','WITH','YOUR','HAS','BEEN','THE','AND','FOR',
  'INDEX','IDX','POINT','POINTS','PIPS','RISK','LOT','LOTS','NEW','NOW','ALL','ANY',
  'TEST','BT','BACKTEST','LIVE','MODE','REAL','DEMO','PAPER','NOTE','MEMO'
];

/* ---------- โหมด LIVE / TEST ---------- */

var MODE_TEST_WORDS = ['TEST','BT','BACKTEST','PAPER','ทดสอบ','ย้อน'];
var MODE_LIVE_WORDS = ['LIVE','REAL','จริง'];

/**
 * หาโหมดจากข้อความ — คืน 'TEST' | 'LIVE' | null (ไม่ได้ระบุ ให้ใช้ค่า default)
 * ตรวจ TEST ก่อน เพราะข้อความมักเขียนว่า "test" ต่อท้ายสัญญาณจริง ๆ ไม่ได้
 */
function detectMode(text) {
  var t = ' ' + String(text).toUpperCase() + ' ';
  function hit(list) {
    for (var i = 0; i < list.length; i++) {
      var w = list[i].toUpperCase();
      if (/^[A-Z]+$/.test(w)) {
        if (new RegExp('(^|[^A-Z])' + w + '([^A-Z]|$)').test(t)) return true;
      } else if (t.indexOf(w) >= 0) return true;
    }
    return false;
  }
  if (hit(MODE_TEST_WORDS)) return MODE_TEST;
  if (hit(MODE_LIVE_WORDS)) return MODE_LIVE;
  return null;
}

/* ---------- วันเวลาย้อนหลัง (สำหรับบันทึกไม้ backtest) ---------- */

var WHEN_STRIP_RE = /@\s*\d{1,4}[-\/]\d{1,2}[-\/]\d{1,4}(?:[\sT]+\d{1,2}[:\.]\d{2})?/g;

/**
 * อ่านวันเวลาย้อนหลังจากข้อความ — เขียนนำหน้าด้วย @
 *   @2026-08-25         → '2026-08-25 00:00:00'
 *   @2026-08-25 14:30   → '2026-08-25 14:30:00'
 *   @25/08/2026 14:30   → '2026-08-25 14:30:00'
 * @return {string|null}
 */
function grabWhen(text) {
  var t = String(text);
  var m = t.match(/@\s*(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})(?:[\sT]+(\d{1,2})[:\.](\d{2}))?/);
  if (m) {
    return m[1] + '-' + pad2_(m[2]) + '-' + pad2_(m[3]) + ' ' +
           pad2_(m[4] || 0) + ':' + pad2_(m[5] || 0) + ':00';
  }
  var d = t.match(/@\s*(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})(?:[\sT]+(\d{1,2})[:\.](\d{2}))?/);
  if (d) {
    return d[3] + '-' + pad2_(d[2]) + '-' + pad2_(d[1]) + ' ' +
           pad2_(d[4] || 0) + ':' + pad2_(d[5] || 0) + ':00';
  }
  return null;
}

/** ตัดโทเคน @วันเวลา ทิ้ง เพื่อไม่ให้ตัวเลขในวันที่ถูกอ่านเป็นราคา */
function stripWhen(text) {
  return String(text).replace(WHEN_STRIP_RE, ' ');
}

// สกุลเงินอ้างอิงที่พบบ่อย — ใช้ยืนยันว่าเป็นคู่เทรดจริง
var QUOTE_CCY = ['USDT','USDC','USD','THB','JPY','EUR','GBP','AUD','CAD','CHF','NZD','BTC','ETH','BUSD'];

/** เดาชื่อสัญลักษณ์จากข้อความอิสระ — เรียงตามความมั่นใจ มาก→น้อย */
function guessSymbol(text) {
  var t = String(text);

  // 1) รูปแบบมี prefix ตลาด เช่น OANDA:XAUUSD, BINANCE:BTCUSDT — มั่นใจสูงสุด
  var m = t.match(/\b[A-Z]{2,12}:([A-Z0-9]{3,12})\b/);
  if (m) return m[1];

  // 2) โทเคนที่ลงท้ายด้วยสกุลเงินอ้างอิง เช่น XAUUSD, BTCUSDT, EURJPY
  var tokens = t.toUpperCase().match(/\b[A-Z][A-Z0-9]{2,11}\b/g) || [];
  for (var i = 0; i < tokens.length; i++) {
    var w = tokens[i];
    if (SYMBOL_STOPWORDS.indexOf(w) >= 0) continue;
    for (var q = 0; q < QUOTE_CCY.length; q++) {
      var suf = QUOTE_CCY[q];
      if (w.length > suf.length && w.slice(-suf.length) === suf) return w;
    }
  }

  // 3) โทเคน 4-8 ตัวอักษรที่ไม่ใช่คำต้องห้าม (เช่น SET, XAUUSD ที่เขียนแปลก)
  for (var j = 0; j < tokens.length; j++) {
    var v = tokens[j];
    if (SYMBOL_STOPWORDS.indexOf(v) >= 0) continue;
    if (v.length >= 4 && v.length <= 8 && /^[A-Z][A-Z0-9]*$/.test(v)) return v;
  }

  // 4) ทางเลือกสุดท้าย: โทเคน 3 ตัวอักษร เช่น SET, DJI
  for (var k = 0; k < tokens.length; k++) {
    if (tokens[k].length === 3 && SYMBOL_STOPWORDS.indexOf(tokens[k]) < 0) return tokens[k];
  }
  return null;
}

/**
 * parseSignal — รับข้อความดิบ คืน object สัญญาณ
 * รองรับ 3 รูปแบบ: JSON / key=value / ข้อความอิสระ
 * @return {{ok:boolean, error:string, symbol:string, tf:string, action:string,
 *           price:number, tp:number, sl:number, secret:string, bartime:string, note:string}}
 */
function parseSignal(raw) {
  var out = { ok: false, error: '', symbol: '', tf: '', action: '',
              price: null, tp: null, sl: null, secret: '', bartime: '', note: '',
              mode: '', when: '', lots: null };
  var text = String(raw || '').trim();
  if (!text) { out.error = 'ข้อความว่าง'; return out; }

  // ---------- 1) ลอง JSON ก่อน ----------
  var jsonPart = text;
  var b1 = text.indexOf('{'), b2 = text.lastIndexOf('}');
  if (b1 >= 0 && b2 > b1) jsonPart = text.substring(b1, b2 + 1);
  var obj = null;
  try { obj = JSON.parse(jsonPart); } catch (e) { obj = null; }

  if (obj && typeof obj === 'object') {
    out.symbol  = String(obj.symbol || obj.ticker || obj.pair || '').toUpperCase();
    out.tf      = String(obj.tf || obj.timeframe || obj.interval || '');
    out.action  = String(obj.action || obj.side || obj.signal || obj.order || '').toUpperCase();
    out.price   = toNum(obj.price !== undefined ? obj.price : obj.close);
    out.tp      = toNum(obj.tp !== undefined ? obj.tp : obj.takeprofit);
    out.sl      = toNum(obj.sl !== undefined ? obj.sl : obj.stoploss);
    out.secret  = String(obj.secret || obj.key || '');
    out.bartime = String(obj.time || obj.bartime || '');
    out.note    = String(obj.note || obj.comment || '');
    out.mode    = obj.mode ? normMode(obj.mode) : '';
    out.when    = String(obj.when || obj.entry_time || '');
    out.lots    = toNum(obj.lots !== undefined ? obj.lots : obj.size);
    if (out.action) out.action = detectAction(out.action) || out.action;
  }

  // ---------- 2) เติมช่องที่ยังว่างด้วย regex ----------
  // ตัด @วันเวลา ออกก่อน ไม่งั้นเลขในวันที่จะถูกอ่านเป็นราคา
  if (!out.when) out.when = grabWhen(text) || '';
  var clean = stripWhen(text);

  if (!out.action) out.action = detectAction(clean) || '';
  if (!out.symbol) out.symbol = String(guessSymbol(clean) || '').toUpperCase();
  // ตัด prefix ตลาด เช่น OANDA:XAUUSD → XAUUSD
  if (out.symbol.indexOf(':') >= 0) out.symbol = out.symbol.split(':').pop();
  if (!out.mode) out.mode = detectMode(clean) || '';

  if (!out.tf) {
    var tfv = grabStr(clean, ['tf','timeframe','interval','tframe']);
    if (!tfv) {
      var mt = clean.match(/\b(\d+\s?(?:m|min|h|hr|D|W|M)\b)/i);
      if (mt) tfv = mt[1].replace(/\s/g, '');
    }
    out.tf = String(tfv || '');
  }
  if (out.price === null) out.price = grab(clean, ['price','entry','close','ราคา','เข้า']);
  if (out.tp === null)    out.tp    = grab(clean, ['tp','take[\\s_]?profit','target','เป้า']);
  if (out.sl === null)    out.sl    = grab(clean, ['sl','stop[\\s_]?loss','stop','ตัดขาดทุน']);
  if (!out.secret)        out.secret = grabStr(clean, ['secret','key','token']) || '';

  // ---------- 3) ตรวจความครบถ้วน ----------
  if (!out.action) { out.error = 'ไม่พบทิศทาง (BUY/SELL/CLOSE) ในข้อความ'; return out; }
  if (!out.symbol) { out.error = 'ไม่พบชื่อสัญลักษณ์ — ต้องใส่ {{ticker}} ในข้อความ Alert'; return out; }
  if (out.action !== 'CLOSE' && out.price === null) {
    out.error = 'ไม่พบราคาเข้า — ต้องใส่ {{close}} ในข้อความ Alert'; return out;
  }
  if (!out.tf) out.tf = '?';
  out.ok = true;
  return out;
}

/** สร้าง id กันสัญญาณซ้ำ (webhook + email อาจมาพร้อมกัน) */
function makeSignalId(sig, bucketMs) {
  // ไม้ย้อนหลังใช้เวลาที่ระบุเป็นตัวแยก จะได้บันทึกหลายไม้ในนาทีเดียวกันได้
  var t = sig.when ? String(sig.when)
        : sig.bartime ? String(sig.bartime)
        : String(Math.floor(Date.now() / (bucketMs || 60000)));
  return [sig.symbol, sig.tf, sig.action, sig.price, sig.mode || '', t].join('|');
}

/*******************************************************
 * 03_Ingest.gs — รับสัญญาณ 2 ทาง
 *   Path A: doPost()  ← TradingView Webhook (Essential ขึ้นไป)
 *   Path B: pollGmail() ← อีเมล Alert (ใช้ได้ทุกแพ็กเกจ)
 *******************************************************/

/**
 * doPost — ใช้ 2 อย่าง
 *   1) Webhook จาก TradingView (ส่ง JSON สัญญาณมาตรง ๆ)
 *   2) อัปโหลดรูปจากแอป (ต้องมี act:'upload' + key ที่ถูกต้อง)
 * ส่งแบบ Content-Type: text/plain จะไม่โดน CORS preflight ของ Apps Script
 */
function doPost(e) {
  var body = '';
  try { body = (e && e.postData && e.postData.contents) ? e.postData.contents : ''; } catch (err) {}

  var obj = null;
  try { obj = JSON.parse(body); } catch (err) { obj = null; }

  if (obj && obj.act) {
    var key = String(cfg('API_KEY', ''));
    if (key && String(obj.key || '') !== key) {
      return jsonOut_({ ok: false, error: 'unauthorized' }, null);
    }
    if (obj.act === 'upload') return jsonOut_(apiAttachImage_(obj), null);
    return jsonOut_({ ok: false, error: 'ไม่รู้จักคำสั่ง ' + obj.act }, null);
  }

  return jsonOut_(ingest(body, 'webhook'), null);
}

// หมายเหตุ: doGet อยู่ในไฟล์ 09_WebApi.gs (ทำหน้าที่เป็น API ให้หน้า Dashboard ด้วย)

/** อ่านอีเมล Alert จาก TradingView (รันทุก 1 นาทีโดย trigger) */
function pollGmail() {
  var from = cfg('EMAIL_FROM', 'noreply@tradingview.com');
  var mins = cfgNum('EMAIL_LOOKBACK_MIN', 15);
  var afterSec = Math.floor((Date.now() - mins * 60000) / 1000);
  var q = 'from:(' + from + ') is:unread after:' + afterSec;

  var threads = GmailApp.search(q, 0, 25);
  for (var i = 0; i < threads.length; i++) {
    var msgs = threads[i].getMessages();
    for (var j = 0; j < msgs.length; j++) {
      var m = msgs[j];
      if (!m.isUnread()) continue;
      var text = (m.getSubject() || '') + '\n' + (m.getPlainBody() || '');
      ingest(text, 'email');
      m.markRead();
    }
  }
}

/**
 * ingest — ทางเข้าหลักของสัญญาณทุกช่องทาง
 * @param {string} raw ข้อความดิบ
 * @param {string} source 'webhook' | 'email' | 'manual'
 */
function ingest(raw, source) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) {
    return { ok: false, error: 'ระบบกำลังประมวลผลสัญญาณอื่น' };
  }

  try {
    var sig = parseSignal(raw);

    // ตรวจรหัสลับเฉพาะ webhook (email มาจาก Gmail ของเราเองอยู่แล้ว)
    var secret = String(cfg('WEBHOOK_SECRET', ''));
    if (source === 'webhook' && secret && sig.secret !== secret) {
      logRaw_(sig, source, 'rejected', raw, 'secret ไม่ถูกต้อง');
      return { ok: false, error: 'unauthorized' };
    }

    if (!sig.ok) {
      logRaw_(sig, source, 'error', raw, sig.error);
      return { ok: false, error: sig.error };
    }

    var id = makeSignalId(sig);
    if (isDuplicate_(id)) {
      logRaw_(sig, source, 'duplicate', raw, '');
      return { ok: true, status: 'duplicate' };
    }

    logRaw_(sig, source, 'accepted', raw, '', id);
    var result = handleSignal(sig, id);   // → 04_Engine.gs
    return { ok: true, status: 'accepted', result: result };

  } catch (err) {
    logRaw_({}, source, 'error', raw, String(err));
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

function logRaw_(sig, source, status, raw, note, id) {
  var sh = ss().getSheetByName(SHEETS.RAW);
  if (!sh) return;
  sh.appendRow([
    id || '', nowStr(), source, sig.symbol || '', sig.tf || '', sig.action || '',
    sig.price, sig.tp, sig.sl, note || sig.note || '', status,
    String(raw || '').substring(0, 2000)
  ]);
}

/** เช็คซ้ำจาก 300 แถวล่าสุด */
function isDuplicate_(id) {
  var sh = ss().getSheetByName(SHEETS.RAW);
  var last = sh.getLastRow();
  if (last < 2) return false;
  var start = Math.max(2, last - 299);
  var ids = sh.getRange(start, 1, last - start + 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (ids[i][0] === id) return true;
  return false;
}

/** ทดสอบยิงสัญญาณเองโดยไม่ต้องรอ TradingView */
function testIngest() {
  var demo = JSON.stringify({
    secret: cfg('WEBHOOK_SECRET',''), symbol: 'XAUUSD', tf: '30',
    action: 'BUY', price: 4440.30, tp: 4460.00, sl: 4430.00
  });
  Logger.log(ingest(demo, 'manual'));
}

/*******************************************************
 * 04_Engine.gs — Paper Trade Engine
 *   เปิดไม้จากสัญญาณ / ปิดไม้เมื่อชน TP-SL หรือมีสัญญาณตรงข้าม
 *******************************************************/

/* ---------- ฟังก์ชันคำนวณ (pure — ทดสอบแยกได้) ---------- */

/** คำนวณผลกำไรของไม้ที่ปิดแล้ว */
function computePnl(side, entry, exit, lots, valuePerPoint, sl) {
  var dir = (side === 'BUY') ? 1 : -1;
  var pts = (exit - entry) * dir;
  var pct = entry ? (pts / entry) * 100 : 0;
  var usd = pts * (valuePerPoint || 1) * (lots || 1);
  var r = null;
  if (sl !== null && sl !== undefined && sl !== '' && !isNaN(sl)) {
    var risk = Math.abs(entry - sl);
    if (risk > 0) r = pts / risk;
  }
  return { pts: pts, pct: pct, usd: usd, r: r };
}

/**
 * ตรวจว่าราคาในช่วงที่ผ่านมาชน TP หรือ SL
 * กติกาอนุรักษ์นิยม: ถ้าแท่งเดียวกันชนทั้งคู่ ให้ถือว่า SL ชนก่อน (worst case)
 */
function checkHit(side, tp, sl, high, low) {
  var hasTp = (tp !== null && tp !== undefined && tp !== '' && !isNaN(tp) && Number(tp) > 0);
  var hasSl = (sl !== null && sl !== undefined && sl !== '' && !isNaN(sl) && Number(sl) > 0);
  var hitTp = false, hitSl = false;

  if (side === 'BUY') {
    if (hasTp && high >= Number(tp)) hitTp = true;
    if (hasSl && low  <= Number(sl)) hitSl = true;
  } else {
    if (hasTp && low  <= Number(tp)) hitTp = true;
    if (hasSl && high >= Number(sl)) hitSl = true;
  }

  if (hitSl) return { hit: true, price: Number(sl), reason: 'SL' };
  if (hitTp) return { hit: true, price: Number(tp), reason: 'TP' };
  return { hit: false, price: null, reason: '' };
}

/* ---------- ตรรกะหลัก ---------- */

function handleSignal(sig, signalId) {
  // LIVE กับ TEST แยกกระเป๋าเด็ดขาด — สัญญาณโหมดหนึ่งจะไม่ไปแตะไม้ของอีกโหมด
  sig.mode = normMode(sig.mode || defaultMode());
  var open = findOpenTrades_(sig.symbol, sig.tf, sig.mode);

  if (sig.action === 'CLOSE') {
    var n = 0;
    open.forEach(function (t) { closeTrade_(t, sig.price, 'SIGNAL', signalId, sig.when); n++; });
    return 'ปิด ' + n + ' ไม้ตามสัญญาณ';
  }

  // สัญญาณตรงข้าม → ปิดไม้เดิมก่อน
  if (cfgBool('CLOSE_ON_OPPOSITE', true)) {
    open.filter(function (t) { return t.side !== sig.action; })
        .forEach(function (t) { closeTrade_(t, sig.price, 'REVERSE', signalId, sig.when); });
  }

  // ห้ามเปิดซ้อนทิศเดียวกัน ถ้าไม่เปิด pyramid
  if (!cfgBool('ALLOW_PYRAMID', false)) {
    var same = findOpenTrades_(sig.symbol, sig.tf, sig.mode)
                 .filter(function (t) { return t.side === sig.action; });
    if (same.length > 0) return 'มีไม้ทิศเดียวกันเปิดอยู่แล้ว — ข้าม';
  }

  return openTrade_(sig, signalId);
}

function openTrade_(sig, signalId) {
  var sh = ss().getSheetByName(SHEETS.TRADES);
  var id = 'T' + Utilities.formatDate(new Date(), tz(), 'yyMMddHHmmss') +
           '-' + Math.floor(Math.random() * 900 + 100);
  var lots = (sig.lots !== null && sig.lots !== undefined && sig.lots > 0)
             ? sig.lots : cfgNum('DEFAULT_LOTS', 1);
  var mode = normMode(sig.mode || defaultMode());
  var entryTime = sig.when || nowStr();     // ไม้ย้อนหลังใช้เวลาที่ผู้ใช้ระบุ

  sh.appendRow([
    id, sig.symbol, sig.tf, sig.action, lots,
    entryTime, sig.price, sig.tp, sig.sl,
    '', '', '', '', '', '', '', '', 'OPEN', signalId, '',
    mode, '', sig.note || ''
  ]);

  CACHE.lastTradeId = id;          // ให้ฝั่งแอปเอา id ไปแนบรูปต่อได้ทันที
  if (cfgBool('NOTIFY_OPEN', true)) notifyOpen(sig, id, lots, mode, entryTime);
  return 'เปิดไม้ ' + id + (mode === MODE_TEST ? ' [TEST]' : '');
}

function closeTrade_(trade, exitPrice, reason, signalId, whenStr) {
  if (exitPrice === null || exitPrice === undefined || isNaN(exitPrice)) {
    exitPrice = getLastPrice_(trade.symbol);
  }
  if (exitPrice === null) return;   // ดึงราคาไม่ได้ → ยังไม่ปิด

  var map = symbolMap_(trade.symbol);
  var p = computePnl(trade.side, trade.entry, exitPrice, trade.lots, map.value_per_point, trade.sl);
  var exitTime = whenStr || nowStr();
  var exitDate = parseDate_(exitTime);
  var mins = Math.round((exitDate - parseDate_(trade.entry_time)) / 60000);
  if (mins < 0) mins = 0;

  var sh = ss().getSheetByName(SHEETS.TRADES);
  sh.getRange(trade.row, 10, 1, 11).setValues([[
    exitTime, exitPrice, reason,
    round_(p.pts, 5), round_(p.pct, 4), round_(p.usd, 2),
    (p.r === null ? '' : round_(p.r, 3)), mins, 'CLOSED', trade.signal_open, signalId || ''
  ]]);

  updateDaily_(exitDate, p.usd, trade.mode);
  if (cfgBool('NOTIFY_CLOSE', true)) notifyClose(trade, exitPrice, reason, p);
}

/**
 * trigger ทุก 5 นาที — ตรวจว่าไม้ที่เปิดอยู่ชน TP/SL หรือยัง
 * ตรวจเฉพาะไม้ LIVE เท่านั้น เพราะไม้ TEST เป็นไม้ย้อนหลัง
 * ราคาปัจจุบันไม่เกี่ยวกัน ต้องปิดเองด้วย /close หรือ CLOSE @วันที่
 */
function monitorOpenTrades() {
  var trades = findOpenTrades_(null, null, MODE_LIVE);
  if (!trades.length) return;

  // จัดกลุ่มตามสัญลักษณ์ เพื่อยิง API ครั้งเดียวต่อคู่
  var bySymbol = {};
  trades.forEach(function (t) {
    if (!bySymbol[t.symbol]) bySymbol[t.symbol] = [];
    bySymbol[t.symbol].push(t);
  });

  Object.keys(bySymbol).forEach(function (sym) {
    var list = bySymbol[sym];
    var oldest = Math.min.apply(null, list.map(function (t) {
      return parseDate_(t.entry_time).getTime();
    }));
    var rng = getRange_(sym, oldest);
    if (!rng) return;

    list.forEach(function (t) {
      var hit = checkHit(t.side, t.tp, t.sl, rng.high, rng.low);
      if (hit.hit) closeTrade_(t, hit.price, hit.reason, '');
    });
  });
}

/* ---------- ตัวช่วยอ่านข้อมูล ---------- */

/**
 * หาไม้ที่ยังเปิดอยู่
 * @param {string} mode 'LIVE' | 'TEST' | null (null = ทุกโหมด)
 */
function findOpenTrades_(symbol, tf, mode) {
  var sh = ss().getSheetByName(SHEETS.TRADES);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, HEADERS.TRADES.length).getValues();
  var want = mode ? normMode(mode) : null;
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    if (String(v[17]) !== 'OPEN') continue;
    if (symbol && String(v[1]) !== symbol) continue;
    if (tf && String(v[2]) !== String(tf)) continue;
    var m = normMode(v[20]);                       // ไม้เก่าที่ยังไม่มีค่า = LIVE
    if (want && m !== want) continue;
    out.push({
      row: i + 2, id: v[0], symbol: String(v[1]), tf: String(v[2]), side: String(v[3]),
      lots: Number(v[4]) || 1, entry_time: v[5], entry: Number(v[6]),
      tp: v[7], sl: v[8], signal_open: v[18],
      mode: m, chart: String(v[21] || ''), note: String(v[22] || '')
    });
  }
  return out;
}

function symbolMap_(symbol) {
  if (!CACHE.map) {
    CACHE.map = {};
    var sh = ss().getSheetByName(SHEETS.MAP);
    var last = sh.getLastRow();
    if (last >= 2) {
      sh.getRange(2, 1, last - 1, 6).getValues().forEach(function (r) {
        if (r[0]) CACHE.map[String(r[0]).toUpperCase()] = {
          provider: String(r[1] || 'yahoo').toLowerCase(),
          api_symbol: String(r[2] || ''),
          value_per_point: Number(r[3]) || 1,
          digits: Number(r[4]) || 2
        };
      });
    }
  }
  return CACHE.map[String(symbol).toUpperCase()] ||
         CACHE.map['*'] ||
         { provider: 'yahoo', api_symbol: '', value_per_point: 1, digits: 2 };
}

/* ---------- ราคาตลาด ---------- */

function getLastPrice_(symbol) {
  var r = getRange_(symbol, Date.now() - 10 * 60000);
  return r ? r.last : null;
}

/** คืนช่วงราคาสูงสุด/ต่ำสุดตั้งแต่ sinceMs จนถึงตอนนี้ */
function getRange_(symbol, sinceMs) {
  var m = symbolMap_(symbol);
  var api = m.api_symbol || symbol;
  try {
    if (m.provider === 'binance')  return binanceRange_(api, sinceMs);
    if (m.provider === 'yahoo')    return yahooRange_(api, sinceMs);
  } catch (e) {
    Logger.log('ดึงราคา ' + symbol + ' ไม่สำเร็จ: ' + e);
  }
  return null;   // provider = manual → ต้องส่งสัญญาณ CLOSE เอง
}

function binanceRange_(sym, sinceMs) {
  var url = 'https://api.binance.com/api/v3/klines?symbol=' + encodeURIComponent(sym) +
            '&interval=5m&limit=500&startTime=' + Math.floor(sinceMs);
  var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) return null;
  var k = JSON.parse(r.getContentText());
  if (!k.length) return null;
  var hi = -Infinity, lo = Infinity;
  k.forEach(function (c) { hi = Math.max(hi, +c[2]); lo = Math.min(lo, +c[3]); });
  return { high: hi, low: lo, last: +k[k.length - 1][4] };
}

function yahooRange_(sym, sinceMs) {
  var url = 'https://query1.finance.yahoo.com/v8/finance/chart/' +
            encodeURIComponent(sym) + '?interval=5m&range=5d';
  var r = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: { 'User-Agent': 'Mozilla/5.0' }
  });
  if (r.getResponseCode() !== 200) return null;
  var j = JSON.parse(r.getContentText());
  var res = j && j.chart && j.chart.result && j.chart.result[0];
  if (!res) return null;
  var ts = res.timestamp || [];
  var q = res.indicators.quote[0];
  var hi = -Infinity, lo = Infinity, last = null;
  for (var i = 0; i < ts.length; i++) {
    if (ts[i] * 1000 < sinceMs) continue;
    if (q.high[i] != null) hi = Math.max(hi, q.high[i]);
    if (q.low[i]  != null) lo = Math.min(lo, q.low[i]);
    if (q.close[i] != null) last = q.close[i];
  }
  if (last === null) {
    last = res.meta && res.meta.regularMarketPrice;
    if (last == null) return null;
    hi = lo = last;
  }
  return { high: hi, low: lo, last: last };
}

/* ---------- utility ---------- */

function round_(n, d) {
  if (n === null || n === undefined || isNaN(n)) return '';
  var f = Math.pow(10, d);
  return Math.round(n * f) / f;
}

function parseDate_(v) {
  if (v instanceof Date) return v;
  var s = String(v).replace(' ', 'T');
  var d = new Date(s);
  return isNaN(d.getTime()) ? new Date() : d;
}

/*******************************************************
 * 05_Notify.gs — แจ้งเตือน Telegram
 *******************************************************/

function tgSend(text) {
  var token = String(cfg('TELEGRAM_TOKEN', ''));
  var chat  = String(cfg('TELEGRAM_CHAT_ID', ''));
  if (!token || !chat) { Logger.log('ยังไม่ได้ตั้ง TELEGRAM_TOKEN / CHAT_ID'); return false; }

  var url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    muteHttpExceptions: true,
    payload: {
      chat_id: chat,
      text: text,
      parse_mode: 'HTML',
      disable_web_page_preview: 'true'
    }
  });
  if (res.getResponseCode() !== 200) {
    Logger.log('Telegram error: ' + res.getContentText());
    return false;
  }
  return true;
}

function notifyOpen(sig, tradeId, lots, mode, entryTime) {
  var d = symbolMap_(sig.symbol).digits;
  var arrow = (sig.action === 'BUY') ? '🟢 <b>BUY</b>' : '🔴 <b>SELL</b>';
  var risk = (sig.sl ? Math.abs(sig.price - sig.sl) : null);
  var rr   = (sig.tp && sig.sl && risk) ? Math.abs(sig.tp - sig.price) / risk : null;
  var isTest = normMode(mode) === MODE_TEST;

  var lines = [
    isTest ? '🧪 <b>บันทึกไม้ TEST</b> (ไม่นับสถิติจริง)' : '📈 <b>เปิดไม้ใหม่</b>',
    arrow + '  ' + sig.symbol + '  <code>' + sig.tf + '</code>',
    '───────────────',
    'เข้าที่ : <b>' + fmt(sig.price, d) + '</b>',
    (sig.tp ? 'TP    : ' + fmt(sig.tp, d) : null),
    (sig.sl ? 'SL    : ' + fmt(sig.sl, d) : null),
    (rr ? 'R:R   : 1 : ' + fmt(rr, 2) : null),
    'ขนาด  : ' + fmt(lots, 2) + ' lot',
    '───────────────',
    '🕐 ' + (entryTime || nowStr()) + '   <code>' + tradeId + '</code>',
    '🖼 ส่งรูปชาร์ตตามมาได้เลย จะแนบให้ไม้นี้'
  ].filter(function (x) { return x; });

  tgSend(lines.join('\n'));
}

function notifyClose(trade, exitPrice, reason, p) {
  var d = symbolMap_(trade.symbol).digits;
  var win = p.usd >= 0;
  var tag = (normMode(trade.mode) === MODE_TEST) ? ' 🧪 TEST' : '';
  var head = (win ? '✅ <b>ปิดไม้ — กำไร</b>' : '❌ <b>ปิดไม้ — ขาดทุน</b>') + tag;
  var reasonTh = { TP: 'ชน Take Profit', SL: 'ชน Stop Loss', MANUAL: 'ปิดเอง',
                   REVERSE: 'สัญญาณกลับทิศ', SIGNAL: 'สัญญาณสั่งปิด' }[reason] || reason;

  var lines = [
    head,
    (trade.side === 'BUY' ? '🟢 BUY' : '🔴 SELL') + '  ' + trade.symbol + '  <code>' + trade.tf + '</code>',
    '───────────────',
    'เข้า  : ' + fmt(trade.entry, d),
    'ออก  : ' + fmt(exitPrice, d) + '  (' + reasonTh + ')',
    'กำไร : <b>' + fmtSigned(p.usd, 2) + ' USD</b>  (' + fmtSigned(p.pct, 2) + '%)',
    (p.r !== null ? 'R     : ' + fmtSigned(p.r, 2) + ' R' : null),
    '───────────────',
    '🕐 ' + nowStr()
  ].filter(function (x) { return x; });

  tgSend(lines.join('\n'));
}

/* ---------- ตัวช่วยตอนติดตั้ง ---------- */

/** ทักบอทใน Telegram 1 ข้อความก่อน แล้วรันฟังก์ชันนี้เพื่อหา Chat ID */
function findMyChatId() {
  var token = String(cfg('TELEGRAM_TOKEN', ''));
  if (!token) { SpreadsheetApp.getUi().alert('กรอก TELEGRAM_TOKEN ในชีต Config ก่อนครับ'); return; }

  var r = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getUpdates',
                            { muteHttpExceptions: true });
  var j = JSON.parse(r.getContentText());
  if (!j.ok || !j.result || !j.result.length) {
    SpreadsheetApp.getUi().alert(
      'ยังไม่พบข้อความ\n\n1) เปิด Telegram หาบอทของคุณ\n2) กด Start แล้วพิมพ์อะไรก็ได้ 1 ข้อความ\n3) กลับมารันเมนูนี้อีกครั้ง');
    return;
  }
  var ids = {};
  j.result.forEach(function (u) {
    var c = (u.message && u.message.chat) || (u.channel_post && u.channel_post.chat);
    if (c) ids[c.id] = (c.title || c.first_name || c.username || '');
  });
  var msg = Object.keys(ids).map(function (k) { return k + '  →  ' + ids[k]; }).join('\n');
  SpreadsheetApp.getUi().alert('Chat ID ที่เจอ:\n\n' + msg + '\n\nนำไปใส่ช่อง TELEGRAM_CHAT_ID ในชีต Config');
}

function testTelegram() {
  var ok = tgSend('🤖 <b>GTPro Signal DB</b>\nเชื่อมต่อสำเร็จแล้วครับ\n🕐 ' + nowStr());
  SpreadsheetApp.getActive().toast(ok ? 'ส่งสำเร็จ — เช็ค Telegram' : 'ส่งไม่สำเร็จ ดู Log', 'GTPro', 8);
}

/*******************************************************
 * 06_Report.gs — สถิติกำไร / ปฏิทินรายวัน / สรุปส่ง Telegram
 *******************************************************/

/* ---------- pure function: สรุปสถิติจากรายการกำไรที่เรียงตามเวลา ---------- */
function summarize(pnlList) {
  var n = pnlList.length;
  var wins = [], losses = [], sum = 0, peak = 0, cum = 0, maxDD = 0;
  for (var i = 0; i < n; i++) {
    var v = Number(pnlList[i]) || 0;
    sum += v;
    if (v >= 0) wins.push(v); else losses.push(v);
    cum += v;
    if (cum > peak) peak = cum;
    var dd = peak - cum;
    if (dd > maxDD) maxDD = dd;
  }
  var grossWin  = wins.reduce(function (a, b) { return a + b; }, 0);
  var grossLoss = Math.abs(losses.reduce(function (a, b) { return a + b; }, 0));
  return {
    trades: n,
    wins: wins.length,
    losses: losses.length,
    winrate: n ? (wins.length / n) * 100 : 0,
    avgWin: wins.length ? grossWin / wins.length : 0,
    avgLoss: losses.length ? -grossLoss / losses.length : 0,
    profitFactor: grossLoss > 0 ? grossWin / grossLoss : (grossWin > 0 ? Infinity : 0),
    expectancy: n ? sum / n : 0,
    maxDD: maxDD,
    total: sum
  };
}

/* ---------- ปฏิทินกำไรรายวัน ---------- */

function updateDaily_(date, pnl, mode) {
  var sh = ss().getSheetByName(SHEETS.DAILY);
  var key = dateKey(date);
  var m = normMode(mode);
  var last = sh.getLastRow();
  var row = -1;
  if (last >= 2) {
    var keys = sh.getRange(2, 1, last - 1, 2).getDisplayValues();
    for (var i = 0; i < keys.length; i++) {
      if (keys[i][0] === key && normMode(keys[i][1]) === m) { row = i + 2; break; }
    }
  }
  if (row < 0) { sh.appendRow([key, m, 0, 0, 0, 0, 0, 0, 0, 0]); row = sh.getLastRow(); }

  var v = sh.getRange(row, 1, 1, 10).getValues()[0];
  var trades = (Number(v[2]) || 0) + 1;
  var wins   = (Number(v[3]) || 0) + (pnl >= 0 ? 1 : 0);
  var losses = (Number(v[4]) || 0) + (pnl <  0 ? 1 : 0);
  var total  = (Number(v[6]) || 0) + pnl;
  var best   = Math.max(Number(v[8]) || -Infinity, pnl);
  var worst  = Math.min(Number(v[9]) ||  Infinity, pnl);

  sh.getRange(row, 1, 1, 10).setValues([[
    key, m, trades, wins, losses, round_(trades ? wins / trades * 100 : 0, 2),
    round_(total, 2), '', round_(best, 2), round_(worst, 2)
  ]]);
  recalcCumulative_();
}

/** ยอดสะสมคิดแยกกระเป๋าตามโหมด — เส้น equity ของ LIVE จะไม่ปนกับ TEST */
function recalcCumulative_() {
  var sh = ss().getSheetByName(SHEETS.DAILY);
  var last = sh.getLastRow();
  if (last < 2) return;
  var rng = sh.getRange(2, 1, last - 1, 8);
  var v = rng.getValues();
  v.sort(function (a, b) {
    var ma = normMode(a[1]), mb = normMode(b[1]);
    if (ma !== mb) return ma < mb ? -1 : 1;
    return String(a[0]) < String(b[0]) ? -1 : 1;
  });
  var cum = {};
  for (var i = 0; i < v.length; i++) {
    var m = normMode(v[i][1]);
    v[i][1] = m;
    cum[m] = (cum[m] || 0) + (Number(v[i][6]) || 0);
    v[i][7] = round_(cum[m], 2);
  }
  rng.setValues(v);
}

/* ---------- คำนวณสถิติใหม่ทั้งหมดจาก Trades ---------- */

function rebuildAll() {
  var sh = ss().getSheetByName(SHEETS.TRADES);
  var last = sh.getLastRow();
  var closed = [];
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, HEADERS.TRADES.length).getValues().forEach(function (v) {
      if (String(v[17]) !== 'CLOSED') return;
      closed.push({ symbol: String(v[1]), tf: String(v[2]), side: String(v[3]),
                    exit_time: v[9], pnl: Number(v[14]) || 0, mode: normMode(v[20]) });
    });
  }
  closed.sort(function (a, b) { return parseDate_(a.exit_time) - parseDate_(b.exit_time); });

  // --- Daily_PL (แยกกระเป๋า LIVE / TEST) ---
  var byDay = {};                       // key = 'MODE|yyyy-MM-dd'
  closed.forEach(function (t) {
    var k = t.mode + '|' + dateKey(parseDate_(t.exit_time));
    if (!byDay[k]) byDay[k] = [];
    byDay[k].push(t.pnl);
  });
  var dsh = ss().getSheetByName(SHEETS.DAILY);
  if (dsh.getLastRow() > 1) dsh.getRange(2, 1, dsh.getLastRow() - 1, 10).clearContent();
  var keys = Object.keys(byDay).sort(), cum = {}, rows = [];
  keys.forEach(function (k) {
    var m = k.split('|')[0], day = k.split('|')[1];
    var s = summarize(byDay[k]);
    cum[m] = (cum[m] || 0) + s.total;
    rows.push([day, m, s.trades, s.wins, s.losses, round_(s.winrate, 2),
               round_(s.total, 2), round_(cum[m], 2),
               round_(Math.max.apply(null, byDay[k]), 2),
               round_(Math.min.apply(null, byDay[k]), 2)]);
  });
  if (rows.length) dsh.getRange(2, 1, rows.length, 10).setValues(rows);

  // --- Stats (ภาพรวม + รายคู่ + รายคู่-TF + ราย side) ---
  // ไม้ TEST มี prefix 'TEST · ' เสมอ จะได้ไม่ปนกับสถิติจริง
  var groups = {};
  closed.forEach(function (t) {
    var pre = (t.mode === MODE_TEST) ? 'TEST · ' : '';
    push_(groups, pre + 'ALL', t.pnl);
    push_(groups, pre + 'SYMBOL: ' + t.symbol, t.pnl);
    push_(groups, pre + 'TF: ' + t.tf, t.pnl);
    push_(groups, pre + t.symbol + ' @ ' + t.tf, t.pnl);
    push_(groups, pre + 'SIDE: ' + t.side, t.pnl);
  });
  if (!groups['ALL']) groups['ALL'] = [];

  var ssh = ss().getSheetByName(SHEETS.STATS);
  if (ssh.getLastRow() > 1) ssh.getRange(2, 1, ssh.getLastRow() - 1, 12).clearContent();
  var out = Object.keys(groups).map(function (k) {
    var s = summarize(groups[k]);
    return [k, s.trades, s.wins, s.losses, round_(s.winrate, 2),
            round_(s.avgWin, 2), round_(s.avgLoss, 2),
            (s.profitFactor === Infinity ? '∞' : round_(s.profitFactor, 2)),
            round_(s.expectancy, 2), round_(s.maxDD, 2), round_(s.total, 2), nowStr()];
  });
  out.sort(function (a, b) { return b[1] - a[1]; });
  if (out.length) ssh.getRange(2, 1, out.length, 12).setValues(out);

  try { buildCalendar(); } catch (e) { Logger.log('สร้างปฏิทินไม่สำเร็จ: ' + e); }
  SpreadsheetApp.getActive().toast('คำนวณสถิติใหม่แล้ว (' + fmt(closed.length, 0) + ' ไม้)', 'GTPro', 6);
}

function push_(obj, key, v) { if (!obj[key]) obj[key] = []; obj[key].push(v); }

/* ---------- สรุปรายวันส่ง Telegram ---------- */

function sendDailySummary() {
  var sh = ss().getSheetByName(SHEETS.TRADES);
  var last = sh.getLastRow();
  if (last < 2) return;
  var today = dateKey(new Date());
  var todays = [], openCount = 0, testToday = 0;

  sh.getRange(2, 1, last - 1, HEADERS.TRADES.length).getValues().forEach(function (v) {
    var m = normMode(v[20]);
    if (String(v[17]) === 'OPEN') { if (m === MODE_LIVE) openCount++; return; }
    if (String(v[17]) !== 'CLOSED') return;
    if (dateKey(parseDate_(v[9])) !== today) return;
    if (m === MODE_TEST) { testToday++; return; }      // สรุปรายวันนับเฉพาะไม้จริง
    todays.push({ symbol: String(v[1]), side: String(v[3]), pnl: Number(v[14]) || 0 });
  });

  var s = summarize(todays.map(function (t) { return t.pnl; }));
  var icon = s.total >= 0 ? '🟢' : '🔴';

  var lines = [
    '📊 <b>สรุปผลประจำวัน</b>  ' + today,
    '───────────────',
    icon + ' กำไรวันนี้ : <b>' + fmtSigned(s.total, 2) + ' USD</b>',
    'จำนวนไม้  : ' + fmt(s.trades, 0) + ' ไม้',
    'ชนะ/แพ้   : ' + fmt(s.wins, 0) + ' / ' + fmt(s.losses, 0) +
      '  (Win rate ' + fmt(s.winrate, 1) + '%)',
    'Profit Factor : ' + (s.profitFactor === Infinity ? '∞' : fmt(s.profitFactor, 2)),
    'ไม้ที่ยังเปิดอยู่ : ' + fmt(openCount, 0) + ' ไม้'
  ];
  if (testToday) lines.push('🧪 ไม้ TEST วันนี้ : ' + fmt(testToday, 0) + ' ไม้ (ไม่นับรวมด้านบน)');

  if (todays.length) {
    var byS = {};
    todays.forEach(function (t) { byS[t.symbol] = (byS[t.symbol] || 0) + t.pnl; });
    lines.push('───────────────', '<b>แยกตามคู่</b>');
    Object.keys(byS).sort(function (a, b) { return byS[b] - byS[a]; })
      .forEach(function (k) {
        lines.push((byS[k] >= 0 ? '  🟩 ' : '  🟥 ') + k + '  ' + fmtSigned(byS[k], 2) + ' USD');
      });
  }

  tgSend(lines.join('\n'));
  rebuildAll();
}

/* ---------- ข้อมูลตัวอย่างไว้ทดสอบระบบ ---------- */

function seedDemoData() {
  var demo = [
    ['XAUUSD','30','BUY',  4420.50, 4440.00, 4410.00, 4440.00, 'TP'],
    ['XAUUSD','30','SELL', 4445.00, 4425.00, 4455.00, 4455.00, 'SL'],
    ['XAUUSD','30','BUY',  4430.00, 4450.00, 4420.00, 4450.00, 'TP'],
    ['BTCUSDT','60','BUY', 78000.00, 79500.00, 77200.00, 79500.00, 'TP'],
    ['BTCUSDT','60','SELL',79200.00, 77800.00, 79900.00, 77800.00, 'TP']
  ];
  var sh = ss().getSheetByName(SHEETS.TRADES);
  demo.forEach(function (d, i) {
    var map = symbolMap_(d[0]);
    var p = computePnl(d[2], d[3], d[6], 1, map.value_per_point, d[5]);
    var t = new Date(Date.now() - (demo.length - i) * 3600000);
    sh.appendRow([
      'DEMO-' + (i + 1), d[0], d[1], d[2], 1,
      Utilities.formatDate(t, tz(), 'yyyy-MM-dd HH:mm:ss'), d[3], d[4], d[5],
      Utilities.formatDate(new Date(t.getTime() + 1800000), tz(), 'yyyy-MM-dd HH:mm:ss'),
      d[6], d[7], round_(p.pts, 5), round_(p.pct, 4), round_(p.usd, 2),
      round_(p.r, 3), 30, 'CLOSED', 'demo', '',
      MODE_LIVE, '', 'ข้อมูลตัวอย่าง'
    ]);
  });
  rebuildAll();
}

/*******************************************************
 * 07_Calendar.gs — ปฏิทินกำไรรายวัน (แบบ Trading Calendar)
 *   สร้างชีต "Calendar" แสดงกำไรรายวันเป็นตารางปฏิทิน
 *   เปลี่ยนเดือนโดยแก้ค่าในช่อง B1 (เช่น 2026-08) แล้วกดเมนู
 *******************************************************/

var TH_DOW = ['อา','จ','อ','พ','พฤ','ศ','ส'];
var TH_MONTH = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];

function buildCalendar() {
  var s = ss();
  var sh = s.getSheetByName('Calendar');
  if (!sh) sh = s.insertSheet('Calendar');

  // ---- เดือนที่จะแสดง ----
  var monthKey = String(sh.getRange('B1').getDisplayValue() || '').trim();
  if (!/^\d{4}-\d{2}$/.test(monthKey)) monthKey = Utilities.formatDate(new Date(), tz(), 'yyyy-MM');
  var year  = parseInt(monthKey.substring(0, 4), 10);
  var month = parseInt(monthKey.substring(5, 7), 10) - 1;

  // ---- โหมดที่จะแสดง (LIVE / TEST) อยู่ในช่อง D1 ----
  var wantMode = normMode(String(sh.getRange('D1').getDisplayValue() || MODE_LIVE));

  // ---- ดึงข้อมูลรายวัน ----
  var daily = {};
  var dsh = s.getSheetByName(SHEETS.DAILY);
  if (dsh && dsh.getLastRow() > 1) {
    dsh.getRange(2, 1, dsh.getLastRow() - 1, 10).getDisplayValues().forEach(function (r) {
      if (!r[0]) return;
      if (normMode(r[1]) !== wantMode) return;
      daily[r[0]] = { trades: Number(r[2]) || 0,
                      pnl: parseFloat(String(r[6]).replace(/,/g, '')) || 0 };
    });
  }

  sh.clear();
  sh.clearFormats();

  // ---- ส่วนหัว ----
  sh.getRange('A1').setValue('เดือน →').setFontWeight('bold');
  sh.getRange('B1').setValue(monthKey).setFontWeight('bold')
    .setBackground('#111827').setFontColor('#34d399').setHorizontalAlignment('center');
  sh.getRange('C1').setValue('โหมด →').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('D1').setValue(wantMode).setFontWeight('bold')
    .setBackground('#111827').setFontColor(wantMode === MODE_TEST ? '#fbbf24' : '#34d399')
    .setHorizontalAlignment('center');
  sh.getRange('E1').setValue('พิมพ์ปี-เดือน (B1) และ LIVE/TEST (D1) แล้วกดเมนู ⚙️ GTPro › สร้างปฏิทินใหม่')
    .setFontColor('#6b7280').setFontSize(9);

  sh.getRange('A2').setValue('Trading Calendar ' + (wantMode === MODE_TEST ? '[TEST] ' : '') +
                             '— ' + TH_MONTH[month] + ' ' + year)
    .setFontSize(16).setFontWeight('bold');

  // ---- สรุปรายเดือน ----
  var mTotal = 0, mTrades = 0, winDays = 0, lossDays = 0;
  Object.keys(daily).forEach(function (k) {
    if (k.indexOf(monthKey) !== 0) return;
    mTotal += daily[k].pnl;
    mTrades += daily[k].trades;
    if (daily[k].pnl > 0) winDays++; else if (daily[k].pnl < 0) lossDays++;
  });

  var kpi = [
    ['กำไรเดือนนี้ (USD)', 'จำนวน Trades', 'วันกำไร / วันขาดทุน'],
    [fmtSigned(mTotal, 2) + ' USD', fmt(mTrades, 0), fmt(winDays, 0) + ' / ' + fmt(lossDays, 0)]
  ];
  sh.getRange(4, 1, 2, 3).setValues(kpi);
  sh.getRange(4, 1, 1, 3).setFontColor('#9ca3af').setFontSize(10);
  sh.getRange(5, 1, 1, 3).setFontSize(18).setFontWeight('bold')
    .setBackground('#0b1220').setFontColor('#e5e7eb');
  sh.getRange(5, 1).setFontColor(mTotal >= 0 ? '#34d399' : '#f87171');

  // ---- หัวตารางวัน ----
  var top = 7;
  sh.getRange(top, 1, 1, 7).setValues([TH_DOW])
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#111827').setFontColor('#9ca3af');

  // ---- เติมวันในเดือน ----
  var first = new Date(year, month, 1);
  var startCol = first.getDay();              // 0 = อาทิตย์
  var daysInMonth = new Date(year, month + 1, 0).getDate();

  var grid = [], colors = [], fonts = [];
  for (var w = 0; w < 6; w++) {
    grid.push(['', '', '', '', '', '', '']);
    colors.push(['#0b1220','#0b1220','#0b1220','#0b1220','#0b1220','#0b1220','#0b1220']);
    fonts.push(['#374151','#374151','#374151','#374151','#374151','#374151','#374151']);
  }

  for (var d = 1; d <= daysInMonth; d++) {
    var idx = startCol + d - 1;
    var r = Math.floor(idx / 7), c = idx % 7;
    if (r > 5) break;
    var key = year + '-' + pad2_(month + 1) + '-' + pad2_(d);
    var rec = daily[key];
    if (rec && rec.trades) {
      grid[r][c] = d + '\n' + fmtSigned(rec.pnl, 2) + ' USD\n' + fmt(rec.trades, 0) + ' trades';
      colors[r][c] = rec.pnl >= 0 ? '#0f2e1f' : '#3b1418';
      fonts[r][c]  = rec.pnl >= 0 ? '#34d399' : '#f87171';
    } else {
      grid[r][c] = d + '\n—';
      colors[r][c] = '#0b1220';
      fonts[r][c] = '#4b5563';
    }
  }

  var body = sh.getRange(top + 1, 1, 6, 7);
  body.setValues(grid).setBackgrounds(colors).setFontColors(fonts)
      .setVerticalAlignment('top').setHorizontalAlignment('left')
      .setWrap(true).setFontSize(10);

  for (var col = 1; col <= 7; col++) sh.setColumnWidth(col, 130);
  for (var row = top + 1; row <= top + 6; row++) sh.setRowHeight(row, 62);
  sh.setHiddenGridlines(true);

  SpreadsheetApp.getActive().toast('สร้างปฏิทิน ' + monthKey + ' แล้ว', 'GTPro', 5);
}
// หมายเหตุ: pad2_ ย้ายไปอยู่ใน 00_Config.gs แล้ว (ใช้ร่วมกันหลายไฟล์)

/*******************************************************
 * 08_Telegram_Intake.gs — รับสัญญาณผ่าน Telegram
 *   ใช้แทน TradingView Alert สำหรับแพ็กเกจ Basic
 *   พิมพ์ในแชตบอท เช่น:  BUY SET 1590 tp 1600 sl 1585
 *   ระบบอ่านทุก 1 นาที แล้วบันทึกเข้า DB ให้อัตโนมัติ
 *******************************************************/

/** ▶ รันครั้งเดียวเพื่อเปิดใช้งานการอ่านข้อความจาก Telegram */
function installTelegramTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'pollTelegram') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pollTelegram').timeBased().everyMinutes(1).create();
  tgSend([
    '✅ <b>เปิดใช้งานการรับสัญญาณผ่าน Telegram แล้ว</b>',
    '',
    'พิมพ์สัญญาณในแชตนี้ได้เลย เช่น',
    '<code>BUY SET 1590 tp 1600 sl 1585</code>',
    '',
    'พิมพ์ /help เพื่อดูคำสั่งทั้งหมด'
  ].join('\n'));
  SpreadsheetApp.getActive().toast('เปิดรับสัญญาณทาง Telegram แล้ว', 'GTPro', 8);
}

/* ---------- ตัวแปลงข้อความสั้นเป็น JSON (pure — ทดสอบแยกได้) ---------- */

/**
 * แปลง "BUY SET 1590 tp 1600 sl 1585" → JSON ที่ parseSignal อ่านได้
 * รูปแบบ: ทิศทาง สัญลักษณ์ ราคา [tp x] [sl y] [tf z]
 */
function normalizeTgText(text, fallbackMode) {
  var t = String(text || '').trim();
  if (t.indexOf('{') >= 0) return t;               // เป็น JSON อยู่แล้ว

  var action = detectAction(t);
  if (!action) return t;                            // ให้ parseSignal ไปฟ้อง error เอง

  // อ่านวันเวลาย้อนหลัง + โหมด แล้วตัด @วันเวลา ทิ้งก่อนอ่านตัวเลขอื่น
  var when = grabWhen(t) || '';
  var mode = detectMode(t) || fallbackMode || '';
  t = stripWhen(t);

  var tp = grab(t, ['tp', 'take[\\s_]?profit', 'target', 'เป้า']);
  var sl = grab(t, ['sl', 'stop[\\s_]?loss', 'ตัดขาดทุน']);
  var tf = grabStr(t, ['tf', 'timeframe', 'interval']);

  // ตัดส่วน tp/sl/tf ออกก่อน เพื่อไม่ให้หยิบเลขพวกนั้นมาเป็นราคาเข้า
  var head = t.replace(
    /\b(tp|take\s*profit|target|sl|stop\s*loss|tf|timeframe|interval)\s*[:=]?\s*[0-9][0-9,\.]*/gi, ' ');

  var price = grab(head, ['price', 'entry', 'ราคา', 'เข้า']);
  if (price === null) {
    var m = head.match(/[0-9][0-9,]*(\.[0-9]+)?/);
    if (m) price = toNum(m[0]);
  }

  return JSON.stringify({
    symbol: symbolAfterAction_(head) || guessSymbol(head) || '',
    tf: tf || 'TG',
    action: action,
    price: price,
    tp: tp,
    sl: sl,
    mode: mode ? normMode(mode) : '',
    when: when
  });
}

/**
 * ในรูปแบบสั้น สัญลักษณ์คือคำถัดจากทิศทางเสมอ — "BUY SET 1590" → SET
 * แม่นกว่าการเดาแบบทั่วไป และกันคำอย่าง slippage มาแย่งตำแหน่ง
 */
function symbolAfterAction_(text) {
  var words = String(text).toUpperCase().match(/[A-Zก-๙][A-Z0-9ก-๙]*/g) || [];
  var isAction = function (w) { return detectAction(w) !== null; };
  for (var i = 0; i < words.length - 1; i++) {
    if (!isAction(words[i])) continue;
    var next = words[i + 1];
    if (!/^[A-Z][A-Z0-9]{1,11}$/.test(next)) return null;   // ไม่ใช่ชื่อสัญลักษณ์
    if (isAction(next) || SYMBOL_STOPWORDS.indexOf(next) >= 0) return null;
    return next;
  }
  return null;
}

/* ---------- ตัวอ่านข้อความจาก Telegram ---------- */

function pollTelegram() {
  var token = String(cfg('TELEGRAM_TOKEN', ''));
  var myChat = String(cfg('TELEGRAM_CHAT_ID', ''));
  if (!token || !myChat) return;

  var props = PropertiesService.getScriptProperties();
  var offset = props.getProperty('TG_OFFSET');
  var url = 'https://api.telegram.org/bot' + token + '/getUpdates?timeout=0&limit=20' +
            (offset ? '&offset=' + offset : '');

  var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return;
  var j = JSON.parse(res.getContentText());
  if (!j.ok || !j.result || !j.result.length) return;

  var lastId = 0;
  j.result.forEach(function (u) {
    lastId = Math.max(lastId, u.update_id);
    var msg = u.message || u.edited_message;
    if (!msg || !msg.chat) return;
    if (String(msg.chat.id) !== myChat) return;     // รับเฉพาะแชตของเจ้าของ

    if (msg.photo && msg.photo.length) {
      handleTgPhoto_(msg);                          // ส่งรูปชาร์ตเข้ามา
    } else if (msg.document && /^image\//.test(String(msg.document.mime_type || ''))) {
      handleTgPhoto_(msg);                          // ส่งเป็นไฟล์ (ไม่บีบอัด)
    } else if (msg.text) {
      handleTgText_(msg.text.trim());
    }
  });

  // ยืนยันว่าอ่านแล้ว เพื่อไม่ให้ประมวลผลซ้ำ
  props.setProperty('TG_OFFSET', String(lastId + 1));
}

/* ---------- รูปชาร์ต ---------- */

/**
 * รับรูปจากแชต แล้วผูกเข้ากับไม้เทรด
 *   • ถ้าใส่ caption เป็นสัญญาณ (เช่น "BUY SET 1590 tp 1600") → เปิดไม้ใหม่แล้วแนบรูปให้
 *   • ถ้าส่งรูปเปล่า ๆ → แนบกับไม้ล่าสุดที่บันทึกไว้
 *   • ถ้าใส่ caption เป็นรหัสไม้ (T2608...) → แนบกับไม้นั้น
 */
function handleTgPhoto_(msg) {
  var caption = String(msg.caption || '').trim();

  // เลือกไฟล์ความละเอียดสูงสุดที่ Telegram ส่งมา
  var fileId = msg.document ? msg.document.file_id
             : msg.photo[msg.photo.length - 1].file_id;

  // 1) caption เป็นสัญญาณ → บันทึกไม้ก่อน แล้วค่อยแนบรูป
  var opened = null;
  if (caption && detectAction(caption)) {
    var res = ingest(normalizeTgText(caption, defaultMode()), 'telegram');
    if (!res.ok) {
      return tgSend('⚠️ <b>อ่านสัญญาณใต้รูปไม่ออก</b>\n' + (res.error || '') +
                    '\n\nรูปยังไม่ได้บันทึก ลองส่งใหม่พร้อมข้อความแบบนี้:\n' +
                    '<code>BUY SET 1590 tp 1600 sl 1585</code>');
    }
    opened = true;
  }

  // 2) หาแถวของไม้ที่จะแนบรูป
  var target = caption ? findTradeRowById_(caption) : null;
  if (!target) target = latestTradeRow_();
  if (!target) return tgSend('ยังไม่มีไม้ให้แนบรูปครับ — ส่งสัญญาณก่อน หรือใส่ข้อความสัญญาณใต้รูปมาพร้อมกัน');

  // 3) โหลดรูปเก็บ Drive
  var sh = ss().getSheetByName(SHEETS.TRADES);
  var tradeId = String(sh.getRange(target, 1).getValue() || 'chart');
  var driveId = saveTelegramPhoto_(fileId, tradeId);
  if (!driveId) return tgSend('❌ บันทึกรูปไม่สำเร็จ — ลองส่งใหม่อีกครั้งครับ');

  attachChart_(target, driveId);

  var sym  = String(sh.getRange(target, 2).getValue() || '');
  var mode = normMode(sh.getRange(target, COL_MODE).getValue());
  tgSend('🖼 <b>แนบรูปเรียบร้อย</b>' + (opened ? ' (บันทึกไม้ใหม่ให้ด้วยแล้ว)' : '') +
         '\nไม้ : <code>' + tradeId + '</code>  ' + sym +
         (mode === MODE_TEST ? '  🧪 TEST' : '') +
         '\nดูรูปย้อนหลังได้ในแอป หน้า “ไม้เทรด”');
}

/** หาแถวของไม้จากรหัส trade_id */
function findTradeRowById_(text) {
  var m = String(text).match(/\bT\d{12}-\d{3}\b|\bDEMO-\d+\b/i);
  if (!m) return null;
  var sh = ss().getSheetByName(SHEETS.TRADES);
  var last = sh.getLastRow();
  if (last < 2) return null;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]).toUpperCase() === m[0].toUpperCase()) return i + 2;
  }
  return null;
}

function handleTgText_(text) {
  var lower = text.toLowerCase();

  if (lower === '/start' || lower === '/help' || lower === 'help') return tgHelp_();
  if (lower === '/today' || lower === 'วันนี้')  return sendDailySummary();
  if (lower === '/open'  || lower === 'ไม้เปิด') return tgOpenTrades_();
  if (lower === '/stats' || lower === 'สถิติ')   return tgStats_();
  if (lower.indexOf('/mode') === 0)              return tgMode_(text);
  if (lower.indexOf('/close') === 0)             return tgClose_(text);
  if (text.charAt(0) === '/') return tgSend('ไม่รู้จักคำสั่งนี้ครับ — พิมพ์ /help เพื่อดูรายการ');

  // ไม่ใช่คำสั่ง → ถือว่าเป็นสัญญาณ
  var res = ingest(normalizeTgText(text, defaultMode()), 'telegram');
  if (!res.ok) {
    tgSend('⚠️ <b>อ่านสัญญาณไม่ออก</b>\n' + (res.error || '') +
           '\n\nรูปแบบที่ถูกต้อง:\n<code>BUY SET 1590 tp 1600 sl 1585</code>');
  } else if (res.status === 'duplicate') {
    tgSend('ℹ️ สัญญาณนี้บันทึกไปแล้ว (กันซ้ำ)');
  }
  // ถ้าสำเร็จ notifyOpen/notifyClose จะส่งข้อความให้เองอยู่แล้ว
}

/** /mode — ดูหรือสลับโหมดเริ่มต้นระหว่าง LIVE กับ TEST */
function tgMode_(text) {
  var rest = String(text).replace(/^\/mode\s*/i, '').trim();
  if (!rest) {
    return tgSend('โหมดปัจจุบัน : <b>' + defaultMode() + '</b>\n\n' +
                  'สลับด้วย <code>/mode test</code> หรือ <code>/mode live</code>\n' +
                  'หรือพิมพ์คำว่า test ต่อท้ายสัญญาณเป็นราย ๆ ก็ได้');
  }
  var want = detectMode(rest);
  if (!want) return tgSend('ระบุ <code>/mode live</code> หรือ <code>/mode test</code> ครับ');

  var sh = ss().getSheetByName(SHEETS.CONFIG);
  var last = sh.getLastRow();
  var row = -1;
  var keys = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0]).trim() === 'DEFAULT_MODE') { row = i + 2; break; }
  }
  if (row < 0) { sh.appendRow(['DEFAULT_MODE', want, 'LIVE = เทรดจริง / TEST = backtest']); }
  else sh.getRange(row, 2).setValue(want);
  CACHE.config = null;

  tgSend(want === MODE_TEST
    ? '🧪 <b>เข้าโหมด TEST แล้ว</b>\nไม้ที่บันทึกต่อจากนี้จะแยกกระเป๋า ไม่ปนสถิติจริง\nกลับด้วย <code>/mode live</code>'
    : '✅ <b>กลับสู่โหมด LIVE แล้ว</b>\nไม้ที่บันทึกต่อจากนี้นับเป็นสถิติจริง');
}

/* ---------- คำสั่งต่าง ๆ ---------- */

function tgHelp_() {
  tgSend([
    '📖 <b>วิธีใช้ GTPro Signal DB</b>',
    '',
    '<b>บันทึกสัญญาณ</b> — พิมพ์ตามนี้',
    '<code>BUY SET 1590 tp 1600 sl 1585</code>',
    '<code>SELL XAUUSD 4440 tp 4420 sl 4450</code>',
    '<code>CLOSE SET 1595</code>',
    '',
    'ใส่ tp/sl หรือไม่ใส่ก็ได้',
    'ถ้าไม่ใส่ ไม้จะปิดตอนมีสัญญาณกลับทิศ',
    '',
    '<b>🖼 แนบรูปชาร์ต</b>',
    'ส่งรูปเข้ามาในแชตนี้ได้เลย',
    '• ใส่ข้อความสัญญาณเป็น caption ใต้รูป → บันทึกไม้ + แนบรูปในทีเดียว',
    '• ส่งรูปเปล่า → แนบกับไม้ล่าสุด',
    '• ใส่รหัสไม้เป็น caption → แนบกับไม้นั้น',
    'รูปเก็บใน Google Drive ของพี่เอง โฟลเดอร์ <b>GTPro_Charts</b>',
    '',
    '<b>🧪 โหมด TEST (backtest)</b>',
    '<code>/mode test</code> — ไม้ต่อจากนี้แยกกระเป๋า ไม่ปนสถิติจริง',
    '<code>/mode live</code> — กลับมาบันทึกเป็นไม้จริง',
    'หรือใส่คำว่า test ต่อท้ายเป็นราย ๆ',
    '<code>BUY SET 1590 tp 1600 test</code>',
    '',
    '<b>📅 บันทึกไม้ย้อนหลัง</b>',
    'ใส่ @ปี-เดือน-วัน (เวลา ใส่หรือไม่ใส่ก็ได้)',
    '<code>BUY SET 1585 tp 1600 test @2026-08-25 10:30</code>',
    '<code>CLOSE SET 1600 test @2026-08-26 14:00</code>',
    '',
    '<b>คำสั่ง</b>',
    '/today — สรุปกำไรวันนี้',
    '/open — ดูไม้ที่ยังเปิดอยู่',
    '/stats — สถิติรวมทั้งหมด',
    '/mode — ดู/สลับ LIVE ↔ TEST',
    '/close SET 1595 — ปิดไม้ที่ราคานี้',
    '/help — เมนูนี้'
  ].join('\n'));
}

function tgOpenTrades_() {
  var list = findOpenTrades_(null, null);
  if (!list.length) return tgSend('📭 ตอนนี้ไม่มีไม้เปิดอยู่ครับ');

  var lines = ['📂 <b>ไม้ที่เปิดอยู่ ' + fmt(list.length, 0) + ' ไม้</b>', '───────────────'];
  list.forEach(function (t) {
    var d = symbolMap_(t.symbol).digits;
    lines.push(
      (t.side === 'BUY' ? '🟢' : '🔴') + ' <b>' + t.symbol + '</b> ' + t.side +
      '  @ ' + fmt(t.entry, d) +
      (t.mode === MODE_TEST ? '  🧪 TEST' : '') +
      (t.chart ? '  🖼' : '') +
      (t.tp ? '\n   TP ' + fmt(t.tp, d) : '') +
      (t.sl ? '   SL ' + fmt(t.sl, d) : '') +
      '\n   <code>' + t.id + '</code>'
    );
  });
  tgSend(lines.join('\n'));
}

function tgStats_() {
  var sh = ss().getSheetByName(SHEETS.STATS);
  if (!sh || sh.getLastRow() < 2) return tgSend('ยังไม่มีสถิติ — ยังไม่มีไม้ที่ปิดแล้วครับ');

  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues();
  var all = null;
  rows.forEach(function (r) { if (String(r[0]) === 'ALL') all = r; });
  if (!all) return tgSend('ยังไม่มีสถิติรวม');

  tgSend([
    '📊 <b>สถิติรวมทั้งหมด</b>',
    '───────────────',
    'จำนวนไม้ : ' + fmt(all[1], 0) + ' ไม้',
    'ชนะ/แพ้ : ' + fmt(all[2], 0) + ' / ' + fmt(all[3], 0) +
      '  (' + fmt(all[4], 1) + '%)',
    'กำไรสุทธิ : <b>' + fmtSigned(all[10], 2) + ' USD</b>',
    'Profit Factor : ' + all[7],
    'Expectancy : ' + fmtSigned(all[8], 2) + ' USD/ไม้',
    'Max Drawdown : ' + fmt(all[9], 2) + ' USD'
  ].join('\n'));
}

function tgClose_(text) {
  var rest = text.replace(/^\/close\s*/i, '').trim();
  if (!rest) return tgSend('ระบุคู่ที่จะปิดด้วยครับ เช่น <code>/close SET 1595</code>');

  var when = grabWhen(rest) || '';
  var mode = detectMode(rest) || defaultMode();
  rest = stripWhen(rest);

  var sym = guessSymbol(rest);
  var price = null;
  var m = rest.replace(sym || '', ' ').match(/[0-9][0-9,]*(\.[0-9]+)?/);
  if (m) price = toNum(m[0]);

  var list = findOpenTrades_(sym, null, mode);
  if (!list.length) {
    return tgSend('ไม่พบไม้ ' + mode + ' ที่เปิดอยู่ของ ' + (sym || rest));
  }

  var n = 0;
  list.forEach(function (t) { closeTrade_(t, price, 'MANUAL', '', when); n++; });
  if (!n) tgSend('ปิดไม่สำเร็จ — ลองระบุราคาด้วย เช่น <code>/close ' + sym + ' 1595</code>');
}

/*******************************************************
 * 09_WebApi.gs — API สำหรับหน้า Dashboard (PWA)
 *
 *   GET  .../exec?api=1&key=XXXX            → JSON
 *   GET  .../exec?api=1&key=XXXX&callback=f → JSONP (กันปัญหา CORS)
 *
 * ▶ Deploy: ทำให้ใช้งานได้ › เว็บแอป › ผู้มีสิทธิ์เข้าถึง = ทุกคน
 *******************************************************/

/** router ของ doGet — ทับตัวเดิมใน 03_Ingest.gs */
function doGet(e) {
  var p = (e && e.parameter) || {};

  // เปิด URL เฉย ๆ ในเบราว์เซอร์ → เช็คว่าระบบยังมีชีวิต
  if (p.api !== '1') {
    return jsonOut_({ ok: true, service: 'GTPro Signal DB', time: nowStr() }, null);
  }

  var key = String(cfg('API_KEY', ''));
  if (key && String(p.key || '') !== key) {
    return jsonOut_({ ok: false, error: 'unauthorized' }, p.callback);
  }

  try {
    // ขอรูปชาร์ต 1 ใบ — ?api=1&key=XXX&img=<driveId>
    if (p.img) return jsonOut_(serveImage_(String(p.img)), p.callback);

    // บันทึก/ปิดไม้จากหน้าแอป
    if (p.act === 'add')    return jsonOut_(apiAddTrade_(p), p.callback);
    if (p.act === 'close')  return jsonOut_(apiCloseTrade_(p), p.callback);
    if (p.act === 'delete') return jsonOut_(apiDeleteTrade_(p), p.callback);
    if (p.act === 'news')   return jsonOut_({ ok: true, news: newsForApp_() }, p.callback);

    return jsonOut_(buildDashboardData_(p.month || '', p.mode || ''), p.callback);
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) }, p.callback);
  }
}

function jsonOut_(obj, callback) {
  var body = JSON.stringify(obj);
  if (callback && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- ประกอบข้อมูลทั้งหมดที่หน้า Dashboard ต้องใช้ ---------- */

function buildDashboardData_(monthKey, modeParam) {
  if (!/^\d{4}-\d{2}$/.test(String(monthKey))) {
    monthKey = Utilities.formatDate(new Date(), tz(), 'yyyy-MM');
  }
  var mode = normMode(modeParam || MODE_LIVE);

  return {
    ok: true,
    generated_at: nowStr(),
    tz: tz(),
    month: monthKey,
    mode: mode,
    default_mode: defaultMode(),
    has_test: hasTestData_(),
    symbols: readSymbols_(),
    default_lots: cfgNum('DEFAULT_LOTS', 1),
    daily: readDaily_(mode),
    open: readOpenWithFloating_(mode),
    history: readHistory_(mode, monthKey),
    stats: readStats_(),
    news: newsForApp_(),
    closed_count: countClosed_()
  };
}

/** ปฏิทิน + equity สะสม มาจากชีตเดียวกัน (กรองตามโหมด) */
function readDaily_(mode) {
  var sh = ss().getSheetByName(SHEETS.DAILY);
  if (!sh || sh.getLastRow() < 2) return [];
  var want = normMode(mode);
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  var out = [];
  rows.forEach(function (r) {
    if (!r[0]) return;
    if (normMode(r[1]) !== want) return;
    out.push({
      date: (r[0] instanceof Date) ? dateKey(r[0]) : String(r[0]),
      trades: Number(r[2]) || 0,
      wins: Number(r[3]) || 0,
      losses: Number(r[4]) || 0,
      winrate: Number(r[5]) || 0,
      pnl: Number(r[6]) || 0,
      cum: Number(r[7]) || 0
    });
  });
  out.sort(function (a, b) { return a.date < b.date ? -1 : 1; });
  return out;
}

/** มีไม้โหมด TEST อยู่ในระบบไหม — แอปจะได้ตัดสินใจว่าจะโชว์ปุ่มสลับหรือไม่ */
function hasTestData_() {
  var sh = ss().getSheetByName(SHEETS.TRADES);
  if (!sh || sh.getLastRow() < 2) return false;
  var v = sh.getRange(2, COL_MODE, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (normMode(v[i][0]) === MODE_TEST) return true;
  return false;
}

/**
 * ประวัติไม้ที่ปิดแล้วของเดือนนั้น — พร้อมรหัสรูปชาร์ต
 * เรียงใหม่สุดขึ้นก่อน จำกัด 200 ไม้กันข้อมูลบวม
 */
function readHistory_(mode, monthKey) {
  var sh = ss().getSheetByName(SHEETS.TRADES);
  if (!sh || sh.getLastRow() < 2) return [];
  var want = normMode(mode);
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.TRADES.length).getValues();
  var out = [];

  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    if (String(v[17]) !== 'CLOSED') continue;
    if (normMode(v[20]) !== want) continue;

    var exitTime = (v[9] instanceof Date)
      ? Utilities.formatDate(v[9], tz(), 'yyyy-MM-dd HH:mm:ss') : String(v[9] || '');
    if (monthKey && exitTime.substring(0, 7) !== monthKey) continue;

    var entryTime = (v[5] instanceof Date)
      ? Utilities.formatDate(v[5], tz(), 'yyyy-MM-dd HH:mm:ss') : String(v[5] || '');

    out.push({
      id: String(v[0]), symbol: String(v[1]), tf: String(v[2]), side: String(v[3]),
      lots: Number(v[4]) || 1,
      entry_time: entryTime, entry: Number(v[6]) || 0,
      tp: (v[7] === '' ? null : Number(v[7])),
      sl: (v[8] === '' ? null : Number(v[8])),
      exit_time: exitTime, exit: Number(v[10]) || 0, reason: String(v[11] || ''),
      pnl: Number(v[14]) || 0,
      pnl_pct: Number(v[13]) || 0,
      r: (v[15] === '' ? null : Number(v[15])),
      duration_min: Number(v[16]) || 0,
      digits: symbolMap_(String(v[1])).digits,
      chart: String(v[21] || ''),
      note: String(v[22] || '')
    });
  }

  out.sort(function (a, b) { return a.exit_time < b.exit_time ? 1 : -1; });
  return out.slice(0, 200);
}

/** ไม้ที่เปิดอยู่ + กำไรลอย (ราคาปัจจุบัน cache 60 วินาที) */
function readOpenWithFloating_(mode) {
  var list = findOpenTrades_(null, null, mode || null);
  if (!list.length) return [];

  var prices = {};
  list.forEach(function (t) {
    if (prices[t.symbol] === undefined) prices[t.symbol] = cachedPrice_(t.symbol);
  });

  return list.map(function (t) {
    var map = symbolMap_(t.symbol);
    var last = prices[t.symbol];
    var f = null;
    if (last !== null && last !== undefined) {
      f = computePnl(t.side, t.entry, last, t.lots, map.value_per_point, t.sl);
    }
    return {
      id: t.id, symbol: t.symbol, tf: t.tf, side: t.side, lots: t.lots,
      mode: t.mode, chart: t.chart, note: t.note,
      entry_time: (t.entry_time instanceof Date)
        ? Utilities.formatDate(t.entry_time, tz(), 'yyyy-MM-dd HH:mm:ss')
        : String(t.entry_time),
      entry: t.entry,
      tp: (t.tp === '' ? null : Number(t.tp)),
      sl: (t.sl === '' ? null : Number(t.sl)),
      digits: map.digits,
      last: (last === undefined ? null : last),
      float_usd: f ? round_(f.usd, 2) : null,
      float_pct: f ? round_(f.pct, 3) : null,
      float_r:   (f && f.r !== null) ? round_(f.r, 2) : null
    };
  });
}

function cachedPrice_(symbol) {
  var c = CacheService.getScriptCache();
  var k = 'px_' + symbol;
  var hit = c.get(k);
  if (hit !== null) return hit === 'NA' ? null : Number(hit);

  var v = null;
  try { v = getLastPrice_(symbol); } catch (e) { v = null; }
  c.put(k, (v === null || v === undefined) ? 'NA' : String(v), 60);
  return v;
}

function readStats_() {
  var sh = ss().getSheetByName(SHEETS.STATS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return {
        scope: String(r[0]), trades: Number(r[1]) || 0,
        wins: Number(r[2]) || 0, losses: Number(r[3]) || 0,
        winrate: Number(r[4]) || 0,
        avg_win: Number(r[5]) || 0, avg_loss: Number(r[6]) || 0,
        profit_factor: String(r[7]),
        expectancy: Number(r[8]) || 0,
        max_dd: Number(r[9]) || 0,
        total: Number(r[10]) || 0
      };
    });
}

function countClosed_() {
  var sh = ss().getSheetByName(SHEETS.TRADES);
  if (!sh || sh.getLastRow() < 2) return 0;
  var v = sh.getRange(2, 18, sh.getLastRow() - 1, 1).getValues();
  var n = 0;
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === 'CLOSED') n++;
  return n;
}

/* ---------- บันทึก / ปิด / ลบไม้ จากหน้าแอป ---------- */

/**
 * เปิดไม้จากฟอร์มในแอป
 * ?api=1&key=..&act=add&symbol=XAUUSD&side=BUY&price=4296&tp=4346.09&sl=4270.74
 *   &tf=30&lots=0.01&mode=TEST&when=2026-09-02%2009:30&note=...
 * ส่งต่อเข้า ingest() ตัวเดียวกับ Telegram/Webhook จะได้กันซ้ำและลง Signals_Raw เหมือนกัน
 */
function apiAddTrade_(p) {
  var side = detectAction(String(p.side || p.action || ''));
  if (!side) return { ok: false, error: 'ต้องระบุทิศเป็น BUY / SELL / CLOSE' };

  var payload = {
    symbol: String(p.symbol || '').toUpperCase().trim(),
    tf:     String(p.tf || 'APP').trim(),
    action: side,
    price:  toNum(p.price),
    tp:     toNum(p.tp),
    sl:     toNum(p.sl),
    lots:   toNum(p.lots),
    mode:   normMode(p.mode || defaultMode()),
    when:   String(p.when || '').trim(),
    note:   String(p.note || '').trim()
  };
  if (!payload.symbol) return { ok: false, error: 'ยังไม่ได้เลือกสัญลักษณ์' };
  if (payload.action !== 'CLOSE' && payload.price === null) {
    return { ok: false, error: 'ยังไม่ได้กรอกราคาเข้า' };
  }
  CACHE.lastTradeId = '';
  var res = ingest(JSON.stringify(payload), 'app');
  res.trade_id = CACHE.lastTradeId || '';   // แอปเอาไปแนบรูปต่อ
  return res;
}

/**
 * แนบรูปให้ไม้เทรด — เรียกผ่าน POST (รูปใหญ่เกินกว่าจะใส่ใน query string)
 * body: {act:'upload', key, id:'T2609..', b64, mime, name}
 */
function apiAttachImage_(o) {
  var row = tradeRowById_(o.id);
  if (!row) return { ok: false, error: 'ไม่พบไม้รหัส ' + (o.id || '(ว่าง)') };

  var b64 = String(o.b64 || '');
  if (!b64) return { ok: false, error: 'ไม่มีข้อมูลรูป' };
  // base64 ยาว 4 ตัวต่อ 3 ไบต์ — กันไฟล์ใหญ่เกินที่ Apps Script ไหว
  if (b64.length > 8 * 1024 * 1024) {
    return { ok: false, error: 'รูปใหญ่เกินไป ลองย่อขนาดก่อน' };
  }

  var sh = ss().getSheetByName(SHEETS.TRADES);
  var old = String(sh.getRange(row, COL_CHART).getValue() || '');
  var driveId = saveImageBytes_(b64, o.mime, String(o.id));
  if (!driveId) return { ok: false, error: 'บันทึกรูปลง Drive ไม่สำเร็จ' };

  attachChart_(row, driveId);
  // เปลี่ยนรูปใหม่ทับของเดิม → ย้ายรูปเก่าลงถังขยะ ไม่ให้ Drive รก
  if (old && old !== driveId) {
    try { DriveApp.getFileById(old).setTrashed(true); } catch (e) {}
  }
  return { ok: true, id: String(o.id), chart: driveId, replaced: !!old };
}

/** ปิดไม้ตามรหัส — ?act=close&id=T2609..&price=4346.09&when=2026-09-02%2014:00 */
function apiCloseTrade_(p) {
  var id = String(p.id || '').trim();
  if (!id) return { ok: false, error: 'ไม่ได้ระบุรหัสไม้' };

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) {
    return { ok: false, error: 'ระบบกำลังประมวลผลอย่างอื่นอยู่' };
  }
  try {
    var t = null;
    findOpenTrades_(null, null, null).forEach(function (x) { if (x.id === id) t = x; });
    if (!t) return { ok: false, error: 'ไม่พบไม้นี้ หรือปิดไปแล้ว' };

    var px = toNum(p.price);
    if (px === null && normMode(t.mode) === MODE_TEST) {
      return { ok: false, error: 'ไม้ TEST ต้องกรอกราคาปิดเอง (ระบบไม่ดึงราคาย้อนหลังให้)' };
    }
    closeTrade_(t, px, String(p.reason || 'MANUAL'), '', String(p.when || '').trim());
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}

/** ลบไม้ที่กรอกผิด — ?act=delete&id=T2609.. (ลบได้เฉพาะไม้ที่ยังเปิดอยู่) */
function apiDeleteTrade_(p) {
  var id = String(p.id || '').trim();
  if (!id) return { ok: false, error: 'ไม่ได้ระบุรหัสไม้' };
  var t = null;
  findOpenTrades_(null, null, null).forEach(function (x) { if (x.id === id) t = x; });
  if (!t) return { ok: false, error: 'ลบได้เฉพาะไม้ที่ยังเปิดอยู่' };
  ss().getSheetByName(SHEETS.TRADES).deleteRow(t.row);
  return { ok: true, id: id };
}

/** รายชื่อสัญลักษณ์จากชีต Symbol_Map — ให้แอปทำ dropdown */
function readSymbols_() {
  var sh = ss().getSheetByName(SHEETS.MAP);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .filter(function (r) { return r[0] && String(r[0]) !== '*'; })
    .map(function (r) {
      return { symbol: String(r[0]).toUpperCase(),
               vpp: Number(r[3]) || 1, digits: Number(r[4]) || 2 };
    });
}

/* ---------- ตัวช่วยตอนติดตั้ง ---------- */

/** สร้าง API key แบบสุ่มลงชีต Config แล้วบอกค่าที่ได้ */
function generateApiKey() {
  var chars = 'abcdefghijkmnpqrstuvwxyz23456789';
  var key = '';
  for (var i = 0; i < 24; i++) key += chars.charAt(Math.floor(Math.random() * chars.length));

  var sh = ss().getSheetByName(SHEETS.CONFIG);
  var last = sh.getLastRow();
  var keys = sh.getRange(2, 1, last - 1, 1).getValues();
  var row = -1;
  for (var j = 0; j < keys.length; j++) if (String(keys[j][0]) === 'API_KEY') { row = j + 2; break; }
  if (row < 0) { sh.appendRow(['API_KEY', key, 'รหัสสำหรับหน้า Dashboard']); row = sh.getLastRow(); }
  else sh.getRange(row, 2).setValue(key);

  CACHE.config = null;
  SpreadsheetApp.getUi().alert(
    'สร้าง API key แล้ว\n\n' + key +
    '\n\nนำไปกรอกในหน้า Dashboard ตอนเปิดครั้งแรก\n(ค่านี้อยู่ในชีต Config แถว API_KEY)');
}

/*******************************************************
 * 10_Charts.gs — รูปชาร์ตประกอบไม้เทรด
 *   รับรูปจากแชต Telegram → เก็บใน Google Drive ของเจ้าของ
 *   แล้วเสิร์ฟกลับให้แอปแบบส่วนตัว (ไม่ต้องแชร์ไฟล์ให้ใคร)
 *******************************************************/

var CHART_FOLDER = 'GTPro_Charts';

/** หาโฟลเดอร์เก็บรูป ถ้ายังไม่มีก็สร้างให้ */
function chartFolder_() {
  if (CACHE.folder) return CACHE.folder;
  var it = DriveApp.getFoldersByName(CHART_FOLDER);
  CACHE.folder = it.hasNext() ? it.next() : DriveApp.createFolder(CHART_FOLDER);
  return CACHE.folder;
}

/**
 * ดึงไฟล์รูปจาก Telegram แล้วบันทึกลง Drive
 * @return {string|null} id ของไฟล์ใน Drive
 */
function saveTelegramPhoto_(fileId, nameHint) {
  var token = String(cfg('TELEGRAM_TOKEN', ''));
  if (!token || !fileId) return null;

  try {
    var meta = UrlFetchApp.fetch(
      'https://api.telegram.org/bot' + token + '/getFile?file_id=' + encodeURIComponent(fileId),
      { muteHttpExceptions: true });
    if (meta.getResponseCode() !== 200) return null;

    var j = JSON.parse(meta.getContentText());
    if (!j.ok || !j.result || !j.result.file_path) return null;

    var bin = UrlFetchApp.fetch(
      'https://api.telegram.org/file/bot' + token + '/' + j.result.file_path,
      { muteHttpExceptions: true });
    if (bin.getResponseCode() !== 200) return null;

    var ext = (j.result.file_path.match(/\.([a-z0-9]+)$/i) || [, 'jpg'])[1];
    var name = (nameHint || 'chart') + '_' +
               Utilities.formatDate(new Date(), tz(), 'yyyyMMdd_HHmmss') + '.' + ext;

    var blob = bin.getBlob().setName(name);
    return chartFolder_().createFile(blob).getId();

  } catch (e) {
    Logger.log('บันทึกรูปไม่สำเร็จ: ' + e);
    return null;
  }
}

/**
 * บันทึกรูปที่แอปส่งมาเป็น base64 ลง Drive
 * @return {string|null} id ของไฟล์ใน Drive
 */
function saveImageBytes_(b64, mime, nameHint) {
  try {
    if (!b64) return null;
    var m = String(mime || 'image/jpeg');
    var ext = m.indexOf('png') >= 0 ? 'png' : m.indexOf('webp') >= 0 ? 'webp' : 'jpg';
    var name = (nameHint || 'chart') + '_' +
               Utilities.formatDate(new Date(), tz(), 'yyyyMMdd_HHmmss') + '.' + ext;
    var blob = Utilities.newBlob(Utilities.base64Decode(b64), m, name);
    return chartFolder_().createFile(blob).getId();
  } catch (e) {
    Logger.log('บันทึกรูปจากแอปไม่สำเร็จ: ' + e);
    return null;
  }
}

/** หาแถวของไม้จากรหัส — ใช้ได้ทั้งไม้ที่เปิดอยู่และปิดไปแล้ว */
function tradeRowById_(id) {
  var sh = ss().getSheetByName(SHEETS.TRADES);
  var last = sh.getLastRow();
  if (last < 2 || !id) return null;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return null;
}

/** ผูกรูปเข้ากับไม้เทรด (คอลัมน์ chart_id) */
function attachChart_(tradeRow, driveId) {
  if (!tradeRow || !driveId) return;
  ss().getSheetByName(SHEETS.TRADES).getRange(tradeRow, COL_CHART).setValue(driveId);
}

/** หาไม้ล่าสุดที่ยังไม่มีรูป — ใช้ตอนส่งรูปตามหลังข้อความ */
function latestTradeRow_() {
  var sh = ss().getSheetByName(SHEETS.TRADES);
  var last = sh.getLastRow();
  return last >= 2 ? last : null;
}

/**
 * ส่งรูปกลับให้แอปเป็น base64 (ไม่ต้องเปิดสิทธิ์ไฟล์ให้สาธารณะ)
 * เรียกผ่าน  ?api=1&key=XXX&img=<driveId>
 */
function serveImage_(driveId) {
  try {
    var f = DriveApp.getFileById(driveId);
    var blob = f.getBlob();
    var bytes = blob.getBytes();
    if (bytes.length > 6 * 1024 * 1024) {
      return { ok: false, error: 'รูปใหญ่เกินไป (' + Math.round(bytes.length / 1024) + ' KB)' };
    }
    return {
      ok: true,
      name: f.getName(),
      mime: blob.getContentType(),
      size: bytes.length,
      b64: Utilities.base64Encode(bytes)
    };
  } catch (e) {
    return { ok: false, error: 'ไม่พบรูปนี้' };
  }
}

/** ล้างรูปที่ไม่มีไม้ไหนอ้างถึงแล้ว (เรียกเองเมื่ออยากเก็บกวาด) */
function cleanupOrphanCharts() {
  var sh = ss().getSheetByName(SHEETS.TRADES);
  var used = {};
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, COL_CHART, sh.getLastRow() - 1, 1).getValues()
      .forEach(function (r) { if (r[0]) used[String(r[0])] = true; });
  }
  var files = chartFolder_().getFiles(), n = 0;
  while (files.hasNext()) {
    var f = files.next();
    if (!used[f.getId()]) { f.setTrashed(true); n++; }
  }
  SpreadsheetApp.getActive().toast('ย้ายรูปที่ไม่ได้ใช้ไปถังขยะ ' + fmt(n, 0) + ' ไฟล์', 'GTPro', 6);
}

/*******************************************************
 * 11_News.gs — ปฏิทินข่าวเศรษฐกิจที่มีผลกับทอง/ดอลลาร์
 *   ดึงฝั่งเซิร์ฟเวอร์ด้วย UrlFetchApp (ไม่ติด CORS เหมือนยิงจากเบราว์เซอร์)
 *   แคชไว้ 30 นาที แอปเรียกผ่าน API เดิม ไม่ต้องยิงเว็บนอกเอง
 *******************************************************/

var NEWS_URL_WEEK = 'https://nfs.faireconomy.media/ff_calendar_thisweek.json';
var NEWS_URL_NEXT = 'https://nfs.faireconomy.media/ff_calendar_nextweek.json';
var NEWS_CACHE_KEY = 'news_v1';
var NEWS_CACHE_SEC = 1800;          // 30 นาที
var NEWS_ALERT_MIN = 30;            // เตือนล่วงหน้ากี่นาที

/**
 * ผลของ "ตัวเลขจริงออกมาสูงกว่าที่คาด" ต่อราคาทอง
 *   -1 = ทองมีแนวโน้มลง (ดอลลาร์แข็ง)
 *   +1 = ทองมีแนวโน้มขึ้น
 *    0 = ตีความจากตัวเลขอย่างเดียวไม่ได้ ต้องฟังเนื้อหา
 */
function goldDirection_(title) {
  var t = String(title || '').toLowerCase();

  // ตัวเลขยิ่งสูง = เศรษฐกิจยิ่งแย่ → ทองขึ้น
  if (/unemployment rate|jobless claims|continuing claims/.test(t)) return 1;

  // การจ้างงาน / เงินเฟ้อ / กิจกรรมเศรษฐกิจ → สูง = ดอลลาร์แข็ง = ทองลง
  if (/non-farm|nonfarm|payroll|employment change|adp/.test(t))         return -1;
  if (/\bcpi\b|\bppi\b|inflation|\bpce\b/.test(t))                      return -1;
  if (/retail sales|\bgdp\b|durable goods|industrial production/.test(t))return -1;
  if (/\bism\b|\bpmi\b|consumer confidence|consumer sentiment/.test(t)) return -1;
  if (/federal funds rate|interest rate decision/.test(t))              return -1;
  if (/average hourly earnings|employment cost/.test(t))                return -1;

  return 0;
}

/** เหตุการณ์ที่ไม่มีตัวเลข แต่สำคัญกับทองมาก (ต้องฟังเนื้อหา) */
function isTalkEvent_(title) {
  return /fomc|powell|fed chair|press conference|statement|minutes|testimony|speaks/i
         .test(String(title || ''));
}

/** ดึงปฏิทินดิบ 2 สัปดาห์ */
function fetchNewsRaw_() {
  var all = [];
  [NEWS_URL_WEEK, NEWS_URL_NEXT].forEach(function (u) {
    try {
      var r = UrlFetchApp.fetch(u, {
        muteHttpExceptions: true, followRedirects: true,
        headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' }
      });
      if (r.getResponseCode() !== 200) return;
      var j = JSON.parse(r.getContentText());
      if (j instanceof Array) all = all.concat(j);
    } catch (e) {
      Logger.log('ดึงปฏิทินข่าวไม่สำเร็จ ' + u + ': ' + e);
    }
  });
  return all;
}

/**
 * แปลงข้อมูลดิบเป็นรูปแบบที่แอปใช้
 * เก็บเฉพาะ USD (ดอลลาร์คุมราคาทอง) ระดับ High/Medium
 */
function normalizeNews_(raw) {
  var out = [];
  for (var i = 0; i < raw.length; i++) {
    var e = raw[i];
    if (String(e.country || '').toUpperCase() !== 'USD') continue;
    var im = String(e.impact || '');
    if (im !== 'High' && im !== 'Medium') continue;

    var d = parseDate_(e.date);
    if (!d || isNaN(d.getTime())) continue;

    out.push({
      id: String(e.title || '').replace(/[^A-Za-z0-9]/g, '').substring(0, 16) + '_' +
          Utilities.formatDate(d, 'UTC', 'yyyyMMddHHmm'),
      title: String(e.title || ''),
      impact: im,
      when: Utilities.formatDate(d, tz(), 'yyyy-MM-dd HH:mm'),
      ts: d.getTime(),
      forecast: String(e.forecast || ''),
      previous: String(e.previous || ''),
      actual: String(e.actual || ''),
      dir: goldDirection_(e.title),
      talk: isTalkEvent_(e.title)
    });
  }
  out.sort(function (a, b) { return a.ts - b.ts; });
  return out;
}

/** อ่านปฏิทิน (มีแคช + ตัวสำรองเผื่อแหล่งข้อมูลล่ม) */
function readNews_(force) {
  var c = CacheService.getScriptCache();
  if (!force) {
    var hit = c.get(NEWS_CACHE_KEY);
    if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  }

  var list = normalizeNews_(fetchNewsRaw_());
  if (list.length) {
    try { c.put(NEWS_CACHE_KEY, JSON.stringify(list), NEWS_CACHE_SEC); } catch (e) {}
    try {
      PropertiesService.getScriptProperties()
        .setProperty('NEWS_BACKUP', JSON.stringify(list).substring(0, 400000));
    } catch (e) {}
    return list;
  }

  // ดึงไม่ได้ → ใช้ชุดล่าสุดที่เคยดึงสำเร็จ ดีกว่าหน้าว่าง
  try {
    var b = PropertiesService.getScriptProperties().getProperty('NEWS_BACKUP');
    if (b) return JSON.parse(b);
  } catch (e) {}
  return [];
}

/**
 * ข่าวที่ส่งให้แอป — ย้อนหลัง 18 ชม. ถึงอีก 7 วัน
 * ใส่ mins (นาทีที่เหลือ, ติดลบ = ผ่านไปแล้ว) ให้แอปนับถอยหลัง/เตือนได้
 */
function newsForApp_() {
  var all = readNews_(false);
  var now = Date.now();
  var from = now - 18 * 3600 * 1000;
  var to   = now + 7 * 24 * 3600 * 1000;

  var out = [];
  for (var i = 0; i < all.length; i++) {
    var e = all[i];
    if (e.ts < from || e.ts > to) continue;
    out.push({
      id: e.id, title: e.title, impact: e.impact, when: e.when,
      forecast: e.forecast, previous: e.previous, actual: e.actual,
      dir: e.dir, talk: e.talk,
      ts: e.ts,                                  // ให้แอปนับถอยหลังเองได้ ไม่เพี้ยนตอนอ่านจากแคช
      mins: Math.round((e.ts - now) / 60000)
    });
  }
  return out;
}

/* ---------- แจ้งเตือนเข้า Telegram ก่อนถึงเวลาข่าว ---------- */

/** ▶ รันครั้งเดียวเพื่อเปิดการเตือนข่าว (ตรวจทุก 5 นาที) */
function installNewsTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'newsWatch') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('newsWatch').timeBased().everyMinutes(5).create();
  readNews_(true);
  tgSend([
    '📰 <b>เปิดการเตือนข่าวเศรษฐกิจแล้ว</b>',
    '',
    'เตือนล่วงหน้า ' + NEWS_ALERT_MIN + ' นาที',
    'เฉพาะข่าว USD ระดับแรง (High) ที่มีผลกับทอง'
  ].join('\n'));
  SpreadsheetApp.getActive().toast('เปิดการเตือนข่าวแล้ว', 'GTPro', 8);
}

/** trigger ทุก 5 นาที — ข่าวแรงที่ใกล้ถึงเวลาและยังไม่เคยเตือน → ส่ง Telegram */
function newsWatch() {
  var list = readNews_(false);
  var now = Date.now();
  var props = PropertiesService.getScriptProperties();
  var sent = {};
  try { sent = JSON.parse(props.getProperty('NEWS_SENT') || '{}'); } catch (e) { sent = {}; }

  var changed = false;
  for (var i = 0; i < list.length; i++) {
    var e = list[i];
    if (e.impact !== 'High') continue;
    var mins = Math.round((e.ts - now) / 60000);
    if (mins < 0 || mins > NEWS_ALERT_MIN) continue;
    if (sent[e.id]) continue;

    tgSend(newsMessage_(e, mins));
    sent[e.id] = now;
    changed = true;
  }

  // ล้างของเก่ากว่า 2 วัน กัน property บวม
  for (var k in sent) {
    if (now - sent[k] > 2 * 24 * 3600 * 1000) { delete sent[k]; changed = true; }
  }
  if (changed) props.setProperty('NEWS_SENT', JSON.stringify(sent));
}

function newsMessage_(e, mins) {
  var bias =
    e.dir === -1 ? 'ออกมา <b>สูงกว่าคาด</b> → ทองมีแนวโน้ม <b>ลง</b> 🔴\n' +
                   'ออกมา <b>ต่ำกว่าคาด</b> → ทองมีแนวโน้ม <b>ขึ้น</b> 🟢'
  : e.dir === 1  ? 'ออกมา <b>สูงกว่าคาด</b> → ทองมีแนวโน้ม <b>ขึ้น</b> 🟢\n' +
                   'ออกมา <b>ต่ำกว่าคาด</b> → ทองมีแนวโน้ม <b>ลง</b> 🔴'
  :                'ต้องฟังเนื้อหา ตีความจากตัวเลขอย่างเดียวไม่ได้';

  return [
    '⚠️ <b>อีก ' + fmt(mins, 0) + ' นาที มีข่าวแรง</b>',
    '',
    '<b>' + e.title + '</b>',
    '🕐 ' + e.when + ' น.',
    (e.forecast ? 'คาดการณ์ : ' + e.forecast : null),
    (e.previous ? 'ครั้งก่อน : ' + e.previous : null),
    '───────────────',
    bias,
    '',
    '<i>ช่วงข่าวแรงราคาเหวี่ยง สเปรดกว้าง ระวังไม้ที่เปิดค้างไว้</i>'
  ].filter(function (x) { return x; }).join('\n');
}

/** ทดสอบเอง — ดูว่าดึงข่าวได้กี่รายการ อันถัดไปคืออะไร */
function testNews() {
  var l = newsForApp_();
  var next = null;
  for (var i = 0; i < l.length; i++) if (l[i].mins >= 0) { next = l[i]; break; }
  SpreadsheetApp.getUi().alert(
    'ข่าวเศรษฐกิจ',
    'ดึงมาได้ ' + l.length + ' รายการ (USD ระดับ High/Medium)\n\n' +
    (next ? 'อันถัดไป: ' + next.title + '\n' + next.when + ' น. (อีก ' +
            next.mins + ' นาที)\nระดับ ' + next.impact
          : 'ไม่มีข่าวข้างหน้าในช่วง 7 วัน'),
    SpreadsheetApp.getUi().ButtonSet.OK);
}

