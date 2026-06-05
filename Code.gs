/**
 * 日経225 寄与度 Top 10（Google Apps Script）
 *
 * スプレッドシート:
 *   https://docs.google.com/spreadsheets/d/1jKzRFVh7UQh_pLoxrLrRTAMXOX3VfczO3ZLVvX-sc9o/edit
 * Apps Script:
 *   https://script.google.com/home/projects/13tfPkuquhp2My2fkcUTFEGQC0dPRwF4c4Dgnjc72uNwlq4joGwJb2Sgu/edit
 *
 * clasp push:
 *   clasp push
 *
 * 初回セットアップ:
 *   setupProperties() で SLACK_WEBHOOK_URL を設定
 *
 * データソース:
 *   - nikkei225jp.com（寄与度・株価・日経225指数）
 *   - https://nikkei225jp.com/chart/nikkei.php
 */

const CONFIG = {
  SPREADSHEET_ID: '1jKzRFVh7UQh_pLoxrLrRTAMXOX3VfczO3ZLVvX-sc9o',
  SHEET_SUMMARY: 'Summary',
  SHEET_ALL: 'All',
  TRIGGER_HOUR_JST: 15,
  TRIGGER_MINUTE_JST: 35,
  DATA_BASE_URL: 'https://nikkei225jp.com',
  DATA_REFERER: 'https://nikkei225jp.com/chart/nikkei.php',
  PATH_KIYO10: '/_data/_nfsWEB/min/country_jp_kiyo10N.js',
  PATH_NK225: '/_data/_nfsWEB/min/country_jp_nk225N.js',
  PATH_INDEX: '/_data/_nfsWEB/ajaxindex/ajax_NDY_min.js',
};

// 東証休場日（yyyy-MM-dd, Asia/Tokyo）
const JP_MARKET_HOLIDAYS = [
  '2026-01-01', '2026-01-02', '2026-01-12', '2026-02-11', '2026-02-23',
  '2026-03-20', '2026-04-29', '2026-05-03', '2026-05-04', '2026-05-05', '2026-05-06',
  '2026-07-20', '2026-08-11', '2026-09-21', '2026-09-22', '2026-09-23',
  '2026-10-12', '2026-11-03', '2026-11-23', '2026-12-31',
  '2027-01-01', '2027-01-02', '2027-01-11', '2027-02-11', '2027-02-23',
  '2027-03-21', '2027-03-22', '2027-04-29', '2027-05-03', '2027-05-04', '2027-05-05',
  '2027-07-19', '2027-08-11', '2027-09-20', '2027-09-23', '2027-10-11',
  '2027-11-03', '2027-11-23', '2027-12-31',
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('日経225')
    .addItem('寄与度を更新', 'updateContribution')
    .addItem('Slack テスト通知', 'testSlackNotification')
    .addItem('初回セットアップ', 'setupSheets')
    .addSeparator()
    .addItem('後場終了トリガーを設定', 'installDailyTrigger')
    .addItem('トリガーを削除', 'removeTriggers')
    .addToUi();
}

/**
 * 初回のみ Apps Script エディタから実行してください。
 * SLACK_WEBHOOK_URL を日経用チャンネルの Webhook に差し替えてから Run。
 */
function setupProperties() {
  PropertiesService.getScriptProperties().setProperties({
    SLACK_WEBHOOK_URL: 'https://hooks.slack.com/services/xxx/xxx/xxx',
  });
  Logger.log('スクリプトプロパティを保存しました');
}

function setupSheets() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, CONFIG.SHEET_SUMMARY);
  ensureSheet_(ss, CONFIG.SHEET_ALL);
  ss.setActiveSheet(ss.getSheetByName(CONFIG.SHEET_SUMMARY));
  safeUiAlert_('シートの準備が完了しました。\n「寄与度を更新」を実行してください。');
}

function installDailyTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('runAfterAfternoonClose')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.TRIGGER_HOUR_JST)
    .nearMinute(CONFIG.TRIGGER_MINUTE_JST)
    .inTimezone('Asia/Tokyo')
    .create();

  try {
    safeUiAlert_(
      `後場終了トリガーを設定しました（JST ${CONFIG.TRIGGER_HOUR_JST}:${String(CONFIG.TRIGGER_MINUTE_JST).padStart(2, '0')}）。\n` +
        '日本の取引日（平日・休場日除く）のみ更新・Slack通知します。'
    );
  } catch (e) {
    Logger.log('トリガー設定完了（UIなし）');
  }
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach((t) => {
    const fn = t.getHandlerFunction();
    if (fn === 'runAfterAfternoonClose' || fn === 'updateContribution') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/** 時間トリガーから呼ばれるエントリポイント（後場終了後） */
function runAfterAfternoonClose() {
  if (!isJpTradingDay_(new Date())) {
    Logger.log('本日は日本取引日ではないためスキップ');
    return;
  }
  updateContribution({ notifySlack: true });
}

function updateContribution(options) {
  options = options || {};
  Logger.log('更新開始');
  const ss = getSpreadsheet_();
  setupSheetsQuiet_(ss);

  Logger.log('nikkei225jp.com からデータ取得中…');
  const data = fetchNikkei225jpData_();
  Logger.log(`取得完了: ${data.allRows.length} 銘柄（更新 ${data.lastTime}）`);

  const updatedAt = buildUpdatedAt_(data.lastTime);

  Logger.log('シート書き込み中…');
  writeAllSheet_(ss, data.allRows, updatedAt, data.nikkei, data.totalContribYen);
  writeSummarySheet_(
    ss,
    data.allRows,
    updatedAt,
    data.nikkei,
    data.totalContribYen,
    data.topUp,
    data.topDown
  );
  Logger.log('更新完了');

  if (options.notifySlack) {
    Logger.log('Slack 通知中…');
    notifySlack_({
      updatedAt,
      nikkei: data.nikkei,
      totalContribYen: data.totalContribYen,
      topUp: data.topUp,
      topDown: data.topDown,
      counts: data.counts,
      spreadsheetUrl: ss.getUrl(),
      isTest: !!options.testSlack,
    });
    Logger.log('Slack 通知完了');
  }

  Logger.log('実行完了');
}

function testSlackNotification() {
  updateContribution({ notifySlack: true, testSlack: true });
  Logger.log('Slack テスト通知を送信しました');
}

// ---------------------------------------------------------------------------
// Slack
// ---------------------------------------------------------------------------

const SLACK = {
  USERNAME: '日経225 Watcher',
  ICON_EMOJI: ':chart_with_upwards_trend:',
  COLOR_UP: '#36a64f',
  COLOR_DOWN: '#e01e5a',
  COLOR_FLAT: '#949494',
};

function notifySlack_(data) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty(
    'SLACK_WEBHOOK_URL'
  );
  if (!webhookUrl) {
    throw new Error(
      'SLACK_WEBHOOK_URL が未設定です。setupProperties() を実行してください。'
    );
  }

  const nikkeiPct = data.nikkei ? data.nikkei.changePct : 0;
  const topUpRow = data.topUp[0];
  const topDownRow = data.topDown[0];
  const isTest = !!data.isTest;

  const headline = data.nikkei
    ? `📊 ${isTest ? 'テスト通知: ' : '後場終了: '}日経225 *${formatPct_(nikkeiPct)}*` +
      ` — 上昇寄与 ${formatStockLabel_(topUpRow)}、下落寄与 ${formatStockLabel_(topDownRow)}`
    : '📊 日経225 寄与度レポート';

  const metaLine =
    '🏷️ 日経225 · 📅 ' +
    formatSlackDate_(data.updatedAt + (isTest ? '（テスト）' : '')) +
    ' · 🕐 後場終了後';

  const nikkeiDetail = data.nikkei
    ? formatNikkeiLine_(data.nikkei) +
      `  ·  合計寄与 *${formatContribYen_(data.totalContribYen)}*`
    : '_日経225 データ取得失敗_';

  const bodyLines = [
    metaLine,
    '',
    nikkeiDetail,
    '',
    '*📈 上昇寄与 Top 10*',
    formatSlackRank_(data.topUp),
    '',
    '*📉 下落寄与 Top 10*',
    formatSlackRank_(data.topDown),
  ];

  if (data.spreadsheetUrl) {
    bodyLines.push('', `<${data.spreadsheetUrl}|📋 スプレッドシートで詳細を見る>`);
  }

  const attachment = {
    color: slackColorForChange_(nikkeiPct),
    title: '📊 日経225 寄与度 Top 10',
    text: bodyLines.join('\n'),
    mrkdwn_in: ['text'],
    footer: 'nikkei225jp.com | 日経225 Watcher',
    ts: Math.floor(Date.now() / 1000),
  };

  if (data.spreadsheetUrl) {
    attachment.title_link = data.spreadsheetUrl;
  }

  const payload = {
    username: SLACK.USERNAME,
    icon_emoji: SLACK.ICON_EMOJI,
    text: headline,
    attachments: [attachment],
  };

  const response = UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() >= 400) {
    throw new Error(
      'Slack 通知失敗: ' + response.getResponseCode() + ' ' + response.getContentText()
    );
  }
}

function formatSlackRank_(rows) {
  if (!rows || rows.length === 0) return '_—_';
  return rows
    .map((r, i) => {
      const contrib = formatContribYen_(r.contributionYen);
      const weight =
        r.weightPct > 0 ? `  W${r.weightPct.toFixed(2)}%` : '';
      return (
        `${String(i + 1).padStart(2, ' ')}. ${formatStockLabel_(r)}` +
        `\n     \`${contrib}\`  ${formatPct_(r.pctChg)}${weight}`
      );
    })
    .join('\n');
}

function formatStockLabel_(row) {
  if (!row) return '—';
  const code = row.code;
  if (!code) return '—';
  const name = shortenCompanyName_(row.company || code);
  return `*${code}*（${name}）`;
}

function shortenCompanyName_(name) {
  return String(name)
    .replace(/（株）/g, '')
    .replace(/\(株\)/g, '')
    .replace(/株式会社/g, '')
    .replace(/ホールディングス/g, 'HD')
    .replace(/グループ/g, 'G')
    .replace(/\s+Common Stock$/i, '')
    .replace(/\s+Capital Stock$/i, '')
    .trim();
}

function formatNikkeiLine_(nikkei) {
  const label = nikkei.marketState === 'estimated' ? '*日経225（概算）*' : '*日経225*';
  if (nikkei.price == null) {
    return `${label} ${formatPct_(nikkei.changePct)}`;
  }
  return `${label} ${formatPct_(nikkei.changePct)}（${formatYen_(nikkei.price)}）`;
}

function formatContribYen_(value) {
  const sign = value >= 0 ? '+' : '';
  return sign + value.toFixed(2) + '円';
}

function formatSlackDate_(updatedAt) {
  const cleaned = String(updatedAt).replace('（テスト）', '').trim();
  const m = cleaned.match(/^(\d{4}-\d{2}-\d{2}) (\d{2}:\d{2})/);
  return m ? `${m[1]} ${m[2]}` : cleaned;
}

function slackColorForChange_(pct) {
  if (pct > 0) return SLACK.COLOR_UP;
  if (pct < 0) return SLACK.COLOR_DOWN;
  return SLACK.COLOR_FLAT;
}

function formatYen_(value) {
  return '¥' + Number(value).toLocaleString('ja-JP', { maximumFractionDigits: 2 });
}

// ---------------------------------------------------------------------------
// Data fetching (nikkei225jp.com)
// ---------------------------------------------------------------------------

function fetchNikkei225jpData_() {
  const headers = {
    'User-Agent': 'Mozilla/5.0',
    Referer: CONFIG.DATA_REFERER,
  };
  const urls = [
    CONFIG.DATA_BASE_URL + CONFIG.PATH_KIYO10,
    CONFIG.DATA_BASE_URL + CONFIG.PATH_NK225,
    CONFIG.DATA_BASE_URL + CONFIG.PATH_INDEX,
  ];
  const responses = UrlFetchApp.fetchAll(
    urls.map((url) => ({ url, headers, muteHttpExceptions: true }))
  );

  responses.forEach((response, i) => {
    if (response.getResponseCode() >= 400) {
      throw new Error('データ取得失敗: ' + urls[i] + ' (' + response.getResponseCode() + ')');
    }
  });

  const kiyo10Text = responses[0].getContentText('UTF-8');
  const nk225Text = responses[1].getContentText('UTF-8');
  const indexText = responses[2].getContentText('UTF-8');

  const weightByCode = {};
  const allRows = parseJsArrayEntries_(nk225Text, 'N2').map((line) => {
    const row = parseN2Row_(line.split('__'));
    weightByCode[row.code] = row.weightPct;
    return row;
  });
  if (allRows.length < 200) {
    throw new Error('構成銘柄データが不足しています: ' + allRows.length + ' 件');
  }

  const topUp = parseJsArrayEntries_(kiyo10Text, 'top10').map((line) =>
    attachWeight_(parseKiyo10Row_(line.split('__')), weightByCode)
  );
  const topDown = parseJsArrayEntries_(kiyo10Text, 'las10').map((line) =>
    attachWeight_(parseKiyo10Row_(line.split('__')), weightByCode)
  );

  const nikkei = parseNikkeiIndex_(indexText, nk225Text);
  const lastTime = parseJsQuotedVar_(nk225Text, 'LastTime') || '';
  const totalContribYen = allRows.reduce((sum, r) => sum + r.contributionYen, 0);

  return {
    allRows,
    topUp,
    topDown,
    nikkei,
    lastTime,
    totalContribYen,
    counts: {
      up: parseJsNumberVar_(nk225Text, 'CntUp'),
      down: parseJsNumberVar_(nk225Text, 'CntDwn'),
      flat: parseJsNumberVar_(nk225Text, 'CntEvn'),
    },
  };
}

function attachWeight_(row, weightByCode) {
  row.weightPct = weightByCode[row.code] || 0;
  return row;
}

function parseKiyo10Row_(fields) {
  if (!fields || fields.length < 6) {
    throw new Error('寄与度 Top10 データの形式が不正です');
  }
  return {
    code: fields[0],
    company: fields[1],
    contributionYen: toNumber_(fields[2]),
    price: toNumber_(fields[3]),
    changeYen: toNumber_(fields[4]),
    pctChg: toNumber_(fields[5]),
    weightPct: 0,
    prevClose: toNumber_(fields[3]) - toNumber_(fields[4]),
  };
}

function parseN2Row_(fields) {
  if (!fields || fields.length < 9) {
    throw new Error('構成銘柄データの形式が不正です');
  }
  const price = toNumber_(fields[3]);
  const changeYen = toNumber_(fields[7]);
  return {
    code: fields[0],
    company: fields[2],
    price,
    weightPct: toNumber_(String(fields[5]).replace('%', '')),
    pctChg: toNumber_(fields[6]),
    changeYen,
    contributionYen: toNumber_(fields[8]),
    prevClose: price - changeYen,
  };
}

function parseNikkeiIndex_(indexText, nk225Text) {
  const match = indexText.match(/NDY111V=([^,;\s]+),NDY111Z=([^,;\s]+)/);
  if (match) {
    const price = toNumber_(match[1]);
    const changePoints = toNumber_(match[2]);
    const prevClose = price - changePoints;
    const changePct = prevClose !== 0 ? (changePoints / prevClose) * 100 : 0;
    return { price, changePoints, changePct, prevClose, marketState: '' };
  }

  const price = parseJsNumberVar_(nk225Text, 'N225kabuka');
  if (price <= 0) return null;
  return {
    price,
    changePoints: null,
    changePct: 0,
    prevClose: null,
    marketState: 'estimated',
  };
}

function parseJsArrayEntries_(text, varName) {
  const re = new RegExp(varName + '\\[(\\d+)\\]="([^"]+)"', 'g');
  const entries = [];
  let match;
  while ((match = re.exec(text)) !== null) {
    entries.push({ index: parseInt(match[1], 10), value: match[2] });
  }
  entries.sort((a, b) => a.index - b.index);
  return entries.map((entry) => entry.value);
}

function parseJsQuotedVar_(text, varName) {
  const match = text.match(new RegExp(varName + '="([^"]*)"'));
  return match ? match[1] : '';
}

function parseJsNumberVar_(text, varName) {
  const match =
    text.match(new RegExp('var ' + varName + '=([^;\\s]+)')) ||
    text.match(new RegExp(varName + '=([^;\\s]+)'));
  return match ? toNumber_(match[1]) : 0;
}

function buildUpdatedAt_(lastTime) {
  const dateStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const time = (lastTime || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm'))
    .replace(/"/g, '')
    .trim();
  return dateStr + ' ' + time + ':00 JST';
}

function toNumber_(value) {
  if (value === '' || value == null) return 0;
  if (typeof value === 'string' && value.charAt(0) === '#') return 0;
  const normalized = String(value).replace(/,/g, '').replace(/\+/g, '');
  const n = Number(normalized);
  return Number.isNaN(n) ? 0 : n;
}

// ---------------------------------------------------------------------------
// Sheet output
// ---------------------------------------------------------------------------

function padRows_(rows, numCols) {
  return rows.map((row) => {
    const padded = row.slice();
    while (padded.length < numCols) {
      padded.push('');
    }
    return padded.slice(0, numCols);
  });
}

function writeSummarySheet_(ss, rows, updatedAt, nikkei, totalContribYen, topUp, topDown) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_SUMMARY);
  sheet.clear();
  const numCols = 7;

  const nikkeiLine = nikkei
    ? formatNikkeiLine_(nikkei)
    : '日経225: 取得失敗';

  const header = padRows_([
    ['日経225 寄与度 Top 10'],
    [`更新: ${updatedAt}`],
    [nikkeiLine, '', '', `合計寄与: ${formatContribYen_(totalContribYen)}`],
    [],
    ['【上昇寄与 Top 10】'],
    ['順位', 'コード', '会社名', '構成率', '騰落率', '寄与度(円)', '現在値'],
  ], numCols);

  const upRows = (topUp || [])
    .slice(0, 10)
    .map((r, i) => [
      i + 1,
      r.code,
      r.company,
      r.weightPct / 100,
      r.pctChg / 100,
      r.contributionYen,
      r.price,
    ]);

  const downHeader = padRows_([
    [],
    ['【下落寄与 Top 10】'],
    ['順位', 'コード', '会社名', '構成率', '騰落率', '寄与度(円)', '現在値'],
  ], numCols);
  const downRows = (topDown || [])
    .slice(0, 10)
    .map((r, i) => [
      i + 1,
      r.code,
      r.company,
      r.weightPct / 100,
      r.pctChg / 100,
      r.contributionYen,
      r.price,
    ]);

  const values = padRows_([...header, ...upRows, ...downHeader, ...downRows], numCols);
  sheet.getRange(1, 1, values.length, numCols).setValues(values);

  sheet.getRange('A1').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A5').setFontWeight('bold');
  sheet.getRange('A' + (8 + upRows.length)).setFontWeight('bold');
  formatSummaryTable_(sheet, 6, upRows.length);
  formatSummaryTable_(sheet, 9 + upRows.length, downRows.length);
  [40, 70, 260, 80, 80, 80, 90].forEach((width, i) => {
    sheet.setColumnWidth(i + 1, width);
  });
  sheet.setFrozenRows(1);
}

function formatSummaryTable_(sheet, startRow, rowCount) {
  if (rowCount === 0) return;
  sheet.getRange(startRow, 4, rowCount, 2).setNumberFormat('0.00%');
  sheet.getRange(startRow, 6, rowCount, 1).setNumberFormat('#,##0.00"円"');
  sheet.getRange(startRow, 7, rowCount, 1).setNumberFormat('¥#,##0');
}

function writeAllSheet_(ss, rows, updatedAt, nikkei, totalContribYen) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_ALL);
  sheet.clear();
  const numCols = 8;

  const header = padRows_([
    ['更新', updatedAt],
    ['日経225', nikkei ? nikkei.price : ''],
    ['日経225騰落率', nikkei ? nikkei.changePct / 100 : ''],
    ['合計寄与(円)', totalContribYen],
    [],
    ['順位', 'コード', '会社名', '構成率', '現在値', '前日終値', '騰落率', '寄与度(円)'],
  ], numCols);

  const body = rows.map((r, i) => [
    i + 1,
    r.code,
    r.company,
    r.weightPct ? r.weightPct / 100 : '',
    r.price,
    r.prevClose,
    r.pctChg !== '' ? r.pctChg / 100 : '',
    r.contributionYen !== '' ? r.contributionYen : '',
  ]);

  const values = padRows_([...header, ...body], numCols);
  sheet.getRange(1, 1, values.length, numCols).setValues(values);
  sheet.getRange('D7:D').setNumberFormat('0.00%');
  sheet.getRange('G7:G').setNumberFormat('0.00%');
  sheet.getRange('H7:H').setNumberFormat('#,##0.00"円"');
  sheet.getRange('E7:F').setNumberFormat('¥#,##0');
  sheet.setFrozenRows(6);
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  const id =
    PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') ||
    CONFIG.SPREADSHEET_ID;
  if (!id) {
    throw new Error('SPREADSHEET_ID が未設定です。');
  }
  return SpreadsheetApp.openById(id);
}

function isJpTradingDay_(date) {
  const tz = 'Asia/Tokyo';
  const dow = parseInt(Utilities.formatDate(date, tz, 'u'), 10);
  if (dow >= 6) return false;
  const dateStr = Utilities.formatDate(date, tz, 'yyyy-MM-dd');
  return !JP_MARKET_HOLIDAYS.includes(dateStr);
}

function ensureSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function setupSheetsQuiet_(ss) {
  ensureSheet_(ss, CONFIG.SHEET_SUMMARY);
  ensureSheet_(ss, CONFIG.SHEET_ALL);
}

function safeUiAlert_(message) {
  Logger.log(message);
}

function formatPct_(value) {
  const sign = value >= 0 ? '+' : '';
  return sign + value.toFixed(2) + '%';
}
