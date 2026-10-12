/**
 * 請求書 自動生成エンジン
 *
 * 何をする:
 *   「取引先マスタ」+「明細」シートから、対象月の請求を「取引先＋請求単位」ごとに
 *   集約し（請求単位が空欄なら月まとめ、指定すれば案件ごとに別の請求書）、
 *   適格請求書(インボイス)の要件を満たす請求書PDFを生成して Drive に保存、
 *   「発行台帳」に番号・金額・PDFリンクを記録する。二重発行はしない。
 *
 * 追加の3機能（2026-08-08 追加。CONFIG.OPTIONS で個別に切れる）:
 *   (1) 決まった日の自動発行 … 時間主導トリガー(dailyIssueCheck)で毎日判定し、
 *       「自動発行設定」シートの発行日に当たる日だけ発行する。随時発行と併用できる。
 *   (2) 実行ログ … 実行日時・実行種別・対象月・件数・成否を「実行ログ」シートに追記する。
 *   (3) 失敗時のエラー通知 … 自動発行が失敗したときだけ、実行者本人（＝自動発行を
 *       設定した利用者本人）にメールで知らせる。宛先は Session.getEffectiveUser()
 *       に固定で、宛先を設定する欄も引数も持たない。請求書そのものを取引先へ送る
 *       機能はこのエンジンには無い（宛先・金額の誤りが社外にそのまま届かないようにするため）。
 *
 * デモモード:
 *   DEMO_MODE=true では、サンプルの取引先・明細を
 *   自動生成して請求書PDFを作る（実データ不要で動作を見せる）。実運用では
 *   シート様式・発行者情報・テンプレートを利用者に合わせて調整する前提。
 *
 * 法令(一次確認 2026-07-13 / 国税庁 No.6625):
 *   適格請求書の記載事項6つ =
 *     (1)発行事業者名+登録番号(T+13桁) (2)取引年月日 (3)取引内容(軽減対象の明示)
 *     (4)税率ごとの対価合計+適用税率 (5)税率ごとの消費税額 (6)宛名(交付先名)
 *   端数処理は「1請求書につき税率ごとに1回」— 行ごとに丸めて合算しない。
 */

// ==== 設定（本番はここを差し替える）====
var CONFIG = {
  DEMO_MODE: true,
  MASTER_SHEET: '取引先マスタ',
  ITEMS_SHEET: '明細',
  LEDGER_SHEET: '発行台帳',
  SCHEDULE_SHEET: '自動発行設定',         // 決まった日の自動発行の設定（シート上で完結）
  RUN_LOG_SHEET: '実行ログ',              // 実行日時・件数・成否の記録
  OUTPUT_FOLDER_NAME: '請求書出力(デモ)', // Drive内。無ければ作る
  TARGET_MONTH: '2026-07',                // 対象請求月 YYYY-MM
  ROUNDING: 'floor',                      // 'floor' | 'round' | 'ceil'（消費税の端数処理）
  PAYMENT_DUE: 'endOfNextMonth',          // 'endOfNextMonth' | 'issuePlus30'
  INVOICE_NO_FORMAT: 'INV-{YYYYMM}-{SEQ3}', // {SEQ3}=通番3桁ゼロ埋め
  TIMEZONE: 'Asia/Tokyo',
  ISSUER: {                               // 発行者(自社)情報
    name: 'Mxhlix（サンプル発行者）',
    registrationNo: 'T1234567890123',     // 適格請求書発行事業者 登録番号 T+13桁
    address: '東京都〇〇区〇〇 1-2-3',
    tel: '000-0000-0000',
    bank: '〇〇銀行 〇〇支店 普通 0000000 名義：Mxhlix'
  },
  OPTIONS: {                              // 追加の3機能。不要なものは false にする
    SCHEDULED_ISSUE: true,                // 決まった日の自動発行
    RUN_LOG: true,                        // 実行ログの記録
    ERROR_MAIL: true                      // 自動発行が失敗したときの通知（宛先は実行者本人で固定）
  }
};

var TAX_REDUCED = 0.08; // 軽減税率
var TAX_STANDARD = 0.10;

/**
 * メイン（随時発行）: 対象月(CONFIG.TARGET_MONTH)の請求書をまとめて生成する。
 * メニュー「請求書」→「請求書を生成」／「明細」シート上のボタンから呼ぶ入口で、
 * 引数も戻り値も無い従来どおりの形。生成結果（PDF・発行台帳・二重発行防止・Logger出力）は変えていない。
 * 加わったのは実行ログの追記だけ（CONFIG.OPTIONS.RUN_LOG が true のとき）。
 * 失敗はそのまま画面に出す（随時発行は人が見ているのでメール通知はしない）。
 */
function generateInvoices() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var month = CONFIG.TARGET_MONTH;
  try {
    var result = issueInvoices_(ss, month);
    logRun_(ss, RUN_KIND_MANUAL, month, result.made, true, '');
  } catch (e) {
    logRun_(ss, RUN_KIND_MANUAL, month, 0, false, errText_(e));
    throw e;
  }
}

/**
 * 発行の本体。随時発行(generateInvoices)と自動発行(dailyIssueCheck)の両方から呼ぶ。
 * 対象月だけを引数で受け取り、それ以外の挙動は従来の generateInvoices と同じ。
 * 戻り: {month, made} … made は今回生成した請求書の件数
 */
function issueInvoices_(ss, targetMonth) {
  if (CONFIG.DEMO_MODE) ensureDemoData_(ss);
  // 生成ボタンはDEMO_MODEに関わらず配置する（本番でも「ボタンかメニューをクリックするだけ」
  // で実行できるようにするため）。ensureGenerateButton_は冪等（既にあれば作らない）なので毎回呼んでも安全。
  ensureGenerateButton_(getOrCreateSheet_(ss, CONFIG.ITEMS_SHEET, ITEMS_HEADERS));

  var master = readMaster_(ss);           // 取引先コード -> マスタ情報
  var groups = readItemsForMonth_(ss, targetMonth); // "取引先コード|請求単位" -> {code, unit, lines[]}
  var ledger = getOrCreateSheet_(ss, CONFIG.LEDGER_SHEET, LEDGER_HEADERS);
  var issued = issuedSet_(ledger);        // 既発行(取引先コード|対象月|請求単位)
  var folder = getOrCreateFolder_(CONFIG.OUTPUT_FOLDER_NAME);

  var made = 0;
  Object.keys(groups).forEach(function (groupKey) {
    var g = groups[groupKey];
    var code = g.code, unit = g.unit;
    var issuedKey = code + '|' + targetMonth + '|' + unit;
    if (issued[issuedKey]) return; // 二重発行しない
    var client = master[code];
    if (!client) { Logger.log('マスタに無い取引先コード: %s（スキップ）', code); return; }

    var lines = g.lines;
    var totals = computeTotals_(lines, CONFIG.ROUNDING);
    var invoiceNo = nextInvoiceNo_(ledger, targetMonth);
    var issueDate = fmtDate_(new Date());
    var dueDate = paymentDue_(new Date(), client.paymentDue);

    var html = buildInvoiceHtml_({
      issuer: CONFIG.ISSUER, client: client, lines: lines, totals: totals,
      invoiceNo: invoiceNo, issueDate: issueDate, dueDate: dueDate,
      targetMonth: targetMonth, unit: unit
    });
    var fileName = invoiceNo + '_' + client.name + (unit ? '_' + unit : '');
    var pdf = Utilities.newBlob(html, MimeType.HTML, fileName + '.html')
      .getAs(MimeType.PDF).setName(fileName + '.pdf');
    var file = folder.createFile(pdf);

    ledger.appendRow([invoiceNo, issueDate, targetMonth, unit, code, client.name,
      totals.subtotal, totals.taxTotal, totals.total, dueDate, file.getUrl()]);
    made++;
  });
  Logger.log('生成完了: %s 件の請求書PDF（DEMO_MODE=%s, 対象月=%s）', made, CONFIG.DEMO_MODE, targetMonth);
  return { month: targetMonth, made: made };
}

// ==== 決まった日の自動発行 ====
var SCHEDULE_HEADERS = ['項目', '設定値', '説明'];
// 「自動発行設定」シートの初期値。利用者はこのシートの「設定値」列だけを書き換える（コードは触らない）
var SCHEDULE_DEFAULTS = [
  ['自動発行', 'ON', 'ON で自動発行します。止めるときは OFF に書き換えてください'],
  ['発行日', 25, '毎月この日に発行します。「月末」と書くと月末に発行します。31 など無い月は末日になります'],
  ['対象月', '前月', '「前月」か「当月」。前月なら 8月25日に 7月分を発行します'],
  ['実行時刻', 9, '0〜23。判定を行う時刻です。変えたらメニュー「請求書」→「自動発行を有効にする」を実行し直してください']
];

/**
 * 時間主導トリガーから毎日呼ばれる入口。「自動発行設定」シートの発行日に当たる日だけ発行する。
 * 随時発行(generateInvoices)とは独立して動くので併用できる。同じ請求書が二重に出ることは
 * 発行台帳による二重発行防止で起きない。
 */
function dailyIssueCheck() {
  if (!CONFIG.OPTIONS.SCHEDULED_ISSUE) { Logger.log('自動発行は無効設定です（CONFIG.OPTIONS.SCHEDULED_ISSUE=false）'); return; }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var setting = readScheduleSettings_(ss);
  var today = new Date();

  if (!setting.enabled) { Logger.log('自動発行が OFF のため実行しません（%s）', fmtDate_(today)); return; }
  if (!isIssueDay_(today, setting.issueDay)) {
    Logger.log('本日は発行日ではありません（発行日の設定: %s / 本日: %s）', setting.issueDay, fmtDate_(today));
    return;
  }

  var month = scheduledTargetMonth_(today, setting.monthMode);
  try {
    var result = issueInvoices_(ss, month);
    logRun_(ss, RUN_KIND_AUTO, month, result.made, true, '');
  } catch (e) {
    logRun_(ss, RUN_KIND_AUTO, month, 0, false, errText_(e));
    notifyFailure_(month, e);
  }
}

/**
 * 自動発行のトリガーを作る（同じ関数の既存トリガーは作り直す）。
 * トリガーは作成した本人のアカウントで動くため、利用者本人が実行する操作。
 */
function createDailyTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var setting = readScheduleSettings_(ss);
  removeDailyTriggers_();
  ScriptApp.newTrigger('dailyIssueCheck').timeBased().everyDays(1).atHour(setting.hour).create();
  Logger.log('自動発行を有効にしました（毎日 %s 時台に判定 / 発行日: %s / 対象月: %s）',
    setting.hour, setting.issueDay, setting.monthMode);
}

/** 自動発行を止める（トリガーだけ削除する。設定シートはそのまま残す）。 */
function disableDailyTrigger() {
  Logger.log('自動発行を止めました（削除したトリガー: %s 件）', removeDailyTriggers_());
}

function removeDailyTriggers_() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyIssueCheck') { ScriptApp.deleteTrigger(t); removed++; }
  });
  return removed;
}

/** 「自動発行設定」シートを読む。無ければ初期値で作る。不正値は既定にフォールバックする。 */
function readScheduleSettings_(ss) {
  var sheet = ensureScheduleSheet_(ss);
  var map = {};
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function (r) {
      var key = String(r[0]).trim();
      if (key) map[key] = r[1];
    });
  }
  var onoff = String(map['自動発行'] == null ? 'ON' : map['自動発行']).trim().toUpperCase();
  var day = (map['発行日'] === '' || map['発行日'] == null) ? 25 : map['発行日'];
  var monthMode = String(map['対象月'] == null ? '前月' : map['対象月']).trim();
  var hour = Number(map['実行時刻']);
  return {
    enabled: (onoff === 'ON' || onoff === 'オン' || onoff === 'TRUE'),
    issueDay: day,
    monthMode: (monthMode === '当月') ? '当月' : '前月',
    hour: (hour >= 0 && hour <= 23) ? Math.floor(hour) : 9
  };
}

function ensureScheduleSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, CONFIG.SCHEDULE_SHEET, SCHEDULE_HEADERS);
  if (sheet.getLastRow() < 2) {
    sheet.getRange(2, 1, SCHEDULE_DEFAULTS.length, SCHEDULE_HEADERS.length).setValues(SCHEDULE_DEFAULTS);
  }
  return sheet;
}

/** 今日が発行日かどうか。「月末」指定と、31日など無い月（末日に寄せる）に対応する。 */
function isIssueDay_(today, issueDay) {
  var lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  var setting = String(issueDay).trim();
  if (setting === '月末') return today.getDate() === lastDay;
  var n = Number(setting);
  if (!n || n < 1 || n > 31) {
    Logger.log('発行日の設定が不正です: %s（自動発行を見送ります）', setting);
    return false;
  }
  return today.getDate() === Math.min(n, lastDay);
}

/** 自動発行の対象月を決める。'前月'（既定）なら 8/25 の実行で 2026-07 を対象にする。 */
function scheduledTargetMonth_(today, monthMode) {
  var d = new Date(today.getFullYear(), today.getMonth(), 1);
  if (monthMode !== '当月') d.setMonth(d.getMonth() - 1);
  return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM');
}

// ==== 実行ログ ====
var RUN_LOG_HEADERS = ['実行日時', '実行種別', '対象月', '対象件数', '結果', 'メッセージ'];
var RUN_KIND_MANUAL = '随時';
var RUN_KIND_AUTO = '自動';

/** 実行の記録を「実行ログ」シートに1行追記する。ログの失敗で発行本体を止めない。 */
function logRun_(ss, kind, month, count, ok, message) {
  if (!CONFIG.OPTIONS.RUN_LOG) return;
  try {
    var sheet = getOrCreateSheet_(ss, CONFIG.RUN_LOG_SHEET, RUN_LOG_HEADERS);
    sheet.appendRow([fmtDateTime_(new Date()), kind, month, count, ok ? '成功' : '失敗', message || '']);
  } catch (e) {
    Logger.log('実行ログの記録に失敗しました: %s', errText_(e));
  }
}

// ==== 失敗時のエラー通知 ====
/**
 * 自動発行が失敗したときだけ、実行者本人へメールで知らせる。
 * 宛先は Session.getEffectiveUser()（＝自動発行のトリガーを設定した利用者本人）に固定。
 * 宛先を設定する欄も引数も設けない。取引先など社外へ送る経路をこのエンジンは持たない
 * （宛先や金額の誤りがそのまま社外に届く形にしないため）。
 */
function notifyFailure_(month, err) {
  if (!CONFIG.OPTIONS.ERROR_MAIL) return false;
  var to = Session.getEffectiveUser().getEmail();
  if (!to) { Logger.log('通知の宛先（実行者のメールアドレス）を取得できませんでした'); return false; }
  var body = [
    '請求書の自動発行が失敗しました。',
    '',
    '対象月: ' + month,
    '発生日時: ' + fmtDateTime_(new Date()),
    'エラー内容: ' + errText_(err),
    '',
    '同じ内容が「実行ログ」シートにも残っています。',
    '原因を取り除いたあと、メニュー「請求書」→「請求書を生成」で発行し直せます',
    '（発行済みの分は二重に発行されません）。'
  ].join('\n');
  try {
    MailApp.sendEmail({
      to: to,
      subject: '[請求書自動生成] 自動発行に失敗しました（' + month + '）',
      body: body
    });
  } catch (e) {
    Logger.log('エラー通知メールの送信に失敗しました: %s', errText_(e));
    return false;
  }
  Logger.log('エラー通知を送信しました: %s', to);
  return true;
}

// ==== UI（メニュー・シート上ボタン）====
/**
 * シンプルトリガー: スプレッドシートを開いたときにカスタムメニューを追加する。
 * シンプルトリガーはスクリプト未認可の状態でも動く（メニュー表示のみ）。
 * メニューから generateInvoices を初めて実行するときに認可を求められる。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('請求書')
    .addItem('請求書を生成', 'generateInvoices')
    .addSeparator()
    .addItem('自動発行を有効にする', 'createDailyTrigger')
    .addItem('自動発行を止める', 'disableDailyTrigger')
    .addToUi();
}

/**
 * 「明細」シートの右上にボタン画像（base64埋め込みPNG）を配置し、
 * クリックで generateInvoices を実行できるようにする。
 * 既に同じボタン（generateInvoices が割り当て済みの画像）があれば二重に挿入しない。
 */
function ensureGenerateButton_(sheet) {
  var already = sheet.getImages().some(function (img) {
    return img.getScript() === 'generateInvoices';
  });
  if (already) return;
  var bytes = Utilities.base64Decode(GEN_BUTTON_PNG_B64);
  var blob = Utilities.newBlob(bytes, MimeType.PNG, '請求書を生成ボタン.png');
  var image = sheet.insertImage(blob, ITEMS_HEADERS.length + 2, 1);
  image.setWidth(160).setHeight(36);
  image.assignScript('generateInvoices');
}

// ボタン画像（320x72pxのRetina想定PNG、#E5880C角丸背景+白太字「請求書を生成」）のbase64。
// シート上ではensureGenerateButton_内でsetWidth/setHeightにより160x36へ縮小して配置する。
var GEN_BUTTON_PNG_B64 =
  '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';

// ==== 計算（コア。テストで検証する純関数）====
/**
 * 明細行から、税率ごとに区分した対価・消費税・合計を計算する。
 * 端数処理は「税率ごとに1回」= 各税率の対価合計に対して1回だけ丸める（国税庁ルール）。
 * lines: [{name, qty, unitPrice(税抜), taxRate}]
 * 戻り: {byRate:{rate:{net,tax}}, subtotal, taxTotal, total}
 */
function computeTotals_(lines, rounding) {
  var byRate = {};
  lines.forEach(function (l) {
    var rate = Number(l.taxRate);
    var net = Number(l.qty) * Number(l.unitPrice);
    if (!byRate[rate]) byRate[rate] = { net: 0, tax: 0 };
    byRate[rate].net += net;
  });
  var subtotal = 0, taxTotal = 0;
  Object.keys(byRate).forEach(function (rate) {
    var net = byRate[rate].net;
    var tax = roundAmount_(net * Number(rate), rounding); // 税率ごとに1回だけ丸める
    byRate[rate].tax = tax;
    subtotal += net;
    taxTotal += tax;
  });
  return { byRate: byRate, subtotal: subtotal, taxTotal: taxTotal, total: subtotal + taxTotal };
}

function roundAmount_(v, rounding) {
  if (rounding === 'round') return Math.round(v);
  if (rounding === 'ceil') return Math.ceil(v);
  return Math.floor(v); // 既定: 切り捨て
}

// ==== 請求書HTML（適格請求書6要件を満たす）====
function buildInvoiceHtml_(d) {
  var rateRows = Object.keys(d.totals.byRate).sort().map(function (rate) {
    var r = d.totals.byRate[rate];
    var label = (Number(rate) === TAX_REDUCED ? '軽減8%' : (Math.round(Number(rate) * 100) + '%'));
    return '<tr><td>' + label + ' 対象</td><td class="num">' + yen_(r.net) +
      '</td><td class="num">' + yen_(r.tax) + '</td></tr>';
  }).join('');

  var showDate = !!d.unit; // 請求単位が指定されたグループ(案件ごと請求)は明細に日付を出す
  var lineRows = d.lines.map(function (l) {
    var mark = (Number(l.taxRate) === TAX_REDUCED ? ' ※' : '');
    var amt = Number(l.qty) * Number(l.unitPrice);
    var dateCell = showDate ? '<td>' + esc_(l.date || '') + '</td>' : '';
    return '<tr>' + dateCell + '<td>' + esc_(l.name) + mark + '</td><td class="num">' + l.qty +
      '</td><td class="num">' + yen_(l.unitPrice) + '</td><td class="num">' + yen_(amt) + '</td></tr>';
  }).join('');

  return '' +
    '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:sans-serif;font-size:12px;color:#111;margin:32px;}' +
    'h1{font-size:22px;letter-spacing:6px;text-align:center;margin:0 0 18px;}' +
    'table{width:100%;border-collapse:collapse;margin-top:8px;}' +
    'th,td{border:1px solid #999;padding:6px 8px;}th{background:#f0f0f0;}' +
    '.num{text-align:right;}.right{text-align:right;}.muted{color:#555;}' +
    '.box{border:1px solid #999;padding:8px 10px;margin-top:6px;}' +
    '.total{font-size:16px;font-weight:bold;}' +
    '</style></head><body>' +
    '<h1>請 求 書</h1>' +
    '<div class="right muted">請求書番号: ' + esc_(d.invoiceNo) + '<br>発行日: ' + d.issueDate + '</div>' +
    '<div style="font-size:15px;font-weight:bold;">' + esc_(d.client.name) + ' 御中</div>' +
    '<div class="muted">' + esc_(d.client.address || '') + '</div>' +
    '<p>下記のとおりご請求申し上げます。（対象: ' + d.targetMonth + (d.unit ? '（' + esc_(d.unit) + '）' : '') + '）</p>' +
    '<div class="box total">ご請求金額（税込）: ' + yen_(d.totals.total) + ' 円' +
    '<span class="muted" style="font-weight:normal;font-size:12px;">　支払期限: ' + d.dueDate + '</span></div>' +
    '<table><tr>' + (showDate ? '<th>日付</th>' : '') + '<th>品目</th><th>数量</th><th>単価(税抜)</th><th>金額(税抜)</th></tr>' + lineRows + '</table>' +
    '<table style="width:60%;margin-top:12px;"><tr><th>税率区分</th><th>対価の額(税抜)</th><th>消費税額</th></tr>' +
    rateRows +
    '<tr><td class="right">合計</td><td class="num">' + yen_(d.totals.subtotal) +
    '</td><td class="num">' + yen_(d.totals.taxTotal) + '</td></tr></table>' +
    '<p class="muted">※ は軽減税率(8%)対象品目です。</p>' +
    '<div class="box"><b>お振込先</b><br>' + esc_(d.issuer.bank) + '</div>' +
    '<div style="margin-top:14px;text-align:right;">' +
    '<b>' + esc_(d.issuer.name) + '</b>　登録番号: ' + esc_(d.issuer.registrationNo) + '<br>' +
    '<span class="muted">' + esc_(d.issuer.address) + '　TEL: ' + esc_(d.issuer.tel) + '</span></div>' +
    '</body></html>';
}

// ==== シート読み込み ====
var MASTER_HEADERS = ['取引先コード', '取引先名', '住所', '登録番号', '振込先', '支払期限'];
var ITEMS_HEADERS = ['対象月', '取引先コード', '請求単位', '日付', '品目', '数量', '単価(税抜)', '税率'];
var LEDGER_HEADERS = ['請求書番号', '発行日', '対象月', '請求単位', '取引先コード', '取引先名',
  '小計(税抜)', '消費税', '合計(税込)', '支払期限', 'PDF'];

function readMaster_(ss) {
  var rows = readRows_(ss, CONFIG.MASTER_SHEET);
  var map = {};
  rows.forEach(function (r) {
    map[String(r['取引先コード'])] = {
      code: String(r['取引先コード']), name: r['取引先名'],
      address: r['住所'], registrationNo: r['登録番号'], bank: r['振込先'],
      paymentDue: r['支払期限'] // 空欄可。「翌月末」「発行日+N」。不正/空欄はCONFIG既定にフォールバック
    };
  });
  return map;
}

/**
 * 対象月の明細を「取引先コード＋請求単位」でグループ化する。
 * 請求単位が空欄のグループ = 月まとめ請求。同じ取引先でも請求単位が違えば別グループ(別PDF)。
 * 戻り: { "取引先コード|請求単位": {code, unit, lines:[{name,qty,unitPrice,taxRate,date}]} }
 */
function readItemsForMonth_(ss, month) {
  var rows = readRows_(ss, CONFIG.ITEMS_SHEET);
  var groups = {};
  rows.forEach(function (r) {
    if (String(r['対象月']) !== month) return;
    var code = String(r['取引先コード']);
    var unit = r['請求単位'] ? String(r['請求単位']) : '';
    var key = code + '|' + unit;
    if (!groups[key]) groups[key] = { code: code, unit: unit, lines: [] };
    groups[key].lines.push({
      name: r['品目'], qty: Number(r['数量']),
      unitPrice: Number(r['単価(税抜)']), taxRate: Number(r['税率']),
      date: fmtItemDate_(r['日付'])
    });
  });
  return groups;
}

function readRows_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getDataRange().getValues();
  var head = values[0];
  return values.slice(1).map(function (row) {
    var o = {};
    head.forEach(function (h, i) { o[h] = row[i]; });
    return o;
  });
}

// ==== 発行台帳・採番 ====
// 台帳の列順: 0請求書番号 1発行日 2対象月 3請求単位 4取引先コード 5取引先名 ...
function issuedSet_(ledger) {
  var set = {};
  if (ledger.getLastRow() < 2) return set;
  var vals = ledger.getRange(2, 1, ledger.getLastRow() - 1, LEDGER_HEADERS.length).getValues();
  vals.forEach(function (r) { set[String(r[4]) + '|' + String(r[2]) + '|' + String(r[3])] = true; });
  return set;
}

function nextInvoiceNo_(ledger, month) {
  var seq = 1;
  if (ledger.getLastRow() >= 2) {
    var vals = ledger.getRange(2, 3, ledger.getLastRow() - 1, 1).getValues(); // 対象月列
    vals.forEach(function (r) { if (String(r[0]) === month) seq++; });
  }
  var ym = month.replace('-', '');
  return CONFIG.INVOICE_NO_FORMAT
    .replace('{YYYYMM}', ym)
    .replace('{SEQ3}', String(seq).padStart(3, '0')); // 1000件目以降は4桁になる（切り詰めて番号が重複しないように）
}

// ==== ユーティリティ ====
/**
 * 支払期限を計算する。取引先マスタの「支払期限」指定を優先し、空欄・不正値は
 * CONFIG.PAYMENT_DUE の既定にフォールバックする(不正値はLoggerで警告)。
 * clientDue: '翌月末' | '発行日+N'(例: '発行日+30') | '' | 不正値
 */
function paymentDue_(issue, clientDue) {
  var setting = clientDue ? String(clientDue).trim() : '';
  if (!setting) return paymentDueByRule_(issue, CONFIG.PAYMENT_DUE);
  if (setting === '翌月末') return paymentDueByRule_(issue, 'endOfNextMonth');
  var m = setting.match(/^発行日\+(\d+)$/);
  if (m) {
    var d = new Date(issue); d.setDate(d.getDate() + Number(m[1])); return fmtDate_(d);
  }
  Logger.log('取引先の支払期限指定が不正です: %s（既定にフォールバック）', setting);
  return paymentDueByRule_(issue, CONFIG.PAYMENT_DUE);
}

// CONFIG.PAYMENT_DUE の既定ルール('endOfNextMonth' | 'issuePlus30')を適用する
function paymentDueByRule_(issue, rule) {
  if (rule === 'issuePlus30') {
    var d = new Date(issue); d.setDate(d.getDate() + 30); return fmtDate_(d);
  }
  // 既定: 翌月末
  var e = new Date(issue.getFullYear(), issue.getMonth() + 2, 0);
  return fmtDate_(e);
}
function getOrCreateFolder_(name) {
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
function getOrCreateSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function fmtDate_(d) { return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd'); }
function fmtDateTime_(d) { return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'); }
// 例外オブジェクト・文字列のどちらでもログとメールに載る形にする
function errText_(e) { return (e && e.message) ? String(e.message) : String(e); }
// 明細シートの「日付」セルはDate型/文字列/空欄のいずれもあり得るため吸収する
function fmtItemDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return fmtDate_(v);
  return String(v);
}
function yen_(n) { return Math.round(Number(n)).toLocaleString('ja-JP'); }
function esc_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ==== デモ用サンプルデータ ====
function ensureDemoData_(ss) {
  var master = getOrCreateSheet_(ss, CONFIG.MASTER_SHEET, MASTER_HEADERS);
  if (master.getLastRow() < 2) {
    // C001=月まとめ請求(支払期限は空欄→既定にフォールバック)、C002=案件ごと請求(支払期限は取引先ごと指定の例)
    master.getRange(2, 1, 2, MASTER_HEADERS.length).setValues([
      ['C001', 'サンプル商事株式会社', '東京都千代田区1-1-1', 'T9010001000001', 'サンプル銀行 本店 普通 1234567', ''],
      ['C002', '合同会社デモワークス', '大阪府大阪市2-2-2', 'T5120002000002', 'デモ信用金庫 支店 普通 7654321', '発行日+30']
    ]);
  }
  var items = getOrCreateSheet_(ss, CONFIG.ITEMS_SHEET, ITEMS_HEADERS);
  if (items.getLastRow() < 2) {
    // C001: 請求単位・日付とも空欄(月まとめ) / C002: 案件A・案件Bで請求単位を分けた随時発行の例
    items.getRange(2, 1, 5, ITEMS_HEADERS.length).setValues([
      ['2026-07', 'C001', '', '', 'Webサイト保守', 1, 30000, 0.10],
      ['2026-07', 'C001', '', '', 'バナー制作', 3, 5000, 0.10],
      ['2026-07', 'C001', '', '', 'ギフトコーヒー豆', 8, 1500, 0.08],
      ['2026-07', 'C002', '案件A', '2026-07-05', '業務自動化スクリプト開発', 1, 80000, 0.10],
      ['2026-07', 'C002', '案件B', '2026-07-20', '保守サポート（月額）', 1, 15000, 0.10]
    ]);
  }
  // 生成ボタンの配置は issueInvoices_ 側で DEMO_MODE によらず行う（ここでは呼ばない）
}
