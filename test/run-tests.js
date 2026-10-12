'use strict';
/**
 * invoice-gen の検証（随時発行 / 決まった日の自動発行 / 実行ログ / 失敗時のエラー通知）。
 * GASはローカルで実行できないため、GASグローバルをモックして Code.gs をそのまま動かす。
 *
 * 実行: node test/run-tests.js（Node標準のみ。依存パッケージなし）
 */
const assert = require('assert');
const { loadCode } = require('./load-code');

// ---- 小さなテストランナー ----
let passed = 0;
const failures = [];
function test(name, fn) {
  try {
    fn();
    passed++;
    console.log('  PASS  ' + name);
  } catch (e) {
    failures.push({ name, error: e });
    console.log('  FAIL  ' + name + '\n        ' + (e && e.message ? e.message : e));
  }
}

// ---- 補助 ----
const ms = (y, m, d, h, mi) => new Date(y, m - 1, d, h || 0, mi || 0, 0).getTime();

function seedSchedule(ss, rows) {
  const sheet = ss.insertSheet('自動発行設定');
  sheet.appendRow(['項目', '設定値', '説明']);
  rows.forEach(r => sheet.appendRow([r[0], r[1], '']));
  return sheet;
}
const SCHEDULE_25TH = [['自動発行', 'ON'], ['発行日', 25], ['対象月', '前月'], ['実行時刻', 9]];

function sheetRows(ss, name) {
  const s = ss.getSheetByName(name);
  return s ? s.rows() : null;
}
function fileCount(folders) {
  return Object.keys(folders).reduce((n, k) => n + folders[k].files.length, 0);
}
// 実行ログの列: 0実行日時 1実行種別 2対象月 3対象件数 4結果 5メッセージ
const LOG = { AT: 0, KIND: 1, MONTH: 2, COUNT: 3, RESULT: 4, MESSAGE: 5 };
// 発行台帳の列: 0請求書番号 1発行日 2対象月 3請求単位 4取引先コード 5取引先名 6小計 7消費税 8合計
const LED = { NO: 0, ISSUED_AT: 1, MONTH: 2, UNIT: 3, CODE: 4, NAME: 5, NET: 6, TAX: 7, TOTAL: 8 };

console.log('invoice-gen ハーネス（随時発行 + 自動発行・実行ログ・エラー通知）\n');

// ============================================================
// 1. 定期実行で発行日に該当 → 発行が走り、実行ログに成功が記録される
// ============================================================
test('1. 発行日に該当する日の定期実行 → 請求書が生成され、実行ログに「自動/成功」が記録される', () => {
  const env = loadCode({ nowMs: ms(2026, 8, 25, 9, 5) });
  seedSchedule(env.ss, SCHEDULE_25TH);

  env.context.dailyIssueCheck();

  // 発行が走った（デモデータの2026-07分＝C001月まとめ + C002案件A/B の3件）
  assert.strictEqual(fileCount(env.folders), 3, 'PDFが3件生成されること');
  const ledger = sheetRows(env.ss, '発行台帳');
  assert.strictEqual(ledger.length, 3, '発行台帳に3行記録されること');
  assert.strictEqual(ledger[0][LED.MONTH], '2026-07', '対象月が前月(2026-07)になること');
  assert.strictEqual(ledger[0][LED.ISSUED_AT], '2026-08-25', '発行日が実行日になること');

  // 実行ログに成功が1行
  const log = sheetRows(env.ss, '実行ログ');
  assert.strictEqual(log.length, 1, '実行ログが1行追記されること');
  assert.strictEqual(log[0][LOG.KIND], '自動');
  assert.strictEqual(log[0][LOG.MONTH], '2026-07');
  assert.strictEqual(log[0][LOG.COUNT], 3, '対象件数が記録されること');
  assert.strictEqual(log[0][LOG.RESULT], '成功');
  assert.strictEqual(log[0][LOG.AT], '2026-08-25 09:05:00', '実行日時が記録されること');

  // 成功時はメールを送らない
  assert.strictEqual(env.calls.mail.length, 0, '成功時はメールを送らないこと');
});

test('1b. 「月末」指定と、31日指定で31日が無い月（末日に寄せる）', () => {
  const endOfMonth = loadCode({ nowMs: ms(2026, 8, 31, 9, 0) });
  seedSchedule(endOfMonth.ss, [['自動発行', 'ON'], ['発行日', '月末'], ['対象月', '当月']]);
  endOfMonth.context.dailyIssueCheck();
  assert.strictEqual(sheetRows(endOfMonth.ss, '実行ログ').length, 1, '月末指定で8/31に実行されること');
  assert.strictEqual(sheetRows(endOfMonth.ss, '実行ログ')[0][LOG.MONTH], '2026-08', '当月指定が効くこと');

  const feb = loadCode({ nowMs: ms(2027, 2, 28, 9, 0) });
  seedSchedule(feb.ss, [['自動発行', 'ON'], ['発行日', 31], ['対象月', '前月']]);
  assert.strictEqual(feb.context.isIssueDay_(new Date(2027, 1, 28), 31), true,
    '31日指定は2月28日（末日）に発行されること');
});

// ============================================================
// 2. 非該当日 → 発行が走らない
// ============================================================
test('2. 発行日でない日の定期実行 → 発行も記録も走らない', () => {
  const env = loadCode({ nowMs: ms(2026, 8, 10, 9, 5) });
  seedSchedule(env.ss, SCHEDULE_25TH);

  env.context.dailyIssueCheck();

  assert.strictEqual(fileCount(env.folders), 0, 'PDFが生成されないこと');
  assert.strictEqual(env.ss.getSheetByName('発行台帳'), null, '発行台帳が作られないこと');
  assert.strictEqual(env.ss.getSheetByName('実行ログ'), null, '実行ログが作られないこと');
  assert.strictEqual(env.ss.getSheetByName('明細'), null, 'デモデータも作られないこと');
  assert.strictEqual(env.calls.mail.length, 0, 'メールを送らないこと');
});

test('2b. 設定が OFF なら発行日でも走らない', () => {
  const env = loadCode({ nowMs: ms(2026, 8, 25, 9, 5) });
  seedSchedule(env.ss, [['自動発行', 'OFF'], ['発行日', 25], ['対象月', '前月']]);

  env.context.dailyIssueCheck();

  assert.strictEqual(fileCount(env.folders), 0, 'PDFが生成されないこと');
  assert.strictEqual(env.ss.getSheetByName('実行ログ'), null, '実行ログが作られないこと');
  assert.strictEqual(env.calls.mail.length, 0, 'メールを送らないこと');
});

// ============================================================
// 3. 発行処理が例外 → ログに失敗 + 客自身宛のメール1通
// ============================================================
test('3. 自動発行が失敗 → 実行ログに「失敗」、実行者本人宛にメール1通だけ送られる', () => {
  const env = loadCode({ nowMs: ms(2026, 8, 25, 9, 5), failPdf: true, effectiveUserEmail: 'user@example.test' });
  seedSchedule(env.ss, SCHEDULE_25TH);

  env.context.dailyIssueCheck(); // 例外は内部で受け止める（トリガー実行なので投げ直さない）

  // ログに失敗が残る
  const log = sheetRows(env.ss, '実行ログ');
  assert.strictEqual(log.length, 1, '実行ログが1行追記されること');
  assert.strictEqual(log[0][LOG.KIND], '自動');
  assert.strictEqual(log[0][LOG.RESULT], '失敗');
  assert.ok(String(log[0][LOG.MESSAGE]).indexOf('PDF変換に失敗しました') >= 0,
    'エラー内容がログに残ること: ' + log[0][LOG.MESSAGE]);

  // 請求書は1件も出ていない
  assert.strictEqual(fileCount(env.folders), 0, 'PDFが生成されないこと');
  assert.strictEqual(sheetRows(env.ss, '発行台帳').length, 0, '発行台帳に記録されないこと');

  // メールはちょうど1通、宛先は Session.getEffectiveUser() と一致
  assert.strictEqual(env.calls.mail.length, 1, 'メールがちょうど1通であること');
  const mail = env.calls.mail[0];
  const sessionEmail = env.context.Session.getEffectiveUser().getEmail();
  assert.strictEqual(sessionEmail, 'user@example.test');
  assert.strictEqual(mail.to, sessionEmail, '宛先が Session.getEffectiveUser() と一致すること');
  assert.strictEqual(mail.to, env.effectiveUserEmail);
  assert.ok(String(mail.subject).indexOf('2026-07') >= 0, '件名に対象月が入ること: ' + mail.subject);
  assert.ok(String(mail.body).indexOf('PDF変換に失敗しました') >= 0, '本文にエラー内容が入ること');
});

test('3b. 通知の宛先はコードに固定。設定シート・CONFIG に宛先欄が存在しない', () => {
  const env = loadCode({ nowMs: ms(2026, 8, 25, 9, 5) });
  const cfg = JSON.stringify(env.context.CONFIG);
  assert.ok(cfg.indexOf('@') < 0, 'CONFIG にメールアドレスらしき値が無いこと');
  ['MAIL', 'NOTIFY_EMAIL', 'TO', 'MAIL_TO', 'RECIPIENT'].forEach(k => {
    assert.strictEqual(env.context.CONFIG[k], undefined, 'CONFIG.' + k + ' が存在しないこと');
  });
  // 設定シートの初期値に宛先の項目が無い（vmコンテキストの配列なので文字列で比較する）
  const keys = env.context.SCHEDULE_DEFAULTS.map(r => r[0]).join(',');
  assert.strictEqual(keys, '自動発行,発行日,対象月,実行時刻', '設定項目に宛先が無いこと');
  // notifyFailure_ の引数は (month, err) の2つだけ＝宛先を渡せない
  assert.strictEqual(env.context.notifyFailure_.length, 2, 'notifyFailure_ が宛先引数を取らないこと');
});

// ============================================================
// 4. 随時発行の既存経路が従来どおり動く
// ============================================================
test('4. 随時発行 generateInvoices() が従来どおり動く（3件生成・採番・税計算・二重発行防止）', () => {
  const env = loadCode({ nowMs: ms(2026, 7, 31, 14, 0) });

  const ret = env.context.generateInvoices();
  assert.strictEqual(ret, undefined, '戻り値は従来どおり undefined であること');

  assert.strictEqual(fileCount(env.folders), 3, 'PDFが3件生成されること');
  const ledger = sheetRows(env.ss, '発行台帳');
  assert.strictEqual(ledger.length, 3, '発行台帳に3行');
  assert.deepStrictEqual(ledger.map(r => r[LED.NO]),
    ['INV-202607-001', 'INV-202607-002', 'INV-202607-003'], '採番が従来どおりであること');

  // C001（月まとめ）: 10% 45,000 → 税4,500 ／ 軽減8% 12,000 → 税960
  const c001 = ledger.filter(r => r[LED.CODE] === 'C001')[0];
  assert.strictEqual(c001[LED.NET], 57000, '小計(税抜)');
  assert.strictEqual(c001[LED.TAX], 5460, '消費税（税率ごとに1回丸め）');
  assert.strictEqual(c001[LED.TOTAL], 62460, '合計(税込)');

  // 案件ごと請求（請求単位あり）が別々に出ている
  const units = ledger.filter(r => r[LED.CODE] === 'C002').map(r => r[LED.UNIT]).sort();
  assert.deepStrictEqual(units, ['案件A', '案件B'], '請求単位ごとに別の請求書が出ること');

  // 実行ログには「随時」の成功が1行
  const log = sheetRows(env.ss, '実行ログ');
  assert.strictEqual(log.length, 1);
  assert.strictEqual(log[0][LOG.KIND], '随時');
  assert.strictEqual(log[0][LOG.RESULT], '成功');
  assert.strictEqual(log[0][LOG.COUNT], 3);

  // 随時発行では自動発行の設定シートを作らない（既存経路に副作用を足さない）
  assert.strictEqual(env.ss.getSheetByName('自動発行設定'), null);
  assert.strictEqual(env.calls.mail.length, 0, '随時発行でメールを送らないこと');

  // 再実行しても二重発行しない
  env.context.generateInvoices();
  assert.strictEqual(fileCount(env.folders), 3, '再実行でPDFが増えないこと');
  assert.strictEqual(sheetRows(env.ss, '発行台帳').length, 3, '再実行で台帳が増えないこと');
  const log2 = sheetRows(env.ss, '実行ログ');
  assert.strictEqual(log2.length, 2, '実行ログは2行目が追記されること');
  assert.strictEqual(log2[1][LOG.COUNT], 0, '2回目の対象件数は0件');
});

test('4b. 自動発行と随時発行の併用（同じ月を続けて実行しても二重に出ない）', () => {
  const env = loadCode({ nowMs: ms(2026, 8, 25, 9, 5) });
  seedSchedule(env.ss, SCHEDULE_25TH);

  env.context.dailyIssueCheck();          // 自動で2026-07分を3件
  env.context.CONFIG.TARGET_MONTH = '2026-07';
  env.context.generateInvoices();         // 随時で同じ月を実行

  assert.strictEqual(fileCount(env.folders), 3, 'PDFは3件のまま');
  assert.strictEqual(sheetRows(env.ss, '発行台帳').length, 3, '台帳は3行のまま');
  const log = sheetRows(env.ss, '実行ログ');
  assert.deepStrictEqual(log.map(r => r[LOG.KIND]), ['自動', '随時'], '両方の実行が記録されること');
  assert.strictEqual(log[1][LOG.COUNT], 0, '随時側の対象件数は0件');
});

test('4c. 随時発行が失敗したときは例外がそのまま画面へ出る（メールは送らない）', () => {
  const env = loadCode({ nowMs: ms(2026, 7, 31, 14, 0), failPdf: true });

  assert.throws(() => env.context.generateInvoices(), /PDF変換に失敗しました/,
    '例外を握りつぶさず投げ直すこと');

  const log = sheetRows(env.ss, '実行ログ');
  assert.strictEqual(log.length, 1);
  assert.strictEqual(log[0][LOG.KIND], '随時');
  assert.strictEqual(log[0][LOG.RESULT], '失敗');
  assert.strictEqual(env.calls.mail.length, 0, '随時発行の失敗ではメールを送らないこと');
});

// ============================================================
// 5. トリガー設定（自動発行の有効化・停止）
// ============================================================
test('5. createDailyTrigger が dailyIssueCheck の日次トリガーを1つだけ作る', () => {
  const env = loadCode({ nowMs: ms(2026, 8, 1, 9, 0) });

  env.context.createDailyTrigger();
  assert.strictEqual(env.triggers.length, 1);
  assert.strictEqual(env.triggers[0].handler, 'dailyIssueCheck');
  assert.strictEqual(env.triggers[0].days, 1);
  assert.strictEqual(env.triggers[0].hour, 9, '設定シートの既定（9時）が使われること');

  // 設定シートが初期値つきで作られる
  const settings = sheetRows(env.ss, '自動発行設定');
  assert.strictEqual(settings.length, 4);
  assert.deepStrictEqual(settings.map(r => r[0]), ['自動発行', '発行日', '対象月', '実行時刻']);

  env.context.createDailyTrigger();
  assert.strictEqual(env.triggers.length, 1, '作り直しても増えないこと');

  env.context.disableDailyTrigger();
  assert.strictEqual(env.triggers.length, 0, '停止でトリガーが消えること');
});

test('5b. 実行時刻の設定がトリガーに反映される', () => {
  const env = loadCode({ nowMs: ms(2026, 8, 1, 9, 0) });
  seedSchedule(env.ss, [['自動発行', 'ON'], ['発行日', 25], ['対象月', '前月'], ['実行時刻', 7]]);
  env.context.createDailyTrigger();
  assert.strictEqual(env.triggers[0].hour, 7);
});

// ============================================================
// 6. 追加機能のスイッチ（CONFIG.OPTIONS）
// ============================================================
test('6. CONFIG.OPTIONS を切ると自動発行・実行ログが動かない（随時発行は動く）', () => {
  const env = loadCode({ nowMs: ms(2026, 8, 25, 9, 5) });
  seedSchedule(env.ss, SCHEDULE_25TH);
  env.context.CONFIG.OPTIONS.SCHEDULED_ISSUE = false;
  env.context.CONFIG.OPTIONS.RUN_LOG = false;

  env.context.dailyIssueCheck();
  assert.strictEqual(fileCount(env.folders), 0, '自動発行が走らないこと');

  env.context.generateInvoices();
  assert.strictEqual(fileCount(env.folders), 3, '随時発行は動くこと');
  assert.strictEqual(env.ss.getSheetByName('実行ログ'), null, 'ログOFFでログシートを作らないこと');
});

// ============================================================
// 7. 社外送信の経路が存在しない（実装レベルの確認）
// ============================================================
test('7. MailApp の呼び出しは notifyFailure_ の1か所だけで、宛先は Session 固定', () => {
  const fs = require('fs');
  const { CODE_PATH } = require('./load-code');
  const src = fs.readFileSync(CODE_PATH, 'utf8');

  const mailCalls = src.match(/MailApp\.\w+/g) || [];
  assert.deepStrictEqual(mailCalls, ['MailApp.sendEmail'], 'MailApp の使用は1か所だけ: ' + mailCalls.join(', '));

  const toLines = src.split('\n').filter(l => /var to =/.test(l));
  assert.strictEqual(toLines.length, 1);
  assert.ok(/Session\.getEffectiveUser\(\)\.getEmail\(\)/.test(toLines[0]),
    '宛先が Session.getEffectiveUser() 固定であること: ' + toLines[0].trim());

  // 取引先へ請求書を送る類の関数・添付が無い
  assert.ok(!/attachments/.test(src), 'メール添付の実装が無いこと');
  assert.ok(!/sendInvoice|mailInvoice|sendToClient/.test(src), '請求書送付の関数が無いこと');
});

// ============================================================
// 8. 本番(DEMO_MODE=false)でもボタン配置経路が存在する
// ============================================================
test('8. 本番(DEMO_MODE=false)でも「請求書を生成」ボタンが明細シートに配置される', () => {
  const env = loadCode({ nowMs: ms(2026, 7, 31, 14, 0) });
  env.context.CONFIG.DEMO_MODE = false;

  env.context.generateInvoices(); // デモデータは作らない経路（明細・取引先マスタは空のまま）

  const items = env.ss.getSheetByName('明細');
  assert.ok(items, '本番実行でも明細シートが用意されること（無ければ作る）');
  const hasButton = items.getImages().some(img => img.getScript() === 'generateInvoices');
  assert.ok(hasButton, 'DEMO_MODE=false でも生成ボタンが明細シートに配置されること');

  // このテストの主眼はボタン配置。取引先マスタ・明細が空でも例外にならず0件で終わる
  const ledger = sheetRows(env.ss, '発行台帳');
  assert.strictEqual(ledger.length, 0, '対象データが無いので発行は0件');
});

test('6b. CONFIG.OPTIONS.ERROR_MAIL を切ると、自動発行が失敗してもメールを送らない', () => {
  const env = loadCode({ nowMs: ms(2026, 8, 25, 9, 5), failPdf: true });
  seedSchedule(env.ss, SCHEDULE_25TH);
  env.context.CONFIG.OPTIONS.ERROR_MAIL = false;

  env.context.dailyIssueCheck();
  assert.strictEqual(env.calls.mail.length, 0, 'メールを送らないこと');
  assert.strictEqual(sheetRows(env.ss, '実行ログ')[0][LOG.RESULT], '失敗', 'ログには失敗が残ること');
});

// ============================================================
// 9. 消費税の端数処理は「税率ごとに1回」（行ごとに丸めると結果が変わる数字で確かめる）
// ============================================================
test('9. 端数処理は税率ごとに1回だけ（行ごとに丸めた場合とは結果が違う数字で確認）', () => {
  const env = loadCode({ nowMs: ms(2026, 7, 31, 14, 0) });
  // 10%対象 105円 × 2行: 行ごとに切り捨てると 10+10=20円、税率ごとに1回なら floor(21)=21円
  const lines = [
    { name: 'a', qty: 1, unitPrice: 105, taxRate: 0.10 },
    { name: 'b', qty: 1, unitPrice: 105, taxRate: 0.10 },
    { name: 'c', qty: 3, unitPrice: 1, taxRate: 0.08 }
  ];
  const floor = env.context.computeTotals_(lines, 'floor');
  assert.strictEqual(floor.byRate['0.1'].net, 210);
  assert.strictEqual(floor.byRate['0.1'].tax, 21, '10%対象の税は 210×0.1 を1回だけ丸めた21円');
  assert.strictEqual(floor.byRate['0.08'].tax, 0, '8%対象 3円の税 0.24円は切り捨てで0円');
  assert.strictEqual(floor.subtotal, 213);
  assert.strictEqual(floor.taxTotal, 21);
  assert.strictEqual(floor.total, 234);

  const ceil = env.context.computeTotals_(lines, 'ceil');
  assert.strictEqual(ceil.byRate['0.08'].tax, 1, '切り上げ指定が効くこと（0.24円→1円）');
  assert.strictEqual(ceil.byRate['0.1'].tax, 21, '割り切れる額は切り上げても変わらないこと');

  const round = env.context.computeTotals_([{ name: 'd', qty: 1, unitPrice: 115, taxRate: 0.10 }], 'round');
  assert.strictEqual(round.taxTotal, 12, '四捨五入指定が効くこと（11.5円→12円）');
});

// ============================================================
// 10. 支払期限（取引先ごとの指定 → 既定へのフォールバック）
// ============================================================
test('10. 支払期限は取引先の指定を優先し、空欄・不正値は既定（翌月末）になる', () => {
  const env = loadCode({ nowMs: ms(2026, 7, 31, 14, 0) });
  const issue = new Date(2026, 6, 31);
  assert.strictEqual(env.context.paymentDue_(issue, ''), '2026-08-31', '空欄は既定の翌月末');
  assert.strictEqual(env.context.paymentDue_(issue, '翌月末'), '2026-08-31');
  assert.strictEqual(env.context.paymentDue_(issue, '発行日+30'), '2026-08-30');
  assert.strictEqual(env.context.paymentDue_(new Date(2026, 0, 31), '翌月末'), '2026-02-28', '2月は28日');
  const before = env.calls.log.length;
  assert.strictEqual(env.context.paymentDue_(issue, '来月10日'), '2026-08-31', '不正値は既定に戻る');
  assert.ok(env.calls.log.length > before, '不正値のときは警告をログに残すこと');
});

// ============================================================
// 11. 請求書HTMLの中身（記載事項とエスケープ）
// ============================================================
test('11. 請求書HTMLに登録番号・軽減税率の印・税率ごとの税額が入り、取引先名はエスケープされる', () => {
  const env = loadCode({ nowMs: ms(2026, 7, 31, 14, 0) });
  const lines = [
    { name: 'コーヒー豆', qty: 2, unitPrice: 1000, taxRate: 0.08, date: '2026-07-05' },
    { name: '保守', qty: 1, unitPrice: 10000, taxRate: 0.10, date: '2026-07-06' }
  ];
  const html = env.context.buildInvoiceHtml_({
    issuer: env.context.CONFIG.ISSUER,
    client: { name: '<script>alert(1)</script>商事', address: '' },
    lines, totals: env.context.computeTotals_(lines, 'floor'),
    invoiceNo: 'INV-202607-009', issueDate: '2026-07-31', dueDate: '2026-08-31',
    targetMonth: '2026-07', unit: '案件X'
  });
  assert.ok(html.indexOf('<script>') < 0, '取引先名のタグがそのまま出ないこと');
  assert.ok(html.indexOf('&lt;script&gt;alert(1)&lt;/script&gt;商事 御中') >= 0, '宛名がエスケープされて入ること');
  assert.ok(html.indexOf('登録番号: T1234567890123') >= 0, '発行者の登録番号');
  assert.ok(html.indexOf('コーヒー豆 ※') >= 0, '軽減税率の品目に ※ が付くこと');
  assert.ok(html.indexOf('<td>軽減8% 対象</td><td class="num">2,000</td><td class="num">160</td>') >= 0, '8%の対価と税額');
  assert.ok(html.indexOf('<td>10% 対象</td><td class="num">10,000</td><td class="num">1,000</td>') >= 0, '10%の対価と税額');
  assert.ok(html.indexOf('<th>日付</th>') >= 0 && html.indexOf('<td>2026-07-05</td>') >= 0, '請求単位ありは日付列が出ること');
  assert.ok(html.indexOf('（対象: 2026-07（案件X））') >= 0, '対象表記に請求単位が付くこと');
});

// ============================================================
// 12. 請求書番号は1か月1,000件目以降も重複しない
// ============================================================
test('12. 同じ月の1,000件目以降は4桁になり、番号が重複しない', () => {
  const env = loadCode({ nowMs: ms(2026, 7, 31, 14, 0) });
  const ledger = env.ss.insertSheet('台帳テスト');
  ledger.appendRow(['請求書番号', '発行日', '対象月']);
  for (let i = 0; i < 998; i++) ledger.appendRow(['', '2026-07-31', '2026-07']);
  assert.strictEqual(env.context.nextInvoiceNo_(ledger, '2026-07'), 'INV-202607-999');
  ledger.appendRow(['', '2026-07-31', '2026-07']);
  assert.strictEqual(env.context.nextInvoiceNo_(ledger, '2026-07'), 'INV-202607-1000');
  ledger.appendRow(['', '2026-07-31', '2026-07']);
  assert.strictEqual(env.context.nextInvoiceNo_(ledger, '2026-07'), 'INV-202607-1001');
  assert.strictEqual(env.context.nextInvoiceNo_(ledger, '2026-08'), 'INV-202608-001', '別の月は001から');
});

// ---- 結果 ----
console.log('\n' + '-'.repeat(60));
if (failures.length) {
  console.log(`RED: ${passed} passed / ${failures.length} failed`);
  failures.forEach(f => {
    console.log('\n[FAIL] ' + f.name);
    console.log(f.error && f.error.stack ? f.error.stack : String(f.error));
  });
  process.exit(1);
}
console.log(`GREEN: ${passed} passed / 0 failed`);
