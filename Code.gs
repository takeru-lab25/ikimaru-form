/**
 * いきまる イベント申込受付（最終版GAS）
 * doGet ?action=schedule : 日程＋現在の申込数 → 受付中/満席/残りわずか＋参加費/場所/補足（JSONP）
 * doGet ?action=profile&userId=XXX : 同じLINEユーザーの「前回の申込内容」を返す（自動入力用・JSONP）
 * doPost : 定員チェック（門番）のうえ「申込ログ」に記録。複数日は日付ごとに1行で記録。新規申込はメール通知。
 *
 * 必要シート（すべて同じスプレッドシート内）：
 *  「日程」   … A:グループ B:日程(表示) C:記録用ラベル D:定員 E:受付 F:参加費 G:場所 H:補足 I:残りわずか  ※1行目は見出し
 *  「申込ログ」… 全体台帳（自動作成・定員管理と自動入力に使用）
 *  ＋ 回ごとに「記録用ラベル」と同名のシート（例：260712 海あそび）にも自動で振り分け記録（当日名簿用・自動作成）
 *  既存の申込を日付別シートへ反映するには backfillEventSheets() を1回だけ実行。
 */

var SHEET_LOG  = "申込ログ";
var SHEET_SCHED= "日程";
var CHANNEL_ACCESS_TOKEN = "";
var NOTIFY_EMAIL = "takeru.asaka@ikimaru.co.jp";   // 新規申込の通知先（空にすると通知オフ）

/* ========== 配信 ========== */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === "schedule") return jsonp_(p.callback, buildSchedule_());
  if (p.action === "profile")  return jsonp_(p.callback, lastByUser_(p.userId));
  return HtmlService.createHtmlOutput(
    '<meta charset="utf-8"><div style="font-family:sans-serif;padding:24px">' +
    'いきまる 申込受付API：稼働中です。</div>'
  ).setTitle("いきまる 申込受付");
}

function buildSchedule_() {
  var counts = countByLabel_();
  var events = readSched_().map(function(r) {
    var remaining = r.cap - (counts[r.label] || 0);
    var full = (!r.open) || remaining <= 0;
    var low  = (!full) && r.lowThresh > 0 && remaining <= r.lowThresh; // 回ごとのしきい値
    return { group:r.group, name:r.name, value:r.label, status:(full?"full":"open"), low:low,
             fee:r.fee, place:r.place, note:r.note };
  });
  return { events: events };
}

// 同じLINEユーザーの最後の申込内容（自動入力用）
function lastByUser_(uid) {
  if (!uid) return { found:false };
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  if (!s || s.getLastRow() < 2) return { found:false };
  var v = s.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][1]) === String(uid)) {
      return { found:true,
        guardian:v[i][3], area:v[i][4], schools:v[i][5], phone:v[i][6],
        emgName:v[i][7], emgRel:v[i][8], emgPhone:v[i][9], children:v[i][10], health:v[i][14] };
    }
  }
  return { found:false };
}

/* ========== 受付（門番つき記録） ========== */
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var need = Number(d.childCount) || splitList_(d.children).length || 1;
    var sched = {}; readSched_().forEach(function(r){ sched[r.label] = r; });
    var counts = countByLabel_();

    var accepted = [], waited = [];
    // 確定希望：空きがあれば確定、満席なら自動的にキャンセル待ちへ回す
    splitList_(d.dates).forEach(function(lab){
      var r = sched[lab], used = counts[lab] || 0;
      if (r && r.open && (used + need) <= r.cap) { accepted.push(lab); counts[lab] = used + need; }
      else waited.push(lab);
    });
    // フロントで満席と分かったうえで選んだ分はキャンセル待ち（定員に数えない）
    splitList_(d.waitlist).forEach(function(lab){ if (waited.indexOf(lab) < 0 && accepted.indexOf(lab) < 0) waited.push(lab); });

    if (accepted.length === 0 && waited.length === 0) return json_({ ok:false, full:true });

    var log = getLog_(), now = new Date();
    // 全体台帳（申込ログ）と 日付別シート の両方へ1行ずつ記録
    accepted.forEach(function(lab){ var row = buildRow_(now, d, need, lab, "確定"); log.appendRow(row); eventSheet_(lab).appendRow(row); });
    waited.forEach(function(lab){ var row = buildRow_(now, d, need, lab, "キャンセル待ち"); log.appendRow(row); eventSheet_(lab).appendRow(row); });

    notifyNewEntry_(d, accepted, waited, need);                    // ★ 新規申込メール通知
    return json_({ ok:true, accepted:accepted, waitlist:waited });
  } catch (err) { return json_({ ok:false, error:String(err) }); }
}

// 1日付＝1行（状態＝確定／キャンセル待ち）の配列を作る
function buildRow_(now, d, need, lab, status) {
  return [
    now, d.userId || "", d.displayName || "", d.guardian || "",
    d.area || "", d.schools || "", d.phone || "",
    d.emgName || "", d.emgRel || "", d.emgPhone || "",
    d.children || "", need, lab, status, d.health || "",
    d.photo || "", d.agree ? "同意" : ""
  ];
}

function splitList_(s) {
  return String(s || "").split(" / ").map(function(x){ return x.trim(); }).filter(String);
}

/* ========== 新規申込メール通知 ========== */
function notifyNewEntry_(d, accepted, waited, need) {
  if (!NOTIFY_EMAIL) return;
  try {
    var tag = (accepted.length ? "確定" : "") + (waited.length ? (accepted.length ? "＋キャンセル待ち" : "キャンセル待ち") : "");
    var subject = "【いきまる】新規申込み（" + tag + "）：" + (d.guardian || "") + "様";
    var body =
      "新規申込みが入りました。\n" +
      "------------------------------\n" +
      (accepted.length ? "■ 確定した参加日\n  " + accepted.join("\n  ") + "\n\n" : "") +
      (waited.length ? "■ キャンセル待ち\n  " + waited.join("\n  ") + "\n\n" : "") +
      "■ お子さま（" + need + "名）\n  " + (d.children || "") + "\n\n" +
      "■ 連絡事項\n  " + (d.health || "（記入なし）") + "\n\n" +
      "■ 保護者氏名\n  " + (d.guardian || "") + "\n" +
      "■ 在住（市区町村）\n  " + (d.area || "") + "\n" +
      "■ 園・学校\n  " + (d.schools || "") + "\n" +
      "■ 電話\n  " + (d.phone || "") + "\n" +
      "■ 緊急連絡先\n  " + (d.emgName || "") + "（" + (d.emgRel || "") + "）" + (d.emgPhone || "") + "\n" +
      "■ 写真掲載\n  " + (d.photo || "") + "\n" +
      "------------------------------\n" +
      "※詳細はスプレッドシート「申込ログ」をご確認ください。";
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
  } catch (err) { /* 通知失敗で受付自体は止めない */ }
}

/* ========== シート読み取り ========== */
function readSched_() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SCHED);
  if (!s || s.getLastRow() < 2) return [];
  var v = s.getDataRange().getValues(), out = [];
  for (var i = 1; i < v.length; i++) {
    var label = v[i][2];
    if (!label) continue;
    out.push({
      group: String(v[i][0] || ""),
      name:  String(v[i][1] || label), label: String(label),
      cap:   Number(v[i][3]) || 0,
      open:  (String(v[i][4]).toUpperCase() !== "OFF" && v[i][4] !== false),
      fee:   String(v[i][5] || ""), place: String(v[i][6] || ""), note: String(v[i][7] || ""),
      lowThresh: Number(v[i][8]) || 0
    });
  }
  return out;
}

function countByLabel_() {
  var counts = {}, s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  if (!s || s.getLastRow() < 2) return counts;
  var v = s.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][13]) === "キャンセル待ち") continue;   // 待ちは定員に数えない
    var n = Number(v[i][11]) || 0;
    String(v[i][12] || "").split(" / ").forEach(function(lab){ lab = lab.trim(); if (lab) counts[lab] = (counts[lab] || 0) + n; });
  }
  return counts;
}

var ENTRY_HEADER = ["受付日時","LINEユーザーID","LINE表示名","保護者氏名","市区町村","園・学校",
  "電話","緊急連絡先氏名","続柄","緊急連絡先電話","お子さま","人数","参加希望日","状態","連絡事項","写真掲載","同意"];

// 申込シート（見出し＋受付日時の表記）を整える共通処理
function ensureEntrySheet_(s) {
  var cur = s.getLastRow() >= 1 ? s.getRange(1, 1, 1, ENTRY_HEADER.length).getValues()[0] : [];
  var same = cur.length === ENTRY_HEADER.length && ENTRY_HEADER.every(function(h, i){ return cur[i] === h; });
  if (!same) s.getRange(1, 1, 1, ENTRY_HEADER.length).setValues([ENTRY_HEADER]);
  // 受付日時（A列）を「6月20日（木） 11:22」表記に統一（実データはDateのまま＝並べ替え可）
  s.getRange(2, 1, Math.max(s.getMaxRows() - 1, 1), 1).setNumberFormat('m"月"d"日（"aaa"） "hh:mm');
  return s;
}

function getLog_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ensureEntrySheet_(ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG));
}

// 回（記録用ラベル）ごとの当日名簿シート
function eventSheet_(label) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = sheetName_(label);
  return ensureEntrySheet_(ss.getSheetByName(name) || ss.insertSheet(name));
}

// シート名に使えない文字を除去（/ \ ? * [ ] :）
function sheetName_(label) {
  return (String(label).replace(/[\/\\\?\*\[\]\:]/g, "-").trim().slice(0, 90)) || "申込";
}

/* ★1回だけ実行：既存の「申込ログ」を日付別シートへ反映（重複しないよう各シートを作り直す。何度実行してもOK） */
function backfillEventSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName(SHEET_LOG);
  if (!log || log.getLastRow() < 2) return;
  var v = log.getDataRange().getValues();
  var byLabel = {};
  for (var i = 1; i < v.length; i++) {
    var lab = String(v[i][12] || "").trim();      // M列＝参加希望日（記録用ラベル）
    if (!lab) continue;
    var row = v[i].slice(0, ENTRY_HEADER.length);
    while (row.length < ENTRY_HEADER.length) row.push("");   // 列数を17に揃える
    (byLabel[lab] = byLabel[lab] || []).push(row);
  }
  Object.keys(byLabel).forEach(function(lab){
    var s = eventSheet_(lab);
    if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, ENTRY_HEADER.length).clearContent();
    s.getRange(2, 1, byLabel[lab].length, ENTRY_HEADER.length).setValues(byLabel[lab]);
  });
}

function pushConfirm_(userId, d, accepted) {
  var text = "お申込みありがとうございます。\n【" + d.children + "】\n希望日：" + accepted.join(" / ") +
    "\n当日は現金でのお支払いをお願いします。前日にあらためてご連絡しますね。";
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
    method:"post", contentType:"application/json",
    headers:{ Authorization:"Bearer " + CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify({ to:userId, messages:[{ type:"text", text:text }] }), muteHttpExceptions:true });
}

/* メール送信権限を一度だけ承認するためのテスト関数（エディタで実行→承認） */
function testMail() {
  MailApp.sendEmail(NOTIFY_EMAIL, "テスト：いきまる申込通知", "メール通知の動作テストです。これが届けば設定OKです。");
}

function jsonp_(callback, obj) {
  var out = JSON.stringify(obj);
  if (callback) return ContentService.createTextOutput(callback + "(" + out + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
