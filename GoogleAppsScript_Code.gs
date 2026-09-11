/**
 * 役割分担ボード用の簡易バックエンド（Google Apps Script）
 * ------------------------------------------------------
 * このコードは、スプレッドシートを「キー：値」の保存庫として使い、
 * 役割分担ボード（HTML版）からの読み書きを受け付けるWebアプリです。
 *
 * 【使い方】
 * 1. このコード全体をコピーし、Apps Scriptのエディタ（Code.gs）に貼り付けて保存
 * 2. 上部の「デプロイ」→「新しいデプロイ」→種類「ウェブアプリ」を選択
 * 3. 「実行するユーザー」＝自分（Me）
 *    「アクセスできるユーザー」＝
 *       ・組織のGoogle Workspaceがあれば「〇〇（組織名）内の全員」
 *       ・まだ無ければ、暫定的に「全員」を選択（URLを知る人だけがアクセスできる状態）
 * 4. デプロイ後に表示される「ウェブアプリのURL」（.../exec で終わるもの）を控えてClaudeに伝える
 *
 * 【あとで組織のアカウントに譲渡する場合】
 * スプレッドシートの共有設定で新しい管理者を「オーナー」に変更し、
 * Apps Scriptも同様に共有・移管すれば、このコードはそのまま使えます。
 */

const SHEET_NAME = "KV";

// 書き込みを許可するメールアドレス（空欄なら誰でも書き込み可＝暫定運用向け）
// 組織のGoogleアカウントが決まったら、ここに列挙してアクセスを絞ってください。
// 例: const ALLOWED_EMAILS = ["leader1@example.com", "admin@example.com"];
const ALLOWED_EMAILS = [];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["key", "value", "updatedAt", "updatedBy"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 全データ取得：GET ?action=getAll
function doGet(e) {
  try {
    const sh = getSheet_();
    const data = sh.getDataRange().getValues();
    const out = {};
    let maxUpdatedAt = "";
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0];
      const value = data[i][1];
      const updatedAt = data[i][2];
      if (!key) continue;
      try { out[key] = JSON.parse(value); } catch (err) { out[key] = value; }
      if (updatedAt && String(updatedAt) > maxUpdatedAt) maxUpdatedAt = String(updatedAt);
    }
    return jsonOut_({ ok: true, data: out, serverTime: new Date().toISOString(), maxUpdatedAt: maxUpdatedAt });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// 1件保存：POST { key, value }（valueは任意のJSON）
function doPost(e) {
  try {
    let email = "";
    try { email = Session.getActiveUser().getEmail(); } catch (err) { email = ""; }

    if (ALLOWED_EMAILS.length && ALLOWED_EMAILS.indexOf(email) === -1) {
      return jsonOut_({ ok: false, error: "forbidden", email: email });
    }

    const body = JSON.parse(e.postData.contents);
    if (!body || !body.key) return jsonOut_({ ok: false, error: "key is required" });

    const sh = getSheet_();
    const range = sh.getDataRange();
    const data = range.getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === body.key) { rowIndex = i + 1; break; }
    }
    const valueStr = JSON.stringify(body.value);
    const now = new Date().toISOString();

    if (rowIndex > 0) {
      sh.getRange(rowIndex, 2, 1, 3).setValues([[valueStr, now, email]]);
    } else {
      sh.appendRow([body.key, valueStr, now, email]);
    }
    return jsonOut_({ ok: true, updatedAt: now });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
