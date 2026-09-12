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
 *
 * 【コードを更新した場合の再デプロイ】
 * Code.gsを保存しただけでは、公開中のURLには反映されません。
 * 「デプロイ」→「デプロイを管理」→ 編集（鉛筆アイコン）→
 * バージョン「新しいバージョン」を選んで「デプロイ」を押してください
 * （URLは変わりません）。
 *
 * 【APIトークンについて】
 * 下のAPI_TOKENは、役割分担ボード側のSYNC_TOKENと必ず同じ値にしてください。
 * これが一致しないアクセスはすべて拒否されます。
 */

const SHEET_NAME = "KV";

// 書き込みを許可するメールアドレス（空欄なら誰でも書き込み可＝暫定運用向け）
// 組織のGoogleアカウントが決まったら、ここに列挙してアクセスを絞ってください。
// 例: const ALLOWED_EMAILS = ["leader1@example.com", "admin@example.com"];
const ALLOWED_EMAILS = [];

// APIトークン（合言葉）。この値と一致しないGET/POSTリクエストは拒否します。
// アプリ側（役割分担ボードのHTML）の SYNC_TOKEN と必ず同じ値にしてください。
const API_TOKEN = "vwhO7Nt49C51xyc-SannBWMhvldfkVkm";

function checkToken_(token) {
  return API_TOKEN && token === API_TOKEN;
}

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

// 全データ取得：GET ?token=API_TOKEN
function doGet(e) {
  try {
    const token = (e && e.parameter && e.parameter.token) || "";
    if (!checkToken_(token)) return jsonOut_({ ok: false, error: "forbidden" });
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

// 1件保存：POST { key, value, token }（valueは任意のJSON）
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (!checkToken_(body && body.token)) return jsonOut_({ ok: false, error: "forbidden" });

    let email = "";
    try { email = Session.getActiveUser().getEmail(); } catch (err) { email = ""; }

    if (ALLOWED_EMAILS.length && ALLOWED_EMAILS.indexOf(email) === -1) {
      return jsonOut_({ ok: false, error: "forbidden", email: email });
    }

    if (!body || !body.key) return jsonOut_({ ok: false, error: "key is required" });

    // 複数端末が同時に書き込んでも行が重複・破損しないよう、書き込み中は他の書き込みを待たせる
    const lock = LockService.getScriptLock();
    lock.waitLock(10000); // 最大10秒待つ。取れなければ下のcatchでエラーを返す
    try {
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
      SpreadsheetApp.flush();
      return jsonOut_({ ok: true, updatedAt: now });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    // ロック待ちタイムアウトなど。アプリ側が自動で再送するので、ここでは失敗を返すだけでよい。
    return jsonOut_({ ok: false, error: String(err) });
  }
}
