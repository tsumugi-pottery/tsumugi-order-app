/**
 * TSUMUGI ORDER CONCEPT SHEET — 受信エンドポイント
 *
 * 役割
 *   1. 送信内容をスプレッドシートに1行 append
 *   2. 添付写真を Google Drive のフォルダに保存（案件ごとにサブフォルダ）
 *   3. 担当者へメール通知
 *
 * 設置手順は SETUP.md を参照してください。
 */

// ── 設定 ──────────────────────────────────────────────
const SHEET_ID   = 'ここにスプレッドシートのIDを入れる';
const DRIVE_ID   = 'ここに保存先Driveフォルダのidを入れる';
const MAIL_TO    = 'orders@example.com';        // カンマ区切りで複数可
const SHEET_NAME = '受注シート';
// ─────────────────────────────────────────────────────

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const folder = saveFiles_(data);
    appendRow_(data, folder);
    notify_(data, folder);
    return json_({ ok: true, ref: data.ref });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  }
}

/** 写真を案件ごとのサブフォルダに保存し、フォルダを返す */
function saveFiles_(data) {
  const root = DriveApp.getFolderById(DRIVE_ID);
  const name = [stamp_(data.submittedAt), data.store || '無題', data.ref].join(' ');
  const folder = root.createFolder(name);

  (data.files || []).forEach(function (f) {
    const parts = String(f.dataUrl).split(',');
    const bytes = Utilities.base64Decode(parts[1]);
    const blob = Utilities.newBlob(bytes, 'image/jpeg', f.key + '.jpg');
    const file = folder.createFile(blob);
    if (f.memo) file.setDescription(f.label + ' — ' + f.memo);
  });

  // 回答全文もテキストで同梱しておくと、Drive だけ見れば内容が分かる
  const body = (data.rows || [])
    .map(function (r) { return [r.section, r.group, r.label, r.value].join('\t'); })
    .join('\n');
  folder.createFile(Utilities.newBlob(body, 'text/plain', 'answers.txt'));

  return folder;
}

/** 1件 = 1行。列は rows の key 順に自動で増える */
function appendRow_(data, folder) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  const fixed = ['受付日時', '受付番号', 'お店の名前', '案件名', 'ご担当者', '連絡先', 'Driveフォルダ', '写真枚数'];
  const keys = (data.rows || []).map(function (r) { return r.key; });
  const labels = (data.rows || []).map(function (r) { return r.label; });

  if (sh.getLastRow() === 0) {
    sh.appendRow(fixed.concat(labels));
    sh.getRange(1, 1, 1, fixed.length + labels.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }

  const map = {};
  (data.rows || []).forEach(function (r) { map[r.key] = r.value; });

  sh.appendRow([
    new Date(data.submittedAt),
    data.ref,
    data.store,
    data.project,
    data.contact,
    data.contactTel,
    folder.getUrl(),
    (data.files || []).length
  ].concat(keys.map(function (k) { return map[k] || ''; })));
}

function notify_(data, folder) {
  const filled = (data.rows || []).filter(function (r) { return r.value; }).length;
  const lines = (data.rows || [])
    .filter(function (r) { return r.value; })
    .map(function (r) { return '・' + r.label + '：' + r.value; })
    .join('\n');

  MailApp.sendEmail({
    to: MAIL_TO,
    subject: '【ORDER SHEET】' + (data.store || '無題') + ' / ' + (data.project || '') + '（' + data.ref + '）',
    body: [
      'コンセプトシートが届きました。',
      '',
      '受付番号：' + data.ref,
      'お店：' + data.store,
      '案件：' + data.project,
      'ご担当：' + data.contact + '（' + data.contactTel + '）',
      '記入項目：' + filled + ' 件 ／ 写真：' + (data.files || []).length + ' 枚',
      '写真フォルダ：' + folder.getUrl(),
      '',
      '────────────',
      lines
    ].join('\n')
  });
}

function stamp_(iso) {
  return Utilities.formatDate(new Date(iso), 'Asia/Tokyo', 'yyyy-MM-dd');
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
