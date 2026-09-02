/**
 * TSUMUGI ORDER CONCEPT SHEET — 受信エンドポイント
 *
 * 役割
 *   1. 送信内容をスプレッドシートに追記（器 1点＝1行）
 *   2. 添付写真を Google Drive のフォルダに保存（案件フォルダ内に器ごとのサブフォルダ）
 *   3. 担当者へメール通知
 *
 * 設置手順は SETUP.md を参照してください。
 */

// ── 設定 ──────────────────────────────────────────────
const SHEET_ID   = 'ここにスプレッドシートのIDを入れる';        // 手順1で作成するスプレッドシートのID
const DRIVE_ID   = '1unUjWdgtqUFw1XVnQ_oGca0AgY0WaOhT';   // 写真の保存先Driveフォルダ
const MAIL_TO    = 's.y.connect.co@gmail.com';             // カンマ区切りで複数可
const MAIL_CC    = 'm.yuri0713@gmail.com';                 // カンマ区切りで複数可（不要なら空文字に）
const SHEET_NAME = '受注シート';
// ─────────────────────────────────────────────────────

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const items = (data.items && data.items.length) ? data.items : [{ rows: [], files: [] }];
    const folder = saveFiles_(data, items);
    items.forEach(function (item, idx) { appendRow_(data, item, idx, items.length, folder); });
    notify_(data, items, folder);
    return json_({ ok: true, ref: data.ref });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  }
}

/** 案件フォルダを作り、共通の写真＋器ごとの写真を保存。全体の回答をanswers.txtにまとめる */
function saveFiles_(data, items) {
  const root = DriveApp.getFolderById(DRIVE_ID);
  const name = [stamp_(data.submittedAt), data.store || '無題', data.ref].join(' ');
  const folder = root.createFolder(name);

  (data.sharedFiles || []).forEach(function (f) { saveOneFile_(folder, f); });

  items.forEach(function (item, idx) {
    if (!item.files || !item.files.length) return;
    const itemFolder = folder.createFolder('器' + (idx + 1));
    item.files.forEach(function (f) { saveOneFile_(itemFolder, f); });
  });

  const lines = (data.sharedRows || [])
    .map(function (r) { return [r.section, r.group, r.label, r.value].join('\t'); });
  items.forEach(function (item, idx) {
    lines.push('');
    lines.push('── 器' + (idx + 1) + ' ──');
    (item.rows || []).forEach(function (r) { lines.push([r.section, r.group, r.label, r.value].join('\t')); });
  });
  folder.createFile(Utilities.newBlob(lines.join('\n'), 'text/plain', 'answers.txt'));

  return folder;
}

function saveOneFile_(folder, f) {
  const parts = String(f.dataUrl).split(',');
  const bytes = Utilities.base64Decode(parts[1]);
  const blob = Utilities.newBlob(bytes, 'image/jpeg', f.key + '.jpg');
  const file = folder.createFile(blob);
  if (f.memo) file.setDescription(f.label + ' — ' + f.memo);
}

/** 器 1点＝1行。店舗・共通情報の列は全行で同じ値になる */
function appendRow_(data, item, idx, total, folder) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  const fixed = ['受付日時', '受付番号', '器 No.', 'お店の名前', '案件名', 'ご担当者', '連絡先', 'Driveフォルダ', '写真枚数(全体)'];
  const sharedLabels = (data.sharedRows || []).map(function (r) { return r.label; });
  const itemLabels = (item.rows || []).map(function (r) { return r.label; });

  if (sh.getLastRow() === 0) {
    const header = fixed.concat(sharedLabels).concat(itemLabels);
    sh.appendRow(header);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }

  const sharedMap = {};
  (data.sharedRows || []).forEach(function (r) { sharedMap[r.key] = r.value; });
  const itemMap = {};
  (item.rows || []).forEach(function (r) { itemMap[r.key] = r.value; });
  const sharedKeys = (data.sharedRows || []).map(function (r) { return r.key; });
  const itemKeys = (item.rows || []).map(function (r) { return r.key; });

  const totalFiles = (data.sharedFiles || []).length +
    (data.items || []).reduce(function (s, it) { return s + (it.files || []).length; }, 0);

  sh.appendRow([
    new Date(data.submittedAt),
    data.ref,
    (idx + 1) + ' / ' + total,
    data.store,
    data.project,
    data.contact,
    data.contactTel,
    folder.getUrl(),
    totalFiles
  ].concat(sharedKeys.map(function (k) { return sharedMap[k] || ''; }))
   .concat(itemKeys.map(function (k) { return itemMap[k] || ''; })));
}

function notify_(data, items, folder) {
  const totalFilled = (data.sharedRows || []).filter(function (r) { return r.value; }).length +
    items.reduce(function (s, it) { return s + (it.rows || []).filter(function (r) { return r.value; }).length; }, 0);
  const totalFiles = (data.sharedFiles || []).length +
    items.reduce(function (s, it) { return s + (it.files || []).length; }, 0);

  const sharedLines = (data.sharedRows || [])
    .filter(function (r) { return r.value; })
    .map(function (r) { return '・' + r.label + '：' + r.value; })
    .join('\n');

  const itemLines = items.map(function (item, idx) {
    const lines = (item.rows || [])
      .filter(function (r) { return r.value; })
      .map(function (r) { return '・' + r.label + '：' + r.value; })
      .join('\n');
    return '【器' + (idx + 1) + '】\n' + (lines || '（未入力）');
  }).join('\n\n');

  MailApp.sendEmail({
    to: MAIL_TO,
    cc: MAIL_CC,
    subject: '【ORDER SHEET】' + (data.store || '無題') + ' / ' + (data.project || '') + '（' + data.ref + '・器' + items.length + '点）',
    body: [
      'コンセプトシートが届きました。',
      '',
      '受付番号：' + data.ref,
      'お店：' + data.store,
      '案件：' + data.project,
      'ご担当：' + data.contact + '（' + data.contactTel + '）',
      '器の点数：' + items.length + ' 点',
      '記入項目：' + totalFilled + ' 件 ／ 写真：' + totalFiles + ' 枚',
      '写真フォルダ：' + folder.getUrl(),
      '',
      '──────────── 店舗・共通情報 ────────────',
      sharedLines || '（未入力）',
      '',
      '──────────── 器の仕様 ────────────',
      itemLines
    ].join('\n')
  });
}

function stamp_(iso) {
  return Utilities.formatDate(new Date(iso), 'Asia/Tokyo', 'yyyy-MM-dd');
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
