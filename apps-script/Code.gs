/**
 * TSUMUGI ORDER CONCEPT SHEET — 受信エンドポイント
 *
 * 役割
 *   1. 送信内容をスプレッドシートに追記（器 1点＝1行、社内の受注管理用）
 *   2. 添付写真を Google Drive のフォルダに保存（案件フォルダ内に器ごとのサブフォルダ）
 *   3. 窯元に渡せる「仕様書」（紙のオーダーシートと同じレイアウト、器ごとに別タブ）を自動生成
 *   4. 担当者へメール通知（仕様書へのリンク付き）
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
    const saved = saveFiles_(data, items);
    const folder = saved.folder;
    items.forEach(function (item, idx) { appendRow_(data, item, idx, items.length, folder); });
    const orderFile = buildOrderSheet_(data, items, folder, saved.sharedFileUrls, saved.itemFileUrls);
    notify_(data, items, folder, orderFile);
    return json_({ ok: true, ref: data.ref });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  }
}

function toMap_(rows) {
  const m = {};
  (rows || []).forEach(function (r) { m[r.key] = r.value; });
  return m;
}

/** 案件フォルダを作り、共通の写真＋器ごとの写真を保存。全体の回答をanswers.txtにまとめる */
function saveFiles_(data, items) {
  const root = DriveApp.getFolderById(DRIVE_ID);
  const name = [stamp_(data.submittedAt), data.store || '無題', data.ref].join(' ');
  const folder = root.createFolder(name);

  const sharedFileUrls = {};
  (data.sharedFiles || []).forEach(function (f) {
    sharedFileUrls[f.key] = saveOneFile_(folder, f).getUrl();
  });

  const itemFileUrls = items.map(function (item, idx) {
    const map = {};
    if (item.files && item.files.length) {
      const itemFolder = folder.createFolder('器' + (idx + 1));
      item.files.forEach(function (f) { map[f.key] = saveOneFile_(itemFolder, f).getUrl(); });
    }
    return map;
  });

  const lines = (data.sharedRows || [])
    .map(function (r) { return [r.section, r.group, r.label, r.value].join('\t'); });
  items.forEach(function (item, idx) {
    lines.push('');
    lines.push('── 器' + (idx + 1) + ' ──');
    (item.rows || []).forEach(function (r) { lines.push([r.section, r.group, r.label, r.value].join('\t')); });
  });
  folder.createFile(Utilities.newBlob(lines.join('\n'), 'text/plain', 'answers.txt'));

  return { folder: folder, sharedFileUrls: sharedFileUrls, itemFileUrls: itemFileUrls };
}

function saveOneFile_(folder, f) {
  const parts = String(f.dataUrl).split(',');
  const bytes = Utilities.base64Decode(parts[1]);
  const blob = Utilities.newBlob(bytes, 'image/jpeg', f.key + '.jpg');
  const file = folder.createFile(blob);
  if (f.memo) file.setDescription(f.label + ' — ' + f.memo);
  return file;
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

  const sharedMap = toMap_(data.sharedRows);
  const itemMap = toMap_(item.rows);
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

// ── 仕様書（窯元向け）の自動生成 ──────────────────────────
// 紙の「TSUMUGI ORDER CONCEPT SHEET」と同じ構成（店舗・ブランドコンセプト → 器仕様 → 製作スケジュール）
// をスプレッドシートとして自動で組み立てる。テンプレートファイルは使わず、このスクリプトだけで完結する。
// 器が複数点ある場合は、器ごとに別シート（タブ）に分ける。各タブは店舗情報・製作スケジュールも
// 含めた自己完結の1枚なので、そのまま窯元に渡せる。

const OS_INK = '#201e1d';
const OS_ACCENT = '#ec3013';
const OS_PANEL = '#f3f2f2';
const OS_WIDTH = 6; // レイアウトの列数（A〜F）

function buildOrderSheet_(data, items, folder, sharedFileUrls, itemFileUrls) {
  const ss = SpreadsheetApp.create('仕様書 ' + (data.store || '無題') + ' ' + data.ref);
  const file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file); // create()は必ずマイドライブ直下に作るので、案件フォルダへ移す

  const defaultSheet = ss.getSheets()[0];
  items.forEach(function (item, idx) {
    const tabName = items.length > 1 ? '器' + (idx + 1) : '仕様書';
    const sh = (idx === 0) ? defaultSheet : ss.insertSheet(idx);
    sh.setName(tabName);
    buildItemSheet_(sh, data, item, idx, items.length, sharedFileUrls, itemFileUrls[idx] || {});
  });

  return file;
}

function buildItemSheet_(sh, data, item, idx, itemCount, sharedFileUrls, fileUrls) {
  sh.setColumnWidths(1, 1, 130);
  sh.setColumnWidths(2, OS_WIDTH - 1, 150);

  const sharedMap = toMap_(data.sharedRows);
  const m = toMap_(item.rows);
  let row = 1;

  row = osTitle_(sh, row, 'TSUMUGI', 'ORDER CONCEPT SHEET — 仕様書' + (itemCount > 1 ? '（器' + (idx + 1) + ' / ' + itemCount + '）' : ''));
  row = osMeta_(sh, row, data, itemCount);
  row++;

  const contactBits = [sharedMap.contact_tel, sharedMap.contact_email].filter(String).join(' / ');

  row = osSection_(sh, row, '01', '店舗・ブランドコンセプト / BRAND & STORE CONCEPT');
  row = osRow2_(sh, row, '案件名', data.project, '作成日', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日'));
  row = osRow2_(sh, row, 'お店の名前', data.store, 'ご担当者', data.contact + (contactBits ? '（' + contactBits + '）' : ''));
  row = osRow1_(sh, row, 'お店のコンセプト', sharedMap.concept);
  row = osRow1_(sh, row, 'デザインの方向', sharedMap.direction);
  row = osRow2_(sh, row, '主な料理', sharedMap.dishes, '使用シーン', sharedMap.scene);
  row = osRow2_(sh, row, '盛付イメージ', sharedMap.plating, '特記事項', sharedMap.notes1);
  row = osPhotoRow_(sh, row, [
    { label: '店舗写真・内装', url: sharedFileUrls.atmos }
  ]);
  row++;

  row = osSection_(sh, row, '02', '器 仕様・デザイン / TABLEWARE SPECIFICATION' + (m.item ? '　－ ' + m.item + ' －' : ''));
  row = osRow2_(sh, row, 'アイテム名', m.item, '使用用途', m.usage);
  row = osPhotoRow_(sh, row, [
    { label: 'REFERENCE 01', url: fileUrls.ref1, memo: m.memo_ref1 },
    { label: 'REFERENCE 02', url: fileUrls.ref2, memo: m.memo_ref2 },
    { label: 'REFERENCE 03', url: fileUrls.ref3, memo: m.memo_ref3 }
  ]);
  row = osRow1_(sh, row, 'サイズ', osSize_(m));
  row = osRow1_(sh, row, '色・釉薬', m.glaze);
  row = osRow1_(sh, row, '素材感・表情', m.texture);
  row = osRow1_(sh, row, 'ロゴ・刻印', m.deco);
  row = osPhotoRow_(sh, row, [
    { label: 'ロゴ・刻印データ', url: fileUrls.deco_file }
  ]);
  row = osRow1_(sh, row, '希望数量', m.qty ? m.qty + ' 個' : '');
  row = osRow2_(sh, row, '希望単価', m.unitprice ? '¥' + m.unitprice + ' / 個' : '', '総予算', m.budget ? '¥' + m.budget : '');
  row = osRow2_(sh, row, 'サンプル製作', m.sample, '食洗機・レンジ', m.dish);
  row = osRow1_(sh, row, '避けたいこと（NG事項）', m.ng);
  row++;

  row = osSection_(sh, row, '03', '納品に関して / DELIVERY');
  row = osRow2_(sh, row, '希望納期', sharedMap.d_due, '納品場所', sharedMap.d_place);
  row = osRow2_(sh, row, '納品先担当者', sharedMap.d_person, '連絡先', sharedMap.d_tel);
  row = osRow1_(sh, row, '備考', sharedMap.remarks);
  row = osRow2_(sh, row, '承認日', '　　　年　　月　　日', '承認者', '');
  row++;

  row = osNote_(sh, row, 'NOTE　陶器は原料・釉薬・焼成条件により、サンプルと量産品の間でも色味・表情・寸法に個体差が生じる場合があります。最終サンプル承認時に許容範囲を確認させてください。');

  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, row - 1, OS_WIDTH).setBorder(true, true, true, true, false, false, OS_INK, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setFrozenRows(0);
}

function osSize_(m) {
  const parts = [];
  if (m.size_d) parts.push('Φ' + m.size_d);
  if (m.size_w) parts.push('W' + m.size_w);
  if (m.size_h) parts.push('H' + m.size_h);
  return parts.length ? parts.join(' × ') + ' mm' : '';
}

/** 帯状の見出し（ロゴエリア）：黒帯にTSUMUGI、下に赤い罫線と英字サブタイトル */
function osTitle_(sh, row, brand, subtitle) {
  sh.getRange(row, 1, 1, OS_WIDTH).merge().setValue(brand)
    .setBackground(OS_INK).setFontColor('#ffffff')
    .setFontSize(22).setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(row, 44);
  row++;

  sh.getRange(row, 1, 1, OS_WIDTH).merge().setValue(subtitle)
    .setFontSize(10).setFontColor(OS_ACCENT).setFontWeight('bold').setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, OS_ACCENT, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.setRowHeight(row, 22);
  row++;

  sh.setRowHeight(row, 6);
  return row + 1;
}

function osMeta_(sh, row, data, itemCount) {
  const labels = ['受付番号', '受付日時', '器 点数'];
  const values = [data.ref, stamp_(data.submittedAt), itemCount + ' 点'];
  labels.forEach(function (l, i) {
    sh.getRange(row, 1 + i * 2).setValue(l).setFontWeight('bold').setFontColor('#605d5d').setFontSize(9);
    sh.getRange(row, 2 + i * 2).setValue(values[i]).setFontSize(9);
  });
  return row + 1;
}

/** セクション見出し：大きな番号バッジ（赤）＋黒帯のタイトル */
function osSection_(sh, row, num, title) {
  sh.getRange(row, 1).setValue(num)
    .setBackground(OS_ACCENT).setFontColor('#ffffff').setFontWeight('bold').setFontSize(15)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(row, 2, 1, OS_WIDTH - 1).merge().setValue(title)
    .setBackground(OS_INK).setFontColor('#ffffff').setFontWeight('bold').setFontSize(11)
    .setVerticalAlignment('middle');
  sh.setRowHeight(row, 30);
  return row + 1;
}

function osLabelCell_(cell, text) {
  cell.setValue(text).setFontWeight('bold').setFontColor('#605d5d').setFontSize(9)
    .setBackground(OS_PANEL).setVerticalAlignment('top');
}
function osValueCell_(range, text) {
  const target = range.getNumColumns() > 1 ? range.merge() : range;
  target.setValue(text || '').setWrap(true).setVerticalAlignment('top').setFontSize(10);
}

/** グリッド線を隠している分、行の下に薄い罫線を引いて表として読めるようにする */
function osRowBorder_(sh, row) {
  sh.getRange(row, 1, 1, OS_WIDTH).setBorder(false, false, true, false, false, false, '#d7d3d3', SpreadsheetApp.BorderStyle.SOLID);
}

/** ラベル/値 のペアを1行に2組 */
function osRow2_(sh, row, l1, v1, l2, v2) {
  const half = Math.floor(OS_WIDTH / 2);
  osLabelCell_(sh.getRange(row, 1), l1);
  osValueCell_(sh.getRange(row, 2, 1, half - 1), v1);
  osLabelCell_(sh.getRange(row, 1 + half), l2);
  osValueCell_(sh.getRange(row, 1 + half + 1, 1, OS_WIDTH - half - 1), v2);
  osRowBorder_(sh, row);
  return row + 1;
}

/** ラベル/値 を1行いっぱいに1組（長文向け） */
function osRow1_(sh, row, l1, v1) {
  osLabelCell_(sh.getRange(row, 1), l1);
  osValueCell_(sh.getRange(row, 2, 1, OS_WIDTH - 1), v1);
  osRowBorder_(sh, row);
  return row + 1;
}

/** 写真参照行：Driveに保存済みの写真へのリンクとメモを並べる */
function osPhotoRow_(sh, row, photos) {
  const cols = Math.floor(OS_WIDTH / photos.length);
  photos.forEach(function (p, i) {
    const c = 1 + i * cols;
    const span = (i === photos.length - 1) ? OS_WIDTH - c + 1 : cols;
    const range = sh.getRange(row, c, 1, span);
    const cell = span > 1 ? range.merge() : range;
    if (p.url) {
      const label = p.label + (p.memo ? '（' + p.memo + '）' : '');
      cell.setFormula('=HYPERLINK("' + p.url + '","📎 ' + label.replace(/"/g, "'") + '")');
    } else {
      cell.setValue(p.label + '：未添付').setFontColor('#9b9797');
    }
    cell.setWrap(true).setVerticalAlignment('top').setFontSize(10);
  });
  osRowBorder_(sh, row);
  return row + 1;
}

function osNote_(sh, row, text) {
  const range = sh.getRange(row, 1, 1, OS_WIDTH).merge();
  range.setValue(text).setWrap(true).setFontSize(9).setFontColor('#605d5d')
    .setBorder(true, false, false, false, false, false, OS_INK, SpreadsheetApp.BorderStyle.SOLID_THICK);
  return row + 1;
}

function notify_(data, items, folder, orderFile) {
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
      'ご担当：' + data.contact + '（' + [data.contactTel, data.contactEmail].filter(Boolean).join(' / ') + '）',
      '器の点数：' + items.length + ' 点',
      '記入項目：' + totalFilled + ' 件 ／ 写真：' + totalFiles + ' 枚',
      '写真フォルダ：' + folder.getUrl(),
      '仕様書（窯元へそのまま共有できます）：' + orderFile.getUrl(),
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
