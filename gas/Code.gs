/**
 * 山創_道具管理 - スプレッドシート連携用 Apps Script
 * (Google のみで完結する構成: HTML Service でフロントエンドも配信)
 *
 * 使い方:
 * 1. 対象のスプレッドシートを開き、拡張機能 > Apps Script を開く
 * 2. デフォルトの Code.gs の中身を全部削除し、このファイルの内容を貼り付ける
 * 3. ファイル > 新規作成 > HTML で「Index」という名前のファイルを作成し、
 *    Index.html の内容を貼り付ける(拡張子は自動で付くのでファイル名は "Index" のみ)
 * 4. 保存して「デプロイ」>「新しいデプロイ」
 *    - 種類: ウェブアプリ
 *    - 実行するユーザー: 自分
 *    - アクセスできるユーザー: 全員
 * 5. 発行されたウェブアプリのURL(https://script.google.com/macros/s/.../exec)を
 *    ブックマークすれば、それがそのままアプリのURLになります
 * 6. iPhone のホーム画面にアイコン付きで置きたい場合は、リポジトリ直下の index.html の
 *    GAS_URL に 5 のURLを設定し、GitHub Pages のURLを「ホーム画面に追加」する
 */

const SHEET_TOOLS = '道具';
const SHEET_HISTORY = '移動履歴';
const SHEET_LOCATIONS = '現在地';
const SHEET_SITE_CHECK = '現場状況確認';

// ホーム画面・タブに出すアイコン画像の公開URL(PNG)。
// Index.html 内の <link rel="icon"> は iframe 内なので効かず、ここで設定する必要がある。
// data: URL は不可。例: Google ドライブで「リンクを知っている全員」に共有した画像の
// ファイルIDを使い 'https://drive.google.com/uc?id=ファイルID&.png'
const FAVICON_URL = '';

function doGet() {
  const output = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('山創 道具管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    // GitHub Pages の入口ページ(index.html)から iframe で埋め込めるようにする
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  if (FAVICON_URL) output.setFaviconUrl(FAVICON_URL);
  return output;
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('シートが見つかりません: ' + name);
  return sh;
}

function sheetToObjects(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 1) return [];
  const headers = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every((c) => c === '' || c === null)) continue;
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = row[idx];
    });
    obj.__row = i + 1; // シート上の実際の行番号(1始まり、ヘッダーが1行目)
    rows.push(obj);
  }
  return rows;
}

function formatDateTime(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function getData() {
  return { tools: getTools(), history: getHistory(), siteChecks: getSiteChecks() };
}

// 複数人が同時に記録・編集した場合に書き込みが競合しないよう、
// シートを更新する処理はこのロックで1件ずつ直列に実行する。
function withScriptLock(fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // 最大30秒、他の人の書き込みが終わるのを待つ
  } catch (e) {
    throw new Error('他の人の操作と重なったため保存できませんでした。少し待ってからもう一度お試しください');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function getTools() {
  const rows = sheetToObjects(getSheet(SHEET_TOOLS));
  rows.forEach((r) => {
    ['作成日時', '更新日時'].forEach((k) => {
      if (r[k] instanceof Date) r[k] = formatDateTime(r[k]);
    });
  });
  return rows;
}

function getHistory() {
  const rows = sheetToObjects(getSheet(SHEET_HISTORY));
  rows.forEach((r) => {
    ['移動日', '作成日時', '更新日時'].forEach((k) => {
      if (r[k] instanceof Date) r[k] = formatDateTime(r[k]);
    });
  });
  return rows;
}

function getSiteChecks() {
  const rows = sheetToObjects(getSheet(SHEET_SITE_CHECK));
  rows.forEach((r) => {
    if (r['チェック日時'] instanceof Date) r['チェック日時'] = formatDateTime(r['チェック日時']);
  });
  return rows;
}

function addMove(record) {
  return withScriptLock(() => addMoveImpl(record));
}

function addMoveImpl(record) {
  if (!record || !record.toolName) throw new Error('道具名は必須です');
  if (!record.from || !record.to) throw new Error('移動元・移動先は必須です');
  const qty = Number(record.qty);
  if (!qty || qty <= 0) throw new Error('数量は1以上の数値で入力してください');

  const tool = getTools().find((t) => t['道具名'] === record.toolName);
  const total = tool ? Number(tool['総数']) || 0 : null;
  if (total != null && qty > total) {
    throw new Error(record.toolName + ' の移動数(' + qty + ')が総数(' + total + ')を超えています');
  }

  const sheet = getSheet(SHEET_HISTORY);
  const now = new Date();
  sheet.appendRow([
    record.recorder || '',
    record.worker || '',
    record.toolName,
    qty,
    record.moveDate || Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    record.from || '',
    record.to || '',
    record.note || '',
    now,
    now,
  ]);
  updateCurrentLocationLedger(record.toolName, record.from, record.to, qty, tool);
  return { ok: true };
}

// 移動記録に合わせて「現在地」シートを更新する。
// 道具1つにつき「場所」ごとに1行を持たせ、移動元は数量を減らす(0になれば行を削除)、
// 移動先は既存行があれば加算、なければ新しい行を追加する。
function updateCurrentLocationLedger(toolName, from, to, qty, tool) {
  const sheet = getSheet(SHEET_LOCATIONS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const col = {};
  headers.forEach((h, i) => { col[h] = i; });
  const now = new Date();

  let fromRow = -1;
  let toRow = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][col['道具名']] !== toolName) continue;
    if (values[i][col['場所']] === from) fromRow = i + 1;
    if (values[i][col['場所']] === to) toRow = i + 1;
  }

  // 移動先: 既存行があれば加算、なければ新規行を追加(既存行の削除より先に行うことで行番号のズレを避ける)
  if (toRow > 0) {
    const current = Number(sheet.getRange(toRow, col['数'] + 1).getValue()) || 0;
    sheet.getRange(toRow, col['数'] + 1).setValue(current + qty);
    sheet.getRange(toRow, col['記録日時'] + 1).setValue(now);
  } else {
    // 列の並びが変わっても崩れないよう、ヘッダー名を頼りに値を配置する
    const newRow = new Array(headers.length).fill('');
    const setByHeader = (name, value) => {
      if (col[name] !== undefined) newRow[col[name]] = value;
    };
    setByHeader('大分類', tool ? tool['大分類'] : '');
    setByHeader('小分類', tool ? tool['小分類'] : '');
    setByHeader('道具名', toolName);
    setByHeader('表示順', tool ? tool['表示順'] : '');
    setByHeader('場所', to);
    setByHeader('数', qty);
    setByHeader('記録日時', now);
    sheet.appendRow(newRow);
  }

  // 移動元: 数量を減らし、0以下になったら行ごと削除する
  if (fromRow > 0) {
    const remaining = (Number(sheet.getRange(fromRow, col['数'] + 1).getValue()) || 0) - qty;
    if (remaining > 0) {
      sheet.getRange(fromRow, col['数'] + 1).setValue(remaining);
      sheet.getRange(fromRow, col['記録日時'] + 1).setValue(now);
    } else {
      sheet.deleteRow(fromRow);
    }
  }

  sortLocationsSheetByOrder(sheet, headers, col);
}

// 現在地シートを表示順(同じ道具内では場所名)で並び替える
function sortLocationsSheetByOrder(sheet, headers, col) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return; // データが0〜1行なら並び替え不要
  const orderCol = col['表示順'] + 1;
  const placeCol = col['場所'] + 1;
  const sortSpec = [{ column: orderCol, ascending: true }];
  if (placeCol > 0) sortSpec.push({ column: placeCol, ascending: true });
  sheet.getRange(2, 1, lastRow - 1, headers.length).sort(sortSpec);
}

function deleteMove(rowIndex) {
  return withScriptLock(() => deleteMoveImpl(rowIndex));
}

function deleteMoveImpl(rowIndex) {
  const row = Number(rowIndex);
  if (!row || row < 2) throw new Error('不正な行番号です');
  const sheet = getSheet(SHEET_HISTORY);
  sheet.deleteRow(row);
  return { ok: true };
}

// 道具シートのヘッダー名 -> 0始まり列インデックスのマップを取得する
function getToolsColumnMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = {};
  headers.forEach((h, i) => { col[h] = i; });
  return { headers: headers, col: col };
}

function addTool(tool) {
  return withScriptLock(() => addToolImpl(tool));
}

function addToolImpl(tool) {
  if (!tool || !tool.name) throw new Error('道具名は必須です');
  const sheet = getSheet(SHEET_TOOLS);
  const { headers, col } = getToolsColumnMap(sheet);

  const lastRow = sheet.getLastRow();
  let nextNo = 1;
  if (lastRow > 1 && col['No'] !== undefined) {
    const lastNo = sheet.getRange(lastRow, col['No'] + 1).getValue();
    nextNo = (Number(lastNo) || lastRow - 1) + 1;
  }
  const now = new Date();
  const newRow = new Array(headers.length).fill('');
  const setByHeader = (name, value) => {
    if (col[name] !== undefined) newRow[col[name]] = value;
  };
  setByHeader('No', nextNo);
  setByHeader('大分類', tool.major || '');
  setByHeader('小分類', tool.minor || '');
  setByHeader('道具名', tool.name);
  setByHeader('未使用時保管場所', tool.storage || '');
  setByHeader('備考', tool.note || '');
  setByHeader('表示順', Number(tool.order) || nextNo);
  setByHeader('総数', Number(tool.total) || 0);
  setByHeader('作成日時', now);
  setByHeader('更新日時', now);
  sheet.appendRow(newRow);
  return { ok: true };
}

function updateTool(tool) {
  return withScriptLock(() => updateToolImpl(tool));
}

function updateToolImpl(tool) {
  if (!tool || !tool.__row) throw new Error('更新対象の行が特定できません');
  const sheet = getSheet(SHEET_TOOLS);
  const { col } = getToolsColumnMap(sheet);
  const row = Number(tool.__row);
  const setByHeader = (name, value) => {
    if (col[name] !== undefined) sheet.getRange(row, col[name] + 1).setValue(value);
  };
  setByHeader('大分類', tool.major || '');
  setByHeader('小分類', tool.minor || '');
  setByHeader('道具名', tool.name || '');
  setByHeader('未使用時保管場所', tool.storage || '');
  setByHeader('備考', tool.note || '');
  setByHeader('表示順', Number(tool.order) || 1);
  setByHeader('総数', Number(tool.total) || 0);
  setByHeader('更新日時', new Date());
  return { ok: true };
}

// 現場状況タブでチェックが入った道具だけを「現場状況確認」シートに1件ずつ記録する
function submitSiteCheck(payload) {
  return withScriptLock(() => submitSiteCheckImpl(payload));
}

function submitSiteCheckImpl(payload) {
  if (!payload || !payload.place) throw new Error('現場が選択されていません');
  if (!payload.checker) throw new Error('チェック者を入力してください');
  if (!payload.items || !payload.items.length) throw new Error('確認済みの道具がありません');

  const sheet = getSheet(SHEET_SITE_CHECK);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = {};
  headers.forEach((h, i) => { col[h] = i; });
  const now = new Date();

  const rows = payload.items.map((item) => {
    if (!item.toolName) throw new Error('道具名が不正です');
    const row = new Array(headers.length).fill('');
    const setByHeader = (name, value) => {
      if (col[name] !== undefined) row[col[name]] = value;
    };
    setByHeader('場所', payload.place);
    setByHeader('道具', item.toolName);
    setByHeader('数量', item.qty === '' || item.qty == null ? '' : Number(item.qty));
    setByHeader('備考', item.note || '');
    setByHeader('チェック日時', now);
    setByHeader('チェック者', payload.checker);
    return row;
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  return { ok: true, count: rows.length };
}
