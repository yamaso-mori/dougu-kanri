/**
 * 山創_道具管理 - スプレッドシート連携用 Apps Script
 * (Google のみで完結する構成: HTML Service でフロントエンドも配信)
 *
 * 使い方:
 * 1. 対象のスプレッドシートを開き、拡張機能 > Apps Script を開く
 * 2. デフォルトの Code.gs の中身を全部削除し、このファイルの内容を貼り付ける
 * 3. ファイル > 新規作成 > HTML で「Index」という名前のファイルを作成し、
 *    Index.html の内容を貼り付ける(拡張子は自動で付くのでファイル名は "Index" のみ)
 * 4. プロジェクトの設定 > スクリプト プロパティ に「PASSCODE」を追加し、
 *    値にアプリの合言葉を設定する(未設定のままだとアプリは使えません)
 * 5. 保存して「デプロイ」>「新しいデプロイ」
 *    - 種類: ウェブアプリ
 *    - 実行するユーザー: 自分
 *    - アクセスできるユーザー: 全員
 * 6. 発行されたウェブアプリのURL(https://script.google.com/macros/s/.../exec)を
 *    ブックマークすれば、それがそのままアプリのURLになります
 * 7. iPhone のホーム画面にアイコン付きで置きたい場合は、リポジトリ直下の index.html の
 *    GAS_URL に 6 のURLを設定し、GitHub Pages のURLを「ホーム画面に追加」する
 *
 * 画面(google.script.run)から呼べるのは、名前の末尾が "_" でない関数だけ。
 * 画面に公開する関数は必ず最初に requirePasscode_ で合言葉を確認し、
 * それ以外の内部関数は末尾に "_" を付けて直接呼べないようにしている。
 */

const SHEET_TOOLS = '道具';
const SHEET_HISTORY = '移動履歴';
const SHEET_LOCATIONS = '現在地';
const SHEET_SITE_CHECK = '現場状況確認';

// 合言葉を保存するスクリプト プロパティの名前
const PROP_PASSCODE = 'PASSCODE';
// 合言葉が違うときのエラーの目印(画面側はこの文字列で再入力画面を出す)
const AUTH_ERROR_MARK = 'AUTH_REQUIRED';

const MAX_NAME_LENGTH = 100;
const MAX_NOTE_LENGTH = 500;
const MAX_QTY = 100000;
const MAX_SITE_CHECK_ITEMS = 500;

// ホーム画面・タブに出すアイコン画像の公開URL(PNG)。
// Index.html 内の <link rel="icon"> は iframe 内なので効かず、ここで設定する必要がある。
// data: URL は不可。例: Google ドライブで「リンクを知っている全員」に共有した画像の
// ファイルIDを使い 'https://drive.google.com/uc?id=ファイルID&.png'
const FAVICON_URL = '';

function doGet() {
  const output = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('アレドコ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    // GitHub Pages の入口ページ(index.html)から iframe で埋め込めるようにする
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  if (FAVICON_URL) output.setFaviconUrl(FAVICON_URL);
  return output;
}

// ---------- 画面に公開する関数(すべて合言葉が必要) ----------

function getData(passcode) {
  requirePasscode_(passcode);
  return { tools: getTools_(), history: getHistory_(), siteChecks: getSiteChecks_() };
}

function addMove(passcode, record) {
  requirePasscode_(passcode);
  return withScriptLock_(() => addMoveImpl_(record));
}

function addTool(passcode, tool) {
  requirePasscode_(passcode);
  return withScriptLock_(() => addToolImpl_(tool));
}

function updateTool(passcode, tool) {
  requirePasscode_(passcode);
  return withScriptLock_(() => updateToolImpl_(tool));
}

// 現場状況タブでチェックが入った道具だけを「現場状況確認」シートに1件ずつ記録する
function submitSiteCheck(passcode, payload) {
  requirePasscode_(passcode);
  return withScriptLock_(() => submitSiteCheckImpl_(payload));
}

// ---------- 合言葉・入力値の検証 ----------

function requirePasscode_(passcode) {
  const expected = String(PropertiesService.getScriptProperties().getProperty(PROP_PASSCODE) || '').trim();
  if (!expected) throw new Error('合言葉が未設定のため利用できません。管理者に連絡してください');
  if (String(passcode == null ? '' : passcode).trim() !== expected) {
    Utilities.sleep(1000); // 総当たりで試されにくいよう、間違えたときは少し待たせる
    throw new Error(AUTH_ERROR_MARK + ': 合言葉が違います');
  }
}

// シートに書き込む文字列を整える。"=" などで始まる値はシート上で数式として
// 実行されてしまうため、先頭に "'" を付けて文字列として保存する
// ("'" はシートの「文字列として扱う」印で、読み出した値には含まれない)。
function safeText_(value, maxLength) {
  const s = String(value == null ? '' : value).trim().slice(0, maxLength || MAX_NAME_LENGTH);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function toCount_(value, label, min) {
  const n = Number(value);
  if (!Number.isInteger(n) || n < min || n > MAX_QTY) {
    throw new Error(label + 'は' + min + '以上の整数で入力してください');
  }
  return n;
}

function toRowNumber_(value, sheet) {
  const row = Number(value);
  if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) {
    throw new Error('更新対象の行が特定できません。画面を更新してからやり直してください');
  }
  return row;
}

// ---------- シートの読み書き ----------

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('シートが見つかりません: ' + name);
  return sh;
}

function sheetToObjects_(sheet) {
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

function formatDateTime_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

// 複数人が同時に記録・編集した場合に書き込みが競合しないよう、
// シートを更新する処理はこのロックで1件ずつ直列に実行する。
function withScriptLock_(fn) {
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

function getTools_() {
  const rows = sheetToObjects_(getSheet_(SHEET_TOOLS));
  rows.forEach((r) => {
    ['作成日時', '更新日時'].forEach((k) => {
      if (r[k] instanceof Date) r[k] = formatDateTime_(r[k]);
    });
  });
  return rows;
}

function getHistory_() {
  const rows = sheetToObjects_(getSheet_(SHEET_HISTORY));
  rows.forEach((r) => {
    ['移動日', '作成日時', '更新日時'].forEach((k) => {
      if (r[k] instanceof Date) r[k] = formatDateTime_(r[k]);
    });
  });
  return rows;
}

function getSiteChecks_() {
  const rows = sheetToObjects_(getSheet_(SHEET_SITE_CHECK));
  rows.forEach((r) => {
    if (r['チェック日時'] instanceof Date) r['チェック日時'] = formatDateTime_(r['チェック日時']);
  });
  return rows;
}

function addMoveImpl_(record) {
  if (!record || !record.toolName) throw new Error('道具名は必須です');
  if (!record.from || !record.to) throw new Error('移動元・移動先は必須です');
  const qty = toCount_(record.qty, '数量', 1);

  // 道具シートに登録されている道具だけを記録できるようにする
  const toolName = String(record.toolName).trim();
  const tool = getTools_().find((t) => String(t['道具名']) === toolName);
  if (!tool) throw new Error('道具が見つかりません: ' + toolName);
  const total = Number(tool['総数']) || 0;
  if (qty > total) {
    throw new Error(toolName + ' の移動数(' + qty + ')が総数(' + total + ')を超えています');
  }

  const from = safeText_(record.from);
  const to = safeText_(record.to);
  const now = new Date();
  const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const moveDate = /^\d{4}-\d{2}-\d{2}$/.test(String(record.moveDate || '')) ? record.moveDate : today;

  getSheet_(SHEET_HISTORY).appendRow([
    safeText_(record.recorder),
    safeText_(record.worker),
    safeText_(toolName),
    qty,
    moveDate,
    from,
    to,
    safeText_(record.note, MAX_NOTE_LENGTH),
    now,
    now,
  ]);
  updateCurrentLocationLedger_(toolName, from, to, qty, tool);
  return { ok: true };
}

// 移動記録に合わせて「現在地」シートを更新する。
// 道具1つにつき「場所」ごとに1行を持たせ、移動元は数量を減らす(0になれば行を削除)、
// 移動先は既存行があれば加算、なければ新しい行を追加する。
function updateCurrentLocationLedger_(toolName, from, to, qty, tool) {
  const sheet = getSheet_(SHEET_LOCATIONS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const col = {};
  headers.forEach((h, i) => { col[h] = i; });
  const now = new Date();

  // シートから読み出した値には safeText_ の "'" が含まれないので、比較は "'" を除いた値で行う
  const plain = (s) => String(s).replace(/^'/, '');
  let fromRow = -1;
  let toRow = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][col['道具名']]) !== toolName) continue;
    if (String(values[i][col['場所']]) === plain(from)) fromRow = i + 1;
    if (String(values[i][col['場所']]) === plain(to)) toRow = i + 1;
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
    setByHeader('大分類', safeText_(tool['大分類']));
    setByHeader('小分類', safeText_(tool['小分類']));
    setByHeader('道具名', safeText_(toolName));
    setByHeader('表示順', tool['表示順']);
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

  sortLocationsSheetByOrder_(sheet, headers, col);
}

// 現在地シートを表示順(同じ道具内では場所名)で並び替える
function sortLocationsSheetByOrder_(sheet, headers, col) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return; // データが0〜1行なら並び替え不要
  const orderCol = col['表示順'] + 1;
  const placeCol = col['場所'] + 1;
  const sortSpec = [{ column: orderCol, ascending: true }];
  if (placeCol > 0) sortSpec.push({ column: placeCol, ascending: true });
  sheet.getRange(2, 1, lastRow - 1, headers.length).sort(sortSpec);
}

// 道具シートのヘッダー名 -> 0始まり列インデックスのマップを取得する
function getToolsColumnMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = {};
  headers.forEach((h, i) => { col[h] = i; });
  return { headers: headers, col: col };
}

function addToolImpl_(tool) {
  if (!tool || !String(tool.name || '').trim()) throw new Error('道具名は必須です');
  const sheet = getSheet_(SHEET_TOOLS);
  const { headers, col } = getToolsColumnMap_(sheet);

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
  setByHeader('大分類', safeText_(tool.major));
  setByHeader('小分類', safeText_(tool.minor));
  setByHeader('道具名', safeText_(tool.name));
  setByHeader('未使用時保管場所', safeText_(tool.storage));
  setByHeader('備考', safeText_(tool.note, MAX_NOTE_LENGTH));
  setByHeader('表示順', tool.order === '' || tool.order == null ? nextNo : toCount_(tool.order, '表示順', 1));
  setByHeader('総数', tool.total === '' || tool.total == null ? 0 : toCount_(tool.total, '総数', 0));
  setByHeader('作成日時', now);
  setByHeader('更新日時', now);
  sheet.appendRow(newRow);
  return { ok: true };
}

function updateToolImpl_(tool) {
  if (!tool || !tool.__row) throw new Error('更新対象の行が特定できません');
  if (!String(tool.name || '').trim()) throw new Error('道具名は必須です');
  const sheet = getSheet_(SHEET_TOOLS);
  const { col } = getToolsColumnMap_(sheet);
  const row = toRowNumber_(tool.__row, sheet);
  const setByHeader = (name, value) => {
    if (col[name] !== undefined) sheet.getRange(row, col[name] + 1).setValue(value);
  };
  setByHeader('大分類', safeText_(tool.major));
  setByHeader('小分類', safeText_(tool.minor));
  setByHeader('道具名', safeText_(tool.name));
  setByHeader('未使用時保管場所', safeText_(tool.storage));
  setByHeader('備考', safeText_(tool.note, MAX_NOTE_LENGTH));
  setByHeader('表示順', tool.order === '' || tool.order == null ? 1 : toCount_(tool.order, '表示順', 1));
  setByHeader('総数', tool.total === '' || tool.total == null ? 0 : toCount_(tool.total, '総数', 0));
  setByHeader('更新日時', new Date());
  return { ok: true };
}

function submitSiteCheckImpl_(payload) {
  if (!payload || !payload.place) throw new Error('現場が選択されていません');
  if (!payload.checker) throw new Error('チェック者を入力してください');
  if (!Array.isArray(payload.items) || !payload.items.length) throw new Error('確認済みの道具がありません');
  if (payload.items.length > MAX_SITE_CHECK_ITEMS) throw new Error('一度に送信できる件数を超えています');

  const sheet = getSheet_(SHEET_SITE_CHECK);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = {};
  headers.forEach((h, i) => { col[h] = i; });
  const now = new Date();

  const rows = payload.items.map((item) => {
    if (!item || !item.toolName) throw new Error('道具名が不正です');
    const row = new Array(headers.length).fill('');
    const setByHeader = (name, value) => {
      if (col[name] !== undefined) row[col[name]] = value;
    };
    setByHeader('場所', safeText_(payload.place));
    setByHeader('道具', safeText_(item.toolName));
    setByHeader('数量', item.qty === '' || item.qty == null ? '' : toCount_(item.qty, '数量', 0));
    setByHeader('備考', safeText_(item.note, MAX_NOTE_LENGTH));
    setByHeader('チェック日時', now);
    setByHeader('チェック者', safeText_(payload.checker));
    return row;
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  return { ok: true, count: rows.length };
}
