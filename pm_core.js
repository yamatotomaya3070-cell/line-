// ==========================================
// プロジェクト進捗管理：共通基盤（定数・ヘルパー・セットアップ）
//   既存 コード.js を壊さない追加モジュール。
//   すべてのトップレベル関数・定数は `pm` / `PM_` 接頭辞で名前空間衝突を回避。
//   依存（既存 コード.js）: getSS, getSheet, setupSheet, getConfig,
//     fmtDate, fmtDT, generateId, safeParseJson
// ==========================================

// ---- シート名（英語名。既存の日本語シートと衝突しない） ----
var PM_SHEET_PROJECTS   = 'projects';
var PM_SHEET_LOGS       = 'project_logs';
var PM_SHEET_ALERTS     = 'project_alerts';
var PM_SHEET_CALENDAR   = 'project_calendar_events';
var PM_SHEET_PENDING    = 'project_pending_updates';
// 設定（グループID等をソースに書かず、運用者が編集できる置き場）。Script Property の代替。
var PM_SHEET_SETTINGS   = 'PM設定';

var PM_SETTINGS_HEADERS = ['キー', '値', '説明'];

// 社内グループ台帳（役割→LINEグループID）。案件グループとは別の「通知先」管理シート。
// 運用者が「グループID」列にIDを記入する（ソースにIDは書かない）。
var PM_SHEET_INTERNAL_GROUPS   = '社内グループ';
var PM_INTERNAL_GROUPS_HEADERS = ['役割', 'グループID', '表示名', '備考', '更新日時'];

// ---- projects シート 列定義（設計書 §6 準拠） ----
var PM_PROJECTS_HEADERS = [
  '案件ID', '案件名', 'クライアント名', '担当者', 'LINEグループID', 'DriveフォルダURL',
  '現在フェーズ', '営業ステータス', '設計・デザインステータス', '施工準備ステータス',
  '施工ステータス', 'お引き渡しステータス', '最終更新日', '次回アクション', '次回アクション期限',
  '打合せ日', '着工予定日', '引き渡し予定日', '売上', '原価', '利益額', '利益率',
  '請求予定日', '請求済日', '請求金額', '入金予定日', '入金確認日', '入金金額', '請求ステータス', '入金ステータス',
  '備考', '更新元メッセージ', '最終更新者', '作成日', '更新日',
  // 見積プロファイル両対応：案件にプロファイルIDを持たせると「案件→プロファイル」で解決、
  // 無ければ会社/トリガーキーワード単位にフォールバック（フェーズ1の est_core で参照）
  '見積プロファイルID',
  // アプリでの「取り消し」＝物理削除でなくアーカイブ。TRUEでボードから隠す（履歴/Bot/連携は壊さない）
  '取消',
  // 区分：案件/社内/その他。ボードは「案件」のみ表示し社内・連絡用グループを除外。
  '区分',
  // 施工フェーズの進捗％（0〜100）。ボードの施工カードから10%刻みで更新。
  '施工進捗',
];

// 案件グループ（1案件に複数グループをぶら下げる・フェーズ1.5）
// ※将来「複数案件を1案件にまとめる(merge)」機能は、この案件グループの付け替え＋取消(アーカイブ)で実現予定。
//   merge本体は後日しっかり相談してから実装する（今回は土台のみ）。
var PM_SHEET_GROUPS   = '案件グループ';
var PM_GROUPS_HEADERS = ['グループID', '案件ID', '種別', '表示名', '備考', '登録日時'];

var PM_LOGS_HEADERS     = ['ログID', '日時', '案件ID', 'intent', '変更項目', '更新元メッセージ', '更新者', 'グループID', '適用区分'];
var PM_ALERTS_HEADERS   = ['アラートID', '案件ID', '種別', 'アラートキー', '検知日時', '通知先', '内容', '状態'];
var PM_CALENDAR_HEADERS = ['案件ID', 'イベント種別', 'イベントID', '開始日時', '終了日時', '担当者', '作成日', '更新日'];
var PM_PENDING_HEADERS  = ['確認ID', '作成日時', '案件ID', '案件名候補', 'type', 'intent', '抽出内容', '更新元メッセージ', '申請者', 'グループID', 'status', '承認者', '処理日時'];

// ---- フェーズ・ステータス マスタ（設計書 §7） ----
var PM_PHASES = {
  '営業':           ['引合／問合せ受付', '現地調査／ヒアリング', '概算見積済', '見積提出（フォロー中）', '契約書確認中', '受注', '保留（長期検討）', '失注'],
  '設計・デザイン': ['未着手', '基本設計中', '実施設計中', 'お客様確認待ち', '修正対応中', '承認済'],
  '施工準備':       ['未着手', '工程表作成中', '工程表完成', '発注／材料手配中', '職人手配中', '近隣挨拶', '原価確定（着工前）', '確認申請中'],
  '施工':           ['未着手', '着工', '施工中', '中間検査', '行政検査', '社内検査', '是正期間'],
  'お引き渡し':     ['未', '完了検査', '引き渡し準備', '引き渡し済', '最終請求済', '入金確認', 'アフター／保証期間'],
};

// フェーズ → projects のステータス列名
var PM_PHASE_COLUMN = {
  '営業': '営業ステータス',
  '設計・デザイン': '設計・デザインステータス',
  '施工準備': '施工準備ステータス',
  '施工': '施工ステータス',
  'お引き渡し': 'お引き渡しステータス',
};

// ---- アラートしきい値（設計書 §9 / 確定値） ----
var PM_STALE_AFTER_QUOTE_DAYS = 7;   // 見積提出（フォロー中）後の放置判定
var PM_STALE_NO_UPDATE_DAYS   = 7;   // 全案件共通の無更新日数（最終更新から7日で初回通知）
var PM_STALE_REMIND_INTERVAL_DAYS = 7; // 放置アラートの再通知間隔（前回通知から7日ごと）

// ==========================================
// 設定・宛先
//   設計方針：グループID・経理グループID等の機密IDは「ソースに書かない」。
//   取得元は ① Script Property ② PM設定シート（運用者が編集）の順でフォールバック。
//   Script Property UIが上限で編集できない環境向けに PM設定シートを用意している。
// ==========================================
function pmProp(key) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  if (v) return v;
  return pmGetSetting(key); // プロパティ未設定なら PM設定シートを参照
}

// PM設定シート（キー/値）から1件取得。無ければ空文字。
function pmGetSetting(key) {
  try {
    var sheet = pmSheet(PM_SHEET_SETTINGS);
    if (!sheet || sheet.getLastRow() <= 1) return '';
    var idx = pmHeaderIndex(sheet);
    var cKey = idx['キー'];
    var cVal = idx['値'];
    if (cKey === undefined || cVal === undefined) return '';
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][cKey]).trim() === String(key).trim()) return String(data[i][cVal] || '').trim();
    }
  } catch (e) { console.error('pmGetSetting error:', e.message); }
  return '';
}

// 社内グループシートから役割（社内/経理/日次通知 等）のグループIDを引く。未設定は空文字。
function pmRoleGroupId(role) {
  try {
    var rows = pmReadObjects(PM_SHEET_INTERNAL_GROUPS);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i]['役割'] || '').trim() === String(role).trim()) {
        var gid = String(rows[i]['グループID'] || '').trim();
        if (gid.indexOf('C') === 0 || gid.indexOf('R') === 0) return gid; // グループ/ルームIDのみ許可
        return '';
      }
    }
  } catch (e) { console.error('pmRoleGroupId error:', e.message); }
  return '';
}

// 役割付き通知先の解決：① Script Property ② 社内グループシート ③ PM設定シート
function pmNamedTarget(role, propKey) {
  var v = PropertiesService.getScriptProperties().getProperty(propKey);
  if (v) return v;
  return pmRoleGroupId(role) || pmGetSetting(propKey);
}

// 金額・請求の「確認待ち」通知先（経理専用グループ。未設定時は社内グループへフォールバック）
function pmAccountingTarget() {
  return pmNamedTarget('経理', 'ACCOUNTING_GROUP_ID') || pmInternalTarget();
}

// 管理者・社内通知先
function pmInternalTarget() {
  return pmNamedTarget('社内', 'INTERNAL_GROUP_ID');
}

// 社内グループシートの指定役割にグループIDを登録／更新（無ければ追記）。
// スクリプトプロパティ PROGRESS_GROUP_ID 等が設定済みだとそちらが優先される点に注意。
function pmSetRoleGroupId(role, groupId, displayName) {
  var gid = String(groupId || '').trim();
  if (gid.indexOf('C') !== 0 && gid.indexOf('R') !== 0) return { ok: false, msg: 'グループ/ルーム内でのみ設定できます' };
  var sheet = pmSheet(PM_SHEET_INTERNAL_GROUPS);
  if (!sheet) return { ok: false, msg: '社内グループシートがありません（pmSetupSheets未実行?）' };
  var idx = pmHeaderIndex(sheet);
  var cRole = idx['役割'], cGid = idx['グループID'];
  if (cRole === undefined || cGid === undefined) return { ok: false, msg: '社内グループシートの列が不正です' };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][cRole] || '').trim() === String(role).trim()) {
      pmWriteRowFields(PM_SHEET_INTERNAL_GROUPS, i + 1, { 'グループID': gid, '表示名': displayName || '', '更新日時': fmtDT(new Date()) });
      return { ok: true, updated: true };
    }
  }
  pmAppendRowFields(PM_SHEET_INTERNAL_GROUPS, { '役割': role, 'グループID': gid, '表示名': displayName || '', '更新日時': fmtDT(new Date()) });
  return { ok: true, updated: false };
}

// 進捗アラート通知先（進捗管理グループ。未設定時は社内グループへフォールバック）
function pmProgressTarget() {
  return pmNamedTarget('進捗', 'PROGRESS_GROUP_ID') || pmInternalTarget();
}

// ==========================================
// 同時実行ロック（承認の二重実行・カレンダー/タスクの二重作成を防止）
//   GASは Webhook再送や複数ボタン連打で同一処理が並走しうる。
//   read-modify-write をこのヘルパーで囲って直列化する。
//   戻り値: { ok:true, result } | { ok:false, error:'lock_timeout' }
// ==========================================
function pmWithLock(fn, timeoutMs) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(timeoutMs || 15000);
  } catch (e) {
    console.error('pmWithLock: ロック取得失敗（タイムアウト）', e.message);
    return { ok: false, error: 'lock_timeout' };
  }
  try {
    return { ok: true, result: fn() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ==========================================
// シートアクセス・行オブジェクト変換（列順非依存）
// ==========================================
function pmSheet(name) { return getSheet(name); }

// ヘッダー行(1行目)から {ヘッダー名: 0基準index} を返す
function pmHeaderIndex(sheet) {
  var map = {};
  if (!sheet || sheet.getLastColumn() === 0) return map;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) map[String(headers[i]).trim()] = i;
  return map;
}

// シート全行を {ヘッダー:値, _row: 行番号(1基準)} の配列で返す
function pmReadObjects(name) {
  var sheet = pmSheet(name);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var values  = sheet.getDataRange().getValues();
  var headers = values[0].map(function(h) { return String(h).trim(); });
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var obj = { _row: r + 1 };
    for (var c = 0; c < headers.length; c++) obj[headers[c]] = values[r][c];
    out.push(obj);
  }
  return out;
}

// 指定行に {ヘッダー:値} を書き込む（存在する列のみ）
function pmWriteRowFields(name, rowNumber, fields) {
  var sheet = pmSheet(name);
  if (!sheet) return;
  var idx = pmHeaderIndex(sheet);
  Object.keys(fields).forEach(function(h) {
    if (idx[h] !== undefined) sheet.getRange(rowNumber, idx[h] + 1).setValue(fields[h]);
  });
}

// {ヘッダー:値} を1行追記（ヘッダー順に整列）
function pmAppendRowFields(name, fields) {
  var sheet = pmSheet(name);
  if (!sheet) return;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
  var row = headers.map(function(h) { return fields[h] !== undefined ? fields[h] : ''; });
  sheet.appendRow(row);
}

// ==========================================
// 日付ユーティリティ
// ==========================================
function pmTodayYmd() { return fmtDate(new Date()); }

// 'YYYY-MM-DD' / Date / 'YYYY/MM/DD ...' を Date に。失敗時 null
function pmParseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var s = String(v).trim();
  var m = s.match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function pmDaysBetween(a, b) {
  if (!a || !b) return null;
  return Math.floor((b.getTime() - a.getTime()) / 86400000);
}

function pmIsBlank(v) { return v === null || v === undefined || String(v).trim() === ''; }

// ==========================================
// セットアップ（シート作成・列追加・トリガー登録）
//   既存 setupSheet() を流用し、非破壊で作成。
// ==========================================
function pmSetupSheets() {
  var ss = getSS();
  setupSheet(ss, PM_SHEET_PROJECTS, PM_PROJECTS_HEADERS, '#6AA84F', null);
  setupSheet(ss, PM_SHEET_LOGS,     PM_LOGS_HEADERS,     '#A2C4C9', null);
  setupSheet(ss, PM_SHEET_ALERTS,   PM_ALERTS_HEADERS,   '#E06666', null);
  setupSheet(ss, PM_SHEET_CALENDAR, PM_CALENDAR_HEADERS, '#F1C232', null);
  setupSheet(ss, PM_SHEET_PENDING,  PM_PENDING_HEADERS,  '#C27BA0', null);
  setupSheet(ss, PM_SHEET_GROUPS,   PM_GROUPS_HEADERS,   '#6FA8DC', null);
  setupSheet(ss, PM_SHEET_SETTINGS, PM_SETTINGS_HEADERS, '#999999', null);
  setupSheet(ss, PM_SHEET_INTERNAL_GROUPS, PM_INTERNAL_GROUPS_HEADERS, '#B4A7D6', null);
  // PM設定シートに既定キーの行を用意（値は運用者が記入。ソースにはIDを書かない）
  pmSeedSettingRows();
  // 社内グループシートに既定の役割行を用意（グループIDは運用者が記入）
  pmSeedInternalGroupRows();
  // 既存「タスク管理」に案件連携用の列を非破壊で追加
  pmEnsureTaskColumns();
  console.log('pmSetupSheets: 完了');
}

// PM設定シートに既定キー行を用意（既存値は保持・不足キーのみ追記）。
// 運用者は「値」列にグループID等を記入する。ソースコードにIDを書かないための置き場。
function pmSeedSettingRows() {
  var sheet = pmSheet(PM_SHEET_SETTINGS);
  if (!sheet) return;
  var existing = {};
  pmReadObjects(PM_SHEET_SETTINGS).forEach(function(r) { existing[String(r['キー']).trim()] = true; });
  var defaults = [
    ['ACCOUNTING_GROUP_ID', '', '経理（金額承認）通知先グループID'],
    ['INTERNAL_GROUP_ID',   '', '社内通知先グループID（経理未設定時のフォールバック）'],
    ['PM_APPROVER_USER_IDS','', '金額承認できるLINEユーザーID（カンマ区切り）。空なら制限なし'],
  ];
  defaults.forEach(function(d) {
    if (!existing[d[0]]) pmAppendRowFields(PM_SHEET_SETTINGS, { 'キー': d[0], '値': d[1], '説明': d[2] });
  });
}

// 社内グループシートに既定の役割行を用意（既存行は保持・不足役割のみ追記）。
// 運用者は「グループID」列に LINEグループID（C…）を記入する。記入すると即反映（push不要）。
function pmSeedInternalGroupRows() {
  var sheet = pmSheet(PM_SHEET_INTERNAL_GROUPS);
  if (!sheet) return;
  var existing = {};
  pmReadObjects(PM_SHEET_INTERNAL_GROUPS).forEach(function(r) { existing[String(r['役割']).trim()] = true; });
  var defaults = [
    ['社内',     '放置アラート等の社内通知先。未記入時は Script Property / PM設定の INTERNAL_GROUP_ID'],
    ['経理',     '請求漏れ・入金未確認アラートと金額承認の通知先。未記入時は「社内」へフォールバック'],
    ['日次通知', '毎朝9時の進捗ダイジェスト送信先。未記入時は既定の善波グループ'],
    ['進捗',     '放置アラートの@メンション通知先（進捗管理グループ）。未記入時は「社内」へフォールバック'],
  ];
  defaults.forEach(function(d) {
    if (!existing[d[0]]) pmAppendRowFields(PM_SHEET_INTERNAL_GROUPS, {
      '役割': d[0], 'グループID': '', '表示名': '', '備考': d[1], '更新日時': fmtDT(new Date()),
    });
  });
}

// 既存タスク管理シートへ「関連案件ID」「作成元」列を追加（不足時のみ・右端）
function pmEnsureTaskColumns() {
  var sheet = getSheet('タスク管理');
  if (!sheet) return;
  var idx = pmHeaderIndex(sheet);
  ['関連案件ID', '作成元'].forEach(function(h) {
    if (idx[h] === undefined) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
    }
  });
}

// 毎朝のアラートチェック用トリガーを登録（既存トリガーと独立）
function pmSetupTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'pmCheckAlerts'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('pmCheckAlerts').timeBased().everyDays(1).atHour(9).create();
  console.log('pmSetupTriggers: pmCheckAlerts を毎朝9時に登録');
}

// セットアップ一括（GASエディタから1回実行）
function pmSetupAll() {
  pmSetupSheets();
  pmSetupTriggers();
}
