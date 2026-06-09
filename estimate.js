// ==========================================
// 工程管理表 → 見積書 月別シート自動生成
// ==========================================
// 設定シート（管理用スプシ内）:
//   - 見積連携設定: 工程⇔見積のペア定義
//   - 見積単価設定: 単価ルール
//   - 見積要確認:   自動処理できなかった行のログ
// ==========================================

var ESTIMATE_LINK_SETTING_SHEET = '見積連携設定';
var ESTIMATE_RATE_SHEET         = '見積単価設定';
var ESTIMATE_CONFIRM_SHEET      = '見積要確認';

var ESTIMATE_LINK_HEADERS    = ['設定ID','表示名','工程管理表ID','工程シート名','見積書ID','テンプレートシート名','担当者表示','単価','有効','備考'];
var ESTIMATE_RATE_HEADERS    = ['設定ID','分類','条件キーワード','単価','単位','備考'];
var ESTIMATE_CONFIRM_HEADERS = ['作成日時','設定ID','対象シート名','物件','図面名','分類','理由'];

// テンプレートのセル位置（東邦家具テンプレに合わせて設定）
// 1ページ目（物件集計）  ヘッダー: No=A 名称=B 数量=F 単位=G 単価=H 金額=I 備考=J
var COVER_ITEM_START_ROW = 14;
var COVER_COL_NO         = 1;
var COVER_COL_NAME       = 2;
var COVER_COL_QTY        = 6;
var COVER_COL_UNIT       = 7;
var COVER_COL_AMOUNT     = 9;
var COVER_COL_NOTE       = 10;
var COVER_MAX_ROWS       = 28;  // 14〜41

// 合計行（テンプレ行10）  "税込合計" C列 / "消費税" F列 / "合計"(=小計) I列
var TOTAL_ROW                = 10;
var TOTAL_COL_TAXINCL        = 3;
var TOTAL_COL_TAX            = 6;
var TOTAL_COL_SUBTOTAL       = 9;
// 1ページ目下部の「小計」（行43 列I）
var COVER_SUBTOTAL_ROW       = 43;
var COVER_SUBTOTAL_VALUE_COL = 9;

// 2ページ目以降（明細）
// テンプレ構造: 行44-52がページヘッダー、行53〜84が page2 明細領域
// 行85-93が page3 ヘッダー、行94以降が page3 明細領域（要確認）
var DETAIL_START_ROW      = 53;
var DETAIL_PAGE2_END_ROW  = 84;
var DETAIL_PAGE3_START_ROW = 92;
var DETAIL_PAGE3_END_ROW  = 126;
var DETAIL_COL_NO         = 1;
var DETAIL_COL_NAME       = 2;
var DETAIL_COL_QTY        = 6;
var DETAIL_COL_UNIT       = 7;
var DETAIL_COL_UNITPRICE  = 8;
var DETAIL_COL_AMOUNT     = 9;
var DETAIL_COL_NOTE       = 10;

var TAX_RATE = 0.10;

// ==========================================
// SECTION: 設定シート操作
// ==========================================

// 「見積連携設定」シートを必要に応じて作成・ヘッダー補完
function ensureEstimateLinkSettingsSheet() {
  var ss = getSS();
  var sh = ss.getSheetByName(ESTIMATE_LINK_SETTING_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ESTIMATE_LINK_SETTING_SHEET);
    sh.appendRow(ESTIMATE_LINK_HEADERS);
    sh.getRange(1, 1, 1, ESTIMATE_LINK_HEADERS.length).setFontWeight('bold').setBackground('#D9EAD3');
    return sh;
  }
  // 既存シート：足りないヘッダーを末尾に追加（既存データを壊さない）
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
  var aliasMap = { '表示名': ['名称'] };  // 旧名「名称」も同義として扱う
  ESTIMATE_LINK_HEADERS.forEach(function(h){
    var alias = aliasMap[h] || [];
    var exists = headers.indexOf(h) !== -1 || alias.some(function(a){ return headers.indexOf(a) !== -1; });
    if (!exists) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
      sh.getRange(1, sh.getLastColumn(), 1, 1).setFontWeight('bold').setBackground('#D9EAD3');
      headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim(); });
    }
  });
  return sh;
}

function getEstimateLinkSettings() {
  var sh = ensureEstimateLinkSettingsSheet();
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).trim(); });

  function col() {
    for (var i = 0; i < arguments.length; i++) {
      var idx = headers.indexOf(arguments[i]);
      if (idx !== -1) return idx;
    }
    return -1;
  }
  var iId      = col('設定ID');
  var iName    = col('表示名', '名称');
  var iProcSs  = col('工程管理表ID');
  var iProcSh  = col('工程シート名');
  var iEstSs   = col('見積書ID');
  var iTpl     = col('テンプレートシート名');
  var iAssign  = col('担当者表示');
  var iPrice   = col('単価');
  var iEnabled = col('有効');
  var iNote    = col('備考');

  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[iId]) continue;
    var enabled = iEnabled !== -1 ? r[iEnabled] : true;
    if (enabled === false || String(enabled).toUpperCase() === 'FALSE') continue;
    var price = iPrice !== -1 ? Number(r[iPrice]) : NaN;
    if (!price || isNaN(price)) price = 5000;
    var name = iName !== -1 ? String(r[iName] || '').trim() : '';
    rows.push({
      settingId:             String(r[iId]).trim(),
      displayName:           name,
      processSpreadsheetId:  String(r[iProcSs] || '').trim(),
      processSheetName:      String(r[iProcSh] || '').trim(),
      estimateSpreadsheetId: String(r[iEstSs] || '').trim(),
      templateSheetName:     String(r[iTpl] || '').trim(),
      assigneeLabel:         iAssign !== -1 ? String(r[iAssign] || '').trim() : '',
      unitPrice:             price,
      enabled:               true,
      note:                  iNote !== -1 ? String(r[iNote] || '').trim() : '',
      // 後方互換（既存コード参照用エイリアス）
      name:          name,
      processSsId:   String(r[iProcSs] || '').trim(),
      processSheet:  String(r[iProcSh] || '').trim(),
      estimateSsId:  String(r[iEstSs] || '').trim(),
      templateSheet: String(r[iTpl] || '').trim(),
    });
  }
  return rows;
}

// 有効な設定の一覧テキストを返す（「見積設定一覧」コマンド用）
function listEstimateSettings() {
  var settings = getEstimateLinkSettings();
  if (!settings.length) {
    return '⚠️ 「見積連携設定」シートに有効な設定がありません。';
  }
  var lines = ['登録済みの見積設定です。', ''];
  settings.forEach(function(s){
    var line = '・' + (s.displayName || '(名称未設定)') + '（' + s.settingId + '）';
    if (s.note) line += ' — ' + s.note;
    lines.push(line);
  });
  lines.push('', '使い方：', (settings[0].displayName || settings[0].settingId) + ' 5月分の見積書作って');
  return lines.join('\n');
}

// 「見積設定一覧」コマンド判定
function isEstimateSettingsListRequest(text) {
  if (!text) return false;
  return /見積設定一覧|見積一覧|見積設定リスト/.test(text);
}

function handleEstimateSettingsListRequest(event, messageText) {
  if (!isEstimateSettingsListRequest(messageText)) return false;
  sendLineReply(event.replyToken, listEstimateSettings());
  return true;
}

// 候補一覧テキスト（該当なし／複数候補時に使用）
function listEstimateSettingCandidates() {
  var settings = getEstimateLinkSettings();
  if (!settings.length) {
    return '⚠️ 「見積連携設定」シートに登録がありません。\n設定を1行追加してから再度お試しください。';
  }
  var lines = ['見積書を作成する対象が見つかりませんでした。', '', '登録済みの対象：'];
  settings.forEach(function(s){
    lines.push('・' + (s.displayName || '(名称未設定)') + '（' + s.settingId + '）');
  });
  lines.push('', '例：「' + (settings[0].displayName || settings[0].settingId) + ' 5月分の見積書作って」のように送ってください。');
  return lines.join('\n');
}

function getEstimateRateSettings() {
  var sh = getSheet(ESTIMATE_RATE_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).trim(); });
  var idx = {};
  ESTIMATE_RATE_HEADERS.forEach(function(h){ idx[h] = headers.indexOf(h); });
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[idx['設定ID']]) continue;
    rows.push({
      settingId: String(r[idx['設定ID']]).trim(),
      category:  String(r[idx['分類']]).trim(),
      keyword:   String(r[idx['条件キーワード']]).trim(),
      rate:      Number(r[idx['単価']]) || 0,
      unit:      String(r[idx['単位']]).trim() || '枚',
      note:      String(r[idx['備考']]).trim(),
    });
  }
  return rows;
}

// ==========================================
// SECTION: メッセージ解析
// ==========================================

function isEstimateCreateRequest(text) {
  if (!text) return false;
  // 「見積作って／作成」「見積書作って／作成」など
  if (/見積(書)?(を)?(作って|作成|作っといて|つくって|お願い)/.test(text)) return true;
  if (/見積(書)?/.test(text) && /(作って|作成|つくって|お願い)/.test(text)) return true;
  // 「見積書 上書き／再作成／更新」など（上書きモードの短縮表現）
  if (/見積(書)?/.test(text) && /(上書き|再作成|作り直し|やり直し|更新)/.test(text)) return true;
  return false;
}

function isEstimateOverwriteRequest(text) {
  if (!text) return false;
  return /上書き|上書きして|再作成|作り直し|やり直し/.test(text);
}

// メッセージから設定を検索
//   戻り値: null（該当なし）／ setting オブジェクト1件 ／ setting 配列（複数候補）
function findEstimateSettingFromMessage(text) {
  if (!text) return null;
  var settings = getEstimateLinkSettings();
  if (!settings.length) return null;
  var t = String(text).normalize('NFKC').toLowerCase();

  // 表示名・設定IDのいずれかが含まれる設定を抽出
  var matches = settings.filter(function(s){
    var name = (s.displayName || '').normalize('NFKC').toLowerCase();
    var id   = (s.settingId   || '').toLowerCase();
    return (name && t.indexOf(name) !== -1) || (id && t.indexOf(id) !== -1);
  });

  if (matches.length === 1) return matches[0];
  if (matches.length >= 2) {
    // 表示名長で一意化を試す（より具体的なものを優先）
    matches.sort(function(a, b){ return (b.displayName || '').length - (a.displayName || '').length; });
    if (matches[0].displayName.length > matches[1].displayName.length) return matches[0];
    return matches;
  }
  // 0件で登録が1件しかなければそれを使う
  if (settings.length === 1) return settings[0];
  return null;
}

// 対象年月を抽出。年指定なしは現在年（1月〜3月は要注意 - 必要に応じてここを調整）
function parseEstimateTargetDate(text) {
  if (!text) return null;
  var t = String(text).normalize('NFKC');
  var year  = null;
  var month = null;
  // 「2026年5月」「2026/5」
  var ym1 = t.match(/(\d{4})年(\d{1,2})月/);
  var ym2 = t.match(/(\d{4})[\/\-](\d{1,2})/);
  if (ym1)      { year = parseInt(ym1[1], 10); month = parseInt(ym1[2], 10); }
  else if (ym2) { year = parseInt(ym2[1], 10); month = parseInt(ym2[2], 10); }
  else {
    // 「5月」のみ
    var m = t.match(/(\d{1,2})月/);
    if (m) {
      month = parseInt(m[1], 10);
      year  = new Date().getFullYear();
      // 注意: 1月の指定で12月以降のデータを期待する等、年跨ぎが必要な場合は呼び出し側で調整
    }
  }
  if (!month || month < 1 || month > 12) return null;
  if (!year)  year = new Date().getFullYear();
  return { year: year, month: month };
}

function buildEstimateSheetName(year, month) {
  return String(year).slice(-2) + '-' + month;
}

// ==========================================
// SECTION: 工程管理表 読み取り
// ==========================================

function getProcessRowsBySetting(setting) {
  var ss = SpreadsheetApp.openById(setting.processSsId);
  var sh = ss.getSheetByName(setting.processSheet);
  if (!sh) throw new Error('工程シート「' + setting.processSheet + '」が見つかりません');
  if (sh.getLastRow() < 2) return { headers: [], rows: [] };
  var data = sh.getDataRange().getValues();
  // ヘッダー行検出（「物件」を含む最初の行）
  var headerRow = 0;
  for (var i = 0; i < Math.min(5, data.length); i++) {
    var has = data[i].some(function(c){ return String(c).indexOf('物件') !== -1; });
    if (has) { headerRow = i; break; }
  }
  return { headers: data[headerRow].map(function(h){ return String(h).trim(); }), rows: data.slice(headerRow + 1) };
}

function normalizeProcessRows(processData) {
  var headers = processData.headers;
  function colIdx(name) {
    for (var i = 0; i < headers.length; i++) if (headers[i] === name) return i;
    // ファジー
    for (var j = 0; j < headers.length; j++) if (headers[j] && headers[j].indexOf(name) !== -1) return j;
    return -1;
  }
  var iProp   = colIdx('物件');
  var iDate   = colIdx('提出日');
  var iDraw   = colIdx('図面名');
  var iCnt    = colIdx('枚数');
  var iTime   = colIdx('製図時間');
  var iCat    = colIdx('分類');
  var iNo     = colIdx('物件番号');
  var iAuthor = colIdx('製図者');

  var result = [];
  processData.rows.forEach(function(r){
    var prop = iProp !== -1 ? String(r[iProp] || '').trim() : '';
    if (!prop) return; // 物件名空欄はスキップ
    var dateVal = iDate !== -1 ? r[iDate] : '';
    var month = null, day = null, year = null;
    if (dateVal instanceof Date) {
      year = dateVal.getFullYear(); month = dateVal.getMonth() + 1; day = dateVal.getDate();
    } else {
      var s = String(dateVal || '').normalize('NFKC');
      var m1 = s.match(/(\d{1,2})月(\d{1,2})日/);
      var m2 = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
      if (m2)      { year = parseInt(m2[1],10); month = parseInt(m2[2],10); day = parseInt(m2[3],10); }
      else if (m1) { month = parseInt(m1[1],10); day = parseInt(m1[2],10); }
    }
    result.push({
      property:    prop,
      propertyNo:  iNo     !== -1 ? String(r[iNo] || '').trim() : '',
      submitDate:  dateVal,
      submitYear:  year,
      submitMonth: month,
      submitDay:   day,
      drawingName: iDraw   !== -1 ? String(r[iDraw] || '').trim() : '',
      count:       iCnt    !== -1 ? Number(r[iCnt])  || 0 : 0,
      time:        iTime   !== -1 ? Number(r[iTime]) || 0 : 0,
      category:    iCat    !== -1 ? String(r[iCat] || '').trim() : '',
      author:      iAuthor !== -1 ? String(r[iAuthor] || '').trim() : '',
    });
  });
  return result;
}

function filterRowsByTargetMonth(rows, year, month) {
  return rows.filter(function(r){
    if (r.submitMonth !== month) return false;
    if (r.submitYear !== null && r.submitYear !== year) return false;
    return true;
  });
}

function groupRowsByProperty(rows) {
  var map = {};
  var order = [];
  rows.forEach(function(r){
    var key = r.propertyNo + '|' + r.property;
    if (!map[key]) { map[key] = { propertyNo: r.propertyNo, property: r.property, rows: [] }; order.push(key); }
    map[key].rows.push(r);
  });
  return order.map(function(k){ return map[k]; });
}

// ==========================================
// SECTION: 単価判定
// ==========================================

function getEstimateRate(settingId, category, drawingName, rateSettings) {
  if (!rateSettings) rateSettings = getEstimateRateSettings();
  var cands = rateSettings.filter(function(r){
    if (r.settingId && r.settingId !== settingId) return false;
    if (r.category && r.category !== category)   return false;
    return true;
  });
  if (!cands.length) return null;
  // キーワード一致を優先（長いキーワードほど具体的）
  var draw = String(drawingName || '');
  var specific = cands.filter(function(r){ return r.keyword && r.keyword !== 'デフォルト' && draw.indexOf(r.keyword) !== -1; });
  if (specific.length) {
    specific.sort(function(a, b){ return b.keyword.length - a.keyword.length; });
    return specific[0];
  }
  // デフォルト
  var def = cands.filter(function(r){ return !r.keyword || r.keyword === 'デフォルト'; });
  if (def.length) return def[0];
  return null;
}

// ==========================================
// SECTION: 集計計算
// ==========================================

// 計算ルール（共通版）：
//   金額 = 製図時間 × setting.unitPrice（新規・修正とも共通）
//   修正: 数量=製図時間 / 単位=H / 単価=unitPrice / 金額=製図時間×unitPrice
//   新規: 数量=枚数      / 単位=枚 / 単価=空欄    / 金額=製図時間×unitPrice
var HOURLY_RATE = 5000; // フォールバックデフォルト

// 同じ図面名＋分類の行を統合する（複数日の同一作業を1行にまとめる）
// 統合した場合は _mergedCount に件数をセットし、名称に ×N を付加できるようにする
function mergeRowsByDrawing(rows) {
  var map = {};
  var order = [];
  rows.forEach(function(r) {
    var key = (r.drawingName || '') + '\0' + (r.category || '');
    if (!map[key]) {
      map[key] = { rows: [], totalTime: 0, totalCount: 0 };
      order.push(key);
    }
    map[key].rows.push(r);
    map[key].totalTime  += (r.time  || 0);
    map[key].totalCount += (r.count || 0);
  });

  var merged = [];
  order.forEach(function(key) {
    var g = map[key];
    if (g.rows.length === 1) {
      merged.push(g.rows[0]);
      return;
    }
    var base = {};
    for (var k in g.rows[0]) base[k] = g.rows[0][k];
    base.time          = g.totalTime;
    base.count         = g.totalCount;
    base._mergedCount  = g.rows.length;
    merged.push(base);
  });
  return merged;
}

function calculateEstimateData(groups, setting) {
  var hourlyRate = (setting && setting.unitPrice) || HOURLY_RATE;
  var settingId  = setting && setting.settingId;
  var properties   = [];
  var confirmItems = [];
  var subtotal     = 0;

  groups.forEach(function(g){
    var totalAmount = 0;
    var newSheets   = 0;
    var fixHours    = 0;
    var detailItems = [];

    var mergedRows = mergeRowsByDrawing(g.rows);
    mergedRows.forEach(function(r){
      var detail  = { name: '', qty: '', unit: '', unitPrice: '', amount: 0 };
      var confirm = '';
      var suffix  = r._mergedCount > 1 ? ' ×' + r._mergedCount : '';

      if (!r.category) {
        confirm = '分類が空欄';
        detail.name = r.drawingName || '(図面名空欄)';
      } else if (r.category === '修正') {
        if (!r.time) {
          confirm = '製図時間が空欄';
          detail.name = (r.drawingName || '') + ' 修正' + suffix;
        } else {
          detail.name      = (r.drawingName || '') + ' 修正' + suffix;
          detail.qty       = r.time;
          detail.unit      = 'H';
          // 時間が1Hのときは単価欄を表示しない（数量=金額になるため重複表示を避ける）
          detail.unitPrice = r.time === 1 ? '' : hourlyRate;
          detail.amount    = r.time * hourlyRate;
          fixHours += r.time;
        }
      } else if (r.category === '新規') {
        if (!r.time) {
          confirm = '製図時間が空欄（金額計算不可）';
          detail.name = (r.drawingName || '') + ' 新規' + suffix;
          detail.qty  = r.count || '';
          detail.unit = '枚';
        } else {
          detail.name      = (r.drawingName || '') + ' 新規' + suffix;
          detail.qty       = r.count || '';
          detail.unit      = '枚';
          detail.unitPrice = ''; // 新規は単価欄空欄（仕様）
          detail.amount    = r.time * hourlyRate;
          newSheets += r.count || 0;
        }
      } else if (r.category === '検討のみ' || r.category === 'WBF') {
        confirm = r.category + '：請求対象不明';
        detail.name = (r.drawingName || '') + ' ' + r.category + suffix;
        detail.qty  = r.count || r.time || '';
        detail.unit = r.count ? '枚' : (r.time ? 'H' : '');
      } else {
        confirm = '未対応の分類「' + r.category + '」';
        detail.name = (r.drawingName || '') + ' ' + r.category + suffix;
      }

      if (confirm) {
        confirmItems.push({
          settingId:   settingId,
          property:    g.property,
          drawingName: r.drawingName,
          category:    r.category,
          reason:      confirm,
        });
      }

      detailItems.push(detail);
      totalAmount += detail.amount;
    });

    // 名称：【物件番号】物件名（新規〇枚／修正〇時間）
    var suffixParts = [];
    if (newSheets > 0) suffixParts.push('新規' + newSheets + '枚');
    if (fixHours  > 0) suffixParts.push('修正' + fixHours + '時間');
    var coverName = '【' + (g.propertyNo || '') + '】' + g.property +
                    (suffixParts.length ? '（' + suffixParts.join('／') + '）' : '');

    properties.push({
      propertyNo:  g.propertyNo,
      property:    g.property,
      coverName:   coverName,
      amount:      totalAmount,
      newSheets:   newSheets,
      fixHours:    fixHours,
      detailItems: detailItems,
    });
    subtotal += totalAmount;
  });

  var tax   = Math.round(subtotal * TAX_RATE);
  var total = subtotal + tax;
  return {
    properties:   properties,
    confirmItems: confirmItems,
    subtotal:     subtotal,
    tax:          tax,
    total:        total,
    rowCount:     groups.reduce(function(s, g){ return s + g.rows.length; }, 0),
    groupCount:   groups.length,
  };
}

// ==========================================
// SECTION: シート操作
// ==========================================

function copyEstimateTemplateSheet(estimateSpreadsheetId, templateSheetName, newSheetName, overwrite) {
  var ss = SpreadsheetApp.openById(estimateSpreadsheetId);
  var existing = ss.getSheetByName(newSheetName);
  if (existing) {
    if (!overwrite) {
      return { existing: true, ss: ss, sheet: existing };
    }
    ss.deleteSheet(existing);
  }
  var tpl = ss.getSheetByName(templateSheetName);
  if (!tpl) throw new Error('テンプレートシート「' + templateSheetName + '」が見つかりません');
  var copy = tpl.copyTo(ss);
  copy.setName(newSheetName);
  // テンプレートシートの直後（右隣）に配置
  ss.setActiveSheet(copy);
  ss.moveActiveSheet(tpl.getIndex() + 1);
  return { existing: false, ss: ss, sheet: copy };
}

// シート全体から指定テキストを含むセルを探す
function findCellByText(sheet, searchText) {
  var data = sheet.getDataRange().getValues();
  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < data[r].length; c++) {
      if (String(data[r][c] || '').indexOf(searchText) !== -1) return { row: r + 1, col: c + 1 };
    }
  }
  return null;
}

function writeEstimateCover(targetSheet, estimateData, setting, year, month) {
  // J3: 作成日（今日の日付 yyyy/MM/dd）
  var tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  targetSheet.getRange('J3').setValue(Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd'));
  // A5: 対象月（例: 5月）
  if (month) targetSheet.getRange('A5').setValue(month);

  var startRow = COVER_ITEM_START_ROW;
  estimateData.properties.forEach(function(p, i){
    if (i >= COVER_MAX_ROWS) return;
    var row = startRow + i;
    targetSheet.getRange(row, COVER_COL_NO).setValue(i + 1);
    targetSheet.getRange(row, COVER_COL_NAME).setValue(p.coverName);
    targetSheet.getRange(row, COVER_COL_QTY).setValue(1);
    targetSheet.getRange(row, COVER_COL_UNIT).setValue('式');
    targetSheet.getRange(row, COVER_COL_AMOUNT).setValue(p.amount);
    targetSheet.getRange(row, COVER_COL_NOTE).setValue(i === 0 ? (setting.assigneeLabel || '') : '同上');
  });
}

function writeEstimateDetails(targetSheet, estimateData, setting) {
  // page2: 53-84, page3: 94-125 を書込領域として、ページヘッダー行はスキップ
  var rowState = { row: DETAIL_START_ROW, pageEnd: DETAIL_PAGE2_END_ROW, nextStart: DETAIL_PAGE3_START_ROW, nextEnd: DETAIL_PAGE3_END_ROW, overflow: false };

  function advance() {
    rowState.row++;
    if (rowState.row > rowState.pageEnd) {
      if (rowState.nextStart) {
        rowState.row     = rowState.nextStart;
        rowState.pageEnd = rowState.nextEnd;
        rowState.nextStart = null;
      } else {
        rowState.overflow = true;
      }
    }
  }

  estimateData.properties.forEach(function(p, pIdx){
    if (rowState.overflow) return;
    // 物件見出し
    targetSheet.getRange(rowState.row, DETAIL_COL_NO).setValue(pIdx + 1);
    targetSheet.getRange(rowState.row, DETAIL_COL_NAME).setValue(p.property);
    advance();
    var seqInProperty = 0;
    p.detailItems.forEach(function(d){
      if (rowState.overflow) return;
      seqInProperty++;
      targetSheet.getRange(rowState.row, DETAIL_COL_NAME).setValue(d.name);
      targetSheet.getRange(rowState.row, DETAIL_COL_QTY).setValue(d.qty === '' ? '' : d.qty);
      targetSheet.getRange(rowState.row, DETAIL_COL_UNIT).setValue(d.unit || '');
      targetSheet.getRange(rowState.row, DETAIL_COL_UNITPRICE).setValue(d.unitPrice === '' ? '' : d.unitPrice);
      targetSheet.getRange(rowState.row, DETAIL_COL_AMOUNT).setValue(d.amount || '');
      advance();
    });
    advance(); // 物件間の空行
  });

  if (rowState.overflow) {
    console.warn('明細件数が page2+page3 の領域を超えました。テンプレを拡張してください。');
  }
}

// 合計反映：テンプレ行10（税込合計/消費税/合計＝小計）と 1ページ目下部の小計（行43）
function writeEstimateTotals(targetSheet, estimateData) {
  targetSheet.getRange(TOTAL_ROW, TOTAL_COL_TAXINCL).setValue(estimateData.total);
  targetSheet.getRange(TOTAL_ROW, TOTAL_COL_TAX).setValue(estimateData.tax);
  targetSheet.getRange(TOTAL_ROW, TOTAL_COL_SUBTOTAL).setValue(estimateData.subtotal);
  // 1ページ目下部の小計
  targetSheet.getRange(COVER_SUBTOTAL_ROW, COVER_SUBTOTAL_VALUE_COL).setValue(estimateData.subtotal);
}

function writeEstimateConfirmItems(confirmItems, targetSheetName) {
  if (!confirmItems || !confirmItems.length) return;
  var sh = getSheet(ESTIMATE_CONFIRM_SHEET);
  if (!sh) { sh = getSS().insertSheet(ESTIMATE_CONFIRM_SHEET); sh.appendRow(ESTIMATE_CONFIRM_HEADERS); }
  var now = fmtDT(new Date());
  confirmItems.forEach(function(c){
    sh.appendRow([now, c.settingId, targetSheetName, c.property, c.drawingName, c.category, c.reason]);
  });
}

// ==========================================
// SECTION: メイン処理
// ==========================================

// 設定 + 年月から見積書を作成する共通処理（LINE非依存）
//   戻り値: { kind: 'success'|'exists'|'noData', sheetName, spreadsheetUrl, data, year, month, setting }
function createEstimateFromProcessSheet(setting, year, month, overwrite) {
  // 必須フィールド検証
  var missing = [];
  if (!setting.processSpreadsheetId)  missing.push('工程管理表ID');
  if (!setting.processSheetName)      missing.push('工程シート名');
  if (!setting.estimateSpreadsheetId) missing.push('見積書ID');
  if (!setting.templateSheetName)     missing.push('テンプレートシート名');
  if (missing.length) {
    return { kind: 'configError', setting: setting, missing: missing };
  }

  var sheetName = buildEstimateSheetName(year, month);

  // 既存チェック（上書きモードでない場合のみ中止）
  var preCheckSs;
  try {
    preCheckSs = SpreadsheetApp.openById(setting.estimateSpreadsheetId);
  } catch (e) {
    return { kind: 'openError', setting: setting, target: '見積書', id: setting.estimateSpreadsheetId, message: e.message };
  }
  if (!overwrite && preCheckSs.getSheetByName(sheetName)) {
    return { kind: 'exists', sheetName: sheetName, spreadsheetUrl: preCheckSs.getUrl(), setting: setting };
  }

  var processData;
  try {
    processData = getProcessRowsBySetting(setting);
  } catch (e) {
    return { kind: 'openError', setting: setting, target: '工程管理表', id: setting.processSpreadsheetId, message: e.message };
  }
  var allRows     = normalizeProcessRows(processData);
  var targetRows  = filterRowsByTargetMonth(allRows, year, month);
  if (!targetRows.length) {
    return { kind: 'noData', sheetName: sheetName, year: year, month: month, setting: setting };
  }

  var groups = groupRowsByProperty(targetRows);
  var data   = calculateEstimateData(groups, setting);

  var copyRes = copyEstimateTemplateSheet(setting.estimateSpreadsheetId, setting.templateSheetName, sheetName, overwrite);
  var target  = copyRes.sheet;

  writeEstimateCover(target, data, setting, year, month);
  writeEstimateDetails(target, data, setting);
  writeEstimateTotals(target, data);
  writeEstimateConfirmItems(data.confirmItems, sheetName);

  return {
    kind:           'success',
    sheetName:      sheetName,
    spreadsheetUrl: copyRes.ss.getUrl(),
    data:           data,
    year:           year,
    month:          month,
    setting:        setting,
  };
}

function handleEstimateCreateRequest(event, messageText) {
  try {
    if (!isEstimateCreateRequest(messageText)) return false;

    // 設定特定
    var matched = findEstimateSettingFromMessage(messageText);
    if (!matched) {
      sendLineReply(event.replyToken, listEstimateSettingCandidates());
      return true;
    }
    if (Array.isArray(matched)) {
      var lines = ['対象が複数該当しました。明示してもう一度送信してください。', ''];
      matched.forEach(function(s){ lines.push('・' + s.displayName + '（' + s.settingId + '）'); });
      lines.push('', '例：「' + matched[0].settingId + ' 5月分の見積書作って」');
      sendLineReply(event.replyToken, lines.join('\n'));
      return true;
    }
    var setting = matched;

    // 対象年月
    var ym = parseEstimateTargetDate(messageText);
    if (!ym) {
      sendLineReply(event.replyToken, '⚠️ 対象年月が読み取れませんでした。「5月分」「2026年5月」のように指定してください。');
      return true;
    }

    var overwrite = isEstimateOverwriteRequest(messageText);
    var result = createEstimateFromProcessSheet(setting, ym.year, ym.month, overwrite);
    if (result.kind === 'configError') {
      sendLineReply(event.replyToken, '⚠️ 設定「' + setting.displayName + '（' + setting.settingId + '）」に必須項目が未設定です：\n・' + result.missing.join('\n・') + '\n\n「見積連携設定」シートを確認してください。');
      return true;
    }
    if (result.kind === 'openError') {
      sendLineReply(event.replyToken, '⚠️ ' + result.target + 'を開けませんでした。\nID：' + result.id + '\n権限・IDが正しいか確認してください。\n\n詳細：' + result.message);
      return true;
    }
    if (result.kind === 'exists') {
      sendLineReply(event.replyToken, '⚠️ 既に「' + result.sheetName + '」の見積書シートがあります。\n\n上書きする場合は「上書き」を含めて再送してください。\n例：「' + setting.displayName + ' ' + ym.month + '月分の見積書を上書きして作って」\n\n見積書：' + result.spreadsheetUrl);
      return true;
    }
    if (result.kind === 'noData') {
      sendLineReply(event.replyToken, '⚠️ 工程管理表に「' + result.year + '年' + result.month + '月」のデータがありませんでした。');
      return true;
    }
    sendLineReply(event.replyToken, formatEstimateCreatedReply({
      setting: setting, year: ym.year, month: ym.month, sheetName: result.sheetName,
      data: result.data, spreadsheetUrl: result.spreadsheetUrl,
    }));
    return true;
  } catch (err) {
    console.error('handleEstimateCreateRequest error:', err.message, err.stack);
    try { sendLineReply(event.replyToken, '⚠️ 見積書作成中にエラーが発生しました：' + err.message); } catch (e) {}
    return true;
  }
}

function formatEstimateCreatedReply(r) {
  var lines = [];
  lines.push(r.setting.name + ' ' + r.sheetName + ' の見積書を作成しました。');
  lines.push('');
  lines.push('対象：' + r.year + '年' + r.month + '月分');
  lines.push('対象件数：' + r.data.rowCount + '件');
  lines.push('対象物件数：' + r.data.groupCount + '件');
  lines.push('小計：¥' + r.data.subtotal.toLocaleString('ja-JP'));
  lines.push('消費税：¥' + r.data.tax.toLocaleString('ja-JP'));
  lines.push('税込合計：¥' + r.data.total.toLocaleString('ja-JP'));
  if (r.data.confirmItems.length) {
    lines.push('');
    lines.push('要確認：' + r.data.confirmItems.length + '件');
    r.data.confirmItems.slice(0, 10).forEach(function(c){
      lines.push('・' + c.property + ' / ' + c.drawingName + ' / ' + c.category + '：' + c.reason);
    });
    if (r.data.confirmItems.length > 10) lines.push('…ほか ' + (r.data.confirmItems.length - 10) + ' 件');
  }
  lines.push('');
  lines.push('見積書：' + r.spreadsheetUrl);
  return lines.join('\n');
}

// ==========================================
// SECTION: セットアップ
// ==========================================

function setupEstimateSheets() {
  var ss = getSS();
  [ESTIMATE_LINK_SETTING_SHEET, ESTIMATE_RATE_SHEET, ESTIMATE_CONFIRM_SHEET].forEach(function(name){
    if (ss.getSheetByName(name)) return;
    var sh = ss.insertSheet(name);
    var headers = name === ESTIMATE_LINK_SETTING_SHEET ? ESTIMATE_LINK_HEADERS
                : name === ESTIMATE_RATE_SHEET         ? ESTIMATE_RATE_HEADERS
                : ESTIMATE_CONFIRM_HEADERS;
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#D9EAD3');
  });
  console.log('見積関連シート作成完了:', ESTIMATE_LINK_SETTING_SHEET, '/', ESTIMATE_RATE_SHEET, '/', ESTIMATE_CONFIRM_SHEET);
}

// 東邦家具のデフォルト設定を投入（存在しなければ追加のみ・既存値は上書きしない）
function setupEstimateForToho() {
  setupEstimateSheets();

  var link = getSheet(ESTIMATE_LINK_SETTING_SHEET);
  var data = link.getDataRange().getValues();
  var hasToho = false;
  for (var i = 1; i < data.length; i++) if (String(data[i][0]).trim() === 'toho') { hasToho = true; break; }
  if (!hasToho) {
    link.appendRow(['toho', '東邦家具', '1MHHVpkdNjTihse-CX4sOZbeGKi71ncKD', '工程管理表',
                    '1Fr5_KPcvTGa7WP7PpQ7glIQglaZgeuv_', 'テンプレート', 'YY-M', '担当者：米田様', true]);
    console.log('「東邦家具」連携設定を追加しました');
  } else {
    console.log('「東邦家具」連携設定は既に存在します');
  }

  var rate = getSheet(ESTIMATE_RATE_SHEET);
  var rateData = rate.getDataRange().getValues();
  function ensureRate(category, keyword, price, unit, note) {
    for (var i = 1; i < rateData.length; i++) {
      if (String(rateData[i][0]).trim() === 'toho' && String(rateData[i][1]).trim() === category && String(rateData[i][2]).trim() === keyword) return;
    }
    rate.appendRow(['toho', category, keyword, price, unit, note]);
  }
  ensureRate('修正', 'デフォルト', 5000,  'H', '修正は時間単価');
  ensureRate('新規', 'デフォルト', 10000, '枚', '新規デフォルト');
  ensureRate('新規', 'カウンター', 15000, '枚', 'カウンター系');
  ensureRate('新規', '建具',       20000, '枚', '建具系');
  console.log('東邦家具の単価設定を追加（既存はそのまま）');
}

// ==========================================
// SECTION: 診断
// ==========================================

// GASエディタ実行ボタン用：東邦家具のテンプレを診断
function diagnoseEstimateTemplateForToho() { diagnoseEstimateTemplate('toho'); }

// テンプレの指定行範囲を表示（2ページ目/3ページ目の構造確認用）
function dumpEstimateTemplateRows(fromRow, toRow) {
  var settings = getEstimateLinkSettings();
  var s = null;
  for (var i = 0; i < settings.length; i++) if (settings[i].settingId === 'toho') { s = settings[i]; break; }
  if (!s) { console.error('toho 設定なし'); return; }
  var ss = SpreadsheetApp.openById(s.estimateSsId);
  var tpl = ss.getSheetByName(s.templateSheet);
  fromRow = fromRow || 40;
  toRow = Math.min(toRow || 133, tpl.getLastRow());
  var d = tpl.getRange(fromRow, 1, toRow - fromRow + 1, tpl.getLastColumn()).getValues();
  for (var i = 0; i < d.length; i++) console.log('行' + (fromRow + i) + ':', JSON.stringify(d[i]));
}
function dumpEstimateTemplatePage2() { dumpEstimateTemplateRows(40, 90); }
function dumpEstimateTemplatePage3() { dumpEstimateTemplateRows(85, 133); }

// 見積書スプシのシート一覧を表示（テンプレ名確認用）
function listEstimateSheetsForToho() {
  var settings = getEstimateLinkSettings();
  var s = null;
  for (var i = 0; i < settings.length; i++) if (settings[i].settingId === 'toho') { s = settings[i]; break; }
  if (!s) { console.error('toho 設定なし'); return; }
  var ss = SpreadsheetApp.openById(s.estimateSsId);
  console.log('見積書スプシ:', ss.getName());
  ss.getSheets().forEach(function(sh){ console.log(' -', JSON.stringify(sh.getName())); });
}

// 工程管理表スプシのシート一覧
function listProcessSheetsForToho() {
  var settings = getEstimateLinkSettings();
  var s = null;
  for (var i = 0; i < settings.length; i++) if (settings[i].settingId === 'toho') { s = settings[i]; break; }
  if (!s) { console.error('toho 設定なし'); return; }
  var ss = SpreadsheetApp.openById(s.processSsId);
  console.log('工程管理表スプシ:', ss.getName());
  ss.getSheets().forEach(function(sh){ console.log(' -', JSON.stringify(sh.getName())); });
}

// 現在の DETAIL_START_ROW 等の値を確認（GAS側のコードが最新か確認用）
function showEstimateConstants() {
  console.log('DETAIL_START_ROW:',       DETAIL_START_ROW);
  console.log('DETAIL_PAGE2_END_ROW:',   DETAIL_PAGE2_END_ROW);
  console.log('DETAIL_PAGE3_START_ROW:', DETAIL_PAGE3_START_ROW);
  console.log('DETAIL_PAGE3_END_ROW:',   DETAIL_PAGE3_END_ROW);
  console.log('COVER_ITEM_START_ROW:',   COVER_ITEM_START_ROW);
  console.log('HOURLY_RATE:',            HOURLY_RATE);
}

// 既存の toho 行を正しい値で上書き（ヘッダー名ベース・列順非依存）
function fixTohoEstimateIds() {
  var sh = ensureEstimateLinkSettingsSheet();
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).trim(); });

  function colNum(name, alias) {
    var i = headers.indexOf(name);
    if (i === -1 && alias) i = headers.indexOf(alias);
    return i === -1 ? -1 : (i + 1);
  }
  var cId       = colNum('設定ID');
  var cName     = colNum('表示名', '名称');
  var cProcSs   = colNum('工程管理表ID');
  var cProcSh   = colNum('工程シート名');
  var cEstSs    = colNum('見積書ID');
  var cTpl      = colNum('テンプレートシート名');
  var cAssign   = colNum('担当者表示');
  var cPrice    = colNum('単価');
  var cEnabled  = colNum('有効');
  var cNote     = colNum('備考');

  var v = {
    settingId:'toho', name:'東邦家具',
    processSs:'1coS4GtIZeeqIZYVGdSHpLUKW4ewyINE2KcIG3AV4q5g', processSh:'工程管理表',
    estSs:'1bwngYD3yaI1gyqdbENYR6BizpworX0_eEA7Bxq0fajY', tpl:'原図',
    assign:'担当者：米田様', price:5000, enabled:true, note:'東邦家具用'
  };

  // 既存行を探す
  for (var i = 1; i < data.length; i++) {
    if (cId !== -1 && String(data[i][cId - 1]).trim() === 'toho') {
      if (cName    !== -1) sh.getRange(i + 1, cName).setValue(v.name);
      if (cProcSs  !== -1) sh.getRange(i + 1, cProcSs).setValue(v.processSs);
      if (cProcSh  !== -1) sh.getRange(i + 1, cProcSh).setValue(v.processSh);
      if (cEstSs   !== -1) sh.getRange(i + 1, cEstSs).setValue(v.estSs);
      if (cTpl     !== -1) sh.getRange(i + 1, cTpl).setValue(v.tpl);
      if (cAssign  !== -1) sh.getRange(i + 1, cAssign).setValue(v.assign);
      if (cPrice   !== -1) sh.getRange(i + 1, cPrice).setValue(v.price);
      if (cEnabled !== -1) sh.getRange(i + 1, cEnabled).setValue(v.enabled);
      if (cNote    !== -1 && !data[i][cNote - 1]) sh.getRange(i + 1, cNote).setValue(v.note);
      console.log('toho 行を更新（ヘッダー名ベース）');
      return;
    }
  }
  // 新規追加
  var row = new Array(headers.length).fill('');
  if (cId      !== -1) row[cId - 1] = v.settingId;
  if (cName    !== -1) row[cName - 1] = v.name;
  if (cProcSs  !== -1) row[cProcSs - 1] = v.processSs;
  if (cProcSh  !== -1) row[cProcSh - 1] = v.processSh;
  if (cEstSs   !== -1) row[cEstSs - 1] = v.estSs;
  if (cTpl     !== -1) row[cTpl - 1] = v.tpl;
  if (cAssign  !== -1) row[cAssign - 1] = v.assign;
  if (cPrice   !== -1) row[cPrice - 1] = v.price;
  if (cEnabled !== -1) row[cEnabled - 1] = v.enabled;
  if (cNote    !== -1) row[cNote - 1] = v.note;
  sh.appendRow(row);
  console.log('toho 行を追加（ヘッダー名ベース）');
}

// テンプレートシートの構造を確認（セル位置調整のため）
function diagnoseEstimateTemplate(settingId) {
  var settings = getEstimateLinkSettings();
  var setting = null;
  for (var i = 0; i < settings.length; i++) if (settings[i].settingId === (settingId || 'toho')) { setting = settings[i]; break; }
  if (!setting) { console.error('設定ID未登録:', settingId); return; }

  var ss = SpreadsheetApp.openById(setting.estimateSsId);
  var tpl = ss.getSheetByName(setting.templateSheet);
  if (!tpl) { console.error('テンプレシート未存在:', setting.templateSheet); return; }
  console.log('対象テンプレ:', ss.getName(), '/', tpl.getName(), '行数:', tpl.getLastRow(), '列数:', tpl.getLastColumn());

  // 「小計」「消費税」「合計」「税込合計」の位置
  ['小計','消費税','税込合計','合計','名称','品名','項目'].forEach(function(label){
    var loc = findCellByText(tpl, label);
    if (loc) console.log('  「' + label + '」位置: 行' + loc.row + ' 列' + loc.col + ' (R' + loc.row + 'C' + loc.col + ')');
  });

  // 先頭25行を表示
  var n = Math.min(25, tpl.getLastRow());
  if (n > 0) {
    var d = tpl.getRange(1, 1, n, tpl.getLastColumn()).getValues();
    for (var i = 0; i < n; i++) {
      console.log('行' + (i + 1) + ':', JSON.stringify(d[i]));
    }
  }
}

// 計算根拠の詳細ログ（どの行がいくらになったか・どの単価ルールが当たったか）
function explainEstimateForToho(month) {
  month = month || 5;
  var setting = findEstimateSettingFromMessage('東邦家具');
  if (!setting) { console.error('toho 設定なし'); return; }
  var ym = { year: new Date().getFullYear(), month: month };
  var processData = getProcessRowsBySetting(setting);
  var all = normalizeProcessRows(processData);
  var target = filterRowsByTargetMonth(all, ym.year, ym.month);
  var rateSettings = getEstimateRateSettings();

  console.log('=== ' + ym.year + '年' + ym.month + '月 計算根拠 ===');
  console.log('対象行数:', target.length);
  var subtotal = 0;
  target.forEach(function(r, i){
    var amount = 0, qty = '', unit = '', up = '', reason = '';
    if (r.category === '修正') {
      if (r.time) { qty = r.time; unit = 'H'; up = HOURLY_RATE; amount = r.time * HOURLY_RATE; reason = '修正'; }
      else { reason = '修正だが製図時間空欄'; }
    } else if (r.category === '新規') {
      if (r.time) { qty = r.count; unit = '枚'; up = ''; amount = r.time * HOURLY_RATE; reason = '新規(時間' + r.time + 'H×' + HOURLY_RATE + ')'; }
      else { reason = '新規だが製図時間空欄'; }
    } else {
      reason = '分類「' + r.category + '」自動計算なし';
    }
    console.log((i + 1) + '. ' + r.property + ' / ' + r.drawingName + ' / ' + r.category +
                ' → 数量' + qty + unit + ' / 単価' + up + ' / 金額¥' + amount + '  [' + reason + ']');
    subtotal += amount;
  });
  console.log('小計:', subtotal, '消費税:', Math.round(subtotal * TAX_RATE), '税込:', subtotal + Math.round(subtotal * TAX_RATE));
}
function explainEstimateForToho_May() { explainEstimateForToho(5); }

// LINE送信を伴わないドライラン
function testEstimateDryRun() {
  var setting = findEstimateSettingFromMessage('東邦家具 5月分の見積書作って');
  console.log('setting:', setting && setting.name);
  var ym = parseEstimateTargetDate('東邦家具 5月分の見積書作って');
  console.log('ym:', ym);
  if (!setting || !ym) return;
  var processData = getProcessRowsBySetting(setting);
  console.log('工程ヘッダー:', processData.headers);
  console.log('工程行数:', processData.rows.length);
  var all = normalizeProcessRows(processData);
  console.log('正規化後:', all.length);
  var filtered = filterRowsByTargetMonth(all, ym.year, ym.month);
  console.log('対象月:', filtered.length);
  var groups = groupRowsByProperty(filtered);
  console.log('物件数:', groups.length);
  var data = calculateEstimateData(groups, getEstimateRateSettings(), setting.settingId);
  console.log('小計:', data.subtotal, '税:', data.tax, '合計:', data.total);
  console.log('要確認:', data.confirmItems.length, '件');
  data.confirmItems.slice(0, 5).forEach(function(c){ console.log(' -', c.property, c.drawingName, c.category, c.reason); });
}
