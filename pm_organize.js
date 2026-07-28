// ==========================================
// スプレッドシート整理（表示整頓・ダッシュボード・projects装飾）
//   方針（重要）:
//   - 既存データ・数式・シート名・列順は壊さない（破壊的操作なし）。
//   - 列の物理並べ替えはしない。固定（freeze）・非表示・列グループのみで整える。
//   - projects は列順が環境で異なりうる（setupSheetが不足列を右端追加するため）。
//     よって列はすべて「ヘッダー名→位置」を実行時に解決する。
//   - 冪等：何度実行しても重複しない（条件付き書式は全置換、グループは一旦解除して再作成）。
//   依存: pm_core.js(getSS/getSheet/pmSheet/pmHeaderIndex/pmReadObjects/pmParseDate/
//         pmTodayYmd/pmIsBlank/PM_SHEET_*), pm_alert.js(pmIsProjectClosed),
//         既存 fmtDate/fmtDT
// ==========================================

// 先頭に並べて表示するシート（存在するもののみ・この順）
var PM_ORG_SHOW_ORDER = [
  'ダッシュボード',
  PM_SHEET_PROJECTS,        // projects
  'タスク管理',
  'スケジュール管理',
  PM_SHEET_PENDING,         // project_pending_updates
  '見積要確認',
];

// 削除せず非表示にするシステム処理用シート（存在するもののみ）
var PM_ORG_HIDE = [
  '仮タスク', 'メッセージログ', PM_SHEET_LOGS, PM_SHEET_ALERTS, PM_SHEET_CALENDAR,
  '完了タスク', '会社情報', '見積連携設定', '見積単価設定',
  // アプリ↔GAS連携のシステムシート（pm_jobs.js）。アプリからのみ触る想定なのでスプシ上は隠す
  '見積ジョブ', 'pmジョブ', 'アプリ権限',
];

var PM_DASHBOARD = 'ダッシュボード';

// 共通ヘッダー装飾（全シート統一）
var PM_HEADER_BG    = '#37474F'; // ダークスレート
var PM_HEADER_FG    = '#FFFFFF';
var PM_HEADER_FONT  = 'Arial';

// ==========================================
// エントリポイント
// ==========================================
function pmOrganizeSpreadsheet() {
  var ss = getSS();

  // 1) ダッシュボード作成＋集計（先頭）
  try { pmRefreshDashboard(); } catch (e) { console.error('pmOrganize dashboard:', e.message); }

  // 2) 表示シートを前面へ並べる＋表示状態に
  try { pmOrderAndShowSheets(ss); } catch (e) { console.error('pmOrganize order:', e.message); }

  // 3) システムシートを非表示（存在時のみ・削除しない）
  try { pmHideSystemSheets(ss); } catch (e) { console.error('pmOrganize hide:', e.message); }

  // 4) projects を見やすく装飾
  try { pmFormatProjectsSheet(); } catch (e) { console.error('pmOrganize projects:', e.message); }

  // 5) 全表示シート共通：ヘッダー固定＋装飾統一
  try { pmUnifyHeaders(ss); } catch (e) { console.error('pmOrganize headers:', e.message); }

  // 仕上げにダッシュボードをアクティブ化
  try { var d = ss.getSheetByName(PM_DASHBOARD); if (d) ss.setActiveSheet(d); } catch (e) {}
  console.log('pmOrganizeSpreadsheet: 完了');
}

// ==========================================
// シートの並べ替え・表示/非表示
// ==========================================
function pmOrderAndShowSheets(ss) {
  var pos = 1;
  PM_ORG_SHOW_ORDER.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;            // 無いものは飛ばす（見積要確認など）
    if (sheet.isSheetHidden()) sheet.showSheet();
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(pos);       // 1始まりの位置へ（冪等：毎回同じ位置に固定）
    pos++;
  });
}

function pmHideSystemSheets(ss) {
  // 表示対象は絶対に隠さない（万一の重複指定対策）
  var keepVisible = {};
  PM_ORG_SHOW_ORDER.forEach(function(n) { keepVisible[n] = true; });
  PM_ORG_HIDE.forEach(function(name) {
    if (keepVisible[name]) return;
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    // 全シートが隠れる事故を防ぐ（最低1枚は可視を保証）
    var visibleCount = ss.getSheets().filter(function(s) { return !s.isSheetHidden(); }).length;
    if (visibleCount <= 1) return;
    if (!sheet.isSheetHidden()) sheet.hideSheet();
  });
}

// ==========================================
// ダッシュボード（集計はコードでスナップショット。再実行で最新化）
// ==========================================
function pmRefreshDashboard() {
  var ss = getSS();
  var sheet = ss.getSheetByName(PM_DASHBOARD);
  if (!sheet) sheet = ss.insertSheet(PM_DASHBOARD, 0);

  var m = pmComputeDashboardMetrics();

  // 生成専用シートのため内容・書式を毎回クリアして作り直す（ユーザーデータは無い）
  sheet.clear();
  try { sheet.setHiddenGridlines(true); } catch (e) {}

  var rows = [];
  rows.push(['📊 WOODBASE ダッシュボード', '', '', '']);
  rows.push(['最終集計：' + fmtDT(new Date()), '', '', '']);
  rows.push(['', '', '', '']);
  rows.push(['指標', '件数', '', '']);
  rows.push(['進行中案件', m.activeProjects, '', '']);
  rows.push(['未完了タスク', m.openTasks, '', '']);
  rows.push(['期限超過タスク', m.overdueTasks, '', '']);
  rows.push(['承認待ち（全体）', m.pendingAll, '', '']);
  rows.push(['　うち請求・入金の要確認', m.pendingMoney, '', '']);
  rows.push(['', '', '', '']);
  rows.push(['📅 直近の予定（' + m.upcoming.length + '件）', '', '', '']);
  rows.push(['日付', '時刻', '予定', '案件']);
  if (m.upcoming.length) {
    m.upcoming.forEach(function(s) { rows.push([s.date, s.time || '終日', s.title, s.project]); });
  } else {
    rows.push(['（なし）', '', '', '']);
  }
  rows.push(['', '', '', '']);
  rows.push(['⏰ 期限の近い未完了タスク（' + m.dueTasks.length + '件）', '', '', '']);
  rows.push(['期日', '残', 'タスク', '担当']);
  if (m.dueTasks.length) {
    m.dueTasks.forEach(function(t) { rows.push([t.due, t.label, t.content, t.assignee]); });
  } else {
    rows.push(['（なし）', '', '', '']);
  }

  // 幅を4列に揃える
  var width = 4;
  var grid = rows.map(function(r) { while (r.length < width) r.push(''); return r.slice(0, width); });
  sheet.getRange(1, 1, grid.length, width).setValues(grid);

  // ---- 装飾（冪等：毎回上書き）----
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 280);
  sheet.setColumnWidth(4, 200);

  sheet.getRange(1, 1, 1, width).merge().setFontSize(16).setFontWeight('bold')
    .setBackground(PM_HEADER_BG).setFontColor(PM_HEADER_FG).setHorizontalAlignment('left');
  sheet.getRange(2, 1, 1, width).merge().setFontColor('#666666').setFontSize(10);

  // 各セクション見出し行を装飾
  pmStyleDashHeaderRow(sheet, 4, width);   // 指標
  pmStyleDashHeaderRow(sheet, 12, width);  // 直近の予定 列見出し
  // 「直近の予定」「期限の近い〜」のセクションタイトル行
  [11, 11 + (m.upcoming.length ? m.upcoming.length : 1) + 2].forEach(function(r) {
    try { sheet.getRange(r, 1, 1, width).setFontWeight('bold').setFontSize(12); } catch (e) {}
  });
  // タスク列見出し行（直近の予定件数に応じて動的）
  var taskHeaderRow = 12 + (m.upcoming.length ? m.upcoming.length : 1) + 3;
  pmStyleDashHeaderRow(sheet, taskHeaderRow, width);

  sheet.setFrozenRows(2);
}

function pmStyleDashHeaderRow(sheet, row, width) {
  try {
    sheet.getRange(row, 1, 1, width).setFontWeight('bold').setBackground('#ECEFF1').setFontColor('#263238');
  } catch (e) {}
}

function pmComputeDashboardMetrics() {
  var today = pmParseDate(pmTodayYmd());

  // projects
  var projects = pmReadObjects(PM_SHEET_PROJECTS);
  var activeProjects = projects.filter(function(p) { return !pmIsProjectClosed(p); }).length;

  // tasks（ヘッダー名で解決）
  var openTasks = 0, overdueTasks = 0, dueTasks = [];
  var tsheet = getSheet('タスク管理');
  if (tsheet && tsheet.getLastRow() > 1) {
    var idx = pmHeaderIndex(tsheet);
    var cContent = idx['タスク内容'], cAssignee = idx['担当者'], cDue = idx['期日'], cStatus = idx['ステータス'];
    var data = tsheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var st = String(data[i][cStatus] || '');
      if (st === 'done' || st === '完了') continue;
      openTasks++;
      var dueRaw = data[i][cDue];
      var due = pmParseDate(dueRaw);
      var dueYmd = due ? fmtDate(due) : '';
      if (due && today && due.getTime() < today.getTime()) {
        overdueTasks++;
        dueTasks.push({ due: dueYmd, content: String(data[i][cContent] || ''), assignee: String(data[i][cAssignee] || ''), label: '超過', _t: due.getTime() });
      } else if (due) {
        var days = Math.round((due.getTime() - today.getTime()) / 86400000);
        dueTasks.push({ due: dueYmd, content: String(data[i][cContent] || ''), assignee: String(data[i][cAssignee] || ''), label: 'あと' + days + '日', _t: due.getTime() });
      }
    }
  }
  dueTasks.sort(function(a, b) { return a._t - b._t; });
  dueTasks = dueTasks.slice(0, 7);

  // pending
  var pendingAll = 0, pendingMoney = 0;
  pmReadObjects(PM_SHEET_PENDING).forEach(function(r) {
    if (String(r['status']) !== 'awaiting_approval') return;
    pendingAll++;
    var t = String(r['type']);
    if (t === 'billing' || t === 'payment' || t === 'amount') pendingMoney++;
  });

  // schedules（直近・今日以降）
  var upcoming = [];
  var ssheet = getSheet('スケジュール管理');
  if (ssheet && ssheet.getLastRow() > 1) {
    var sidx = pmHeaderIndex(ssheet);
    var cTitle = sidx['予定タイトル'], cDate = sidx['日付'], cStart = sidx['開始時間'], cProj = sidx['案件名'];
    var sdata = ssheet.getDataRange().getValues();
    for (var j = 1; j < sdata.length; j++) {
      var d = pmParseDate(sdata[j][cDate]);
      if (!d || (today && d.getTime() < today.getTime())) continue;
      upcoming.push({
        date: fmtDate(d),
        time: String(sdata[j][cStart] || ''),
        title: String(sdata[j][cTitle] || ''),
        project: String(sdata[j][cProj] || ''),
        _t: d.getTime(),
      });
    }
  }
  upcoming.sort(function(a, b) { return a._t - b._t; });
  upcoming = upcoming.slice(0, 7);

  return {
    activeProjects: activeProjects,
    openTasks: openTasks,
    overdueTasks: overdueTasks,
    pendingAll: pendingAll,
    pendingMoney: pendingMoney,
    upcoming: upcoming,
    dueTasks: dueTasks,
  };
}

// ==========================================
// projects シートの装飾（列はヘッダー名で解決・物理移動なし）
// ==========================================
// 論理グループ（折りたたみ対象）。実際の列順に応じて連続区間ごとにグループ化する。
var PM_PROJ_GROUP_DATES  = ['打合せ日', '着工予定日', '引き渡し予定日'];
var PM_PROJ_GROUP_MONEY  = ['売上', '原価', '利益額', '利益率', '請求予定日', '請求済日', '請求金額', '入金予定日', '入金確認日', '入金金額'];
var PM_PROJ_GROUP_SYSTEM = ['LINEグループID', 'DriveフォルダURL', '更新元メッセージ', '作成日', '更新日'];
var PM_PROJ_GROUP_STATUS = ['営業ステータス', '設計・デザインステータス', '施工準備ステータス', '施工ステータス', 'お引き渡しステータス'];
var PM_PROJ_WRAP_COLS    = ['次回アクション', '備考', '更新元メッセージ'];

function pmFormatProjectsSheet() {
  var sheet = getSheet(PM_SHEET_PROJECTS);
  if (!sheet || sheet.getLastRow() === 0) { console.log('pmFormatProjectsSheet: projects未作成'); return; }
  var idx = pmHeaderIndex(sheet);
  var lastCol = sheet.getLastColumn();
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var maxRows = sheet.getMaxRows();

  // 行・列の固定（左の主要列を見えたまま横スクロール）
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2); // 案件ID・案件名

  // フィルター（冪等：既存を外して付け直す）
  try {
    var existing = sheet.getFilter();
    if (existing) existing.remove();
    sheet.getRange(1, 1, lastRow, lastCol).createFilter();
  } catch (e) { console.warn('projects filter:', e.message); }

  // 列幅
  pmSetColWidth(sheet, idx, '案件ID', 90);
  pmSetColWidth(sheet, idx, '案件名', 220);
  pmSetColWidth(sheet, idx, 'クライアント名', 150);
  pmSetColWidth(sheet, idx, '担当者', 90);
  pmSetColWidth(sheet, idx, '現在フェーズ', 110);
  pmSetColWidth(sheet, idx, '最終更新日', 100);
  pmSetColWidth(sheet, idx, '次回アクション', 220);
  pmSetColWidth(sheet, idx, '次回アクション期限', 120);
  pmSetColWidth(sheet, idx, '備考', 260);

  // 長文列は折り返し
  PM_PROJ_WRAP_COLS.forEach(function(h) {
    if (idx[h] !== undefined) {
      try { sheet.getRange(2, idx[h] + 1, maxRows - 1, 1).setWrap(true); } catch (e) {}
    }
  });

  // 列グループ（冪等：一旦すべて解除→連続区間ごとに作成）
  try {
    pmResetColumnGroups(sheet, lastCol);
    pmGroupColumns(sheet, idx, PM_PROJ_GROUP_STATUS);
    pmGroupColumns(sheet, idx, PM_PROJ_GROUP_DATES);
    pmGroupColumns(sheet, idx, PM_PROJ_GROUP_MONEY);
    pmGroupColumns(sheet, idx, PM_PROJ_GROUP_SYSTEM);
  } catch (e) { console.warn('projects column groups:', e.message); }

  // 条件付き書式（フェーズ／請求／入金）。冪等：全ルール置換。
  try { pmApplyProjectsConditionalFormat(sheet, idx, maxRows); } catch (e) { console.warn('projects CF:', e.message); }
}

function pmSetColWidth(sheet, idx, header, width) {
  if (idx[header] !== undefined) {
    try { sheet.setColumnWidth(idx[header] + 1, width); } catch (e) {}
  }
}

// 既存の列グループを全解除（冪等化のため）
function pmResetColumnGroups(sheet, lastCol) {
  for (var c = 1; c <= lastCol; c++) {
    var depth = 0;
    try { depth = sheet.getColumnGroupDepth(c); } catch (e) { depth = 0; }
    if (depth > 0) {
      try { sheet.getRange(1, c, 1, 1).shiftColumnGroupDepth(-depth); } catch (e) {}
    }
  }
}

// ヘッダー名集合を、実際の列順での「連続区間」ごとにグループ化する（非連続でも安全）
function pmGroupColumns(sheet, idx, headerNames) {
  var cols = headerNames
    .map(function(h) { return idx[h] !== undefined ? idx[h] + 1 : null; })
    .filter(function(c) { return c !== null; })
    .sort(function(a, b) { return a - b; });
  if (!cols.length) return;
  var runStart = cols[0], prev = cols[0];
  for (var i = 1; i <= cols.length; i++) {
    var cur = cols[i];
    if (cur === prev + 1) { prev = cur; continue; }
    // 区間 [runStart..prev] を1つのグループに
    var num = prev - runStart + 1;
    if (num >= 1) {
      try {
        sheet.getRange(1, runStart, 1, num).shiftColumnGroupDepth(1);
        // 既定で折りたたんでおく（見やすさ優先）
        var grp = sheet.getColumnGroup(runStart, 1);
        if (grp) grp.collapse();
      } catch (e) { console.warn('group ' + runStart + '+' + num + ':', e.message); }
    }
    runStart = cur; prev = cur;
  }
}

function pmApplyProjectsConditionalFormat(sheet, idx, maxRows) {
  var rules = [];
  var nRows = maxRows - 1;

  var phaseColors = {
    '営業': '#FCE5CD', '設計・デザイン': '#D9EAD3', '施工準備': '#FFF2CC',
    '施工': '#CFE2F3', 'お引き渡し': '#D9D2E9',
  };
  if (idx['現在フェーズ'] !== undefined) {
    var pr = sheet.getRange(2, idx['現在フェーズ'] + 1, nRows, 1);
    Object.keys(phaseColors).forEach(function(k) {
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(k).setBackground(phaseColors[k]).setRanges([pr]).build());
    });
  }

  var billColors = { '未請求': '#F4CCCC', '一部請求': '#FFF2CC', '請求済': '#D9EAD3' };
  if (idx['請求ステータス'] !== undefined) {
    var br = sheet.getRange(2, idx['請求ステータス'] + 1, nRows, 1);
    Object.keys(billColors).forEach(function(k) {
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(k).setBackground(billColors[k]).setRanges([br]).build());
    });
  }

  var payColors = { '未入金': '#F4CCCC', '一部入金': '#FFF2CC', '入金済': '#D9EAD3' };
  if (idx['入金ステータス'] !== undefined) {
    var yr = sheet.getRange(2, idx['入金ステータス'] + 1, nRows, 1);
    Object.keys(payColors).forEach(function(k) {
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(k).setBackground(payColors[k]).setRanges([yr]).build());
    });
  }

  sheet.setConditionalFormatRules(rules); // 全置換（冪等）
}

// ==========================================
// 全表示シート共通：ヘッダー固定＋装飾統一
// ==========================================
function pmUnifyHeaders(ss) {
  PM_ORG_SHOW_ORDER.forEach(function(name) {
    if (name === PM_DASHBOARD) return; // ダッシュボードは専用装飾済み
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastColumn() === 0) return;
    try {
      sheet.setFrozenRows(1);
      var hr = sheet.getRange(1, 1, 1, sheet.getLastColumn());
      hr.setFontWeight('bold').setFontFamily(PM_HEADER_FONT)
        .setBackground(PM_HEADER_BG).setFontColor(PM_HEADER_FG)
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
    } catch (e) { console.warn('unifyHeaders ' + name + ':', e.message); }
  });
}
