// ==========================================
// ProjectDrive：案件フォルダの自動作成（冪等）とURL保存
//   新階層（§2）: ROOT(DRIVE_FOLDER_ID) / 案件名   ← ここまでを保証する
//     ※ 店舗名/6_データ は店舗登録時 pmEnsureStoreRecord が作る（pm_stores/pm_folders）。
//   旧構成「ROOT/プロジェクト管理/{案件ID}_{案件名}/[01_見積..]」は廃止（余計な階層を作らない）。
//   依存: pm_core.js, pm_manager.js, pm_stores.js(pmEnsureCaseFolderUnlocked_), pm_folders.js
// ==========================================

// 旧構成の名残（後片付け・E2Eテストの掃除でのみ参照）。新規作成には使わない。
var PM_DRIVE_PARENT_NAME = 'プロジェクト管理';
var PM_DRIVE_SUBFOLDERS  = ['01_見積', '02_図面', '03_契約書', '04_写真', '05_請求書', '06_議事録'];

// 案件フォルダ（新階層 ROOT/案件名）を保証する共通サービス。
//   - source は監査ログ用（LINE / progress-board 等）。
//   - DBの driveProjectFolderId/Url が既にあれば必ず再利用（再送・リトライ・二重クリックで新規作成しない・§6/§10）。
//   - ロックは取らない（pmEnsureCaseFolderUnlocked_ を利用）。呼び出し元がロック内でも安全（入れ子回避）。
//   - 旧「プロジェクト管理」層は作らない。
function pmEnsureProjectFolder(projectId, projectName, opts) {
  opts = opts || {};
  // デプロイ直後でも、Drive状態列だけは非破壊で不足分を追加する。
  pmEnsureDriveColumns_();
  var proj = pmGetProjectById(projectId);
  if (!proj) return { ok: false, status: 'error', error: '案件が見つかりません' };
  // 新階層フォルダが既に記録済みなら再利用
  if (!pmIsBlank(proj['driveProjectFolderId']) || !pmIsBlank(proj['driveProjectFolderUrl'])) {
    return { ok: true, status: 'ready',
      id:  String(proj['driveProjectFolderId']  || ''),
      url: String(proj['driveProjectFolderUrl'] || '') };
  }

  var caseName = projectName || proj['案件名'];
  try {
    var r = pmEnsureCaseFolderUnlocked_(projectId, caseName, null, {
      source: opts.source || 'unknown', actor: opts.actor || '',
      rootId: opts.rootId, templateId: opts.templateId,
    });
    if (!r.ok) return pmSaveDriveError_(proj, r.error || '案件フォルダ作成失敗');

    // 後方互換：旧Drive状態列にも同じ案件フォルダを反映（ダッシュボードの「開く」/フラット検索が参照）。
    var result = { 'DriveフォルダID': r.caseFolderId, 'DriveフォルダURL': r.caseFolderUrl,
      'Driveフォルダ名': String(caseName || ''), 'Drive作成日時': new Date(),
      'Drive同期ステータス': 'ready', 'Drive同期エラー': '', '更新日': pmTodayYmd() };
    try { if (proj._row) pmWriteRowFields(PM_SHEET_PROJECTS, proj._row, result); } catch (e) {}
    try { pmAddLog(projectId, 'drive_folder_ready', { source: opts.source || 'unknown', folderId: r.caseFolderId, hierarchy: '案件名(新階層)' }, '', opts.actor || '', ''); } catch (ignore) {}
    return { ok: true, status: 'ready', id: r.caseFolderId, url: r.caseFolderUrl, name: String(caseName || '') };
  } catch (err) {
    console.error('pmEnsureProjectFolder [' + (opts.source || 'unknown') + '] project=' + projectId + ':', err.message);
    return pmSaveDriveError_(proj, err.message);
  }
}

function pmEnsureDriveColumns_() {
  var sheet = pmSheet(PM_SHEET_PROJECTS);
  if (!sheet || sheet.getLastRow() === 0) return;
  var have = pmHeaderIndex(sheet);
  ['DriveフォルダID', 'Driveフォルダ名', 'Drive作成日時', 'Drive同期ステータス', 'Drive同期エラー'].forEach(function(header) {
    if (have[header] === undefined) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header).setFontWeight('bold');
      have[header] = sheet.getLastColumn() - 1;
    }
  });
}

function pmSaveDriveError_(proj, message) {
  var safeMessage = String(message || '不明なエラー').slice(0, 500);
  pmWriteRowFields(PM_SHEET_PROJECTS, proj._row, { 'Drive同期ステータス': 'error', 'Drive同期エラー': safeMessage, '更新日': pmTodayYmd() });
  return { ok: false, status: 'error', error: safeMessage };
}

// Google Driveで扱いづらい制御文字・パス区切りを除去し、空白を正規化する。
function pmProjectFolderName_(projectId, projectName) {
  var clean = String(projectName || '').replace(/[\\/\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim();
  clean = clean.slice(0, 150) || '案件名未設定';
  return String(projectId || 'NO-ID') + '_' + clean;
}

// 空白（全角含む）を除去した比較用文字列
function pmNormalizeFolderName_(s) {
  return String(s || '').replace(/[\s　]/g, '').trim();
}

// 旧フラット構成（DRIVE_FOLDER_ID直下＝案件名そのまま）のフォルダを検索（作成はしない）。
//   projects.案件名だけでなく、案件IDで名寄せできる「プロジェクト管理」の正式名称／略称も
//   候補に含める（フォルダ作成時の名前が今の案件名と表記ゆれで一致しないケースを拾うため）。
//   完全一致を優先、無ければ部分一致（どちらかがどちらかを含む）で妥協的にヒットさせる。
function pmFindExistingFlatFolder(projectId, projectName) {
  var config = getConfig();
  if (!config.DRIVE_FOLDER_ID) return null;
  var parent;
  try { parent = DriveApp.getFolderById(config.DRIVE_FOLDER_ID); } catch (e) { return null; }

  var candidates = [projectName];
  try {
    var dictRows = pmReadObjects('プロジェクト管理');
    for (var i = 0; i < dictRows.length; i++) {
      if (String(dictRows[i]['案件ID'] || '').trim() === String(projectId).trim()) {
        if (dictRows[i]['正式名称']) candidates.push(dictRows[i]['正式名称']);
        if (dictRows[i]['略称']) candidates.push(dictRows[i]['略称']);
        break;
      }
    }
  } catch (e) {}

  var normCandidates = [];
  candidates.forEach(function(c) {
    var n = pmNormalizeFolderName_(c);
    if (n.length >= 2 && normCandidates.indexOf(n) === -1) normCandidates.push(n);
  });
  if (!normCandidates.length) return null;

  var it = parent.getFolders();
  var partial = null;
  while (it.hasNext()) {
    var f = it.next();
    var fname = pmNormalizeFolderName_(f.getName());
    if (!fname) continue;
    for (var j = 0; j < normCandidates.length; j++) {
      var c = normCandidates[j];
      if (fname === c) return f.getUrl(); // 完全一致は即決定
      if (!partial && (fname.indexOf(c) !== -1 || c.indexOf(fname) !== -1)) partial = f.getUrl();
    }
  }
  return partial;
}
