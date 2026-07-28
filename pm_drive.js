// ==========================================
// ProjectDrive：案件フォルダの自動作成（冪等）とURL保存
//   共有ドライブ/(DRIVE_FOLDER_ID)/プロジェクト管理/{案件ID}_{案件名}/[01_見積..06_議事録]
//   依存: pm_core.js, pm_manager.js, 既存 getConfig / getOrCreateFolder
// ==========================================

var PM_DRIVE_PARENT_NAME = 'プロジェクト管理';
var PM_DRIVE_SUBFOLDERS  = ['01_見積', '02_図面', '03_契約書', '04_写真', '05_請求書', '06_議事録'];

// 案件フォルダを保証。既に projects に DriveフォルダURL があれば作成しない。
// 戻り値: フォルダURL（作成/既存）または null
function pmEnsureProjectFolder(projectId, projectName) {
  var proj = pmGetProjectById(projectId);
  if (!proj) return null;
  if (!pmIsBlank(proj['DriveフォルダURL'])) return proj['DriveフォルダURL']; // 冪等

  var config = getConfig();
  if (!config.DRIVE_FOLDER_ID) { console.log('pmEnsureProjectFolder: DRIVE_FOLDER_ID未設定'); return null; }

  try {
    var parent = getOrCreateFolder(config.DRIVE_FOLDER_ID, PM_DRIVE_PARENT_NAME);
    if (!parent) return null;
    var folderName = projectId + '_' + (projectName || '');
    var projectFolder = getOrCreateFolder(parent.getId(), folderName);
    if (!projectFolder) return null;

    PM_DRIVE_SUBFOLDERS.forEach(function(sub) {
      getOrCreateFolder(projectFolder.getId(), sub);
    });

    var url = projectFolder.getUrl();
    pmWriteRowFields(PM_SHEET_PROJECTS, proj._row, { 'DriveフォルダURL': url, '更新日': pmTodayYmd() });
    return url;
  } catch (err) {
    console.error('pmEnsureProjectFolder error:', err.message);
    return null;
  }
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
