// ==========================================
// pm_linefiles.js — LINE受信ファイルを店舗フォルダへ「種別振り分け」保存＋台帳記録（冪等）
//   振り分け（§3/§4）:
//     写真・動画(image/video) → 店舗/5_写真
//     その他ファイル(file/PDF/スプシ等) → 店舗/6_データ
//   店舗特定（§5）:
//     グループ紐付け店舗 / 単一店舗 → 確定保存
//     曖昧・未登録 → 誤保存せず「未紐付け」で案件フォルダ(or _未分類)へ退避し台帳に記録
//   冪等（§10）: LINEメッセージIDが台帳にあれば二重保存しない
//   依存: pm_core.js, pm_folders.js(pmFolderConfig_/pmFindDataSubfolder_/pmFindChildFolder_/PM_DATA_SUBFOLDER_NAME),
//         pm_stores.js(pmDetermineStoreForGroup_/pmEnsureStoreRecord/pmEnsureCaseFolder/pmDriveFolderAlive_),
//         pm_manager.js(pmGetProjectByName), コード.js(getConfig/getProjectNameByGroupId/detectProjectFromRecentLogs/
//         getOrCreateFolder/getSS/setupSheet/fmtDT)
// ==========================================

var PM_SHEET_FILES = 'ファイル台帳';
var PM_FILES_HEADERS = [
  '保存日時', '種別', '案件ID', '案件名', '店舗ID', '店舗名',
  'ファイル名', 'DriveファイルID', 'DriveURL', '保存先フォルダ', '保存先URL',
  'LINEメッセージID', 'グループID', '状態', '備考',
];

// 写真・動画の保存先サブフォルダ名（雛型に一致＝アンダースコア）
var PM_PHOTO_SUBFOLDER_NAME = '5_写真';

function pmEnsureFilesSheet() {
  var ss = getSS();
  setupSheet(ss, PM_SHEET_FILES, PM_FILES_HEADERS, '#E69138', null);
}

// LINEメッセージIDで既存保存を検索（冪等判定）。無ければ null
function pmFileAlreadySaved_(messageId) {
  if (pmIsBlank(messageId)) return null;
  var mid = String(messageId).trim();
  var rows = pmReadObjects(PM_SHEET_FILES);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['LINEメッセージID']).trim() === mid) return rows[i];
  }
  return null;
}

// メッセージ種別 → 保存先サブフォルダ名（写真/動画=5_写真、その他=6_データ）
function pmSubfolderForType_(messageType) {
  return (messageType === 'image' || messageType === 'video') ? PM_PHOTO_SUBFOLDER_NAME : PM_DATA_SUBFOLDER_NAME;
}

// 店舗フォルダ配下の指定サブフォルダ（5_写真/6_データ）を表記ゆれ込みで探し、無ければ作成
function pmEnsureNamedSubfolder_(storeFolder, name) {
  // 6_データは専用の緩い一致（-/_/空白を融合）、それ以外は完全一致→正規化一致
  if (pmDataFolderKey_(name) === pmDataFolderKey_(PM_DATA_SUBFOLDER_NAME)) {
    var d = pmFindDataSubfolder_(storeFolder);
    if (d) return d;
  } else {
    var f = pmFindChildFolder_(storeFolder, name);
    if (f) return f;
  }
  return storeFolder.createFolder(name);
}

// 台帳へ1行記録（失敗しても保存自体は継続）
function pmRecordFile_(o) {
  try {
    pmAppendRowFields(PM_SHEET_FILES, {
      '保存日時': fmtDT(new Date()),
      '種別': o.type || '', '案件ID': o.projectId || '', '案件名': o.projectName || '',
      '店舗ID': o.storeId || '', '店舗名': o.storeName || '',
      'ファイル名': o.fileName || '', 'DriveファイルID': o.fileId || '', 'DriveURL': o.url || '',
      '保存先フォルダ': o.folderName || '', '保存先URL': o.folderUrl || '',
      'LINEメッセージID': o.messageId || '', 'グループID': o.groupId || '',
      '状態': o.status || '', '備考': o.note || '',
    });
  } catch (e) { console.error('pmRecordFile_:', e.message); }
}

// ------------------------------------------
// 中核: LINEファイルを保存（種別振り分け・店舗特定・未紐付け保留・冪等）
//   messageType: 'image' | 'video' | 'file' | ...
//   戻り: { ok, saved, held?, duplicate?, fileId, url, error }
// ------------------------------------------
function pmSaveLineFile(messageType, messageId, fileName, groupId, timestamp, blob) {
  try {
    pmEnsureFilesSheet();

    // 冪等: 既に保存済みならスキップ（Webhook再送対策・§10）
    var dup = pmFileAlreadySaved_(messageId);
    if (dup) return { ok: true, saved: false, duplicate: true, fileId: dup['DriveファイルID'], url: dup['DriveURL'] };

    var cfg      = pmFolderConfig_();
    var when     = timestamp || new Date();
    var dateStr  = Utilities.formatDate(when, 'Asia/Tokyo', 'yyyyMMdd');
    var safeName = dateStr + '_' + (fileName || messageId);
    var subName  = pmSubfolderForType_(messageType);

    // 案件を解決（グループ紐付け → 直近ログ）
    var projectName = getProjectNameByGroupId(groupId) || detectProjectFromRecentLogs(groupId) || '';
    var proj      = projectName ? pmGetProjectByName(projectName) : null;
    var projectId = proj ? String(proj['案件ID'] || '') : '';

    // 店舗を特定（グループ紐付け店舗 / 単一店舗 / 曖昧 / 未登録）
    var storeInfo = projectId ? pmDetermineStoreForGroup_(projectId, groupId) : { status: 'no_store' };

    // ① 店舗が確定 → 店舗/サブフォルダ（5_写真 or 6_データ）へ保存
    if (storeInfo.status === 'ok' && storeInfo.store) {
      var store = storeInfo.store;
      var storeFolderId = store['driveStoreFolderId'];
      // フォルダIDが無い/死んでいる → 店舗レコード保証で作り直す
      if (pmIsBlank(storeFolderId) || !pmDriveFolderAlive_(storeFolderId)) {
        var ens = pmEnsureStoreRecord(projectId, store['店舗名'], { caseName: projectName, groupId: groupId, source: 'line-file' });
        if (ens.ok && ens.folders) storeFolderId = ens.folders.storeFolderId;
      }
      if (!pmIsBlank(storeFolderId)) {
        var storeFolder = DriveApp.getFolderById(storeFolderId);
        var sub = pmEnsureNamedSubfolder_(storeFolder, subName);
        blob.setName(safeName);
        var file = sub.createFile(blob);
        pmRecordFile_({
          type: messageType, projectId: projectId, projectName: projectName,
          storeId: store['店舗ID'], storeName: store['店舗名'],
          fileName: safeName, fileId: file.getId(), url: file.getUrl(),
          folderName: sub.getName(), folderUrl: sub.getUrl(),
          messageId: messageId, groupId: groupId, status: '保存済',
        });
        return { ok: true, saved: true, fileId: file.getId(), url: file.getUrl() };
      }
    }

    // ② 案件は判るが店舗が未確定（曖昧/未登録）→ 案件フォルダへ退避＋「未紐付け」記録（誤保存しない・§5）
    if (projectId && !pmIsBlank(projectName)) {
      var cf = pmEnsureCaseFolder(projectId, projectName, { source: 'line-file' });
      if (cf.ok) {
        var caseFolder = DriveApp.getFolderById(cf.caseFolderId);
        blob.setName(safeName);
        var f2 = caseFolder.createFile(blob);
        pmRecordFile_({
          type: messageType, projectId: projectId, projectName: projectName, storeId: '', storeName: '',
          fileName: safeName, fileId: f2.getId(), url: f2.getUrl(),
          folderName: caseFolder.getName(), folderUrl: caseFolder.getUrl(),
          messageId: messageId, groupId: groupId, status: '未紐付け',
          note: (storeInfo.status === 'ambiguous' ? '店舗が複数・未特定' : '店舗未登録'),
        });
        return { ok: true, saved: true, held: true, fileId: f2.getId(), url: f2.getUrl() };
      }
    }

    // ③ 案件も不明 → 親/_未分類 へ退避＋「未紐付け」記録
    var holdRoot = getOrCreateFolder(cfg.rootId, '_未分類');
    if (holdRoot) {
      blob.setName(safeName);
      var f3 = holdRoot.createFile(blob);
      pmRecordFile_({
        type: messageType, projectId: '', projectName: '', storeId: '', storeName: '',
        fileName: safeName, fileId: f3.getId(), url: f3.getUrl(),
        folderName: holdRoot.getName(), folderUrl: holdRoot.getUrl(),
        messageId: messageId, groupId: groupId, status: '未紐付け', note: '案件未特定',
      });
      return { ok: true, saved: true, held: true, fileId: f3.getId(), url: f3.getUrl() };
    }

    return { ok: false, error: '保存先を決定できませんでした（DRIVE_FOLDER_ID未設定）' };
  } catch (e) {
    console.error('pmSaveLineFile:', e.message);
    return { ok: false, error: e.message };
  }
}

// ------------------------------------------
// 未紐付けファイルの再割り当て（P4の土台）：台帳の1件を正しい案件/店舗へ移動＋記録更新。
//   messageId で台帳を特定 → 店舗の 5_写真/6_データ へ Drive 移動 → 状態を「保存済」に更新。
//   既存ファイルの削除はしない（moveTo による移動のみ）。
// ------------------------------------------
function pmRelinkLineFile(messageId, projectId, storeName) {
  var row = pmFileAlreadySaved_(messageId);
  if (!row) return { ok: false, error: '台帳に該当メッセージがありません' };
  if (pmIsBlank(storeName)) return { ok: false, error: '店舗名が必要です' };

  var proj = projectId ? pmGetProjectById(projectId) : null;
  var caseName = proj ? proj['案件名'] : row['案件名'];
  var sr = pmEnsureStoreRecord(projectId, storeName, { caseName: caseName, source: 'relink' });
  if (!sr.ok) return { ok: false, error: '店舗解決失敗: ' + sr.error };

  try {
    var file = DriveApp.getFileById(row['DriveファイルID']);
    var storeFolder = DriveApp.getFolderById(sr.folders.storeFolderId);
    var sub = pmEnsureNamedSubfolder_(storeFolder, pmSubfolderForType_(row['種別']));
    file.moveTo(sub);
    if (row._row) {
      pmWriteRowFields(PM_SHEET_FILES, row._row, {
        '案件ID': projectId || '', '案件名': caseName || '', '店舗ID': sr.storeId || '', '店舗名': storeName,
        '保存先フォルダ': sub.getName(), '保存先URL': sub.getUrl(), 'DriveURL': file.getUrl(),
        '状態': '保存済', '備考': '再割り当て',
      });
    }
    return { ok: true, url: file.getUrl(), folder: sub.getName() };
  } catch (e) { return { ok: false, error: e.message }; }
}

// 未紐付けファイル一覧（管理画面/管理者用）
function pmListUnlinkedFiles() {
  return pmReadObjects(PM_SHEET_FILES).filter(function (r) {
    return String(r['状態']).trim() === '未紐付け';
  });
}
