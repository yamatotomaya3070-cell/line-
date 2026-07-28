// ==========================================
// pm_stores.js  —  店舗(store)エンティティ + DriveフォルダID台帳（Phase 2）
//   1案件に複数店舗をぶら下げる。各店舗は自分の Drive フォルダ(店舗/6_データ)IDを保持。
//   フォルダは pm_folders.js(pmResolveStoreFolders) で生成・特定し、IDをここに保存。
//   次回以降は保存済みIDを優先利用（案件名/店舗名の再検索を避ける・§6）。
//   依存: pm_core.js, コード.js(setupSheet,generateId,fmtDT), pm_manager.js(pmGetProjectById),
//         pm_folders.js(pmResolveStoreFolders)
// ==========================================

var PM_SHEET_STORES = 'stores';
var PM_STORES_HEADERS = [
  '店舗ID', '案件ID', '案件名', '店舗名', 'LINEグループID',
  'driveStoreFolderId', 'driveStoreFolderUrl', 'driveDataFolderId', 'driveDataFolderUrl',
  '表示名', '状態', '備考', '登録日時', '更新日時',
];

// ------------------------------------------
// シート/列の保証（非破壊）。projects の新階層列も併せて追加。
// ------------------------------------------
function pmEnsureStoreSheet() {
  var ss = getSS();
  setupSheet(ss, PM_SHEET_STORES, PM_STORES_HEADERS, '#3D85C6', null);
  // projects に driveProjectFolderId/Url を非破壊追加（PM_PROJECTS_HEADERS に定義済み）
  setupSheet(ss, PM_SHEET_PROJECTS, PM_PROJECTS_HEADERS, '#6AA84F', null);
}

// ------------------------------------------
// 参照系
// ------------------------------------------
function pmListStoresByProject(projectId) {
  if (pmIsBlank(projectId)) return [];
  var pid = String(projectId).trim();
  return pmReadObjects(PM_SHEET_STORES).filter(function (s) {
    return String(s['案件ID']).trim() === pid && !pmIsTruthyFlag_(s['状態'], 'deleted');
  });
}

function pmGetStoreById(storeId) {
  if (pmIsBlank(storeId)) return null;
  var sid = String(storeId).trim();
  var rows = pmReadObjects(PM_SHEET_STORES);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['店舗ID']).trim() === sid) return rows[i];
  }
  return null;
}

// 案件ID + 店舗名（正規化一致）で店舗行を探す。無ければ null
function pmFindStoreRow_(projectId, storeName) {
  var pid = String(projectId || '').trim();
  var target = pmFolderNormalize_(storeName);
  var rows = pmReadObjects(PM_SHEET_STORES);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['案件ID']).trim() !== pid) continue;
    if (pmFolderNormalize_(rows[i]['店舗名']) === target) return rows[i];
  }
  return null;
}

function pmIsTruthyFlag_(v, word) {
  var s = String(v == null ? '' : v).trim().toLowerCase();
  return word ? s === String(word).toLowerCase() : (s === 'true' || s === '1' || s === 'yes');
}

// ------------------------------------------
// 中核: 案件ID + 店舗名 から 店舗レコード＋Driveフォルダ を保証する（冪等）
//   - 店舗行が無ければ作成。あれば再利用。
//   - 保存済み driveDataFolderId が生きていれば再解決を省略（ID優先・§6）。
//   - フォルダは pmResolveStoreFolders で ROOT/案件名/店舗名/6_データ を特定/作成。
//   - 案件フォルダIDは projects.driveProjectFolderId/Url にも保存。
//   opts: { caseName, groupId, displayName, source, actor, forceResolve }
//   戻り値: { ok, storeId, store(行), folders(resolve結果), error }
// ------------------------------------------
function pmEnsureStoreRecord(projectId, storeName, opts) {
  opts = opts || {};
  if (pmIsBlank(storeName)) return { ok: false, error: '店舗名が空' };

  pmEnsureStoreSheet();

  // 案件名の解決（opts優先→projectsから）
  var caseName = opts.caseName;
  var project = null;
  if (!pmIsBlank(projectId)) {
    try { project = pmGetProjectById(projectId); } catch (e) {}
    if (project && pmIsBlank(caseName)) caseName = project['案件名'];
  }
  if (pmIsBlank(caseName)) return { ok: false, error: '案件名を特定できません（projectId=' + projectId + '）' };

  var locked = pmWithLock(function () {
    var now = fmtDT(new Date());
    var existing = pmFindStoreRow_(projectId, storeName);

    // 既存で driveDataFolderId が生きていれば、フォルダ再解決を省略できる
    if (existing && !opts.forceResolve && !pmIsBlank(existing['driveDataFolderId']) && pmDriveFolderAlive_(existing['driveDataFolderId'])) {
      return { ok: true, storeId: existing['店舗ID'], store: existing, folders: {
        ok: true,
        storeFolderId: existing['driveStoreFolderId'], storeFolderUrl: existing['driveStoreFolderUrl'],
        dataFolderId: existing['driveDataFolderId'], dataFolderUrl: existing['driveDataFolderUrl'],
        reused: true,
      } };
    }

    // フォルダを特定/作成（ROOT/案件名/店舗名/6_データ）
    // 既に pmWithLock 内なので Unlocked 版を使う（入れ子ロック回避）。
    // rootId/templateId は opts で上書き可（テスト時は新親を指定）。
    var f = pmResolveStoreFoldersUnlocked_(caseName, storeName, {
      projectId: projectId || '', storeId: existing ? existing['店舗ID'] : '',
      source: opts.source || 'store-record', actor: opts.actor || '',
      rootId: opts.rootId, templateId: opts.templateId,
    });
    if (!f.ok) return { ok: false, error: 'フォルダ解決失敗: ' + f.error };

    // 案件フォルダIDを projects へ保存（新階層）
    if (project && project._row) {
      var pf = {};
      if (pmIsBlank(project['driveProjectFolderId'])) pf['driveProjectFolderId'] = f.caseFolderId;
      if (pmIsBlank(project['driveProjectFolderUrl'])) pf['driveProjectFolderUrl'] = f.caseFolderUrl;
      if (Object.keys(pf).length) { try { pmWriteRowFields(PM_SHEET_PROJECTS, project._row, pf); } catch (e) {} }
    }

    // 店舗行 upsert
    var storeId = existing ? existing['店舗ID'] : generateId();
    var fields = {
      '店舗ID': storeId, '案件ID': projectId || '', '案件名': caseName, '店舗名': String(storeName).trim(),
      'driveStoreFolderId': f.storeFolderId, 'driveStoreFolderUrl': f.storeFolderUrl,
      'driveDataFolderId': f.dataFolderId, 'driveDataFolderUrl': f.dataFolderUrl,
      '状態': 'active', '更新日時': now,
    };
    if (!pmIsBlank(opts.groupId)) fields['LINEグループID'] = String(opts.groupId).trim();
    if (!pmIsBlank(opts.displayName)) fields['表示名'] = String(opts.displayName).trim();

    if (existing && existing._row) {
      pmWriteRowFields(PM_SHEET_STORES, existing._row, fields);
    } else {
      fields['登録日時'] = now;
      pmAppendRowFields(PM_SHEET_STORES, fields);
    }

    try {
      pmAddLog(projectId || '', 'store_folder_ensured', {
        storeId: storeId, store: storeName, dataFolderId: f.dataFolderId,
        created: !existing, source: opts.source || 'store-record',
      }, '', opts.actor || '', '');
    } catch (ig) {}

    return { ok: true, storeId: storeId, store: pmFindStoreRow_(projectId, storeName), folders: f };
  }, 30000);

  if (!locked.ok) return { ok: false, error: '店舗登録ロック取得失敗' };
  return locked.result;
}

// 店舗未指定の案件でも、案件フォルダ(ROOT/案件名)だけは用意して projects に保存。
//   店舗フォルダ・6_データは店舗登録時(pmEnsureStoreRecord)に作る。
function pmEnsureCaseFolder(projectId, caseName, opts) {
  opts = opts || {};
  var cfg = pmFolderConfig_(opts);
  if (pmIsBlank(cfg.rootId)) return { ok: false, error: 'DRIVE_FOLDER_ID未設定' };
  if (pmIsBlank(caseName))   return { ok: false, error: '案件名が空' };

  var locked = pmWithLock(function () {
    var root;
    try { root = DriveApp.getFolderById(cfg.rootId); }
    catch (e) { return { ok: false, error: '親フォルダ取得不可: ' + e.message }; }
    var caseR = pmEnsureChildFolder_(root, caseName);
    if (!pmIsBlank(projectId)) {
      var proj = null;
      try { proj = pmGetProjectById(projectId); } catch (e) {}
      if (proj && proj._row) {
        var pf = {};
        if (pmIsBlank(proj['driveProjectFolderId']))  pf['driveProjectFolderId']  = caseR.folder.getId();
        if (pmIsBlank(proj['driveProjectFolderUrl'])) pf['driveProjectFolderUrl'] = caseR.folder.getUrl();
        if (Object.keys(pf).length) { try { pmWriteRowFields(PM_SHEET_PROJECTS, proj._row, pf); } catch (e) {} }
      }
    }
    return { ok: true, caseFolderId: caseR.folder.getId(), caseFolderUrl: caseR.folder.getUrl(), created: caseR.created };
  }, 30000);
  if (!locked.ok) return { ok: false, error: '案件フォルダ作成ロック取得失敗' };
  return locked.result;
}

// DriveフォルダIDが生存しているか（ゴミ箱含め取得可否）
function pmDriveFolderAlive_(folderId) {
  if (pmIsBlank(folderId)) return false;
  try { var f = DriveApp.getFolderById(folderId); return !f.isTrashed(); }
  catch (e) { return false; }
}

// ==========================================
// テスト（管理者が手動実行）— 一時レコードを作って検証し、後片付け
// ==========================================
function pmTestStoreRecord() {
  var lines = [];
  function out(s) { lines.push(s); }
  var suffix = Utilities.getUuid().slice(0, 6);
  var caseName = '★TEST_削除可_' + suffix;
  var storeName = '検証店舗A';
  var fakePid = 'TESTPID_' + suffix;

  out('== 店舗レコード保存テスト ==');
  pmEnsureStoreSheet();
  out('storesシート/列を保証しました。ヘッダー: ' + PM_STORES_HEADERS.join(', '));

  // テストは新親フォルダを直接指定（本番設定=旧フォルダにはadmin書込権が無いため）
  var testRoot = (typeof DIAG_NEW_ROOT_ID !== 'undefined') ? DIAG_NEW_ROOT_ID : '14QoHs_nlMwN7IeY-SGgQ9HX51lnsKmZl';
  var testTpl  = (typeof DIAG_TEMPLATE_ID !== 'undefined') ? DIAG_TEMPLATE_ID  : '1eM6BVsC2T496Ve03vVa4O6gNSKS1dR-X';

  // caseName を明示指定（projects に依存しない）
  var r = pmEnsureStoreRecord(fakePid, storeName, { caseName: caseName, source: 'test', rootId: testRoot, templateId: testTpl });
  out('1回目: ' + JSON.stringify({ ok: r.ok, storeId: r.storeId, error: r.error, dataFolderId: r.folders && r.folders.dataFolderId }));

  var r2 = pmEnsureStoreRecord(fakePid, storeName, { caseName: caseName, source: 'test', rootId: testRoot, templateId: testTpl });
  out('2回目(冪等): 同一店舗ID=' + (r.ok && r2.ok && r.storeId === r2.storeId) + ' reused=' + (r2.folders && r2.folders.reused));

  var row = pmFindStoreRow_(fakePid, storeName);
  out('保存行: ' + (row ? JSON.stringify({
    店舗ID: row['店舗ID'], 案件名: row['案件名'], 店舗名: row['店舗名'],
    driveDataFolderUrl: row['driveDataFolderUrl'],
  }) : '(見つからない)'));

  // 後片付け: フォルダをゴミ箱へ + 店舗行を削除
  try {
    if (r.folders && r.folders.caseFolderId) { DriveApp.getFolderById(r.folders.caseFolderId).setTrashed(true); out('後片付け: テスト案件フォルダをゴミ箱へ'); }
  } catch (e) { out('フォルダ後片付けERR: ' + e.message); }
  try {
    if (row && row._row) { getSheet(PM_SHEET_STORES).deleteRow(row._row); out('後片付け: テスト店舗行を削除'); }
  } catch (e) { out('行後片付けERR: ' + e.message); }

  var report = lines.join('\n');
  console.log(report);
  return report;
}
