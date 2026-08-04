// ==========================================
// pm_folders.js  —  新Drive階層リゾルバ（Phase 1）
//   階層: ROOT(DRIVE_FOLDER_ID) / 案件名 / 店舗名 / (【原本】雛形の中身を複製)
//   雛型: TEMPLATE_FOLDER_ID（例: 【原本】雛形 = 0_図面,1_見積,…,6_データ）
//   LINE添付の保存先サブフォルダ = 6_データ（雛型の実表記＝アンダースコアに統一）
//
//   方針:
//   - 冪等: 台帳ID優先→完全一致→正規化一致→無ければ作成（§5/§10）。
//   - 既存フォルダは移動・統合・削除しない。見つかれば再利用するだけ。
//   - 同時実行/Webhook再送でも二重作成しないよう pmWithLock で直列化。
//   依存: pm_core.js(getConfig,pmWithLock,pmProp,pmIsBlank), pm_manager.js(pmAddLog)
// ==========================================

// 新Drive親フォルダ / 雛型テンプレの正規ID。ここが唯一の定義（実行時はScript Property優先で参照）。
//   ※コード内の複数箇所ハードコードを避けるため、テスト/セットアップもこの定数を使う（§1）。
var PM_NEW_DRIVE_ROOT_ID  = '14QoHs_nlMwN7IeY-SGgQ9HX51lnsKmZl';
var PM_TEMPLATE_FOLDER_ID = '1eM6BVsC2T496Ve03vVa4O6gNSKS1dR-X';

// LINEから届いた写真・ファイルの保存先サブフォルダ名（雛型内の該当フォルダ名と一致させる）
var PM_DATA_SUBFOLDER_NAME = '6_データ';

// ------------------------------------------
// 設定取得（ハードコードせず Script Property / PM設定シートから）
// ------------------------------------------
function pmFolderConfig_(opts) {
  opts = opts || {};
  var c = (typeof getConfig === 'function') ? getConfig() : {};
  return {
    rootId:     opts.rootId     || c.DRIVE_FOLDER_ID   || pmProp('DRIVE_FOLDER_ID'),
    templateId: opts.templateId || c.TEMPLATE_FOLDER_ID || pmProp('TEMPLATE_FOLDER_ID'),
  };
}

// ------------------------------------------
// 正規化（比較用）: 前後/全角空白除去・全角半角統一・小文字化
//   ※ 案件名/店舗名は固有名詞のため、ハイフン/アンダースコアの融合まではしない
// ------------------------------------------
function pmFolderNormalize_(s) {
  var t = String(s == null ? '' : s);
  try { t = t.normalize('NFKC'); } catch (e) {}   // 全角英数記号→半角
  t = t.replace(/[\s　]+/g, '');               // 空白(半角/全角/改行)除去
  return t.toLowerCase().trim();
}

// 親フォルダ内で name に一致する子フォルダを探す（完全一致→正規化一致）。無ければ null
function pmFindChildFolder_(parentFolder, name) {
  if (!parentFolder || pmIsBlank(name)) return null;
  var exact = parentFolder.getFoldersByName(String(name));
  if (exact.hasNext()) return exact.next();
  var target = pmFolderNormalize_(name);
  if (!target) return null;
  var it = parentFolder.getFolders();
  while (it.hasNext()) {
    var f = it.next();
    if (pmFolderNormalize_(f.getName()) === target) return f;
  }
  return null;
}

// find-or-create 子フォルダ。戻り値 { folder, created }
function pmEnsureChildFolder_(parentFolder, name) {
  var found = pmFindChildFolder_(parentFolder, name);
  if (found) return { folder: found, created: false };
  return { folder: parentFolder.createFolder(String(name).trim()), created: true };
}

// ------------------------------------------
// 雛型（TEMPLATE_FOLDER_ID）の中身を dest へ冪等コピー（再帰）
//   同名サブフォルダ/ファイルがあれば再利用・スキップ（二重コピーしない）
// ------------------------------------------
function pmCopyTemplateInto_(destFolder, templateId) {
  if (pmIsBlank(templateId)) throw new Error('TEMPLATE_FOLDER_ID未設定');
  var tpl = DriveApp.getFolderById(templateId);
  pmCopyFolderContents_(tpl, destFolder);
}

function pmCopyFolderContents_(src, dest) {
  var subs = src.getFolders();
  while (subs.hasNext()) {
    var s = subs.next();
    var existing = pmFindChildFolder_(dest, s.getName());
    var target = existing || dest.createFolder(s.getName());
    pmCopyFolderContents_(s, target);           // 再帰
  }
  var files = src.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    if (!dest.getFilesByName(f.getName()).hasNext()) {
      f.makeCopy(f.getName(), dest);
    }
  }
}

// 「6_データ」フォルダを表記ゆれ込みで探す（-, _, 空白, 各種ダッシュ, 全半角を融合）。
//   「6-データ」「6 データ」等の別表記でも再利用し、重複作成を防ぐ（§5）。
//   ※案件名/店舗名は固有名詞なので融合しない。データフォルダ専用の緩い一致。
function pmDataFolderKey_(s) {
  var t = String(s == null ? '' : s);
  try { t = t.normalize('NFKC'); } catch (e) {}
  return t.replace(/[\s　_\-‐-―－]+/g, '').toLowerCase();
}

function pmFindDataSubfolder_(storeFolder) {
  if (!storeFolder) return null;
  var exact = storeFolder.getFoldersByName(PM_DATA_SUBFOLDER_NAME);
  if (exact.hasNext()) return exact.next();
  var target = pmDataFolderKey_(PM_DATA_SUBFOLDER_NAME);
  var it = storeFolder.getFolders();
  while (it.hasNext()) {
    var f = it.next();
    if (pmDataFolderKey_(f.getName()) === target) return f;
  }
  return null;
}

// 店舗フォルダ配下に「6_データ」を保証する（雛型があれば複製、無ければ最低限そのフォルダだけ作る）
function pmEnsureDataSubfolder_(storeFolder, templateId) {
  var data = pmFindDataSubfolder_(storeFolder);
  if (data) return { folder: data, created: false, templateCopied: false };
  var copied = false;
  if (!pmIsBlank(templateId)) {
    try { pmCopyTemplateInto_(storeFolder, templateId); copied = true; }
    catch (e) { console.warn('pmEnsureDataSubfolder_ 雛型複製失敗: ' + e.message); }
  }
  data = pmFindDataSubfolder_(storeFolder);
  if (!data) data = storeFolder.createFolder(PM_DATA_SUBFOLDER_NAME); // フォールバック
  return { folder: data, created: true, templateCopied: copied };
}

// ------------------------------------------
// メイン: 案件名・店舗名から {案件/店舗/6_データ} フォルダを保証し ID/URL を返す
//   opts: { rootId, templateId, projectId, storeId, source, actor }
//   戻り値: { ok, caseFolderId/Url, storeFolderId/Url, dataFolderId/Url,
//             caseCreated, storeCreated, templateCopied, error }
// ------------------------------------------
function pmResolveStoreFolders(caseName, storeName, opts) {
  opts = opts || {};
  var cfg = pmFolderConfig_(opts);
  var invalid = pmValidateFolderInputs_(cfg, caseName, storeName);
  if (invalid) return invalid;

  var locked = pmWithLock(function () {
    return pmResolveStoreFoldersUnlocked_(caseName, storeName, opts, cfg);
  }, 30000);

  if (!locked.ok) return { ok: false, error: 'フォルダ作成ロック取得失敗' };
  return locked.result;
}

function pmValidateFolderInputs_(cfg, caseName, storeName) {
  if (pmIsBlank(cfg.rootId)) return { ok: false, error: 'DRIVE_FOLDER_ID未設定' };
  if (pmIsBlank(caseName))   return { ok: false, error: '案件名が空' };
  if (pmIsBlank(storeName))  return { ok: false, error: '店舗名が空' };
  return null;
}

// ロックを取らない中核処理。既に pmWithLock 内にいる呼び出し元（pm_stores 等）から使う。
function pmResolveStoreFoldersUnlocked_(caseName, storeName, opts, cfg) {
  opts = opts || {};
  cfg = cfg || pmFolderConfig_(opts);
  var invalid = pmValidateFolderInputs_(cfg, caseName, storeName);
  if (invalid) return invalid;

  var root;
  try { root = DriveApp.getFolderById(cfg.rootId); }
  catch (e) { return { ok: false, error: '親フォルダ取得不可: ' + e.message }; }

  var caseR  = pmEnsureChildFolder_(root, caseName);
  var storeR = pmEnsureChildFolder_(caseR.folder, storeName);
  var dataR  = pmEnsureDataSubfolder_(storeR.folder, cfg.templateId);

  if (caseR.created || storeR.created || dataR.created) {
    try {
      pmAddLog(opts.projectId || '', 'drive_folder_autocreated', {
        case: caseName, store: storeName,
        caseCreated: caseR.created, storeCreated: storeR.created,
        templateCopied: dataR.templateCopied, storeId: opts.storeId || '',
        source: opts.source || 'unknown',
      }, '', opts.actor || '', '');
    } catch (ig) {}
  }

  return {
    ok: true,
    caseFolderId:  caseR.folder.getId(),  caseFolderUrl:  caseR.folder.getUrl(),
    storeFolderId: storeR.folder.getId(), storeFolderUrl: storeR.folder.getUrl(),
    dataFolderId:  dataR.folder.getId(),  dataFolderUrl:  dataR.folder.getUrl(),
    caseCreated: caseR.created, storeCreated: storeR.created,
    templateCopied: dataR.templateCopied,
  };
}

// ==========================================
// セットアップ / テスト（管理者が手動実行）
// ==========================================

// 一度だけ実行: 新Drive親フォルダ・雛型IDを Script Property に設定する。
//   ※ 実行すると以後の保存先が新親フォルダに切り替わる。Phase 3 と合わせて実行推奨。
function pmApplyDriveConfig() {
  var p = PropertiesService.getScriptProperties();
  var before = { DRIVE_FOLDER_ID: p.getProperty('DRIVE_FOLDER_ID'), TEMPLATE_FOLDER_ID: p.getProperty('TEMPLATE_FOLDER_ID') };
  p.setProperty('DRIVE_FOLDER_ID', PM_NEW_DRIVE_ROOT_ID);
  p.setProperty('TEMPLATE_FOLDER_ID', PM_TEMPLATE_FOLDER_ID);
  var msg = '設定完了\n  before: ' + JSON.stringify(before)
          + '\n  after : DRIVE_FOLDER_ID=' + PM_NEW_DRIVE_ROOT_ID + ' / TEMPLATE_FOLDER_ID=' + PM_TEMPLATE_FOLDER_ID;
  console.log(msg);
  return msg;
}

// 新階層の生成を実機検証（本番親フォルダにテスト案件を作り、確認後に自動ゴミ箱移動で自己掃除）。
//   Drive設定を切り替えずに、既知IDを直接指定して検証する。
function pmTestNewHierarchy() {
  var rootId = PM_NEW_DRIVE_ROOT_ID;
  var tplId  = PM_TEMPLATE_FOLDER_ID;
  var suffix = Utilities.getUuid().slice(0, 6);
  var caseName = '★TEST_削除可_' + suffix;
  var storeName = 'テスト店舗';

  var lines = [];
  function out(s) { lines.push(s); }
  out('== 新階層 生成テスト ==');
  out('root=' + rootId + ' / template=' + tplId);
  out('case=' + caseName + ' / store=' + storeName);

  var r = pmResolveStoreFolders(caseName, storeName, { rootId: rootId, templateId: tplId, source: 'test' });
  out('結果: ' + JSON.stringify(r, null, 2));

  if (r.ok) {
    // 店舗フォルダ配下（雛型が複製されているか）を列挙
    try {
      var store = DriveApp.getFolderById(r.storeFolderId);
      var names = [];
      var it = store.getFolders();
      while (it.hasNext()) names.push(it.next().getName());
      names.sort();
      out('店舗フォルダ直下のサブフォルダ(' + names.length + '件): ' + names.join(', '));
      out('→ 期待: 雛型【原本】雛形の構成（… , ' + PM_DATA_SUBFOLDER_NAME + '）');
    } catch (e) { out('サブフォルダ列挙ERR: ' + e.message); }

    // 冪等性チェック: もう一度呼んで created=false になるはず
    var r2 = pmResolveStoreFolders(caseName, storeName, { rootId: rootId, templateId: tplId, source: 'test' });
    out('冪等再実行: caseCreated=' + r2.caseCreated + ' storeCreated=' + r2.storeCreated + ' templateCopied=' + r2.templateCopied + '（全て false が期待値）');

    // 後片付け: テスト案件フォルダをゴミ箱へ
    try { DriveApp.getFolderById(r.caseFolderId).setTrashed(true); out('後片付け: テスト案件フォルダをゴミ箱へ移動しました'); }
    catch (e) { out('後片付けERR(手動削除してください): ' + e.message); }
  }

  var report = lines.join('\n');
  console.log(report);
  return report;
}
