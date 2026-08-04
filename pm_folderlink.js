// ==========================================
// pm_folderlink.js — 既存Driveフォルダを案件/店舗に照合してDBへ紐付ける（P5 §7）
//   方針（非破壊）:
//     - 既存フォルダは移動・統合・削除しない。IDをDBに記録するだけ。
//     - まず dry-run（何も書き換えず一覧表示）→ 確認後に apply で driveProjectFolderId 等を書き込む。
//     - 店舗はフラットな合体名（例「焼肉こじま 藤沢店」）が多く自動判定が不確実なため、
//       案件(projects)の紐付けを自動、店舗は手動紐付け関数を用意する。
//   依存: pm_core.js, pm_folders.js(pmFolderConfig_/pmFolderNormalize_),
//         pm_manager.js(pmGetProjectById), pm_stores.js(pmGetStoreById)
// ==========================================

// 案件フォルダ紐付け（既存フォルダ→projects.driveProjectFolderId）。
//   commit=false（既定）: DRY-RUN。何も書き換えない。
//   commit=true        : 割り当てを driveProjectFolderId/Url に書き込む（フォルダは移動/削除しない）。
//   includeBrand=false（既定）: 「完全一致（表記ゆれ正規化）＋1対1」のみ（誤爆防止）。
//   includeBrand=true : 部分一致（ブランドフォルダ）や多対一（例: プラージュに複数店）も割り当てる。
//     ※ エディタからは引数なしの専用ラッパー（下記）で呼ぶと安全・簡単。
function pmLinkExistingFolders(commit, includeBrand) {
  var lines = [];
  function out(s) { lines.push(s); }
  var cfg = pmFolderConfig_();
  if (pmIsBlank(cfg.rootId)) { out('DRIVE_FOLDER_ID未設定'); console.log(lines.join('\n')); return lines.join('\n'); }

  var root;
  try { root = DriveApp.getFolderById(cfg.rootId); }
  catch (e) { out('親フォルダ取得不可: ' + e.message); console.log(lines.join('\n')); return lines.join('\n'); }

  out('== 既存フォルダ 案件紐付け ' + (commit ? '【APPLY・書込あり】' : '【DRY-RUN・書込なし】') +
      (includeBrand ? ' 〈ブランド部分一致=ON〉' : ' 〈完全一致1対1のみ〉') + ' ==');

  var children = [];
  var it = root.getFolders();
  while (it.hasNext()) children.push(it.next());
  out('親直下フォルダ数: ' + children.length);

  var folderByKey = {}; // normKey -> [folder,...]
  children.forEach(function (f) {
    var k = pmFolderNormalize_(f.getName());
    if (!k) return;
    (folderByKey[k] = folderByKey[k] || []).push(f);
  });

  var projects = pmReadObjects(PM_SHEET_PROJECTS);
  var already = 0;
  var plan = [];              // {p, name, folder, type}
  var claimByFolderId = {};   // folderId -> [案件名,...]（多対一検出）
  var suggestions = [];       // {name, hint}

  projects.forEach(function (p) {
    var name = String(p['案件名'] || '').trim();
    if (!name) return;
    if (!pmIsBlank(p['driveProjectFolderId'])) { already++; return; }

    var key = pmFolderNormalize_(name);
    var exact = key ? folderByKey[key] : null;
    if (exact && exact.length) {
      plan.push({ p: p, name: name, folder: exact[0], type: '完全一致' });
      (claimByFolderId[exact[0].getId()] = claimByFolderId[exact[0].getId()] || []).push(name);
    } else {
      var best = pmBestPartialFolder_(children, key);
      if (best && includeBrand) {
        plan.push({ p: p, name: name, folder: best, type: '部分一致' });
        (claimByFolderId[best.getId()] = claimByFolderId[best.getId()] || []).push(name);
      } else {
        suggestions.push({ name: name, hint: best ? best.getName() : null });
      }
    }
  });

  var applied = 0, assignedN = 0;
  var conflictFolders = {}; // folder名 -> [案件名]（includeBrand=false時のみ除外対象）
  plan.forEach(function (pl) {
    var claimants = claimByFolderId[pl.folder.getId()] || [];
    var shared = claimants.length > 1;
    // 完全一致1対1モードでは、多対一は衝突として自動割当しない
    if (shared && !includeBrand) { conflictFolders[pl.folder.getName()] = claimants; return; }
    assignedN++;
    out('  ✓ [' + pl.type + '] ' + pl.name + ' → 「' + pl.folder.getName() + '」' + (shared ? '（' + claimants.length + '案件で共有）' : '') + ' (' + pl.folder.getId() + ')');
    if (commit && pl.p._row) {
      try {
        pmWriteRowFields(PM_SHEET_PROJECTS, pl.p._row, { 'driveProjectFolderId': pl.folder.getId(), 'driveProjectFolderUrl': pl.folder.getUrl() });
        applied++;
      } catch (e) { out('    書込ERR: ' + e.message); }
    }
  });

  out('');
  out('割当: ' + assignedN + ' / 既に紐付け済: ' + already + (commit ? ' / 書込: ' + applied : '') +
      ' / 衝突(除外): ' + Object.keys(conflictFolders).length + ' / 未一致: ' + suggestions.length);

  var confKeys = Object.keys(conflictFolders);
  if (confKeys.length) {
    out('');
    out('-- ⚠️ 衝突：同じフォルダに複数案件が一致（手動 or ブランドONで割当）--');
    confKeys.slice(0, 20).forEach(function (fname) { out('  「' + fname + '」← ' + conflictFolders[fname].join(' / ')); });
  }

  if (suggestions.length) {
    out('');
    out('-- 未一致（' + (includeBrand ? '候補フォルダ無し' : '部分候補・手動確認') + '）--');
    suggestions.slice(0, 40).forEach(function (s) { out('  ' + s.name + (s.hint ? ' → 候補: ' + s.hint : ' → 候補なし')); });
    if (suggestions.length > 40) out('  …他 ' + (suggestions.length - 40) + ' 件');
  }

  out('');
  out('手動紐付け: 案件 pmLinkProjectFolder(案件ID, フォルダURL) / 店舗 pmLinkStoreFolder(店舗ID, フォルダURL)');
  out(commit ? '※ 記録のみ（フォルダは移動/削除なし）。' : '※ DRY-RUN：何も書き換えていません。');

  var report = lines.join('\n');
  console.log(report);
  return report;
}

// 部分一致の最良フォルダ：フォルダ名が案件名の部分文字列（＝ブランド）を優先し、最長を選ぶ。無ければ null。
function pmBestPartialFolder_(children, key) {
  if (!key || key.length < 2) return null;
  var best = null, bestLen = -1;
  for (var i = 0; i < children.length; i++) {
    var fn = pmFolderNormalize_(children[i].getName());
    if (fn.length < 2) continue;
    if (fn.indexOf(key) !== -1 || key.indexOf(fn) !== -1) {
      if (fn.length > bestLen) { best = children[i]; bestLen = fn.length; }
    }
  }
  return best;
}

// ---- エディタのドロップダウンから引数なしで呼べる専用ラッパー ----
function pmLinkPreview()          { return pmLinkExistingFolders(false, false); } // 完全一致のみ・確認
function pmLinkApply()            { return pmLinkExistingFolders(true,  false); } // 完全一致のみ・反映
function pmLinkPreviewWithBrand() { return pmLinkExistingFolders(false, true);  } // ブランド含む・確認
function pmLinkApplyWithBrand()   { return pmLinkExistingFolders(true,  true);  } // ブランド含む・反映

// DriveフォルダURL or ID から フォルダIDを取り出す
function pmExtractFolderId_(s) {
  s = String(s || '').trim();
  var m = s.match(/folders\/([A-Za-z0-9_\-]+)/);
  if (m) return m[1];
  var m2 = s.match(/[?&]id=([A-Za-z0-9_\-]+)/);
  if (m2) return m2[1];
  return s; // 既にID想定
}

// 手動紐付け（案件）：既存フォルダのURL/IDを案件に結びつける。フォルダは動かさない。
function pmLinkProjectFolder(projectId, folderIdOrUrl) {
  var proj = pmGetProjectById(projectId);
  if (!proj || !proj._row) return 'projectId不明: ' + projectId;
  var id = pmExtractFolderId_(folderIdOrUrl);
  var f;
  try { f = DriveApp.getFolderById(id); } catch (e) { return 'フォルダ取得不可: ' + e.message; }
  pmWriteRowFields(PM_SHEET_PROJECTS, proj._row, { 'driveProjectFolderId': f.getId(), 'driveProjectFolderUrl': f.getUrl() });
  var msg = 'OK(案件): ' + proj['案件名'] + ' → 「' + f.getName() + '」';
  console.log(msg);
  return msg;
}

// 手動紐付け（店舗）：既存フォルダのURL/IDを店舗に結びつける。フォルダは動かさない。
//   店舗の driveStoreFolderId と、配下の 6_データ があれば driveDataFolderId も記録。
function pmLinkStoreFolder(storeId, folderIdOrUrl) {
  var store = pmGetStoreById(storeId);
  if (!store || !store._row) return 'storeId不明: ' + storeId;
  var id = pmExtractFolderId_(folderIdOrUrl);
  var f;
  try { f = DriveApp.getFolderById(id); } catch (e) { return 'フォルダ取得不可: ' + e.message; }

  var fields = { 'driveStoreFolderId': f.getId(), 'driveStoreFolderUrl': f.getUrl() };
  // 配下に 6_データ があれば data フォルダも記録（無ければ空のまま。作成はしない）
  var data = pmFindDataSubfolder_(f);
  if (data) { fields['driveDataFolderId'] = data.getId(); fields['driveDataFolderUrl'] = data.getUrl(); }

  pmWriteRowFields(PM_SHEET_STORES, store._row, fields);
  var msg = 'OK(店舗): ' + store['案件名'] + ' / ' + store['店舗名'] + ' → 「' + f.getName() + '」' + (data ? '（6_データ検出）' : '（6_データ未検出）');
  console.log(msg);
  return msg;
}
