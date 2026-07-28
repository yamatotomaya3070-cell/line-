// ==========================================
// フェーズ1.5: 案件グループ（1案件に複数グループをぶら下げる）
//   - プロジェクト管理(辞書)の施主/業者2列の限界を解消。グループは何個でも持てる。
//   - 案件IDで projects と結ぶ。種別(施主/業者/社内/元請/行政/その他)を付与。
//   - 既存データを壊さない（追記のみ）。案件名解決は getProjectNameByGroupId が本テーブルを優先参照。
//   - エントリ:
//       backfillProjectGroups()        … DRY-RUN（変更なし・計画のみ）
//       backfillProjectGroupsCommit()  … 実適用
//   全トップレベル名は pm 接頭辞で衝突回避。
// ==========================================

// グループID → 案件名（案件グループ→案件ID→projects/辞書 で解決）。未一致は null。
function pmGetProjectNameByGroupIdFromGroups(groupId) {
  if (!groupId) return null;
  var sh = getSheet(PM_SHEET_GROUPS);
  if (!sh || sh.getLastRow() < 2) return null;
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).trim(); });
  var cGid = headers.indexOf('グループID');
  var cPid = headers.indexOf('案件ID');
  if (cGid === -1 || cPid === -1) return null;
  var pid = '';
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][cGid] || '').trim() === String(groupId).trim()) { pid = String(data[i][cPid] || '').trim(); break; }
  }
  if (!pid) return null;
  return pmGetProjectNameById(pid);
}

// 案件ID → 案件名（projects優先、無ければ辞書の正式名称）。
function pmGetProjectNameById(projectId) {
  if (!projectId) return null;
  var rows = pmReadObjects(PM_SHEET_PROJECTS);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['案件ID'] || '').trim() === String(projectId).trim()) return String(rows[i]['案件名'] || '') || null;
  }
  // フォールバック: プロジェクト管理(辞書)の案件IDから正式名称
  var sh = getSheet('プロジェクト管理');
  if (sh && sh.getLastRow() >= 2) {
    var data = sh.getDataRange().getValues();
    var h = data[0].map(function(x){ return String(x).trim(); });
    var cId = h.indexOf('案件ID'), cFormal = h.indexOf('正式名称');
    if (cId !== -1 && cFormal !== -1) {
      for (var j = 1; j < data.length; j++) {
        if (String(data[j][cId] || '').trim() === String(projectId).trim()) return String(data[j][cFormal] || '') || null;
      }
    }
  }
  return null;
}

// 既存「案件グループ」にそのグループIDが既にあるか
function pmGroupExists_(data, cGid, groupId) {
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][cGid] || '').trim() === String(groupId).trim()) return true;
  }
  return false;
}

// 辞書(プロジェクト管理)の施主/業者列 → 案件グループへ移行。commit=falsy(既定)はDRY-RUN。
function backfillProjectGroups(commit) {
  var dryRun = !commit;
  var dict = getSheet('プロジェクト管理');
  var grp  = getSheet(PM_SHEET_GROUPS);
  if (!dict) { console.log('プロジェクト管理シートがありません'); return; }
  if (!grp)  { console.log('案件グループシートがありません。先に setup() を実行してください。'); return; }
  if (dict.getLastRow() < 2) { console.log('プロジェクト管理にデータがありません'); return; }

  var d = dict.getDataRange().getValues();
  var dh = d[0].map(function(h){ return String(h).trim(); });
  var cId     = dh.indexOf('案件ID');
  var cFormal = dh.indexOf('正式名称');
  var cNushi  = dh.indexOf('グループID（施主）');
  var cGyo    = dh.indexOf('グループID（業者）');
  if (cId === -1) { console.log('辞書に「案件ID」列がありません。先に backfillProjectIds を実行してください。'); return; }

  var g = grp.getDataRange().getValues();
  var gh = (g[0] || PM_GROUPS_HEADERS).map(function(h){ return String(h).trim(); });
  var gGid = gh.indexOf('グループID');
  if (g.length < 1) g = [PM_GROUPS_HEADERS];

  var now = (typeof pmTodayYmd === 'function') ? pmTodayYmd() : '';
  var rep = { added: 0, exists: 0, noId: 0, details: [] };
  var toAppend = [];

  function plan(gid, pid, kind, name) {
    if (!gid) return;
    if (!pid) { rep.noId++; return; }
    if (pmGroupExists_(g, gGid, gid) || toAppend.some(function(r){ return r[0] === gid; })) { rep.exists++; return; }
    toAppend.push([gid, pid, kind, name, '', now]); // グループID,案件ID,種別,表示名,備考,登録日時
    rep.added++;
    rep.details.push((kind) + ': ' + name + ' / ' + gid + ' → ' + pid + (dryRun ? '（予定）' : ''));
  }

  for (var i = 1; i < d.length; i++) {
    var pid  = String(d[i][cId] || '').trim();
    var name = cFormal !== -1 ? String(d[i][cFormal] || '').trim() : '';
    if (cNushi !== -1) plan(String(d[i][cNushi] || '').trim(), pid, '施主', name);
    if (cGyo   !== -1) plan(String(d[i][cGyo]   || '').trim(), pid, '業者', name);
  }

  if (!dryRun && toAppend.length) {
    grp.getRange(grp.getLastRow() + 1, 1, toAppend.length, PM_GROUPS_HEADERS.length).setValues(toAppend);
  }

  console.log((dryRun ? '[DRY-RUN 変更なし] ' : '[適用済] ') +
    '追加:' + rep.added + ' / 既存:' + rep.exists + ' / 案件ID無で対象外:' + rep.noId);
  rep.details.forEach(function(x){ console.log('  ' + x); });
  if (dryRun) console.log('▶ 問題なければ backfillProjectGroupsCommit を実行して適用してください。');
  return rep;
}

function backfillProjectGroupsCommit() { return backfillProjectGroups(true); }
