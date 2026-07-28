// ==========================================
// プロジェクト名整理 一括適用（2026-07-07 仕分け確定版）
//   仕分け表: プロジェクト管理_仕分け作業表_2026-07-07.md
//   使い方: GASエディタで
//     applyReorg()        … DRY-RUN（変更なし・何が起きるかログに出すだけ）
//     applyReorgCommit()  … 実適用
//   内容:
//     A. 改名（プロジェクト管理.正式名称 と projects.案件名 の両方）
//     B. まとめグループを「案件グループ」へ追記
//     C/D. タスク専用・社内グループを「案件グループ」へ台帳記録
//     E. 不要案件の行削除（プロジェクト管理・projects 両方）
//   冪等: 何度実行しても同じ結果（改名済みはスキップ、登録済みグループはスキップ、
//         削除済み行はスキップ）。
//   実行後に backfillProjectGroups() → backfillProjectGroupsCommit() を実行すること
//   （残した案件の施主/業者列のグループIDを案件グループへ移すため）。
// ==========================================

// A. 改名（案件ID → 新しい案件名）。そのままの案件は載せない。
var REORG_RENAMES = [
  ['2026-003', '理容プラージュ　帯広'],
  ['2026-005', 'UMI宮古'],
  ['2026-013', 'Amuyo渋谷店'],
  ['2026-021', 'ラフリー天王寺'],
  ['2026-027', 'べんてん屋'],
  ['2026-028', '理容プラージュ　砺波'],
  ['2026-029', '美容プラージュ　川西'],
  ['2026-036', 'ことくしゃちょー内装'],
  ['2026-038', '美容プラージュ　東海'],
  ['2026-039', '美容プラージュ　知多半田'],
  ['2026-040', '理容プラージュ　釧路桂木'],
  ['2026-041', '美容プラージュ　今治小泉'],
  ['2026-042', '理容プラージュ　利府'],
  ['2026-043', '理容プラージュ　出雲'],
  ['2026-045', '美容プラージュ　旭川永山'],
  ['2026-046', '理容プラージュ　武豊里中'],
  ['2026-047', '理容プラージュ　旭川永山'],
  ['2026-048', '理容プラージュ　大和田'],
  ['2026-049', '理容美容プラージュ　岐阜'],
  ['2026-050', '理容プラージュ　蒲田'],
  ['2026-053', '理容プラージュ　豊田高橋'],
];

// B. まとめ（案件にぶら下げるグループ）: [グループID, 案件ID, 種別, 表示名]
// C/D. 台帳（案件ID空欄＝進捗管理対象外。タスクはグループ名のまま動く）
var REORG_GROUPS = [
  // --- B. まとめ ---
  ['C09549bd01c9268511ad549c40a371764', '2026-005', '社内', '宮古島UMITERRACE是正チーム'],
  ['Cc9fd62e92842aca3d8f0f86f7aa2a47f', '2026-021', '施主', 'ラフリー本部'],
  ['Cfa18bb4faa562f4305a6a470245c213a', '2026-021', '業者', 'ラフリー天王寺（ダイノック）'],
  ['C1ec82464904c6d51ea6c467cbd3f1449', '2026-021', '業者', 'ラフリー天王寺（電気屋・山内）'],
  ['Cab11234b7071e8902720a5af35383653', '2026-021', '業者', 'ラフリー天王寺（清水・HIRO）'],
  // --- C. タスク専用グループ（業者・施主・行政） ---
  ['C4db58afd8296dc6f392964056bc0b997', '', '業者', 'プラージュ建設加藤'],
  ['Cb851bdd98fc5316520a446de7a453f64', '', '業者', '柴田建設（旭川永山）'],
  ['C763cf0f9dd386d0f16ddc12d1438a89d', '', '業者', 'UMITERRACE宮古島'],
  ['C9b3bf966f9f9a58000daee0ab4b38d72', '', '施主', 'UMITERRACE/385ホテル'],
  ['Caab4c642dae021080347b29d581bc339', '', '業者', 'UMITERRACEサイン（タイペックス）'],
  ['Cf68c3cf7b2a0577afb2ebb1bf523f7fa', '', '施主', '津田歯科クリニック'],
  ['Ce9dbe398fc443e33a75592aa8f42db42', '', '業者', '雨のち晴れクリニック武蔵小杉'],
  ['C5889df50cbe5b8813655514aa903cce5', '', '社内', 'アムヨ/雨晴/武蔵小杉 社内'],
  ['Ce7b409f51c9a0533024451461df715ed', '', '施主', 'Amuyo（てんか・シン）'],
  ['C84df4447ae6c56dfba37ebe7650b3ea7', '', '業者', '安田陳列（家具配送）'],
  ['C86bdb559212d773aaba18e8bd77a1f14', '', '業者', 'プラージュ建設佐野'],
  ['C91bc2398d4e578a685757ef26b1d1eae', '', '業者', 'プラージュ建設糸洲'],
  ['Cf991d32865459ae86fd5399c9be6607a', '', '業者', 'カエルデザイン/中井建設'],
  ['Cdfead3d4bedbc18e93b3ff5399315c08', '', '業者', '中井建設'],
  ['Cc095c20df5b6d59a50d134494a6219d0', '', '業者', 'プラージュ建設小林奨伍'],
  ['C04901189e304fd3f2cfe4ada91901ac7', '', '業者', '山本建設（岐阜）'],
  ['Cc2f403959cdefdf6006f945be0bb51c1', '', '業者', '河合（外注作図）'],
  ['Cd082ba7096bf37f7d8be34b050c4c7c8', '', '行政', '水道局協議（宮古島）'],
  // --- D. 社内グループ ---
  ['C064410a492da8938d05b3105fb90c58a', '', '社内', 'WBG開発テスト'],
  ['Caeea6c76ec54ef9f800ae13a2958db92', '', '社内', 'WOODBASE社内経理'],
  ['Ce064982f9932d5f5bb7e1c52725376f2', '', '社内', 'WBF内部チーム'],
  ['C5d464a41877ed749da53c487b40205f8', '', '社内', 'WBF事務スタッフ'],
  ['C13e8ef03a8cbdb7e9bb9eeaab5dc631a', '', '社内', 'WBグループ内部'],
  ['Ca265505ee79949dfd8da151760f6c0bc', '', '社内', 'ヤマトのグループ'],
];

// E. 行削除する案件ID（プロジェクト管理・projects の両方から）
//   まとめ元(006,022-025) + 消す(20件) + 社内(031-035,044,054)
var REORG_DELETE_IDS = [
  '2026-002', '2026-004', '2026-006', '2026-007', '2026-008', '2026-009',
  '2026-010', '2026-011', '2026-012', '2026-014', '2026-015', '2026-016',
  '2026-017', '2026-018', '2026-019', '2026-020', '2026-022', '2026-023',
  '2026-024', '2026-025', '2026-026', '2026-030', '2026-031', '2026-032',
  '2026-033', '2026-034', '2026-035', '2026-037', '2026-044', '2026-051',
  '2026-052', '2026-054',
];

function applyReorg(commit) {
  var dryRun = !commit;
  var tag = dryRun ? '[DRY-RUN 変更なし] ' : '[適用] ';
  console.log('===== プロジェクト名整理 一括適用 ' + (dryRun ? '（DRY-RUN）' : '（実適用）') + ' =====');

  var dict = getSheet('プロジェクト管理');
  var proj = getSheet(PM_SHEET_PROJECTS);
  if (!dict) { console.log('❌ プロジェクト管理シートがありません。中止。'); return; }
  if (!proj) { console.log('❌ projects シートがありません。先に setup() を実行してください。中止。'); return; }

  // --- 0. 案件グループシートを確保 ---
  var grp = getSheet(PM_SHEET_GROUPS);
  if (!grp) {
    if (dryRun) {
      console.log(tag + '案件グループシートが無いので作成します');
    } else {
      grp = getSS().insertSheet(PM_SHEET_GROUPS);
      grp.getRange(1, 1, 1, PM_GROUPS_HEADERS.length).setValues([PM_GROUPS_HEADERS]);
      console.log(tag + '案件グループシートを作成しました');
    }
  }

  // --- 1. 改名（両シート・冪等） ---
  console.log('----- 1. 改名 -----');
  var renamed = reorgRenameById_(dict, '正式名称', dryRun, tag) + reorgRenameById_(proj, '案件名', dryRun, tag);
  console.log('改名 ' + renamed + ' 件');

  // --- 2. 案件グループへ追記（登録済みはスキップ・冪等） ---
  console.log('----- 2. 案件グループ登録 -----');
  var added = 0, skipped = 0;
  var existingGids = {};
  if (grp && grp.getLastRow() >= 2) {
    var g = grp.getDataRange().getValues();
    var gh = g[0].map(function(h){ return String(h).trim(); });
    var cGid = gh.indexOf('グループID');
    for (var i = 1; i < g.length; i++) {
      var gid = String(g[i][cGid] || '').trim();
      if (gid) existingGids[gid] = true;
    }
  }
  var now = (typeof pmTodayYmd === 'function') ? pmTodayYmd() : '';
  var toAppend = [];
  REORG_GROUPS.forEach(function(r) {
    if (existingGids[r[0]]) { skipped++; return; }
    toAppend.push([r[0], r[1], r[2], r[3], '整理2026-07-07', now]);
    added++;
    console.log(tag + '登録: ' + r[2] + ' / ' + r[3] + (r[1] ? ' → ' + r[1] : '（案件なし・台帳のみ）'));
  });
  if (!dryRun && toAppend.length && grp) {
    grp.getRange(grp.getLastRow() + 1, 1, toAppend.length, PM_GROUPS_HEADERS.length).setValues(toAppend);
  }
  console.log('案件グループ 追加 ' + added + ' 件 / 登録済みスキップ ' + skipped + ' 件');

  // --- 3. 行削除（両シート・下から削除） ---
  console.log('----- 3. 行削除 -----');
  var del = reorgDeleteByIds_(dict, dryRun, tag) + reorgDeleteByIds_(proj, dryRun, tag);
  console.log('削除 ' + del + ' 行');

  console.log('===== 完了 =====');
  if (dryRun) {
    console.log('▶ 内容が仕分け表どおりなら applyReorgCommit() を実行してください。');
  } else {
    console.log('▶ 次に backfillProjectGroups()（DRY-RUN）→ backfillProjectGroupsCommit() を実行して、');
    console.log('  残した案件の施主/業者列のグループIDを案件グループへ移してください。');
  }
}

function applyReorgCommit() { return applyReorg(true); }

// シート内の案件ID一致行の名前列を REORG_RENAMES に沿って更新（冪等）。更新件数を返す。
function reorgRenameById_(sheet, nameHeader, dryRun, tag) {
  if (sheet.getLastRow() < 2) return 0;
  var data = sheet.getDataRange().getValues();
  var h = data[0].map(function(x){ return String(x).trim(); });
  var cId = h.indexOf('案件ID');
  var cName = h.indexOf(nameHeader);
  if (cId === -1 || cName === -1) {
    console.log('⚠ ' + sheet.getName() + ' に 案件ID/' + nameHeader + ' 列が見つかりません。スキップ。');
    return 0;
  }
  var map = {};
  REORG_RENAMES.forEach(function(r){ map[r[0]] = r[1]; });
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][cId] || '').trim();
    if (!map[id]) continue;
    var cur = String(data[i][cName] || '').trim();
    if (cur === map[id]) continue; // 改名済み
    console.log(tag + sheet.getName() + ' ' + id + ': 「' + cur + '」→「' + map[id] + '」');
    if (!dryRun) sheet.getRange(i + 1, cName + 1).setValue(map[id]);
    count++;
  }
  return count;
}

// シート内の REORG_DELETE_IDS 一致行を下から削除（冪等）。削除件数を返す。
function reorgDeleteByIds_(sheet, dryRun, tag) {
  if (sheet.getLastRow() < 2) return 0;
  var data = sheet.getDataRange().getValues();
  var h = data[0].map(function(x){ return String(x).trim(); });
  var cId = h.indexOf('案件ID');
  if (cId === -1) {
    console.log('⚠ ' + sheet.getName() + ' に 案件ID 列が見つかりません。スキップ。');
    return 0;
  }
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][cId] || '').trim();
    if (REORG_DELETE_IDS.indexOf(id) === -1) continue;
    var label = String(data[i][1] || data[i][0] || '').trim();
    console.log(tag + sheet.getName() + ' 行削除: ' + id + ' ' + label);
    rows.push(i + 1);
  }
  if (!dryRun) {
    for (var j = rows.length - 1; j >= 0; j--) sheet.deleteRow(rows[j]);
  }
  return rows.length;
}
