// ==========================================
// フェーズ1: 案件IDで「二面一体」化（プロジェクト管理 ⇔ projects）
//   - 辞書(プロジェクト管理)と進捗(projects)を案件IDで1対1リンクする。
//   - 既存データを壊さない。すべて追記/空欄埋めのみ。
//   - エントリ: backfillProjectIds()  ※引数なし=DRY-RUN(変更なし・計画のみ)
//              backfillProjectIds(true) で実適用。
//   全トップレベル名は pm/backfill 接頭辞で衝突回避。
// ==========================================

// 辞書(プロジェクト管理)の該当案件行に案件IDを刻む（空欄のときだけ）。
// 案件ID列が無い(setup未実行)なら何もしない＝安全。live作成経路から呼ばれる。
function pmStampProjectIdInDict(name, projectId) {
  if (!name || !projectId) return;
  var sh = getSheet('プロジェクト管理');
  if (!sh || sh.getLastRow() < 1) return;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim(); });
  var cId = headers.indexOf('案件ID');
  var cFormal = headers.indexOf('正式名称');
  var cAbbr = headers.indexOf('略称');
  if (cId === -1 || cFormal === -1) return;
  var key = String(name).replace(/\s/g, '');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var f = String(data[i][cFormal] || '').replace(/\s/g, '');
    var a = cAbbr !== -1 ? String(data[i][cAbbr] || '').replace(/\s/g, '') : '';
    if ((f && f === key) || (a && a === key)) {
      if (!String(data[i][cId] || '').trim()) sh.getRange(i + 1, cId + 1).setValue(projectId);
      return;
    }
  }
}

// 一括バックフィル。commit=falsy(既定)はDRY-RUN。
function backfillProjectIds(commit) {
  var dryRun = !commit;
  var sh = getSheet('プロジェクト管理');
  if (!sh) { console.log('プロジェクト管理シートがありません'); return; }
  if (sh.getLastRow() < 2) { console.log('プロジェクト管理にデータがありません'); return; }

  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).trim(); });
  var cId = headers.indexOf('案件ID');
  var cFormal = headers.indexOf('正式名称');
  if (cId === -1) { console.log('「案件ID」列がありません。先に setup() を実行してください。'); return; }
  if (cFormal === -1) { console.log('「正式名称」列がありません'); return; }

  // projects: 案件名(空白除去) -> 案件ID
  var nameToId = {};
  pmReadObjects(PM_SHEET_PROJECTS).forEach(function(r){
    var n = String(r['案件名'] || '').replace(/\s/g, '');
    if (n && r['案件ID']) nameToId[n] = r['案件ID'];
  });

  var rep = { linked: 0, created: 0, already: 0, skipped: 0, details: [] };

  for (var i = 1; i < data.length; i++) {
    var rowNum = i + 1;
    var formal = String(data[i][cFormal] || '').trim();
    if (String(data[i][cId] || '').trim()) { rep.already++; continue; }
    if (!formal) { rep.skipped++; rep.details.push('行' + rowNum + ': 正式名称が空→スキップ'); continue; }

    var key = formal.replace(/\s/g, '');
    var id = nameToId[key];
    if (id) {
      if (!dryRun) sh.getRange(rowNum, cId + 1).setValue(id);
      rep.linked++;
      rep.details.push('行' + rowNum + ': ' + formal + ' → 既存projects ' + id + (dryRun ? '（予定）' : ''));
    } else {
      if (!dryRun) {
        var newId = pmGenProjectId();
        pmAppendRowFields(PM_SHEET_PROJECTS, {
          '案件ID': newId, '案件名': formal, 'クライアント名': '',
          '現在フェーズ': '営業', '請求ステータス': '未請求', '入金ステータス': '未入金',
          '最終更新日': pmTodayYmd(), '備考': 'backfillで作成', '作成日': pmTodayYmd(), '更新日': pmTodayYmd(),
        });
        nameToId[key] = newId;
        sh.getRange(rowNum, cId + 1).setValue(newId);
        rep.created++;
        rep.details.push('行' + rowNum + ': ' + formal + ' → projects新規作成 ' + newId);
      } else {
        rep.created++;
        rep.details.push('行' + rowNum + ': ' + formal + ' → projects新規作成（予定）');
      }
    }
  }

  console.log((dryRun ? '[DRY-RUN 変更なし] ' : '[適用済] ') +
    '連携:' + rep.linked + ' / projects新規:' + rep.created + ' / 既にID有:' + rep.already + ' / スキップ:' + rep.skipped);
  rep.details.forEach(function(d){ console.log('  ' + d); });
  if (dryRun) console.log('▶ 問題なければ backfillProjectIdsCommit を実行して適用してください。');
  return rep;
}

// GASエディタの「実行」ボタンは引数を渡せないため、適用用のラッパーを用意。
//   backfillProjectIds()        … 引数なし＝DRY-RUN（変更なし・計画のみ）
//   backfillProjectIdsCommit()  … 実適用（ドロップダウンから実行可）
function backfillProjectIdsCommit() {
  return backfillProjectIds(true);
}

// ==========================================
// 区分の自動分類（案件/社内/その他）
//   - 社内・連絡用グループをボードから除外するため projects.区分 を埋める。
//   - 既に区分が入っている行は触らない（手動調整を尊重）。空欄のみ自動判定。
//   - classifyProjectKinds()        … DRY-RUN（変更なし・計画のみ）
//     classifyProjectKindsCommit()  … 実適用
// ==========================================
// ユーザー判定(2026-06-30)に合わせて調整：
//  - 「社内」「是正」「水道局/協議/行政」は除外（是正=施工手直し/水道局協議=実案件、アムヨ…社内=案件）
//  - 真の社内は WBF/WBG/経理/内部 で正確に検出（社内経理→経理、WBグループ内部→内部 で漏れない）
var PM_INTERNAL_KEYWORDS = ['WBF', 'WBG', '経理', '内部', '事務スタッフ'];
var PM_OTHER_KEYWORDS    = [];  // その他は自動判定しない（必要時に手動 or 参加時UIで設定）

function pmGuessKind_(name, remark) {
  var s = String(name || '') + ' ' + String(remark || '');
  for (var i = 0; i < PM_INTERNAL_KEYWORDS.length; i++) if (s.indexOf(PM_INTERNAL_KEYWORDS[i]) !== -1) return '社内';
  for (var j = 0; j < PM_OTHER_KEYWORDS.length; j++) if (s.indexOf(PM_OTHER_KEYWORDS[j]) !== -1) return 'その他';
  return '案件';
}

function classifyProjectKinds(commit) {
  var dryRun = !commit;
  var sh = getSheet(PM_SHEET_PROJECTS);
  if (!sh || sh.getLastRow() < 2) { console.log('projectsにデータがありません'); return; }
  var data = sh.getDataRange().getValues();
  var h = data[0].map(function(x){ return String(x).trim(); });
  var cKind = h.indexOf('区分');
  var cName = h.indexOf('案件名');
  var cRemark = h.indexOf('備考');
  if (cKind === -1) { console.log('「区分」列がありません。先に setup() を実行してください。'); return; }
  if (cName === -1) { console.log('「案件名」列がありません'); return; }

  var rep = { 案件: 0, 社内: 0, その他: 0, skipped: 0, details: [] };
  for (var i = 1; i < data.length; i++) {
    var rowNum = i + 1;
    if (String(data[i][cKind] || '').trim()) { rep.skipped++; continue; } // 既設定は尊重
    var name = String(data[i][cName] || '').trim();
    if (!name) { rep.skipped++; continue; }
    var kind = pmGuessKind_(name, cRemark !== -1 ? data[i][cRemark] : '');
    rep[kind]++;
    if (kind !== '案件') rep.details.push(kind + ': ' + name + (dryRun ? '（予定）' : ''));
    if (!dryRun) sh.getRange(rowNum, cKind + 1).setValue(kind);
  }

  console.log((dryRun ? '[DRY-RUN 変更なし] ' : '[適用済] ') +
    '案件:' + rep.案件 + ' / 社内:' + rep.社内 + ' / その他:' + rep.その他 + ' / 既設定スキップ:' + rep.skipped);
  console.log('  （除外対象＝社内・その他の一覧）');
  rep.details.forEach(function(d){ console.log('  ' + d); });
  if (dryRun) console.log('▶ 案件のはずが社内/その他に入っている等あれば、適用後にシートで直してOK。問題なければ classifyProjectKindsCommit を実行。');
  return rep;
}

function classifyProjectKindsCommit() { return classifyProjectKinds(true); }
