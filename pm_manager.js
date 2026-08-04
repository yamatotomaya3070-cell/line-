// ==========================================
// ProjectManager：projects シートのCRUD・案件解決・ステータス正規化・自動更新適用
//   依存: pm_core.js, 既存 identifyProject / getProjectCandidates /
//         getProjectNameByGroupId / registerNewProject / getProjectData
// ==========================================

// ---- 案件ID採番（YYYY-NNN） ----
function pmGenProjectId() {
  var year = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy');
  var rows = pmReadObjects(PM_SHEET_PROJECTS);
  var max = 0;
  rows.forEach(function(r) {
    var m = String(r['案件ID'] || '').match(new RegExp('^' + year + '-(\\d+)$'));
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return year + '-' + String(max + 1).padStart(3, '0');
}

// ---- 取得 ----
function pmGetProjectByName(name) {
  if (!name) return null;
  var rows = pmReadObjects(PM_SHEET_PROJECTS);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['案件名']).trim() === String(name).trim()) return rows[i];
  }
  return null;
}

function pmGetProjectById(projectId) {
  if (!projectId) return null;
  var rows = pmReadObjects(PM_SHEET_PROJECTS);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['案件ID']).trim() === String(projectId).trim()) return rows[i];
  }
  return null;
}

// projects に案件レコードを保証（無ければ作成）。既存「プロジェクト管理」にも登録して
// グループ紐付け・identifyProject を機能させる。戻り値: projects の行オブジェクト
function pmEnsureProjectRecord(name, opts) {
  opts = opts || {};
  // 同時実行時にも案件ID・Driveフォルダを一組だけ作る。
  var locked = pmWithLock(function() { return pmEnsureProjectRecordUnlocked_(name, opts); }, 30000);
  if (!locked.ok) throw new Error('案件作成のロックを取得できませんでした');
  return locked.result;
}

function pmEnsureProjectRecordUnlocked_(name, opts) {
  var existing = pmGetProjectByName(name);
  if (existing) return existing;
  var now = pmTodayYmd();
  var id = pmGenProjectId();
  pmAppendRowFields(PM_SHEET_PROJECTS, {
    '案件ID': id,
    '案件名': name,
    'クライアント名': opts.client || '',
    '担当者': opts.assignee || '',
    'LINEグループID': opts.groupId || '',
    '現在フェーズ': opts.phase || '営業',
    '請求ステータス': '未請求',
    '入金ステータス': '未入金',
    '最終更新日': now,
    '備考': opts.remark || '',
    '最終更新者': opts.updatedBy || '',
    '作成日': now,
    '更新日': now,
  });
  // 既存「プロジェクト管理」へも反映（グループ紐付け用）。
  // Driveフォルダは下の pmEnsureProjectFolder が正規構成で1つだけ作るため、
  // registerNewProject 側のフラットなフォルダ作成は抑止（二重フォルダ防止：設計書§5）。
  // ※registerNewProject が失敗しても projects 側は作成済み。整合のためログを残し処理は継続。
  try {
    registerNewProject(name, opts.groupId || '', true);
    // 辞書(プロジェクト管理)に同一案件IDを刻む＝projectsと二面一体（フェーズ1）。
    // 案件ID列が未作成(setup未実行)なら pmStampProjectIdInDict は何もしない（安全）。
    pmStampProjectIdInDict(name, id);
  }
  catch (e) { console.error('pmEnsureProjectRecord registerNewProject 失敗（プロジェクト管理シート未反映の可能性）:', e.message); }
  // 案件フォルダ（新階層 ROOT/案件名）を冪等作成。pmEnsureProjectFolder は旧「プロジェクト管理」層を作らない。
  // 店舗名/6_データ は店舗登録時 pmEnsureStoreRecord が作る。skipDriveFolder=true でフォルダ生成自体を抑止できる。
  if (!opts.skipDriveFolder) {
    try { pmEnsureProjectFolder(id, name, { source: opts.source || 'unknown', actor: opts.updatedBy || '' }); }
    catch (e) { console.error('pmEnsureProjectRecord pmEnsureProjectFolder:', e.message); }
  }
  return pmGetProjectByName(name);
}

// ---- 自動確定のしきい値（設計書§13 / 確定仕様） ----
var PM_AUTO_RESOLVE_MIN    = 90; // この信頼度未満は自動確定しない（候補提示）
var PM_SECOND_CANDIDATE_MAX = 70; // 2番手がこの値以上なら「実質1件」とみなさず候補提示

// 2つの案件名が実質同一を指すか（部分一致で寄せ判定）
function pmNamesLooselyEqual(a, b) {
  var x = String(a || '').replace(/\s/g, '');
  var y = String(b || '').replace(/\s/g, '');
  if (!x || !y) return false;
  return x === y || x.indexOf(y) !== -1 || y.indexOf(x) !== -1;
}

// グループ紐付け案件(bound)と本文/ヒントの案件名が食い違うか判定。
//  食い違う（本文が別の既存案件を高確度で指す）→ 候補配列を返す。食い違わない→null。
function pmBoundNameConflict(text, hint, bound, groupId) {
  var probe = String(hint || '').trim();
  if (!probe) return null;                       // 本文に案件名の手掛かりなし → boundを信頼
  if (pmNamesLooselyEqual(probe, bound)) return null; // 本文の案件名 == bound → 矛盾なし

  // グループ紐付けを外して本文ヒント単体で識別（boundに引っ張られないよう groupId は渡さない）
  var idH = identifyProject(probe, '', probe);
  if (idH && idH.confidence >= PM_AUTO_RESOLVE_MIN && idH.name && idH.name !== '未分類' &&
      !pmNamesLooselyEqual(idH.name, bound)) {
    // 紐付け案件と本文案件が別物 → 人間に選ばせる
    return [
      { name: bound,     confidence: 95, reason: 'グループ紐付け案件' },
      { name: idH.name,  confidence: idH.confidence, reason: '本文の案件名' },
    ];
  }
  return null;
}

// ---- 案件解決（誤名寄せ防止） ----
//  戻り値: { status:'resolved', name } | { status:'ambiguous', candidates } | { status:'none' }
//  ルール（確定仕様）:
//   - グループ紐付け済み → 原則確定。ただし本文が別案件を高確度で指すなら確認に回す。
//   - 自動確定は信頼度 90 以上、かつ実質1件（2番手が弱い）場合のみ。
//   - 類似案件が複数 → 候補提示。
function pmResolveProject(text, groupId, hint) {
  // グループ紐付け済み
  var bound = getProjectNameByGroupId(groupId);
  if (bound) {
    var conflict = pmBoundNameConflict(text, hint, bound, groupId);
    if (conflict) return { status: 'ambiguous', candidates: conflict };
    return { status: 'resolved', name: bound };
  }

  var candidates = getProjectCandidates(text, groupId, hint, 4);
  if (!candidates.length) return { status: 'none' };

  var top = candidates[0];
  var second = candidates[1];
  var effectivelySingle = !second || second.confidence < PM_SECOND_CANDIDATE_MAX ||
                          pmNamesLooselyEqual(top.name, second.name);
  if (top && top.confidence >= PM_AUTO_RESOLVE_MIN && top.name && top.name !== '未分類' && effectivelySingle) {
    return { status: 'resolved', name: top.name };
  }
  return { status: 'ambiguous', candidates: candidates };
}

// ---- ステータス正規化（フェーズ内の正規ステータスへ寄せる） ----
//  戻り値: { phase, status, matched(boolean) }
function pmNormalizeStatus(phase, status) {
  var result = { phase: phase || '', status: status || '', matched: false };
  if (!status) return result;
  var clean = String(status).replace(/\s/g, '');

  // フェーズ指定があればその中から、無ければ全フェーズから探索
  var phasesToScan = phase && PM_PHASES[phase] ? [phase] : Object.keys(PM_PHASES);
  for (var p = 0; p < phasesToScan.length; p++) {
    var list = PM_PHASES[phasesToScan[p]];
    for (var i = 0; i < list.length; i++) {
      var cand = list[i].replace(/\s/g, '');
      if (clean === cand || clean.indexOf(cand) !== -1 || cand.indexOf(clean) !== -1) {
        return { phase: phasesToScan[p], status: list[i], matched: true };
      }
    }
  }
  return result; // 未一致：元の値を保持（呼び出し側で備考化を検討）
}

// ---- 自動更新の適用（安全項目のみ。金額/請求/入金は含めない） ----
//  parsed: pm_parser の戻り値。projectName: 確定済み案件名。
//  戻り値: { projectId, applied: {変更項目}, project: 行オブジェクト }
function pmApplyUpdate(parsed, projectName, sender, groupId, srcMsg) {
  var rec = pmEnsureProjectRecord(projectName, {
    assignee: parsed.assignee || '',
    groupId: groupId || '',
    updatedBy: sender || '',
  });
  if (!rec) return null;

  var fields = {};
  var applied = {};

  // フェーズ・ステータス
  if (parsed.status || parsed.phase) {
    var norm = pmNormalizeStatus(parsed.phase, parsed.status);
    if (norm.phase) { fields['現在フェーズ'] = norm.phase; applied['現在フェーズ'] = norm.phase; }
    if (norm.status) {
      if (norm.matched) {
        var col = PM_PHASE_COLUMN[norm.phase] || '営業ステータス';
        fields[col] = norm.status; applied[col] = norm.status;
      } else {
        // 正規ステータス未一致 → ステータス列は更新せず備考に原文を残す（誤更新防止）。
        // 通常はRouter側で確認待ちに回るため、ここは二重の安全策。
        fields['備考'] = (rec['備考'] ? rec['備考'] + ' / ' : '') + '要確認ステータス:' + parsed.status;
        applied['要確認ステータス'] = parsed.status;
      }
    }
  }

  // 安全項目
  var safeMap = {
    'next_action': '次回アクション',
    'next_action_due_date': '次回アクション期限',
    'meeting_date': '打合せ日',
    'construction_start_date': '着工予定日',
    'handover_date': '引き渡し予定日',
    'assignee': '担当者',
    'client_name': 'クライアント名',
  };
  Object.keys(safeMap).forEach(function(k) {
    if (parsed[k]) { fields[safeMap[k]] = parsed[k]; applied[safeMap[k]] = parsed[k]; }
  });
  if (parsed.remark) {
    fields['備考'] = (fields['備考'] || rec['備考'] || '');
    fields['備考'] = (fields['備考'] ? fields['備考'] + ' / ' : '') + parsed.remark;
    applied['備考'] = parsed.remark;
  }

  // メタ情報（常に更新）
  var today = pmTodayYmd();
  fields['最終更新日']   = today;
  fields['最終更新者']   = sender || '';
  fields['更新元メッセージ'] = srcMsg || '';
  fields['更新日']       = today;

  pmWriteRowFields(PM_SHEET_PROJECTS, rec._row, fields);
  pmAddLog(rec['案件ID'], parsed.intent || 'project_update', applied, srcMsg, sender, groupId, 'auto');

  return { projectId: rec['案件ID'], applied: applied, project: pmGetProjectById(rec['案件ID']) };
}

// ---- プロジェクトアーカイブ ----
function pmArchiveProject(projectId, updatedBy, kind) {
  if (!projectId) return { ok: false, msg: '案件IDが必要です' };
  var locked = pmWithLock(function() {
    var sh = getSheet(PM_SHEET_PROJECTS);
    if (!sh) return { ok: false, msg: 'projectsシートが見つかりません' };
    var data = sh.getDataRange().getValues();
    if (!data || !data.length) return { ok: false, msg: 'projectsシートが空です' };
    var h = data[0].map(function(x) { return String(x).trim(); });
    var cId = h.indexOf('案件ID');
    var cDel = h.indexOf('取消');
    var cUpd = h.indexOf('最終更新日');
    var cWho = h.indexOf('最終更新者');
    var cUpdate = h.indexOf('更新日');
    if (cId === -1) return { ok: false, msg: '案件ID列がありません（setup実行を）' };
    if (cDel === -1) return { ok: false, msg: '取消列がありません（setup実行を）' };
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][cId] || '').trim() === String(projectId).trim()) {
        var current = String(data[i][cDel] || '').trim().toUpperCase();
        if (current === 'TRUE' || current === '1' || current === '✓') {
          return { ok: true, msg: '既にアーカイブ済みです' };
        }
        // mark cancel flag (legacy cell write for compatibility)
        sh.getRange(i + 1, cDel + 1).setValue('TRUE');
        if (cUpd !== -1) sh.getRange(i + 1, cUpd + 1).setValue(pmTodayYmd());
        if (cUpdate !== -1) sh.getRange(i + 1, cUpdate + 1).setValue(pmTodayYmd());
        if (cWho !== -1 && String(updatedBy || '').trim()) sh.getRange(i + 1, cWho + 1).setValue(String(updatedBy).trim());

        // additional archive metadata (will only write if those headers exist)
        var archiveFields = {
          '完了日': pmTodayYmd(),
          '完了者': String(updatedBy || ''),
          '施工進捗': '100'
        };
        try { pmWriteRowFields(PM_SHEET_PROJECTS, i + 1, archiveFields); } catch (e3) {}

        // log all applied changes
        var appliedChanges = { '取消': 'TRUE' };
        try {
          if (archiveFields['完了日']) appliedChanges['完了日'] = archiveFields['完了日'];
          if (archiveFields['完了者']) appliedChanges['完了者'] = archiveFields['完了者'];
          if (archiveFields['施工進捗']) appliedChanges['施工進捗'] = archiveFields['施工進捗'];
        } catch (e4) {}
        try { pmAddLog(projectId, 'archive', appliedChanges, '', updatedBy || '', '', kind || 'auto'); } catch (e2) {}
        return { ok: true };
      }
    }
    return { ok: false, msg: '案件が見つかりません' };
  }, 30000);
  if (!locked.ok) return { ok: false, msg: 'アーカイブ処理が混み合っています。再試行してください' };
  return locked.result;
}

// ---- アーカイブ解除（案件再開） ----
function pmUnarchiveProject(projectId, updatedBy, kind) {
  if (!projectId) return { ok: false, msg: '案件IDが必要です' };
  var locked = pmWithLock(function() {
    var sh = getSheet(PM_SHEET_PROJECTS);
    if (!sh) return { ok: false, msg: 'projectsシートが見つかりません' };
    var data = sh.getDataRange().getValues();
    if (!data || !data.length) return { ok: false, msg: 'projectsシートが空です' };
    var h = data[0].map(function(x) { return String(x).trim(); });
    var cId = h.indexOf('案件ID');
    var cDel = h.indexOf('取消');
    var cUpd = h.indexOf('最終更新日');
    var cWho = h.indexOf('最終更新者');
    var cUpdate = h.indexOf('更新日');
    if (cId === -1) return { ok: false, msg: '案件ID列がありません（setup実行を）' };
    if (cDel === -1) return { ok: false, msg: '取消列がありません（setup実行を）' };
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][cId] || '').trim() === String(projectId).trim()) {
        var current = String(data[i][cDel] || '').trim().toUpperCase();
        if (!(current === 'TRUE' || current === '1' || current === '✓')) {
          // 二重解除防止：既にアーカイブでない場合は操作失敗扱いにする
          return { ok: false, msg: 'アーカイブされていません' };
        }
        sh.getRange(i + 1, cDel + 1).setValue('');
        if (cUpd !== -1) sh.getRange(i + 1, cUpd + 1).setValue(pmTodayYmd());
        if (cUpdate !== -1) sh.getRange(i + 1, cUpdate + 1).setValue(pmTodayYmd());
        if (cWho !== -1 && String(updatedBy || '').trim()) sh.getRange(i + 1, cWho + 1).setValue(String(updatedBy).trim());

        // clear archive metadata where possible
        var clearFields = { '完了日': '', '完了者': '' };
        try { pmWriteRowFields(PM_SHEET_PROJECTS, i + 1, clearFields); } catch (e5) {}

        // NOTE: 施工進捗の単純クリアは行わない（完了前進捗の保存・統一再計算未実装のため）。
        // TODO: implement recalculation or restore of pre-completion progress on unarchive.
        try { pmAddLog(projectId, 'unarchive', { '取消': '', '完了日': '', '完了者': '' }, '', updatedBy || '', '', kind || 'auto'); } catch (e6) {}
        return { ok: true };
      }
    }
    return { ok: false, msg: '案件が見つかりません' };
  }, 30000);
  if (!locked.ok) return { ok: false, msg: 'アーカイブ解除が混み合っています。再試行してください' };
  return locked.result;
}

// ---- アーカイブ一覧取得 ----
function pmListArchivedProjects() {
  var rows = pmReadObjects(PM_SHEET_PROJECTS);
  return rows.filter(function(r) {
    var canceled = String(r['取消'] || '').trim().toUpperCase();
    return (canceled === 'TRUE' || canceled === '1' || canceled === '✓' || canceled === 'YES');
  });
}

// ---- 進捗率統一（プロジェクト単位の再計算） ----
// 設計: 案件の `施工進捗` は店舗単位（PM_SHEET_STORES の '施工進捗'）の平均で算出する。
// - 存在する店舗の進捗を数値で集計し、平均（四捨五入）を projects シートに書き戻す。
// - 店舗が無い場合は既存の案件フィールドを変更しない。
function pmRecalculateProjectProgress(projectId) {
  if (!projectId) return { ok: false, msg: '案件IDが必要です' };
  try {
    var stores = pmReadObjects('stores'); // PM_SHEET_STORES 名は 'stores' または '店舗' 環境に依存
    var vals = [];
    stores.forEach(function(s) {
      if (String(s['案件ID'] || '').trim() === String(projectId).trim()) {
        var p = Number(s['施工進捗']);
        if (!isNaN(p)) vals.push(p);
      }
    });
    if (!vals.length) return { ok: false, msg: '対象案件に店舗進捗がありません' };
    var sum = vals.reduce(function(a, b) { return a + b; }, 0);
    var avg = Math.round(sum / vals.length);
    var rec = pmGetProjectById(projectId);
    if (!rec) return { ok: false, msg: '案件が見つかりません' };
    pmWriteRowFields(PM_SHEET_PROJECTS, rec._row, { '施工進捗': avg, '最終更新日': pmTodayYmd() });
    pmAddLog(projectId, 'progress_recalc', { '施工進捗': avg }, '', 'system', '', 'auto');
    return { ok: true, progress: avg };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function pmRecalculateAllProjects() {
  var rows = pmReadObjects(PM_SHEET_PROJECTS);
  var results = [];
  rows.forEach(function(r) {
    var res = pmRecalculateProjectProgress(r['案件ID']);
    results.push({ projectId: r['案件ID'], result: res });
  });
  return results;
}

// ---- 複数人タスク：タスク作成・担当更新のユーティリティ ----
// タスク管理シートのフォーマットに依存（既存は 'タスク管理' を使用）。
// 既存レコードの担当者列は文字列として複数名をカンマ区切り等で保持する方針。
function pmCreateTaskMulti(task) {
  // task: { projectId, projectName, content, assignees: [name,...], deadline }
  if (!task || !task.content) return { ok: false, msg: 'タスク内容が必要です' };
  var sheet = getSheet('タスク管理');
  if (!sheet) return { ok: false, msg: 'タスク管理シートが見つかりません' };
  var assignees = (task.assignees || []).filter(function(x){return String(x||'').trim();}).map(function(s){return String(s).trim();});
  var assigneeStr = assignees.join(' / ');
  var id = generateId();
  var row = [id, task.projectName || '', task.content, assigneeStr, task.deadline || '', 'confirmed'];
  sheet.appendRow(row);
  return { ok: true, taskId: id };
}

function pmAssignTaskMultiple(taskId, assignees) {
  if (!taskId) return { ok: false, msg: 'taskIdが必要です' };
  var sheet = getSheet('タスク管理');
  if (!sheet) return { ok: false, msg: 'タスク管理シートが見つかりません' };
  var data = sheet.getDataRange().getValues();
  var idx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '') === String(taskId)) { idx = i + 1; break; }
  }
  if (idx === -1) return { ok: false, msg: 'タスクが見つかりません' };
  var s = (assignees || []).filter(function(x){return String(x||'').trim();}).join(' / ');
  sheet.getRange(idx, 4).setValue(s); // 担当者は4列目（既存フォーマット依存）
  return { ok: true };
}

// ---- LINE 活動履歴（軽量ログ） ----
// 追加シート: PM_SHEET_LINE_ACTIVITY
function pmRecordLineActivity(evt) {
  // evt: { ts, type, userId, displayName, groupId, text, projectId, raw }
  try {
    pmAppendRowFields(PM_SHEET_LINE_ACTIVITY, {
      '日時': fmtDT(new Date(evt.ts || Date.now())),
      '種別': evt.type || '',
      'ユーザーID': evt.userId || '',
      '表示名': evt.displayName || '',
      'グループID': evt.groupId || '',
      'メッセージ': evt.text || '',
      '案件ID': evt.projectId || '',
      'raw': JSON.stringify(evt.raw || {})
    });
    return { ok: true };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function pmListLineActivities(projectId, limit) {
  limit = Number(limit) || 100;
  var rows = pmReadObjects(PM_SHEET_LINE_ACTIVITY);
  if (projectId) rows = rows.filter(function(r){ return String(r['案件ID']||'') === String(projectId); });
  rows.sort(function(a,b){ return new Date(b['日時']).getTime() - new Date(a['日時']).getTime(); });
  return rows.slice(0, limit);
}

// ---- Drive アーカイブ dry-run 設計（移動は行わない） ----
function pmDriveArchiveDryRun(projectId) {
  // returns a detailed plan of actions that WOULD be taken for Drive archive (dry-run only)
  var proj = pmGetProjectById(projectId);
  if (!proj) return { ok: false, msg: '案件が見つかりません' };
  // collect candidate folders from known fields (no API calls)
  var candidates = [];
  if (proj['driveProjectFolderId']) candidates.push({ key: 'driveProjectFolderId', id: proj['driveProjectFolderId'], url: proj['driveProjectFolderUrl'] || '' });
  if (proj['DriveフォルダID']) candidates.push({ key: 'DriveフォルダID', id: proj['DriveフォルダID'], url: proj['DriveフォルダURL'] || '' });

  // design: move to archive folder structure: /Archive/YYYY/<案件ID>_案件名
  var year = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy');
  var destPath = '/Archive/' + year + '/' + projectId + '_' + (proj['案件名'] || '').replace(/\//g, '_');

  var actions = candidates.map(function(c) {
    return {
      from: c,
      to: { path: destPath },
      note: 'dry-run: no changes made',
      preState: { fieldKey: c.key, id: c.id, url: c.url }
    };
  });

  // prepare rollback design: since we do not perform moves, this is a textual plan
  var rollbackSteps = actions.map(function(a) {
    return {
      action: 'restore',
      description: 'If an archive move succeeded then failed later, move back from ' + a.to.path + ' to original folder recorded in ' + a.preState.fieldKey,
      restoreFrom: a.to.path,
      restoreTo: a.preState.id || ('original path from ' + a.preState.fieldKey),
      note: 'Requires Drive API + credentials; ensure idempotent checks and retry/backoff.'
    };
  });

  var preSnapshot = {
    projectId: projectId,
    projectName: proj['案件名'] || '',
    timestamp: fmtDT(new Date()),
    projectFields: {
      driveProjectFolderId: proj['driveProjectFolderId'] || '',
      driveProjectFolderUrl: proj['driveProjectFolderUrl'] || '',
      DriveフォルダID: proj['DriveフォルダID'] || '',
      DriveフォルダURL: proj['DriveフォルダURL'] || ''
    }
  };

  return {
    ok: true,
    projectId: projectId,
    projectName: proj['案件名'] || '',
    candidates: candidates,
    actions: actions,
    preSnapshot: preSnapshot,
    rollbackPlan: { summary: 'Move files/folders back to original locations using recorded IDs/paths', steps: rollbackSteps }
  };
}

// ---- Drive アーカイブ実行設計（dry-run を基にした実行前チェック用） ----
function pmDriveArchivePlan(projectId) {
  // returns a checklist and preconditions required before doing actual moves
  var dry = pmDriveArchiveDryRun(projectId);
  if (!dry.ok) return dry;
  var checks = [];
  if (!dry.candidates.length) checks.push({ ok: false, msg: '移動対象となる Drive フォルダ ID が見つかりません' });
  else checks.push({ ok: true, msg: dry.candidates.length + ' 個の候補フォルダを確認' });

  checks.push({ ok: true, msg: 'アーカイブ先パス: ' + (dry.actions && dry.actions[0] && dry.actions[0].to.path || '') });
  checks.push({ ok: true, msg: '事前スナップショットを取得済み（返却データ参照）' });
  checks.push({ ok: false, msg: '実行には Drive API 権限と運用承認が必要（dry-run のまま実行はしない）' });

  return { ok: true, projectId: projectId, planChecks: checks, dryRun: dry };
}

// ---- 履歴記録 ----
function pmAddLog(projectId, intent, changes, srcMsg, updatedBy, groupId, kind) {
  pmAppendRowFields(PM_SHEET_LOGS, {
    'ログID': generateId(),
    '日時': fmtDT(new Date()),
    '案件ID': projectId || '',
    'intent': intent || '',
    '変更項目': JSON.stringify(changes || {}),
    '更新元メッセージ': srcMsg || '',
    '更新者': updatedBy || '',
    'グループID': groupId || '',
    '適用区分': kind || 'auto',
  });
}

// ---- 表示用：適用結果の要約テキスト ----
function pmFormatApplied(projectId, applied) {
  var lines = ['✅ 案件「' + projectId + '」を更新しました。'];
  Object.keys(applied).forEach(function(k) {
    lines.push('・' + k + '：' + applied[k]);
  });
  return lines.join('\n');
}
