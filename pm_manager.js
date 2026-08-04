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
  var sh = getSheet(PM_SHEET_PROJECTS);
  var data = sh.getDataRange().getValues();
  var h = data[0].map(function(x) { return String(x).trim(); });
  var cId = h.indexOf('案件ID');
  var cDel = h.indexOf('取消');
  var cUpd = h.indexOf('最終更新日');
  var cWho = h.indexOf('最終更新者');
  var cUpdate = h.indexOf('更新日');
  if (cDel === -1) return { ok: false, msg: '取消列がありません（setup実行を）' };
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][cId] || '').trim() === String(projectId).trim()) {
      var current = String(data[i][cDel] || '').trim().toUpperCase();
      if (current === 'TRUE' || current === '1' || current === '✓') {
        return { ok: true, msg: '既にアーカイブ済みです' };
      }
      sh.getRange(i + 1, cDel + 1).setValue('TRUE');
      if (cUpd !== -1) sh.getRange(i + 1, cUpd + 1).setValue(pmTodayYmd());
      if (cUpdate !== -1) sh.getRange(i + 1, cUpdate + 1).setValue(pmTodayYmd());
      if (cWho !== -1 && String(updatedBy || '').trim()) sh.getRange(i + 1, cWho + 1).setValue(String(updatedBy).trim());
      try { pmAddLog(projectId, 'archive', { '取消': 'TRUE' }, '', updatedBy || '', '', kind || 'auto'); } catch (e2) {}
      return { ok: true };
    }
  }
  return { ok: false, msg: '案件が見つかりません' };
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
