// ==========================================
// 確認待ちワークフロー：金額・請求・入金・案件曖昧 の人間承認
//   依存: pm_core.js, pm_manager.js, 既存 sendLineMessage / sendQuickReplyPush /
//         sendLineReply / generateId / fmtDT / safeParseJson
// ==========================================

// 確認待ちへ登録。type: 'billing'|'amount'|'resolve'。戻り値: 確認ID
function pmQueuePending(type, parsed, projectId, projectName, sender, groupId, srcMsg, candidates) {
  var pid = generateId();
  pmAppendRowFields(PM_SHEET_PENDING, {
    '確認ID': pid,
    '作成日時': fmtDT(new Date()),
    '案件ID': projectId || '',
    '案件名候補': candidates ? JSON.stringify(candidates) : (projectName || ''),
    'type': type,
    'intent': parsed.intent || '',
    '抽出内容': JSON.stringify(parsed),
    '更新元メッセージ': srcMsg || '',
    '申請者': sender || '',
    'グループID': groupId || '',
    'status': type === 'resolve' ? 'awaiting_project' : 'awaiting_approval',
    '承認者': '',
    '処理日時': '',
  });
  return pid;
}

function pmGetPending(pendingId) {
  if (!pendingId) return null;
  var rows = pmReadObjects(PM_SHEET_PENDING);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['確認ID']).trim() === String(pendingId).trim()) return rows[i];
  }
  return null;
}

function pmSetPendingStatus(pendingId, status, approver) {
  var rec = pmGetPending(pendingId);
  if (!rec) return;
  pmWriteRowFields(PM_SHEET_PENDING, rec._row, {
    'status': status,
    '承認者': approver || '',
    '処理日時': fmtDT(new Date()),
  });
}

// 経理グループへ確認依頼を通知（承認/却下のQuick Reply付き）
function pmNotifyAccountingPending(pendingId, parsed, projectName, srcMsg) {
  var target = pmAccountingTarget();
  if (!target) { console.warn('pmNotifyAccountingPending: 通知先未設定'); return; }
  var amt = parsed.amount && parsed.amount.value != null
    ? (parsed.amount.type || '金額') + '：' + Number(parsed.amount.value).toLocaleString() + '円'
    : '';
  var dateLines = [];
  if (parsed.billing_due_date)        dateLines.push('請求予定日：' + parsed.billing_due_date);
  if (parsed.billing_done_date)       dateLines.push('請求済日：' + parsed.billing_done_date);
  if (parsed.payment_due_date)        dateLines.push('入金予定日：' + parsed.payment_due_date);
  if (parsed.payment_confirmed_date)  dateLines.push('入金確認日：' + parsed.payment_confirmed_date);
  if (parsed.billing_kind)            dateLines.push('請求種別：' + parsed.billing_kind);

  var body = '【確認待ち】金額・請求/入金の更新候補です。\n' +
    '案件：' + (projectName || '（未特定）') + '\n' +
    (amt ? amt + '\n' : '') +
    (dateLines.length ? dateLines.join('\n') + '\n' : '') +
    '\n報告原文：' + (srcMsg || '') +
    '\n\n内容を確認して承認/却下してください。';

  sendQuickReplyPush(target, body, [
    { type: 'action', action: { type: 'postback', label: '✅ 承認', data: 'action=pm_approve&id=' + pendingId } },
    { type: 'action', action: { type: 'postback', label: '❌ 却下', data: 'action=pm_reject&id=' + pendingId } },
  ]);
}

// ==========================================
// 承認権限（金額・請求・入金は経理/管理者のみ）
//   取得元: ① Script Property / PM設定 'PM_APPROVER_USER_IDS'（カンマ区切り）
//          ② メンバー管理シートの「権限/役割」列に 経理/管理者 等
//   どちらも未設定なら後方互換で許可（警告）。本番は①の設定を推奨。
// ==========================================
function pmListApproverUserIds() {
  var raw = pmProp('PM_APPROVER_USER_IDS');
  if (!raw) return [];
  return String(raw).split(/[,、\s]+/).map(function(s) { return s.trim(); }).filter(Boolean);
}

// メンバー管理の権限列で承認者か。列が無ければ null（判定不能）。
function pmHasApproverRole(userId) {
  if (!userId) return false;
  var sheet = getSheet('メンバー管理');
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var idx = pmHeaderIndex(sheet);
  var cId = idx['LINE ユーザーID'];
  var roleCol;
  Object.keys(idx).forEach(function(h) { if (/権限|役割|ロール|role/i.test(h)) roleCol = idx[h]; });
  if (cId === undefined || roleCol === undefined) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][cId] || '') === String(userId)) {
      return /経理|管理者|管理|admin|approver|承認/i.test(String(data[i][roleCol] || ''));
    }
  }
  return false;
}

// 金額・請求・入金の承認可否
function pmIsApprover(userId) {
  var list = pmListApproverUserIds();
  if (list.length) return list.indexOf(String(userId)) !== -1;
  var roleOk = pmHasApproverRole(userId);
  if (roleOk !== null) return roleOk;
  console.warn('pmIsApprover: 承認者未設定。PM_APPROVER_USER_IDS かメンバー管理の権限列を設定してください。暫定で許可します。');
  return true;
}

// 金額が絡む確認タイプ（承認権限が必要）
function pmIsFinancialType(type) {
  return type === 'billing' || type === 'payment' || type === 'amount';
}

// 金額/請求/入金を projects へ確定反映（承認後）。戻り値: applied マップ
//   updatedBy: 最終更新者に記録する名前（申請者。空なら列を触らない＝後方互換）
function pmApplyFinancialFields(parsed, proj, updatedBy) {
  var fields = {};
  var applied = {};
  if (parsed.amount && parsed.amount.value != null) {
    var typeCol = { '売上': '売上', '原価': '原価', '請求': '請求金額', '入金': '入金金額' }[parsed.amount.type];
    if (typeCol) { fields[typeCol] = parsed.amount.value; applied[typeCol] = parsed.amount.value; }
    // 利益額・利益率（売上・原価が揃えば計算）
    var sales = parsed.amount.type === '売上' ? parsed.amount.value : Number(proj['売上'] || 0);
    var cost  = parsed.amount.type === '原価' ? parsed.amount.value : Number(proj['原価'] || 0);
    if (sales && cost) {
      fields['利益額'] = sales - cost; applied['利益額'] = sales - cost;
      fields['利益率'] = Math.round((sales - cost) / sales * 1000) / 10 + '%';
    }
  }
  if (parsed.billing_due_date)  { fields['請求予定日'] = parsed.billing_due_date; applied['請求予定日'] = parsed.billing_due_date; }
  if (parsed.billing_done_date) { fields['請求済日'] = parsed.billing_done_date; fields['請求ステータス'] = '請求済'; applied['請求済日'] = parsed.billing_done_date; }
  else if (parsed.billing_due_date && pmIsBlank(proj['請求ステータス'])) { fields['請求ステータス'] = '未請求'; }
  if (parsed.payment_due_date)       { fields['入金予定日'] = parsed.payment_due_date; applied['入金予定日'] = parsed.payment_due_date; }
  if (parsed.payment_confirmed_date) { fields['入金確認日'] = parsed.payment_confirmed_date; fields['入金ステータス'] = '入金済'; applied['入金確認日'] = parsed.payment_confirmed_date; }
  else if (parsed.payment_due_date && pmIsBlank(proj['入金ステータス'])) { fields['入金ステータス'] = '未入金'; }

  var today = pmTodayYmd();
  fields['最終更新日'] = today;
  fields['更新日'] = today;
  if (updatedBy) fields['最終更新者'] = updatedBy;
  pmWriteRowFields(PM_SHEET_PROJECTS, proj._row, fields);
  return applied;
}

// 承認の中核（UI非依存）。LINE経路（pmApprovePending）とアプリジョブ経路（pm_jobs.js）の両方から呼ぶ。
//   - 金額/請求/入金（financial）: 経理/管理者のみ。
//   - project_field（低信頼度の進捗）: 報告者本人/メンバーが内容を確認して反映。
//   - 二重実行・再処理防止: LockService で直列化し、status を再確認してから1回だけ適用。
//   canApproveFinancialFn: 金額系を承認できるか返す関数（LINE=ユーザーID判定／アプリ=メール役割判定と文脈が異なるため注入）
//   戻り値 code: lock_timeout | notfound | done | forbidden | parse_error | no_project | apply_failed | ok_project | ok_financial
function pmApprovePendingCore(pendingId, approver, canApproveFinancialFn) {
  var locked = pmWithLock(function() {
    var rec = pmGetPending(pendingId);
    if (!rec) return { code: 'notfound' };
    if (rec['status'] !== 'awaiting_approval') return { code: 'done', status: rec['status'] };

    // 権限チェック（金額系のみ）。NGなら状態を変えずに離脱。
    if (pmIsFinancialType(rec['type']) && !canApproveFinancialFn()) return { code: 'forbidden' };

    var parsed = safeParseJson(rec['抽出内容']);
    if (!parsed || typeof parsed !== 'object') return { code: 'parse_error' };

    // project_field（進捗の確認反映）
    if (rec['type'] === 'project_field') {
      var pname = rec['案件名候補'] || parsed.project_name || '';
      if (!pname) return { code: 'no_project' };
      var r = pmApplyUpdate(parsed, pname, approver, rec['グループID'], rec['更新元メッセージ']);
      if (!r) return { code: 'apply_failed' };
      pmSetPendingStatus(pendingId, 'applied', approver);
      try { pmSyncCalendar(r.projectId); } catch (e) { console.error('pmApprovePending(project_field) calendar:', e.message); }
      try { pmCreateFollowupTask(r.projectId, pname, parsed, approver, rec['グループID']); } catch (e) { console.error('pmApprovePending(project_field) taskbridge:', e.message); }
      return { code: 'ok_project', projectName: pname, applied: r.applied };
    }

    // 金額/請求/入金
    var proj = pmGetProjectById(rec['案件ID']);
    if (!proj) return { code: 'no_project' };
    var applied = pmApplyFinancialFields(parsed, proj, String(rec['申請者'] || '').trim() || approver);
    pmAddLog(rec['案件ID'], parsed.intent || 'billing_update', applied, rec['更新元メッセージ'], approver, rec['グループID'], 'approved');
    pmSetPendingStatus(pendingId, 'applied', approver);
    try { pmSyncCalendar(rec['案件ID']); } catch (e) { console.error('pmApprovePending calendar:', e.message); }
    return { code: 'ok_financial', projectName: proj['案件名'] || rec['案件ID'], applied: applied };
  });

  if (!locked.ok) return { code: 'lock_timeout' };
  return locked.result;
}

// LINE経路の承認（従来どおり）。中核処理は pmApprovePendingCore に委譲し、結果をLINE返信に変換する。
function pmApprovePending(pendingId, approver, replyToken, userId) {
  var r = pmApprovePendingCore(pendingId, approver, function() { return pmIsApprover(userId); });
  switch (r.code) {
    case 'lock_timeout': sendLineReply(replyToken, '混雑のため処理できませんでした。少し時間をおいて再度お試しください。'); return;
    case 'notfound':    sendLineReply(replyToken, '確認データが見つかりませんでした。'); return;
    case 'done':        sendLineReply(replyToken, 'この確認は既に処理済みです（' + r.status + '）。'); return;
    case 'forbidden':   sendLineReply(replyToken, '⛔ 金額・請求・入金の承認は経理／管理者のみ可能です。'); return;
    case 'parse_error': sendLineReply(replyToken, '抽出内容の解析に失敗しました。手動で確認してください。'); return;
    case 'no_project':  sendLineReply(replyToken, '対象案件が見つかりませんでした。'); return;
    case 'apply_failed':sendLineReply(replyToken, '反映に失敗しました。'); return;
    case 'ok_project':
    case 'ok_financial':
      sendLineReply(replyToken, '✅ 承認しました。案件「' + r.projectName + '」に確定反映しました。\n' +
        Object.keys(r.applied).map(function(k) { return '・' + k + '：' + r.applied[k]; }).join('\n'));
      return;
  }
}

// 却下の中核（UI非依存）。戻り値 code: lock_timeout | notfound | done | forbidden | ok
function pmRejectPendingCore(pendingId, approver, canApproveFinancialFn) {
  var locked = pmWithLock(function() {
    var rec = pmGetPending(pendingId);
    if (!rec) return { code: 'notfound' };
    // 処理済み（applied/rejected）は再処理不可
    if (rec['status'] !== 'awaiting_approval' && rec['status'] !== 'awaiting_project') return { code: 'done', status: rec['status'] };
    if (pmIsFinancialType(rec['type']) && !canApproveFinancialFn()) return { code: 'forbidden' };
    pmSetPendingStatus(pendingId, 'rejected', approver);
    return { code: 'ok' };
  });
  if (!locked.ok) return { code: 'lock_timeout' };
  return locked.result;
}

// LINE経路の却下（従来どおり）。
function pmRejectPending(pendingId, approver, replyToken, userId) {
  var r = pmRejectPendingCore(pendingId, approver, function() { return pmIsApprover(userId); });
  switch (r.code) {
    case 'lock_timeout': sendLineReply(replyToken, '混雑のため処理できませんでした。少し時間をおいて再度お試しください。'); return;
    case 'notfound':  sendLineReply(replyToken, '確認データが見つかりませんでした。'); return;
    case 'done':      sendLineReply(replyToken, 'この確認は既に処理済みです（' + r.status + '）。'); return;
    case 'forbidden': sendLineReply(replyToken, '⛔ 金額・請求・入金の却下は経理／管理者のみ可能です。'); return;
    case 'ok':        sendLineReply(replyToken, '❌ 却下しました。projectsへの反映は行いません。'); return;
  }
}
