// ==========================================
// ProjectTaskBridge：次回アクション → 既存タスク管理へ自動作成（重複防止）
//   依存: pm_core.js, 既存 getSheet / fmtDT / generateId / getProjectNameByGroupId
//   注: カレンダー登録は pm_calendar.js (pmSyncCalendar) が担当するため
//       ここでは addTaskToCalendar は呼ばない（二重登録防止）。
// ==========================================

function pmCreateFollowupTask(projectId, projectName, parsed, sender, groupId) {
  if (!parsed.next_action || !parsed.next_action_due_date) return false;

  var sheet = getSheet('タスク管理');
  if (!sheet) return false;

  var content  = projectName + 'の' + parsed.next_action;
  var assignee = parsed.assignee || sender || '';
  var due      = parsed.next_action_due_date;

  // follow-up task support: preserve assignee user IDs when present, but do not add columns automatically.
  // TODO: さらなる multi-user / assignee ID フロー対応は将来の検討。
  var assigneeIds = Array.isArray(parsed.assignee_user_ids) ? parsed.assignee_user_ids : [];

  // 重複検出と追記を同時実行で直列化（並走でも二重作成しない）
  var res = pmWithLock(function() {
    var idx  = pmHeaderIndex(sheet);
    var data = sheet.getDataRange().getValues();
    var cRelId   = idx['関連案件ID'];
    var cContent = idx['タスク内容'];
    var cDue     = idx['期日'];
    var cStatus  = idx['ステータス'];
    // 列が揃っていない（pmSetupSheets未実行など）と重複検出が無効化され
    // タスクが毎回増殖するため、安全側で作成を中止して列追加を促す
    if (cRelId === undefined || cContent === undefined || cDue === undefined || cStatus === undefined) {
      console.warn('pmCreateFollowupTask: 必要列が不足（pmEnsureTaskColumns未実行?）。作成を中止');
      return false;
    }
    // 重複防止：同一(関連案件ID × タスク内容 × 期日)で未完了が存在 → 作成しない
    for (var i = 1; i < data.length; i++) {
      var st = String(data[i][cStatus] || '');
      if (st === 'done' || st === '完了') continue;
      var rowDue = data[i][cDue] instanceof Date ? fmtDate(data[i][cDue]) : String(data[i][cDue] || '').slice(0, 10);
      if (String(data[i][cRelId] || '') === String(projectId) &&
          String(data[i][cContent] || '') === content && rowDue === due) {
        return false; // 重複
      }
    }
    pmAppendRowFields('タスク管理', {
      'タスクID': generateId(),
      '案件名': projectName,
      'タスク内容': content,
      '担当者': assignee,
      '担当者ID': assigneeIds.length ? assigneeIds.join('、') : '',
      '期日': due,
      'ステータス': 'confirmed',
      '作成日時': fmtDT(new Date()),
      'グループID': groupId || '',
      '緊急度': '',
      '関連案件ID': projectId || '',
      '作成元': parsed.intent || 'project_update',
    });
    return true;
  });
  return res.ok ? res.result : false;
}
