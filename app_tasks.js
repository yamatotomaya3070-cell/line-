// ==========================================
// タスク画面（GASウェブアプリ）— チェックボックスで選んで一括完了
//   配信: doGet(?page=tasks) → tasks.html（ルーティングは app_dashboard.js）
//   LINEのタップ完了（task_done postback）と同じタスク管理シートを更新するため、
//   どちらの経路で完了しても整合が保たれる。
// ==========================================

// 未完了タスク一覧（期日超過 → 期日近い順 → 期日未定）。担当者フィルタ候補も返す
function appGetTasks() {
  var open  = collectOpenTasks(null);
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var list  = open.map(function(t) {
    var d = '';
    if (t.deadline instanceof Date) d = Utilities.formatDate(t.deadline, 'Asia/Tokyo', 'yyyy-MM-dd');
    else if (t.deadline) d = String(t.deadline).slice(0, 10);
    return {
      id: t.taskId, project: t.project, content: t.content, assignee: t.assignee,
      deadline: d, overdue: !!(d && d < today),
    };
  });
  list.sort(function(a, b) {
    if (a.overdue !== b.overdue) return a.overdue ? -1 : 1;
    if (!a.deadline && !b.deadline) return 0;
    if (!a.deadline) return 1;
    if (!b.deadline) return -1;
    return a.deadline < b.deadline ? -1 : (a.deadline > b.deadline ? 1 : 0);
  });
  var assignees = [];
  list.forEach(function(t) {
    String(t.assignee || '').split(/[、,，\/／・\s]+/).forEach(function(name) {
      name = name.trim();
      if (name && assignees.indexOf(name) === -1) assignees.push(name);
    });
  });
  return { tasks: list, assignees: assignees };
}

// 選択したタスクを一括で完了にする
function appCompleteTasks(taskIds) {
  if (!taskIds || !taskIds.length) return { ok: false, msg: 'タスクが選択されていません', count: 0 };
  var sheet = getSheet('タスク管理');
  if (!sheet || sheet.getLastRow() <= 1) return { ok: false, msg: 'タスクがありません', count: 0 };
  var wanted = {};
  taskIds.forEach(function(id) { wanted[String(id)] = true; });
  var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  var count = 0;
  for (var i = 0; i < data.length; i++) {
    if (!wanted[String(data[i][0])]) continue;
    var st = String(data[i][5] || '');
    if (st === 'done' || st === '完了') continue;
    sheet.getRange(i + 2, 6).setValue('done');
    count++;
  }
  return { ok: count > 0, msg: count > 0 ? '' : '対象が見つかりませんでした（完了済みの可能性）', count: count };
}

// 一括完了の取り消し（完了直後の「元に戻す」用）
function appReopenTasks(taskIds) {
  if (!taskIds || !taskIds.length) return { ok: false, count: 0 };
  var sheet = getSheet('タスク管理');
  if (!sheet || sheet.getLastRow() <= 1) return { ok: false, count: 0 };
  var wanted = {};
  taskIds.forEach(function(id) { wanted[String(id)] = true; });
  var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  var count = 0;
  for (var i = 0; i < data.length; i++) {
    if (!wanted[String(data[i][0])]) continue;
    sheet.getRange(i + 2, 6).setValue('confirmed');
    count++;
  }
  return { ok: count > 0, count: count };
}
