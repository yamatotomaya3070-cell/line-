// ==========================================
// Router：LINE報告の入口判定・確認応答・postback振り分け（全体オーケストレーション）
//   コード.js への接続点はこの2関数のみ：
//     pmRouteText(ev, text, sender, userId, groupId) → 処理したら true
//     pmRoutePostback(ev, userId, groupId)           → 処理したら true
//   依存: pm_core / pm_parser / pm_manager / pm_pending / pm_taskbridge /
//         pm_calendar / pm_drive, 既存 sendLineReply / sendQuickReply / parseParams
// ==========================================

// ---- テキスト入口 ----
function pmRouteText(ev, text, sender, userId, groupId) {
  if (!text) return false;
  var intent = pmClassifyIntent(text);
  if (!intent) return false; // PMプレフィックス非該当 → 既存処理に委譲

  var body = pmStripPrefix(text);

  // 店舗登録は「案件解決＋店舗名」だけで完結（Gemini抽出は不要）
  if (intent === 'store_add') {
    pmHandleStoreAdd(ev.replyToken, body, sender, userId, groupId, text);
    return true;
  }

  if (!body) {
    sendLineReply(ev.replyToken, '報告内容を続けてご記入ください。\n例：「案件更新 〇〇ホテル、見積提出しました。6/25にフォローします。」');
    return true;
  }

  var parsed = pmParseReport(body, intent);

  if (intent === 'project_create') {
    pmHandleCreate(ev.replyToken, parsed, sender, userId, groupId, body);
    return true;
  }
  pmHandleReport(ev.replyToken, parsed, sender, userId, groupId, body);
  return true;
}

// ---- 新規案件登録 ----
function pmHandleCreate(replyToken, parsed, sender, userId, groupId, srcMsg) {
  var name = parsed.project_name;
  if (!name) {
    sendLineReply(replyToken, '案件名を読み取れませんでした。\n例：「新規案件 案件名:〇〇ホテル客室改装 クライアント:ABC運営 担当:濱田」');
    return;
  }
  var existing = pmGetProjectByName(name);
  if (existing) {
    // 既存案件 → 更新として扱う
    pmHandleReport(replyToken, parsed, sender, userId, groupId, srcMsg, name);
    return;
  }

  var rec = pmEnsureProjectRecord(name, {
    client: parsed.client_name || '',
    assignee: parsed.assignee || sender || '',
    groupId: groupId || '',
    phase: parsed.phase || '営業',
    updatedBy: sender || '',
  });
  pmAddLog(rec['案件ID'], 'project_create', { '案件名': name, 'クライアント名': parsed.client_name || '', '担当者': parsed.assignee || '' }, srcMsg, sender, groupId, 'auto');

  // Driveフォルダ作成（冪等・新階層 親/案件名）。戻り値はオブジェクトなので .url を取り出す。
  var driveRes = pmEnsureProjectFolder(rec['案件ID'], name, { source: 'line', actor: sender || '' });
  var driveUrl = (driveRes && driveRes.ok) ? driveRes.url : '';

  // 金額（売上予定など）→ 確認待ち
  var pendingNote = '';
  if (parsed.amount && parsed.amount.value != null) {
    var pid = pmQueuePending('amount', parsed, rec['案件ID'], name, sender, groupId, srcMsg);
    pmNotifyAccountingPending(pid, parsed, name, srcMsg);
    pendingNote = '\n※金額（' + (parsed.amount.raw || parsed.amount.value) + '）は確認待ちに送りました。';
  }

  sendLineReply(replyToken,
    '🆕 新規案件を登録しました。\n案件ID：' + rec['案件ID'] + '\n案件名：' + name +
    (parsed.client_name ? '\nクライアント：' + parsed.client_name : '') +
    (parsed.assignee ? '\n担当：' + parsed.assignee : '') +
    (driveUrl ? '\nDrive：' + driveUrl : '') +
    pendingNote +
    '\n\n📍 続けて店舗を登録：「店舗追加 〇〇店」と送ってください（店舗フォルダを作成します）。');
}

// ---- 店舗登録（LINE）----
//  「店舗追加 <店舗名>」→ 案件を解決し、店舗フォルダ(親/案件名/店舗名/6_データ)生成＋台帳登録。
//  案件はグループ紐付けから解決。生成済みの店舗は以降 選択肢（pm_store_pick）として提示できる。
function pmHandleStoreAdd(replyToken, storeName, sender, userId, groupId, srcMsg) {
  storeName = String(storeName || '').trim();
  if (!storeName) {
    sendLineReply(replyToken, '店舗名を入れてください。\n例：「店舗追加 藤沢店」');
    return;
  }

  // 案件を解決（グループ紐付け優先）。曖昧/未特定なら案内して終了。
  var resolved = pmResolveProject(srcMsg, groupId, '');
  if (resolved.status !== 'resolved') {
    sendLineReply(replyToken,
      'この店舗をひもづける案件を特定できませんでした。\n' +
      '先に「新規案件 案件名:〇〇」で案件を登録するか、案件が紐づくグループから送ってください。');
    return;
  }

  var projectName = resolved.name;
  var proj = pmGetProjectByName(projectName);
  var projectId = proj ? String(proj['案件ID'] || '') : '';

  var sr = pmEnsureStoreRecord(projectId, storeName, {
    caseName: projectName, groupId: groupId, source: 'line', actor: sender || '',
  });
  if (!sr.ok) { sendLineReply(replyToken, '店舗登録に失敗しました：' + (sr.error || '')); return; }

  try { pmAddLog(projectId, 'store_create', { '案件名': projectName, '店舗名': storeName }, srcMsg, sender, groupId, 'auto'); } catch (e) {}

  var dataUrl = (sr.folders && (sr.folders.dataFolderUrl || sr.folders.storeFolderUrl)) || '';
  var stores  = projectId ? pmListStoresByProject(projectId) : [];
  var names   = stores.map(function (s) { return String(s['店舗名']); }).filter(function (n) { return !!n; });
  sendLineReply(replyToken,
    '🏪 店舗を登録しました。\n案件：' + projectName + '\n店舗：' + storeName +
    (dataUrl ? '\n保存先：' + dataUrl : '') +
    (names.length > 1 ? '\n\nこの案件の店舗：' + names.join(' / ') : ''));
}

// ---- 進捗/請求/入金 報告 ----
//  forcedName: 既に案件確定済みの場合に指定（postback解決後など）
function pmHandleReport(replyToken, parsed, sender, userId, groupId, srcMsg, forcedName) {
  var projectName = forcedName;
  if (!projectName) {
    var resolved = pmResolveProject(srcMsg, groupId, parsed.project_name);
    if (resolved.status === 'resolved') {
      projectName = resolved.name;
    } else if (resolved.status === 'ambiguous') {
      pmAskProjectChoice(replyToken, parsed, sender, groupId, srcMsg, resolved.candidates);
      return;
    } else {
      // 候補なし → 新規作成を促す
      sendLineReply(replyToken,
        '該当する案件が見つかりませんでした。\n新規登録する場合は「新規案件 案件名:〇〇」とお送りください。');
      return;
    }
  }
  pmApplyResolved(replyToken, parsed, projectName, sender, userId, groupId, srcMsg);
}

// 案件曖昧時：候補をQuick Replyで提示し確認待ちへ
function pmAskProjectChoice(replyToken, parsed, sender, groupId, srcMsg, candidates) {
  var pid = pmQueuePending('resolve', parsed, '', '', sender, groupId, srcMsg, candidates);
  var items = [];
  candidates.slice(0, 3).forEach(function(c) {
    items.push({ type: 'action', action: { type: 'postback',
      label: String(c.name).slice(0, 20),
      data: 'action=pm_pick&id=' + pid + '&p=' + encodeURIComponent(c.name) } });
  });
  items.push({ type: 'action', action: { type: 'postback', label: '🆕 新規案件として登録', data: 'action=pm_pick&id=' + pid + '&p=__new__' } });
  sendQuickReply(replyToken,
    '該当する案件が複数あります。更新する案件をお選びください。', items);
}

// 進捗更新を自動反映してよいか。NGなら理由文字列、OKならnull（確定仕様：低信頼度・マスタ不一致は確認待ち）
function pmProgressHoldReason(parsed) {
  if (typeof parsed.confidence === 'number' && parsed.confidence < 70) return 'AIの信頼度が基準未満';
  if (parsed.status) {
    var norm = pmNormalizeStatus(parsed.phase, parsed.status);
    if (norm.status && !norm.matched) return 'ステータスが許可マスタに一致しない';
  }
  return null;
}

// 確認待ち提示用：parsed の主要項目を短く要約
function pmSummarizeParsed(parsed) {
  var lines = [];
  if (parsed.phase)  lines.push('・フェーズ：' + parsed.phase);
  if (parsed.status) lines.push('・ステータス：' + parsed.status);
  if (parsed.next_action) lines.push('・次回アクション：' + parsed.next_action + (parsed.next_action_due_date ? '（期限 ' + parsed.next_action_due_date + '）' : ''));
  if (parsed.assignee) lines.push('・担当：' + parsed.assignee);
  if (parsed.remark)   lines.push('・備考：' + parsed.remark);
  return lines.length ? lines.join('\n') : '（抽出項目なし）';
}

// 予定追加（calendar_update）：時刻を保持してカレンダー登録。projectsの進捗フェーズ等は変更しない。
function pmHandleCalendarAdd(replyToken, parsed, projectName, sender, groupId, srcMsg) {
  var date = parsed.event_date || parsed.meeting_date || parsed.next_action_due_date;
  if (!date) {
    sendLineReply(replyToken, '予定の日付を読み取れませんでした。\n例：「予定追加 〇〇案件 7/3 14時 現地打合せ」');
    return;
  }
  var title = parsed.event_title || parsed.next_action || '予定';
  var startTime = parsed.event_time || '';
  var schedule = {
    project_name: projectName,
    title:        title,
    date:         date,
    startTime:    startTime,
    endTime:      parsed.event_end_time || '',
    location:     parsed.location || '',
    attendees:    parsed.assignee || sender || '',
    description:  parsed.remark || '',
    group_id:     groupId || '',
    datetime:     new Date(),
  };

  // 同時実行・二重作成防止：シート重複チェック＋書込み＋カレンダー登録をロックで直列化
  var res = pmWithLock(function() {
    if (isDuplicateSchedule(title, date, groupId)) return 'dup';
    writeSchedule(schedule);   // スケジュール管理シート（registerNewProjectは既存案件ならスキップ）
    addToCalendar(schedule);   // 既存関数：開始時刻を保持し、時刻付きは重複チェックする
    return 'ok';
  });
  if (!res.ok) { sendLineReply(replyToken, '混雑のため登録できませんでした。少し時間をおいて再度お送りください。'); return; }
  if (res.result === 'dup') {
    sendLineReply(replyToken, '同じ予定が既に登録済みのため、追加しませんでした：' + title + '（' + date + '）');
    return;
  }
  sendLineReply(replyToken,
    '📅 予定を登録しました。\n' + title + '：' + date + (startTime ? ' ' + startTime : '') + '\n案件：' + projectName);
}

// 案件確定後：自動更新の適用＋カレンダー＋タスク連携＋（金額は確認待ち）
function pmApplyResolved(replyToken, parsed, projectName, sender, userId, groupId, srcMsg) {
  var intent = parsed.intent;

  // 予定追加：時刻を保持してカレンダー登録（進捗は変更しない）
  if (intent === 'calendar_update') {
    pmHandleCalendarAdd(replyToken, parsed, projectName, sender, groupId, srcMsg);
    return;
  }

  // 請求・入金は自動確定しない → まるごと確認待ち
  if (intent === 'billing_update' || intent === 'payment_update') {
    var rec = pmEnsureProjectRecord(projectName, { groupId: groupId, updatedBy: sender });
    var pendType = intent === 'billing_update' ? 'billing' : 'payment';
    var pid = pmQueuePending(pendType, parsed, rec['案件ID'], projectName, sender, groupId, srcMsg);
    pmNotifyAccountingPending(pid, parsed, projectName, srcMsg);
    sendLineReply(replyToken,
      '🧾 ' + projectName + ' の' + (intent === 'billing_update' ? '請求' : '入金') + '情報を確認待ちに登録しました。\n経理の承認後に反映されます。');
    return;
  }

  // 進捗更新：低信頼度・ステータスマスタ不一致は自動反映せず、報告者へ確認を依頼（確定仕様§2）
  var holdReason = pmProgressHoldReason(parsed);
  if (holdReason) {
    // 案件IDは承認時に確定（ここでprojects/フォルダを先行作成しない）
    var hpid = pmQueuePending('project_field', parsed, '', projectName, sender, groupId, srcMsg);
    sendQuickReply(replyToken,
      '⚠️ ' + projectName + ' の進捗更新は確認が必要です（' + holdReason + '）。\n' +
      pmSummarizeParsed(parsed) + '\n\nこの内容で反映してよろしいですか？',
      [
        { type: 'action', action: { type: 'postback', label: '✅ 反映する', data: 'action=pm_approve&id=' + hpid } },
        { type: 'action', action: { type: 'postback', label: '❌ 取消',     data: 'action=pm_reject&id=' + hpid } },
      ]);
    return;
  }

  // 進捗更新（安全項目を自動反映）
  var result = pmApplyUpdate(parsed, projectName, sender, groupId, srcMsg);
  if (!result) { sendLineReply(replyToken, '更新に失敗しました。'); return; }
  var projectId = result.projectId;

  // カレンダー同期
  try { pmSyncCalendar(projectId); } catch (e) { console.error('pmApplyResolved calendar:', e.message); }

  // 既存タスク管理へフォロータスク作成
  var taskCreated = false;
  try { taskCreated = pmCreateFollowupTask(projectId, projectName, parsed, sender, groupId); } catch (e) { console.error('pmApplyResolved taskbridge:', e.message); }

  // 金額が混在 → 確認待ち
  var amountNote = '';
  if (parsed.amount && parsed.amount.value != null) {
    var apid = pmQueuePending('amount', parsed, projectId, projectName, sender, groupId, srcMsg);
    pmNotifyAccountingPending(apid, parsed, projectName, srcMsg);
    amountNote = '\n※金額（' + (parsed.amount.raw || parsed.amount.value) + '）は確認待ちに送りました。';
  }

  var msg = pmFormatApplied(projectName, result.applied);
  if (taskCreated) msg += '\n📋 フォロータスクを作成しました：' + projectName + 'の' + parsed.next_action + '（期限 ' + parsed.next_action_due_date + '）';
  sendLineReply(replyToken, msg + amountNote);
}

// ---- postback 入口 ----
function pmRoutePostback(ev, userId, groupId) {
  var data   = (ev.postback && ev.postback.data) ? ev.postback.data : '';
  var params = parseParams(data);
  var action = params.action;
  if (!action || action.indexOf('pm_') !== 0) return false;

  var approver = pmMemberNameByUserId(userId) || userId;

  if (action === 'pm_approve') { pmApprovePending(params.id, approver, ev.replyToken, userId); return true; }
  if (action === 'pm_reject')  { pmRejectPending(params.id, approver, ev.replyToken, userId); return true; }

  if (action === 'pm_pick') {
    var pickedRaw = params.p ? decodeURIComponent(params.p) : '';
    // 二重押下対策：状態確認→applied化をロックで直列化（先に1回だけ確定）
    var claim = pmWithLock(function() {
      var rec = pmGetPending(params.id);
      if (!rec) return { code: 'notfound' };
      if (rec['status'] !== 'awaiting_project') return { code: 'done' };
      pmSetPendingStatus(params.id, 'applied', approver);
      return { code: 'ok', rec: rec };
    });
    if (!claim.ok) { sendLineReply(ev.replyToken, '混雑のため処理できませんでした。少し時間をおいて再度お試しください。'); return true; }
    if (claim.result.code === 'notfound') { sendLineReply(ev.replyToken, '確認データが見つかりませんでした。'); return true; }
    if (claim.result.code === 'done') { sendLineReply(ev.replyToken, 'この案件選択は既に処理済みです。'); return true; }

    var rec = claim.result.rec;
    var parsed = safeParseJson(rec['抽出内容']);

    if (pickedRaw === '__new__') {
      var newName = parsed.project_name || '';
      if (!newName) { sendLineReply(ev.replyToken, '新規案件名を読み取れませんでした。「新規案件 案件名:〇〇」とお送りください。'); return true; }
      pmHandleCreate(ev.replyToken, parsed, rec['申請者'], userId, rec['グループID'], rec['更新元メッセージ']);
      return true;
    }
    pmApplyResolved(ev.replyToken, parsed, pickedRaw, rec['申請者'], userId, rec['グループID'], rec['更新元メッセージ']);
    return true;
  }

  // 既存店舗を選択肢から選ぶ：pid（案件ID）+ s（店舗名）→ 店舗フォルダを保証（冪等）
  if (action === 'pm_store_pick') {
    var pid   = params.pid ? decodeURIComponent(params.pid) : '';
    var sname = params.s   ? decodeURIComponent(params.s)   : '';
    if (!sname) { sendLineReply(ev.replyToken, '店舗を特定できませんでした。'); return true; }
    var sproj = pid ? pmGetProjectById(pid) : null;
    var scase = sproj ? sproj['案件名'] : '';
    var ssr = pmEnsureStoreRecord(pid, sname, { caseName: scase, groupId: groupId, source: 'line-pick', actor: approver });
    if (!ssr.ok) { sendLineReply(ev.replyToken, '店舗の準備に失敗しました：' + (ssr.error || '')); return true; }
    var sUrl = (ssr.folders && (ssr.folders.dataFolderUrl || ssr.folders.storeFolderUrl)) || '';
    sendLineReply(ev.replyToken, '🏪 店舗「' + sname + '」を選択しました。' + (sUrl ? '\n保存先：' + sUrl : ''));
    return true;
  }

  return false;
}

// 案件の既存店舗を Quick Reply の選択肢にして提示する共通導線。
//   他フロー（写真受信時の店舗特定など）からも再利用できる。stores が無ければ false。
function pmPromptStoreChoice(replyToken, projectId, promptText) {
  if (!projectId) return false;
  var stores = pmListStoresByProject(projectId);
  if (!stores.length) return false;
  var items = stores.slice(0, 12).map(function (s) {
    return { type: 'action', action: { type: 'postback',
      label: String(s['店舗名'] || '店舗').slice(0, 20),
      data: 'action=pm_store_pick&pid=' + encodeURIComponent(projectId) + '&s=' + encodeURIComponent(s['店舗名']) } };
  });
  sendQuickReply(replyToken, promptText || '店舗を選んでください（新規は「店舗追加 店舗名」）。', items);
  return true;
}

// LINEユーザーID → メンバー名（承認者表示用）
function pmMemberNameByUserId(userId) {
  if (!userId) return '';
  var sheet = getSheet('メンバー管理');
  if (!sheet || sheet.getLastRow() <= 1) return '';
  var idx = pmHeaderIndex(sheet);
  var cName = idx['名前'];
  var cId   = idx['LINE ユーザーID'];
  if (cName === undefined || cId === undefined) return '';
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][cId] || '') === userId) return String(data[i][cName] || '');
  }
  return '';
}
