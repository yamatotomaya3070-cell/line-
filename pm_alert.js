// ==========================================
// ProjectAlert：放置・請求漏れ・入金未確認 の検知と通知（毎朝トリガー）
//   重複送信は project_alerts の「アラートキー」で防止（1案件・1種別・1日1回）。
//   依存: pm_core.js, pm_manager.js, 既存 getSheet / sendLineMessage / fmtDT
// ==========================================

// 案件が完了/失注フェーズか（放置アラート対象外）
function pmIsProjectClosed(proj) {
  var handover = String(proj['お引き渡しステータス'] || '');
  var sales    = String(proj['営業ステータス'] || '');
  if (sales.indexOf('失注') !== -1) return true;
  if (['引き渡し済', '最終請求済', '入金確認', 'アフター／保証期間'].indexOf(handover) !== -1) return true;
  return false;
}

// メンバー名 → LINE ユーザーID（DM通知用）
function pmMemberUserId(name) {
  if (!name) return '';
  var sheet = getSheet('メンバー管理');
  if (!sheet || sheet.getLastRow() <= 1) return '';
  var idx  = pmHeaderIndex(sheet);
  var cName = idx['名前'];
  var cId   = idx['LINE ユーザーID'];
  if (cName === undefined || cId === undefined) return '';
  var clean = String(name).replace('さん', '').trim();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var n = String(data[i][cName] || '').trim();
    if (n && (n === clean || clean.indexOf(n) !== -1 || n.indexOf(clean) !== -1)) {
      var uid = String(data[i][cId] || '');
      if (uid && uid.indexOf('U') === 0) return uid;
    }
  }
  return '';
}

// 既存アラートキーの集合を一括取得（毎朝チェックの先頭で1回構築）
function pmLoadAlertKeys() {
  var set = {};
  var rows = pmReadObjects(PM_SHEET_ALERTS);
  for (var i = 0; i < rows.length; i++) set[String(rows[i]['アラートキー'])] = true;
  return set;
}

// 案件ID → 直近の放置アラート検知日（Date）。再通知間隔の判定に使う。
function pmLoadLastStaleDates() {
  var map = {};
  var rows = pmReadObjects(PM_SHEET_ALERTS);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['種別'] || '') !== 'stale') continue;
    var pid = String(rows[i]['案件ID'] || '');
    if (!pid) continue;
    var d = pmParseDate(rows[i]['検知日時']);
    if (!d) continue;
    if (!map[pid] || d.getTime() > map[pid].getTime()) map[pid] = d;
  }
  return map;
}

// 放置アラートを今日通知してよいか。初回（履歴なし）はOK、以降は前回通知から
// PM_STALE_REMIND_INTERVAL_DAYS 日以上あいていればOK（毎日流さないための間引き）。
function pmShouldRemindStale(projectId, today, lastStaleDates) {
  var last = lastStaleDates[String(projectId)];
  if (!last) return true;
  var since = pmDaysBetween(last, today);
  return since === null || since >= PM_STALE_REMIND_INTERVAL_DAYS;
}

// メンション付きアラート（グループ宛て1通）。text内の {who} が @メンションになる。
// メンバー管理にLINE IDが無い人は名前の文字列にフォールバック。重複防止・履歴はpmSendAlertと同様。
function pmSendAlertMention(projectId, type, key, body, groupId, mentionUserId, fallbackName, targetLabel, knownKeys) {
  if (knownKeys ? knownKeys[key] : pmLoadAlertKeys()[key]) return false;
  if (knownKeys) knownKeys[key] = true;
  var sent = false;
  if (groupId) { sendLineMentionMessage(groupId, body, mentionUserId, fallbackName); sent = true; }
  pmAppendRowFields(PM_SHEET_ALERTS, {
    'アラートID': generateId(),
    '案件ID': projectId,
    '種別': type,
    'アラートキー': key,
    '検知日時': fmtDT(new Date()),
    '通知先': targetLabel || '',
    '内容': body.split('{who}').join(fallbackName || '担当者'),
    '状態': sent ? 'sent' : 'no_target',
  });
  return sent;
}

// アラート送信（重複防止＋履歴記録）。knownKeys を渡すと走査を省略。
function pmSendAlert(projectId, type, key, body, targets, targetLabel, knownKeys) {
  if (knownKeys ? knownKeys[key] : pmLoadAlertKeys()[key]) return false;
  if (knownKeys) knownKeys[key] = true;
  var sent = false;
  targets.forEach(function(t) {
    if (t) { sendLineMessage(t, body); sent = true; }
  });
  pmAppendRowFields(PM_SHEET_ALERTS, {
    'アラートID': generateId(),
    '案件ID': projectId,
    '種別': type,
    'アラートキー': key,
    '検知日時': fmtDT(new Date()),
    '通知先': targetLabel || '',
    '内容': body,
    '状態': sent ? 'sent' : 'no_target',
  });
  return sent;
}

// ---- 毎朝トリガー本体 ----
function pmCheckAlerts() {
  var sheet = getSheet(PM_SHEET_PROJECTS);
  if (!sheet) { console.log('pmCheckAlerts: projectsシートなし。pmSetupSheets未実行?'); return; }
  var rows = pmReadObjects(PM_SHEET_PROJECTS);
  var today = pmParseDate(pmTodayYmd());
  var todayYmd = pmTodayYmd();
  var internal = pmInternalTarget();
  var accounting = pmAccountingTarget();
  var progress = pmProgressTarget();
  var knownKeys = pmLoadAlertKeys();
  var lastStaleDates = pmLoadLastStaleDates();

  var count = 0;
  rows.forEach(function(proj) {
    var projectId = proj['案件ID'];
    if (!projectId) return;
    var name = proj['案件名'] || projectId;
    var assignee = proj['担当者'] || '';
    var assigneeId = pmMemberUserId(assignee);

    // ---- 放置アラート ----
    if (!pmIsProjectClosed(proj)) {
      var staleReason = null;
      var lastUpd = pmParseDate(proj['最終更新日']);
      var sales = String(proj['営業ステータス'] || '');
      var daysNoUpdate = lastUpd ? pmDaysBetween(lastUpd, today) : null;

      if (sales.indexOf('見積提出') !== -1 && daysNoUpdate !== null && daysNoUpdate >= PM_STALE_AFTER_QUOTE_DAYS) {
        staleReason = '見積提出後' + daysNoUpdate + '日間更新がありません';
      } else if (daysNoUpdate !== null && daysNoUpdate >= PM_STALE_NO_UPDATE_DAYS) {
        staleReason = '最終更新から' + daysNoUpdate + '日間更新がありません';
      } else {
        var due = pmParseDate(proj['次回アクション期限']);
        if (due && pmDaysBetween(due, today) > 0) {
          staleReason = '次回アクション期限（' + fmtDate(due) + '）を過ぎています';
        }
      }

      // 毎日流さない：前回の放置通知から PM_STALE_REMIND_INTERVAL_DAYS 日未満なら今日はスキップ
      if (staleReason && !pmShouldRemindStale(projectId, today, lastStaleDates)) staleReason = null;

      if (staleReason) {
        // 進捗管理グループへ担当者@メンション付きで1通（個人DMは送らない）
        var body = '【放置アラート】\n' +
          '{who}さん担当の「' + name + '」が、' + staleReason + '。\n' +
          '現在フェーズ：' + (proj['現在フェーズ'] || '') + (sales ? '（' + sales + '）' : '') + '\n' +
          '次回アクション：' + (proj['次回アクション'] || '未設定') +
          (proj['次回アクション期限'] ? '（期限 ' + fmtDate(pmParseDate(proj['次回アクション期限'])) + '）' : '');
        var key = projectId + ':stale:' + todayYmd;
        if (pmSendAlertMention(projectId, 'stale', key, body, progress, assigneeId, assignee || '担当未設定', '進捗グループ(メンション)', knownKeys)) count++;
      }
    }

    // ---- 請求漏れアラート ----
    var billDue = pmParseDate(proj['請求予定日']);
    var billStatus = String(proj['請求ステータス'] || '');
    if (billDue && pmDaysBetween(billDue, today) > 0 && pmIsBlank(proj['請求済日']) && billStatus !== '請求済') {
      var bbody = '【請求漏れアラート】\n' + name + 'の請求予定日（' + fmtDate(billDue) + '）を過ぎています。\n' +
        'ステータス：' + (billStatus || '未請求') + '\n担当：' + (assignee || '未設定');
      var bkey = projectId + ':billing_overdue:' + todayYmd;
      if (pmSendAlert(projectId, 'billing_overdue', bkey, bbody, [assigneeId, accounting], '担当者DM/経理', knownKeys)) count++;
    }

    // ---- 入金未確認アラート ----
    var payDue = pmParseDate(proj['入金予定日']);
    var payStatus = String(proj['入金ステータス'] || '');
    if (payDue && pmDaysBetween(payDue, today) > 0 && pmIsBlank(proj['入金確認日']) && payStatus !== '入金済') {
      var pbody = '【入金未確認アラート】\n' + name + 'の入金予定日（' + fmtDate(payDue) + '）を過ぎています。\n' +
        'ステータス：' + (payStatus || '未入金') + '\n担当：' + (assignee || '未設定');
      var pkey = projectId + ':payment_overdue:' + todayYmd;
      if (pmSendAlert(projectId, 'payment_overdue', pkey, pbody, [assigneeId, accounting], '担当者DM/経理', knownKeys)) count++;
    }
  });

  console.log('pmCheckAlerts: ' + count + '件のアラートを送信しました');
}
