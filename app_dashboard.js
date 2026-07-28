// ==========================================
// 進捗ボード（GASウェブアプリ）— 自社GAS配信の自作ダッシュボード
//   doGet でHTMLを配信。データは google.script.run で取得/更新（実行は所有者権限）。
//   アクセス保護: ?key=<DASHBOARD_TOKEN(Scriptプロパティ)> が一致しないと表示しない。
//   既存のLINE Webhook(doPost/匿名公開)はそのまま。doGet を足すだけ＝非破壊。
// ==========================================

function doGet(e) {
  try {
    var token = String(PropertiesService.getScriptProperties().getProperty('DASHBOARD_TOKEN') || '').trim();
    var key = (e && e.parameter) ? String(e.parameter.key || '').trim() : '';
    if (token && key !== token) {
      return HtmlService.createHtmlOutput(
        '<div style="font-family:sans-serif;padding:48px;text-align:center;color:#333">' +
        '<h2>🔒 アクセスキーが必要です</h2>' +
        '<p>URLの末尾に <code>?key=（合言葉）</code> を付けて開いてください。</p></div>'
      ).setTitle('進捗ボード');
    }
    // ?page=billing → 請求・入金 / ?page=tasks → タスク / それ以外 → 進捗ボード
    var page = (e && e.parameter) ? String(e.parameter.page || '').trim() : '';
    var pageFile = (page === 'billing') ? 'billing' : (page === 'tasks') ? 'tasks' : 'dashboard';
    var titles = { billing: 'WOODBASE 請求・入金', tasks: 'WOODBASE タスク', dashboard: 'WOODBASE 進捗ボード' };
    // ページ間リンク用URLをテンプレートに注入（キー必須運用でも遷移できるように）
    var baseUrl = ScriptApp.getService().getUrl();
    var keyQuery = token ? ('&key=' + encodeURIComponent(key)) : '';
    var t = HtmlService.createTemplateFromFile(pageFile);
    t.boardUrl = baseUrl + '?page=board' + keyQuery;
    t.billingUrl = baseUrl + '?page=billing' + keyQuery;
    t.tasksUrl = baseUrl + '?page=tasks' + keyQuery;
    // タスク画面の初期絞り込み（LINEから ?who=名前 で開いた場合）
    t.whoParam = (e && e.parameter) ? String(e.parameter.who || '').trim() : '';
    return t.evaluate()
      .setTitle(titles[pageFile])
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    // エラーは画面に出す（ログを見なくて済むように）
    return HtmlService.createHtmlOutput(
      '<div style="font-family:monospace;padding:24px;color:#b00020">' +
      '<h3>doGet エラー</h3><pre style="white-space:pre-wrap">' +
      String(err && err.stack ? err.stack : err) + '</pre></div>'
    );
  }
}

// 区分=社内/その他、取消=TRUE を除外した案件のみ返す
function appGetData() {
  var phases = Object.keys(PM_PHASES);
  var rows = pmReadObjects(PM_SHEET_PROJECTS);
  var list = [];
  rows.forEach(function(r) {
    var kind = String(r['区分'] || '').trim();
    if (kind === '社内' || kind === 'その他') return;
    var canceled = String(r['取消'] || '').trim().toUpperCase();
    if (canceled === 'TRUE' || canceled === '1' || canceled === '✓' || canceled === 'YES') return;
    var phase = String(r['現在フェーズ'] || '営業');
    list.push({
      id:         String(r['案件ID'] || ''),
      name:       String(r['案件名'] || ''),
      client:     String(r['クライアント名'] || ''),
      assignee:   String(r['担当者'] || ''),
      phase:      phase,
      // 現在フェーズに対応するサブステータス（営業ステータス等の列値）
      status:     String(r[PM_PHASE_COLUMN[phase] || ''] || ''),
      progress:   String(r['施工進捗'] || ''),
      nextAction: String(r['次回アクション'] || ''),
      nextDue:    String(r['次回アクション期限'] || ''),
      // 閲覧限定アプリのため日付・ステータスは表示可（金額そのものは出さない）
      meetingDate:    String(r['打合せ日'] || ''),
      startDue:       String(r['着工予定日'] || ''),
      deliveryDue:    String(r['引き渡し予定日'] || ''),
      billingStatus:  String(r['請求ステータス'] || ''),
      paymentStatus:  String(r['入金ステータス'] || ''),
      note:           String(r['備考'] || ''),
      driveUrl:       String(r['DriveフォルダURL'] || ''),
    });
  });
  // statusMaster: フェーズ→サブステータス候補（カードのドロップダウン用）
  return { phases: phases, statusMaster: PM_PHASES, projects: list };
}

// シート値の日付をyyyy-MM-ddに整形（文字列ならそのまま）
function appFmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(v == null ? '' : v).trim();
}

// 請求・入金ページ用データ。金額そのものは閲覧アプリには出さない方針を維持し、
// ステータスと日付のみ返す。対象案件の絞り込みは appGetData と同条件。
function appGetBillingData() {
  var phases = Object.keys(PM_PHASES);
  var rows = pmReadObjects(PM_SHEET_PROJECTS);
  var list = [];
  rows.forEach(function(r) {
    var kind = String(r['区分'] || '').trim();
    if (kind === '社内' || kind === 'その他') return;
    var canceled = String(r['取消'] || '').trim().toUpperCase();
    if (canceled === 'TRUE' || canceled === '1' || canceled === '✓' || canceled === 'YES') return;
    list.push({
      id:            String(r['案件ID'] || ''),
      name:          String(r['案件名'] || ''),
      client:        String(r['クライアント名'] || ''),
      phase:         String(r['現在フェーズ'] || '営業'),
      billingStatus: String(r['請求ステータス'] || ''),
      billingDue:    appFmtDate_(r['請求予定日']),
      billingDone:   appFmtDate_(r['請求済日']),
      paymentStatus: String(r['入金ステータス'] || ''),
      paymentDue:    appFmtDate_(r['入金予定日']),
      paymentDone:   appFmtDate_(r['入金確認日']),
    });
  });
  return { phases: phases, projects: list };
}

// アプリからの更新者名（ブラウザ側で入力済みの名前）。ログには「アプリ:名前」で残す。
function appWho_(who) { return String(who || '').trim().slice(0, 30); }
function appActor_(who) { var w = appWho_(who); return w ? 'アプリ:' + w : 'webapp'; }

// フェーズ変更（自動更新OK項目）。最終更新日・最終更新者を更新しログを残す。
function appSetPhase(projectId, phase, who) {
  if (!projectId || !phase) return { ok: false, msg: '引数不足' };
  var sh = getSheet(PM_SHEET_PROJECTS);
  var data = sh.getDataRange().getValues();
  var h = data[0].map(function(x) { return String(x).trim(); });
  var cId = h.indexOf('案件ID'), cPhase = h.indexOf('現在フェーズ'), cUpd = h.indexOf('最終更新日'), cWho = h.indexOf('最終更新者');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][cId] || '').trim() === String(projectId).trim()) {
      sh.getRange(i + 1, cPhase + 1).setValue(phase);
      if (cUpd !== -1) sh.getRange(i + 1, cUpd + 1).setValue(pmTodayYmd());
      if (cWho !== -1 && appWho_(who)) sh.getRange(i + 1, cWho + 1).setValue(appWho_(who));
      try { pmAddLog(projectId, 'phase_update', { '現在フェーズ': phase }, '', appActor_(who), '', 'auto'); } catch (e2) {}
      return { ok: true };
    }
  }
  return { ok: false, msg: '案件が見つかりません' };
}

// 担当者変更（自動更新OK項目）。最終更新日・最終更新者を更新しログを残す。
function appSetAssignee(projectId, assignee, who) {
  if (!projectId) return { ok: false, msg: '引数不足' };
  assignee = String(assignee || '').trim();
  var sh = getSheet(PM_SHEET_PROJECTS);
  var data = sh.getDataRange().getValues();
  var h = data[0].map(function(x) { return String(x).trim(); });
  var cId = h.indexOf('案件ID'), cAssignee = h.indexOf('担当者'), cUpd = h.indexOf('最終更新日'), cWho = h.indexOf('最終更新者');
  if (cAssignee === -1) return { ok: false, msg: '担当者列がありません（setup実行を）' };
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][cId] || '').trim() === String(projectId).trim()) {
      sh.getRange(i + 1, cAssignee + 1).setValue(assignee);
      if (cUpd !== -1) sh.getRange(i + 1, cUpd + 1).setValue(pmTodayYmd());
      if (cWho !== -1 && appWho_(who)) sh.getRange(i + 1, cWho + 1).setValue(appWho_(who));
      try { pmAddLog(projectId, 'assignee_update', { '担当者': assignee }, '', appActor_(who), '', 'auto'); } catch (e2) {}
      return { ok: true };
    }
  }
  return { ok: false, msg: '案件が見つかりません' };
}

// サブステータス変更（現在フェーズに対応するステータス列へ書き込み）。
// LINE経由のpmApplyUpdateと同じ列を使うため、両経路の整合が保たれる。
function appSetStatus(projectId, status, who) {
  if (!projectId) return { ok: false, msg: '引数不足' };
  status = String(status || '').trim();
  var sh = getSheet(PM_SHEET_PROJECTS);
  var data = sh.getDataRange().getValues();
  var h = data[0].map(function(x) { return String(x).trim(); });
  var cId = h.indexOf('案件ID'), cPhase = h.indexOf('現在フェーズ'), cUpd = h.indexOf('最終更新日'), cWho = h.indexOf('最終更新者');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][cId] || '').trim() !== String(projectId).trim()) continue;
    var phase = String(data[i][cPhase] || '営業');
    var col = PM_PHASE_COLUMN[phase];
    if (!col) return { ok: false, msg: 'フェーズが不明です: ' + phase };
    // マスタ外の値は誤更新防止のため拒否（空=未設定は許可）
    if (status && (PM_PHASES[phase] || []).indexOf(status) === -1) {
      return { ok: false, msg: '「' + phase + '」にないステータスです: ' + status };
    }
    var cStatus = h.indexOf(col);
    if (cStatus === -1) return { ok: false, msg: col + ' 列がありません（setup実行を）' };
    sh.getRange(i + 1, cStatus + 1).setValue(status);
    if (cUpd !== -1) sh.getRange(i + 1, cUpd + 1).setValue(pmTodayYmd());
    if (cWho !== -1 && appWho_(who)) sh.getRange(i + 1, cWho + 1).setValue(appWho_(who));
    var chg = {}; chg[col] = status;
    try { pmAddLog(projectId, 'status_update', chg, '', appActor_(who), '', 'auto'); } catch (e2) {}
    return { ok: true };
  }
  return { ok: false, msg: '案件が見つかりません' };
}

// 施工進捗％（0〜100）。ボードの施工カードから更新。
function appSetProgress(projectId, pct, who) {
  if (!projectId) return { ok: false, msg: '引数不足' };
  var n = Number(pct);
  if (isNaN(n) || n < 0 || n > 100) return { ok: false, msg: '進捗は0〜100で指定してください' };
  var sh = getSheet(PM_SHEET_PROJECTS);
  var data = sh.getDataRange().getValues();
  var h = data[0].map(function(x) { return String(x).trim(); });
  var cId = h.indexOf('案件ID'), cProg = h.indexOf('施工進捗'), cUpd = h.indexOf('最終更新日'), cWho = h.indexOf('最終更新者');
  if (cProg === -1) return { ok: false, msg: '施工進捗列がありません（setup実行を）' };
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][cId] || '').trim() === String(projectId).trim()) {
      sh.getRange(i + 1, cProg + 1).setValue(n);
      if (cUpd !== -1) sh.getRange(i + 1, cUpd + 1).setValue(pmTodayYmd());
      if (cWho !== -1 && appWho_(who)) sh.getRange(i + 1, cWho + 1).setValue(appWho_(who));
      try { pmAddLog(projectId, 'progress_update', { '施工進捗': n }, '', appActor_(who), '', 'auto'); } catch (e2) {}
      return { ok: true };
    }
  }
  return { ok: false, msg: '案件が見つかりません' };
}

// Driveフォルダを開く（新規作成はしない・既存フォルダの検索のみ）。
//   ①projects.DriveフォルダURL（pm_drive.jsの正規フォルダ）があればそれを使う。
//   ②無ければ、LINE添付保存(saveFileToDrive)・旧registerNewProjectが使う
//     「DRIVE_FOLDER_ID直下＝案件名そのままのフォルダ」を、表記ゆれ込みであいまい検索する
//     （pmFindExistingFlatFolder：案件名／プロジェクト管理の正式名称・略称も候補にする）。
//   ③どちらも無ければ「未作成」を返す（フォルダは作らない）。
function appGetDriveUrl(projectId) {
  if (!projectId) return { ok: false, msg: '引数不足' };
  var proj = pmGetProjectById(projectId);
  if (!proj) return { ok: false, msg: '案件が見つかりません' };
  try {
    if (!pmIsBlank(proj['DriveフォルダURL'])) return { ok: true, url: proj['DriveフォルダURL'] };

    var config = getConfig();
    if (!config.DRIVE_FOLDER_ID) return { ok: false, msg: 'DRIVE_FOLDER_ID未設定です' };
    var url = pmFindExistingFlatFolder(projectId, proj['案件名']);
    if (url) return { ok: true, url: url };

    return { ok: false, msg: 'Driveフォルダが見つかりませんでした（案件名の表記ゆれの可能性。フォルダ名をご確認ください）' };
  } catch (e) {
    return { ok: false, msg: 'エラー：' + e.message };
  }
}

// 取り消し（物理削除でなくアーカイブ）
function appArchiveProject(projectId, who) {
  if (!projectId) return { ok: false };
  var sh = getSheet(PM_SHEET_PROJECTS);
  var data = sh.getDataRange().getValues();
  var h = data[0].map(function(x) { return String(x).trim(); });
  var cId = h.indexOf('案件ID'), cDel = h.indexOf('取消');
  if (cDel === -1) return { ok: false, msg: '取消列がありません（setup実行を）' };
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][cId] || '').trim() === String(projectId).trim()) {
      sh.getRange(i + 1, cDel + 1).setValue('TRUE');
      try { pmAddLog(projectId, 'archive', { '取消': 'TRUE' }, '', appActor_(who), '', 'auto'); } catch (e2) {}
      return { ok: true };
    }
  }
  return { ok: false };
}

// 請求/入金入力（承認制は維持）：LINEの「請求更新」「入金更新」と同じ経路
//   （pmQueuePending→経理へ通知→承認で確定）に乗せる。ここではprojectsへ直接書き込まない。
//   金額・予定日・確定日のいずれか1つ以上が必須。
var APP_AMOUNT_KIND_CONF = {
  billing: { pendingType: 'billing', intent: 'billing_update', amountType: '請求',
             dueField: 'billing_due_date', doneField: 'billing_done_date', srcLabel: '請求入力' },
  payment: { pendingType: 'payment', intent: 'payment_update', amountType: '入金',
             dueField: 'payment_due_date', doneField: 'payment_confirmed_date', srcLabel: '入金入力' },
};

function appSubmitAmountRequest_(kind, projectId, payload) {
  var conf = APP_AMOUNT_KIND_CONF[kind];
  payload = payload || {};
  if (!projectId) return { ok: false, msg: '案件が不明です' };
  var proj = pmGetProjectById(projectId);
  if (!proj) return { ok: false, msg: '案件が見つかりません' };

  var requester = String(payload.requester || '').trim();
  if (!requester) return { ok: false, msg: '申請者名は必須です' };

  var amountRaw = payload.amount;
  var hasAmount = amountRaw !== '' && amountRaw !== null && amountRaw !== undefined;
  var amountVal = hasAmount ? Number(amountRaw) : null;
  if (hasAmount && isNaN(amountVal)) return { ok: false, msg: '金額は数値で入力してください' };

  var dueDate  = String(payload.dueDate || '').trim();
  var doneDate = String(payload.doneDate || '').trim();
  if (!hasAmount && !dueDate && !doneDate) {
    return { ok: false, msg: '金額・予定日・確定日のいずれかを入力してください' };
  }

  var parsed = { intent: conf.intent };
  if (hasAmount)  parsed.amount = { type: conf.amountType, value: amountVal, raw: String(amountRaw) };
  if (dueDate)    parsed[conf.dueField]  = dueDate;
  if (doneDate)   parsed[conf.doneField] = doneDate;

  var pid = pmQueuePending(conf.pendingType, parsed, projectId, proj['案件名'], requester, '', 'アプリ: ' + conf.srcLabel + '（' + requester + '）');
  try { pmNotifyAccountingPending(pid, parsed, proj['案件名'], 'アプリからの' + conf.srcLabel + '（' + requester + '）'); }
  catch (e) { console.error('appSubmitAmountRequest_(' + kind + ') notify:', e.message); }
  return { ok: true, pendingId: pid };
}

function appSubmitBillingRequest(projectId, payload) { return appSubmitAmountRequest_('billing', projectId, payload); }
function appSubmitPaymentRequest(projectId, payload) { return appSubmitAmountRequest_('payment', projectId, payload); }

// 新規案件追加（既存 pmEnsureProjectRecord を再利用＝案件ID採番・辞書同期・Driveフォルダ）
function appAddProject(payload) {
  payload = payload || {};
  var name = String(payload.name || '').trim();
  if (!name) return { ok: false, msg: '案件名は必須です' };
  if (pmGetProjectByName(name)) return { ok: false, msg: '同名の案件が既にあります' };
  var rec = pmEnsureProjectRecord(name, {
    client:    String(payload.client || ''),
    assignee:  String(payload.assignee || ''),
    phase:     String(payload.phase || '営業'),
    updatedBy: appWho_(payload.who) || 'webapp',
  });
  return { ok: true, id: rec ? String(rec['案件ID'] || '') : '' };
}
