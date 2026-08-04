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
    // ?page=app → 進捗管理アプリ(新UI) / ?page=billing → 請求・入金 / ?page=tasks → タスク / それ以外 → 進捗ボード
    var page = (e && e.parameter) ? String(e.parameter.page || '').trim() : '';
    var pageFile = (page === 'app') ? 'pm_app' : (page === 'billing') ? 'billing' : (page === 'tasks') ? 'tasks' : 'dashboard';
    var titles = { pm_app: 'WOODBASE 進捗管理', billing: 'WOODBASE 請求・入金', tasks: 'WOODBASE タスク', dashboard: 'WOODBASE 進捗ボード' };
    // ページ間リンク用URLをテンプレートに注入（キー必須運用でも遷移できるように）
    var baseUrl = ScriptApp.getService().getUrl();
    var keyQuery = token ? ('&key=' + encodeURIComponent(key)) : '';
    var t = HtmlService.createTemplateFromFile(pageFile);
    t.boardUrl = baseUrl + '?page=board' + keyQuery;
    t.billingUrl = baseUrl + '?page=billing' + keyQuery;
    t.tasksUrl = baseUrl + '?page=tasks' + keyQuery;
    t.appUrl = baseUrl + '?page=app' + keyQuery;
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

// intent → 表示用ラベル（タイムライン）
var APP_INTENT_LABEL = {
  phase_update:    'フェーズを更新', status_update: '状況を更新', assignee_update: '担当者を変更',
  progress_update: '進捗を更新',     archive:       '案件をアーカイブ',
  billing_update:  '請求を更新',     payment_update:'入金を更新',
};
// 日時セルを yyyy-MM-dd HH:mm に。Dateでも文字列でも扱えるように。
function appFmtLogDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  return String(v == null ? '' : v).trim();
}
function appLogTime_(v) {
  if (v instanceof Date) return v.getTime();
  var d = new Date(String(v || '').replace(/-/g, '/'));
  return isNaN(d.getTime()) ? 0 : d.getTime();
}
// 案件ID→案件名 の辞書（タイムライン/お知らせの突き合わせ用）
function appProjectNameMap_() {
  var m = {};
  pmReadObjects(PM_SHEET_PROJECTS).forEach(function(r) { m[String(r['案件ID'] || '')] = String(r['案件名'] || ''); });
  return m;
}

// 更新履歴（project_logs）を新しい順に返す。案件名を突き合わせて表示用に整形。
function appGetTimeline(limit) {
  limit = Number(limit) || 40;
  if (!getSheet(PM_SHEET_LOGS)) return [];
  var names = appProjectNameMap_();
  var out = pmReadObjects(PM_SHEET_LOGS).map(function(r) {
    var changes = {};
    try { changes = JSON.parse(r['変更項目'] || '{}'); } catch (e) {}
    var parts = Object.keys(changes).map(function(k) { return k + '：' + changes[k]; });
    var intent = String(r['intent'] || '');
    return {
      projectId: String(r['案件ID'] || ''),
      name:      names[String(r['案件ID'] || '')] || String(r['案件ID'] || ''),
      intent:    intent,
      label:     APP_INTENT_LABEL[intent] || intent || '更新',
      detail:    parts.join('／'),
      who:       String(r['更新者'] || ''),
      date:      appFmtLogDate_(r['日時']),
      _t:        appLogTime_(r['日時']),
      kind:      String(r['適用区分'] || ''),
    };
  });
  out.sort(function(a, b) { return b._t - a._t; });
  return out.slice(0, limit);
}

// お知らせ＝アラート(project_alerts)を新しい順に。案件名を突き合わせ。
function appGetNotices(limit) {
  limit = Number(limit) || 30;
  if (!getSheet(PM_SHEET_ALERTS)) return [];
  var names = appProjectNameMap_();
  var rows = pmReadObjects(PM_SHEET_ALERTS).map(function(r) {
    return {
      projectId: String(r['案件ID'] || ''),
      name:      names[String(r['案件ID'] || '')] || '',
      type:      String(r['種別'] || ''),
      body:      String(r['内容'] || ''),
      date:      appFmtLogDate_(r['検知日時']),
      _t:        appLogTime_(r['検知日時']),
      status:    String(r['状態'] || ''),
    };
  });
  rows.sort(function(a, b) { return b._t - a._t; });
  return rows.slice(0, limit);
}

// 区分=社内/その他、取消=TRUE を除外した案件のみ返す
function appGetData() {
  var phases = Object.keys(PM_PHASES);
  var rows = pmReadObjects(PM_SHEET_PROJECTS);
  var storesByProject = appStoresByProject_();   // 案件ID→店舗配列（一括・N+1回避）
  var list = [];
  rows.forEach(function(r) {
    var kind = String(r['区分'] || '').trim();
    if (kind === '社内' || kind === 'その他') return;
    var canceled = String(r['取消'] || '').trim().toUpperCase();
    if (canceled === 'TRUE' || canceled === '1' || canceled === '✓' || canceled === 'YES') return;
    var phase = String(r['現在フェーズ'] || '営業');
    var pid = String(r['案件ID'] || '');
    list.push({
      id:         pid,
      name:       String(r['案件名'] || ''),
      client:     String(r['クライアント名'] || ''),
      assignee:   String(r['担当者'] || ''),
      phase:      phase,
      // 現在フェーズに対応するサブステータス（営業ステータス等の列値）
      status:     String(r[PM_PHASE_COLUMN[phase] || ''] || ''),
      updated:    appFmtDate_(r['最終更新日']),
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
      driveStatus:    String(r['Drive同期ステータス'] || (r['DriveフォルダURL'] ? 'ready' : 'missing')),
      // 新階層：案件フォルダURL（案件名クリックで開く・§6）＋店舗一覧（店舗名クリックで店舗フォルダ）
      caseFolderUrl:  String(r['driveProjectFolderUrl'] || ''),
      stores:         storesByProject[pid] || [],
    });
  });
  // statusMaster: フェーズ→サブステータス候補（カードのドロップダウン用）
  return { phases: phases, statusMaster: PM_PHASES, projects: list };
}

// stores シートを一括読みして 案件ID→[{storeId,name,storeFolderUrl,dataFolderUrl}] にまとめる
function appStoresByProject_() {
  var map = {};
  var rows = (typeof pmReadObjects === 'function' && getSheet(PM_SHEET_STORES)) ? pmReadObjects(PM_SHEET_STORES) : [];
  rows.forEach(function(s) {
    if (String(s['状態'] || '').trim().toLowerCase() === 'deleted') return;
    var pid = String(s['案件ID'] || '').trim();
    if (!pid) return;
    (map[pid] = map[pid] || []).push({
      storeId:        String(s['店舗ID'] || ''),
      name:           String(s['店舗名'] || ''),
      storeFolderUrl: String(s['driveStoreFolderUrl'] || ''),
      dataFolderUrl:  String(s['driveDataFolderUrl'] || ''),
      // 店舗ごとの進捗（無ければ空＝営業扱い）
      phase:          String(s['現在フェーズ'] || ''),
      status:         String(s['店舗ステータス'] || ''),
      progress:       String(s['施工進捗'] || ''),
      updated:        appFmtDate_(s['進捗更新日'] || s['更新日時']),
    });
  });
  return map;
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

// 管理者向け：未作成または失敗した案件だけを、既存の共通Driveサービスで再実行する。
function appCreateDriveFolder(projectId, who) {
  if (!projectId) return { ok: false, msg: '引数不足' };
  var locked = pmWithLock(function() {
    var proj = pmGetProjectById(projectId);
    if (!proj) return { ok: false, msg: '案件が見つかりません' };
    var result = pmEnsureProjectFolder(projectId, proj['案件名'], { source: 'progress-board-retry', actor: appActor_(who) });
    return result.ok ? { ok: true, url: result.url } : { ok: false, msg: 'Drive作成エラー' };
  }, 30000);
  if (!locked.ok) return { ok: false, msg: 'Drive作成が混み合っています。再度お試しください' };
  return locked.result;
}

// 取り消し（物理削除でなくアーカイブ）
function appArchiveProject(projectId, who) {
  if (!projectId) return { ok: false };
  return pmArchiveProject(projectId, appActor_(who), 'auto');
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

// 新規案件追加。案件ID採番・辞書同期は既存 pmEnsureProjectRecord を再利用しつつ、
// Driveは新階層(ROOT/案件名/店舗名/6_データ)を使う（旧構成は skipDriveFolder で停止）。
// payload.stores（配列 or 改行/カンマ区切り文字列）で店舗を同時登録できる（§2/§6）。
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
    source:    'progress-board',
    skipDriveFolder: true,   // 新階層を使うため旧「プロジェクト管理/{ID}_{名}」は作らない
  });
  var projectId = rec ? String(rec['案件ID'] || '') : '';

  var stores = appParseStoreNames_(payload.stores);
  var storeResults = [];
  if (stores.length) {
    stores.forEach(function(sn) {
      var sr = pmEnsureStoreRecord(projectId, sn, { caseName: name, source: 'progress-board', actor: appActor_(payload.who) });
      storeResults.push({ store: sn, ok: !!sr.ok, storeId: sr.storeId || '', dataFolderUrl: (sr.folders && sr.folders.dataFolderUrl) || '', error: sr.error || '' });
    });
  } else {
    // 店舗未指定でも案件フォルダは用意（案件名クリックで開けるように）
    pmEnsureCaseFolder(projectId, name, { source: 'progress-board', actor: appActor_(payload.who) });
  }
  return { ok: true, id: projectId, stores: storeResults };
}

// 既存案件へ店舗を追加（ダッシュボードの「店舗追加」用）。フォルダ生成＋台帳保存。
function appAddStore(projectId, storeName, who) {
  if (!projectId || !String(storeName || '').trim()) return { ok: false, msg: '案件と店舗名が必要です' };
  var proj = pmGetProjectById(projectId);
  if (!proj) return { ok: false, msg: '案件が見つかりません' };
  var sr = pmEnsureStoreRecord(projectId, String(storeName).trim(), { caseName: proj['案件名'], source: 'progress-board', actor: appActor_(who) });
  if (!sr.ok) return { ok: false, msg: sr.error || '店舗登録に失敗しました' };
  return { ok: true, storeId: sr.storeId, storeFolderUrl: sr.folders && sr.folders.storeFolderUrl, dataFolderUrl: sr.folders && sr.folders.dataFolderUrl };
}

// 店舗名の複数入力を配列へ（改行/カンマ/読点区切り、空・重複除去）
function appParseStoreNames_(raw) {
  if (!raw) return [];
  var arr = Array.isArray(raw) ? raw : String(raw).split(/[\n,、，]+/);
  var seen = {}, out = [];
  arr.forEach(function(s) { var t = String(s == null ? '' : s).trim(); if (t && !seen[t]) { seen[t] = 1; out.push(t); } });
  return out;
}

// ==========================================
// 店舗ごとの進捗（stores シートへ書き込み）。案件と同じフェーズ体系を使う。
//   列が未追加の環境でも動くよう、書込前に pmEnsureStoreSheet で列を保証する。
// ==========================================
function appSetStoreField_(storeId, colName, value, who, intent, validate) {
  if (!storeId) return { ok: false, msg: '店舗が不明です' };
  pmEnsureStoreSheet();  // 進捗列を非破壊で保証
  var sh = getSheet(PM_SHEET_STORES);
  if (!sh) return { ok: false, msg: 'storesシートがありません' };
  var data = sh.getDataRange().getValues();
  var h = data[0].map(function(x) { return String(x).trim(); });
  var cId = h.indexOf('店舗ID'), cCol = h.indexOf(colName),
      cUpd = h.indexOf('進捗更新日'), cWho = h.indexOf('進捗更新者'), cPid = h.indexOf('案件ID'), cName = h.indexOf('店舗名');
  if (cCol === -1) return { ok: false, msg: colName + ' 列がありません' };
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][cId] || '').trim() !== String(storeId).trim()) continue;
    if (validate) { var v = validate(String(data[i][h.indexOf('現在フェーズ')] || '')); if (v) return v; }
    sh.getRange(i + 1, cCol + 1).setValue(value);
    if (cUpd !== -1) sh.getRange(i + 1, cUpd + 1).setValue(pmTodayYmd());
    if (cWho !== -1 && appWho_(who)) sh.getRange(i + 1, cWho + 1).setValue(appWho_(who));
    var chg = { '店舗': String(data[i][cName] || '') }; chg[colName] = value;
    try { pmAddLog(String(data[i][cPid] || ''), intent, chg, '', appActor_(who), '', 'auto'); } catch (e2) {}
    return { ok: true };
  }
  return { ok: false, msg: '店舗が見つかりません' };
}

function appSetStorePhase(storeId, phase, who) {
  if (!phase || !PM_PHASES[phase]) return { ok: false, msg: 'フェーズが不正です' };
  return appSetStoreField_(storeId, '現在フェーズ', phase, who, 'store_phase_update');
}
function appSetStoreStatus(storeId, status, who) {
  status = String(status || '').trim();
  return appSetStoreField_(storeId, '店舗ステータス', status, who, 'store_status_update', function(phase) {
    if (status && phase && (PM_PHASES[phase] || []).indexOf(status) === -1) {
      return { ok: false, msg: '「' + phase + '」にないステータスです: ' + status };
    }
    return null;
  });
}
function appSetStoreProgress(storeId, pct, who) {
  var n = Number(pct);
  if (isNaN(n) || n < 0 || n > 100) return { ok: false, msg: '進捗は0〜100で指定してください' };
  return appSetStoreField_(storeId, '施工進捗', n, who, 'store_progress_update');
}

// ==========================================
// LINEグループ紐付け（アプリUI用・§LINE連携）
//   自動通知は行わない。既存の linkGroupToProject / 案件グループ台帳 / メッセージログ を再利用する。
// ==========================================

// この案件に紐付くLINEグループ＋紐付け候補（メッセージログの既知グループ）を返す
function appGetLineGroups(projectId) {
  var proj = pmGetProjectById(projectId);
  if (!proj) return { ok: false, msg: '案件が見つかりません' };
  var name = String(proj['案件名'] || '');
  var linked = [];
  // ① 案件グループ台帳（1案件に複数グループ可）
  if (getSheet(PM_SHEET_GROUPS)) {
    pmReadObjects(PM_SHEET_GROUPS).forEach(function(g) {
      if (String(g['案件ID'] || '') === String(projectId)) {
        linked.push({ groupId: String(g['グループID'] || ''), kind: String(g['種別'] || ''), label: String(g['表示名'] || '') });
      }
    });
  }
  // ② projects.LINEグループID（従来の単一列）も拾う
  var primary = String(proj['LINEグループID'] || '').trim();
  if (primary && !linked.some(function(x) { return x.groupId === primary; })) linked.push({ groupId: primary, kind: '主', label: '' });

  return { ok: true, projectName: name, linked: linked, candidates: appKnownLineGroups_(24) };
}

// メッセージログから既知グループを集計（最近の活動順・現在の紐付け案件名つき）
function appKnownLineGroups_(limit) {
  var sh = getSheet('メッセージログ');
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var gid = String(data[i][1] || '').trim();
    if (!gid) continue;
    var o = map[gid] || (map[gid] = { groupId: gid, last: '', count: 0, sample: '' });
    o.count++;
    var dt = data[i][0];
    var dts = (dt instanceof Date) ? Utilities.formatDate(dt, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : String(dt || '');
    if (dts > o.last) { o.last = dts; o.sample = String(data[i][3] || '').slice(0, 30); }
  }
  var arr = Object.keys(map).map(function(k) {
    var o = map[k];
    try { o.linkedTo = getProjectNameByGroupId(o.groupId) || ''; } catch (e) { o.linkedTo = ''; }
    return o;
  });
  arr.sort(function(a, b) { return a.last < b.last ? 1 : a.last > b.last ? -1 : 0; });
  return arr.slice(0, limit || 24);
}

// 案件にLINEグループを紐付け（既存 linkGroupToProject＋台帳upsertを利用）。物理送信なし。
function appLinkLineGroup(projectId, groupId) {
  if (!projectId || !String(groupId || '').trim()) return { ok: false, msg: '案件とグループが必要です' };
  var proj = pmGetProjectById(projectId);
  if (!proj) return { ok: false, msg: '案件が見つかりません' };
  var name = String(proj['案件名'] || '');
  var gid = String(groupId).trim();
  try {
    if (typeof linkGroupToProject === 'function') linkGroupToProject(gid, name);
    // プロジェクト管理(辞書)に無い案件でも台帳を確実に更新（linkGroupToProjectが早期returnするケースの保険）
    if (typeof upsertProjectGroupLink_ === 'function') upsertProjectGroupLink_(gid, name);
  } catch (e) {
    return { ok: false, msg: '紐付けエラー：' + e.message };
  }
  try { pmAddLog(projectId, 'line_link', { 'LINEグループ': gid }, '', 'webapp', gid, 'auto'); } catch (e2) {}
  return { ok: true };
}

// 紐付け解除（案件グループ台帳の該当行の案件IDを空に）。物理送信なし。
function appUnlinkLineGroup(projectId, groupId) {
  if (!projectId || !String(groupId || '').trim()) return { ok: false, msg: '引数不足' };
  var sh = getSheet(PM_SHEET_GROUPS);
  if (!sh || sh.getLastRow() < 2) return { ok: false, msg: '案件グループ台帳がありません' };
  var data = sh.getDataRange().getValues();
  var h = data[0].map(function(x) { return String(x).trim(); });
  var cGid = h.indexOf('グループID'), cPid = h.indexOf('案件ID');
  if (cGid === -1 || cPid === -1) return { ok: false, msg: '台帳の列が不正です' };
  var gid = String(groupId).trim();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][cGid] || '').trim() === gid && String(data[i][cPid] || '').trim() === String(projectId).trim()) {
      sh.getRange(i + 1, cPid + 1).clearContent();
      try { pmAddLog(projectId, 'line_unlink', { 'LINEグループ': gid }, '', 'webapp', gid, 'auto'); } catch (e2) {}
      return { ok: true };
    }
  }
  return { ok: false, msg: '該当の紐付けが見つかりません' };
}
