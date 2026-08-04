// ==========================================
// アプリ↔GAS ジョブキュー（AppSheet設計 パイロット1/2 §4・§5）
//   UI（AppSheet等）は「ジョブ行を1行追加」するだけ。毎分トリガー runAppJobs() が
//   受付ジョブを処理し、状態/結果URL/メッセージをジョブ行へ書き戻し、依頼者へLINE通知する。
//   方針:
//   - 既存を壊さない: 帳票生成・承認・カレンダー・Driveは既存関数をそのまま再利用。
//   - 冪等・多重実行安全: クレーム（受付→処理中）を pmWithLock で直列化。処理はロック外。
//   - シート未セットアップなら何もしない（安全にスキップ）。
//   依存: pm_core.js(pmSheet/pmReadObjects/pmWriteRowFields/pmWithLock/pmIsBlank/pmInternalTarget),
//         pm_pending.js(pmApprovePendingCore/pmRejectPendingCore), pm_manager.js(pmGetProjectById/pmAddLog),
//         pm_calendar.js(pmSyncCalendar), pm_drive.js(pmEnsureProjectFolder),
//         estimate.js(getEstimateLinkSettings/createEstimateFromProcessSheet),
//         estimate_pioneer.js(createEstimateSheets/runEstimateTestOneStore),
//         コード.js(getSS/setupSheet/generateId/fmtDT/sendLineMessage)
// ==========================================

// ---- シート定義 ----
var PM_SHEET_EST_JOBS   = '見積ジョブ';
var PM_EST_JOBS_HEADERS = ['ジョブID', '依頼日時', '依頼者', '設定ID', '対象月', '上書き', 'テストモード', '状態', '結果URL', 'メッセージ', '処理開始', '処理完了'];

var PM_SHEET_PM_JOBS    = 'pmジョブ';
var PM_PM_JOBS_HEADERS  = ['ジョブID', '依頼日時', '依頼者', '種別', '案件ID', '確認ID', 'パラメータ', '状態', '結果URL', 'メッセージ', '処理開始', '処理完了'];

// アプリ利用者の権限マスタ（メール→役割）。コードに埋め込まずシートで増減できる（設計v1 §7-3）。
// LINEユーザーID を書くとジョブ完了時にその人へDM通知（空なら社内グループへフォールバック）。
var PM_SHEET_APP_ROLES   = 'アプリ権限';
var PM_APP_ROLES_HEADERS = ['メール', '役割', '表示名', 'LINEユーザーID', '備考'];
var PM_APP_ROLE_ADMIN      = '管理者';
var PM_APP_ROLE_ACCOUNTING = '経理';

// ---- ジョブ状態 ----
var PM_JOB_STATE_RECEIVED = '受付';
var PM_JOB_STATE_RUNNING  = '処理中';
var PM_JOB_STATE_DONE     = '完了';
var PM_JOB_STATE_ERROR    = 'エラー';

// ---- 実行制御 ----
var PM_JOBS_TIME_BUDGET_MS       = 270000;  // 1回の実行で使う時間上限 4.5分（GAS上限6分に余裕）
var PM_JOBS_STALE_MS             = 1800000; // 処理中のまま30分超はタイムアウト回収
var PM_JOBS_MAX_ESTIMATE_PER_RUN = 2;       // 帳票生成は重いので1回2件まで
var PM_JOBS_MAX_PM_PER_RUN       = 10;

// ==========================================
// セットアップ（GASエディタから pmJobsSetupAll() を1回実行）
// ==========================================
function pmJobsSetupSheets() {
  var ss = getSS();
  setupSheet(ss, PM_SHEET_EST_JOBS,  PM_EST_JOBS_HEADERS,  '#3D85C6', null);
  setupSheet(ss, PM_SHEET_PM_JOBS,   PM_PM_JOBS_HEADERS,   '#674EA7', null);
  setupSheet(ss, PM_SHEET_APP_ROLES, PM_APP_ROLES_HEADERS, '#E69138', null);
  console.log('pmJobsSetupSheets: 完了');
}

// runAppJobs の毎分トリガーを登録（重複登録しない）
function pmJobsSetupTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'runAppJobs'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('runAppJobs').timeBased().everyMinutes(1).create();
  console.log('pmJobsSetupTriggers: runAppJobs を毎分実行で登録');
}

function pmJobsSetupAll() {
  pmJobsSetupSheets();
  pmJobsSetupTriggers();
}

// ==========================================
// アプリ権限（メール→役割）
//   アプリ権限シートに行が1つも無い間は「制限なし」（既存 PM_APPROVER_USER_IDS と同じ思想。
//   試運転を止めないため）。行を入れた時点で登録メールのみ許可の厳格運用に切り替わる。
// ==========================================
function appGetRoleRow(email) {
  if (!email) return null;
  var key = String(email).trim().toLowerCase();
  if (!key) return null;
  var rows = pmReadObjects(PM_SHEET_APP_ROLES);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['メール'] || '').trim().toLowerCase() === key) return rows[i];
  }
  return null;
}

function appRoleRowCount() {
  var n = 0;
  pmReadObjects(PM_SHEET_APP_ROLES).forEach(function(r) { if (!pmIsBlank(r['メール'])) n++; });
  return n;
}

// アプリ権限の表示名（未登録・表示名空欄ならメールのまま）。申請者・最終更新者の記録に使う。
function appDisplayName(email) {
  var row = appGetRoleRow(email);
  var name = row ? String(row['表示名'] || '').trim() : '';
  return name || String(email || '').trim();
}

// ジョブ実行の可否（見積発行・カレンダー等の一般操作）
function appCanRunJobs(email) {
  if (appRoleRowCount() === 0) return true; // 未運用時は制限なし
  return appGetRoleRow(email) !== null;
}

// 金額/請求/入金の承認・却下の可否（経理/管理者のみ）
function appCanApproveFinancial(email) {
  if (appRoleRowCount() === 0) return true; // 未運用時は制限なし
  var row = appGetRoleRow(email);
  if (!row) return false;
  var role = String(row['役割'] || '').trim();
  return role === PM_APP_ROLE_ADMIN || role === PM_APP_ROLE_ACCOUNTING;
}

// 依頼者への完了/失敗通知。アプリ権限のLINEユーザーIDがあればDM、無ければ社内グループへ。
// 通知失敗はジョブ結果に影響させない。
function appNotifyRequester(email, text) {
  try {
    var row = appGetRoleRow(email);
    var lineId = row ? String(row['LINEユーザーID'] || '').trim() : '';
    if (lineId) { sendLineMessage(lineId, text); return; }
    var internal = pmInternalTarget();
    if (internal) sendLineMessage(internal, (email ? '【' + email + ' さんの依頼】\n' : '') + text);
  } catch (e) {
    console.error('appNotifyRequester:', e.message);
  }
}

// ==========================================
// 純ロジック（テスト対象）
// ==========================================

// 対象月のパース: 'YYYY-MM' / 'YYYY/M' / 'YYYY年M月' / '当月' / '先月' / Date → {year, month}。不正は null
function pmJobsParseTargetMonth(v, nowDate) {
  var now = nowDate || new Date();
  if (v instanceof Date && !isNaN(v.getTime())) return { year: v.getFullYear(), month: v.getMonth() + 1 };
  if (v === null || v === undefined) return null;
  var s = String(v).trim();
  if (!s) return null;
  if (s === '当月' || s === '今月') return { year: now.getFullYear(), month: now.getMonth() + 1 };
  if (s === '先月') {
    var p = new Date(now.getTime());
    p.setDate(1);
    p.setMonth(p.getMonth() - 1);
    return { year: p.getFullYear(), month: p.getMonth() + 1 };
  }
  var m = s.match(/^(\d{4})\s*[-\/年.]\s*(\d{1,2})\s*月?$/);
  if (!m) return null;
  var month = parseInt(m[2], 10);
  if (month < 1 || month > 12) return null;
  return { year: parseInt(m[1], 10), month: month };
}

// チェックボックス/Yes-No系のゆるい真偽判定（AppSheet・手入力の表記ゆれ吸収）
function pmJobsTruthy(v) {
  if (v === true) return true;
  var s = String(v === null || v === undefined ? '' : v).trim().toUpperCase();
  return s === 'TRUE' || s === '1' || s === 'YES' || s === 'Y' || s === 'ON' || s === '✓' || s === 'はい';
}

// 金額申請ジョブ（請求申請/入金申請）のパラメータ解析。
//   JSON形式:      {"金額":1000000,"予定日":"2026-08-01","確定日":"2026-08-05"}
//   key=value形式: 金額=1000000; 予定日=2026-08-01（区切りは ; 、改行。=の代わりに : も可）
//   すべて任意（金額/予定日/確定日のいずれか必須のチェックは appSubmitAmountRequest_ 側で行う）。
//   戻り値: { ok:true, amount, dueDate, doneDate } | { ok:false, msg }
function pmJobsParseAmountParams(v) {
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (!s) return { ok: true, amount: '', dueDate: '', doneDate: '' };

  var obj = safeParseJson(s);
  if (obj && typeof obj === 'object') {
    var amt = obj['金額'] !== undefined ? obj['金額'] : obj['amount'];
    return {
      ok: true,
      amount:   (amt === null || amt === undefined) ? '' : String(amt).replace(/[,，円\s]/g, ''),
      dueDate:  String(obj['予定日'] || obj['dueDate'] || '').trim(),
      doneDate: String(obj['確定日'] || obj['doneDate'] || '').trim(),
    };
  }

  var out = { amount: '', dueDate: '', doneDate: '' };
  var known = false;
  s.split(/[;、\n]+/).forEach(function(part) {
    var m = String(part).trim().match(/^(金額|amount|予定日|dueDate|確定日|doneDate)\s*[=＝:：]\s*(.+)$/i);
    if (!m) return;
    known = true;
    var val = m[2].trim();
    if (/^(金額|amount)$/i.test(m[1]))       out.amount = val.replace(/[,，円\s]/g, '');
    else if (/^(予定日|dueDate)$/i.test(m[1])) out.dueDate = val;
    else                                       out.doneDate = val;
  });
  if (!known) return { ok: false, msg: 'パラメータが読み取れません。例: 金額=1000000; 予定日=2026-08-01' };
  return { ok: true, amount: out.amount, dueDate: out.dueDate, doneDate: out.doneDate };
}

// 処理中のまま放置されたジョブか（クラッシュ・タイムアウト回収用）
function pmJobsIsStale(startedVal, now) {
  var d = startedVal instanceof Date
    ? startedVal
    : (pmIsBlank(startedVal) ? null : new Date(String(startedVal)));
  if (!d || isNaN(d.getTime())) return true; // 開始時刻が読めない処理中行は回収対象
  return (now.getTime() - d.getTime()) > PM_JOBS_STALE_MS;
}

// ==========================================
// ジョブのクレームと共通ランナー
// ==========================================

// 受付（または状態空欄）のジョブを最大 maxCount 件、処理中に更新して返す。
// あわせて: ジョブID空欄の自動採番、処理中のまま30分超のタイムアウト回収。
function pmJobsClaim(sheetName, maxCount) {
  var locked = pmWithLock(function() {
    var rows = pmReadObjects(sheetName);
    var claimed = [];
    var now = new Date();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var state = String(r['状態'] || '').trim();
      if (state === PM_JOB_STATE_RUNNING) {
        if (pmJobsIsStale(r['処理開始'], now)) {
          pmWriteRowFields(sheetName, r._row, {
            '状態': PM_JOB_STATE_ERROR,
            'メッセージ': 'タイムアウト（処理が完了しませんでした。再実行する場合は状態を「受付」に戻してください）',
            '処理完了': fmtDT(now),
          });
        }
        continue;
      }
      if (claimed.length >= maxCount) continue;
      if (state === '' || state === PM_JOB_STATE_RECEIVED) {
        var fields = { '状態': PM_JOB_STATE_RUNNING, '処理開始': fmtDT(now), '処理完了': '' };
        if (pmIsBlank(r['ジョブID'])) { fields['ジョブID'] = generateId(); r['ジョブID'] = fields['ジョブID']; }
        if (pmIsBlank(r['依頼日時'])) { fields['依頼日時'] = fmtDT(now); }
        pmWriteRowFields(sheetName, r._row, fields);
        claimed.push(r);
      }
    }
    return claimed;
  }, 30000);
  return locked.ok ? locked.result : [];
}

// キュー1本を処理する共通ランナー。processor(job) は
// { state: PM_JOB_STATE_DONE|PM_JOB_STATE_ERROR, url?, message? } を返す。
function pmJobsRunQueue(sheetName, maxCount, processor, label, deadlineMs) {
  if (!pmSheet(sheetName)) return 0; // シート未セットアップ＝安全にスキップ
  var deadline = deadlineMs || (Date.now() + PM_JOBS_TIME_BUDGET_MS);
  var jobs = pmJobsClaim(sheetName, maxCount);
  var processed = 0;
  for (var i = 0; i < jobs.length; i++) {
    var job = jobs[i];
    if (Date.now() > deadline) {
      // 時間切れ: クレーム済み未処理分は受付へ戻す（次回実行で処理）
      pmWriteRowFields(sheetName, job._row, { '状態': PM_JOB_STATE_RECEIVED, '処理開始': '' });
      continue;
    }
    var res;
    try {
      res = processor(job);
    } catch (e) {
      console.error('pmJobsRunQueue(' + label + ') error:', e.message, e.stack);
      res = { state: PM_JOB_STATE_ERROR, message: 'エラー: ' + e.message };
    }
    if (!res || !res.state) res = { state: PM_JOB_STATE_ERROR, message: '処理結果が不明です' };
    pmWriteRowFields(sheetName, job._row, {
      '状態': res.state,
      '結果URL': res.url || '',
      'メッセージ': res.message || '',
      '処理完了': fmtDT(new Date()),
    });
    var head = (res.state === PM_JOB_STATE_DONE ? '✅ ' + label + 'ジョブ完了' : '⚠️ ' + label + 'ジョブ失敗');
    appNotifyRequester(String(job['依頼者'] || '').trim(),
      head + '\n' + (res.message || '') + (res.url ? '\n' + res.url : ''));
    processed++;
  }
  return processed;
}

// ==========================================
// エントリポイント（毎分トリガー）
// ==========================================
function runAppJobs() {
  var deadline = Date.now() + PM_JOBS_TIME_BUDGET_MS;
  var nEst = runEstimateJobs(deadline);
  var nPm  = runPmJobs(deadline);
  if (nEst + nPm > 0) console.log('runAppJobs: 見積=' + nEst + '件 / pm=' + nPm + '件 処理');
}

// 見積ジョブ（パイロット1 §4）。単体実行も可。
function runEstimateJobs(deadlineMs) {
  return pmJobsRunQueue(PM_SHEET_EST_JOBS, PM_JOBS_MAX_ESTIMATE_PER_RUN, pmJobsProcessEstimateJob, '見積', deadlineMs);
}

// pmジョブ（パイロット2 §5: 承認/却下・カレンダー同期・Driveフォルダ）。単体実行も可。
function runPmJobs(deadlineMs) {
  return pmJobsRunQueue(PM_SHEET_PM_JOBS, PM_JOBS_MAX_PM_PER_RUN, pmJobsProcessPmJob, 'pm', deadlineMs);
}

// ==========================================
// 見積ジョブの処理
//   発行方式はプロファイル（見積連携設定.発行方式）で吸収（設計v1 §7-4）:
//   - 月別: 対象月必須。createEstimateFromProcessSheet を呼ぶ（東邦系）
//   - 完了行一括: 対象月無視。createEstimateSheets / テストモード時 runEstimateTestOneStore（パイオニア系）
// ==========================================
function pmJobsProcessEstimateJob(job) {
  var requester = String(job['依頼者'] || '').trim();
  if (!appCanRunJobs(requester)) {
    return { state: PM_JOB_STATE_ERROR, message: '依頼者がアプリ権限に未登録です: ' + (requester || '(空欄)') };
  }

  var settingId = String(job['設定ID'] || '').trim();
  if (!settingId) return { state: PM_JOB_STATE_ERROR, message: '設定IDが未指定です' };
  var setting = null;
  getEstimateLinkSettings().forEach(function(s) { if (!setting && s.settingId === settingId) setting = s; });
  if (!setting) {
    return { state: PM_JOB_STATE_ERROR, message: '設定「' + settingId + '」が見つからないか無効です（見積連携設定シートを確認）' };
  }

  if (setting.issueMode === '完了行一括') {
    var isTest = pmJobsTruthy(job['テストモード']);
    var log = isTest ? runEstimateTestOneStore() : createEstimateSheets();
    if (log.error) return { state: PM_JOB_STATE_ERROR, message: log.error };
    if (!log.created) {
      return { state: PM_JOB_STATE_ERROR, message: '対象行がありませんでした' + (log.note ? '（' + log.note + '）' : '') };
    }
    var files = log.files || [];
    var lines = files.map(function(f) { return '・' + f.store + '：' + f.url; });
    return {
      state: PM_JOB_STATE_DONE,
      url: files.length ? files[0].url : '',
      message: (isTest ? '[テスト先頭1店舗] ' : '') + setting.displayName + '：' + log.created + '件の見積書を作成\n' + lines.join('\n'),
    };
  }

  // 月別（既定）
  var ym = pmJobsParseTargetMonth(job['対象月']);
  if (!ym) {
    return { state: PM_JOB_STATE_ERROR, message: '対象月が読み取れません（例: 2026-05）: ' + String(job['対象月'] || '(空欄)') };
  }
  var overwrite = pmJobsTruthy(job['上書き']);
  var result = createEstimateFromProcessSheet(setting, ym.year, ym.month, overwrite);
  switch (result.kind) {
    case 'success':
      return {
        state: PM_JOB_STATE_DONE,
        url: result.spreadsheetUrl,
        message: setting.displayName + '：「' + result.sheetName + '」を作成しました',
      };
    case 'exists':
      return {
        state: PM_JOB_STATE_ERROR,
        url: result.spreadsheetUrl,
        message: '既に「' + result.sheetName + '」があります。上書き=TRUE にして再実行してください',
      };
    case 'noData':
      return { state: PM_JOB_STATE_ERROR, message: '工程管理表に ' + result.year + '年' + result.month + '月 のデータがありません' };
    case 'configError':
      return { state: PM_JOB_STATE_ERROR, message: '設定に必須項目が未設定: ' + result.missing.join('、') };
    case 'openError':
      return { state: PM_JOB_STATE_ERROR, message: result.target + 'を開けません（ID: ' + result.id + '）: ' + result.message };
    default:
      return { state: PM_JOB_STATE_ERROR, message: '不明な結果です: ' + String(result.kind) };
  }
}

// ==========================================
// pmジョブの処理（種別で分岐）
//   承認/却下: 確認ID 必須。金額系は経理/管理者のみ（アプリ権限シートで判定）
//   請求申請/入金申請: 案件ID 必須。projectsへは書かず承認待ちに登録→経理へLINE通知
//   カレンダー同期 / Driveフォルダ: 案件ID 必須。既存関数を再利用（冪等）
// ==========================================
var PM_JOB_PENDING_CODE_MESSAGES = {
  lock_timeout: '混雑のため処理できませんでした（自動で再実行されます。状態を「受付」に戻してください）',
  notfound:     '確認データが見つかりませんでした',
  done:         'この確認は既に処理済みです',
  forbidden:    '金額・請求・入金の承認/却下は経理・管理者のみ可能です',
  parse_error:  '抽出内容の解析に失敗しました。手動で確認してください',
  no_project:   '対象案件が見つかりませんでした',
  apply_failed: '反映に失敗しました',
};

function pmJobsProcessPmJob(job) {
  var requester = String(job['依頼者'] || '').trim();
  if (!appCanRunJobs(requester)) {
    return { state: PM_JOB_STATE_ERROR, message: '依頼者がアプリ権限に未登録です: ' + (requester || '(空欄)') };
  }

  var kind = String(job['種別'] || '').trim();
  var projectId = String(job['案件ID'] || '').trim();

  if (kind === '承認' || kind === '却下') {
    var pendingId = String(job['確認ID'] || '').trim();
    if (!pendingId) return { state: PM_JOB_STATE_ERROR, message: '確認IDが未指定です' };
    var canFn = function() { return appCanApproveFinancial(requester); };
    var r = (kind === '承認')
      ? pmApprovePendingCore(pendingId, requester, canFn)
      : pmRejectPendingCore(pendingId, requester, canFn);
    if (r.code === 'ok_project' || r.code === 'ok_financial') {
      var appliedText = Object.keys(r.applied).map(function(k) { return '・' + k + '：' + r.applied[k]; }).join('\n');
      return { state: PM_JOB_STATE_DONE, message: '承認し「' + r.projectName + '」へ確定反映しました\n' + appliedText };
    }
    if (r.code === 'ok') {
      return { state: PM_JOB_STATE_DONE, message: '却下しました（projectsへは反映しません）' };
    }
    var msg = PM_JOB_PENDING_CODE_MESSAGES[r.code] || ('処理できませんでした: ' + r.code);
    if (r.code === 'done' && r.status) msg += '（' + r.status + '）';
    return { state: PM_JOB_STATE_ERROR, message: msg };
  }

  // 金額系の申請（AppSheetの金額申請フォーム→ジョブ行）。依頼者=USEREMAIL() がそのまま申請者になる。
  // 確定はしない: pmQueuePending→経理グループへLINE通知→承認後に反映（ダッシュボードと同一経路）。
  if (kind === '請求申請' || kind === '入金申請') {
    if (!projectId) return { state: PM_JOB_STATE_ERROR, message: '案件IDが未指定です' };
    if (!requester) return { state: PM_JOB_STATE_ERROR, message: '依頼者（メール）が未指定です' };
    var params = pmJobsParseAmountParams(job['パラメータ']);
    if (!params.ok) return { state: PM_JOB_STATE_ERROR, message: params.msg };
    var sub = appSubmitAmountRequest_(kind === '請求申請' ? 'billing' : 'payment', projectId, {
      requester: appDisplayName(requester),
      amount:    params.amount,
      dueDate:   params.dueDate,
      doneDate:  params.doneDate,
    });
    if (!sub.ok) return { state: PM_JOB_STATE_ERROR, message: sub.msg || '申請できませんでした' };
    return { state: PM_JOB_STATE_DONE, message: '経理へ承認依頼を送りました（確認ID: ' + sub.pendingId + '）。承認後に反映されます' };
  }

  if (kind === 'カレンダー同期') {
    if (!projectId) return { state: PM_JOB_STATE_ERROR, message: '案件IDが未指定です' };
    if (!pmGetProjectById(projectId)) return { state: PM_JOB_STATE_ERROR, message: '案件が見つかりません: ' + projectId };
    pmSyncCalendar(projectId);
    return { state: PM_JOB_STATE_DONE, message: 'カレンダーを同期しました（' + projectId + '）' };
  }

  if (kind === 'Driveフォルダ') {
    if (!projectId) return { state: PM_JOB_STATE_ERROR, message: '案件IDが未指定です' };
    var proj = pmGetProjectById(projectId);
    if (!proj) return { state: PM_JOB_STATE_ERROR, message: '案件が見つかりません: ' + projectId };
    var res = pmEnsureProjectFolder(projectId, proj['案件名'], { source: 'app-job', actor: requester });
    if (!res || !res.ok) return { state: PM_JOB_STATE_ERROR, message: (res && res.error) || 'フォルダを作成できませんでした（DRIVE_FOLDER_ID 設定を確認）' };
    var url = res.url;
    pmAddLog(projectId, 'app_job', { 'DriveフォルダURL': url }, 'アプリ: Driveフォルダ作成', requester, '', 'auto');
    return { state: PM_JOB_STATE_DONE, url: url, message: '「' + (proj['案件名'] || projectId) + '」のDriveフォルダを用意しました' };
  }

  return { state: PM_JOB_STATE_ERROR, message: '未対応の種別です: ' + (kind || '(空欄)') + '（承認/却下/請求申請/入金申請/カレンダー同期/Driveフォルダ）' };
}
