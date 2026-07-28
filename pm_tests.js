// ==========================================
// プロジェクト進捗管理：テスト・診断（GASエディタから実行）
//   依存: pm_* 各モジュール
// ==========================================

// intent判定（プレフィックス）の確認

function pmTestClassifyIntent() {
  var cases = [
    ['案件更新 〇〇ホテル、見積提出しました。6/25にフォローします。', 'project_update'],
    ['新規案件 案件名:〇〇ホテル客室改装 クライアント:ABC運営 担当:濱田', 'project_create'],
    ['請求更新 〇〇ホテル、中間金を7/1に請求予定です。', 'billing_update'],
    ['入金更新 〇〇ホテル、7/10入金予定。', 'payment_update'],
    ['予定追加 〇〇ホテル 7/3 14時 現地打合せ', 'calendar_update'],
    ['濱田さん 見積もりを4月25日までにお願いします', null], // 既存タスクへ委譲
    ['ありがとうございます', null],
  ];
  cases.forEach(function(c) {
    var got = pmClassifyIntent(c[0]);
    console.log((got === c[1] ? '✅' : '❌') + ' "' + c[0].slice(0, 24) + '..." → ' + got + ' (期待:' + c[1] + ')');
  });
}

// ステータス正規化の確認
function pmTestStatusMapping() {
  var cases = [
    ['営業', '見積提出', '見積提出（フォロー中）'],
    ['営業', '受注しました', '受注'],
    ['施工', '着工しました', '着工'],
    [null, '基本設計中', '基本設計中'],
  ];
  cases.forEach(function(c) {
    var r = pmNormalizeStatus(c[0], c[1]);
    console.log((r.status === c[2] ? '✅' : '⚠️') + ' phase=' + c[0] + ' "' + c[1] + '" → phase=' + r.phase + ' status=' + r.status + ' matched=' + r.matched);
  });
}

// Gemini抽出（要 GEMINI_API_KEY）。projectsへは書き込まない。
function pmTestParseProjectUpdate() {
  var parsed = pmParseReport('〇〇ホテル、見積提出しました。6/25にフォローします。', 'project_update');
  console.log(JSON.stringify(parsed, null, 2));
  console.log('needs_confirmation:', parsed.needs_confirmation);
}

function pmTestParseBilling() {
  var parsed = pmParseReport('〇〇ホテル、中間金を7/1に請求予定です。', 'billing_update');
  console.log(JSON.stringify(parsed, null, 2));
  console.log('needs_confirmation（trueであるべき）:', parsed.needs_confirmation);
}

// セットアップの非破壊性：2回実行して列が増殖しないこと
function pmTestSetupSheetsNonDestructive() {
  pmSetupSheets();
  var before = getSheet(PM_SHEET_PROJECTS).getLastColumn();
  pmSetupSheets();
  var after = getSheet(PM_SHEET_PROJECTS).getLastColumn();
  console.log((before === after ? '✅' : '❌') + ' projects列数 before=' + before + ' after=' + after);
  var taskIdx = pmHeaderIndex(getSheet('タスク管理'));
  console.log('タスク管理 関連案件ID列:', taskIdx['関連案件ID'] !== undefined ? '✅' : '❌',
    '/ 作成元列:', taskIdx['作成元'] !== undefined ? '✅' : '❌');
}

// 案件解決（読み取りのみ）：曖昧時に候補が複数返ることを確認
function pmTestResolveAmbiguous() {
  var r = pmResolveProject('プラージュの件です', '', '');
  console.log('status:', r.status);
  if (r.candidates) r.candidates.forEach(function(c) { console.log('  候補:', c.name, c.confidence, c.reason); });
}

// 接続診断：必要シート・プロパティの有無
function pmDiagnose() {
  [PM_SHEET_PROJECTS, PM_SHEET_LOGS, PM_SHEET_ALERTS, PM_SHEET_CALENDAR, PM_SHEET_PENDING, PM_SHEET_SETTINGS, PM_SHEET_INTERNAL_GROUPS].forEach(function(n) {
    console.log((getSheet(n) ? '✅' : '❌ 未作成') + ' シート: ' + n);
  });
  console.log('経理通知先:', pmAccountingTarget() ? '設定済（' + (pmRoleGroupId('経理') ? '社内グループシート' : 'プロパティ/PM設定 or 社内へフォールバック') + '）' : '未設定');
  console.log('社内通知先:', pmInternalTarget() ? '設定済（' + (pmRoleGroupId('社内') ? '社内グループシート' : 'プロパティ/PM設定') + '）' : '未設定');
  console.log('日次通知先（役割）:', pmRoleGroupId('日次通知') ? '社内グループシートで設定済' : '未設定（プロパティ or 既定の善波グループ）');
  console.log('PM_APPROVER_USER_IDS:', pmListApproverUserIds().length ? ('設定済（' + pmListApproverUserIds().length + '名）') : '未設定（暫定で全員許可・本番は要設定）');
  console.log('CALENDAR_ID:', getConfig().CALENDAR_ID ? '設定済' : '未設定');
  console.log('DRIVE_FOLDER_ID:', getConfig().DRIVE_FOLDER_ID ? '設定済' : '未設定');
  var triggers = ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === 'pmCheckAlerts'; });
  console.log('pmCheckAlertsトリガー:', triggers.length ? '✅ 登録済' : '❌ 未登録（pmSetupTriggers実行）');
}

// ==========================================
// 追加修正の単体テスト（GASエディタから実行・LINE送信なし）
// ==========================================

// 1) Webhook冪等性：同一イベントの2回目はスキップされる（CacheService使用）
function pmTestWebhookIdempotency() {
  var ev = { type: 'message', webhookEventId: 'TEST-WHEID-' + generateId(), message: { id: 'm1', type: 'text' }, timestamp: 1 };
  var first  = isWebhookEventProcessed(ev);
  var second = isWebhookEventProcessed(ev);
  console.log((first === false ? '✅' : '❌') + ' 1回目=未処理(false期待): ' + first);
  console.log((second === true ? '✅' : '❌') + ' 2回目=処理済(true期待): ' + second);
  // フォールバックキー（webhookEventId無し）も確認
  var ev2 = { type: 'message', message: { id: 'MID-' + generateId(), type: 'text' }, timestamp: 2, source: { groupId: 'g1' } };
  console.log((isWebhookEventProcessed(ev2) === false ? '✅' : '❌') + ' fallback 1回目=false');
  console.log((isWebhookEventProcessed(ev2) === true ? '✅' : '❌') + ' fallback 2回目=true');
}

// 2) 進捗の確認待ち判定：低信頼度／ステータス不一致は hold
function pmTestProgressHold() {
  var cases = [
    [{ confidence: 95, status: '見積提出' }, null,  '高信頼度＋正規ステータス → 自動反映'],
    [{ confidence: 40, status: '見積提出' }, '理由', '低信頼度 → 確認待ち'],
    [{ confidence: 95, status: 'よくわからない状態' }, '理由', 'マスタ不一致 → 確認待ち'],
    [{ confidence: 95 }, null, 'ステータスなし高信頼度 → 自動反映'],
  ];
  cases.forEach(function(c) {
    var r = pmProgressHoldReason(c[0]);
    var ok = (c[1] === null) ? (r === null) : (!!r);
    console.log((ok ? '✅' : '❌') + ' ' + c[2] + ' → ' + (r || 'null'));
  });
}

// 3) 承認権限：PM_APPROVER_USER_IDS を一時設定し、許可/不許可を確認（最後に元へ戻す）
function pmTestApproverAuth() {
  var props = PropertiesService.getScriptProperties();
  var prev = props.getProperty('PM_APPROVER_USER_IDS');
  try {
    props.setProperty('PM_APPROVER_USER_IDS', 'Uallowed1, Uallowed2');
    console.log((pmIsApprover('Uallowed1') === true ? '✅' : '❌') + ' 登録ユーザーは承認可');
    console.log((pmIsApprover('Uunknown') === false ? '✅' : '❌') + ' 未登録ユーザーは承認不可（権限テスト本命）');
    console.log((pmIsFinancialType('billing') && pmIsFinancialType('amount') && pmIsFinancialType('payment') && !pmIsFinancialType('project_field')) ? '✅ financialType判定OK' : '❌ financialType判定NG');
  } finally {
    if (prev === null) props.deleteProperty('PM_APPROVER_USER_IDS');
    else props.setProperty('PM_APPROVER_USER_IDS', prev);
  }
}

// 4) スケジュール重複防止（読み取りのみ・シートに依存）。直近の同名同日があれば true。
function pmTestScheduleDedupSmoke() {
  var sheet = getSheet('スケジュール管理');
  if (!sheet || sheet.getLastRow() <= 1) { console.log('スケジュール管理シートが空のためスキップ'); return; }
  var row = sheet.getDataRange().getValues()[1];
  var title = row[2], date = row[3] instanceof Date ? fmtDate(row[3]) : String(row[3]).slice(0, 10), gid = row[9];
  console.log((isDuplicateSchedule(title, date, gid) === true ? '✅' : '❌') + ' 既存予定は重複検出される: ' + title + '/' + date);
  console.log((isDuplicateSchedule('絶対に存在しない予定XYZ', '2099-01-01', gid) === false ? '✅' : '❌') + ' 未登録予定は重複扱いされない');
}

// 5) 名称ゆるい一致＆グループ紐付け衝突の判定ロジック（純粋関数）
function pmTestNameMatching() {
  console.log((pmNamesLooselyEqual('〇〇ホテル客室改装', '〇〇ホテル') === true ? '✅' : '❌') + ' 部分一致=同一扱い');
  console.log((pmNamesLooselyEqual('Aビル', 'Bビル') === false ? '✅' : '❌') + ' 別案件=不一致');
}

// 追加修正テストを一括実行
function pmRunNewTests() {
  console.log('--- Webhook冪等性 ---');      pmTestWebhookIdempotency();
  console.log('--- 進捗hold判定 ---');        pmTestProgressHold();
  console.log('--- 承認権限 ---');            pmTestApproverAuth();
  console.log('--- 名称一致 ---');            pmTestNameMatching();
  console.log('--- スケジュール重複 ---');    pmTestScheduleDedupSmoke();
}

// ==========================================
// 結合(E2E)テスト：LINE送信なしで全書き込み経路を検証
//   テスト案件名で実データを作成 → 最後に pmE2ECleanup() で後始末
//   ※ Driveフォルダ・カレンダー予定は実際に作成されます（cleanupで削除）
// ==========================================
var PM_E2E_NAME = 'E2Eテスト案件_自動';

function pmE2ETest() {
  pmE2ECleanup(); // 前回分を掃除してから開始
  console.log('=== E2E開始: ' + PM_E2E_NAME + ' ===');

  // ① 新規案件
  var rec = pmEnsureProjectRecord(PM_E2E_NAME, { client: 'テスト商事', assignee: '濱田', groupId: 'TESTGROUP', updatedBy: 'tester' });
  var pid = rec['案件ID'];
  console.log('① 案件作成: 案件ID=' + pid + ' / クライアント=' + rec['クライアント名'] + ' / 担当=' + rec['担当者']);

  // ② Driveフォルダ（冪等確認）
  var url1 = pmEnsureProjectFolder(pid, PM_E2E_NAME);
  var url2 = pmEnsureProjectFolder(pid, PM_E2E_NAME);
  console.log('② Drive: ' + (url1 || '(DRIVE_FOLDER_ID未設定でスキップ)'));
  console.log('   冪等(2回目が同じURL): ' + (url1 === url2 ? '✅' : '❌'));

  // ③ 進捗更新（安全項目の自動反映）
  var due = fmtDate(new Date(Date.now() + 7 * 86400000));
  var parsed = { intent: 'project_update', phase: '営業', status: '見積提出', next_action: 'フォロー', next_action_due_date: due, assignee: '濱田' };
  var r = pmApplyUpdate(parsed, PM_E2E_NAME, 'tester', 'TESTGROUP', 'E2E 案件更新');
  console.log('③ 進捗反映: ' + JSON.stringify(r.applied));

  // ④ フォロータスク（重複防止）
  var t1 = pmCreateFollowupTask(pid, PM_E2E_NAME, parsed, 'tester', 'TESTGROUP');
  var t2 = pmCreateFollowupTask(pid, PM_E2E_NAME, parsed, 'tester', 'TESTGROUP');
  console.log('④ フォロータスク: 1回目=' + t1 + '(true期待) / 2回目=' + t2 + '(false=重複防止期待)');

  // ⑤ カレンダー同期（冪等）
  pmSyncCalendar(pid);
  pmSyncCalendar(pid);
  var calRows = pmReadObjects(PM_SHEET_CALENDAR).filter(function(x) { return String(x['案件ID']) === String(pid); });
  console.log('⑤ カレンダーイベント数: ' + calRows.length + '(重複作成なし=種別ごと1件ずつ期待)');

  // ⑥ 請求 確認待ち → 承認（LINE送信なし: replyToken空）
  var bdue = fmtDate(new Date(Date.now() + 10 * 86400000));
  var bparsed = { intent: 'billing_update', billing_due_date: bdue, billing_kind: '中間金', amount: { type: '請求', value: null, raw: null } };
  var qid = pmQueuePending('billing', bparsed, pid, PM_E2E_NAME, 'tester', 'TESTGROUP', 'E2E 請求更新');
  console.log('⑥ 確認待ち登録: status=' + pmGetPending(qid)['status'] + '(awaiting_approval期待)');
  pmApprovePending(qid, 'tester', '');
  var p1 = pmGetProjectById(pid);
  console.log('   承認後: 請求予定日=' + p1['請求予定日'] + ' / 請求ステータス=' + p1['請求ステータス'] + ' / pending=' + pmGetPending(qid)['status'] + '(applied期待)');

  // ⑦ 金額承認（売上・原価→利益自動計算）
  pmApprovePending(pmQueuePending('amount', { intent: 'project_update', amount: { type: '売上', value: 50000000, raw: '5000万' } }, pid, PM_E2E_NAME, 'tester', 'TESTGROUP', 'E2E売上'), 'tester', '');
  pmApprovePending(pmQueuePending('amount', { intent: 'project_update', amount: { type: '原価', value: 30000000, raw: '3000万' } }, pid, PM_E2E_NAME, 'tester', 'TESTGROUP', 'E2E原価'), 'tester', '');
  var p2 = pmGetProjectById(pid);
  console.log('⑦ 金額承認後: 売上=' + p2['売上'] + ' 原価=' + p2['原価'] + ' 利益額=' + p2['利益額'] + '(2000万期待) 利益率=' + p2['利益率'] + '(40%期待)');

  // ⑧ ログ確認
  var logs = pmReadObjects(PM_SHEET_LOGS).filter(function(x) { return String(x['案件ID']) === String(pid); });
  console.log('⑧ project_logs 件数: ' + logs.length + ' / 適用区分: ' + logs.map(function(l) { return l['適用区分']; }).join(','));

  console.log('=== E2E完了。確認後 pmE2ECleanup() で後始末してください ===');
}

// 放置/請求漏れ/入金未確認の「判定だけ」をドライラン（実通知なし）
//  テスト案件の請求予定日を過去にして請求漏れが立つか確認
function pmE2EAlertDryRun() {
  var rec = pmGetProjectByName(PM_E2E_NAME);
  if (!rec) { console.log('先に pmE2ETest() を実行してください'); return; }
  // 請求予定日を昨日に、最終更新日を20日前に書き換え（漏れ＆放置を発火させる）
  pmWriteRowFields(PM_SHEET_PROJECTS, rec._row, {
    '請求予定日': fmtDate(new Date(Date.now() - 1 * 86400000)),
    '請求済日': '',
    '請求ステータス': '未請求',
    '最終更新日': fmtDate(new Date(Date.now() - 20 * 86400000)),
    '営業ステータス': '見積提出（フォロー中）',
  });
  var p = pmGetProjectByName(PM_E2E_NAME);
  var today = pmParseDate(pmTodayYmd());
  var lastUpd = pmParseDate(p['最終更新日']);
  var billDue = pmParseDate(p['請求予定日']);
  console.log('放置判定(見積後7日以上): ' + (pmDaysBetween(lastUpd, today) >= PM_STALE_AFTER_QUOTE_DAYS ? '✅発火' : '❌'));
  console.log('請求漏れ判定: ' + (billDue && pmDaysBetween(billDue, today) > 0 && pmIsBlank(p['請求済日']) ? '✅発火' : '❌'));
  console.log('※実通知は pmCheckAlerts() を実行すると担当者DM＋社内/経理グループへ飛びます（テスト案件のみ反応）');
}

// E2Eテストの後始末（projects/logs/pending/calendar/タスク/プロジェクト管理/Driveを削除）
function pmE2ECleanup() {
  var rec = pmGetProjectByName(PM_E2E_NAME);
  var pid = rec ? rec['案件ID'] : null;
  var n = 0;

  // カレンダー実イベント削除
  if (pid) {
    try {
      var cfg = getConfig();
      var cal = cfg.CALENDAR_ID ? CalendarApp.getCalendarById(cfg.CALENDAR_ID) : null;
      pmReadObjects(PM_SHEET_CALENDAR).filter(function(x) { return String(x['案件ID']) === String(pid); })
        .forEach(function(x) { try { if (cal && x['イベントID']) { var ev = cal.getEventById(x['イベントID']); if (ev) ev.deleteEvent(); } } catch (e) {} });
    } catch (e) {}
  }

  // 各シートの行削除
  n += pmDeleteRowsWhere(PM_SHEET_PROJECTS, '案件名', PM_E2E_NAME);
  if (pid) {
    n += pmDeleteRowsWhere(PM_SHEET_LOGS, '案件ID', pid);
    n += pmDeleteRowsWhere(PM_SHEET_PENDING, '案件ID', pid);
    n += pmDeleteRowsWhere(PM_SHEET_CALENDAR, '案件ID', pid);
    n += pmDeleteRowsWhere('タスク管理', '関連案件ID', pid);
  }
  n += pmDeleteRowsWhere('プロジェクト管理', '正式名称', PM_E2E_NAME);

  // Driveフォルダをゴミ箱へ
  if (pid) {
    try {
      var cfg2 = getConfig();
      if (cfg2.DRIVE_FOLDER_ID) {
        var parent = getOrCreateFolder(cfg2.DRIVE_FOLDER_ID, PM_DRIVE_PARENT_NAME);
        if (parent) {
          var it = parent.getFoldersByName(pid + '_' + PM_E2E_NAME);
          while (it.hasNext()) { it.next().setTrashed(true); }
        }
      }
    } catch (e) {}
  }
  console.log('pmE2ECleanup: ' + n + '行を削除（テスト案件 ' + PM_E2E_NAME + '）');
}

// 指定列の値が一致する行を削除（ヘッダー名指定・下から走査）
function pmDeleteRowsWhere(sheetName, header, value) {
  var sheet = getSheet(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return 0;
  var idx = pmHeaderIndex(sheet);
  var col = idx[header];
  if (col === undefined) return 0;
  var data = sheet.getDataRange().getValues();
  var deleted = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][col]).trim() === String(value).trim()) { sheet.deleteRow(i + 1); deleted++; }
  }
  return deleted;
}

// ==========================================
// アプリ↔GAS ジョブキュー（pm_jobs.js）のテスト
// ==========================================

// 対象月パース・真偽判定・スタック判定（純ロジック・シート不要）
function pmTestJobsPureLogic() {
  var now = new Date(2026, 6, 7); // 2026-07-07
  var monthCases = [
    ['2026-05',    2026, 5],
    ['2026/5',     2026, 5],
    ['2026年5月',  2026, 5],
    ['2026.12',    2026, 12],
    ['当月',       2026, 7],
    ['先月',       2026, 6],
    ['2026-13',    null, null],
    ['5月',        null, null],
    ['',           null, null],
    ['abc',        null, null],
  ];
  monthCases.forEach(function(c) {
    var r = pmJobsParseTargetMonth(c[0], now);
    var ok = (c[1] === null) ? (r === null) : (r && r.year === c[1] && r.month === c[2]);
    console.log((ok ? '✅' : '❌') + ' 対象月 "' + c[0] + '" → ' + JSON.stringify(r));
  });

  var truthyCases = [[true, true], ['TRUE', true], ['1', true], ['はい', true], ['✓', true],
    [false, false], ['FALSE', false], ['', false], [null, false], ['いいえ', false]];
  truthyCases.forEach(function(c) {
    var got = pmJobsTruthy(c[0]);
    console.log((got === c[1] ? '✅' : '❌') + ' truthy(' + JSON.stringify(c[0]) + ') → ' + got);
  });

  var old = new Date(now.getTime() - PM_JOBS_STALE_MS - 60000);
  var recent = new Date(now.getTime() - 60000);
  console.log((pmJobsIsStale(old, now) === true ? '✅' : '❌') + ' 31分前開始 → stale');
  console.log((pmJobsIsStale(recent, now) === false ? '✅' : '❌') + ' 1分前開始 → not stale');
  console.log((pmJobsIsStale('', now) === true ? '✅' : '❌') + ' 開始不明 → stale（回収）');

  // 金額申請パラメータ（請求申請/入金申請ジョブ）: [入力, ok, amount, dueDate, doneDate]
  var amountCases = [
    ['{"金額":1000000,"予定日":"2026-08-01"}',        true, '1000000', '2026-08-01', ''],
    ['{"金額":"1,000,000円","確定日":"2026/8/5"}',    true, '1000000', '',           '2026/8/5'],
    ['金額=1000000; 予定日=2026-08-01',               true, '1000000', '2026-08-01', ''],
    ['金額：500000',                                  true, '500000',  '',           ''],
    ['確定日=2026-08-05',                             true, '',        '',           '2026-08-05'],
    ['',                                              true, '',        '',           ''],
    ['よくわからない文字列',                          false, null,     null,         null],
  ];
  amountCases.forEach(function(c) {
    var r = pmJobsParseAmountParams(c[0]);
    var ok = c[1]
      ? (r.ok && r.amount === c[2] && r.dueDate === c[3] && r.doneDate === c[4])
      : (r.ok === false);
    console.log((ok ? '✅' : '❌') + ' 金額パラメータ "' + c[0] + '" → ' + JSON.stringify(r));
  });
}

// セットアップの非破壊性：2回実行して列が増殖しないこと
function pmTestJobsSetupNonDestructive() {
  pmJobsSetupSheets();
  var b1 = getSheet(PM_SHEET_EST_JOBS).getLastColumn();
  var b2 = getSheet(PM_SHEET_PM_JOBS).getLastColumn();
  var b3 = getSheet(PM_SHEET_APP_ROLES).getLastColumn();
  pmJobsSetupSheets();
  console.log((getSheet(PM_SHEET_EST_JOBS).getLastColumn() === b1 ? '✅' : '❌') + ' 見積ジョブ列数不変: ' + b1);
  console.log((getSheet(PM_SHEET_PM_JOBS).getLastColumn() === b2 ? '✅' : '❌') + ' pmジョブ列数不変: ' + b2);
  console.log((getSheet(PM_SHEET_APP_ROLES).getLastColumn() === b3 ? '✅' : '❌') + ' アプリ権限列数不変: ' + b3);
}

// 権限判定：アプリ権限シートの状態に応じた許可/拒否（読み取りのみ）
function pmTestJobsRoles() {
  var n = appRoleRowCount();
  console.log('アプリ権限 登録行数: ' + n + (n === 0 ? '（0件=制限なしモード）' : '（登録メールのみ許可）'));
  if (n === 0) {
    console.log((appCanRunJobs('anyone@example.com') === true ? '✅' : '❌') + ' 未運用時: 実行可');
    console.log((appCanApproveFinancial('anyone@example.com') === true ? '✅' : '❌') + ' 未運用時: 承認可');
  } else {
    console.log((appCanRunJobs('not-registered@example.com') === false ? '✅' : '❌') + ' 未登録メール: 実行不可');
    console.log((appCanApproveFinancial('not-registered@example.com') === false ? '✅' : '❌') + ' 未登録メール: 承認不可');
    pmReadObjects(PM_SHEET_APP_ROLES).forEach(function(r) {
      if (pmIsBlank(r['メール'])) return;
      console.log('  ' + r['メール'] + ' 役割=' + r['役割'] +
        ' 実行=' + appCanRunJobs(r['メール']) + ' 金額承認=' + appCanApproveFinancial(r['メール']));
    });
  }
}

// E2E（書き込みあり・要ジョブシート）：不正ジョブ1件を投入→runAppJobs→エラー書き戻しを確認→行削除
function pmTestJobsE2EInvalidJob() {
  pmJobsSetupSheets();
  var jobId = 'TEST-JOB-' + generateId();
  pmAppendRowFields(PM_SHEET_EST_JOBS, {
    'ジョブID': jobId, '依頼日時': fmtDT(new Date()), '依頼者': 'test@example.com',
    '設定ID': '存在しない設定ID', '対象月': '2026-05', '状態': PM_JOB_STATE_RECEIVED,
  });
  runEstimateJobs();
  var row = null;
  pmReadObjects(PM_SHEET_EST_JOBS).forEach(function(r) { if (String(r['ジョブID']) === jobId) row = r; });
  var ok = row && String(row['状態']).trim() === PM_JOB_STATE_ERROR && String(row['メッセージ']).indexOf('見つからない') !== -1;
  console.log((ok ? '✅' : '❌') + ' 不正設定IDジョブ → エラー書き戻し: ' + (row ? row['状態'] + ' / ' + row['メッセージ'] : '行なし'));
  console.log('削除: ' + pmDeleteRowsWhere(PM_SHEET_EST_JOBS, 'ジョブID', jobId) + '行');
}
