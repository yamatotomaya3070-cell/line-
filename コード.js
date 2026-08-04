// ==========================================
// WOODBASE 専属秘書AI v3.1
// 設計原則：ユーザーに投稿ルールを強制しない
//   - フリーテキストから抽出
//   - 案件は信頼度スコアで識別（フォーマット不要）
//   - 確認は必須・選択UIで完結（入力ゼロ）
// ==========================================

// ===== CONFIG =====
function getConfig() {
  var p = PropertiesService.getScriptProperties();
  return {
    LINE_CHANNEL_ACCESS_TOKEN: p.getProperty('LINE_CHANNEL_ACCESS_TOKEN'),
    LINE_CHANNEL_SECRET:       p.getProperty('LINE_CHANNEL_SECRET'),
    GEMINI_API_KEY:            p.getProperty('GEMINI_API_KEY'),
    SPREADSHEET_ID:            p.getProperty('SPREADSHEET_ID'),
    INTERNAL_GROUP_ID:         p.getProperty('INTERNAL_GROUP_ID'),
    DRIVE_FOLDER_ID:           p.getProperty('DRIVE_FOLDER_ID'),
    TEMPLATE_FOLDER_ID:        p.getProperty('TEMPLATE_FOLDER_ID'),
    CALENDAR_ID:               p.getProperty('CALENDAR_ID'),
  };
}

function getSS()        { return SpreadsheetApp.openById(getConfig().SPREADSHEET_ID); }
function getSheet(name) { return getSS().getSheetByName(name); }
function okRes()        { return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT); }

// ==========================================
// SECTION 1: Webhook メイン
// ==========================================

// LINE Channel Secret を Script Properties から取得
function getLineChannelSecret() {
  return PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_SECRET');
}

// HMAC-SHA256 と Base64 を使った安全な等価判定（タイミング攻撃耐性）
function timingSafeEqual(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string') return false;
  if (a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

// LINE Webhook 署名検証
// 注意：GAS doPost(e) では HTTPヘッダー（X-Line-Signature）に直接アクセスできない。
// そのため以下の二段構えで実装する：
//   ① URLクエリ ?sig=<署名> または ?token=<シークレット> による検証（推奨）
//      → LINEコンソールのWebhook URLに ?token=xxx を付けて運用する
//   ② LINE_CHANNEL_SECRET が未設定の場合は無検証（既存運用との互換性のため）
//
// より堅牢にしたい場合は Cloudflare Worker 等の前段プロキシで X-Line-Signature を
// 検証し、検証済みリクエストにのみ ?token= を付けて GAS に転送する構成を推奨。
function verifyLineWebhook(e) {
  var secret = getLineChannelSecret();
  // シークレット未設定 → 通過
  if (!secret) return true;

  var providedToken = e && e.parameter && (e.parameter.token || e.parameter.t);
  var providedSig   = e && e.parameter && e.parameter.sig;

  // トークンも署名もURLに無い → 通過（警告のみ）
  // ※LINE WebhookはデフォルトでURLパラメータを付けないため、未設定運用との互換性を維持
  if (!providedToken && !providedSig) {
    console.warn('LINE Webhook: URLにtoken/sig無し → 通過（推奨: Webhook URL末尾に ?token=<シークレット> を追加）');
    return true;
  }

  // トークンが渡されている → 厳密照合
  if (providedToken) {
    if (timingSafeEqual(String(providedToken), secret)) return true;
    console.error('LINE Webhook 署名検証失敗: token不一致');
    return false;
  }

  // 署名が渡されている → HMAC照合
  if (providedSig && e.postData && e.postData.contents) {
    try {
      var raw = Utilities.computeHmacSha256Signature(e.postData.contents, secret);
      var expected = Utilities.base64Encode(raw);
      if (timingSafeEqual(String(providedSig), expected)) return true;
      console.error('LINE Webhook 署名検証失敗: sig不一致');
    } catch (err) { console.error('HMAC計算エラー:', err.message); }
  }
  return false;
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) return okRes();
    // 署名検証（失敗時は処理せず終了）
    if (!verifyLineWebhook(e)) return okRes();
    var events = JSON.parse(e.postData.contents).events || [];
    for (var i = 0; i < events.length; i++) {
      // 冪等性：LINEのWebhook再送（同一イベントの複数配信）で二重処理しない
      if (isWebhookEventProcessed(events[i])) {
        console.log('Webhook再送をスキップ:', webhookEventKey(events[i]));
        continue;
      }
      try { handleEvent(events[i]); } catch (err) { console.error('event error:', err.message); }
    }
  } catch (err) { console.error('doPost error:', err.message); }
  return okRes();
}

// Webhookイベントの一意キー（再送判定用）。webhookEventId 優先、無ければ種別＋メッセージID＋タイムスタンプ。
function webhookEventKey(ev) {
  if (!ev) return '';
  if (ev.webhookEventId) return 'we:' + ev.webhookEventId;
  var parts = [
    ev.type || '',
    (ev.message && ev.message.id) || '',
    (ev.postback && ev.postback.data) || '',
    ev.replyToken || '',
    ev.timestamp || '',
    (ev.source && (ev.source.userId || ev.source.groupId || ev.source.roomId)) || '',
  ];
  return 'ev:' + parts.join('|');
}

// 既に処理済みのWebhookイベントか判定し、未処理なら「処理済み」として記録する（check-and-set）。
//   - CacheService（最大6時間）で記録。LINEの再送は通常数分〜数十分以内。
//   - 同時並走（同一イベントの同時2配信）対策に LockService で直列化。
function isWebhookEventProcessed(ev) {
  var key = webhookEventKey(ev);
  if (!key) return false; // キーが取れない場合は従来どおり処理（取りこぼし防止）
  var cache = CacheService.getScriptCache();
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    try { lock.waitLock(5000); locked = true; } catch (e) { locked = false; }
    if (cache.get(key)) return true;          // 既に処理済み
    cache.put(key, '1', 21600);               // 6時間保持（再送ウィンドウを十分カバー）
    return false;
  } catch (err) {
    console.error('isWebhookEventProcessed error:', err.message);
    return false; // 判定不能時は処理を継続（重複より取りこぼしを避ける）
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}

// 単体テスト：HMAC計算が期待通り動くか確認
function testVerifyLineSignature() {
  var secret = 'test_secret_value';
  var body   = '{"events":[]}';
  var raw    = Utilities.computeHmacSha256Signature(body, secret);
  var sig    = Utilities.base64Encode(raw);
  console.log('期待される署名:', sig);
  // 一致テスト
  console.log('timingSafeEqual同値:',  timingSafeEqual(sig, sig)        === true ? '✅' : '❌');
  console.log('timingSafeEqual不一致:', timingSafeEqual(sig, sig + 'x') === false ? '✅' : '❌');
  console.log('timingSafeEqual長さ違い:', timingSafeEqual(sig, '')      === false ? '✅' : '❌');
}

function handleEvent(ev) {
  var userId  = ev.source.userId;
  var groupId = ev.source.groupId || ev.source.roomId || userId;
  var ts      = new Date(ev.timestamp);
  var isGroup = ev.source.type === 'group' || ev.source.type === 'room';

  // ボタン押下（Quick Reply）
  if (ev.type === 'postback') { handlePostback(ev, userId, groupId); return; }

  // 友達追加 → メンバー自動登録
  if (ev.type === 'follow') {
    if (userId) registerMember(userId);
    return;
  }

  // ボットがグループに参加 → 案件選択UIを表示
  if (ev.type === 'join') {
    handleBotJoinGroup(ev.replyToken, groupId);
    return;
  }

  // グループメンバー参加 → メンバー自動登録
  if (ev.type === 'memberJoined') {
    var members = ev.joined && ev.joined.members || [];
    members.forEach(function(m) { if (m.userId) registerMember(m.userId, groupId); });
    return;
  }

  if (ev.type !== 'message') return;

  // ファイル・画像・動画
  if (['image', 'file', 'video'].indexOf(ev.message.type) !== -1) {
    saveFileToDrive(ev.message.id, ev.message.fileName || ev.message.type, groupId, ts, ev.message.type);
    return;
  }
  if (ev.message.type !== 'text') return;

  var text   = ev.message.text;
  var sender = getMemberNameByUserId(userId);
  if (!sender) { registerMember(userId, groupId); sender = getMemberNameByUserId(userId) || '不明'; }

  saveMessageLog(groupId, sender, text, ts);

  // グループ紐付け用 新規プロジェクト名入力モード
  var linkGroupId = getNewProjectLinkMode(userId);
  if (linkGroupId) {
    clearNewProjectLinkMode(userId);
    finalizeNewProjectLink(text, linkGroupId, userId, ev.replyToken);
    return;
  }

  // 新規プロジェクト名入力モードのチェック（修正モードより優先）
  var newProjBatchId = getNewProjectMode(userId);
  if (newProjBatchId) {
    clearNewProjectMode(userId);
    finalizeNewProject(text, newProjBatchId, groupId, userId, ev.replyToken);
    return;
  }

  // 修正（再入力）モードのチェック
  var retryBatchId = getRetryMode(userId);
  if (retryBatchId) {
    clearRetryMode(userId);
    reprocessMessage(text, retryBatchId, groupId, userId, sender, ts, ev.replyToken);
    return;
  }

  if (!isGroup) { handleDM(ev, userId, groupId, sender, text, ts); return; }
  handleGroup(ev, groupId, userId, sender, text, ts);
}

// ==========================================
// SECTION 2: 1対1メッセージ
// ==========================================
function handleDM(ev, userId, groupId, sender, text, ts) {
  var justRegistered = registerMember(userId);
  if (handleGoogleCalendarConnectRequest(ev, userId, text)) return;
  if (justRegistered) {
    sendLineReply(ev.replyToken, 'WOODBASE秘書AIにご登録いただきました。\nタスクが割り当てられた際にはお知らせいたします。\n\n「残タスクは？」「今週の予定は？」とお送りいただくとご確認いただけます。');
    return;
  }
  // 見積設定一覧コマンド
  if (handleEstimateSettingsListRequest(ev, text)) return;
  // パイオニア什器 見積書作成（制作一覧→店舗別シート。東邦家具判定より先に）
  if (handleEstimateCommand(ev, text)) return;
  // 見積書作成リクエスト（進捗質問より先に処理）
  if (handleEstimateCreateRequest(ev, text)) return;
  // 進捗管理表への自然文クエリ（完了報告誤判定より先に処理）
  if (handleProgressQuestionRequest(ev, text)) return;
  // プロジェクト進捗管理（案件更新/新規案件/請求更新/入金更新 等のプレフィックス）
  if (pmRouteText(ev, text, sender, userId, groupId)) return;
  // 「URL」等 → 進捗ボード/請求・入金/タスク の各URLを返す
  if (isDashboardUrlRequest(text)) { handleDashboardUrlRequest(ev.replyToken, sender); return; }
  // 残タスク一覧（タップで完了できるボタン付き）
  if (isTaskListRequest(text)) { handleTaskListWithButtons(ev.replyToken, sender); return; }
  if (isCompletionReport(text)) { handleCompletion(text, sender, ev.replyToken); return; }
  if (isSummaryRequest(text)) { sendLineReply(ev.replyToken, '【会話まとめ】\n' + summarizeChat(userId)); return; }
  // ボールホルダー照会
  if (isBallHolderQuery(text)) { sendLineReply(ev.replyToken, answerBallHolderQuery(text, groupId)); return; }
  // 案件メモ 照会
  if (isMemoQuery(text)) { sendLineReply(ev.replyToken, answerMemoQuery(text, groupId)); return; }
  // 案件メモ 保存
  if (isMemoRequest(text)) {
    var dmMemoReply = handleMemoSave(text, groupId, sender);
    if (dmMemoReply) { sendLineReply(ev.replyToken, dmMemoReply); return; }
  }
  // 住所・会社名の更新（例: 「雨晴れの住所：東京都〇〇」）
  var dmFieldUpd = parseProjectFieldUpdate(text);
  if (dmFieldUpd) {
    var dmUpdName = updateProjectField(dmFieldUpd.projectQuery, dmFieldUpd.field, dmFieldUpd.value);
    if (dmUpdName) {
      sendLineReply(ev.replyToken, '✅ ' + dmUpdName + 'の' + dmFieldUpd.field + 'を登録しました。\n' + dmFieldUpd.value);
      return;
    }
  }
  // 住所・会社名の照会（例: 「雨晴れの住所は？」）
  if (isAddressOrCompanyQuery(text)) {
    var dmAddrAns = answerAddressOrCompanyQuery(text);
    if (dmAddrAns) { sendLineReply(ev.replyToken, dmAddrAns); return; }
  }
  sendLineReply(ev.replyToken, answerQueryForMember(text, sender));
}

// ==========================================
// SECTION 3: グループメッセージ
//
// 設計原則：グループは静かに保つ（水面下で動く）
//   - 通常メッセージ → タスク・予定を自動検出。検出したら送信者のDMで確認（グループには返信しない）
//   - @WBG メンション → handleMentionCommand → 照会・コマンドに応答（タスク登録はしない）
// ==========================================
function handleGroup(ev, groupId, userId, sender, text, ts) {
  var mention      = ev.message.mention;
  var botMentioned = isBotMentioned(mention);

  // テキストプレフィックスでもメンション扱い（PC版LINEで@候補が出ない場合の回避策）
  // 「@WBG」「＠WBG」「/WBG」「/wbg」「秘書:」「秘書：」を文頭に書けば反応
  var textPrefixMatch = text && text.match(/^\s*(@WBG|＠WBG|\/WBG|\/wbg|秘書[:：])\s*/i);
  if (textPrefixMatch) {
    botMentioned = true;
    // プレフィックス部分を除去して後続処理に渡す（既存のテキスト解析を壊さないため）
    ev.message.text = text.slice(textPrefixMatch[0].length);
    text = ev.message.text;
  }

  // プロジェクト進捗管理：明示プレフィックス（案件更新/新規案件/請求更新 等）は
  // メンション不要で反応（pmClassifyIntentが非該当なら即false＝既存挙動を維持）
  if (pmRouteText(ev, text, sender, userId, groupId)) return;

  // @WBG メンション → 照会・コマンドに応答（タスク登録はしない）
  if (botMentioned) {
    handleMentionCommand(ev, groupId, userId, sender, text, ts);
    return;
  }

  // 通常メッセージ → 水面下でタスク・予定を自動検出し、即登録（グループには返信しない）
  detectTasksSilently(ev, groupId, userId, sender, text, ts);
}

// 通常メッセージからタスク・予定を自動検出し、即登録する（グループはサイレント）
//   - メンション不要。タスク・予定らしき文面が出たら反応する＝「初めの」挙動
//   - 確認を挟まず即登録。タスク・予定とも同じ流れで登録する
//   - @メンションされた人は担当者／参加者に加え、DM通知・カレンダー招待で予定を反映する
//   - グループには一切返信しない＝水面下で動く
function detectTasksSilently(ev, groupId, userId, sender, text, ts) {
  if (!userId) return;                       // 送信元が不明なら何もしない
  if (shouldSkipExtraction(text)) return;    // 短文・絵文字のみ等はGemini呼び出しをスキップ

  var extracted = extractWithGemini(text, groupId, ts, sender);
  if (!extracted || (extracted.tasks.length + extracted.schedules.length === 0)) return;

  // @メンションされた人 → タスク担当者・予定参加者に追加（DM通知・カレンダー反映用）
  var mentionedNames = getMentionedMemberNames(ev);
  var mentionedIds   = getMentionedMemberIds(ev);
  if (mentionedNames.length || mentionedIds.length) {
    extracted.tasks.forEach(function(t) {
      t.assignee  = mergeNames(t.assignee, mentionedNames);
      t.assignee_user_ids = mergeUserIds(t.assignee_user_ids, mentionedIds);
    });
    extracted.schedules.forEach(function(s) {
      s.attendees = mergeNames(s.attendees, mentionedNames);
      s.attendees_user_ids = mergeUserIds(s.attendees_user_ids, mentionedIds);
    });
  }

  // 保存先プロジェクト：紐付け済み > 高確度の既存案件 > 未分類（誤った新規案件の自動生成は避ける）
  var projectName = resolveProjectForAutoCommit(groupId, text, extracted.projectNameHint);

  // 確認を挟まず即登録（commitDirectly がタスク・予定の両方を処理）
  commitDirectly(extracted, projectName, groupId, userId);
}

// 自動登録用の保存先プロジェクトを決定する
//   - グループ紐付け済み → その案件
//   - 未紐付けでも本文/ヒントが既存案件に高確度一致 → その案件
//   - それ以外 → 未分類（存在しない案件名での新規自動生成を防ぐ）
function resolveProjectForAutoCommit(groupId, text, hint) {
  var bound = getProjectNameByGroupId(groupId);
  if (bound) return bound;
  var cands = getProjectCandidates(text, groupId, hint, 1);
  if (cands.length && cands[0].confidence >= 70) return cands[0].name;
  return '未分類';
}

// メッセージで@メンションされたメンバー名一覧（Bot自身は除外）
function getMentionedMemberNames(ev) {
  var names = [];
  var mention = ev && ev.message && ev.message.mention;
  if (!mention || !mention.mentionees || !mention.mentionees.length) return names;
  var botId = getCachedBotUserId();
  mention.mentionees.forEach(function(m) {
    if (m.isSelf === true) return;
    if (botId && m.userId === botId) return;
    if (!m.userId) return;
    var name = getMemberNameByUserId(m.userId);
    if (name && names.indexOf(name) === -1) names.push(name);
  });
  return names;
}

function getMentionedMemberIds(ev) {
  var ids = [];
  var mention = ev && ev.message && ev.message.mention;
  if (!mention || !mention.mentionees || !mention.mentionees.length) return ids;
  var botId = getCachedBotUserId();
  mention.mentionees.forEach(function(m) {
    if (m.isSelf === true) return;
    if (botId && m.userId === botId) return;
    if (!m.userId) return;
    if (ids.indexOf(m.userId) === -1) ids.push(m.userId);
  });
  return ids;
}

function mergeUserIds(existing, additions) {
  if (!Array.isArray(existing)) existing = [];
  if (!Array.isArray(additions)) additions = [];
  var seen = {};
  existing.forEach(function(id) { if (id) seen[id] = true; });
  additions.forEach(function(id) { if (id && !seen[id]) { seen[id] = true; existing.push(id); } });
  return existing;
}

// 既存の担当者/参加者文字列に、追加メンバー名を重複なく結合（「、」区切り）
function mergeNames(existing, additions) {
  var out = [];
  splitMemberNames(existing).forEach(function(n) { if (out.indexOf(n) === -1) out.push(n); });
  (additions || []).forEach(function(n) {
    var nn = normalizeMemberName(n);
    if (nn && out.indexOf(nn) === -1) out.push(nn);
  });
  return out.join('、');
}

// Bot宛メンション時のコマンド処理（グループへの返信あり）
function handleMentionCommand(ev, groupId, userId, sender, text, ts) {
  // ⓪ 「ここに進捗アラート」→ このグループを放置アラート通知先に登録（進捗質問より先に）
  if (isSetProgressAlertHereRequest(text)) {
    handleSetProgressAlertHere(ev.replyToken, groupId);
    return;
  }
  // ⓪-b 既存グループでも案件の紐付けを見直せるようにする。
  //       ボット参加時と異なり、すでに紐付け済みでも選択UIを表示する。
  if (isProjectLinkRequest(text)) {
    sendProjectLinkUI(ev.replyToken, groupId, true);
    return;
  }
  // ① Docsサマリー更新（優先判定：「サマリー更新」「Docsまとめて」など）
  if (isDocSummaryRequest(text)) {
    handleDocSummaryRequest(ev.replyToken, text, groupId, sender);
    return;
  }
  // ② 会話まとめ（「まとめて」「要約して」など）
  if (isSummaryRequest(text)) {
    sendLineReply(ev.replyToken, '【会話まとめ】\n' + summarizeChat(groupId));
    return;
  }
  // ② パイオニア什器 見積書作成（東邦家具判定より先に）
  if (handleEstimateCommand(ev, text)) return;
  // ② 見積書作成リクエスト（進捗質問より先に処理）
  if (handleEstimateCreateRequest(ev, text)) return;
  // ③ 進捗管理表への自然文クエリ（完了報告誤判定より先に処理）
  if (handleProgressQuestionRequest(ev, text)) return;
  // ④ 完了報告
  if (isCompletionReport(text)) {
    handleCompletion(text, sender, ev.replyToken);
    return;
  }
  // ⑤ 仮タスク確認
  if (text.includes('仮タスク')) {
    handlePendingList(ev.replyToken, groupId);
    return;
  }
  // ⑥ ボールホルダー照会（例: 「誰がボール持ってる？」「今誰待ち？」）
  if (isBallHolderQuery(text)) {
    sendLineReply(ev.replyToken, answerBallHolderQuery(text, groupId));
    return;
  }
  // ⑦ 案件メモ 照会（例: 「メモは？」「特記事項教えて」）
  if (isMemoQuery(text)) {
    sendLineReply(ev.replyToken, answerMemoQuery(text, groupId));
    return;
  }
  // ⑧ 案件メモ 保存（例: 「メモして：現場は土日NG」「重要：予算上限〇〇万」）
  if (isMemoRequest(text)) {
    var memoReply = handleMemoSave(text, groupId, sender);
    if (memoReply) { sendLineReply(ev.replyToken, memoReply); return; }
  }
  // ④-b 住所・会社名 更新（例: 「雨晴れの住所：東京都〇〇」）
  var grpFieldUpd = parseProjectFieldUpdate(text);
  if (grpFieldUpd) {
    var grpUpdName = updateProjectField(grpFieldUpd.projectQuery, grpFieldUpd.field, grpFieldUpd.value);
    if (grpUpdName) {
      sendLineReply(ev.replyToken, '✅ ' + grpUpdName + 'の' + grpFieldUpd.field + 'を登録しました。\n' + grpFieldUpd.value);
      return;
    }
  }
  // ④-c 住所・会社名 照会（例: 「雨晴れの住所は？」）
  if (isAddressOrCompanyQuery(text)) {
    var grpAddrAns = answerAddressOrCompanyQuery(text);
    if (grpAddrAns) { sendLineReply(ev.replyToken, grpAddrAns); return; }
  }
  // ④-d' 残タスク一覧（タップで完了できるボタン付き・全体表示）
  if (isTaskListRequest(text)) {
    handleTaskListWithButtons(ev.replyToken, null);
    return;
  }
  // ④-d'' 「URL」等 → 進捗ボード/請求・入金/タスク の各URLを返す
  if (isDashboardUrlRequest(text)) {
    handleDashboardUrlRequest(ev.replyToken, sender);
    return;
  }
  // ④-d''' 「ここに進捗アラート」→ このグループを放置アラートの通知先に登録
  if (isSetProgressAlertHereRequest(text)) {
    handleSetProgressAlertHere(ev.replyToken, groupId);
    return;
  }
  // ④-d タスク・スケジュール照会
  if (isQuery(text)) {
    sendLineReply(ev.replyToken, answerQuery(text));
    return;
  }
  // ⑤ タスク・スケジュール登録はメンションでは行わない
  //    （通常メッセージから detectTasksSilently が水面下で自動検出する）
  // ⑥ 何も該当しない → Geminiに会話履歴を渡して文脈に沿った返答を生成
  sendLineReply(ev.replyToken, answerWithContext(text, groupId, sender));
}

// グループ会話履歴・タスク状況を踏まえて自由に応答する（コマンド非該当時のフォールバック）
function answerWithContext(question, groupId, sender) {
  var config = getConfig();
  var today  = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日');

  // 直近のグループ会話履歴（最大30件）
  var logSheet = getSheet('メッセージログ');
  var historyText = '（会話履歴なし）';
  if (logSheet && logSheet.getLastRow() > 1) {
    var rows = logSheet.getDataRange().getValues().slice(1)
      .filter(function(r) { return r[1] === groupId; })
      .slice(-30);
    if (rows.length) {
      historyText = rows.map(function(r) {
        return '[' + r[0] + '] ' + r[2] + '：' + r[3];
      }).join('\n');
    }
  }

  // このグループに紐付く案件名
  var projectName = getProjectNameByGroupId(groupId) || '';
  var taskList    = projectName ? buildProjectTaskListText(projectName) : '';
  var schedList   = projectName ? buildProjectSchedListText(projectName) : '';

  var prompt =
    'あなたはWOODBASE・Fの専属秘書AIです。今日：' + today + '\n' +
    '相手は「' + sender + '」さんで、グループ「' + (projectName || '案件未設定') + '」での発言です。\n' +
    '\n【あなたの役割】\n' +
    '・グループの会話の流れを踏まえて、自然で具体的な返答をすること。\n' +
    '・必ず丁寧な敬語（です・ます調）を使うこと。フランクな口調や「〜だよ」「〜だね」などは絶対に使わない。\n' +
    '・「コマンド一覧」のような形式的な定型文は出さないこと。\n' +
    '・回答は3〜4文以内で簡潔に。\n' +
    '・状況がわからない場合は、何を確認すべきか具体的に問い返すこと。\n' +
    '\n【このグループの直近会話】\n' + historyText +
    (taskList  ? '\n\n【案件「' + projectName + '」の未完了タスク】\n' + taskList  : '') +
    (schedList ? '\n\n【案件「' + projectName + '」の直近の予定】\n' + schedList : '') +
    '\n\n' + sender + 'さんからのメッセージ：\n' + question;

  return callGemini(config.GEMINI_API_KEY, prompt, 0.4) ||
    '申し訳ございません。うまくお答えできませんでした。';
}

// 案件単位のタスク一覧テキスト
function buildProjectTaskListText(projectName) {
  var sheet = getSheet('タスク管理');
  if (!sheet || sheet.getLastRow() <= 1) return '';
  return sheet.getDataRange().getValues().slice(1)
    .filter(function(r) {
      return r[1] === projectName && String(r[5] || '') !== 'done' && String(r[5] || '') !== '完了';
    })
    .map(function(r) {
      var dl = r[4] ? Utilities.formatDate(new Date(r[4]), 'Asia/Tokyo', 'M/d') : '期日未定';
      return '・[' + (r[3] || '担当未定') + '] ' + r[2] + '（' + dl + '）';
    }).join('\n');
}

// 案件単位の直近スケジュール一覧テキスト
function buildProjectSchedListText(projectName) {
  var sheet = getSheet('スケジュール管理');
  if (!sheet || sheet.getLastRow() <= 1) return '';
  var todayYmd = fmtDate(new Date());
  return sheet.getDataRange().getValues().slice(1)
    .filter(function(r) {
      var d = r[3] instanceof Date ? fmtDate(r[3]) : String(r[3]).slice(0, 10);
      return r[1] === projectName && d >= todayYmd;
    })
    .slice(0, 10)
    .map(function(r) {
      var d = r[3] instanceof Date ? fmtDate(r[3]) : String(r[3]).slice(0, 10);
      return '・' + d.slice(5).replace('-', '/') + ' ' + r[2] + (r[4] ? ' ' + r[4] + '〜' : '');
    }).join('\n');
}

// Bot自身へのメンションか判定（isSelf + botUserId の二重チェック）
function isBotMentioned(mention) {
  if (!mention || !mention.mentionees || !mention.mentionees.length) return false;
  var botId = getCachedBotUserId();
  return mention.mentionees.some(function(m) {
    return m.isSelf === true || (botId && m.userId === botId);
  });
}

// ボット自身のuserIdをキャッシュ取得
function getCachedBotUserId() {
  var props  = PropertiesService.getScriptProperties();
  var cached = props.getProperty('BOT_USER_ID');
  if (cached) return cached;
  try {
    var config = getConfig();
    var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/info', {
      headers: { 'Authorization': 'Bearer ' + config.LINE_CHANNEL_ACCESS_TOKEN },
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() === 200) {
      var info = JSON.parse(res.getContentText());
      if (info.userId) { props.setProperty('BOT_USER_ID', info.userId); return info.userId; }
    }
  } catch(e) { console.error('getCachedBotUserId error:', e.message); }
  return null;
}

// グループのisSummaryRequest（メンション必須チェックを削除・handleMentionCommandで保証済み）
function isSummaryRequest(text) {
  return ['まとめて', '箇条書き', '会話内容', '要約して', '議事録'].some(function(k) { return text.includes(k); });
}

// ==========================================
// SECTION 4: ルールベースフィルタ
// フォーマット不要・自然文から意図を検出
// ==========================================
function ruleBasedFilter(text) {
  // タスクを示す自然な動詞・表現（フォーマット依存なし）
  var TASK_KW = [
    'お願い', 'やります', 'やっとく', 'やっておく', 'します', 'しておく', 'しとく',
    '対応', '提出', '送って', '送ります', '用意', '手配', '依頼', '作成', '確認して',
    '連絡して', 'てください', 'ておいて', 'とく', 'やる', '担当', '頼む', '頼んで',
  ];
  // スケジュールを示す自然な表現
  var SCHED_KW = [
    '打ち合わせ', '打合せ', 'ミーティング', '会議', '現場', '検査', '引渡',
    '引き渡し', '訪問', 'アポ', '来週', '今週', '明日', '明後日',
    '時から', '時に', '時〜', '時ごろ', '予定', 'スケジュール',
  ];
  return TASK_KW.some(function(k) { return text.includes(k); }) ||
    SCHED_KW.some(function(k) { return text.includes(k); });
}

// 日付・担当者ヒントのみルールで抽出（案件名はidentifyProjectが担当）
function extractHints(text) {
  var hints = {};

  // @メンション → 担当者候補
  var atMatch = text.match(/@([^\s　、。,，\n]+)/);
  if (atMatch) hints.assignee = atMatch[1];

  // 日付表現をルールで高精度に変換（今日の日付ベース）
  var today = new Date();
  var dateRules = [
    [/(\d{1,2})月(\d{1,2})日/, function(m) { return today.getFullYear() + '-' + pad2(m[1]) + '-' + pad2(m[2]); }],
    [/(\d{1,2})\/(\d{1,2})/,   function(m) { return today.getFullYear() + '-' + pad2(m[1]) + '-' + pad2(m[2]); }],
    [/明日/,    function() { return fmtDate(new Date(today.getTime() + 86400000)); }],
    [/明後日/,  function() { return fmtDate(new Date(today.getTime() + 172800000)); }],
    [/今週中/,  function() { return fmtDate(getWeekEnd(today)); }],
    [/来週/,    function() { return fmtDate(new Date(today.getTime() + 7 * 86400000)); }],
    [/再来週/,  function() { return fmtDate(new Date(today.getTime() + 14 * 86400000)); }],
  ];
  for (var i = 0; i < dateRules.length; i++) {
    var m = text.match(dateRules[i][0]);
    if (m) { hints.dueDate = dateRules[i][1](m); break; }
  }

  return hints;
}

function detectUrgency(text) {
  var URGENT_KW = ['至急', '急ぎ', '急いで', '急いで', 'ASAP', 'asap', '今すぐ', '今日中', '最短', '緊急', 'GW前', 'GW中', '明日まで', '本日中'];
  return URGENT_KW.some(function(k) { return text.includes(k); });
}

function pad2(s)   { return String(s).padStart(2, '0'); }
function fmtDate(d){ return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd'); }
function fmtDT(d)  { return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'); }
function getWeekEnd(d) { var r = new Date(d), day = r.getDay(); r.setDate(r.getDate() + (day === 0 ? 0 : 7 - day)); return r; }

// ==========================================
// SECTION 5: Gemini抽出
// プロンプトはフリーテキスト前提・フォーマット不要
// ==========================================
function extractWithGemini(message, groupId, timestamp, senderName) {
  var config = getConfig();
  var today  = Utilities.formatDate(timestamp, 'Asia/Tokyo', 'yyyy年MM月dd日');
  var hints  = extractHints(message);

  var hintLines = [];
  if (hints.assignee) hintLines.push('担当者ヒント（@メンションから）：' + hints.assignee);
  if (hints.dueDate)  hintLines.push('期日ヒント（日付表現から）：' + hints.dueDate);
  var hintBlock = hintLines.length > 0 ? '\n【ルール抽出ヒント（高信頼度・優先使用）】\n' + hintLines.join('\n') + '\n' : '';

  // プロンプトはフリーテキスト・口語前提
  // 案件名はここでは抽出しない（identifyProjectが担当）
  // projectNameHintだけ補助的に取得する
  var prompt =
    'あなたは建築会社WOODBASEの専属秘書AIです。\n' +
    '今日：' + today + '　送信者：' + senderName + '\n' +
    hintBlock +
    '\n【重要】このメッセージはフォーマット不要の自然な会話です。\n' +
    '口語・略語・省略表現から意図を読み取ってください。\n' +
    '\n【担当者の読み取り方】\n' +
    '・「〇〇さんに頼んでおいて」「〇〇に確認して」→ assigneeは〇〇\n' +
    '・単独作業（資料作成・提出・見積等）→ メンションされた人のみ\n' +
    '・双方向参加（打合せ・現場確認等）→ 送信者(' + senderName + ')も追加\n' +
    '・担当者不明なら空文字（勝手に補完しない）\n' +
    '\n【タスクの読み取り方】\n' +
    '・「〜やっとく」「〜しておく」「〜お願い」→ タスク\n' +
    '・「〜承知しました」「〜やります」→ タスク受諾として記録\n' +
    '・ただの雑談・報告・返事はタスクにしない\n' +
    '\n【スケジュールの読み取り方】\n' +
    '・「〜に会いましょう」「〜で打ち合わせ」「〜に現場行く」→ スケジュール\n' +
    '・日時が含まれているものを対象にする\n' +
    '\n【案件名ヒントについて】\n' +
    '・テキストから読み取れる物件名・案件名・施主名があれば projectNameHint に入れる\n' +
    '・確信がなければ空文字にする（誤判定しない）\n' +
    '\n【複数タスク・スケジュール】必ず配列で個別出力。絶対にまとめない。\n' +
    '\n【日付の確定判定（dateConfirmed・必須）】\n' +
    '・「7/25」「明日」「来週火曜」「5月10日」など日にちが1日に特定できる → dateConfirmed: true\n' +
    '・「来週あたり」「近いうち」「そのうち」「今月中旬」「〜頃」「日程未定」「たぶん」など曖昧・仮の表現 → dateConfirmed: false\n' +
    '・曖昧な場合も deadline/date にはおおよその日付を入れてよいが、dateConfirmed は必ず false にする\n' +
    '・日付が全く無い場合も dateConfirmed: false\n' +
    '\n【出力（JSONのみ・前置き不要）】\n' +
    '{\n' +
    '  "projectNameHint": "テキストから読み取れる案件名（確信なければ空文字）",\n' +
    '  "tasks": [\n' +
    '    {"assignee":"","taskContent":"","deadline":"yyyy-MM-dd or \'\'","dateConfirmed":true|false}\n' +
    '  ],\n' +
    '  "schedules": [\n' +
    '    {"title":"","date":"yyyy-MM-dd","startTime":"HH:mm or \'\'","endTime":"HH:mm or \'\'","location":"","attendees":"","description":"","dateConfirmed":true|false}\n' +
    '  ]\n' +
    '}\n' +
    '\nタスクもスケジュールもなければ両方空配列。\n' +
    '\nメッセージ：\n' + message;

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + config.GEMINI_API_KEY;

  try {
    var res   = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.1, thinkingConfig: { thinkingBudget: 0 } } }),
    });
    var raw   = geminiText(res.getContentText());
    if (!raw) return { tasks: [], schedules: [], projectNameHint: '' };
    var match = raw.match(/\{[\s\S]*\}/);
    if (!match) return { tasks: [], schedules: [], projectNameHint: '' };

    var parsed = JSON.parse(match[0]);
    var urgent = detectUrgency(message);
    return {
      projectNameHint: parsed.projectNameHint || '',
      tasks: (Array.isArray(parsed.tasks) ? parsed.tasks : []).map(function(t) {
        // dateConfirmed未返却（旧レスポンス等）は期日ありなら確定扱いで従来動作を維持
        return { assignee: t.assignee || '', taskContent: t.taskContent || '', deadline: t.deadline || '', dateConfirmed: t.dateConfirmed === undefined ? !!t.deadline : t.dateConfirmed === true, groupId: groupId, datetime: timestamp, urgency: urgent ? '高' : '' };
      }),
      schedules: (Array.isArray(parsed.schedules) ? parsed.schedules : []).map(function(s) {
        return { title: s.title || '', date: s.date || '', startTime: s.startTime || '', endTime: s.endTime || '', location: s.location || '', attendees: s.attendees || '', description: s.description || '', dateConfirmed: s.dateConfirmed === undefined ? !!s.date : s.dateConfirmed === true, groupId: groupId, datetime: timestamp };
      }),
    };
  } catch (err) {
    console.error('extractWithGemini error:', err.message);
    return { tasks: [], schedules: [], projectNameHint: '' };
  }
}

// ==========================================
// SECTION 6: 案件識別（信頼度スコア付き）
// フォーマット不要・自然文から推定
// ==========================================

// 戻り値: { name: '案件名', confidence: 0-100 }
// confidence >= 70 → 確認画面で案件名を表示
// confidence < 70  → 案件選択UIを先に表示
function identifyProject(text, groupId, geminiHint) {
  // レベル1: グループIDから直接マッチ（グループが案件に紐づいている場合）
  var byGroup = getProjectNameByGroupId(groupId);
  if (byGroup) return { name: byGroup, confidence: 95 };

  var projects = getProjectData();
  if (!projects.length) return { name: '未分類', confidence: 0 };

  var cleanText = text.replace(/\s/g, '');
  var best = { name: '未分類', confidence: 0 };

  for (var i = 0; i < projects.length; i++) {
    var abbr   = String(projects[i][0] || '').replace(/\s/g, '');
    var formal = String(projects[i][1] || '').replace(/\s/g, '');
    if (!formal) continue;

    var score = 0;

    // レベル2: テキスト中に正式名称・略称が完全に含まれる
    if (cleanText.includes(formal)) { score = 90; }
    else if (abbr.length >= 2 && cleanText.includes(abbr)) { score = 82; }
    else {
      // レベル3: 最長共通部分文字列（3文字以上で部分一致）
      var sub = longestCommonSubstring(formal, cleanText);
      if (sub >= 5)      score = 72;
      else if (sub >= 4) score = 60;
      else if (sub >= 3) score = 45;

      // 略称でも試す
      if (abbr.length >= 2) {
        var subAbbr = longestCommonSubstring(abbr, cleanText);
        if (subAbbr >= abbr.length) score = Math.max(score, 75);
        else if (subAbbr >= 3)      score = Math.max(score, 55);
      }
    }

    if (score > best.confidence) best = { name: formal, confidence: score };
  }

  // レベル4: Geminiヒントで補完（信頼度は低め・上限65）
  if (best.confidence < 60 && geminiHint) {
    var hintMatch = matchProjectByName(geminiHint, projects);
    if (hintMatch.confidence > best.confidence) {
      best = { name: hintMatch.name, confidence: Math.min(hintMatch.confidence, 65) };
    }
  }

  return best.confidence >= 40 ? best : { name: '未分類', confidence: 0 };
}

// プロジェクト名文字列から最長共通部分文字列長を返す（2文字以上の連続一致）
function longestCommonSubstring(s1, s2) {
  var max = 0;
  for (var i = 0; i < s1.length - 1; i++) {
    for (var j = i + 2; j <= s1.length; j++) {
      var sub = s1.slice(i, j);
      if (s2.includes(sub) && sub.length > max) max = sub.length;
    }
  }
  return max;
}

// 案件名文字列をプロジェクトリストと照合
function matchProjectByName(name, projects) {
  if (!name || !projects.length) return { name: '未分類', confidence: 0 };
  var cleanName = name.replace(/\s/g, '');
  var best = { name: '未分類', confidence: 0 };
  for (var i = 0; i < projects.length; i++) {
    var formal = String(projects[i][1] || '').replace(/\s/g, '');
    var abbr   = String(projects[i][0] || '').replace(/\s/g, '');
    if (!formal) continue;
    var score = 0;
    if (formal === cleanName || abbr === cleanName) score = 90;
    else if (formal.includes(cleanName) || cleanName.includes(formal)) score = 70;
    else if (abbr.length >= 2 && (abbr.includes(cleanName) || cleanName.includes(abbr))) score = 65;
    if (score > best.confidence) best = { name: formal, confidence: score };
  }
  return best;
}

// プロジェクト候補を信頼度順に返す（最大limit件）
// AIによる自動確定は行わず、ユーザーに選んでもらうための候補リスト
function getProjectCandidates(text, groupId, geminiHint, limit) {
  limit = limit || 4;
  var projects = getProjectData();
  if (!projects.length) return [];

  var cleanText = String(text || '').replace(/\s/g, '');
  var cleanHint = String(geminiHint || '').replace(/\s/g, '');
  var scored = [];

  for (var i = 0; i < projects.length; i++) {
    var formal = String(projects[i][1] || '').replace(/\s/g, '');
    var abbr   = String(projects[i][0] || '').replace(/\s/g, '');
    var status = String(projects[i][4] || '');
    if (!formal) continue;
    if (status && status !== '進行中') continue;

    var score = 0;
    var reason = '';
    if (cleanText && cleanText.includes(formal))                       { score = 90; reason = '正式名一致'; }
    else if (abbr.length >= 2 && cleanText.includes(abbr))             { score = 82; reason = '略称一致'; }
    else {
      var sub = longestCommonSubstring(formal, cleanText);
      if (sub >= 5)      { score = 72; reason = '部分一致' + sub + '文字'; }
      else if (sub >= 4) { score = 60; reason = '部分一致' + sub + '文字'; }
      else if (sub >= 3) { score = 45; reason = '部分一致' + sub + '文字'; }
      if (abbr.length >= 2) {
        var subA = longestCommonSubstring(abbr, cleanText);
        if (subA >= abbr.length && 75 > score) { score = 75; reason = '略称含有'; }
        else if (subA >= 3 && 55 > score)      { score = 55; reason = '略称部分一致'; }
      }
    }
    if (cleanHint) {
      if (formal === cleanHint || abbr === cleanHint) {
        if (score < 85) { score = 85; reason = 'AIヒント一致'; }
      } else if (formal.includes(cleanHint) || cleanHint.includes(formal)) {
        if (score < 68) { score = 68; reason = 'AIヒント部分一致'; }
      }
    }
    if (score > 0) scored.push({ name: formal, confidence: score, reason: reason });
  }

  scored.sort(function(a, b) { return b.confidence - a.confidence; });

  // 候補が少ない場合は進行中の他案件で埋める（誤名寄せ防止のため必ず複数提示）
  if (scored.length < Math.min(limit, 3)) {
    var active = getActiveProjects(limit + 2);
    for (var k = 0; k < active.length && scored.length < limit; k++) {
      var n = active[k];
      var exists = scored.some(function(s) { return s.name === n; });
      if (!exists) scored.push({ name: n, confidence: 0, reason: '進行中案件' });
    }
  }

  return scored.slice(0, limit);
}

// 進行中のプロジェクト一覧（最新N件・Quick Reply用）
function getActiveProjects(limit) {
  var sheet = getSheet('プロジェクト管理');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getDataRange().getValues().slice(1)
    .filter(function(r) { return r[1] && (r[4] === '進行中' || !r[4]); })
    .map(function(r) { return String(r[1]); })
    .slice(0, limit || 10);
}

// ==========================================
// SECTION 7: 確認フロー UI
// ==========================================

// 仮タスクシートに保存してbatchIdを返す
// 列: A=batch_id, B=連番, C=種別, D=group_id, E=user_id,
//     F=案件名, G=内容/タイトル, H=担当者/参加者, I=期日/日時,
//     J=作成日時, K=元メッセージ, L=追加JSON
function storePending(extracted, projectName, groupId, userId, rawMessage) {
  var batchId = generateId();
  var sheet   = getSheet('仮タスク');
  var now     = fmtDT(new Date());
  var idx     = 0;

  extracted.tasks.forEach(function(t) {
    idx++;
    sheet.appendRow([batchId, idx, 'task', groupId, userId,
      projectName || '', t.taskContent || '', t.assignee || '', t.deadline || '',
      now, idx === 1 ? rawMessage : '', JSON.stringify({ urgency: t.urgency || '', dateConfirmed: t.dateConfirmed !== false, assignee_user_ids: t.assignee_user_ids || [] })]);
  });
  extracted.schedules.forEach(function(s) {
    idx++;
    var extra = JSON.stringify({ endTime: s.endTime || '', location: s.location || '', description: s.description || '', dateConfirmed: s.dateConfirmed !== false, attendees_user_ids: s.attendees_user_ids || [] });
    sheet.appendRow([batchId, idx, 'schedule', groupId, userId,
      projectName || '', s.title || '', s.attendees || '',
      (s.date || '') + (s.startTime ? ' ' + s.startTime : ''),
      now, idx === 1 ? rawMessage : '', extra]);
  });

  return batchId;
}

// 案件名が確定している場合の確認画面（ワンタップ登録）
function sendConfirmUI(replyToken, extracted, projectName, batchId) {
  if (!replyToken) return;
  var lines = ['【確認】以下を登録しますか？', '案件：' + (projectName || '未分類')];

  extracted.tasks.forEach(function(t, i) {
    lines.push(
      '\n📋 タスク' + (extracted.tasks.length > 1 ? (i + 1) : '') +
      '\n担当：' + (t.assignee || '未定') +
      '\n内容：' + t.taskContent +
      (t.deadline ? '\n期日：' + t.deadline + (t.dateConfirmed === false ? '（仮）' : '') : '')
    );
  });
  extracted.schedules.forEach(function(s, i) {
    lines.push(
      '\n📅 予定' + (extracted.schedules.length > 1 ? (i + 1) : '') +
      '\n日時：' + s.date + (s.dateConfirmed === false ? '（仮）' : '') + (s.startTime ? ' ' + s.startTime + '〜' : '') +
      '\n内容：' + s.title
    );
  });

  sendQuickReply(replyToken, lines.join('\n'), [
    { type: 'action', action: { type: 'postback', label: '✅ 登録', data: 'action=register&batch=' + batchId } },
    { type: 'action', action: { type: 'postback', label: '✏️ 修正', data: 'action=edit&batch=' + batchId } },
    { type: 'action', action: { type: 'postback', label: '❌ 無視', data: 'action=ignore&batch=' + batchId } },
  ]);
}

// 案件選択UI：候補プロジェクト + 「新規プロジェクト作成」を必ず表示
// candidatesが空の場合はgetProjectCandidatesで自動取得
function sendProjectSelectUI(replyToken, extracted, batchId, candidates, groupId, geminiHint) {
  if (!replyToken) return;
  if (!candidates) candidates = getProjectCandidates('', groupId || '', geminiHint || '', 4);

  var summary = buildItemSummary(extracted);
  var lines   = ['【保存先プロジェクトをご選択ください】', '', summary, ''];
  if (candidates.length) {
    lines.push('候補：');
    candidates.forEach(function(c, i) { lines.push((i + 1) + '. ' + c.name); });
    lines.push((candidates.length + 1) + '. 🆕 新規プロジェクトとして作成');
  } else {
    lines.push('既存の候補プロジェクトが見つかりませんでした。');
    lines.push('「🆕 新規プロジェクト作成」をご選択ください。');
  }
  lines.push('\n※類似名でも別案件として管理したい場合は新規作成をお選びください。');

  var items = candidates.map(function(c) {
    var label = c.name.length > 18 ? c.name.slice(0, 17) + '…' : c.name;
    return {
      type: 'action',
      action: {
        type: 'postback',
        label: label,
        data: 'action=set_project&batch=' + batchId + '&p=' + encodeURIComponent(c.name),
        displayText: c.name + 'に保存します',
      },
    };
  });
  items.push({
    type: 'action',
    action: {
      type: 'postback',
      label: '🆕 新規プロジェクト作成',
      data: 'action=new_project&batch=' + batchId,
      displayText: '新規プロジェクトを作成します',
    },
  });
  items.push({
    type: 'action',
    action: {
      type: 'postback',
      label: '❌ 無視',
      data: 'action=ignore&batch=' + batchId,
    },
  });

  sendQuickReply(replyToken, lines.join('\n'), items);
}

// 抽出内容を短くまとめたテキスト（案件選択UIに表示）
function buildItemSummary(extracted) {
  var lines = [];
  extracted.tasks.forEach(function(t, i) {
    lines.push((t.urgency === '高' ? '🚨 ' : '📋 ') + t.taskContent + (t.assignee ? '（' + t.assignee + '）' : '') + (t.deadline ? ' ／ ' + t.deadline : ''));
  });
  extracted.schedules.forEach(function(s) {
    lines.push('📅 ' + s.title + ' ／ ' + s.date + (s.startTime ? ' ' + s.startTime : ''));
  });
  return lines.join('\n');
}

// LINE Quick Reply 送信（共通）
function sendQuickReply(replyToken, text, items) {
  if (!replyToken) return;
  var config  = getConfig();
  var safeText = text.length > 4900 ? text.slice(0, 4900) + '…' : text;
  var payload = {
    replyToken: replyToken,
    messages: [{ type: 'text', text: safeText, quickReply: { items: items.slice(0, 13) } }],
  };
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method: 'post', muteHttpExceptions: true,
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + config.LINE_CHANNEL_ACCESS_TOKEN },
      payload: JSON.stringify(payload),
    });
  } catch (err) { console.error('sendQuickReply error:', err.message); }
}

// Push型 Quick Reply（DM送信用・replyTokenなし）
function sendQuickReplyPush(targetId, text, items) {
  if (!targetId) return;
  var config   = getConfig();
  var safeText = text.length > 4900 ? text.slice(0, 4900) + '…' : text;
  var payload  = {
    to: targetId,
    messages: [{ type: 'text', text: safeText, quickReply: { items: items.slice(0, 13) } }],
  };
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post', muteHttpExceptions: true,
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + config.LINE_CHANNEL_ACCESS_TOKEN },
      payload: JSON.stringify(payload),
    });
  } catch (err) { console.error('sendQuickReplyPush error:', err.message); }
}

// タスク確認をグループではなく送信者のDMへ送信
function sendConfirmToDM(userId, extracted, proj, batchId, groupId, text) {
  if (!userId) return;

  var lines = ['📝 タスク／予定を検出いたしました。'];
  lines.push(buildItemSummary(extracted));

  var items = [];

  if (proj.confidence >= 70) {
    // 案件が確定済み → 確認UIをDMで
    lines.push('\n案件：' + proj.name + '\n\n登録してよろしいですか？');
    items = [
      { type: 'action', action: { type: 'postback', label: '✅ 登録', data: 'action=register&batch=' + batchId } },
      { type: 'action', action: { type: 'postback', label: '✏️ 修正', data: 'action=edit&batch=' + batchId } },
      { type: 'action', action: { type: 'postback', label: '❌ 無視', data: 'action=ignore&batch=' + batchId } },
    ];
  } else {
    // 案件不明 → 候補 + 新規作成 から選択（メッセージ文面・案件ヒントから候補を順位付け）
    lines.push('\n保存先プロジェクトをご選択ください。');
    var cands = getProjectCandidates(text || '', groupId || '', (extracted && extracted.projectNameHint) || '', 4);
    cands.forEach(function(c) {
      var label = c.name.length > 18 ? c.name.slice(0, 17) + '…' : c.name;
      items.push({ type: 'action', action: { type: 'postback', label: label, data: 'action=set_project&batch=' + batchId + '&p=' + encodeURIComponent(c.name), displayText: c.name + 'に保存します' } });
    });
    items.push({ type: 'action', action: { type: 'postback', label: '🆕 新規プロジェクト作成', data: 'action=new_project&batch=' + batchId } });
    items.push({ type: 'action', action: { type: 'postback', label: '❌ 無視', data: 'action=ignore&batch=' + batchId } });
  }

  sendQuickReplyPush(userId, lines.join('\n'), items);
}

// ==========================================
// SECTION 8: Postbackハンドリング
// ==========================================
function handlePostback(ev, userId, groupId) {
  var data    = (ev.postback && ev.postback.data) ? ev.postback.data : '';
  // プロジェクト進捗管理の承認/却下/案件選択（pm_*）を最優先で処理
  if (pmRoutePostback(ev, userId, groupId)) return;
  var params  = parseParams(data);
  var action  = params.action;
  var batchId = params.batch;
  if (!action) return;

  // 即登録後の取消
  if (action === 'cancel') {
    var cancelKey = params.key;
    if (!cancelKey) { sendLineReply(ev.replyToken, '取消情報が見つかりませんでした。'); return; }
    var raw = PropertiesService.getScriptProperties().getProperty(cancelKey);
    if (!raw) { sendLineReply(ev.replyToken, '取消期限（10分）が過ぎています。'); return; }
    var info = safeParseJson(raw);
    if (Date.now() - info.ts > 10 * 60 * 1000) {
      sendLineReply(ev.replyToken, '取消期限（10分）が過ぎています。');
      return;
    }
    // タスク削除
    var taskSheet = getSheet('タスク管理');
    if (taskSheet && info.taskIds && info.taskIds.length) {
      var td = taskSheet.getDataRange().getValues();
      for (var ti = td.length - 1; ti >= 1; ti--) {
        if (info.taskIds.indexOf(String(td[ti][0])) !== -1) taskSheet.deleteRow(ti + 1);
      }
    }
    // スケジュール削除
    var schSheet = getSheet('スケジュール管理');
    if (schSheet && info.scheds && info.scheds.length) {
      var sd = schSheet.getDataRange().getValues();
      for (var si = sd.length - 1; si >= 1; si--) {
        var key2 = (sd[si][2] || '') + '|' + (sd[si][3] instanceof Date ? fmtDate(sd[si][3]) : String(sd[si][3]).slice(0, 10));
        if (info.scheds.indexOf(key2) !== -1) schSheet.deleteRow(si + 1);
      }
    }
    PropertiesService.getScriptProperties().deleteProperty(cancelKey);
    sendLineReply(ev.replyToken, '↩️ 登録を取り消しました。');
    return;
  }

  // 仮タスク全削除（グループID単位）
  if (action === 'ignore_all') {
    var targetGroup = decodeURIComponent(params.group || '');
    if (!targetGroup) { sendLineReply(ev.replyToken, '削除対象が見つかりませんでした。'); return; }
    var sheet0 = getSheet('仮タスク');
    if (sheet0 && sheet0.getLastRow() > 1) {
      var d0 = sheet0.getDataRange().getValues();
      for (var x = d0.length - 1; x >= 1; x--) {
        if (d0[x][3] === targetGroup) sheet0.deleteRow(x + 1);
      }
    }
    sendLineReply(ev.replyToken, '🗑 確認待ちの仮タスクを全て削除しました。');
    return;
  }

  // batchId不要アクションを先に処理（到達不能バグ修正）
  if (action === 'task_done') {
    handleTaskDonePostback(ev, params);
    return;
  }
  if (action === 'doc_summary') {
    handleDocSummaryPostback(ev, params);
    return;
  }
  if (action === 'link_group') {
    var linkProject = decodeURIComponent(params.p || '');
    var linkGroup   = decodeURIComponent(params.g || '');
    if (!linkProject || !linkGroup) { sendLineReply(ev.replyToken, '案件の紐付けに失敗しました。'); return; }
    linkGroupToProject(linkGroup, linkProject);
    sendLineReply(ev.replyToken, '✅ このグループを「' + linkProject + '」に紐付けました。\nメッセージ・ファイルは自動で振り分けられます。');
    return;
  }

  // グループ参加時の「新規プロジェクト作成」
  if (action === 'new_project_link') {
    var npLinkGroup = decodeURIComponent(params.g || '');
    if (!npLinkGroup) { sendLineReply(ev.replyToken, 'グループIDが取得できませんでした。'); return; }
    setNewProjectLinkMode(userId, npLinkGroup);
    sendLineReply(ev.replyToken,
      '【新規プロジェクト作成】\nプロジェクト名を次のメッセージでご入力ください（10分以内、「キャンセル」で中止）。');
    return;
  }

  if (!batchId) return;

  if (action === 'register') {
    var items = getPendingItems(batchId);
    if (!items.length) { sendLineReply(ev.replyToken, '確認データが見つかりませんでした。'); return; }
    var registered = commitPendingItems(items);
    deletePendingItems(batchId);
    sendLineReply(ev.replyToken, buildRegisteredMsg(registered));

  } else if (action === 'ignore') {
    deletePendingItems(batchId);
    sendLineReply(ev.replyToken, '❌ キャンセルしました。登録はされていません。');

  } else if (action === 'edit') {
    var items2 = getPendingItems(batchId);
    if (!items2.length) { sendLineReply(ev.replyToken, '確認データが見つかりませんでした。'); return; }
    var rawMsg = items2[0][10] || '（元のメッセージ不明）';
    setRetryMode(userId, batchId);
    sendLineReply(ev.replyToken,
      '✏️ 修正内容をお送りください。\n\n【元のメッセージ】\n' + rawMsg +
      '\n\n修正内容をお送りいただくと再度ご確認いたします。（10分以内）');

  } else if (action === 'set_project') {
    // 既存プロジェクト選択 → 即保存（候補提示UIで既に選んでもらっているため確認は省略可）
    var projectName = decodeURIComponent(params.p || '未分類');
    var items3      = getPendingItems(batchId);
    if (!items3.length) { sendLineReply(ev.replyToken, '確認データが見つかりませんでした。'); return; }

    applyProjectToPending(batchId, projectName);
    var refreshed3 = getPendingItems(batchId);
    var registered3 = commitPendingItems(refreshed3);
    deletePendingItems(batchId);
    sendLineReply(ev.replyToken,
      '✅ 『' + projectName + '』に保存しました。\n\n' + buildRegisteredMsg(registered3));

  } else if (action === 'new_project') {
    // 新規プロジェクト作成 → 次のメッセージをプロジェクト名として受け取る
    var itemsNP = getPendingItems(batchId);
    if (!itemsNP.length) { sendLineReply(ev.replyToken, '確認データが見つかりませんでした。'); return; }
    setNewProjectMode(userId, batchId);
    sendLineReply(ev.replyToken,
      '【新規プロジェクト作成】\n保存先のプロジェクト名を次のメッセージでご入力ください（10分以内、「キャンセル」で中止）。');
  }
  // 注: doc_summary / link_group は batchId 不要のため上方で先に処理される
}

function parseParams(data) {
  var result = {};
  data.split('&').forEach(function(p) { var kv = p.split('='); if (kv[0]) result[kv[0]] = kv[1] || ''; });
  return result;
}

// ==========================================
// SECTION 9: 仮タスク操作
// ==========================================
// 仮タスク一覧をグループに表示（batchId単位でボタンを出す）
function handlePendingList(replyToken, groupId) {
  var sheet = getSheet('仮タスク');
  if (!sheet || sheet.getLastRow() <= 1) {
    sendLineReply(replyToken, '✅ 現在、確認待ちのタスク・予定はございません。');
    return;
  }
  var data = sheet.getDataRange().getValues().slice(1);
  // このグループのbatchIdを収集（重複なし）
  var seen = {};
  var batches = [];
  data.forEach(function(r) {
    if (r[3] === groupId && !seen[r[0]]) {
      seen[r[0]] = true;
      batches.push({ batchId: r[0], project: r[5], content: r[6], type: r[2], createdAt: r[9] });
    }
  });

  if (!batches.length) {
    sendLineReply(replyToken, '✅ このグループの確認待ちタスク・予定はございません。');
    return;
  }

  var lines = ['📋 確認待ち一覧（' + batches.length + '件）\nタップで確認・登録できます。'];
  batches.forEach(function(b, i) {
    var icon = b.type === 'schedule' ? '📅' : '📌';
    lines.push('\n' + (i + 1) + '. ' + icon + ' ' + (b.project || '案件未定') + '：' + String(b.content).slice(0, 25));
  });

  var items = batches.slice(0, 13).map(function(b, i) {
    var label = ((i + 1) + '. ' + String(b.content).slice(0, 15)).slice(0, 20);
    return { type: 'action', action: { type: 'postback', label: label, data: 'action=register&batch=' + b.batchId } };
  });
  items.push({ type: 'action', action: { type: 'postback', label: '🗑 全て削除', data: 'action=ignore_all&group=' + groupId } });

  sendQuickReply(replyToken, lines.join('\n'), items);
}

function getPendingItems(batchId) {
  var sheet = getSheet('仮タスク');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getDataRange().getValues().filter(function(r) { return r[0] === batchId; });
}

function deletePendingItems(batchId) {
  var sheet = getSheet('仮タスク');
  if (!sheet || sheet.getLastRow() <= 1) return;
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === batchId) sheet.deleteRow(i + 1);
  }
}

// 予定の5W1H不足チェック → 不足項目を日本語で質問文にして返す（なければnull）
function buildAmbiguityQuestion(extracted) {
  if (!extracted.schedules || !extracted.schedules.length) return null;
  var s = extracted.schedules[0];
  var missing = [];
  if (!s.date && !s.startTime) missing.push('いつ・何時から');
  if (!s.location)             missing.push('どこで');
  if (!s.attendees)            missing.push('誰と（参加者）');
  if (!missing.length) return null;
  return '📅 予定を検出しましたが、以下が不明です。\n「' + (s.title || '予定') + '」\n\n不明な項目：' +
    missing.join('、') + '\n\nこのまま送っていただくか、補足をご返信ください。';
}

// 確認なしで直接登録（グループ自動登録用）
function commitDirectly(extracted, projectName, groupId, userId) {
  var registered = { tasks: [], schedules: [] };
  extracted.tasks.forEach(function(t) {
    if (isDuplicateTask(t.taskContent, t.assignee, groupId)) return;
    var task = {
      task_id:         generateId(),
      project_name:    projectName,
      content:         t.taskContent || '',
      assignee:        t.assignee    || '',
      assignee_user_ids: t.assignee_user_ids || [],
      due_date:        t.deadline    || '',
      status:          'confirmed',
      created_at:      fmtDT(new Date()),
      group_id:        groupId,
      urgency:         t.urgency     || '',
      date_confirmed:  t.dateConfirmed !== false,
    };
    writeTask(task);
    notifyAssignee(task);
    registered.tasks.push(task);
  });
  extracted.schedules.forEach(function(s) {
    // 重複防止：同グループ・同タイトル・同日付の予定は二重登録しない（シート/カレンダー/DMの重複を防ぐ）
    if (isDuplicateSchedule(s.title, s.date, groupId)) return;
    var schedule = {
      project_name: projectName,
      title:        s.title       || '',
      date:         s.date        || '',
      startTime:    s.startTime   || '',
      endTime:      s.endTime     || '',
      location:     s.location    || '',
      attendees:    s.attendees   || '',
      description:  s.description || '',
      group_id:     groupId,
      datetime:     new Date(),
      date_confirmed: s.dateConfirmed !== false,
    };
    writeSchedule(schedule);
    // 日程未確定の予定はカレンダー登録せずDM通知のみ
    var meetLink = schedule.date_confirmed ? addToCalendar(schedule) : '';
    notifySchedulePeople(schedule, meetLink);
    registered.schedules.push(schedule);
  });
  return registered;
}

// 取消用に登録済みIDをScriptPropertiesへ保存（10分TTL）
function saveCancelInfo(registered) {
  var key     = 'CANCEL_' + generateId();
  var taskIds = registered.tasks.map(function(t) { return t.task_id; });
  var scheds  = registered.schedules.map(function(s) { return s.title + '|' + s.date; });
  PropertiesService.getScriptProperties().setProperty(
    key, JSON.stringify({ taskIds: taskIds, scheds: scheds, ts: Date.now() })
  );
  return key;
}

// 案件名を仮タスクシートの該当行に適用（F列）
function applyProjectToPending(batchId, projectName) {
  var sheet = getSheet('仮タスク');
  if (!sheet || sheet.getLastRow() <= 1) return;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === batchId) sheet.getRange(i + 1, 6).setValue(projectName);
  }
}

// 仮タスク行から確認UI用オブジェクトを再構築
// 仮タスク列: [0]=batch_id, [1]=idx, [2]=type, [3]=group_id, [4]=user_id,
//             [5]=案件名, [6]=内容/title, [7]=担当者/参加者, [8]=期日/日時,
//             [9]=作成日時, [10]=元msg, [11]=extraJSON
function rebuildExtractedFromPending(items, projectName) {
  var extracted = { tasks: [], schedules: [] };
  items.forEach(function(row) {
    var extra = safeParseJson(row[11]);
    if (row[2] === 'task') {
      extracted.tasks.push({ taskContent: row[6], assignee: row[7], deadline: row[8], dateConfirmed: extra.dateConfirmed !== false, assignee_user_ids: extra.assignee_user_ids || [] });
    } else if (row[2] === 'schedule') {
      var parts = String(row[8]).split(' ');
      extracted.schedules.push({ title: row[6], date: parts[0] || '', startTime: parts[1] || '', attendees: row[7], dateConfirmed: extra.dateConfirmed !== false, attendees_user_ids: extra.attendees_user_ids || [] });
    }
  });
  return extracted;
}

// 仮タスクを本番シートへ確定書き込み
function commitPendingItems(items) {
  var registered = { tasks: [], schedules: [] };
  items.forEach(function(row) {
    var type        = row[2];
    var groupId     = row[3];
    var projectName = row[5];
    var content     = row[6];
    var assignee    = row[7];
    var dateOrDl    = row[8];
    var extra       = safeParseJson(row[11]);

    if (type === 'task') {
      if (isDuplicateTask(content, assignee, groupId)) return;
      var task = {
        task_id:         generateId(),
        project_name:    projectName,
        content:         content,
        assignee:        assignee,
        assignee_user_ids: extra.assignee_user_ids || [],
        due_date:        dateOrDl,
        status:          'confirmed',
        created_at:      fmtDT(new Date()),
        group_id:        groupId,
        urgency:         extra.urgency || '',
        date_confirmed:  extra.dateConfirmed !== false,
      };
      writeTask(task);
      notifyAssignee(task);
      registered.tasks.push(task);

    } else if (type === 'schedule') {
      var parts     = String(dateOrDl).split(' ');
      var schedule  = {
        project_name: projectName,
        title:        content,
        date:         parts[0] || '',
        startTime:    parts[1] || '',
        endTime:      extra.endTime    || '',
        location:     extra.location   || '',
        attendees:    assignee,
        description:  extra.description || '',
        group_id:     groupId,
        datetime:     new Date(),
        date_confirmed: extra.dateConfirmed !== false,
      };
      writeSchedule(schedule);
      // 日程未確定の予定はカレンダー登録せずDM通知のみ
      var meetLink = schedule.date_confirmed ? addToCalendar(schedule) : '';
      notifySchedulePeople(schedule, meetLink);
      registered.schedules.push(schedule);
    }
  });
  return registered;
}

function buildRegisteredMsg(registered) {
  var lines = ['✅ 登録が完了しました。'];
  registered.tasks.forEach(function(t) {
    lines.push('\n' + (t.urgency === '高' ? '🚨 緊急タスク' : '📋 タスク') + '\n案件：' + (t.project_name || '未分類') +
      '\n担当：' + (t.assignee || '未定') +
      '\n内容：' + t.content +
      (t.due_date ? '\n期日：' + t.due_date + (t.date_confirmed === false ? '（仮・カレンダー未登録）' : '') : ''));
  });
  registered.schedules.forEach(function(s) {
    lines.push('\n📅 予定\n案件：' + (s.project_name || '未分類') +
      '\n日時：' + s.date + (s.date_confirmed === false ? '（仮・カレンダー未登録）' : '') + (s.startTime ? ' ' + s.startTime + '〜' : '') +
      '\n内容：' + s.title);
  });
  return lines.join('\n');
}

function notifyAssignee(task) {
  var dueLabel = task.due_date
    ? task.due_date + (task.date_confirmed === false ? '（仮・カレンダー未登録）' : '')
    : '未定';
  var msg = (task.urgency === '高' ? '🚨【至急】タスクのご連絡\n' : '【タスクのご連絡】\n') +
    '案件：' + (task.project_name || '未分類') +
    '\n内容：' + task.content +
    '\n期日：' + dueLabel +
    '\n\nご確認のほどよろしくお願いいたします。';
  if (task.assignee_user_ids && task.assignee_user_ids.length) {
    notifyMembersByUserIds(task.assignee_user_ids, msg);
  } else {
    notifyMembersByNames(task.assignee, msg);
  }
}

// 予定の参加者全員に予定内容をDM通知（カレンダー招待と併せて予定を反映）
function notifySchedulePeople(schedule, meetLink) {
  if (!schedule || (!schedule.attendees && !(schedule.attendees_user_ids && schedule.attendees_user_ids.length))) return;
  var isTentative = schedule.date_confirmed === false;
  var msg = '【予定のご連絡】\n案件：' + (schedule.project_name || '未分類') +
    '\n予定：' + schedule.title +
    '\n日時：' + (schedule.date ? schedule.date + (isTentative ? '（仮）' : '') : '未定') + (schedule.startTime ? ' ' + schedule.startTime : '') +
    (schedule.location ? '\n場所：' + schedule.location : '') +
    (meetLink ? '\n会議URL：' + meetLink : '') +
    (isTentative
      ? '\n\n※日程が未確定のためカレンダーには登録していません。確定後に改めてご連絡ください。'
      : '\n\nカレンダーに登録しました。ご確認のほどよろしくお願いいたします。');
  if (schedule.attendees_user_ids && schedule.attendees_user_ids.length) {
    notifyMembersByUserIds(schedule.attendees_user_ids, msg);
  } else {
    notifyMembersByNames(schedule.attendees, msg);
  }
}

// 名前リスト（userId配列）を分解し、各メンバーのDMへ通知（同一userIdへの重複送信は防止）
function notifyMembersByUserIds(userIds, message) {
  var seen = {};
  (Array.isArray(userIds) ? userIds : []).forEach(function(uid) {
    if (!uid || seen[uid]) return;
    seen[uid] = true;
    sendLineMessage(uid, message);
  });
}

// 名前文字列（複数名可）を分解し、各メンバーのDMへ通知（同一userIdへの重複送信は防止）
function notifyMembersByNames(namesText, message) {
  var seen = {};
  splitMemberNames(namesText).forEach(function(name) {
    var uid = getMemberUserId(name);
    if (uid && !seen[uid]) { seen[uid] = true; sendLineMessage(uid, message); }
  });
}

// ==========================================
// SECTION 10: 修正（retry）モード管理
// ==========================================
function setRetryMode(userId, batchId) {
  PropertiesService.getScriptProperties().setProperty(
    'RETRY_' + userId,
    JSON.stringify({ batchId: batchId, ts: Date.now() })
  );
}
function getRetryMode(userId) {
  var raw = PropertiesService.getScriptProperties().getProperty('RETRY_' + userId);
  if (!raw) return null;
  try {
    var obj = JSON.parse(raw);
    if (Date.now() - obj.ts > 10 * 60 * 1000) { clearRetryMode(userId); return null; } // 10分TTL
    return obj.batchId;
  } catch (e) { return null; }
}
function clearRetryMode(userId) {
  PropertiesService.getScriptProperties().deleteProperty('RETRY_' + userId);
}

// グループ紐付け用 新規プロジェクト名入力モード（10分TTL）
function setNewProjectLinkMode(userId, groupId) {
  PropertiesService.getScriptProperties().setProperty(
    'NEWPROJLINK_' + userId,
    JSON.stringify({ groupId: groupId, ts: Date.now() })
  );
}
function getNewProjectLinkMode(userId) {
  var raw = PropertiesService.getScriptProperties().getProperty('NEWPROJLINK_' + userId);
  if (!raw) return null;
  try {
    var obj = JSON.parse(raw);
    if (Date.now() - obj.ts > 10 * 60 * 1000) { clearNewProjectLinkMode(userId); return null; }
    return obj.groupId;
  } catch (e) { return null; }
}
function clearNewProjectLinkMode(userId) {
  PropertiesService.getScriptProperties().deleteProperty('NEWPROJLINK_' + userId);
}

// プロジェクト名を受け取り、登録＋グループ紐付けを実行
function finalizeNewProjectLink(projectName, groupId, userId, replyToken) {
  projectName = String(projectName || '').trim();
  if (/^(キャンセル|cancel|中止|やめる)$/i.test(projectName)) {
    sendLineReply(replyToken, '❌ 新規プロジェクト作成を中止しました。');
    return;
  }
  if (!projectName || projectName.length > 60) {
    setNewProjectLinkMode(userId, groupId);
    sendLineReply(replyToken, '⚠️ プロジェクト名が無効です（1〜60文字）。もう一度ご入力ください。');
    return;
  }
  registerNewProject(projectName, groupId);
  linkGroupToProject(groupId, projectName);
  sendLineReply(replyToken,
    '✅ 新規プロジェクト「' + projectName + '」を作成し、このグループに紐付けました。\nメッセージ・ファイルは自動で振り分けられます。');
}

// 新規プロジェクト名入力モード（10分TTL）
function setNewProjectMode(userId, batchId) {
  PropertiesService.getScriptProperties().setProperty(
    'NEWPROJ_' + userId,
    JSON.stringify({ batchId: batchId, ts: Date.now() })
  );
}
function getNewProjectMode(userId) {
  var raw = PropertiesService.getScriptProperties().getProperty('NEWPROJ_' + userId);
  if (!raw) return null;
  try {
    var obj = JSON.parse(raw);
    if (Date.now() - obj.ts > 10 * 60 * 1000) { clearNewProjectMode(userId); return null; }
    return obj.batchId;
  } catch (e) { return null; }
}
function clearNewProjectMode(userId) {
  PropertiesService.getScriptProperties().deleteProperty('NEWPROJ_' + userId);
}

// 新規プロジェクト名を受け取り、プロジェクト登録＋仮タスク保存をまとめて実行
function finalizeNewProject(projectName, batchId, groupId, userId, replyToken) {
  projectName = String(projectName || '').trim();

  // キャンセル系入力で中止
  if (/^(キャンセル|cancel|中止|やめる)$/i.test(projectName)) {
    deletePendingItems(batchId);
    sendLineReply(replyToken, '❌ 新規プロジェクト作成を中止しました。仮タスクも削除しました。');
    return;
  }
  if (!projectName) {
    setNewProjectMode(userId, batchId);
    sendLineReply(replyToken, 'プロジェクト名が空です。もう一度ご入力ください。');
    return;
  }
  if (projectName.length > 60) {
    setNewProjectMode(userId, batchId);
    sendLineReply(replyToken, 'プロジェクト名が長すぎます（60文字以内）。もう一度ご入力ください。');
    return;
  }

  var items = getPendingItems(batchId);
  if (!items.length) {
    sendLineReply(replyToken, '⚠️ 確認データが見つかりませんでした（時間切れの可能性があります）。');
    return;
  }

  registerNewProject(projectName, groupId);
  applyProjectToPending(batchId, projectName);
  var refreshed  = getPendingItems(batchId);
  var registered = commitPendingItems(refreshed);
  deletePendingItems(batchId);

  sendLineReply(replyToken,
    '✅ 新規プロジェクト『' + projectName + '』を作成し、以下を保存しました。\n\n' +
    buildRegisteredMsg(registered));
}

function reprocessMessage(text, oldBatchId, groupId, userId, sender, ts, replyToken) {
  // 元メッセージを削除前に取得（曖昧質問への補足入力に使用）
  var pendingItems = getPendingItems(oldBatchId);
  var originalMsg  = (pendingItems.length && pendingItems[0][10]) ? String(pendingItems[0][10]) : '';
  deletePendingItems(oldBatchId);

  if (shouldSkipExtraction(text)) {
    sendLineReply(replyToken, 'タスク・スケジュールを検出できませんでした。お手数ですが、もう一度ご入力いただけますでしょうか。');
    return;
  }

  // 元メッセージがある場合 → 補足情報として結合して再抽出（ruleBasedFilter不要）
  // 元メッセージがない場合 → ruleBasedFilterを通す
  var textForExtraction = originalMsg ? (originalMsg + '\n補足：' + text) : text;
  if (!originalMsg && !ruleBasedFilter(text)) {
    sendLineReply(replyToken, 'タスク・スケジュールを検出できませんでした。お手数ですが、もう一度ご入力いただけますでしょうか。');
    return;
  }

  var extracted = extractWithGemini(textForExtraction, groupId, ts, sender);
  if (!extracted || (extracted.tasks.length + extracted.schedules.length === 0)) {
    sendLineReply(replyToken, 'タスク・スケジュールが見つかりませんでした。恐れ入りますが、内容を確認の上もう一度お試しください。');
    return;
  }
  var boundProject = getProjectNameByGroupId(groupId);
  if (boundProject) {
    // グループ紐付け済み → 確認UI（案件は確定として表示）
    var batchIdB = storePending(extracted, boundProject, groupId, userId, text);
    sendConfirmUI(replyToken, extracted, boundProject, batchIdB);
    return;
  }
  // 未紐付け → 候補+新規作成の選択UIへ
  var batchId = storePending(extracted, '', groupId, userId, text);
  var allProjects = getProjectData();
  if (!allProjects.length) {
    setNewProjectMode(userId, batchId);
    sendLineReply(replyToken,
      '【新規プロジェクト作成】\n保存先のプロジェクト名を次のメッセージでご入力ください（10分以内、「キャンセル」で中止）。');
    return;
  }
  var candidates = getProjectCandidates(text, groupId, extracted.projectNameHint, 4);
  sendProjectSelectUI(replyToken, extracted, batchId, candidates, groupId, extracted.projectNameHint);
}

// ==========================================
// SECTION 11: データ書き込み
// ==========================================

// タスク管理シート列: task_id|案件名|タスク内容|担当者|期日|ステータス|作成日時|グループID|緊急度
function writeTask(task) {
  var sheet = getSheet('タスク管理');
  if (!sheet) return;

  var res = pmWithLock(function() {
    // Legacy 9-column append: do not auto-migrate sheet schema during normal task registration.
    sheet.appendRow([
      task.task_id || '', task.project_name || '', task.content || '', task.assignee || '',
      task.due_date || '', task.status || 'confirmed', task.created_at || '', task.group_id || '',
      task.urgency || '',
    ]);

    var rowNumber = sheet.getLastRow();
    var idx = pmHeaderIndex(sheet);
    try {
      if (idx['担当者ID'] !== undefined) {
        sheet.getRange(rowNumber, idx['担当者ID'] + 1)
          .setValue(Array.isArray(task.assignee_user_ids) ? task.assignee_user_ids.join('、') : '');
      }
      if (idx['関連案件ID'] !== undefined) {
        sheet.getRange(rowNumber, idx['関連案件ID'] + 1)
          .setValue(task.related_project_id || '');
      }
      if (idx['作成元'] !== undefined) {
        sheet.getRange(rowNumber, idx['作成元'] + 1)
          .setValue(task.source || '');
      }
    } catch (e) {
      console.warn('writeTask: 拡張列追記に失敗しました。row=' + rowNumber + ', task_id=' + (task.task_id || '') + ', error=' + e.message);
      // 既存9列タスク登録を維持するため、拡張列追記失敗時は例外を再スローしません。
    }
  });

  if (!res || res.ok === false) {
    console.error('writeTask: ロック取得/実行に失敗しました。', res && res.error);
    return;
  }

  if (task.project_name && task.project_name !== '未分類') {
    try {
      registerNewProject(task.project_name, task.group_id);
    } catch (e) {
      console.error('writeTask: registerNewProject failed', { task_id: task.task_id || '', action: 'registerNewProject', error: e.message });
    }
  }
  // 期日未確定（date_confirmed === false）のタスクはカレンダー登録しない（DM通知のみ）
  if (task.date_confirmed !== false) {
    try {
      addTaskToCalendar(task);
    } catch (e) {
      console.error('writeTask: addTaskToCalendar failed', { task_id: task.task_id || '', action: 'addTaskToCalendar', error: e.message });
    }
  }
}

// スケジュール管理シート列: 登録日時|案件名|予定タイトル|日付|開始時間|終了時間|場所|参加者|詳細|グループID
function writeSchedule(schedule) {
  getSheet('スケジュール管理').appendRow([
    fmtDT(schedule.datetime || new Date()),
    schedule.project_name || '', schedule.title || '', schedule.date || '',
    schedule.startTime || '', schedule.endTime || '', schedule.location || '',
    schedule.attendees || '', schedule.description || '', schedule.group_id || '',
  ]);
  if (schedule.project_name && schedule.project_name !== '未分類') registerNewProject(schedule.project_name, schedule.group_id);
}

// メッセージログシート列: 日時|グループID|送信者|メッセージ
function saveMessageLog(groupId, senderName, text, timestamp) {
  try {
    var sheet = getSheet('メッセージログ');
    if (sheet) sheet.appendRow([fmtDT(timestamp), groupId, senderName, text]);
  } catch (err) { console.error('saveMessageLog error:', err.message); }
}

// ==========================================
// SECTION 12: Google Calendar
// ==========================================

// オンライン面談の判定キーワード（タイトル・場所・備考のいずれかに含まれれば対象）
var ONLINE_MEETING_KEYWORDS = [
  'オンライン', 'リモート', 'ビデオ会議', 'ビデオ通話', 'テレビ会議',
  'web会議', 'web面談', 'web打合せ', 'web打ち合わせ', 'google meet', 'グーグルミート',
];

function isOnlineSchedule(schedule) {
  var text = ((schedule.title || '') + ' ' + (schedule.location || '') + ' ' + (schedule.description || '')).toLowerCase();
  return ONLINE_MEETING_KEYWORDS.some(function(kw) { return text.indexOf(kw) !== -1; });
}

// 本文に既に会議URL（Zoom/Meet/Teams/Webex）があればそれを優先し、Meetの二重発行を防ぐ
function findExistingConferenceUrl(schedule) {
  var text = (schedule.location || '') + ' ' + (schedule.description || '');
  var m = text.match(/https?:\/\/\S*(zoom\.us|meet\.google\.com|teams\.microsoft\.com|webex\.com)\S*/i);
  return m ? m[0] : '';
}

// Calendar API (Advanced Service) で Google Meet リンク付き予定を作成し、MeetのURLを返す
function createMeetEvent(calendarId, title, start, end, description, location, guestEmails) {
  var resource = {
    summary: title,
    description: description,
    location: location || 'オンライン',
    start: { dateTime: start.toISOString(), timeZone: 'Asia/Tokyo' },
    end:   { dateTime: end.toISOString(),   timeZone: 'Asia/Tokyo' },
    attendees: guestEmails.map(function(email) { return { email: email }; }),
    conferenceData: {
      createRequest: {
        requestId: Utilities.getUuid(),
        conferenceSolutionKey: { type: 'hangoutsMeet' },
      },
    },
  };
  var ev = Calendar.Events.insert(resource, calendarId, {
    conferenceDataVersion: 1,
    sendUpdates: guestEmails.length ? 'all' : 'none',
  });
  if (ev.hangoutLink) return ev.hangoutLink;
  var eps = ev.conferenceData && ev.conferenceData.entryPoints;
  return (eps && eps.length && eps[0].uri) || '';
}

// 戻り値: 会議URL（オンライン面談でMeetを発行 or 本文中の既存URL）。なければ ''
function addToCalendar(schedule) {
  try {
    var config = getConfig();
    if (!schedule.date || !config.CALENDAR_ID) {
      console.log('addToCalendar スキップ: date=' + schedule.date + ' CALENDAR_ID=' + (config.CALENDAR_ID ? '設定済' : '未設定'));
      return '';
    }
    var calendar = CalendarApp.getCalendarById(config.CALENDAR_ID);
    if (!calendar) {
      console.error('addToCalendar: カレンダーが見つかりません CALENDAR_ID=' + config.CALENDAR_ID);
      return '';
    }

    var title = '【' + (schedule.project_name || 'WOODBASE') + '】' + schedule.title;
    var date  = new Date(schedule.date);
    var guestList = getMemberGoogleEmailsFromNames(schedule.attendees);
    var guests = guestList.join(',');
    var description = buildCalendarDescription(schedule.description, schedule.attendees);
    var options = { location: schedule.location || '', description: description };
    if (guests) { options.guests = guests; options.sendInvites = true; }

    if (schedule.startTime) {
      var sp    = schedule.startTime.split(':');
      var start = new Date(date); start.setHours(parseInt(sp[0]), parseInt(sp[1]), 0);
      var end   = new Date(start);
      if (schedule.endTime) {
        var ep = schedule.endTime.split(':'); end.setHours(parseInt(ep[0]), parseInt(ep[1]), 0);
      } else {
        end.setHours(start.getHours() + 1, start.getMinutes(), 0);
      }
      // 重複チェック：同日・同時刻・完全同一タイトルのみスキップ
      if (hasSimilarCalendarEvent(calendar, title, start, end, false)) {
        console.log('addToCalendar 重複スキップ:', title);
        return '';
      }
      // オンライン面談 → Google Meet リンク付きで作成（既存の会議URLがあれば発行しない）
      var existingUrl = findExistingConferenceUrl(schedule);
      if (!existingUrl && isOnlineSchedule(schedule)) {
        try {
          var meetLink = createMeetEvent(config.CALENDAR_ID, title, start, end, description, schedule.location, guestList);
          console.log('カレンダー登録完了(Meet付き):', title + ' / ' + meetLink + (guests ? ' / 招待=' + guests : ''));
          return meetLink;
        } catch (meetErr) {
          console.warn('Meet付き登録に失敗、通常登録にフォールバック:', meetErr.message);
        }
      }
      calendar.createEvent(title, start, end, options);
      console.log('カレンダー登録完了:', title + (guests ? ' / 招待=' + guests : ''));
      return existingUrl;
    }
    // 終日イベントは重複チェックしない（同日に同名予定が複数登録されるケースは少ない）
    calendar.createAllDayEvent(title, date, options);
    console.log('カレンダー登録完了:', title + (guests ? ' / 招待=' + guests : ''));
    return '';
  } catch (err) { console.error('addToCalendar error:', err.message); return ''; }
}

function addTaskToCalendar(task) {
  try {
    var config = getConfig();
    if (!task || !task.due_date || !config.CALENDAR_ID) return;
    var calendar = CalendarApp.getCalendarById(config.CALENDAR_ID);
    if (!calendar) return;

    var date = new Date(task.due_date);
    if (isNaN(date.getTime())) return;

    var title = '【' + (task.project_name || 'WOODBASE') + '】' + task.content;
    var guests = getMemberGoogleEmailsFromNames(task.assignee).join(',');
    var description = buildCalendarDescription(
      'タスク期日\n担当者: ' + (task.assignee || '未定') + '\nステータス: ' + (task.status || 'confirmed'),
      task.assignee
    );
    var options = { description: description };
    if (guests) {
      options.guests = guests;
      options.sendInvites = true;
    }
    if (hasSimilarCalendarEvent(calendar, title, date, null, true)) return;
    calendar.createAllDayEvent(title, date, options);
    console.log('タスク期日カレンダー登録:', title + (guests ? ' / guests=' + guests : ''));
  } catch (err) { console.error('addTaskToCalendar error:', err.message); }
}

function buildCalendarDescription(baseDescription, attendeesText) {
  var lines = [];
  if (baseDescription) lines.push(baseDescription);
  if (attendeesText) lines.push('参加者/担当者: ' + attendeesText);
  lines.push('WOODBASE LINE秘書AIから自動登録');
  return lines.join('\n');
}

function hasSimilarCalendarEvent(calendar, title, start, end, allDay) {
  try {
    var events = allDay ? calendar.getEventsForDay(start) : calendar.getEvents(start, end);
    for (var i = 0; i < events.length; i++) {
      if (events[i].getTitle() === title) {
        console.log('カレンダー重複スキップ:', title);
        return true;
      }
    }
  } catch (e) {
    console.warn('カレンダー重複チェック失敗:', e.message);
  }
  return false;
}

// 会話履歴(2026/04/16〜2026/04/28)から抽出した未来の予定を一括登録
// スケジュール管理シート＋Googleカレンダーに反映
// タスクはタスク管理シートに反映
function bulkImportHistorySchedules() {
  // ===== スケジュール（確定日付あり） =====
  var schedules = [
    { project_name: 'プラージュ東海店',           title: '東海店ガラス・電気・内装工事',  date: '2026-04-30', startTime: '',      location: '', attendees: '加藤大誠', description: '' },
    { project_name: '知多半田',                   title: 'パネル工事',                    date: '2026-05-01', startTime: '',      location: '', attendees: '',         description: '' },
    { project_name: 'Amuyo',                      title: 'カチオンしごき',                date: '2026-05-02', startTime: '08:30', location: '', attendees: 'HIRO',     description: '' },
    { project_name: 'Amuyo',                      title: '塗装工事一期',                  date: '2026-05-03', startTime: '',      location: '', attendees: '大鐘勇気', description: '5/3〜5/6' },
    { project_name: 'UMITERRACE宮古島',           title: 'WBグループ宮古島作業',          date: '2026-05-07', startTime: '',      location: '宮古島', attendees: '南忠則', description: '5/7〜5/16' },
    { project_name: 'UMITERRACE宮古島',           title: 'NTT光回線下見',                 date: '2026-05-07', startTime: '13:00', location: '宮古島', attendees: '伊藤健一090-7831-4251', description: '本館→別館配線ルート再下見' },
    { project_name: 'UMITERRACE宮古島',           title: 'サイン製作完了予定',            date: '2026-05-08', startTime: '',      location: '広島工場', attendees: 'タイペックス', description: 'バイブレーション切り文字' },
    { project_name: 'UMITERRACE宮古島',           title: '撮影',                          date: '2026-05-09', startTime: '',      location: '', attendees: '',         description: '5/9〜5/10' },
    { project_name: '釧路桂木',                   title: '家具納品',                      date: '2026-05-11', startTime: '10:00', location: '釧路市鶴野東3町1番1号 SKワークス', attendees: '鈴木きよひで090-3397-4139', description: '4t1車で理容のみ' },
    { project_name: '雨のち晴れクリニック',       title: '照明器具納品',                  date: '2026-05-11', startTime: '15:00', location: '川崎市中原区新丸子東2-886 1F', attendees: 'アレン株式会社小川080-9676-8634', description: '4t車で配送' },
    { project_name: 'UMITERRACE宮古島',           title: '消防検査',                      date: '2026-05-12', startTime: '14:30', location: '宮古島', attendees: '魚見', description: '' },
    { project_name: 'UMITERRACE宮古島',           title: 'サイン引渡し最終',              date: '2026-05-14', startTime: '',      location: '宮古島', attendees: 'タイペックス', description: '' },
    { project_name: '美容今治小泉',               title: '家具納品',                      date: '2026-05-15', startTime: '',      location: '店舗', attendees: 'カエルデザイン辰巳070-1212-6392', description: '5/18に変更可能性あり' },
    { project_name: '理容利府店',                 title: '待合椅子完成・納品',            date: '2026-05-15', startTime: '',      location: '', attendees: 'プラージュ建設佐野', description: '' },
    { project_name: '理容出雲店',                 title: '追加待合椅子2台納品',           date: '2026-05-16', startTime: '',      location: '店舗', attendees: '', description: '家財便・時間指定不可' },
    { project_name: 'Amuyo',                      title: '什器搬入',                      date: '2026-05-18', startTime: '',      location: '渋谷松濤', attendees: '', description: '' },
    { project_name: '関西版収録',                 title: '関西版収録設営',                date: '2026-05-19', startTime: '08:00', location: '', attendees: '清水淳之介', description: '8〜9時設営' },
    { project_name: '美容旭川永山',               title: '家具納品',                      date: '2026-05-20', startTime: '08:30', location: '店舗', attendees: '柴田建設', description: '4t1車' },
    { project_name: '理容武豊里中',               title: '家具納品',                      date: '2026-05-21', startTime: '09:00', location: '店舗', attendees: '中井建設', description: '' },
    { project_name: '理容旭川永山',               title: '家具納品',                      date: '2026-05-22', startTime: '08:30', location: '店舗', attendees: '柴田建設', description: '4t1車' },
    { project_name: '理容プラージュ砺波',         title: '搬入',                          date: '2026-05-22', startTime: '',      location: '店舗', attendees: '', description: '' },
    { project_name: 'Amuyo',                      title: '竣工クリーニング',              date: '2026-05-23', startTime: '',      location: '渋谷松濤', attendees: '', description: '' },
    { project_name: 'Amuyo',                      title: '全体予備日',                    date: '2026-05-24', startTime: '',      location: '', attendees: '', description: '' },
    { project_name: '雨のち晴れクリニック',       title: 'お引渡し',                      date: '2026-05-24', startTime: '',      location: '武蔵小杉', attendees: '', description: '' },
    { project_name: 'Amuyo',                      title: 'お引渡し',                      date: '2026-05-25', startTime: '15:00', location: '渋谷松濤', attendees: '', description: '' },
    { project_name: 'Amuyo',                      title: 'オープン目標',                  date: '2026-05-30', startTime: '',      location: '渋谷松濤', attendees: 'てんか・シン', description: '' },
    { project_name: '理容大和田',                 title: '家具納期',                      date: '2026-06-09', startTime: '',      location: '', attendees: '柴田建設', description: '9席バージョン' },
    { project_name: '関西版収録',                 title: '関西版収録設営',                date: '2026-06-09', startTime: '08:00', location: '', attendees: '清水淳之介', description: '' },
    { project_name: '関西版収録',                 title: '関西版収録設営',                date: '2026-07-14', startTime: '08:00', location: '', attendees: '清水淳之介', description: '' },
    { project_name: '理容美容岐阜',               title: 'お祝いの花納品',                date: '2026-07-08', startTime: '16:00', location: '岐阜市殿町6-612', attendees: '山本建設', description: 'オープン前日' },
    { project_name: '関西版収録',                 title: '関西版収録設営',                date: '2026-08-10', startTime: '08:00', location: '', attendees: '清水淳之介', description: '' },
    { project_name: '関西版収録',                 title: '関西版収録設営',                date: '2026-09-08', startTime: '08:00', location: '', attendees: '清水淳之介', description: '' },
    { project_name: '関西版収録',                 title: '関西版収録設営',                date: '2026-10-13', startTime: '08:00', location: '', attendees: '清水淳之介', description: '' },
    { project_name: '関西版収録',                 title: '関西版収録設営',                date: '2026-11-10', startTime: '08:00', location: '', attendees: '清水淳之介', description: '' },
    { project_name: '関西版収録',                 title: '関西版収録設営',                date: '2026-12-08', startTime: '08:00', location: '', attendees: '清水淳之介', description: '' },
  ];

  // ===== タスク（期限あり） =====
  var tasks = [
    { project_name: 'WOODBASE社内経理', content: '4月分立替経費を取りまとめて提出',          assignee: '相馬俊之・土井愛加・ラッキー池田・清水淳之介', due_date: '2026-04-29' },
    { project_name: 'WOODBASE社内経理', content: 'Aun3月分汲み取り請求書PDF共有',           assignee: 'ラッキー池田',                                 due_date: '2026-04-27' },
    { project_name: 'プラージュ蒲田',   content: '請負契約書 リーガルチェック反映後の捺印', assignee: '土井愛加',                                     due_date: '2026-05-01' },
    { project_name: 'プラージュ西山',   content: '4月請求予定分の見積書提出',               assignee: '相馬俊之',                                     due_date: '2026-05-01' },
    { project_name: '東邦作図',         content: '4月東邦作図見積書作成',                   assignee: '善波明日香',                                   due_date: '2026-05-02' },
    { project_name: 'UMITERRACE宮古島', content: '受水槽手配・見積書提出',                  assignee: 'ラッキー池田',                                 due_date: '2026-05-08' },
    { project_name: 'UMITERRACE宮古島', content: '電気容量申請（津田クリニック）',          assignee: 'hayato',                                       due_date: '2026-05-02' },
    { project_name: '雨のち晴れクリニック', content: '建具・枠 発注',                        assignee: '土井愛加',                                     due_date: '2026-05-04' },
    { project_name: 'Amuyo',            content: '建具品番・サイン仕様の最終決定',          assignee: 'TATSUYA ODA・土井愛加',                        due_date: '2026-04-30' },
    { project_name: 'プラージュ豊田高橋', content: '理容豊田高橋 待合椅子納品(AM9時)',      assignee: '安田陳列',                                     due_date: '2026-05-15' },
    { project_name: 'プラージュ蒲田',   content: 'プラージュ蒲田 契約書 アクティリビング返送', assignee: '土井愛加',                                   due_date: '2026-04-24' },
    { project_name: 'UMITERRACE宮古島', content: 'リモートロック設置（GOAL HD仕様）',       assignee: 'ラッキー池田',                                 due_date: '2026-05-08' },
  ];

  var added_s = 0, added_t = 0;
  schedules.forEach(function(s) {
    s.datetime = new Date();
    s.group_id = '';
    writeSchedule(s);
    addToCalendar(s);
    added_s++;
    Utilities.sleep(500); // カレンダーAPIレート対策
  });

  tasks.forEach(function(t) {
    var task = {
      task_id:      generateId(),
      project_name: t.project_name,
      content:      t.content,
      assignee:     t.assignee,
      due_date:     t.due_date,
      status:       'confirmed',
      created_at:   fmtDT(new Date()),
      group_id:     '',
    };
    writeTask(task);
    added_t++;
  });

  console.log('スケジュール追加:', added_s + '件 / タスク追加:', added_t + '件');
}

// ==========================================
// SECTION 13: ファイル保存
// ==========================================
function saveFileToDrive(messageId, fileName, groupId, timestamp, messageType) {
  try {
    var config  = getConfig();
    var res     = UrlFetchApp.fetch('https://api-data.line.me/v2/bot/message/' + messageId + '/content', {
      headers: { 'Authorization': 'Bearer ' + config.LINE_CHANNEL_ACCESS_TOKEN }, muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) return;

    var blob = res.getBlob();

    // 新: 店舗フォルダへ種別振り分け保存（写真/動画→5_写真、その他→6_データ）＋台帳記録＋冪等（pm_linefiles）。
    //     店舗が特定できない場合は「未紐付け」で退避。失敗時のみ従来のフラット保存にフォールバック。
    if (typeof pmSaveLineFile === 'function') {
      var r = pmSaveLineFile(messageType || 'file', messageId, fileName, groupId, timestamp, blob);
      if (r && r.ok) {
        console.log('ファイル保存(新):', r.duplicate ? '重複スキップ' : (r.held ? '未紐付け退避 ' : '') + (r.url || ''));
        return;
      }
      console.warn('pmSaveLineFile失敗→従来保存にフォールバック:', r && r.error);
    }

    // フォールバック（従来）: 親/案件名 にフラット保存
    var dateStr     = Utilities.formatDate(timestamp, 'Asia/Tokyo', 'yyyyMMdd');
    var safeName    = dateStr + '_' + (fileName || messageId);
    var projectName = getProjectNameByGroupId(groupId) || detectProjectFromRecentLogs(groupId) || '未分類';
    var folder      = getOrCreateFolder(config.DRIVE_FOLDER_ID, projectName);
    if (folder) {
      blob.setName(safeName);
      folder.createFile(blob);
      console.log('ファイル保存(従来):', projectName + '/' + safeName);
    }
  } catch (err) { console.error('saveFileToDrive error:', err.message); }
}

function getOrCreateFolder(parentId, name) {
  if (!parentId) return null;
  try {
    var parent   = DriveApp.getFolderById(parentId);
    var existing = parent.getFoldersByName(name);
    return existing.hasNext() ? existing.next() : parent.createFolder(name);
  } catch (e) {
    console.warn('getOrCreateFolder skipped:', e.message);
    return null;
  }
}

// メッセージログにある未紐付けグループに案件選択UIを送信
function promptAllGroupsToLinkProject() {
  var sheet = getSheet('メッセージログ');
  if (!sheet || sheet.getLastRow() <= 1) { console.log('ログなし'); return; }

  var data     = sheet.getDataRange().getValues().slice(1);
  var groupIds = {};
  data.forEach(function(r) {
    var gid = String(r[1] || '');
    if (gid && gid.startsWith('C')) groupIds[gid] = true;
  });

  var projects = getActiveProjects(11);
  if (!projects.length) { console.log('案件がありません'); return; }

  var sent = 0;
  Object.keys(groupIds).forEach(function(gid) {
    // すでに紐付け済みならスキップ
    if (getProjectNameByGroupId(gid)) {
      console.log('登録済みスキップ:', gid);
      return;
    }

    var items = projects.map(function(name) {
      var label = name.length > 20 ? name.slice(0, 19) + '…' : name;
      return { type: 'action', action: { type: 'postback', label: label,
        data: 'action=link_group&p=' + encodeURIComponent(name) + '&g=' + encodeURIComponent(gid) } };
    });
    items.push({ type: 'action', action: { type: 'postback', label: '🆕 新規プロジェクト作成',
      data: 'action=new_project_link&g=' + encodeURIComponent(gid) } });

    sendQuickReplyPush(gid, 'このグループはどの案件に該当しますか？タップでご選択ください。', items);
    sent++;
    Utilities.sleep(500);
  });

  console.log('送信完了:', sent + 'グループ');
}

// メッセージログに履歴がある全グループのメンバーを一括登録
// メッセージログから抽出したグループID群を対象にメンバー同期
function syncAllGroupMembersFromLogs() {
  var sheet = getSheet('メッセージログ');
  if (!sheet || sheet.getLastRow() <= 1) { console.log('ログなし'); return 0; }

  var data     = sheet.getDataRange().getValues().slice(1);
  var groupIds = {};
  data.forEach(function(r) {
    var gid = String(r[1] || '');
    if (gid && gid.startsWith('C')) groupIds[gid] = true;
  });

  var ids = Object.keys(groupIds);
  console.log('[Logs] 対象グループ数:', ids.length);

  var total = 0;
  ids.forEach(function(gid) {
    console.log('[Logs] 同期中:', gid);
    try {
      var count = syncGroupMembers(gid);
      total += count || 0;
    } catch(e) {
      console.warn('[Logs] スキップ:', gid, e.message);
    }
    Utilities.sleep(500);
  });

  console.log('[Logs] 全グループ同期完了 / 新規登録合計:', total + '人');
  return total;
}

// グループの全メンバーをメンバー管理シートに一括同期
function syncGroupMembers(groupId) {
  var config = getConfig();
  if (!groupId) groupId = config.INTERNAL_GROUP_ID;
  if (!groupId) { console.error('groupIdが必要です'); return; }

  var url  = 'https://api.line.me/v2/bot/group/' + groupId + '/members/ids';
  var ids  = [];

  // ページネーション対応
  var next = null;
  do {
    var reqUrl = url + (next ? '?start=' + next : '');
    var res    = UrlFetchApp.fetch(reqUrl, {
      headers: { 'Authorization': 'Bearer ' + config.LINE_CHANNEL_ACCESS_TOKEN },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      console.error('members/ids API error', res.getResponseCode(), res.getContentText().slice(0, 300));
      break;
    }
    var body = JSON.parse(res.getContentText());
    ids  = ids.concat(body.memberIds || []);
    next = body.next || null;
  } while (next);

  console.log('グループメンバー取得:', ids.length + '人');

  var registered = 0;
  ids.forEach(function(uid) {
    if (registerMember(uid, groupId)) registered++;
    Utilities.sleep(200); // API制限対策
  });

  console.log('新規登録:', registered + '人');
  return registered;
}

// 公式LINEを追加しているフォロワー全員をメンバー管理に一括登録
function registerAllFollowers() {
  var config = getConfig();
  var token  = config.LINE_CHANNEL_ACCESS_TOKEN;
  var ids    = [];
  var next   = null;

  do {
    var url = 'https://api.line.me/v2/bot/followers/ids' + (next ? '?start=' + next : '');
    var res = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) { console.error('followers API error:', res.getContentText()); break; }
    var body = JSON.parse(res.getContentText());
    ids  = ids.concat(body.userIds || []);
    next = body.next || null;
  } while (next);

  console.log('フォロワー取得:', ids.length + '人');
  var registered = 0;
  ids.forEach(function(uid) {
    if (registerMember(uid)) registered++;
    Utilities.sleep(200);
  });
  console.log('新規登録:', registered + '人（既登録はスキップ）');
  return registered;
}

// 全フォロワーへブロードキャスト送信（未認証アカウントでも使用可）
// 受け取ったユーザーが返信すると handleDM → registerMember で自動登録される
// 実行前に LINE 公式アカウントの送信枠を確認すること（フォロワー数 ≒ 消費通数）
function broadcastRegistrationPrompt(customText) {
  var config = getConfig();
  var token  = config.LINE_CHANNEL_ACCESS_TOKEN;
  if (!token) { console.error('LINE_CHANNEL_ACCESS_TOKENが未設定'); return; }

  var defaultText =
    '【WBG君（AI秘書）からのお知らせ】\n' +
    '\n' +
    'WOODBASE秘書AIに登録するため、このメッセージに「OK」とご返信ください！\n' +
    '\n' +
    '登録するとタスク・スケジュール通知があなたのDMに届くようになります。\n' +
    '\n' +
    '※既に登録済みの方は再返信しても問題ありません。';

  var payload = {
    messages: [{ type: 'text', text: customText || defaultText }],
  };

  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', {
    method:  'post',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  console.log('HTTPステータス:', res.getResponseCode());
  console.log('レスポンス:', res.getContentText().slice(0, 500));
  if (res.getResponseCode() === 200) {
    console.log('✅ 全フォロワーに送信完了。返信が届き次第、メンバー管理に自動登録されます。');
  }
}

// プロジェクト管理シートの全グループIDからメンバーを一括登録
// プロジェクト管理シート由来のグループID群を対象にメンバー同期
function syncAllGroupMembersFromProjects() {
  var sheet = getSheet('プロジェクト管理');
  if (!sheet || sheet.getLastRow() <= 1) { console.error('プロジェクト管理シートにデータがありません'); return 0; }

  var data    = sheet.getDataRange().getValues().slice(1);
  var groupIds = {};
  data.forEach(function(r) {
    [r[2], r[3]].forEach(function(gid) {
      if (gid && String(gid).startsWith('C')) groupIds[String(gid)] = true;
    });
  });

  var ids = Object.keys(groupIds);
  console.log('[Projects] 対象グループ数:', ids.length);
  var total = 0;
  ids.forEach(function(gid) {
    console.log('[Projects] --- グループ:', gid);
    try { total += syncGroupMembers(gid) || 0; } catch(e) { console.warn('[Projects] スキップ:', gid, e.message); }
  });
  console.log('[Projects] 合計新規登録:', total + '人');
  return total;
}

// 統合関数（外部呼び出し互換のため名前は維持）
function syncAllGroupMembers() {
  var fromLogs     = syncAllGroupMembersFromLogs() || 0;
  var fromProjects = syncAllGroupMembersFromProjects() || 0;
  console.log('合計新規登録:', (fromLogs + fromProjects) + '人 (Logs:' + fromLogs + ' / Projects:' + fromProjects + ')');
}

// ボットがグループ参加時 → 案件選択UIを表示
function handleBotJoinGroup(replyToken, groupId) {
  // グループメンバーを一括登録
  try { syncGroupMembers(groupId); } catch(e) { console.warn('メンバー同期エラー:', e.message); }

  var existing = getProjectNameByGroupId(groupId);
  if (existing) {
    sendLineReply(replyToken, 'このグループはすでに「' + existing + '」に紐付けられています。');
    return;
  }
  sendProjectLinkUI(replyToken, groupId, false);
}

// @メンションで「案件確認」と送った時、およびボット参加時に表示する案件紐付けUI。
// force=true の場合は、既存の紐付けがあっても選び直せる。
function sendProjectLinkUI(replyToken, groupId, force) {
  var existing = getProjectNameByGroupId(groupId);
  var projects = getActiveProjects(11);
  if (!projects.length) {
    sendLineReply(replyToken, 'WOODBASE秘書AIです。\nまずスプレッドシートの「プロジェクト管理」に案件をご登録ください。');
    return;
  }
  var items = projects.map(function(name) {
    var label = name.length > 20 ? name.slice(0, 19) + '…' : name;
    return { type: 'action', action: { type: 'postback', label: label,
      data: 'action=link_group&p=' + encodeURIComponent(name) + '&g=' + encodeURIComponent(groupId) } };
  });
  items.push({ type: 'action', action: { type: 'postback', label: '🆕 新規プロジェクト作成',
    data: 'action=new_project_link&g=' + encodeURIComponent(groupId) } });
  var intro = force
    ? '【案件確認】\n現在の紐付け：' + (existing || '未設定') + '\nこのグループの案件を選び直してください。'
    : 'WOODBASE秘書AIです。\nこのグループはどの案件に該当しますか？タップでご選択ください。';
  sendQuickReply(replyToken, intro, items);
}

// グループでボットにメンションして「案件確認」と送るための判定。
// 通常の「案件更新」等とは分け、誤って紐付けUIを開かないよう確認・登録系の明示語だけに限定する。
function isProjectLinkRequest(text) {
  var t = String(text || '').trim();
  // 実際のLINEメンションでは本文先頭に「@Bot名」が残るため、完全一致にはしない。
  return /案件\s*(?:を\s*)?(?:確認|登録|紐付け|ひも付け)(?:して)?/.test(t);
}

// グループIDをプロジェクト管理シートに登録
function linkGroupToProject(groupId, projectName) {
  var sheet = getSheet('プロジェクト管理');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  var targetRow = 0;
  for (var i = 1; i < data.length; i++) {
    // 選び直し時は、以前の案件に残った同じグループIDを外す。
    if (data[i][2] === groupId) sheet.getRange(i + 1, 3).clearContent();
    if (data[i][3] === groupId) sheet.getRange(i + 1, 4).clearContent();
    if (String(data[i][1] || '') === projectName) targetRow = i + 1;
  }
  if (!targetRow) return;

  // 既存の別グループを消さないよう、空いている従来列にだけ記録する。
  var target = sheet.getRange(targetRow, 3, 1, 2).getValues()[0];
  if (!target[0] || target[0] === groupId) sheet.getRange(targetRow, 3).setValue(groupId);
  else if (!target[1] || target[1] === groupId) sheet.getRange(targetRow, 4).setValue(groupId);

  // 新しい「案件グループ」台帳がある場合はこちらも更新する。案件解決はこの台帳を優先する。
  try { upsertProjectGroupLink_(groupId, projectName); } catch (e) { console.warn('案件グループ更新エラー:', e.message); }
  console.log('グループ紐付け:', projectName, groupId);
}

// 案件グループ台帳のグループIDを選択案件へ付け替える（台帳未整備時は従来動作のみ）。
function upsertProjectGroupLink_(groupId, projectName) {
  if (typeof PM_SHEET_GROUPS === 'undefined' || typeof pmGetProjectByName !== 'function') return;
  var project = pmGetProjectByName(projectName);
  var sheet = getSheet(PM_SHEET_GROUPS);
  if (!project || !sheet || sheet.getLastRow() < 1) return;
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var gidCol = headers.indexOf('グループID');
  var pidCol = headers.indexOf('案件ID');
  if (gidCol < 0 || pidCol < 0) return;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][gidCol] || '') === String(groupId)) {
      sheet.getRange(i + 1, pidCol + 1).setValue(project['案件ID']);
      return;
    }
  }
  sheet.appendRow([groupId, project['案件ID'], 'その他', '', 'LINEからの案件確認で紐付け', fmtDT(new Date())]);
}

function getProjectNameByGroupId(groupId) {
  if (!groupId) return null;
  // 案件グループ(新・1案件複数グループ対応)を優先。未整備/未一致なら従来の施主/業者2列にフォールバック＝非破壊。
  try {
    var viaGroups = pmGetProjectNameByGroupIdFromGroups(groupId);
    if (viaGroups) return viaGroups;
  } catch (e) { /* 案件グループ未整備でも従来動作を維持 */ }
  var sheet = getSheet('プロジェクト管理');
  if (!sheet) return null;
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === groupId || data[i][3] === groupId) return data[i][1] || data[i][0];
  }
  return null;
}

// 直近ログから案件名を推定（ファイル送信時など）
function detectProjectFromRecentLogs(groupId) {
  var sheet = getSheet('メッセージログ');
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var data     = sheet.getDataRange().getValues();
  var filtered = data.slice(1).filter(function(r) { return r[1] === groupId; }).slice(-20);
  var projects = getProjectData();
  // 直近のメッセージと案件名を照合
  for (var i = filtered.length - 1; i >= 0; i--) {
    var text = String(filtered[i][3] || '');
    var match = matchProjectByName(text, projects);
    if (match.confidence >= 70) return match.name;
  }
  return null;
}

// skipDriveFolder=true のとき Driveフォルダを作らない。
//   PM（プロジェクト進捗管理）経由の新規案件は pmEnsureProjectFolder() が
//   新階層「ROOT/案件名」の案件フォルダを1つだけ作るため、
//   ここでのフラットなフォルダ作成（重複の原因）を抑止する。
//   既存のタスク自動登録(writeTask)等から呼ばれる場合は従来どおり ROOT/案件名 を作成する。
function registerNewProject(projectName, groupId, skipDriveFolder) {
  // LINE/Webhookからの新規案件も進捗ボードと同じ案件作成サービスへ集約する。
  // pmEnsureProjectRecord 内部からは skipDriveFolder=true で呼ばれるため再帰しない。
  if (!skipDriveFolder && typeof pmEnsureProjectRecord === 'function') {
    pmEnsureProjectRecord(projectName, { groupId: groupId || '', source: 'line' });
    return;
  }
  var sheet = getSheet('プロジェクト管理');
  if (!sheet) return;
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === projectName || data[i][0] === projectName) return;
  }
  sheet.appendRow([projectName, projectName, groupId || '', '', '進行中', '自動登録', '', '']);
  if (!skipDriveFolder) {
    var config = getConfig();
    if (config.DRIVE_FOLDER_ID) getOrCreateFolder(config.DRIVE_FOLDER_ID, projectName);
  }
  console.log('新規プロジェクト登録:', projectName);
}

// 会話履歴から抽出したグループID→プロジェクトの初期マッピングを一括投入
// 既に同じグループIDが登録済みの場合はスキップ
function bulkRegisterProjectsFromHistory() {
  // [略称, 正式名称, グループID（施主）, グループID（業者）, ステータス, 備考]
  var rows = [
    ['東海・知多半田', 'プラージュ東海店・知多半田',           '', 'C4db58afd8296dc6f392964056bc0b997', '進行中', 'プラージュ建設加藤'],
    ['帯広',           '理容プラージュ帯広店',                '', 'Ce72787f92d5dbfbf07d62a9c97bfdd8f', '進行中', ''],
    ['旭川永山',       '理容美容プラージュ旭川永山',          '', 'Cb851bdd98fc5316520a446de7a453f64', '進行中', '柴田建設'],
    ['二ツ亀',         '二ツ亀（外注作図）',                  '', 'Cf28403603197b83c18df671bd5d6bcc6', '進行中', '作図業者'],
    ['宮古是正',       '宮古島UMITERRACE是正チーム',          '', 'C09549bd01c9268511ad549c40a371764', '進行中', '社内'],
    ['UMI宮古',        'UMITERRACE宮古島',                    '', 'C763cf0f9dd386d0f16ddc12d1438a89d', '進行中', ''],
    ['UMI/385',        'UMITERRACE/385ホテル',                'C9b3bf966f9f9a58000daee0ab4b38d72', '', '進行中', '施主'],
    ['UMIサイン',      'UMITERRACEサイン（タイペックス）',    '', 'Caab4c642dae021080347b29d581bc339', '進行中', ''],
    ['津田歯科',       '津田歯科クリニック',                  'Cf68c3cf7b2a0577afb2ebb1bf523f7fa', '', '進行中', '施主'],
    ['雨晴れ',         '雨のち晴れクリニック武蔵小杉',        '', 'Ce9dbe398fc443e33a75592aa8f42db42', '進行中', ''],
    ['アムヨ社内',     'アムヨ/雨晴/武蔵小杉 社内',           '', 'C5889df50cbe5b8813655514aa903cce5', '進行中', '社内'],
    ['Amuyo',          'Amuyo渋谷松濤',                       '', 'C51c7a19eb0dc8157948b4a30bace367e', '進行中', ''],
    ['Amuyo施主',      'Amuyo（てんか・シン）',               'Ce7b409f51c9a0533024451461df715ed', '', '進行中', '施主'],
    ['安田陳列',       '安田陳列（家具配送）',                '', 'C84df4447ae6c56dfba37ebe7650b3ea7', '進行中', '業者'],
    ['佐野店舗群',     'プラージュ建設佐野（複数店舗）',      '', 'C86bdb559212d773aaba18e8bd77a1f14', '進行中', ''],
    ['糸洲店舗群',     'プラージュ建設糸洲（今治小泉/小幡）', '', 'C91bc2398d4e578a685757ef26b1d1eae', '進行中', ''],
    ['中井建設1',      'カエルデザイン/中井建設',             '', 'Cf991d32865459ae86fd5399c9be6607a', '進行中', ''],
    ['中井建設2',      '中井建設',                            '', 'Cdfead3d4bedbc18e93b3ff5399315c08', '進行中', ''],
    ['小林店舗群',     'プラージュ建設小林奨伍',              '', 'Cc095c20df5b6d59a50d134494a6219d0', '進行中', ''],
    ['ラフリー天王寺', 'ラフリー天王寺（クロス・カーテン）',  '', 'Cf1ba725429a43e1f62b4677895c0d611', '進行中', ''],
    ['ラフリー本部',   'ラフリー本部',                        'Cc9fd62e92842aca3d8f0f86f7aa2a47f', '', '進行中', '施主'],
    ['ラフリーDi-noc', 'ラフリー天王寺（ダイノック）',        '', 'Cfa18bb4faa562f4305a6a470245c213a', '進行中', ''],
    ['ラフリー電気',   'ラフリー天王寺（電気屋・山内）',      '', 'C1ec82464904c6d51ea6c467cbd3f1449', '進行中', ''],
    ['ラフリー清水',   'ラフリー天王寺（清水・HIRO）',        '', 'Cab11234b7071e8902720a5af35383653', '進行中', ''],
    ['岐阜',           '理容美容プラージュ岐阜（山本建設）',  '', 'C04901189e304fd3f2cfe4ada91901ac7', '進行中', ''],
    ['べんてん屋',     '飲食店改修（べんてん屋）',            'C6866d8e822256f01529359d6c6937aa4', '', '進行中', '施主'],
    ['砺波',           '理容プラージュ砺波',                  '', 'Cb2c5b6b75107f65f4ef9595b004923af', '進行中', ''],
    ['川西',           '美容プラージュ川西',                  '', 'Cfaef42d7c9b407ae18390f01127731a1', '進行中', ''],
    ['河合',           '河合（外注作図）',                    '', 'Cc2f403959cdefdf6006f945be0bb51c1', '進行中', '作図業者'],
    ['WBG開発',        'WBG開発テスト',                       '', 'C064410a492da8938d05b3105fb90c58a', '進行中', '社内'],
    ['経理',           'WOODBASE社内経理',                    '', 'Caeea6c76ec54ef9f800ae13a2958db92', '進行中', '社内'],
    ['WBF',            'WBF内部チーム',                       '', 'Ce064982f9932d5f5bb7e1c52725376f2', '進行中', '社内'],
    ['WBF事務',        'WBF事務スタッフ',                     '', 'C5d464a41877ed749da53c487b40205f8', '進行中', '社内'],
    ['WB内部',         'WBグループ内部',                      '', 'C13e8ef03a8cbdb7e9bb9eeaab5dc631a', '進行中', '社内'],
    ['ことく',         'ことくしゃちょー（加盟店）',          'C3a991d9ab7f0e2bbbed0bd5574c841eb', '', '進行中', '施主'],
    ['水道局',         '水道局協議（宮古島）',                '', 'Cd082ba7096bf37f7d8be34b050c4c7c8', '進行中', '行政'],
  ];

  var sheet = getSheet('プロジェクト管理');
  if (!sheet) { console.error('プロジェクト管理シートが存在しません'); return; }

  var existing = {};
  if (sheet.getLastRow() > 1) {
    sheet.getDataRange().getValues().slice(1).forEach(function(r) {
      [r[2], r[3]].forEach(function(g) { if (g) existing[String(g)] = true; });
    });
  }

  var added = 0, skipped = 0;
  rows.forEach(function(row) {
    var gid = row[2] || row[3];
    if (existing[gid]) { skipped++; return; }
    sheet.appendRow(row);
    added++;
  });
  console.log('追加:', added + '件 / スキップ:', skipped + '件');
}

// ==========================================
// SECTION 14: メンバー管理
// ==========================================
function handleGoogleCalendarConnectRequest(ev, userId, text) {
  var cleanText = String(text || '').trim();
  if (!cleanText) return false;

  if (isEmailAddress(cleanText)) {
    saveMemberGoogleEmail(userId, cleanText);
    sendLineReply(ev.replyToken,
      'Googleカレンダー連携用のメールアドレスを登録しました。\n\n今後、担当タスクや参加予定が登録された場合は、共通カレンダーの予定に招待します。');
    return true;
  }

  if (isGoogleCalendarConnectKeyword(cleanText)) {
    sendLineReply(ev.replyToken,
      'Googleカレンダー連携をご希望の場合は、予定を受け取りたいGoogleアカウントのメールアドレスをこのトークに送ってください。\n\n例：name@gmail.com\n\n登録後も、会社の共通カレンダーへの登録はこれまで通り継続します。');
    return true;
  }
  return false;
}

function isGoogleCalendarConnectKeyword(text) {
  return ['Googleカレンダー連携', 'googleカレンダー連携', 'カレンダー連携', 'カレンダー登録', 'Googleカレンダー登録', 'googleカレンダー登録'].some(function(k) {
    return text.indexOf(k) !== -1;
  });
}

function isEmailAddress(text) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(text || '').trim());
}

function sendGoogleCalendarConnectPromptToAllMembers() {
  var sheet = getSheet('メンバー管理');
  if (!sheet || sheet.getLastRow() <= 1) { console.log('送信対象メンバーなし'); return 0; }
  var data = sheet.getDataRange().getValues();
  var sent = 0;
  for (var i = 1; i < data.length; i++) {
    var userId = data[i][1];
    if (!userId || String(userId).indexOf('U') !== 0) continue;
    sendGoogleCalendarConnectPrompt(userId);
    sent++;
    Utilities.sleep(200);
  }
  console.log('Googleカレンダー連携案内を送信:', sent + '人');
  return sent;
}

// GASエディタの実行関数一覧で選びやすい短縮名
function sendCalendarPromptAll() {
  return sendGoogleCalendarConnectPromptToAllMembers();
}

function sendGoogleCalendarConnectPrompt(userId) {
  if (!userId) return;
  var config = getConfig();
  var text = 'GoogleカレンダーにWBGの予定を登録したい方は、Googleアカウントのメールアドレスを返信してください。\n\n今後タスクや予定として登録された際に、個人のカレンダーにも招待で予定が届くようになります。';
  var message = {
    type: 'text',
    text: text,
    quickReply: {
      items: [{
        type: 'action',
        action: { type: 'message', label: 'カレンダー登録', text: 'カレンダー登録' },
      }],
    },
  };
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      muteHttpExceptions: true,
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + config.LINE_CHANNEL_ACCESS_TOKEN },
      payload: JSON.stringify({ to: userId, messages: [message] }),
    });
  } catch (err) { console.error('sendGoogleCalendarConnectPrompt error:', err.message); }
}

function registerMember(userId, groupId) {
  var sheet = getSheet('メンバー管理');
  if (!sheet) return false;
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (data[i][1] === userId) return false; }

  var displayName = '未設定';
  try {
    var config = getConfig();
    var token  = config.LINE_CHANNEL_ACCESS_TOKEN;
    // グループ文脈ではグループメンバーAPIを先に試す
    if (groupId && groupId !== userId) {
      var gRes = UrlFetchApp.fetch('https://api.line.me/v2/bot/group/' + groupId + '/member/' + userId, {
        headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true,
      });
      if (gRes.getResponseCode() === 200) displayName = JSON.parse(gRes.getContentText()).displayName || '未設定';
    }
    // 取得できなければ通常プロフィールAPIにフォールバック
    if (displayName === '未設定') {
      var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/' + userId, {
        headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true,
      });
      if (res.getResponseCode() === 200) displayName = JSON.parse(res.getContentText()).displayName || '未設定';
    }
  } catch (e) {}

  sheet.appendRow([displayName, userId, '社内', '', '']);
  console.log('メンバー登録:', displayName);
  return true;
}

function getMemberNameByUserId(userId) {
  var sheet = getSheet('メンバー管理');
  if (!sheet) return null;
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (data[i][1] === userId) return data[i][0] || null; }
  return null;
}

function getMemberUserId(name) {
  if (!name) return null;
  var sheet = getSheet('メンバー管理');
  if (!sheet) return null;
  var clean = name.replace('さん', '').trim();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && String(data[i][0]).includes(clean)) return data[i][1] || null;
  }
  return null;
}

// Googleメール列のインデックスを返す（読み取り専用・列自動作成なし）
// 列が存在しない場合は -1 を返す
function getGoogleEmailColIdx(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
  return headers.indexOf('Googleメール'); // 0-indexed、なければ -1
}

// Googleメール列を確保して列番号（1-indexed）を返す（書き込み用）
function ensureGoogleEmailCol(sheet) {
  var idx = getGoogleEmailColIdx(sheet);
  if (idx !== -1) return idx + 1;
  var newCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, newCol).setValue('Googleメール');
  return newCol;
}

function saveMemberGoogleEmail(userId, email) {
  var sheet = getSheet('メンバー管理');
  if (!sheet || !userId || !isEmailAddress(email)) return false;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === userId) {
      var col = ensureGoogleEmailCol(sheet); // 書き込み時のみ列を確保
      sheet.getRange(i + 1, col).setValue(String(email).trim());
      console.log('Googleメール登録:', userId + ' / ' + email);
      return true;
    }
  }
  // 未登録の場合はメンバー登録してから再度保存（再帰は1回のみ）
  registerMember(userId);
  var data2 = sheet.getDataRange().getValues();
  for (var j = 1; j < data2.length; j++) {
    if (data2[j][1] === userId) {
      var col2 = ensureGoogleEmailCol(sheet);
      sheet.getRange(j + 1, col2).setValue(String(email).trim());
      console.log('Googleメール登録（新規メンバー）:', userId + ' / ' + email);
      return true;
    }
  }
  return false;
}

function getMemberGoogleEmail(name) {
  if (!name) return null;
  var sheet = getSheet('メンバー管理');
  if (!sheet) return null;
  var colIdx = getGoogleEmailColIdx(sheet); // 読み取り専用・列自動作成なし
  if (colIdx === -1) return null; // Googleメール列が存在しなければ即return
  var clean = normalizeMemberName(name);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var memberName = normalizeMemberName(data[i][0]);
    if (memberName && memberName === clean && isEmailAddress(data[i][colIdx])) return String(data[i][colIdx]).trim();
  }
  for (var j = 1; j < data.length; j++) {
    var memberName2 = normalizeMemberName(data[j][0]);
    if (memberName2 && (memberName2.indexOf(clean) !== -1 || clean.indexOf(memberName2) !== -1) && isEmailAddress(data[j][colIdx])) {
      return String(data[j][colIdx]).trim();
    }
  }
  return null;
}

function getMemberGoogleEmailsFromNames(namesText) {
  if (!namesText) return [];
  var seen = {};
  var emails = [];
  splitMemberNames(namesText).forEach(function(name) {
    var email = getMemberGoogleEmail(name);
    if (email && !seen[email.toLowerCase()]) {
      seen[email.toLowerCase()] = true;
      emails.push(email);
    }
  });
  return emails;
}

function splitMemberNames(namesText) {
  return String(namesText || '')
    .split(/[、,，・\/／\s]+/)
    .map(function(v) { return normalizeMemberName(v); })
    .filter(function(v) { return !!v; });
}

function normalizeMemberName(name) {
  return String(name || '')
    .replace(/さん/g, '')
    .replace(/[0-9０-９+\-ー－()\s]/g, '')
    .trim();
}

// ==========================================
// SECTION 15: クエリ応答
// ==========================================
function isQuery(text) {
  return ['残タスク', 'タスクは', 'タスクある', '何件', '進捗', 'どうなってる', '教えて', 'スケジュール', '今週', '来週', '予定',
          '休み', '定休', '営業', '会社', '住所', '電話', 'GW', 'ゴールデン', 'お盆', '年末年始', '連休'].some(function(k) { return text.includes(k); });
}

// 会社情報シートを読み取りテキスト化（質問プロンプト注入用）
function buildCompanyInfoText() {
  var sheet = getSheet('会社情報');
  if (!sheet || sheet.getLastRow() <= 1) return '（未登録）';
  return sheet.getDataRange().getValues().slice(1)
    .filter(function(r) { return r[0] && r[1]; })
    .map(function(r) { return '・' + r[0] + '：' + r[1]; })
    .join('\n');
}

// ==========================================
// プロジェクト 会社名・住所 照会・更新
// ==========================================

// プロジェクト管理シートの列インデックスをヘッダー名から動的に取得
function getProjColIndex(headers, colName) {
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === colName) return i;
  }
  return -1;
}

// 住所・会社名に関するクエリかどうか判定
function isAddressOrCompanyQuery(text) {
  return ['住所', '所在地', '会社名', '施主名'].some(function(k) { return text.includes(k); });
}

// 住所・会社名の更新リクエストを解析
// 例: "雨晴れの住所：東京都〇〇" / "津田歯科の会社名：株式会社ABC"
// 戻り値: { projectQuery, field, value } or null
function parseProjectFieldUpdate(text) {
  var m = text.match(/^(.+?)(?:案件)?の(住所|所在地|会社名|施主名)[はをに]?[：:\s]+(.+)/);
  if (!m) return null;
  var field = (m[2] === '住所' || m[2] === '所在地') ? '住所' : '会社名';
  return { projectQuery: m[1].trim(), field: field, value: m[3].trim() };
}

// プロジェクト管理シートの指定フィールドを更新
// 戻り値: 更新成功なら案件名（文字列）、案件未発見なら false
function updateProjectField(projectQuery, fieldName, value) {
  var sheet = getSheet('プロジェクト管理');
  if (!sheet || sheet.getLastRow() <= 1) return false;
  var allVals = sheet.getDataRange().getValues();
  var colIdx  = getProjColIndex(allVals[0], fieldName);
  if (colIdx < 0) return false;
  var cleanQ  = String(projectQuery).normalize('NFKC').replace(/\s/g, '');
  for (var i = 1; i < allVals.length; i++) {
    var abbr = String(allVals[i][0] || '').normalize('NFKC').replace(/\s/g, '');
    var full  = String(allVals[i][1] || '').normalize('NFKC').replace(/\s/g, '');
    if ((abbr.length >= 2 && cleanQ.includes(abbr)) ||
        (full.length  >= 2 && cleanQ.includes(full))) {
      sheet.getRange(i + 1, colIdx + 1).setValue(value);
      return String(allVals[i][1] || allVals[i][0]);
    }
  }
  return false;
}

// クエリに該当する案件の会社名・住所を返す
// 戻り値: { name, company, address } or null（案件不明なら null）
function lookupProjectInfo(query) {
  var sheet = getSheet('プロジェクト管理');
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var allVals = sheet.getDataRange().getValues();
  var headers = allVals[0];
  var rows    = allVals.slice(1);
  var colCo   = getProjColIndex(headers, '会社名');
  var colAddr = getProjColIndex(headers, '住所');
  var cleanQ  = String(query || '').normalize('NFKC').replace(/\s/g, '');
  var best = null, bestScore = 0;
  rows.forEach(function(r) {
    var abbr = String(r[0] || '').normalize('NFKC').replace(/\s/g, '');
    var full  = String(r[1] || '').normalize('NFKC').replace(/\s/g, '');
    var score = 0;
    if (abbr.length >= 2 && cleanQ.includes(abbr)) score = Math.max(score, abbr.length * 2);
    if (full.length  >= 2 && cleanQ.includes(full))  score = Math.max(score, full.length  * 2);
    if (score > bestScore) {
      bestScore = score;
      best = {
        name:    String(r[1] || r[0]),
        company: colCo   >= 0 ? String(r[colCo]   || '') : '',
        address: colAddr >= 0 ? String(r[colAddr]  || '') : '',
      };
    }
  });
  return bestScore >= 4 ? best : null;
}

// 住所・会社名クエリへの返信テキストを生成
// null を返した場合は通常フローへ流す（案件不明時）
function answerAddressOrCompanyQuery(query) {
  var info = lookupProjectInfo(query);
  if (!info) return null;
  var lines = ['📋 ' + info.name];
  if (info.company) lines.push('会社名：' + info.company);
  if (info.address) lines.push('住所：'   + info.address);
  if (!info.company && !info.address) {
    lines.push('（会社名・住所は未登録です）');
    lines.push('登録する場合：\n  ' + info.name + 'の住所：〇〇〇');
  }
  return lines.join('\n');
}

// 進行中案件の会社名・住所一覧（Geminiコンテキスト用）
function buildProjectAddressListText() {
  var sheet = getSheet('プロジェクト管理');
  if (!sheet || sheet.getLastRow() <= 1) return '';
  var allVals = sheet.getDataRange().getValues();
  var headers = allVals[0];
  var rows    = allVals.slice(1);
  var colCo   = getProjColIndex(headers, '会社名');
  var colAddr = getProjColIndex(headers, '住所');
  if (colCo < 0 && colAddr < 0) return '';
  var lines = rows
    .filter(function(r) { return r[1] && (r[4] === '進行中' || !r[4]); })
    .filter(function(r) { return (colCo >= 0 && r[colCo]) || (colAddr >= 0 && r[colAddr]); })
    .map(function(r) {
      var parts = ['・' + (r[1] || r[0])];
      if (colCo   >= 0 && r[colCo])   parts.push('会社名：' + r[colCo]);
      if (colAddr >= 0 && r[colAddr]) parts.push('住所：'  + r[colAddr]);
      return parts.join(' ／ ');
    });
  return lines.join('\n');
}

// ==========================================
// 案件メモ（特記事項）管理
// ==========================================

// 明示的なメモ保存依頼かどうか判定
function isMemoRequest(text) {
  return ['メモして', 'メモお願い', '覚えておいて', '記録して', '特記事項', '備考として',
          'メモ：', 'メモ:', '特記：', '特記:', '注意：', '注意:', '重要：', '重要:'].some(function(k) {
    return text.includes(k);
  });
}

// メモ照会クエリかどうか判定
function isMemoQuery(text) {
  return ['メモは', 'メモ一覧', 'メモ教えて', '特記事項は', '注意点は', '備考は', '覚えてる', 'メモある'].some(function(k) {
    return text.includes(k);
  });
}

// Geminiでメモ内容を1〜2文に整形
function extractMemoContent(text, config) {
  var prompt =
    '以下のメッセージから「案件メモ・特記事項」として保存すべき要点を1〜2文で簡潔に抽出してください。\n' +
    '前置き・挨拶は除き、要点のみ出力してください。\n\nメッセージ：' + text;
  return callGemini(config.GEMINI_API_KEY, prompt, 0.1) || text.slice(0, 200);
}

// 案件メモをシートに保存
function saveMemo(projectName, content, sender, groupId) {
  var sheet = getSheet('案件メモ');
  if (!sheet) return false;
  sheet.appendRow([new Date(), projectName || '未分類', content, sender || '', groupId || '']);
  return true;
}

// 案件メモを取得（最新20件・案件名でフィルタ）
function getMemoList(projectName) {
  var sheet = getSheet('案件メモ');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var rows = sheet.getDataRange().getValues().slice(1);
  if (projectName) {
    rows = rows.filter(function(r) {
      var rName = String(r[1] || '');
      return rName === projectName || rName.includes(projectName) || projectName.includes(rName);
    });
  }
  return rows.slice(-20);
}

// メモ照会への返信テキストを生成
function answerMemoQuery(text, groupId) {
  var projectName = getProjectNameByGroupId(groupId) || '';
  if (!projectName) {
    var info = lookupProjectInfo(text);
    if (info) projectName = info.name;
  }
  var memos = getMemoList(projectName);
  if (!memos.length) {
    return '📝 ' + (projectName || '案件') + 'のメモは現在ありません。\n' +
      '登録する場合：\n@WBG メモ：〇〇〇';
  }
  var lines = ['📝 ' + (projectName || '全案件') + 'のメモ（' + memos.length + '件）：'];
  memos.forEach(function(r) {
    var d = r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Tokyo', 'M/d') : String(r[0]).slice(5, 10);
    lines.push('・[' + d + '] ' + r[2]);
  });
  return lines.join('\n');
}

// メモ保存を実行して返信テキストを返す（失敗時は null）
function handleMemoSave(text, groupId, sender) {
  var config      = getConfig();
  var projectName = getProjectNameByGroupId(groupId) || '';
  var content     = extractMemoContent(text, config);
  if (!content) return null;
  saveMemo(projectName, content, sender, groupId);
  return '📝 メモを保存しました。\n案件：' + (projectName || '未分類') + '\n内容：' + content;
}

// 単独で案件メモシートを作成（setup()なしで実行可能）
function addMemoSheet() {
  var ss    = getSS();
  var sheet = ss.getSheetByName('案件メモ');
  if (sheet) { console.log('ℹ️ 案件メモシートは既に存在します'); return; }
  setupSheet(ss, '案件メモ',
    ['登録日時', '案件名', '内容', '発言者', 'グループID'],
    '#B4A7D6', [140, 160, 400, 100, 160]);
  console.log('✅ 案件メモシートを作成しました');
}

// ==========================================
// ボールホルダー管理（誰が今ボールを持っているか）
// ==========================================

// ボールホルダークエリかどうか判定
function isBallHolderQuery(text) {
  return ['ボール', '誰待ち', '誰が待ち', '誰が持って', '確認待ち誰', '今誰待ち', '誰がボール'].some(function(k) {
    return text.includes(k);
  });
}

// 案件の未完了タスクを担当者ごとに集計してボールホルダーリストを返す
// 戻り値: [{ name, tasks: [{content, deadline, overdue}] }, ...]  期日が近い順
function getBallHolders(projectName) {
  var sheet = getSheet('タスク管理');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var rows = sheet.getDataRange().getValues().slice(1).filter(function(r) {
    var status = String(r[5] || '');
    if (status === 'done' || status === '完了') return false;
    if (projectName && r[1] !== projectName) return false;
    return !!(r[2] && r[3]); // タスク内容と担当者が両方あるもの
  });

  // 担当者ごとにグループ化
  var holders = {};
  rows.forEach(function(r) {
    var name    = String(r[3] || '').trim();
    var content = String(r[2] || '');
    var dl      = r[4] instanceof Date ? r[4] : (r[4] ? new Date(r[4]) : null);
    var overdue = dl ? dl < today : false;
    if (!holders[name]) holders[name] = [];
    holders[name].push({ content: content, deadline: dl, overdue: overdue });
  });

  // 期日超過タスクがある担当者を先頭に、タスク数の多い順で並べる
  return Object.keys(holders).map(function(name) {
    return { name: name, tasks: holders[name] };
  }).sort(function(a, b) {
    var aOver = a.tasks.filter(function(t) { return t.overdue; }).length;
    var bOver = b.tasks.filter(function(t) { return t.overdue; }).length;
    if (bOver !== aOver) return bOver - aOver;
    return b.tasks.length - a.tasks.length;
  });
}

// ボールホルダークエリへの返信テキストを生成
function answerBallHolderQuery(text, groupId) {
  var projectName = getProjectNameByGroupId(groupId) || '';
  // テキストから案件名を補完
  if (!projectName) {
    var info = lookupProjectInfo(text);
    if (info) projectName = info.name;
  }

  var holders = getBallHolders(projectName);
  if (!holders.length) {
    return '⚽ ' + (projectName || '全案件') + 'の未完了タスクはありません。';
  }

  var lines = ['⚽ ' + (projectName ? projectName + 'の' : '') + 'ボールホルダー：'];
  holders.forEach(function(h) {
    var overCount = h.tasks.filter(function(t) { return t.overdue; }).length;
    var label = h.name + '（' + h.tasks.length + '件';
    if (overCount > 0) label += '・期日超過' + overCount + '件';
    label += '）';
    lines.push('\n👤 ' + label);
    h.tasks.slice(0, 3).forEach(function(t) {
      var dlStr = t.deadline
        ? Utilities.formatDate(t.deadline, 'Asia/Tokyo', 'M/d') + (t.overdue ? '⚠️' : '')
        : '期日未定';
      lines.push('  ・' + t.content + '（' + dlStr + '）');
    });
    if (h.tasks.length > 3) lines.push('  …他' + (h.tasks.length - 3) + '件');
  });
  return lines.join('\n');
}

function answerQuery(question) {
  var config       = getConfig();
  var today        = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日');
  var projAddrList = buildProjectAddressListText();
  var prompt  = 'あなたは株式会社WOOD BASE Fの秘書AIです。以下のデータを元に質問に簡潔に日本語で答えてください。\n今日：' + today +
    '\n\n【会社情報】\n' + buildCompanyInfoText() +
    (projAddrList ? '\n\n【案件の会社名・住所】\n' + projAddrList : '') +
    '\n\n【未完了タスク】\n' + buildTaskListText(null) +
    '\n\n【スケジュール】\n' + buildSchedListText() +
    '\n\n質問：' + question;
  return callGemini(config.GEMINI_API_KEY, prompt, 0.2) || 'うまく答えられませんでした。';
}

function getRecentLogsForUser(memberName, limit) {
  var sheet = getSheet('メッセージログ');
  if (!sheet || sheet.getLastRow() <= 1) return '';
  var data = sheet.getDataRange().getValues().slice(1);
  var name = memberName.replace('さん', '');
  var matched = data.filter(function(r) {
    return String(r[2] || '').includes(name) || String(r[3] || '').includes(name);
  }).slice(-(limit || 6));
  if (!matched.length) return '';
  return matched.map(function(r) {
    return '[' + r[0] + '] ' + r[2] + '：' + r[3];
  }).join('\n');
}

function answerQueryForMember(question, memberName) {
  var config       = getConfig();
  var today        = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日');
  var recentLogs   = getRecentLogsForUser(memberName, 6);
  var logBlock     = recentLogs
    ? '\n\n【' + memberName + 'さんの最近の会話（記憶）】\n' + recentLogs
    : '';
  var projAddrList = buildProjectAddressListText();

  var prompt =
    'あなたはWOODBASE・Fの専属メンターAIです。相手は「' + memberName + '」さんです。今日：' + today + '\n' +
    '\n【あなたの役割と姿勢】\n' +
    '・現場で働くメンバーの頼れる存在として、温かく丁寧に対応すること。\n' +
    '・メンバーが抱えるプレッシャーや悩みをまず受け止め、共感を示してから情報を伝えること。\n' +
    '・正解や指示をトップダウンで押しつけるのではなく、「いかがでしょうか？」などコーチング型の問いかけを自然に交えること。\n' +
    '・過去の会話履歴がある場合は、以前の相談内容を覚えているかのように自然に触れること。\n' +
    '・必ず丁寧な敬語（です・ます調）を使うこと。フランクな口調や「〜だよ」「〜だね」などの表現は絶対に使わないこと。\n' +
    '・回答は簡潔に。箇条書きより自然な文章を優先すること。\n' +
    logBlock +
    (projAddrList ? '\n\n【案件の会社名・住所】\n' + projAddrList : '') +
    '\n\n【' + memberName + 'さんの未完了タスク】\n' + buildTaskListText(memberName) +
    '\n\n【直近のスケジュール】\n' + buildSchedListText() +
    '\n\n' + memberName + 'さんからのメッセージ：\n' + question;

  return callGemini(config.GEMINI_API_KEY, prompt, 0.4) || 'うまくお答えできませんでした。恐れ入りますが、もう一度お試しください。';
}

// 会社情報シートをセットアップ（公式サイトから取得した実データで初期化）
// 既に値が入っている項目は上書きしない（運用中の編集を保護）
function setupCompanyInfoSheet() {
  var ss    = getSS();
  var sheet = ss.getSheetByName('会社情報');
  if (!sheet) {
    sheet = ss.insertSheet('会社情報');
    sheet.appendRow(['項目', '値']);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#A4C2F4');
    sheet.setColumnWidths(1, 1, 140);
    sheet.setColumnWidths(2, 1, 400);
  }

  // 公式サイト(woodbasef.com/pages/about)から取得した情報＋既知の岡山支店情報
  var defaults = [
    ['会社名',           '株式会社WOODBASE・F'],
    ['代表者',           '南 忠則（代表取締役）'],
    ['設立',             '2016年スタート'],
    ['従業員数',         '20名'],
    ['企業理念',         '夢現の扉を開く'],
    ['本社住所',         '〒580-0001 大阪府松原市若林1丁目4-3'],
    ['本社電話',         ''],
    ['本社FAX',          ''],
    ['岡山支店住所',     '〒700-0942 岡山県岡山市南区豊成3-21-3'],
    ['岡山支店電話',     '086-264-1005'],
    ['営業時間',         '平日 9:00〜18:00'],
    ['定休日',           '土日祝'],
    ['GW休業',           ''],
    ['お盆休業',         ''],
    ['年末年始休業',     ''],
    ['事業内容',         'オーダーメイド家具の設計・製造・販売／建築工事・内装・設備・空間プロデュース／海外事業（家具製造・販売）'],
    ['その他事業',       'ふるさと納税返礼品（オーダー家具チケット）'],
    ['ウェブサイト',     'https://woodbasef.com'],
    ['備考',             ''],
  ];

  // 既存項目を辞書化
  var existing = {};
  if (sheet.getLastRow() > 1) {
    sheet.getDataRange().getValues().slice(1).forEach(function(r, i) {
      existing[r[0]] = { row: i + 2, value: r[1] };
    });
  }

  var added = 0, updated = 0;
  defaults.forEach(function(d) {
    var key = d[0], val = d[1];
    if (existing[key]) {
      // 既存項目で値が空欄なら埋める。既に値が入っていれば保護
      if (!existing[key].value && val) {
        sheet.getRange(existing[key].row, 2).setValue(val);
        updated++;
      }
    } else {
      sheet.appendRow([key, val]);
      added++;
    }
  });
  console.log('会社情報セットアップ：新規' + added + '件 / 空欄補完' + updated + '件');
}

// タスク管理シート列: [0]=task_id, [1]=案件名, [2]=内容, [3]=担当者, [4]=期日, [5]=ステータス, [6]=作成日時, [7]=group_id
function buildTaskListText(memberName) {
  var sheet = getSheet('タスク管理');
  if (!sheet || sheet.getLastRow() <= 1) return 'なし';
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  var rows = data.filter(function(r) {
    var st = String(r[5] || '');
    if (st === 'done' || st === '完了') return false;
    if (!r[2]) return false;
    if (memberName) return String(r[3] || '').includes(memberName.replace('さん', ''));
    return true;
  });
  if (!rows.length) return 'なし';
  return rows.map(function(r, i) {
    var dl = r[4] ? Utilities.formatDate(new Date(r[4]), 'Asia/Tokyo', 'M/d') : '期日未定';
    return (i + 1) + '. [' + r[1] + '] ' + r[3] + '：' + r[2] + '（' + dl + '・' + r[5] + '）';
  }).join('\n');
}

function buildSchedListText() {
  var sheet = getSheet('スケジュール管理');
  if (!sheet || sheet.getLastRow() <= 1) return 'なし';
  var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  if (!data.length) return 'なし';
  return data.map(function(r, i) {
    var time = r[4] ? ' ' + r[4] + '〜' : '';
    var att  = r[7] ? '　参加者：' + r[7] : '';
    return (i + 1) + '. [' + r[1] + '] ' + r[2] + ' ' + r[3] + time + ' 場所：' + r[6] + att;
  }).join('\n');
}

// Gemini共通呼び出し
function callGemini(apiKey, prompt, temperature) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
  try {
    var res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: temperature || 0.2, thinkingConfig: { thinkingBudget: 0 } } }),
    });
    return geminiText(res.getContentText()) || null;
  } catch (err) {
    console.error('callGemini error:', err.message);
    return null;
  }
}

// ==========================================
// SECTION 16: 完了報告・要約
// ==========================================
function isCompletionReport(text) {
  return ['完了', '終わりました', 'できました', '終わった', '完成しました', 'やりました'].some(function(k) { return text.includes(k); });
}

function handleCompletion(text, memberName, replyToken) {
  var sheet = getSheet('タスク管理');
  if (!sheet || sheet.getLastRow() <= 1) { sendLineReply(replyToken, '現在タスクは登録されていません。'); return; }

  var data     = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  var clean    = (memberName || '').replace('さん', '').trim();
  var myTasks  = [];
  for (var i = 0; i < data.length; i++) {
    var st = String(data[i][5] || '');
    if (st === 'done' || st === '完了') continue;
    if (!data[i][2]) continue;
    if (clean && !String(data[i][3] || '').includes(clean)) continue;
    myTasks.push({ row: i + 2, content: data[i][2], project: data[i][1] });
  }

  if (!myTasks.length) { sendLineReply(replyToken, '現在、未完了のタスクはございません。'); return; }

  var target = myTasks[0];
  if (myTasks.length > 1) {
    var config = getConfig();
    var list   = myTasks.map(function(t, i) { return (i + 1) + '. [' + t.project + '] ' + t.content; }).join('\n');
    var ans    = callGemini(config.GEMINI_API_KEY, '「' + text + '」が完了報告しているタスクを1つ選び番号のみ答えてください。\n' + list, 0);
    var num    = parseInt(ans);
    if (num >= 1 && num <= myTasks.length) target = myTasks[num - 1];
  }

  sheet.getRange(target.row, 6).setValue('done');
  sendLineReply(replyToken, '✅ 完了を記録いたしました。お疲れ様でございます。\n案件：' + target.project + '\nタスク：' + target.content);
}

// ---- タスク一覧のタップ完了（quickReplyボタン） ----
function isTaskListRequest(text) {
  if (!text) return false;
  return ['残タスク', 'タスク一覧', '未完了タスク'].some(function(k) { return text.includes(k); });
}

// 未完了タスク行を取得（memberName指定でその担当分に絞る）。task_id欠損の古い行はその場で採番して書き戻す
function collectOpenTasks(memberName) {
  var sheet = getSheet('タスク管理');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  var clean = (memberName || '').replace('さん', '').trim();
  var out   = [];
  for (var i = 0; i < data.length; i++) {
    var st = String(data[i][5] || '');
    if (st === 'done' || st === '完了') continue;
    if (!data[i][2]) continue;
    if (clean && !String(data[i][3] || '').includes(clean)) continue;
    var id = String(data[i][0] || '');
    if (!id) { id = generateId(); sheet.getRange(i + 2, 1).setValue(id); }
    out.push({ taskId: id, project: String(data[i][1] || '未分類'), content: String(data[i][2] || ''), assignee: String(data[i][3] || ''), deadline: data[i][4] });
  }
  return out;
}

// 「残タスクは？」→ 番号付き一覧＋タップで完了できるボタン（quickReplyは最大13件）
function handleTaskListWithButtons(replyToken, memberName) {
  var tasks = collectOpenTasks(memberName);
  var scope = memberName ? memberName.replace('さん', '') + 'さん' : '全体';
  if (!tasks.length && memberName) { tasks = collectOpenTasks(null); scope = '全体'; }
  if (!tasks.length) { sendLineReply(replyToken, '✅ 未完了のタスクはございません。'); return; }

  var lines = ['【' + scope + 'の未完了タスク】\n下のボタンをタップすると完了にできます。'];
  tasks.forEach(function(t, i) {
    var dl = t.deadline ? Utilities.formatDate(new Date(t.deadline), 'Asia/Tokyo', 'M/d') : '期日未定';
    lines.push((i + 1) + '. [' + t.project + '] ' + t.content + '（' + (t.assignee ? t.assignee + '・' : '') + dl + '）');
  });
  if (tasks.length > 13) lines.push('※ボタンは先頭13件のみです。完了後にもう一度「残タスク」とお送りください。');
  var pageUrl = tasksPageUrl_(memberName);
  if (pageUrl) lines.push('\n📱 まとめて完了はタスク画面から：\n' + pageUrl);

  var items = tasks.slice(0, 13).map(function(t, i) {
    var label = '✅' + (i + 1) + ' ' + t.content;
    if (label.length > 20) label = label.slice(0, 19) + '…';
    return { type: 'action', action: {
      type: 'postback', label: label,
      data: 'action=task_done&id=' + encodeURIComponent(t.taskId),
      displayText: '「' + t.content.slice(0, 30) + '」を完了にします',
    } };
  });
  sendQuickReply(replyToken, lines.join('\n'), items);
}

// タスク画面（?page=tasks）のURL。DASHBOARD_TOKEN設定時は key 付き、
// who指定時はその人のタスクに初期絞り込みされる
function tasksPageUrl_(who) {
  try {
    var url = ScriptApp.getService().getUrl();
    if (!url) return '';
    var token = String(PropertiesService.getScriptProperties().getProperty('DASHBOARD_TOKEN') || '').trim();
    var q = url + '?page=tasks' + (token ? '&key=' + encodeURIComponent(token) : '');
    var name = String(who || '').replace('さん', '').trim();
    if (name) q += '&who=' + encodeURIComponent(name);
    return q;
  } catch (e) { return ''; }
}

// 各画面のURL（page='' で進捗ボード）。DASHBOARD_TOKEN設定時は key 付き。
function dashboardPageUrl_(page) {
  try {
    var url = ScriptApp.getService().getUrl();
    if (!url) return '';
    var token = String(PropertiesService.getScriptProperties().getProperty('DASHBOARD_TOKEN') || '').trim();
    var q = url + (page ? '?page=' + page : '?page=board');
    if (token) q += '&key=' + encodeURIComponent(token);
    return q;
  } catch (e) { return ''; }
}

// 「URL」「進捗ボード」等 → 各画面URLの案内を返す
function isDashboardUrlRequest(text) {
  if (!text) return false;
  var t = String(text).trim();
  if (/^(url|ｕｒｌ|ＵＲＬ|link|リンク)$/i.test(t)) return true;
  // 「進捗ボードのURL」「ボード開いて」「管理画面のリンク」等
  if (/(url|ｕｒｌ|ＵＲＬ|リンク|開き|ひらき|ちょうだい|教え|おしえ)/i.test(t) &&
      /(進捗|ボード|管理画面|ダッシュ|請求|入金|タスク画面)/.test(t)) return true;
  return false;
}

// 「ここに進捗アラート」「放置アラートをここに」等 → 送信先変更コマンド判定
function isSetProgressAlertHereRequest(text) {
  if (!text) return false;
  var t = String(text).trim();
  var hasAlert = /(放置|進捗)?(アラート|通知|リマインド)/.test(t);
  var hasHere  = /(ここ|こちら|このグループ)/.test(t);
  var hasSet   = /(に|へ|にして|送|設定|変更|登録)/.test(t);
  return hasAlert && hasHere && hasSet;
}

// このグループを放置アラート（進捗）の通知先として登録
function handleSetProgressAlertHere(replyToken, groupId) {
  if (!groupId) { sendLineReply(replyToken, 'グループ／ルーム内で送ってください（個人トークでは設定できません）。'); return; }
  var res = pmSetRoleGroupId('進捗', groupId, 'LINEで設定');
  if (!res.ok) { sendLineReply(replyToken, '設定に失敗しました：' + (res.msg || '不明なエラー')); return; }
  // スクリプトプロパティが設定されているとシートより優先されるため注意喚起
  var propOverride = String(PropertiesService.getScriptProperties().getProperty('PROGRESS_GROUP_ID') || '').trim();
  var note = propOverride
    ? '\n\n⚠️ ただし現在スクリプトプロパティ PROGRESS_GROUP_ID が設定されており、そちらが優先されます。切り替えを有効にするにはこの設定を解除してください。'
    : '';
  sendLineReply(replyToken,
    '✅ このグループを放置アラートの通知先に登録しました。\n今後の放置アラートはここに届きます。' + note);
}

// 進捗ボード／タスク／請求・入金 の各URLをまとめて返信
function handleDashboardUrlRequest(replyToken, who) {
  var board   = dashboardPageUrl_('');
  var tasks   = tasksPageUrl_(who);
  var billing = dashboardPageUrl_('billing');
  if (!board) { sendLineReply(replyToken, 'URLを取得できませんでした。ウェブアプリのデプロイ状態をご確認ください。'); return; }
  var lines = [
    '🗂 WOODBASE 各画面のURLです。',
    '',
    '📋 進捗ボード',
    board,
    '',
    '🧾 請求・入金',
    billing,
  ];
  if (tasks) { lines.push('', '✅ タスク', tasks); }
  sendLineReply(replyToken, lines.join('\n'));
}

// タップ完了 postback（action=task_done&id=<task_id>）
function handleTaskDonePostback(ev, params) {
  var taskId = decodeURIComponent(params.id || '');
  if (!taskId) { sendLineReply(ev.replyToken, 'タスクを特定できませんでした。'); return; }
  var sheet = getSheet('タスク管理');
  if (!sheet || sheet.getLastRow() <= 1) { sendLineReply(ev.replyToken, 'タスクが見つかりませんでした。'); return; }
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) !== taskId) continue;
    var st = String(data[i][5] || '');
    if (st === 'done' || st === '完了') {
      sendLineReply(ev.replyToken, '✅ このタスクは既に完了しています。\nタスク：' + data[i][2]);
      return;
    }
    sheet.getRange(i + 2, 6).setValue('done');
    sendLineReply(ev.replyToken, '✅ 完了を記録いたしました。お疲れ様でございます。\n案件：' + (data[i][1] || '未分類') + '\nタスク：' + data[i][2]);
    return;
  }
  sendLineReply(ev.replyToken, 'タスクが見つかりませんでした（削除済みの可能性がございます）。');
}

function shouldSkipExtraction(text) {
  if (!text || text.length <= 3) return true;
  if (!/[ぁ-んァ-ヶー一-龯a-zA-Z0-9]/.test(text)) return true;

  // 短い定型返事
  var SHORT_NOISE = ['了解', 'はい', 'おk', 'ok', 'OK', 'なるほど', 'ありがとう', 'お疲れ', 'お世話', 'よろしく', 'わかった'];
  if (text.length <= 10 && SHORT_NOISE.some(function(w) { return text.includes(w); })) return true;

  // 長くても内容のない返事パターン
  var NOISE_PATTERNS = [
    /^(了解|承知|わかりました|わかった|了解です|承知しました|了解しました)[!！。\s]*$/,
    /^(ありがとう|ありがとうございます|ありがとうございました)[!！。\s]*$/,
    /^(お疲れ様|お疲れさまです|お疲れ様でした)[!！。\s]*$/,
    /^(おはよう|おはようございます|こんにちは|こんばんは)[!！。\s]*$/,
    /^(はい|いいえ|そうです|そうですね|ですね)[!！。\s]*$/,
    /^(確認します|確認しました|見ました|みました)[!！。\s]*$/,
    /^(いいですね|いいね|👍|🙏|😊)[!！。\s]*$/,
  ];
  if (NOISE_PATTERNS.some(function(p) { return p.test(text.trim()); })) return true;

  return false;
}

// 同グループ・同担当者・同内容のタスクが48時間以内に登録済みか確認
function isDuplicateTask(content, assignee, groupId) {
  var sheet = getSheet('タスク管理');
  if (!sheet || sheet.getLastRow() <= 1) return false;
  var data    = sheet.getDataRange().getValues();
  var cutoff  = Date.now() - 48 * 60 * 60 * 1000;
  var normNew = String(content || '').replace(/\s/g, '').slice(0, 30);

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[7] !== groupId) continue;
    var createdAt = new Date(row[6]);
    if (isNaN(createdAt) || createdAt.getTime() < cutoff) continue;
    if (row[3] !== assignee) continue;
    var normExist = String(row[2]).replace(/\s/g, '').slice(0, 30);
    if (normNew === normExist) return true;
  }
  return false;
}

// 同グループ・同タイトル・同日付の予定が既にスケジュール管理シートにあるか確認
//   スケジュールは「重複防止」が無かったため、再送・連投で二重登録＋二重DM＋二重カレンダーが起きていた。
//   列: [0]=登録日時, [1]=案件名, [2]=予定タイトル, [3]=日付, ... [9]=グループID
function isDuplicateSchedule(title, date, groupId) {
  var sheet = getSheet('スケジュール管理');
  if (!sheet || sheet.getLastRow() <= 1) return false;
  var data    = sheet.getDataRange().getValues();
  var normNew = String(title || '').replace(/\s/g, '');
  var dnew    = String(date || '').slice(0, 10);
  if (!normNew || !dnew) return false; // タイトルか日付が無ければ判定不可（重複扱いしない）
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[9] || '') !== String(groupId)) continue;
    var rowDate = row[3] instanceof Date ? fmtDate(row[3]) : String(row[3] || '').slice(0, 10);
    if (rowDate !== dnew) continue;
    if (String(row[2] || '').replace(/\s/g, '') === normNew) return true;
  }
  return false;
}

// 後方互換
function saveChatHistory(groupId, senderName, text, timestamp) { saveMessageLog(groupId, senderName, text, timestamp); }

function summarizeChat(groupId) {
  var config  = getConfig();
  var sheet   = getSheet('メッセージログ');
  if (!sheet || sheet.getLastRow() <= 1) return '履歴がまだありません。';
  var data     = sheet.getDataRange().getValues();
  var filtered = data.slice(1).filter(function(r) { return r[1] === groupId; }).slice(-50);
  if (!filtered.length) return '履歴がまだありません。';
  var histText = filtered.map(function(r) { return '[' + r[0] + '] ' + r[2] + '：' + r[3]; }).join('\n');
  return callGemini(config.GEMINI_API_KEY,
    'WOODBASEの会話履歴を箇条書きでまとめてください。決定事項・依頼・タスク・スケジュールを中心に。前置き不要。\n\n' + histText,
    0.2) || 'まとめられませんでした。';
}

// ==========================================
// SECTION 17: リマインド・レポート
// ==========================================
function checkDeadlines() {
  var config = getConfig();
  var sheet  = getSheet('タスク管理');
  if (!sheet || sheet.getLastRow() <= 1) return;

  var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  var today = new Date(); today.setHours(0, 0, 0, 0);

  for (var i = 0; i < data.length; i++) {
    var row      = data[i];
    var project  = row[1], content  = row[2], assignee = row[3];
    var dueRaw   = row[4], status   = String(row[5] || ''), groupId = row[7];

    if (!dueRaw || status === 'done' || status === '完了') continue;

    var due  = new Date(dueRaw); due.setHours(0, 0, 0, 0);
    var diff = Math.round((due - today) / 86400000);
    var dlStr = Utilities.formatDate(due, 'Asia/Tokyo', 'M月d日');
    var msg  = null;

    if (diff === 1) {
      msg = '【明日が期日です】\n' + assignee + 'さん、明日が期日のタスクがございます。ご確認をお願いいたします。\n案件：' + project + '\nタスク：' + content;
    } else if (diff === 0) {
      msg = '【本日が期日です】\n' + assignee + 'さん、本日が期日のタスクがございます。ご確認をお願いいたします。\n案件：' + project + '\nタスク：' + content;
    } else if (diff < 0) {
      msg = '【期日を超過しています】\n' + assignee + 'さん、期日を' + Math.abs(diff) + '日超過しているタスクがございます。ご対応をお願いいたします。\n案件：' + project + '\nタスク：' + content + '\n期日：' + dlStr;
    }

    if (msg) {
      if (groupId) sendLineMessage(groupId, msg);
      if (config.INTERNAL_GROUP_ID && config.INTERNAL_GROUP_ID !== groupId) sendLineMessage(config.INTERNAL_GROUP_ID, msg);
      var uid = getMemberUserId(assignee);
      if (uid) sendLineMessage(uid, msg);
    }
  }

  // 前日スケジュールリマインド
  var schedSheet = getSheet('スケジュール管理');
  if (!schedSheet || schedSheet.getLastRow() <= 1) return;
  var tomorrow    = new Date(today); tomorrow.setDate(today.getDate() + 1);
  var tomorrowYmd = fmtDate(tomorrow);
  var schedData   = schedSheet.getRange(2, 1, schedSheet.getLastRow() - 1, 10).getValues();

  for (var j = 0; j < schedData.length; j++) {
    var sr      = schedData[j];
    var dateVal = sr[3];
    if (!dateVal) continue;
    var dateStr = dateVal instanceof Date ? fmtDate(dateVal) : String(dateVal).slice(0, 10);
    if (dateStr !== tomorrowYmd) continue;
    var schedMsg = '【明日の予定】\n案件：' + (sr[1] || '未定') + '\n予定：' + sr[2] +
      '\n日時：' + dateStr + (sr[4] ? ' ' + sr[4] + '〜' : '') +
      '\n場所：' + (sr[6] || '未定') + '\n参加者：' + (sr[7] || '未定');
    sendLineMessage(sr[9] || config.INTERNAL_GROUP_ID, schedMsg);
  }

  checkStaleTasks();
}

// ⑧ 連絡漏れ検出：タスク登録から3日経過しても未完了のタスクをリマインド
function checkStaleTasks() {
  var config  = getConfig();
  var sheet   = getSheet('タスク管理');
  if (!sheet || sheet.getLastRow() <= 1) return;

  var data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  var now     = new Date();
  var cutoff  = 3 * 24 * 60 * 60 * 1000; // 3日

  for (var i = 0; i < data.length; i++) {
    var row      = data[i];
    var project  = row[1], content = row[2], assignee = row[3];
    var status   = String(row[5] || '');
    var createdAt = new Date(row[6]);
    var groupId  = row[7];

    if (!assignee || status === 'done' || status === '完了') continue;
    if (isNaN(createdAt)) continue;
    if (now - createdAt < cutoff) continue;

    // 送信済みフラグ（スクリプトプロパティで管理）
    var flagKey = 'STALE_' + row[0];
    if (PropertiesService.getScriptProperties().getProperty(flagKey)) continue;

    var msg = '【進捗ご確認のお願い】\n' + assignee + 'さん、下記タスクの状況をご確認いただけますでしょうか。\n案件：' + (project || '未分類') +
      '\nタスク：' + content + '\n\n完了された場合は「' + content.slice(0, 15) + ' 完了」とご送信ください。';

    var uid = getMemberUserId(assignee);
    if (uid) sendLineMessage(uid, msg);
    if (groupId) sendLineMessage(groupId, '【状況確認中】' + assignee + 'さんへ：「' + content.slice(0, 20) + '」の対応状況をご確認中です。');

    PropertiesService.getScriptProperties().setProperty(flagKey, '1');
  }
}

function sendWeeklyReport() {
  var config = getConfig();
  var now    = new Date();
  var dow    = now.getDay() === 0 ? 7 : now.getDay();
  var monday = new Date(now); monday.setDate(now.getDate() - dow + 1); monday.setHours(0, 0, 0, 0);
  var sunday = new Date(monday); sunday.setDate(monday.getDate() + 6);
  var mStr   = Utilities.formatDate(monday, 'Asia/Tokyo', 'M/d');
  var sStr   = Utilities.formatDate(sunday, 'Asia/Tokyo', 'M/d');
  var mYmd   = fmtDate(monday);
  var sYmd   = fmtDate(sunday);

  var schedSheet = getSheet('スケジュール管理');
  var schedList  = 'なし';
  if (schedSheet && schedSheet.getLastRow() > 1) {
    var rows = schedSheet.getRange(2, 1, schedSheet.getLastRow() - 1, 9).getValues()
      .filter(function(r) {
        var d = r[3] instanceof Date ? fmtDate(r[3]) : String(r[3]).slice(0, 10);
        return d >= mYmd && d <= sYmd;
      })
      .map(function(r) {
        var d = r[3] instanceof Date ? fmtDate(r[3]) : String(r[3]).slice(0, 10);
        return '・' + d.slice(5).replace('-', '/') + ' [' + r[1] + '] ' + r[2];
      });
    if (rows.length) schedList = rows.join('\n');
  }

  sendLineMessage(config.INTERNAL_GROUP_ID,
    '【週次レポート ' + mStr + '〜' + sStr + '】\n\n▼ 今週の予定\n' + schedList + '\n\n▼ 未完了タスク\n' + buildTaskListText(null));
}

function archiveCompletedTasks() {
  var sheet = getSheet('タスク管理');
  if (!sheet || sheet.getLastRow() <= 1) return;

  var archSheet = getSheet('完了タスク');
  if (!archSheet) {
    archSheet = getSS().insertSheet('完了タスク');
    archSheet.appendRow(['タスクID','案件名','タスク内容','担当者','期日','ステータス','作成日時','グループID','緊急度','完了日時']);
    archSheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#B7B7B7').setFontColor('#FFFFFF');
  } else {
    // 既存シート：不足ヘッダーを補完
    var lastCol = archSheet.getLastColumn();
    var hdrs    = archSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
    ['緊急度','完了日時'].forEach(function(h){
      if (hdrs.indexOf(h) === -1) {
        archSheet.getRange(1, archSheet.getLastColumn() + 1).setValue(h);
        archSheet.getRange(1, archSheet.getLastColumn()).setFontWeight('bold').setBackground('#B7B7B7').setFontColor('#FFFFFF');
      }
    });
  }

  var data  = sheet.getDataRange().getValues();
  var toDel = [];
  var now   = fmtDT(new Date());
  for (var i = data.length - 1; i >= 1; i--) {
    var st = String(data[i][5] || '');
    if (st !== 'done' && st !== '完了') continue;
    // タスク管理は9列（緊急度含む）想定。完了日時を末尾に追加してアーカイブ
    var row = data[i].slice(0, 9);
    while (row.length < 9) row.push(''); // 緊急度欠落のレガシー行に対応
    row.push(now); // 完了日時
    archSheet.appendRow(row);
    toDel.push(i + 1);
  }
  toDel.forEach(function(r) { sheet.deleteRow(r); });
  console.log('アーカイブ: ' + toDel.length + '件');
}

// ==========================================
// SECTION 18: LINE送信
// ==========================================
function sendLineReply(replyToken, text) {
  if (!replyToken) return;
  var config = getConfig();
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method: 'post', muteHttpExceptions: true,
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + config.LINE_CHANNEL_ACCESS_TOKEN },
      payload: JSON.stringify({ replyToken: replyToken, messages: [{ type: 'text', text: text.length > 4900 ? text.slice(0, 4900) + '…' : text }] }),
    });
  } catch (err) { console.error('sendLineReply error:', err.message); }
}

function sendLineMessage(targetId, text) {
  if (!targetId) return;
  var config = getConfig();
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post', muteHttpExceptions: true,
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + config.LINE_CHANNEL_ACCESS_TOKEN },
      payload: JSON.stringify({ to: targetId, messages: [{ type: 'text', text: text.length > 4900 ? text.slice(0, 4900) + '…' : text }] }),
    });
  } catch (err) { console.error('sendLineMessage error:', err.message); }
}

// メンション付きプッシュ（textV2）。text内の {who} がメンションに置き換わる。
// mentionUserId が無い（メンバー管理に未登録等）場合は名前入りの通常テキストで送る。
function sendLineMentionMessage(targetId, text, mentionUserId, fallbackName) {
  if (!targetId) return;
  if (!mentionUserId) {
    sendLineMessage(targetId, text.split('{who}').join(fallbackName || '担当者'));
    return;
  }
  var config = getConfig();
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post', muteHttpExceptions: true,
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + config.LINE_CHANNEL_ACCESS_TOKEN },
      payload: JSON.stringify({
        to: targetId,
        messages: [{
          type: 'textV2',
          text: text.length > 4900 ? text.slice(0, 4900) + '…' : text,
          substitution: { who: { type: 'mention', mentionee: { type: 'user', userId: mentionUserId } } },
        }],
      }),
    });
  } catch (err) { console.error('sendLineMentionMessage error:', err.message); }
}

// ==========================================
// SECTION 19: ユーティリティ
// ==========================================
function generateId() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function safeParseJson(str) { try { return JSON.parse(str || '{}') || {}; } catch (e) { return {}; } }
function getProjectData() { var s = getSheet('プロジェクト管理'); return (s && s.getLastRow() > 1) ? s.getDataRange().getValues().slice(1) : []; }

// Geminiレスポンスからテキストを安全に取得（candidates未定義・空・安全フィルタ等に対応）
function geminiText(responseBody) {
  try {
    var data = JSON.parse(responseBody);
    var c    = data && data.candidates && data.candidates[0];
    if (!c) {
      // 安全フィルタ・quota超過等でcandidatesが空の場合
      console.warn('Gemini: no candidates. finishReason:', data && data.promptFeedback ? JSON.stringify(data.promptFeedback) : 'unknown');
      return '';
    }
    return (c.content && c.content.parts && c.content.parts[0] && c.content.parts[0].text) ? c.content.parts[0].text : '';
  } catch (e) {
    console.error('geminiText parse error:', e.message);
    return '';
  }
}

// ==========================================
// SECTION 20: 初期セットアップ
// ==========================================
// 安全版 setup：既存データを絶対に消さない（運用中再実行OK）
//   - シートがなければ作成
//   - シートが空ならヘッダー書込
//   - 既にヘッダーがある場合、不足列のみ右側に追加
//   - clearContents() は呼ばない
function setup() {
  var ss = getSS();

  setupSheet(ss, 'タスク管理',
    ['タスクID', '案件名', 'タスク内容', '担当者', '期日', 'ステータス', '作成日時', 'グループID', '緊急度'],
    '#4A86E8', [120, 160, 280, 100, 100, 80, 140, 160, 80]);
  var taskSheet = getSheet('タスク管理');
  try {
    taskSheet.getRange(2, 6, 1000, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['pending', 'confirmed', 'done'], true).build()
    );
  } catch (e) { /* 既存検証は無視 */ }

  setupSheet(ss, 'スケジュール管理',
    ['登録日時', '案件名', '予定タイトル', '日付', '開始時間', '終了時間', '場所', '参加者', '詳細', 'グループID'],
    '#E67C73', [140, 160, 200, 100, 80, 80, 160, 160, 200, 160]);

  setupSheet(ss, '仮タスク',
    ['バッチID', '連番', '種別', 'グループID', 'ユーザーID', '案件名', '内容/タイトル', '担当者/参加者', '期日/日時', '作成日時', '元メッセージ', '追加データ'],
    '#FFD966', [120, 50, 70, 160, 160, 140, 280, 120, 140, 140, 300, 200]);

  setupSheet(ss, 'メンバー管理',
    ['名前', 'LINE ユーザーID', '役割', '備考', 'Googleメール'],
    '#34A853', [120, 220, 120, 200, 220]);
  var mem = getSheet('メンバー管理');
  if (mem.getLastRow() <= 1) {
    mem.appendRow(['濱田', 'Uxxxxxxxxxx', '社内', '', '']);
    mem.appendRow(['織田', 'Uxxxxxxxxxx', '社内', '', '']);
  }

  // 案件ID列＝projects と二面一体化するリンクキー（フェーズ1）。右端へ非破壊追加。
  setupSheet(ss, 'プロジェクト管理',
    ['略称', '正式名称', 'グループID（施主）', 'グループID（業者）', 'ステータス', '備考', '会社名', '住所', '案件ID'],
    '#F6B26B', [80, 200, 180, 180, 80, 200, 150, 250, 100]);
  var proj = getSheet('プロジェクト管理');
  if (proj.getLastRow() <= 1) proj.appendRow(['雨晴れ', '雨のち晴れクリニック', '', '', '進行中', '']);

  setupSheet(ss, '案件メモ',
    ['登録日時', '案件名', '内容', '発言者', 'グループID'],
    '#B4A7D6', [140, 160, 400, 100, 160]);

  setupSheet(ss, 'メッセージログ',
    ['日時', 'グループID', '送信者', 'メッセージ'],
    '#9FC5E8', [140, 200, 100, 400]);

  setupSheet(ss, '完了タスク',
    ['タスクID', '案件名', 'タスク内容', '担当者', '期日', 'ステータス', '作成日時', 'グループID', '緊急度', '完了日時'],
    '#B7B7B7', [120, 160, 280, 100, 100, 80, 140, 160, 80, 140]);

  // プロジェクト進捗管理シート（projects / project_logs / 他）を非破壊で作成
  try { pmSetupSheets(); } catch (e) { console.error('pmSetupSheets:', e.message); }

  // スプレッドシートの表示整頓（ダッシュボード・表示/非表示・projects装飾）。冪等。
  try { pmOrganizeSpreadsheet(); } catch (e) { console.error('pmOrganizeSpreadsheet:', e.message); }

  try {
    SpreadsheetApp.getUi().alert(
      '✅ セットアップ完了（既存データは保護されています）\n\n次の手順：\n' +
      '1. メンバー管理シートにLINEユーザーIDを入力\n' +
      '2. プロジェクト管理シートに案件情報を入力\n' +
      '3. スクリプトプロパティに INTERNAL_GROUP_ID / DRIVE_FOLDER_ID / LINE_CHANNEL_SECRET を追加\n' +
      '4. createTriggers を実行'
    );
  } catch (e) { /* UIなし環境（トリガー実行等）はスキップ */ }
}

// 旧式の破壊的セットアップ（既存データ全消去）
// 通常は実行しないでください。新規環境の初期化用です。
function resetAllSheetsDangerously() {
  if (typeof SpreadsheetApp.getUi === 'function') {
    var ui = SpreadsheetApp.getUi();
    var res = ui.alert(
      '⚠️ 警告：破壊的初期化',
      '全シートの内容を消去します。本当に実行しますか？\n（運用データが完全に失われます）',
      ui.ButtonSet.YES_NO
    );
    if (res !== ui.Button.YES) { console.log('破壊的初期化をキャンセル'); return; }
  }
  var ss = getSS();
  var defs = [
    ['タスク管理', ['タスクID','案件名','タスク内容','担当者','期日','ステータス','作成日時','グループID','緊急度']],
    ['スケジュール管理', ['登録日時','案件名','予定タイトル','日付','開始時間','終了時間','場所','参加者','詳細','グループID']],
    ['仮タスク', ['バッチID','連番','種別','グループID','ユーザーID','案件名','内容/タイトル','担当者/参加者','期日/日時','作成日時','元メッセージ','追加データ']],
    ['メンバー管理', ['名前','LINE ユーザーID','役割','備考','Googleメール']],
    ['プロジェクト管理', ['略称','正式名称','グループID（施主）','グループID（業者）','ステータス','備考','会社名','住所']],
    ['案件メモ', ['登録日時','案件名','内容','発言者','グループID']],
    ['メッセージログ', ['日時','グループID','送信者','メッセージ']],
    ['完了タスク', ['タスクID','案件名','タスク内容','担当者','期日','ステータス','作成日時','グループID','緊急度','完了日時']],
  ];
  defs.forEach(function(d){
    var sh = ss.getSheetByName(d[0]);
    if (!sh) sh = ss.insertSheet(d[0]);
    sh.clearContents();
    sh.appendRow(d[1]);
  });
  console.log('破壊的初期化を実行しました');
}

// シート作成・ヘッダー追加（既存データを絶対に削除しない）
function setupSheet(ss, name, headers, color, widths) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    // 新規作成：ヘッダー書込・装飾
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    var hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setFontWeight('bold').setBackground(color).setFontColor('#FFFFFF').setHorizontalAlignment('center');
    if (widths) widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
    return;
  }
  // 既存シート：ヘッダー存在チェック
  if (sheet.getLastRow() === 0) {
    // 完全空 → ヘッダーだけ追加
    sheet.appendRow(headers);
    var hr2 = sheet.getRange(1, 1, 1, headers.length);
    hr2.setFontWeight('bold').setBackground(color).setFontColor('#FFFFFF').setHorizontalAlignment('center');
    if (widths) widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
    return;
  }
  // 既存ヘッダーあり → 不足列のみ右側に追加
  var lastCol     = sheet.getLastColumn();
  var curHeaders  = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
  headers.forEach(function(h){
    if (curHeaders.indexOf(h) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
      var newCol = sheet.getLastColumn();
      sheet.getRange(1, newCol).setFontWeight('bold').setBackground(color).setFontColor('#FFFFFF').setHorizontalAlignment('center');
      curHeaders.push(h);
    }
  });
}

// ==========================================
// SECTION 21: トリガー
// ==========================================
function createTriggers() {
  var fns = ['checkDeadlines', 'sendWeeklyReport', 'archiveCompletedTasks', 'sendDailyScheduleNotification'];
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return fns.indexOf(t.getHandlerFunction()) !== -1; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('checkDeadlines').timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger('sendWeeklyReport').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  ScriptApp.newTrigger('archiveCompletedTasks').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(23).create();
  ScriptApp.newTrigger('sendDailyScheduleNotification').timeBased().everyDays(1).atHour(9).create();

  // プロジェクト進捗管理のアラート（放置/請求漏れ/入金未確認）を毎朝9時に登録
  try { pmSetupTriggers(); } catch (e) { console.error('pmSetupTriggers:', e.message); }

  SpreadsheetApp.getUi().alert('完了：\n・毎朝9時 タスクリマインド＋前日通知\n・毎朝9時 日次予定通知（進捗管理グループ）\n・毎朝9時 案件アラート（放置/請求漏れ/入金未確認）\n・毎週月曜8時 週次レポート\n・毎週日曜23時 完了タスクアーカイブ');
}

// ==========================================
// SECTION 22: テスト
// ==========================================

// カレンダー接続確認 + テストイベント作成
function testCalendarConnection() {
  var config = getConfig();
  if (!config.CALENDAR_ID) { console.error('CALENDAR_ID が未設定です'); return; }
  console.log('CALENDAR_ID:', config.CALENDAR_ID);
  var calendar = CalendarApp.getCalendarById(config.CALENDAR_ID);
  if (!calendar) { console.error('カレンダーが見つかりません'); return; }
  console.log('カレンダー名:', calendar.getName());

  // 今日にテストイベントを作成
  var today = new Date();
  var start = new Date(today); start.setHours(12, 0, 0);
  var end   = new Date(today); end.setHours(12, 30, 0);
  try {
    var ev = calendar.createEvent('【テスト】LINE秘書AI接続確認', start, end, { description: 'テスト用。削除してください。' });
    console.log('✅ テストイベント作成成功:', ev.getTitle(), ev.getStartTime());
    ev.deleteEvent(); // すぐ削除
    console.log('✅ テストイベント削除済み');
  } catch (err) {
    console.error('❌ イベント作成失敗:', err.message);
  }
}

function testGetMemberGoogleEmail() {
  var sheet = getSheet('メンバー管理');
  if (!sheet) { console.error('メンバー管理シートがありません'); return; }
  var colIdx = getGoogleEmailColIdx(sheet);
  console.log('Googleメール列インデックス(0-based):', colIdx, colIdx === -1 ? '（列未作成）' : '');
  console.log('濱田:', getMemberGoogleEmail('濱田'));
  console.log('織田:', getMemberGoogleEmail('織田さん'));
}

function testAddTaskToCalendarWithGuest() {
  addTaskToCalendar({
    project_name: 'テスト',
    content: 'Googleカレンダー連携テスト',
    assignee: '濱田',
    due_date: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'),
    status: 'confirmed',
  });
}

// Gemini API疎通確認（実行してHTTPステータスと生レスポンスを確認）
function testGeminiRaw() {
  var config = getConfig();
  if (!config.GEMINI_API_KEY) { console.error('GEMINI_API_KEYが設定されていません'); return; }

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + config.GEMINI_API_KEY;
  var res = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ contents: [{ parts: [{ text: 'こんにちは' }] }], generationConfig: { thinkingConfig: { thinkingBudget: 0 } } }),
  });

  console.log('HTTPステータス:', res.getResponseCode());
  console.log('レスポンス:', res.getContentText().slice(0, 800));
}

// ルールベースフィルタの動作確認
function testFilter() {
  var cases = [
    ['了解です！', false],
    ['ありがとうございます', false],
    ['明日田中さんに連絡しとく', true],
    ['来週の火曜に打ち合わせしよう', true],
    ['見積もりやっておきます', true],
    ['現場確認お願いします', true],
  ];
  cases.forEach(function(c) {
    var result = ruleBasedFilter(c[0]);
    console.log((result === c[1] ? '✅' : '❌') + ' "' + c[0] + '" → ' + result + ' (期待:' + c[1] + ')');
  });
}

// 案件識別の動作確認
function testIdentifyProject() {
  // プロジェクト管理シートに「雨のち晴れクリニック」（略称：雨晴れ）が登録済み前提
  var cases = [
    '雨晴れの件で田中さんに連絡します',
    '来週クリニックの現場確認します',
    '山田邸の見積もりやっておきます', // 未登録 → 未分類
    '明日打ち合わせしよう',            // 案件なし → 未分類
  ];
  cases.forEach(function(text) {
    var proj = identifyProject(text, 'dummy-group', '');
    console.log('"' + text + '" → ' + proj.name + '（信頼度:' + proj.confidence + '）');
  });
}

// Gemini抽出のフリーテキスト対応確認
function testExtractFreeText() {
  var cases = [
    '明日田中さんに連絡しとく',
    '来週の火曜14時から現場確認しよう',
    '見積もり週末までにやっておきます',
    '打合せ5月10日10時ね、場所は事務所で',
  ];
  cases.forEach(function(msg) {
    var result = extractWithGemini(msg, 'test-group', new Date(), '濱田');
    console.log('入力: "' + msg + '"');
    console.log('結果:', JSON.stringify({ hint: result.projectNameHint, tasks: result.tasks, schedules: result.schedules }, null, 2));
    console.log('---');
  });
}

// プロジェクト候補抽出のテスト（LINE送信なし）
function testProjectCandidates() {
  var cases = [
    'WOODBASE LP制作の件で田中さんに連絡します',
    'WOODBASE LINE秘書Botのバグ対応お願いします',
    'WOODBASEの保守対応進めます',
    '雨晴れの件で打ち合わせしよう',
    '全く関係ない話題のメッセージ',
  ];
  cases.forEach(function(text) {
    var cands = getProjectCandidates(text, 'test-group', '', 5);
    console.log('入力: "' + text + '"');
    if (!cands.length) {
      console.log('  → 候補なし（新規作成プロンプトに進む）');
    } else {
      cands.forEach(function(c, i) {
        console.log('  ' + (i + 1) + '. ' + c.name + ' (信頼度:' + c.confidence + ' / ' + c.reason + ')');
      });
    }
    console.log('---');
  });
}

// 新規プロジェクト作成フローのドライラン（シート書き込みあり・LINE送信なし）
function testNewProjectFlow() {
  var text    = '新案件の打ち合わせを明日10時から事務所で田中さんとやる';
  var groupId = 'test-group-newproj';
  var userId  = 'test-user-newproj';

  var extracted = extractWithGemini(text, groupId, new Date(), '濱田');
  console.log('抽出:', JSON.stringify({ tasks: extracted.tasks, schedules: extracted.schedules }, null, 2));

  var batchId = storePending(extracted, '', groupId, userId, text);
  console.log('仮タスク保存 batchId:', batchId);

  // ユーザーが「新規プロジェクト作成」を選んだと仮定
  setNewProjectMode(userId, batchId);
  console.log('NEWPROJモード設定 → 取得:', getNewProjectMode(userId));

  // ユーザーが新規名「テスト新規案件」を入力したと仮定
  // 実環境ではsendLineReply経由なのでここではコメント化:
  // finalizeNewProject('テスト新規案件', batchId, groupId, userId, null);

  // ドライランのため仮タスクをクリーンアップ
  clearNewProjectMode(userId);
  deletePendingItems(batchId);
  console.log('テスト終了（仮タスクは削除済み）');
}

// プロジェクト管理シートに会社名・住所列を直接追加（setup()を使わない単独版）
function addProjectColumns() {
  var sheet = getSheet('プロジェクト管理');
  if (!sheet) { console.log('ERROR: プロジェクト管理シートが見つかりません'); return; }
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  console.log('現在のヘッダー: ' + headers.join(', '));
  [['会社名', 150], ['住所', 250]].forEach(function(def) {
    var colName = def[0], width = def[1];
    if (headers.indexOf(colName) === -1) {
      var newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(colName)
        .setFontWeight('bold').setBackground('#F6B26B').setFontColor('#FFFFFF').setHorizontalAlignment('center');
      sheet.setColumnWidth(newCol, width);
      console.log('✅ ' + colName + ' を列' + newCol + 'に追加しました');
    } else {
      console.log('ℹ️ ' + colName + ' は既に存在します（スキップ）');
    }
  });
}

// === セキュリティ・データ保護関連テスト ===

// setup() が既存データを消さないことを確認（破壊テスト：テスト後にデータ復元される）
function testSafeSetupDoesNotDeleteData() {
  var ss = getSS();
  var sh = ss.getSheetByName('タスク管理');
  if (!sh || sh.getLastRow() < 2) {
    console.log('タスク管理にデータがないため、別の方法で確認します');
    var rowsBefore = sh ? sh.getLastRow() : 0;
    setup();
    var rowsAfter = getSheet('タスク管理').getLastRow();
    console.log(rowsAfter >= rowsBefore ? '✅ データ件数維持' : '❌ データ消失！');
    return;
  }
  var before = sh.getLastRow();
  var sampleRow = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0];
  console.log('実行前: ' + before + '行 / サンプル行先頭=' + sampleRow[0]);
  setup();
  var after = sh.getLastRow();
  var sampleAfter = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0];
  console.log('実行後: ' + after + '行 / サンプル行先頭=' + sampleAfter[0]);
  console.log((before === after && String(sampleRow[0]) === String(sampleAfter[0])) ? '✅ データ完全保持' : '❌ データに変化あり');
}

// タスク管理・完了タスクに緊急度列があるか
function testTaskHeadersIncludeUrgency() {
  ['タスク管理', '完了タスク'].forEach(function(name){
    var sh = getSheet(name);
    if (!sh) { console.log('❌ ' + name + ' シート未存在'); return; }
    var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim(); });
    console.log(name + ' ヘッダー:', hdrs.join(' / '));
    console.log(hdrs.indexOf('緊急度') !== -1 ? '✅ 緊急度あり' : '❌ 緊急度なし');
    if (name === '完了タスク') console.log(hdrs.indexOf('完了日時') !== -1 ? '✅ 完了日時あり' : '❌ 完了日時なし');
  });
}

// postback 分岐ロジック確認（実呼び出しなし）
function testPostbackRouting() {
  var cases = [
    { data: 'action=doc_summary&p=foo', needBatch: false },
    { data: 'action=link_group&p=foo&g=Cxxxx', needBatch: false },
    { data: 'action=cancel&key=abc', needBatch: false },
    { data: 'action=ignore_all&group=Cxxxx', needBatch: false },
    { data: 'action=register&batch=abc', needBatch: true },
    { data: 'action=set_project&batch=abc&p=foo', needBatch: true },
    { data: 'action=new_project&batch=abc', needBatch: true },
  ];
  cases.forEach(function(c){
    var p = parseParams(c.data);
    var hasBatch = !!p.batch;
    console.log('action=' + p.action + ' batch必要:' + c.needBatch + ' / batchあり:' + hasBatch);
  });
}

// syncAllGroupMembers が分割関数を両方呼ぶか
function testSyncAllGroupMembersNoDuplicate() {
  var hasLogs     = typeof syncAllGroupMembersFromLogs     === 'function';
  var hasProjects = typeof syncAllGroupMembersFromProjects === 'function';
  var hasCombined = typeof syncAllGroupMembers             === 'function';
  console.log('syncAllGroupMembersFromLogs:',     hasLogs     ? '✅' : '❌');
  console.log('syncAllGroupMembersFromProjects:', hasProjects ? '✅' : '❌');
  console.log('syncAllGroupMembers:',             hasCombined ? '✅' : '❌');
}

// 案件識別→確認フローの統合テスト（シート書き込みあり・LINE送信なし）
function testFullFlow() {
  var text      = '明日田中さんに見積もり送っておくよ。あと来週月曜10時から打ち合わせしよう';

  var sender    = '濱田';
  var groupId   = 'test-group';
  var userId    = 'test-user';

  var extracted = extractWithGemini(text, groupId, new Date(), sender);
  console.log('抽出:', JSON.stringify({ hint: extracted.projectNameHint, tasks: extracted.tasks, schedules: extracted.schedules }, null, 2));

  var proj      = identifyProject(text, groupId, extracted.projectNameHint);
  console.log('案件識別:', proj.name, '信頼度:', proj.confidence);

  var batchId   = storePending(extracted, proj.name, groupId, userId, text);
  console.log('仮タスク保存完了 batchId:', batchId);
  console.log('信頼度' + proj.confidence + ' → ' + (proj.confidence >= 70 ? '確認UI' : '案件選択UI') + 'を表示（LINE送信はスキップ）');
}


function listSheetUsage() {
  getSS().getSheets().forEach(function(s){
    console.log(s.getName(), 'rows:', s.getLastRow(), 'cols:', s.getLastColumn());
  });
}

function renameLegacyChatHistory() {
  var ss = getSS();
  var oldName = '会話履歴';
  var newName = '_OLD_会話履歴_20260518';
  var sh = ss.getSheetByName(oldName);
  if (!sh) { console.log('対象シート「' + oldName + '」は既にありません'); return; }
  if (ss.getSheetByName(newName)) { console.log('既に「' + newName + '」が存在します。中断'); return; }
  sh.setName(newName);
  console.log('リネーム完了: ' + oldName + ' → ' + newName + ' (rows=' + sh.getLastRow() + ')');
}

// メッセージログのドライラン分析（変更なし）
function analyzeMessageLog() {
  var sh = getSheet('メッセージログ');
  if (!sh || sh.getLastRow() <= 1) { console.log('ログなし'); return; }
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  var now = new Date();
  var buckets = { '7日以内': 0, '8〜30日': 0, '31〜60日': 0, '61〜90日': 0, '90日超': 0, '日付不明': 0 };
  data.forEach(function(r){
    var v = r[0];
    var d = v instanceof Date ? v : new Date(String(v).replace(/-/g, '/'));
    if (isNaN(d)) { buckets['日付不明']++; return; }
    var diff = Math.floor((now - d) / 86400000);
    if (diff <= 7) buckets['7日以内']++;
    else if (diff <= 30) buckets['8〜30日']++;
    else if (diff <= 60) buckets['31〜60日']++;
    else if (diff <= 90) buckets['61〜90日']++;
    else buckets['90日超']++;
  });
  console.log('総行数:', data.length);
  Object.keys(buckets).forEach(function(k){ console.log('  ' + k + ': ' + buckets[k] + '行'); });
  console.log('→ 60日超を退避すると ' + (buckets['61〜90日'] + buckets['90日超']) + '行が移動対象');
}

// 60日より古いメッセージログをアーカイブシートに退避
function archiveOldMessageLog() {
  var ss = getSS();
  var sh = ss.getSheetByName('メッセージログ');
  if (!sh || sh.getLastRow() <= 1) { console.log('ログなし'); return; }

  var THRESHOLD_DAYS = 60;
  var archName = 'メッセージログ_アーカイブ';
  var arch = ss.getSheetByName(archName);
  if (!arch) {
    arch = ss.insertSheet(archName);
    arch.appendRow(['日時', 'グループID', '送信者', 'メッセージ']);
    arch.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#CCCCCC').setFontColor('#FFFFFF');
  }

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  var now = new Date();
  var cutoffMs = THRESHOLD_DAYS * 86400000;
  var toArchive = [];
  var toDeleteRows = [];

  for (var i = 0; i < data.length; i++) {
    var v = data[i][0];
    var d = v instanceof Date ? v : new Date(String(v).replace(/-/g, '/'));
    if (isNaN(d)) continue;
    if ((now - d) > cutoffMs) {
      toArchive.push(data[i]);
      toDeleteRows.push(i + 2); // シート上の行番号
    }
  }

  if (!toArchive.length) { console.log('退避対象なし（60日超のログなし）'); return; }

  // アーカイブシートへ一括追記
  arch.getRange(arch.getLastRow() + 1, 1, toArchive.length, 4).setValues(toArchive);

  // 元シートから後ろから削除（行番号ズレ防止）
  for (var j = toDeleteRows.length - 1; j >= 0; j--) {
    sh.deleteRow(toDeleteRows[j]);
  }

  console.log('退避完了: ' + toArchive.length + '行 → ' + archName);
  console.log('メッセージログ残: ' + (sh.getLastRow() - 1) + '行');
}

// ==========================================
// SECTION 23: 日次予定通知（毎朝9時 → 進捗管理グループ）
// 送信先: ① Script Property DAILY_NOTIFY_GROUP_ID ② 社内グループシート（役割「日次通知」） ③ 既定値
// ==========================================

var DAILY_NOTIFY_GROUP_ID_DEFAULT = 'Cc9886c70f353c706b76991070c3695bc';

function getDailyNotifyGroupId() {
  var byRole = (typeof pmRoleGroupId === 'function') ? pmRoleGroupId('日次通知') : '';
  return PropertiesService.getScriptProperties().getProperty('DAILY_NOTIFY_GROUP_ID') || byRole || DAILY_NOTIFY_GROUP_ID_DEFAULT;
}

// 診断：毎朝9時の自動送信が「実際に動く状態か」を確認する（送信はしない）。
// ① 登録済みトリガー一覧（毎朝9時のものがあるか） ② 送信先グループID
// ③ 今送られる内容のプレビュー（予定＋進捗リマインド）
// GASエディタで diagnoseDailyNotify を実行し、実行ログを共有してください。
function diagnoseDailyNotify() {
  console.log('===== 毎朝9時 自動送信 診断 =====');

  // ① トリガー
  var triggers = ScriptApp.getProjectTriggers();
  console.log('登録トリガー数: ' + triggers.length);
  var found9 = [];
  triggers.forEach(function(t) {
    var fn   = t.getHandlerFunction();
    var type = String(t.getEventType());
    var info = 'fn=' + fn + ' / type=' + type;
    try {
      if (t.getEventType() === ScriptApp.EventType.CLOCK) {
        // atHour は読めないので、ハンドラ名で判定
        info += ' (時間主導)';
      }
    } catch (e) {}
    console.log(' - ' + info);
    if (fn === 'sendDailyScheduleNotification' || fn === 'checkDeadlines') found9.push(fn);
  });
  console.log('→ 日次通知系トリガー: ' + (found9.length ? found9.join(', ') : '【なし＝自動送信は動いていません】'));

  // ② 送信先
  var props = PropertiesService.getScriptProperties();
  console.log('\n送信先 DAILY_NOTIFY_GROUP_ID(プロパティ): ' + (props.getProperty('DAILY_NOTIFY_GROUP_ID') || '(未設定→デフォルト使用)'));
  console.log('実際の送信先(解決後): ' + getDailyNotifyGroupId());

  // ③ 送信内容プレビュー（送信はしない）
  console.log('\n----- ① 本日の予定（プレビュー） -----');
  try { console.log(buildDailyScheduleText()); } catch (e) { console.error('buildDailyScheduleText error: ' + e.message); }
  console.log('\n----- ② 進捗管理表 未完了リマインド（プレビュー） -----');
  try { console.log(buildDailyProgressReminderText() || '（未完了なし）'); } catch (e) { console.error('buildDailyProgressReminderText error: ' + e.message); }

  console.log('\n===== 診断ここまで =====');
}

// 本日の予定テキストを組み立てる
// 出力形式:
//   ●本日M/D(曜)の予定
//   ～HH:MM
//   ・予定タイトル
//   ...
function buildDailyScheduleText() {
  var today    = new Date();
  var todayYmd = fmtDate(today);
  var WDAYS    = ['日', '月', '火', '水', '木', '金', '土'];
  var dateLabel = (today.getMonth() + 1) + '/' + today.getDate() + '(' + WDAYS[today.getDay()] + ')';

  var lines = ['●本日' + dateLabel + 'の予定'];

  // 時刻セルを「H:mm」に整形する。
  // スプレッドシートの「時刻のみ」セルは内部的に 1899/12/30 のDate値になるため、
  // String() するとDateのフルテキストになりLINEに化けて出る。Dateは時刻だけ取り出す。
  function fmtTimeCell(v) {
    if (v == null || v === '') return '';
    if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'H:mm');
    return String(v).trim().replace(/：/g, ':');
  }

  // スケジュール管理から本日の予定を取得（開始時間でソート）
  // 列: [0]=登録日時, [1]=案件名, [2]=予定タイトル, [3]=日付, [4]=開始時間, [5]=終了時間, [6]=場所, [7]=参加者
  var schedSheet = getSheet('スケジュール管理');
  var events = [];
  if (schedSheet && schedSheet.getLastRow() > 1) {
    events = schedSheet.getDataRange().getValues().slice(1)
      .filter(function(r) {
        var d = r[3] instanceof Date ? fmtDate(r[3]) : String(r[3]).slice(0, 10);
        return d === todayYmd;
      })
      .sort(function(a, b) {
        // 時刻は "H:mm" に整形してからゼロ埋め比較（8:00 と 13:00 を正しく並べる）
        function sortKey(r) {
          var s = fmtTimeCell(r[4]) || fmtTimeCell(r[5]);
          var m = s.match(/^(\d{1,2}):(\d{2})/);
          return m ? (('0' + m[1]).slice(-2) + m[2]) : 'zzzz';
        }
        var ta = sortKey(a), tb = sortKey(b);
        return ta < tb ? -1 : ta > tb ? 1 : 0;
      });
  }

  // タスク管理から本日期限の未完了タスクを取得
  // 列: [0]=task_id, [1]=案件名, [2]=内容, [3]=担当者, [4]=期日, [5]=ステータス, ..., [8]=緊急度
  var taskSheet = getSheet('タスク管理');
  var tasks = [];
  if (taskSheet && taskSheet.getLastRow() > 1) {
    tasks = taskSheet.getDataRange().getValues().slice(1)
      .filter(function(r) {
        if (!r[2]) return false;
        if (r[5] === 'done' || r[5] === '完了') return false;
        var dl = r[4] ? (r[4] instanceof Date ? fmtDate(r[4]) : String(r[4]).slice(0, 10)) : '';
        return dl === todayYmd;
      });
  }

  if (!events.length && !tasks.length) {
    lines.push('本日の予定・タスクはありません。');
    return lines.join('\n');
  }

  // 予定を「終了時間」でタイムブロックにグループ化して出力
  // 終了時間あり → 「～HH:MM」、開始時間のみ → 「HH:MM〜」、なし → 時間なしでそのまま
  var groups = [];
  var labelIndex = {};

  events.forEach(function(r) {
    var endTime   = fmtTimeCell(r[5]);
    var startTime = fmtTimeCell(r[4]);
    var title     = String(r[2] || '').trim();
    var project   = String(r[1] || '').trim();
    var location  = String(r[6] || '').trim();

    var timeLabel = endTime   ? ('～' + endTime)
                  : startTime ? (startTime + '〜')
                  : '';

    var display = title;
    if (project && project !== '未分類') display = project + ' ' + display;
    if (location) display += '（' + location + '）';

    if (timeLabel in labelIndex) {
      groups[labelIndex[timeLabel]].items.push(display);
    } else {
      labelIndex[timeLabel] = groups.length;
      groups.push({ label: timeLabel, items: [display] });
    }
  });

  groups.forEach(function(g) {
    if (g.label) lines.push(g.label);
    g.items.forEach(function(it) { lines.push('・' + it); });
  });

  // 本日期限タスクをブロック末尾に追記
  if (tasks.length) {
    if (events.length) lines.push('');
    lines.push('【本日期限のタスク】');
    tasks.forEach(function(r) {
      var prefix   = (String(r[8] || '') === '高') ? '🚨 ' : '・';
      var assignee = r[3] ? '（' + r[3] + '）' : '';
      lines.push(prefix + r[2] + assignee);
    });
  }

  return lines.join('\n');
}

// 進捗管理表の未完了一覧テキストを組み立てる（毎朝リマインド用）
function buildDailyProgressReminderText() {
  try {
    var answer = answerProgressQuestion('未完了全部教えて');
    return answer || null;
  } catch (err) {
    console.error('buildDailyProgressReminderText error:', err.message);
    return null;
  }
}

// 毎朝9時に実行されるエントリーポイント
// 毎朝9時：善波グループ（DAILY_NOTIFY_GROUP_ID）へ進捗管理表ダイジェストを送る。
// 内容＝進捗管理表の「今月＋翌月 施工分」で未完了の項目（buildSenbaProgressDigest）。
// ※ 進捗管理表とタスク管理は別物。タスクの期日通知は sendTaskDeadlineNotifications が担当。
function sendDailyScheduleNotification() {
  try {
    var groupId = getDailyNotifyGroupId();

    var digest = buildSenbaProgressDigest();
    if (digest) {
      sendLineMessage(groupId, digest);
    } else {
      sendLineMessage(groupId, '本日時点で、今月・翌月施工分の未完了項目はありません。');
    }

    console.log('善波グループへ進捗ダイジェストを送信しました → ' + groupId);
  } catch (err) {
    console.error('sendDailyScheduleNotification error:', err.message);
  }
}

// 毎朝9時：本日／明日が期日のタスクを、担当者本人＋その案件グループにだけ通知する。
// 社内全体グループ（INTERNAL_GROUP_ID）への一斉送信はしない。タスク管理シートのみ参照。
// dryRun=true のときは送信せず、対象と送信先をログ出力するだけ（テスト用）。
function sendTaskDeadlineNotifications(dryRun) {
  var sheet = getSheet('タスク管理');
  if (!sheet || sheet.getLastRow() <= 1) { console.log('タスク管理：対象なし'); return; }

  var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var sent  = 0, hit = 0;

  for (var i = 0; i < data.length; i++) {
    var row     = data[i];
    var project = row[1], content = row[2], assignee = row[3];
    var dueRaw  = row[4], status  = String(row[5] || ''), groupId = row[7];

    if (!content || !dueRaw) continue;
    if (status === 'done' || status === '完了') continue;

    var due = new Date(dueRaw);
    if (isNaN(due.getTime())) continue;
    due.setHours(0, 0, 0, 0);
    var diff = Math.round((due - today) / 86400000);

    var msg = null;
    if (diff === 1) {
      msg = '【明日が期日です】\n' + (assignee || '') + 'さん、明日が期日のタスクがございます。ご確認をお願いいたします。\n案件：' + project + '\nタスク：' + content;
    } else if (diff === 0) {
      msg = '【本日が期日です】\n' + (assignee || '') + 'さん、本日が期日のタスクがございます。ご確認をお願いいたします。\n案件：' + project + '\nタスク：' + content;
    }
    if (!msg) continue;
    hit++;

    // 送信先＝案件グループ＋担当者本人のDM（重複排除）。社内全体には送らない。
    var targets = {};
    if (groupId) targets[groupId] = true;
    String(assignee || '').split(/[、,・／\/\s]+/).forEach(function(name) {
      var nm = name.trim();
      if (!nm) return;
      var uid = getMemberUserId(nm);
      if (uid) targets[uid] = true;
    });

    var tos = Object.keys(targets);
    if (dryRun) {
      console.log('[対象] diff=' + diff + ' / ' + project + ' / ' + content + ' → ' + (tos.length ? tos.join(', ') : '送信先なし(担当者未登録・グループID空)'));
    } else {
      tos.forEach(function(to) { sendLineMessage(to, msg); sent++; });
    }
  }

  console.log('タスク期日通知: 該当' + hit + '件 / ' + (dryRun ? '（ドライラン：送信なし）' : '送信' + sent + '件'));
}

// 毎朝9時の自動送信を有効化（重複は消してから1つずつ登録）。1回実行すればOK。
function setupDailyTriggers() {
  var fns = ['sendDailyScheduleNotification', 'sendTaskDeadlineNotifications'];
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return fns.indexOf(t.getHandlerFunction()) !== -1; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('sendDailyScheduleNotification').timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger('sendTaskDeadlineNotifications').timeBased().everyDays(1).atHour(9).create();

  console.log('✅ 毎朝9時トリガー登録完了: ' + fns.join(', '));
}

// 送信せず中身だけ確認するテスト（GASエディタで実行）。
function testDailyOutputs() {
  console.log('===== ① 善波グループ：進捗ダイジェスト =====');
  console.log(buildSenbaProgressDigest() || '（今月・翌月施工分の未完了なし）');
  console.log('\n===== ② タスク期日通知（ドライラン：送信しません） =====');
  sendTaskDeadlineNotifications(true);
}

// テスト用：LINE送信せずコンソールに出力して確認
function testDailyScheduleNotification() {
  console.log('=== ① 本日の予定 ===');
  console.log(buildDailyScheduleText());
  console.log('=== ② 進捗リマインド ===');
  console.log(buildDailyProgressReminderText() || '（未完了なし）');
}

// 送信テスト：グループIDとトークンが正しいか確認（レスポンスコードをログ出力）
function testSendToProgressGroup() {
  var config  = getConfig();
  var groupId = getDailyNotifyGroupId();
  console.log('送信先グループID:', groupId);
  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', muteHttpExceptions: true,
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + config.LINE_CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify({ to: groupId, messages: [{ type: 'text', text: '📋 テスト送信です' }] }),
  });
  console.log('HTTPステータス:', res.getResponseCode());
  console.log('レスポンス:', res.getContentText().slice(0, 300));
}

// トリガーを登録（GASエディタから1回実行するだけでOK）
function setupDailyScheduleTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'sendDailyScheduleNotification'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('sendDailyScheduleNotification')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  console.log('✅ 毎朝9時 日次予定通知トリガーを設定しました（グループ: ' + getDailyNotifyGroupId() + '）');
}
