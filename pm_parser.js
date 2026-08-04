// ==========================================
// ProjectParser：intent判定（プレフィックス）＋ Gemini構造化抽出
//   依存: pm_core.js, 既存 callGemini / getConfig / fmtDate
// ==========================================

// ---- プレフィックスによる intent 判定（MVP：ゆるいルール / 高精度） ----
//  該当しなければ null（＝既存処理に委譲）
//  順序重要：明示語（〇〇更新/報告）を先に。素の「請求」「入金」は
//  区切り必須にして「請求書はどこ？」等の通常文に誤反応しないようにする。
var PM_PREFIX_RULES = [
  { re: /^[\s　]*(案件更新|進捗更新|進捗報告)[\s　:：]*/, intent: 'project_update' },
  { re: /^[\s　]*(新規案件|案件登録)[\s　:：]*/,        intent: 'project_create' },
  { re: /^[\s　]*(店舗追加|店舗登録)[\s　:：]*/,        intent: 'store_add' },
  { re: /^[\s　]*(請求更新|請求報告)[\s　:：]*/,        intent: 'billing_update' },
  { re: /^[\s　]*(入金更新|入金報告)[\s　:：]*/,        intent: 'payment_update' },
  { re: /^[\s　]*(予定追加|カレンダー追加)[\s　:：]*/,   intent: 'calendar_update' },
  { re: /^[\s　]*(請求)[\s　:：]+/,                     intent: 'billing_update' },
  { re: /^[\s　]*(入金)[\s　:：]+/,                     intent: 'payment_update' },
];

function pmClassifyIntent(text) {
  if (!text) return null;
  for (var i = 0; i < PM_PREFIX_RULES.length; i++) {
    if (PM_PREFIX_RULES[i].re.test(text)) return PM_PREFIX_RULES[i].intent;
  }
  return null;
}

// プレフィックスを除去した本文を返す
function pmStripPrefix(text) {
  for (var i = 0; i < PM_PREFIX_RULES.length; i++) {
    if (PM_PREFIX_RULES[i].re.test(text)) return text.replace(PM_PREFIX_RULES[i].re, '').trim();
  }
  return String(text || '').trim();
}

// ---- Geminiで構造化抽出（設計書 §6 のJSONスキーマ） ----
//  intent はプレフィックス判定値を優先採用するため引数で渡す。
function pmParseReport(body, intent) {
  var config = getConfig();
  var today  = fmtDate(new Date());
  var phaseGuide = Object.keys(PM_PHASES).map(function(p) {
    return p + '：' + PM_PHASES[p].join(' / ');
  }).join('\n');

  var prompt =
    'あなたは建築・内装会社の案件管理アシスタントです。今日は ' + today + '（Asia/Tokyo）です。\n' +
    '以下のLINE報告文から案件情報を抽出し、JSONのみを出力してください（前置き・コードフェンス禁止）。\n' +
    '日付は必ず YYYY-MM-DD 形式に正規化。「6/25」など年が無い場合は今日以降の直近の日付に補完。相対表現（来週/明日等）も今日基準で解決。\n' +
    '【重要】日にちが1日に特定できる場合のみ日付フィールドに出力する。「〜頃」「〜あたり」「中旬/下旬」「来月中」「未定」「調整中」など曖昧・仮の日程は日付フィールドに入れず null とし、原文の表現を remark に残す（曖昧な日付をカレンダー登録しないため）。\n' +
    '時刻は HH:mm（24時間制）に正規化（「14時」→「14:00」「14時半」→「14:30」）。終了時刻が無ければ event_end_time は null。\n' +
    '「予定追加」系の文では、予定の名称を event_title、日付を event_date、開始時刻を event_time に入れる（例「予定追加 〇〇 7/3 14時 現地打合せ」→ event_title:"現地打合せ", event_date:"2026-07-03", event_time:"14:00"）。\n' +
    '値が読み取れない項目は null。金額は数値(円)に換算し raw に原文を残す（例「5000万」→ 50000000）。\n' +
    '\n【フェーズと許可ステータス】\n' + phaseGuide + '\n' +
    '\n【出力JSONスキーマ】\n' +
    '{\n' +
    '  "project_name": string|null,\n' +
    '  "client_name": string|null,\n' +
    '  "assignee": string|null,\n' +
    '  "phase": "営業"|"設計・デザイン"|"施工準備"|"施工"|"お引き渡し"|null,\n' +
    '  "status": string|null,\n' +
    '  "next_action": string|null,\n' +
    '  "next_action_due_date": "YYYY-MM-DD"|null,\n' +
    '  "meeting_date": "YYYY-MM-DD"|null,\n' +
    '  "event_title": string|null,\n' +
    '  "event_date": "YYYY-MM-DD"|null,\n' +
    '  "event_time": "HH:mm"|null,\n' +
    '  "event_end_time": "HH:mm"|null,\n' +
    '  "construction_start_date": "YYYY-MM-DD"|null,\n' +
    '  "handover_date": "YYYY-MM-DD"|null,\n' +
    '  "billing_due_date": "YYYY-MM-DD"|null,\n' +
    '  "billing_done_date": "YYYY-MM-DD"|null,\n' +
    '  "payment_due_date": "YYYY-MM-DD"|null,\n' +
    '  "payment_confirmed_date": "YYYY-MM-DD"|null,\n' +
    '  "billing_kind": string|null,\n' +
    '  "amount": { "type": "売上"|"原価"|"請求"|"入金"|null, "value": number|null, "raw": string|null },\n' +
    '  "remark": string|null,\n' +
    '  "confidence": number,\n' +
    '  "confirmation_reason": string|null\n' +
    '}\n' +
    '\n【報告文】\n' + body;

  var raw = callGemini(config.GEMINI_API_KEY, prompt, 0);
  var parsed = pmExtractJson(raw);
  if (!parsed) {
    // フォールバック：抽出失敗 → 確認必須で最小情報のみ
    parsed = { confidence: 0, confirmation_reason: 'AI抽出に失敗しました' };
  }
  parsed.intent = intent || parsed.intent || 'unknown';

  // 確認必須フラグの強制ルール（設計書 §6）
  var amountPresent = parsed.amount && parsed.amount.value !== null && parsed.amount.value !== undefined;
  parsed.needs_confirmation =
    parsed.intent === 'billing_update' ||
    parsed.intent === 'payment_update' ||
    !!amountPresent ||
    (typeof parsed.confidence === 'number' && parsed.confidence < 70);

  return parsed;
}

// Geminiの返答テキストからJSONを安全に取り出す
function pmExtractJson(raw) {
  if (!raw) return null;
  var s = String(raw).trim();
  // コードフェンス除去
  s = s.replace(/^```(json)?/i, '').replace(/```$/g, '').trim();
  var start = s.indexOf('{');
  var end   = s.lastIndexOf('}');
  if (start === -1 || end === -1 || end <= start) return null;
  try { return JSON.parse(s.slice(start, end + 1)); }
  catch (e) { console.error('pmExtractJson parse error:', e.message); return null; }
}
