// ==========================================
// 進捗管理表 LINE自然文クエリ
// ==========================================
// シート名: 「進捗管理表」
// 想定列: 店舗名 | 施工月 | 区分 | 担当者 | 依頼日 | 図面チェック | 発注 | 見積 | ...
// 実装方針:
//   - ユーザー質問文から対象列・月・店舗・担当者を推定
//   - ヘッダー文字列はファジーマッチ（部分一致＋共通部分文字列）
//   - ○系の値は完了。空欄/△/確認中/×/未/保留/対応中/差戻し/要確認 は通知対象
// ==========================================

var PROGRESS_SHEET_NAME = '進捗管理表';

// 進捗管理表シートを取得（別スプシ対応）
// スクリプトプロパティ PROGRESS_SPREADSHEET_ID が設定されていればそのスプシを開く。
// 未設定なら既存のメインスプシ（SPREADSHEET_ID）を使う。
// シート名は PROGRESS_SHEET_NAME_OVERRIDE があればそちら、なければ PROGRESS_SHEET_NAME。
function getProgressSheet() {
  var props      = PropertiesService.getScriptProperties();
  var extId      = props.getProperty('PROGRESS_SPREADSHEET_ID');
  var sheetName  = props.getProperty('PROGRESS_SHEET_NAME_OVERRIDE') || PROGRESS_SHEET_NAME;
  try {
    var ss = extId ? SpreadsheetApp.openById(extId) : getSS();
    return ss.getSheetByName(sheetName);
  } catch (err) {
    console.error('getProgressSheet error:', err.message);
    return null;
  }
}

// 完了とみなすテキスト値
var PROGRESS_DONE_TEXTS = ['丸', 'まる', 'マル', '済', '済み', '完了', '完', 'ok', 'OK', 'done', 'Done', '✓', '✔', '✅', 'Ⓞ', '対応済', '済済'];

// 丸系のUnicode文字（円・丸記号のあらゆる異体字）
// U+25CB ○ / U+3007 〇 / U+25EF ◯ / U+25CE ◎ / U+26AA ⚪ / U+26AB ⚫ / U+25CF ●
// U+2B24 ⬤ / U+2B55 ⭕ / U+24EA Ⓞ / U+24C4 Ⓞ / U+25C9 ◉ / U+25EC ◬ / U+2299 ⊙
var PROGRESS_CIRCLE_REGEX = /^[○〇◯◎⚪⚫●⬤⭕⓪Ⓞ◉⊙]+$/;

// 完了判定（trimして照合・各種丸記号と完了テキストを許容）
function isProgressDone(value) {
  var v = String(value == null ? '' : value).replace(/^[\s　]+|[\s　]+$/g, '');
  if (!v) return false;
  if (PROGRESS_CIRCLE_REGEX.test(v)) return true;
  var lc = v.toLowerCase();
  for (var i = 0; i < PROGRESS_DONE_TEXTS.length; i++) {
    var d = PROGRESS_DONE_TEXTS[i];
    if (v === d || lc === String(d).toLowerCase()) return true;
  }
  return false;
}

// 既存互換のため変数名は残す（未使用）
var PROGRESS_DONE_VALUES = PROGRESS_DONE_TEXTS;

function progressStatusLabel(value) {
  var v = String(value == null ? '' : value).trim();
  return v ? v : '未入力';
}

function normalizeHeader(s) {
  // NFKC で半角カナ→全角カナ・全角英数→半角英数を正規化
  return String(s || '')
    .normalize('NFKC')
    .replace(/[\s　・/／\\\-_（）()\[\]【】]/g, '')
    .toLowerCase();
}

// 複数の候補名を順に試して見つかったヘッダーインデックスを返す
function findMatchingHeaderMulti(headers, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var idx = findMatchingHeader(headers, candidates[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}

// 質問文中にフレーズが含まれるか（NFKC正規化して比較）
function textContainsPhrase(text, phrase) {
  var t = String(text || '').normalize('NFKC');
  var p = String(phrase || '').normalize('NFKC');
  return p && t.indexOf(p) !== -1;
}

// ヘッダー配列から keyword に最も近い列インデックスを返す（見つからなければ-1）
function findMatchingHeader(headers, keyword) {
  if (!keyword || !headers || !headers.length) return -1;
  var k = normalizeHeader(keyword);
  if (!k) return -1;
  var best = -1, bestScore = 0;
  for (var i = 0; i < headers.length; i++) {
    var h = normalizeHeader(headers[i]);
    if (!h) continue;
    var score = 0;
    if (h === k)                              score = 100;
    else if (h.includes(k) || k.includes(h))  score = 80;
    else {
      var sub = longestCommonSubstring(h, k);
      if (sub >= 3)      score = 55 + sub * 3;
      else if (sub >= 2) score = 30;
    }
    if (score > bestScore) { bestScore = score; best = i; }
  }
  return bestScore >= 50 ? best : -1;
}

// 質問文 → 対象列・フィルタ条件を抽出
function analyzeProgressQuestion(text, headers) {
  var analysis = {
    targetColumns: [],    // 注目するチェック列名（複数可）
    monthFilter: null,    // 5 → 5月
    storeFilter: null,    // 「○○店」
    assigneeFilter: null, // 「田中」
    statusFilter: null,   // 「確認中」など特定状態
    mode: 'specific',     // 'specific' or 'all'
  };
  var t = String(text || '');

  // 月抽出
  var mm = t.match(/(\d{1,2})月/);
  if (mm) analysis.monthFilter = parseInt(mm[1], 10);
  if (/今月/.test(t)) analysis.monthFilter = new Date().getMonth() + 1;
  if (/来月/.test(t)) {
    var nm = new Date(); nm.setMonth(nm.getMonth() + 1);
    analysis.monthFilter = nm.getMonth() + 1;
  }

  // 特定状態キーワード
  var statusKeywords = ['確認中', '対応中', '保留', '差戻し', '差し戻し', '要確認', '△'];
  for (var si = 0; si < statusKeywords.length; si++) {
    if (t.indexOf(statusKeywords[si]) !== -1 && /(だけ|のみ|ある|教え|表示|どこ)/.test(t)) {
      analysis.statusFilter = statusKeywords[si];
      break;
    }
  }

  // 担当者「〇〇さん」
  var am = t.match(/([一-龥ぁ-んァ-ヶーA-Za-z]{1,8})さん/);
  if (am) analysis.assigneeFilter = am[1];

  // 店舗「〇〇店」抽出（漢字/カタカナ/英数のみ・ひらがな除外で「中の店舗」等の誤検出防止）
  // また「店舗」「店」単独語、「美容店」「理容店」は除外
  var sm = t.match(/([一-龥ァ-ヶーｦ-ﾟA-Za-z0-9]{1,12})店(?!舗)/);
  if (sm && sm[1] && !/^(美容|理容|美|理)$/.test(sm[1])) {
    analysis.storeFilter = sm[1];
  }

  // 全体モード
  if (/全部|全て|まとめて|一覧/.test(t)) analysis.mode = 'all';

  // 対象列推定：フレーズ→候補ヘッダー（複数可）。最初に見つかったヘッダーを採用。
  // 「発注」「納品」「見積」など複数列に対応する語は、複数列を同時に対象にする。
  var phraseMap = [
    { phrases: ['図面チェック', '図面確認', '図面'],
      headers: ['図面チェック'], multi: false },
    { phrases: ['依頼日', '依頼した日', '依頼の日', '依頼'],
      headers: ['依頼日'], multi: false },
    { phrases: ['制作一覧', '一覧入力', '一覧'],
      headers: ['制作一覧入力'], multi: false },
    { phrases: ['安陳見積もり', '安陳見積', '安陳の見積'],
      headers: ['安陳見積もり', '安陳見積'], multi: false },
    { phrases: ['安陳入力', '安陳'],
      headers: ['安陳入力'], multi: false },
    { phrases: ['スクショ見積', 'スクショ'],
      headers: ['スクショ見積'], multi: false },
    { phrases: ['金額調整', '金額確認', '金額'],
      headers: ['金額調整確認'], multi: false },
    { phrases: ['パイオニア見積', 'パイオニア', 'ﾊﾟｲｵﾆｱ'],
      headers: ['ﾊﾟｲｵﾆｱ見積', 'パイオニア見積'], multi: false },
    { phrases: ['発注書'],
      headers: ['発注書'], multi: false },
    { phrases: ['制作発注'],
      headers: ['制作発注'], multi: false },
    { phrases: ['発注'],
      headers: ['制作発注', '発注書'], multi: true },
    { phrases: ['納品日時', '納品日', '納品時間'],
      headers: ['納品日時'], multi: false },
    { phrases: ['納品場所'],
      headers: ['納品場所'], multi: false },
    { phrases: ['納品'],
      headers: ['納品日時', '納品場所'], multi: true },
    { phrases: ['完工日', '完工'],
      headers: ['完工日'], multi: false },
    { phrases: ['見積もり', '見積り', '見積'],
      headers: ['見積'], multi: false },
    { phrases: ['搬入'],
      headers: ['搬入'], multi: false },
    { phrases: ['設置'],
      headers: ['設置'], multi: false },
    { phrases: ['担当者', '担当'],
      headers: ['担当者'], multi: false },
    { phrases: ['区分', '美容', '理容', '理/美', '美/理'],
      headers: ['区分', '理/美', '美/理'], multi: false },
  ];

  var picked = {}; // 重複防止
  for (var pi = 0; pi < phraseMap.length; pi++) {
    var entry = phraseMap[pi];
    var hit = false;
    for (var pj = 0; pj < entry.phrases.length; pj++) {
      if (textContainsPhrase(t, entry.phrases[pj])) { hit = true; break; }
    }
    if (!hit) continue;

    if (entry.multi) {
      // 複数列同時に対象（例：発注 → 制作発注 + 発注書）
      entry.headers.forEach(function(hn) {
        var ix = findMatchingHeader(headers, hn);
        if (ix !== -1 && !picked[headers[ix]]) { picked[headers[ix]] = true; analysis.targetColumns.push(headers[ix]); }
      });
    } else {
      var ix = findMatchingHeaderMulti(headers, entry.headers);
      if (ix !== -1 && !picked[headers[ix]]) { picked[headers[ix]] = true; analysis.targetColumns.push(headers[ix]); }
    }
  }

  return analysis;
}

// 行抽出
// rows: シートの全データ行（ヘッダー除く）
// headers: ヘッダー文字列配列
function extractProgressItemsByQuestion(rows, headers, analysis) {
  var idxStore   = findMatchingHeaderMulti(headers, ['店舗名', '店舗', '店名']);
  var idxMonth   = findMatchingHeader(headers, '施工月');
  var idxKubun   = findMatchingHeaderMulti(headers, ['区分', '理/美', '美/理', '理美', '美理']);
  var idxAssign  = findMatchingHeaderMulti(headers, ['担当者', '担当']);
  var idxReqDate = findMatchingHeader(headers, '依頼日');

  // チェック列の範囲：図面チェック以降。なければ依頼日の次以降。それもなければ全列。
  var checkStart = findMatchingHeader(headers, '図面チェック');
  if (checkStart === -1 && idxReqDate !== -1) checkStart = idxReqDate + 1;
  if (checkStart === -1) checkStart = 0;

  // 日付系列は「空欄のみ通知」（値が入っていれば完了扱い）
  var dateColIdxs = {};
  ['依頼日', '納品日時', '完工日'].forEach(function(n) {
    var i = findMatchingHeader(headers, n);
    if (i !== -1) dateColIdxs[i] = true;
  });

  var results = [];

  rows.forEach(function(row) {
    var store = idxStore !== -1 ? String(row[idxStore] || '').trim() : '';
    if (!store) return;
    // ヘッダー行・テンプレート行をスキップ（シート末尾の見本行対策）
    if (store === '店舗名' || store === '店舗' || store === '店名') return;

    // 月フィルタ（Date・「5月」「３月」全角数字・「5/1」「2025/5/1」いずれも対応）
    if (analysis.monthFilter !== null && idxMonth !== -1) {
      var rawMonth = row[idxMonth];
      var monthVal = null;
      if (rawMonth instanceof Date) {
        monthVal = rawMonth.getMonth() + 1;
      } else {
        var mvStr = String(rawMonth == null ? '' : rawMonth).normalize('NFKC');
        var monthMatch = mvStr.match(/(\d{1,2})月/) || mvStr.match(/\/(\d{1,2})\b/) || mvStr.match(/^(\d{1,2})$/);
        if (monthMatch) monthVal = parseInt(monthMatch[1], 10);
      }
      if (monthVal !== analysis.monthFilter) return;
    }
    // 店舗フィルタ
    if (analysis.storeFilter && store.indexOf(analysis.storeFilter) === -1) return;
    // 担当者フィルタ
    if (analysis.assigneeFilter && idxAssign !== -1) {
      if (String(row[idxAssign] || '').indexOf(analysis.assigneeFilter) === -1) return;
    }

    function fmtCell(v) {
      if (v == null) return '';
      if (v instanceof Date) {
        return Utilities.formatDate(v, 'Asia/Tokyo', 'M/d');
      }
      return String(v).trim();
    }
    function fmtMonthCell(v) {
      if (v == null) return '';
      if (v instanceof Date) return (v.getMonth() + 1) + '月';
      var s = String(v).trim();
      if (/^\d{1,2}$/.test(s)) return s + '月';
      return s;
    }
    var item = {
      store:       store,
      month:       idxMonth   !== -1 ? fmtMonthCell(row[idxMonth])      : '',
      kubun:       idxKubun   !== -1 ? fmtCell(row[idxKubun])            : '',
      assignee:    idxAssign  !== -1 ? fmtCell(row[idxAssign])           : '',
      requestDate: idxReqDate !== -1 ? fmtCell(row[idxReqDate])          : '',
      pending: [],
    };

    function pushPending(colIdx, colName) {
      var v = row[colIdx];
      // 日付列：Date値あり/日付文字列は完了。空欄は未入力。それ以外（確認中等）は通常判定。
      if (dateColIdxs[colIdx]) {
        if (v instanceof Date) return;
        var s = String(v == null ? '' : v).trim();
        if (!s) {
          if (!analysis.statusFilter) item.pending.push({ column: colName, value: '未入力' });
          return;
        }
        if (/^\d{4}[\/\-年]\d{1,2}[\/\-月]\d{1,2}/.test(s) || /^\d{1,2}[\/月]\d{1,2}/.test(s)) return;
        if (isProgressDone(s)) return;
        if (analysis.statusFilter && s !== analysis.statusFilter) return;
        item.pending.push({ column: colName, value: s });
        return;
      }
      if (isProgressDone(v)) return;
      var label = progressStatusLabel(v);
      if (analysis.statusFilter && label !== analysis.statusFilter) return;
      item.pending.push({ column: colName, value: label });
    }

    if (analysis.targetColumns.length) {
      analysis.targetColumns.forEach(function(colName) {
        var ci = headers.indexOf(colName);
        if (ci !== -1) pushPending(ci, colName);
      });
    } else {
      // 全チェック列を見る。属性列・空ヘッダー・属性名の重複列は除外
      var attrNames = ['施工月','店舗名','店舗','店名','担当者','担当','依頼日','区分','理/美','美/理'];
      var attrNormSet = {};
      attrNames.forEach(function(a){ attrNormSet[normalizeHeader(a)] = true; });
      var skipCols = {};
      [idxStore, idxMonth, idxKubun, idxAssign, idxReqDate].forEach(function(i) { if (i !== -1) skipCols[i] = true; });
      for (var ci2 = checkStart; ci2 < headers.length; ci2++) {
        if (skipCols[ci2]) continue;
        var hname = String(headers[ci2] || '').trim();
        if (!hname) continue;
        if (attrNormSet[normalizeHeader(hname)]) continue;
        pushPending(ci2, hname);
      }
    }

    if (item.pending.length) results.push(item);
  });

  return results;
}

// 回答テキスト生成
function formatProgressQuestionAnswer(results, analysis) {
  if (!results.length) return '該当する未完了・要確認の店舗はありません。';

  var headerLine = '';
  if (analysis.statusFilter) {
    headerLine = '【状態「' + analysis.statusFilter + '」の項目】';
  } else if (analysis.targetColumns.length === 1) {
    headerLine = '【' + analysis.targetColumns[0] + '：未完了・要確認】';
  } else if (analysis.monthFilter !== null) {
    headerLine = '【' + analysis.monthFilter + '月施工分：未完了・要確認】';
  } else {
    headerLine = '【未完了・要確認】';
  }

  var MAX = 30;
  var lines = [headerLine];
  results.slice(0, MAX).forEach(function(it) {
    lines.push('');
    lines.push('■ ' + it.store);
    if (it.month)       lines.push('施工月：' + it.month);
    if (it.kubun)       lines.push('区分：' + it.kubun);
    if (it.assignee)    lines.push('担当：' + it.assignee);

    if (analysis.targetColumns.length === 1 && it.pending.length === 1) {
      lines.push('状態：' + it.pending[0].value);
    } else if (it.pending.length === 1 && analysis.targetColumns.length === 0 && analysis.statusFilter) {
      lines.push(it.pending[0].column + '：' + it.pending[0].value);
    } else {
      lines.push('未完了：');
      it.pending.forEach(function(p) { lines.push('・' + p.column + '：' + p.value); });
    }
  });
  if (results.length > MAX) {
    var rest = results.slice(MAX).map(function(it) { return it.store; }).join('、');
    lines.push('\n…ほか ' + (results.length - MAX) + ' 件：' + rest);
  }
  return lines.join('\n');
}

// 進捗質問エントリ：質問でなければnull、質問なら回答テキスト返却
function answerProgressQuestion(text) {
  var sheet = getProgressSheet();
  if (!sheet || sheet.getLastRow() < 2) return null;

  var data       = sheet.getDataRange().getValues();
  var headerRowIdx = findHeaderRowIndex(sheet);
  var headers    = data[headerRowIdx].map(function(h){ return String(h == null ? '' : h).trim(); });
  var rows       = data.slice(headerRowIdx + 1);

  var analysis = analyzeProgressQuestion(text, headers);

  // タスク登録メッセージは絶対に進捗質問扱いしない
  var hasTaskVerb = /(お願い|頼みま|頼む|依頼しま|依頼しと|してお|しといて|しといで|送っと|送って[くだ]|連絡しと|連絡して|対応して|登録して|入れと|やっとく|やってお|やっとき|やります|でお願い|完了したよ|完了です|終わ(りまし|った))/.test(text);
  if (hasTaskVerb) return null;

  // 明示的進捗キーワード
  var explicitProgress = /未完了|まとめて|一覧|全部|全て|進捗管理|残ってる|残ってない|○ついて(い)?ない|まる(つい|入っ)てない|空欄/.test(text);

  // クエリ動詞（質問形）
  var hasQueryVerb =
       /(ある[?？]?|ない[?？]?|教え|まだ|どこ|どれ|どう|何件|何個|いくつ|表示|出して|出せ|だして|だせ)/.test(text)
    || /[?？]/.test(text);

  var hasFilter =
       analysis.targetColumns.length > 0
    || analysis.monthFilter   !== null
    || analysis.storeFilter   !== null
    || analysis.assigneeFilter!== null
    || analysis.statusFilter  !== null;

  // 進捗質問として扱う条件：
  //   ① 明示キーワード（未完了/まとめて等）が含まれる → 即発火
  //   ② フィルタ条件あり ＋ クエリ動詞 ＋ 進捗関連シグナル
  var progressSignal = analysis.targetColumns.length > 0 || analysis.statusFilter !== null
                     || /進捗|どうなって|状況|状態|店舗/.test(text);
  var triggered = explicitProgress || (hasFilter && hasQueryVerb && progressSignal);

  if (!triggered) return null;

  var results = extractProgressItemsByQuestion(rows, headers, analysis);
  return formatProgressQuestionAnswer(results, analysis);
}

// LINE event 入口（handleMentionCommand / handleDM から呼ぶ）
// 返り値: true=回答済み, false=対象外
function handleProgressQuestionRequest(event, messageText) {
  try {
    var answer = answerProgressQuestion(messageText);
    if (!answer) return false;
    sendLineReply(event.replyToken, answer);
    return true;
  } catch (err) {
    console.error('handleProgressQuestionRequest error:', err.message);
    return false;
  }
}

// ==========================================
// 毎朝9時：善波グループ向け 進捗管理表ダイジェスト
// 「今月＋翌月の施工分」で「まだ完了していない項目」を一覧にする。
// 進捗管理表には日次の予定日が無い（依頼日のみ実日付・他はテキスト）ため、
// 「翌日分」は施工月（今月＋翌月）の未完了で近似する。
// ==========================================
function buildSenbaProgressDigest() {
  var sheet = getProgressSheet();
  if (!sheet || sheet.getLastRow() < 2) return null;

  var data         = sheet.getDataRange().getValues();
  var headerRowIdx = findHeaderRowIndex(sheet);
  var headers      = data[headerRowIdx].map(function(h){ return String(h == null ? '' : h).trim(); });
  var rows         = data.slice(headerRowIdx + 1);

  var now    = new Date();
  var curM   = now.getMonth() + 1;
  var nextM  = now.getMonth() === 11 ? 1 : curM + 1;
  var months = [curM, nextM];

  // 月ごとに未完了抽出（既存の抽出ロジックを再利用）
  var results = [];
  months.forEach(function(m) {
    var analysis = {
      targetColumns: [], monthFilter: m, storeFilter: null,
      assigneeFilter: null, statusFilter: null, mode: 'all',
    };
    extractProgressItemsByQuestion(rows, headers, analysis).forEach(function(it) {
      results.push(it);
    });
  });

  if (!results.length) return null;

  var lines = ['●進捗管理表 未完了（' + curM + '月・' + nextM + '月 施工分）'];
  var MAX = 40;
  results.slice(0, MAX).forEach(function(it) {
    lines.push('');
    var head = '■ ' + it.store;
    if (it.month)    head += '（' + it.month + (it.kubun ? '・' + it.kubun : '') + '）';
    else if (it.kubun) head += '（' + it.kubun + '）';
    lines.push(head);
    if (it.assignee) lines.push('担当：' + it.assignee);
    it.pending.forEach(function(p) { lines.push('・' + p.column + '：' + p.value); });
  });
  if (results.length > MAX) {
    var rest = results.slice(MAX).map(function(it){ return it.store; }).join('、');
    lines.push('\n…ほか ' + (results.length - MAX) + ' 件：' + rest);
  }
  return lines.join('\n');
}

// ==========================================
// テスト
// ==========================================
function testProgressQuestion() {
  var sheet = getProgressSheet();
  if (!sheet) { console.error('進捗管理表シートが見つかりません（PROGRESS_SPREADSHEET_ID / シート名を確認）'); return; }
  console.log('対象スプシ:', sheet.getParent().getName(), '/ シート:', sheet.getName());
  var cases = [
    '図面チェックできてない店舗ある？',
    '依頼日入ってない店舗ある？',
    '5月の図面チェックまだのところ教えて',
    '確認中の店舗だけ教えて',
    '6月施工で未完了ある？',
    '田中さん担当で未完了ある？',
    '未完了全部教えて',
  ];
  cases.forEach(function(q) {
    console.log('Q: ' + q);
    var ans = answerProgressQuestion(q);
    console.log(ans || '（対象外と判定）');
    console.log('---');
  });
}

// GASエディタの実行ボタン用：このプロジェクトの進捗管理スプシ・シート名を直接設定
function setupProgressSpreadsheetForWoodbase() {
  setupProgressSpreadsheet('1zzA2qSoKZoTBp81BvH4Vl36TdJUuy1F4prjzcKUX-uE', 'プラージュ家具　進捗管理表');
}

// 別スプシを進捗管理表として登録（GASエディタから1回実行）
// 例: setupProgressSpreadsheet('1zzA2qSoKZoTBp81BvH4Vl36TdJUuy1F4prjzcKUX-uE', '進捗管理表');
function setupProgressSpreadsheet(spreadsheetId, sheetName) {
  if (!spreadsheetId) { console.error('spreadsheetId が必要です'); return; }
  var props = PropertiesService.getScriptProperties();
  props.setProperty('PROGRESS_SPREADSHEET_ID', String(spreadsheetId));
  if (sheetName) props.setProperty('PROGRESS_SHEET_NAME_OVERRIDE', String(sheetName));

  // 動作確認：開けるか・対象シートが存在するか
  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var target = sheetName || PROGRESS_SHEET_NAME;
    var sh = ss.getSheetByName(target);
    if (!sh) {
      console.error('スプシは開けましたが「' + target + '」シートが見つかりません。シート名を確認してください。');
      console.log('利用可能シート:', ss.getSheets().map(function(s){ return s.getName(); }).join(', '));
      return;
    }
    console.log('✅ 設定完了:', ss.getName(), '/', sh.getName(), '/ 行数:', sh.getLastRow());
  } catch (err) {
    console.error('開けませんでした。共有設定（GAS実行アカウントに閲覧権限）と ID を確認してください。', err.message);
  }
}

// 進捗管理表の現在の接続先を確認
function showProgressSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  console.log('PROGRESS_SPREADSHEET_ID:', props.getProperty('PROGRESS_SPREADSHEET_ID') || '(未設定 → メインスプシを使用)');
  console.log('PROGRESS_SHEET_NAME_OVERRIDE:', props.getProperty('PROGRESS_SHEET_NAME_OVERRIDE') || '(未設定 → ' + PROGRESS_SHEET_NAME + ')');
  var sh = getProgressSheet();
  if (sh) console.log('解決後:', sh.getParent().getName(), '/', sh.getName(), '/ 行数:', sh.getLastRow());
  else    console.log('シート解決失敗');
}

// 診断：シートの先頭5行を表示してヘッダー行を特定する
function diagnoseProgressSheet() {
  var sh = getProgressSheet();
  if (!sh) { console.error('シート未接続'); return; }
  console.log('シート:', sh.getParent().getName(), '/', sh.getName());
  console.log('行数:', sh.getLastRow(), '列数:', sh.getLastColumn());
  var n = Math.min(8, sh.getLastRow());
  var data = sh.getRange(1, 1, n, sh.getLastColumn()).getValues();
  for (var i = 0; i < data.length; i++) {
    console.log('行' + (i + 1) + ':', JSON.stringify(data[i].map(function(c){
      if (c instanceof Date) return Utilities.formatDate(c, 'Asia/Tokyo', 'yyyy/M/d');
      return c;
    })));
  }
}

// 1つのスプレッドシートの全タブ＋中身プレビューをログ出力する内部ヘルパー
function dumpOneSpreadsheet_(ss, labelPrefix) {
  var sheets = ss.getSheets();
  console.log('\n========================================');
  console.log('=== ' + labelPrefix + ': ' + ss.getName() + ' / タブ数: ' + sheets.length + ' ===');
  console.log('タブ一覧: ' + sheets.map(function(s){ return s.getName(); }).join(' | '));

  var PREVIEW_ROWS = 20;
  var PREVIEW_COLS = 14;

  sheets.forEach(function(sh) {
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    console.log('\n----- タブ「' + sh.getName() + '」 行数:' + lastRow + ' 列数:' + lastCol + ' -----');
    if (lastRow < 1 || lastCol < 1) { console.log('（空タブ）'); return; }

    var rows = Math.min(PREVIEW_ROWS, lastRow);
    var cols = Math.min(PREVIEW_COLS, lastCol);
    var data = sh.getRange(1, 1, rows, cols).getValues();
    for (var i = 0; i < data.length; i++) {
      var cells = data[i].map(function(c) {
        if (c == null || c === '') return '';
        if (c instanceof Date) return Utilities.formatDate(c, 'Asia/Tokyo', 'yyyy/M/d HH:mm');
        return String(c).replace(/\n/g, '⏎');
      });
      console.log('R' + (i + 1) + ': ' + JSON.stringify(cells));
    }
    if (lastRow > rows) console.log('… ほか ' + (lastRow - rows) + ' 行省略');
  });
}

// 診断：進捗管理表の「全列・全行」をダンプする。
// 「翌日でまだできてない部分」をどの日付列で判定するか決めるために、
// 納品日時・完工日などが実際にどんな値（Date か 自由テキストか）かを全行確認する。
// GASエディタで dumpProgressSheetFull を実行 → 実行ログを共有してください。
function dumpProgressSheetFull() {
  var sh = getProgressSheet();
  if (!sh) { console.error('進捗管理表シートに接続できません'); return; }
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  console.log('=== 進捗管理表 全ダンプ:', sh.getParent().getName(), '/', sh.getName(), '/ 行:', lastRow, '列:', lastCol, '===');

  var data = sh.getRange(1, 1, lastRow, lastCol).getValues();
  for (var i = 0; i < data.length; i++) {
    var cells = data[i].map(function(c) {
      if (c == null || c === '') return '';
      if (c instanceof Date) return '[D]' + Utilities.formatDate(c, 'Asia/Tokyo', 'yyyy/M/d HH:mm');
      return String(c).replace(/\n/g, '⏎');
    });
    console.log('R' + (i + 1) + ': ' + JSON.stringify(cells));
  }
  console.log('=== ここまで（[D]=Date値, それ以外はテキスト）。全部コピーして共有してください ===');
}

// 診断：進捗管理スプレッドシート内の「全タブ」と各タブの中身プレビューを出力する。
// 「いつも送ってる予定メッセージ」がどのタブ・どの範囲にあるかを特定するために使う。
// GASエディタからこの関数を1回実行し、実行ログ（表示 → 実行ログ）を全部コピーして共有してください。
function dumpScheduleSpreadsheet() {
  var props = PropertiesService.getScriptProperties();

  // ① 進捗管理スプレッドシート
  var extId = props.getProperty('PROGRESS_SPREADSHEET_ID')
            || '1zzA2qSoKZoTBp81BvH4Vl36TdJUuy1F4prjzcKUX-uE';
  try {
    dumpOneSpreadsheet_(SpreadsheetApp.openById(extId), '進捗管理スプレッドシート');
  } catch (err) {
    console.error('進捗管理スプレッドシートを開けませんでした:', err.message);
  }

  // ② メインスプレッドシート（スケジュール管理・タスク管理などがある方）
  var mainId = props.getProperty('SPREADSHEET_ID');
  if (mainId && mainId !== extId) {
    try {
      dumpOneSpreadsheet_(SpreadsheetApp.openById(mainId), 'メインスプレッドシート');
    } catch (err) {
      console.error('メインスプレッドシートを開けませんでした:', err.message);
    }
  } else if (!mainId) {
    console.log('\n（SPREADSHEET_ID 未設定のためメインスプレッドシートはスキップ）');
  }

  console.log('\n=== ここまで。このログを全部コピーして共有してください ===');
}

// ヘッダー行のインデックスを自動推定（「店舗名」または「図面チェック」を含む最初の行）
function findHeaderRowIndex(sh) {
  var maxScan = Math.min(10, sh.getLastRow());
  var data = sh.getRange(1, 1, maxScan, sh.getLastColumn()).getValues();
  var keywords = ['店舗名', '店舗', '図面チェック', '図面', '施工月', '依頼日'];
  for (var i = 0; i < data.length; i++) {
    var row = data[i].map(function(c){ return String(c == null ? '' : c).normalize('NFKC'); });
    var hits = 0;
    keywords.forEach(function(k){
      if (row.some(function(c){ return c.indexOf(k) !== -1; })) hits++;
    });
    if (hits >= 2) return i; // 0-indexed
  }
  return 0;
}

// ヘッダーマッチのみ単体テスト（シート不要）
function testFindMatchingHeader() {
  var headers = ['店舗名', '施工月', '美容/理容区分', '担当者', '依頼日', '図面チェック', '発注', '見積', '納品'];
  var cases = [
    ['図面',     '図面チェック'],
    ['図面確認', '図面チェック'],
    ['担当',     '担当者'],
    ['依頼した日', '依頼日'],
    ['美容',     '美容/理容区分'],
    ['区分',     '美容/理容区分'],
    ['発注',     '発注'],
    ['存在しない列', null],
  ];
  cases.forEach(function(c) {
    var idx = findMatchingHeader(headers, c[0]);
    var got = idx === -1 ? null : headers[idx];
    console.log((got === c[1] ? '✅' : '❌') + ' "' + c[0] + '" → ' + got + ' (期待:' + c[1] + ')');
  });
}
