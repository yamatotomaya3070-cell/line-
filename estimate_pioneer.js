// ============================================================
// パイオニア什器：制作一覧 → 店舗別 見積書 自動作成
// ※「状況」列は見積とは無関係（店舗名のある全行が対象）
// ------------------------------------------------------------
// ★ 既存のLINEチャットボット処理／既存 estimate.js とは完全に独立。
//   - doPost / 署名検証 / handleEvent / handleDM には一切手を加えない。
//   - グローバル定数の衝突を避けるため、設定はすべて EST_PIONEER 配下に閉じる。
//   - 補助関数は末尾 "_"（外部から直接呼ばない前提）。
//   - 手動実行のエントリーポイント: createEstimateSheets()
// ============================================================

// ---- 設定（裸のグローバル定数を作らず1オブジェクトに集約）----
var EST_PIONEER = {
  SOURCE_SS_ID:    '1AiI58tt-CLmlcgghsI_nyH3BVOn48LDr3erqNUR9gBA', // 参照元スプレッドシート
  SOURCE_SHEET:    '制作一覧',
  TEMPLATE_SS_ID:  '1AITs5qf2A3-y5LLHN9hYi1mB4iWZ52F6FwyoIBBr2-s', // 見積書テンプレート
  TEMPLATE_SHEET:  '原本',  // テンプレ(マスター)シート名。実体は『原本』だった。
  TAX_RATE:        0.10,
  UNIT_LABEL:      '台',
  OUTPUT_FOLDER_ID: '',  // 空ならマイドライブ直下に出力。フォルダ指定したい場合はIDを入れる。
  TZ:              'Asia/Tokyo',
};

// ============================================================
// SECTION: メイン処理（手動実行）
// ============================================================
// 本番：全店舗ぶんの見積書（シート）を作成
function createEstimateSheets() { return runEstimateCore_(Infinity, ''); }

// テスト：先頭1店舗だけ作成（本番前の動作確認用。プルダウンに表示される）
function runEstimateTestOneStore() { return runEstimateCore_(1, ''); }

// requestText: LINEメッセージ等。制作一覧に存在する店舗名が含まれていれば、その店舗だけ作成する。
function runEstimateCore_(maxStores, requestText) {
  var log = { created: 0, stores: 0, detailRows: 0, skipped: 0, files: [], error: '', note: '', requestedStores: [] };

  // 1〜2. 参照元スプレッドシート／制作一覧シートを開く
  var srcSheet;
  try {
    var srcSs = SpreadsheetApp.openById(EST_PIONEER.SOURCE_SS_ID);
    srcSheet = srcSs.getSheetByName(EST_PIONEER.SOURCE_SHEET);
  } catch (e) {
    log.error = '参照元スプレッドシートを開けません: ' + e.message;
    console.error('❌ ' + log.error);
    return log;
  }
  if (!srcSheet) {
    log.error = '「' + EST_PIONEER.SOURCE_SHEET + '」シートが見つかりません。';
    console.error('❌ ' + log.error);
    return log;
  }

  // 3〜5. ヘッダー自動判定 → 店舗名あり行を抽出 → 店舗名空欄スキップ
  var extracted = getEstimateSourceRows_(srcSheet);
  if (extracted.missingCols.length) {
    log.error = '必須列が見つかりません: ' + extracted.missingCols.join(' / ');
    console.error('❌ ' + log.error);
    return log;
  }
  log.skipped = extracted.skipped;

  // 店舗指定（メッセージに制作一覧上の店舗名が含まれていればその店舗だけ対象にする）
  var requested = findRequestedStores_(requestText, extracted.allStores);
  log.requestedStores = requested;

  var targetRows = extracted.rows;
  if (!targetRows.length) {
    log.note = '制作一覧に店舗名の入った行が0件のため、見積書は作成しませんでした。スキップ行数: ' + log.skipped;
    console.log(log.note);
    return log;
  }
  // 6. 店舗名ごとにグループ化
  var allGroups = groupEstimateRowsByStore_(targetRows);

  // 店舗指定らしき語（「〇〇店」）があるのに一致しなかった場合は、
  // 意図しない全店舗作成をせず、候補一覧を返して止める
  if (!requested.length) {
    var hinted = extractStoreHints_(requestText);
    if (hinted.length) {
      var candidateStores = allGroups.map(function (g) { return g.store; });
      log.note = '店舗「' + hinted.join('、') + '」が制作一覧の「店舗名」と一致しませんでした。\n' +
                 '登録のある店舗：' + (candidateStores.length ? candidateStores.join('、') : '（なし）') + '\n' +
                 '上記の店舗名で再送してください。';
      console.log(log.note);
      return log;
    }
  }
  if (requested.length) {
    var availableStores = allGroups.map(function (g) { return g.store; });
    var want = {};
    requested.forEach(function (s) { want[normCmp_(s)] = true; });
    allGroups = allGroups.filter(function (g) { return want[normCmp_(g.store)]; });
    if (!allGroups.length) {
      log.note = '指定された店舗（' + requested.join('、') + '）の行が制作一覧に見つからないため、見積書は作成しませんでした。\n' +
                 '登録のある店舗：' + (availableStores.length ? availableStores.join('、') : '（なし）');
      console.log(log.note);
      return log;
    }
    console.log('★店舗指定：' + requested.join('、') + ' のみ作成します');
  }
  var groups = isFinite(maxStores) ? allGroups.slice(0, maxStores) : allGroups;
  if (isFinite(maxStores)) console.log('★テストモード：先頭 ' + groups.length + ' 店舗のみ作成します（全 ' + allGroups.length + ' 店舗中）');
  log.stores = groups.length;
  log.detailRows = groups.reduce(function (n, g) { return n + g.items.length; }, 0);

  // 7〜11. 店舗ごとに見積書作成（1店舗のエラーで全体を止めない）
  groups.forEach(function (g) {
    try {
      var title = buildEstimateTitle_(g.rimei, g.store);                 // 8
      var copy  = copyEstimateTemplate_(g.store, g.rimei);               // 7
      if (!copy) { console.warn('⚠️ ' + g.store + '：テンプレコピー失敗のためスキップ'); return; }

      writeEstimateHeader_(copy.sheet, title);                          // 9
      var totals = writeEstimateDetails_(copy.sheet, g.items);          // 10
      calculateEstimateTotals_(copy.sheet, totals);                     // 11

      log.created++;
      log.files.push({ store: g.store, name: copy.fileName, url: copy.url });
    } catch (err) {
      console.error('⚠️ 店舗「' + g.store + '」でエラー（スキップして継続）: ' + err.message);
    }
  });

  // 12. ログ出力
  logEstimateResult_(log);
  return log;
}

// ============================================================
// SECTION: 制作一覧 読み取り
// ============================================================
// 店舗名あり の行を、ヘッダー名ベースで抽出（「状況」列は見積とは無関係）。
// 戻り値: { rows:[...], skipped:Number, missingCols:[String] }
function getEstimateSourceRows_(sheet) {
  if (sheet.getLastRow() < 2) return { rows: [], skipped: 0, missingCols: [], allStores: [] };

  var values = sheet.getDataRange().getValues();
  var headerRow = findHeaderRowIndex_(values);
  var headers   = values[headerRow].map(normHeader_);

  function col() { // 別名候補から最初に一致した列indexを返す（-1=なし）
    for (var a = 0; a < arguments.length; a++) {
      var t = normHeader_(arguments[a]);
      var i = headers.indexOf(t);
      if (i !== -1) return i;
      for (var j = 0; j < headers.length; j++) {
        if (headers[j] && headers[j].indexOf(t) !== -1) return j; // 部分一致フォールバック
      }
    }
    return -1;
  }

  // NFKC正規化済みなので Ｄ→D、（）→() に揃っている
  var c = {
    store:   col('店舗名'),
    rimei:   col('理/美', '理美', '理/美区分'),
    name:    col('什器名'),
    qty:     col('台数'),
    d:       col('D'),
    w:       col('W'),
    h:       col('H'),
    sUnit:   col('P指値(単価)'),
    sAmount: col('P指値(金額)'),
    unit:    col('単価'),
    amount:  col('金額'),
  };

  // 必須列チェック（D/W/H・状況は任意、それ以外は必須）
  var required = { 店舗名: c.store, '理/美': c.rimei, 什器名: c.name, 台数: c.qty, 単価: c.unit, 金額: c.amount };
  var missing = [];
  Object.keys(required).forEach(function (k) { if (required[k] === -1) missing.push(k); });
  if (missing.length) return { rows: [], skipped: 0, missingCols: missing, allStores: [] };

  var rows = [], skipped = 0, allStores = [], storeSeen = {};
  for (var r = headerRow + 1; r < values.length; r++) {
    var row = values[r];
    var store = String(row[c.store] == null ? '' : row[c.store]).trim();
    // 店舗名の全量リスト（LINEの店舗指定判定に使う）
    if (store && !storeSeen[store]) { storeSeen[store] = true; allStores.push(store); }
    // 「状況」列は見積とは無関係。店舗名のある全行が対象。
    if (!store) {
      // 内容があるのに店舗名だけ空欄の行をカウント（完全な空行は数えない）
      if (String(row[c.name] == null ? '' : row[c.name]).trim()) skipped++;
      continue;
    }

    rows.push({
      store:   store,
      rimei:   String(row[c.rimei] == null ? '' : row[c.rimei]).trim(),
      name:    String(row[c.name]  == null ? '' : row[c.name]).trim(),
      qty:     toNum_(row[c.qty]),
      d:       cellStr_(row[c.d]),
      w:       cellStr_(row[c.w]),
      h:       cellStr_(row[c.h]),
      sUnit:   c.sUnit   !== -1 ? toNum_(row[c.sUnit])   : '',
      sAmount: c.sAmount !== -1 ? toNum_(row[c.sAmount]) : '',
      unit:    toNum_(row[c.unit]),
      amount:  toNum_(row[c.amount]),
    });
  }
  return { rows: rows, skipped: skipped, missingCols: [], allStores: allStores };
}

// メッセージ内に含まれる店舗名を、制作一覧の店舗名リストから抽出する。
// 「九条店」のように末尾に「店」を付けた表記でも一致するようにする。
// 長い店舗名から順に照合し、一致した部分はテキストから除去する
// （例：「西新宿」を含むメッセージで、別店舗の「新宿」まで誤一致するのを防ぐ）。
// 完全包含で一致しない場合は、「〇〇店」の〇〇が店舗名の一部に一致するかで再照合する
// （例：メッセージ「九条店」⇔ シート「イオンモール九条」）。
function findRequestedStores_(text, storeList) {
  if (!text || !storeList || !storeList.length) return [];
  var t = normCmp_(text);
  var sorted = storeList.slice().sort(function (a, b) {
    return normCmp_(b).length - normCmp_(a).length;
  });
  var found = {};
  sorted.forEach(function (store) {
    var s = normCmp_(store);
    if (!s) return;
    var hit = '';
    if (t.indexOf(s) !== -1) {
      hit = s;
    } else {
      var stripped = s.replace(/店$/, '');
      if (stripped && stripped !== s && t.indexOf(stripped) !== -1) hit = stripped;
    }
    if (hit) {
      found[store] = true;
      t = t.split(hit).join('　');  // 一致部分を除去し、より短い店舗名の重複一致を防ぐ
    }
  });
  var exact = storeList.filter(function (store) { return found[store] === true; });
  if (exact.length) return exact;

  // フォールバック：「〇〇店」の〇〇を含む店舗名を探す
  var hints = extractStoreHints_(text);
  if (!hints.length) return [];
  return storeList.filter(function (store) {
    var s = normCmp_(store);
    return hints.some(function (h) { return s.indexOf(h) !== -1; });
  });
}

// メッセージ中の「〇〇店」という表記から店舗名候補（〇〇の部分）を抜き出す。
// 「1店舗」「全店」などの一般表現や、コマンド語を含むものは除外する。
function extractStoreHints_(text) {
  if (!text) return [];
  var t = normCmp_(text);
  var hints = [], m;
  var re = /([^\s、。,．・のをはがで「」()（）]{1,15})店/g;
  while ((m = re.exec(t)) !== null) {
    var h = m[1];
    // 先頭に付いたコマンド語・冠語を除去（「パイオニア九条店」→「九条」）
    h = h.replace(/^(パイオニア|プラージュ|理容|美容|什器|見積書|見積|指定)+/, '');
    if (!h) continue;
    if (/^[0-9一二三四五六七八九十全各当支本]+$/.test(h)) continue;  // 「1店舗」「全店」等
    if (hints.indexOf(h) === -1) hints.push(h);
  }
  return hints;
}

// 「店舗名」を含む行をヘッダー行とみなす（先頭10行以内）。無ければ0行目。
function findHeaderRowIndex_(values) {
  var limit = Math.min(10, values.length);
  for (var i = 0; i < limit; i++) {
    var has = values[i].some(function (cell) { return normHeader_(cell).indexOf('店舗名') !== -1; });
    if (has) return i;
  }
  return 0;
}

// ============================================================
// SECTION: 店舗グループ化
// ============================================================
function groupEstimateRowsByStore_(rows) {
  var map = {}, order = [];
  rows.forEach(function (r) {
    if (!map[r.store]) { map[r.store] = { store: r.store, rimei: '', rows: [] }; order.push(r.store); }
    var g = map[r.store];
    g.rows.push(r);
    if (!g.rimei && r.rimei) g.rimei = r.rimei; // グループの理/美は最初の非空値
  });

  return order.map(function (store) {
    var g = map[store];
    var items = g.rows.map(function (r, idx) {
      var amount = r.amount;
      // 金額が空で台数・単価があれば 金額 = 台数 × 単価
      if (amount === '' && r.qty !== '' && r.unit !== '') amount = r.qty * r.unit;
      return {
        no:       idx + 1,
        itemName: buildItemNameWithSize_(r.name, r.d, r.w, r.h),
        qty:      r.qty,
        unitLabel: EST_PIONEER.UNIT_LABEL,
        unitPrice: r.unit,
        amount:   amount,
        sAmount:  r.sAmount,
        note:     '',
      };
    });
    return { store: store, rimei: g.rimei, items: items };
  });
}

// ============================================================
// SECTION: タイトル・名称生成
// ============================================================
function buildEstimateTitle_(rimei, store) {
  var head = 'プラージュ';
  var r = String(rimei || '').trim();
  if (r.indexOf('理') !== -1)      head = '理容プラージュ';
  else if (r.indexOf('美') !== -1) head = '美容プラージュ';
  return head + '　' + store + '店　什器工事';
}

// {什器名}　D{D}　W{W}　H{H} （区切りは全角スペース。空欄の寸法は省略。全部空なら什器名のみ）
// 例: 「レジ台　D640　W900　H950」「時間指定配送」（寸法なし）
function buildItemNameWithSize_(name, d, w, h) {
  name = String(name || '').trim();
  var parts = [];
  if (notBlankDim_(d)) parts.push('D' + String(d).trim());
  if (notBlankDim_(w)) parts.push('W' + String(w).trim());
  if (notBlankDim_(h)) parts.push('H' + String(h).trim());
  if (!parts.length) return name;
  return (name ? name + '　' : '') + parts.join('　');
}
function notBlankDim_(v) { return v !== '' && v != null && String(v).trim() !== ''; }

// ============================================================
// SECTION: テンプレート(原本)を同じファイル内にコピー（既存運用に合わせる）
// ============================================================
// 別ファイルは作らず、テンプレファイルの中で『原本』シートを複製して
// 店舗別シート（例: 「理容　九条」）を増やす。
// 既存同名シートがあれば上書きせず、シート名末尾に _yyyyMMdd_HHmm を付与。
function copyEstimateTemplate_(store, rimei) {
  var tplSs = SpreadsheetApp.openById(EST_PIONEER.TEMPLATE_SS_ID);
  var tpl   = tplSs.getSheetByName(EST_PIONEER.TEMPLATE_SHEET);
  if (!tpl) throw new Error('テンプレートシート「' + EST_PIONEER.TEMPLATE_SHEET + '」が見つかりません');

  var sheetName = resolveEstimateSheetName_(tplSs, buildEstimateSheetName_(rimei, store));

  // 原本を同じスプレッドシート内に複製 → リネーム → 原本の右隣に配置
  var copied = tpl.copyTo(tplSs);
  copied.setName(sheetName);
  tplSs.setActiveSheet(copied);
  tplSs.moveActiveSheet(tpl.getIndex() + 1);

  var url = tplSs.getUrl() + '#gid=' + copied.getSheetId();
  return { sheet: copied, fileName: sheetName, url: url, id: copied.getSheetId() };
}

// 店舗別シート名（既存運用の「理容　九条」「美容　姫路花田」形式に合わせる）
function buildEstimateSheetName_(rimei, store) {
  var r = String(rimei || '').trim();
  if (r.indexOf('理') !== -1)      return '理容　' + store;
  if (r.indexOf('美') !== -1)      return '美容　' + store;
  return store; // 理/美 不明はそのまま店舗名
}

// 同名シートが既にあれば日時サフィックスを付ける（上書きしない）
function resolveEstimateSheetName_(ss, baseName) {
  if (!ss.getSheetByName(baseName)) return baseName;
  var stamp = Utilities.formatDate(new Date(), EST_PIONEER.TZ, 'yyyyMMdd_HHmm');
  return baseName + '_' + stamp;
}

// ============================================================
// SECTION: ヘッダー（タイトル文言）転記
// ============================================================
// 「理容or美容プラージュ　店　什器工事」等のプレースホルダーを実タイトルに置換。
function writeEstimateHeader_(sheet, title) {
  var values   = sheet.getDataRange().getValues();
  var formulas = sheet.getDataRange().getFormulas();
  var replaced = 0;
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      // 数式セル（例: B13 = =B5）は触らない。実体セルだけ置換すればミラーは自動追従。
      if (formulas[r][c]) continue;
      var s = String(values[r][c] == null ? '' : values[r][c]);
      if (s.indexOf('プラージュ') !== -1 && s.indexOf('什器工事') !== -1) {
        sheet.getRange(r + 1, c + 1).setValue(title);
        replaced++;
      }
    }
  }
  if (!replaced) console.warn('⚠️ タイトル用プレースホルダーが見つかりませんでした（テンプレ文言を確認してください）。');
}

// ============================================================
// SECTION: 明細転記
// ============================================================
// 完成見本に合わせて転記する。
//   1. ページ1明細域（itemStart〜小計の手前）を一旦クリア（旧D W H・旧P合計・旧注記・=F*H を消す）
//   2. 明細を連続で記入（No/名称/数量/単位/単価/金額(I,値)/P指値(L)/金額(M,値)）
//   3. 明細直後に P合計／金額合計／P合計-金額合計、その下に黄色の※注記
//   4. 小計・消費税・税込合計はテンプレ数式が自動計算（I列の値を集計）
// 戻り値: { subtotal, totalRows, totalCells, amountCol }
function writeEstimateDetails_(sheet, items) {
  var layout = detectEstimateLayout_(sheet);
  var cols = layout.cols;
  if (!cols || !layout.itemStartRow || !layout.subtotalRow1) {
    console.warn('⚠️ テンプレの明細構造を検出できませんでした。テンプレを確認してください。');
    return { subtotal: 0, totalRows: 0, totalCells: layout.totalCells, amountCol: layout.amountCol };
  }

  var startRow    = layout.itemStartRow;     // 例: 14
  var subtotalRow = layout.subtotalRow1;     // 例: 44
  var SUMMARY_RESERVE = 3;                    // P合計ラベル/値/注記 の3行ぶん
  var maxItemRows = subtotalRow - startRow - SUMMARY_RESERVE;
  var n = items.length;
  if (n > maxItemRows) {
    console.warn('⚠️ 明細が1ページ上限(' + maxItemRows + '行)を超えました。' + items.length + '件中 ' + maxItemRows + '件まで記入します。残りは手動対応してください。');
    n = maxItemRows;
  }

  // 1. ページ1明細域をクリア（罫線・結合は保持しつつ、テンプレ由来の塗り/文字色は除去）
  //    clearContent() は値だけ消すため、テンプレ『原本』に残る「黄色塗り＋赤文字」の
  //    サンプル注記が、空セルの書式（黄色背景・赤文字）だけ残って見えてしまう。
  //    背景色・文字色も既定に戻して消す（本来の※注記は明細直後に入れ直す）。
  var clearCount = subtotalRow - startRow;   // startRow〜subtotalRow-1
  if (clearCount > 0) {
    var clearRange = sheet.getRange(startRow, 1, clearCount, 14);
    clearRange.clearContent();
    clearRange.setBackground(null);   // 黄色塗りつぶしを除去
    clearRange.setFontColor(null);    // 赤文字を既定色に戻す
  }

  // 2. 明細を連続記入
  var subtotal = 0;
  for (var i = 0; i < n; i++) {
    var it  = items[i];
    var row = startRow + i;
    var amt = (it.amount !== '') ? it.amount
            : (it.qty !== '' && it.unitPrice !== '') ? it.qty * it.unitPrice : '';

    forceSetCell_(sheet, row, cols.no,        i + 1);
    forceSetCell_(sheet, row, cols.name,      it.itemName);
    forceSetCell_(sheet, row, cols.qty,       it.qty);
    forceSetCell_(sheet, row, cols.unitLabel, it.unitLabel);
    forceSetCell_(sheet, row, cols.unitPrice, it.unitPrice);
    forceSetCell_(sheet, row, cols.amount,    amt);          // I列: 金額（値）
    forceSetCell_(sheet, row, cols.sAmount,   it.sAmount);   // L列: P指値（金額）
    forceSetCell_(sheet, row, cols.amount2,   amt);          // M列: 金額（値・P指値セクション）

    if (typeof amt === 'number') subtotal += amt;
  }

  // 3. 明細直後に P合計／金額合計／P合計-金額合計 と 黄色の※注記
  writeEstimatePSummaryAndNote_(sheet, cols, startRow + n, items.slice(0, n));

  return { subtotal: subtotal, totalRows: n, totalCells: layout.totalCells, amountCol: layout.amountCol };
}

// 明細直後の P合計／金額合計／P合計-金額合計（ラベル行＋値行）と、黄色の※注記を書く
var ESTIMATE_NOTE_TEXT = '※　見積もり項目に記載の無いものは別途となります。';
var ESTIMATE_NOTE_BG   = '#ffff00';
function writeEstimatePSummaryAndNote_(sheet, cols, labelRow, items) {
  var L = cols.sAmount || 12;          // P指値（金額）列
  var M = cols.amount2 || 13;          // 金額列（P指値セクション）
  var N = M + 1;                       // P合計-金額合計列
  var valueRow = labelRow + 1;
  var noteRow  = valueRow + 1;

  var pTotal = 0, hasP = false, amtTotal = 0;
  items.forEach(function (it) {
    if (typeof it.sAmount === 'number') { pTotal += it.sAmount; hasP = true; }
    var a = (it.amount !== '') ? it.amount
          : (it.qty !== '' && it.unitPrice !== '') ? it.qty * it.unitPrice : 0;
    if (typeof a === 'number') amtTotal += a;
  });

  forceSetCell_(sheet, labelRow, L, 'P合計');
  forceSetCell_(sheet, labelRow, M, '金額合計');
  forceSetCell_(sheet, labelRow, N, 'P合計-金額合計');
  forceSetCell_(sheet, valueRow, L, hasP ? pTotal : '');
  forceSetCell_(sheet, valueRow, M, amtTotal);
  forceSetCell_(sheet, valueRow, N, (hasP ? pTotal : 0) - amtTotal);

  // ※注記（黄色塗りつぶし）。名称列(B)に記載。
  var noteCol = cols.name || 2;
  forceSetCell_(sheet, noteRow, noteCol, ESTIMATE_NOTE_TEXT);
  sheet.getRange(noteRow, noteCol).setBackground(ESTIMATE_NOTE_BG);
}

// セル書き込み（空欄なら空文字でクリア。テンプレの 0 やプレースホルダーを残さない）
function forceSetCell_(sheet, row, col, value) {
  if (!col || col < 1) return;
  sheet.getRange(row, col).setValue(value === '' || value == null ? '' : value);
}
// セルをクリア（列が無効なら何もしない）
function clearCell_(sheet, row, col) {
  if (!col || col < 1) return;
  sheet.getRange(row, col).setValue('');
}

// ----- テンプレ構造検出 -----
// このテンプレは金額(I列)が各行 =SUM(Fx*Hx) の数式。
// 「金額セルが自分の行の F*H 数式を持つ行」を明細スロットとみなす。
//   - 見出し行（名称に『プラージュ／什器工事』を含む）は除外
//   - 小計行(SUBTOTAL/SUM範囲)・ヘッダー行は F*H ではないので自然に除外される
function detectEstimateLayout_(sheet) {
  var values   = sheet.getDataRange().getValues();
  var formulas = sheet.getDataRange().getFormulas();

  // 列マップは最初の明細ヘッダー行から取得
  var cols = null;
  for (var r = 0; r < values.length; r++) {
    if (rowIsDetailHeader_(values[r])) { cols = buildDetailColMap_(values[r].map(normHeader_)); break; }
  }
  if (!cols) return { cols: null, slots: [], itemStartRow: 0, subtotalRow1: 0, totalCells: detectTotalCells_(values, 0), amountCol: 0 };

  var amountCol = cols.amount || 9;
  var nameCol   = cols.name   || 2;
  var slots = [];
  for (var rr = 0; rr < values.length; rr++) {
    var f = String(formulas[rr][amountCol - 1] || '');
    // 自己行参照の F<row>*H<row>（金額=数量×単価）を持つ行＝明細スロット
    if (!(new RegExp('F' + (rr + 1) + '\\*H' + (rr + 1))).test(f)) continue;
    var nm = String(values[rr][nameCol - 1] || '');
    // 見出し行（…プラージュ…什器工事）と注記行（※/別途）はスロットから除外
    if (nm.indexOf('プラージュ') !== -1 || nm.indexOf('什器工事') !== -1) continue;
    if (nm.indexOf('※') !== -1 || nm.indexOf('別途') !== -1) continue;
    slots.push({ row: rr + 1, cols: cols });
  }

  var itemStartRow = slots.length ? slots[0].row : 0;
  // ページ1の「小計」行（明細域の終端）。最初に出現する「小計」セルの行。
  var subtotalRow1 = 0;
  for (var sr = 0; sr < values.length; sr++) {
    for (var sc = 0; sc < values[sr].length; sc++) {
      if (normHeader_(values[sr][sc]) === '小計') { subtotalRow1 = sr + 1; break; }
    }
    if (subtotalRow1) break;
  }

  return {
    cols: cols, slots: slots, itemStartRow: itemStartRow, subtotalRow1: subtotalRow1,
    totalCells: detectTotalCells_(values, amountCol), amountCol: amountCol,
  };
}

// ヘッダー行から各列のindex(1始まり)を作る
function buildDetailColMap_(rowN) {
  function find() {
    for (var a = 0; a < arguments.length; a++) {
      for (var i = 0; i < rowN.length; i++) {
        if (rowN[i] === arguments[a]) return i + 1;
      }
    }
    // 部分一致フォールバック
    for (var a2 = 0; a2 < arguments.length; a2++) {
      for (var j = 0; j < rowN.length; j++) {
        if (rowN[j] && rowN[j].indexOf(arguments[a2]) !== -1) return j + 1;
      }
    }
    return 0;
  }
  // N番目の完全一致を返す（「金額」が複数列ある構成に対応）
  function findNth(name, n) {
    var t = normHeader_(name), count = 0;
    for (var i = 0; i < rowN.length; i++) {
      if (rowN[i] === t) { count++; if (count === n) return i + 1; }
    }
    return 0;
  }
  // 金額は2か所：主明細=最初の「金額」(I列)、P指値セクション=2番目の「金額」(M列)。
  // sAmount は「P指値(金額)」(L列)。
  return {
    no:        find('No.', 'No', 'no'),
    name:      find('名称'),
    qty:       find('数量'),
    unitLabel: find('単位'),
    unitPrice: find('単価'),
    amount:    find('金額'),
    amount2:   findNth('金額', 2),
    note:      find('備考'),
    sAmount:   find('P指値(金額)', 'P指値（金額）', 'P指値'),
  };
}

function rowIsDetailHeader_(row) {
  var rowN = row.map(normHeader_);
  return rowN.some(function (x) { return x.indexOf('名称') !== -1; }) &&
         rowN.some(function (x) { return x.indexOf('数量') !== -1; }) &&
         rowN.some(function (x) { return x.indexOf('単価') !== -1; });
}

function rowHasTotalKeyword_(row) {
  return row.some(function (cell) {
    var s = normHeader_(cell);
    return s.indexOf('小計') !== -1 || s.indexOf('消費税') !== -1 ||
           s.indexOf('税込') !== -1 || s === '合計' || s.indexOf('合計') !== -1;
  });
}

// 小計・消費税・税込合計ラベルの「金額を入れるべきセル」を推定。
// ラベルセルの右隣以降で最初の数値セルを値入力先とする（テンプレ行10: 税込合計=C / 消費税=F / 合計=I）。
function detectTotalCells_(values, amountCol) {
  var out = { subtotal: null, tax: null, total: null };
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var s = normHeader_(values[r][c]);
      if (!s) continue;
      var target = { row: r + 1, col: findValueColRight_(values[r], c, amountCol) };
      // 判定順が重要：「税込合計」は '税込' を含むので最初に。'小計' と '合計'(単独) は小計扱い。
      if (s.indexOf('税込') !== -1)        { if (!out.total)    out.total    = target; }
      else if (s.indexOf('消費税') !== -1) { if (!out.tax)      out.tax      = target; }
      else if (s.indexOf('小計') !== -1)   { if (!out.subtotal) out.subtotal = target; }
      else if (s === '合計')               { if (!out.subtotal) out.subtotal = target; }
    }
  }
  return out;
}

// ラベルセル(列index c)の右隣以降で、最初に「数値が入っているセル」の列番号(1始まり)を返す。
// 見つからなければ amountCol、それも無ければ c+2 を仮の入力先とする。
function findValueColRight_(row, c, amountCol) {
  for (var k = c + 1; k < row.length; k++) {
    if (typeof row[k] === 'number') return k + 1;
  }
  return amountCol || (c + 2);
}

// ============================================================
// SECTION: 合計反映（既存数式を尊重）
// ============================================================
function calculateEstimateTotals_(sheet, totals) {
  var subtotal = totals.subtotal || 0;
  var tax      = Math.round(subtotal * EST_PIONEER.TAX_RATE);
  var total    = subtotal + tax;
  var cells    = totals.totalCells || {};

  // 既存数式があるセルは尊重し、値で上書きしない
  putTotalIfNoFormula_(sheet, cells.subtotal, subtotal);
  putTotalIfNoFormula_(sheet, cells.tax,      tax);
  putTotalIfNoFormula_(sheet, cells.total,    total);
}

function putTotalIfNoFormula_(sheet, cell, value) {
  if (!cell || !cell.col || cell.col < 1) return;
  var range = sheet.getRange(cell.row, cell.col);
  if (range.getFormula()) return; // 数式があるなら触らない
  range.setValue(value);
}

// ============================================================
// SECTION: 結果ログ
// ============================================================
function logEstimateResult_(log) {
  console.log('==== 見積書 自動作成 結果 ====');
  console.log('作成した見積書の件数: ' + log.created);
  console.log('対象店舗数: '          + log.stores);
  console.log('対象明細行数: '        + log.detailRows);
  console.log('スキップした行数: '    + log.skipped);
  if (log.files.length) {
    console.log('---- 作成シート ----');
    log.files.forEach(function (f) {
      console.log('・' + f.store + ' → シート「' + f.name + '」');
      console.log('   URL: ' + f.url);
    });
  } else {
    console.log('（作成されたシートはありません）');
  }
}

// ============================================================
// SECTION: 将来のLINE連携用ラッパー（今は配線しない）
// ------------------------------------------------------------
// 将来「見積作成」等で呼べるようにする場合は、既存 handleDM 内の分岐に
//   if (handleEstimateCommand(ev, text)) return;
// を1行追加するだけにする（doPost 本体や署名検証は触らない）。
// ============================================================
function handleEstimateCommand(event, messageText) {
  if (!isPioneerEstimateCommand_(messageText)) return false;
  var reply;
  try {
    var isTest = isPioneerEstimateTestCommand_(messageText);
    // メッセージに店舗名が含まれていれば、その店舗だけ作成される
    var log = runEstimateCore_(isTest ? 1 : Infinity, messageText);
    reply = formatPioneerEstimateReply_(log, isTest);
  } catch (e) {
    console.error('handleEstimateCommand error: ' + e.message + ' ' + (e.stack || ''));
    reply = '⚠️ 見積書作成中にエラーが発生しました: ' + e.message;
  }
  if (event && event.replyToken && typeof sendLineReply === 'function') {
    sendLineReply(event.replyToken, reply);
  }
  return true;
}

// 「パイオニア 見積…作って」等で発火。照会・進捗質問では発火しない（作成動詞が必須）。
function isPioneerEstimateCommand_(text) {
  if (!text) return false;
  var t = String(text);
  if (!/パイオニア|什器/.test(t)) return false;   // 対象が明示されていること
  if (!/見積/.test(t)) return false;
  // 作成系の動詞があるときだけ（「どうなってる？」等の進捗質問では発火させない）
  return /(作って|作成|作っといて|つくって|生成|出して|お願い)/.test(t);
}

// テスト指定（先頭1店舗だけ作成）。「テスト」「お試し」「1店舗」等を含むとき。
function isPioneerEstimateTestCommand_(text) {
  if (!text) return false;
  return /テスト|ﾃｽﾄ|お試し|おためし|1店舗|一店舗|試し/.test(String(text));
}

// 実行結果(log)をLINE返信用テキストに整形
function formatPioneerEstimateReply_(log, isTest) {
  if (!log)      return '⚠️ 見積書作成の結果を取得できませんでした。GASの実行ログをご確認ください。';
  if (log.error) return '⚠️ 見積書を作成できませんでした：\n' + log.error;
  if (!log.created) {
    return 'ℹ️ ' + (log.note || '対象行が見つからず、見積書は作成されませんでした。');
  }
  var lines = [];
  lines.push((isTest ? '【テスト】' : '') + 'パイオニア什器の見積書を作成しました。');
  lines.push('');
  if (log.requestedStores && log.requestedStores.length) {
    lines.push('指定店舗：' + log.requestedStores.join('、'));
  }
  lines.push('作成件数：' + log.created + '件');
  lines.push('対象店舗数：' + log.stores + '店舗');
  lines.push('明細行数：' + log.detailRows + '行');
  if (log.skipped) lines.push('スキップ：' + log.skipped + '行（店舗名空欄）');
  if (log.files && log.files.length) {
    lines.push('');
    lines.push('作成シート：');
    log.files.slice(0, 20).forEach(function (f) { lines.push('・' + f.store + '（' + f.name + '）'); });
    if (log.files.length > 20) lines.push('…ほか ' + (log.files.length - 20) + ' 件');
    var baseUrl = String(log.files[0].url || '').split('#')[0];
    if (baseUrl) { lines.push(''); lines.push('見積書：' + baseUrl); }
  }
  if (isTest) {
    lines.push('');
    lines.push('※テスト実行（先頭1店舗のみ）。全店舗作成するには「パイオニア 見積書作って」と送ってください。');
  }
  return lines.join('\n');
}

// ============================================================
// SECTION: 共通ユーティリティ（このファイル内に閉じる）
// ============================================================
// ヘッダー比較用：全角半角統一（NFKC）+ 空白除去
function normHeader_(v) {
  return String(v == null ? '' : v).normalize('NFKC').replace(/\s+/g, '').trim();
}
// 値比較用：NFKC + trim（内部空白は残す）
function normCmp_(v) {
  return String(v == null ? '' : v).normalize('NFKC').trim();
}
// 数値化：空欄は '' を返す。カンマ・通貨記号を許容。数値化不能も '' 。
function toNum_(v) {
  if (v === '' || v == null) return '';
  if (typeof v === 'number') return v;
  var s = String(v).replace(/[,¥\s　]/g, '').trim();
  if (s === '') return '';
  var n = Number(s);
  return isNaN(n) ? '' : n;
}
// 文字列化（寸法用）：数値はそのまま文字列に、空欄は ''
function cellStr_(v) {
  if (v === '' || v == null) return '';
  return String(v).trim();
}

// ============================================================
// SECTION: 診断（実テンプレ／実シートを開いて構造を確認）
//   GASエディタで diagnoseEstimateTemplate_ を実行 → ログで当たりを確認。
//   検出が外れていたら detect* の条件か EST_PIONEER のシート名を調整。
// ============================================================
// ★ GASエディタの実行プルダウン用（"_"なし＝一覧に表示される）。
//   こちらを選んで「▷ 実行」すると診断が走ります。
function runEstimateDiagnosis() {
  diagnoseEstimateTemplate_();
}

function diagnoseEstimateTemplate_() {
  // --- 参照元 制作一覧 ---
  try {
    var src = SpreadsheetApp.openById(EST_PIONEER.SOURCE_SS_ID).getSheetByName(EST_PIONEER.SOURCE_SHEET);
    if (src) {
      var hr = findHeaderRowIndex_(src.getDataRange().getValues());
      console.log('[制作一覧] ヘッダー行 index=' + hr + ' / 内容: ' +
        JSON.stringify(src.getRange(hr + 1, 1, 1, src.getLastColumn()).getValues()[0]));
    } else {
      console.log('[制作一覧] シートが見つかりません: ' + EST_PIONEER.SOURCE_SHEET);
    }
  } catch (e) { console.log('[制作一覧] 開けません: ' + e.message); }

  // --- テンプレート ---
  try {
    var tpl = SpreadsheetApp.openById(EST_PIONEER.TEMPLATE_SS_ID).getSheetByName(EST_PIONEER.TEMPLATE_SHEET);
    if (!tpl) {
      console.log('[テンプレ] シートが見つかりません: ' + EST_PIONEER.TEMPLATE_SHEET);
      console.log('[テンプレ] 候補シート: ' + SpreadsheetApp.openById(EST_PIONEER.TEMPLATE_SS_ID).getSheets().map(function (s) { return s.getName(); }).join(' / '));
      return;
    }
    var layout = detectEstimateLayout_(tpl);
    console.log('[テンプレ] 検出スロット数: ' + layout.slots.length + ' / 金額列: ' + layout.amountCol);
    if (layout.slots.length) {
      console.log('[テンプレ] 列マップ: ' + JSON.stringify(layout.slots[0].cols));
      console.log('[テンプレ] 明細スロット先頭5行の行番号: ' +
        layout.slots.slice(0, 5).map(function (s) { return s.row; }).join(','));
    }
    console.log('[テンプレ] 合計セル: ' + JSON.stringify(layout.totalCells));

    // --- 数式の有無を確認（書き込みで数式を壊さないか判断するため）---
    var fRange = tpl.getRange(1, 1, Math.min(46, tpl.getLastRow()), Math.min(13, tpl.getLastColumn()));
    var formulas = fRange.getFormulas();
    var fLines = [];
    for (var fr = 0; fr < formulas.length; fr++) {
      for (var fc = 0; fc < formulas[fr].length; fc++) {
        if (formulas[fr][fc]) fLines.push('R' + (fr + 1) + 'C' + (fc + 1) + '=' + formulas[fr][fc]);
      }
    }
    console.log('[テンプレ] 数式セル(1〜46行): ' + (fLines.length ? fLines.join('  /  ') : '（数式なし＝全て値）'));

    var n = Math.min(20, tpl.getLastRow());
    var d = tpl.getRange(1, 1, n, tpl.getLastColumn()).getValues();
    for (var i = 0; i < n; i++) console.log('行' + (i + 1) + ': ' + JSON.stringify(d[i]));

    // --- 注記「※…」と各種「合計」ラベルの位置・背景色を調べる ---
    var lastRow = tpl.getLastRow(), lastCol = tpl.getLastColumn();
    var allV  = tpl.getRange(1, 1, lastRow, lastCol).getValues();
    var allBg = tpl.getRange(1, 1, lastRow, lastCol).getBackgrounds();
    var notes = [], totals = [];
    for (var rr = 0; rr < lastRow; rr++) {
      for (var cc = 0; cc < lastCol; cc++) {
        var s = String(allV[rr][cc] == null ? '' : allV[rr][cc]);
        if (!s) continue;
        if (s.indexOf('※') !== -1 || s.indexOf('別途') !== -1) {
          notes.push('R' + (rr + 1) + 'C' + (cc + 1) + ' bg=' + allBg[rr][cc] + ' 「' + s.slice(0, 30) + '」');
        }
        if (/小計|合計|P合計|金額合計/.test(s)) {
          totals.push('R' + (rr + 1) + 'C' + (cc + 1) + ' bg=' + allBg[rr][cc] + ' 「' + s + '」');
        }
      }
    }
    console.log('[テンプレ] 注記(※/別途)セル: ' + (notes.join('  /  ') || 'なし'));
    console.log('[テンプレ] 合計ラベルセル: ' + (totals.join('  /  ') || 'なし'));

    // --- 末尾20行をダンプ（最終合計まわりの構造確認）---
    var bStart = Math.max(1, lastRow - 19);
    var bd = tpl.getRange(bStart, 1, lastRow - bStart + 1, lastCol).getValues();
    var bf = tpl.getRange(bStart, 1, lastRow - bStart + 1, lastCol).getFormulas();
    for (var bi = 0; bi < bd.length; bi++) {
      var rowNo = bStart + bi;
      var merged = bd[bi].map(function (v, idx) { return bf[bi][idx] ? ('=' + bf[bi][idx]) : v; });
      console.log('行' + rowNo + ': ' + JSON.stringify(merged));
    }
  } catch (e) { console.log('[テンプレ] 開けません: ' + e.message); }
}

// build: pioneer-estimate v1
