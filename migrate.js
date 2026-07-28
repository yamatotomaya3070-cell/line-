// ==========================================
// アカウント移行ツール
// 使い方: GASエディタで関数を選んで▶実行するだけ
// ==========================================

// ① まずこれを実行して現状確認
function checkOwnership() {
  var folderId = getConfig().DRIVE_FOLDER_ID;
  var folder = DriveApp.getFolderById(folderId);
  var myEmail = Session.getActiveUser().getEmail();

  console.log('=== 現在の実行アカウント ===');
  console.log('メール: ' + myEmail);
  console.log('フォルダ名: ' + folder.getName());
  console.log('フォルダオーナー: ' + folder.getOwner().getEmail());

  var fileCount = 0;
  var notMineCount = 0;
  countFiles(folder, myEmail, function(owned) {
    fileCount++;
    if (!owned) notMineCount++;
  });

  console.log('総ファイル数: ' + fileCount);
  console.log('オーナーが自分でないファイル: ' + notMineCount);
  console.log(notMineCount === 0 ? '✅ 全ファイル移行済み' : '⚠️ まだ移行が必要なファイルがあります');
}

function countFiles(folder, email, callback) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    try { callback(f.getOwner().getEmail() === email); } catch(e) { callback(false); }
  }
  var subs = folder.getFolders();
  while (subs.hasNext()) { countFiles(subs.next(), email, callback); }
}

// ② オーナー移行を実行（新アカウントのメールを入力してから実行）
function transferOwnershipToNewAccount() {
  var NEW_ACCOUNT_EMAIL = 'admin@woodbasef.com';

  if (!NEW_ACCOUNT_EMAIL) {
    console.log('❌ NEW_ACCOUNT_EMAIL を入力してから実行してください');
    return;
  }

  var folderId = getConfig().DRIVE_FOLDER_ID;
  var folder = DriveApp.getFolderById(folderId);

  console.log('移行開始: ' + folder.getName() + ' → ' + NEW_ACCOUNT_EMAIL);

  var result = { success: 0, skip: 0, error: 0 };
  transferFolder(folder, NEW_ACCOUNT_EMAIL, result);

  console.log('=== 完了 ===');
  console.log('成功: ' + result.success + ' 件');
  console.log('スキップ: ' + result.skip + ' 件');
  console.log('エラー: ' + result.error + ' 件');
  console.log('次にcheckOwnership()を実行して確認してください');
}

function transferFolder(folder, newOwner, result) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    try {
      if (file.getOwner().getEmail() !== newOwner) {
        file.setOwner(newOwner);
        result.success++;
        console.log('✅ ' + file.getName());
      } else {
        result.skip++;
      }
    } catch(e) {
      result.error++;
      console.log('⚠️ スキップ: ' + file.getName() + ' (' + e.message + ')');
    }
  }
  var subs = folder.getFolders();
  while (subs.hasNext()) {
    transferFolder(subs.next(), newOwner, result);
  }
}

// ==========================================
// 仮タスク → タスク管理 一括移行
// 対象：期日が2026-05-01以降 または 期日なし
// GASエディタで「migratePendingFromMay1」を選んで▶実行
// ==========================================
function migratePendingFromMay1() {
  var pendingSheet = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')
  ).getSheetByName('仮タスク');
  var taskSheet = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')
  ).getSheetByName('タスク管理');
  var schedSheet = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')
  ).getSheetByName('スケジュール管理');

  if (!pendingSheet || !taskSheet) { console.log('シートが見つかりません'); return; }

  var data       = pendingSheet.getDataRange().getValues();
  var cutoff     = '2026-05-01';
  var now        = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  var rowsToDel  = [];
  var taskCount  = 0;
  var schedCount = 0;

  // 列: [0]=batch_id [1]=idx [2]=type [3]=group_id [4]=user_id
  //     [5]=案件名 [6]=内容/title [7]=担当者 [8]=期日/日時 [9]=作成日時 [10]=元msg [11]=extraJSON
  for (var i = data.length - 1; i >= 1; i--) {
    var row      = data[i];
    var type     = String(row[2]);
    var dateStr  = String(row[8] || '').slice(0, 10); // YYYY-MM-DD部分だけ取得
    var hasDate  = /^\d{4}-\d{2}-\d{2}/.test(dateStr);

    // 期日が過去（5/1より前）のものはスキップ
    if (hasDate && dateStr < cutoff) continue;

    if (type === 'task') {
      var extra = {};
      try { extra = JSON.parse(String(row[11] || '{}')); } catch(e) {}
      taskSheet.appendRow([
        generateMigrateId(),   // task_id
        row[5] || '未分類',    // 案件名
        row[6] || '',          // タスク内容
        row[7] || '',          // 担当者
        hasDate ? dateStr : '', // 期日
        'confirmed',           // ステータス
        now,                   // 作成日時
        row[3] || '',          // グループID
        extra.urgency || '',   // 緊急度
      ]);
      taskCount++;
    } else if (type === 'schedule' && schedSheet) {
      var extra2 = {};
      try { extra2 = JSON.parse(String(row[11] || '{}')); } catch(e) {}
      var parts = String(row[8] || '').split(' ');
      schedSheet.appendRow([
        now,                      // 登録日時
        row[5] || '未分類',       // 案件名
        row[6] || '',             // 予定タイトル
        parts[0] || '',           // 日付
        parts[1] || '',           // 開始時間
        extra2.endTime || '',     // 終了時間
        extra2.location || '',    // 場所
        row[7] || '',             // 参加者
        extra2.description || '', // 詳細
        row[3] || '',             // グループID
      ]);
      schedCount++;
    }

    rowsToDel.push(i + 1);
  }

  // 仮タスクから削除（下から順に）
  rowsToDel.forEach(function(r) { pendingSheet.deleteRow(r); });

  console.log('✅ 移行完了');
  console.log('タスク登録: ' + taskCount + '件');
  console.log('予定登録:   ' + schedCount + '件');
  console.log('仮タスク削除: ' + rowsToDel.length + '行');
}

function generateMigrateId() {
  return Math.random().toString(36).slice(2, 14);
}
