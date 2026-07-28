// ==========================================
// admin@woodbasef.com への移行ヘルパー（コピー方式）
// ------------------------------------------
// 背景: woodbasegroup@gmail.com(個人) と admin@woodbasef.com(Workspace) は
//       別組織のため、Drive の所有権は直接移行できない（setOwner 不可）。
//       → 「admin の Drive / 共有ドライブにコピーして、Script Properties を
//          新しいIDに差し替える」のが確実な方法。
//
// 【最重要】必ず admin@woodbasef.com でログインした GASエディタから実行すること。
//          コピーしたファイルは実行ユーザーの所有になるため。
//
// 使い方（上から順に GASエディタで関数を選んで▶実行）:
//   1. migrateCheck()             … 現状確認（破壊なし）
//   2. migrateCopySpreadsheet()   … スプレッドシートをコピーして ID 差し替え
//   3. migrateCopyDriveFolder()   … 案件フォルダツリーをコピーして ID 差し替え
//   4. migrateCreateCalendar()    … admin 所有カレンダーを作成して ID 差し替え
//   （問題があれば migrateRollback() で旧IDに戻す）
//
// 旧IDは "OLD_SPREADSHEET_ID" のように退避保存するので、いつでもロールバック可。
// ==========================================

// 移行先（共有ドライブのフォルダID等）。空ならログインユーザー(admin)のマイドライブ直下。
// ★共有ドライブを使う場合はここに共有ドライブ内フォルダのIDを入れてから実行。
var MIGRATE_DEST_FOLDER_ID = '0AC5ryYAAXxS8Uk9PVA';

var MIGRATE_TARGET_EMAIL = 'admin@woodbasef.com';

// ------------------------------------------
// 共通ユーティリティ
// ------------------------------------------
function _mProps() { return PropertiesService.getScriptProperties(); }

function _mGet(key) { return _mProps().getProperty(key); }

// 旧値を OLD_ に退避してから新値を保存
function _mSwap(key, newValue) {
  var props = _mProps();
  var old = props.getProperty(key);
  if (old && !props.getProperty('OLD_' + key)) {
    props.setProperty('OLD_' + key, old);
  }
  props.setProperty(key, newValue);
  console.log('  ' + key + ': ' + old + ' → ' + newValue);
}

function _mDestParent() {
  if (MIGRATE_DEST_FOLDER_ID) {
    return DriveApp.getFolderById(MIGRATE_DEST_FOLDER_ID);
  }
  return DriveApp.getRootFolder(); // 実行ユーザー(admin)のマイドライブ
}

function _mActiveEmail() {
  // getEffectiveUser() は「実行者本人」なのでクロスドメインでも確実に取れる。
  try {
    var eff = Session.getEffectiveUser().getEmail();
    if (eff) return eff;
  } catch (e1) {}
  // フォールバック（同ドメイン時のみ取れる）
  try {
    var act = Session.getActiveUser().getEmail();
    if (act) return act;
  } catch (e2) {}
  return '';
}

// 実行アカウントが admin か確認。違えばエラーで止める。
function _mAssertAdmin() {
  var email = _mActiveEmail();
  if (!email) {
    // クロスドメインだと getEmail() が空になることがある。
    // 共有ドライブ(0AC5...)への書き込みは admin 以外では権限エラーになるため、
    // 特定できない場合は警告のみで続行（誤アカウントなら自然にコピーで失敗する）。
    console.log('⚠️ 実行アカウントを特定できませんでした（クロスドメインの仕様）。');
    console.log('   admin@woodbasef.com で実行している前提で続行します。');
    console.log('   違うアカウントなら共有ドライブへの書き込みで権限エラーになります。');
    return '(特定不可)';
  }
  if (email !== MIGRATE_TARGET_EMAIL) {
    throw new Error(
      '実行アカウントが ' + email + ' です。' +
      MIGRATE_TARGET_EMAIL + ' でログインし直してから実行してください。'
    );
  }
  return email;
}

// ------------------------------------------
// 1. 現状確認（破壊なし）
// ------------------------------------------
function migrateCheck() {
  console.log('=== 移行前チェック ===');
  console.log('実行アカウント: ' + _mActiveEmail());
  console.log('移行先メール  : ' + MIGRATE_TARGET_EMAIL);
  console.log('コピー先      : ' + (MIGRATE_DEST_FOLDER_ID
    ? ('指定フォルダ ' + MIGRATE_DEST_FOLDER_ID) : 'admin のマイドライブ直下'));
  console.log('');

  _mReportResource('SPREADSHEET_ID', function (id) {
    var file = DriveApp.getFileById(id);
    return file.getName() + ' / owner: ' + _mOwnerEmail(file);
  });
  _mReportResource('DRIVE_FOLDER_ID', function (id) {
    var folder = DriveApp.getFolderById(id);
    return folder.getName() + ' / owner: ' + _mOwnerEmail(folder);
  });
  _mReportResource('CALENDAR_ID', function (id) {
    var cal = CalendarApp.getCalendarById(id);
    return cal ? cal.getName() : '(アクセス不可)';
  });

  console.log('');
  console.log('問題なければ migrateCopySpreadsheet() → migrateCopyDriveFolder() → migrateCreateCalendar() の順で実行してください。');
}

function _mReportResource(key, describe) {
  var id = _mGet(key);
  if (!id) { console.log(key + ': 未設定'); return; }
  try {
    console.log(key + ': ' + describe(id));
    console.log('  (id=' + id + ')');
  } catch (e) {
    console.log(key + ': アクセスできません — ' + e.message);
    console.log('  → 旧アカウント側で admin に閲覧共有してください (id=' + id + ')');
  }
}

function _mOwnerEmail(fileOrFolder) {
  try {
    var owner = fileOrFolder.getOwner();
    return owner ? owner.getEmail() : '共有ドライブ(組織所有)';
  } catch (e) {
    return '共有ドライブ(組織所有)';
  }
}

// ------------------------------------------
// 2. スプレッドシートをコピー
// ------------------------------------------
function migrateCopySpreadsheet() {
  _mAssertAdmin();
  var oldId = _mGet('SPREADSHEET_ID');
  if (!oldId) { console.log('SPREADSHEET_ID 未設定のため中止'); return; }

  var src = DriveApp.getFileById(oldId);
  var dest = _mDestParent();
  var newName = src.getName() + ' (admin移行 ' + _mStamp() + ')';

  console.log('スプレッドシートをコピー中: ' + src.getName());
  var copy = src.makeCopy(newName, dest);
  var newId = copy.getId();

  console.log('=== スプレッドシート コピー完了 ===');
  _mSwap('SPREADSHEET_ID', newId);
  console.log('新シートURL: ' + copy.getUrl());
  console.log('※このシートを案件メンバー/グループに共有し直してください。');
}

// ------------------------------------------
// 3. 案件フォルダツリーをコピー（中断・再開対応）
//    ・約5分で安全に中断し、もう一度実行すると続きから再開
//    ・コピー先に同名ファイルが既にあればスキップ（再実行しても二重コピーしない）
//    ・全部終わったときだけ DRIVE_FOLDER_ID を切替
// ------------------------------------------
var MIGRATE_TIME_LIMIT_MS = 4 * 60 * 1000; // 4分（5分トリガー間隔と重ならないよう手前で中断）

function migrateCopyDriveFolder() {
  _mAssertAdmin();
  var props = _mProps();
  var oldId = _mGet('DRIVE_FOLDER_ID');
  if (!oldId) { console.log('DRIVE_FOLDER_ID 未設定のため中止'); return; }

  var srcFolder = DriveApp.getFolderById(oldId);

  // 中断・再開用：コピー先ルートを保存しておき、再実行時は同じ場所に続ける
  var newRoot;
  var savedRootId = props.getProperty('MIGRATE_DRIVE_NEWROOT');
  if (savedRootId) {
    newRoot = DriveApp.getFolderById(savedRootId);
    console.log('再開: 既存コピー先を継続使用 → ' + newRoot.getName());
  } else {
    newRoot = _mDestParent().createFolder(srcFolder.getName());
    props.setProperty('MIGRATE_DRIVE_NEWROOT', newRoot.getId());
    console.log('新規コピー先を作成: ' + newRoot.getName());
  }

  var ctx = { start: new Date().getTime(), files: 0, skipped: 0, folders: 0, errors: 0, timedOut: false };
  console.log('フォルダツリーをコピー開始: ' + srcFolder.getName());

  _mCopyFolderInto(srcFolder, newRoot, ctx);

  console.log('--- 進捗 ---');
  console.log('コピー: ' + ctx.files + ' / スキップ(既存): ' + ctx.skipped +
              ' / 新規フォルダ: ' + ctx.folders + ' / エラー: ' + ctx.errors);

  if (ctx.timedOut) {
    console.log('⏸ 時間制限のため中断しました。');
    console.log('   もう一度 migrateCopyDriveFolder() を実行すると続きから再開します。');
    return;
  }

  // 完了：IDを切替してテンポラリを掃除
  _mSwap('DRIVE_FOLDER_ID', newRoot.getId());
  props.deleteProperty('MIGRATE_DRIVE_NEWROOT');
  console.log('=== フォルダ コピー完了 ===');
  console.log('新フォルダURL: ' + newRoot.getUrl());
}

function _mTimeUp(ctx) {
  return (new Date().getTime() - ctx.start) > MIGRATE_TIME_LIMIT_MS;
}

function _mCopyFolderInto(srcFolder, destFolder, ctx) {
  if (ctx.timedOut) return;

  // 既存ファイル名を「一括取得してメモリ照合」（API呼び出しを激減させ再開を高速化）
  var existingFiles = {};
  var df = destFolder.getFiles();
  while (df.hasNext()) { existingFiles[df.next().getName()] = true; }

  var files = srcFolder.getFiles();
  while (files.hasNext()) {
    if (_mTimeUp(ctx)) { ctx.timedOut = true; return; }
    var f = files.next();
    var name = f.getName();
    try {
      if (existingFiles[name]) {
        ctx.skipped++;
      } else {
        f.makeCopy(name, destFolder);
        existingFiles[name] = true; // 同一フォルダ内の同名重複コピーを防止
        ctx.files++;
        if (ctx.files % 20 === 0) console.log('  ...' + ctx.files + ' 件コピー');
      }
    } catch (e) {
      ctx.errors++;
      console.log('⚠️ コピー失敗: ' + name + ' (' + e.message + ')');
    }
  }

  // サブフォルダも一括取得してメモリ照合（同名は作らず再利用）
  var existingSubs = {};
  var ds = destFolder.getFolders();
  while (ds.hasNext()) { var d = ds.next(); existingSubs[d.getName()] = d; }

  var subs = srcFolder.getFolders();
  while (subs.hasNext()) {
    if (_mTimeUp(ctx)) { ctx.timedOut = true; return; }
    var sub = subs.next();
    var sname = sub.getName();
    var newSub = existingSubs[sname];
    if (!newSub) {
      newSub = destFolder.createFolder(sname);
      existingSubs[sname] = newSub;
      ctx.folders++;
    }
    _mCopyFolderInto(sub, newSub, ctx);
    if (ctx.timedOut) return;
  }
}

// ------------------------------------------
// 3'. 自動コピー（おすすめ）
//     migrateStartAutoCopy() を1回実行すれば、5分ごとに自動でバックグラウンド実行。
//     エディタを閉じてもOK。全部終わると自動で停止する。
//     手動で止めたいときは migrateStopAutoCopy()。
// ------------------------------------------
function migrateStartAutoCopy() {
  _mAssertAdmin();
  _mDeleteAutoTriggers();
  ScriptApp.newTrigger('migrateAutoCopyTick').timeBased().everyMinutes(5).create();
  console.log('=== 自動コピー開始 ===');
  console.log('5分ごとにバックグラウンドで実行します。エディタは閉じてOK。');
  console.log('完了すると自動停止します。進捗は左メニュー「実行数」で確認できます。');
  console.log('途中で止めるときは migrateStopAutoCopy() を実行。');
  console.log('--- まず1回目を実行 ---');
  migrateAutoCopyTick();
}

function migrateAutoCopyTick() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) { console.log('別の実行が稼働中のためスキップ'); return; }
  try {
    // 既に完了済みなら停止
    if (!_mGet('MIGRATE_DRIVE_NEWROOT') && _mGet('OLD_DRIVE_FOLDER_ID')) {
      console.log('コピーは既に完了済み。自動トリガーを停止します。');
      _mDeleteAutoTriggers();
      return;
    }
    migrateCopyDriveFolder();
    // 完了判定：MIGRATE_DRIVE_NEWROOT が消えていれば完了
    if (!_mGet('MIGRATE_DRIVE_NEWROOT')) {
      console.log('✅ コピー完了を検知。自動トリガーを停止します。');
      _mDeleteAutoTriggers();
    }
  } finally {
    lock.releaseLock();
  }
}

function migrateStopAutoCopy() {
  _mDeleteAutoTriggers();
  console.log('自動コピーのトリガーを停止しました。');
}

function _mDeleteAutoTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'migrateAutoCopyTick') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// ------------------------------------------
// 4. admin 所有カレンダーを作成
// ------------------------------------------
function migrateCreateCalendar() {
  _mAssertAdmin();
  var name = 'WOODBASE 案件カレンダー (admin)';
  var newCal = CalendarApp.createCalendar(name, { timeZone: 'Asia/Tokyo' });
  console.log('=== カレンダー作成完了 ===');
  console.log('名称: ' + name);
  _mSwap('CALENDAR_ID', newCal.getId());
  console.log('※過去の予定は自動では移りません。必要なら旧カレンダーから ICS で書き出して取り込んでください。');
  console.log('※今後ボットが作る予定はこの新カレンダーに入ります。');
}

// ------------------------------------------
// 4'. 新規作成せず admin のメインカレンダーを使う場合はこちら
//     （migrateCreateCalendar() の代わりに実行）
// ------------------------------------------
function migrateUseAdminPrimaryCalendar() {
  _mAssertAdmin();
  var calId = MIGRATE_TARGET_EMAIL; // admin@woodbasef.com = admin のメインカレンダー
  var cal = CalendarApp.getCalendarById(calId);
  if (!cal) {
    console.log('⚠️ ' + calId + ' のカレンダーにアクセスできません。');
    console.log('   admin@woodbasef.com で実行しているか確認してください。');
    return;
  }
  console.log('=== admin メインカレンダーを使用 ===');
  console.log('名称: ' + cal.getName());
  _mSwap('CALENDAR_ID', calId);
  console.log('※今後ボットが作る予定は admin のメインカレンダーに入ります。');
}

// ------------------------------------------
// ロールバック（旧IDに戻す）
// ------------------------------------------
function migrateRollback() {
  var keys = ['SPREADSHEET_ID', 'DRIVE_FOLDER_ID', 'CALENDAR_ID'];
  var props = _mProps();
  console.log('=== ロールバック ===');
  keys.forEach(function (key) {
    var old = props.getProperty('OLD_' + key);
    if (old) {
      props.setProperty(key, old);
      props.deleteProperty('OLD_' + key);
      console.log('  ' + key + ' を旧IDに戻しました: ' + old);
    } else {
      console.log('  ' + key + ': 退避された旧IDなし（変更されていません）');
    }
  });
  console.log('完了。コピーで作成したファイル/カレンダーは手動で削除してください。');
}

function _mStamp() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm');
}
