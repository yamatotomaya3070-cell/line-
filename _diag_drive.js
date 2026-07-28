// ============================================================
// 一時診断ツール（読み取り専用）
//   目的: 新Drive親フォルダ / 【原本】雛形 の実構成とアクセス権を確認する。
//   - Drive・スプレッドシートを一切変更しない（getName/getFolders等の読取のみ）。
//   - Phase 1 着手前の事前確認用。移行dry-run実装後に削除可。
//   実行: GASエディタで diagDriveStructure() を実行し、ログ(Ctrl+Enter)を確認。
// ============================================================

var DIAG_NEW_ROOT_ID = '14QoHs_nlMwN7IeY-SGgQ9HX51lnsKmZl';       // 新しい大もと親フォルダ
var DIAG_TEMPLATE_ID = '1eM6BVsC2T496Ve03vVa4O6gNSKS1dR-X';       // 【原本】雛形

function diagDriveStructure() {
  var lines = [];
  function out(s) { lines.push(s); }

  // --- 実行ユーザーと現行設定 ---
  var effEmail = '(取得不可)';
  try { effEmail = Session.getEffectiveUser().getEmail() || '(空)'; } catch (e) { effEmail = 'ERR:' + e.message; }
  out('■ 実行ユーザー(effective): ' + effEmail);
  try {
    var curRoot = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
    out('■ 現在の DRIVE_FOLDER_ID (Script Property): ' + (curRoot || '(未設定)'));
    out('■ 新親フォルダ(希望): ' + DIAG_NEW_ROOT_ID + ' / 一致=' + (curRoot === DIAG_NEW_ROOT_ID));
  } catch (e) { out('■ Script Property読取ERR: ' + e.message); }

  // --- 新親フォルダ ---
  out('\n===== 新親フォルダ ' + DIAG_NEW_ROOT_ID + ' =====');
  out(diagDescribeFolder_(DIAG_NEW_ROOT_ID, effEmail));
  out('-- 直下フォルダ一覧（既存案件の有無確認）--');
  out(diagListChildren_(DIAG_NEW_ROOT_ID, 30));

  // --- 【原本】雛形 ---
  out('\n===== 【原本】雛形 ' + DIAG_TEMPLATE_ID + ' =====');
  out(diagDescribeFolder_(DIAG_TEMPLATE_ID, effEmail));
  out('-- 雛型ツリー（最大3階層）--');
  out(diagTree_(DIAG_TEMPLATE_ID, 3, ''));

  var report = lines.join('\n');
  console.log(report);
  return report;
}

// 権限を確認したいアカウント（運用=admin。個人=woodbasegroup も参考表示）
var DIAG_CHECK_EMAILS = ['admin@woodbasef.com', 'woodbasegroup@gmail.com', 'woodbase.hamada@gmail.com'];

// フォルダの基本情報とアクセス権（読取のみ）
function diagDescribeFolder_(id, effEmail) {
  try {
    var f = DriveApp.getFolderById(id);
    // 各アカウントの個別アクセス権（EDIT/OWNER なら書込可）
    var perEmail = DIAG_CHECK_EMAILS.map(function (email) {
      var a;
      try { a = f.getAccess(email); a = a ? String(a) : 'NONE'; }
      catch (e2) { a = 'ERR:' + e2.message; }
      return '    - ' + email + ' : ' + a;
    }).join('\n');
    // リンク共有の状態（anyone系がEDITなら誰でも書込可）
    var link = '(取得不可)';
    try { link = String(f.getSharingAccess()) + ' / ' + String(f.getSharingPermission()); }
    catch (eL) { link = 'ERR:' + eL.message; }
    var editors = '(取得不可)';
    try { editors = f.getEditors().map(function (u) { return u.getEmail(); }).join(', ') || '(なし)'; }
    catch (e3) { editors = 'getEditorsERR:' + e3.message + '（共有ドライブの可能性）'; }
    return [
      '  name : ' + f.getName(),
      '  url  : ' + f.getUrl(),
      '  アカウント別アクセス:',
      perEmail,
      '  リンク共有(Access/Permission) : ' + link,
      '  Editors : ' + editors,
    ].join('\n');
  } catch (e) {
    return '  ★アクセス不可/存在しない: ' + e.message + '\n  → admin@woodbasef.com へ「編集者」共有が必要な可能性';
  }
}

function diagListChildren_(id, limit) {
  try {
    var f = DriveApp.getFolderById(id);
    var it = f.getFolders();
    var names = [];
    while (it.hasNext() && names.length < limit) names.push('  └ ' + it.next().getName());
    if (it.hasNext()) names.push('  … (他多数)');
    return names.length ? names.join('\n') : '  (子フォルダなし)';
  } catch (e) { return '  一覧取得ERR: ' + e.message; }
}

// 再帰ツリー（フォルダ名のみ・読取）
function diagTree_(id, depth, indent) {
  try {
    var f = DriveApp.getFolderById(id);
    var lines = [indent + '📁 ' + f.getName()];
    if (depth <= 0) return lines.join('\n');
    var it = f.getFolders();
    var count = 0;
    while (it.hasNext() && count < 50) {
      var child = it.next();
      lines.push(diagTree_(child.getId(), depth - 1, indent + '  '));
      count++;
    }
    // ファイル数だけ数える（名前は出さない）
    try {
      var fit = f.getFiles(); var fc = 0;
      while (fit.hasNext() && fc < 1000) { fit.next(); fc++; }
      if (fc > 0) lines.push(indent + '  (ファイル ' + fc + '件)');
    } catch (eF) {}
    return lines.join('\n');
  } catch (e) { return indent + '★ERR: ' + e.message; }
}
