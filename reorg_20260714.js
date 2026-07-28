// ==========================================
// 2026-07-14 ワンショット: 社内グループシートへ既存グループIDを反映
//   IDはソースに書かず、プロジェクト管理(辞書)から正式名称で引き当てる。
//     社内     ← 'WBグループ内部'
//     経理     ← 'WOODBASE社内経理'
//     日次通知 ← getDailyNotifyGroupId()（現行の善波グループ解決値）
//   通知先の解決は Script Property が最優先のため、実行後に
//   プロパティとの食い違いがあればログへ警告を出す（プロパティは変更しない）。
//   再実行しても安全（同じ値の上書き＋ログのみ）。
// ==========================================
function pmApplyInternalGroups20260714() {
  // シートが無ければ作成し、役割行を用意（非破壊）
  setupSheet(getSS(), PM_SHEET_INTERNAL_GROUPS, PM_INTERNAL_GROUPS_HEADERS, '#B4A7D6', null);
  pmSeedInternalGroupRows();

  var plan = [
    { role: '社内',     name: 'WBグループ内部',   gid: pmDictGroupIdByName_('WBグループ内部') },
    { role: '経理',     name: 'WOODBASE社内経理', gid: pmDictGroupIdByName_('WOODBASE社内経理') },
    { role: '日次通知', name: '善波グループ',     gid: getDailyNotifyGroupId() },
  ];

  var rows = pmReadObjects(PM_SHEET_INTERNAL_GROUPS);
  plan.forEach(function(p) {
    if (!p.gid) { console.warn('⚠ ' + p.role + ': ' + p.name + ' のグループIDが辞書に見つからずスキップ'); return; }
    var row = null;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i]['役割'] || '').trim() === p.role) { row = rows[i]; break; }
    }
    if (!row) { console.warn('⚠ 役割行が見つかりません: ' + p.role); return; }
    var before = String(row['グループID'] || '').trim();
    pmWriteRowFields(PM_SHEET_INTERNAL_GROUPS, row._row, {
      'グループID': p.gid,
      '表示名': p.name,
      '更新日時': fmtDT(new Date()),
    });
    console.log('✅ ' + p.role + ' ← ' + p.name + ' (' + p.gid + ')' +
      (before && before !== p.gid ? ' ※旧値 ' + before + ' を上書き' : ''));
  });

  // Script Property が優先されるため、食い違いを警告（プロパティ自体は触らない）
  var props = PropertiesService.getScriptProperties();
  var internalProp = props.getProperty('INTERNAL_GROUP_ID') || '';
  var sheetInternal = pmRoleGroupId('社内');
  if (internalProp && sheetInternal && internalProp !== sheetInternal) {
    console.warn('⚠ Script Property INTERNAL_GROUP_ID（' + internalProp + '）とシートの社内（' + sheetInternal +
      '）が不一致。現状はプロパティが優先されます。シート側を正としたい場合はプロパティを削除してください。');
  }
  var accProp = props.getProperty('ACCOUNTING_GROUP_ID') || '';
  if (accProp && accProp !== pmRoleGroupId('経理')) {
    console.warn('⚠ Script Property ACCOUNTING_GROUP_ID（' + accProp + '）がシートの経理と不一致。プロパティが優先されます。');
  }
  console.log('完了。pmDiagnose() で解決状況を確認できます。');
}

// プロジェクト管理(辞書)の正式名称 → グループID（施主/業者のうち入っている方）
function pmDictGroupIdByName_(formalName) {
  var sheet = getSheet('プロジェクト管理');
  if (!sheet || sheet.getLastRow() < 2) return '';
  var data = sheet.getDataRange().getValues();
  var h = data[0].map(function(x) { return String(x).trim(); });
  var cFormal = h.indexOf('正式名称');
  var cNushi  = h.indexOf('グループID（施主）');
  var cGyo    = h.indexOf('グループID（業者）');
  if (cFormal === -1) return '';
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][cFormal] || '').trim() === String(formalName).trim()) {
      var gid = (cGyo !== -1 ? String(data[i][cGyo] || '').trim() : '') ||
                (cNushi !== -1 ? String(data[i][cNushi] || '').trim() : '');
      if (gid) return gid;
    }
  }
  return '';
}
