// ==========================================
// ProjectCalendar：projects の日付項目を Googleカレンダーへ同期
//   案件ID×イベント種別 でイベントIDを保持し、重複作成せず更新する（冪等）。
//   依存: pm_core.js, pm_manager.js, 既存 getConfig
// ==========================================

// イベント種別 → projects の日付列・表示ラベル
var PM_CAL_TYPES = [
  { type: 'next_action',        col: '次回アクション期限', label: '次回アクション' },
  { type: 'meeting',            col: '打合せ日',           label: '打合せ' },
  { type: 'construction_start', col: '着工予定日',         label: '着工' },
  { type: 'handover',           col: '引き渡し予定日',     label: '引き渡し' },
  { type: 'billing_due',        col: '請求予定日',         label: '請求予定' },
  { type: 'payment_due',        col: '入金予定日',         label: '入金予定' },
];

function pmGetCalEventRow(projectId, type) {
  var rows = pmReadObjects(PM_SHEET_CALENDAR);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['案件ID']) === String(projectId) && String(rows[i]['イベント種別']) === type) return rows[i];
  }
  return null;
}

// 追跡行を削除（案件ID×種別）。日付削除時のクリーンアップ用。
function pmDeleteCalEventRow(projectId, type) {
  var sheet = pmSheet(PM_SHEET_CALENDAR);
  if (!sheet || sheet.getLastRow() <= 1) return;
  var idx = pmHeaderIndex(sheet);
  var cId = idx['案件ID'], cType = idx['イベント種別'];
  if (cId === undefined || cType === undefined) return;
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][cId]) === String(projectId) && String(data[i][cType]) === type) sheet.deleteRow(i + 1);
  }
}

function pmSyncCalendar(projectId) {
  var config = getConfig();
  if (!config.CALENDAR_ID) { console.log('pmSyncCalendar: CALENDAR_ID未設定のためスキップ'); return; }
  var calendar = CalendarApp.getCalendarById(config.CALENDAR_ID);
  if (!calendar) { console.error('pmSyncCalendar: カレンダーが見つかりません'); return; }

  // 同時実行で同じ案件×種別のイベントが二重作成されないよう直列化
  var res = pmWithLock(function() { pmSyncCalendarInner(calendar, projectId); return true; });
  if (!res.ok) console.warn('pmSyncCalendar: ロック取得できずスキップ（次回更新で再同期されます）');
}

function pmSyncCalendarInner(calendar, projectId) {
  var proj = pmGetProjectById(projectId);
  if (!proj) return;
  var projName = proj['案件名'] || projectId;
  var now = fmtDT(new Date());

  PM_CAL_TYPES.forEach(function(def) {
    var dateVal = proj[def.col];
    // 日付が削除された → 既存のカレンダー予定を削除し、追跡行も除去する
    if (pmIsBlank(dateVal)) {
      var stale = pmGetCalEventRow(projectId, def.type);
      if (stale && stale['イベントID']) {
        try { var sev = calendar.getEventById(stale['イベントID']); if (sev) sev.deleteEvent(); } catch (e) { console.warn('pmSyncCalendar 削除失敗(' + def.type + '):', e.message); }
        pmDeleteCalEventRow(projectId, def.type);
        console.log('pmSyncCalendar: 日付削除に伴いカレンダー予定を削除 (' + def.type + ')');
      }
      return;
    }
    var date = pmParseDate(dateVal);
    if (!date) return;
    var ymd = fmtDate(date);
    var title = '【' + projName + '】' + def.label;

    var existing = pmGetCalEventRow(projectId, def.type);
    try {
      if (existing && existing['イベントID']) {
        // 日付が同じなら何もしない（保存値がDate/文字列いずれでも正規化して比較）
        var prevDate = pmParseDate(existing['開始日時']);
        var prevYmd = prevDate ? fmtDate(prevDate) : '';
        if (prevYmd === ymd) return;
        // 日付変更 → 既存イベントを更新（取得できなければ作り直し）
        var ev = null;
        try { ev = calendar.getEventById(existing['イベントID']); } catch (e) { ev = null; }
        if (ev) {
          ev.setAllDayDate(date);
          ev.setTitle(title);
          pmWriteRowFields(PM_SHEET_CALENDAR, existing._row, { '開始日時': ymd, '更新日': now });
          return;
        }
        // イベントが消えていた → 新規作成して行を更新
        var newEv2 = calendar.createAllDayEvent(title, date, { description: 'WOODBASE LINE秘書AI 自動登録（案件進捗）' });
        pmWriteRowFields(PM_SHEET_CALENDAR, existing._row, { 'イベントID': newEv2.getId(), '開始日時': ymd, '更新日': now });
        return;
      }
      // 新規作成
      var newEv = calendar.createAllDayEvent(title, date, { description: 'WOODBASE LINE秘書AI 自動登録（案件進捗）' });
      pmAppendRowFields(PM_SHEET_CALENDAR, {
        '案件ID': projectId,
        'イベント種別': def.type,
        'イベントID': newEv.getId(),
        '開始日時': ymd,
        '終了日時': ymd,
        '担当者': proj['担当者'] || '',
        '作成日': now,
        '更新日': now,
      });
    } catch (err) {
      console.error('pmSyncCalendar error (' + def.type + '):', err.message);
    }
  });
}
