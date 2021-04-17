import { LoadCalendarList } from './usecase/LoadCalendarList';
import { InsertRoutineSchedules } from './usecase/InsertRoutineSchedules';
import { DeleteOwnedSchedules } from './usecase/DeleteOwnedSchedules';

// -----------------------------------------------------------------------------
// イベント
// -----------------------------------------------------------------------------

/**
 * スプレッドシートを開いたときに実行するイベントハンドラ
 */
function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('<<カスタム>>')
    .addItem('①カレンダーリストの読み込み', 'loadCalendarList')
    .addItem('②スケジュール登録', 'insertRoutineSchedules')
    .addSeparator()
    .addItem('※注意※スケジュール削除', 'deleteOwnedSchedules')
    .addToUi();
}

// -----------------------------------------------------------------------------
// メソッド
// -----------------------------------------------------------------------------
/**
 * カレンダーリストの読み込み
 */
function loadCalendarList(): void {
  const loadCalendarList = new LoadCalendarList();
  loadCalendarList.execute();
}

/**
 *  指定日の予定を全て削除する
 *    ※自分が所有しているカレンダーのみ対象
 */
function deleteOwnedSchedules(): void {
  const deleteOwnedSchedules = new DeleteOwnedSchedules();
  deleteOwnedSchedules.excute();
}

/**
 * 表示中のシートから予定登録
 */
function insertRoutineSchedules(): void {
  const insertRoutineSchedules = new InsertRoutineSchedules();
  insertRoutineSchedules.excute();
}

/***
 * WebAPI Getメソッド
 */
function doGet(e): GoogleAppsScript.Content.TextOutput {
  const dateAdd: string = e.parameter.add;
  let targetDate = new Date();
  if (dateAdd) {
    targetDate.setDate(targetDate.getDate() + Number(dateAdd));
  }

  // アクディブシートから予定登録
  const insertRoutineSchedules = new InsertRoutineSchedules(targetDate);
  insertRoutineSchedules.excute();

  // 結果を返す
  const msg = `「${targetDate.getFullYear()}/${
    targetDate.getMonth() + 1
  }/${targetDate.getDate()}」の予定を登録しました。`;
  console.log(msg);

  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.TEXT);
  out.setContent(JSON.stringify(msg));

  return out;
}

/***
 * Getメソッドのデバッグ用
 */
function debugDoGet() {
  const e = {
    parameter: {
      add: '4',
    },
  };
  doGet(e);
}
