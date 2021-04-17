import { LoadCalendarList } from './usecase/LoadCalendarList';
import { InsertRoutineSchedules } from './usecase/InsertRoutineSchedules';
import { DeleteOwnedSchedules } from './usecase/DeleteOwnedSchedules';

// -----------------------------------------------------------------------------
// イベント
// -----------------------------------------------------------------------------

/**
 * スプレッドシートを開いたときに実行するイベントハンドラ
 */
function onOpen (): void {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('<<カスタム>>')
    .addItem('①カレンダーリストの読み込み', 'loadCalendarList')
    .addItem('②スケジュール登録', 'insertRoutineSchedules')
    .addSeparator()
    .addItem('※注意※スケジュール削除', 'deleteOwnedSchedules')
    .addToUi();
};

// -----------------------------------------------------------------------------
// メソッド
// -----------------------------------------------------------------------------
/**
 * カレンダーリストの読み込み
 */
function loadCalendarList(): void {
  const loadCalendarList = new LoadCalendarList();
  loadCalendarList.execute();
};

/**
 *  指定日の予定を全て削除する
 *    ※自分が所有しているカレンダーのみ対象
 */
function deleteOwnedSchedules(): void {
  const deleteOwnedSchedules = new DeleteOwnedSchedules();
  deleteOwnedSchedules.excute();
};

/**
 * 表示中のシートから予定登録
 */
function insertRoutineSchedules (): void {
  const insertRoutineSchedules = new InsertRoutineSchedules();
  insertRoutineSchedules.excute();
};
