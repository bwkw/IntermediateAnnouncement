const SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('日程');
const CALENDAR = CalendarApp.getCalendarById('*******');
const CHECKBOX_COLUMN = 2;


/**
 * スプレッドシートの特定行から登録する中間発表の情報を取得
 * 
 * @param {int} row - 行
 * @return {string} title, {date} startDateTime, {date} endDateTime, {string} host, {string} sub_host
 */
function fetchSpreadSheetRegisterIntermediateAnnouncement(row) {
  var admission_period = SHEET.getRange(row, 3).getValue();
  var admission_count = SHEET.getRange(row, 4).getValue();
  var title = admission_period + ' ' + admission_count + '回目 ' + '中間発表';
  var date = SHEET.getRange(row, 5).getValue();
  date = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
  var startTime = SHEET.getRange(row, 6).getValue();
  startTime = Utilities.formatDate(startTime, 'Asia/Tokyo', 'HH:mm');
  var endTime = SHEET.getRange(row, 7).getValue();
  endTime = Utilities.formatDate(endTime, 'Asia/Tokyo', 'HH:mm');
  var startDateTime = date + ' ' + startTime;
  var endDateTime = date + ' ' + endTime;
  var host = SHEET.getRange(row, 8).getValue();
  var sub_host = SHEET.getRange(row, 9).getValue();
  var zoom_url = SHEET.getRange(row, 10).getValue();
  return {title, startDateTime, endDateTime, host, sub_host, zoom_url};
}

/**
 * スプレッドシートの特定行から削除する中間発表の情報を取得
 * 
 * @param {int} row - 行
 * @return {date} startDateTime, {date} endDateTime
 */
function fetchSpreadSheetDeleteIntermediateAnnouncement(row) {
  var date = SHEET.getRange(row, 5).getValue();
  date = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
  var startTime = SHEET.getRange(row, 6).getValue();
  startTime = Utilities.formatDate(startTime, 'Asia/Tokyo', 'HH:mm');
  var endTime = SHEET.getRange(row, 7).getValue();
  endTime = Utilities.formatDate(endTime, 'Asia/Tokyo', 'HH:mm');
  var startDateTime = date + ' ' + startTime;
  var endDateTime = date + ' ' + endTime;
  return {startDateTime, endDateTime};
}

/**
 * Googleカレンダーにイベントを登録する
 * 
 * @param {string} title - イベント名
 * @param {date} startDateTime - 開始時間
 * @param {date} endDateTime - 終了時間
 * @param {string} host - 司会
 * @param {string} sub_host - サブ司会
 * @param {string} zoom_url - zoomのurl
 */
function registerCalendarEvent(title, startDateTime, endDateTime, host, sub_host, zoom_url) {
  var startDateTime = new Date(startDateTime);
  var endDateTime = new Date(endDateTime);
  var option = {
    description: zoom_url + '\n\n司会：' + host + '\nサブ司会：' + sub_host
  }
  CALENDAR.createEvent(title, startDateTime, endDateTime, option);
}

/**
 * Googleカレンダーからイベントを削除する
 * 
 * @param {date} startDateTime - 開始時間
 * @param {date} endDateTime - 終了時間
 */
function deleteCalendarEvent(startDateTime, endDateTime) {
  var startDateTime = new Date(startDateTime);
  var endDateTime = new Date(endDateTime);
  var events = CALENDAR.getEvents(startDateTime, endDateTime);
  events[0].deleteEvent();
}

/**
 * mainの登録処理
 * 
 * @param {int} row - 行
 */
function main_register(row) {
  var {title, startDateTime, endDateTime, host, sub_host, zoom_url} = fetchSpreadSheetRegisterIntermediateAnnouncement(row);
  registerCalendarEvent(title, startDateTime, endDateTime, host, sub_host, zoom_url);
}

/**
 * mainの削除処理
 * 
 * @param {int} row - 行
 */
function main_delete(row) {
  var {startDateTime, endDateTime} = fetchSpreadSheetDeleteIntermediateAnnouncement(row);
  deleteCalendarEvent(startDateTime, endDateTime);
}

/**
 * スプレッドシートのチェックボックスにチェックが付けられたら、main_register処理実行
 * スプレッドシートのチェックボックスからチェックが外されたら、main_delete処理実行
 * 
 */
function changeSpreadSheetIntermediateAnnouncementEvent() {
  var activeCell = SHEET.getActiveCell();
  var activeCellColumn = activeCell.getColumn();
  var activeCellRow = activeCell.getRow();
  var activeCellValue = activeCell.getValue();
  if (activeCellColumn == CHECKBOX_COLUMN) {
    if (activeCellValue == true) {
      main_register(activeCellRow);
    } else {
      main_delete(activeCellRow);
    }
  }
}
