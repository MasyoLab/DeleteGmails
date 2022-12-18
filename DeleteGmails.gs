const SLEEP_TIME = 1000;
/** コーヒーのスタンプ */
const STAMP_COFFEE = ':coffee:';
/** シートID */
var SHEET_ID = null;
/** シートGID */
var SHEET_GID = null;
/** ウェブフック */
var DISCORD_WEBHOOKS = null;
var APP_DEBUG = false;

/** gasのpost関数 */
function doPost(e) {
  APP_DEBUG = isNull(e);

  var sp = PropertiesService.getScriptProperties();
  SHEET_ID = sp.getProperty('spreadsheets_id');
  SHEET_GID = sp.getProperty('spreadsheets_gid');
  DISCORD_WEBHOOKS = sp.getProperty('discord_webhooks');

  main();
}

/** メインの処理 */
function main() {
  var spreadSheet = SpreadsheetApp.openById(SHEET_ID);
  var sheet = spreadSheet.getSheets().find(v => v.getSheetId() == SHEET_GID);
  if (isNull(sheet)) {
    return;
  }

  var actionArray = new Array();
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  for (var i = 0; i < values.length; i++) {
    actionArray.push(values[i][0]);
  }

  if (APP_DEBUG) {
    Logger.log(SHEET_ID);
    Logger.log(SHEET_GID);
    Logger.log(DISCORD_WEBHOOKS);
    Logger.log(actionArray);
    return;
  }

  startMessage();

  // 検索コマンドを指定する
  for (var i = 0; i < actionArray.length; i++) {
    deleteGmails(actionArray[i]);
  }

  endMessage();
}

/** 指定コマンドで検索したメールを削除する処理 */
function deleteGmails(actionStr) {
  var deleteThreads = GmailApp.search(actionStr);
  if (isNull(deleteThreads) || deleteThreads.length == 0) {
    return;
  }

  var logStr = STAMP_COFFEE + '\n';
  logStr += 'start deleteAction : ' + actionStr + '\n';
  logStr += '該当件数 : ' + deleteThreads.length + '件';
  postDiscordBOT(logStr);

  // 該当メールを全て削除
  for (var i = 0; i < deleteThreads.length; i++) {
    deleteThreads[i].moveToTrash();
  }

  logStr = STAMP_COFFEE + '\n';
  logStr += 'end deleteAction : ' + actionStr + ' done';
  postDiscordBOT(logStr);
}
  
/** Discord API を実行する */
function postDiscordBOT(str) {
  // 投稿するチャット内容と設定
  var message = {
    "content": str, // チャット本文
    "tts": false  // ロボットによる読み上げ機能を無効化
  }
  var param = {
    'method': 'POST',
    'headers': { 'Content-type': "application/json" },
    'payload': JSON.stringify(message)
  }
  
  UrlFetchApp.fetch(DISCORD_WEBHOOKS, param);
  Utilities.sleep(SLEEP_TIME); // 0.3秒間の間に、連続でAPIを実行できないので待機処理
}

function startMessage() {
  var logStr = STAMP_COFFEE + '\n';
  logStr += 'start DeleteGmails';
  postDiscordBOT(logStr);
}

function endMessage() {
  var logStr = STAMP_COFFEE + '\n';
  logStr += 'end DeleteGmails\n';
  logStr += 'all done';
  postDiscordBOT(logStr);
}

/** null */
function isNull(value) {
  return value == null;
}
