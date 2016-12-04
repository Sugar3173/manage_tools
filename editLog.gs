/***********************************************************************
* @file editLog.gs
*
* @brief function of edit Logger
*
* @author Shohei.Sugano
*
* @date 2016.12.04
*
* @copyright (c) 2016 aba Co.,Ltd.
*
***********************************************************************/

/**
 * edit Logger
 * 
 * @param sheet[in] events {@see https://developers.google.com/apps-script/understanding_events?hl=ja}
 * @attention need setting trigger
 */
function editLog(e) {
    //Spreadsheet Servces 
    //http://googlestyle.client.jp/sheet/sheet.html
/*	
    //Log保存用シートの名前
	var logSheetName = 'Log';

	// スプレッドシート
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	// スプレッドシート名
	var ssName = ss.getName();

	// 選択シート
	var sheet = ss.getActiveSheet();
	// 選択シート名
	var sheetName = sheet.getName();
	// Logシートなら何もしない
	if (sheetName == logSheetName) {
		return;
	}

	// 選択セル範囲
	var range = sheet.getActiveRange();
	// セル範囲の行番号
	var rowIndex = range.getRowIndex();
	// セル範囲の列番号
	var colIndex = range.getColumnIndex();

	// getRange(始点行, 始点列, 取得する行数, 取得する列数)
	var v = sheet.getRange(rowIndex, colIndex, 1, 1).getValue();
	//内容が空だ
	if (v == '') {
		v = '※削除かな？';
	}

	//更新者情報は法人向けGoogle Appsの同一ドメインでないと取得できないかも？ふつうの @gmail.com だと無理かも？
	//https://productforums.google.com/forum/#!topic/docs/5D23Os_NIAc

	//更新者のメールアドレス
	var email = Session.getActiveUser().getEmail();

	//ここからLogシートに書き込み
	//Log保存用シート
	var logSheet = ss.getSheetByName(logSheetName);
	//引数で指定した行の前の行に1行追加
    var row = 2;
    logSheet.insertRowBefore(row);

	//日付
	logSheet.getRange(row, 1).setNumberFormat('yyyy/mm/dd(ddd)');
	logSheet.getRange(row, 1).setValue(new Date());
	//時刻
	logSheet.getRange(row, 2).setNumberFormat('h:mm:ss');
	logSheet.getRange(row, 2).setValue(new Date());
	//更新者
	logSheet.getRange(row, 3).setValue(email);
	//シート名
	logSheet.getRange(row, 4).setValue(sheetName);
	//行番号
	logSheet.getRange(row, 5).setValue(rowIndex);
	//列番号　#TODO アルファベットに変換するメソッドを作成できないか
	logSheet.getRange(row, 6).setValue(colIndex);
	//変更セルの内容(Stringフォーマットにする)
	logSheet.getRange(row, 7).setNumberFormat('@');  
	logSheet.getRange(row, 7).setValue(v);  
*/
/*
	//Slackに通知する場合
	//tokenを取得(ボタン押す)→ https://api.slack.com/web
	//channels:id取得(ボタン押す)→ https://api.slack.com/methods/channels.list/test
	//先に更新通知用チャンネルを作っておいた方が良いかも？

	var token = 'ここにtoken';
	var channel = 'ここにchannelId';
	var userName = 'Spreadsheets'; // Slack投稿時に使われる好きな名前

	//Slackに通知するテキスト
	var text = '「'+sheetName + '」シート変更: ' + email + ' [' + rowIndex + ',' + colIndex + '] ' + v;

	UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
		method: 'post',
		payload: {
			token: token,
			channel: channel,
			username: userName,
			text: text
		}
	});
	//実行トランスクリプトに「実行に失敗: fetch を呼び出す権限がありません」
	//と出た場合、メニューから実行して承認してやる必要あり？
	//こちらもやはりGoogle Apps for work でないとだめかも？
	//https://code.google.com/p/google-apps-script-issues/issues/detail?id=677
*/
}

