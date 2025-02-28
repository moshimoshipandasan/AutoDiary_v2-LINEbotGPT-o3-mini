// 環境変数(スクリプトのプロパティ)へのAPI設定
const apikey = PropertiesService.getScriptProperties().getProperty('apikey');
const linetoken = PropertiesService.getScriptProperties().getProperty('linetoken');

// 使用するAPIの定義
const LINE_REPLY_URL = 'https://api.line.me/v2/bot/message/reply';
const CHAT_GPT_URL = "https://api.openai.com/v1/chat/completions";
const CHAT_GPT_VER = "o3-mini";
// const CHAT_GPT_TEMPRATURE = 0.3;//会話の創造性を向上させる(0~1)

// スプレッドシートの情報
const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEET = SS.getSheetByName('プロンプト');
const SHEET_CHECK = SS.getSheetByName('検証');
const SHEET_LOG = SS.getSheetByName('ログ');
const SHEET_USER = SS.getSheetByName('ユーザー');
const SHEET_INFO = SS.getSheetByName('情報');

// 遡るLINEの過去の履歴数
const MAX_COUNT_LOG = 10;

// セリフ
const firstWord = "初めまして、よろしくお願いします";
const nonText =  "ごめんなさい。メッセージだけ送ってね";

//メニュー
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('アクティブセルlineUserIdにメッセージを送る', 'sendMessageToActiveCellUser')
    .addToUi();
}

// LINEからのメッセージ受信
function doPost(e) {
  const data = JSON.parse(e.postData.contents).events[0]; // LINEからのイベント情報を取得
  if (!data) {
    return SHEET_CHECK.appendRow(["Webhook設定接続確認しました。"]);
  }
  const replyToken = data.replyToken; // リプライトークン
  const lineUserId = data.source.userId; // LINE ユーザーIDを取得
  const userSheetName = findUserSheetName(lineUserId); // ユーザー個別のシート
  const dataType = data.type; // データのタイプ(メッセージ、フォロー、アンフォローなど)
  // ユーザーがフォローした際(友達追加時)の処理
  if (dataType === "follow") {
    const UserData = findUser(lineUserId);
    if (typeof UserData === "undefined") {
      addUser(lineUserId); // ユーザーシートにユーザーを追加
    }
    // ユーザーに歓迎のメッセージを送信(必要に応じてカスタマイズ)
    sendMessage(replyToken, firstWord);
    return; // ここで処理を終了
  }
  // メッセージ受信時の処理
  if (dataType === "message") {
    const messageData = data.message;
    const messageType = messageData.type; // メッセージのタイプ(テキスト、画像、スタンプなど)
    // テキストメッセージの場合の処理
    if (messageType === "text") {
      const postMessage = messageData.text; // ユーザーから受け取ったテキストメッセージ
      // データ生成&LINEに送信
      const totalMessages = chatGPTLog(userSheetName, postMessage);
      const replyText = chatGPT(lineUserId, totalMessages); // ChatGPT関数を使って応答テキストを生成
      sendMessage(replyToken, replyText); // LINEに応答を送信
      // ログに追加
      debugLog(lineUserId, postMessage, replyText); // 大元のログシートに追加
      debugLogIndividual(userSheetName, postMessage, replyText); // 個別ユーザーシートのログに追加
    } else {
      // テキスト以外のメッセージの場合、サポートしていないメッセージタイプである旨をユーザーに通知
      sendMessage(replyToken, nonText);
    }
  }
  // その他のdataTypeに対する処理があればここに追加
}

// LINEに返答
function sendMessage(replyToken, replyText) {
  const postData = {
    "replyToken" : replyToken,
    "messages" : [
      {
        "type" : "text",
        "text" : replyText
      }
    ]
  };  
  const headers = {
    "Content-Type" : "application/json; charset=UTF-8",
    "Authorization" : `Bearer ${linetoken}`
  };
  const options = {
    "method" : "POST",
    "headers" : headers,
    "payload" : JSON.stringify(postData)
  };
  return UrlFetchApp.fetch(LINE_REPLY_URL, options);
}

// ChatGPTのレスポンスを確認する関数
function checkChatGPTResponse(responseText) {
  // レスポンスの詳細をログ出力
  console.log("ChatGPTレスポンス受信:", responseText);

  // responseTextが未定義または空の場合のチェック
  if (!responseText) {
    console.log("ChatGPTレスポンスエラー: レスポンスが空です");
    return "申し訳ありません。応答を受信できませんでした。";
  }

  try {
    // JSONパースを試みる
    const json = JSON.parse(responseText);
    console.log("パースされたJSON:", JSON.stringify(json, null, 2));

    // レスポンスの構造を確認
    if (!json.choices) {
      console.log("ChatGPTレスポンスエラー: choicesが存在しません");
      return responseText;
    }

    if (!json.choices[0]) {
      console.log("ChatGPTレスポンスエラー: choices[0]が存在しません");
      return responseText;
    }

    // choices[0]の内容をログ出力
    console.log("choices[0]の内容:", json.choices[0]);

    if (json.choices[0].message && json.choices[0].message.content) {
      // 正常なJSONレスポンスの場合
      const content = json.choices[0].message.content.trim();
      console.log("ChatGPTレスポンス正常: ", content);
      return content;
    } else if (typeof json.choices[0] === 'string') {
      // choicesが直接文字列の場合
      console.log("ChatGPTレスポンス (文字列): ", json.choices[0]);
      return json.choices[0];
    } else if (json.choices[0].text) {
      // 古い形式のレスポンス
      console.log("ChatGPTレスポンス (text): ", json.choices[0].text);
      return json.choices[0].text;
    }

    console.log("ChatGPTレスポンスエラー: 不明なレスポンス形式");
    return responseText;
  } catch (error) {
    // JSONパースエラーの場合
    console.log("ChatGPTレスポンスエラー: JSONパースエラー");
    console.log("エラー詳細:", error);
    console.log("受信したレスポンス:", responseText);
    
    // テキストとして返却
    return responseText;
  }
}

// ChatGPTへの指示
function chatGPT(userId, totalMessages) {
  const constraints = SHEET.getRange(1, 1).getValue(); // スプシからプロンプト
  const userName = getUserDisplayName(userId); // ユーザーのプロフィール名を取得
  const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'); // 現在の日時を取得
  const info = getInformationSheetDataAsCSV();
  totalMessages.unshift({
    "role": "system",
    "content": `${constraints}\n
# ユーザーの名前:${userName}\n
# 現在時刻:${date}\n
# 過去の日記:${info}
`
  });  //プロンプト
  const requestOptions = {
    "method": "post",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer "+ apikey
    },
    "payload": JSON.stringify({
      "model": CHAT_GPT_VER,
      "messages": totalMessages,
      "reasoning_effort": "medium"
      // "temperature": CHAT_GPT_TEMPRATURE
    })
  }
  try {
    const response = UrlFetchApp.fetch(CHAT_GPT_URL, requestOptions);
    const responseText = response.getContentText();
    return checkChatGPTResponse(responseText);
  } catch (error) {
    console.log("ChatGPT API呼び出しエラー:", error);
    return "申し訳ありません。ChatGPTとの通信中にエラーが発生しました。";
  }
}

// ログの記録処理
//「ログ」に追加
function debugLog(userId, text, replyText) {
  const UserData = findUser(userId); // ユーザーシートにデータがあるか確認
  typeof UserData === "undefined" ? addUser(userId) : userUseChat(userId); // ユーザーシートにデータがなければユーザー追加、あれば投稿数だけ追加
  const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'); // 現在の日付を取得
  SHEET_LOG.appendRow([userId, UserData, text, replyText, date]); // ログシートに情報追加
}

// 個別のログシートに追加
function debugLogIndividual(sheetName, text, replyText) {
  const individualSheetLog = SS.getSheetByName(sheetName);
  const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'); // 現在の日付を取得
  individualSheetLog.appendRow([text, replyText, date]);
}

function addUser(userId) {
  const userName = getUserDisplayName(userId); // ユーザーのプロフィール名を取得
  const userIMG  = getUserDisplayIMG(userId); // ユーザーのプロフィール画像URLを取得
  const sheetName = userName + userId.substring(1, 5); // シート名を userName + userIdの2~5文字で設定
  // 新しいシートを追加し、シート名を設定
  if (!SS.getSheetByName(sheetName)) {
    const newUserSheet = SS.insertSheet(sheetName);    
    // ユーザー名の列を含むヘッダー行を追加
    newUserSheet.appendRow(['送信メッセージ', '応答メッセージ', '日付']);
    // 新しいシートの現在の行数と列数を取得する
    const maxRows = newUserSheet.getMaxRows();
    const maxCols = newUserSheet.getMaxColumns();
    // 余分な列を削除する
    if (maxCols > 3) newUserSheet.deleteColumns(4, maxCols - 3);
    // 余分な行を削除する(1行のヘッダーを残して、他のすべての行を削除)
    if (maxRows > 1) newUserSheet.deleteRows(2, maxRows - 1);
  }
  // ユーザー情報シートにユーザー情報を記録
  const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'); // 現在の日付を取得
  SHEET_USER.appendRow([userId, userName, userIMG, 0, sheetName, date, date]);
  return sheetName; // 作成または取得したシート名を返す
}

// ユーザーのシートを更新 
function userUseChat(userId) {
  // 送信したユーザー先のユーザーを検索
  const textFinder = SHEET_USER.createTextFinder(userId);
  const ranges = textFinder.findAll();
  // ユーザーが存在しない場合エラー
  if (!ranges[0])
    SHEET_USER.appendRow([userId, "???", '', 1]);
  // 投稿数プラス1
  const timesFinder = SHEET_USER.createTextFinder('投稿数');
  const timesRanges = timesFinder.findAll();
  const timesRow    = ranges[0].getRow();
  const timesColumn = timesRanges[0].getColumn();
  const times = SHEET_USER.getRange(timesRow, timesColumn).getValue() + 1;
  SHEET_USER.getRange(timesRow, timesColumn).setValue(times);
  // 更新日時を更新
  const updateDateFinder = SHEET_USER.createTextFinder('更新日時');
  const updateDateRanges = updateDateFinder.findAll();
  const updateDateRow    = ranges[0].getRow();
  const updateDateColumn = updateDateRanges[0].getColumn();
  const updateDate = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'); // 現在の日付を取得
  SHEET_USER.getRange(updateDateRow, updateDateColumn).setValue(updateDate);
  // 更新順に並び替え
/**
  const numColumn = SHEET_USER.getLastColumn(); // 最後列の列番号を取得
  const numRow    = SHEET_USER.getLastRow() - 1;  // 最後行の行番号を取得
  const dataRange = SHEET_USER.getRange(2, 1, numRow, numColumn);
  dataRange.sort([{column: updateDateColumn, ascending: false}]); // 更新日付順に並び替え
 */
}

// ユーザーIDに基づいてユーザーのシート名を検索する関数
function findUserSheetName(userId) {
  const userData = SHEET_USER.getDataRange().getValues();
  for (let i = 0; i < userData.length; i++) {
    if (userData[i][0] === userId) {
      return userData[i][4]; // 対応するシート名を返す
    }
  }
  // ユーザーが見つからない場合、シートを作成するかエラーをスロー
  return addUser(userId); // シートが存在しない場合は新しいシートを作成してその名前を返す
}

// 過去の履歴を取得
function chatGPTLog(UserSheetName, postMessage) {
  let individualSheetLog = SS.getSheetByName(UserSheetName);
  // シートが見つからない場合、新しいシートを作成
  if (!individualSheetLog) {
    individualSheetLog = SS.insertSheet(UserSheetName);
    
    // 新しいシートにヘッダーを追加
    individualSheetLog.appendRow(['送信メッセージ', '応答メッセージ', '日付']);
    
    // 新しいシートの現在の行数と列数を取得する
    const maxRows = individualSheetLog.getMaxRows();
    const maxCols = individualSheetLog.getMaxColumns();
    
    // 余分な列を削除する
    if (maxCols > 3) individualSheetLog.deleteColumns(4, maxCols - 3);
    
    // 余分な行を削除する(1行のヘッダーを残して、他のすべての行を削除)
    if (maxRows > 1) individualSheetLog.deleteRows(2, maxRows - 1);
  }
  let values = individualSheetLog.getDataRange().getValues();
  // valuesを逆順にする
  values = values.reverse();
  values.pop();
  // valuesをmapで処理し、MAX_COUNT_LOGの制限を超えないようにメッセージを追加
  const totalMessages = values.slice(0, MAX_COUNT_LOG).flatMap(value => [
    {"role": "assistant", "content": value[1]},
    {"role": "user", "content": value[0]}
  ]);
  // 最新のメッセージを追加
  totalMessages.reverse().push({"role": "user", "content": postMessage});
  // console.log(totalMessages);
  return totalMessages;
}

function chatGPTLogOld(UserSheetName, postMessage) {
  const individualSheetLog = SS.getSheetByName(UserSheetName);
  let values = individualSheetLog.getDataRange().getValues();  
  // valuesを逆順にする
  values = values.reverse();
  let totalMessages = [];
  // valuesを最初から順に処理し、MAX_COUNT_LOGの制限を超えないようにメッセージを追加
  for (let i = 0; i < values.length - 1; i++) {
    if (i >= MAX_COUNT_LOG) break; // 過去の履歴を遡る回数の制限 
    // valuesのインデックスは0から始まるため、ユーザーとアシスタントのメッセージを適切に追加
    totalMessages.unshift({"role": "assistant", "content": values[i][1]});
    totalMessages.unshift({"role": "user", "content": values[i][0]});
  }
  // 最新のメッセージを追加
  totalMessages.push({"role": "user", "content": postMessage});
  // console.log(totalMessages);
  return totalMessages;
}

// メンバーとしてユーザー登録されているか検索
function findUser(uid) {
  return getUserData().reduce(function(uuid, row) { return uuid || (row.key === uid && row.value); }, false) || undefined;
}

// LINEのAPIを使ってユーザーの名前を取得
function getUserDisplayName(userId) {
  const url = 'https://api.line.me/v2/bot/profile/' + userId;
  const userProfile = UrlFetchApp.fetch(url,{
    'headers': {
      'Authorization' : `Bearer ${linetoken}`,
    },
  });
  return JSON.parse(userProfile).displayName;
}

// ユーザーのプロフィール画像取得 
function getUserDisplayIMG(userId) {
  const url = 'https://api.line.me/v2/bot/profile/' + userId
  const userProfile = UrlFetchApp.fetch(url,{
    'headers': {
      'Authorization' : `Bearer ${linetoken}`,
    },
  });
  return JSON.parse(userProfile).pictureUrl;
}

// スプレッドシートを並び替え(対象のシートのカラムを降順に変更)
function dataSort(sortSheet,columnNumber) {
  const numColumn = sortSheet.getLastColumn(); // 最後列の列番号を取得
  const numRow    = sortSheet.getLastRow() - 1;  // 最後行の行番号を取得
  const dataRange = sortSheet.getRange(2, 1, numRow, numColumn);
  dataRange.sort([{column: columnNumber, ascending: false}]); // 降順に並び替え
}

// ユーザー情報取得
function getUserData() {
  const data = SHEET_USER.getDataRange().getValues();
  return data.map(function(row) { return {key: row[0], value: row[1]}; });
}

// 情報の取得
function getInformationSheetDataAsCSV() {
  const data = SHEET_INFO.getDataRange().getValues();
  let csvData = '';
  data.forEach(function(row) {
    csvData += row.join('  ') + '\n';
  });
  return csvData;
}

// 個別メッセージ
// 特定のユーザーにメッセージを送信する
function sendMessageToUser(userId, message) {
  const postData = {
    "to" : userId,
    "messages" : [
      {
        "type" : "text",
        "text" : message
      }
    ]
  };
  
  const headers = {
    "Content-Type" : "application/json; charset=UTF-8",
    "Authorization" : `Bearer ${linetoken}`
  };
  
  const options = {
    "method" : "POST",
    "headers" : headers,
    "payload" : JSON.stringify(postData)
  };
  
  const url = 'https://api.line.me/v2/bot/message/push';
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    Logger.log('Message sent successfully to user: ' + userId);
  } catch (error) {
    Logger.log('Error sending message to user: ' + error);
  }
}

// ダイアログボックスを開いてユーザーにメッセージを入力させる
function promptForMessage() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('メッセージの送信', 'ユーザーに送信するメッセージを入力してください:', ui.ButtonSet.OK_CANCEL);

  // ユーザーがOKを押した場合、入力したメッセージを返す
  if (response.getSelectedButton() == ui.Button.OK) {
    return response.getResponseText();
  } else {
    return null; // キャンセルされた場合はnullを返す
  }
}

// アクティブなセルのuserIdにメッセージを送信する
function sendMessageToActiveCellUser() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeCell = sheet.getActiveCell();
  const userId = activeCell.getValue(); // アクティブセルの値を取得
  
  if (!userId) {
    SpreadsheetApp.getUi().alert('アクティブなセルにユーザーIDがありません。');
    return;
  }
  
  const message = promptForMessage(); // ダイアログボックスを開いてメッセージを取得
  
  if (message) {
    sendMessageToUser(userId, message); // メッセージを送信
  } else {
    SpreadsheetApp.getUi().alert('メッセージの送信がキャンセルされました。');
  }
}
