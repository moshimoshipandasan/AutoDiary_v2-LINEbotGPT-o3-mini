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

// LINEからのメッセージ受信（非同期処理を導入）
function doPost(e) {
  const data = JSON.parse(e.postData.contents).events[0]; // LINEからのイベント情報を取得
  if (!data) {
    // Webhookの確認
    CacheService.getScriptCache().put('webhook_check', 'true', 600); // 10分間キャッシュ
    SHEET_CHECK.appendRow(["Webhook設定接続確認しました。"]);
    return ContentService.createTextOutput(JSON.stringify({'status': 'ok'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const replyToken = data.replyToken; // リプライトークン
  const lineUserId = data.source.userId; // LINE ユーザーIDを取得
  const dataType = data.type; // データのタイプ(メッセージ、フォロー、アンフォローなど)
  
  try {
    // 即時処理が必要な部分を実行
    if (dataType === "message") {
      const messageData = data.message;
      const messageType = messageData.type;
      
      if (messageType === "text") {
        const postMessage = messageData.text;
        const userSheetName = findUserSheetName(lineUserId);
        const totalMessages = chatGPTLog(userSheetName, postMessage);
        const replyText = chatGPT(lineUserId, totalMessages);
        sendMessage(replyToken, replyText);
        
        // 即時にログを記録（トリガーを使用せず）
        debugLog(lineUserId, postMessage, replyText); // 大元のログシートに追加
        debugLogIndividual(userSheetName, postMessage, replyText); // 個別ユーザーシートのログに追加
      } else {
        // テキスト以外のメッセージの場合
        sendMessage(replyToken, nonText);
      }
    } else if (dataType === "follow") {
      // フォロー時の処理
      const UserData = findUser(lineUserId);
      if (typeof UserData === "undefined") {
        // ユーザー追加処理
        const userSheetName = addUser(lineUserId);
        
        // キャッシュに保存
        const cache = CacheService.getScriptCache();
        cache.put('sheetName_' + lineUserId, userSheetName, 3600);
      }
      // 歓迎メッセージを送信
      sendMessage(replyToken, firstWord);
    }
  } catch (error) {
    console.error('doPost処理エラー: ' + error.message);
    // エラーログを記録
    try {
      SHEET_CHECK.appendRow(["エラー発生", error.message, new Date()]);
    } catch (e) {
      console.error('エラーログ記録失敗: ' + e.message);
    }
  }
  
  // 即座にLINEサーバーに200 OKを返す
  return ContentService.createTextOutput(JSON.stringify({'status': 'ok'}))
    .setMimeType(ContentService.MimeType.JSON);
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
  };
  try {
    const response = UrlFetchApp.fetch(CHAT_GPT_URL, requestOptions);
    const responseText = response.getContentText();
    return checkChatGPTResponse(responseText);
  } catch (error) {
    console.log("ChatGPT API呼び出しエラー:", error);
    return "申し訳ありません。ChatGPTとの通信中にエラーが発生しました。";
  }
}

// ログの記録処理（バッチ処理と非同期処理を導入）
//「ログ」に追加
function debugLog(userId, text, replyText) {
  // ロックを取得して競合を防止
  const lock = LockService.getScriptLock();
  try {
    // ロックを取得（最大5秒待機、10秒間ロック）
    lock.tryLock(5000);
    
    const UserData = findUser(userId); // ユーザーシートにデータがあるか確認
    typeof UserData === "undefined" ? addUser(userId) : userUseChat(userId); // ユーザーシートにデータがなければユーザー追加、あれば投稿数だけ追加
    
    const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'); // 現在の日付を取得
    
    // 一括書き込み用の配列
    const logValues = [[userId, UserData, text, replyText, date]];
    
    // 次の空き行を取得して一括書き込み
    const nextRow = SHEET_LOG.getLastRow() + 1;
    SHEET_LOG.getRange(nextRow, 1, 1, 5).setValues(logValues);
    
    // ログ記録成功をコンソールに出力
    console.log('ログ記録成功: ' + userId);
  } catch (error) {
    console.error('ログ記録エラー: ' + error.message);
    // エラーログを記録
    try {
      SHEET_CHECK.appendRow(["ログ記録エラー", userId, error.message, new Date()]);
    } catch (e) {
      console.error('エラーログ記録失敗: ' + e.message);
    }
  } finally {
    // 必ずロックを解放
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

// 個別のログシートに追加（非同期処理を導入）
function debugLogIndividual(sheetName, text, replyText) {
  // ロックを取得して競合を防止
  const lock = LockService.getScriptLock();
  try {
    // ロックを取得（最大5秒待機、10秒間ロック）
    lock.tryLock(5000);
    
    const individualSheetLog = SS.getSheetByName(sheetName);
    const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'); // 現在の日付を取得
    
    // 一括書き込み用の配列
    const logValues = [[text, replyText, date]];
    
    // 次の空き行を取得して一括書き込み
    const nextRow = individualSheetLog.getLastRow() + 1;
    individualSheetLog.getRange(nextRow, 1, 1, 3).setValues(logValues);
    
    // ログ記録成功をコンソールに出力
    console.log('個別ログ記録成功: ' + sheetName);
  } catch (error) {
    console.error('個別ログ記録エラー: ' + error.message);
    // エラーログを記録
    try {
      SHEET_CHECK.appendRow(["個別ログ記録エラー", sheetName, error.message, new Date()]);
    } catch (e) {
      console.error('エラーログ記録失敗: ' + e.message);
    }
  } finally {
    // 必ずロックを解放
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

// ユーザー追加（非同期処理とロック機能を導入）
function addUser(userId) {
  // ロックを取得して競合を防止
  const lock = LockService.getScriptLock();
  try {
    // ロックを取得（最大10秒待機、30秒間ロック）
    lock.tryLock(10000);
    
    const userName = getUserDisplayName(userId); // ユーザーのプロフィール名を取得
    const userIMG  = getUserDisplayIMG(userId); // ユーザーのプロフィール画像URLを取得
    const sheetName = userName + userId.substring(1, 5); // シート名を userName + userIdの2~5文字で設定
    
    // 新しいシートを追加し、シート名を設定
    if (!SS.getSheetByName(sheetName)) {
      const newUserSheet = SS.insertSheet(sheetName);    
      
      // ヘッダー行とシート設定を一括で行う
      const headerValues = [['送信メッセージ', '応答メッセージ', '日付']];
      newUserSheet.getRange(1, 1, 1, 3).setValues(headerValues);
      
      // 新しいシートの現在の行数と列数を取得する
      const maxRows = newUserSheet.getMaxRows();
      const maxCols = newUserSheet.getMaxColumns();
      
      // 余分な列を削除する
      if (maxCols > 3) newUserSheet.deleteColumns(4, maxCols - 3);
      
      // 余分な行を削除する(1行のヘッダーを残して、他のすべての行を削除)
      if (maxRows > 1) newUserSheet.deleteRows(2, maxRows - 1);
    }
    
    // ユーザー情報シートにユーザー情報を記録（一括書き込み）
    const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'); // 現在の日付を取得
    const userValues = [[userId, userName, userIMG, 0, sheetName, date, date]];
    
    // 次の空き行を取得
    const nextRow = SHEET_USER.getLastRow() + 1;
    SHEET_USER.getRange(nextRow, 1, 1, 7).setValues(userValues);
    
    return sheetName; // 作成または取得したシート名を返す
  } finally {
    // 必ずロックを解放
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

// ユーザーのシートを更新（非同期処理とロック機能を導入）
function userUseChat(userId) {
  // ロックを取得して競合を防止
  const lock = LockService.getScriptLock();
  try {
    // ロックを取得（最大5秒待機、10秒間ロック）
    lock.tryLock(5000);
    
    // 送信したユーザー先のユーザーを検索
    const textFinder = SHEET_USER.createTextFinder(userId);
    const ranges = textFinder.findAll();
    
    // ユーザーが存在しない場合、新規追加
    if (!ranges[0]) {
      const userValues = [[userId, "???", '', 1]];
      const nextRow = SHEET_USER.getLastRow() + 1;
      SHEET_USER.getRange(nextRow, 1, 1, 4).setValues(userValues);
      return;
    }
    
    // 投稿数プラス1
    const timesFinder = SHEET_USER.createTextFinder('投稿数');
    const timesRanges = timesFinder.findAll();
    const timesRow    = ranges[0].getRow();
    const timesColumn = timesRanges[0].getColumn();
    const times = SHEET_USER.getRange(timesRow, timesColumn).getValue() + 1;
    
    // 更新日時を更新
    const updateDateFinder = SHEET_USER.createTextFinder('更新日時');
    const updateDateRanges = updateDateFinder.findAll();
    const updateDateRow    = ranges[0].getRow();
    const updateDateColumn = updateDateRanges[0].getColumn();
    const updateDate = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'); // 現在の日付を取得
    
    // 一括更新
    SHEET_USER.getRange(timesRow, timesColumn).setValue(times);
    SHEET_USER.getRange(updateDateRow, updateDateColumn).setValue(updateDate);
    
    // 更新順に並び替え（コメントアウトされているが、必要に応じて有効化）
    /**
    const numColumn = SHEET_USER.getLastColumn(); // 最後列の列番号を取得
    const numRow    = SHEET_USER.getLastRow() - 1;  // 最後行の行番号を取得
    const dataRange = SHEET_USER.getRange(2, 1, numRow, numColumn);
    dataRange.sort([{column: updateDateColumn, ascending: false}]); // 更新日付順に並び替え
    */
  } finally {
    // 必ずロックを解放
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

// ユーザーIDに基づいてユーザーのシート名を検索する関数（キャッシュ機能を追加）
function findUserSheetName(userId) {
  // キャッシュからシート名を取得を試みる
  const cache = CacheService.getScriptCache();
  const cachedSheetName = cache.get('sheetName_' + userId);
  
  if (cachedSheetName) {
    return cachedSheetName;
  }
  
  // キャッシュにない場合はスプレッドシートから検索
  const userData = SHEET_USER.getDataRange().getValues();
  for (let i = 0; i < userData.length; i++) {
    if (userData[i][0] === userId) {
      // 見つかったシート名をキャッシュに保存（1時間有効）
      cache.put('sheetName_' + userId, userData[i][4], 3600);
      return userData[i][4]; // 対応するシート名を返す
    }
  }
  
  // ユーザーが見つからない場合、シートを作成
  const newSheetName = addUser(userId);
  // 新しいシート名をキャッシュに保存
  cache.put('sheetName_' + userId, newSheetName, 3600);
  return newSheetName;
}

// 過去の履歴を取得（最適化）
function chatGPTLog(UserSheetName, postMessage) {
  let individualSheetLog = SS.getSheetByName(UserSheetName);
  // シートが見つからない場合、新しいシートを作成
  if (!individualSheetLog) {
    individualSheetLog = SS.insertSheet(UserSheetName);
    
    // 新しいシートにヘッダーを追加（一括書き込み）
    const headerValues = [['送信メッセージ', '応答メッセージ', '日付']];
    individualSheetLog.getRange(1, 1, 1, 3).setValues(headerValues);
    
    // 新しいシートの現在の行数と列数を取得する
    const maxRows = individualSheetLog.getMaxRows();
    const maxCols = individualSheetLog.getMaxColumns();
    
    // 余分な列を削除する
    if (maxCols > 3) individualSheetLog.deleteColumns(4, maxCols - 3);
    
    // 余分な行を削除する(1行のヘッダーを残して、他のすべての行を削除)
    if (maxRows > 1) individualSheetLog.deleteRows(2, maxRows - 1);
  }
  
  // データを一度に取得して処理
  let values = individualSheetLog.getDataRange().getValues();
  
  // valuesを逆順にする
  values = values.reverse();
  values.pop(); // ヘッダー行を削除
  
  // valuesをmapで処理し、MAX_COUNT_LOGの制限を超えないようにメッセージを追加
  const totalMessages = values.slice(0, MAX_COUNT_LOG).flatMap(value => [
    {"role": "assistant", "content": value[1]},
    {"role": "user", "content": value[0]}
  ]);
  
  // 最新のメッセージを追加
  totalMessages.reverse().push({"role": "user", "content": postMessage});
  
  return totalMessages;
}

// 古い実装は残しておくが、使用しない
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

// メンバーとしてユーザー登録されているか検索（キャッシュ機能を追加）
function findUser(uid) {
  // キャッシュからユーザーデータを取得を試みる
  const cache = CacheService.getScriptCache();
  const cachedUserData = cache.get('userData_' + uid);
  
  if (cachedUserData) {
    return cachedUserData;
  }
  
  // キャッシュにない場合はスプレッドシートから検索
  const userData = getUserData().reduce(function(uuid, row) { 
    return uuid || (row.key === uid && row.value); 
  }, false) || undefined;
  
  // 見つかった場合、キャッシュに保存（10分間有効）
  if (userData) {
    cache.put('userData_' + uid, userData, 600);
  }
  
  return userData;
}

// LINEのAPIを使ってユーザーの名前を取得（キャッシュ機能を追加）
function getUserDisplayName(userId) {
  // キャッシュからユーザー名を取得を試みる
  const cache = CacheService.getScriptCache();
  const cachedName = cache.get('displayName_' + userId);
  
  if (cachedName) {
    return cachedName;
  }
  
  // キャッシュにない場合はLINE APIから取得
  const url = 'https://api.line.me/v2/bot/profile/' + userId;
  const userProfile = UrlFetchApp.fetch(url, {
    'headers': {
      'Authorization' : `Bearer ${linetoken}`,
    },
  });
  
  const displayName = JSON.parse(userProfile).displayName;
  
  // 取得した名前をキャッシュに保存（1時間有効）
  cache.put('displayName_' + userId, displayName, 3600);
  
  return displayName;
}

// ユーザーのプロフィール画像取得（キャッシュ機能を追加）
function getUserDisplayIMG(userId) {
  // キャッシュからプロフィール画像URLを取得を試みる
  const cache = CacheService.getScriptCache();
  const cachedIMG = cache.get('profileIMG_' + userId);
  
  if (cachedIMG) {
    return cachedIMG;
  }
  
  // キャッシュにない場合はLINE APIから取得
  const url = 'https://api.line.me/v2/bot/profile/' + userId;
  const userProfile = UrlFetchApp.fetch(url, {
    'headers': {
      'Authorization' : `Bearer ${linetoken}`,
    },
  });
  
  const pictureUrl = JSON.parse(userProfile).pictureUrl;
  
  // 取得した画像URLをキャッシュに保存（1時間有効）
  cache.put('profileIMG_' + userId, pictureUrl, 3600);
  
  return pictureUrl;
}

// スプレッドシートを並び替え(対象のシートのカラムを降順に変更)
function dataSort(sortSheet, columnNumber) {
  // ロックを取得して競合を防止
  const lock = LockService.getScriptLock();
  try {
    // ロックを取得（最大5秒待機、10秒間ロック）
    lock.tryLock(5000);
    
    const numColumn = sortSheet.getLastColumn(); // 最後列の列番号を取得
    const numRow    = sortSheet.getLastRow() - 1;  // 最後行の行番号を取得
    
    if (numRow > 0) {
      const dataRange = sortSheet.getRange(2, 1, numRow, numColumn);
      dataRange.sort([{column: columnNumber, ascending: false}]); // 降順に並び替え
    }
  } finally {
    // 必ずロックを解放
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

// ユーザー情報取得（キャッシュ機能を追加）
function getUserData() {
  // キャッシュからユーザーデータを取得を試みる
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('allUserData');
  
  if (cachedData) {
    return JSON.parse(cachedData);
  }
  
  // キャッシュにない場合はスプレッドシートから取得
  const data = SHEET_USER.getDataRange().getValues();
  const userData = data.map(function(row) { 
    return {key: row[0], value: row[1]}; 
  });
  
  // 取得したデータをキャッシュに保存（5分間有効）
  cache.put('allUserData', JSON.stringify(userData), 300);
  
  return userData;
}

// 情報の取得（キャッシュ機能を追加）
function getInformationSheetDataAsCSV() {
  // キャッシュから情報データを取得を試みる
  const cache = CacheService.getScriptCache();
  const cachedInfo = cache.get('infoSheetData');
  
  if (cachedInfo) {
    return cachedInfo;
  }
  
  // キャッシュにない場合はスプレッドシートから取得
  const data = SHEET_INFO.getDataRange().getValues();
  let csvData = '';
  data.forEach(function(row) {
    csvData += row.join('  ') + '\n';
  });
  
  // 取得したデータをキャッシュに保存（5分間有効）
  cache.put('infoSheetData', csvData, 300);
  
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
