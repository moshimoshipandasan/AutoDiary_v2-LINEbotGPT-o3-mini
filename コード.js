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
const nonText =  "ごめんね。メッセージだけ送ってね";

//メニュー
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('アクティブセルlineUserIdにメッセージを送る', 'sendMessageToActiveCellUser')
    .addToUi();
}

// リクエスト処理のキャッシュ（メモリ内キャッシュ）
const requestCache = {};
const CACHE_EXPIRY = 60 * 1000; // キャッシュの有効期限（ミリ秒）

/**
 * メッセージをキューに追加する関数
 * 
 * 非同期処理の中核となる機能。受信したメッセージのリクエストIDをキューに追加し、
 * バックグラウンドで順次処理できるようにします。これにより複数のリクエストが同時に
 * 来た場合でも、処理の競合を避け、順番に処理することができます。
 * 
 * @param {string} requestId - 処理対象のリクエストを識別するためのID
 */
function addMessageToQueue(requestId) {
  // ScriptCacheを使用してキューを管理（サーバー再起動後も維持される）
  const cache = CacheService.getScriptCache();
  const queueKey = 'message_queue';
  
  // 現在のキューを取得（既存のキューがなければ空の配列を作成）
  let queue = [];
  const queueData = cache.get(queueKey);
  
  if (queueData) {
    queue = JSON.parse(queueData);
  }
  
  // リクエストIDをキューに追加（FIFOキュー - 先入れ先出し）
  queue.push(requestId);
  
  // 更新されたキューをキャッシュに保存（6時間有効）
  // 注: 長時間の処理が必要な場合は、この有効期限を調整する
  cache.put(queueKey, JSON.stringify(queue), 21600); // 6時間キャッシュを保持
  
  console.log(`メッセージをキューに追加しました: ${requestId}, キューサイズ: ${queue.length}`);
}

/**
 * LINEからのメッセージ受信（非同期処理対応版）
 * 
 * Webhookからのリクエストを受け取り、非同期処理を行うためのエントリーポイント。
 * 重要な特徴:
 * 1. 受信したメッセージをキューに追加し、バックグラウンドで処理
 * 2. 即時レスポンスを返すことでLINEプラットフォームのタイムアウトを回避
 * 3. 重要なメッセージは即時処理を開始
 * 
 * LINEプラットフォームは応答に3秒のタイムアウト制限があり、それを超えると
 * エラーとみなされます。この関数は即座に200 OKを返すことでタイムアウトを回避し、
 * 実際の処理（ChatGPT APIの呼び出しなど時間のかかる処理）は
 * バックグラウンドで非同期に行います。
 * 
 * @param {Object} e - Webhookからのリクエストデータ
 * @return {Object} 即時レスポンス
 */
function doPost(e) {
  const data = JSON.parse(e.postData.contents).events[0]; // LINEからのイベント情報を取得
  if (!data) {
    return SHEET_CHECK.appendRow(["Webhook設定接続確認しました。"]);
  }
  
  // 非同期処理のためにリクエストIDを生成（一意のIDを使用）
  const requestId = Utilities.getUuid();
  
  try {
    // ステップ1: リクエスト情報をキャッシュに保存
    // 後で非同期処理するために、リクエストデータを一時的に保存します
    CacheService.getScriptCache().put(
      requestId, 
      JSON.stringify({
        timestamp: new Date().getTime(),
        data: data
      }),
      21600 // 6時間キャッシュを保持（秒単位）
    );
    
    // ステップ2: 即時処理を行う
    // キューには追加せず、即時に処理を行います
    // これにより、メッセージが2回処理される問題を解決しつつ、即時応答も可能にします
    processMessageAsync(requestId);
    
    // 注意: キューには追加しない
    // processQueuedMessages関数による二重処理を防ぐため、キューには追加しません
    // 以前はここで addMessageToQueue(requestId) を呼び出していたが、削除した
    
    // ステップ3: 即時レスポンスを返す（HTTP 200）
    // ここが重要: LINEプラットフォームのタイムアウト（3秒）を回避するため、
    // 実際の処理（ChatGPT APIの呼び出しなど）を待たずに即座に応答を返します
    // これにより、ユーザーからのメッセージが「配信済み」として扱われ、
    // 実際の応答は後から非同期で送信されます
    return ContentService.createTextOutput(JSON.stringify({
      status: "processing",
      requestId: requestId
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.log("エラー発生:", error);
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 非同期メッセージ処理（トリガーで実行）
function processMessageAsync(requestId) {
  // キャッシュからリクエストデータを取得
  const cachedData = CacheService.getScriptCache().get(requestId);
  if (!cachedData) {
    console.log("キャッシュデータが見つかりません: " + requestId);
    return;
  }
  
  const requestData = JSON.parse(cachedData);
  const data = requestData.data;
  
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
      
      // 処理中メッセージを先に送信（オプション）
      // sendMessage(replyToken, "メッセージを処理中です...");
      
      try {
        // データ生成&LINEに送信（非同期処理）
        const totalMessages = chatGPTLog(userSheetName, postMessage);
        
        // キャッシュをチェック（同じユーザーからの同じメッセージの場合）
        const cacheKey = `${lineUserId}_${postMessage}`;
        const cachedResponse = requestCache[cacheKey];
        
        let replyText;
        if (cachedResponse && (new Date().getTime() - cachedResponse.timestamp < CACHE_EXPIRY)) {
          // キャッシュから応答を取得
          console.log("キャッシュから応答を取得: " + cacheKey);
          replyText = cachedResponse.response;
        } else {
          // ChatGPT APIを呼び出して応答を取得
          replyText = chatGPT(lineUserId, totalMessages);
          
          // 応答をキャッシュに保存
          requestCache[cacheKey] = {
            timestamp: new Date().getTime(),
            response: replyText
          };
        }
        
        // LINEに応答を送信
        sendMessage(replyToken, replyText);
        
        // 処理済みフラグをキャッシュに保存（二重処理防止）
        const processedKey = `processed_${requestId}`;
        CacheService.getScriptCache().put(processedKey, 'true', 21600); // 6時間キャッシュを保持
        
        // ログに追加（非同期で実行可能）
        debugLog(lineUserId, postMessage, replyText);
        debugLogIndividual(userSheetName, postMessage, replyText);
      } catch (error) {
        console.log("メッセージ処理エラー:", error);
        sendMessage(replyToken, "申し訳ありません。エラーが発生しました。しばらくしてからもう一度お試しください。");
      }
    } else {
      // テキスト以外のメッセージの場合、サポートしていないメッセージタイプである旨をユーザーに通知
      sendMessage(replyToken, nonText);
    }
  }
  // その他のdataTypeに対する処理があればここに追加
}

/**
 * トリガーを削除する関数
 * 
 * この関数は、キュー処理用のトリガーを削除します。
 * 二重処理の問題を解決するために、キュー処理を無効化します。
 * スクリプトエディタから手動で実行する必要があります。
 */
function setupTriggers() {
  // 既存のトリガーをクリア
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processQueuedMessages') {
      ScriptApp.deleteTrigger(triggers[i]);
      deletedCount++;
    }
  }
  
  // トリガーが削除されたことを確認
  if (deletedCount > 0) {
    console.log(`${deletedCount}個のキュー処理用トリガーを削除しました`);
    return `${deletedCount}個のキュー処理用トリガーを削除しました。これにより二重処理の問題が解決されます。`;
  } else {
    console.log("削除すべきトリガーが見つかりませんでした");
    return "削除すべきトリガーが見つかりませんでした。既にトリガーが削除されている可能性があります。";
  }
  
  // 注意: 以前はここで新しいトリガーを作成していましたが、
  // 二重処理の問題を解決するために、トリガーの作成を停止しました。
  // 即時処理のみを行うことで、メッセージが2回処理される問題を解決します。
}

/**
 * トリガーの状態を確認する関数
 * 
 * この関数は、現在設定されているトリガーの状態を確認します。
 * キュー処理用のトリガーが正しく設定されているかを確認できます。
 */
function checkTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const queueTriggers = triggers.filter(trigger => 
    trigger.getHandlerFunction() === 'processQueuedMessages'
  );
  
  if (queueTriggers.length > 0) {
    return "キュー処理用トリガーが設定されています。数: " + queueTriggers.length;
  } else {
    return "キュー処理用トリガーが設定されていません。setupTriggers()関数を実行してください。";
  }
}

/**
 * キューに入っているメッセージを処理する関数（改良版）
 * 
 * この関数は定期的に実行され（トリガーによって1分ごとなど）、キューに溜まった
 * メッセージを取り出して処理します。非同期処理の中心的な役割を担い、
 * 複数のリクエストを効率的に処理します。
 * 
 * 主な特徴:
 * 1. キューからバッチ単位でメッセージを取り出し、並行処理
 * 2. ChatGPTリクエストを一括処理して効率化
 * 3. エラーハンドリングとリトライ機能
 */
function processQueuedMessages() {
  // ステップ1: キャッシュからキューを取得
  const cache = CacheService.getScriptCache();
  const queueKey = 'message_queue';
  const queueData = cache.get(queueKey);
  
  // キューが空の場合は処理終了
  if (!queueData) {
    console.log("処理すべきメッセージはありません");
    return;
  }
  
  const queue = JSON.parse(queueData);
  if (queue.length === 0) {
    console.log("キューは空です");
    return;
  }
  
  console.log(`キュー内のメッセージ数: ${queue.length}`);
  
  // ステップ2: バッチサイズを決定し、キューから取り出す
  // 同時処理数を制限することで、リソース使用量を管理し安定性を確保
  const batchSize = Math.min(5, queue.length);
  const batch = queue.splice(0, batchSize);
  
  console.log(`バッチ処理するメッセージ数: ${batch.length}`);
  
  // ステップ3: ChatGPTリクエストをバッチ処理するためのリクエスト配列を準備
  // 複数のChatGPTリクエストを一括で処理することで効率化
  const chatGPTRequests = [];
  
      // ステップ4: バッチ内の各リクエストを処理
      batch.forEach(requestId => {
        try {
          // 処理済みフラグをチェック（二重処理防止）
          const processedKey = `processed_${requestId}`;
          const isProcessed = cache.get(processedKey);
          
          if (isProcessed) {
            console.log(`リクエスト ${requestId} は既に処理済みです。スキップします。`);
            return;
          }
          
          // キャッシュからリクエストデータを取得
          const cachedData = cache.get(requestId);
          if (!cachedData) {
            console.log(`キャッシュデータが見つかりません: ${requestId}`);
            return;
          }
          
          const requestData = JSON.parse(cachedData);
          const data = requestData.data;
          
          // メッセージタイプのリクエストをChatGPTバッチ処理用に収集
          // テキストメッセージのみをChatGPTバッチ処理の対象とする
          if (data.type === "message" && data.message.type === "text") {
            const lineUserId = data.source.userId;
            const userSheetName = findUserSheetName(lineUserId);
            const postMessage = data.message.text;
            const totalMessages = chatGPTLog(userSheetName, postMessage);
            
            // バッチ処理用のリクエスト情報を収集
            chatGPTRequests.push({
              userId: lineUserId,
              totalMessages: totalMessages,
              requestId: requestId,
              replyToken: data.replyToken,
              postMessage: postMessage,
              userSheetName: userSheetName
            });
      } else {
        // ChatGPT以外のリクエスト（フォローイベントなど）は個別に処理
        processMessageAsync(requestId);
      }
    } catch (error) {
      console.log(`リクエスト処理エラー (${requestId}): ${error}`);
    }
  });
  
  // ステップ5: ChatGPTリクエストを一括処理
  // 複数のリクエストを一度に処理することで、APIコール回数を削減し効率化
  if (chatGPTRequests.length > 0) {
    console.log(`ChatGPTバッチ処理するリクエスト数: ${chatGPTRequests.length}`);
    
    try {
      // バッチ処理実行 - 複数のChatGPTリクエストを並列処理
      const results = batchProcessChatGPTRequests(chatGPTRequests);
      
      // ステップ6: 結果を処理して各ユーザーに返信
      results.forEach(result => {
        if (result.success) {
          // 成功した場合、対応するリクエストを検索
          const request = chatGPTRequests.find(req => req.requestId === result.requestId);
          if (request) {
            // LINEに応答を送信
            sendMessage(request.replyToken, result.response);
            
            // 処理済みフラグをキャッシュに保存（二重処理防止）
            const processedKey = `processed_${request.requestId}`;
            cache.put(processedKey, 'true', 21600); // 6時間キャッシュを保持
            
            // ログに追加
            debugLog(request.userId, request.postMessage, result.response);
            debugLogIndividual(request.userSheetName, request.postMessage, result.response);
            
            // キャッシュに応答を保存（同じリクエストが来た場合に再利用）
            const cacheKey = `${request.userId}_${request.postMessage}`;
            requestCache[cacheKey] = {
              timestamp: new Date().getTime(),
              response: result.response
            };
          }
        } else {
          // 失敗した場合、個別に処理してエラーメッセージを返信
          console.log(`バッチ処理失敗 (${result.requestId}): ${result.error}`);
          const request = chatGPTRequests.find(req => req.requestId === result.requestId);
          if (request) {
            sendMessage(request.replyToken, "申し訳ありません。エラーが発生しました。しばらくしてからもう一度お試しください。");
          }
        }
      });
    } catch (error) {
      // 全体的なエラーが発生した場合の処理
      console.log(`ChatGPTバッチ処理全体エラー: ${error}`);
      // すべてのリクエストに対してエラーメッセージを返信
      chatGPTRequests.forEach(request => {
        sendMessage(request.replyToken, "申し訳ありません。エラーが発生しました。しばらくしてからもう一度お試しください。");
      });
    }
  }
  
  // ステップ7: 更新されたキューを保存
  // 処理済みのリクエストを除いた残りのキューを保存
  if (queue.length > 0) {
    cache.put(queueKey, JSON.stringify(queue), 21600);
  } else {
    cache.remove(queueKey);
  }
  
  console.log("キュー処理完了");
}

/**
 * LINEに返答を送信する関数（非同期対応版）
 * 
 * 非同期処理の最終段階として、処理結果をLINEユーザーに返信します。
 * リトライ機能を備えており、一時的なネットワークエラーや
 * LINE APIの一時的な障害に対して耐性があります。
 * 
 * 特徴:
 * 1. 最大3回のリトライ
 * 2. 指数バックオフによる再試行間隔の調整
 * 3. 詳細なエラーログ
 * 4. トークン有効期限チェック
 * 5. プッシュメッセージへのフォールバック
 * 
 * @param {string} replyToken - LINEからのリプライトークン（24時間有効）
 * @param {string} replyText - ユーザーに送信するテキストメッセージ
 * @param {string} [userId] - ユーザーID（リプライトークンが無効な場合のフォールバック用）
 * @return {Object|null} 成功時はレスポンスオブジェクト、失敗時はnull
 */
function sendMessage(replyToken, replyText, userId = null) {
  // テキストが長すぎる場合は分割（LINEの制限は5000文字）
  if (replyText.length > 4000) {
    console.log("テキストが長すぎるため分割します");
    const firstPart = replyText.substring(0, 4000);
    const secondPart = replyText.substring(4000);
    
    // 最初の部分を送信
    const result1 = sendMessage(replyToken, firstPart);
    
    // 残りの部分はプッシュメッセージとして送信（リプライトークンは1回しか使えないため）
    if (userId && secondPart.length > 0) {
      setTimeout(() => {
        sendMessageToUser(userId, secondPart);
      }, 500); // 少し遅延させて送信
    }
    
    return result1;
  }
  
  // リトライ機能付き送信
  const maxRetries = 3;
  let retryCount = 0;
  let success = false;
  
  while (retryCount < maxRetries && !success) {
    try {
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
        "payload" : JSON.stringify(postData),
        "muteHttpExceptions": true // エラーを黙らせて例外をキャッチできるようにする
      };
      
      // LINE Messaging APIにリクエスト送信
      const response = UrlFetchApp.fetch(LINE_REPLY_URL, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      // 成功（2xx）の場合
      if (responseCode >= 200 && responseCode < 300) {
        console.log("LINE メッセージ送信成功");
        success = true;
        return response;
      } else {
        // APIエラーの場合（4xx、5xxなど）
        console.log(`LINE API エラー (${responseCode}): ${responseText}`);
        
        // トークンが無効または期限切れの場合（リプライトークンは24時間有効）
        if (responseCode === 400 && responseText.includes("Invalid reply token")) {
          console.log("リプライトークンが無効です。プッシュメッセージにフォールバックします。");
          
          // ユーザーIDが提供されている場合、プッシュメッセージを試みる
          if (userId) {
            return sendMessageToUser(userId, replyText);
          } else {
            console.log("フォールバック用のユーザーIDがありません。メッセージを送信できません。");
            return null;
          }
        }
        
        retryCount++;
        
        if (retryCount < maxRetries) {
          // 指数バックオフ（再試行間隔を徐々に長くする）
          // 1回目: 1秒、2回目: 2秒、3回目: 4秒...
          Utilities.sleep(1000 * Math.pow(2, retryCount));
        }
      }
    } catch (error) {
      // ネットワークエラーなどの例外
      console.log(`LINE メッセージ送信エラー: ${error}`);
      retryCount++;
      
      if (retryCount < maxRetries) {
        // 指数バックオフ
        Utilities.sleep(1000 * Math.pow(2, retryCount));
      }
    }
  }
  
  // すべてのリトライが失敗した場合
  if (!success) {
    console.log(`LINE メッセージ送信失敗: 最大リトライ回数(${maxRetries})に達しました`);
    
    // 最後の手段としてプッシュメッセージを試みる
    if (userId) {
      console.log("プッシュメッセージにフォールバックします");
      return sendMessageToUser(userId, replyText);
    }
    
    return null;
  }
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

// ChatGPTへの指示（非同期対応版）
function chatGPT(userId, totalMessages) {
  // リトライ機能付きChatGPT API呼び出し
  const maxRetries = 3;
  let retryCount = 0;
  
  // ユーザー情報とプロンプトの準備
  const constraints = SHEET.getRange(1, 1).getValue(); // スプシからプロンプト
  const userName = getUserDisplayName(userId); // ユーザーのプロフィール名を取得
  const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'); // 現在の日時を取得
  const info = getInformationSheetDataAsCSV();
  
  // システムプロンプトを追加
  totalMessages.unshift({
    "role": "system",
    "content": `${constraints}\n
# ユーザーの名前:${userName}\n
# 現在時刻:${date}\n
# 過去の日記:${info}
`
  });
  
  // リクエストオプションの設定
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
    }),
    "muteHttpExceptions": true // エラーを黙らせて例外をキャッチできるようにする
  };
  
  // リトライループ
  while (retryCount < maxRetries) {
    try {
      const response = UrlFetchApp.fetch(CHAT_GPT_URL, requestOptions);
      const responseCode = response.getResponseCode();
      
      if (responseCode >= 200 && responseCode < 300) {
        const responseText = response.getContentText();
        return checkChatGPTResponse(responseText);
      } else {
        console.log(`ChatGPT API エラー (${responseCode}): ${response.getContentText()}`);
        retryCount++;
        
        if (retryCount < maxRetries) {
          // 指数バックオフ（再試行間隔を徐々に長くする）
          Utilities.sleep(1000 * Math.pow(2, retryCount));
        }
      }
    } catch (error) {
      console.log("ChatGPT API呼び出しエラー:", error);
      retryCount++;
      
      if (retryCount < maxRetries) {
        // 指数バックオフ
        Utilities.sleep(1000 * Math.pow(2, retryCount));
      }
    }
  }
  
  // すべてのリトライが失敗した場合
  console.log(`ChatGPT API呼び出し失敗: 最大リトライ回数(${maxRetries})に達しました`);
  return "申し訳ありません。ChatGPTとの通信中にエラーが発生しました。しばらくしてからもう一度お試しください。";
}

/**
 * 複数のChatGPTリクエストを一括処理する関数
 * 
 * 非同期処理の効率化のための重要な機能。複数のユーザーからのリクエストを
 * 一度にまとめて処理することで、APIコール回数を削減し、全体的な処理時間を短縮します。
 * 
 * 特徴:
 * 1. 複数リクエストの並列処理（UrlFetchApp.fetchAllを使用）
 * 2. バッチサイズの制限による安定性確保
 * 3. 詳細なエラーハンドリングと結果の追跡
 * 
 * この関数は非同期処理アーキテクチャの中核部分で、複数ユーザーが同時にアクセスした場合の
 * パフォーマンスを大幅に向上させます。
 * 
 * @param {Array} requests - 処理すべきChatGPTリクエストの配列
 * @return {Array} 各リクエストの処理結果を含む配列
 */
function batchProcessChatGPTRequests(requests) {
  // ステップ1: バッチサイズを決定（リソース使用量を制限）
  // 同時に処理するリクエスト数を制限することで、APIレート制限やメモリ使用量を管理
  const batchSize = Math.min(5, requests.length);
  const batch = requests.slice(0, batchSize);
  
  // ステップ2: 各リクエストのURLFetchApp.fetchを準備
  // 各ユーザーのリクエストに対して、適切なプロンプトとパラメータを設定
  const httpRequests = batch.map(req => {
    const { userId, totalMessages, requestId } = req;
    
    // ユーザー情報とプロンプトの準備
    const constraints = SHEET.getRange(1, 1).getValue();
    const userName = getUserDisplayName(userId);
    const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const info = getInformationSheetDataAsCSV();
    
    // システムプロンプトを追加（各ユーザー固有の情報を含む）
    totalMessages.unshift({
      "role": "system",
      "content": `${constraints}\n
# ユーザーの名前:${userName}\n
# 現在時刻:${date}\n
# 過去の日記:${info}
`
    });
    
    // リクエストオプションの設定（ChatGPT APIへのリクエスト準備）
    return {
      url: CHAT_GPT_URL,
      options: {
        "method": "post",
        "headers": {
          "Content-Type": "application/json",
          "Authorization": "Bearer "+ apikey
        },
        "payload": JSON.stringify({
          "model": CHAT_GPT_VER,
          "messages": totalMessages,
          "reasoning_effort": "medium"
        }),
        "muteHttpExceptions": true
      },
      requestId: requestId // 後で結果を追跡するためのID
    };
  });
  
  try {
    // ステップ3: 一括リクエスト実行の準備
    const fetchRequests = httpRequests.map(req => ({ 
      url: req.url, 
      options: req.options,
      requestId: req.requestId 
    }));
    
    // ステップ4: fetchAllを使用して並列リクエスト実行
    // これが非同期処理の核心部分 - 複数のリクエストを同時に処理
    // 通常のfetchを使うと逐次処理になるが、fetchAllを使うことで並列処理が可能
    const responses = UrlFetchApp.fetchAll(fetchRequests.map(req => ({ 
      url: req.url, 
      options: req.options 
    })));
    
    // ステップ5: レスポンスを処理して結果を返す
    const results = responses.map((response, index) => {
      const requestId = fetchRequests[index].requestId;
      try {
        const responseCode = response.getResponseCode();
        
        // 成功の場合（HTTPステータスコード2xx）
        if (responseCode >= 200 && responseCode < 300) {
          const responseText = response.getContentText();
          return {
            requestId: requestId,
            success: true,
            response: checkChatGPTResponse(responseText)
          };
        } else {
          // APIエラーの場合（4xx、5xxなど）
          console.log(`バッチ処理エラー (${responseCode}): ${response.getContentText()}`);
          return {
            requestId: requestId,
            success: false,
            error: `API エラー (${responseCode})`
          };
        }
      } catch (error) {
        // 例外が発生した場合
        console.log(`バッチ処理例外: ${error}`);
        return {
          requestId: requestId,
          success: false,
          error: error.toString()
        };
      }
    });
    
    return results;
  } catch (error) {
    // 全体的なエラーが発生した場合（ネットワーク障害など）
    console.log(`バッチ処理全体エラー: ${error}`);
    return httpRequests.map(req => ({
      requestId: req.requestId,
      success: false,
      error: error.toString()
    }));
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
