// --- 定数定義 ---
const SHEET_NAME = 'todo';

const REQUIRED_HEADERS_MAP = {
  'No': 'no',
  'Todo': 'todo',
  '期限': 'deadline',
  '難易度': 'difficulty'
};

const headerRange = 'A1:H1';

const DEADLINE_CATEGORIES = {
  OVERDUE: { label: '❗期限切れ --------' },
  TODAY: { label: '⏰ 本日期限 --------' },
  WITHIN_3_DAYS: { label: '⚠️ 3日以内 --------' },
  WITHIN_7_DAYS: { label: '📌 7日以内 --------' },
};

const INITIAL_NOTIFICATION_MSG = 'Todoの情報をお知らせします\n\n';
const NO_TASKS_MSG = '通知対象のToDoはありません。';


// !!! LINE Messaging APIの設定 !!!

// ステップ1で取得した「チャネルアクセストークン」
const LINE_CHANNEL_ACCESS_TOKEN = 'YOUR_LINE_CHANNEL_ACCESS_TOKEN'; // ポートフォリオ用にマスキング
const LINE_USER_ID = 'YOUR_LINE_USER_ID'; // ポートフォリオ用にマスキング


// --- メイン関数 ---
function checkDeadlinesAndNotify() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log(`エラー: スプレッドシート「${SHEET_NAME}」が見つかりません。`);
    return;
  }

  const headerValues = sheet.getRange(headerRange).getValues()[0];
  const headerIndex = getHeaderIndex(headerValues, REQUIRED_HEADERS_MAP);

  // 必須ヘッダーの存在チェック
  const missingHeaders = Object.keys(REQUIRED_HEADERS_MAP).filter(
    header => headerIndex[REQUIRED_HEADERS_MAP[header]] === -1
  );
  if (missingHeaders.length > 0) {
    Logger.log(`必要なヘッダーが見つかりません: ${missingHeaders.join(', ')}。`);
    return;
  }

  const data = getValidRangeData(sheet);
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const categorizedTasks = {
    overdue: [],
    today: [],
    within3Days: [],
    within7Days: []
  };

  for (const row of data) {
    const no = row[headerIndex.no];
    const todo = row[headerIndex.todo];
    const deadline = row[headerIndex.deadline];
    
    // 日付が無効な場合のチェックを追加
    if (!(deadline instanceof Date) || isNaN(deadline.getTime())) {
      Logger.log(`警告: No: ${no} のTodo「${todo}」の期限が無効な日付です。スキップします。`);
      continue;
    }
    deadline.setHours(0, 0, 0, 0);


    const diffDays = Math.ceil((deadline.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
    if (diffDays < 0) { /** 期限切れTodo */
      categorizedTasks.overdue.push({ no: no, todo: todo });

    } else if (diffDays === 0) { /** 本日期限のTodo */
      categorizedTasks.today.push({ no: no, todo: todo });

    } else if (diffDays <= 3) { /** 期限まであと3日のTodo */
      categorizedTasks.within3Days.push({ no: no, todo: todo });

    } else if (diffDays <= 7) { /** 期限まであと7日のTodo */
      categorizedTasks.within7Days.push({ no: no, todo: todo });
    }
  }

  let msg = INITIAL_NOTIFICATION_MSG;
  msg = appendCategoryMessage(msg, categorizedTasks.overdue, DEADLINE_CATEGORIES.OVERDUE.label);
  msg = appendCategoryMessage(msg, categorizedTasks.today, DEADLINE_CATEGORIES.TODAY.label);
  msg = appendCategoryMessage(msg, categorizedTasks.within3Days, DEADLINE_CATEGORIES.WITHIN_3_DAYS.label);
  msg = appendCategoryMessage(msg, categorizedTasks.within7Days, DEADLINE_CATEGORIES.WITHIN_7_DAYS.label);
  
  if (
    msg.trim() === INITIAL_NOTIFICATION_MSG.trim() ||
    msg.trim() === 'Todoの情報をお知らせします'
  ) {
    msg = NO_TASKS_MSG;
  }

  // LINE Messaging APIで通知を送信する
  sendLinePushMessage(msg);
}


// --- ヘルパー関数 ---

/**
 * ヘッダーから各列のインデックスを取得する
 * @param {Array<string>} headerRow - スプレッドシートのヘッダー行の値
 * @param {Object<string, string>} requiredHeadersMap - 必須ヘッダー名とその内部的なキーのマッピング
 * * @returns {Object<string, number>} 各内部キーとそれに対応するインデックスのマッピング
 */
function getHeaderIndex(headerRow, requiredHeadersMap) {
  const indexMap = {};
  headerRow.forEach((h, i) => {
    // マップに存在するヘッダー名であれば、対応する内部キーでインデックスを保存
    const internalKey = Object.keys(requiredHeadersMap).find(key => key === h);
    if (internalKey) indexMap[requiredHeadersMap[internalKey]] = i;
  });

  // 以下は、必要な内部キーすべてに対して、存在チェック（≒ -1 をセット）を保証する
  const result = {};
  for (const internalKey of Object.values(requiredHeadersMap)) {
    result[internalKey] = indexMap[internalKey] !== undefined ? indexMap[internalKey] : -1;
  }
  return result;
};

/**
 * 特定のカテゴリのToDoリストをメッセージに追加する
 * * @param {string} currentMsg - 現在のメッセージ文字列
 * @param {Array<Object>} todos - 該当カテゴリのToDoリスト
 * @param {string} headerLabel - カテゴリの見出しテキスト
 * * @returns {string} 更新されたメッセージ文字列
 */
function appendCategoryMessage(currentMsg, todos, headerLabel) {
  if(todos.length > 0) {
    currentMsg += `${headerLabel}\n`;
    todos.forEach(todo => currentMsg += `${todo.no}. ${todo.todo}\n`);
    currentMsg += '\n';
  }
  return currentMsg;
}


// --- LINE Messaging API 用 ---

/**
 * LINE Messaging APIを通じてPushメッセージを送信する
 * @param {string} messageText - 送信するメッセージ本文
 */
function sendLinePushMessage(messageText) {

  if (!LINE_CHANNEL_ACCESS_TOKEN) {
    Logger.log('エラー: LINEチャネルアクセストークンが設定されていません。');
    return;
  }
  if (!LINE_USER_ID) {
    Logger.log('エラー: LINEユーザーIDが設定されていません。');
    return;
  }

  const LINE_MESSAGING_API_URL = 'https://api.line.me/v2/bot/message/push';

  const payload = {
    to: LINE_USER_ID,
    messages: [
      {
        type: 'text',
        text: messageText,
      },
    ],
  };

  const options = {
    'method': 'post',
    'headers': {
      'Content-type': 'application/json',
      'Authorization': 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN,
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true, // エラー時も例外を投げずにレスポンスを取得する
  };

  try {
    const response = UrlFetchApp.fetch(LINE_MESSAGING_API_URL, options);
    const responseText = response.getContentText();

    // LINE Messaging APIの成功ステータスは通常200ですが、エラー情報も含まれる
    if (response.getResponseCode() === 200) {
      Logger.log('LINE Pushメッセージを送信しました。');

    } else {
      Logger.log(`LINE Pushメッセージの送信に失敗しました。ステータス: ${response.getResponseCode()}, レスポンス: ${responseText}`);
    }

  } catch (e) {
    Logger.log('LINE Pushメッセージの送信中に例外が発生しました: ' + e.message);
  }
}

// 【LINEユーザーIDをWebhook経由で取得する場合の関数】
function doPost(e) {
  const json = JSON.parse(e.postData.contents);
  Logger.log(json); // 受信したWebhookイベント全体をログに出力

  if (json.events && json.events.length > 0) {
    const userId = json.events[0].source.userId;
    Logger.log('!!! LINE User ID: ' + userId + ' !!!');
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
}