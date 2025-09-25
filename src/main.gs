// ChatworkAPIトークンを設定
const CHATWORK_API =
  PropertiesService.getScriptProperties().getProperty("CHATWORK_API");
// ChatworkルームIDを設定
const CHATWORK_ROOM_ID =
  PropertiesService.getScriptProperties().getProperty("CHATWORK_ROOM_ID");

/**
 * 今日の作業件数をカウントする関数
 */
function countTodayRows() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("管理シート");
  const values = sheet.getRange("A2:A").getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0); // 今日の日付のみを比較するため時刻をリセット
  let count = 0;

  for (let i = 0; i < values.length; i++) {
    const cellValue = values[i][0];
    if (cellValue instanceof Date) {
      let d = new Date(cellValue);
      d.setHours(0, 0, 0, 0);
      if (d.getTime() === today.getTime()) {
        count++;
      }
    }
  }
  console.log(count);
  return count;
}

/**
 * 作業開始をChatworkに通知する関数
 */
function notifyStart() {
  try {
    // ChatworkAPIクライアントを作成
    const client = ChatWorkClient.factory({ token: CHATWORK_API });

    // 送信するメッセージを作成
    const message = `お疲れ様です。\n本日の作業開始いたします。`;

    // Chatworkにメッセージを送信
    const response = client.sendMessage({
      room_id: CHATWORK_ROOM_ID,
      body: message,
    });

    // 送信完了を通知
    SpreadsheetApp.getUi().alert(
      "成功",
      "Chatworkに作業開始の通知を送信しました！",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    // エラーハンドリング
    console.error("エラー:", error);
    SpreadsheetApp.getUi().alert(
      "エラー",
      "メッセージの送信に失敗しました。\nAPIトークンとルームIDを確認してください。",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 作業終了をChatworkに通知する関数
 */
function notifyEnd() {
  try {
    const client = ChatWorkClient.factory({ token: CHATWORK_API });

    const count = countTodayRows();

    const message = `お疲れ様です。本日${count}件出品作業しました。\nご確認よろしくお願いいたします。`;

    const response = client.sendMessage({
      room_id: CHATWORK_ROOM_ID,
      body: message,
    });

    SpreadsheetApp.getUi().alert(
      "成功",
      "Chatworkに作業終了の通知を送信しました！",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    console.error("エラー:", error);
    SpreadsheetApp.getUi().alert(
      "エラー",
      "メッセージの送信に失敗しました。\nAPIトークンとルームIDを確認してください。",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
