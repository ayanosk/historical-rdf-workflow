//以下の関数のみを実行
function cleanTopicData() {
  Logger.log("=== トピックデータのクリーニング開始 ===");

  clearTopicFieldsIfAttendanceOnly();
  clearInvalidMentionedAsSubsequent();

  Logger.log("=== トピックデータのクリーニング完了 ===");
}

// 以下、補助関数

function clearTopicFieldsIfAttendanceOnly() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("統合データ");
  if (!sheet) {
    Logger.log("統合データシートが見つかりません");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // データなし

  const fValues = sheet.getRange(2, 6, lastRow - 1).getValues(); // F列: 出席かどうか
  const aRange = sheet.getRange(2, 1, lastRow - 1); // A列: トピックID
  const dRange = sheet.getRange(2, 4, lastRow - 1); // D列: トピック順

  const aValues = aRange.getValues();
  const dValues = dRange.getValues();

  let changed = false;

  for (let i = 0; i < fValues.length; i++) {
    const status = fValues[i][0];
    if (status === "出席") {
      aValues[i][0] = "";
      dValues[i][0] = "";
      changed = true;
    }
  }

  if (changed) {
    aRange.setValues(aValues);
    dRange.setValues(dValues);
    Logger.log("✅ 出席行のトピックIDとトピック順を削除しました。");
  } else {
    Logger.log("出席該当行はありませんでした。");
  }
}

// 不要なMentionedAsSubsequentを削除する
function clearInvalidMentionedAsSubsequent() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("統合データ");
  if (!sheet) {
    Logger.log("統合データシートが見つかりません");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const topicIds = sheet.getRange(2, 1, lastRow - 1).getValues(); // A列: トピックID
  const mentionedValues = sheet.getRange(2, 5, lastRow - 1).getValues(); // E列: mentionedAsSubsequent

  const topicIdSet = new Set(topicIds.map(row => row[0]));

  let changed = false;
  for (let i = 0; i < mentionedValues.length; i++) {
    const val = mentionedValues[i][0];
    if (val && !topicIdSet.has(val)) {
      mentionedValues[i][0] = ""; // 存在しないIDなら削除
      changed = true;
    }
  }

  if (changed) {
    sheet.getRange(2, 5, mentionedValues.length, 1).setValues(mentionedValues);
    Logger.log("✅ mentionedAsSubsequent 列に存在しないトピックIDが削除されました。");
  } else {
    Logger.log("mentionedAsSubsequent はすべて有効なトピックIDです。");
  }
}
