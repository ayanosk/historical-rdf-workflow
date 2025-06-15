function handleTopicManagement(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const column = range.getColumn();
  const row = range.getRow();

  // B列が編集された場合のみ処理を実行
  if (column === 2) {
    copyValue(sheet, 'J2', `J${row}`); // J2の値をJ列の編集行にコピー
    copyValue(sheet, 'K2', `K${row}`); // K2の値をK列の編集行にコピー
    copyValue(sheet, 'L2', `L${row}`); // L2の値をL列の編集行にコピー
  }

  if (row >= 3) {
    // B列とC列の両方に値が入力されている場合にA列にトピックIDを設定
    const dateValue = sheet.getRange(row, 2).getValue();
    const numericValue = sheet.getRange(row, 3).getValue();

    if (dateValue && numericValue) {
      const topicID = generateTopicID(dateValue, numericValue);
      sheet.getRange(row, 1).setValue(topicID);
    }
  }

  // 開始頁（G列）が編集された場合にアノテーションURI（E列）の値を設定
  if (column === 7) {
    const baseurl = sheet.getRange(2, 5).getValue(); // E2の値
    const gValue = sheet.getRange(row, 7).getValue(); // G列の値
    const eValue = baseurl + gValue.toString() + "/edit";
    sheet.getRange(row, 5).setValue(eValue); // E列に値を設定
  }

}


// トピックIDを生成する
function generateTopicID(dateValue, numericValue) {
  // 日付をYYYYMMDD形式に変換
  const formattedDate = Utilities.formatDate(new Date(dateValue), Session.getScriptTimeZone(), 'yyyyMMdd');

  // 数値を4桁の文字列に変換
  const formattedNumber = (numericValue * 10).toFixed(0).padStart(4, '0');

  return `${formattedDate}-${formattedNumber}`;
}