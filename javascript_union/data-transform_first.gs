// 統合データ処理フローの一括実行関数。これだけ実行すればOK
// 初年のみ
function runAllTopicDataProcessing() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("統合データ");
  if (!sheet) {
    Logger.log("統合データシートが見つかりません");
    return;
  }

  // ✅ C列（3列目）を整数表示に設定（2行目以降）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const cRange = sheet.getRange(2, 3, lastRow - 1); // C列: col = 3
    cRange.setNumberFormat("0"); // 整数フォーマット
    Logger.log("✅ C列の表示形式を整数に設定しました。");
  }

  Logger.log("=== 統合データ 処理開始 ===");

  generateAssemblyID();
  insertMentionedAsSubsequent();
  clearMentionedAsSubsequentIfFirstTopic();

  insertScLabelPairColumn("トピック種別", "sc_recogito", "recogito", "recogito", "Recogito");
  insertScLabelPairColumn("付与チェック", "sc_gallica", "gallica", "gallica", "gallica");
  insertScLabelPairColumn("資料URI", "sc_iiif", "iiif", "iiif", "iiif");
  insertScLabelPairColumn("IIIF manifest URI", "sc_orig", "original", "original", "original");

  insertSourceAndSourceIdColumns();

  Logger.log("=== 統合データ 処理完了 ===");
}


// ここから先は処理用の関数
function generateAssemblyID() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("統合データ");
  if (!sheet) {
    Logger.log("統合データシートが見つかりません");
    return;
  }

  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let dateColIndex = headers.indexOf("集会の日付");
  if (dateColIndex === -1) {
    Logger.log("「集会の日付」列が見つかりません");
    return;
  }

  sheet.insertColumnAfter(1);
  sheet.getRange(1, 2).setValue("AssemblyID");
  dateColIndex += 1;

  const numRows = sheet.getLastRow();
  const dateValues = sheet.getRange(2, dateColIndex + 1, numRows - 1).getValues();
  const idValues = dateValues.map(([date]) => {
    if (!(date instanceof Date)) return [""];
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const dd = String(date.getDate()).padStart(2, "0");
    return [`ass${yyyy}${mm}${dd}`];
  });

  sheet.getRange(2, 2, idValues.length, 1).setValues(idValues);
  Logger.log("✅ AssemblyID の生成が完了しました。");
}

function insertMentionedAsSubsequent() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("統合データ");
  if (!sheet) {
    Logger.log("統合データシートが見つかりません");
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dateColIndex = headers.indexOf("集会の日付");
  const orderColIndex = headers.indexOf("トピック順");
  const topicIdColIndex = 0;

  if (dateColIndex === -1 || orderColIndex === -1) {
    Logger.log("「集会の日付」または「トピック順」列が見つかりません");
    return;
  }

  const insertColIndex = orderColIndex + 1;
  sheet.insertColumnBefore(insertColIndex + 1);
  sheet.getRange(1, insertColIndex + 1).setValue("mentionedAsSubsequent");

  const numRows = sheet.getLastRow();
  const topicIds = sheet.getRange(2, topicIdColIndex + 1, numRows - 1).getValues();
  const topicOrders = sheet.getRange(2, orderColIndex + 2, numRows - 1).getValues();

  const output = [];
  for (let i = 0; i < topicIds.length; i++) {
    const topicId = topicIds[i][0];
    const topicOrder = topicOrders[i][0];
    if (!topicId) {
      Logger.log(`A列が空のため、${i + 2}行目で処理を終了します。`);
      break;
    }
    if (topicOrder === 1 || topicOrder === "1") {
      output.push([""]);
    } else {
      const prevTopicId = topicIds[i - 1]?.[0] || "";
      output.push([prevTopicId]);
    }
  }

  sheet.getRange(2, insertColIndex + 1, output.length, 1).setValues(output);
  Logger.log("✅ mentionedAsSubsequent 列の生成が完了しました。");
}

function clearMentionedAsSubsequentIfFirstTopic() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("統合データ");
  if (!sheet) {
    Logger.log("統合データシートが見つかりません");
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const topicOrderColIndex = headers.indexOf("トピック順");
  const mentionedColIndex = headers.indexOf("mentionedAsSubsequent");
  const topicIdColIndex = 0;

  if (topicOrderColIndex === -1 || mentionedColIndex === -1) {
    Logger.log("「トピック順」または「mentionedAsSubsequent」列が見つかりません");
    return;
  }

  const numRows = sheet.getLastRow();
  const topicIds = sheet.getRange(2, topicIdColIndex + 1, numRows - 1).getValues();
  const topicOrders = sheet.getRange(2, topicOrderColIndex + 1, numRows - 1).getValues();
  const mentionedVals = sheet.getRange(2, mentionedColIndex + 1, numRows - 1).getValues();

  const updated = [];
  for (let i = 0; i < topicIds.length; i++) {
    const topicId = topicIds[i][0];
    const topicOrder = topicOrders[i][0];
    if (!topicId) {
      Logger.log(`A列が空のため、${i + 2}行目で処理を終了します。`);
      break;
    }
    if (topicOrder === 1 || topicOrder === "1") {
      updated.push([""]);
    } else {
      updated.push([mentionedVals[i][0]]);
    }
  }

  sheet.getRange(2, mentionedColIndex + 1, updated.length, 1).setValues(updated);
  Logger.log("✅ mentionedAsSubsequent 列の置換処理が完了しました。");
}

function insertScLabelPairColumn(anchorHeader, scHeader, labelHeader, suffix, labelValue) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("統合データ");
  if (!sheet) {
    Logger.log("統合データシートが見つかりません");
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const anchorColIndex = headers.indexOf(anchorHeader);
  if (anchorColIndex === -1) {
    Logger.log(`「${anchorHeader}」列が見つかりません`);
    return;
  }

  const insertCol = anchorColIndex + 1;
  sheet.insertColumnAfter(insertCol);
  sheet.insertColumnAfter(insertCol);

  sheet.getRange(1, insertCol + 1).setValue(scHeader);
  sheet.getRange(1, insertCol + 2).setValue(labelHeader);

  const numRows = sheet.getLastRow();
  const topicIds = sheet.getRange(2, 1, numRows - 1, 1).getValues();

  const scValues = [];
  const labelValues = [];

  for (let i = 0; i < topicIds.length; i++) {
    const topicId = topicIds[i][0];
    if (!topicId) {
      Logger.log(`A列が空のため、${i + 2}行目で処理を終了します。`);
      break;
    }
    scValues.push([`sc${topicId}${suffix}`]);
    labelValues.push([labelValue]);
  }

  sheet.getRange(2, insertCol + 1, scValues.length, 1).setValues(scValues);
  sheet.getRange(2, insertCol + 2, labelValues.length, 1).setValues(labelValues);

  Logger.log(`✅ ${scHeader} および ${labelHeader} 列の生成が完了しました。`);
}

function insertSourceAndSourceIdColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("統合データ");
  if (!sheet) {
    Logger.log("統合データシートが見つかりません");
    return;
  }

  const sourceSpreadsheetId = '189MGYsRkeTaEudD4CXibdhm9NfDLRZ7nYEGdwGcAfgA'; // これはいずれ統合して入力不要にしたい
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const gallicaColIndex = headers.indexOf("Gallica資料名");

  if (gallicaColIndex === -1) {
    Logger.log("「Gallica資料名」列が見つかりません");
    return;
  }

  const insertCol = gallicaColIndex + 1;
  sheet.insertColumnAfter(insertCol);
  sheet.insertColumnAfter(insertCol);
  sheet.getRange(1, insertCol + 1).setValue("source");
  sheet.getRange(1, insertCol + 2).setValue("sourceID");

  const numRows = sheet.getLastRow();
  const dateStr = sheet.getRange("C2").getValue();
  const origVolume = sheet.getRange("P2").getValue();

  let sourceyear = "0000";
  if (typeof dateStr === 'string' || dateStr instanceof String) {
    const match = dateStr.match(/^(\d{4})/);
    if (match) sourceyear = match[1];
  } else if (dateStr instanceof Date) {
    sourceyear = String(dateStr.getFullYear());
  }

  const sourcename = `pv${sourceyear}_${origVolume}`;
  const sourceValues = Array(numRows - 1).fill([sourcename]);
  const sourceIdValues = Array(numRows - 1).fill([sourceSpreadsheetId]);

  sheet.getRange(2, insertCol + 1, sourceValues.length, 1).setValues(sourceValues);
  sheet.getRange(2, insertCol + 2, sourceIdValues.length, 1).setValues(sourceIdValues);

  Logger.log("✅ source および sourceID 列の生成が完了しました。");
}
