// 統合データの再処理（列の追加は行わない）
// 2年目以降はこちらを実行
// updateSourceAndSourceIdValues()のsourceSpreadSheetIdは毎回修正の上使うこと
function adjustTopicDataProcessing() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("統合データ");
  if (!sheet) {
    Logger.log("統合データシートが見つかりません");
    return;
  }

  // ✅ D列（4列目）を整数表示に設定（2行目以降）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const cRange = sheet.getRange(2, 4, lastRow - 1); // D列: col = 4
    cRange.setNumberFormat("0");
    Logger.log("✅ D列の表示形式を整数に設定しました。");
  }

  Logger.log("=== 統合データ 調整処理開始 ===");

  generateAssemblyID_Adjust(); // 再生成（列追加なし）
  insertMentionedAsSubsequent_Adjust(); // 再生成（列追加なし）
  clearMentionedAsSubsequentIfFirstTopic_Adjust(); //再生成（列追加なし）

  updateScLabelValues_Adjust("トピック種別", "sc_recogito", "recogito", "Recogito");
  updateScLabelValues_Adjust("付与チェック", "sc_gallica", "gallica", "gallica");
  updateScLabelValues_Adjust("資料URI", "sc_iiif", "iiif", "iiif");
  updateScLabelValues_Adjust("IIIF manifest URI", "sc_orig", "original", "original");

  // updateSourceAndSourceIdValues();
  updateSourceAndSourceIdValues_Adjust

  Logger.log("=== 統合データ 調整処理完了 ===");
}

// 以下、補助関数

function generateAssemblyID_Adjust() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("統合データ");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dateColIndex = headers.indexOf("集会の日付");
  const assemblyColIndex = headers.indexOf("AssemblyID");

  if (dateColIndex === -1 || assemblyColIndex === -1) {
    Logger.log("「集会の日付」または「AssemblyID」列が見つかりません");
    return;
  }

  const numRows = sheet.getLastRow();
  const dateValues = sheet.getRange(2, dateColIndex + 1, numRows - 1).getValues();

  const idValues = dateValues.map(([date]) => {
    if (!(date instanceof Date)) return [""];
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const dd = String(date.getDate()).padStart(2, "0");
    return [`ass${yyyy}${mm}${dd}`];
  });

  sheet.getRange(2, assemblyColIndex + 1, idValues.length, 1).setValues(idValues);
  Logger.log("✅ AssemblyID を再生成しました。");
}


function insertMentionedAsSubsequent_Adjust() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("統合データ");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const topicIdColIndex = 0; // A列: トピックID
  const secondColIndex = 1;  // B列: 制御終了判定
  const orderColIndex = headers.indexOf("トピック順");
  const mentionedColIndex = headers.indexOf("mentionedAsSubsequent");

  if (orderColIndex === -1 || mentionedColIndex === -1) {
    Logger.log("「トピック順」または「mentionedAsSubsequent」列が見つかりません");
    return;
  }

  const numRows = sheet.getLastRow();
  const allValues = sheet.getRange(2, 1, numRows - 1, Math.max(topicIdColIndex, secondColIndex, orderColIndex, mentionedColIndex) + 1).getValues();

  const output = [];

  for (let i = 0; i < allValues.length; i++) {
    const row = allValues[i];
    const topicId = row[topicIdColIndex];
    const controlVal = row[secondColIndex];
    const topicOrder = row[orderColIndex];

    if (controlVal === "" || controlVal === null) break; // B列が空であれば終了

    if (topicOrder === "" || topicOrder === null) {
      output.push([""]); // トピック順が空欄の行はスキップ
      continue;
    }

    if (topicOrder === 1 || topicOrder === "1") {
      output.push([""]);
    } else {
      const prevTopicId = i > 0 ? allValues[i - 1][topicIdColIndex] || "" : "";
      output.push([prevTopicId]);
    }
  }

  // 出力長に応じて書き込み
  if (output.length > 0) {
    sheet.getRange(2, mentionedColIndex + 1, output.length, 1).setValues(output);
    Logger.log("✅ mentionedAsSubsequent を再生成しました（B列の空白行まで）。");
  } else {
    Logger.log("⚠️ 有効なデータ行がなかったため、再生成をスキップしました。");
  }
}

function clearMentionedAsSubsequentIfFirstTopic_Adjust() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("統合データ");
  if (!sheet) {
    Logger.log("統合データシートが見つかりません");
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const topicOrderColIndex = headers.indexOf("トピック順");
  const mentionedColIndex = headers.indexOf("mentionedAsSubsequent");
  const topicIdColIndex = 0; // A列
  const controlColIndex = 1; // B列（空欄になるまで繰り返す）

  if (topicOrderColIndex === -1 || mentionedColIndex === -1) {
    Logger.log("「トピック順」または「mentionedAsSubsequent」列が見つかりません");
    return;
  }

  const numRows = sheet.getLastRow();
  const allValues = sheet.getRange(2, 1, numRows - 1, Math.max(topicOrderColIndex, mentionedColIndex, controlColIndex) + 1).getValues();

  const updated = [];

  for (let i = 0; i < allValues.length; i++) {
    const row = allValues[i];
    const topicId = row[topicIdColIndex];
    const controlVal = row[controlColIndex];
    const topicOrder = row[topicOrderColIndex];
    const currentMentioned = row[mentionedColIndex];

    if (controlVal === "" || controlVal === null) break; // B列が空になったら終了
    if (!topicId) {
      updated.push([currentMentioned]); // A列が空なら何もしない（スキップ）
      continue;
    }

    if (topicOrder === 1 || topicOrder === "1") {
      updated.push([""]);
    } else {
      updated.push([currentMentioned]);
    }
  }

  if (updated.length > 0) {
    sheet.getRange(2, mentionedColIndex + 1, updated.length, 1).setValues(updated);
    Logger.log("✅ mentionedAsSubsequent の 'トピック順 = 1' 行のクリア処理が完了しました（B列の空白まで）。");
  } else {
    Logger.log("⚠️ 書き換え対象の行が見つかりませんでした。");
  }
}


function updateScLabelValues_Adjust(anchorHeader, scHeader, suffix, labelValue) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("統合データ");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const topicIdColIndex = 0; // A列
  const controlColIndex = 1; // B列（終了判定）

  const scIndex = headers.indexOf(scHeader);
  const labelIndex = headers.indexOf(suffix);

  if (scIndex === -1 || labelIndex === -1) {
    Logger.log(`列 ${scHeader} または ${suffix} が見つかりません`);
    return;
  }

  const numRows = sheet.getLastRow();
  const allValues = sheet.getRange(2, 1, numRows - 1, Math.max(scIndex, labelIndex, controlColIndex) + 1).getValues();

  const scValues = [];
  const labelValues = [];

  for (let i = 0; i < allValues.length; i++) {
    const row = allValues[i];
    const topicId = row[topicIdColIndex];
    const controlVal = row[controlColIndex];

    if (controlVal === "" || controlVal === null) break;
    if (!topicId) {
      scValues.push([""]);
      labelValues.push([""]);
      continue;
    }

    scValues.push([`sc${topicId}${suffix}`]);
    labelValues.push([labelValue]);
  }

  if (scValues.length > 0) {
    sheet.getRange(2, scIndex + 1, scValues.length, 1).setValues(scValues);
    sheet.getRange(2, labelIndex + 1, labelValues.length, 1).setValues(labelValues);
    Logger.log(`✅ ${scHeader} / ${suffix} 列の更新が完了しました（${scValues.length}行処理）。`);
  } else {
    Logger.log(`⚠️ 有効な行が見つからなかったため、${scHeader} / ${suffix} の更新をスキップしました。`);
  }
}


function updateSourceAndSourceIdValues_Adjust() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("統合データ");
  const sourceSpreadsheetId = "転送元スプレッドシートのID"; // 修正の上使用する

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const sourceCol = headers.indexOf("source");
  const sourceIdCol = headers.indexOf("sourceID");

  if (sourceCol === -1 || sourceIdCol === -1) {
    Logger.log("❌ 'source' または 'sourceID' 列が見つかりません");
    return;
  }

  const numRows = sheet.getLastRow();
  const allValues = sheet.getRange(2, 1, numRows - 1, sheet.getLastColumn()).getValues();

  let startRow = -1;
  for (let i = 0; i < allValues.length; i++) {
    const qVal = allValues[i][16]; // Q列（17列目, index 16）
    if (qVal === "" || qVal === null) {
      startRow = i + 2; // シート上の行番号（1-based）
      break;
    }
  }

  if (startRow === -1) {
    Logger.log("⚠️ Q列が空の行が見つかりませんでした。処理をスキップします。");
    return;
  }

  // C列（3列目, index 2）、P列（16列目, index 15）を取得
  const dateStr = sheet.getRange(startRow, 3).getValue();  // C列
  const origVolume = sheet.getRange(startRow, 16).getValue(); // P列

  let sourceyear = "0000";
  if (typeof dateStr === "string" || dateStr instanceof String) {
    const match = dateStr.match(/^(\d{4})/);
    if (match) sourceyear = match[1];
  } else if (dateStr instanceof Date) {
    sourceyear = String(dateStr.getFullYear());
  }

  const sourcename = `pv${sourceyear}_${origVolume}`;

  const sourceRange = sheet.getRange(startRow, sourceCol + 1, numRows - startRow + 1, 1);
  const sourceIdRange = sheet.getRange(startRow, sourceIdCol + 1, numRows - startRow + 1, 1);

  const rowCount = numRows - startRow + 1;
  const sourceValues = Array(rowCount).fill([sourcename]);
  const sourceIdValues = Array(rowCount).fill([sourceSpreadsheetId]);

  sourceRange.setValues(sourceValues);
  sourceIdRange.setValues(sourceIdValues);

  Logger.log(`✅ source / sourceID の再設定完了 (${startRow}行目から下 ${rowCount}行に適用)`);
}