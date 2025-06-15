// この関数だけ実行する
function mergeAllTopicData() {
  Logger.log("=== トピック統合処理 開始 ===");

  mergeTopicIdManagementSlowly();
  mergeSheetsByTopicIdWithLogs();
  mergeInitialReportLinks();

  Logger.log("=== トピック統合処理 完了 ===");
}


// 以下は処理材料のため個別の実行は不要
function mergeTopicIdManagementSlowly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("トピックID全体管理");
  const targetSheet = ss.getSheetByName("統合データ");

  if (!sourceSheet || !targetSheet) {
    Logger.log("必要なシートが見つかりません");
    return;
  }

  // 統合データのヘッダー行を取得し、列名→インデックスの辞書を作成
  const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  const targetHeaderIndex = {};
  targetHeaders.forEach((header, i) => {
    targetHeaderIndex[header] = i;
  });

  // トピック全体管理のデータ全体取得
  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[0];

  // 3行目以降を1行ずつ処理
  for (let i = 2; i < sourceData.length; i++) {
    const row = sourceData[i];

    // A列（トピックID）が空なら終了
    if (!row[0]) {
      Logger.log(`A列が空のため、${i + 1}行目で処理を終了します。`);
      break;
    }

    const newRow = new Array(targetHeaders.length).fill("");
    row.forEach((cell, colIndex) => {
      const header = sourceHeaders[colIndex];
      if (header in targetHeaderIndex) {
        const targetCol = targetHeaderIndex[header];
        newRow[targetCol] = cell;
      }
    });

    // 統合先の最終行の次に1行書き込み
    const lastRow = targetSheet.getLastRow();
    targetSheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);

    Utilities.sleep(200); // 処理の間隔を空ける
  }

  Logger.log("mergeTopicIdManagementSlowly() の実行が完了しました。");
}

function mergeSheetsByTopicIdWithLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName("統合データ");

  if (!targetSheet) {
    Logger.log("統合データシートが見つかりません");
    return;
  }

  const targetData = targetSheet.getDataRange().getValues();
  const targetHeaders = targetData[0];

  // 統合データのヘッダー → 列番号（0始まり）
  const targetHeaderIndex = {};
  targetHeaders.forEach((h, i) => targetHeaderIndex[h.trim()] = i);

  // トピックID → 行番号（1始まり）
  const topicIdToRow = {};
  for (let i = 2; i < targetData.length; i++) {
    const topicId = targetData[i][0];
    if (topicId) topicIdToRow[topicId] = i + 1;
  }

  // 処理対象のシート一覧
  const sheetNames = [
    "研究報告・実験・査読",
    "投票・推薦",
    "会員の任命"
  ];

  Logger.log("実行開始");

  sheetNames.forEach(sheetName => {
    const sourceSheet = ss.getSheetByName(sheetName);
    if (!sourceSheet) {
      Logger.log(`⚠️ シート '${sheetName}' が見つかりません。スキップします。`);
      return;
    }

    const sourceData = sourceSheet.getDataRange().getValues();
    const sourceHeaders = sourceData[0].map(h => h.trim());

    let rowCount = 0;

    for (let i = 2; i < sourceData.length; i++) {
      const row = sourceData[i];
      const topicId = row[0];

      if (!topicId) {
        Logger.log(`ℹ️ シート「${sheetName}」の ${i + 1} 行目で A列が空のため処理を終了しました。処理行数: ${rowCount}`);
        break;
      }

      const targetRowNum = topicIdToRow[topicId];
      if (!targetRowNum) {
        Logger.log(`⚠️ トピックID「${topicId}」（シート: ${sheetName}）が統合データに見つかりませんでした。スキップします。`);
        continue;
      }

      const updates = [];
      const updateColumns = [];

      row.forEach((cell, colIndex) => {
        const header = sourceHeaders[colIndex];
        if (header in targetHeaderIndex && colIndex !== 0) {
          const targetCol = targetHeaderIndex[header];
          updates[targetCol] = cell;
          updateColumns.push(targetCol);
        }
      });

      updateColumns.forEach(colIndex => {
        targetSheet.getRange(targetRowNum, colIndex + 1).setValue(updates[colIndex]);
      });

      rowCount++;
      Utilities.sleep(100);
    }

    Logger.log(`✅ シート「${sheetName}」の統合が完了しました（処理行数: ${rowCount}）。`);
  });

  Logger.log("✅ すべてのシートの統合処理が完了しました。");
}

function mergeInitialReportLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("初出研究報告管理");
  const targetSheet = ss.getSheetByName("統合データ");

  if (!sourceSheet || !targetSheet) {
    Logger.log("必要なシートが見つかりません");
    return;
  }

  // 統合データのヘッダー取得と列マッピング
  const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  const targetHeaderIndex = {};
  targetHeaders.forEach((h, i) => targetHeaderIndex[h.trim()] = i);

  // 「初出のトピックID」列のインデックスを取得
  const outputColIndex = targetHeaderIndex["初出報告ID"];
  if (outputColIndex === undefined) {
    Logger.log("統合データに『初出報告ID』という列が見つかりません。処理を中止します。");
    return;
  }

  // トピックID → 行番号のマッピングを構築（"統合データ"）
  const targetData = targetSheet.getDataRange().getValues();
  const topicIdToRow = {};
  for (let i = 2; i < targetData.length; i++) {
    const topicId = targetData[i][0]; // A列（トピックID）
    if (topicId) {
      topicIdToRow[topicId] = i + 1; // 1-based
    }
  }

  const sourceData = sourceSheet.getDataRange().getValues();

  let count = 0;
  for (let i = 2; i < sourceData.length; i++) {
    const row = sourceData[i];
    const initialReportId = row[0];   // A列（初出報告ID）
    const linkedTopicId = row[1];     // B列（初出トピックID）

    if (!initialReportId) {
      Logger.log(`ℹ️ シート「初出研究報告管理」の ${i + 1} 行目で A列が空のため処理を終了しました。処理件数: ${count}`);
      break;
    }

    if (!linkedTopicId || !(linkedTopicId in topicIdToRow)) {
      Logger.log(`⚠️ トピックID「${linkedTopicId}」が統合データに見つかりません（初出報告ID: ${initialReportId}）。スキップします。`);
      continue;
    }

    const targetRowNum = topicIdToRow[linkedTopicId];
    targetSheet.getRange(targetRowNum, outputColIndex + 1).setValue(initialReportId);
    count++;

    Utilities.sleep(100); // 念のため遅延を挿入
  }

  Logger.log(`✅ 「初出研究報告管理」シートの統合作業が完了しました（転送数: ${count}）。`);
}


