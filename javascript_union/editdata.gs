function createPersonIntegratedSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("人物管理");
  const targetSheet = ss.getSheetByName("人物統合データ");

  if (!sourceSheet) {
    Logger.log("⚠️ 人物管理シートが見つかりません");
    return;
  }
  if (!targetSheet) {
    Logger.log("⚠️ 人物統合データシートが見つかりません");
    return;
  }

  // === ソースデータ取得（3行目以降）===
  const sourceData = sourceSheet.getDataRange().getValues().slice(2); // index 2 = 3行目
  if (sourceData.length === 0) {
    Logger.log("⚠️ 人物管理にコピー対象のデータがありません");
    return;
  }

  // === "人物統合データ" のA列を走査して、最初の空白行を取得 ===
  const targetAValues = targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1).getValues();
  let insertRow = targetAValues.findIndex(row => row[0] === "" || row[0] === null);

  if (insertRow === -1) {
    insertRow = targetAValues.length;
  }

  const startRow = insertRow + 2; // データは2行目から始まるので +2

  // === A列・B列のハイフン削除 ===
  for (let i = 0; i < sourceData.length; i++) {
    for (let j = 0; j < 2; j++) {
      if (typeof sourceData[i][j] === 'string') {
        sourceData[i][j] = sourceData[i][j].replace(/-/g, '');
      }
    }
  }

  // === データ貼り付け ===
  targetSheet.getRange(startRow, 1, sourceData.length, sourceData[0].length).setValues(sourceData);

  Logger.log(`✅ 人物統合データに ${sourceData.length} 行を ${startRow} 行目から追加しました（ヘッダー変更なし）。`);
}
