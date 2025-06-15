// 初年のみ

function collectUniqueHeaders() {
  const sheetNames = [
    "トピックID全体管理",
    "研究報告・実験・査読",
    "初出研究報告管理",
    "投票・推薦",
    "会員の任命"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headerSet = new Set();

  // 各シートの1行目からヘッダーを収集（重複を除外）
  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      Logger.log(`シート '${name}' が見つかりませんでした`);
      return;
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    headers.forEach(h => {
      if (h !== "") headerSet.add(h); // 空白セルは除く
    });
  });

  // Setから配列に変換し、ソートも可能
  const uniqueHeaders = Array.from(headerSet);

  // 統合先シートに書き込み
  let targetSheet = ss.getSheetByName("統合データ");
  if (!targetSheet) {
    targetSheet = ss.insertSheet("統合データ");
  } else {
    targetSheet.clear(); // 既存内容を削除
  }

  targetSheet.getRange(1, 1, 1, uniqueHeaders.length).setValues([uniqueHeaders]);

  Logger.log("統合済みヘッダーの書き込みが完了しました。");
}
