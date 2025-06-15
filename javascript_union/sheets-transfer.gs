function copySheetsFromSource() {
  const sourceSpreadsheetId = '転送元スプレッドシートのID'; // 転送元スプレッドシートのIDを入れる
  const sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);

  // 転送先（このスクリプトが紐づくスプレッドシート）
  const destinationSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // コピー対象のシート名一覧
  const sheetNames = [
    "トピックID全体管理",
    "人物管理",
    "研究報告・実験・査読",
    "初出研究報告管理",
    "投票・推薦",
    "会員の任命"
  ];

  // 各シートを複製
  sheetNames.forEach(name => {
    const sourceSheet = sourceSpreadsheet.getSheetByName(name);
    if (!sourceSheet) {
      Logger.log(`シート '${name}' は見つかりませんでした`);
      return;
    }

    // すでに同名のシートが存在する場合は削除（必要に応じてコメントアウト可能）
    const existingSheet = destinationSpreadsheet.getSheetByName(name);
    if (existingSheet) {
      destinationSpreadsheet.deleteSheet(existingSheet);
    }

    // シートを複製し、名前を元と同じに設定
    const copiedSheet = sourceSheet.copyTo(destinationSpreadsheet);
    copiedSheet.setName(name);
  });

  Logger.log('シートのコピーが完了しました。');
}
