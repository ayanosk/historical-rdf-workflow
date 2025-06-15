function onEdit(e) {
  e.source.getSheetByName("ログ").appendRow([new Date()]);

  const range = e.range;  // 編集された範囲
  const sheet = range.getSheet();
  const column = range.getColumn();  // 編集された列番号
  const row = range.getRow();  // 編集された行番号
  const sheetName1 = "トピックID全体管理"; 
  const sheetName2 = "人物管理";  
  const sheetName3 = "研究報告・実験・査読"; 
  const sheetName4 = "初出研究報告管理";  
  const sheetName5 = "投票・推薦";
  const sheetName6 = "会員の任命";
  const sheetName7 = "出席";

  // トピックID全体管理シートを処理
  if (sheet.getName() === sheetName1) {
    handleTopicManagement(e);
    if (column === 13) { // M列の編集をチェック
      handleSheet1Edit(range, sheetName1, sheetName3, sheetName5, sheetName6, sheetName7);
    }
  }
  
  // 人物管理シートを処理
  if (sheet.getName() === sheetName2) {
    handlePersonManagement(e);
    if (column === 15) { // O列の編集をチェック
      handleSheet2Edit(range, sheetName4, sheetName2);
    }
  }
  
  // 研究報告・実験・査読シートを処理
  if (sheet.getName() === sheetName3) {
    handleResearchManagement(e);
    if (column === 14) { // N列の編集をチェック
      handleSheet3Edit(range, sheetName3, sheetName4);
    }
  }
}

// 任意のセルの内容を特定のセルにコピーする関数
function copyValue(sheet, sourceRange, targetRange) {
  const value = sheet.getRange(sourceRange).getValue(); // 入力セルの値を取得
  sheet.getRange(targetRange).setValue(value); // 出力セルに値を設定
}

// 指定したセルを入力不可にし、背景色をlightgrayに変更する関数（デフォルト）
function disableCell(sheet, targetRange) {
  const range = sheet.getRange(targetRange);
  range.setBackground('lightgray');
  range.setValue(range.getValue()); // セルの内容を保持
  
  // 入力規則を設定して入力不可にする
  const rule = SpreadsheetApp.newDataValidation()
    .requireTextDoesNotContain('')  // 空文字列が含まれない場合にエラー
    .setHelpText('このセルは入力不可です')
    .build();
  range.setDataValidation(rule);
}

// 指定したセルを入力可能にし、背景色をクリアにする関数
function enableCell(sheet, targetRange) {
  const range = sheet.getRange(targetRange);
  range.setBackground(null); // 背景色をクリアにする
  
  // 入力規則を解除して入力可能にする
  range.setDataValidation(null);
}

// 指定したセルの値を基に列を検索し、マッチする行の指定した列に値を挿入する関数
function searchAndInsertValue(sheet, searchColumn, searchValue, targetColumn, resultRow, resultColumn) {
  const data = sheet.getRange(`${searchColumn}:${searchColumn}`).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === searchValue) {
      const sourceRow = i + 1; // 行番号は1から始まるため、インデックスに1を足す
      const sourceValue = sheet.getRange(`${targetColumn}${sourceRow}`).getValue();
      sheet.getRange(`${resultColumn}${resultRow}`).setValue(sourceValue);
      break;
    }
  }
}

function setDropdownList(sheet, range, options) {
  const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options)
      .setAllowInvalid(false)
      .build();
  sheet.getRange(range).setDataValidation(rule);
}
