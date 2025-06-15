function handleResearchManagement(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const column = range.getColumn();
  const row = range.getRow();
  const sheetName2 = "人物管理";
  const sheetName3 = "研究報告・実験・査読";
  const sheetName4 = "初出研究報告管理";

  // アクティブなスプレッドシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet3 = ss.getSheetByName(sheetName3);
  const sheet2 = ss.getSheetByName(sheetName2);
  const sheet4 = ss.getSheetByName(sheetName4);

  // 現在のシートがsheetName3でない場合は処理を終了
  if (!sheet3 || sheet.getName() !== sheetName3) {
    return;
  }

  // F列に新しい値が入力されたときのn-1行目の処理
  if (column === 6 && row > 3) {
    handlePreviousRow(sheet3, sheet2, row - 1);
  }

  // F列で"本人報告"が選択されたらG列をdisableCell、そうでなければenableCell
  if (column === 6) { // F列
    const value = range.getValue();
    if (value === "本人報告") {
      disableCell(sheet3, `G${row}`);
    } else {
      enableCell(sheet3, `G${row}`);
    }
  }

  // H列でチェックボックスにチェックが入ったらI列をdisableCell、そうでなければenableCell
  if (column === 8) { // H列
    const value = range.getValue();
    if (value === true) {
      disableCell(sheet3, `I${row}`);
    } else {
      enableCell(sheet3, `I${row}`);
    }
  }

}

function generateNewID(row) {
  const year = 1716;
  const idNumber = row * 10;
  return `t${year}-${String(idNumber).padStart(7, '0')}`;
}

function handlePreviousRow(sheet3, sheet2, row) {
  const aValue = sheet3.getRange(`A${row}`).getValue();
  const data = sheet2.getRange("B:G").getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === aValue) { // B列に一致する値を見つける
      if (data[i][4] === "報告者") { // F列の値が"報告者"
        const gValue = data[i][5]; // G列の値を取得
        sheet3.getRange(`J${row}`).setValue(gValue); // J列に値を設定
        break;
      }
    }
  }
}

function updateKColumnBasedOnIValue(sheet3, sheet4, row, value) {
  const data = sheet4.getRange("A:B").getValues();
  Logger.log(`Searching for value: ${value}`); // ログを追加して検索値を確認
  for (let i = 0; i < data.length; i++) {
    Logger.log(`Row ${i + 1} - A: ${data[i][0]}, B: ${data[i][1]}`); // ログを追加して各行のデータを確認
    if (data[i][1] === value) { // B列に一致する値を見つける
      const aValue = data[i][0]; // A列の値を取得
      Logger.log(`Found matching value in row ${i + 1}, A: ${aValue}`); // ログを追加して一致する行を確認
      sheet3.getRange(`K${row}`).setValue(aValue); // K列に値を設定
      break;
    }
  }
}

function disableCell(sheet, targetRange) {
  const range = sheet.getRange(targetRange);
  range.setBackground('lightgray');
  range.setValue(range.getValue()); // セルの内容を保持
  
  // 入力規則を設定して入力不可にする
  const rule = SpreadsheetApp.newDataValidation()
    .requireTextDoesNotContain('') // 空文字列が含まれない場合にエラー
    .setHelpText('このセルは入力不可です')
    .build();
  range.setDataValidation(rule);
}

function enableCell(sheet, targetRange) {
  const range = sheet.getRange(targetRange);
  range.setBackground(null); // 背景色をクリアにする
  
  // 入力規則を解除して入力可能にする
  range.setDataValidation(null);
}
