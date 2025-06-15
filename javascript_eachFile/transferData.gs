// シート1の編集を処理する関数
function handleSheet1Edit(range, sheetName1, sheetName3, sheetName5, sheetName6, sheetName7) {
  const sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName1);
  const sheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName3);
  const sheet5 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName5);
  const sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName6);
  const sheet7 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName7);
  const row = range.getRow();
  const mValue = sheet1.getRange(row, 13).getValue();

  if (mValue === true) {
    const dValue = sheet1.getRange(row, 4).getValue();

    if (["研究報告", "実験・提示", "査読依頼", "査読結果"].includes(dValue)) {
      const values = getSheet1Values(sheet1, row);
      const targetRow = findEmptyRow(sheet3, 1);
      sheet3.getRange(targetRow, 1, 1, 5).setValues([values]);
    } else if (dValue === "会員の任命") {
      const values = getSheet1ValuesForAppointment(sheet1, row);
      const targetRow = findEmptyRow(sheet6, 1);
      sheet6.getRange(targetRow, 1, 1, 5).setValues([values]);
    } else if (dValue === "投票・推薦") {
      const values = getSheet1ValuesForVoting(sheet1, row);
      const targetRow = findEmptyRow(sheet5, 1);
      sheet5.getRange(targetRow, 1, 1, 4).setValues([values]);
    } else if (dValue === "出席") {
      const values = getSheet1ValuesForAttendance(sheet1, row);
      const targetRow = findEmptyRow(sheet7, 1);
      sheet7.getRange(targetRow, 1, 1, 4).setValues([values]);
    }
  }
}

// シート3の編集を処理する関数
function handleSheet3Edit(range, sheetName3, sheetName4) {
  const sheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName3);
  const sheet4 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName4);
  const row = range.getRow();
  const nValue = sheet3.getRange(row, 14).getValue();

  if (nValue === true) {
    const hValue = sheet3.getRange(row, 8).getValue();

    if (hValue === true) {
      const values = getSheet3Values(sheet3, row);
      const targetRow = findEmptyRow(sheet4, 3);
      const id = generateFirstTopicId(sheet4, targetRow);
      sheet4.getRange(targetRow, 1).setValue(id);
      sheet4.getRange(targetRow, 2, 1, 5).setValues([values]);
    }
  }
}

// シート2の編集を処理する関数
function handleSheet2Edit(range, sheetName4, sheetName2) {
  const sheet4 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName4);
  const sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName2);
  const row = range.getRow();
  const oValue = sheet2.getRange(row, 15).getValue();

  if (oValue === true) {
    const fValue = sheet2.getRange(row, 6).getValue();

    if (fValue === "報告者") {
      const bValue = sheet2.getRange(row, 2).getValue();
      const gValue = sheet2.getRange(row, 7).getValue();
      const targetRow = findRowByValue(sheet4, 2, bValue, 3);
      if (targetRow) {
        sheet4.getRange(targetRow, 5).setValue(gValue);
      }
    }
  }
}

// シート1から必要な値を取得する関数（研究報告、実験・提示、査読依頼、査読結果用）
function getSheet1Values(sheet, row) {
  return [
    sheet.getRange(row, 1).getValue(),
    sheet.getRange(row, 2).getValue(),
    sheet.getRange(row, 7).getValue(),
    sheet.getRange(row, 9).getValue(),
    sheet.getRange(row, 4).getValue()
  ];
}

// シート1から必要な値を取得する関数（会員の任命用）
function getSheet1ValuesForAppointment(sheet, row) {
  return [
    sheet.getRange(row, 1).getValue(),
    sheet.getRange(row, 2).getValue(),
    sheet.getRange(row, 7).getValue(),
    sheet.getRange(row, 9).getValue(),
    sheet.getRange(row, 2).getValue() // B列の値を再度使用
  ];
}

// シート1から必要な値を取得する関数（投票・推薦、出席用）
function getSheet1ValuesForVoting(sheet, row) {
  return [
    sheet.getRange(row, 1).getValue(),
    sheet.getRange(row, 2).getValue(),
    sheet.getRange(row, 7).getValue(),
    sheet.getRange(row, 9).getValue()
  ];
}

// シート1から必要な値を取得する関数（出席用）
function getSheet1ValuesForAttendance(sheet, row) {
  return [
    sheet.getRange(row, 1).getValue(),
    sheet.getRange(row, 2).getValue(),
    sheet.getRange(row, 7).getValue(),
    sheet.getRange(row, 9).getValue()
  ];
}

// シート3から必要な値を取得する関数
function getSheet3Values(sheet, row) {
  return [
    sheet.getRange(row, 1).getValue(),
    sheet.getRange(row, 2).getValue(),
    sheet.getRange(row, 5).getValue(),
    sheet.getRange(row, 27).getValue(),
    sheet.getRange(row, 12).getValue()
  ];
}

// 指定した列が空の最初の行を見つける関数
function findEmptyRow(sheet, startRow) {
  const lastRow = sheet.getLastRow();
  for (let i = startRow; i <= lastRow; i++) {
    if (!sheet.getRange(i, 1).getValue()) {
      return i;
    }
  }
  return lastRow + 1;
}

// 指定した値を持つ行を見つける関数
function findRowByValue(sheet, column, value, startRow) {
  const lastRow = sheet.getLastRow();
  for (let i = startRow; i <= lastRow; i++) {
    if (sheet.getRange(i, column).getValue() === value) {
      return i;
    }
  }
  return null;
}

// 新しい初回報告IDを生成する関数
function generateFirstTopicId(sheet, row) {
  if (row > 3) {
    const previousId = sheet.getRange(row - 1, 1).getValue();
    return incrementId(previousId);
  }
  return "t1716-0000001";
}

// IDをインクリメントする関数
function incrementId(previousId) {
  const prefix = previousId.split('-')[0];
  const number = parseInt(previousId.split('-')[1], 10);
  const newNumber = number + 10;
  return `${prefix}-${('0000000' + newNumber).slice(-7)}`;
}
