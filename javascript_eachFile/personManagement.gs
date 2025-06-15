function handlePersonManagement(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  const column = range.getColumn();
  
  const topicSheet = e.source.getSheetByName("トピックID全体管理");
  const personSheet = e.source.getSheetByName("人物管理");
  const initialResearchReportSheet = e.source.getSheetByName("初出研究報告管理");
  
  // B列の3行目以降の変更を処理
  if (sheet.getName() === "人物管理" && column === 2 && row >= 3) {
    const searchValue = range.getValue();  // 入力された値

    // トピックID全体管理シートで検索し、対応する値を人物管理シートに挿入
    searchAndInsertValue(topicSheet, "A", searchValue, "G", row, "C", personSheet);
    searchAndInsertValue(topicSheet, "A", searchValue, "I", row, "D", personSheet);
    searchAndInsertValue(topicSheet, "A", searchValue, "B", row, "E", personSheet);
  }
  
  // G列の値が変更されたときにIDを生成
  if (sheet.getName() === "人物管理" && column === 7 && row >= 3) {
    const dateValue = personSheet.getRange(`E${row}`).getValue();
    const year = new Date(dateValue).getFullYear();
    const generatedID = generatePersonID(year, row);
    personSheet.getRange(`A${row}`).setValue(generatedID);
  }

  // F列の値が変更されたときの処理
  if (sheet.getName() === "人物管理" && column === 6 && row >= 3) {
    const value = range.getValue();
    if (value === "被任命者") {
      ["I", "J", "K"].forEach(col => enableCell(sheet, `${col}${row}`));
      ["H", "L", "M"].forEach(col => disableCell(sheet, `${col}${row}`));
      setDropdownLists(sheet, row);
    } else if (value === "候補者") {
      ["I", "J", "K", "L", "M"].forEach(col => enableCell(sheet, `${col}${row}`));
      ["H"].forEach(col => disableCell(sheet, `${col}${row}`));
      setDropdownLists(sheet, row);
    } else if (value === "除名・退任者") {
      ["H", "I", "J", "K"].forEach(col => enableCell(sheet, `${col}${row}`));
      ["L", "M"].forEach(col => disableCell(sheet, `${col}${row}`));
      setDropdownLists(sheet, row);
    } else {
      ["H", "I", "J", "K", "L", "M"].forEach(col => disableCell(sheet, `${col}${row}`));
    }
  }

  // ここに、I列の値によってK列の値を調整する処理を入れたい（保留）
}

// ここから個別関数

function searchAndInsertValue(sheet, searchColumn, searchValue, targetColumn, resultRow, resultColumn, resultSheet) {
  const data = sheet.getRange(`${searchColumn}:${searchColumn}`).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === searchValue) {
      const sourceRow = i + 1; // 行番号は1から始まるため、インデックスに1を足す
      const sourceValue = sheet.getRange(`${targetColumn}${sourceRow}`).getValue();
      resultSheet.getRange(`${resultColumn}${resultRow}`).setValue(sourceValue);
      break;
    }
  }
}

// A列に人物管理IDを生成する
function generatePersonID(year, row) {
  const idNumber = row * 10;
  return `p${year}-${String(idNumber).padStart(7, '0')}`;
}

// I列,J列,K列にプルダウンリストを設定する
function setDropdownLists(sheet, row) {
  const positionSheetName = "職階・役職";
  const fieldSheetName = "分野";  
  const positionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(positionSheetName);
  const fieldSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(fieldSheetName);
    
  const positions = positionSheet.getRange("A:A").getValues().flat().filter(String);  // 空の値を除去
  const fields = fieldSheet.getRange("A:A").getValues().flat().filter(String);  // 空の値を除去
    
  // Logger.log(`Positions for dropdown: ${positions}`);
  // Logger.log(`Fields for dropdown: ${fields}`);

  if (positions.length > 0) {
    ["I", "J"].forEach(col => setDropdownList(sheet, `${col}${row}`, positions));
  } else {
    Logger.log('No positions found in the position sheet.');
  }

  if (fields.length > 0) {
    setDropdownList(sheet, `K${row}`, fields);
  } else {
    Logger.log('No fields found in the field sheet.');
  }
}

// 初出研究報告管理シートの更新処理
function updateInitialResearchReport(sheet, searchValue, newValue) {
  const data = sheet.getRange("B:B").getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === searchValue) {
      const targetRow = i + 1; // 行番号は1から始まるため、インデックスに1を足す
      sheet.getRange(`E${targetRow}`).setValue(newValue);
      break;
    }
  }
}