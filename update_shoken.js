function recordDailyStock() {
  const targetSheets = ["時系列円", "時系列ドル", "時系列保有数"];  // ← 対象シートを複数指定
  //const targetSheets = ["時系列保有数"];  // ← 対象シートを複数指定

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  targetSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      console.log("Sheet not found:", sheetName);
      return;
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // --- 1. 次の行に数式をコピー ---
    const prevRowRange = sheet.getRange(lastRow, 1, 1, lastCol);     // 最新行
    const newRowRange = sheet.getRange(lastRow + 1, 1, 1, lastCol);  // 次の行

    // 数式ごとコピー
    prevRowRange.copyTo(newRowRange);

    // --- 2. 最新行（前の行）を値として固定 ---
    prevRowRange.copyTo(prevRowRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

    // --- 3. 表示形式だけを次の行 → 前の行にコピー ---
    newRowRange.copyTo(prevRowRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

    console.log("Processed:", sheetName);
    recordDailyStock2()
  });
}

function recordDailyStock2() {
  const targetSheets = ["時系列円", "時系列ドル", "時系列保有数"];  // ← 対象シートを複数指定
  // const targetSheets = ["時系列保有数"];  // ← 対象シートを複数指定

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  targetSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      console.log("Sheet not found:", sheetName);
      return;
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // --- 1. 次の行に数式をコピー ---
    const prevRowRange = sheet.getRange(lastRow - 1, 1, 1, lastCol);     // 最新行
    const newRowRange = sheet.getRange(lastRow, 1, 1, lastCol);  // 次の行

    // --- 3. 表示形式だけを次の行 → 前の行にコピー ---
    newRowRange.copyTo(prevRowRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

    console.log("Processed:", sheetName);
  });
}


function doGet(e) {
  const page = e.parameter.page;
  const sheetName = e.parameter.sheet;

  // --- holdingsForm ページを表示 ---
  if (page === "holdings") {
    const template = HtmlService.createTemplateFromFile('holdingsForm');
    return template.evaluate()
      .setTitle('保有数アップデートフォーム');
  }

  // --- sheet 表示ページ ---
  if (!sheetName) {
    return HtmlService.createHtmlOutput("Sheet name is required.");
  }

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return HtmlService.createHtmlOutput("Sheet not found.");
  }

  const range = (sheetName === "まとめ")
    ? sheet.getRange("A1:E11")
    : sheet.getDataRange();

  const values = range.getDisplayValues();
  const backgrounds = range.getBackgrounds();
  const fontColors = range.getFontColors();
  const numberFormats = range.getNumberFormats();

  const template = HtmlService.createTemplateFromFile("index");
  template.values = values;
  template.bg = backgrounds;
  template.fc = fontColors;
  template.nf = numberFormats;
  return template.evaluate().setTitle(sheetName);
}



// 現在の保有数を取得してフォームに渡す
function getCurrentHoldings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("銘柄一覧"); // 保有数が記録されているシート
  const data = sheet.getDataRange().getValues();

  // 配列を作成: { broker: "Charles Schwab", account: "一般口座", name: "SCHB", quantity: 3569.4332 }
  const holdings = data.slice(1).map(row => ({
    broker: row[0],    // 証券会社
    account: row[1],   // 口座名
    name: row[2],      // 銘柄名
    quantity: row[3]   // 現在の保有数
  }));

  return holdings;
}

// フォームから送信された保有数を更新
function updateHoldings(updates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("銘柄一覧");
  const data = sheet.getDataRange().getValues();

  updates.forEach(update => {
    // 証券会社・口座・銘柄名で一致する行を検索
    for (let i = 1; i < data.length; i++) {
      if (
        data[i][0] === update.broker &&
        data[i][1] === update.account &&
        data[i][2] === update.name
      ) {
        sheet.getRange(i + 1, 4).setValue(parseFloat(update.quantity)); // 保有数更新（4列目）
        break;
      }
    }
  });

  return "保有数が更新されました！";
}
