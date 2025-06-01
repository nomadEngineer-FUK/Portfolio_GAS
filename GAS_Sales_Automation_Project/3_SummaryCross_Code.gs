/**
 * メイン関数: クロス軸売上集計の全体を統括します。
 * すべてのクロス集計を一つのシートに出力します。
 */
function generateAllCrossAxisSummaries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetService = new SpreadsheetService(ss);
  const salesData = spreadsheetService.getDataRangeValues(CONFIG.SALES_SHEET_NAME);
  const outputSheet = spreadsheetService.getSheet(CONFIG.OUTPUT_SHEET_NAME_CROSS, true);

  let currentRow = 1; // 出力開始行を管理する変数

  // 例: カテゴリ別月別集計
  currentRow = _processGenericCrossSummary(
    spreadsheetService,
    salesData,
    outputSheet,
    "カテゴリ",  // 行軸 (rowHeaderName)
    "日付",     // 列軸 (colHeaderName)
    currentRow // 現在の出力開始行
  );

  // 例: 販売経路別月別集計 (新しいクロス集計を追加する場合)
  currentRow = _processGenericCrossSummary(
    spreadsheetService,
    salesData,
    outputSheet,
    "販売経路",  // 行軸
    "日付",     // 列軸
    currentRow // 現在の出力開始行
  );

  // カテゴリ×販売経路
  currentRow = _processGenericCrossSummary(
    spreadsheetService,
    salesData,
    outputSheet,
    "カテゴリ", // 行軸
    "販売経路", // 列軸
    currentRow // 現在の出力開始行
  );

  // 区分（法人／個人）×販売経路
  currentRow = _processCustomerTypeByChannelSummary(
    spreadsheetService,
    salesData,
    outputSheet,
    "名称", // 顧客区分を抽出する元の列
    "販売経路",
    currentRow // 現在の出力開始行
  );

  // 既存: 顧客区分別月別集計
  currentRow = _processCustomerTypeByMonthSummary(
    spreadsheetService,
    salesData,
    outputSheet,
    "名称", // 顧客区分を抽出する元の列
    "日付",
    currentRow // 現在の出力開始行
  );

  Logger.log("全てのクロス軸売上集計が完了しました。");
}


/**
 * 汎用的なクロス集計を処理し、指定されたシートに出力します。
 */
function _processGenericCrossSummary(spreadsheetService, salesData, outputSheet, rowHeaderName, colHeaderName, startRow) {
  Logger.log(`クロス集計を開始: ${rowHeaderName} vs ${colHeaderName} (行: ${startRow})`);
  // CrossTabSummary のコンストラクタに軸名を渡す
  const crossTabSummary = new CrossTabSummary(salesData, rowHeaderName, colHeaderName);
  const { headerRow, dataRows, totalRow } = crossTabSummary.getSummaryDataForOutput();

  // ヘッダーを書き込み (月が変わるとヘッダー列数が変わる可能性があるため、常に書き込み)
  outputSheet.getRange(startRow, 1, 1, headerRow.length).setValues([headerRow]);

  // データを出力
  const numDataRows = dataRows.length;
  const numCols = headerRow.length;
  spreadsheetService.clearAndWriteCrossDataBlock(outputSheet, dataRows, totalRow, numCols, startRow + 1); // ヘッダーの下から出力

  // 罫線と背景色の適用
  _applyCrossTabFormatting(outputSheet, numDataRows, numCols, spreadsheetService, startRow);

  return startRow + numDataRows + 3; // 次の集計の開始行を返す (ヘッダー + データ + 合計 + 空白行)
}


/**
 * 顧客区分別月別集計を処理し、指定されたシートに出力します。
 */
function _processCustomerTypeByMonthSummary(spreadsheetService, salesData, outputSheet, customerColName, colHeaderName, startRow) {
  Logger.log(`クロス集計を開始: 顧客区分 vs ${colHeaderName} (行: ${startRow})`);

  // 顧客区分（法人/個人）をキーとするための専用のCrossTabSummaryをインスタンス化
  const crossTabSummary = new CrossTabSummary(
    salesData,
    customerColName,
    colHeaderName,
    (customerValue) => {
      if (!customerValue || customerValue.toString().trim() === "") return "不明";
      return customerValue.toString().trim().startsWith("co") ? "法人" : "個人";
    },
    (dateValue) => {
      if (!dateValue) return "不明な日付";
      try {
        const date = new Date(dateValue);
        if (isNaN(date.getTime())) return "無効な日付";
        return `${date.getFullYear()}/${("0" + (date.getMonth() + 1)).slice(-2)}`;
      } catch (e) {
        return "日付エラー";
      }
    }
  );

  const { headerRow, dataRows, totalRow } = crossTabSummary.getSummaryDataForOutput();

  // ヘッダーを書き込み
  outputSheet.getRange(startRow, 1, 1, headerRow.length).setValues([headerRow]);

  // データを出力
  const numDataRows = dataRows.length;
  const numCols = headerRow.length;
  spreadsheetService.clearAndWriteCrossDataBlock(outputSheet, dataRows, totalRow, numCols, startRow + 1); // ヘッダーの下から出力

  // 罫線と背景色の適用
  _applyCrossTabFormatting(outputSheet, numDataRows, numCols, spreadsheetService, startRow);
  return startRow + numDataRows + 3; // 次の集計の開始行を返す
}

/**
 * 顧客区分別 販売経路別 集計を処理し、指定されたシートに出力します。
 * _processCustomerTypeByMonthSummary と同様に、顧客区分をキーに変換します。
 */
function _processCustomerTypeByChannelSummary(spreadsheetService, salesData, outputSheet, customerColName, colHeaderName, startRow) {
  Logger.log(`クロス集計を開始: 顧客区分 vs ${colHeaderName} (行: ${startRow})`);

  // 顧客区分（法人/個人）をキーとするための専用のCrossTabSummaryをインスタンス化
  const crossTabSummary = new CrossTabSummary(
    salesData,
    customerColName,
    colHeaderName,
    (customerValue) => {
      if (!customerValue || customerValue.toString().trim() === "") return "不明";
      return customerValue.toString().trim().startsWith("co") ? "法人" : "個人";
    },
    // 販売経路はそのまま使用するため、colKeyTransformer はデフォルト（trim）でOK
    // または明示的に (val) => String(val).trim() を指定しても良い
  );

  const { headerRow, dataRows, totalRow } = crossTabSummary.getSummaryDataForOutput();

  // ヘッダーを書き込み
  outputSheet.getRange(startRow, 1, 1, headerRow.length).setValues([headerRow]);

  // データを出力
  const numDataRows = dataRows.length;
  const numCols = headerRow.length;
  spreadsheetService.clearAndWriteCrossDataBlock(outputSheet, dataRows, totalRow, numCols, startRow + 1); // ヘッダーの下から出力

  // 罫線と背景色の適用
  _applyCrossTabFormatting(outputSheet, numDataRows, numCols, spreadsheetService, startRow);
  return startRow + numDataRows + 3; // 次の集計の開始行を返す
}


/**
 * クロス集計シートの罫線と背景色を適用します。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} outputSheet - 対象シート
 * @param {number} numDataRows - データ行数 (合計行は含まない)
 * @param {number} numCols - 列数
 * @param {SpreadsheetService} spreadsheetService - SpreadsheetServiceインスタンス
 * @param {number} startRow - この集計ブロックの開始行 (ヘッダー行の行番号)
 */
function _applyCrossTabFormatting(outputSheet, numDataRows, numCols, spreadsheetService, startRow) {
  // 全体の罫線 (ヘッダー含む全データ + 合計行)
  const totalRows = numDataRows + 2; // ヘッダー1行 + データ行数 + 合計行1行
  spreadsheetService.applyBorders(outputSheet, startRow, 1, totalRows, numCols);

  // 背景色 (ヘッダー行全体)
  // startRow 行目から、1列目から numCols 列目まで、1行分に背景色を適用
  spreadsheetService.applyBackgroundColor(outputSheet, startRow, 1, 1, numCols, CONFIG.COLORS.LIGHT_GRAY);

  // 背景色 (行軸のカテゴリ列)
  // データがある場合のみ、ヘッダーの次の行からデータ終了行まで、1列目（行軸の列）に背景色
  if (numDataRows > 0) {
    // 罫線と背景色の適用: applyBackgroundColor(sheet, startRow, startCol, numRows, numCols, color)
    // この行はA列のみに色を塗るべきなので、numColsは1
    spreadsheetService.applyBackgroundColor(outputSheet, startRow + 1, 1, numDataRows, 1, CONFIG.COLORS.LIGHT_GRAY);
  }
  
  // 合計行の背景色 (A列から合計列まで)
  spreadsheetService.applyBackgroundColor(outputSheet, startRow + numDataRows + 1, 1, 1, numCols, CONFIG.COLORS.LIGHT_GRAY);
}
