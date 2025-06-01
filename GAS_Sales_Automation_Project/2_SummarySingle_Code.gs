/**
 * メイン関数: 単軸売上集計の全体を統括します。
 */
function generateSingleAxisSummaryRefactored() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetService = new SpreadsheetService(ss);
  const outputSheet = spreadsheetService.getSheet(CONFIG.OUTPUT_SHEET_NAME_SINGLE, true);

  // 1. ヘッダーを初期設定（シートが新規の場合や、ヘッダーが消えた場合に備える）
  _initializeHeaders(outputSheet);

  // 2. 売上データを取得し、集計クラスを初期化
  const salesData = spreadsheetService.getDataRangeValues(CONFIG.SALES_SHEET_NAME);
  const salesSummary = new SalesSummary(salesData);

  // 3. 各集計ブロックを処理
  let currentColumn = 1;

  // 月別集計
  currentColumn = _processSummaryBlock(
    outputSheet, spreadsheetService, salesSummary.summarizeByMonth(),
    CONFIG.HEADERS.MONTHLY, currentColumn
  );

  // カテゴリ別集計
  currentColumn = _processSummaryBlock(
    outputSheet, spreadsheetService, salesSummary.summarizeByCategory(),
    CONFIG.HEADERS.CATEGORY, currentColumn
  );

  // 販売経路別集計
  currentColumn = _processSummaryBlock(
    outputSheet, spreadsheetService, salesSummary.summarizeByChannel(),
    CONFIG.HEADERS.CHANNEL, currentColumn
  );

  // 売上先（法人／個人）の集計
  _processCustomerSummaryBlock(
    outputSheet, spreadsheetService, salesSummary.summarizeByCustomer(),
    CONFIG.HEADERS.CUSTOMER, currentColumn
  );

  Logger.log("単軸売上集計が完了しました。");
}


/**
 * 出力シートのヘッダーを初期設定します。
 * 既にヘッダーが存在する場合は上書きしません。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} outputSheet - 出力シート
 * @private
 */
function _initializeHeaders(outputSheet) {
  // 1行目の既存の値をチェックし、空の場合のみ書き込む
  const existingHeaders = outputSheet.getRange(1, 1, 1, 16).getValues()[0];

  const headerConfigs = [
    { startCol: 1, config: CONFIG.HEADERS.MONTHLY },
    { startCol: 5, config: CONFIG.HEADERS.CATEGORY },
    { startCol: 9, config: CONFIG.HEADERS.CHANNEL },
    { startCol: 13, config: CONFIG.HEADERS.CUSTOMER }
  ];

  headerConfigs.forEach(item => {
    const headerText = item.config.text;
    const startCol = item.startCol;
    const numCols = item.config.cols;

    // 現在のセルが空かどうかをチェック (最初のセルだけを簡易的にチェック)
    if (!existingHeaders[startCol - 1] || existingHeaders[startCol - 1].toString().trim() === "") {
        // ヘッダーを書き込む前に、その範囲のフォーマットもクリア
        outputSheet.getRange(1, startCol, 1, numCols).clearFormat();
        outputSheet.getRange(1, startCol, 1, numCols).setValues([headerText]);
        Logger.log(`ヘッダーを書き込みました: ${headerText} (列: ${startCol})`);

    } else {
        Logger.log(`ヘッダーは既に存在します: ${existingHeaders[startCol - 1]} (列: ${startCol})`);
    }
  });
}


/**
 * 個別の集計ブロック（月別、カテゴリ別、販売経路別）を処理し、シートに出力
 * @param {GoogleAppsScript.Spreadsheet.Sheet} outputSheet - 出力シート
 * @param {SpreadsheetService}            spreadsheetService - SpreadsheetServiceインスタンス
 * @param {Object} summaryMap - 集計結果のマップ (例: ymMap, categoryMap)
 * @param {Object} headerConfig - ヘッダー設定 ({ text: Array<string>, cols: number })
 * @param {number} startColumn - このブロックの開始列
 * * @returns {number} 次のブロックの開始列
 * @private
 */
function _processSummaryBlock(outputSheet, spreadsheetService, summaryMap, headerConfig, startColumn) {
  const values = Object.entries(summaryMap)
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([key, val]) => [key, val.count, val.total]);

  spreadsheetService.clearAndWriteDataBlock(outputSheet, startColumn, values, headerConfig.cols);

  // 罫線と背景色の適用
  // ヘッダー行にも罫線と背景色を適用
  spreadsheetService.applyBorders(outputSheet, 1, startColumn, values.length + 1, headerConfig.cols);
  spreadsheetService.applyBackgroundColor(outputSheet, 1, startColumn, 1, headerConfig.cols, CONFIG.COLORS.LIGHT_GRAY);

  // データ行の左端（カテゴリ名など）のみに背景色を適用
  if (values.length > 0) {
    spreadsheetService.applyBackgroundColor(outputSheet, 2, startColumn, values.length, 1, CONFIG.COLORS.LIGHT_GRAY);
  } 
  
  return startColumn + headerConfig.cols + 1; // 次のブロックの開始列を返す
}

/**
 * 売上先（法人／個人）の集計を処理し、シートに出力
 * 他の集計とは出力形式が異なるため、個別関数としている
 * * @param {GoogleAppsScript.Spreadsheet.Sheet} outputSheet - 出力シート
 * @param {SpreadsheetService} spreadsheetService - SpreadsheetServiceインスタンス
 * @param {Object} customerMap - 顧客集計結果のマップ
 * @param {Object} headerConfig - ヘッダー設定 ({ text: Array<string>, cols: number })
 * @param {number} startColumn - このブロックの開始列
 * @private
 */
function _processCustomerSummaryBlock(outputSheet, spreadsheetService, customerMap, headerConfig, startColumn) {
  const customerOutputValues = [];
  
  // 1. 法人・個人の小計を先に計算して追加
  ["法人", "個人"].forEach(type => {
    const group = customerMap[type];
    let subtotalCount = 0;
    let subtotalAmount = 0;

    for (const [id, val] of Object.entries(group)) {
      subtotalCount += val.count;
      subtotalAmount += val.total;
    }
    customerOutputValues.push([type, "（小計）", subtotalCount, subtotalAmount]);
  });

  // 2. 各IDの明細を追加
  ["法人", "個人"].forEach(type => {
    const group = customerMap[type];
    const entries = Object.entries(group).sort((a, b) => a[0].localeCompare(b[0]));
    
    for (const [id, val] of entries) {
      customerOutputValues.push([type, id, val.count, val.total]);
    }
  });
  
  spreadsheetService.clearAndWriteDataBlock(outputSheet, startColumn, customerOutputValues, headerConfig.cols);
  
  // 罫線と背景色の適用
  // ヘッダー行にも罫線と背景色を適用
  spreadsheetService.applyBorders(outputSheet, 1, startColumn, customerOutputValues.length + 1, headerConfig.cols);
  spreadsheetService.applyBackgroundColor(outputSheet, 1, startColumn, 1, headerConfig.cols, CONFIG.COLORS.LIGHT_GRAY);

  // データ行の左端（区分、名称）のみに背景色を適用
  if (customerOutputValues.length > 0) {
    // 顧客集計は2列（区分、名称）が左端なので、numCols を 2 にする
    spreadsheetService.applyBackgroundColor(outputSheet, 2, startColumn, customerOutputValues.length, 2, CONFIG.COLORS.LIGHT_GRAY);
  } 
}
