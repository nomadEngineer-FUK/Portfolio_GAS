class SpreadsheetService {
  constructor(spreadsheet) {
    this.ss = spreadsheet;
  }

  getSheet(sheetName, createIfNotExist = false) {
    let sheet = this.ss.getSheetByName(sheetName);
    if (!sheet && createIfNotExist) {
      sheet = this.ss.insertSheet(sheetName);
    }
    return sheet;
  }

  getDataRangeValues(sheetName) {
    const sheet = this.getSheet(sheetName);
    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません。`);
    }
    return sheet.getDataRange().getValues();
  }

  /**
   * 数値にカンマ区切りを適用します。
   * @param {number|string} value - フォーマットする値
   * @returns {string|number} フォーマットされた文字列、または元の値（数値でない場合）
   */
  _formatNumberWithCommas(value) {
    const num = Number(value);
    if (!isNaN(num) && (typeof value === 'number' || (typeof value === 'string' && value.match(/^-?\d+(\.\d+)?$/)))) {
      return num.toLocaleString();
    }
    return value;
  }

  /**
   * データ配列内の数値をカンマ区切りにフォーマットします。
   * @param {Array<Array<any>>} data - フォーマットする二次元配列
   * @returns {Array<Array<any>>} フォーマットされた二次元配列
   */
  _formatDataValues(data) {
    return data.map(row => {
      return row.map(cell => this._formatNumberWithCommas(cell));
    });
  }

  /**
   * 指定された範囲（2行目以降）にデータ行を書き込みます。
   * 書き込み前に、該当範囲の既存データとフォーマットをクリアします。
   * ヘッダーの書き込みは行いません。
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
   * @param {number} startCol - 開始列
   * @param {Array<Array<any>>} dataValues - 書き込むデータ（ヘッダーを除く）
   * @param {number} numCols - データブロックの列数（ヘッダーと一致するはず）
   */
  clearAndWriteDataBlock(sheet, startCol, dataValues, numCols) {
    const clearRows = Math.max(dataValues.length + 5, 100);
    const rangeToClear = sheet.getRange(2, startCol, clearRows, numCols);
    rangeToClear.clearFormat();
    rangeToClear.clearContent();

    if (dataValues.length > 0) {
      const formattedValues = this._formatDataValues(dataValues);
      sheet.getRange(2, startCol, formattedValues.length, formattedValues[0].length).setValues(formattedValues);
    }
  }

  /**
   * クロス集計のデータと合計行をシートに書き込みます。
   * 書き込み前に、データと合計行の範囲をクリアします。
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
   * @param {Array<Array<any>>} dataRows - データ行
   * @param {Array<any>} totalRow - 合計行
   * @param {number} numCols - 書き込む列数
   * @param {number} startRow - 書き込み開始行 (ヘッダー行を含む)
   */
  clearAndWriteCrossDataBlock(sheet, dataRows, totalRow, numCols, startRow) {
    const totalContentRows = dataRows.length + 1; // データ行 + 合計行
    const clearMaxRows = Math.max(totalContentRows + 5, 100);
    const rangeToClear = sheet.getRange(startRow, 1, clearMaxRows, numCols);
    rangeToClear.clearFormat();
    rangeToClear.clearContent();

    // データ行の書き込み
    if (dataRows.length > 0) {
      const formattedDataRows = this._formatDataValues(dataRows);
      sheet.getRange(startRow, 1, formattedDataRows.length, formattedDataRows[0].length).setValues(formattedDataRows);
    }
    // 合計行の書き込み
    const formattedTotalRow = totalRow.map(cell => this._formatNumberWithCommas(cell));
    sheet.getRange(startRow + dataRows.length, 1, 1, formattedTotalRow.length).setValues([formattedTotalRow]);
  }

  /**
   * 指定された範囲に四方を囲む罫線を適用します。
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
   * @param {number} row - 開始行
   * @param {number} column - 開始列
   * @param {number} numRows - 行数
   * @param {number} numColumns - 列数
   */
  applyBorders(sheet, row, column, numRows, numColumns) {
    const range = sheet.getRange(row, column, numRows, numColumns);
    range.setBorder(true, true, true, true, true, true);
  }

  /**
   * 指定された範囲に背景色を適用します。
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
   * @param {number} row - 開始行
   * @param {number} column - 開始列
   * @param {number} numRows - 行数
   * @param {number} numColumns - 列数
   * @param {string} color - 背景色（例: "#F0F0F0"）
   */
  applyBackgroundColor(sheet, row, column, numRows, numColumns, color) {
    const range = sheet.getRange(row, column, numRows, numColumns);
    range.setBackground(color);
  }
}