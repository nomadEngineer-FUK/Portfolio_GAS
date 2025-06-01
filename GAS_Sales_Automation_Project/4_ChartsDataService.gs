/**
 * ウェブアプリケーションとしてアクセスされたときに実行されるメイン関数。
 * HTMLコンテンツを返します。
 */
function doGet() {
  return HtmlService.createTemplateFromFile('ChartsSidebar') // ChartsSidebar.html の内容を返す
      .evaluate()
      .setTitle('売上分析グラフ'); // ウィンドウのタイトル
}

/**
 * 折れ線グラフ用のデータを取得し、整形して返します。
 * データソース: Summary_Crossシートの「販売経路×日付」ブロック
 * @returns {Array<Array<any>>} Google Charts DataTable形式のデータ
 */
function getSalesChannelByMonthData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (typeof CONFIG === 'undefined' || !CONFIG.OUTPUT_SHEET_NAME_CROSS) {
      Logger.log("エラー: CONFIG.OUTPUT_SHEET_NAME_CROSS が定義されていません。");
      return [];
    }
    const outputSheetCross = ss.getSheetByName(CONFIG.OUTPUT_SHEET_NAME_CROSS);

    if (!outputSheetCross) {
      Logger.log("エラー: Summary_Cross シートが見つかりません。シート名: " + CONFIG.OUTPUT_SHEET_NAME_CROSS);
      return [];
    }

    const allCrossData = outputSheetCross.getDataRange().getValues();

    let startRow = -1;
    let startCol = -1;
    const targetHeader = "カテゴリ＼日付";

    for (let r = 0; r < allCrossData.length; r++) {
      for (let c = 0; c < allCrossData[r].length; c++) {
        if (allCrossData[r][c] && allCrossData[r][c] === targetHeader) {
          startRow = r;
          startCol = c;
          break;
        }
      }
      if (startRow !== -1) break;
    }

    if (startRow === -1) {
      Logger.log(`警告: クロス集計ヘッダー「${targetHeader}」が見つかりません。`);
      return [];
    }

    const header = allCrossData[startRow].slice(startCol);

    const dataRows = [];
    let currentRow = startRow + 1;
    while (currentRow < allCrossData.length &&
           (allCrossData[currentRow][startCol] !== "" || allCrossData[currentRow].some(cell => String(cell).trim() !== ""))) {
      if (allCrossData[currentRow][startCol] === "合計") break;
      dataRows.push(allCrossData[currentRow].slice(startCol));
      currentRow++;
    }

    if (dataRows.length === 0) {
      Logger.log(`警告: ${targetHeader} のデータ行が見つかりません。`);
      return [];
    }

    const monthColumns = header.slice(1, header.length > 1 ? header.length - 1 : 1);
    const channelNames = dataRows.map(row => row[0]);

    if (channelNames.length === 0 || monthColumns.length === 0) {
      Logger.log("警告: カテゴリまたは月カラムが取得できませんでした。");
      return [];
    }

    const chartHeader = ['年月', ...channelNames];
    const chartData = [];

    for (let i = 0; i < monthColumns.length; i++) {
      const monthObject = monthColumns[i];
      let monthString;

      // 日付オブジェクトを文字列に変換
      if (monthObject instanceof Date) {
        if (!isNaN(monthObject.getTime())) { // 有効な日付オブジェクトか確認
          // YYYY-MM-DD 形式の文字列に変換
          const year = monthObject.getFullYear();
          const month = (monthObject.getMonth() + 1).toString().padStart(2, '0'); // 月は0から始まるので+1
          const day = monthObject.getDate().toString().padStart(2, '0');
          monthString = `${year}-${month}-${day}`;
          // Logger.log(`Converted Date: ${monthObject} -> String: ${monthString}`);
        } else {
          Logger.log(`警告: 無効なDateオブジェクトがmonthColumnsに含まれていました: ${monthObject}`);
          monthString = "無効な日付"; // またはエラーとして処理、行をスキップなど
        }
      } else {
        // Dateオブジェクトでなければ、そのまま文字列として扱う (あるいはエラー処理)
        monthString = String(monthObject);
        Logger.log(`警告: monthColumnがDateオブジェクトではありませんでした: ${monthObject}。文字列として使用します。`);
      }

      const rowValues = [monthString]; // 変換後の文字列を使用
      for (let j = 0; j < channelNames.length; j++) {
        let salesAmount = parseFloat(String(dataRows[j][i + 1]).replace(/,/g, ''));
        if (isNaN(salesAmount)) {
          salesAmount = 0;
        }
        rowValues.push(salesAmount);
      }
      chartData.push(rowValues);
    }

    const returnValue = [chartHeader, ...chartData];
    if (returnValue.length < 2 || !Array.isArray(returnValue[0]) || returnValue[0].length === 0) {
        Logger.log("警告: getSalesChannelByMonthData の最終的な戻り値が不十分です。");
    }
    Logger.log("getSalesChannelByMonthData: 正常にデータを返します (日付は文字列化済み)。");
    return returnValue;

  } catch (e) {
    Logger.log("getSalesChannelByMonthData で致命的なエラー: " + e.toString() + " Stack: " + e.stack);
    return []; // エラー時は必ず空配列を返す
  }
}

/**
 * 棒グラフ用のデータを取得し、整形して返します。
 * データソース: Summary_Crossシートの「カテゴリ×販売経路」ブロック
 * 内容: 各販売経路でのカテゴリ別売上内訳 (積み上げ棒グラフ)
 * @returns {Array<Array<any>>} Google Charts DataTable形式のデータ
 */
function getSalesChannelByCategoryData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetService = new SpreadsheetService(ss);
    const outputSheetCross = spreadsheetService.getSheet(CONFIG.OUTPUT_SHEET_NAME_CROSS);

    if (!outputSheetCross) {
      Logger.log("エラー: Summary_Cross シートが見つかりません。");
      return [];
    }

    let startRow = -1;
    let startCol = -1;
    const targetHeader = "カテゴリ＼販売経路";
    const allCrossData = outputSheetCross.getDataRange().getValues();

    for (let r = 0; r < allCrossData.length; r++) {
      for (let c = 0; c < allCrossData[r].length; c++) {
        if (allCrossData[r][c] === targetHeader) {
          startRow = r;
          startCol = c;
          break;
        }
      }
      if (startRow !== -1) break;
    }

    if (startRow === -1) {
      Logger.log(`警告: クロス集計ヘッダー「${targetHeader}」が見つかりません。`);
      return [];
    }

    const fullHeaderRowFromSheet = allCrossData[startRow].slice(startCol);

    let cleanFullHeaderRow = [...fullHeaderRowFromSheet];
    while (cleanFullHeaderRow.length > 0 && String(cleanFullHeaderRow[cleanFullHeaderRow.length - 1]).trim() === "") {
      cleanFullHeaderRow.pop();
    }

    // 'カテゴリ＼販売経路' を除き、かつ 末尾の '合計' を除いた販売経路名だけを抽出
    const channelNames = cleanFullHeaderRow.slice(1, cleanFullHeaderRow.length - 1);

    const originalDataRows = [];
    let currentRow = startRow + 1;
    while (currentRow < allCrossData.length &&
           (allCrossData[currentRow][startCol] !== "" || allCrossData[currentRow].some(cell => String(cell).trim() !== ""))) {
      if (String(allCrossData[currentRow][startCol]).trim() === "合計") break;
      originalDataRows.push(allCrossData[currentRow].slice(startCol, startCol + 1 + channelNames.length));
      currentRow++;
    }

    if (originalDataRows.length === 0) {
      Logger.log(`警告: ${targetHeader} のデータ行が見つかりません。`);
      return [];
    }

    const categoryNames = originalDataRows.map(row => row[0]);
    const chartHeader = ['販売経路', ...categoryNames];

    const transposedData = [];
    for (let colIdx = 0; colIdx < channelNames.length; colIdx++) {
      const channelName = channelNames[colIdx];
      const rowValues = [channelName];
      for (let rowIdx = 0; rowIdx < originalDataRows.length; rowIdx++) {
        // originalDataRows[rowIdx] は [カテゴリ名, 売上1, 売上2, ...] の形式
        // 数値データは originalDataRows[rowIdx][colIdx + 1] になるはず
        // channelNames のインデックスが colIdx なので、
        // originalDataRows のスライス方法と合わせて確認
        const salesAmount = parseFloat(String(originalDataRows[rowIdx][colIdx + 1]).replace(/,/g, '')) || 0;
        rowValues.push(salesAmount);
      }
      transposedData.push(rowValues);
    }

    if (!chartHeader || chartHeader.length === 0 || transposedData.length === 0) {
      Logger.log("警告: getSalesChannelByCategoryData の最終データが不十分です。");
      return [];
    }
    
    return [chartHeader, ...transposedData];

  } catch (e) {
    Logger.log(`getSalesChannelByCategoryData でエラーが発生しました: ${e.toString()} Stack: ${e.stack}`);
    return [];
  }
}

/**
 * 円グラフ用のデータを取得し、整形して返します。
 * データソース: Summary_SingleシートのM列「区分」とP列「金額」から、法人と個人の小計を直接取得
 * 内容: 法人／個人の総売上比率
 * @returns {Array<Array<any>>} Google Charts DataTable形式のデータ
 */
function getCustomerTypeRatioData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // CONFIGオブジェクトとシート名の確認
    if (typeof CONFIG === 'undefined' || !CONFIG.OUTPUT_SHEET_NAME_SINGLE) {
      Logger.log("エラー: CONFIG.OUTPUT_SHEET_NAME_SINGLE が定義されていません。");
      return [];
    }
    const sheetName = CONFIG.OUTPUT_SHEET_NAME_SINGLE; // 例: "Summary_Single"
    const outputSheetSingle = ss.getSheetByName(sheetName);

    if (!outputSheetSingle) {
      Logger.log(`エラー: シート "${sheetName}" が見つかりません。`);
      return [];
    }

    // ヘッダー行を取得 (1行目全体を想定)
    // スプレッドシートの列 M, N, O, P はそれぞれインデックスで 12, 13, 14, 15
    // 必要なのは M列の「区分」と P列の「金額」ですが、円グラフヘッダーは固定でOK
    // const headerRow = outputSheetSingle.getRange(1, 1, 1, outputSheetSingle.getLastColumn()).getValues()[0];
    // Logger.log("ヘッダー行: " + JSON.stringify(headerRow));


    // データは2行目と3行目から直接取得
    // M列 (区分), N列 (名称), P列 (金額) を想定
    // M列は13番目の列 (インデックス12)
    // N列は14番目の列 (インデックス13)
    // P列は16番目の列 (インデックス15)

    // 2行目のデータを取得 (法人の小計を想定)
    const corporateDataRow = outputSheetSingle.getRange(2, 13, 1, 4).getValues()[0]; // M2:P2 の範囲
    // 3行目のデータを取得 (個人の小計を想定)
    const personalDataRow = outputSheetSingle.getRange(3, 13, 1, 4).getValues()[0];  // M3:P3 の範囲

    let corporateTotal = 0;
    let personalTotal = 0;

    // 法人データの検証と金額取得
    // corporateDataRow[0] はM列、corporateDataRow[1] はN列、corporateDataRow[3] はP列
    if (corporateDataRow && corporateDataRow[0] === "法人" && corporateDataRow[1] === "（小計）") {
      corporateTotal = parseFloat(String(corporateDataRow[3]).replace(/,/g, '')) || 0;
    } else {
      Logger.log("警告: 2行目が期待した法人小計データではありません。");
      Logger.log("2行目 M列: " + corporateDataRow[0] + ", N列: " + corporateDataRow[1]);
    }

    // 個人データの検証と金額取得
    // personalDataRow[0] はM列、personalDataRow[1] はN列、personalDataRow[3] はP列
    if (personalDataRow && personalDataRow[0] === "個人" && personalDataRow[1] === "（小計）") {
      personalTotal = parseFloat(String(personalDataRow[3]).replace(/,/g, '')) || 0;
    } else {
      Logger.log("警告: 3行目が期待した個人小計データではありません。");
      Logger.log("3行目 M列: " + personalDataRow[0] + ", N列: " + personalDataRow[1]);
    }

    if (corporateTotal === 0 && personalTotal === 0) {
      Logger.log("警告: 法人および個人の小計データから有効な金額が取得できませんでした。");
      // 取得したデータをログに出力して確認
      Logger.log(`法人 (M2:P2): ${JSON.stringify(corporateDataRow)}, 金額P2: ${corporateDataRow ? corporateDataRow[3] : 'N/A'}`);
      Logger.log(`個人 (M3:P3): ${JSON.stringify(personalDataRow)}, 金額P3: ${personalDataRow ? personalDataRow[3] : 'N/A'}`);
      return [['区分', '売上合計'], ['法人', 0], ['個人', 0]]; // データがない場合でもグラフの枠は表示
    }
    Logger.log(`法人合計: ${corporateTotal}, 個人合計: ${personalTotal}`);

    return [
      ['区分', '売上合計'],
      ['法人', corporateTotal],
      ['個人', personalTotal]
    ];

  } catch (e) {
    Logger.log(`getCustomerTypeRatioData でエラーが発生しました: ${e.toString()} Stack: ${e.stack}`);
    return [];
  }
}
