/**
 * クロス集計のロジックをカプセル化するクラス。
 * 任意の2つの軸（行と列）を指定して集計できます。
 */
class CrossTabSummary {
  /**
   * @param {Array<Array<any>>} salesData - 元の売上データ (ヘッダー含む)
   * @param {string} rowHeaderName - 行ヘッダーとして使用するDBシートの列名
   * @param {string} colHeaderName - 列ヘッダーとして使用するDBシートの列名
   * @param {function(any): string} [rowKeyTransformer] - 行キーを変換する関数（例: "co..."を"法人"に）
   * @param {function(any): string} [colKeyTransformer] - 列キーを変換する関数（例: Dateオブジェクトを"YYYY/MM"に）
   */
  constructor(salesData, rowHeaderName, colHeaderName, rowKeyTransformer = (val) => String(val).trim(), colKeyTransformer) {
    this.salesData = salesData;
    this.headers = salesData[0];
    this.rows = salesData.slice(1);

    this.rowHeaderName = rowHeaderName;
    this.colHeaderName = colHeaderName;
    this.rowKeyTransformer = rowKeyTransformer;

    // colKeyTransformer が明示的に渡されなかった場合、colHeaderNameが"日付"なら年月形式に変換
    if (colKeyTransformer) {
      this.colKeyTransformer = colKeyTransformer;
    } else if (colHeaderName === "日付") {
      this.colKeyTransformer = (dateValue) => {
        if (!dateValue) return "不明な日付";
        try {
          const date = new Date(dateValue);
          if (isNaN(date.getTime())) return "無効な日付";
          return `${date.getFullYear()}/${("0" + (date.getMonth() + 1)).slice(-2)}`;
        } catch (e) {
          return "日付エラー";
        }
      };
    } else {
      this.colKeyTransformer = (val) => String(val).trim();
    }


    this.indexes = {
      row: this.headers.indexOf(rowHeaderName),
      col: this.headers.indexOf(colHeaderName),
      amount: this.headers.indexOf("金額"),
    };

    // 必要なインデックスが存在するかチェック
    if (this.indexes.row === -1 || this.indexes.col === -1 || this.indexes.amount === -1) {
      const missing = [];
      if (this.indexes.row === -1) missing.push(rowHeaderName);
      if (this.indexes.col === -1) missing.push(colHeaderName);
      if (this.indexes.amount === -1) missing.push("金額");
      throw new Error(`必要なヘッダーが見つかりません: ${missing.join(', ')}。DBシートのヘッダー名と一致しているか確認してください。`);
    }
  }

  /**
   * クロス集計を実行し、シート出力用の整形済みデータを返します。
   * @returns {{headerRow: Array<string>, dataRows: Array<Array<any>>, totalRow: Array<any>}} 整形済みデータ
   */
  getSummaryDataForOutput() {
    const crossMap = {}; // {rowKey: {colKey: 合計}}
    const allColKeys = new Set();
    const allRowKeys = new Set();

    for (const row of this.rows) {
      const rowValue = row[this.indexes.row];
      const colValue = row[this.indexes.col];
      const amount = parseFloat(row[this.indexes.amount]) || 0;

      // キーの変換とバリデーション
      const rowKey = this.rowKeyTransformer(rowValue);
      const colKey = this.colKeyTransformer(colValue);

      if (!rowKey || rowKey === "不明" || rowKey === "無効" || rowKey === "エラー") {
        // Logger.log(`警告: 行キー変換エラーまたは無効な行キー: ${rowValue}. 行キー: ${rowKey}. Row: ${JSON.stringify(row)}`);
        continue;
      }
      if (!colKey || colKey === "不明な日付" || colKey === "無効な日付" || colKey === "日付エラー") {
        // Logger.log(`警告: 列キー変換エラーまたは無効な列キー: ${colValue}. 列キー: ${colKey}. Row: ${JSON.stringify(row)}`);
        continue;
      }

      if (!crossMap[rowKey]) crossMap[rowKey] = {};
      crossMap[rowKey][colKey] = (crossMap[rowKey][colKey] || 0) + amount;

      allColKeys.add(colKey);
      allRowKeys.add(rowKey);
    }

    // 日付形式の場合は 'YYYY/MM' でソートされるので問題ない
    const sortedColKeys = Array.from(allColKeys).sort();
    const sortedRowKeys = Array.from(allRowKeys).sort();

    // ヘッダー行の構築
    const headerRow = [`${this.rowHeaderName}＼${this.colHeaderName}`, ...sortedColKeys, "合計"];

    const dataRows = [];
    const numColKeys = sortedColKeys.length;

    // 各行キーごとの行生成
    for (const rowKey of sortedRowKeys) {
      const colValues = sortedColKeys.map(colKey => crossMap[rowKey][colKey] || 0);
      const rowSum = colValues.reduce((a, b) => a + b, 0);
      dataRows.push([rowKey, ...colValues, rowSum]);
    }

    // 合計行の生成
    const totalRow = ["合計"];
    for (let i = 0; i < numColKeys; i++) {
      let sum = 0;
      for (const row of dataRows) {
        sum += row[i + 1]; // +1 は行キーをスキップ
      }
      totalRow.push(sum);
    }
    totalRow.push(totalRow.slice(1).reduce((a, b) => a + b, 0)); // 総合計

    return { headerRow, dataRows, totalRow };
  }
}