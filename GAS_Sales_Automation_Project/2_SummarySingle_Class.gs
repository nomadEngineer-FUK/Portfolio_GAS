class SalesSummary {
  constructor(salesData) {
    this.salesData = salesData;
    this.headers = salesData[0];
    this.rows = salesData.slice(1);

    this.indexes = {
      date: this.headers.indexOf("日付"),
      category: this.headers.indexOf("カテゴリ"),
      channel: this.headers.indexOf("販売経路"),
      customer: this.headers.indexOf("名称"),
      amount: this.headers.indexOf("金額"),
    };
    Logger.log(`Headers: ${this.headers}`);
    Logger.log(`Indexes: ${JSON.stringify(this.indexes)}`);
    
    // 取得したインデックスが -1 でないか確認
    if (Object.values(this.indexes).some(index => index === -1)) {
      Logger.log("警告: ヘッダーが見つからない列があります。シートのヘッダー名と一致しているか確認してください。");
    }
  }

  // 各軸の集計を共通で処理するヘルパー関数
  _summarizeBy(keyName, map, processKey = (value) => value) { // keyIndex を keyName に変更
    const keyIndex = this.indexes[keyName]; // インデックスをここで取得
    if (keyIndex === -1) {
      Logger.log(`警告: 集計キー "${keyName}" の列が見つかりません。この集計はスキップされます。`);
      return map;
    }
    for (const row of this.rows) {
      // 該当する列のデータが存在しない場合はスキップ
      if (row[keyIndex] === undefined || row[keyIndex] === null) {
        // Logger.log(`スキップ: 列 ${this.headers[keyIndex]} のデータが空です。Row: ${JSON.stringify(row)}`);
        continue;
      }
      const key = processKey(row[keyIndex]);
      const amount = parseFloat(row[this.indexes.amount]) || 0; // 金額は常にamountIndexから取得

      if (!map[key]) map[key] = { count: 0, total: 0 };
      map[key].count++;
      map[key].total += amount;
    }
    return map;
  }

  summarizeByMonth() {
    const ymMap = {};
    return this._summarizeBy('date', ymMap, (dateValue) => {
      // 日付の処理も堅牢に
      if (!dateValue) return "不明な日付";
      try {
        const date = new Date(dateValue);
        if (isNaN(date.getTime())) { // 無効な日付をチェック
          Logger.log(`警告: 無効な日付データ: ${dateValue}. Row: ${JSON.stringify(row)}`);
          return "無効な日付";
        }
        return `${date.getFullYear()}/${("0" + (date.getMonth() + 1)).slice(-2)}`;
      } catch (e) {
        Logger.log(`エラー: 日付処理中に問題が発生しました: ${dateValue}. エラー: ${e.message}`);
        return "日付エラー";
      }
    });
  }

  summarizeByCategory() {
    const categoryMap = {};
    return this._summarizeBy('category', categoryMap);
  }

  summarizeByChannel() {
    const channelMap = {};
    return this._summarizeBy('channel', channelMap);
  }

  summarizeByCustomer() {
    const customerMap = { "法人": {}, "個人": {} };
    const customerIndex = this.indexes.customer;
    const amountIndex = this.indexes.amount;

    if (customerIndex === -1) {
      Logger.log("警告: '売上先（購入者）' 列が見つかりません。顧客集計はスキップされます。");
      return customerMap;
    }
    if (amountIndex === -1) {
      Logger.log("警告: '金額' 列が見つかりません。顧客集計はスキップされます。");
      return customerMap;
    }

    for (const row of this.rows) {
      const customer = row[customerIndex];
      const amount = parseFloat(row[amountIndex]) || 0;

      if (customer === undefined || customer === null || customer === "") { // 空文字列も考慮
        // Logger.log(`売上先（購入者）が空の行をスキップしました: ${JSON.stringify(row)}`);
        continue; // この行の処理をスキップ
      }

      const customerStr = String(customer).trim();
      if (customerStr === "") {
        // Logger.log(`売上先（購入者）が空白のみの行をスキップしました: ${JSON.stringify(row)}`);
        continue;
      }

      const type = customerStr.startsWith("co") ? "法人" : "個人";
      
      // Logger.log(`処理中の行: customer=${customerStr}, amount=${amount}, type=${type}`); // 各行の処理を確認

      if (!customerMap[type][customerStr]) {
        customerMap[type][customerStr] = { count: 0, total: 0 };
      }
      customerMap[type][customerStr].count++;
      customerMap[type][customerStr].total += amount;
    }
    // Logger.log(`最終的な customerMap: ${JSON.stringify(customerMap)}`); // 集計結果を確認
    return customerMap;
  }
}