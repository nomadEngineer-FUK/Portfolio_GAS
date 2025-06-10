# 売上データ分析システム

## 概要

このGoogle Apps Script (GAS) プロジェクトは、Google スプレッドシートに蓄積された売上データを多角的に集計し、その結果をGoogle Charts API を利用したウェブアプリケーション形式のダッシュボードで視覚的に表示するシステムです。

## 主要機能

### 1. 売上データ集計機能

売上データを様々な切り口で集計し、ビジネス上の意思決定をサポートします。

* **単軸集計**:
    * 月別、カテゴリ別、販売経路別の売上を集計します。
    * 顧客（法人/個人）別の売上を集計し、内訳も表示します。
    * 集計結果は整形されたGoogleスプレッドシートに出力されます。
* **クロス集計**:
    * 任意の二つの軸（例: カテゴリ vs 月、販売経路 vs 月、カテゴリ vs 販売経路、顧客区分 vs 月、顧客区分 vs 販売経路）で売上をクロス集計します。
    * 特に「顧客区分（法人/個人）」を軸としたクロス集計も可能です。
    * 集計結果は、見やすく整形されたGoogleスプレッドシートに出力されます。

### 2. 売上分析グラフ表示機能 (Web UI)

Google Charts API を利用し、集計された売上データを視覚的に理解しやすいグラフで表示します。

* **カテゴリ別 月次売上推移 (折れ線グラフ)**: クロス集計シートから「カテゴリ＼日付」のデータを読み込み、時間の経過に伴うカテゴリごとの売上トレンドを把握できます。
* **販売経路別 カテゴリ構成 (積み上げ棒グラフ)**: クロス集計シートから「カテゴリ＼販売経路」のデータを読み込み、各販売経路におけるカテゴリ売上の割合を比較できます。
* **法人／個人売上比率 (円グラフ)**: 単軸集計シートから法人と個人の小計を直接取得し、全体売上における法人顧客と個人顧客の貢献度を一目で確認できます。
* **応答性とエラーハンドリング**: データがない場合やGASの実行エラーが発生した場合にも、ユーザーに分かりやすいメッセージをウェブページ上に表示します。

## 動作環境

* Google Workspaceアカウント
* Google Apps Script (GAS)
* Google Charts API (Web UIのグラフ描画のため)

## ファイル構成

このプロジェクトは以下の主要なGASスクリプトおよびHTMLファイルで構成されています。

* `index.html`: 売上分析グラフを表示するウェブアプリケーションのフロントエンド。
* `1_config.gs`: 各種設定（シート名、ヘッダー定義、色など）をまとめた定数ファイル。
* `1_excuteAllFunction.gs`: 主要な実行関数を含むファイル。
* `1_spreadsheetService.gs`: スプレッドシート操作の共通ユーティリティを提供するクラス。データの読み書き、フォーマット適用などを担当。
* `2_SummarySingle_Class.gs`: 単軸集計ロジックをカプセル化したクラス。
* `2_SummarySingle_Code.gs`: 単軸集計のメイン処理と、`SummarySingle_Class`を使用したシート出力。
* `3_SummaryCross_Class.gs`: クロス集計ロジックをカプセル化するクラス。
* `3_SummaryCross_Code.gs`: クロス集計のメイン処理と、`SummaryCross_Class`を使用したシート出力。
* `4_ChartsDataService.gs`: ウェブアプリ (`index.html`) から呼び出され、Google Charts用のデータを提供するGAS関数群 (`getSalesChannelByMonthData`, `getSalesChannelByCategoryData`, `getCustomerTypeRatioData` など) を含む。

## セットアップ

### 1. スプレッドシートの準備

新しいGoogleスプレッドシートを作成し、以下のシートを準備します。

* **`DB` シート**: 元の売上データが格納されます。
    * 必須ヘッダー: `日付`, `カテゴリ`, `販売経路`, `名称`, `金額`
    * 例: A1:日付, B1:カテゴリ, C1:販売経路, D1:名称, E1:金額, ...
* **`Summary_Single` シート**: 単軸集計の結果が出力されます。
* **`Summary_Cross` シート**: クロス集計の結果が出力されます。

### 2. GASプロジェクトの作成とコードの配置

1.  上記で作成したスプレッドシートを開き、「拡張機能」>「Apps Script」を選択してApps Scriptエディタを開きます。
2.  新しいプロジェクトを作成し、任意の名前を付けます。
3.  提供された以下のコードを、それぞれのファイル名でスクリプトエディタ内に作成・貼り付けます。
    * `index.html`
    * `ChartsDataService.gs`
    * `config.gs`
    * `excuteAllFunction.gs`
    * `spreadsheetService.gs`
    * `SummarySingle_Class.gs`
    * `SummarySingle_Code.gs`
    * `SummaryCross_Class.gs`
    * `SummaryCross_Code.gs`

    **注意**: ウェブアプリのエントリーポイントとして機能する`doGet()`関数が、現在`ChartsDataService.gs`内で`ChartsSidebar.html`を返すように設定されています。`index.html`の内容をウェブアプリとして表示したい場合は、`ChartsDataService.gs`の`doGet()`関数を以下のように修正してください。

    ```javascript
    function doGet() {
      return HtmlService.createTemplateFromFile('index') // index.html の内容を返す
        .evaluate()
        .setTitle('売上分析グラフ'); // ウィンドウのタイトル
    }
    ```

### 3. Webアプリのデプロイ

1.  Apps Scriptエディタで、「デプロイ」>「新しいデプロイ」を選択します。
2.  「種類を選択」で「ウェブアプリ」を選択します。
3.  以下の設定を行います。
    * **実行ユーザー**: 「自分」
    * **アクセスできるユーザー**: 「全員」（または制限したい場合は適宜設定）
    * **デプロイ名**: 例「売上分析ダッシュボード」
4.  「デプロイ」をクリックします。
5.  デプロイ完了後、表示される「ウェブアプリのURL」を控えておきます。このURLにアクセスすると、売上分析のWeb UIが表示されます。

### 4. トリガーの設定 (GASスクリプト実行用)

#### 売上集計の実行トリガー

集計処理を自動化する場合に設定します。手動で実行する場合は不要です。

1.  Apps Scriptエディタの左側メニューで「トリガー」アイコンをクリックします。
2.  「トリガーを追加」ボタンをクリックします。
3.  以下の設定を行います。
    * **実行する関数を選択**: 例: `generateSingleAxisSummaryRefactored` (単軸集計のメイン関数)
    * **イベントのソースを選択**: 「時間主導型」
    * **時間ベースのトリガーのタイプを選択**: 例: 「月ベースのタイマー」や「週ベースのタイマー」
    * **時間の間隔を選択**: データ更新の頻度に合わせて設定
4.  同様に、クロス集計のメイン関数（例: `generateAllCrossAxisSummaries`）についてもトリガーを設定します。

## 処理フロー

### 1. 売上データ集計 (GASスクリプト実行時)

* `generateSingleAxisSummaryRefactored()` または `generateAllCrossAxisSummaries()` が実行されます。
* `SpreadsheetService` クラスが初期化され、スプレッドシートの操作を抽象化します。
* `config.gs` で定義されたシート名から元データ (`DB`シート) が読み込まれます。
* `SalesSummary` クラス（単軸の場合）または `CrossTabSummary` クラス（クロス集計の場合）が売上データで初期化され、指定された軸に基づいてデータ集計ロジックを実行します。
* 集計されたデータは、`SpreadsheetService` を介して、それぞれの出力シート (`Summary_Single`、`Summary_Cross`) に書き込まれます。
* 書き込みの際、数値のカンマ区切りフォーマット、罫線、背景色などの視覚的な整形も適用されます。

### 2. 売上分析グラフ表示 (Web UIアクセス時)

* ユーザーがウェブアプリのURLにアクセスすると、`index.html` が `doGet()` によって返され、ブラウザに読み込まれます。
* `index.html` 内のJavaScriptが `google.charts.load` でGoogle Charts APIをロードし、`google.script.run` を介して`ChartsDataService.gs`内のデータ取得関数を呼び出します。
* `ChartsDataService.gs` 内の `getSalesChannelByMonthData()`, `getSalesChannelByCategoryData()`, `getCustomerTypeRatioData()` 関数がそれぞれ実行され、`Summary_Cross`または`Summary_Single`シートから整形済みのグラフデータ（ヘッダー含む二次元配列）を返します。
* HTML側のJavaScriptは、受け取ったデータを使用してGoogle Charts APIで折れ線グラフ、積み上げ棒グラフ、円グラフを動的に描画します。
* データがない場合やグラフ描画時にエラーが発生した場合は、適切なメッセージが表示されます。

## 開発者向け情報

### クラス構造

* `SpreadsheetService`: スプレッドシートへのアクセスとデータ操作（読み取り、書き込み、フォーマット）を一元管理する汎用的なクラス。
* `SalesSummary`: 単軸集計ロジックをカプセル化し、月別、カテゴリ別、販売経路別、顧客別の集計メソッドを提供。
* `CrossTabSummary`: クロス集計ロジックをカプセル化し、任意の2つの軸での集計と出力データ整形を提供。

### 定数 (`config.gs`)

* `CONFIG.SALES_SHEET_NAME`: 元データのシート名 (`DB`)
* `CONFIG.OUTPUT_SHEET_NAME_SINGLE`: 単軸集計の出力シート名 (`Summary_Single`)
* `CONFIG.OUTPUT_SHEET_NAME_CROSS`: クロス集計の出力シート名 (`Summary_Cross`)
* `CONFIG.HEADERS.MONTHLY`, `CONFIG.HEADERS.CATEGORY`, `CONFIG.HEADERS.CHANNEL`, `CONFIG.HEADERS.CUSTOMER`: 各集計ブロックのヘッダー定義
* `CONFIG.COLORS.LIGHT_GRAY`: スプレッドシートの背景色に使用するカラーコード

### 主な関数

* `getSalesChannelByMonthData()` (`ChartsDataService.gs`): 折れ線グラフ用データを取得。
* `getSalesChannelByCategoryData()` (`ChartsDataService.gs`): 積み上げ棒グラフ用データを取得。
* `getCustomerTypeRatioData()` (`ChartsDataService.gs`): 円グラフ用データを取得。
* `doGet()` (`ChartsDataService.gs`): Webアプリのエントリーポイント。
* `generateSingleAxisSummaryRefactored()` (`SummarySingle_Code.gs`): 単軸集計のメインエントリポイント。
* `generateAllCrossAxisSummaries()` (`SummaryCross_Code.gs`): クロス集計のメインエントリポイント。
