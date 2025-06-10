# Todo進捗トラッカー & LINE通知システム

## 概要

このプロジェクトは、Google スプレッドシートをデータベースとして活用し、**Todoの進捗状況をウェブアプリケーション形式で可視化**するフロントエンド（GAS Web App）と、**期限が迫ったTodoをLINEに通知**するバックエンド（GASスクリプト）で構成されています。

Todoの管理、進捗の追跡、そして重要な期限の見落とし防止を目的とした、Google Workspace環境向けの総合的なタスク管理ソリューションです。

## 主要機能

### 1. Todo進捗トラッカー (GAS Web App)

* **進捗の可視化**: スプレッドシートのTodoリストに基づき、全体の完了度をパーセンテージとプログレスバーで表示します。
* **難易度に基づくスコアリング**: 各Todoに設定された難易度（★☆☆、★★☆、★★★）に応じて点数を割り当て、合計点数と完了点数から進捗を計算します。
* **ステージ表示**: 進捗率に応じて「Stage 1」から「Complete」までのステージが動的に表示されます。
* **Web UI**: スプレッドシートの情報を基に、シンプルなWebインターフェースで進捗状況をリアルタイムに確認できます。
* **完了時の演出**: 進捗率が95%を超えると、クラッカーが🎉打ち上がる演出があります（Web UI上）。

### 2. LINE期限通知システム (GASスクリプト)

* **Todo期限の自動監視**: スプレッドシート内の未完了のTodoを定期的にチェックします。
* **期限カテゴリ分類**: 期限切れ、本日、3日以内、7日以内の4つのカテゴリにTodoを自動で分類します。
* **LINE通知**: 分類された期限情報に基づき、LINE Messaging APIを通じて指定のLINEユーザーへプッシュ通知を送信します。
* **必要なヘッダーの検証**: スプレッドシートのヘッダーが正しく設定されているかを検証し、エラーを防止します。

## 動作環境

* Google Workspaceアカウント
* Google Apps Script (GAS)
* Google Drive API (v2)
* LINE Developers アカウント（LINE Messaging APIの利用のため）

## セットアップ

### 1. スプレッドシートの準備

1.  新しいGoogleスプレッドシートを作成し、名前を付けます（例: `MyTodoManager`）。
2.  このスプレッドシート内に、以下の2つのシートを作成します。
    * **`todo` シート**: Todoリストのメインデータが格納されます。
        * **A列**: 完了ステータス (チェックボックスを推奨)
        * **C列**: 種別 (例: 「開発」「学習」など) - `validateInput` 関数で必須チェック
        * **G列**: 難易度 (例: `★☆☆`, `★★☆`, `★★★`) - `validateInput` 関数で必須チェック
        * その他、B列, F列, I列などもデータ入力に使用されます。
    * **`ToEmiratesStadium` シート**: LINE通知対象のTodoリストが格納されます。（プロジェクト名の変更に応じてシート名も変更してください。）
        * **A1セルに以下のヘッダーを含む行を設定**: `No`, `Todo`, `期限`, `難易度`
            * **A列**: 完了ステータス (チェックボックスを推奨)
            * **C列**: `Todo` (タスク内容)
            * **D列**: `期限` (日付形式)
            * **E列**: `難易度`
            * その他、`No`列やその他の列も利用されます。

### 2. GASプロジェクトの作成とコードの配置

1.  上記で作成したスプレッドシートを開き、「拡張機能」>「Apps Script」を選択してApps Scriptエディタを開きます。
2.  新しいプロジェクトを作成し、名前を付けます（例: `TodoProgressApp`）。
3.  以下のファイルをそれぞれ作成し、対応するコードを貼り付けます。
    * `Code.gs` (または任意の`.gs`ファイル): `doGet`, `getProgressData` などの進捗トラッカー関連関数
    * `lineAPI.gs` (または任意の`.gs`ファイル): `checkDeadlinesAndNotify` などのLINE通知関連関数
    * `index.html`: Web UIのためのHTML、CSS、JavaScript

### 3. 必要なGoogleサービスとLINE Messaging APIの設定

1.  **Google Drive APIの有効化**:
    * Apps Scriptエディタの左側メニューで「サービス」の横にある「＋」アイコンをクリックします。
    * 「**Drive API**」を探してクリックし、「追加」ボタンをクリックします。(バージョンは `v2` を想定)
2.  **LINE Messaging APIの設定**:
    * [LINE Developers コンソール](https://developers.line.biz/console/) にアクセスし、プロバイダーとチャネルを作成します。
    * 作成したチャネルの「Messaging API設定」タブから「**チャネルアクセストークン（長期）**」を取得します。
    * `lineAPI.gs` ファイル内の `LINE_CHANNEL_ACCESS_TOKEN` 定数に、取得したトークンを設定します。
    * 通知を受け取りたいLINEユーザーの**ユーザーID**を取得します。これは、`doPost` 関数（Webhook経由）や、LINEのBOTアカウントに話しかけてログから取得する方法などがあります。
    * `lineAPI.gs` ファイル内の `LINE_USER_ID` 定数に、取得したユーザーIDを設定します。

### 4. Webアプリのデプロイ (進捗トラッカー用)

1.  Apps Scriptエディタで、「デプロイ」>「新しいデプロイ」を選択します。
2.  「種類を選択」で「ウェブアプリ」を選択します。
3.  以下の設定を行います。
    * **実行ユーザー**: 「自分」
    * **アクセスできるユーザー**: 「全員」（または制限したい場合は適宜設定）
4.  「デプロイ」をクリックします。
5.  デプロイ完了後、表示される「ウェブアプリのURL」を控えておきます。このURLにアクセスすると、進捗トラッカーのWeb UIが表示されます。

### 5. トリガーの設定 (LINE通知用)

LINE通知を定期的に実行するには、GASのトリガーを設定します。

1.  Apps Scriptエディタの左側メニューで「トリガー」アイコン（時計のような形）をクリックします。
2.  「トリガーを追加」ボタンをクリックします。
3.  以下の設定を行います。
    * **実行する関数を選択**: `checkDeadlinesAndNotify`
    * **イベントのソースを選択**: 「時間主導型」
    * **時間ベースのトリガーのタイプを選択**: 例: 「日付ベースのタイマー」
    * **時間の間隔を選択**: 例: 「1日おき」「午前8時～9時」など、通知したい頻度と時間を設定
4.  「保存」をクリックします。

## 処理フロー

### Todo進捗トラッカー (Web UIアクセス時)

1.  ユーザーがデプロイされたウェブアプリのURLにアクセスします。
2.  `doGet()` 関数が実行され、`index.html` ファイルの内容をブラウザに表示します。
3.  `index.html` 内のJavaScriptが `google.script.run.withSuccessHandler().getProgressData()` を呼び出します。
4.  GAS側の `getProgressData()` 関数が実行されます。
    * アクティブなスプレッドシートの `todo` シートからデータを取得します。
    * C列とG列の入力バリデーションを行います。
    * 各Todoの難易度に基づき点数を計算し、完了しているTodoの点数を集計します。
    * 全体の進捗率と、それに応じた「ステージ」を決定します。
5.  `getProgressData()` から返された進捗データが `index.html` のJavaScriptに渡されます。
6.  JavaScriptがプログレスバー、ステージ表示、および必要に応じてクラッカー演出を更新し、ユーザーに表示します。

### LINE期限通知システム (トリガー実行時)

1.  設定されたトリガーによって、定期的に `checkDeadlinesAndNotify()` 関数が実行されます。
2.  アクティブなスプレッドシートの `ToEmiratesStadium` シートからTodoデータを読み込みます。
3.  「No」「Todo」「期限」「難易度」の必須ヘッダーが存在するかを検証します。
4.  各Todoの完了ステータスと期限日をチェックします。
5.  未完了のTodoについて、現在の日付との比較に基づき「期限切れ」「本日」「3日以内」「7日以内」のいずれかのカテゴリに分類します。
6.  分類されたTodo情報をもとに、LINE Messaging APIで送信するメッセージを作成します。
7.  作成されたメッセージを `sendLinePushMessage()` 関数を通じて、設定されたLINEユーザーIDにプッシュ通知として送信します。
8.  通知対象のTodoがない場合は、「通知対象のToDoはありません。」というメッセージを送信します。

## 開発者向け情報

### スプレッドシートの構造

* **`todo` シート**: 進捗トラッカーのデータソース
    * `A`列: `isDone` (Boolean: チェックボックス)
    * `C`列: `typeForValidation` (String: 種別)
    * `G`列: `difficultyForValidation` (String: 難易度, "★☆☆", "★★☆", "★★★")
* **`ToEmiratesStadium` シート**: LINE通知のデータソース
    * `A`列: `isCompleted` (Boolean: チェックボックス)
    * `No`列: `no` (String/Number)
    * `Todo`列: `todo` (String: タスク内容)
    * `期限`列: `deadline` (Date: 期限日)
    * `難易度`列: `difficulty` (String: 難易度)

### 主要関数

#### `Code.gs`
* `doGet()`: Webアプリのエントリーポイント。
* `getProgressData()`: スプレッドシートから進捗データを取得し、計算、バリデーションを行うメインロジック。
* `getValidRangeData(sheet)`: 有効なデータ範囲を動的に取得。
* `validateInput(data)`: 特定の列の入力漏れをチェック。
* `calculateScore(data)`: Todoの難易度に基づき、合計点と完了点、進捗率を計算。
* `convertDifficultyToScore(difficulty)`: 難易度文字列を数値スコアに変換。
* `getStageKey(progress)`: 進捗率に応じたステージキー（例: "Stage 1"）を返す。
* `getStage(progress)`: Web UIに表示するステージの文言を返す。

#### `lineAPI.gs`
* `checkDeadlinesAndNotify()`: 期限チェックとLINE通知のメイン関数。
* `getHeaderIndex(headerRow, requiredHeadersMap)`: ヘッダー行から必要な列のインデックスを取得。
* `appendCategoryMessage(currentMsg, todos, headerLabel)`: LINEメッセージにカテゴリ別Todoを追加。
* `sendLinePushMessage(messageText)`: LINE Messaging APIを呼び出してメッセージを送信。
* `doPost(e)`: (参考用) LINE Webhookからのリクエストを処理し、ユーザーIDを取得する例。

## 著作権・ライセンス

[必要に応じてライセンス情報を記載してください。例: MIT License]
