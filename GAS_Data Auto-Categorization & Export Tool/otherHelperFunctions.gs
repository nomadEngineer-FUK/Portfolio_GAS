// --- 汎用ヘルパー関数群 ---

/**
 * 値をスプレッドシートで条件付きで強制的に文字列として表示するための形式に変換します。
 * 値が空でなく、かつ半角数字のみで構成されている場合、先頭にシングルクォーテーション「'」を付加します。
 * それ以外の場合は、トリムされた文字列をそのまま返します。
 * nullやundefinedは空文字列として扱います。
 * 
 * @param {any} value 元の値。
 * @return {string} フォーマットされた文字列。
 */
function formatValueForTextOutput_(value) {
  if (value === null || value === undefined) {
    return "";
  }
  const stringValue = value.toString().trim();

  if (stringValue === "") {
    return "";
  }

  if (/^\d+$/.test(stringValue)) {
    return "'" + stringValue;
  } else {
    return stringValue;
  }
}

/**
 * 設定シートの指定されたセルから設定値を取得します。
 * 値が空または取得できない場合は、ユーザーにエラーアラートを表示し、nullを返します。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} configSheet 設定シートのオブジェクト。
 * @param {string} cellA1Notation 設定値が記載されているセルのA1表記 (例: "B1", "C5")。
 * @param {string} settingDisplayName アラートメッセージに表示する設定項目の名前 (例: "出力先フォルダURL")。
 * @return {string|null} セルから取得した設定値 (文字列)。取得失敗時は null。
 * @private
 */
function getRequiredSettingOrAlert_(configSheet, cellA1Notation, settingDisplayName) {
  const value = getSettingFromCell_(configSheet, cellA1Notation);
  if (!value) {
    showErrorAlert(ERROR_TYPES.SETTING_MISSING, {
      settingName: settingDisplayName,
      cell: cellA1Notation
    });
    return null;
  }
  return value;
}

/**
 * 設定シートの指定されたセルから値を取得し、文字列として返します。
 * 値が空の場合は null を返します。値の取得中にエラーが発生した場合も null を返し、エラーをログに記録します。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} configSheet 設定シートのオブジェクト。
 * @param {string} cellA1Notation 値が記載されているセルのA1表記 (例: "B1", "C5")。
 * @return {string|null} セルから取得した設定値 (トリムされた文字列)。値が空またはエラー時は null。
 * @private
 */
function getSettingFromCell_(configSheet, cellA1Notation) {
  try {
    const value = configSheet.getRange(cellA1Notation).getValue().toString().trim();
    return value || null; // 空白の場合はnullを返す
  } catch (e) {
    Logger.log(`セル「${cellA1Notation}」からの値の取得中にエラー: ${e.message}`);
    return null;
  }
}

/**
 * 指定されたエラータイプと詳細に基づいて、ユーザーにエラーメッセージのアラートを表示します。
 * エラータイプに応じて、アラートのタイトルとメッセージの内容が動的に生成されます。
 * 未定義のエラータイプの場合は、汎用的なエラーメッセージを表示し、ログに記録します。
 *
 * @param {string} errorType エラーの種類を示す文字列。ERROR_TYPESオブジェクトのいずれかのキーを想定。
 * @param {object} [details={}] エラーメッセージに含める詳細情報。プロパティはエラータイプによって異なる。
 * @param {string} [details.settingName] (SETTING_MISSING, SETTING_INVALID) 設定項目の表示名。
 * @param {string} [details.cell] (SETTING_MISSING, SETTING_INVALID) 設定シート内のセル参照 (例: "B1")。
 * @param {string} [details.value] (SETTING_INVALID) 設定の現在の値やファイル名。
 * @param {string} [details.reason] (SETTING_INVALID, FILE_CONVERSION_ERROR, TARGET_FILE_STRUCTURE_ERROR) エラーの理由。
 * @param {string} [details.example] (SETTING_INVALID) 設定値の正しい例。
 * @param {string} [details.fileName] (FILE_CONVERSION_ERROR, TARGET_FILE_STRUCTURE_ERROR, COLUMN_NOT_FOUND) 関連するファイル名。
 * @param {string} [details.sheetName] (TARGET_FILE_STRUCTURE_ERROR, COLUMN_NOT_FOUND) 関連するシート名。
 * @param {string} [details.columns] (COLUMN_NOT_FOUND) 見つからなかった列名 (カンマ区切り)。
 * @param {string} [details.errorMessage] (UNEXPECTED_ERROR) 具体的なエラーメッセージ。
 * 
 * @return {void}
 */
function showErrorAlert(errorType, details = {}) {
  const ui = SpreadsheetApp.getUi();
  let title = "エラー";
  let message = "";

  switch (errorType) {
    case ERROR_TYPES.TARGET_FILE_STRUCTURE_ERROR:
      title = "読み取り対象ファイルの構造エラー";
      message = `読み取り対象のファイルまたはシートの構造に問題があります。\n`;
      if (details.fileName) message += `ファイル名: ${details.fileName}\n`;
      if (details.sheetName) message += `シート名: ${details.sheetName}\n`;
      if (details.reason) message += `\n問題点: ${details.reason}\n`;
      message += `\n対応：\nファイルの形式や、処理対象シート（"yymmdd"形式で、2行目からデータがあるかなど）を確認してください。`;
      break;
    
    // ... 他のエラーケースは変更なしのため、ここでは省略（元のコードのままでOK） ...
    case ERROR_TYPES.SETTING_MISSING:
      title = "設定の確認が必要です";
      message = `設定項目「${details.settingName || '不明な項目'}」が、設定シートのセル「${details.cell || '不明なセル'}」に記載されていません。\n\n対応：\n該当セルに必要な情報を入力してください。`;
      break;
    case ERROR_TYPES.SETTING_INVALID:
      title = "設定内容に誤りがあります";
      message = `設定項目「${details.settingName || '不明な項目'}」（セル「${details.cell || '不明なセル'}」）の記載内容が正しくありません。\n`;
      if (details.value) message += `現在の値/ファイル名: ${details.value}\n`;
      if (details.reason) message += `\n問題点:\n ${details.reason}\n`;
      else message += `\n問題点:\n 指定された情報（URLやIDなど）が正しくないか、アクセス権限がない可能性があります。\n`;
      message += `\n対応：\n設定内容を確認し、修正してください。`;
      if (details.example) message += `\n例： ${details.example}`;
      break;
    case ERROR_TYPES.FILE_CONVERSION_ERROR:
      title = "ファイル変換エラー";
      message = `ファイル「${details.fileName || '不明なファイル'}」の処理中に変換エラーが発生しました。\n`;
      if (details.reason) message += `\n問題点:\n ${details.reason}\n`;
      message += `\n対応：\nDrive APIサービスが有効になっていること、およびファイルの形式が破損していないことを確認してください。それでも解決しない場合は管理者に連絡してください。`;
      break;
    case ERROR_TYPES.COLUMN_NOT_FOUND:
      title = "必要な列が見つかりません";
      message = `ファイル「${details.fileName || '不明'}」内のシート「${details.sheetName || '不明なシート'}」から、必要なデータ列（${details.columns || '不明な列'}）のいずれかが見つかりませんでした。\n\n対応：\n読み取り対象ファイルの列名が正しいか確認してください。`;
      break;
    case ERROR_TYPES.UNEXPECTED_ERROR:
      title = "予期せぬエラー";
      message = `処理中に予期せぬエラーが発生しました。\n`;
      if (details.errorMessage) message += `エラー詳細:\n ${details.errorMessage}\n`;
      message += `\n対応：\n問題が解決しない場合は、エラー詳細を管理者に連絡してください。`;
      break;
    default:
      title = "不明なエラー";
      message = `不明なエラーが発生しました。詳細はログを確認してください。`;
      if (errorType) message += ` (エラータイプ: ${errorType})`;
      Logger.log(`showErrorAlert に不明な errorType: ${errorType}, details: ${JSON.stringify(details)}`);
  }
  ui.alert(title, message, ui.ButtonSet.OK);
}


/**
 * スプレッドシートから最新の日付を持つ対象シート ("yymmdd"形式) を取得します。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss スプレッドシートオブジェクト
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null} 最新のシートオブジェクト、またはnull
 */
function getLatestTargetSheet_(ss) {
  const sheets = ss.getSheets();
  const targetSheets = sheets.filter(sheet => {
    const name = sheet.getName();
    return /^\d{6}$/.test(name); // "yymmdd" 形式 (例: 250609) のシート名を正規表現で探す
  });

  if (targetSheets.length === 0) return null;

  // シート名を降順にソートして最新のものを取得 (例: "250609" > "250608")
  targetSheets.sort((a, b) => b.getName().localeCompare(a.getName()));
  return targetSheets[0];
}

/**
 * ヘッダー行の配列を受け取り、定義済みの必須カラム名 (TARGET_COLUMNS) に基づいて、
 * 各カラムの0ベースのインデックスを特定します。
 *
 * @param {Array<string>} headerRow ヘッダー行の文字列配列。各要素が列名を表します。
 * @return {{ID: number, STUDENT_CODE: number, ACTION_TYPE: number}}
 * 必須カラムのキー (ID, STUDENT_CODE, ACTION_TYPE) と、
 * それぞれに対応するヘッダー行内でのインデックス (見つからない場合は -1) を持つオブジェクト。
 * @private
 */
function getColumnIndices_(headerRow) {
  const indices = { ID: -1, STUDENT_CODE: -1, ACTION_TYPE: -1 };
  headerRow.forEach((header, index) => {
    if (header === TARGET_COLUMNS.ID) indices.ID = index;
    else if (header === TARGET_COLUMNS.STUDENT_CODE) indices.STUDENT_CODE = index;
    else if (header === TARGET_COLUMNS.ACTION_TYPE) indices.ACTION_TYPE = index;
  });
  return indices;
}


/**
 * 読み取ったデータを処理し、「連携」「連携対象外」「休会」の3つのデータに分類します。
 * @param {any[][]} dataRows データ行の2次元配列 (ヘッダー除く)
 * @param {object} colIndices カラムインデックスオブジェクト
 * @return {{linkageData: any[][], notLinkageData: any[][], kyukaiData: any[][]}} 分類された3つのデータセット
 */
function processData_(dataRows, colIndices) {
  const linkageData = [];
  const notLinkageData = [];
  const kyukaiData = [];

  dataRows.forEach(row => {
    const rawStudentCode = row[colIndices.STUDENT_CODE] ? row[colIndices.STUDENT_CODE].toString().trim() : "";
    const rawId = row[colIndices.ID] ? row[colIndices.ID].toString().trim() : "";
    const actionType = row[colIndices.ACTION_TYPE] ? row[colIndices.ACTION_TYPE].toString().trim() : "";

    let studentCodeForOutput;
    let idForOutput;

    switch (actionType) {
      case ACTION_TYPES.AUTO_LINK:
        if (rawStudentCode === "" || rawId === "") break;
        studentCodeForOutput = formatValueForTextOutput_(rawStudentCode);
        idForOutput = formatValueForTextOutput_(rawId);
        linkageData.push([studentCodeForOutput, idForOutput]);
        break;

      case ACTION_TYPES.AUTO_LINK_EXCLUDE:
        if (rawStudentCode === "") break;
        studentCodeForOutput = formatValueForTextOutput_(rawStudentCode);
        IdForOutput = formatValueForTextOutput_(rawId);
        notLinkageData.push([studentCodeForOutput, IdForOutput]);
        break;

      case ACTION_TYPES.KYUKAI:
        if (rawId === "") break;
        idForOutput = formatValueForTextOutput_(rawId);
        kyukaiData.push([idForOutput]);
        break;

      case ACTION_TYPES.DO_NOTHING:
        break;

      default:
        Logger.log(`不明な対応内容: ${actionType} (行データ: ${row.join(", ")})`);
        break;
    }
  });

  return { linkageData, notLinkageData, kyukaiData };
}


/**
 * 【修正】新規スプレッドシートを作成し、3種類の処理済みデータをそれぞれのシートに書き込みます。
 * @param {GoogleAppsScript.Drive.Folder} outputFolder 出力先フォルダオブジェクト
 * @param {string} readSheetNameFromSource 読み取り対象だったシート名 (yymmdd形式)。出力ファイル名生成に使用。
 * @param {any[][]} linkageData 「連携」シート用データ
 * @param {any[][]} notLinkageData 「連携対象外」シート用データ
 * @param {any[][]} kyukaiData 「休会」シート用データ
 * @return {string|null} 作成されたスプレッドシートのURL、失敗時はnull
 */
function createAndPopulateOutputSpreadsheet_(outputFolder, readSheetNameFromSource, linkageData, notLinkageData, kyukaiData) {
  let newSsId = null;
  try {
    // 出力ファイル名は読み取りシート名から直接生成 (例: 振分結果_250609)
    const outputFileName = `${OUTPUT_FILE_PREFIX_NEW}${readSheetNameFromSource}`;

    const newSs = SpreadsheetApp.create(outputFileName);
    newSsId = newSs.getId();
    const newFile = DriveApp.getFileById(newSsId);
    newFile.moveTo(outputFolder);

    // --- 1. 「連携」シートの処理 ---
    const linkageSheet = newSs.getSheetByName("シート1");
    linkageSheet.setName(OUTPUT_SHEET_LINKAGE);
    linkageSheet.clearContents();
    const linkageHeaders = [[TARGET_COLUMNS.STUDENT_CODE, TARGET_COLUMNS.ID]];
    linkageSheet.getRange(1, 1, 1, 2).setValues(linkageHeaders).setFontWeight("bold");
    if (linkageData && linkageData.length > 0) {
      linkageSheet.getRange(2, 1, linkageData.length, linkageData[0].length).setValues(linkageData);
    } else {
      linkageSheet.getRange(2, 1).setValue("該当データなし");
    }

    // --- 2. 「連携対象外」シートの処理 ---
    const notLinkageSheet = newSs.insertSheet(OUTPUT_SHEET_NOT_LINKAGE);
    notLinkageSheet.clearContents();
    const notLinkageHeaders = [[TARGET_COLUMNS.STUDENT_CODE, TARGET_COLUMNS.ID]];
    notLinkageSheet.getRange(1, 1, 1, 2).setValues(notLinkageHeaders).setFontWeight("bold");
    if (notLinkageData && notLinkageData.length > 0) {
      notLinkageSheet.getRange(2, 1, notLinkageData.length, notLinkageData[0].length).setValues(notLinkageData);
    } else {
      notLinkageSheet.getRange(2, 1).setValue("該当データなし");
    }

    // --- 3. 「休会」シートの処理 ---
    const kyukaiSheet = newSs.insertSheet(OUTPUT_SHEET_KYUKAI);
    kyukaiSheet.clearContents();
    const kyukaiHeaders = [[TARGET_COLUMNS.ID]];
    kyukaiSheet.getRange(1, 1, 1, 1).setValues(kyukaiHeaders).setFontWeight("bold");
    if (kyukaiData && kyukaiData.length > 0) {
      kyukaiSheet.getRange(2, 1, kyukaiData.length, kyukaiData[0].length).setValues(kyukaiData);
    } else {
      kyukaiSheet.getRange(2, 1).setValue("該当データなし");
    }

    return newSs.getUrl();
  } catch (e) {
    Logger.log(`出力スプレッドシート作成・書き込みエラー: ${e.message}\n${e.stack}`);
    if (newSsId) {
      try {
        DriveApp.getFileById(newSsId).setTrashed(true);
        Logger.log(`エラー発生のため、作成途中だったファイル (ID: ${newSsId}) をゴミ箱に移動しました。`);
      } catch (errTrash) {
        Logger.log(`作成途中ファイルの削除失敗 (ID: ${newSsId}): ${errTrash.message}`);
      }
    }
    return null;
  }
}


/**
 * 実行履歴をログシートに記録します。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss スクリプトホストファイル
 * @param {string} readSheetName 読み取ったシート名 (yymmdd形式)
 * @param {string} newSpreadsheetUrl 作成されたスプレッドシートのURL
 */
function logExecution_(ss, readSheetName, newSpreadsheetUrl) {
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    logSheet.appendRow(["実行日時", "読み取りシート名", "作成ファイルURL"]);
    logSheet.setFrozenRows(1);
  }
  const executionTime = new Date();
  const formattedExecutionTime = Utilities.formatDate(executionTime, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  const displaySheetName = "'" + readSheetName;
  logSheet.appendRow([formattedExecutionTime, displaySheetName, newSpreadsheetUrl]);
}
