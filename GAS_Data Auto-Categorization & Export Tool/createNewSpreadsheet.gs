/**
 * メイン処理関数：全体の処理フローを制御します。
 * 
 * このスクリプトは Drive API を使用します。
 * Apps Script エディタの「サービス」の横にある「＋」アイコンをクリックし、
 * 「Drive API」を追加して有効にしてください。(バージョンは v2 を想定しています)
 */

/**
 * @description
 * データ処理のメインプロセスを実行します。
 * この関数は、設定の読み込み、ソースデータの準備、データ処理、結果の出力、
 * 実行ログの記録、後処理（一時ファイルの削除）までの一連の流れを統括します。
 *
 * 主な処理の流れは以下の通りです。
 * 1. 設定シートから設定情報を読み込み、検証します。
 * 2. ソースデータを準備します。これにはExcelファイルのGoogleスプレッドシートへの変換が含まれる場合があります。
 * 3. 準備したデータを加工します。
 * 4. 加工後のデータを新しいスプレッドシートに出力します。
 * 5. 処理結果をログシートに記録し、ユーザーにUIで通知します。
 *
 * @throws {Error} 処理中に予期せぬエラーが発生した場合、エラー内容をログに出力し、
 * ユーザーにアラートで通知します。エラーの例としては、設定の不備、ファイルの読み取り失敗、
 * APIの制限超過などが考えられます。
 *
 * @returns {void} この関数は直接の戻り値を持ちませんが、処理の最後に成功または失敗を示すUIアラートを表示します。
 *
 * @see {@link getConfigSheetOrAlert_}
 * @see {@link loadAndValidateSettings_}
 * @see {@link prepareSourceData_}
 * @see {@link processData_}
 * @see {@link createAndPopulateOutputSpreadsheet_}
 * @see {@link logExecution_}
 * @see {@link showErrorAlert}
 */
function mainProcessData() {
  let tempConvertedSpreadsheetId = null; // 一時変換されたスプレッドシートのIDを格納

  try {
    const scriptHostSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = getConfigSheetOrAlert_(scriptHostSpreadsheet);
    if (!configSheet) return;

    const settings = loadAndValidateSettings_(configSheet);
    if (!settings) return;
    tempConvertedSpreadsheetId = settings.tempConvertedSpreadsheetId; // クリーンアップ用にIDを保持

    // sourceDataInfo には読み取り対象となったシート名(yymmdd)が含まれる
    const sourceDataInfo = prepareSourceData_(settings.sourceSpreadsheet, settings.originalFileName);
    if (!sourceDataInfo) return;

    // データを「連携」「連携対象外」「休会」の3つに分類
    const processedData = processData_(sourceDataInfo.dataRows, sourceDataInfo.columnIndices);

    // 出力ファイル生成。ファイル名には元のファイル名ではなく、読み取ったシート名を使用
    const newSpreadsheetUrl = createAndPopulateOutputSpreadsheet_(
      settings.outputFolder,
      sourceDataInfo.targetSheetName, // yymmdd形式のシート名
      processedData.linkageData,
      processedData.notLinkageData,
      processedData.kyukaiData
    );

    if (newSpreadsheetUrl) {
      logExecution_(scriptHostSpreadsheet, sourceDataInfo.targetSheetName, newSpreadsheetUrl);
      SpreadsheetApp.getUi().alert("成功", `処理が完了しました。\n作成されたファイル: ${newSpreadsheetUrl}`, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      showErrorAlert(ERROR_TYPES.UNEXPECTED_ERROR, {
        errorMessage: "出力ファイルの作成に失敗しました。詳細はログを確認してください。"
      });
    }

  } catch (e) {
    showErrorAlert(ERROR_TYPES.UNEXPECTED_ERROR, {
      errorMessage: `メイン処理でエラーが発生しました: ${e.message}` + (e.stack ? `\nスタックトレース: ${e.stack}` : '')
    });
    Logger.log(`致命的なエラーが発生しました: ${e.message}\n${e.stack}`);
  } finally {
    if (tempConvertedSpreadsheetId) {
      try {
        Logger.log(`一時変換されたスプレッドシート (ID: ${tempConvertedSpreadsheetId}) を削除します。`);
        DriveApp.getFileById(tempConvertedSpreadsheetId).setTrashed(true);
        Logger.log(`一時変換されたスプレッドシート (ID: ${tempConvertedSpreadsheetId}) をゴミ箱に移動しました。`);
      } catch (trashError) {
        Logger.log(`一時変換されたスプレッドシート (ID: ${tempConvertedSpreadsheetId}) の削除に失敗しました: ${trashError.message}`);
      }
    }
  }
}



// --- メイン処理からのヘルパー関数群 ---
/**
 * 現在アクティブなスプレッドシート（スクリプトホストファイル）から、
 * 定数 CONFIG_SHEET_NAME で定義された名前の設定シートを取得します。
 * 設定シートが見つからない場合は、ユーザーにアラートを表示し、エラーログを記録して null を返します。
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} scriptHostSpreadsheet スクリプトが実行されているホストスプレッドシートオブジェクト。
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null} 見つかった設定シートオブジェクト。見つからない場合は null。
 * @private
 */
function getConfigSheetOrAlert_(scriptHostSpreadsheet) {
  const configSheet = scriptHostSpreadsheet.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) {
    SpreadsheetApp.getUi().alert("設定エラー", `スクリプトホストファイルに「${CONFIG_SHEET_NAME}」シートが見つかりません。\n作成してB1セルに読み取り対象ファイルのURL、B2セルに出力先フォルダURLを記載してください。`, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(`設定シート「${CONFIG_SHEET_NAME}」が見つかりません。`);
    return null;
  }
  return configSheet;
}


/**
 * 設定シートから各種設定を読み込み、検証し、処理に必要なオブジェクトを準備します。
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} configSheet 設定シートオブジェクト
 * @return {object|null} 処理に必要な設定情報オブジェクト、またはエラー時は null
 */
function loadAndValidateSettings_(configSheet) {
  const outputFolder = getOutputFolderFromSettings_(configSheet);
  if (!outputFolder) return null;

  const sourceFileInfo = getSourceFileInfoFromSettings_(configSheet);
  if (!sourceFileInfo) return null;

  const spreadsheetPreparationResult = prepareSourceSpreadsheet_(sourceFileInfo);
  if (!spreadsheetPreparationResult) return null;

  return {
    outputFolder: outputFolder,
    sourceSpreadsheet: spreadsheetPreparationResult.spreadsheet,
    tempConvertedSpreadsheetId: spreadsheetPreparationResult.tempConvertedSpreadsheetId,
    originalFileName: sourceFileInfo.name // 元のファイル名を引き続き使用
  };
}

/**
 * 設定シートから出力先フォルダの情報を読み込み、検証してフォルダオブジェクトを返します。
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} configSheet 設定シートオブジェクト
 * @return {GoogleAppsScript.Drive.Folder|null} Folderオブジェクト、またはエラー時は null
 */
function getOutputFolderFromSettings_(configSheet) {
  const outputFolderUrl = getRequiredSettingOrAlert_(configSheet, OUTPUT_FOLDER_URL_CELL, "出力先フォルダのリンク(URL)");
  if (!outputFolderUrl) return null;

  const lowerOutputFolderUrl = outputFolderUrl.toLowerCase();

  if (!(lowerOutputFolderUrl.startsWith("https://docs.google.com/drive/folders/") ||
        lowerOutputFolderUrl.startsWith("https://drive.google.com/drive/folders/"))) {
    showErrorAlert(ERROR_TYPES.SETTING_INVALID, {
      settingName: "出力先フォルダのリンク(URL)", cell: OUTPUT_FOLDER_URL_CELL, value: outputFolderUrl,
      reason: "Google DriveのフォルダURLの形式が正しくありません。",
      example: "https://drive.google.com/drive/folders/ID または https://docs.google.com/drive/folders/ID"
    });
    return null;
  }

  let outputFolderId;
  const folderMatch = outputFolderUrl.match(/folders\/([a-zA-Z0-9_-]+)/);
  if (folderMatch && folderMatch[1]) {
    outputFolderId = folderMatch[1];
  } else {
    showErrorAlert(ERROR_TYPES.SETTING_INVALID, {
      settingName: "出力先フォルダのリンク(URL)", cell: OUTPUT_FOLDER_URL_CELL, value: outputFolderUrl,
      reason: "URLからフォルダIDを正しく抽出できませんでした。"
    });
    return null;
  }

  try {
    const outputFolder = DriveApp.getFolderById(outputFolderId);
    Logger.log(`出力先フォルダ: ${outputFolder.getName()} (ID: ${outputFolderId})`);
    return outputFolder;

  } catch (e) {
    showErrorAlert(ERROR_TYPES.SETTING_INVALID, {
      settingName: "出力先フォルダ", cell: OUTPUT_FOLDER_URL_CELL,
      value: `(抽出ID: ${outputFolderId}, 元URL: ${outputFolderUrl})`,
      reason: `指定されたフォルダが開けませんでした。URLが正しいか、アクセス権があるか確認してください。\n\n詳細:\n ${e.message}`
    });
    return null;
  }
}

/**
 * 設定シートから読み取り対象ファイルの情報を読み込み、検証してファイル関連情報を返します。
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} configSheet 設定シートオブジェクト
 * @return {{file: GoogleAppsScript.Drive.File, id: string, name: string, mimeType: string, url: string}|null}
 * ファイルオブジェクト、ID、名前、MIMEタイプ、元のURLを含むオブジェクト、またはエラー時は null
 */
function getSourceFileInfoFromSettings_(configSheet) {
  const sourceFileUrl = getRequiredSettingOrAlert_(configSheet, SOURCE_FILE_URL_CELL, "読み取り対象ファイルのURL");
  if (!sourceFileUrl) return null;

  let sourceFileId;
  const fileIdMatch = sourceFileUrl.match(/[-\w]{25,}/);

  if (fileIdMatch && fileIdMatch[0]) {
    sourceFileId = fileIdMatch[0];

  } else {
    showErrorAlert(ERROR_TYPES.SETTING_INVALID, {
      settingName: "読み取り対象ファイルのURL", cell: SOURCE_FILE_URL_CELL, value: sourceFileUrl,
      reason: "URLからファイルIDを正しく抽出できませんでした。Google Driveのファイル共有リンクを使用してください。"
    });
    return null;
  }

  try {
    const sourceFile = DriveApp.getFileById(sourceFileId);
    const originalFileName = sourceFile.getName();
    const mimeType = sourceFile.getMimeType();
    Logger.log(`読み取り対象ファイル: ${originalFileName} (ID: ${sourceFileId}, MIMEタイプ: ${mimeType})`);
    return { file: sourceFile, id: sourceFileId, name: originalFileName, mimeType: mimeType, url: sourceFileUrl };

  } catch (e) {
    showErrorAlert(ERROR_TYPES.SETTING_INVALID, {
      settingName: "読み取り対象ファイル", cell: SOURCE_FILE_URL_CELL, value: sourceFileUrl,
      reason: `指定されたファイル (ID: ${sourceFileId}) が開けませんでした。URLが正しいか、アクセス権があるか確認してください。\n\n詳細:\n ${e.message}`
    });
    return null;
  }
}

/**
 * 提供されたファイル情報に基づき、処理可能なスプレッドシートオブジェクトを準備します。
 * Excelの場合は変換処理を行い、Googleスプレッドシートの場合は直接開きます。
 * 
 * @param {{file: GoogleAppsScript.Drive.File, id: string, name: string, mimeType: string, url: string}} fileInfo ファイル情報オブジェクト
 * @return {{spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, tempConvertedSpreadsheetId: string|null}|null}
 * Spreadsheetオブジェクトと一時変換ID (該当する場合) を含むオブジェクト、またはエラー時は null
 */
function prepareSourceSpreadsheet_(fileInfo) {
  let spreadsheet = null;
  let tempConvertedSpreadsheetId = null;
  const { file, id: sourceFileId, name: originalFileName, mimeType, url: sourceFileUrl } = fileInfo;

  if (mimeType === MimeType.MICROSOFT_EXCEL ||
      mimeType === MimeType.MICROSOFT_EXCEL_X ||
      mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
    Logger.log(`Excelファイル (${mimeType}) を検出しました。「${originalFileName}」をGoogleスプレッドシート形式に変換します。`);

    try {
      const parents = file.getParents(); // fileInfo.file を使用
      if (!parents.hasNext()) {
        Logger.log(`エラー: ソースファイル (ID: ${sourceFileId}, 名前: ${originalFileName}) に親フォルダが見つかりません。`);
        showErrorAlert(ERROR_TYPES.FILE_CONVERSION_ERROR, {
            fileName: originalFileName,
            reason: `ソースファイルの親フォルダ情報を取得できませんでした。ファイルの場所やアクセス権を確認してください。`
        });
        return null;
      }

      const parentFolderId = parents.next().getId();
      Logger.log(`ソースファイル「${originalFileName}」の親フォルダID: ${parentFolderId}`);

      Logger.log(`Drive API を使用してファイルメタデータを取得試行 (ID: ${sourceFileId})`);
      const fileMetadataByDriveAPI = Drive.Files.get(sourceFileId, { supportsAllDrives: true });
      Logger.log(`Drive API によるファイルメタデータ取得成功: ${fileMetadataByDriveAPI.title} (ID: ${fileMetadataByDriveAPI.id})`);

      const tempSpreadsheetName = `_temp_conversion_${originalFileName}_${new Date().getTime()}`;
      const convertedFileResource = {
        title: tempSpreadsheetName,
        mimeType: MimeType.GOOGLE_SHEETS,
        parents: [{ id: parentFolderId }]
      };

      const optionalArgs = { supportsAllDrives: true };
      Logger.log(`Drive.Files.copy を呼び出します。Resource: ${JSON.stringify(convertedFileResource)}, Source ID: ${sourceFileId}, OptionalArgs: ${JSON.stringify(optionalArgs)}`);

      const convertedFile = Drive.Files.copy(convertedFileResource, sourceFileId, optionalArgs);
      tempConvertedSpreadsheetId = convertedFile.id;
      spreadsheet = SpreadsheetApp.openById(tempConvertedSpreadsheetId);
      Logger.log(`ファイル「${originalFileName}」を一時スプレッドシート「${tempSpreadsheetName}」(ID: ${tempConvertedSpreadsheetId}) に変換しました。保存場所の親ID: ${parentFolderId}`);

    } catch (e) {
      Logger.log(`Excel「${originalFileName}」からスプレッドシートへの変換中にエラー: ${e.message}\n${e.stack}`);

      let reasonDetail = `ExcelファイルのGoogleスプレッドシートへの変換に失敗しました。\n\n詳細:\n ${e.message}`;

      if (e.message.includes(sourceFileId) && e.message.toLowerCase().includes("file not found")) {
          reasonDetail = `Drive API が指定されたファイルID (${sourceFileId}) でファイルを見つけられませんでした (共有ドライブ内のファイル)。\n\nエラー詳細:\n ${e.message}`;

      } else if (e.message.toLowerCase().includes("does not support importing from this file type") || e.message.toLowerCase().includes("unable to convert")) {
          reasonDetail = `指定されたExcelファイル (${originalFileName}) は、変換がサポートされていない形式か、ファイル内容に問題がある可能性があります。\n\nエラー詳細:\n ${e.message}`;
      }

      showErrorAlert(ERROR_TYPES.FILE_CONVERSION_ERROR, { fileName: originalFileName, reason: reasonDetail });
      return null;
    }

  } else if (mimeType === MimeType.GOOGLE_SHEETS) {
    Logger.log(`Googleスプレッドシート「${originalFileName}」を直接開きます。`);

    try {
      spreadsheet = SpreadsheetApp.openById(sourceFileId);

    } catch (e) {
      showErrorAlert(ERROR_TYPES.SETTING_INVALID, {
        settingName: "読み取り対象スプレッドシート", cell: SOURCE_FILE_URL_CELL, value: sourceFileUrl,
        reason: `指定されたスプレッドシート「${originalFileName}」が開けませんでした。\n\n詳細:\n ${e.message}`
      });
      return null;
    }

  } else {
    showErrorAlert(ERROR_TYPES.SETTING_INVALID, {
      settingName: "読み取り対象ファイルの形式", cell: SOURCE_FILE_URL_CELL, value: originalFileName,
      reason: `対応していないファイル形式です (${mimeType})。Excelファイル (xlsx, xls) または Googleスプレッドシートを指定してください。`
    });
    return null;
  }

  if (spreadsheet) {
    Logger.log(`読み取りに使用するスプレッドシート (変換後または直接): ${spreadsheet.getName()} (URL: ${spreadsheet.getUrl()})`);
    return { spreadsheet: spreadsheet, tempConvertedSpreadsheetId: tempConvertedSpreadsheetId };
  }
  // ここには到達しないはずだが念のため
  Logger.log(`prepareSourceSpreadsheet_ の終端で spreadsheet オブジェクトが null です。mimeType: ${mimeType}`);
  return null;
}


/**
 * 読み取り対象のデータを準備します。最新シートの特定、ヘッダー検証など。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sourceSpreadsheet 処理対象のスプレッドシート
 * @param {string} originalSourceFileName 元のファイル名 (エラー表示用)
 * @return {object|null} 成功時は {targetSheetName, dataRows, columnIndices}, 失敗時は null
 */
function prepareSourceData_(sourceSpreadsheet, originalSourceFileName) {
  const targetSheet = getLatestTargetSheet_(sourceSpreadsheet);
  if (!targetSheet) {
    showErrorAlert(ERROR_TYPES.TARGET_FILE_STRUCTURE_ERROR, {
      fileName: originalSourceFileName,
      reason: `読み取り対象のシート（"yymmdd"形式のシート名）が見つかりませんでした。ファイル「${originalSourceFileName}」の内容を確認してください。`
    });
    return null;
  }
  const targetSheetName = targetSheet.getName();
  Logger.log(`読み取り対象シート: ${targetSheetName} (元ファイル: ${originalSourceFileName})`);

  const sheetData = targetSheet.getDataRange().getValues();
  if (sheetData.length < 2) {
    showErrorAlert(ERROR_TYPES.TARGET_FILE_STRUCTURE_ERROR, {
      fileName: originalSourceFileName,
      sheetName: targetSheetName,
      reason: `シート「${targetSheetName}」にヘッダー行以外の処理可能なデータがありません。`
    });
    return null;
  }
  const headerRow = sheetData[0];

  const columnIndices = getColumnIndices_(headerRow);
  const missingCols = [];
  if (columnIndices.ID === -1) missingCols.push(TARGET_COLUMNS.ID);
  if (columnIndices.STUDENT_CODE === -1) missingCols.push(TARGET_COLUMNS.STUDENT_CODE);
  if (columnIndices.ACTION_TYPE === -1) missingCols.push(TARGET_COLUMNS.ACTION_TYPE);

  if (missingCols.length > 0) {
    showErrorAlert(ERROR_TYPES.COLUMN_NOT_FOUND, {
      fileName: originalSourceFileName,
      sheetName: targetSheetName,
      columns: missingCols.join(", ")
    });
    return null;
  }

  return {
    targetSheetName: targetSheetName,
    dataRows: sheetData.slice(1), // ヘッダー行を除いたデータ
    columnIndices: columnIndices
  };
}
