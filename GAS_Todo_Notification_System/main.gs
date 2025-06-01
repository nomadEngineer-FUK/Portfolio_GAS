function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

/**
 * メイン関数
 * 
 * 1. 有効な範囲のデータを取得
 * 2. バリデーション（C列：種別 と G列：難易度）
 * 3. 「満点」「得点」「該当のステージ」を取得
 * 
 */
function getProgressData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('todo');
  const data = getValidRangeData(sheet);

  // B列とF列のいずれかが空欄のセルがあれば、該当行を返す
  // 全て値が入力されていれば[0]を返す
  const resultOfValidation = validateInput(data);

  if (resultOfValidation) {
    const msg = `${resultOfValidation} 行目が未入力です。処理を中止します。`;
    Logger.log(msg);
    throw new Error(msg);
  }

  return calculateScore(data);
}

/**
 * 有効なデータ範囲（※1）を取得
 * 
 * ※1 有効なデータ範囲
 * 　A2セル～I列の最終行（※2）
 * 
 * ※2 最終行
 * 　C列またはG列における、データが存在する一番下の行
 * 　（それ以外の列のデータは最終行の判定には用いない）
 * 
 */
function getValidRangeData(sheet) {
  const values = sheet.getRange("A2:H").getValues();

  // C列 と G列のいずれかにデータがある行を抽出
  const filtered = values
    .map((row, idx) => ({ row, idx}))
    .filter(item => item.row[2] || item.row[6]);

    if (filtered.length === 0) return [];

    const lastIndex = filtered[filtered.length - 1].idx + 1;
    return sheet.getRange(2, 1, lastIndex, 8).getValues();
}

/**
 * 入力のバリデーション
 * 
 * C列もしくはG列で、空欄のセルがある場合に処理を中断
 * 
 */
function validateInput(data) {

  // バリデーション：B列とF列がすべて埋まっているか確認
  for (let i = 0; i < data.length; i++) {

    const row = data[i];
    const typeForValidation = row[2];
    const difficultyForValidation = row[6];

    // 未入力の行数をreturn（複数の未入力がある場合、最初に見つかった行だけreturnする）
    if (!typeForValidation || !difficultyForValidation) return i + 2; 
  }
  return 0; // 全て正しく入力されている場合
}

/**
 * Todo全体の点数と完了済みTodoの点数を取得
 */
function calculateScore(data) {

  let total = 0;
  let done = 0;

  for (let i = 0; i < data.length; i++) {
    const currentRow = data[i];

    // C列が空白の行はスキップする
    if (!currentRow[2]) continue;

    const isDone = currentRow[0]; // A列のチェックボックス（完了/未完了）
    const difficulty = currentRow[6]; // G列の難易度

    const score = convertDifficultyToScore(difficulty); // 難易度に応じた配点を取得

    // スコアが0の場合（難易度が設定されていないなど）も加算しない
    if (score === 0 && difficulty !== "") { // 難易度が空白でなく、スコアが0の場合はエラーの可能性があるのでログに出す
        Logger.log(`警告: Difficulty "${difficulty}" in row ${i + 2} resulted in a score of 0.`);
    }

    total += score; // 全体の合計値
    
    // isDone が true（チェックボックスがオン）の場合に加算
    if (isDone === true) done += score;
  }

  // 進捗率
  const rawProgress = total === 0 ? 0 : done / total; // 0 ～ 1
  const progress = Math.round(rawProgress * 100); // 0％～100%
  
  return {
    progress,
    stageKey: getStageKey(rawProgress),
    displayStage: getStage(rawProgress)
  };
}

/**
 * 難易度別に点数を返す
 */
function convertDifficultyToScore(difficulty) {

  if (difficulty === "★☆☆") return 1;
  if (difficulty === "★★☆") return 2;
  if (difficulty === "★★★") return 3;

  return 0;
}

/**
 * ステージをセット
 */
function getStageKey(progress) {
  if (progress < 0.2) return "Stage 1";
  if (progress < 0.4) return "Stage 2";
  if (progress < 0.6) return "Stage 3";
  if (progress < 0.8) return "Stage 4";
  
  return "Complete";
}

/**
 * アプリ上に表示する文言を取得
 */
function getStage(progress) {
  const key = getStageKey(progress);
  if (key === "Complete") return "🎉 Congratulations!! You've completed all tasks! 🎉";
  return key;
}