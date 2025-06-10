// --- 定数定義 ---
const CONFIG_SHEET_NAME = "設定";     // 設定を記載するシート名
const SOURCE_FILE_URL_CELL = "B1";   // 設定シート内の読み取り対象ファイルURL記載セル
const OUTPUT_FOLDER_URL_CELL = "B2"; // 設定シート内の出力先フォルダの【リンク(URL)専用】記載セル
const LOG_SHEET_NAME = "実行履歴";    // 実行履歴を記録するシート名

// 読み取り対象のカラム名
const TARGET_COLUMNS = {
  ID: "ID",
  STUDENT_CODE: "生徒コード",
  ACTION_TYPE: "対応内容"
};

// 対応内容の種別
const ACTION_TYPES = {
  AUTO_LINK: "自動連携対象",
  AUTO_LINK_EXCLUDE: "自動連携対象外",
  KYUKAI: "休会",
  DO_NOTHING: "何もしない"
};

// 出力ファイル関連
const OUTPUT_FILE_PREFIX_NEW = "振分結果_";
const OUTPUT_SHEET_LINKAGE = "連携";
const OUTPUT_SHEET_NOT_LINKAGE = "連携対象外";
const OUTPUT_SHEET_KYUKAI = "休会";

// エラータイプの定義
const ERROR_TYPES = {
  SETTING_MISSING: 'SETTING_MISSING',
  SETTING_INVALID: 'SETTING_INVALID',
  TARGET_FILE_STRUCTURE_ERROR: 'TARGET_FILE_STRUCTURE_ERROR',
  COLUMN_NOT_FOUND: 'COLUMN_NOT_FOUND',
  UNEXPECTED_ERROR: 'UNEXPECTED_ERROR',
  FILE_CONVERSION_ERROR: 'FILE_CONVERSION_ERROR' // Excelからスプレッドシートへの変換エラー
};
