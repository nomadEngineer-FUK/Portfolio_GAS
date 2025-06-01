const CONFIG = {
  SALES_SHEET_NAME: "DB",
  OUTPUT_SHEET_NAME_SINGLE: "Summary_Single",
  OUTPUT_SHEET_NAME_CROSS: "Summary_Cross",

  HEADERS: {
    MONTHLY: { text: ["年月", "件数", "売上合計"], cols: 3 },
    CATEGORY: { text: ["カテゴリ", "件数", "売上合計"], cols: 3 },
    CHANNEL: { text: ["販売経路", "件数", "売上合計"], cols: 3 },
    CUSTOMER: { text: ["区分", "名称", "件数", "金額"], cols: 4 },
  },
  COLORS: {
    LIGHT_GRAY: "#F0F0F0"
  }
};