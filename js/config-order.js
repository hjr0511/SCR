/**
 * 秩序評分系統前端設定（生輔組專用試算表）
 *
 * 請把「秩序」Google 試算表裡 Apps Script 網頁應用程式網址貼到 GAS_WEB_APP_URL。
 * 不要填整潔系統的網址。
 *
 * 後端部署步驟：
 * 1. 開一份新的 Google 試算表（可複製整潔的「教室清單」）
 * 2. 擴充功能 → Apps Script，貼上專案裡的 Code-order.gs 後儲存
 * 3. 部署 → 新增部署 → 類型選「網頁應用程式」
 * 4. 執行身分：我
 * 5. 存取權：任何人
 * 6. 把部署網址貼到下方
 */
window.APP_CONFIG = {
  GAS_WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbxTOl-OX-LpDhYUMkoFoVJc5Zp3cXUwTiv9EJxwr0xo-tGmV5Mv9Hb-krgqDfBVQUVTUA/exec',
  AUTH_STORAGE_KEY: 'ORDER_SCORE_AUTH',
  LOGIN_TYPE_KEY: 'ORDER_SCORE_LOGIN_TYPE',
  CONFIG_FILE: 'js/config-order.js'
};
