/**
 * GitHub Pages 前端設定
 *
 * 請把 Google Apps Script「網頁應用程式」網址貼到 GAS_WEB_APP_URL。
 *
 * 後端部署步驟：
 * 1. 在綁定本系統的 Google 試算表開啟「擴充功能 → Apps Script」
 * 2. 貼上 / 更新 Code.gs 後儲存
 * 3. 部署 → 新增部署 → 類型選「網頁應用程式」
 * 4. 執行身分：我
 * 5. 存取權：任何人（GitHub 靜態頁才能呼叫）
 * 6. 授權試算表、Drive、文件權限
 * 7. 把部署網址貼到下方（格式：https://script.google.com/macros/s/xxxxx/exec）
 */
window.APP_CONFIG = {
  GAS_WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbxosRh8HQ3axllsry4jjKb2k5c8Bo6bGzTQA_w82BBJx4hDC6YqRSmJvzA2XYqRzgZv/exec'
};
