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
  GAS_WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbwT9n_mkDX_uYK4oQqk83wLV2u4z65ycd9KceQ1jNyt_K7IPL9ZE5EzWx16-bwoMZ0TsQ/exec',
  AUTH_STORAGE_KEY: 'SCHOOL_SCORE_AUTH',
  LOGIN_TYPE_KEY: 'SCHOOL_SCORE_LOGIN_TYPE',
  CONFIG_FILE: 'js/config.js'
};

/** 整潔缺失項目。每項扣 1 分。1D 為資源、一般垃圾桶。 */
window.CLEAN_SCORE_ITEMS = [
  { code: '1A1', area: '教室區', name: '講桌、講台' },
  { code: '1A2', area: '教室區', name: '黑板' },
  { code: '1B', area: '教室區', name: '桌椅' },
  { code: '1C', area: '教室區', name: '教室地板' },
  { code: '1D', area: '教室區', name: '資源、一般垃圾桶' },
  { code: '1E', area: '教室區', name: '走廊地板（含水溝）' },
  { code: '1F', area: '教室區', name: '清掃用具' },
  { code: '1G1', area: '教室區', name: '窗戶' },
  { code: '1G2', area: '教室區', name: '窗台' },
  { code: '1G3', area: '教室區', name: '門' },
  { code: '1H', area: '教室區', name: '儲櫃' },
  { code: '1J', area: '教室區', name: '餐車' },
  { code: '1K', area: '教室區', name: '其它（請註明）', needsNote: true },
  { code: '2A1', area: '外掃區', name: '樹葉/落葉堆（含水溝）' },
  { code: '2A2', area: '外掃區', name: '樹枝' },
  { code: '2A3', area: '外掃區', name: '垃圾打包' },
  { code: '2B', area: '外掃區', name: '人為垃圾' },
  { code: '2C', area: '外掃區', name: '樓梯（請註明樓層）', needsNote: true },
  { code: '2D1', area: '外掃區', name: '洗手台' },
  { code: '2D2', area: '外掃區', name: '飲水機' },
  { code: '2E', area: '外掃區', name: '清掃用具' },
  { code: '2F', area: '外掃區', name: '其它（請註明）', needsNote: true },
  { code: '3A', area: '廁所區', name: '地板垃圾、髒污' },
  { code: '3B', area: '廁所區', name: '蹲、坐式便池（含地板及垃圾）' },
  { code: '3C', area: '廁所區', name: '掛牆式小便池（含外殼）' },
  { code: '3D', area: '廁所區', name: '洗手台、拖把台、鏡子' },
  { code: '3E', area: '廁所區', name: '垃圾桶' },
  { code: '3F', area: '廁所區', name: '清掃用具' }
];

function cleanNotesCodeAt_(source, index, code) {
  const wanted = String(code || '');
  if (!wanted || index < 0 || index + wanted.length > source.length) return false;
  if (source.substr(index, wanted.length).toUpperCase() !== wanted.toUpperCase()) return false;
  const prev = index > 0 ? source.charAt(index - 1) : '';
  const next = source.charAt(index + wanted.length);
  if (prev && /[0-9A-Za-z]/.test(prev)) return false;
  if (next && /[0-9A-Za-z]/.test(next)) return false;
  return true;
}

function cleanNotesHasItemCode_(text, code) {
  const source = String(text || '');
  const wanted = String(code || '');
  if (!wanted) return false;
  for (let i = 0; i <= source.length - wanted.length; i++) {
    if (cleanNotesCodeAt_(source, i, wanted)) return true;
  }
  return false;
}

/** 備註中的項目代號，大小寫都算（1d 與 1D 相同）。 */
window.cleanScoreItemsInNotes = function(text) {
  const source = String(text || '');
  const items = window.CLEAN_SCORE_ITEMS.slice().sort(function(a, b) {
    return b.code.length - a.code.length;
  });
  const found = [];
  const seen = {};
  for (let i = 0; i < source.length; i++) {
    let matched = null;
    for (let j = 0; j < items.length; j++) {
      if (!cleanNotesCodeAt_(source, i, items[j].code)) continue;
      matched = items[j];
      break;
    }
    if (!matched) continue;
    if (!seen[matched.code]) {
      seen[matched.code] = true;
      found.push(matched);
    }
    i += matched.code.length - 1;
  }
  return found;
};

window.cleanNotesHasItemCode = cleanNotesHasItemCode_;

/** 備註裡的項目代號改成「代號 名稱」。寫了「三年級評二年級」且未寫 1D 時，補上 1D 資源、一般垃圾桶扣1。 */
window.formatCleanScoreNotes = function(notes) {
  const raw = notes == null ? '' : String(notes);
  if (!raw) return '';
  const items = window.CLEAN_SCORE_ITEMS.slice().sort(function(a, b) {
    return b.code.length - a.code.length;
  });
  let out = '';
  for (let i = 0; i < raw.length; i++) {
    let matched = null;
    for (let j = 0; j < items.length; j++) {
      const item = items[j];
      if (!cleanNotesCodeAt_(raw, i, item.code)) continue;
      matched = item;
      break;
    }
    if (!matched) {
      out += raw.charAt(i);
      continue;
    }
    const after = raw.slice(i + matched.code.length).replace(/^\s+/, '');
    out += matched.code;
    if (after.indexOf(matched.name) !== 0) out += ' ' + matched.name;
    i += matched.code.length - 1;
  }
  if (out.indexOf('三年級評二年級') >= 0 && !cleanNotesHasItemCode_(out, '1D')) {
    out += '，1D 資源、一般垃圾桶扣1';
  }
  return out;
};
