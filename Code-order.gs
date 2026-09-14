/**
 * 學校秩序評分系統（生輔組）
 * 午休 A–G 加扣分、集會／巡堂登記；可即拍即上傳佐證照片（不另扣分）。
 *
 * 請綁在「秩序專用」Google 試算表，部署成網頁應用程式：
 * 執行身分＝我，存取權＝任何人。前端網址請填入 js/config-order.js。
 */

// 工作表名稱常數
const SHEET_NAMES = {
  CLASSROOMS: '教室清單',      // 教室順序工作表
  SCORES: '評分記錄',          // 評分記錄工作表
  SETTINGS: '系統設定'         // 系統設定工作表（用於存儲密碼等設定）
};

// 年級評比：秩序以年級為單位，學生評本年級（不跨年級互評）
const GRADE_MAPPING = {
  '一年級': '一年級',
  '二年級': '二年級',
  '三年級': '三年級'
};

// 評分區域（秩序不使用區域，保留一項以相容既有欄位）
const AREAS = ['教室'];

// 評比項目
const TIME_SLOTS = ['午休狀況', '重要集會', '課間巡堂', '其他'];

// 每週統計基準分：週總分 = 75 + 當週評分加減
const ORDER_WEEKLY_BASE_SCORE = 75;

// 評分人員類型
const EVALUATOR_TYPES = ['風紀股長', '校安人員', '師長', '巡堂教師', '生輔組', '校長及業管'];

// 添加一個全局緩存來存儲資料夾引用（參考 uploadFileToDrive 的成功策略）
const folderCache = {};
const ORDER_PHOTO_FOLDER_NAME = '學校秩序評分照片';
const ORDER_PHOTO_FOLDER_PROP = 'ORDER_PHOTO_FOLDER_ID';

function getOrderPhotoFolder_() {
  if (folderCache[ORDER_PHOTO_FOLDER_NAME]) {
    return folderCache[ORDER_PHOTO_FOLDER_NAME];
  }

  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty(ORDER_PHOTO_FOLDER_PROP);
  if (savedId) {
    try {
      const saved = DriveApp.getFolderById(savedId);
      folderCache[ORDER_PHOTO_FOLDER_NAME] = saved;
      return saved;
    } catch (e) {
      props.deleteProperty(ORDER_PHOTO_FOLDER_PROP);
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(ss.getId());
  const parents = file.getParents();
  const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folders = parentFolder.getFoldersByName(ORDER_PHOTO_FOLDER_NAME);
  let folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = parentFolder.createFolder(ORDER_PHOTO_FOLDER_NAME);
  }
  folderCache[ORDER_PHOTO_FOLDER_NAME] = folder;
  try {
    props.setProperty(ORDER_PHOTO_FOLDER_PROP, folder.getId());
  } catch (e) {}
  return folder;
}

/**
 * 初始化工作表（首次執行時創建必要的工作表）
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 創建教室清單工作表（如果不存在）
  let classroomsSheet = ss.getSheetByName(SHEET_NAMES.CLASSROOMS);
  if (!classroomsSheet) {
    classroomsSheet = ss.insertSheet(SHEET_NAMES.CLASSROOMS);
    // 設定標題列
    // 欄位：教室編號（樓層／位置，可重複）、教室名稱（班級，唯一）、順序
    classroomsSheet.getRange(1, 1, 1, 3).setValues([['教室編號', '教室名稱', '順序']]);
    classroomsSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    
    const exampleData = [
      ['仁愛五樓', '資訊二忠', 1],
      ['仁愛五樓', '資訊二孝', 2],
      ['仁愛五樓', '資訊一忠', 3],
      ['仁愛五樓', '資訊一孝', 4]
    ];
    classroomsSheet.getRange(2, 1, exampleData.length, 3).setValues(exampleData);
  }
  
  // 創建評分記錄工作表（如果不存在）
  let scoresSheet = ss.getSheetByName(SHEET_NAMES.SCORES);
  if (!scoresSheet) {
    scoresSheet = ss.insertSheet(SHEET_NAMES.SCORES);
    // 設定標題列
    scoresSheet.getRange(1, 1, 1, 15).setValues([[
      '評分時間', 
      '評分年級', 
      '被評年級',
      '評分人員類型',
      '評分人員', 
      '教室編號', 
      '教室名稱', 
      '評比地點',
      '評比項目',
      '扣分',
      '加分',
      '總分（加分-扣分）', 
      '備註',
      '細項說明',
      '照片連結'
    ]]);
    scoresSheet.getRange(1, 1, 1, 15).setFontWeight('bold');
  } else {
    dropOrderDirectDeductionColumn_(scoresSheet);
    const headers = scoresSheet.getRange(1, 1, 1, scoresSheet.getLastColumn()).getValues()[0];
    const hasPhotoLinkColumn = headers.includes('照片連結');
    
    if (!hasPhotoLinkColumn) {
      // 如果沒有「照片連結」欄位，添加它
      const lastCol = scoresSheet.getLastColumn();
      scoresSheet.getRange(1, lastCol + 1).setValue('照片連結');
      scoresSheet.getRange(1, lastCol + 1).setFontWeight('bold');
    }
  }
  
  // 創建系統設定工作表（如果不存在）
  let settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
    // 設定標題列
    settingsSheet.getRange(1, 1, 1, 2).setValues([['設定項目', '設定值']]);
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    
    // 設定預設密碼（可以修改）
    settingsSheet.getRange(2, 1, 1, 2).setValues([['評分系統密碼', '1234']]);
    // 設定學生密碼（第二組密碼）
    settingsSheet.getRange(3, 1, 1, 2).setValues([['學生密碼', '5678']]);
    settingsSheet.getRange(4, 1, 1, 2).setValues([['管理員密碼', '9999']]);
  }
}

function dropOrderDirectDeductionColumn_(sheet) {
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let i = headers.length - 1; i >= 0; i--) {
    if (String(headers[i] || '').trim() === '直接扣分') {
      sheet.deleteColumn(i + 1);
    }
  }
}

function getOrderScoresSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.SCORES);
  if (!sheet) {
    initializeSheets();
    sheet = ss.getSheetByName(SHEET_NAMES.SCORES);
  } else {
    dropOrderDirectDeductionColumn_(sheet);
  }
  return sheet;
}

function getOrderScoreColMap_(headers) {
  const names = (headers || []).map(function(h) {
    return String(h || '').trim();
  });
  function findCol(label, fallback) {
    const i = names.indexOf(label);
    return i >= 0 ? i : fallback;
  }
  function findTotalCol(fallback) {
    for (let i = 0; i < names.length; i++) {
      if (names[i] === '總分' || names[i].indexOf('總分') === 0) return i;
    }
    return fallback;
  }
  const hasDirect = names.indexOf('直接扣分') >= 0;
  return {
    timestamp: findCol('評分時間', 0),
    evaluatorGrade: findCol('評分年級', 1),
    evaluatedGrade: findCol('被評年級', 2),
    evaluatorType: findCol('評分人員類型', 3),
    evaluator: findCol('評分人員', 4),
    classroomId: findCol('教室編號', 5),
    classroomName: findCol('教室名稱', 6),
    area: findCol('評比地點', 7),
    timeSlot: findCol('評比項目', 8),
    deduction: findCol('扣分', 9),
    bonus: findCol('加分', hasDirect ? 11 : 10),
    totalScore: findTotalCol(hasDirect ? 12 : 11),
    notes: findCol('備註', hasDirect ? 13 : 12),
    itemSummary: findCol('細項說明', hasDirect ? 14 : 13),
    photoLinks: findCol('照片連結', hasDirect ? 15 : 14)
  };
}

function applyOrderWeeklyBaseScore_(item) {
  if (!item) return item;
  const delta = Number(item.totalScore) || 0;
  item.scoreDelta = delta;
  item.totalScore = ORDER_WEEKLY_BASE_SCORE + delta;
  return item;
}

function applyOrderWeeklyBaseToMap_(map) {
  Object.keys(map || {}).forEach(function(key) {
    applyOrderWeeklyBaseScore_(map[key]);
  });
}

function applyOrderWeeklyBaseToGradeGroups_(gradeGroups) {
  Object.keys(gradeGroups || {}).forEach(function(grade) {
    applyOrderWeeklyBaseToMap_(gradeGroups[grade]);
  });
}

function getOrderWeekStartKey_(value) {
  const d = value instanceof Date ? new Date(value.getTime()) : new Date(value);
  if (isNaN(d.getTime())) return '';
  const dayOfWeek = d.getDay();
  const diff = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
  d.setDate(d.getDate() + diff);
  d.setHours(0, 0, 0, 0);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getOrderWeekKeysInRange_(start, end) {
  const keys = [];
  const firstKey = getOrderWeekStartKey_(start);
  const lastKey = getOrderWeekStartKey_(end);
  if (!firstKey || !lastKey) return keys;
  const monday = new Date(firstKey + ' 00:00:00');
  const last = new Date(lastKey + ' 00:00:00');
  while (monday.getTime() <= last.getTime()) {
    keys.push(Utilities.formatDate(monday, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    monday.setDate(monday.getDate() + 7);
  }
  return keys;
}

/**
 * 取得所有年級清單
 * @return {Array} 年級清單
 */
function getGrades() {
  return ['一年級', '二年級', '三年級'];
}

/**
 * 取得評分區域清單
 * @return {Array} 評分區域清單
 */
function getAreas() {
  return AREAS;
}

/**
 * 取得評分時段清單
 * @return {Array} 評分時段清單
 */
function getTimeSlots() {
  return TIME_SLOTS;
}

/**
 * 取得評分人員類型清單
 * @return {Array} 評分人員類型清單
 */
function getEvaluatorTypes() {
  return EVALUATOR_TYPES;
}

/**
 * 取得教室清單（依順序排列）
 * @param {string} grade 評分年級（選填，用於過濾被評年級）
 * @return {Array} 教室清單陣列
 */
function getClassrooms(grade) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CLASSROOMS);
  
  if (!sheet) {
    initializeSheets();
    return getClassrooms(grade);
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return [];
  }
  
  // 建立年級格式對照表（支援多種格式）
  const gradeFormatMap = {
    '一年級': ['一年級', '1年級', '1', '一', '一級'],
    '二年級': ['二年級', '2年級', '2', '二', '二級'],
    '三年級': ['三年級', '3年級', '3', '三', '三級']
  };
  
  // 從教室名稱中提取年級（支援多種格式）
  function extractGradeFromName(name) {
    if (!name) return null;
    
    // 優先匹配完整年級名稱（如「一年一班」「二年一班」「三年一班」）
    if (name.includes('一年')) return '一年級';
    if (name.includes('二年')) return '二年級';
    if (name.includes('三年')) return '三年級';
    
    // 匹配數字格式（如「1年一班」「2年一班」「3年一班」）
    if (name.includes('1年') || (name.match(/^1/) && !name.match(/^10/))) return '一年級';
    if (name.includes('2年') || (name.match(/^2/) && !name.match(/^20/))) return '二年級';
    if (name.includes('3年') || (name.match(/^3/) && !name.match(/^30/))) return '三年級';
    
    // 匹配中間的年級字（如「電機一忠」「電機二孝」「電機三忠」）
    // 使用更精確的匹配，避免匹配到其他位置的數字
    const gradeMatch = name.match(/([一二三])/);
    if (gradeMatch) {
      const char = gradeMatch[1];
      if (char === '一') return '一年級';
      if (char === '二') return '二年級';
      if (char === '三') return '三年級';
    }
    
    // 匹配中間的數字年級（如「電機1忠」「電機2孝」）
    const numMatch = name.match(/([123])(?![0-9])/); // 匹配1、2、3，但後面不能是數字
    if (numMatch) {
      const num = numMatch[1];
      if (num === '1') return '一年級';
      if (num === '2') return '二年級';
      if (num === '3') return '三年級';
    }
    
    return null;
  }
  
  // 檢查年級是否匹配（支援多種格式）
  function isGradeMatch(roomGrade, targetGrade) {
    if (!roomGrade || !targetGrade) return false;
    const roomGradeTrim = String(roomGrade).trim();
    const targetVariants = gradeFormatMap[targetGrade] || [targetGrade];
    return targetVariants.some(variant => roomGradeTrim === variant);
  }
  
  // 跳過標題列，取得資料
  // 工作表結構：教室編號、教室名稱、順序（沒有年級欄位）
  let classrooms = data.slice(1)
    .map(row => {
      const id = String(row[0] || '').trim();
      const name = String(row[1] || '').trim();
      const order = Number(row[2]) || 999; // 順序在第3欄
      
      // 從教室名稱提取年級（因為工作表沒有年級欄位）
      let grade = extractGradeFromName(name);
      
      // 如果無法從名稱提取，嘗試從年級欄位讀取（向後相容）
      if (!grade && row.length > 3) {
        grade = String(row[3] || '').trim();
      }
      
      return {
        id: id,
        name: name,
        grade: grade || '',
        order: order
      };
    })
    .filter(room => room.id && room.name && room.grade); // 必須有年級才能過濾
  
  // 如果指定了評分年級，則過濾出對應的被評年級
  if (grade && GRADE_MAPPING[grade]) {
    const targetGrade = GRADE_MAPPING[grade];
    
    classrooms = classrooms.filter(room => {
      // 檢查年級是否匹配（年級已從名稱提取或從欄位讀取）
      return isGradeMatch(room.grade, targetGrade);
    });
  }
  
  // 依順序排序
  classrooms.sort((a, b) => a.order - b.order);
  
  return classrooms;
}

/**
 * 上傳照片到 Google Drive（採用 uploadFileToDrive 的成功策略）
 * @param {string} base64Data 照片的 base64 資料
 * @param {string} filename 檔案名稱
 * @return {string} 照片的 Google Drive 連結
 */
function uploadPhotoToDrive(base64Data, filename) {
  try {
    if (!base64Data || !filename) {
      throw new Error('參數無效');
    }

    const folder = getOrderPhotoFolder_();
    const commaIndex = base64Data.indexOf(',');
    const base64Content = commaIndex >= 0 ? base64Data.substring(commaIndex + 1) : base64Data;
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Content), 'image/jpeg', filename);
    return folder.createFile(blob).getUrl();
  } catch (error) {
    const errorMsg = '上傳照片失敗：' + error.toString();
    Logger.log(errorMsg);
    
    let detailedError = errorMsg;
    const errorStr = error.toString().toLowerCase();
    if (errorStr.includes('permission') || errorStr.includes('權限') || errorStr.includes('access')) {
      detailedError += ' 可能的原因：1) Google Workspace 管理員限制了 Drive API 存取權限；2) 需要授權「查看和管理您的 Google Drive 檔案」權限；3) 帳戶沒有創建資料夾的權限。請聯繫 Google Workspace 管理員檢查相關設定。';
    } else if (errorStr.includes('quota') || errorStr.includes('配額')) {
      detailedError += ' 可能的原因：Google Drive 儲存空間不足。';
    } else if (errorStr.includes('rate') || errorStr.includes('限制')) {
      detailedError += ' 可能的原因：上傳頻率過高，請稍後再試。';
    }
    throw new Error(detailedError);
  }
}

/**
 * 清除資料夾緩存（參考 uploadFileToDrive）
 */
function clearFolderCache() {
  Object.keys(folderCache).forEach(key => {
    delete folderCache[key];
  });
  Logger.log('資料夾緩存已清除');
}


/**
 * 取得照片資料夾連結
 * @return {string} 照片資料夾的 Google Drive 連結
 */
function getPhotoFolderLink() {
  try {
    return getOrderPhotoFolder_().getUrl();
  } catch (error) {
    Logger.log('取得照片資料夾連結失敗：' + error.toString());
    return '';
  }
}

/**
 * 儲存評分記錄
 * @param {Object} scoreData 評分資料物件
 * @return {Object} 執行結果
 */
/**
 * 檢查 Google Drive 權限
 * @return {Object} 權限檢查結果
 */
function checkDrivePermission() {
  try {
    // 嘗試存取 Drive
    DriveApp.getFolders();
    return {
      hasPermission: true,
      message: '權限正常'
    };
  } catch (error) {
    const errorMsg = error.toString() || '';
    if (errorMsg.includes('權限') || errorMsg.includes('permission') || errorMsg.includes('DriveApp') || errorMsg.includes('authorization')) {
      return {
        hasPermission: false,
        message: '需要授權 Google Drive 存取權限',
        instructions: '請在 Google Apps Script 編輯器中，點擊「執行」按鈕，選擇任意函數執行，然後授權「查看和管理您的 Google Drive 檔案」權限。'
      };
    }
    return {
      hasPermission: false,
      message: '檢查權限時發生錯誤：' + errorMsg
    };
  }
}

/**
 * 單獨上傳照片（用於處理大量照片資料）
 * @param {string} photoBase64 照片的 base64 資料
 * @param {string} filename 檔案名稱
 * @return {string} 照片的 Google Drive 連結
 */
function uploadSinglePhoto(photoBase64, filename) {
  try {
    // 檢查參數（優化：減少日誌記錄）
    if (!filename || filename === 'undefined') {
      throw new Error('檔案名稱無效：' + filename);
    }
    
    if (!photoBase64 || photoBase64.length === 0) {
      throw new Error('照片資料為空');
    }
    
    // 直接調用上傳函數
    const link = uploadPhotoToDrive(photoBase64, filename);
    
    if (link && link.trim() !== '') {
      return link;
    } else {
      throw new Error('照片上傳失敗，返回空連結');
    }
  } catch (error) {
    const errorMsg = error.toString() || '';
    
    // 如果是權限錯誤，提供更詳細的訊息
    if (errorMsg.includes('權限') || errorMsg.includes('permission') || errorMsg.includes('DriveApp') || errorMsg.includes('authorization')) {
      throw new Error('權限錯誤：需要授權 Google Drive 存取權限。請在 Google Apps Script 編輯器中執行一次腳本來完成權限授權。');
    }
    
    // 拋出錯誤以便前端可以捕獲
    throw error;
  }
}

/**
 * 接收照片分塊（JSONP 無法一次傳送大張 base64）
 */
function uploadPhotoChunk(filename, index, total, chunk) {
  const cache = CacheService.getScriptCache();
  const safe = String(filename || '').replace(/[^A-Za-z0-9._-]/g, '_');
  cache.put('ph_' + safe + '_' + Number(index), String(chunk || ''), 600);
  cache.put('phm_' + safe, String(total), 600);
  return { success: true, index: Number(index), total: Number(total) };
}

/**
 * 組合分塊後上傳到 Drive
 */
function finalizePhotoUpload(filename) {
  const cache = CacheService.getScriptCache();
  const safe = String(filename || '').replace(/[^A-Za-z0-9._-]/g, '_');
  const total = Number(cache.get('phm_' + safe) || 0);
  if (!total) {
    throw new Error('找不到照片分塊，請重新上傳');
  }
  const parts = [];
  for (let i = 0; i < total; i++) {
    const part = cache.get('ph_' + safe + '_' + i);
    if (part === null || part === undefined) {
      throw new Error('照片分塊缺失：' + (i + 1) + '/' + total);
    }
    parts.push(part);
  }
  for (let i = 0; i < total; i++) {
    cache.remove('ph_' + safe + '_' + i);
  }
  cache.remove('phm_' + safe);
  return uploadSinglePhoto(parts.join(''), filename);
}

function orderScoreDateKey_(value) {
  try {
    const d = value instanceof Date ? value : new Date(value);
    if (isNaN(d.getTime())) return '';
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (err) {
    return '';
  }
}

function isOrderTeacherEvaluator_(type) {
  const text = String(type || '').trim();
  if (!text) return false;
  if (text.indexOf('風紀股長') >= 0) return false;
  return true;
}

function findOwnOrderScoreRow_(sheet, scoreData, category) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 0;
  const todayKey = orderScoreDateKey_(new Date());
  const evaluator = String(scoreData.evaluator || '未指定').trim();
  const evaluatorType = String(scoreData.evaluatorType || '').trim();
  const classroomName = String(scoreData.classroomName || '').trim();
  const classroomId = String(scoreData.classroomId || '').trim();
  const cat = String(category || '').trim();

  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if (orderScoreDateKey_(row[0]) !== todayKey) continue;
    if (String(row[4] || '').trim() !== evaluator) continue;
    if (String(row[3] || '').trim() !== evaluatorType) continue;
    if (String(row[8] || '').trim() !== cat) continue;
    const rowName = String(row[6] || '').trim();
    const rowId = String(row[5] || '').trim();
    const sameClass = classroomName ? rowName === classroomName : (!!classroomId && rowId === classroomId);
    if (!sameClass) continue;
    return i + 1;
  }
  return 0;
}

function isOrderEvaluatorTypeAllowed_(loginType, evaluatorType) {
  const type = String(evaluatorType || '').trim();
  if (!type) return false;
  if (loginType === 'student') {
    return type.indexOf('風紀股長') >= 0;
  }
  if (loginType === 'teacher') {
    return type.indexOf('師長') >= 0 || type.indexOf('巡堂') >= 0 || type.indexOf('校安') >= 0;
  }
  if (loginType === 'admin') {
    return type.indexOf('生輔組') >= 0 || type.indexOf('校長') >= 0;
  }
  return false;
}

function saveScore(scoreData, authPassword) {
  try {
    const loginType = getAuthLoginType_(authPassword);
    if (!loginType) {
      return { success: false, message: '未授權：請先從查詢頁輸入密碼進入評分系統' };
    }
    if (!isOrderEvaluatorTypeAllowed_(loginType, scoreData && scoreData.evaluatorType)) {
      return { success: false, message: '評分人員類型與登入身分不符' };
    }

    const sheet = getOrderScoresSheet_();
    
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const category = String(scoreData.timeSlot || scoreData.category || '').trim();
    const items = scoreData.items || {};
    const notes = String(scoreData.notes || '').trim();
    
    let bonus = Number(scoreData.bonus) || 0;
    let deduction = Number(scoreData.directDeduction) || 0;
    let itemSummary = String(scoreData.itemSummary || '').trim();
    
    const isNap = category.indexOf('午休') >= 0;
    const hasItemFlags = items.A || items.B || items.C || items.D || items.E || items.F || items.G;
    const itemBased = isNap || hasItemFlags;
    if (itemBased) {
      bonus = 0;
      deduction = 0;
      const parts = [];
      if (items.A) {
        bonus += 5;
        parts.push('A全班午睡+5');
      } else if (items.B) {
        bonus += 3;
        parts.push('B安靜休息+3');
      }
      if (items.C) {
        deduction += 3;
        parts.push('C未關燈-3');
      }
      const walkCount = Number(items.D) || 0;
      if (walkCount > 0) {
        deduction += walkCount;
        parts.push('D任意走動' + walkCount + '人');
      }
      const groupCount = Number(items.E) || 0;
      if (groupCount > 0) {
        deduction += groupCount;
        parts.push('E群聚' + groupCount);
      }
      const noiseCount = Number(items.F) || 0;
      if (noiseCount > 0) {
        deduction += noiseCount;
        parts.push('F吵鬧交談' + noiseCount);
      }
      const otherDeduction = Math.min(3, Math.max(0, Number(items.G) || 0));
      if (otherDeduction > 0) {
        deduction += otherDeduction;
        parts.push('G其他行為-' + otherDeduction);
      }
      itemSummary = parts.join('；');
    }
    
    const totalScore = bonus - deduction;
    const photoLinks = String(scoreData.photoLinks || '').trim();
    
    let evaluatedGrade = '';
    function extractGradeFromName(name) {
      if (!name) return '';
      const text = String(name);
      if (text.includes('一年')) return '一年級';
      if (text.includes('二年')) return '二年級';
      if (text.includes('三年')) return '三年級';
      if (text.includes('1年') || (text.match(/^1/) && !text.match(/^10/))) return '一年級';
      if (text.includes('2年') || (text.match(/^2/) && !text.match(/^20/))) return '二年級';
      if (text.includes('3年') || (text.match(/^3/) && !text.match(/^30/))) return '三年級';
      const gradeMatch = text.match(/([一二三])/);
      if (gradeMatch) {
        const char = gradeMatch[1];
        if (char === '一') return '一年級';
        if (char === '二') return '二年級';
        if (char === '三') return '三年級';
      }
      return '';
    }
    if (scoreData.classroomName) {
      evaluatedGrade = extractGradeFromName(scoreData.classroomName);
    }
    if (!evaluatedGrade && scoreData.evaluatorGrade) {
      evaluatedGrade = GRADE_MAPPING[scoreData.evaluatorGrade] || scoreData.evaluatorGrade;
    }
    
    const newRow = [
      timestamp,
      scoreData.evaluatorGrade || '',
      evaluatedGrade,
      scoreData.evaluatorType || '',
      scoreData.evaluator || '未指定',
      scoreData.classroomId || '',
      scoreData.classroomName || '',
      scoreData.area || '教室',
      category,
      deduction,
      bonus,
      totalScore,
      notes,
      itemSummary,
      photoLinks
    ];
    
    let replaced = false;
    if (isOrderTeacherEvaluator_(scoreData.evaluatorType)) {
      const existingRow = findOwnOrderScoreRow_(sheet, scoreData, category);
      if (existingRow) {
        sheet.getRange(existingRow, 1, 1, newRow.length).setValues([newRow]);
        replaced = true;
      }
    }
    if (!replaced) {
      sheet.appendRow(newRow);
    }
    
    return {
      success: true,
      message: replaced
        ? '已用本次評分覆蓋您今天稍早的紀錄（未更動其他老師）'
        : '評分記錄已儲存',
      replaced: replaced,
      timestamp: timestamp,
      totalScore: totalScore,
      itemSummary: itemSummary
    };
  } catch (error) {
    Logger.log('儲存失敗：' + error.toString());
    return {
      success: false,
      message: '儲存失敗：' + error.toString()
    };
  }
}


/**
 * 取得評分統計（可選功能）
 * @param {string} classroomId 教室編號（選填）
 * @return {Object} 統計資料
 */
function getScoreStatistics(classroomId) {
  const sheet = getOrderScoresSheet_();
  
  if (!sheet) {
    return { totalRecords: 0, averageDeduction: 0 };
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { totalRecords: 0, averageDeduction: 0 };
  }
  
  let records = data.slice(1);
  
  // 如果指定教室，則過濾
  if (classroomId) {
    records = records.filter(row => row[2] === classroomId);
  }
  
  const totalRecords = records.length;
  const totalDeduction = records.reduce((sum, row) => sum + (row[5] || 0), 0);
  const averageDeduction = totalRecords > 0 ? (totalDeduction / totalRecords).toFixed(2) : 0;
  
  return {
    totalRecords: totalRecords,
    totalDeduction: totalDeduction,
    averageDeduction: parseFloat(averageDeduction)
  };
}

/**
 * 取得評分記錄（用於查看頁面）
 * @param {string} classroomId 教室編號（選填）
 * @param {string} classroomName 教室名稱（選填）
 * @param {string} grade 年級（選填）
 * @return {Array} 評分記錄陣列
 */
function getScoreRecords(classroomId, classroomName, grade) {
  try {
    const sheet = getOrderScoresSheet_();
    
    if (!sheet) {
      Logger.log('找不到評分記錄工作表');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log('評分記錄工作表沒有資料（只有標題列）');
      return [];
    }
    
    Logger.log('總共有 ' + (data.length - 1) + ' 筆記錄');
    Logger.log('查詢條件：classroomId=' + classroomId + ', classroomName=' + classroomName + ', grade=' + grade);
    
    // 取得標題列
    const headers = data[0];
    const col = getOrderScoreColMap_(headers);
    Logger.log('標題列：' + headers.join(', '));
    
    // 取得資料列（跳過標題列）
    let records = data.slice(1);
    Logger.log('過濾前記錄數：' + records.length);
    
    // 主要以教室名稱過濾（教室編號可能是樓層如「仁愛五樓」，會重複）
    if (classroomName && classroomName.trim() !== '') {
      const beforeCount = records.length;
      const nameTrim = String(classroomName).trim();
      records = records.filter(row => {
        const rowClassroomName = String(row[6] || '').trim();
        return rowClassroomName === nameTrim;
      });
      Logger.log('根據教室名稱過濾：' + beforeCount + ' -> ' + records.length);
    } else if (classroomId && classroomId.trim() !== '') {
      const beforeCount = records.length;
      const idTrim = String(classroomId).trim();
      records = records.filter(row => String(row[5] || '').trim() === idTrim);
      Logger.log('根據教室編號過濾：' + beforeCount + ' -> ' + records.length);
    }
    
    if (grade && grade.trim() !== '') {
      // 建立年級格式對照表（支援多種格式，例如「一年級」「1年級」「1」「一」等）
      const gradeFormatMap = {
        '一年級': ['一年級', '1年級', '1', '一', '一級'],
        '二年級': ['二年級', '2年級', '2', '二', '二級'],
        '三年級': ['三年級', '3年級', '3', '三', '三級']
      };
      
      const gradeTrim = String(grade).trim();
      const targetVariants = gradeFormatMap[gradeTrim] || [gradeTrim];
      
      // 從教室名稱中推斷年級（支援職科班級名稱，例如「資訊一忠」「製圖二忠」「汽車三孝」）
      function extractGradeFromNameForRecord(name) {
        if (!name) return '';
        const text = String(name);
        
        // 優先匹配「一年」「二年」「三年」
        if (text.includes('一年')) return '一年級';
        if (text.includes('二年')) return '二年級';
        if (text.includes('三年')) return '三年級';
        
        // 再看是否有獨立出現的 1/2/3（避免 10、20、30 等房號）
        const numMatch = text.match(/([123])(?![0-9])/);
        if (numMatch) {
          if (numMatch[1] === '1') return '一年級';
          if (numMatch[1] === '2') return '二年級';
          if (numMatch[1] === '3') return '三年級';
        }
        
        // 最後看是否有「一／二／三」這三個字
        const chMatch = text.match(/([一二三])/);
        if (chMatch) {
          if (chMatch[1] === '一') return '一年級';
          if (chMatch[1] === '二') return '二年級';
          if (chMatch[1] === '三') return '三年級';
        }
        
        return '';
      }
      
      // 被評年級在第3欄（索引2），教室名稱在第7欄（索引6）
      const beforeCount = records.length;
      records = records.filter(row => {
        const rawGrade = String(row[2] || '').trim();   // 被評年級欄位（可能為空）
        const classroomName = String(row[6] || '').trim();
        
        // 先用欄位中的被評年級；如果是空的，就從教室名稱推斷
        const effectiveGrade = rawGrade || extractGradeFromNameForRecord(classroomName);
        const match = targetVariants.includes(effectiveGrade);
        
        if (!match) {
          Logger.log('年級不匹配：期望=' + gradeTrim +
                     ' (可接受: ' + targetVariants.join('、') + '), 原始欄位=' + rawGrade +
                     ', 推斷年級=' + effectiveGrade + ', 教室名稱=' + classroomName);
        }
        return match;
      });
      Logger.log('根據年級過濾：' + beforeCount + ' -> ' + records.length);
    }
    
    Logger.log('過濾後記錄數：' + records.length);
    
    // 如果沒有任何過濾條件，返回所有記錄
    if (!classroomId && !classroomName && !grade) {
      Logger.log('沒有指定過濾條件，返回所有記錄');
    }
    
    // 轉換為物件陣列（確保所有值都是可序列化的基本類型）
    const result = records.map((row, index) => {
      try {
        // 確保所有值都是字符串或數字，避免 Date 對象等無法序列化的類型
        const timestamp = row[0] ? String(row[0]) : '';
        const evaluatorGrade = row[1] ? String(row[1]) : '';
        const evaluatedGrade = row[2] ? String(row[2]) : '';
        const evaluatorType = row[3] ? String(row[3]) : '';
        const evaluator = row[4] ? String(row[4]) : '';
        const classroomId = row[5] ? String(row[5]) : '';
        const classroomName = row[6] ? String(row[6]) : '';
        const area = row[7] ? String(row[7]) : '';
        const timeSlot = row[8] ? String(row[8]) : '';
        const photoDeduction = Number(row[col.deduction]) || 0;
        const bonus = Number(row[col.bonus]) || 0;
        const totalScore = Number(row[col.totalScore]) || 0;
        const notes = row[col.notes] ? String(row[col.notes]) : '';
        const itemSummary = row[col.itemSummary] ? String(row[col.itemSummary]) : '';
        const photoLinks = row[col.photoLinks] ? String(row[col.photoLinks]) : '';
        
        return {
          timestamp: timestamp,
          evaluatorGrade: evaluatorGrade,
          evaluatedGrade: evaluatedGrade,
          evaluatorType: evaluatorType,
          evaluator: evaluator,
          classroomId: classroomId,
          classroomName: classroomName,
          area: area,
          timeSlot: timeSlot,
          photoDeduction: photoDeduction,
          directDeduction: 0,
          bonus: bonus,
          totalScore: totalScore,
          notes: notes,
          itemSummary: itemSummary,
          photoLinks: photoLinks
        };
      } catch (err) {
        Logger.log('轉換記錄 ' + index + ' 時發生錯誤：' + err.toString());
        Logger.log('錯誤堆疊：' + (err.stack || '無堆疊資訊'));
        return null;
      }
    }).filter(record => record !== null);
    
    Logger.log('成功轉換 ' + result.length + ' 筆記錄');
    
    // 依時間排序（最新的在前）
    result.sort((a, b) => {
      try {
        const timeA = new Date(a.timestamp);
        const timeB = new Date(b.timestamp);
        return timeB - timeA;
      } catch (err) {
        return 0;
      }
    });
    
    Logger.log('查詢完成，返回 ' + result.length + ' 筆記錄');
    
    // 確保返回的數據可以正確序列化
    try {
      JSON.stringify(result);
      return result;
    } catch (serializeError) {
      Logger.log('序列化失敗：' + serializeError.toString());
      return [];
    }
  } catch (error) {
    Logger.log('getScoreRecords 發生錯誤：' + error.toString());
    Logger.log('錯誤堆疊：' + (error.stack || '無堆疊資訊'));
    return [];
  }
}

/**
 * 取得每週評分統計（按年級分組，計算特優和優等）
 * @param {string} weekStartDate 週開始日期 (yyyy-MM-dd 格式，選填，預設為本週)
 * @return {Object} 週統計資料，包含各年級的特優和優等
 */
function getWeeklyStatistics(weekStartDate) {
  try {
    const sheet = getOrderScoresSheet_();
    
    if (!sheet) {
      return { error: '找不到評分記錄工作表' };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { error: '沒有評分記錄' };
    }
    const col = getOrderScoreColMap_(data[0]);
    
    // 計算週的開始和結束日期
    let weekStart;
    let weekEnd;
    
    if (weekStartDate) {
      weekStart = new Date(weekStartDate + ' 00:00:00');
      weekEnd = new Date(weekStart);
      weekEnd.setDate(weekEnd.getDate() + 6);
      weekEnd.setHours(23, 59, 59, 999);
    } else {
      // 預設為本週（週一到週日）
      const today = new Date();
      const dayOfWeek = today.getDay(); // 0=週日, 1=週一, ..., 6=週六
      const diff = dayOfWeek === 0 ? -6 : 1 - dayOfWeek; // 調整為週一開始
      weekStart = new Date(today);
      weekStart.setDate(today.getDate() + diff);
      weekStart.setHours(0, 0, 0, 0);
      weekEnd = new Date(weekStart);
      weekEnd.setDate(weekStart.getDate() + 6);
      weekEnd.setHours(23, 59, 59, 999);
    }
    
    // 取得資料列（跳過標題列）
    const records = data.slice(1);
    
    // 過濾本週的記錄
    const weeklyRecords = records.filter(row => {
      try {
        const recordDate = new Date(row[0]); // 評分時間在第1欄（索引0）
        return recordDate >= weekStart && recordDate <= weekEnd;
      } catch (err) {
        return false;
      }
    });
    
    // 按年級和「教室名稱」分組，計算總分
    // 這裡只用「教室名稱」作為分組依據，同名教室視為同一班級來統計
    const gradeGroups = {};
    
    // 從教室名稱提取年級的函數（改進版：優先匹配完整年級，避免誤判）
    function extractGradeFromName(name) {
      if (!name) return '';
      const text = String(name);
      // 優先匹配完整年級名稱（如「一年一班」「二年一班」「三年一班」）
      if (text.includes('一年')) return '一年級';
      if (text.includes('二年')) return '二年級';
      if (text.includes('三年')) return '三年級';
      // 匹配數字格式（如「1年一班」「2年一班」「3年一班」）
      if (text.includes('1年') || (text.match(/^1/) && !text.match(/^10/))) return '一年級';
      if (text.includes('2年') || (text.match(/^2/) && !text.match(/^20/))) return '二年級';
      if (text.includes('3年') || (text.match(/^3/) && !text.match(/^30/))) return '三年級';
      // 最後匹配單個年級字（如「電機一忠」「電機二孝」「電機三忠」）
      // 使用正則表達式精確匹配，避免誤判
      const gradeMatch = text.match(/([一二三])/);
      if (gradeMatch) {
        const char = gradeMatch[1];
        if (char === '一') return '一年級';
        if (char === '二') return '二年級';
        if (char === '三') return '三年級';
      }
      return '';
    }
    
    weeklyRecords.forEach(row => {
      let grade = String(row[2] || '').trim(); // 被評年級
      const classroomId = String(row[5] || '').trim();   // 教室編號（例如：仁愛五樓）
      const classroomName = String(row[6] || '').trim(); // 教室名稱（例如：控制一忠、冷凍一孝）
      const totalScore = Number(row[col.totalScore]) || 0;           // 單次總分
      
      // 使用「教室名稱」作為分組 key，只依名稱統計
      const classroomKey = classroomName;
      
      // 如果被評年級為空，從教室名稱推斷
      if (!grade && classroomName) {
        grade = extractGradeFromName(classroomName);
      }
      
      // 沒有年級或沒有教室名稱就不計算（編號可能是樓層、不適合作唯一鍵）
      if (!grade || !classroomName) return;
      
      // 初始化年級組
      if (!gradeGroups[grade]) {
        gradeGroups[grade] = {};
      }
      
      // 初始化教室（以 classroomKey 為索引）
      if (!gradeGroups[grade][classroomKey]) {
        gradeGroups[grade][classroomKey] = {
          classroomId: classroomId,
          classroomName: classroomName,
          grade: grade,
          totalScore: 0,
          recordCount: 0
        };
      }
      
      // 累加總分
      gradeGroups[grade][classroomKey].totalScore += totalScore;
      gradeGroups[grade][classroomKey].recordCount += 1;
    });

    applyOrderWeeklyBaseToGradeGroups_(gradeGroups);
    
    // 轉換為陣列並排序，找出特優和優等（處理同分名次）
    const result = {};
    
    Object.keys(gradeGroups).forEach(grade => {
      const classrooms = Object.values(gradeGroups[grade]);
      
      // 按總分降序排序（分數越高越好）
      classrooms.sort((a, b) => b.totalScore - a.totalScore);
      
      // 依總分計算名次：
      // - 同分者名次相同
      // - 例如第4、5、6名同分，三者皆為第4名，下一名為第7名
      const rankedClassrooms = [];
      let lastScore = null;
      let currentRank = 0;
      
      for (let i = 0; i < classrooms.length; i++) {
        const cls = classrooms[i];
        if (lastScore === null || cls.totalScore !== lastScore) {
          currentRank = i + 1; // 新分數，名次 = 目前索引 + 1
          lastScore = cls.totalScore;
        }
        rankedClassrooms.push({
          rank: currentRank,
          ...cls
        });
      }
      
      // 特優：名次 <= 3 的所有班級（可能超過3個，因為同分並列）
      const top3 = rankedClassrooms.filter(c => c.rank <= 3);
      
      // 優等：從名次 >= 4 開始往後取「最多三個班級」
      // 說明：不再侷限在 4~6 名，避免出現只抓到 1 名的情況
      const excellentCandidates = rankedClassrooms.filter(c => c.rank >= 4);
      const excellent3 = excellentCandidates.slice(0, 3);
      
      // 待改進：最後「三個位置」往後擴展，將同分者一起納入
      let bottom3 = [];
      if (rankedClassrooms.length <= 3) {
        bottom3 = rankedClassrooms.slice(); // 班級數小於等於3，全列入
      } else {
        const thresholdIndex = Math.max(0, rankedClassrooms.length - 3); // 第三名倒數的位置
        const thresholdScore = rankedClassrooms[thresholdIndex].totalScore;
        // 分數「小於等於」第三名倒數的分數者，全部列入待改進
        bottom3 = rankedClassrooms.filter(c => c.totalScore <= thresholdScore);
      }
      
      result[grade] = {
        grade: grade,
        totalClassrooms: rankedClassrooms.length,
        top3: top3,              // 特優（名次1-3）
        excellent3: excellent3,  // 優等（名次4-6）
        bottom3: bottom3,        // 待改進（最後三個名次）
        allClassrooms: rankedClassrooms
      };
    });
    
    return {
      success: true,
      weekStart: Utilities.formatDate(weekStart, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      weekEnd: Utilities.formatDate(weekEnd, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      weekBaseScore: ORDER_WEEKLY_BASE_SCORE,
      statistics: result
    };
  } catch (error) {
    Logger.log('getWeeklyStatistics 發生錯誤：' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 匯出每週排名到試算表工作表（可再由使用者下載成 Excel）
 * 會建立/覆蓋一個名稱為「每週排名_yyyyMMdd」的工作表，內容為各年級完整排名。
 *
 * @param {string} weekStartDate 週開始日期 (yyyy-MM-dd，選填，空白=本週)
 * @return {Object} 匯出結果
 */
function exportWeeklyStatisticsToSheet(weekStartDate) {
  try {
    const weekly = getWeeklyStatistics(weekStartDate);
    
    if (!weekly || weekly.error || weekly.success === false) {
      return {
        success: false,
        message: '無法取得每週統計資料：' + (weekly && weekly.error ? weekly.error : '未知錯誤')
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const weekLabel = weekly.weekStart || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const sheetName = '每週排名_' + weekLabel.replace(/-/g, '');
    
    // 如果工作表已存在就清空，否則新建
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }
    
    const statistics = weekly.statistics || {};
    const grades = ['一年級', '二年級', '三年級'];
    let row = 1;
    
    // 寫入總標題
    sheet.getRange(row, 1, 1, 4).merge();
    sheet.getRange(row, 1).setValue('每週排名（統計期間：' + weekly.weekStart + ' ~ ' + weekly.weekEnd + '）');
    sheet.getRange(row, 1).setFontWeight('bold');
    row += 2;
    
    grades.forEach(function(grade) {
      const gradeStats = statistics[grade];
      if (!gradeStats || !gradeStats.allClassrooms || gradeStats.allClassrooms.length === 0) {
        return;
      }
      
      // 年級小標題
      sheet.getRange(row, 1, 1, 4).merge();
      sheet.getRange(row, 1).setValue(grade + '（共 ' + gradeStats.totalClassrooms + ' 間教室）');
      sheet.getRange(row, 1).setFontWeight('bold');
      row += 1;
      
      // 表頭
      sheet.getRange(row, 1).setValue('名次');
      sheet.getRange(row, 2).setValue('教室');
      sheet.getRange(row, 3).setValue('總分');
      sheet.getRange(row, 4).setValue('記錄數');
      sheet.getRange(row, 1, 1, 4).setFontWeight('bold');
      row += 1;
      
      // 資料列
      gradeStats.allClassrooms.forEach(function(classroom) {
        const displayName = (classroom.classroomName || classroom.classroomId) + ' (' + classroom.classroomId + ')';
        sheet.getRange(row, 1).setValue('第' + classroom.rank + '名');
        sheet.getRange(row, 2).setValue(displayName);
        sheet.getRange(row, 3).setValue(classroom.totalScore);
        sheet.getRange(row, 4).setValue(classroom.recordCount);
        row += 1;
      });
      
      row += 2; // 年級之間留空行
    });
    
    // 自動調整欄寬
    sheet.autoResizeColumns(1, 4);
    
    return {
      success: true,
      message: '已匯出每週排名到工作表：「' + sheetName + '」',
      sheetName: sheetName,
      weekStart: weekly.weekStart,
      weekEnd: weekly.weekEnd
    };
  } catch (error) {
    Logger.log('exportWeeklyStatisticsToSheet 發生錯誤：' + error.toString());
    return {
      success: false,
      message: '匯出每週排名時發生錯誤：' + error.toString()
    };
  }
}

/** 秩序績優班級 PDF 標題、附記與服務科／普通科註4規則 */
function getWeeklyHonorPdfConfig_() {
  return {
    contestShort: '秩序',
    specialKeywords: ['普通', '服務'],
    useNote4: true,
    notes: [
      '一、 週評比取前 6 名，週評分總表如附件。',
      '二、 每週二前由學務處統一公佈績優班級及名次。',
      '三、 每週三利用集會時機統一頒發獎狀，如無集會時機，則由學務主任召集受獎班級班長頒發。',
      '四、 服務科及普通科之評分，不與一般科班評比排序，惟仍依得分並列名次'
    ]
  };
}

/**
 * 匯出每週排名報表為官方「績優班級」單頁 PDF
 * 版面比照學務處紙本：標題在框線內、高一／高二／高三並排、
 * 空白年級斜線、附記與承辦人／學務主任／校長簽核欄。
 *
 * @param {string} weekStartDate 週開始日期 (yyyy-MM-dd，選填，空白=本週)
 * @return {Object} { success, message, docUrl, weekStart, weekEnd }
 */
function exportWeeklyStatisticsPdf(weekStartDate) {
  try {
    const config = getWeeklyHonorPdfConfig_();
    const weekly = getWeeklyStatistics(weekStartDate);
    if (!weekly || weekly.error || weekly.success === false) {
      return {
        success: false,
        message: '無法取得每週統計資料：' + (weekly && weekly.error ? weekly.error : '未知錯誤')
      };
    }

    const range = getWeekDateRange_(weekStartDate);
    const weekStart = range.weekStart;
    const weekEnd = range.weekEnd;
    const gradeGroups = collectWeeklyGradeScores_(weekStart, weekEnd);
    const grades = ['一年級', '二年級', '三年級'];
    const rankedData = rankWeeklyGradeGroups_(gradeGroups, grades);
    const awardLists = buildHonorAwardLists_(rankedData, grades, config);
    const meta = getHonorWeekMeta_(weekStart);

    const pdfName = meta.schoolYear + '-' + meta.semester + '_第' + meta.weekNum + '週' + config.contestShort + '績優班級.pdf';
    trashDriveFilesByName_(pdfName);

    let pdfFile;
    try {
      pdfFile = exportHonorHtmlToPdfFile_(buildHonorFormHtml_(meta, awardLists, grades, config), pdfName);
    } catch (htmlErr) {
      Logger.log('HTML 楷體 PDF 失敗，改用試算表匯出：' + htmlErr.toString());
      const ss = SpreadsheetApp.create('_tmp_honor_form_' + new Date().getTime());
      const sheet = ss.getSheets()[0];
      sheet.setName('績優班級');
      fillHonorFormSheet_(sheet, meta, awardLists, grades, config);
      SpreadsheetApp.flush();
      pdfFile = exportSheetToPdfFile_(ss, sheet, pdfName);
      try {
        DriveApp.getFileById(ss.getId()).setTrashed(true);
      } catch (trashErr) {
        Logger.log('暫存試算表刪除失敗：' + trashErr.toString());
      }
    }
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId = pdfFile.getId();
    const docUrl = pdfFile.getUrl();
    const downloadUrl = 'https://drive.google.com/uc?export=download&id=' + fileId;
    Logger.log('績優班級 PDF 已建立：' + pdfName + ' ' + docUrl);
    return {
      success: true,
      message: '已產生每週排名 PDF',
      docUrl: docUrl,
      downloadUrl: downloadUrl,
      fileId: fileId,
      weekStart: weekly.weekStart,
      weekEnd: weekly.weekEnd
    };
  } catch (error) {
    Logger.log('exportWeeklyStatisticsPdf 發生錯誤：' + error.toString());
    Logger.log('錯誤堆疊：' + (error.stack || '無堆疊資訊'));
    return {
      success: false,
      message: '匯出每週排名 PDF 時發生錯誤：' + error.toString()
    };
  }
}

function getWeekDateRange_(weekStartDate) {
  let weekStart;
  let weekEnd;
  if (weekStartDate) {
    weekStart = new Date(weekStartDate + ' 00:00:00');
    weekEnd = new Date(weekStart);
    weekEnd.setDate(weekEnd.getDate() + 6);
    weekEnd.setHours(23, 59, 59, 999);
  } else {
    const today = new Date();
    const dayOfWeek = today.getDay();
    const diff = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
    weekStart = new Date(today);
    weekStart.setDate(today.getDate() + diff);
    weekStart.setHours(0, 0, 0, 0);
    weekEnd = new Date(weekStart);
    weekEnd.setDate(weekStart.getDate() + 6);
    weekEnd.setHours(23, 59, 59, 999);
  }
  return { weekStart: weekStart, weekEnd: weekEnd };
}

function extractGradeFromClassroomName_(name) {
  if (!name) return '';
  const text = String(name);
  if (text.indexOf('一年') >= 0) return '一年級';
  if (text.indexOf('二年') >= 0) return '二年級';
  if (text.indexOf('三年') >= 0) return '三年級';
  if (text.indexOf('1年') >= 0 || (text.match(/^1/) && !text.match(/^10/))) return '一年級';
  if (text.indexOf('2年') >= 0 || (text.match(/^2/) && !text.match(/^20/))) return '二年級';
  if (text.indexOf('3年') >= 0 || (text.match(/^3/) && !text.match(/^30/))) return '三年級';
  const gradeMatch = text.match(/([一二三])/);
  if (gradeMatch) {
    if (gradeMatch[1] === '一') return '一年級';
    if (gradeMatch[1] === '二') return '二年級';
    if (gradeMatch[1] === '三') return '三年級';
  }
  return '';
}

function collectWeeklyGradeScores_(weekStart, weekEnd) {
  const sheet = getOrderScoresSheet_();
  if (!sheet) {
    throw new Error('找不到評分記錄工作表');
  }
  const data = sheet.getDataRange().getValues();
  const col = getOrderScoreColMap_(data[0]);
  const records = data.slice(1);
  const weeklyRecords = records.filter(function(row) {
    try {
      const recordDate = new Date(row[0]);
      return recordDate >= weekStart && recordDate <= weekEnd;
    } catch (err) {
      return false;
    }
  });

  const gradeGroups = {};
  weeklyRecords.forEach(function(row) {
    let grade = String(row[2] || '').trim();
    const classroomName = String(row[6] || '').trim();
    const totalScore = Number(row[col.totalScore]) || 0;
    const recordDate = new Date(row[0]);
    if (!grade && classroomName) {
      grade = extractGradeFromClassroomName_(classroomName);
    }
    if (!grade || !classroomName) return;
    if (!gradeGroups[grade]) gradeGroups[grade] = {};
    if (!gradeGroups[grade][classroomName]) {
      gradeGroups[grade][classroomName] = {
        classroomName: classroomName,
        monday: 0,
        tuesday: 0,
        wednesday: 0,
        thursday: 0,
        friday: 0,
        saturday: 0,
        totalScore: 0
      };
    }
    const dayOfWeek = recordDate.getDay();
    const target = gradeGroups[grade][classroomName];
    if (dayOfWeek === 1) target.monday += totalScore;
    else if (dayOfWeek === 2) target.tuesday += totalScore;
    else if (dayOfWeek === 3) target.wednesday += totalScore;
    else if (dayOfWeek === 4) target.thursday += totalScore;
    else if (dayOfWeek === 5) target.friday += totalScore;
    else if (dayOfWeek === 6) target.saturday += totalScore;
    target.totalScore += totalScore;
  });
  applyOrderWeeklyBaseToGradeGroups_(gradeGroups);
  return gradeGroups;
}

function rankWeeklyGradeGroups_(gradeGroups, grades) {
  const rankedData = {};
  grades.forEach(function(grade) {
    if (!gradeGroups[grade] || Object.keys(gradeGroups[grade]).length === 0) {
      rankedData[grade] = { top3: [], excellent3: [], allClassrooms: [] };
      return;
    }
    const classrooms = Object.values(gradeGroups[grade]);
    classrooms.sort(function(a, b) { return b.totalScore - a.totalScore; });
    let lastScore = null;
    let currentRank = 0;
    classrooms.forEach(function(cls, index) {
      if (lastScore === null || cls.totalScore !== lastScore) {
        currentRank = index + 1;
        lastScore = cls.totalScore;
      }
      cls.rank = currentRank;
    });
    const top3 = classrooms.filter(function(c) { return c.rank <= 3; });
    const excellent3 = classrooms.filter(function(c) { return c.rank >= 4; }).slice(0, 3);
    rankedData[grade] = { top3: top3, excellent3: excellent3, allClassrooms: classrooms };
  });
  return rankedData;
}

function isHonorSpecialClass_(name, keywords) {
  const text = String(name || '');
  for (let i = 0; i < keywords.length; i++) {
    if (text.indexOf(keywords[i]) >= 0) return true;
  }
  return false;
}

function applyAwardRanks_(classrooms) {
  let lastScore = null;
  let currentRank = 0;
  classrooms.forEach(function(cls, index) {
    if (lastScore === null || cls.totalScore !== lastScore) {
      currentRank = index + 1;
      lastScore = cls.totalScore;
    }
    cls.awardRank = currentRank;
  });
}

function buildHonorAwardLists_(rankedData, grades, config) {
  const lists = {};
  grades.forEach(function(grade) {
    const all = (rankedData[grade] && rankedData[grade].allClassrooms) ? rankedData[grade].allClassrooms.slice() : [];
    const regular = [];
    const special = [];
    all.forEach(function(cls) {
      if (isHonorSpecialClass_(cls.classroomName, config.specialKeywords)) {
        special.push(cls);
      } else {
        regular.push(cls);
      }
    });
    applyAwardRanks_(regular);

    const rows = [];
    regular.filter(function(c) { return c.awardRank <= 3; }).forEach(function(c) {
      rows.push({ label: '特優', name: formatHonorClassName_(c.classroomName) });
    });
    regular.filter(function(c) { return c.awardRank >= 4 && c.awardRank <= 6; }).forEach(function(c) {
      rows.push({ label: '優等', name: formatHonorClassName_(c.classroomName) });
    });

    const awardees = regular.filter(function(c) { return c.awardRank <= 6; });
    const cutoff = awardees.length ? awardees[awardees.length - 1].totalScore : null;
    const topRegular = regular.filter(function(c) { return c.awardRank <= 3; });
    const topCutoff = topRegular.length ? topRegular[topRegular.length - 1].totalScore : null;
    special.forEach(function(c) {
      if (cutoff === null || c.totalScore < cutoff) return;
      const isTop = topCutoff !== null && c.totalScore >= topCutoff;
      const base = isTop ? '特優' : '優等';
      rows.push({
        label: base + (config.useNote4 ? '註4' : ''),
        name: formatHonorClassName_(c.classroomName)
      });
    });
    lists[grade] = rows;
  });
  return lists;
}

function getHonorWeekMeta_(weekStart) {
  const tz = Session.getScriptTimeZone();
  const year = weekStart.getFullYear();
  const month = weekStart.getMonth() + 1;
  const schoolYear = month >= 9 ? year - 1911 : year - 1912;
  const semester = (month >= 9 || month <= 1) ? 1 : 2;
  let semRef;
  if (semester === 1) {
    const y = month <= 1 ? year - 1 : year;
    semRef = mondayOfWeekContaining_(y, 9, 1);
  } else {
    semRef = mondayOfWeekContaining_(year, 2, 16);
  }
  const weekNum = Math.round((weekStart.getTime() - semRef.getTime()) / (7 * 24 * 60 * 60 * 1000)) + 1;
  const friday = new Date(weekStart);
  friday.setDate(weekStart.getDate() + 4);
  const dateRange = Utilities.formatDate(weekStart, tz, 'M/d') + '-' +
    Utilities.formatDate(friday, tz, 'M/d');
  return {
    schoolYear: schoolYear,
    semester: semester,
    weekNum: Math.max(1, weekNum),
    dateRange: dateRange
  };
}

function mondayOfWeekContaining_(year, month1to12, day) {
  const d = new Date(year, month1to12 - 1, day);
  d.setHours(0, 0, 0, 0);
  const dow = d.getDay();
  const diff = dow === 0 ? -6 : 1 - dow;
  d.setDate(d.getDate() + diff);
  return d;
}

function spaceCjkTitle_(text) {
  return String(text).split('').join(' ');
}

function formatHonorClassName_(name) {
  const text = String(name || '').trim();
  if (!text || /\s/.test(text)) return text;
  const matched = text.match(/^(.+?)([一二三123])([忠孝仁愛甲乙丙丁])$/);
  if (matched) return matched[1] + ' ' + matched[2] + ' ' + matched[3];
  return text;
}

function formatHonorSubtitle_(meta) {
  const year = String(meta.schoolYear).split('').join(' ');
  return year + ' 學年度第 ' + meta.semester + ' 學期第 ' + meta.weekNum +
    ' 週(' + meta.dateRange + ')';
}

function honorTitleHtml_(contestShort) {
  const full = '中正高工生活榮譽競賽' + contestShort + '評比績優班級';
  const start = full.indexOf(contestShort);
  const end = start + String(contestShort).length;
  const parts = [];
  for (let i = 0; i < full.length; i++) {
    const ch = escapeHonorHtml_(full.charAt(i));
    if (i >= start && i < end) {
      parts.push('<span style="background:#d9d9d9;">' + ch + '</span>');
    } else {
      parts.push(ch);
    }
  }
  return parts.join(' ');
}

/** 試算表後備匯出：用中文名稱才對得到 Google 內建楷體 */
function honorFormSheetFont_() {
  return '標楷體';
}

/**
 * Google 試算表把 DFKai-SB 換成細明體。改由 Google 文件以「標楷體」匯出，
 * 外觀才接近學務處紙本。
 */
function honorFormPdfFont_() {
  return '標楷體';
}

function escapeHonorHtml_(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function honorLabelHtml_(label) {
  const text = String(label || '');
  const noteIdx = text.indexOf('註');
  if (noteIdx > 0) {
    return escapeHonorHtml_(text.substring(0, noteIdx)) +
      '<span style="font-size:10pt;vertical-align:super;">' + escapeHonorHtml_(text.substring(noteIdx)) + '</span>';
  }
  return escapeHonorHtml_(text);
}

function honorCellStyle_(extra) {
  return 'border:1px solid #000;font-family:標楷體,Iansui,DFKai-SB,serif;' +
    'text-align:center;vertical-align:middle;' + (extra || '');
}

function buildHonorFormHtml_(meta, awardLists, grades, config) {
  const title = honorTitleHtml_(config.contestShort);
  const subtitle = formatHonorSubtitle_(meta);
  const gradeNames = ['高一', '高二', '高三'];
  let maxRows = 8;
  grades.forEach(function(grade) {
    maxRows = Math.max(maxRows, (awardLists[grade] || []).length);
  });
  const empty = grades.map(function(grade) {
    return !(awardLists[grade] && awardLists[grade].length);
  });

  let html = '<html><head><meta charset="UTF-8"></head><body style="color:#000;">';
  html += '<table style="width:100%;border-collapse:collapse;border:2.25pt solid #000;">';
  html += '<tr><td colspan="6" style="' + honorCellStyle_('font-size:22pt;font-weight:bold;padding:10px 6px;') + '">' +
    title + '</td></tr>';
  html += '<tr><td colspan="6" style="' + honorCellStyle_('font-size:20pt;font-weight:bold;padding:8px 6px;') + '">' +
    escapeHonorHtml_(subtitle) + '</td></tr><tr>';
  gradeNames.forEach(function(name) {
    html += '<td colspan="2" style="' + honorCellStyle_('font-size:16pt;font-weight:bold;padding:4px;') + '">' + name + '</td>';
  });
  html += '</tr><tr>';
  gradeNames.forEach(function() {
    html += '<td style="' + honorCellStyle_('font-size:16pt;font-weight:bold;width:13%;') + '">名次</td>';
    html += '<td style="' + honorCellStyle_('font-size:16pt;font-weight:bold;width:20%;') + '">班級</td>';
  });
  html += '</tr>';

  for (let i = 0; i < maxRows; i++) {
    html += '<tr>';
    grades.forEach(function(grade, gi) {
      const rows = awardLists[grade] || [];
      if (empty[gi]) {
        if (i === 0) {
          html += '<td colspan="2" rowspan="' + maxRows + '" style="' + honorCellStyle_('font-size:72pt;') + '">／</td>';
        }
        return;
      }
      if (i < rows.length) {
        html += '<td style="' + honorCellStyle_('font-size:16pt;height:30px;') + '">' + honorLabelHtml_(rows[i].label) + '</td>';
        html += '<td style="' + honorCellStyle_('font-size:16pt;') + '">' + escapeHonorHtml_(rows[i].name) + '</td>';
      } else {
        html += '<td style="' + honorCellStyle_('height:30px;') + '"></td>';
        html += '<td style="' + honorCellStyle_('') + '"></td>';
      }
    });
    html += '</tr>';
  }

  html += '<tr><td colspan="6" style="' + honorCellStyle_('font-size:15pt;font-weight:bold;padding:4px;') + '">附記</td></tr>';
  const notes = config.notes.map(function(line) { return escapeHonorHtml_(line); }).join('<br>');
  html += '<tr><td colspan="6" style="' + honorCellStyle_('text-align:left;font-size:12pt;padding:8px 10px;line-height:1.55;') + '">' + notes + '</td></tr>';
  html += '</table>';
  html += '<table style="width:100%;margin-top:18px;border:none;"><tr>';
  ['承辦人', '學務主任', '校長'].forEach(function(label) {
    html += '<td style="font-family:標楷體,Iansui,serif;text-align:center;font-size:16pt;border:none;width:33%;">' +
      label + '</td>';
  });
  html += '</tr></table></body></html>';
  return html;
}

function applyHonorDocFont_(doc) {
  const font = honorFormPdfFont_();
  const body = doc.getBody();
  try {
    body.editAsText().setFontFamily(font);
  } catch (err) {}
  const tables = body.getTables();
  for (let t = 0; t < tables.length; t++) {
    const table = tables[t];
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      for (let c = 0; c < row.getNumCells(); c++) {
        try {
          row.getCell(c).editAsText().setFontFamily(font);
        } catch (cellErr) {}
      }
    }
  }
}

function exportHonorHtmlToPdfFile_(html, filename) {
  const title = '_tmp_honor_html_' + new Date().getTime();
  const boundary = '-------314159265358979323846';
  const delim = '\r\n--' + boundary + '\r\n';
  const close = '\r\n--' + boundary + '--';
  const metadata = { name: title, mimeType: MimeType.GOOGLE_DOCS };
  const payload = delim +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(metadata) + delim +
    'Content-Type: text/html; charset=UTF-8\r\n\r\n' +
    html + close;
  const createResp = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
    {
      method: 'post',
      contentType: 'multipart/related; boundary="' + boundary + '"',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      payload: payload,
      muteHttpExceptions: true
    }
  );
  if (createResp.getResponseCode() >= 300) {
    throw new Error('HTML 轉文件失敗：' + createResp.getResponseCode() + ' ' + createResp.getContentText());
  }
  const created = JSON.parse(createResp.getContentText());
  const docId = created.id;
  if (!docId) {
    throw new Error('HTML 轉文件未取得檔案 ID');
  }
  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    body.setMarginTop(40);
    body.setMarginBottom(32);
    body.setMarginLeft(40);
    body.setMarginRight(40);
    applyHonorDocFont_(doc);
    doc.saveAndClose();
    Utilities.sleep(1000);
    const pdfResp = UrlFetchApp.fetch(
      'https://docs.google.com/document/d/' + docId + '/export?format=pdf',
      {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      }
    );
    if (pdfResp.getResponseCode() !== 200) {
      throw new Error('文件匯出 PDF 失敗：' + pdfResp.getResponseCode());
    }
    return DriveApp.createFile(pdfResp.getBlob().setName(filename));
  } finally {
    try {
      DriveApp.getFileById(docId).setTrashed(true);
    } catch (trashErr) {
      Logger.log('暫存文件刪除失敗：' + trashErr.toString());
    }
  }
}

function trashDriveFilesByName_(name) {
  const files = DriveApp.getFilesByName(name);
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
}

function setHonorRankLabel_(cell, label, fontFamily) {
  cell.setHorizontalAlignment('center').setVerticalAlignment('middle');
  cell.setFontFamily(fontFamily).setFontColor('#000000');
  const noteIdx = label.indexOf('註');
  if (noteIdx > 0) {
    const rich = SpreadsheetApp.newRichTextValue()
      .setText(label)
      .setTextStyle(SpreadsheetApp.newTextStyle().setFontFamily(fontFamily).setFontSize(16).build())
      .setTextStyle(noteIdx, label.length, SpreadsheetApp.newTextStyle().setFontFamily(fontFamily).setFontSize(9).build())
      .build();
    cell.setRichTextValue(rich);
  } else {
    cell.setFontSize(16).setValue(label);
  }
}

function fillHonorFormSheet_(sheet, meta, awardLists, grades, config) {
  const font = honorFormSheetFont_();
  const minDataRows = 8;
  let maxRows = minDataRows;
  grades.forEach(function(grade) {
    maxRows = Math.max(maxRows, (awardLists[grade] || []).length);
  });

  const titleRow = 1;
  const subRow = 2;
  const gradeHeaderRow = 3;
  const colHeaderRow = 4;
  const dataStart = 5;
  const dataEnd = dataStart + maxRows - 1;
  const noteTitleRow = dataEnd + 1;
  const noteBodyRow = dataEnd + 2;
  const signRow = noteBodyRow + 2;
  const lastCol = 6;

  sheet.setHiddenGridlines(true);
  [78, 132, 78, 132, 78, 132].forEach(function(w, i) {
    sheet.setColumnWidth(i + 1, w);
  });
  sheet.hideColumns(7, 14);

  const title = spaceCjkTitle_('中正高工生活榮譽競賽' + config.contestShort + '評比績優班級');
  const subtitle = formatHonorSubtitle_(meta);

  sheet.getRange(titleRow, 1, 1, lastCol).merge();
  sheet.getRange(titleRow, 1).setValue(title)
    .setFontFamily(font).setFontSize(22).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(titleRow, 42);

  sheet.getRange(subRow, 1, 1, lastCol).merge();
  sheet.getRange(subRow, 1).setValue(subtitle)
    .setFontFamily(font).setFontSize(20).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(subRow, 36);

  const gradeNames = ['高一', '高二', '高三'];
  gradeNames.forEach(function(name, i) {
    const col = i * 2 + 1;
    sheet.getRange(gradeHeaderRow, col, 1, 2).merge();
    sheet.getRange(gradeHeaderRow, col).setValue(name)
      .setFontFamily(font).setFontSize(16).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange(colHeaderRow, col).setValue('名次');
    sheet.getRange(colHeaderRow, col + 1).setValue('班級');
  });
  sheet.getRange(colHeaderRow, 1, 1, lastCol)
    .setFontFamily(font).setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(gradeHeaderRow, 30);
  sheet.setRowHeight(colHeaderRow, 28);

  for (let i = 0; i < maxRows; i++) {
    const row = dataStart + i;
    sheet.setRowHeight(row, 30);
    grades.forEach(function(grade, gi) {
      const rows = awardLists[grade] || [];
      if (rows.length === 0) return;
      const col = gi * 2 + 1;
      if (i < rows.length) {
        setHonorRankLabel_(sheet.getRange(row, col), rows[i].label, font);
        sheet.getRange(row, col + 1).setValue(rows[i].name)
          .setFontFamily(font).setFontSize(16)
          .setHorizontalAlignment('center').setVerticalAlignment('middle');
      }
    });
  }

  grades.forEach(function(grade, gi) {
    const rows = awardLists[grade] || [];
    if (rows.length > 0) return;
    const col = gi * 2 + 1;
    sheet.getRange(dataStart, col, maxRows, 2).merge();
    sheet.getRange(dataStart, col).setValue('／')
      .setFontFamily(font).setFontSize(72)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFontColor('#000000');
  });

  sheet.getRange(noteTitleRow, 1, 1, lastCol).merge();
  sheet.getRange(noteTitleRow, 1).setValue('附記')
    .setFontFamily(font).setFontSize(15).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(noteTitleRow, 28);

  sheet.getRange(noteBodyRow, 1, 1, lastCol).merge();
  sheet.getRange(noteBodyRow, 1).setValue(config.notes.join('\n'))
    .setFontFamily(font).setFontSize(12)
    .setHorizontalAlignment('left').setVerticalAlignment('top')
    .setWrap(true);
  sheet.setRowHeight(noteBodyRow, 120);

  const formLastRow = noteBodyRow;
  const formRange = sheet.getRange(titleRow, 1, formLastRow, lastCol);
  formRange.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  formRange.setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(titleRow, 1, formLastRow, lastCol)
    .setBackground('#FFFFFF').setFontColor('#000000');

  sheet.setRowHeight(noteBodyRow + 1, 18);
  sheet.getRange(signRow, 1, 1, 2).merge();
  sheet.getRange(signRow, 3, 1, 2).merge();
  sheet.getRange(signRow, 5, 1, 2).merge();
  sheet.getRange(signRow, 1).setValue('承辦人');
  sheet.getRange(signRow, 3).setValue('學務主任');
  sheet.getRange(signRow, 5).setValue('校長');
  sheet.getRange(signRow, 1, 1, lastCol)
    .setFontFamily(font).setFontSize(16)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(false, false, false, false, false, false);
  sheet.setRowHeight(signRow, 36);
}

function exportSheetToPdfFile_(ss, sheet, filename) {
  const gid = sheet.getSheetId();
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
    '/export?exportFormat=pdf&format=pdf&size=A4&portrait=true&fitw=true' +
    '&sheetnames=false&printtitle=false&pagenum=UNDEFINED&gridlines=false&fzr=false' +
    '&gid=' + gid +
    '&top_margin=0.55&bottom_margin=0.45&left_margin=0.55&right_margin=0.55' +
    '&horizontal_alignment=CENTER&vertical_alignment=TOP';
  try {
    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() === 200) {
      return DriveApp.createFile(resp.getBlob().setName(filename));
    }
    Logger.log('PDF export HTTP ' + resp.getResponseCode() + '，改用 getAs');
  } catch (err) {
    Logger.log('UrlFetch PDF 失敗：' + err.toString() + '，改用 getAs');
  }
  return DriveApp.createFile(ss.getAs(MimeType.PDF).setName(filename));
}

/**
 * 匯出每週排名為 CSV 字串，給前端下載
 * 格式類似 PDF：班級名稱、星期一～星期六、總分、排名、名次（週三無評分時顯示 0）
 *
 * @param {string} weekStartDate 週開始日期 (yyyy-MM-dd，選填，空白=本週)
 * @return {Object} { success, message, filename, csvContent }
 */
function exportWeeklyStatisticsCsv(weekStartDate) {
  try {
    const sheet = getOrderScoresSheet_();
    
    if (!sheet) {
      return {
        success: false,
        message: '找不到評分記錄工作表'
      };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return {
        success: false,
        message: '沒有評分記錄'
      };
    }
    const col = getOrderScoreColMap_(data[0]);
    
    // 計算週的開始和結束日期
    let weekStart;
    let weekEnd;
    
    if (weekStartDate) {
      weekStart = new Date(weekStartDate + ' 00:00:00');
      weekEnd = new Date(weekStart);
      weekEnd.setDate(weekEnd.getDate() + 6);
      weekEnd.setHours(23, 59, 59, 999);
    } else {
      const today = new Date();
      const dayOfWeek = today.getDay();
      const diff = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
      weekStart = new Date(today);
      weekStart.setDate(today.getDate() + diff);
      weekStart.setHours(0, 0, 0, 0);
      weekEnd = new Date(weekStart);
      weekEnd.setDate(weekStart.getDate() + 6);
      weekEnd.setHours(23, 59, 59, 999);
    }
    
    // 取得資料列（跳過標題列）
    const records = data.slice(1);
    
    // 過濾本週的記錄
    const weeklyRecords = records.filter(row => {
      try {
        const recordDate = new Date(row[0]);
        return recordDate >= weekStart && recordDate <= weekEnd;
      } catch (err) {
        return false;
      }
    });
    
    // 從教室名稱提取年級的函數
    function extractGradeFromName(name) {
      if (!name) return '';
      const text = String(name);
      if (text.includes('一年')) return '一年級';
      if (text.includes('二年')) return '二年級';
      if (text.includes('三年')) return '三年級';
      if (text.includes('1年') || (text.match(/^1/) && !text.match(/^10/))) return '一年級';
      if (text.includes('2年') || (text.match(/^2/) && !text.match(/^20/))) return '二年級';
      if (text.includes('3年') || (text.match(/^3/) && !text.match(/^30/))) return '三年級';
      const gradeMatch = text.match(/([一二三])/);
      if (gradeMatch) {
        const char = gradeMatch[1];
        if (char === '一') return '一年級';
        if (char === '二') return '二年級';
        if (char === '三') return '三年級';
      }
      return '';
    }
    
    // 按年級和教室名稱分組，計算每日分數
    const gradeGroups = {};
    
    weeklyRecords.forEach(row => {
      let grade = String(row[2] || '').trim();
      const classroomId = String(row[5] || '').trim();
      const classroomName = String(row[6] || '').trim();
      const totalScore = Number(row[col.totalScore]) || 0;
      const recordDate = new Date(row[0]);
      
      const classroomKey = classroomName;
      
      if (!grade && classroomName) {
        grade = extractGradeFromName(classroomName);
      }
      
      if (!grade || !classroomName) return;
      
      if (!gradeGroups[grade]) {
        gradeGroups[grade] = {};
      }
      
      if (!gradeGroups[grade][classroomKey]) {
        gradeGroups[grade][classroomKey] = {
          classroomId: classroomId,
          classroomName: classroomName,
          grade: grade,
          monday: 0,
          tuesday: 0,
          wednesday: 0,
          thursday: 0,
          friday: 0,
          saturday: 0,
          totalScore: 0,
          recordCount: 0
        };
      }
      
      // 取得星期幾（0=週日, 1=週一, ..., 6=週六）
      const dayOfWeek = recordDate.getDay();
      
      // 根據星期幾累加分數
      if (dayOfWeek === 1) {
        gradeGroups[grade][classroomKey].monday += totalScore;
      } else if (dayOfWeek === 2) {
        gradeGroups[grade][classroomKey].tuesday += totalScore;
      } else if (dayOfWeek === 3) {
        gradeGroups[grade][classroomKey].wednesday += totalScore;
      } else if (dayOfWeek === 4) {
        gradeGroups[grade][classroomKey].thursday += totalScore;
      } else if (dayOfWeek === 5) {
        gradeGroups[grade][classroomKey].friday += totalScore;
      } else if (dayOfWeek === 6) {
        gradeGroups[grade][classroomKey].saturday += totalScore;
      }
      
      gradeGroups[grade][classroomKey].totalScore += totalScore;
      gradeGroups[grade][classroomKey].recordCount += 1;
    });

    applyOrderWeeklyBaseToGradeGroups_(gradeGroups);
    
    // 計算排名並生成 CSV
    const grades = ['一年級', '二年級', '三年級'];
    const lines = [];
    
    // 格式化日期為中文格式（例如：10/27-10/31）
    const formatDateRange = function(start, end) {
      const startStr = Utilities.formatDate(start, Session.getScriptTimeZone(), 'M/d');
      const endStr = Utilities.formatDate(end, Session.getScriptTimeZone(), 'M/d');
      return startStr + '-' + endStr;
    };
    
    // 計算學年度和學期（簡化處理，假設從 9 月開始為第一學期）
    const year = weekStart.getFullYear();
    const month = weekStart.getMonth() + 1;
    const schoolYear = month >= 9 ? year - 1911 : year - 1912; // 民國年
    const semester = month >= 9 || month <= 1 ? 1 : 2;
    
    // 計算週數（簡化處理，從學期開始計算）
    const weekNum = Math.floor((weekStart - new Date(year - (month >= 9 ? 0 : 1), 8, 1)) / (7 * 24 * 60 * 60 * 1000)) + 1;
    
    // ========== 第一頁：績優班級 ==========
    lines.push('中正高工生活榮譽競賽秩序評比績優班級');
    lines.push(schoolYear + '學年度第' + semester + '學期第' + weekNum + '週(' + formatDateRange(weekStart, weekEnd) + ')');
    lines.push(''); // 空行
    
    // 先計算所有年級的排名
    const rankedData = {};
    grades.forEach(function(grade) {
      if (!gradeGroups[grade]) {
        rankedData[grade] = { top3: [], excellent3: [], allClassrooms: [] };
        return;
      }
      
      const classrooms = Object.values(gradeGroups[grade]);
      
      // 按總分降序排序
      classrooms.sort((a, b) => b.totalScore - a.totalScore);
      
      // 計算排名（同分同名次）
      let lastScore = null;
      let currentRank = 0;
      
      classrooms.forEach((cls, index) => {
        if (lastScore === null || cls.totalScore !== lastScore) {
          currentRank = index + 1;
          lastScore = cls.totalScore;
        }
        cls.rank = currentRank;
      });
      
      // 特優：名次 <= 3
      const top3 = classrooms.filter(c => c.rank <= 3);
      
      // 優等：從名次 >= 4 開始往後取最多三個班級
      const excellentCandidates = classrooms.filter(c => c.rank >= 4);
      const excellent3 = excellentCandidates.slice(0, 3);
      
      rankedData[grade] = {
        top3: top3,
        excellent3: excellent3,
        allClassrooms: classrooms
      };
    });
    
    // 第一頁表格標頭
    lines.push('高一,,高二,,高三');
    lines.push('名次,班級,名次,班級,名次,班級');
    
    // 找出三個年級中特優和優等的最大數量
    const maxTop3 = Math.max(
      rankedData['一年級'].top3.length,
      rankedData['二年級'].top3.length,
      rankedData['三年級'].top3.length
    );
    const maxExcellent3 = Math.max(
      rankedData['一年級'].excellent3.length,
      rankedData['二年級'].excellent3.length,
      rankedData['三年級'].excellent3.length
    );
    
    // 輸出特優
    for (let i = 0; i < maxTop3; i++) {
      const row = [];
      grades.forEach(function(grade) {
        const top3 = rankedData[grade].top3;
        if (i < top3.length) {
          row.push('前三名', '"' + top3[i].classroomName.replace(/"/g, '""') + '"');
        } else {
          row.push('', '');
        }
      });
      lines.push(row.join(','));
    }
    
    // 輸出優等
    for (let i = 0; i < maxExcellent3; i++) {
      const row = [];
      grades.forEach(function(grade) {
        const excellent3 = rankedData[grade].excellent3;
        if (i < excellent3.length) {
          const isSpecial = excellent3[i].classroomName.includes('普通'); // 普通科標記
          row.push('四至六名' + (isSpecial ? '*' : ''), '"' + excellent3[i].classroomName.replace(/"/g, '""') + '"');
        } else {
          row.push('', '');
        }
      });
      lines.push(row.join(','));
    }
    
    lines.push(''); // 空行
    lines.push('附記');
    lines.push('一、週評比取前6名，若因名次重複而超過6個班級，則增額授獎；每日評分細項如共享雲端資料夾附件所示。');
    lines.push('二、每週由學務處統一公佈績優班級及名次。');
    lines.push('三、利用集會時機統一頒發獎狀，如無集會時機，則由學務主任召集受獎班級風紀股長頒發或放班級櫃。');
    lines.push('四、普通科教室區納入評比與排名，不占名額。');
    lines.push(''); // 空行
    lines.push(''); // 空行
    
    // ========== 第二頁：詳細評分明細 ==========
    lines.push('中正高工生活榮譽競賽秩序評比績優班級 ' + schoolYear + '學年度第' + semester + '學期第' + weekNum + '週(' + formatDateRange(weekStart, weekEnd) + ')');
    lines.push(['', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '總分', '排名', '名次'].join(','));
    
    grades.forEach(function(grade) {
      if (!rankedData[grade] || !rankedData[grade].allClassrooms) return;
      
      const classrooms = rankedData[grade].allClassrooms;
      
      // 生成 CSV 行
      classrooms.forEach(function(classroom) {
        const name = classroom.classroomName || '';
        const rankLabel = classroom.rank || '';
        const rankText = classroom.rank <= 3 ? '前三名' : (classroom.rank >= 4 && classroom.rank <= 6 ? '四至六名' : '');
        
        const row = [
          '"' + name.replace(/"/g, '""') + '"',
          classroom.monday,
          classroom.tuesday,
          classroom.wednesday,
          classroom.thursday,
          classroom.friday,
          classroom.saturday,
          classroom.totalScore,
          rankText || '',
          rankLabel || ''
        ];
        lines.push(row.join(','));
      });
    });
    
    // 使用 UTF-8 BOM，讓 Excel 開啟時顯示中文正常
    const csvContent = '\uFEFF' + lines.join('\r\n');
    
    const weekLabel = Utilities.formatDate(weekStart, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const filename = '每週排名_' + weekLabel.replace(/-/g, '') + '.csv';
    
    const result = {
      success: true,
      message: '已產生每週排名 CSV',
      filename: filename,
      csvContent: csvContent,
      weekStart: Utilities.formatDate(weekStart, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      weekEnd: Utilities.formatDate(weekEnd, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    };

    // 同時在試算表新增／更新「每週排名_YYYYMMDD」分頁
    try {
      const sheetResult = exportWeeklyStatisticsToSheet(weekLabel);
      if (sheetResult && sheetResult.success) {
        result.sheetName = sheetResult.sheetName;
        result.message = '已產生每週排名 CSV，並新增工作表「' + sheetResult.sheetName + '」';
      } else if (sheetResult && sheetResult.message) {
        result.sheetMessage = sheetResult.message;
      }
    } catch (sheetErr) {
      result.sheetMessage = sheetErr.toString();
    }

    return result;
  } catch (error) {
    Logger.log('exportWeeklyStatisticsCsv 發生錯誤：' + error.toString());
    return {
      success: false,
      message: '匯出每週排名 CSV 時發生錯誤：' + error.toString()
    };
  }
}

/**
 * 針對指定教室名稱（例如「控制一忠」）做單週觀察：
 * - 列出本週所有相關紀錄
 * - 計算總分加總與筆數
 * 執行方式：在 Apps Script 編輯器中選擇 debugSingleClassWeek('控制一忠') 或直接在函式欄選 debugSingleClassWeek 後按執行。
 *
 * @param {string} classroomNameFilter 教室名稱關鍵字，例如「控制一忠」
 * @param {string} weekStartDate 週開始日期 (yyyy-MM-dd，可選，空白=本週)
 * @return {Object} 調試結果（在執行結果視窗中可直接看到）
 */
function debugSingleClassWeek(classroomNameFilter, weekStartDate) {
  try {
    if (!classroomNameFilter || String(classroomNameFilter).trim() === '') {
      classroomNameFilter = '控制一忠'; // 預設
    }
    classroomNameFilter = String(classroomNameFilter).trim();

    const sheet = getOrderScoresSheet_();
    if (!sheet) {
      return { success: false, message: '找不到評分記錄工作表' };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: false, message: '沒有評分記錄' };
    }
    const col = getOrderScoreColMap_(data[0]);

    // 計算週的開始與結束日期
    let weekStart;
    let weekEnd;
    if (weekStartDate) {
      weekStart = new Date(weekStartDate + ' 00:00:00');
      weekEnd = new Date(weekStart);
      weekEnd.setDate(weekEnd.getDate() + 6);
      weekEnd.setHours(23, 59, 59, 999);
    } else {
      const today = new Date();
      const dayOfWeek = today.getDay(); // 0=週日, 1=週一, ..., 6=週六
      const diff = dayOfWeek === 0 ? -6 : 1 - dayOfWeek; // 調整為週一開始
      weekStart = new Date(today);
      weekStart.setDate(today.getDate() + diff);
      weekStart.setHours(0, 0, 0, 0);
      weekEnd = new Date(weekStart);
      weekEnd.setDate(weekStart.getDate() + 6);
      weekEnd.setHours(23, 59, 59, 999);
    }

    const records = data.slice(1); // 去掉標題列

    // 篩出本週且教室名稱包含指定關鍵字的紀錄
    const matched = [];
    let sumTotalScore = 0;

    records.forEach(row => {
      try {
        const time = row[0];          // 評分時間
        const evalGrade = row[1];     // 評分年級
        const beEvalGrade = row[2];   // 被評年級
        const evaluatorType = row[3]; // 評分人員類型
        const evaluator = row[4];     // 評分人員
        const classroomId = row[5];   // 教室編號
        const classroomName = row[6]; // 教室名稱
        const area = row[col.area];          // 區域
        const timeSlot = row[col.timeSlot];      // 時段
        const photoCount = row[col.deduction];    // 扣分
        const bonus = row[col.bonus];
        const totalScore = Number(row[col.totalScore]) || 0; // 單次總分

        // 時間在本週內？
        const recordDate = new Date(time);
        if (!(recordDate >= weekStart && recordDate <= weekEnd)) {
          return;
        }

        // 教室名稱包含關鍵字？
        if (!classroomName || String(classroomName).indexOf(classroomNameFilter) === -1) {
          return;
        }

        // 累加
        sumTotalScore += totalScore;

        matched.push({
          time: String(time),
          evaluatorGrade: String(evalGrade || ''),
          evaluatedGrade: String(beEvalGrade || ''),
          evaluatorType: String(evaluatorType || ''),
          evaluator: String(evaluator || ''),
          classroomId: String(classroomId || ''),
          classroomName: String(classroomName || ''),
          area: String(area || ''),
          timeSlot: String(timeSlot || ''),
          photoCount: Number(photoCount) || 0,
          bonus: Number(bonus) || 0,
          totalScore: totalScore
        });
      } catch (e) {
        // 單筆錯誤不影響其他紀錄
      }
    });

    return {
      success: true,
      classroomNameFilter: classroomNameFilter,
      weekStart: Utilities.formatDate(weekStart, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      weekEnd: Utilities.formatDate(weekEnd, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      recordCount: matched.length,
      sumTotalScore: sumTotalScore,
      records: matched
    };
  } catch (error) {
    return {
      success: false,
      message: 'debugSingleClassWeek 發生錯誤：' + error.toString()
    };
  }
}

/**
 * 取得多個班級的週比較數據（包含週一到週五的分數）
 * 注意：這裡以「教室名稱」為比較與彙總的主鍵（避免同樓層不同班被合併）
 * @param {Array} classroomNames 教室名稱陣列
 * @param {string} weekStartDate 週開始日期 (yyyy-MM-dd 格式，選填，預設為本週)
 * @param {string} grade 年級（選填，用於驗證教室是否屬於該年級）
 * @return {Object} 比較數據，包含每個班級的週一到週五分數和總分
 */
function getClassroomComparison(classroomIds, weekStartDate, grade) {
  try {
    const sheet = getOrderScoresSheet_();
    
    if (!sheet) {
      return { success: false, error: '找不到評分記錄工作表' };
    }
    
    if (!classroomIds || classroomIds.length === 0) {
      return { success: false, error: '請至少選擇一個班級' };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: false, error: '沒有評分記錄' };
    }
    const col = getOrderScoreColMap_(data[0]);
    
    // 計算週的開始和結束日期
    let weekStart;
    let weekEnd;
    
    if (weekStartDate) {
      weekStart = new Date(weekStartDate + ' 00:00:00');
      weekEnd = new Date(weekStart);
      weekEnd.setDate(weekEnd.getDate() + 6);
      weekEnd.setHours(23, 59, 59, 999);
    } else {
      // 預設為本週（週一到週日）
      const today = new Date();
      const dayOfWeek = today.getDay(); // 0=週日, 1=週一, ..., 6=週六
      const diff = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
      weekStart = new Date(today);
      weekStart.setDate(today.getDate() + diff);
      weekStart.setHours(0, 0, 0, 0);
      weekEnd = new Date(weekStart);
      weekEnd.setDate(weekStart.getDate() + 6);
      weekEnd.setHours(23, 59, 59, 999);
    }
    
    // 取得資料列（跳過標題列）
    const records = data.slice(1);
    
    // 過濾本週的記錄，並只保留選中的「教室名稱」對應的班級
    const weeklyRecords = records.filter(row => {
      try {
        const recordDate = new Date(row[0]); // 評分時間在第1欄（索引0）
        const classroomName = String(row[6] || '').trim(); // 教室名稱在第7欄（索引6）
        return recordDate >= weekStart && recordDate <= weekEnd && classroomIds.includes(classroomName);
      } catch (err) {
        return false;
      }
    });
    
    // 初始化每個班級的數據結構（以教室名稱為 key）
    const classroomData = {};
    classroomIds.forEach(name => {
      const key = String(name || '').trim();
      if (!key) return;
      classroomData[key] = {
        classroomId: '',      // 之後從記錄中補上實際教室編號
        classroomName: key,   // 直接使用教室名稱作為主鍵
        monday: 0,
        tuesday: 0,
        wednesday: 0,
        thursday: 0,
        friday: 0,
        weekend: 0, // 週六週日
        totalScore: 0,
        recordCount: 0
      };
    });
    
    // 處理每筆記錄（依教室名稱彙總）
    weeklyRecords.forEach(row => {
      const classroomId = String(row[5] || '').trim();   // 教室編號在第6欄
      const classroomName = String(row[6] || '').trim(); // 教室名稱在第7欄
      const totalScore = Number(row[col.totalScore]) || 0; // 總分
      const recordDate = new Date(row[0]);
      
      if (!classroomData[classroomName]) return;
      
      // 更新教室名稱（如果還沒有）
      if (!classroomData[classroomName].classroomName && classroomName) {
        classroomData[classroomName].classroomName = classroomName;
      }
      // 補上教室編號（如果還沒有）
      if (!classroomData[classroomName].classroomId && classroomId) {
        classroomData[classroomName].classroomId = classroomId;
      }
      
      // 取得星期幾（0=週日, 1=週一, ..., 6=週六）
      const dayOfWeek = recordDate.getDay();
      
      // 根據星期幾累加分數
      if (dayOfWeek === 1) {
        classroomData[classroomName].monday += totalScore;
      } else if (dayOfWeek === 2) {
        classroomData[classroomName].tuesday += totalScore;
      } else if (dayOfWeek === 3) {
        classroomData[classroomName].wednesday += totalScore;
      } else if (dayOfWeek === 4) {
        classroomData[classroomName].thursday += totalScore;
      } else if (dayOfWeek === 5) {
        classroomData[classroomName].friday += totalScore;
      } else {
        classroomData[classroomName].weekend += totalScore;
      }
      
      // 累加總分
      classroomData[classroomName].totalScore += totalScore;
      classroomData[classroomName].recordCount += 1;
    });

    applyOrderWeeklyBaseToMap_(classroomData);
    
    // 如果指定了年級，驗證教室是否屬於該年級
    if (grade && grade.trim() !== '') {
      // 取得所有教室資訊
      const allClassrooms = getAllClassrooms();
      const classroomGradeMapById = {};
      const classroomGradeMapByName = {};
      
      // 建立 教室ID / 教室名稱 到年級的對應表
      allClassrooms.forEach(room => {
        if (room) {
          if (room.id) {
            classroomGradeMapById[room.id] = room.grade || '';
          }
          if (room.name) {
            classroomGradeMapByName[room.name] = room.grade || '';
          }
        }
      });
      
      // 從教室名稱提取年級的函數
      function extractGradeFromName(name) {
        if (!name) return '';
        const text = String(name);
        if (text.includes('一年')) return '一年級';
        if (text.includes('二年')) return '二年級';
        if (text.includes('三年')) return '三年級';
        if (text.includes('1年') || (text.match(/^1/) && !text.match(/^10/))) return '一年級';
        if (text.includes('2年') || (text.match(/^2/) && !text.match(/^20/))) return '二年級';
        if (text.includes('3年') || (text.match(/^3/) && !text.match(/^30/))) return '三年級';
        if (text.includes('一')) return '一年級';
        if (text.includes('二')) return '二年級';
        if (text.includes('三')) return '三年級';
        return '';
      }
      
      // 檢查年級是否匹配
      function isGradeMatch(roomGrade, targetGrade) {
        if (!roomGrade || roomGrade === '') return false;
        const roomGradeTrim = String(roomGrade).trim();
        const targetGradeTrim = String(targetGrade).trim();
        if (roomGradeTrim === targetGradeTrim) return true;
        
        const gradeVariants = {
          '一年級': ['一年級', '1年級', '1', '一', '一級', '一年'],
          '二年級': ['二年級', '2年級', '2', '二', '二級', '二年'],
          '三年級': ['三年級', '3年級', '3', '三', '三級', '三年']
        };
        
        const targetVariants = gradeVariants[targetGradeTrim] || [targetGradeTrim];
        return targetVariants.some(variant => roomGradeTrim === variant || roomGradeTrim.includes(variant) || variant.includes(roomGradeTrim));
      }
      
      // 過濾掉不屬於所選年級的教室
      Object.keys(classroomData).forEach(classroomNameKey => {
        const data = classroomData[classroomNameKey];
        let roomGrade = '';
        
        // 先用教室名稱找年級
        if (data.classroomName && classroomGradeMapByName[data.classroomName]) {
          roomGrade = classroomGradeMapByName[data.classroomName];
        }
        // 再用教室編號找
        if (!roomGrade && data.classroomId && classroomGradeMapById[data.classroomId]) {
          roomGrade = classroomGradeMapById[data.classroomId];
        }
        
        // 如果還是沒有年級資訊，從教室名稱推斷
        if (!roomGrade && data.classroomName) {
          roomGrade = extractGradeFromName(data.classroomName);
        }
        
        // 如果年級不匹配，移除該教室
        if (!isGradeMatch(roomGrade, grade)) {
          delete classroomData[classroomNameKey];
        }
      });
    }
    
    // 轉換為陣列並按總分排序（分數越高越好，即扣分越少越好）
    const result = Object.values(classroomData).sort((a, b) => b.totalScore - a.totalScore);
    
    // 計算週一到週五的日期字串
    const mondayDate = new Date(weekStart);
    const tuesdayDate = new Date(weekStart);
    tuesdayDate.setDate(tuesdayDate.getDate() + 1);
    const wednesdayDate = new Date(weekStart);
    wednesdayDate.setDate(wednesdayDate.getDate() + 2);
    const thursdayDate = new Date(weekStart);
    thursdayDate.setDate(thursdayDate.getDate() + 3);
    const fridayDate = new Date(weekStart);
    fridayDate.setDate(fridayDate.getDate() + 4);
    
    return {
      success: true,
      weekStart: Utilities.formatDate(weekStart, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      weekEnd: Utilities.formatDate(weekEnd, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      mondayDate: Utilities.formatDate(mondayDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      tuesdayDate: Utilities.formatDate(tuesdayDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      wednesdayDate: Utilities.formatDate(wednesdayDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      thursdayDate: Utilities.formatDate(thursdayDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      fridayDate: Utilities.formatDate(fridayDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      weekBaseScore: ORDER_WEEKLY_BASE_SCORE,
      classrooms: result
    };
  } catch (error) {
    Logger.log('getClassroomComparison 發生錯誤：' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 取得學期總成績統計（按年級分組，計算特優和優等）
 * 學期成績 = 各週總分加總 ÷ 有評分的週數；該週無人評分則不列入，跳過的班當週以 75 分計。
 * @param {string} startDate 開始日期 (yyyy-MM-dd 格式，必填)
 * @param {string} endDate 結束日期 (yyyy-MM-dd 格式，必填)
 * @param {string} grade 年級（選填，如果指定則只返回該年級的統計）
 * @return {Object} 學期統計資料，包含各年級的特優和優等
 */
function getSemesterStatistics(startDate, endDate, grade) {
  try {
    const sheet = getOrderScoresSheet_();
    
    if (!sheet) {
      return { success: false, error: '找不到評分記錄工作表' };
    }
    
    if (!startDate || !endDate) {
      return { success: false, error: '請指定開始日期和結束日期' };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: false, error: '沒有評分記錄' };
    }
    const col = getOrderScoreColMap_(data[0]);
    
    // 計算日期範圍
    const semesterStart = new Date(startDate + ' 00:00:00');
    const semesterEnd = new Date(endDate + ' 23:59:59');
    
    if (semesterStart > semesterEnd) {
      return { success: false, error: '開始日期不能晚於結束日期' };
    }
    
    // 取得資料列（跳過標題列）
    const records = data.slice(1);
    
    // 過濾學期範圍內的記錄
    const semesterRecords = records.filter(row => {
      try {
        const recordDate = new Date(row[0]); // 評分時間在第1欄（索引0）
        return recordDate >= semesterStart && recordDate <= semesterEnd;
      } catch (err) {
        return false;
      }
    });
    
    // 按年級和「教室名稱」分組，計算總分
    const gradeGroups = {};
    
    // 從教室名稱提取年級的函數
    function extractGradeFromName(name) {
      if (!name) return '';
      const text = String(name);
      if (text.includes('一年')) return '一年級';
      if (text.includes('二年')) return '二年級';
      if (text.includes('三年')) return '三年級';
      if (text.includes('1年') || (text.match(/^1/) && !text.match(/^10/))) return '一年級';
      if (text.includes('2年') || (text.match(/^2/) && !text.match(/^20/))) return '二年級';
      if (text.includes('3年') || (text.match(/^3/) && !text.match(/^30/))) return '三年級';
      if (text.includes('一')) return '一年級';
      if (text.includes('二')) return '二年級';
      if (text.includes('三')) return '三年級';
      return '';
    }
    
    semesterRecords.forEach(row => {
      let recordGrade = String(row[2] || '').trim(); // 被評年級
      const classroomId = String(row[5] || '').trim();   // 教室編號
      const classroomName = String(row[6] || '').trim(); // 教室名稱
      const totalScore = Number(row[col.totalScore]) || 0;           // 單次總分
      
      // 如果被評年級為空，從教室名稱推斷
      if (!recordGrade && classroomName) {
        recordGrade = extractGradeFromName(classroomName);
      }
      
      // 如果指定了年級，只處理該年級的記錄
      if (grade && grade.trim() !== '') {
        const gradeVariants = {
          '一年級': ['一年級', '1年級', '1', '一', '一級', '一年'],
          '二年級': ['二年級', '2年級', '2', '二', '二級', '二年'],
          '三年級': ['三年級', '3年級', '3', '三', '三級', '三年']
        };
        const targetVariants = gradeVariants[grade.trim()] || [grade.trim()];
        const isMatch = targetVariants.some(variant => 
          recordGrade === variant || recordGrade.includes(variant) || variant.includes(recordGrade)
        );
        if (!isMatch) return;
      }
      
      // 沒有年級或沒有教室名稱就不計算，避免錯誤分類
      if (!recordGrade || !classroomName) return;
      
      // 使用「教室名稱」作為分組 key，只依名稱統計
      const classroomKey = classroomName;
      
      // 初始化年級組
      if (!gradeGroups[recordGrade]) {
        gradeGroups[recordGrade] = {};
      }
      
      // 初始化教室（以 classroomName 為索引）
      if (!gradeGroups[recordGrade][classroomKey]) {
        gradeGroups[recordGrade][classroomKey] = {
          classroomId: classroomId,
          classroomName: classroomName,
          grade: recordGrade,
          weekDeltas: {},
          recordCount: 0
        };
      }

      const weekKey = getOrderWeekStartKey_(row[0]);
      if (weekKey) {
        gradeGroups[recordGrade][classroomKey].weekDeltas[weekKey] =
          (Number(gradeGroups[recordGrade][classroomKey].weekDeltas[weekKey]) || 0) + totalScore;
      }
      gradeGroups[recordGrade][classroomKey].recordCount += 1;
    });

    try {
      getAllClassrooms().forEach(function(room) {
        const recordGrade = String(room.grade || '').trim();
        const classroomName = String(room.name || '').trim();
        if (!recordGrade || !classroomName) return;
        if (grade && grade.trim() !== '') {
          const gradeVariants = {
            '一年級': ['一年級', '1年級', '1', '一', '一級', '一年'],
            '二年級': ['二年級', '2年級', '2', '二', '二級', '二年'],
            '三年級': ['三年級', '3年級', '3', '三', '三級', '三年']
          };
          const targetVariants = gradeVariants[grade.trim()] || [grade.trim()];
          const isMatch = targetVariants.some(function(variant) {
            return recordGrade === variant || recordGrade.indexOf(variant) >= 0 || variant.indexOf(recordGrade) >= 0;
          });
          if (!isMatch) return;
        }
        if (!gradeGroups[recordGrade]) return;
        if (!gradeGroups[recordGrade][classroomName]) {
          gradeGroups[recordGrade][classroomName] = {
            classroomId: room.id || '',
            classroomName: classroomName,
            grade: recordGrade,
            weekDeltas: {},
            recordCount: 0
          };
        }
      });
    } catch (seedErr) {
      Logger.log('學期教室清單補齊失敗：' + seedErr.toString());
    }

    Object.keys(gradeGroups).forEach(function(gradeKey) {
      const weekSet = {};
      Object.keys(gradeGroups[gradeKey]).forEach(function(classKey) {
        const deltas = gradeGroups[gradeKey][classKey].weekDeltas || {};
        Object.keys(deltas).forEach(function(weekKey) {
          weekSet[weekKey] = true;
        });
      });
      const weekKeys = Object.keys(weekSet);
      const weekCount = weekKeys.length;
      Object.keys(gradeGroups[gradeKey]).forEach(function(classKey) {
        const item = gradeGroups[gradeKey][classKey];
        let sumWeekly = 0;
        weekKeys.forEach(function(weekKey) {
          sumWeekly += ORDER_WEEKLY_BASE_SCORE + (Number(item.weekDeltas[weekKey]) || 0);
        });
        item.scoredWeeks = weekCount;
        item.totalScore = weekCount > 0 ? Math.round((sumWeekly / weekCount) * 10) / 10 : 0;
        delete item.weekDeltas;
      });
    });
    
    // 轉換為陣列並排序，找出特優和優等（處理同分名次）
    const result = {};
    
    Object.keys(gradeGroups).forEach(gradeKey => {
      const classrooms = Object.values(gradeGroups[gradeKey]);
      
      // 按總分降序排序（分數越高越好）
      classrooms.sort((a, b) => b.totalScore - a.totalScore);
      
      // 依總分計算名次：
      // - 同分者名次相同
      // - 例如第4、5、6名同分，三者皆為第4名，下一名為第7名
      const rankedClassrooms = [];
      let lastScore = null;
      let currentRank = 0;
      
      for (let i = 0; i < classrooms.length; i++) {
        const cls = classrooms[i];
        if (lastScore === null || cls.totalScore !== lastScore) {
          currentRank = i + 1;
          lastScore = cls.totalScore;
        }
        rankedClassrooms.push({
          rank: currentRank,
          ...cls
        });
      }
      
      // 特優：學期第1名（同分並列）
      const top3 = rankedClassrooms.filter(c => c.rank <= 1);
      
      // 優勝：學期第2-3名
      const excellent3 = rankedClassrooms.filter(c => c.rank >= 2 && c.rank <= 3);
      
      // 待改進：最後「三個位置」往後擴展，將同分者一起納入
      let bottom3 = [];
      if (rankedClassrooms.length <= 3) {
        bottom3 = rankedClassrooms.slice();
      } else {
        const thresholdIndex = Math.max(0, rankedClassrooms.length - 3);
        const thresholdScore = rankedClassrooms[thresholdIndex].totalScore;
        bottom3 = rankedClassrooms.filter(c => c.totalScore <= thresholdScore);
      }
      
      result[gradeKey] = {
        grade: gradeKey,
        totalClassrooms: rankedClassrooms.length,
        top3: top3,
        excellent3: excellent3,
        bottom3: bottom3,
        allClassrooms: rankedClassrooms
      };
    });
    
    return {
      success: true,
      startDate: Utilities.formatDate(semesterStart, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      endDate: Utilities.formatDate(semesterEnd, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      weekBaseScore: ORDER_WEEKLY_BASE_SCORE,
      statistics: result
    };
  } catch (error) {
    Logger.log('getSemesterStatistics 發生錯誤：' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 取得所有教室清單（用於查看頁面的下拉選單）
 * @return {Array} 教室清單陣列
 */
function getAllClassrooms() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CLASSROOMS);
  
  if (!sheet) {
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return [];
  }
  
  // 跳過標題列，取得資料
  const classrooms = data.slice(1)
    .map(row => {
      const id = String(row[0] || '').trim();
      const name = String(row[1] || '').trim();
      const order = Number(row[2]) || 999;
      
      // 從教室名稱提取年級（支援「資訊一忠」「電機二孝」等職科班名）
      let grade = '';
      if (name.includes('一年')) grade = '一年級';
      else if (name.includes('二年')) grade = '二年級';
      else if (name.includes('三年')) grade = '三年級';
      else if (name.includes('1年') || (name.match(/^1/) && !name.match(/^10/))) grade = '一年級';
      else if (name.includes('2年') || (name.match(/^2/) && !name.match(/^20/))) grade = '二年級';
      else if (name.includes('3年') || (name.match(/^3/) && !name.match(/^30/))) grade = '三年級';
      else {
        const gradeMatch = name.match(/([一二三])/);
        if (gradeMatch) {
          if (gradeMatch[1] === '一') grade = '一年級';
          else if (gradeMatch[1] === '二') grade = '二年級';
          else if (gradeMatch[1] === '三') grade = '三年級';
        }
      }
      
      // 如果無法從名稱提取，嘗試從年級欄位讀取（向後相容）
      if (!grade && row.length > 3) {
        grade = String(row[3] || '').trim();
      }
      
      return {
        id: id,
        name: name,
        grade: grade || '',
        order: order
      };
    })
    .filter(room => room.id && room.name)
    .sort((a, b) => a.order - b.order);
  
  return classrooms;
}

/**
 * Web App 入口：供 GitHub 靜態前端以 JSON API 呼叫。
 * 前端使用 POST + text/plain，避免瀏覽器 CORS 預檢失敗。
 */
function doGet(e) {
  const page = e && e.parameter && e.parameter.page;
  if (page === 'bridge') {
    return HtmlService.createHtmlOutput(getBridgeHtml_())
      .setTitle('School Score API Bridge')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return handleApiRequest(e);
}

function doPost(e) {
  return handleApiRequest(e);
}

function jsonOutput_(data, callback) {
  const json = JSON.stringify(data);
  if (callback) {
    const safeName = String(callback).replace(/[^A-Za-z0-9_$.]/g, '');
    return ContentService
      .createTextOutput(safeName + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function handleApiRequest(e) {
  let callback = '';
  try {
    let request = {};
    if (e && e.parameter && e.parameter.callback) {
      callback = e.parameter.callback;
    }
    if (e && e.postData && e.postData.contents) {
      request = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      if (e.parameter.payload) {
        try {
          request = JSON.parse(e.parameter.payload);
        } catch (parsePayloadErr) {
          request = {};
        }
      }
      if (!request.action) {
        request.action = e.parameter.action || '';
        request.authPassword = request.authPassword || e.parameter.authPassword || '';
        if (e.parameter.args) {
          try {
            request.args = JSON.parse(e.parameter.args);
          } catch (parseArgsErr) {
            request.args = [];
          }
        }
      }
      if (typeof request.args === 'string') {
        try {
          request.args = JSON.parse(request.args);
        } catch (parseErr) {
          request.args = [];
        }
      }
    }

    const action = request.action || '';
    const args = Array.isArray(request.args) ? request.args : [];
    const authPassword = request.authPassword || '';
    const result = dispatchAction_(action, args, authPassword);
    return jsonOutput_(result, callback);
  } catch (error) {
    return jsonOutput_({
      __exception: true,
      success: false,
      message: error.message || error.toString()
    }, callback);
  }
}

function getAuthLoginType_(password) {
  const inputPwd = String(password || '').trim();
  if (!inputPwd) return '';
  const teacherPassword = String(getScoreSystemPassword() || '').trim();
  const studentPassword = String(getStudentPassword() || '').trim();
  const adminPassword = String(getAdminPassword() || '').trim();
  if (inputPwd === teacherPassword) return 'teacher';
  if (inputPwd === studentPassword) return 'student';
  if (inputPwd === adminPassword) return 'admin';
  return '';
}

function requireAuthPassword_(password) {
  if (getAuthLoginType_(password)) return true;
  const inputPwd = String(password || '').trim();
  if (!inputPwd) {
    throw new Error('未授權：請先從查詢頁輸入密碼進入評分系統');
  }
  throw new Error('未授權：密碼錯誤');
}

function requireAdminPassword_(password) {
  const inputPwd = String(password || '').trim();
  if (!inputPwd) {
    throw new Error('未授權：請先輸入管理員密碼');
  }
  const adminPassword = String(getAdminPassword() || '').trim();
  if (inputPwd === adminPassword) {
    return true;
  }
  throw new Error('未授權：此功能僅限管理員');
}

function dispatchAction_(action, args, authPassword) {
  const mutatingActions = {
    saveScore: true,
    uploadSinglePhoto: true,
    uploadPhotoChunk: true,
    finalizePhotoUpload: true
  };
  const adminActions = {
    exportWeeklyStatisticsCsv: true,
    exportWeeklyStatisticsPdf: true,
    exportWeeklyStatisticsToSheet: true
  };
  if (adminActions[action]) {
    requireAdminPassword_(authPassword);
  } else if (mutatingActions[action]) {
    requireAuthPassword_(authPassword);
  }

  switch (action) {
    case 'getGrades':
      return getGrades();
    case 'getAreas':
      return getAreas();
    case 'getTimeSlots':
      return getTimeSlots();
    case 'getEvaluatorTypes':
      return getEvaluatorTypes();
    case 'getClassrooms':
      return getClassrooms(args[0]);
    case 'getAllClassrooms':
      return getAllClassrooms();
    case 'getPhotoFolderLink':
      return getPhotoFolderLink();
    case 'uploadSinglePhoto':
      return uploadSinglePhoto(args[0], args[1]);
    case 'uploadPhotoChunk':
      return uploadPhotoChunk(args[0], args[1], args[2], args[3]);
    case 'finalizePhotoUpload':
      return finalizePhotoUpload(args[0]);
    case 'saveScore':
      return saveScore(args[0], authPassword);
    case 'getScoreRecords':
      return getScoreRecords(args[0] || '', args[1] || '', args[2] || '');
    case 'verifyScoreSystemPassword':
      return verifyScoreSystemPassword(args[0]);
    case 'getClassroomComparison':
      return getClassroomComparison(args[0], args[1], args[2]);
    case 'getWeeklyStatistics':
      return getWeeklyStatistics(args[0]);
    case 'exportWeeklyStatisticsCsv':
      return exportWeeklyStatisticsCsv(args[0]);
    case 'exportWeeklyStatisticsPdf':
      return exportWeeklyStatisticsPdf(args[0]);
    case 'exportWeeklyStatisticsToSheet':
      return exportWeeklyStatisticsToSheet(args[0]);
    case 'getSemesterStatistics':
      return getSemesterStatistics(args[0], args[1], args[2]);
    case 'getViewScoresUrl':
      return '';
    case 'getScoreSystemUrl':
      return '';
    case 'ping':
      return { success: true, message: 'ok' };
    default:
      return {
        __exception: true,
        success: false,
        message: '未知的 action：' + action
      };
  }
}

/**
 * 供前端橋接頁呼叫的唯一公開入口（其餘函式不要從 iframe 直接呼叫）。
 */
function apiCall(action, args, authPassword) {
  return dispatchAction_(action, args || [], authPassword || '');
}

/**
 * 取得評分系統密碼（第一組密碼，給老師/管理員使用）
 * @return {string} 密碼
 */
function getAuthPasswords_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('scr_auth_pwds');
  if (hit) {
    try {
      const parsed = JSON.parse(hit);
      if (parsed && typeof parsed.teacher === 'string') return parsed;
    } catch (e) {}
  }
  const passwords = { teacher: '1234', student: '5678', admin: '9999' };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    if (!sheet) {
      initializeSheets();
      sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    }
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0] || '').trim();
      const val = String(data[i][1] || '').trim();
      if (key === '評分系統密碼') passwords.teacher = val || passwords.teacher;
      else if (key === '學生密碼') passwords.student = val || passwords.student;
      else if (key === '管理員密碼') passwords.admin = val || passwords.admin;
    }
  } catch (error) {
    Logger.log('取得密碼失敗：' + error.toString());
  }
  cache.put('scr_auth_pwds', JSON.stringify(passwords), 300);
  return passwords;
}

function getScoreSystemPassword() {
  return getAuthPasswords_().teacher;
}

function getStudentPassword() {
  return getAuthPasswords_().student;
}

function getAdminPassword() {
  return getAuthPasswords_().admin;
}

function verifyScoreSystemPassword(inputPassword) {
  try {
    const passwords = getAuthPasswords_();
    const inputPwd = String(inputPassword || '').trim();
    let loginType = '';
    let isCorrect = false;
    if (inputPwd && inputPwd === String(passwords.teacher || '').trim()) {
      isCorrect = true;
      loginType = 'teacher';
    } else if (inputPwd && inputPwd === String(passwords.student || '').trim()) {
      isCorrect = true;
      loginType = 'student';
    } else if (inputPwd && inputPwd === String(passwords.admin || '').trim()) {
      isCorrect = true;
      loginType = 'admin';
    }
    return {
      success: isCorrect,
      message: isCorrect ? '密碼正確' : '密碼錯誤，請重新輸入',
      url: '',
      loginType: loginType
    };
  } catch (error) {
    Logger.log('驗證密碼失敗：' + error.toString());
    return {
      success: false,
      message: '驗證密碼時發生錯誤：' + error.toString(),
      url: '',
      loginType: ''
    };
  }
}

/**
 * 獲取評分系統URL
 * @return {string} 評分系統的URL
 */
function getScoreSystemUrl() {
  try {
    const service = ScriptApp.getService();
    if (service) {
      return service.getUrl() + '?page=score';
    }
    return '';
  } catch (error) {
    Logger.log('獲取評分系統URL失敗：' + error.toString());
    return '';
  }
}

/**
 * 獲取評分結果查詢頁 URL
 * @return {string} 查詢頁的 Web App URL
 */
function getViewScoresUrl() {
  try {
    const service = ScriptApp.getService();
    if (service) {
      return service.getUrl() + '?page=view';
    }
    return '';
  } catch (error) {
    Logger.log('獲取查詢頁URL失敗：' + error.toString());
    return '';
  }
}

/**
 * GitHub 靜態頁嵌入的 API 橋接頁：以 google.script.run 執行後端函式，再用 postMessage 回傳。
 */
function getBridgeHtml_() {
  return [
    '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body>',
    '<script>',
    '(function() {',
    '  function reply(source, payload) {',
    '    try { if (source) source.postMessage(payload, "*"); } catch (err) {}',
    '    try { window.top.postMessage(payload, "*"); } catch (err) {}',
    '    try { window.parent.postMessage(payload, "*"); } catch (err) {}',
    '  }',
    '  window.addEventListener("message", function(e) {',
    '    var msg = e.data || {};',
    '    if (!msg || msg.type !== "gas-call") return;',
    '    var action = String(msg.action || "");',
    '    var args = Array.isArray(msg.args) ? msg.args : [];',
    '    var id = msg.id;',
    '    var authPassword = msg.authPassword || "";',
    '    google.script.run',
    '      .withSuccessHandler(function(result) {',
    '        if (result && result.__exception) {',
    '          reply(e.source, { type: "gas-result", id: id, ok: false, message: result.message || "後端發生錯誤" });',
    '          return;',
    '        }',
    '        reply(e.source, { type: "gas-result", id: id, ok: true, result: result });',
    '      })',
    '      .withFailureHandler(function(error) {',
    '        var message = (error && error.message) ? error.message : String(error);',
    '        reply(e.source, { type: "gas-result", id: id, ok: false, message: message });',
    '      })',
    '      .apiCall(action, args, authPassword);',
    '  });',
    '  function pingReady() {',
    '    var payload = { type: "gas-ready" };',
    '    try { window.top.postMessage(payload, "*"); } catch (err) {}',
    '    try { window.parent.postMessage(payload, "*"); } catch (err) {}',
    '  }',
    '  pingReady();',
    '  setTimeout(pingReady, 50);',
    '  setTimeout(pingReady, 300);',
    '})();',
    '</script></body></html>'
  ].join('');
}

