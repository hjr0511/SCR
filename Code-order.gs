/**
 * 學校秩序評分系統（生輔組）
 * 午休 A–G 加扣分、集會／巡堂登記；不拍照。
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

// 評分人員類型
const EVALUATOR_TYPES = ['風紀股長', '校安人員', '師長', '巡堂教師', '生輔組', '校長及業管'];

// 添加一個全局緩存來存儲資料夾引用（參考 uploadFileToDrive 的成功策略）
const folderCache = {};

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
      '直接扣分',
      '加分',
      '總分（加分-扣分）', 
      '備註',
      '細項說明'
    ]]);
    scoresSheet.getRange(1, 1, 1, 15).setFontWeight('bold');
  } else {
    // 如果工作表已存在，檢查是否有「照片連結」欄位
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
    // 優化：簡化驗證，僅檢查關鍵條件以提升速度
    if (!base64Data || !filename) {
      throw new Error('參數無效');
    }
    
    const folderName = '學校整潔評分照片';
    let folder = null;
    
    // 先檢查緩存（優化：優先使用緩存，避免重複查找和 API 調用）
    if (folderCache[folderName]) {
      folder = folderCache[folderName];
    } else {
      // 取得當前試算表的檔案 ID（僅在需要時調用）
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const file = DriveApp.getFileById(ss.getId());
      
      // 取得試算表的父資料夾
      const parents = file.getParents();
      const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
      
      // 優化：使用 getFoldersByName 直接查找，比循環查找更快
      const folders = parentFolder.getFoldersByName(folderName);
      if (folders.hasNext()) {
        folder = folders.next();
        folderCache[folderName] = folder;
      } else {
        // 如果找不到，創建新資料夾
        try {
          folder = parentFolder.createFolder(folderName);
          folderCache[folderName] = folder;
        } catch (createError) {
          // 如果創建資料夾失敗，使用父資料夾作為備選方案
          folder = parentFolder;
        }
      }
    }
    
    if (!folder) {
      throw new Error('無法找到或創建照片資料夾');
    }
    
    // 處理 base64 資料並解碼（優化：使用更高效的方法）
    const commaIndex = base64Data.indexOf(',');
    const base64Content = commaIndex >= 0 ? base64Data.substring(commaIndex + 1) : base64Data;
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Content), 'image/jpeg', filename);
    
    // 上傳檔案到 Drive（優化：直接上傳並獲取 URL，完全移除權限設置以最大化速度）
    const uploadedFile = folder.createFile(blob);
    
    // 優化：完全移除權限設置以提升速度（權限可以在後台或需要時再設置）
    // 如果需要公開訪問，可以在上傳後通過其他方式設置，不阻塞上傳流程
    
    // 立即回傳檔案連結
    return uploadedFile.getUrl();
  } catch (error) {
    const errorMsg = '上傳照片失敗：' + error.toString();
    Logger.log(errorMsg);
    Logger.log('錯誤堆疊：' + (error.stack || '無堆疊資訊'));
    Logger.log('錯誤類型：' + (error.name || '未知'));
    Logger.log('錯誤訊息：' + (error.message || '無訊息'));
    Logger.log('執行用戶：' + Session.getActiveUser().getEmail());
    
    // 提供更詳細的錯誤訊息，幫助診斷學校 Google Workspace 的問題
    let detailedError = errorMsg;
    const errorStr = error.toString().toLowerCase();
    if (errorStr.includes('permission') || errorStr.includes('權限') || errorStr.includes('access')) {
      detailedError += ' 可能的原因：1) Google Workspace 管理員限制了 Drive API 存取權限；2) 需要授權「查看和管理您的 Google Drive 檔案」權限；3) 帳戶沒有創建資料夾的權限。請聯繫 Google Workspace 管理員檢查相關設定。';
    } else if (errorStr.includes('quota') || errorStr.includes('配額')) {
      detailedError += ' 可能的原因：Google Drive 儲存空間不足。';
    } else if (errorStr.includes('rate') || errorStr.includes('限制')) {
      detailedError += ' 可能的原因：上傳頻率過高，請稍後再試。';
    }
    
    // 拋出錯誤以便上層可以捕獲詳細資訊
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssFile = DriveApp.getFileById(ss.getId());
    const folderName = '學校整潔評分照片';
    
    // 取得試算表所在的資料夾
    let parentFolder;
    const parentFolders = ssFile.getParents();
    if (parentFolders.hasNext()) {
      parentFolder = parentFolders.next();
      const folders = parentFolder.getFoldersByName(folderName);
      if (folders.hasNext()) {
        return folders.next().getUrl();
      }
    } else {
      // 試算表在根目錄
      const folders = DriveApp.getFoldersByName(folderName);
      if (folders.hasNext()) {
        return folders.next().getUrl();
      }
    }
    
    // 如果資料夾不存在，返回空字串
    return '';
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

function saveScore(scoreData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.SCORES);
    
    if (!sheet) {
      initializeSheets();
      sheet = ss.getSheetByName(SHEET_NAMES.SCORES);
    }
    
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const category = String(scoreData.timeSlot || scoreData.category || '').trim();
    const items = scoreData.items || {};
    const notes = String(scoreData.notes || '').trim();
    
    let bonus = Number(scoreData.bonus) || 0;
    let deduction = Number(scoreData.directDeduction) || 0;
    let itemSummary = String(scoreData.itemSummary || scoreData.photoLinks || '').trim();
    
    const isNap = category.indexOf('午休') >= 0;
    const hasItemFlags = items.A || items.B || items.C || items.D || items.E || items.F || items.G;
    if (isNap || hasItemFlags) {
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
      const otherDeduction = Number(items.G) || 0;
      if (otherDeduction > 0) {
        deduction += otherDeduction;
        parts.push('G其他行為-' + otherDeduction);
      }
      itemSummary = parts.join('；');
    }
    
    const totalScore = bonus - deduction;
    
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
      deduction,
      bonus,
      totalScore,
      notes,
      itemSummary
    ];
    
    sheet.appendRow(newRow);
    
    return {
      success: true,
      message: '評分記錄已儲存',
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SCORES);
  
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SCORES);
    
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
        const photoDeduction = Number(row[9]) || 0;
        const directDeduction = Number(row[10]) || 0;
        const bonus = Number(row[11]) || 0;
        const totalScore = Number(row[12]) || 0;
        const notes = row[13] ? String(row[13]) : '';
        const photoLinks = row[14] ? String(row[14]) : '';
        
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
          directDeduction: directDeduction,
          bonus: bonus,
          totalScore: totalScore,
          notes: notes,
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SCORES);
    
    if (!sheet) {
      return { error: '找不到評分記錄工作表' };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { error: '沒有評分記錄' };
    }
    
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
    // row[2] = 被評年級, row[5] = 教室編號, row[6] = 教室名稱, row[12] = 總分
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
      const totalScore = Number(row[12]) || 0;           // 單次總分
      
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

/**
 * 匯出每週排名報表為 Google 文件
 * 流程：
 * 1. 取得每週統計資料（特優、優等、所有班級）
 * 2. 取得每日分數資料（週一到週六）
 * 3. 建立 Google 文件，按照 PDF 格式排版：
 *    - 第一頁：績優班級表格（高一、高二、高三並排）
 *    - 第二頁：詳細評分明細（所有班級的每日分數）
 * 4. 設定為「知道連結的人可檢視」
 * 5. 回傳文件連結
 *
 * @param {string} weekStartDate 週開始日期 (yyyy-MM-dd，選填，空白=本週)
 * @return {Object} { success, message, docUrl, weekStart, weekEnd }
 */
function exportWeeklyStatisticsPdf(weekStartDate) {
  try {
    // 取得每週統計資料
    const weekly = getWeeklyStatistics(weekStartDate);
    if (!weekly || weekly.error || weekly.success === false) {
      return {
        success: false,
        message: '無法取得每週統計資料：' + (weekly && weekly.error ? weekly.error : '未知錯誤')
      };
    }

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

    // 取得每日分數資料
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SCORES);
    if (!sheet) {
      return { success: false, message: '找不到評分記錄工作表' };
    }
    const data = sheet.getDataRange().getValues();
    const records = data.slice(1);
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
    Logger.log('開始處理每週記錄，共 ' + weeklyRecords.length + ' 筆');
    
    weeklyRecords.forEach((row, index) => {
      try {
        let grade = String(row[2] || '').trim();
        const classroomName = String(row[6] || '').trim();
        const totalScore = Number(row[12]) || 0;
        const recordDate = new Date(row[0]);
        
        if (!grade && classroomName) {
          grade = extractGradeFromName(classroomName);
        }
        if (!grade || !classroomName) {
          Logger.log('記錄 ' + index + ' 跳過：年級=' + grade + ', 教室名稱=' + classroomName);
          return;
        }
        
        const classroomKey = classroomName;
        if (!gradeGroups[grade]) {
          gradeGroups[grade] = {};
        }
        if (!gradeGroups[grade][classroomKey]) {
          gradeGroups[grade][classroomKey] = {
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
      } catch (err) {
        Logger.log('處理記錄 ' + index + ' 時發生錯誤：' + err.toString());
      }
    });
    
    Logger.log('gradeGroups 處理完成，包含年級：' + Object.keys(gradeGroups).join(', '));

    // 計算排名
    const statistics = weekly.statistics || {};
    const grades = ['一年級', '二年級', '三年級'];
    const rankedData = {};
    
    Logger.log('開始計算排名，gradeGroups 包含的年級：' + Object.keys(gradeGroups).join(', '));
    Logger.log('每週記錄數：' + weeklyRecords.length);
    
    grades.forEach(function(grade) {
      if (!gradeGroups[grade] || Object.keys(gradeGroups[grade]).length === 0) {
        Logger.log('年級 ' + grade + ' 沒有資料');
        rankedData[grade] = { top3: [], excellent3: [], allClassrooms: [] };
        return;
      }
      const classrooms = Object.values(gradeGroups[grade]);
      Logger.log('年級 ' + grade + ' 有 ' + classrooms.length + ' 個班級');
      classrooms.sort((a, b) => b.totalScore - a.totalScore);
      let lastScore = null;
      let currentRank = 0;
      classrooms.forEach((cls, index) => {
        if (lastScore === null || cls.totalScore !== lastScore) {
          currentRank = index + 1;
          lastScore = cls.totalScore;
        }
        cls.rank = currentRank;
      });
      const top3 = classrooms.filter(c => c.rank <= 3);
      const excellentCandidates = classrooms.filter(c => c.rank >= 4);
      const excellent3 = excellentCandidates.slice(0, 3);
      rankedData[grade] = { top3: top3, excellent3: excellent3, allClassrooms: classrooms };
      Logger.log('年級 ' + grade + ' 排名完成：特優 ' + top3.length + ' 個，優等 ' + excellent3.length + ' 個，總共 ' + classrooms.length + ' 個班級');
    });

    // 計算學年度和週數
    const year = weekStart.getFullYear();
    const month = weekStart.getMonth() + 1;
    const schoolYear = month >= 9 ? year - 1911 : year - 1912;
    const semester = month >= 9 || month <= 1 ? 1 : 2;
    const weekNum = Math.floor((weekStart - new Date(year - (month >= 9 ? 0 : 1), 8, 1)) / (7 * 24 * 60 * 60 * 1000)) + 1;
    const formatDateRange = function(start, end) {
      const startStr = Utilities.formatDate(start, Session.getScriptTimeZone(), 'M/d');
      const endStr = Utilities.formatDate(end, Session.getScriptTimeZone(), 'M/d');
      return startStr + '-' + endStr;
    };

    // 建立 Google 文件
    const weekLabel = weekly.weekStart.replace(/-/g, '');
    const docName = '每週排名_' + weekLabel;
    
    Logger.log('開始建立 Google 文件：' + docName);
    
    // 檢查並刪除舊檔案
    const files = DriveApp.getFilesByName(docName);
    let deletedCount = 0;
    while (files.hasNext()) {
      files.next().setTrashed(true);
      deletedCount++;
    }
    if (deletedCount > 0) {
      Logger.log('已刪除 ' + deletedCount + ' 個同名檔案');
    }
    
    // 建立新文件
    // 注意：DocumentApp.create() 需要正確的 OAuth2 權限
    // 如果遇到權限錯誤，請在 Apps Script 編輯器中執行一次任意函式來授權
    let doc;
    try {
      // 嘗試建立文件
      doc = DocumentApp.create(docName);
      Logger.log('Google 文件已建立，ID：' + doc.getId());
    } catch (createError) {
      const errorMsg = createError.toString();
      Logger.log('建立 Google 文件失敗：' + errorMsg);
      
      // 如果是 OAuth2 權限錯誤，提供更清楚的錯誤訊息
      if (errorMsg.includes('OAuth2') || errorMsg.includes('scope') || errorMsg.includes('permission') || 
          errorMsg.includes('權限') || errorMsg.includes('DocumentApp.create') || 
          errorMsg.includes('https://www.googleapis.com/auth/documents')) {
        throw new Error('無法建立 Google 文件：權限不足。\n\n解決方法：\n1. 開啟 Google Apps Script 編輯器（https://script.google.com）\n2. 找到此專案並開啟\n3. 點擊「執行」按鈕，選擇任意函式（例如：initializeSheets）\n4. 點擊「授權」按鈕，允許以下權限：\n   - 查看、編輯、建立和刪除您的 Google 文件\n   - 查看和管理您的 Google Drive 檔案\n5. 授權完成後，重新嘗試匯出功能。\n\n如果問題持續，請聯繫 Google Workspace 管理員檢查是否限制了相關 API 權限。');
      }
      throw new Error('無法建立 Google 文件：' + errorMsg);
    }
    
    const body = doc.getBody();
    body.clear();

    // 設定文件格式
    body.setMarginTop(72);
    body.setMarginBottom(72);
    body.setMarginLeft(72);
    body.setMarginRight(72);

    // ========== 第一頁：績優班級 ==========
    // 標題
    const title = body.appendParagraph('中正高工生活榮譽競賽秩序評比績優班級');
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    title.editAsText().setFontSize(18).setBold(true);

    // 副標題
    const subtitle = body.appendParagraph(schoolYear + '學年度第' + semester + '學期第' + weekNum + '週(' + formatDateRange(weekStart, weekEnd) + ')');
    subtitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    subtitle.editAsText().setFontSize(14);
    body.appendParagraph(''); // 空行

    // 建立表格：高一、高二、高三並排
    const table = body.appendTable();
    // 設定表格置中對齊
    try {
      const tableParent = table.getParent();
      if (tableParent && tableParent.getType && tableParent.getType() === DocumentApp.ElementType.PARAGRAPH) {
        tableParent.asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      }
    } catch (e) {
      Logger.log('設定表格對齊時發生錯誤：' + e.toString());
    }
    const headerRow = table.appendTableRow();
    headerRow.appendTableCell('高一').setBackgroundColor('#E8E8E8').setWidth(100);
    headerRow.appendTableCell('').setWidth(80);
    headerRow.appendTableCell('高二').setBackgroundColor('#E8E8E8').setWidth(100);
    headerRow.appendTableCell('').setWidth(80);
    headerRow.appendTableCell('高三').setBackgroundColor('#E8E8E8').setWidth(100);
    headerRow.appendTableCell('').setWidth(80);
    headerRow.getCell(0).editAsText().setBold(true);
    headerRow.getCell(2).editAsText().setBold(true);
    headerRow.getCell(4).editAsText().setBold(true);

    const subHeaderRow = table.appendTableRow();
    subHeaderRow.appendTableCell('名次').setBackgroundColor('#F0F0F0');
    subHeaderRow.appendTableCell('班級').setBackgroundColor('#F0F0F0');
    subHeaderRow.appendTableCell('名次').setBackgroundColor('#F0F0F0');
    subHeaderRow.appendTableCell('班級').setBackgroundColor('#F0F0F0');
    subHeaderRow.appendTableCell('名次').setBackgroundColor('#F0F0F0');
    subHeaderRow.appendTableCell('班級').setBackgroundColor('#F0F0F0');

    // 找出最大行數（確保 rankedData 存在）
    const maxTop3 = Math.max(
      (rankedData['一年級'] && rankedData['一年級'].top3 ? rankedData['一年級'].top3.length : 0),
      (rankedData['二年級'] && rankedData['二年級'].top3 ? rankedData['二年級'].top3.length : 0),
      (rankedData['三年級'] && rankedData['三年級'].top3 ? rankedData['三年級'].top3.length : 0)
    );
    const maxExcellent3 = Math.max(
      (rankedData['一年級'] && rankedData['一年級'].excellent3 ? rankedData['一年級'].excellent3.length : 0),
      (rankedData['二年級'] && rankedData['二年級'].excellent3 ? rankedData['二年級'].excellent3.length : 0),
      (rankedData['三年級'] && rankedData['三年級'].excellent3 ? rankedData['三年級'].excellent3.length : 0)
    );
    
    Logger.log('最大特優行數：' + maxTop3 + '，最大優等行數：' + maxExcellent3);

    // 輸出特優
    if (maxTop3 > 0) {
      for (let i = 0; i < maxTop3; i++) {
        const row = table.appendTableRow();
        grades.forEach(function(grade) {
          const top3 = rankedData[grade] && rankedData[grade].top3 ? rankedData[grade].top3 : [];
          if (i < top3.length) {
            row.appendTableCell('前三名').setWidth(60);
            row.appendTableCell(top3[i].classroomName).setWidth(80);
          } else {
            row.appendTableCell('').setWidth(60);
            row.appendTableCell('').setWidth(80);
          }
        });
      }
    } else {
      // 如果沒有特優資料，顯示提示
      const row = table.appendTableRow();
      row.appendTableCell('（本週無前六名班級）');
      row.appendTableCell('');
      row.appendTableCell('（本週無前六名班級）');
      row.appendTableCell('');
      row.appendTableCell('（本週無前六名班級）');
      row.appendTableCell('');
    }

    // 輸出優等
    if (maxExcellent3 > 0) {
      for (let i = 0; i < maxExcellent3; i++) {
        const row = table.appendTableRow();
        grades.forEach(function(grade) {
          const excellent3 = rankedData[grade] && rankedData[grade].excellent3 ? rankedData[grade].excellent3 : [];
          if (i < excellent3.length) {
            const isSpecial = excellent3[i].classroomName.includes('普通');
            row.appendTableCell('四至六名' + (isSpecial ? '*' : '')).setWidth(60);
            row.appendTableCell(excellent3[i].classroomName).setWidth(80);
          } else {
            row.appendTableCell('').setWidth(60);
            row.appendTableCell('').setWidth(80);
          }
        });
      }
    }

    // 將附記直接添加到上面的表格中，確保接在一起且寬度一致
    // 使用函數來安全地設定邊框
    function setCellBorder(cell) {
      try {
        // 確保單元格有內容後再設定邊框
        if (cell.getNumChildren() > 0) {
          cell.setBorderColor('#000000');
          cell.setBorderWidth(1);
        } else {
          // 如果沒有內容，先添加一個段落
          cell.appendParagraph(' ');
          cell.setBorderColor('#000000');
          cell.setBorderWidth(1);
        }
      } catch (e) {
        Logger.log('設定單元格邊框時發生錯誤：' + e.toString());
      }
    }
    
    // 將附記直接添加到上面的表格中，確保接在一起且寬度一致
    // 附記標題（使用6個欄位，第一個欄位設定寬度，其他設為0）
    const noteTitleRow = table.appendTableRow();
    const noteTitleCell = noteTitleRow.appendTableCell('附記');
    noteTitleCell.editAsText().setBold(true);
    // 設定第一個欄位寬度為總寬度
    try {
      noteTitleCell.setWidth(520); // 總寬度：100+80+100+80+100+80 = 520
    } catch (e) {
      Logger.log('設定附記標題寬度時發生錯誤：' + e.toString());
    }
    // 添加其他5個空欄位，寬度設為0
    for (let i = 1; i < 6; i++) {
      const emptyCell = noteTitleRow.appendTableCell('');
      try {
        emptyCell.setWidth(0);
      } catch (e) {
        Logger.log('設定空欄位寬度時發生錯誤：' + e.toString());
      }
    }
    // 設定邊框
    setCellBorder(noteTitleCell);
    for (let i = 1; i < 6; i++) {
      try {
        setCellBorder(noteTitleRow.getCell(i));
      } catch (e) {
        Logger.log('設定附記標題欄位邊框時發生錯誤：' + e.toString());
      }
    }
    
    // 附記內容（四條合併在同一框線內，使用6個欄位）
    const noteContentRow = table.appendTableRow();
    const noteContentCell = noteContentRow.appendTableCell('一、週評比取前6名，若因名次重複而超過6個班級，則增額授獎；每日評分細項如共享雲端資料夾附件所示。');
    // 設定第一個欄位寬度為總寬度
    try {
      noteContentCell.setWidth(520); // 總寬度：100+80+100+80+100+80 = 520
    } catch (e) {
      Logger.log('設定附記內容寬度時發生錯誤：' + e.toString());
    }
    // 添加其他三條內容作為段落
    noteContentCell.appendParagraph('二、每週由學務處統一公佈績優班級及名次。');
    noteContentCell.appendParagraph('三、利用集會時機統一頒發獎狀，如無集會時機，則由學務主任召集受獎班級風紀股長頒發或放班級櫃。');
    noteContentCell.appendParagraph('四、普通科教室區納入評比與排名，不占名額。');
    // 添加其他5個空欄位，寬度設為0
    for (let i = 1; i < 6; i++) {
      const emptyCell = noteContentRow.appendTableCell('');
      try {
        emptyCell.setWidth(0);
      } catch (e) {
        Logger.log('設定空欄位寬度時發生錯誤：' + e.toString());
      }
    }
    // 設定邊框
    setCellBorder(noteContentCell);
    for (let i = 1; i < 6; i++) {
      try {
        setCellBorder(noteContentRow.getCell(i));
      } catch (e) {
        Logger.log('設定附記內容欄位邊框時發生錯誤：' + e.toString());
      }
    }
    
    // 簽核欄位：承辦人、學務主任、校長（同一行並排顯示，使用段落而非表格）
    body.appendParagraph(''); // 空行
    const signaturePara = body.appendParagraph('承辦人' + '\t\t\t\t\t\t\t\t\t\t' + '學務主任' + '\t\t\t\t\t\t\t\t\t\t' + '校長');
    signaturePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    // 分頁
    body.appendPageBreak();

    // ========== 第二頁：詳細評分明細 ==========
    const detailTitle = body.appendParagraph('中正高工生活榮譽競賽秩序評比績優班級 ' + schoolYear + '學年度第' + semester + '學期第' + weekNum + '週(' + formatDateRange(weekStart, weekEnd) + ')');
    detailTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    detailTitle.editAsText().setFontSize(14).setBold(true);
    body.appendParagraph(''); // 空行

    // 建立詳細表格
    const detailTable = body.appendTable();
    // 設定表格置中對齊
    try {
      const detailTableParent = detailTable.getParent();
      if (detailTableParent && detailTableParent.getType && detailTableParent.getType() === DocumentApp.ElementType.PARAGRAPH) {
        detailTableParent.asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      }
    } catch (e) {
      Logger.log('設定詳細表格對齊時發生錯誤：' + e.toString());
    }
    const detailHeaderRow = detailTable.appendTableRow();
    detailHeaderRow.appendTableCell('').setBackgroundColor('#E8E8E8');
    detailHeaderRow.appendTableCell('星期一').setBackgroundColor('#E8E8E8');
    detailHeaderRow.appendTableCell('星期二').setBackgroundColor('#E8E8E8');
    detailHeaderRow.appendTableCell('星期三').setBackgroundColor('#E8E8E8');
    detailHeaderRow.appendTableCell('星期四').setBackgroundColor('#E8E8E8');
    detailHeaderRow.appendTableCell('星期五').setBackgroundColor('#E8E8E8');
    detailHeaderRow.appendTableCell('星期六').setBackgroundColor('#E8E8E8');
    detailHeaderRow.appendTableCell('總分').setBackgroundColor('#E8E8E8');
    detailHeaderRow.appendTableCell('排名').setBackgroundColor('#E8E8E8');
    detailHeaderRow.appendTableCell('名次').setBackgroundColor('#E8E8E8');
    // 設置每個單元格的粗體（使用 getNumCells() 和 getCell() 方法）
    const numCells = detailHeaderRow.getNumCells();
    for (let i = 0; i < numCells; i++) {
      detailHeaderRow.getCell(i).editAsText().setBold(true);
    }

    let hasAnyData = false;
    grades.forEach(function(grade) {
      if (!rankedData[grade] || !rankedData[grade].allClassrooms || rankedData[grade].allClassrooms.length === 0) {
        Logger.log('年級 ' + grade + ' 沒有詳細資料可寫入');
        return;
      }
      const classrooms = rankedData[grade].allClassrooms;
      Logger.log('年級 ' + grade + ' 準備寫入 ' + classrooms.length + ' 個班級的詳細資料');
      hasAnyData = true;
      classrooms.forEach(function(classroom) {
        const row = detailTable.appendTableRow();
        const rankText = classroom.rank <= 3 ? '前三名' : (classroom.rank >= 4 && classroom.rank <= 6 ? '四至六名' : '');
        row.appendTableCell(classroom.classroomName || '');
        row.appendTableCell((classroom.monday || 0).toString());
        row.appendTableCell((classroom.tuesday || 0).toString());
        row.appendTableCell((classroom.wednesday || 0).toString());
        row.appendTableCell((classroom.thursday || 0).toString());
        row.appendTableCell((classroom.friday || 0).toString());
        row.appendTableCell((classroom.saturday || 0).toString());
        row.appendTableCell((classroom.totalScore || 0).toString());
        row.appendTableCell(rankText || '');
        row.appendTableCell((classroom.rank || '').toString());
      });
    });
    
    // 如果沒有任何資料，顯示提示訊息
    if (!hasAnyData) {
      const row = detailTable.appendTableRow();
      row.appendTableCell('（本週無評分記錄）');
      row.appendTableCell('');
      row.appendTableCell('');
      row.appendTableCell('');
      row.appendTableCell('');
      row.appendTableCell('');
      row.appendTableCell('');
      row.appendTableCell('');
      row.appendTableCell('');
      row.appendTableCell('');
      Logger.log('警告：沒有任何資料可寫入詳細表格');
    }

    // 儲存文件
    try {
      doc.saveAndClose();
      Logger.log('文件已儲存並關閉');
    } catch (saveError) {
      Logger.log('儲存文件時發生錯誤：' + saveError.toString());
      throw new Error('無法儲存 Google 文件：' + saveError.toString());
    }

    // 取得文件檔案並設定權限
    let docFile;
    let docUrl;
    try {
      docFile = DriveApp.getFileById(doc.getId());
      Logger.log('已取得文件檔案，名稱：' + docFile.getName());
      
      // 設定權限
      docFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      Logger.log('已設定文件權限為「知道連結的人可檢視」');
      
      // 取得文件 URL
      docUrl = docFile.getUrl();
      Logger.log('文件 URL：' + docUrl);
      
      if (!docUrl || docUrl.trim() === '') {
        throw new Error('無法取得文件 URL');
      }
    } catch (fileError) {
      Logger.log('取得文件檔案或 URL 時發生錯誤：' + fileError.toString());
      throw new Error('無法取得 Google 文件連結：' + fileError.toString());
    }

    Logger.log('Google 文件建立完成：' + docName);
    Logger.log('文件 ID：' + doc.getId());
    Logger.log('文件連結：' + docUrl);

    return {
      success: true,
      message: '已產生每週排名 Google 文件',
      docUrl: docUrl,
      fileId: docFile.getId(),
      weekStart: weekly.weekStart,
      weekEnd: weekly.weekEnd
    };
  } catch (error) {
    Logger.log('exportWeeklyStatisticsPdf 發生錯誤：' + error.toString());
    Logger.log('錯誤堆疊：' + (error.stack || '無堆疊資訊'));
    return {
      success: false,
      message: '匯出每週排名 Google 文件時發生錯誤：' + error.toString()
    };
  }
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SCORES);
    
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
      const totalScore = Number(row[12]) || 0;
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

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SCORES);
    if (!sheet) {
      return { success: false, message: '找不到評分記錄工作表' };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: false, message: '沒有評分記錄' };
    }

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
        const area = row[7];          // 區域
        const timeSlot = row[8];      // 時段
        const photoCount = row[9];    // 扣分項目數
        const directDeduction = row[10];
        const bonus = row[11];
        const totalScore = Number(row[12]) || 0; // 單次總分

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
          directDeduction: Number(directDeduction) || 0,
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SCORES);
    
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
      const totalScore = Number(row[12]) || 0; // 總分在第13欄（索引12）
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
 * @param {string} startDate 開始日期 (yyyy-MM-dd 格式，必填)
 * @param {string} endDate 結束日期 (yyyy-MM-dd 格式，必填)
 * @param {string} grade 年級（選填，如果指定則只返回該年級的統計）
 * @return {Object} 學期統計資料，包含各年級的特優和優等
 */
function getSemesterStatistics(startDate, endDate, grade) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SCORES);
    
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
    // row[2] = 被評年級, row[5] = 教室編號, row[6] = 教室名稱, row[12] = 總分
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
      const totalScore = Number(row[12]) || 0;           // 單次總分
      
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
          classroomId: classroomId,       // 保留原始編號做參考（例如樓層）
          classroomName: classroomName,   // 班級名稱，作為主要顯示與分組依據
          grade: recordGrade,
          totalScore: 0,
          recordCount: 0
        };
      }
      
      // 累加總分
      gradeGroups[recordGrade][classroomKey].totalScore += totalScore;
      gradeGroups[recordGrade][classroomKey].recordCount += 1;
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

function requireAuthPassword_(password) {
  const inputPwd = String(password || '').trim();
  if (!inputPwd) {
    throw new Error('未授權：請先從查詢頁輸入密碼進入評分系統');
  }
  const teacherPassword = String(getScoreSystemPassword() || '').trim();
  const studentPassword = String(getStudentPassword() || '').trim();
  const adminPassword = String(getAdminPassword() || '').trim();
  if (inputPwd === teacherPassword || inputPwd === studentPassword || inputPwd === adminPassword) {
    return true;
  }
  throw new Error('未授權：密碼錯誤');
}

function dispatchAction_(action, args, authPassword) {
  const mutatingActions = {
    saveScore: true,
    uploadSinglePhoto: true,
    uploadPhotoChunk: true,
    finalizePhotoUpload: true
  };
  if (mutatingActions[action]) {
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
    case 'uploadSinglePhoto':
      return uploadSinglePhoto(args[0], args[1]);
    case 'uploadPhotoChunk':
      return uploadPhotoChunk(args[0], args[1], args[2], args[3]);
    case 'finalizePhotoUpload':
      return finalizePhotoUpload(args[0]);
    case 'saveScore':
      return saveScore(args[0]);
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
function getScoreSystemPassword() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    
    if (!sheet) {
      initializeSheets();
      sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      // 如果沒有資料，返回預設密碼
      return '1234';
    }
    
    // 查找「評分系統密碼」設定
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === '評分系統密碼') {
        return String(data[i][1] || '').trim();
      }
    }
    
    // 如果找不到，返回預設密碼
    return '1234';
  } catch (error) {
    Logger.log('取得密碼失敗：' + error.toString());
    return '1234'; // 預設密碼
  }
}

/**
 * 取得學生密碼（第二組密碼，給學生使用）
 * @return {string} 密碼
 */
function getStudentPassword() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    
    if (!sheet) {
      initializeSheets();
      sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      // 如果沒有資料，返回預設學生密碼
      return '5678';
    }
    
    // 查找「學生密碼」設定
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === '學生密碼') {
        return String(data[i][1] || '').trim();
      }
    }
    
    // 如果找不到，返回預設學生密碼
    return '5678';
  } catch (error) {
    Logger.log('取得學生密碼失敗：' + error.toString());
    return '5678'; // 預設學生密碼
  }
}

/**
 * 取得管理員密碼（第三組密碼，給衛生組使用）
 * @return {string} 密碼
 */
function getAdminPassword() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    
    if (!sheet) {
      initializeSheets();
      sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      // 如果沒有資料，返回預設管理員密碼
      return '9999';
    }
    
    // 查找「管理員密碼」設定
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === '管理員密碼') {
        return String(data[i][1] || '').trim();
      }
    }
    
    // 如果找不到，返回預設管理員密碼
    return '9999';
  } catch (error) {
    Logger.log('取得管理員密碼失敗：' + error.toString());
    return '9999'; // 預設管理員密碼
  }
}

/**
 * 驗證評分系統密碼
 * @param {string} inputPassword 使用者輸入的密碼
 * @return {Object} 驗證結果（包含正確的跳轉URL和登入類型）
 */
function verifyScoreSystemPassword(inputPassword) {
  try {
    const teacherPassword = getScoreSystemPassword();
    const studentPassword = getStudentPassword();
    const adminPassword = getAdminPassword();
    const inputPwd = String(inputPassword || '').trim();
    
    let loginType = ''; // 'teacher'、'student' 或 'admin'
    let isCorrect = false;
    
    // 檢查是否為第一組密碼（老師）
    if (inputPwd === String(teacherPassword).trim()) {
      isCorrect = true;
      loginType = 'teacher';
    }
    // 檢查是否為第二組密碼（學生）
    else if (inputPwd === String(studentPassword).trim()) {
      isCorrect = true;
      loginType = 'student';
    }
    // 檢查是否為第三組密碼（管理員/衛生組）
    else if (inputPwd === String(adminPassword).trim()) {
      isCorrect = true;
      loginType = 'admin';
    }
    
    // 獲取Web App的URL
    let scoreSystemUrl = '';
    try {
      const service = ScriptApp.getService();
      if (service) {
        // 在 URL 中加入登入類型參數
        scoreSystemUrl = service.getUrl() + '?page=score&loginType=' + loginType;
      } else {
        // 如果無法獲取服務URL，使用當前URL的基礎部分
        scoreSystemUrl = '';
      }
    } catch (e) {
      Logger.log('獲取服務URL失敗：' + e.toString());
    }
    
    return {
      success: isCorrect,
      message: isCorrect ? '密碼正確' : '密碼錯誤，請重新輸入',
      url: scoreSystemUrl,
      loginType: loginType // 返回登入類型
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
    '    try { source.postMessage(payload, "*"); } catch (err) {}',
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
    '  try { window.parent.postMessage({ type: "gas-ready" }, "*"); } catch (err) {}',
    '})();',
    '</script></body></html>'
  ].join('');
}

