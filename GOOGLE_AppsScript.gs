/**
 * ============================================================================
 * 施工抽查表單產生器 —— 雲端同步 Apps Script 範本
 * ============================================================================
 * 使用方式：
 *   1. 開一份新的 Google 試算表（或您既有的試算表）
 *   2. 選單「擴充功能」→「Apps Script」
 *   3. 把這個檔案的全部內容貼進去（取代預設的 myFunction 內容）
 *   4. 畫面右上角「部署」→「新增部署作業」
 *        類型：網頁應用程式
 *        執行身分：我 (您自己的帳號)
 *        誰可以存取：所有人
 *   5. 部署完成後複製「網頁應用程式網址」，貼到系統的「☁️ 雲端同步設定」裡
 *   6. 回到系統，按「🔌 測試連線」確認成功即可
 *
 * 會自動建立兩個分頁：
 *   「紀錄表資料」：一列＝一個管理項目
 *   「照片」      ：一列＝一張照片，並用 =IMAGE() 公式顯示縮圖
 * 照片檔案本身會存到您 Google 雲端硬碟中、下方 DRIVE_FOLDER_NAME 指定的資料夾（會自動建立）
 * ============================================================================
 */

// ========================= 設定區：可依需求修改 =========================
const SHEET_RECORD_NAME = '紀錄表資料';
const SHEET_PHOTO_NAME = '照片';
const DRIVE_FOLDER_NAME = '施工抽查照片'; // 所有同步的照片都會存到這個 Google 雲端硬碟資料夾
// ==========================================================================

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ success: false, message: '沒有收到任何資料 (postData 為空)' });
    }

    const payload = JSON.parse(e.postData.contents);

    // 「測試連線」按鈕會送 { ping: true }，這裡直接回應成功，不寫入任何資料
    if (payload.ping) {
      return jsonResponse({ success: true, message: 'pong' });
    }

    // 系統在切換「分項工程(工項)」時會送 { action: 'nextFormNo', categoryCode }，
    // 查詢該工項目前「紀錄表資料」分頁中已使用過的最大流水號，並回傳下一號 (各工項各自獨立編號)，
    // 讓使用者不需自己手動確認流水號、也不會跳號或重複。
    if (payload.action === 'nextFormNo') {
      const nextNo = getNextFormNumber(payload.categoryCode || '');
      return jsonResponse({ success: true, nextNo: nextNo });
    }

    const submissionId = payload.submissionId || Utilities.getUuid();
    const header = payload.header || {};
    const items = payload.items || [];
    const photos = payload.photos || [];

    if (items.length > 0) {
      writeRecordRows(submissionId, header, items);
    }
    if (photos.length > 0) {
      writePhotoRows(submissionId, header, photos);
    }

    return jsonResponse({ success: true, submissionId: submissionId });
  } catch (err) {
    // 任何錯誤都回傳 success:false + 錯誤訊息，前端會顯示出來，不會讓使用者一頭霧水
    return jsonResponse({ success: false, message: String(err && err.message ? err.message : err) });
  }
}

// 瀏覽器直接打開網址 (GET) 時顯示的提示，方便確認網址有正確部署
function doGet(e) {
  return jsonResponse({ success: true, message: '此端點僅接受系統自動送出的 POST 請求，請直接在系統的「雲端同步設定」中使用「測試連線」按鈕。' });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 查詢「紀錄表資料」分頁中，指定工項代碼 (例如 "E08") 目前已使用過的最大流水號，並回傳下一號。
// 表單編號格式固定為「{工項代碼}-{4碼流水號}」(例如 E08-0037)；每個工項代碼各自獨立編號，
// 互不影響。若該分頁尚不存在，或找不到任何符合前綴的既有編號，則回傳該工項的預設起始號 {代碼}-0001。
function getNextFormNumber(categoryCode) {
  const fallback = categoryCode + '-0001';
  if (!categoryCode) return fallback;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECORD_NAME);
  if (!sheet || sheet.getLastRow() < 2) return fallback;

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const noColIdx = headerRow.indexOf('表單編號');
  if (noColIdx === -1) return fallback;

  const values = sheet.getRange(2, noColIdx + 1, sheet.getLastRow() - 1, 1).getValues();
  const prefix = categoryCode + '-';
  let maxSeq = 0;
  values.forEach(function (row) {
    const val = String(row[0] || '');
    if (val.indexOf(prefix) === 0) {
      const num = parseInt(val.substring(prefix.length), 10);
      if (!isNaN(num) && num > maxSeq) maxSeq = num;
    }
  });

  return prefix + String(maxSeq + 1).padStart(4, '0');
}

// 取得或建立指定名稱的分頁；若是新建立的分頁，會先寫入標題列並凍結
function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f1f5f9');
  }
  return sheet;
}

// 寫入「紀錄表資料」分頁：一列＝一個管理項目
function writeRecordRows(submissionId, header, items) {
  const headers = [
    '提交時間', '提交ID', '公司名稱', '表單編號', '分項工程(工項)', '工程名稱',
    '施工廠商', '抽查位置', '抽查人員', '抽查日期', '佐證文件',
    '施工階段', '施工流程', '管理項目', '抽查標準', '現場實測值', '抽查結果', '備註'
  ];
  const sheet = getOrCreateSheet(SHEET_RECORD_NAME, headers);
  const now = new Date();

  const rows = items.map(function (item) {
    return [
      now, submissionId,
      header.company || '', header.no || '', header.subProjectLabel || '', header.projName || '',
      header.contractor || '', header.location || '', header.inspector || '', header.dateText || '', header.proofDoc || '',
      item.stage || '', item.process || '', item.item || '', item.standard || '',
      item.actual || '', item.result || '', item.note || ''
    ];
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  }
}

// 寫入「照片」分頁：一列＝一張照片，並上傳到 Google 雲端硬碟指定資料夾
function writePhotoRows(submissionId, header, photos) {
  const headers = ['提交時間', '提交ID', '工程名稱', '分項工程', '照片說明', 'Drive 連結', '縮圖預覽'];
  const sheet = getOrCreateSheet(SHEET_PHOTO_NAME, headers);
  const now = new Date();
  const folder = getOrCreateDriveFolder(DRIVE_FOLDER_NAME);

  photos.forEach(function (photo) {
    let driveUrl = '';
    let imageFormula = '';
    try {
      const file = uploadPhotoToDrive(folder, photo, submissionId);
      driveUrl = file.getUrl();
      const directViewUrl = 'https://drive.google.com/uc?export=view&id=' + file.getId();
      imageFormula = '=IMAGE("' + directViewUrl + '")';
    } catch (uploadErr) {
      driveUrl = '上傳失敗：' + String(uploadErr);
    }
    sheet.appendRow([
      now, submissionId, header.projName || '', header.subProjectLabel || '',
      photo.desc || '', driveUrl, imageFormula
    ]);
  });
}

// 取得或建立指定名稱的 Google 雲端硬碟資料夾
function getOrCreateDriveFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

// 把一張照片 (base64 dataURL) 解碼後存成 Drive 檔案，並設定為「知道連結的人都能檢視」
function uploadPhotoToDrive(folder, photo, submissionId) {
  const dataUrl = photo.dataUrl || '';
  const match = dataUrl.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/);
  if (!match) throw new Error('不支援的照片格式 (非 base64 圖片 dataURL)');

  const mimeType = match[1];
  const base64Data = match[2];
  const bytes = Utilities.base64Decode(base64Data);
  const ext = (mimeType.split('/')[1] || 'jpg').replace('jpeg', 'jpg');
  const safeDesc = String(photo.desc || 'photo').replace(/[\\/:*?"<>|]/g, '_').slice(0, 30);
  const fileName = submissionId + '_' + safeDesc + '.' + ext;

  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file;
}
