/**
 * ------------------------------------------------------------------
 * PEA Mechanical Job Management System - Backend (Stable)
 * ------------------------------------------------------------------
 */

const SPREADSHEET_ID = '128p2wrkeCBJR4--kuTpxdBY8zP1qlJG9bCMrjFA52II'; 
const IMAGE_FOLDER_ID = '1Tkzv2cmHYayyRPRKqWIahdD3xdjA7EY5'; 

const DB_CONFIG = {
  SHEET_NAMES: {
    USERS: 'Account',
    JOBS: 'งาน_กบค.',
    OPERATIONS: 'การออกปฏิบัติงาน', 
    PARTS: 'เบิกอะไหล่'
  },
  COLUMNS: {
    USERS: {
      ID: 'รหัสพนักงาน',
      PASS: 'Password',
      NAME: 'ชื่อ-นามสกุล',
      ROLE: 'Level', 
      DEPT: 'แผนก',
      POS: 'ตำแหน่ง',
      IMG: 'รูปภาพ'
    },
    JOBS: {
      ID: 'เลขงาน',
      DESC: 'รายละเอียด',
      JOB_TYPE: 'ประเภทงาน', // Added Job Type
      PENDING_TYPE: 'ประเภทงานค้าง',
      PENDING_DESC: 'รายละเอียดงานค้าง',
      DATE_RECV: 'วันที่รับงาน',
      DURATION: 'ระยะเวลางานค้าง (เดือน)',
      SENDER: 'ผู้ส่งงาน',
      RESPONSIBLE: 'ผู้รับผิดชอบงาน',
      STATUS: 'สถานะ',
      APPROVER: 'ผู้อนุมัติ',
      CLOSE_DATE: 'วันที่ปิดงาน',
      DEPT: 'แผนก',
      ATTACH: 'เอกสารแนบ'
    },
    OPERATIONS: {
      JOB_ID: 'เลขงาน',
      DATE: 'วันที่ปฏิบัติงาน',
      WORKER: 'ผู้ปฏิบัติงาน',
      LOCATION: 'สถานที่',
      DETAIL: 'รายละเอียดการซ่อม'
    },
    PARTS: {
      JOB_ID: 'เลขงาน',
      ITEM: 'รายการอะไหล่',
      QTY: 'จำนวน',
      PRICE: 'ราคา',
      DATE: 'วันที่เบิก'
    }
  }
};

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ระบบบริหารจัดการงานเครื่องกล (PEA)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function _getDb() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function _getDataAsObjects(sheetName) {
  try {
    const ss = _getDb();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];
    
    const headers = data[0].map(h => h.trim());
    const rows = data.slice(1);
    
    return rows.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index]; 
      });
      return obj;
    });
  } catch (e) {
    Logger.log("Error reading " + sheetName + ": " + e.toString());
    return [];
  }
}

// --- IMAGE HELPERS ---
function _getImageUrlMap() {
  const imageMap = {};
  try {
    const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      imageMap[file.getName()] = file.getThumbnailLink().replace('=s220', '=s400');
    }
  } catch (e) {}
  return imageMap;
}

function _extractFilename(path) {
  if (!path || path === '-' || path === '') return null;
  return path.split('/').pop().trim();
}

function _findImageFile(path) {
  if (!path || path === '-' || path === '') return null;
  try {
    let rawFilename = path.split('/').pop().trim();
    const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
    let files = folder.getFilesByName(rawFilename);
    if (files.hasNext()) return files.next();
  } catch (e) {}
  return null;
}

function _getImageBase64(path) {
  const file = _findImageFile(path);
  if (file) {
    try {
      const blob = file.getBlob();
      const base64 = Utilities.base64Encode(blob.getBytes());
      return 'data:' + blob.getContentType() + ';base64,' + base64;
    } catch (e) {}
  }
  return null;
}

function saveUserImage(base64Data, fileName) {
  try {
    const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
    const contentType = base64Data.substring(5, base64Data.indexOf(';'));
    const bytes = Utilities.base64Decode(base64Data.substring(base64Data.indexOf('base64,') + 7));
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "Account_Images/" + file.getName();
  } catch (e) {
    throw new Error("Upload Failed: " + e.toString());
  }
}

function _createUserMap() {
  const rawUsers = _getDataAsObjects(DB_CONFIG.SHEET_NAMES.USERS);
  const COLS = DB_CONFIG.COLUMNS.USERS;
  const userMap = {};
  rawUsers.forEach(u => {
    const uid = String(u[COLS.ID] || '').trim();
    if(uid) {
      userMap[uid] = { name: u[COLS.NAME], imgPath: u[COLS.IMG] };
    }
  });
  return userMap;
}

// --- APIs ---

function loginUser(username, password) {
  try {
    const users = _getDataAsObjects(DB_CONFIG.SHEET_NAMES.USERS);
    const COLS = DB_CONFIG.COLUMNS.USERS;
    const foundUser = users.find(u => String(u[COLS.ID] || '').trim() === String(username).trim() && String(u[COLS.PASS] || '').trim() === String(password).trim());
    
    if (foundUser) {
      let systemRole = 'Staff';
      const level = String(foundUser[COLS.ROLE] || '').toLowerCase();
      const dept = String(foundUser[COLS.DEPT] || '');

      if (level.includes('admin')) systemRole = 'Admin';
      else if (level.includes('manager')) systemRole = 'Approver';
      else if (dept === 'ผจส.') systemRole = 'Dispatcher';

      const avatarData = _getImageBase64(foundUser[COLS.IMG]);

      return { 
        status: 'success', 
        user: {
          id: foundUser[COLS.ID],
          name: foundUser[COLS.NAME],
          role: systemRole, 
          dept: dept,
          position: foundUser[COLS.POS],
          avatar: avatarData
        }
      };
    }
    return { status: 'error', message: 'รหัสพนักงานหรือรหัสผ่านไม่ถูกต้อง' };
  } catch (e) {
    return { status: 'error', message: 'Server Error: ' + e.toString() };
  }
}

function getAllUsers() {
  try {
    const rawUsers = _getDataAsObjects(DB_CONFIG.SHEET_NAMES.USERS);
    const COLS = DB_CONFIG.COLUMNS.USERS;
    
    const users = rawUsers.map(u => ({
      id: String(u[COLS.ID] || ''),
      password: String(u[COLS.PASS] || ''),
      name: String(u[COLS.NAME] || ''),
      role: String(u[COLS.ROLE] || 'User'),
      dept: String(u[COLS.DEPT] || ''),
      position: String(u[COLS.POS] || ''),
      rawImgPath: u[COLS.IMG] || '' 
    }));
    return { status: 'success', data: users };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function getOneUserImage(path) {
  try {
    const base64 = _getImageBase64(path);
    return { status: 'success', image: base64 };
  } catch (e) {
    return { status: 'error' };
  }
}

function updateUser(userId, updatedData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = _getDb();
    const sheet = ss.getSheetByName(DB_CONFIG.SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0].map(h => h.trim());
    const COLS = DB_CONFIG.COLUMNS.USERS;
    
    const idIndex = headers.indexOf(COLS.ID);
    if (idIndex === -1) return { status: 'error', message: 'ID Column not found' };
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIndex]) === String(userId)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { status: 'error', message: 'User not found' };
    
    if (updatedData.newImageBase64 && updatedData.newImageName) {
       const newPath = saveUserImage(updatedData.newImageBase64, updatedData.newImageName);
       updatedData.img = newPath; 
    }

    const mapFields = {
      name: COLS.NAME, role: COLS.ROLE, dept: COLS.DEPT,
      position: COLS.POS, password: COLS.PASS, img: COLS.IMG
    };
    
    for (const [key, colName] of Object.entries(mapFields)) {
      if (updatedData[key] !== undefined && updatedData[key] !== null) {
        const colIndex = headers.indexOf(colName);
        if (colIndex > -1) sheet.getRange(rowIndex, colIndex + 1).setValue(updatedData[key]);
      }
    }
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getAllJobs() {
  try {
    const rawJobs = _getDataAsObjects(DB_CONFIG.SHEET_NAMES.JOBS);
    const rawUsers = _getDataAsObjects(DB_CONFIG.SHEET_NAMES.USERS);
    const COLS = DB_CONFIG.COLUMNS.JOBS;
    const USER_COLS = DB_CONFIG.COLUMNS.USERS;

    const imgMap = _getImageUrlMap();
    const userMap = {};
    rawUsers.forEach(u => {
      const uid = String(u[USER_COLS.ID] || '').trim();
      const filename = _extractFilename(u[USER_COLS.IMG]);
      const avatarUrl = filename && imgMap[filename] ? imgMap[filename] : null;
      if(uid) userMap[uid] = { name: u[USER_COLS.NAME], avatar: avatarUrl };
    });

    const standardJobs = rawJobs.map(job => {
      const respId = String(job[COLS.RESPONSIBLE] || '').trim();
      let respObj = { id: respId, name: respId || '-', avatar: null };
      if(userMap[respId]) {
        respObj.name = userMap[respId].name;
        respObj.avatar = userMap[respId].avatar;
      }

      const appId = String(job[COLS.APPROVER] || '').trim();
      let appObj = { id: appId, name: appId || '-', avatar: null };
      if(userMap[appId]) {
        appObj.name = userMap[appId].name;
        appObj.avatar = userMap[appId].avatar;
      }

      return {
        id: String(job[COLS.ID] || '-'),                 
        desc: String(job[COLS.DESC] || ''),
        jobType: String(job[COLS.JOB_TYPE] || '-'),
        pendingType: String(job[COLS.PENDING_TYPE] || '-'),
        pendingDesc: String(job[COLS.PENDING_DESC] || '-'),
        date: String(job[COLS.DATE_RECV] || ''),
        duration: String(job[COLS.DURATION] || '-'),
        sender: String(job[COLS.SENDER] || '-'),
        responsible: respObj,
        approver: appObj,
        status: String(job[COLS.STATUS] || 'รอดำเนินการ'),
        closeDate: String(job[COLS.CLOSE_DATE] || ''),
        dept: String(job[COLS.DEPT] || '-'),
      };
    });

    const validJobs = standardJobs.filter(j => j.id !== '-' && j.id !== '');
    return { status: 'success', data: validJobs.reverse() };
  } catch (e) {
    return { status: 'error', message: 'Backend Error: ' + e.message };
  }
}

function getJobDetailWithRelations(jobId) {
  try {
    const rawOps = _getDataAsObjects(DB_CONFIG.SHEET_NAMES.OPERATIONS);
    const rawParts = _getDataAsObjects(DB_CONFIG.SHEET_NAMES.PARTS);
    const rawUsers = _getDataAsObjects(DB_CONFIG.SHEET_NAMES.USERS);
    const OP_COLS = DB_CONFIG.COLUMNS.OPERATIONS;
    const PART_COLS = DB_CONFIG.COLUMNS.PARTS;
    const USER_COLS = DB_CONFIG.COLUMNS.USERS;
    
    const imgMap = _getImageUrlMap();
    const userMap = {};
    rawUsers.forEach(u => {
      const uid = String(u[USER_COLS.ID] || '').trim();
      const filename = _extractFilename(u[USER_COLS.IMG]);
      const avatarUrl = filename && imgMap[filename] ? imgMap[filename] : null;
      if(uid) userMap[uid] = { name: u[USER_COLS.NAME], avatar: avatarUrl };
    });

    const ops = rawOps.filter(op => String(op[OP_COLS.JOB_ID]) === String(jobId))
      .map(op => {
        const workerRaw = String(op[OP_COLS.WORKER] || '');
        const workerIds = workerRaw.split(',').map(s => s.trim());
        const workers = workerIds.map(wid => {
           if(userMap[wid]) return { id: wid, name: userMap[wid].name, imgPath: userMap[wid].avatar };
           return { id: wid, name: wid, imgPath: '' };
        });
        return { date: op[OP_COLS.DATE], workers: workers, location: op[OP_COLS.LOCATION], detail: op[OP_COLS.DETAIL] };
      });

    const parts = rawParts.filter(p => String(p[PART_COLS.JOB_ID]) === String(jobId))
      .map(p => ({ item: p[PART_COLS.ITEM], qty: p[PART_COLS.QTY], price: p[PART_COLS.PRICE], date: p[PART_COLS.DATE] }));
    
    return { status: 'success', data: { operations: ops, parts: parts } };
  } catch (e) {
    return { status: 'success', data: { operations: [], parts: [] } };
  }
}

function createNewJob(formData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    const ss = _getDb();
    const sheet = ss.getSheetByName(DB_CONFIG.SHEET_NAMES.JOBS);
    const COLS = DB_CONFIG.COLUMNS.JOBS;
    const lastRow = sheet.getLastRow();
    const year = (new Date().getFullYear() + 543).toString().substr(2);
    const newId = `${year}/${lastRow.toString().padStart(4, '0')}`;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const newRow = headers.map(h => {
      if (h === COLS.ID) return newId;
      if (h === COLS.STATUS) return 'ระหว่างดำเนินการ';
      if (h === COLS.DATE_RECV) return new Date();
      if (h === COLS.DEPT) return formData.dept;
      if (h === COLS.SENDER) return formData.sender;
      if (h === COLS.DESC) return formData.desc;
      return '';
    });
    sheet.appendRow(newRow);
    return { status: 'success', message: 'สร้างงานสำเร็จ', newJobId: newId };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function updateJobStatus(jobId, newStatus) {
  const ss = _getDb();
  const sheet = ss.getSheetByName(DB_CONFIG.SHEET_NAMES.JOBS);
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0].map(h => h.trim());
  const COLS = DB_CONFIG.COLUMNS.JOBS;
  const idColIndex = headers.indexOf(COLS.ID);
  const statusColIndex = headers.indexOf(COLS.STATUS);
  const closeDateIndex = headers.indexOf(COLS.CLOSE_DATE);
  
  if (idColIndex === -1) return { status: 'error', message: 'Not found ID Column' };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idColIndex]) === String(jobId)) {
      sheet.getRange(i + 1, statusColIndex + 1).setValue(newStatus);
      if (newStatus === 'ปิดงาน' && closeDateIndex > -1) {
        sheet.getRange(i + 1, closeDateIndex + 1).setValue(new Date());
      }
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'Job not found' };
}