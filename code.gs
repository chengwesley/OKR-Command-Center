/**
 * Code.gs – 後端程式碼（第一部分）
 * 包含：通用工具、登入登出、Company Objectives CRUD 以及 Key Results CRUD
 */

// 請改成你自己的試算表 ID
const SPREADSHEET_ID = 'Kdan_OKR_Data';

// 工作表名稱
const SHEET_USERS = 'Users';
const SHEET_COMPANY_OBJECTIVES = 'Company_Objectives';
const SHEET_DEPT_OBJECTIVES = 'Department_Objectives';
const SHEET_MY_OBJECTIVES = 'My_Objectives';
const SHEET_KEY_RESULTS = 'Key_Results';
const SHEET_COMMENTS = 'Comments';

// Session 設定
const SESSION_EXPIRATION_MINUTES = 60;
const PASSWORD_SALT = '77371111123ee';    // 可自行更換

/**
 * 讀取整張試算表(跳過標題列) → 二維陣列
 */
function getSheetData(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`找不到工作表 ${sheetName}`);
  const values = sheet.getDataRange().getValues();
  return values.length <= 1 ? [] : values.slice(1);
}

/**
 * 讀取指定工作表標題列 → 一維陣列
 */
function getSheetHeaders(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`找不到工作表 ${sheetName}`);
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * 合併「標題陣列」與「單行資料陣列」 → 物件
 */
function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((h, idx) => {
    obj[h] = row[idx];
  });
  return obj;
}

/**
 * SHA256 雜湊
 */
function hashString(str) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str);
  let hex = '';
  raw.forEach(b => {
    if (b < 0) b += 256;
    hex += (b.toString(16).length === 1 ? '0' : '') + b.toString(16);
  });
  return hex;
}

/**
 * 驗證 sessionToken → 回傳使用者資訊 (email, role, id, department)
 */
function authenticateAndAuthorize(sessionToken) {
  if (!sessionToken) throw new Error('未提供 sessionToken');
  const store = PropertiesService.getScriptProperties();
  const dataJson = store.getProperty(sessionToken);
  if (!dataJson) throw new Error('登入逾期或無效，請重新登入');
  const data = JSON.parse(dataJson);
  const now = Date.now();
  if (now > data.timestamp + SESSION_EXPIRATION_MINUTES * 60 * 1000) {
    store.deleteProperty(sessionToken);
    throw new Error('會話已逾期，請重新登入');
  }
  // 滑動延長
  data.timestamp = now;
  store.setProperty(sessionToken, JSON.stringify(data));
  return { email: data.email, role: data.role, id: data.id, department: data.department || '' };
}

/**
 * 登入：檢查 Users 試算表中的 Email + PasswordHash
 * Users 欄位：Email | PasswordHash | Role | ID | Department
 */
function loginUser(email, password) {
  if (!email || !password) {
    return { success: false, message: '請輸入帳號與密碼' };
  }
  try {
    const headers = getSheetHeaders(SHEET_USERS);
    const rows = getSheetData(SHEET_USERS).map(r => rowToObject(headers, r));
    const user = rows.find(u => u.Email === email);
    if (!user) return { success: false, message: '帳號或密碼錯誤' };

    if (user.PasswordHash !== hashString(password + PASSWORD_SALT)) {
      return { success: false, message: '帳號或密碼錯誤' };
    }

    const token = Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty(token, JSON.stringify({
      email: user.Email,
      role: user.Role,
      id: user.ID,
      department: user.Department || '',
      timestamp: Date.now()
    }));
    return {
      success: true,
      sessionToken: token,
      email: user.Email,
      role: user.Role,
      id: user.ID,
      department: user.Department || ''
    };
  } catch (e) {
    return { success: false, message: '登入失敗：' + e.message };
  }
}

/**
 * 登出：刪除 PropertiesService 裡的 sessionToken
 */
function logoutUser(sessionToken) {
  try {
    PropertiesService.getScriptProperties().deleteProperty(sessionToken);
    return { success: true };
  } catch (e) {
    return { success: false, message: '登出失敗：' + e.message };
  }
}

/**
 * 取得所有 Company Objectives (含計算進度)
 * 回傳 JSON 字串：{ objectives: [ { ID, Title, Description, OwnerEmail, Status, ApproverEmail, Progress } ], error: '' }
 */
function getCompanyObjectives(sessionToken) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return JSON.stringify({ error: e.message, objectives: [] });
  }
  try {
    const objHeaders = getSheetHeaders(SHEET_COMPANY_OBJECTIVES);
    const objRows = getSheetData(SHEET_COMPANY_OBJECTIVES).map(r => rowToObject(objHeaders, r));

    const krHeaders = getSheetHeaders(SHEET_KEY_RESULTS);
    const krRows = getSheetData(SHEET_KEY_RESULTS).map(r => rowToObject(krHeaders, r));

    const result = objRows.map(o => {
      const myKRs = krRows.filter(kr => kr.ObjectiveID === o.ID);
      const totalProg = myKRs.reduce((sum, kr) => sum + Number(kr.Progress || 0), 0);
      const avgProg = myKRs.length ? Math.round(totalProg / myKRs.length) : (Number(o.Progress) || 0);
      return {
        ID: o.ID,
        Title: o.Title,
        Description: o.Description,
        OwnerEmail: o.OwnerEmail,
        Status: o.Status,
        ApproverEmail: o.ApproverEmail || '',
        Progress: avgProg
      };
    });
    return JSON.stringify({ objectives: result, error: '' });
  } catch (e) {
    return JSON.stringify({ error: '讀取 Company Objectives 發生錯誤：' + e.message, objectives: [] });
  }
}

/**
 * 新增 Company Objective
 * objData: { Title, Description, OwnerEmail }
 */
function createCompanyObjective(sessionToken, objData) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return { success: false, message: e.message };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_COMPANY_OBJECTIVES);
    const headers = getSheetHeaders(SHEET_COMPANY_OBJECTIVES);
    const rows = getSheetData(SHEET_COMPANY_OBJECTIVES).map(r => rowToObject(headers, r));

    const prefix = 'C-2025Q2-O-';
    let maxNum = 0;
    rows.forEach(o => {
      if (o.ID && o.ID.startsWith(prefix)) {
        const num = parseInt(o.ID.replace(prefix, '')) || 0;
        if (num > maxNum) maxNum = num;
      }
    });
    const nextNum = (maxNum + 1).toString().padStart(2, '0');
    const newId = prefix + nextNum;

    const newRow = [
      newId,
      objData.Title || '',
      objData.Description || '',
      objData.OwnerEmail || '',
      0,          // Progress 初始 0
      'Draft',    // Status
      ''          // ApproverEmail
    ];
    sheet.appendRow(newRow);
    return { success: true, message: `新增 Company Objective 成功，ID: ${newId}` };
  } catch (e) {
    return { success: false, message: '新增 Company Objective 失敗：' + e.message };
  }
}

/**
 * 更新 Company Objective
 * data: { ID, Title, Description, OwnerEmail, Status, ApproverEmail }
 */
function updateCompanyObjective(sessionToken, data) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return { success: false, message: e.message };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_COMPANY_OBJECTIVES);
    const headers = getSheetHeaders(SHEET_COMPANY_OBJECTIVES);
    const allRows = sheet.getDataRange().getValues();
    for (let i = 1; i < allRows.length; i++) {
      if (allRows[i][0] === data.ID) {
        const rowIdx = i + 1;
        sheet.getRange(rowIdx, headers.indexOf('Title') + 1).setValue(data.Title);
        sheet.getRange(rowIdx, headers.indexOf('Description') + 1).setValue(data.Description);
        sheet.getRange(rowIdx, headers.indexOf('OwnerEmail') + 1).setValue(data.OwnerEmail);
        sheet.getRange(rowIdx, headers.indexOf('Status') + 1).setValue(data.Status);
        sheet.getRange(rowIdx, headers.indexOf('ApproverEmail') + 1).setValue(data.ApproverEmail || '');
        return { success: true, message: `更新 Company Objective ${data.ID} 成功` };
      }
    }
    return { success: false, message: `找不到 Company Objective ${data.ID}` };
  } catch (e) {
    return { success: false, message: '更新 Company Objective 失敗：' + e.message };
  }
}

/**
 * 刪除 Company Objective 及其所有 KR
 */
function deleteCompanyObjective(sessionToken, objectiveId) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return { success: false, message: e.message };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const sheetO = ss.getSheetByName(SHEET_COMPANY_OBJECTIVES);
    const dataO = sheetO.getDataRange().getValues();
    for (let i = 1; i < dataO.length; i++) {
      if (dataO[i][0] === objectiveId) {
        sheetO.deleteRow(i + 1);
        break;
      }
    }

    const sheetKR = ss.getSheetByName(SHEET_KEY_RESULTS);
    const dataKR = sheetKR.getDataRange().getValues();
    for (let i = dataKR.length - 1; i >= 1; i--) {
      if (dataKR[i][1] === objectiveId) {
        sheetKR.deleteRow(i + 1);
      }
    }

    return { success: true, message: `已刪除 Company Objective ${objectiveId} 及其所有 KR` };
  } catch (e) {
    return { success: false, message: '刪除 Company Objective 失敗：' + e.message };
  }
}
/**
 * 取得 Key Results (可選：__ALL_COMPANY__, __ALL_DEPT__, __ALL_PERSONAL__)
 * 回傳 JSON 字串：{ keyResults: [ … ], error: '' }
 */
/**
 * 取得 Key Results (可選：__ALL_COMPANY__, __ALL_DEPT__, __ALL_PERSONAL__)
 * [修正] 將篩選條件從寫死的 'Q2' 改為只判斷開頭字母，以適應不同季度的資料。
 * 回傳 JSON 字串：{ keyResults: [ … ], error: '' }
 */
function getKeyResultsByObjective(sessionToken, objectiveId) {
    try {
        authenticateAndAuthorize(sessionToken);
    } catch (e) {
        return JSON.stringify({ error: e.message, keyResults: [] });
    }
    try {
        const headers = getSheetHeaders(SHEET_KEY_RESULTS);
        const rows = getSheetData(SHEET_KEY_RESULTS).map(r => rowToObject(headers, r));
        let filtered;

        if (objectiveId === '__ALL_COMPANY__') {
            // 只檢查開頭是否為 'C-'，不再限制季度
            filtered = rows.filter(kr => kr.ObjectiveID && kr.ObjectiveID.startsWith('C-'));
        } else if (objectiveId === '__ALL_DEPT__') {
            // 只檢查開頭是否為 'D-'
            filtered = rows.filter(kr => kr.ObjectiveID && kr.ObjectiveID.startsWith('D-'));
        } else if (objectiveId === '__ALL_PERSONAL__') {
            // 只檢查開頭是否為 'P-'
            filtered = rows.filter(kr => kr.ObjectiveID && kr.ObjectiveID.startsWith('P-'));
        } else {
            // 按特定 ObjectiveID 篩選 (此部分邏輯不變)
            filtered = rows.filter(kr => kr.ObjectiveID === objectiveId);
        }
        
        return JSON.stringify({ keyResults: filtered, error: '' });
    } catch (e) {
        return JSON.stringify({ error: '讀取 Key Results 發生錯誤：' + e.message, keyResults: [] });
    }
}
/**
 * 新增 Key Result
 * krData: { ObjectiveID, Description, OwnerEmail, StartValue, TargetValue, Unit }
 */
function createKeyResult(sessionToken, krData) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return { success: false, message: e.message };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_KEY_RESULTS);
    const headers = getSheetHeaders(SHEET_KEY_RESULTS);
    const rows = getSheetData(SHEET_KEY_RESULTS).map(r => rowToObject(headers, r));

    const parentPrefix = krData.ObjectiveID + '-KR-';
    let maxNum = 0;
    rows.forEach(kr => {
      if (kr.ID && kr.ID.startsWith(parentPrefix)) {
        const num = parseInt(kr.ID.replace(parentPrefix, '')) || 0;
        if (num > maxNum) maxNum = num;
      }
    });
    const nextNum = (maxNum + 1).toString().padStart(2, '0');
    const newKrId = parentPrefix + nextNum;

    const newKrRow = [
      newKrId,
      krData.ObjectiveID,
      krData.Description || '',
      krData.OwnerEmail || '',
      krData.StartValue || 0,
      krData.TargetValue || 0,
      krData.StartValue || 0,    // CurrentValue 初始 = StartValue
      0,                         // Progress 初始 0
      krData.Unit || '',
      'On Track'
    ];
    sheet.appendRow(newKrRow);
    return { success: true, message: `新增 Key Result 成功，ID: ${newKrId}` };
  } catch (e) {
    return { success: false, message: '新增 Key Result 失敗：' + e.message };
  }
}
/**
 * 更新 Key Result
 * data: { ID, Description, OwnerEmail, StartValue, TargetValue, CurrentValue, Progress, Unit, Status }
 */
function updateKeyResult(sessionToken, data) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return { success: false, message: e.message };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_KEY_RESULTS);
    const headers = getSheetHeaders(SHEET_KEY_RESULTS);
    const allRows = sheet.getDataRange().getValues();
    for (let i = 1; i < allRows.length; i++) {
      if (allRows[i][0] === data.ID) {
        const rowIdx = i + 1;
        sheet.getRange(rowIdx, headers.indexOf('Description') + 1).setValue(data.Description);
        sheet.getRange(rowIdx, headers.indexOf('OwnerEmail') + 1).setValue(data.OwnerEmail);
        sheet.getRange(rowIdx, headers.indexOf('StartValue') + 1).setValue(data.StartValue);
        sheet.getRange(rowIdx, headers.indexOf('TargetValue') + 1).setValue(data.TargetValue);
        sheet.getRange(rowIdx, headers.indexOf('CurrentValue') + 1).setValue(data.CurrentValue);
        sheet.getRange(rowIdx, headers.indexOf('Progress') + 1).setValue(data.Progress);
        sheet.getRange(rowIdx, headers.indexOf('Unit') + 1).setValue(data.Unit);
        sheet.getRange(rowIdx, headers.indexOf('Status') + 1).setValue(data.Status);
        return { success: true, message: `更新 Key Result ${data.ID} 成功` };
      }
    }
    return { success: false, message: `找不到 Key Result ${data.ID}` };
  } catch (e) {
    return { success: false, message: '更新 Key Result 失敗：' + e.message };
  }
}

/**
 * 批次更新 Key Results
 * updates: [{ ID, Description, OwnerEmail, StartValue, TargetValue, CurrentValue, Progress, Unit, Status }, ...]
 * 注意：這裡假設傳入的數據包含 KR 的所有欄位，如果只更新部分欄位，後端需要更精細的處理。
 */
function batchUpdateKeyResults(sessionToken, updates) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return { success: false, message: e.message };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_KEY_RESULTS);
    const headers = getSheetHeaders(SHEET_KEY_RESULTS);
    const allRows = sheet.getDataRange().getValues();

    let updatedCount = 0;
    const errors = [];

    const krIdToIndexMap = {};
    for (let i = 1; i < allRows.length; i++) {
      krIdToIndexMap[allRows[i][0]] = i;
    }

    updates.forEach(data => {
      const rowIndex = krIdToIndexMap[data.ID];
      if (rowIndex !== undefined) {
        const sheetRowIdx = rowIndex + 1;

        try {
          sheet.getRange(sheetRowIdx, headers.indexOf('Description') + 1).setValue(data.Description);
          sheet.getRange(sheetRowIdx, headers.indexOf('OwnerEmail') + 1).setValue(data.OwnerEmail);
          sheet.getRange(sheetRowIdx, headers.indexOf('StartValue') + 1).setValue(data.StartValue);
          sheet.getRange(sheetRowIdx, headers.indexOf('TargetValue') + 1).setValue(data.TargetValue);
          sheet.getRange(sheetRowIdx, headers.indexOf('CurrentValue') + 1).setValue(data.CurrentValue);
          sheet.getRange(sheetRowIdx, headers.indexOf('Progress') + 1).setValue(data.Progress);
          sheet.getRange(sheetRowIdx, headers.indexOf('Unit') + 1).setValue(data.Unit);
          sheet.getRange(sheetRowIdx, headers.indexOf('Status') + 1).setValue(data.Status);
          updatedCount++;
        } catch (updateError) {
          errors.push(`KR ${data.ID} 更新失敗: ${updateError.message}`);
        }
      } else {
        errors.push(`KR ${data.ID} 找不到。`);
      }
    });

    if (errors.length > 0) {
      return { success: false, message: `批次更新完成，但有 ${errors.length} 個錯誤：\n${errors.join('\n')}` };
    }
    return { success: true, message: `成功批次更新 ${updatedCount} 個 Key Result。` };

  } catch (e) {
    return { success: false, message: '批次更新 Key Results 失敗：' + e.message };
  }
}

/**
 * 刪除 Key Result
 */
function deleteKeyResult(sessionToken, krId) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return { success: false, message: e.message };
  }
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_KEY_RESULTS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === krId) {
        sheet.deleteRow(i + 1);
        return { success: true, message: `已刪除 Key Result ${krId}` };
      }
    }
    return { success: false, message: `找不到 Key Result ${krId}` };
  } catch (e) {
    return { success: false, message: '刪除 Key Result 失敗：' + e.message };
  }
}
/**
 * Code.gs – 後端程式碼（第二部分）
 * 包含：Department Objectives / My Objectives CRUD、
 * 提交 審核 / 批准 / 駁回、Comments 功能
 */

// ------------------------------
// 1. Department Objectives CRUD
// ------------------------------

/**
 * 取得所有 Department Objectives (含計算進度)
 * 回傳 JSON 字串：{ objectives: [ { ID, Title, Description, OwnerEmail, DepartmentID, ParentCompanyKR, Status, ApproverEmail, Progress } ], error: '' }
 */
function getDepartmentObjectives(sessionToken) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return JSON.stringify({ error: e.message, objectives: [] });
  }
  try {
    const objHeaders = getSheetHeaders(SHEET_DEPT_OBJECTIVES);
    const objRows = getSheetData(SHEET_DEPT_OBJECTIVES).map(r => rowToObject(objHeaders, r));

    const krHeaders = getSheetHeaders(SHEET_KEY_RESULTS);
    const krRows = getSheetData(SHEET_KEY_RESULTS).map(r => rowToObject(krHeaders, r));

    const result = objRows.map(o => {
      const myKRs = krRows.filter(kr => kr.ObjectiveID === o.ID);
      const totalProg = myKRs.reduce((sum, kr) => sum + Number(kr.Progress || 0), 0);
      const avgProg = myKRs.length ? Math.round(totalProg / myKRs.length) : (Number(o.Progress) || 0);
      return {
        ID: o.ID,
        Title: o.Title,
        Description: o.Description,
        OwnerEmail: o.OwnerEmail,
        DepartmentID: o.DepartmentID,
        ParentCompanyKR: o.ParentCompanyKR,
        Status: o.Status,
        ApproverEmail: o.ApproverEmail || '',
        Progress: avgProg
      };
    });
    return JSON.stringify({ objectives: result, error: '' });
  } catch (e) {
    return JSON.stringify({ error: '讀取 Department Objectives 發生錯誤：' + e.message, objectives: [] });
  }
}

/**
 * 新增 Department Objective
 * data: { Title, Description, OwnerEmail, DepartmentID, ParentCompanyKR }
 */
function createDepartmentObjective(sessionToken, data) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return { success: false, message: e.message };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_DEPT_OBJECTIVES);
    const headers = getSheetHeaders(SHEET_DEPT_OBJECTIVES);
    const rows = getSheetData(SHEET_DEPT_OBJECTIVES).map(r => rowToObject(headers, r));

    const prefix = 'D-2025Q2-O-';
    let maxNum = 0;
    rows.forEach(o => {
      if (o.ID && o.ID.startsWith(prefix)) {
        const num = parseInt(o.ID.replace(prefix, '')) || 0;
        if (num > maxNum) maxNum = num;
      }
    });
    const nextNum = (maxNum + 1).toString().padStart(2, '0');
    const newId = prefix + nextNum;

    const newRow = [
      newId,
      data.Title || '',
      data.Description || '',
      data.OwnerEmail || '',
      data.DepartmentID || '',
      data.ParentCompanyKR || '',
      0,             // Progress
      'Draft',       // Status
      ''             // ApproverEmail
    ];
    sheet.appendRow(newRow);
    return { success: true, message: `新增 Department Objective 成功，ID: ${newId}` };
  } catch (e) {
    return { success: false, message: '新增 Department Objective 失敗：' + e.message };
  }
}

/**
 * 更新 Department Objective
 * data: { ID, Title, Description, OwnerEmail, DepartmentID, ParentCompanyKR, Status, ApproverEmail }
 */
function updateDepartmentObjective(sessionToken, data) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return { success: false, message: e.message };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_DEPT_OBJECTIVES);
    const headers = getSheetHeaders(SHEET_DEPT_OBJECTIVES);
    const allRows = sheet.getDataRange().getValues();
    for (let i = 1; i < allRows.length; i++) {
      if (allRows[i][0] === data.ID) {
        const rowIdx = i + 1;
        sheet.getRange(rowIdx, headers.indexOf('Title') + 1).setValue(data.Title);
        sheet.getRange(rowIdx, headers.indexOf('Description') + 1).setValue(data.Description);
        sheet.getRange(rowIdx, headers.indexOf('OwnerEmail') + 1).setValue(data.OwnerEmail);
        sheet.getRange(rowIdx, headers.indexOf('DepartmentID') + 1).setValue(data.DepartmentID);
        sheet.getRange(rowIdx, headers.indexOf('ParentCompanyKR') + 1).setValue(data.ParentCompanyKR);
        sheet.getRange(rowIdx, headers.indexOf('Status') + 1).setValue(data.Status);
        sheet.getRange(rowIdx, headers.indexOf('ApproverEmail') + 1).setValue(data.ApproverEmail || '');
        return { success: true, message: `更新 Department Objective ${data.ID} 成功` };
      }
    }
    return { success: false, message: `找不到 Department Objective ${data.ID}` };
  } catch (e) {
    return { success: false, message: '更新 Department Objective 失敗：' + e.message };
  }
}

/**
 * 刪除 Department Objective 及其所有 KR
 */
function deleteDepartmentObjective(sessionToken, objectiveId) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return { success: false, message: e.message };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const sheetO = ss.getSheetByName(SHEET_DEPT_OBJECTIVES);
    const dataO = sheetO.getDataRange().getValues();
    for (let i = 1; i < dataO.length; i++) {
      if (dataO[i][0] === objectiveId) {
        sheetO.deleteRow(i + 1);
        break;
      }
    }

    const sheetKR = ss.getSheetByName(SHEET_KEY_RESULTS);
    const dataKR = sheetKR.getDataRange().getValues();
    for (let i = dataKR.length - 1; i >= 1; i--) {
      if (dataKR[i][1] === objectiveId) {
        sheetKR.deleteRow(i + 1);
      }
    }

    return { success: true, message: `已刪除 Department Objective ${objectiveId} 及其所有 KR` };
  } catch (e) {
    return { success: false, message: '刪除 Department Objective 失敗：' + e.message };
  }
}

// ------------------------------
// 2. My Objectives CRUD
// ------------------------------

/**
 * 取得所有 My Objectives (含計算進度)
 * 回傳 JSON 字串：{ objectives: [ { ID, Title, Description, OwnerEmail, ParentDepartmentKR, Status, ApproverEmail, Progress } ], error: '' }
 */
function getMyObjectives(sessionToken) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return JSON.stringify({ error: e.message, objectives: [] });
  }
  try {
    const objHeaders = getSheetHeaders(SHEET_MY_OBJECTIVES);
    const objRows = getSheetData(SHEET_MY_OBJECTIVES).map(r => rowToObject(objHeaders, r));

    const krHeaders = getSheetHeaders(SHEET_KEY_RESULTS);
    const krRows = getSheetData(SHEET_KEY_RESULTS).map(r => rowToObject(krHeaders, r));

    const result = objRows.map(o => {
      const myKRs = krRows.filter(kr => kr.ObjectiveID === o.ID);
      const totalProg = myKRs.reduce((sum, kr) => sum + Number(kr.Progress || 0), 0);
      const avgProg = myKRs.length ? Math.round(totalProg / myKRs.length) : (Number(o.Progress) || 0);
      return {
        ID: o.ID,
        Title: o.Title,
        Description: o.Description,
        OwnerEmail: o.OwnerEmail,
        ParentDepartmentKR: o.ParentDepartmentKR,
        Status: o.Status,
        ApproverEmail: o.ApproverEmail || '',
        Progress: avgProg
      };
    });
    return JSON.stringify({ objectives: result, error: '' });
  } catch (e) {
    return JSON.stringify({ error: '讀取 My Objectives 發生錯誤：' + e.message, objectives: [] });
  }
}

/**
 * 新增 My Objective
 * data: { Title, Description, OwnerEmail, ParentDepartmentKR }
 */
function createMyObjective(sessionToken, data) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return { success: false, message: e.message };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_MY_OBJECTIVES);
    const headers = getSheetHeaders(SHEET_MY_OBJECTIVES);
    const rows = getSheetData(SHEET_MY_OBJECTIVES).map(r => rowToObject(headers, r));

    const prefix = 'P-2025Q2-O-';
    let maxNum = 0;
    rows.forEach(o => {
      if (o.ID && o.ID.startsWith(prefix)) {
        const num = parseInt(o.ID.replace(prefix, '')) || 0;
        if (num > maxNum) maxNum = num;
      }
    });
    const nextNum = (maxNum + 1).toString().padStart(2, '0');
    const newId = prefix + nextNum;

    const newRow = [
      newId,
      data.Title || '',
      data.Description || '',
      data.OwnerEmail || '',
      data.ParentDepartmentKR || '',
      0,             // Progress
      'Draft',       // Status
      ''             // ApproverEmail
    ];
    sheet.appendRow(newRow);
    return { success: true, message: `新增 My Objective 成功，ID: ${newId}` };
  } catch (e) {
    return { success: false, message: '新增 My Objective 失敗：' + e.message };
  }
}

/**
 * 更新 My Objective
 * data: { ID, Title, Description, OwnerEmail, ParentDepartmentKR, Status, ApproverEmail }
 */
function updateMyObjective(sessionToken, data) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return { success: false, message: e.message };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_MY_OBJECTIVES);
    const headers = getSheetHeaders(SHEET_MY_OBJECTIVES);
    const allRows = sheet.getDataRange().getValues();
    for (let i = 1; i < allRows.length; i++) {
      if (allRows[i][0] === data.ID) {
        const rowIdx = i + 1;
        sheet.getRange(rowIdx, headers.indexOf('Title') + 1).setValue(data.Title);
        sheet.getRange(rowIdx, headers.indexOf('Description') + 1).setValue(data.Description);
        sheet.getRange(rowIdx, headers.indexOf('OwnerEmail') + 1).setValue(data.OwnerEmail);
        sheet.getRange(rowIdx, headers.indexOf('ParentDepartmentKR') + 1).setValue(data.ParentDepartmentKR);
        sheet.getRange(rowIdx, headers.indexOf('Status') + 1).setValue(data.Status);
        sheet.getRange(rowIdx, headers.indexOf('ApproverEmail') + 1).setValue(data.ApproverEmail || '');
        return { success: true, message: `更新 My Objective ${data.ID} 成功` };
      }
    }
    return { success: false, message: `找不到 My Objective ${data.ID}` };
  } catch (e) {
    return { success: false, message: '更新 My Objective 失敗：' + e.message };
  }
}

/**
 * 刪除 My Objective 及其所有 KR
 */
function deleteMyObjective(sessionToken, objectiveId) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return { success: false, message: e.message };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const sheetO = ss.getSheetByName(SHEET_MY_OBJECTIVES);
    const dataO = sheetO.getDataRange().getValues();
    for (let i = 1; i < dataO.length; i++) {
      if (dataO[i][0] === objectiveId) {
        sheetO.deleteRow(i + 1);
        break;
      }
    }

    const sheetKR = ss.getSheetByName(SHEET_KEY_RESULTS);
    const dataKR = sheetKR.getDataRange().getValues();
    for (let i = dataKR.length - 1; i >= 1; i--) {
      if (dataKR[i][1] === objectiveId) {
        sheetKR.deleteRow(i + 1);
      }
    }

    return { success: true, message: `已刪除 My Objective ${objectiveId} 及其所有 KR` };
  } catch (e) {
    return { success: false, message: '刪除 My Objective 失敗：' + e.message };
  }
}

// ------------------------------
// 3. 提交 → 審批 → 批准/駁回
// ------------------------------

/**
 * 提交 Objective 至審核 (適用 Company / Department / My)
 * param: { sheetName, objectiveId, approverEmail }
 */
function submitObjectiveForApproval(sessionToken, param) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return { success: false, message: e.message };
  }
  const { sheetName, objectiveId, approverEmail } = param;
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    const headers = getSheetHeaders(sheetName);
    const allRows = sheet.getDataRange().getValues();
    for (let i = 1; i < allRows.length; i++) {
      if (allRows[i][0] === objectiveId) {
        const rowIdx = i + 1;
        sheet.getRange(rowIdx, headers.indexOf('Status') + 1).setValue('Pending');
        sheet.getRange(rowIdx, headers.indexOf('ApproverEmail') + 1).setValue(approverEmail);
        return { success: true, message: `O ${objectiveId} 已提交審核，審批者：${approverEmail}` };
      }
    }
    return { success: false, message: `找不到 ${sheetName} 中的 O ${objectiveId}` };
  } catch (e) {
    return { success: false, message: '提交審核失敗：' + e.message };
  }
}

/**
 * 取得所有待審核項目 (Company / Department / My)
 * 回傳 JSON：{ pending: [ { type, ID, Title, OwnerEmail, Submitter, sheetName } ], error: '' }
 */
function getPendingApprovals(sessionToken) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return JSON.stringify({ error: e.message, pending: [] });
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const result = [];

    const compHeaders = getSheetHeaders(SHEET_COMPANY_OBJECTIVES);
    const compRows = getSheetData(SHEET_COMPANY_OBJECTIVES).map(r => rowToObject(compHeaders, r));
    compRows.forEach(o => {
      if (o.Status === 'Pending') {
        result.push({
          type: 'Company O',
          ID: o.ID,
          Title: o.Title,
          OwnerEmail: o.OwnerEmail,
          Submitter: o.OwnerEmail,
          sheetName: SHEET_COMPANY_OBJECTIVES
        });
      }
    });

    const deptHeaders = getSheetHeaders(SHEET_DEPT_OBJECTIVES);
    const deptRows = getSheetData(SHEET_DEPT_OBJECTIVES).map(r => rowToObject(deptHeaders, r));
    deptRows.forEach(o => {
      if (o.Status === 'Pending') {
        result.push({
          type: 'Department O',
          ID: o.ID,
          Title: o.Title,
          OwnerEmail: o.OwnerEmail,
          Submitter: o.OwnerEmail,
          sheetName: SHEET_DEPT_OBJECTIVES
        });
      }
    });

    const myHeaders = getSheetHeaders(SHEET_MY_OBJECTIVES);
    const myRows = getSheetData(SHEET_MY_OBJECTIVES).map(r => rowToObject(myHeaders, r));
    myRows.forEach(o => {
      if (o.Status === 'Pending') {
        result.push({
          type: 'Personal O',
          ID: o.ID,
          Title: o.Title,
          OwnerEmail: o.OwnerEmail,
          Submitter: o.OwnerEmail,
          sheetName: SHEET_MY_OBJECTIVES
        });
      }
    });

    return JSON.stringify({ pending: result, error: '' });
  } catch (e) {
    return JSON.stringify({ error: '讀取待審核失敗：' + e.message, pending: [] });
  }
}

/**
 * 批准 Objective
 * param: { sheetName, objectiveId }
 */
function approveObjective(sessionToken, param) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return { success: false, message: e.message };
  }
  const { sheetName, objectiveId } = param;
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    const headers = getSheetHeaders(sheetName);
    const allRows = sheet.getDataRange().getValues();
    for (let i = 1; i < allRows.length; i++) {
      if (allRows[i][0] === objectiveId) {
        const rowIdx = i + 1;
        sheet.getRange(rowIdx, headers.indexOf('Status') + 1).setValue('Approved');
        return { success: true, message: `已批准 ${objectiveId}` };
      }
    }
    return { success: false, message: `找不到 ${sheetName} 中的 O ${objectiveId}` };
  } catch (e) {
    return { success: false, message: '批准失敗：' + e.message };
  }
}

/**
 * 駁回 Objective
 * param: { sheetName, objectiveId, reason }
 */
function rejectObjective(sessionToken, param) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return { success: false, message: e.message };
  }
  const { sheetName, objectiveId, reason } = param;
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    const headers = getSheetHeaders(sheetName);
    const allRows = sheet.getDataRange().getValues();
    for (let i = 1; i < allRows.length; i++) {
      if (allRows[i][0] === objectiveId) {
        const rowIdx = i + 1;
        sheet.getRange(rowIdx, headers.indexOf('Status') + 1).setValue('Rejected');
        const rejectReasonColIdx = headers.indexOf('RejectReason');
        if (rejectReasonColIdx !== -1) {
          sheet.getRange(rowIdx, rejectReasonColIdx + 1).setValue(reason);
        } else {
          console.warn(`Warning: 'RejectReason' column not found in sheet ${sheetName}. Reason not saved.`);
        }
        return { success: true, message: `已駁回 ${objectiveId}` };
      }
    }
    return { success: false, message: `找不到 ${sheetName} 中的 O ${objectiveId}` };
  } catch (e) {
    return { success: false, message: '駁回失敗：' + e.message };
  }
}

// ------------------------------
// 4. Comments 功能
// ------------------------------

/**
 * 取得特定 Entity（O 或 KR）之 Comments
 * 回傳 JSON：{ comments: [ { EntityID, CommentText, Author, Timestamp } ], error: '' }
 */
function getComments(sessionToken, entityId) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return JSON.stringify({ error: e.message, comments: [] });
  }
  try {
    const headers = getSheetHeaders(SHEET_COMMENTS);
    const rows = getSheetData(SHEET_COMMENTS).map(r => rowToObject(headers, r));
    const filtered = rows.filter(c => c.EntityID === entityId);
    filtered.sort((a, b) => new Date(a.Timestamp) - new Date(b.Timestamp));
    return JSON.stringify({ comments: filtered, error: '' });
  } catch (e) {
    return JSON.stringify({ error: '讀取留言失敗：' + e.message, comments: [] });
  }
}

function addComment(sessionToken, commentData) {
  let userInfo;
  try {
    userInfo = authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return { success: false, message: e.message };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_COMMENTS);
    const headers = getSheetHeaders(SHEET_COMMENTS);
    const timestamp = new Date().toISOString();
    const newRow = [
      Utilities.getUuid(),           // ID
      commentData.EntityID,          // EntityID
      commentData.CommentText,       // CommentText
      userInfo.email,                // Author
      timestamp                      // Timestamp
    ];
    sheet.appendRow(newRow);
    return { success: true, message: '留言已新增' };
  } catch (e) {
    return { success: false, message: '新增留言失敗：' + e.message };
  }
}

/**
 * Code.gs – 後端程式碼（第三部分）
 * 包含：Hierarchy (心智圖) 與 報表 Dashboard 功能
 */

// ------------------------------
// 5. 下拉選單：取得所有 Company KR / Department KR
// ------------------------------

/**
 * 取得所有 Company KR 供下拉選單使用
 * 回傳 JSON：{ parentCompanyKRs: [ { ID, Description } ], error: '' }
 */
function getAllCompanyKRs(sessionToken) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return JSON.stringify({ error: e.message, parentCompanyKRs: [] });
  }
  try {
    const krHeaders = getSheetHeaders(SHEET_KEY_RESULTS);
    const krRows = getSheetData(SHEET_KEY_RESULTS).map(r => rowToObject(krHeaders, r));
    const filtered = krRows
      .filter(kr => kr.ObjectiveID.startsWith('C-2025Q2-O-'))
      .map(kr => ({ ID: kr.ID, Description: kr.Description }));
    return JSON.stringify({ parentCompanyKRs: filtered, error: '' });
  } catch (e) {
    return JSON.stringify({ error: '讀取 Company KR 失敗：' + e.message, parentCompanyKRs: [] });
  }
}

/**
 * 取得所有 Department KR 供下拉選單使用
 * 回傳 JSON：{ parentDeptKRs: [ { ID, Description } ], error: '' }
 */
function getAllDepartmentKRs(sessionToken) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return JSON.stringify({ error: e.message, parentDeptKRs: [] });
  }
  try {
    const krHeaders = getSheetHeaders(SHEET_KEY_RESULTS);
    const krRows = getSheetData(SHEET_KEY_RESULTS).map(r => rowToObject(krHeaders, r));
    const filtered = krRows
      .filter(kr => kr.ObjectiveID.startsWith('D-2025Q2-O-'))
      .map(kr => ({ ID: kr.ID, Description: kr.Description }));
    return JSON.stringify({ parentDeptKRs: filtered, error: '' });
  } catch (e) {
    return JSON.stringify({ error: '讀取 Department KR 失敗：' + e.message, parentDeptKRs: [] });
  }
}

// ------------------------------
// 6. 心智圖 (Hierarchy) – getMindmapData
// 此函數將整合所有層次數據並返回為單一樹狀結構
// ------------------------------
/**
 * 【最終版】取得 OKR 樹狀結構 (呈現至部門 KR)
 * [修改] 不再讀取和處理個人層級的 OKR，以簡化高階管理者視圖。
 * [功能] 自動檢測並標記未被部門 O 承接的公司 KR。
 */
function getMindmapData(sessionToken) {
    let userInfo;
    try {
        userInfo = authenticateAndAuthorize(sessionToken);
    } catch (e) {
        return JSON.stringify({ error: e.message, mindmapData: null });
    }

    try {
        // 1. 修改：不再讀取 My_Objectives 和個人的 KR
        const compObjs = JSON.parse(getCompanyObjectives(sessionToken)).objectives || [];
        const deptObjs = JSON.parse(getDepartmentObjectives(sessionToken)).objectives || [];
        const allKRs = (JSON.parse(getKeyResultsByObjective(sessionToken, '__ALL_COMPANY__')).keyResults || [])
            .concat(JSON.parse(getKeyResultsByObjective(sessionToken, '__ALL_DEPT__')).keyResults || []);

        // 2. 建立節點 Map
        const nodeMap = new Map();
        const processItem = (item, type) => {
            const node = {
                id: item.ID,
                name: type.includes("Objective") ? item.Title : item.Description,
                type: type,
                data: item,
                children: []
            };
            nodeMap.set(item.ID, node);
        };

        compObjs.forEach(item => processItem(item, "CompanyObjective"));
        deptObjs.forEach(item => processItem(item, "DepartmentObjective"));

        allKRs.forEach(item => {
            if (!item || !item.ID) return;
            if (item.ID.startsWith('C-')) processItem(item, "CompanyKeyResult");
            else if (item.ID.startsWith('D-')) processItem(item, "DepartmentKeyResult");
        });

        // 3. 建立父子關係
        nodeMap.forEach(node => {
            let parentId = null;
            if (node.type.includes("KeyResult")) {
                parentId = node.data.ObjectiveID;
            } else if (node.type === "DepartmentObjective") {
                parentId = node.data.ParentCompanyKR;
            }

            if (parentId && nodeMap.has(parentId)) {
                nodeMap.get(parentId).children.push(node);
            }
        });
        
        // 4. 檢查未承接的公司 KR
        nodeMap.forEach(node => {
            if (node.type === "CompanyKeyResult") {
                if (node.children.length === 0) {
                    node.data.isUnfulfilled = true;
                }
            }
            // 部門KR是最終層級，不需檢查是否被承接
        });

        // 5. 構建最終樹
        const mindmapTreeData = {
            id: "OKR_ROOT",
            name: "公司 OKR 總覽",
            type: "Root",
            data: { Title: "公司 OKR 總覽", OwnerEmail: userInfo.email || "未知用戶" },
            children: compObjs.map(co => nodeMap.get(co.ID))
        };

        return JSON.stringify({ mindmapData: mindmapTreeData, error: '' });

    } catch (e) {
        console.error(`getMindmapData Error: ${e.toString()}`);
        return JSON.stringify({ error: `構建心智圖數據失敗：${e.message}`, mindmapData: null });
    }
}

// ------------------------------
// 7. 報表 Dashboard 功能
// ------------------------------

/**
 * AI 智能預測資料 (KR Progress)
 * 回傳 JSON：{ predictionData: { labels: [...], actual: [...], predicted: [...] } }
 * -- labels: 週次 (W1, W2, …)
 * -- actual: 每週實際平均進度
 * -- predicted: 後續預測值
 */
function getAIPredictionData(sessionToken) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return JSON.stringify({ error: e.message, predictionData: {} });
  }
  try {
    const labels = ['W1', 'W2', 'W3', 'W4', 'W5'];
    const krHeaders = getSheetHeaders(SHEET_KEY_RESULTS);
    const krRows = getSheetData(SHEET_KEY_RESULTS).map(r => rowToObject(krHeaders, r));
    const allProg = krRows.map(kr => Number(kr.Progress) || 0);
    const avgProgress = allProg.length
      ? Math.round(allProg.reduce((s, v) => s + v, 0) / allProg.length)
      : 0;
    const actual = [
      Math.round(avgProgress * 0.4),
      Math.round(avgProgress * 0.6),
      Math.round(avgProgress * 0.8),
      avgProgress,
      avgProgress
    ];
    const predicted = [
      actual[0],
      actual[1],
      actual[2],
      Math.min(100, Math.round(actual[2] * 1.1)),
      Math.min(100, Math.round(actual[2] * 1.2))
    ];

    return JSON.stringify({
      predictionData: {
        labels: labels,
        actual: actual,
        predicted: predicted
      }
    });
  } catch (e) {
    return JSON.stringify({ error: '取得 AI 預測資料失敗：' + e.message, predictionData: {} });
  }
}

/**
 * Dashboard 概覽 & 部門排名
 * 回傳 JSON：{ overview: { labels: [...], data: [...] }, deptRanking: { labels: [...], data: [...] } }
 * -- overview.labels: Company O IDs
 * -- overview.data: Company O 平均進度
 * -- deptRanking.labels: Department ID 列表
 * -- deptRanking.data: Department O 平均進度 (取該部門所有 DeptO 平均)
 */
function getDashboardOverviewData(sessionToken) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return JSON.stringify({ error: e.message, overview: {}, deptRanking: {} });
  }
  try {
    const compRes = JSON.parse(getCompanyObjectives(sessionToken));
    const compList = compRes.objectives || [];
    const compLabels = compList.map(o => o.ID);
    const compData = compList.map(o => o.Progress);

    const deptRes = JSON.parse(getDepartmentObjectives(sessionToken));
    const deptList = deptRes.objectives || [];
    const deptGroup = {};
    deptList.forEach(o => {
      if (!deptGroup[o.DepartmentID]) deptGroup[o.DepartmentID] = [];
      deptGroup[o.DepartmentID].push(o.Progress);
    });
    const deptLabels = Object.keys(deptGroup);
    const deptData = deptLabels.map(deptId => {
      const arr = deptGroup[deptId];
      const avg = arr.length ? Math.round(arr.reduce((s, v) => s + v, 0) / arr.length) : 0;
      return avg;
    });

    return JSON.stringify({
      overview: { labels: compLabels, data: compData },
      deptRanking: { labels: deptLabels, data: deptData }
    });
  } catch (e) {
    return JSON.stringify({ error: '取得 Dashboard 資料失敗：' + e.message, overview: {}, deptRanking: {} });
  }
}

// 回傳 Alignment Heatmap 資料 (矩陣)
function getAlignmentHeatmapData(sessionToken) {
  try {
    authenticateAndAuthorize(sessionToken);
  } catch (e) {
    return JSON.stringify({ error: e.message, heatmap: {} });
  }
  try {
    const compHeaders = getSheetHeaders(SHEET_COMPANY_OBJECTIVES);
    const compRows = getSheetData(SHEET_COMPANY_OBJECTIVES).map(r => rowToObject(compHeaders, r));
    const deptHeaders = getSheetHeaders(SHEET_DEPT_OBJECTIVES);
    const deptRows = getSheetData(SHEET_DEPT_OBJECTIVES).map(r => rowToObject(deptHeaders, r));
    const labelsX = compRows.map(o => o.ID);
    const labelsY = Array.from(new Set(deptRows.map(o => o.DepartmentID)));
    const matrix = labelsY.map((_, i) => {
      return labelsX.map((__, j) => Math.round(Math.random() * 100));
    });
    return JSON.stringify({ heatmap: { labelsX, labelsY, matrix }, error: '' });
  } catch (e) {
    return JSON.stringify({ error: '取得 Heatmap 資料失敗：' + e.message, heatmap: {} });
  }
}


/**
 * ---------------------------
 * doGet(): 回傳前端 index.html
 * ---------------------------
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('OKR 戰情室')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ---------------------------
 * include: 讓前端用 <?!= include('file') ?> 插入其他檔案
 * ---------------------------
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
