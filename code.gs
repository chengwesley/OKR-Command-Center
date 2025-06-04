// Code.gs - Google Apps Script 後端程式碼

// 您的 Google Sheet ID
// 請將 'YOUR_GOOGLE_SHEET_ID_HERE' 替換為您實際的 Google Sheets 試算表 ID。
// 試算表 ID 可以在瀏覽器網址列中找到，位於 /d/ 和 /edit 之間。
const SPREADSHEET_ID = '1rVCQ3EnWG-s4lLupa2k4wsydLoCSvEoYegVIY06n22E'; // <--- !!! 請務必將此處替換為您實際的試算表 ID !!!

// 各個工作表的名稱定義
// 使用 Object.freeze() 確保此物件及其屬性在運行時不會被意外修改。
// 請確保您的 Google Sheets 中實際的工作表名稱與此處定義的完全一致（包括大小寫）。
const SHEET_NAMES = Object.freeze({
    USERS: 'Users',
    DEPARTMENTS: 'Departments', // 新增部門工作表
    COMPANY_OBJECTIVES: 'Company_Objectives',
    MY_OBJECTIVES: 'My_Objectives',
    DEPARTMENT_OBJECTIVES: 'Department_Objectives',
    KEY_RESULTS: 'Key_Results',
    KEY_RESULTS_UPDATES: 'Key_Results_Updates',
    OKR_PERIODS: 'OKR_Periods', // OKR 週期管理工作表
    COMMENTS: 'Comments' // 評論工作表
});

// 會話令牌的有效時間 (分鐘)
const SESSION_EXPIRATION_MINUTES = 60; // 60 分鐘

// 用於密碼哈希的鹽值 (Salt) - **請替換為一個隨機的、足夠長的字串！**
// 這是非常重要的安全措施，切勿使用預設值！
const PASSWORD_SALT = '11111111111';


/**
 * 輔助函數：從指定工作表讀取所有數據 (跳過標題行)
 * 如果工作表不存在或為空，將拋出錯誤或返回空陣列。
 * @param {string} sheetName - 工作表名稱
 * @returns {Array<Array<any>>} - 數據陣列，每行是一個子陣列
 */
function getSheetData(sheetName) {
    // Logger.log(`[getSheetData] 嘗試獲取工作表數據: '${sheetName}' (類型: ${typeof sheetName})`); // 頻繁呼叫，暫時註釋掉日誌

    // 新增防禦性檢查：確保 sheetName 是有效的字串
    if (typeof sheetName !== 'string' || sheetName.trim() === '') {
        throw new Error(`無效的工作表名稱參數: '${sheetName}' (類型: ${typeof sheetName})。請確保傳遞有效字串。`);
    }

    // 新增防禦性檢查：確保 SPREADSHEET_ID 已設定且有效
    if (!SPREADSHEET_ID || SPREADSHEET_ID === 'YOUR_GOOGLE_SHEET_ID_HERE') {
        throw new Error('試算表 ID 未設定或不正確。請在 Code.gs 中設定 SPREADSHEET_ID。');
    }

    let spreadsheet;
    try {
        spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (e) {
        throw new Error(`無法打開試算表 ID: '${SPREADSHEET_ID}'。請檢查 ID 是否正確，以及 Apps Script 是否有權限訪問。錯誤: ${e.message}`);
    }
    
    // 確保試算表物件已成功打開
    if (!spreadsheet) { // 這通常不會發生，因為 openById 會拋出錯誤
        throw new Error(`無法打開試算表 ID: ${SPREADSHEET_ID}。請檢查 ID 和權限。`);
    }

    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
        throw new Error(`工作表 "${sheetName}" 不存在。請檢查名稱是否正確。`);
    }
    const range = sheet.getDataRange();
    const values = range.getValues();
    if (values.length === 0 || values[0].length === 0) { // 檢查是否完全為空
        // Logger.log(`[getSheetData] 工作表 "${sheetName}" 為空或無數據。`); // 頻繁呼叫，暫時註釋掉日誌
        return [];
    }
    // Logger.log(`[getSheetData] 成功獲取工作表 "${sheetName}" 數據，共 ${values.length - 1} 行。`); // 頻繁呼叫，暫時註釋掉日誌
    return values.slice(1); // 跳過標題行
}

/**
 * 輔助函數：從指定工作表讀取標題行
 * @param {string} sheetName - 工作表名稱
 * @returns {Array<string>} - 標題行陣列
 */
function getSheetHeaders(sheetName) {
    // Logger.log(`[getSheetHeaders] 嘗試獲取工作表標題: '${sheetName}' (類型: ${typeof sheetName})`); // 頻繁呼叫，暫時註釋掉日誌

    // 新增防禦性檢查：確保 sheetName 是有效的字串
    if (typeof sheetName !== 'string' || sheetName.trim() === '') {
        throw new Error(`無效的工作表名稱參數: '${sheetName}' (類型: ${typeof sheetName})。請確保傳遞有效字串。`);
    }

    // 新增防禦性檢查：確保 SPREADSHEET_ID 已設定且有效
    if (!SPREADSHEET_ID || SPREADSHEET_ID === 'YOUR_GOOGLE_SHEET_ID_HERE') {
        throw new Error('試算表 ID 未設定或不正確。請在 Code.gs 中設定 SPREADSHEET_ID。');
    }

    let spreadsheet;
    try {
        spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (e) {
        throw new Error(`無法打開試算表 ID: '${SPREADSHEET_ID}'。請檢查 ID 是否正確，以及 Apps Script 是否有權限訪問。錯誤: ${e.message}`);
    }

    // 確保試算表物件已成功打開
    if (!spreadsheet) { // 這通常不會發生，因為 openById 會拋出錯誤
        throw new Error(`無法打開試算表 ID: ${SPREADSHEET_ID}。請檢查 ID 和權限。`);
    }

    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
        throw new Error(`工作表 "${sheetName}" 不存在。請檢查名稱是否正確。`);
    }
    // 讀取第一行，直到最後一列
    const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    const headers = range.getValues()[0];
    // 過濾掉空標題，確保標題是有效字串
    const filteredHeaders = headers.filter(header => typeof header === 'string' && header.trim() !== '');
    if (filteredHeaders.length === 0) {
        throw new Error(`工作表 "${sheetName}" 的標題行為空。`);
    }
    // Logger.log(`[getSheetHeaders] 成功獲取工作表 "${sheetName}" 標題: ${filteredHeaders.join(', ')}`); // 頻繁呼叫，暫時註釋掉日誌
    return filteredHeaders;
}

/**
 * 輔助函數：將單行數據和標題轉換為物件
 * @param {Array<string>} headers - 標題行陣列
 * @param {Array<any>} row - 單行數據陣列
 * @returns {Object} - 轉換後的物件
 */
function rowToObject(headers, row) {
    const obj = {};
    headers.forEach((header, index) => {
        obj[header] = row[index];
    });
    return obj;
}

/**
 * 輔助函數：對字串進行 SHA-256 哈希處理
 * @param {string} text - 要哈希的字串
 * @returns {string} - 哈希後的十六進位字串
 */
function hashString(text) {
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text);
    let hexString = '';
    for (let i = 0; i < digest.length; i++) {
        let byte = digest[i]; // <--- 將 const 改為 let
        if (byte < 0) { // Convert signed byte to unsigned
            byte += 256;
        }
        const hex = byte.toString(16);
        hexString += (hex.length === 1 ? '0' : '') + hex;
    }
    return hexString;
}

/**
 * 輔助函數：驗證會話令牌並獲取用戶資訊
 * @param {string} sessionToken - 從前端傳遞的會話令牌
 * @param {Array<string>} [requiredRoles] - 可選，操作所需的角色陣列
 * @returns {Object} - 包含 userEmail, userRole, userDepartment, userId 的用戶物件
 * @throws {Error} - 如果令牌無效、過期或權限不足
 */
function authenticateAndAuthorize(sessionToken, requiredRoles) {
    Logger.log(`[Auth] 嘗試驗證令牌: ${sessionToken}`);
    if (!sessionToken) {
        throw new Error("未提供會話令牌。請重新登入。");
    }

    const scriptProperties = PropertiesService.getScriptProperties();
    const sessionDataJson = scriptProperties.getProperty(sessionToken);

    if (!sessionDataJson) {
        throw new Error("會話令牌無效或已過期。請重新登入。");
    }

    const sessionData = JSON.parse(sessionDataJson);
    const currentTime = new Date().getTime();
    const expirationTime = sessionData.timestamp + (SESSION_EXPIRATION_MINUTES * 60 * 1000);

    if (currentTime > expirationTime) {
        scriptProperties.deleteProperty(sessionToken); // 清除過期令牌
        throw new Error("會話已過期。請重新登入。");
    }

    // 更新令牌的 timestamp 以延長會話 (簡單的滑動會話)
    sessionData.timestamp = currentTime;
    scriptProperties.setProperty(sessionToken, JSON.stringify(sessionData));

    const userEmail = sessionData.email;
    const userRole = sessionData.role;
    const userDepartment = sessionData.department;
    const userId = sessionData.id;

    Logger.log(`[Auth] 令牌驗證成功。用戶: ${userEmail}, 角色: ${userRole}`);

    // 檢查角色權限
    if (requiredRoles && requiredRoles.length > 0) {
        if (!requiredRoles.includes(userRole)) {
            throw new Error(`權限不足。您的角色 (${userRole}) 無權執行此操作。所需角色: ${requiredRoles.join(', ')}`);
        }
    }

    return { userEmail, userRole, userDepartment, userId };
}


/**
 * 處理 Web 應用程式的 GET 請求。
 * 當前端載入時，此函數會被呼叫以提供 HTML 頁面。
 * @param {GoogleAppsScript.Events.DoGet} e - 請求事件物件
 * @returns {GoogleAppsScript.HTML.HtmlOutput} - HTML 內容
 */
function doGet(e) {
    Logger.log("[doGet] 函數開始執行。");
    
    // 在自定義登入模式下，doGet 不再從 Session 獲取用戶 Email 或角色
    // 這些資訊將由前端登入後通過會話令牌傳遞
    const userEmail = ''; // 預設為空，前端會處理未登入狀態
    const userRole = ''; // 預設為空

    const template = HtmlService.createTemplateFromFile('index');
    template.userEmail = userEmail; // 前端會檢查是否為空，引導登入
    template.userRole = userRole;   // 前端會檢查是否為空，引導登導

    Logger.log("[doGet] 函數結束執行，返回 HTML 模板。");
    return template.evaluate()
        .setTitle('Kdan OKR 系統')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // 允許在 iframe 中顯示
}

/**
 * 處理用戶登入請求
 * @param {string} email - 用戶 Email (作為用戶名)
 * @param {string} password - 用戶密碼
 * @returns {Object} - 包含 success, message, sessionToken, userEmail, userRole, userDepartment, userId 的物件
 */
function loginUser(email, password) {
    Logger.log(`[loginUser] 嘗試登入用戶: ${email}`);
    try {
        const userHeaders = getSheetHeaders(SHEET_NAMES.USERS);
        const userData = getSheetData(SHEET_NAMES.USERS);
        const userRow = userData.find(row => rowToObject(userHeaders, row).Email === email);

        if (!userRow) {
            return { success: false, message: "用戶名或密碼不正確。" };
        }

        const user = rowToObject(userHeaders, userRow);
        
        // 檢查用戶是否啟用
        if (user.IsActive === false || user.IsActive === 'FALSE') {
            return { success: false, message: "您的帳號已被禁用，請聯繫管理員。" };
        }

        const hashedPassword = hashString(password + PASSWORD_SALT); // 使用鹽值哈希密碼
        if (hashedPassword !== user.PasswordHash) { // 假設 Users 表中有 PasswordHash 欄位
            return { success: false, message: "用戶名或密碼不正確。" };
        }

        // 登入成功，生成會話令牌
        const sessionToken = Utilities.getUuid();
        const sessionData = {
            email: user.Email,
            role: user.Role,
            department: user.Department,
            id: user.ID, // 用戶簡短 ID
            timestamp: new Date().getTime()
        };
        PropertiesService.getScriptProperties().setProperty(sessionToken, JSON.stringify(sessionData));
        Logger.log(`[loginUser] 用戶 ${email} 登入成功，生成令牌: ${sessionToken}`);

        return {
            success: true,
            message: "登入成功！",
            sessionToken: sessionToken,
            userEmail: user.Email,
            userRole: user.Role,
            userDepartment: user.Department,
            userId: user.ID
        };

    } catch (error) {
        Logger.log(`[loginUser] 登入失敗: ${error.message}`);
        return { success: false, message: `登入失敗: ${error.message}` };
    }
}

/**
 * 處理用戶登出請求
 * @param {string} sessionToken - 會話令牌
 * @returns {Object} - 包含 success 和 message 的物件
 */
function logoutUser(sessionToken) {
    Logger.log(`[logoutUser] 嘗試登出用戶，令牌: ${sessionToken}`);
    try {
        PropertiesService.getScriptProperties().deleteProperty(sessionToken);
        return { success: true, message: "登出成功。" };
    } catch (error) {
        Logger.log(`[logoutUser] 登出失敗: ${error.message}`);
        return { success: false, message: `登出失敗: ${error.message}` };
    }
}

/**
 * 從 Google Sheets 獲取所有相關的 OKR 數據，並組合成前端所需的格式。
 * 此函數會處理多個工作表的數據讀取和關聯。
 * @param {string} sessionToken - 會話令牌
 * @returns {string} - JSON 字串，包含 companyOkrs, myOkrs, departmentOkrs, allObjectives 的物件，或包含 error 訊息的物件。
 */
function getOKRs(sessionToken) {
    Logger.log("[getOKRs] 函數開始執行。");
    
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['Chairman', 'OKR_Admin', 'HR_Admin', 'Department_Manager', 'CLevel_Exec', 'Employee']);
    } catch (e) {
        Logger.log(`[getOKRs] 認證失敗: ${e.message}`);
        return JSON.stringify({ error: e.message });
    }

    // 新增：明確記錄 SHEET_NAMES 物件的內容
    Logger.log(`[getOKRs] SHEET_NAMES 物件內容: ${JSON.stringify(SHEET_NAMES)}`);

    // **新增：在調用前，明確檢查 SPREADSHEET_ID**
    if (!SPREADSHEET_ID || SPREADSHEET_ID === 'YOUR_GOOGLE_SHEET_ID_HERE') {
        const errorMessage = '致命錯誤: 試算表 ID 未設定或不正確。請在 Code.gs 中設定 SPREADSHEET_ID。';
        Logger.log(`[getOKRs] ${errorMessage}`);
        return JSON.stringify({ error: errorMessage }); // <--- 返回 JSON 字串
    }

    let result = {
        companyOkrs: [],
        myOkrs: [],
        departmentOkrs: [],
        allObjectives: [], // 新增：用於 Admin 頁面和承接關係查找
        error: '' // 初始化為空字串
    };

    // 新增：嚴格檢查 SHEET_NAMES 的每個屬性
    for (const key in SHEET_NAMES) {
        if (Object.hasOwnProperty.call(SHEET_NAMES, key)) {
            const sheetNameValue = SHEET_NAMES[key];
            if (typeof sheetNameValue !== 'string' || sheetNameValue.trim() === '') {
                const errorMessage = `配置錯誤: SHEET_NAMES.${key} 未正確定義或為空。請檢查 Code.gs 中的 SHEET_NAMES 常數。值: '${sheetNameValue}' (類型: ${typeof sheetNameValue})`;
                Logger.log(`[getOKRs] 致命錯誤: ${errorMessage}`);
                result.error = errorMessage;
                return JSON.stringify(result); // <--- 返回 JSON 字串
            }
        }
    }
    Logger.log("[getOKRs] SHEET_NAMES 配置檢查通過。");


    try { // 最外層的 try-catch，確保總能返回一個物件
        let allKRs = [];
        try {
            const keyResultsSheetName = SHEET_NAMES.KEY_RESULTS;
            Logger.log(`[getOKRs] 即將呼叫 getSheetHeaders，參數為: '${keyResultsSheetName}' (類型: ${typeof keyResultsSheetName})`);
            const krHeaders = getSheetHeaders(keyResultsSheetName);
            
            Logger.log(`[getOKRs] 即將呼叫 getSheetData，參數為: '${keyResultsSheetName}' (類型: ${typeof keyResultsSheetName})`);
            allKRs = getSheetData(keyResultsSheetName).map(row => rowToObject(krHeaders, row));
            Logger.log(`[getOKRs] 成功載入 ${allKRs.length} 個 Key Results。`);
        } catch (e) {
            Logger.log(`[getOKRs] 錯誤: 獲取 Key_Results 數據失敗: ${e.message}`);
            result.error += `獲取 Key_Results 數據失敗: ${e.message}; `;
            // 這裡不直接返回，嘗試載入其他數據，以便前端能看到部分內容和錯誤訊息
        }

        // 獲取所有 Objective，並計算其進度，關聯 KRs
        const allObjectives = []; // 用於存儲所有 Objective，方便承接關係查找

        // 處理公司 OKR
        try {
            const companyObjectivesSheetName = SHEET_NAMES.COMPANY_OBJECTIVES;
            const companyObjHeaders = getSheetHeaders(companyObjectivesSheetName);
            const companyObjsData = getSheetData(companyObjectivesSheetName);
            const companyObjs = companyObjsData.map(row => {
                const obj = rowToObject(companyObjHeaders, row);
                obj.objective_type = 'Company'; // 新增類型屬性
                obj.krs = (obj.KRs ? String(obj.KRs).split(',').map(id => id.trim()).filter(id => id !== '') : [])
                           .map(krId => allKRs.find(kr => kr.ID === krId)).filter(Boolean); // 查找並過濾掉 undefined
                
                // 計算 Objective 進度 (所有 KR 的平均進度)
                const totalProgress = obj.krs.reduce((sum, kr) => sum + (kr.Progress ? parseFloat(kr.Progress) : 0), 0);
                obj.Progress = obj.krs.length > 0 ? (totalProgress / obj.krs.length) : (obj.Progress ? parseFloat(obj.Progress) : 0); // 如果沒有 KRs，使用表格中的進度
                return obj;
            });
            result.companyOkrs = companyObjs;
            allObjectives.push(...companyObjs); // 加入總 Objective 列表
            Logger.log(`[getOKRs] 成功載入 ${result.companyOkrs.length} 個公司 OKR。`);
        } catch (e) {
            Logger.log(`[getOKRs] 錯誤: 獲取公司 OKR 失敗: ${e.message}`);
            result.error += `獲取公司 OKR 失敗: ${e.message}; `;
        }

        // 處理部門 OKR
        try {
            const departmentObjectivesSheetName = SHEET_NAMES.DEPARTMENT_OBJECTIVES;
            const deptObjHeaders = getSheetHeaders(departmentObjectivesSheetName);
            const departmentObjsData = getSheetData(departmentObjectivesSheetName);
            const departmentObjs = departmentObjsData.map(row => {
                const obj = rowToObject(deptObjHeaders, row);
                obj.objective_type = 'Department'; // 新增類型屬性
                obj.krs = (obj.KRs ? String(obj.KRs).split(',').map(id => id.trim()).filter(id => id !== '') : [])
                           .map(krId => allKRs.find(kr => kr.ID === krId)).filter(Boolean);
                
                const totalProgress = obj.krs.reduce((sum, kr) => sum + (kr.Progress ? parseFloat(kr.Progress) : 0), 0);
                obj.Progress = obj.krs.length > 0 ? (totalProgress / obj.krs.length) : (obj.Progress ? parseFloat(obj.Progress) : 0);
                return obj;
            });
            result.departmentOkrs = departmentObjs;
            allObjectives.push(...departmentObjs); // 加入總 Objective 列表
            Logger.log(`[getOKRs] 成功載入 ${result.departmentOkrs.length} 個部門 OKR。`);
        } catch (e) {
            Logger.log(`[getOKRs] 錯誤: 獲取部門 OKR 失敗: ${e.message}`);
            result.error += `獲取部門 OKR 失敗: ${e.message}; `;
        }

        // 處理個人 OKR (根據當前用戶 Email)
        const currentUserEmail = authUser.userEmail; // 使用認證後的用戶 Email
        try {
            const myObjectivesSheetName = SHEET_NAMES.MY_OBJECTIVES;
            const myObjHeaders = getSheetHeaders(myObjectivesSheetName);
            const myObjsData = getSheetData(myObjectivesSheetName);
            const myObjs = myObjsData
                .map(row => {
                    const obj = rowToObject(myObjHeaders, row);
                    obj.objective_type = 'My'; // 新增類型屬性
                    obj.krs = (obj.KRs ? String(obj.KRs).split(',').map(id => id.trim()).filter(id => id !== '') : [])
                               .map(krId => allKRs.find(kr => kr.ID === krId)).filter(Boolean);
                    
                    const totalProgress = obj.krs.reduce((sum, kr) => sum + (kr.Progress ? parseFloat(kr.Progress) : 0), 0);
                    obj.Progress = obj.krs.length > 0 ? (totalProgress / obj.krs.length) : (obj.Progress ? parseFloat(obj.Progress) : 0);
                    return obj;
                });
                // .filter(obj => obj.OwnerEmail === currentUserEmail); // 這裡不再過濾，前端會過濾

            result.myOkrs = myObjs;
            allObjectives.push(...myObjs); // 加入總 Objective 列表
            Logger.log(`[getOKRs] 成功載入 ${result.myOkrs.length} 個個人 OKR。`);
        } catch (e) {
            Logger.log(`[getOKRs] 錯誤: 獲取個人 OKR 失敗: ${e.message}`);
            result.error += `獲取個人 OKR 失敗: ${e.message}; `;
        }

        // 獲取所有 Objective，用於後續的父子關係追溯和 Admin 頁面
        result.allObjectives = allObjectives; // 將所有 Objective 也傳遞給前端，方便 Admin 頁面處理

    } catch (outerError) {
        // 捕獲 getOKRs 函數頂層的任何意外錯誤
        Logger.log(`[getOKRs] 錯誤: getOKRs 函數發生未預期的錯誤: ${outerError.message}`);
        result.error = (result.error ? result.error + "; " : "") + `getOKRs 函數未預期錯誤: ${outerError.message}`;
    }
    
    Logger.log("[getOKRs] 函數結束執行，返回結果。");
    return JSON.stringify(result); // <--- 將結果物件轉換為 JSON 字串
}

/**
 * 獲取所有用戶的 Email 和名稱，以及所有部門名稱。
 * 用於前端的下拉選單。
 * @param {string} sessionToken - 會話令牌
 * @returns {string} JSON 字串，包含 users 和 departments 陣列。
 */
function getUsersAndDepartments(sessionToken) {
    Logger.log("[getUsersAndDepartments] 函數開始執行。");
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['Chairman', 'OKR_Admin', 'HR_Admin', 'Department_Manager', 'CLevel_Exec', 'Employee']); // 所有用戶都可以獲取列表
    } catch (e) {
        Logger.log(`[getUsersAndDepartments] 認證失敗: ${e.message}`);
        return JSON.stringify({ error: e.message });
    }

    let result = {
        users: [],
        departments: [],
        error: ''
    };
    try {
        const userHeaders = getSheetHeaders(SHEET_NAMES.USERS);
        const userData = getSheetData(SHEET_NAMES.USERS);
        result.users = userData.map(row => {
            const userObj = rowToObject(userHeaders, row);
            return { email: userObj.Email, name: userObj.Name, department: userObj.Department, role: userObj.Role, id: userObj.ID }; // 包含用戶ID
        });
        Logger.log(`[getUsersAndDepartments] 成功載入 ${result.users.length} 個用戶。`);
    } catch (e) {
        Logger.log(`[getUsersAndDepartments] 錯誤: 獲取用戶數據失敗: ${e.message}`);
        result.error += `獲取用戶數據失敗: ${e.message}; `;
    }

    try {
        // 從 Departments 工作表獲取部門列表
        const deptHeaders = getSheetHeaders(SHEET_NAMES.DEPARTMENTS);
        const deptData = getSheetData(SHEET_NAMES.DEPARTMENTS);
        result.departments = deptData.map(row => rowToObject(deptHeaders, row)); // 獲取所有部門物件
        Logger.log(`[getUsersAndDepartments] 成功載入 ${result.departments.length} 個部門。`);
    } catch (e) {
        Logger.log(`[getUsersAndDepartments] 錯誤: 獲取部門數據失敗: ${e.message}`);
        result.error += `獲取部門數據失敗: ${e.message}; `;
    }
    Logger.log("[getUsersAndDepartments] 函數結束執行。");
    return JSON.stringify(result);
}

/**
 * 獲取指定類型（公司、部門、個人）的 Objective 列表，用於父級目標的選擇。
 * @param {string} sessionToken - 會話令牌
 * @param {string} objectiveType - Objective 類型 ('Company', 'Department', 'My')
 * @returns {string} JSON 字串，包含 objectives 陣列。
 */
function getObjectivesForSelection(sessionToken, objectiveType) {
    Logger.log(`[getObjectivesForSelection] 函數開始執行。類型: ${objectiveType}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['Chairman', 'OKR_Admin', 'HR_Admin', 'Department_Manager', 'CLevel_Exec', 'Employee']); // 所有用戶都可以獲取列表
    } catch (e) {
        Logger.log(`[getObjectivesForSelection] 認證失敗: ${e.message}`);
        return JSON.stringify({ error: e.message });
    }

    let result = {
        objectives: [],
        error: ''
    };
    let sheetName = '';
    let objHeaders = [];
    let objData = [];

    try {
        if (objectiveType === 'Company') {
            sheetName = SHEET_NAMES.COMPANY_OBJECTIVES;
        } else if (objectiveType === 'Department') {
            sheetName = SHEET_NAMES.DEPARTMENT_OBJECTIVES;
        } else if (objectiveType === 'My') {
            sheetName = SHEET_NAMES.MY_OBJECTIVES;
        } else {
            throw new Error(`無效的 Objective 類型: ${objectiveType}`);
        }

        objHeaders = getSheetHeaders(sheetName);
        objData = getSheetData(sheetName);

        result.objectives = objData.map(row => {
            const obj = rowToObject(objHeaders, row);
            return { id: obj.ID, title: obj.Title, ownerEmail: obj.OwnerEmail };
        });
        Logger.log(`[getObjectivesForSelection] 成功載入 ${result.objectives.length} 個 ${objectiveType} Objective。`);

    } catch (e) {
        Logger.log(`[getObjectivesForSelection] 錯誤: 獲取 ${objectiveType} Objective 失敗: ${e.message}`);
        result.error += `獲取 ${objectiveType} Objective 失敗: ${e.message}; `;
    }
    Logger.log("[getObjectivesForSelection] 函數結束執行。");
    return JSON.stringify(result);
}

/**
 * 輔助函數：生成下一個 Objective 的 ID
 * @param {string} type - Objective 類型 ('Company', 'Department', 'My')
 * @param {string} periodId - 週期 ID (例如 '2025Q2')
 * @param {string} departmentId - 部門 ID (例如 'PROD') (如果適用)
 * @param {string} userId - 用戶 ID (例如 'u1') (如果適用)
 * @returns {string} 新生成的 Objective ID
 */
function _generateNextObjectiveId(type, periodId, departmentId, userId) {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheetName;
    let prefix;
    let existingIds = [];

    if (type === 'Company') {
        sheetName = SHEET_NAMES.COMPANY_OBJECTIVES;
        prefix = `C-${periodId}-O-`;
    } else if (type === 'Department') {
        sheetName = SHEET_NAMES.DEPARTMENT_OBJECTIVES;
        if (!departmentId) throw new Error("生成部門 Objective ID 時缺少部門 ID。");
        prefix = `D-${periodId}-${departmentId}-O-`;
    } else if (type === 'My') {
        sheetName = SHEET_NAMES.MY_OBJECTIVES;
        if (!userId) throw new Error("生成個人 Objective ID 時缺少用戶 ID。");
        prefix = `I-${periodId}-${userId}-O-`;
    } else {
        throw new Error(`無效的 Objective 類型無法生成 ID: ${type}`);
    }

    try {
        const headers = getSheetHeaders(sheetName);
        const data = getSheetData(sheetName);
        existingIds = data.map(row => rowToObject(headers, row).ID)
                          .filter(id => id && String(id).startsWith(prefix));
    } catch (e) {
        Logger.log(`[ID Gen] 獲取現有 Objective ID 失敗 (${sheetName}): ${e.message}`);
        // 如果工作表不存在或為空，則從 1 開始
        existingIds = [];
    }

    let maxNum = 0;
    existingIds.forEach(id => {
        const parts = String(id).split('-O-');
        if (parts.length > 1) {
            const numPart = parts[1]; // 假設編號在 -O- 後
            const num = parseInt(numPart, 10);
            if (!isNaN(num) && num > maxNum) {
                maxNum = num;
            }
        }
    });

    const nextNum = maxNum + 1;
    return `${prefix}${String(nextNum).padStart(2, '0')}`;
}

/**
 * 輔助函數：生成下一個 Key Result 的 ID
 * @param {string} parentObjectiveId - 父級 Objective 的 ID (例如 'C-2025Q2-O-01', 'D-2025Q2-PROD-O-01', 'I-2025Q2-u1-O-01')
 * @returns {string} 新生成的 Key Result ID
 */
function _generateNextKeyResultId(parentObjectiveId) {
    if (!parentObjectiveId) throw new Error("生成 Key Result ID 時缺少父級 Objective ID。");

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const krSheet = spreadsheet.getSheetByName(SHEET_NAMES.KEY_RESULTS);
    let existingKrIds = [];

    try {
        const headers = getSheetHeaders(SHEET_NAMES.KEY_RESULTS);
        const data = getSheetData(SHEET_NAMES.KEY_RESULTS);
        existingKrIds = data.map(row => rowToObject(headers, row).ID)
                            .filter(id => id && String(id).startsWith(`${parentObjectiveId}-KR-`));
    } catch (e) {
        Logger.log(`[ID Gen] 獲取現有 Key Result ID 失敗 (${parentObjectiveId}): ${e.message}`);
        existingKrIds = [];
    }

    let maxNum = 0;
    existingKrIds.forEach(id => {
        const parts = String(id).split('-KR-');
        if (parts.length > 1) {
            const numPart = parts[1]; // 假設編號在 -KR- 後
            const num = parseInt(numPart, 10);
            if (!isNaN(num) && num > maxNum) {
                maxNum = num;
            }
        }
    });

    const nextNum = maxNum + 1;
    return `${parentObjectiveId}-KR-${String(nextNum).padStart(2, '0')}`;
}


/**
 * 新增 Objective 到對應的工作表。
 * @param {string} sessionToken - 會話令牌
 * @param {Object} objData - 包含 Objective 資訊的物件。
 * - Title (string)
 * - Description (string)
 * - OwnerEmail (string)
 * - Period (string)
 * - Type (string: 'Company', 'Department', 'My')
 * - ParentObjectiveID (string, 可選)
 * - DepartmentID (string, 僅限 Type 為 'Department' 時需要，來自 Departments.ID)
 * @returns {Object} - 包含 success 狀態和 message 的物件。
 */
function createObjective(sessionToken, objData) {
    Logger.log(`[createObjective] 函數開始執行。objData: ${JSON.stringify(objData)}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['OKR_Admin', 'Chairman', 'CLevel_Exec', 'Department_Manager']); // 假設這些角色可以創建 Objective
    } catch (e) {
        Logger.log(`[createObjective] 認證失敗: ${e.message}`);
        return { success: false, message: e.message };
    }

    let targetSheetName;
    let headers;
    let newRow = [];
    let objId;

    try {
        // 獲取用戶 ID 和部門 ID 以便生成 Objective ID
        let userId = '';
        let departmentId = '';

        const userHeaders = getSheetHeaders(SHEET_NAMES.USERS);
        const allUsers = getSheetData(SHEET_NAMES.USERS).map(row => rowToObject(userHeaders, row));
        const ownerUser = allUsers.find(u => u.Email === objData.OwnerEmail);
        if (ownerUser) {
            userId = ownerUser.ID;
        } else {
            throw new Error(`找不到負責人 ${objData.OwnerEmail} 的用戶 ID。`);
        }

        if (objData.Type === 'Department' && objData.DepartmentID) {
            departmentId = objData.DepartmentID;
        }

        // 生成 Objective ID
        objId = _generateNextObjectiveId(objData.Type, objData.Period, departmentId, userId);
        Logger.log(`[createObjective] 生成的 Objective ID: ${objId}`);

        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

        if (objData.Type === 'Company') {
            targetSheetName = SHEET_NAMES.COMPANY_OBJECTIVES;
            headers = getSheetHeaders(targetSheetName);
            newRow = [
                objId,
                objData.Title,
                objData.Description,
                objData.OwnerEmail,
                objData.Period,
                0, // Progress 預設為 0
                'Draft', // Status 預設為 Draft
                '' // KRs 預設為空字串
            ];
        } else if (objData.Type === 'Department') {
            targetSheetName = SHEET_NAMES.DEPARTMENT_OBJECTIVES;
            headers = getSheetHeaders(targetSheetName);
            newRow = [
                objId,
                objData.Title,
                objData.Description,
                objData.OwnerEmail,
                objData.Period,
                0, // Progress 預設為 0
                'Draft', // Status 預設為 Draft
                objData.DepartmentID || '', // 使用部門 ID
                objData.ParentObjectiveID || '', // 父級 Objective ID
                '' // KRs 預設為空字串
            ];
        } else if (objData.Type === 'My') {
            targetSheetName = SHEET_NAMES.MY_OBJECTIVES;
            headers = getSheetHeaders(targetSheetName);
            newRow = [
                objId,
                objData.Title,
                objData.Description,
                objData.OwnerEmail,
                objData.Period,
                0, // Progress 預設為 0
                'Draft', // Status 預設為 Draft
                objData.ParentObjectiveID || '', // 父級 Objective ID
                '' // KRs 預設為空字串
            ];
        } else {
            throw new Error(`無效的 Objective 類型: ${objData.Type}`);
        }

        const targetSheet = spreadsheet.getSheetByName(targetSheetName);
        if (!targetSheet) {
            throw new Error(`工作表 "${targetSheetName}" 不存在。`);
        }

        // 確保新行數據的長度與標題長度匹配，不足的用空字串填充
        while (newRow.length < headers.length) {
            newRow.push('');
        }

        targetSheet.appendRow(newRow);
        Logger.log(`[createObjective] 成功新增 Objective: ${objId} 到 ${targetSheetName}`);
        return { success: true, message: `Objective "${objData.Title}" 新增成功！`, id: objId };

    } catch (error) {
        Logger.log(`[createObjective] 錯誤: ${error.message}`);
        return { success: false, message: `新增 Objective 失敗: ${error.message}` };
    }
}

/**
 * 新增 Key Result 到 Key_Results 工作表，並更新其父級 Objective 的 KRs 欄位。
 * @param {string} sessionToken - 會話令牌
 * @param {Object} krData - 包含 Key Result 資訊的物件。
 * - ObjectiveID (string) - 父級 Objective 的 ID
 * - Description (string)
 * - OwnerEmail (string)
 * - MetricType (string)
 * - StartValue (number)
 * - TargetValue (number)
 * - Unit (string)
 * @returns {Object} - 包含 success 狀態和 message 的物件。
 */
function createKeyResult(sessionToken, krData) {
    Logger.log(`[createKeyResult] 函數開始執行。krData: ${JSON.stringify(krData)}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['Chairman', 'OKR_Admin', 'CLevel_Exec', 'Department_Manager', 'Employee']); // 假設這些角色可以創建 KR
    } catch (e) {
        Logger.log(`[createKeyResult] 認證失敗: ${e.message}`);
        return { success: false, message: e.message };
    }

    const krId = _generateNextKeyResultId(krData.ObjectiveID); // 生成新的 KR ID

    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const krSheet = spreadsheet.getSheetByName(SHEET_NAMES.KEY_RESULTS);
        if (!krSheet) {
            throw new Error(`工作表 "${SHEET_NAMES.KEY_RESULTS}" 不存在。`);
        }

        const krHeaders = getSheetHeaders(SHEET_NAMES.KEY_RESULTS);
        const newKrRow = [
            krId,
            krData.ObjectiveID,
            krData.Description,
            krData.OwnerEmail,
            krData.MetricType,
            krData.StartValue,
            krData.TargetValue,
            krData.StartValue, // CurrentValue 預設為 StartValue
            krData.Unit,
            0, // Progress 預設為 0
            'On Track', // ConfidenceLevel 預設為 On Track
            new Date().toLocaleString(), // LastUpdated
            authUser.userEmail, // LastUpdatedByEmail
            '' // Comment
        ];

        // 確保新行數據的長度與標題長度匹配，不足的用空字串填充
        while (newKrRow.length < krHeaders.length) {
            newKrRow.push('');
        }

        krSheet.appendRow(newKrRow);
        Logger.log(`[createKeyResult] 成功新增 Key Result: ${krId}`);

        // --- 更新父級 Objective 的 KRs 欄位 ---
        const objSheets = [
            { name: SHEET_NAMES.COMPANY_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.COMPANY_OBJECTIVES) },
            { name: SHEET_NAMES.MY_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.MY_OBJECTIVES) },
            { name: SHEET_NAMES.DEPARTMENT_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.DEPARTMENT_OBJECTIVES) }
        ];

        let parentObjectiveFound = false;
        for (const objSheetInfo of objSheets) {
            const currentObjSheet = spreadsheet.getSheetByName(objSheetInfo.name);
            const objDataRange = currentObjSheet.getDataRange();
            const objValues = objDataRange.getValues();

            for (let i = 1; i < objValues.length; i++) { // 從第二行開始
                const obj = rowToObject(objSheetInfo.headers, objValues[i]);
                if (obj.ID === krData.ObjectiveID) {
                    const krColIndex = objSheetInfo.headers.indexOf('KRs') + 1;
                    if (krColIndex > 0) {
                        let currentKRs = obj.KRs ? String(obj.KRs).trim() : '';
                        if (currentKRs !== '') {
                            currentKRs += `,${krId}`;
                        } else {
                            currentKRs = krId;
                        }
                        currentObjSheet.getRange(i + 1, krColIndex).setValue(currentKRs);
                        Logger.log(`[createKeyResult] 成功更新父級 Objective (${krData.ObjectiveID}) 的 KRs 欄位。新 KRs: ${currentKRs}`);
                        parentObjectiveFound = true;
                        break; // 找到父級 Objective 後跳出循環
                    }
                }
            }
            if (parentObjectiveFound) break; // 如果在某個工作表找到並更新，則跳出外層循環
        }

        if (!parentObjectiveFound) {
            Logger.log(`[createKeyResult] 警告: 未找到 ID 為 ${krData.ObjectiveID} 的父級 Objective 來更新 KRs 欄位。`);
        } else {
            // 如果父級 Objective 找到並更新了 KRs 欄位，則重新計算其進度
            calculateObjectiveProgressAndScore(krData.ObjectiveID);
        }

        return { success: true, message: `Key Result "${krData.Description}" 新增成功！`, id: krId };

    } catch (error) {
        Logger.log(`[createKeyResult] 錯誤: ${error.message}`);
        return { success: false, message: `新增 Key Result 失敗: ${error.message}` };
    }
}


/**
 * 更新 Key Result 的當前值和進度到 Google Sheets。
 * 此函數會更新 Key_Results 工作表並在 Key_Results_Updates 中記錄歷史。
 * @param {string} sessionToken - 會話令牌
 * @param {string} krId - 要更新的 Key Result 的 ID
 * @param {number} newValue - 新的當前值
 * @param {string} comment - 用戶提供的更新評論
 * @returns {Object} - 包含 success 狀態和 message 的物件。
 */
function updateKeyResultProgress(sessionToken, krId, newValue, comment) {
    Logger.log(`[updateKeyResultProgress] 函數開始執行。KR ID: ${krId}, 新值: ${newValue}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['Chairman', 'OKR_Admin', 'CLevel_Exec', 'Department_Manager', 'Employee']); // 假設這些角色可以更新 KR
    } catch (e) {
        Logger.log(`[updateKeyResultProgress] 認證失敗: ${e.message}`);
        return { success: false, message: e.message };
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const krSheet = spreadsheet.getSheetByName(SHEET_NAMES.KEY_RESULTS);
        const krUpdateSheet = spreadsheet.getSheetByName(SHEET_NAMES.KEY_RESULTS_UPDATES);
        
        if (!krSheet || !krUpdateSheet) {
            throw new Error("Key Results 或 Key Results Updates 工作表不存在，無法更新。");
        }

        const krHeaders = getSheetHeaders(SHEET_NAMES.KEY_RESULTS);
        const krData = krSheet.getDataRange().getValues(); // 包含標題行

        let rowIndexToUpdate = -1;
        let oldKrValue = null;
        let targetValue = 0;
        let startValue = 0;
        let krOwnerEmail = '';
        let parentObjectiveId = ''; // 用於後續更新父級 Objective 進度

        // 尋找要更新的 KR 行
        for (let i = 1; i < krData.length; i++) { // 從第二行開始 (跳過標題)
            const kr = rowToObject(krHeaders, krData[i]);
            if (kr.ID === krId) {
                rowIndexToUpdate = i + 1; // Apps Script 的行號是從 1 開始
                oldKrValue = parseFloat(kr.CurrentValue); // 確保是數字
                targetValue = parseFloat(kr.TargetValue); // 確保是數字
                startValue = parseFloat(kr.StartValue); // 確保是數字
                krOwnerEmail = kr.OwnerEmail;
                parentObjectiveId = kr.ObjectiveID; // 獲取父級 Objective ID
                break;
            }
        }

        if (rowIndexToUpdate === -1) {
            Logger.log(`[updateKeyResultProgress] 警告: 找不到 ID 為 ${krId} 的關鍵成果。`);
            return { success: false, message: `找不到 ID 為 ${krId} 的關鍵成果。` };
        }

        // 權限檢查：只有 KR 負責人、部門主管、C-level主管、OKR管理員、董事長才能更新
        let hasEditPermission = false;
        if (authUser.userEmail === krOwnerEmail) { // KR 負責人
            hasEditPermission = true;
        } else if (['OKR_Admin', 'Chairman'].includes(authUser.userRole)) { // OKR 管理員或董事長
            hasEditPermission = true;
        } else if (authUser.userRole === 'Department_Manager' && authUser.userDepartment === krOwnerEmail.split('@')[0].toUpperCase().replace(/[^A-Z0-9]/g, '')) { // 假設部門經理只能管理自己部門的KR
            // 這需要更複雜的邏輯來判斷部門經理是否是KR負責人的主管
            // 更精確的判斷需要查詢 Users 表格中 KR 負責人的部門，再比對部門經理的部門
            hasEditPermission = true; // 簡化處理
        } else if (authUser.userRole === 'CLevel_Exec') { // C-level 主管可以更新所有 KR
            hasEditPermission = true;
        }
        
        if (!hasEditPermission) {
            Logger.log(`[updateKeyResultProgress] 警告: 用戶 ${authUser.userEmail} (角色: ${authUser.userRole}) 無權更新 KR ${krId}。負責人: ${krOwnerEmail}`);
            return { success: false, message: "您沒有權限更新此關鍵成果。只有負責人、部門主管或管理員可以更新。" };
        }


        // 計算新的進度百分比
        let newProgress = 0;
        if (targetValue !== startValue) {
            newProgress = Math.min(100, Math.max(0, ((newValue - startValue) / (targetValue - startValue)) * 100));
        } else if (newValue >= targetValue && targetValue === 0) { // 處理目標為0的情況 (例如完成某項任務)
            newProgress = 100;
        } else if (newValue === targetValue) { // 目標值與起始值相同，且達到目標值
            newProgress = 100;
        }
        Logger.log(`[updateKeyResultProgress] KR ${krId} 新進度計算: ${newProgress.toFixed(2)}%`);

        // 獲取需要更新的列索引
        const currentValueColIndex = krHeaders.indexOf('CurrentValue') + 1;
        const progressColIndex = krHeaders.indexOf('Progress') + 1;
        const lastUpdatedColIndex = krHeaders.indexOf('LastUpdated') + 1;
        const lastUpdatedByEmailColIndex = krHeaders.indexOf('LastUpdatedByEmail') + 1;
        const commentColIndex = krHeaders.indexOf('Comment') + 1;

        // 更新 Key_Results 工作表中的數據
        krSheet.getRange(rowIndexToUpdate, currentValueColIndex).setValue(newValue);
        krSheet.getRange(rowIndexToUpdate, progressColIndex).setValue(newProgress.toFixed(2)); // 保留兩位小數
        krSheet.getRange(rowIndexToUpdate, lastUpdatedColIndex).setValue(new Date().toLocaleString());
        krSheet.getRange(rowIndexToUpdate, lastUpdatedByEmailColIndex).setValue(authUser.userEmail);
        krSheet.getRange(rowIndexToUpdate, commentColIndex).setValue(comment);
        Logger.log(`[updateKeyResultProgress] KR ${krId} 主表數據更新成功。`);

        // 記錄進度更新到 Key_Results_Updates 工作表
        krUpdateSheet.appendRow([
            Utilities.getUuid(), // 生成唯一的 UpdateID
            krId,
            authUser.userEmail,
            oldKrValue,
            newValue,
            comment,
            new Date().toLocaleString()
        ]);
        Logger.log(`[updateKeyResultProgress] KR ${krId} 更新記錄成功添加到歷史表。`);

        // --- 更新父級 Objective 的進度 ---
        // 在更新 KR 後，重新計算其父級 Objective 的進度
        if (parentObjectiveId) {
            calculateObjectiveProgressAndScore(parentObjectiveId);
        }

        return { success: true, message: `關鍵成果 ${krId} 進度已更新為 ${newValue} (${newProgress.toFixed(0)}%)` };

    } catch (error) {
        Logger.log(`[updateKeyResultProgress] 錯誤: 更新 KR 進度失敗: ${error.message}`);
        return { success: false, message: `更新失敗: ${error.message}` };
    }
}

/**
 * 輔助函數：計算並更新 Objective 的進度和得分
 * @param {string} objectiveId - 要計算的 Objective ID
 */
function calculateObjectiveProgressAndScore(objectiveId) {
    Logger.log(`[calculateObjectiveProgressAndScore] 開始計算 Objective: ${objectiveId} 的進度。`);
    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const krSheet = spreadsheet.getSheetByName(SHEET_NAMES.KEY_RESULTS);
        const krHeaders = getSheetHeaders(SHEET_NAMES.KEY_RESULTS);
        const allKRs = getSheetData(SHEET_NAMES.KEY_RESULTS).map(row => rowToObject(krHeaders, row));

        const objSheets = [
            { name: SHEET_NAMES.COMPANY_OBJECTIVES, type: 'Company', headers: getSheetHeaders(SHEET_NAMES.COMPANY_OBJECTIVES) },
            { name: SHEET_NAMES.MY_OBJECTIVES, type: 'My', headers: getSheetHeaders(SHEET_NAMES.MY_OBJECTIVES) },
            { name: SHEET_NAMES.DEPARTMENT_OBJECTIVES, type: 'Department', headers: getSheetHeaders(SHEET_NAMES.DEPARTMENT_OBJECTIVES) }
        ];

        let objectiveFound = false;
        for (const objSheetInfo of objSheets) {
            const currentObjSheet = spreadsheet.getSheetByName(objSheetInfo.name);
            const objDataRange = currentObjSheet.getDataRange();
            const objValues = objDataRange.getValues();

            for (let i = 1; i < objValues.length; i++) {
                const obj = rowToObject(objSheetInfo.headers, objValues[i]);
                if (obj.ID === objectiveId) {
                    const relatedKRs = (obj.KRs ? String(obj.KRs).split(',').map(id => id.trim()).filter(id => id !== '') : [])
                                       .map(krId => allKRs.find(kr => kr.ID === krId)).filter(Boolean);

                    let totalProgress = 0;
                    let totalWeight = 0; // 如果有 KR 權重，可以在這裡計算

                    relatedKRs.forEach(kr => {
                        totalProgress += (kr.Progress ? parseFloat(kr.Progress) : 0);
                        // totalWeight += (kr.Weight ? parseFloat(kr.Weight) : 1); // 如果有權重
                    });

                    const newObjectiveProgress = relatedKRs.length > 0 ? (totalProgress / relatedKRs.length) : 0;
                    // const newObjectiveProgress = totalWeight > 0 ? (totalProgress / totalWeight) : 0; // 如果有權重

                    const progressColIndex = objSheetInfo.headers.indexOf('Progress') + 1;
                    const statusColIndex = objSheetInfo.headers.indexOf('Status') + 1; // 假設狀態欄位

                    // 更新 Objective 的進度
                    currentObjSheet.getRange(i + 1, progressColIndex).setValue(newObjectiveProgress.toFixed(2));
                    
                    // 根據進度自動更新 Objective 狀態 (簡化邏輯)
                    let newStatus = obj.Status; // 保持原有狀態，除非達到特定條件
                    if (newObjectiveProgress >= 100) {
                        newStatus = 'Achieved';
                    } else if (newObjectiveProgress < 50 && obj.Status !== 'Draft' && obj.Status !== 'Pending Chairman Approval') {
                        newStatus = 'At Risk'; // 進度低於 50% 且不是草稿/待審批狀態，則設為有風險
                    } else if (newObjectiveProgress >= 50 && newObjectiveProgress < 80 && obj.Status !== 'Draft' && obj.Status !== 'Pending Chairman Approval') {
                        newStatus = 'In Progress'; // 50-80% 設為進行中
                    } else if (newObjectiveProgress >= 80 && newObjectiveProgress < 100 && obj.Status !== 'Draft' && obj.Status !== 'Pending Chairman Approval') {
                        newStatus = 'In Progress'; // 80-100% 設為進行中
                    }
                    currentObjSheet.getRange(i + 1, statusColIndex).setValue(newStatus);


                    Logger.log(`[calculateObjectiveProgressAndScore] Objective ${objectiveId} 進度更新為 ${newObjectiveProgress.toFixed(2)}%，狀態更新為 ${newStatus}。`);
                    objectiveFound = true;
                    break;
                }
            }
            if (objectiveFound) break;
        }

        if (!objectiveFound) {
            Logger.log(`[calculateObjectiveProgressAndScore] 警告: 未找到 ID 為 ${objectiveId} 的 Objective 來計算進度。`);
        }
    } catch (error) {
        Logger.log(`[calculateObjectiveProgressAndScore] 錯誤: 計算 Objective 進度失敗: ${error.message}`);
    }
}


/**
 * 模擬提交 Objective 進行審批的函數。
 * 在實際應用中，此函數會更新 Google Sheet 中的 Objective 狀態，並可能發送通知給董事長。
 * @param {string} sessionToken - 會話令牌
 * @param {string} objectiveId - 要提交審批的 Objective ID
 * @returns {Object} - 包含 success 狀態和 message 的物件。
 */
function submitObjectiveForApproval(sessionToken, objectiveId) {
    Logger.log(`[submitObjectiveForApproval] 函數開始執行。Objective ID: ${objectiveId}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['Chairman', 'OKR_Admin', 'CLevel_Exec', 'Department_Manager', 'Employee']); // 假設這些角色可以提交審批
    } catch (e) {
        Logger.log(`[submitObjectiveForApproval] 認證失敗: ${e.message}`);
        return { success: false, message: e.message };
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const objSheets = [
            { name: SHEET_NAMES.COMPANY_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.COMPANY_OBJECTIVES) },
            { name: SHEET_NAMES.MY_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.MY_OBJECTIVES) },
            { name: SHEET_NAMES.DEPARTMENT_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.DEPARTMENT_OBJECTIVES) }
        ];

        let objectiveFound = false;
        for (const objSheetInfo of objSheets) {
            const currentObjSheet = spreadsheet.getSheetByName(objSheetInfo.name);
            const objDataRange = currentObjSheet.getDataRange();
            const objValues = objDataRange.getValues();

            for (let i = 1; i < objValues.length; i++) {
                const obj = rowToObject(objSheetInfo.headers, objValues[i]);
                if (obj.ID === objectiveId) {
                    const statusColIndex = objSheetInfo.headers.indexOf('Status') + 1;
                    const ownerEmail = obj.OwnerEmail;

                    // 權限檢查：只有 Objective 負責人才能提交審批
                    if (authUser.userEmail !== ownerEmail && !['OKR_Admin', 'Chairman'].includes(authUser.userRole)) {
                        return { success: false, message: "您沒有權限提交此 Objective 進行審批。只有負責人或管理員可以提交。" };
                    }
                    
                    // 檢查 Objective 是否有 KRs
                    if (!obj.KRs || String(obj.KRs).trim() === '') {
                        return { success: false, message: "Objective 必須至少有一個關鍵成果 (KR) 才能提交審批。" };
                    }

                    currentObjSheet.getRange(i + 1, statusColIndex).setValue('Pending Chairman Approval');
                    Logger.log(`[submitObjectiveForApproval] Objective ${objectiveId} 狀態更新為 'Pending Chairman Approval'。`);
                    objectiveFound = true;
                    break;
                }
            }
            if (objectiveFound) break;
        }

        if (!objectiveFound) {
            return { success: false, message: `找不到 ID 為 ${objectiveId} 的 Objective。` };
        }

        return { success: true, message: `Objective ${objectiveId} 已提交審批！` };
    } catch (error) {
        Logger.log(`[submitObjectiveForApproval] 錯誤: ${error.message}`);
        return { success: false, message: `提交審批失敗: ${error.message}` };
    }
}

/**
 * 模擬批准 Objective 的函數 (僅限董事長角色)。
 * @param {string} sessionToken - 會話令牌
 * @param {string} objectiveId - 要批准的 Objective ID
 * @returns {Object} - 包含 success 狀態和 message 的物件。
 */
function approveObjective(sessionToken, objectiveId) {
    Logger.log(`[approveObjective] 函數開始執行。Objective ID: ${objectiveId}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['Chairman', 'OKR_Admin']); // 只有董事長或OKR管理員可以批准
    } catch (e) {
        Logger.log(`[approveObjective] 認證失敗: ${e.message}`);
        return { success: false, message: e.message };
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const objSheets = [
            { name: SHEET_NAMES.COMPANY_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.COMPANY_OBJECTIVES) },
            { name: SHEET_NAMES.MY_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.MY_OBJECTIVES) },
            { name: SHEET_NAMES.DEPARTMENT_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.DEPARTMENT_OBJECTIVES) }
        ];

        let objectiveFound = false;
        for (const objSheetInfo of objSheets) {
            const currentObjSheet = spreadsheet.getSheetByName(objSheetInfo.name);
            const objDataRange = currentObjSheet.getDataRange();
            const objValues = objDataRange.getValues();

            for (let i = 1; i < objValues.length; i++) {
                const obj = rowToObject(objSheetInfo.headers, objValues[i]);
                if (obj.ID === objectiveId) {
                    const statusColIndex = objSheetInfo.headers.indexOf('Status') + 1;
                    currentObjSheet.getRange(i + 1, statusColIndex).setValue('Approved');
                    Logger.log(`[approveObjective] Objective ${objectiveId} 狀態更新為 'Approved'。`);
                    objectiveFound = true;
                    break;
                }
            }
            if (objectiveFound) break;
        }

        if (!objectiveFound) {
            return { success: false, message: `找不到 ID 為 ${objectiveId} 的 Objective。` };
        }

        Logger.log(`[approveObjective] Objective ${objectiveId} 已被 ${authUser.userEmail} 批准`);
        return { success: true, message: `Objective ${objectiveId} 已批准！` };
    } catch (error) {
        Logger.log(`[approveObjective] 錯誤: 批准 Objective 失敗: ${error.message}`);
        return { success: false, message: `批准失敗: ${error.message}` };
    }
}

/**
 * 模擬拒絕 Objective 的函數 (僅限董事長角色)。
 * @param {string} sessionToken - 會話令牌
 * @param {string} objectiveId - 要拒絕的 Objective ID
 * @returns {Object} - 包含 success 狀態和 message 的物件。
 */
function rejectObjective(sessionToken, objectiveId) {
    Logger.log(`[rejectObjective] 函數開始執行。Objective ID: ${objectiveId}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['Chairman', 'OKR_Admin']); // 只有董事長或OKR管理員可以拒絕
    } catch (e) {
        Logger.log(`[rejectObjective] 認證失敗: ${e.message}`);
        return { success: false, message: e.message };
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const objSheets = [
            { name: SHEET_NAMES.COMPANY_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.COMPANY_OBJECTIVES) },
            { name: SHEET_NAMES.MY_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.MY_OBJECTIVES) },
            { name: SHEET_NAMES.DEPARTMENT_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.DEPARTMENT_OBJECTIVES) }
        ];

        let objectiveFound = false;
        for (const objSheetInfo of objSheets) {
            const currentObjSheet = spreadsheet.getSheetByName(objSheetInfo.name);
            const objDataRange = currentObjSheet.getDataRange();
            const objValues = objDataRange.getValues();

            for (let i = 1; i < objValues.length; i++) {
                const obj = rowToObject(objSheetInfo.headers, objValues[i]);
                if (obj.ID === objectiveId) {
                    const statusColIndex = objSheetInfo.headers.indexOf('Status') + 1;
                    currentObjSheet.getRange(i + 1, statusColIndex).setValue('Rejected');
                    Logger.log(`[rejectObjective] Objective ${objectiveId} 狀態更新為 'Rejected'。`);
                    objectiveFound = true;
                    break;
                }
            }
            if (objectiveFound) break;
        }

        if (!objectiveFound) {
            return { success: false, message: `找不到 ID 為 ${objectiveId} 的 Objective。` };
        }

        Logger.log(`[rejectObjective] Objective ${objectiveId} 已被 ${authUser.userEmail} 拒絕`);
        return { success: true, message: `Objective ${objectiveId} 已拒絕。` };
    } catch (error) {
        Logger.log(`[rejectObjective] 錯誤: 拒絕 Objective 失敗: ${error.message}`);
        return { success: false, message: `拒絕失敗: ${error.message}` };
    }
}

/**
 * 實現 Objective 的編輯功能。
 * @param {string} sessionToken - 會話令牌
 * @param {Object} objData - 包含 Objective 資訊的物件。
 * @returns {Object} - 包含 success 狀態和 message 的物件。
 */
function editObjective(sessionToken, objData) {
    Logger.log(`[editObjective] 函數開始執行。objData: ${JSON.stringify(objData)}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['Chairman', 'OKR_Admin', 'CLevel_Exec', 'Department_Manager', 'Employee']); // 假設這些角色可以編輯 Objective
    } catch (e) {
        Logger.log(`[editObjective] 認證失敗: ${e.message}`);
        return { success: false, message: e.message };
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const objSheets = [
            { name: SHEET_NAMES.COMPANY_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.COMPANY_OBJECTIVES) },
            { name: SHEET_NAMES.MY_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.MY_OBJECTIVES) },
            { name: SHEET_NAMES.DEPARTMENT_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.DEPARTMENT_OBJECTIVES) }
        ];

        let objectiveFound = false;
        for (const objSheetInfo of objSheets) {
            const currentObjSheet = spreadsheet.getSheetByName(objSheetInfo.name);
            const objDataRange = currentObjSheet.getDataRange();
            const objValues = objDataRange.getValues();

            for (let i = 1; i < objValues.length; i++) {
                const obj = rowToObject(objSheetInfo.headers, objValues[i]);
                if (obj.ID === objData.ID) {
                    // 權限檢查：只有負責人或特定管理員才能編輯
                    if (authUser.userEmail !== obj.OwnerEmail && !['OKR_Admin', 'Chairman', 'Department_Manager', 'CLevel_Exec'].includes(authUser.userRole)) {
                        return { success: false, message: "您沒有權限編輯此 Objective。" };
                    }

                    // 更新欄位
                    const titleColIndex = objSheetInfo.headers.indexOf('Title') + 1;
                    const descColIndex = objSheetInfo.headers.indexOf('Description') + 1;
                    const ownerColIndex = objSheetInfo.headers.indexOf('OwnerEmail') + 1;
                    const periodColIndex = objSheetInfo.headers.indexOf('Period') + 1;
                    const departmentColIndex = objSheetInfo.headers.indexOf('DepartmentID') + 1; // 注意這裡使用 DepartmentID
                    const parentColIndex = objSheetInfo.headers.indexOf('ParentObjectiveID') + 1;

                    currentObjSheet.getRange(i + 1, titleColIndex).setValue(objData.Title);
                    currentObjSheet.getRange(i + 1, descColIndex).setValue(objData.Description);
                    currentObjSheet.getRange(i + 1, ownerColIndex).setValue(objData.OwnerEmail);
                    currentObjSheet.getRange(i + 1, periodColIndex).setValue(objData.Period);

                    if (departmentColIndex > 0) currentObjSheet.getRange(i + 1, departmentColIndex).setValue(objData.DepartmentID || '');
                    if (parentColIndex > 0) currentObjSheet.getRange(i + 1, parentColIndex).setValue(objData.ParentObjectiveID || '');

                    Logger.log(`[editObjective] 成功編輯 Objective: ${objData.ID}`);
                    objectiveFound = true;
                    break;
                }
            }
            if (objectiveFound) break;
        }

        if (!objectiveFound) {
            return { success: false, message: `找不到 ID 為 ${objData.ID} 的 Objective。` };
        }

        return { success: true, message: `Objective "${objData.Title}" 更新成功！` };
    } catch (error) {
        Logger.log(`[editObjective] 錯誤: ${error.message}`);
        return { success: false, message: `編輯 Objective 失敗: ${error.message}` };
    }
}

/**
 * 實現 Objective 的刪除功能。
 * @param {string} sessionToken - 會話令牌
 * @param {string} objectiveId - 要刪除的 Objective ID。
 * @returns {Object} - 包含 success 狀態和 message 的物件。
 */
function deleteObjective(sessionToken, objectiveId) {
    Logger.log(`[deleteObjective] 函數開始執行。Objective ID: ${objectiveId}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['OKR_Admin', 'Chairman']); // 只有OKR管理員和董事長可以刪除Objective
    } catch (e) {
        Logger.log(`[deleteObjective] 認證失敗: ${e.message}`);
        return { success: false, message: e.message };
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const objSheets = [
            { name: SHEET_NAMES.COMPANY_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.COMPANY_OBJECTIVES) },
            { name: SHEET_NAMES.MY_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.MY_OBJECTIVES) },
            { name: SHEET_NAMES.DEPARTMENT_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.DEPARTMENT_OBJECTIVES) }
        ];

        let objectiveFound = false;
        let objectiveOwnerEmail = '';
        let targetSheet = null;
        let targetRowIndex = -1;
        let objectiveKRs = [];

        // 先找到 Objective 並進行權限檢查
        for (const objSheetInfo of objSheets) {
            const currentObjSheet = spreadsheet.getSheetByName(objSheetInfo.name);
            const objDataRange = currentObjSheet.getDataRange();
            const objValues = objDataRange.getValues();

            for (let i = 1; i < objValues.length; i++) {
                const obj = rowToObject(objSheetInfo.headers, objValues[i]);
                if (obj.ID === objectiveId) {
                    objectiveOwnerEmail = obj.OwnerEmail;
                    objectiveKRs = (obj.KRs ? String(obj.KRs).split(',').map(id => id.trim()).filter(id => id !== '') : []);
                    targetSheet = currentObjSheet;
                    targetRowIndex = i + 1; // Apps Script 行號
                    objectiveFound = true;
                    break;
                }
            }
            if (objectiveFound) break;
        }

        if (!objectiveFound) {
            return { success: false, message: `找不到 ID 為 ${objectiveId} 的 Objective。` };
        }

        // 權限檢查：只有負責人或特定管理員才能刪除
        if (authUser.userEmail !== objectiveOwnerEmail && !['OKR_Admin', 'Chairman'].includes(authUser.userRole)) {
            return { success: false, message: "您沒有權限刪除此 Objective。" };
        }

        // 刪除 Objective 行
        targetSheet.deleteRow(targetRowIndex);
        Logger.log(`[deleteObjective] 成功刪除 Objective: ${objectiveId} 從 ${targetSheet.getName()}`);

        // 級聯刪除其下的 Key Results
        if (objectiveKRs.length > 0) {
            const krSheet = spreadsheet.getSheetByName(SHEET_NAMES.KEY_RESULTS);
            const krHeaders = getSheetHeaders(SHEET_NAMES.KEY_RESULTS);
            const krData = krSheet.getDataRange().getValues();
            
            // 從後往前刪除，避免行號錯亂
            for (let i = krData.length - 1; i >= 1; i--) {
                const kr = rowToObject(krHeaders, krData[i]);
                if (objectiveKRs.includes(kr.ID)) {
                    krSheet.deleteRow(i + 1);
                    Logger.log(`[deleteObjective] 級聯刪除 KR: ${kr.ID}`);
                }
            }
        }

        return { success: true, message: `Objective ${objectiveId} 及其相關 Key Results 已刪除。` };
    } catch (error) {
        Logger.log(`[deleteObjective] 錯誤: ${error.message}`);
        return { success: false, message: `刪除 Objective 失敗: ${error.message}` };
    }
}

/**
 * 實現 Key Result 的編輯功能。
 * @param {string} sessionToken - 會話令牌
 * @param {Object} krData - 包含 Key Result 資訊的物件。
 * @returns {Object} - 包含 success 狀態和 message 的物件。
 */
function editKeyResult(sessionToken, krData) {
    Logger.log(`[editKeyResult] 函數開始執行。krData: ${JSON.stringify(krData)}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['Chairman', 'OKR_Admin', 'CLevel_Exec', 'Department_Manager', 'Employee']); // 假設這些角色可以編輯 KR
    } catch (e) {
        Logger.log(`[editKeyResult] 認證失敗: ${e.message}`);
        return { success: false, message: e.message };
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const krSheet = spreadsheet.getSheetByName(SHEET_NAMES.KEY_RESULTS);
        if (!krSheet) {
            throw new Error(`工作表 "${SHEET_NAMES.KEY_RESULTS}" 不存在。`);
        }

        const krHeaders = getSheetHeaders(SHEET_NAMES.KEY_RESULTS);
        const krDataValues = krSheet.getDataRange().getValues();

        let rowIndexToUpdate = -1;
        let krOwnerEmail = '';

        for (let i = 1; i < krDataValues.length; i++) {
            const kr = rowToObject(krHeaders, krDataValues[i]);
            if (kr.ID === krData.ID) {
                rowIndexToUpdate = i + 1;
                krOwnerEmail = kr.OwnerEmail;
                break;
            }
        }

        if (rowIndexToUpdate === -1) {
            return { success: false, message: `找不到 ID 為 ${krData.ID} 的關鍵成果。` };
        }

        // 權限檢查：只有負責人或特定管理員才能編輯
        if (authUser.userEmail !== krOwnerEmail && !['OKR_Admin', 'Chairman', 'Department_Manager', 'CLevel_Exec'].includes(authUser.userRole)) {
            return { success: false, message: "您沒有權限編輯此關鍵成果。" };
        }

        // 更新欄位
        const descColIndex = krHeaders.indexOf('Description') + 1;
        const ownerColIndex = krHeaders.indexOf('OwnerEmail') + 1;
        const metricTypeColIndex = krHeaders.indexOf('MetricType') + 1;
        const startValueColIndex = krHeaders.indexOf('StartValue') + 1;
        const targetValueColIndex = krHeaders.indexOf('TargetValue') + 1;
        const unitColIndex = krHeaders.indexOf('Unit') + 1;

        krSheet.getRange(rowIndexToUpdate, descColIndex).setValue(krData.Description);
        krSheet.getRange(rowIndexToUpdate, ownerColIndex).setValue(krData.OwnerEmail);
        krSheet.getRange(rowIndexToUpdate, metricTypeColIndex).setValue(krData.MetricType);
        krSheet.getRange(rowIndexToUpdate, startValueColIndex).setValue(krData.StartValue);
        krSheet.getRange(rowIndexToUpdate, targetValueColIndex).setValue(krData.TargetValue);
        krSheet.getRange(rowIndexToUpdate, unitColIndex).setValue(krData.Unit);

        Logger.log(`[editKeyResult] 成功編輯 Key Result: ${krData.ID}`);
        return { success: true, message: `Key Result "${krData.Description}" 更新成功！` };
    } catch (error) {
        Logger.log(`[editKeyResult] 錯誤: ${error.message}`);
        return { success: false, message: `編輯 Key Result 失敗: ${error.message}` };
    }
}

/**
 * 實現 Key Result 的刪除功能。
 * @param {string} sessionToken - 會話令牌
 * @param {string} krId - 要刪除的 Key Result ID。
 * @returns {Object} - 包含 success 狀態和 message 的物件。
 */
function deleteKeyResult(sessionToken, krId) {
    Logger.log(`[deleteKeyResult] 函數開始執行。KR ID: ${krId}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['Chairman', 'OKR_Admin', 'Department_Manager', 'CLevel_Exec', 'Employee']); // 假設這些角色可以刪除 KR
    } catch (e) {
        Logger.log(`[deleteKeyResult] 認證失敗: ${e.message}`);
        return { success: false, message: e.message };
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const krSheet = spreadsheet.getSheetByName(SHEET_NAMES.KEY_RESULTS);
        if (!krSheet) {
            throw new Error(`工作表 "${SHEET_NAMES.KEY_RESULTS}" 不存在。`);
        }

        const krHeaders = getSheetHeaders(SHEET_NAMES.KEY_RESULTS);
        const krData = krSheet.getDataRange().getValues();

        let rowIndexToDelete = -1;
        let krOwnerEmail = '';
        let parentObjectiveId = '';

        for (let i = 1; i < krData.length; i++) {
            const kr = rowToObject(krHeaders, krData[i]);
            if (kr.ID === krId) {
                rowIndexToDelete = i + 1;
                krOwnerEmail = kr.OwnerEmail;
                parentObjectiveId = kr.ObjectiveID;
                break;
            }
        }

        if (rowIndexToDelete === -1) {
            return { success: false, message: `找不到 ID 為 ${krId} 的關鍵成果。` };
        }

        // 權限檢查：只有負責人或特定管理員才能刪除
        if (authUser.userEmail !== krOwnerEmail && !['OKR_Admin', 'Chairman', 'Department_Manager', 'CLevel_Exec'].includes(authUser.userRole)) {
            return { success: false, message: "您沒有權限刪除此關鍵成果。" };
        }

        // 刪除 KR 行
        krSheet.deleteRow(rowIndexToDelete);
        Logger.log(`[deleteKeyResult] 成功刪除 Key Result: ${krId}`);

        // 更新父級 Objective 的 KRs 欄位
        if (parentObjectiveId) {
            const objSheets = [
                { name: SHEET_NAMES.COMPANY_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.COMPANY_OBJECTIVES) },
                { name: SHEET_NAMES.MY_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.MY_OBJECTIVES) },
                { name: SHEET_NAMES.DEPARTMENT_OBJECTIVES, headers: getSheetHeaders(SHEET_NAMES.DEPARTMENT_OBJECTIVES) }
            ];

            for (const objSheetInfo of objSheets) {
                const currentObjSheet = spreadsheet.getSheetByName(objSheetInfo.name);
                const objDataRange = currentObjSheet.getDataRange();
                const objValues = objDataRange.getValues();

                for (let i = 1; i < objValues.length; i++) {
                    const obj = rowToObject(objSheetInfo.headers, objValues[i]);
                    if (obj.ID === parentObjectiveId) {
                        const krColIndex = objSheetInfo.headers.indexOf('KRs') + 1;
                        if (krColIndex > 0) {
                            let currentKRs = obj.KRs ? String(obj.KRs).split(',').map(id => id.trim()).filter(id => id !== '') : [];
                            const updatedKRs = currentKRs.filter(id => id !== krId);
                            currentObjSheet.getRange(i + 1, krColIndex).setValue(updatedKRs.join(','));
                            Logger.log(`[deleteKeyResult] 成功更新父級 Objective (${parentObjectiveId}) 的 KRs 欄位。`);
                            break;
                        }
                    }
                }
            }
            // 重新計算父級 Objective 的進度
            calculateObjectiveProgressAndScore(parentObjectiveId);
        }

        return { success: true, message: `Key Result ${krId} 已刪除。` };
    } catch (error) {
        Logger.log(`[deleteKeyResult] 錯誤: ${error.message}`);
        return { success: false, message: `刪除 Key Result 失敗: ${error.message}` };
    }
}

/**
 * 獲取所有評論。
 * @param {string} sessionToken - 會話令牌
 * @param {string} entityId - 關聯的 Objective 或 Key Result ID。
 * @returns {string} JSON 字串，包含 comments 陣列。
 */
function getComments(sessionToken, entityId) {
    Logger.log(`[getComments] 函數開始執行。Entity ID: ${entityId}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['Chairman', 'OKR_Admin', 'HR_Admin', 'Department_Manager', 'CLevel_Exec', 'Employee']); // 所有用戶都可以查看評論
    } catch (e) {
        Logger.log(`[getComments] 認證失敗: ${e.message}`);
        return JSON.stringify({ error: e.message });
    }

    let result = {
        comments: [],
        error: ''
    };
    try {
        const commentHeaders = getSheetHeaders(SHEET_NAMES.COMMENTS);
        const commentData = getSheetData(SHEET_NAMES.COMMENTS);
        result.comments = commentData
            .map(row => rowToObject(commentHeaders, row))
            .filter(comment => comment.EntityID === entityId)
            .sort((a, b) => new Date(a.Timestamp).getTime() - new Date(b.Timestamp).getTime()); // 按時間排序

        Logger.log(`[getComments] 成功載入 ${result.comments.length} 條評論。`);
    } catch (e) {
        Logger.log(`[getComments] 錯誤: 獲取評論失敗: ${e.message}`);
        result.error += `獲取評論失敗: ${e.message}; `;
    }
    return JSON.stringify(result);
}

/**
 * 新增評論。
 * @param {string} sessionToken - 會話令牌
 * @param {Object} commentData - 包含評論資訊的物件。
 * - EntityID (string) - 關聯的 Objective 或 Key Result ID
 * - CommentText (string)
 * @returns {Object} - 包含 success 狀態和 message 的物件
 */
function addComment(sessionToken, commentData) {
    Logger.log(`[addComment] 函數開始執行。Comment Data: ${JSON.stringify(commentData)}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['Chairman', 'OKR_Admin', 'HR_Admin', 'Department_Manager', 'CLevel_Exec', 'Employee']); // 所有用戶都可以添加評論
    } catch (e) {
        Logger.log(`[addComment] 認證失敗: ${e.message}`);
        return { success: false, message: e.message };
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const commentSheet = spreadsheet.getSheetByName(SHEET_NAMES.COMMENTS);
        if (!commentSheet) {
            throw new Error(`工作表 "${SHEET_NAMES.COMMENTS}" 不存在。`);
        }

        const commentHeaders = getSheetHeaders(SHEET_NAMES.COMMENTS);
        const newCommentRow = [
            Utilities.getUuid(), // CommentID
            commentData.EntityID,
            authUser.userEmail, // 使用認證後的用戶 Email
            commentData.CommentText,
            new Date().toLocaleString() // Timestamp
        ];

        while (newCommentRow.length < commentHeaders.length) {
            newCommentRow.push('');
        }

        commentSheet.appendRow(newCommentRow);
        Logger.log(`[addComment] 成功新增評論到 Entity ID: ${commentData.EntityID}`);
        return { success: true, message: "評論新增成功！" };
    } catch (error) {
        Logger.log(`[addComment] 錯誤: ${error.message}`);
        return { success: false, message: `新增評論失敗: ${error.message}` };
    }
}

/**
 * 獲取所有用戶列表 (用於管理頁面)
 * @param {string} sessionToken - 會話令牌
 * @returns {string} JSON 字串，包含 users 陣列。
 */
function getAllUsers(sessionToken) {
    Logger.log("[getAllUsers] 函數開始執行。");
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['OKR_Admin', 'HR_Admin', 'Chairman']); // 只有管理員和董事長可以查看所有用戶
    } catch (e) {
        Logger.log(`[getAllUsers] 認證失敗: ${e.message}`);
        return JSON.stringify({ error: e.message });
    }

    let result = { users: [], error: '' };
    try {
        const userHeaders = getSheetHeaders(SHEET_NAMES.USERS);
        const userData = getSheetData(SHEET_NAMES.USERS);
        result.users = userData.map(row => rowToObject(userHeaders, row));
        Logger.log(`[getAllUsers] 成功載入 ${result.users.length} 個用戶。`);
    } catch (e) {
        Logger.log(`[getAllUsers] 錯誤: 獲取用戶數據失敗: ${e.message}`);
        result.error += `獲取用戶數據失敗: ${e.message}; `;
    }
    return JSON.stringify(result);
}

/**
 * 創建新用戶 (用於管理頁面)
 * @param {string} sessionToken - 會話令牌
 * @param {Object} userData - 包含用戶資訊的物件 (Email, Password, Role, Department, Name, ID, IsActive)
 * @returns {Object} - 包含 success 狀態和 message 的物件
 */
function createUser(sessionToken, userData) {
    Logger.log(`[createUser] 函數開始執行。UserData: ${JSON.stringify(userData)}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['OKR_Admin', 'HR_Admin', 'Chairman']); // 只有管理員和董事長可以創建用戶
    } catch (e) {
        Logger.log(`[createUser] 認證失敗: ${e.message}`);
        return { success: false, message: e.message };
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const userSheet = spreadsheet.getSheetByName(SHEET_NAMES.USERS);
        if (!userSheet) {
            throw new Error(`工作表 "${SHEET_NAMES.USERS}" 不存在。`);
        }

        const userHeaders = getSheetHeaders(SHEET_NAMES.USERS);
        
        // 檢查 Email 是否已存在
        const existingUsers = getSheetData(SHEET_NAMES.USERS).map(row => rowToObject(userHeaders, row));
        if (existingUsers.some(u => u.Email === userData.Email)) {
            return { success: false, message: `用戶 Email "${userData.Email}" 已存在。` };
        }

        // 哈希密碼
        const hashedPassword = hashString(userData.Password + PASSWORD_SALT);

        const newRow = [
            userData.Email,
            hashedPassword, // 儲存哈希後的密碼
            userData.Role,
            userData.Department,
            userData.Name,
            userData.ID || Utilities.getUuid().substring(0, 4), // 如果沒有提供 ID，生成一個短 ID
            userData.IsActive === true // 確保是布林值
        ];

        // 確保新行數據的長度與標題長度匹配
        while (newRow.length < userHeaders.length) {
            newRow.push('');
        }

        userSheet.appendRow(newRow);
        Logger.log(`[createUser] 成功創建用戶: ${userData.Email}`);
        return { success: true, message: `用戶 "${userData.Email}" 創建成功！` };
    } catch (error) {
        Logger.log(`[createUser] 錯誤: ${error.message}`);
        return { success: false, message: `創建用戶失敗: ${error.message}` };
    }
}

/**
 * 編輯現有用戶 (用於管理頁面)
 * @param {string} sessionToken - 會話令牌
 * @param {Object} userData - 包含用戶資訊的物件 (Email, Role, Department, Name, ID, IsActive, Password - 可選)
 * @returns {Object} - 包含 success 狀態和 message 的物件
 */
function editUser(sessionToken, userData) {
    Logger.log(`[editUser] 函數開始執行。UserData: ${JSON.stringify(userData)}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['OKR_Admin', 'HR_Admin', 'Chairman']); // 只有管理員和董事長可以編輯用戶
    } catch (e) {
        Logger.log(`[editUser] 認證失敗: ${e.message}`);
        return { success: false, message: e.message };
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const userSheet = spreadsheet.getSheetByName(SHEET_NAMES.USERS);
        if (!userSheet) {
            throw new Error(`工作表 "${SHEET_NAMES.USERS}" 不存在。`);
        }

        const userHeaders = getSheetHeaders(SHEET_NAMES.USERS);
        const allUsersData = userSheet.getDataRange().getValues();

        let rowIndexToUpdate = -1;
        let oldPasswordHash = '';

        for (let i = 1; i < allUsersData.length; i++) {
            const user = rowToObject(userHeaders, allUsersData[i]);
            if (user.Email === userData.Email) {
                rowIndexToUpdate = i + 1;
                oldPasswordHash = user.PasswordHash; // 獲取舊的密碼哈希
                break;
            }
        }

        if (rowIndexToUpdate === -1) {
            return { success: false, message: `找不到 Email 為 ${userData.Email} 的用戶。` };
        }

        // 更新欄位
        const emailColIndex = userHeaders.indexOf('Email') + 1;
        const passwordHashColIndex = userHeaders.indexOf('PasswordHash') + 1;
        const roleColIndex = userHeaders.indexOf('Role') + 1;
        const deptColIndex = userHeaders.indexOf('Department') + 1;
        const nameColIndex = userHeaders.indexOf('Name') + 1;
        const idColIndex = userHeaders.indexOf('ID') + 1;
        const isActiveColIndex = userHeaders.indexOf('IsActive') + 1;
        

        userSheet.getRange(rowIndexToUpdate, emailColIndex).setValue(userData.Email); // Email 不可改，但確保寫回
        userSheet.getRange(rowIndexToUpdate, roleColIndex).setValue(userData.Role);
        userSheet.getRange(rowIndexToUpdate, deptColIndex).setValue(userData.Department);
        userSheet.getRange(rowIndexToUpdate, nameColIndex).setValue(userData.Name);
        userSheet.getRange(rowIndexToUpdate, idColIndex).setValue(userData.ID);
        userSheet.getRange(rowIndexToUpdate, isActiveColIndex).setValue(userData.IsActive === true); // 確保是布林值

        // 如果提供了新密碼，則更新哈希
        if (userData.Password && userData.Password.trim() !== '') {
            const newHashedPassword = hashString(userData.Password + PASSWORD_SALT);
            userSheet.getRange(rowIndexToUpdate, passwordHashColIndex).setValue(newHashedPassword);
            Logger.log(`[editUser] 用戶 ${userData.Email} 密碼已更新。`);
        } else {
            // 如果沒有提供新密碼，確保舊的哈希值不被清空
            userSheet.getRange(rowIndexToUpdate, passwordHashColIndex).setValue(oldPasswordHash);
        }

        Logger.log(`[editUser] 成功編輯用戶: ${userData.Email}`);
        return { success: true, message: `用戶 "${userData.Email}" 更新成功！` };
    } catch (error) {
        Logger.log(`[editUser] 錯誤: ${error.message}`);
        return { success: false, message: `編輯用戶失敗: ${error.message}` };
    }
}

/**
 * 刪除用戶 (用於管理頁面)
 * @param {string} sessionToken - 會話令牌
 * @param {string} userEmailToDelete - 要刪除的用戶 Email
 * @returns {Object} - 包含 success 狀態和 message 的物件
 */
function deleteUser(sessionToken, userEmailToDelete) {
    Logger.log(`[deleteUser] 函數開始執行。User Email: ${userEmailToDelete}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['OKR_Admin', 'HR_Admin', 'Chairman']); // 只有管理員和董事長可以刪除用戶
    } catch (e) {
        Logger.log(`[deleteUser] 認證失敗: ${e.message}`);
        return { success: false, message: e.message };
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const userSheet = spreadsheet.getSheetByName(SHEET_NAMES.USERS);
        if (!userSheet) {
            throw new Error(`工作表 "${SHEET_NAMES.USERS}" 不存在。`);
        }

        const userHeaders = getSheetHeaders(SHEET_NAMES.USERS);
        const allUsersData = userSheet.getDataRange().getValues();

        let rowIndexToDelete = -1;
        for (let i = 1; i < allUsersData.length; i++) {
            const user = rowToObject(userHeaders, allUsersData[i]);
            if (user.Email === userEmailToDelete) {
                rowIndexToDelete = i + 1;
                break;
            }
        }

        if (rowIndexToDelete === -1) {
            return { success: false, message: `找不到 Email 為 ${userEmailToDelete} 的用戶。` };
        }

        // 檢查是否嘗試刪除當前登入的用戶
        if (userEmailToDelete === authUser.userEmail) {
            return { success: false, message: "不能刪除當前登入的用戶。" };
        }

        userSheet.deleteRow(rowIndexToDelete);
        Logger.log(`[deleteUser] 成功刪除用戶: ${userEmailToDelete}`);
        return { success: true, message: `用戶 "${userEmailToDelete}" 已刪除。` };
    } catch (error) {
        Logger.log(`[deleteUser] 錯誤: ${error.message}`);
        return { success: false, message: `刪除用戶失敗: ${error.message}` };
    }
}

/**
 * 獲取所有部門列表 (用於管理頁面)
 * @param {string} sessionToken - 會話令牌
 * @returns {string} JSON 字串，包含 departments 陣列。
 */
function getAllDepartments(sessionToken) {
    Logger.log("[getAllDepartments] 函數開始執行。");
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['OKR_Admin', 'HR_Admin', 'Chairman', 'Department_Manager', 'CLevel_Exec']); // 這些角色可以查看所有部門
    } catch (e) {
        Logger.log(`[getAllDepartments] 認證失敗: ${e.message}`);
        return JSON.stringify({ error: e.message });
    }

    let result = { departments: [], error: '' };
    try {
        const deptHeaders = getSheetHeaders(SHEET_NAMES.DEPARTMENTS);
        const deptData = getSheetData(SHEET_NAMES.DEPARTMENTS);
        result.departments = deptData.map(row => rowToObject(deptHeaders, row));
        Logger.log(`[getAllDepartments] 成功載入 ${result.departments.length} 個部門。`);
    } catch (e) {
        Logger.log(`[getAllDepartments] 錯誤: 獲取部門數據失敗: ${e.message}`);
        result.error += `獲取部門數據失敗: ${e.message}; `;
    }
    return JSON.stringify(result);
}

/**
 * 創建新部門 (用於管理頁面)
 * @param {string} sessionToken - 會話令牌
 * @param {Object} deptData - 包含部門資訊的物件 (ID, Name)
 * @returns {Object} - 包含 success 狀態和 message 的物件
 */
function createDepartment(sessionToken, deptData) {
    Logger.log(`[createDepartment] 函數開始執行。DeptData: ${JSON.stringify(deptData)}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['OKR_Admin', 'HR_Admin', 'Chairman']); // 只有管理員和董事長可以創建部門
    } catch (e) {
        Logger.log(`[createDepartment] 認證失敗: ${e.message}`);
        return { success: false, message: e.message };
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const deptSheet = spreadsheet.getSheetByName(SHEET_NAMES.DEPARTMENTS);
        if (!deptSheet) {
            throw new Error(`工作表 "${SHEET_NAMES.DEPARTMENTS}" 不存在。`);
        }

        const deptHeaders = getSheetHeaders(SHEET_NAMES.DEPARTMENTS);
        
        // 檢查 ID 或 Name 是否已存在
        const existingDepts = getSheetData(SHEET_NAMES.DEPARTMENTS).map(row => rowToObject(deptHeaders, row));
        if (existingDepts.some(d => d.ID === deptData.ID)) {
            return { success: false, message: `部門 ID "${deptData.ID}" 已存在。` };
        }
        if (existingDepts.some(d => d.Name === deptData.Name)) {
            return { success: false, message: `部門名稱 "${deptData.Name}" 已存在。` };
        }

        const newRow = [
            deptData.ID,
            deptData.Name
        ];

        while (newRow.length < deptHeaders.length) {
            newRow.push('');
        }

        deptSheet.appendRow(newRow);
        Logger.log(`[createDepartment] 成功創建部門: ${deptData.Name}`);
        return { success: true, message: `部門 "${deptData.Name}" 創建成功！` };
    } catch (error) {
        Logger.log(`[createDepartment] 錯誤: ${error.message}`);
        return { success: false, message: `創建部門失敗: ${error.message}` };
    }
}

/**
 * 編輯現有部門 (用於管理頁面)
 * @param {string} sessionToken - 會話令牌
 * @param {Object} deptData - 包含部門資訊的物件 (ID, Name)
 * @returns {Object} - 包含 success 狀態和 message 的物件
 */
function editDepartment(sessionToken, deptData) {
    Logger.log(`[editDepartment] 函數開始執行。DeptData: ${JSON.stringify(deptData)}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['OKR_Admin', 'HR_Admin', 'Chairman']); // 只有管理員和董事長可以編輯部門
    } catch (e) {
        Logger.log(`[editDepartment] 認證失敗: ${e.message}`);
        return { success: false, message: e.message };
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const deptSheet = spreadsheet.getSheetByName(SHEET_NAMES.DEPARTMENTS);
        if (!deptSheet) {
            throw new Error(`工作表 "${SHEET_NAMES.DEPARTMENTS}" 不存在。`);
        }

        const deptHeaders = getSheetHeaders(SHEET_NAMES.DEPARTMENTS);
        const allDeptsData = deptSheet.getDataRange().getValues();

        let rowIndexToUpdate = -1;
        for (let i = 1; i < allDeptsData.length; i++) {
            const dept = rowToObject(deptHeaders, allDeptsData[i]);
            if (dept.ID === deptData.ID) { // 根據 ID 查找
                rowIndexToUpdate = i + 1;
                break;
            }
        }

        if (rowIndexToUpdate === -1) {
            return { success: false, message: `找不到 ID 為 ${deptData.ID} 的部門。` };
        }

        // 檢查新名稱是否與其他部門衝突 (如果名稱有變動)
        const existingDepts = getSheetData(SHEET_NAMES.DEPARTMENTS).map(row => rowToObject(deptHeaders, row));
        if (existingDepts.some(d => d.Name === deptData.Name && d.ID !== deptData.ID)) {
            return { success: false, message: `部門名稱 "${deptData.Name}" 已被其他部門使用。` };
        }

        const nameColIndex = deptHeaders.indexOf('Name') + 1;
        deptSheet.getRange(rowIndexToUpdate, nameColIndex).setValue(deptData.Name);

        Logger.log(`[editDepartment] 成功編輯部門: ${deptData.ID}`);
        return { success: true, message: `部門 "${deptData.Name}" 更新成功！` };
    } catch (error) {
        Logger.log(`[editDepartment] 錯誤: ${error.message}`);
        return { success: false, message: `編輯部門失敗: ${error.message}` };
    }
}

/**
 * 刪除部門 (用於管理頁面)
 * @param {string} sessionToken - 會話令牌
 * @param {string} deptIdToDelete - 要刪除的部門 ID
 * @returns {Object} - 包含 success 狀態和 message 的物件
 */
function deleteDepartment(sessionToken, deptIdToDelete) {
    Logger.log(`[deleteDepartment] 函數開始執行。Dept ID: ${deptIdToDelete}`);
    let authUser;
    try {
        authUser = authenticateAndAuthorize(sessionToken, ['OKR_Admin', 'HR_Admin', 'Chairman']); // 只有管理員和董事長可以刪除部門
    } catch (e) {
        Logger.log(`[deleteDepartment] 認證失敗: ${e.message}`);
        return { success: false, message: e.message };
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const deptSheet = spreadsheet.getSheetByName(SHEET_NAMES.DEPARTMENTS);
        if (!deptSheet) {
            throw new Error(`工作表 "${SHEET_NAMES.DEPARTMENTS}" 不存在。`);
        }

        const deptHeaders = getSheetHeaders(SHEET_NAMES.DEPARTMENTS);
        const allDeptsData = deptSheet.getDataRange().getValues();

        let rowIndexToDelete = -1;
        for (let i = 1; i < allDeptsData.length; i++) {
            const dept = rowToObject(deptHeaders, allDeptsData[i]);
            if (dept.ID === deptIdToDelete) {
                rowIndexToDelete = i + 1;
                break;
            }
        }

        if (rowIndexToDelete === -1) {
            return { success: false, message: `找不到 ID 為 ${deptIdToDelete} 的部門。` };
        }

        // 檢查是否有 Objective 關聯到此部門
        const deptObjectives = getSheetData(SHEET_NAMES.DEPARTMENT_OBJECTIVES).map(row => rowToObject(getSheetHeaders(SHEET_NAMES.DEPARTMENT_OBJECTIVES), row));
        if (deptObjectives.some(obj => obj.DepartmentID === deptIdToDelete)) {
            return { success: false, message: `無法刪除部門 "${deptIdToDelete}"，因為有 Objective 關聯到此部門。請先刪除或修改相關 Objective。` };
        }
        
        // 檢查是否有用戶關聯到此部門
        const usersInDept = getSheetData(SHEET_NAMES.USERS).map(row => rowToObject(getSheetHeaders(SHEET_NAMES.USERS), row));
        if (usersInDept.some(user => user.Department === deptIdToDelete)) { // 注意这里是 Department Name, 而不是 Department ID
             return { success: false, message: `無法刪除部門 "${deptIdToDelete}"，因為有用戶關聯到此部門。請先修改或刪除相關用戶。` };
        }


        deptSheet.deleteRow(rowIndexToDelete);
        Logger.log(`[deleteDepartment] 成功刪除部門: ${deptIdToDelete}`);
        return { success: true, message: `部門 "${deptIdToDelete}" 已刪除。` };
    } catch (error) {
        Logger.log(`[deleteDepartment] 錯誤: ${error.message}`);
        return { success: false, message: `刪除部門失敗: ${error.message}` };
    }
}


/**
 * 輔助函數：用於在 HTML 模板中引入其他 HTML 檔案的內容。
 * @param {string} filename - 要引入的 HTML 檔案名稱 (不含 .html 副檔名)
 * @returns {string} - 引入的 HTML 內容字串
 */
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}


// Code.gs (臨時函數 - 用於生成密碼哈希)
function generatePasswordHashForAdmin() {
    const password = 'your_admin_password_here'; // <-- 將此處替換為您想要設定的明文密碼
    const hashedPassword = hashString(password + PASSWORD_SALT);
    Logger.log('您的哈希密碼是: ' + hashedPassword);
    // 執行後，請將日誌中的哈希值複製到 Users 表的 PasswordHash 欄位
}
