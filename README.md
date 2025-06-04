# Kdan OKR 目標管理平台

這是一個基於 Google Apps Script 和 Google Sheets 開發的輕量級 OKR (Objectives and Key Results) 目標管理平台。它旨在解決企業在目標設定、追蹤、對齊和協作過程中遇到的數據分散、進度不即時和報告生成複雜等問題。

## 專案概述

本系統將 OKR 方法論數位化，提供一個中心化的平台來管理公司、部門和個人的目標與關鍵成果。透過直觀的介面和自動化的數據處理，提升組織的透明度、協作效率和目標達成率。

## 主要功能

本系統目前已實作以下核心功能：

* **自定義用戶登入與會話管理：** 系統不再依賴 Google 原生登入，而是透過自定義的用戶名 (Email) 和密碼進行身份驗證，並使用會話令牌管理登入狀態。

* **角色基礎的權限控制 (RBAC)：** 根據用戶在 `Users` 工作表中定義的角色（例如 `Chairman`, `OKR_Admin`, `Employee` 等），動態顯示側邊欄導航連結和操作按鈕，並在後端進行嚴格的權限檢查。

* **分層級 OKR 數據顯示：**

    * **儀表板：** 提供公司整體 OKR 概覽、部門 OKR 健康度、個人待辦事項。

    * **我的 OKR：** 顯示當前用戶負責的個人目標及其關鍵成果。

    * **部門 OKR：** 顯示部門目標及其關鍵成果。

    * **公司 OKR：** 顯示公司層級目標及其關鍵成果，並可查看承接部門。

* **邏輯化 OKR 編號自動生成：**

    * **公司 Objective：** `C-YYYYQQ-O-XX` (例如 `C-2025Q2-O-01`)

    * **公司 Key Result：** `C-YYYYQQ-O-XX-KR-YY`

    * **部門 Objective：** `D-YYYYQQ-[部門代碼]-O-XX` (例如 `D-2025Q2-PROD-O-01`)

    * **部門 Key Result：** `D-YYYYQQ-[部門代碼]-O-XX-KR-YY`

    * **個人 Objective：** `I-YYYYQQ-[用戶ID]-O-XX` (例如 `I-2025Q2-u1-O-01`)

    * **個人 Key Result：** `I-YYYYQQ-[用戶ID]-O-XX-KR-YY`

* **Key Result 進度更新：** 允許負責人更新 Key Result 的當前值，系統自動計算進度百分比，並匯總到父級 Objective。

* **Objective/Key Result 新增：**

    * 透過模態框表單，新增公司、部門、個人層級的 Objective。

    * 新增 Key Result，並自動關聯到父級 Objective。

    * 新增時可指派負責人、選擇所屬部門和父級目標。

* **用戶管理 (Admin/HR_Admin/Chairman 權限)：** 在「管理」頁面提供介面，允許新增、編輯、刪除用戶，設定其 Email、密碼、角色、部門、姓名和短 ID。

* **部門管理 (Admin/HR_Admin/Chairman 權限)：** 在「管理」頁面提供介面，允許新增、編輯、刪除部門。

* **Objective 審批流程：**

    * Objective 負責人可將 `Draft` 狀態的 Objective **提交審批**。

    * `Chairman` (董事長) 和 `OKR_Admin` 角色可在「管理」頁面查看所有待審批的 Objective，並進行**批准**或**拒絕**操作。

* **評論系統：** 在 Objective/Key Result 詳細頁面，用戶可以添加評論，所有評論將被記錄。

* **Toast 通知：** 提供友好的成功、失敗和提示訊息。

* **基於 Google Sheets 的數據儲存：** 所有 OKR 數據、用戶資訊、評論等都儲存在一個 Google Sheets 試算表中。

## 設定與安裝

要部署和運行此 OKR 系統，請按照以下步驟操作：

### 步驟 1：準備 Google Sheets 試算表

1.  **創建新的 Google Sheets 檔案：**

    * 前往 [Google Sheets](https://sheets.google.com/)。

    * 點擊 `+` 建立一個空白試算表。

    * 將其命名為 `Kdan OKR Data` (或您喜歡的名稱)。

2.  **複製試算表 ID：**

    * 從瀏覽器網址列複製試算表的 ID。它位於 `/d/` 和 `/edit` 之間。

    * 例如：`https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit#gid=0`

    * 您的 ID 就是 `YOUR_SPREADSHEET_ID`。請將此 ID 記下來，稍後會用到。

3.  **創建必要的工作表 (Sheet)：**

    * 在試算表底部，點擊 `+` 號來新增工作表。請確保工作表名稱**完全精確**地匹配以下列表（包括大小寫和底線）：

        * `Users`

        * `Departments`

        * `Company_Objectives`

        * `My_Objectives`

        * `Department_Objectives`

        * `Key_Results`

        * `Key_Results_Updates`

        * `OKR_Periods`

        * `Comments`

4.  **設定每個工作表的欄位 (第一行標題)：**

    * **`Users` 工作表：**
        `Email`, `PasswordHash`, `Role`, `Department`, `Name`, `ID`, `IsActive`

        * **注意：** `PasswordHash` 欄位用於儲存密碼的哈希值。`ID` 欄位必須是唯一的短代碼 (例如 `u1`, `u2`)，用於生成個人 OKR ID。`Department` 欄位應填寫部門的 `Name` (例如 `產品部`)。

    * **`Departments` 工作表：**
        `ID`, `Name`

        * **範例：** `PROD`, `產品部`；`MKT`, `行銷部`。`ID` 欄位是部門的簡短代碼，用於生成部門 OKR ID。

    * **`Company_Objectives` 工作表：**
        `ID`, `Title`, `Description`, `OwnerEmail`, `Period`, `Progress`, `Status`, `KRs`

    * **`My_Objectives` 工作表：**
        `ID`, `Title`, `Description`, `OwnerEmail`, `Period`, `Progress`, `Status`, `ParentObjectiveID`, `KRs`

    * **`Department_Objectives` 工作表：**
        `ID`, `Title`, `Description`, `OwnerEmail`, `Period`, `Progress`, `Status`, `DepartmentID`, `ParentObjectiveID`, `KRs`

        * **注意：** `DepartmentID` 欄位應填寫 `Departments` 工作表中的 `ID` (例如 `PROD`)。

    * **`Key_Results` 工作表：**
        `ID`, `ObjectiveID`, `Description`, `OwnerEmail`, `MetricType`, `StartValue`, `TargetValue`, `CurrentValue`, `Unit`, `Progress`, `ConfidenceLevel`, `LastUpdated`, `LastUpdatedByEmail`, `Comment`

    * **`Key_Results_Updates` 工作表：**
        `UpdateID`, `KeyResultID`, `UpdatedByEmail`, `OldValue`, `NewValue`, `Comment`, `Timestamp`

    * **`OKR_Periods` 工作表：**
        `ID`, `Name`, `StartDate`, `EndDate`, `IsActive`

        * **範例：** `2025Q1`, `2025 第一季度`, `2025/01/01`, `2025/03/31`, `FALSE`

    * **`Comments` 工作表：**
        `ID`, `EntityID`, `CommenterEmail`, `CommentText`, `Timestamp`

5.  **設定第一個管理員帳號 (非常重要！)**

    * 在您的 Apps Script 專案中，打開 `Code.gs` 檔案。

    * 找到 `const PASSWORD_SALT = 'YOUR_SUPER_SECURE_RANDOM_SALT_STRING_HERE_PLEASE_CHANGE_ME_NOW_AND_KEEP_IT_SECRET';` 這行。

    * **請務必將 `'YOUR_SUPER_SECURE_RANDOM_SALT_STRING_HERE_PLEASE_CHANGE_ME_NOW_AND_KEEP_IT_SECRET'` 替換為一個您自己設定的、隨機的、足夠長且複雜的字串。** (例如：`MyKdanOKR@SecretSalt#2025!`) **請務必記住這個鹽值，因為它不能更改。**

    * 在 `Code.gs` 檔案中，新增一個**臨時函數**來生成密碼哈希：

        ```javascript
        // Code.gs (臨時函數 - 用於生成密碼哈希)
        function generatePasswordHashForAdmin() {
            const password = 'your_admin_password_here'; // <-- 將此處替換為您想要設定的明文密碼
            const hashedPassword = hashString(password + PASSWORD_SALT);
            Logger.log('您的哈希密碼是: ' + hashedPassword);
        }
        ```

    * **執行此臨時函數：** 在 Apps Script 編輯器中，選擇 `generatePasswordHashForAdmin` 函數並點擊執行。

    * **複製哈希值：** 查看 Apps Script 的「執行作業」日誌，複製生成的哈希值。

    * **在 `Users` 工作表中填入第一個管理員帳號：** 使用您剛剛生成的哈希值，在 `Users` 工作表中填入一個管理員帳號（例如：`admin@yourcompany.com`，`PasswordHash` 貼上哈希值，`Role` 設為 `OKR_Admin` 或 `Chairman`，並填寫其他欄位）。

    * **刪除臨時函數：** 在成功設定第一個管理員帳號並測試登入後，請**務必**將 `generatePasswordHashForAdmin` 這個臨時函數從 `Code.gs` 中刪除。

6.  **填入其他範例數據：**

    * 在每個工作表中填入一些範例數據，確保它們符合新的編號邏輯和欄位要求。

### 步驟 2：創建 Apps Script 專案並複製程式碼

1.  **創建新的 Apps Script 專案：**

    * 前往 [Google Apps Script](https://script.google.com/)。

    * 點擊 `+ 新增專案`。

2.  **複製後端程式碼 (`Code.gs`)：**

    * 將以下提供的**完整後端程式碼**複製到您的 Apps Script 專案中的 `Code.gs` 檔案。

    * **最重要：** 在 `Code.gs` 檔案的頂部，找到 `const SPREADSHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE';` 這行，並將 `'YOUR_GOOGLE_SHEET_ID_HERE'` 替換為您在**步驟 1.2** 中複製的 Google Sheets 試算表 ID。

3.  **複製前端程式碼 (`index.html`)：**

    * 在 Apps Script 編輯器中，點擊左側檔案列表旁邊的 `+` 號，選擇 `HTML 檔案`。

    * 將檔案命名為 `index.html`。

    * 將以下提供的**完整前端程式碼**複製到 `index.html` 檔案中。

4.  **儲存專案：** 點擊 Apps Script 編輯器上方的儲存圖示 (磁碟片)。

### 步驟 3：部署 Web 應用程式

1.  **部署 Web 應用程式：**

    * 在 Apps Script 編輯器中，點擊頂部菜單欄的 `部署 (Deploy) > 新增部署 (New deployment)`。

    * 選擇類型為 `Web 應用程式 (Web app)`。

    * **執行身份 (Execute as):** 選擇 **`我 (My self)`**。

        * **重要：** 為了讓 Apps Script 能夠讀寫您的 Google Sheet (`SPREADSHEET_ID`)，Web 應用程式需要以部署者的身份運行，這樣它才擁有對該試算表的權限。用戶登入後的權限控制將在應用程式內部實現。

    * **存取權 (Who has access):** 選擇 **`任何人 (Anyone)`**。

        * **重要：** 選擇 `任何人` 才能讓未登入的用戶看到登入頁面。系統內部的權限控制將在登入後根據用戶角色進行。

    * 點擊 `部署 (Deploy)`。

2.  **授權 Apps Script：**

    * 第一次部署時，會彈出授權對話框。點擊「授權存取」。

    * 選擇您的 Google 帳號。

    * 點擊「允許」以授予 Apps Script 訪問 Google Drive 和其他服務的權限。

3.  **獲取 Web 應用程式 URL：**

    * 部署成功後，您將獲得一個「網頁應用程式 URL」。複製這個 URL。

### 步驟 4：測試與驗證

1.  **訪問 Web 應用程式：** 在瀏覽器中打開您複製的 Web 應用程式 URL。

2.  **登入：** 系統將首先顯示登入頁面。使用您在 `Users` 工作表中設定的 Email 和密碼進行登入。

3.  **檢查功能：**

    * **儀表板：** 查看數據是否正確載入。

    * **側邊欄：** 檢查導航連結是否根據您的角色正確顯示/隱藏。

    * **新增 Objective/KR：** 嘗試新增不同類型的 Objective 和 Key Result，並檢查 Google Sheets 中是否正確生成了帶有新編號的數據。

    * **更新進度：** 嘗試更新 Key Result 進度，並檢查 Google Sheets 中的數據和歷史更新是否正確。

    * **審批：** 如果您是 `Chairman` 或 `OKR_Admin` 角色，嘗試在「管理」頁面審批 Objective。

    * **評論：** 嘗試在詳情頁面添加評論。

    * **用戶管理 (Admin/HR_Admin/Chairman 權限)：** 訪問「管理」頁面，嘗試新增、編輯、刪除用戶。

    * **部門管理 (Admin/HR_Admin/Chairman 權限)：** 訪問「管理」頁面，嘗試新增、編輯、刪除部門。

4.  **查看 Apps Script 日誌：**

    * 在 Apps Script 編輯器中，點擊左側導航欄的「執行作業 (Executions)」圖標，或從菜單欄選擇 `執行 > 日誌 (Executions > Logs)`。

    * 檢查是否有任何錯誤訊息，或 `Logger.log` 輸出的調試信息。

## 效能優化與注意事項

* **Google Sheets 讀寫優化：** 後端程式碼已盡量採用批量讀寫 (`getDataRange().getValues()`, `appendRow()`)，以減少 Apps Script API 呼叫次數，提升與 Google Sheets 交互的效率。

* **記憶體中數據處理：** 數據一旦從 Google Sheets 讀取到 Apps Script 記憶體中，後續的篩選、關聯和計算都在記憶體中進行，減少了對 Sheets 的頻繁訪問。

* **會話管理：** 使用 `PropertiesService` 儲存會話令牌，避免每次請求都重新驗證用戶名和密碼，減少 Sheets 讀寫次數。

* **前端數據防禦性：** 前端程式碼中增加了大量的空值檢查和 `Array.isArray()` 判斷，確保即使後端返回的數據結構不完全符合預期，前端也不會輕易崩潰。

* **響應式 UI：** 採用 Tailwind CSS 實現響應式佈局，確保在不同設備上都有良好的用戶體驗。

* **安全警告：** 自定義的用戶名/密碼登入系統在 Google Sheets 中儲存用戶憑證存在安全風險。**不建議用於處理敏感數據或在正式生產環境中大規模部署。** 建議未來考慮整合 Firebase Authentication 等專業的身份驗證服務。

**請您將上述的 README 內容保存下來，並按照其中的步驟進行操作。**

如果您在部署、設定或測試過程中遇到任何問題，請隨時提出，我會盡力協助您。
