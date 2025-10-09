# 時空學園自動上題系統(Boringman擴充版)

> **智能化教育內容管理助手** - 從 Word 文檔到網頁表單的全自動化解決方案

**🏫 時空學園官網：** [Time Warrior Academy](https://sites.google.com/aiv.com.tw/time-warrior-academy/)  
**🔗 基於原作者專案擴充：** [olyueen/automatic](https://github.com/olyueen/automatic?fbclid=IwY2xjawNUQzBleHRuA2FlbQIxMABicmlkETFCYUhJb3hRSVFHQnFzbVdNAR4WS-XchZQ2nA2fIedC7OL90Ny66PW0H2i1RDWTas-Gsf2UIsm6Wc7pZHjZDw_aem_DvnsGraqDHmuwaRwAJ_9QA)

這個工具使用 Node.js + Puppeteer 技術，能夠：
- **智能解析** Word 文檔中的教育任務和知識卡片內容
- **自動圖片處理** 解析 docx 中的圖片並自動上傳（新功能！）
- **自動化填寫** 網頁表單（支援複雜的下拉選單和富文本編輯器）
- **批量處理** 最多 12 張知識卡片
- **循環模式** 連續處理多個 Word 檔案
- **智能匹配** 自動處理學科、等級、課綱等下拉選單映射

---

## 系統需求

- **Node.js** 18+ 版本
- **作業系統** Windows / macOS（支援 Chrome/Edge/Safari 瀏覽器）
- **Word 文檔** (.docx 格式)
- **網路連線** (需訪問目標教育平台)

---

## 快速開始

### 1️⃣ 環境準備

**安裝 Node.js**
```bash
# 檢查是否已安裝
node -v
npm -v

# 若未安裝，請到官網下載 LTS 版本
start https://nodejs.org/
```

**安裝專案依賴**
```bash
npm install
```

### 2️⃣ 準備 Word 文檔

將要處理的 `.docx` 檔案放到 `docs/` 資料夾中：
```
docs/
├── 第一題.docx
├── 第二題.docx
└── 第三題.docx
```

### 3️⃣ 基本使用

```bash
# 最簡單的用法（推薦）- 完整流程 + 自動圖片
node index.js -w ./docs/第一題.docx

# 只處理任務頁面
node index.js -w ./docs/第一題.docx --mode task

# 只處理卡片頁面
node index.js -w ./docs/第一題.docx --mode card

# 停用自動圖片功能
node index.js -w ./docs/第一題.docx --no-image

# 循環處理多個檔案
node index.js -w ./docs/第一題.docx --loop
```

---

## 詳細操作流程

### **任務頁面自動化**

程式會自動啟動 Chrome 瀏覽器並導航到教育平台，然後：

1. **自動填寫文字欄位**：
   - ✅ 任務名稱
   - ✅ 線索文本（支援富文本編輯器）
   - ✅ 解答說明（自動清理格式和括號註解）
   - ✅ 整體鷹架提示 1-3（自動截取25字限制）

2. **智能處理下拉選單**：
   - **主要學科**：自動匹配（國語、英語、數學、自然、社會等）
   - **等級**：智能識別（如「八年級A」→「08A」）
   - **世紀**：自動轉換（如「21世紀」、「第20世紀」）
   - **課綱**：課綱欄位使用真實打字模擬，避免自動完成干擾
   - **地點**：地區匹配（台灣、中國、美國等）

3. **特殊處理機制**：
   ```javascript
   // 課綱欄位使用真實打字，每字符間隔5ms
   await page.type(selector, value, { delay: 5 });
   
   // 自動清理括號註解和系統標記
   text = stripBracketNotesBlock(cleanTextBase(text));
   ```

### **知識卡片批量處理**

程式會依序處理最多 **12 張卡片**（編號 1-1 到 3-4）：

#### **每張卡片的處理步驟**：

1. **自動填寫基本資訊**：
   - ✅ 卡片名稱
   - ✅ 文字內容（自動清理和格式化）
   - ✅ 學科下拉選單匹配
   - ✅ 類別下拉選單（人物、事件、時間、地點等）

2. **自動圖片處理**：
   - **智能解析**：自動從 docx 檔案中提取圖片
   - **圖片關聯**：按順序將圖片與卡片關聯（1-1 對應第1張圖片，1-2 對應第2張圖片...）
   - **自動儲存**：圖片儲存到 `temp_images/` 資料夾
   - **自動上傳**：在上傳圖片對話框中自動選擇對應圖片

3. **智能圖片搜尋功能**：
   ```
   👉 請在卡片頁按【搜尋現有圖片】打開視窗，準備好後按 Enter...
   ```
   - 自動將**卡片名稱**填入搜尋框
   - 讓您快速查找現有圖片資源
   - 避免重複上傳相同圖片

4. **自動化圖片上傳**：
   ```
   卡片 1-1 包含來自 docx 的圖片，將嘗試自動上傳
   請在卡片頁按【上傳圖片】打開視窗，準備好後按 Enter 繼續（將自動上傳圖片）...
   ```
   - 圖片名稱：自動填入卡片名稱
   - 圖片描述：自動填入卡片文字內容
   - **自動選擇檔案**：程式會自動選擇對應的圖片檔案

#### **進度顯示範例**：
```
📂 解析 Word: ./docs/第一題.docx
📸 從 docx 檔案中提取到 12 張圖片
🖼️  卡片 1-1 (祕密投票) 已關聯圖片
...
🎴 開始填寫卡片 1...
✅ cardTitle: 祕密投票
✅ cardDescription: 讓投票人能自由表達投票意志的原則。
✅ cardSubjectId (dropdown): 公民
🖼️  卡片 1-1 包含來自 docx 的圖片，將嘗試自動上傳
📸 嘗試自動上傳從 docx 解析的圖片...
✅ 圖片已自動上傳
🎉 卡片 1 填寫完成！
```

---

## 命令列參數詳解

| 參數 | 別名 | 類型 | 預設值 | 說明 |
|------|------|------|--------|------|
| `--word` | `-w` | string | - | **必須**，指定要處理的 Word 檔案路徑 |
| `--mode` | - | string | `task+card` | 處理模式：`task`、`card`、`task+card` |
| `--auto-image` | `-i` | boolean | `true` | 自動解析並上傳 docx 中的圖片 |
| `--no-image` | - | boolean | `false` | 停用自動圖片功能 |
| `--loop` | - | boolean | `false` | 啟用循環模式，連續處理多個檔案 |
| `--close` | - | boolean | `false` | 處理完成後自動關閉瀏覽器 |
| `--url` | - | string | - | 自訂起始頁面 URL |

### **循環模式詳解**

啟用 `--loop` 後，第一個檔案處理完成會提示：
```
  請輸入下一個 Word 檔路徑（直接 Enter 結束）：
```

**省時技巧**：直接將 Word 檔案從檔案總管**拖曳**到終端機視窗，路徑會自動輸入！

### **使用情境範例**

```bash
# 情境1：最簡單用法（推薦）- 完整流程 + 自動圖片
node index.js -w docs/第一題.docx

# 情境2：批量處理多個檔案（含自動圖片）
node index.js -w docs/第一題.docx --loop

# 情境3：只更新任務內容，跳過卡片
node index.js -w docs/第一題.docx --mode task

# 情境4：完整流程但停用自動圖片（需手動上傳）
node index.js -w docs/第一題.docx --no-image

# 情境5：指定起始頁面並自動關閉
node index.js -w docs/第一題.docx --url "https://adl.edu.tw/twa-admin/edit/missionEdit" --close

# 情境6：使用舊版完整語法（仍然支援）
node index.js --word docs/第一題.docx --mode task+card --auto-image
```

### **自動圖片功能說明**

#### **圖片處理流程**：
1. **解析階段**：程式自動從 docx 檔案中提取所有圖片
2. **關聯階段**：按順序將圖片與卡片關聯（第1張圖片→卡片1-1，第2張圖片→卡片1-2...）
3. **儲存階段**：圖片暫存到 `temp_images/` 資料夾
4. **上傳階段**：在卡片編輯時自動選擇對應圖片上傳

#### **執行時的提示訊息**：
```bash
📸 從 docx 檔案中提取到 12 張圖片
🖼️  卡片 1-1 (祕密投票) 已關聯圖片
🖼️  卡片 1-2 (公開投票) 已關聯圖片
...
📸 嘗試自動上傳從 docx 解析的圖片...
💾 圖片已儲存至: temp_images/card_xxx.png
✅ 圖片已自動上傳
```

---

## 配置檔案說明

###  `mapping.json` - 欄位映射配置

這個檔案定義了 Word 文檔內容如何對應到網頁表單欄位：

```json
{
  "task": {
    "任務名稱": { "name": "name", "label": ["任務名稱"] },
    "主要學科": { "name": "mainSubjectId", "label": ["主要學科"] },
    "課綱": { "name": "syllabus", "label": ["課綱"] },
    "線索文本": { "name": "description", "label": ["線索文本","線索","內容"], "rich": true }
  },
  "card": {
    "卡片名稱": { "name": "cardTitle", "label": ["卡片名稱"] },
    "文字內容": { "name": "cardDescription", "label": ["文字內容"], "rich": true }
  },
  "dropdown": {
    "subject": { "國語":"國語", "數學":"數學", "自然":"自然" },
    "level": { "八年級A":"08A", "九年級B":"09B" },
    "century": { "21世紀":"21世紀", "第20世紀":"20世紀" }
  }
}
```

#### **配置選項說明**：
- `name`: 對應的 HTML 表單欄位名稱
- `label`: 可能的標籤文字（支援多個別名）
- `rich`: 是否為富文本編輯器
- `dropdown`: 下拉選單的值映射關係

---

###  **注意事項**
1. **標題必須完全匹配** `mapping.json` 中定義的 key
2. **卡片順序** 程式會按照出現順序處理，最多12張
3. **內容清理** 程式會自動移除：
   - 頁碼數字
   - 括號註解 `（限25字以內）`
   - 系統標記 `【系統】`、`知識卡分類` 等

---

##  技術特色

###  **智能文字處理**
```javascript
// 自動清理括號註解
function stripBracketNotesBlock(s) {
  // 移除整行括號說明
  if (/^[（(][^）)]+[）)]$/.test(trimmed)) continue;
  
  // 移除段落開頭的括號部分
  line = line.replace(/^[（(][^）)]+[）)]\s*/, "");
}

// 智能截斷處理
function pruneAfterMarkers(text) {
  const markers = ["【系統】", "知識卡分類", "第一組知識卡"];
  // 在系統標記處截斷
}
```

###  **精確的下拉選單處理**
```javascript
// 支援完全匹配和部分匹配
const targetOption = options.find(opt => {
  const text = (opt.textContent || "").trim();
  return text === val; // 完全匹配優先
});

// 智能ID映射
const fieldMappings = {
  '主要學科': 'mui-component-select-mainSubjectId',
  '等級': 'mui-component-select-level',
  '世紀': 'mui-component-select-centuryId'
};
```

###  **強健的錯誤處理**
- **多層備援**：name 屬性 → 標籤查找 → 富文本處理
- **等待機制**：智能等待頁面元素載入
- **值驗證**：確認輸入值是否正確設置

---

##  效能與限制

###  **支援功能**
- [x] Word 文檔解析（.docx 格式）
- [x] 自動表單填寫（文字輸入框、富文本編輯器）
- [x] 下拉選單智能匹配
- [x] 批量卡片處理（最多12張）
- [x] 圖片搜尋和上傳輔助
- [x] 循環處理多檔案
- [x] 瀏覽器會話保持（保持登入狀態）

###  **已知限制**
- [ ] 下拉選單需要 **手動最終確認**
- [ ] 圖片實際上傳需要 **手動操作**
- [ ] 正確答案勾選需要 **手動處理**
- [ ] 線索關鍵詞和鷹架提示需要 **手動填寫**
- [ ] 依賴特定網站的 HTML 結構（MUI 組件）

###  **適用場景**
- ✅ 大量重複性教育內容錄入
- ✅ 標準化題目格式處理
- ✅ 批量知識卡片建立
- ✅ 減少手動複製貼上工作

---

##  故障排除

### 常見問題與解決方案

#### ❌ **「找不到 mapping.json」**
```bash
❌ 找不到 mapping.json 或 JSON 損壞
```
**解決方案**：確保 `mapping.json` 檔案存在於專案根目錄，且格式正確。

#### ❌ **「下拉選單找不到選項」**
```bash
⚠️ mainSubjectId 找不到選項 "自然科學"，可用選項: 國語,英語,數學,自然,社會
```
**解決方案**：檢查 `mapping.json` 中的 `dropdown.subject` 映射，或手動選擇正確選項。

#### ❌ **「課綱輸入失敗」**
```bash
❌ [課綱真實打字] 失敗: 找不到可見的課綱輸入框
```
**解決方案**：
1. 確保已正確導航到任務建立頁面
2. 確認頁面已完全載入
3. 檢查是否有多個課綱輸入框干擾

#### ❌ **「瀏覽器啟動失敗」**
```bash
Error: Could not find Chrome or Edge
```
**解決方案**：
1. 確保已安裝 Chrome 或 Edge 瀏覽器
2. 檢查瀏覽器安裝路徑是否標準
3. 嘗試手動指定瀏覽器路徑

#### 🆕 **「自動圖片功能問題」**
```bash
📷 docx 檔案中未找到圖片，將使用手動上傳模式
```
**可能原因**：
1. docx 檔案中沒有嵌入圖片
2. 圖片格式不支援（僅支援 PNG、JPG）
3. 圖片位於文檔的特殊位置

**解決方案**：
- 使用 `--no-image` 停用自動圖片功能
- 確認 docx 檔案中的圖片是直接插入的（非連結）
- 手動上傳模式仍可正常使用

###  **偵錯模式**

啟用詳細日誌輸出：
```bash
DEBUG=* node index.js --word ./docs/第一題.docx
```

---

##  更新日誌

### v2.0.0 (最新版本) 🆕
-  **新增自動圖片功能**：自動從 docx 解析圖片並上傳
-  **簡化 CLI 語法**：支援 `-w` 簡寫，默認啟用自動圖片
-  **暫存圖片管理**：圖片儲存到 `temp_images/` 資料夾
-  **智能圖片關聯**：按順序自動關聯圖片與卡片
-  **增強錯誤處理**：改善圖片上傳失敗時的提示

### v1.2.0
-  新增課綱欄位真實打字模擬
-  改善下拉選單匹配邏輯
-  修復卡片順序處理問題
-  優化等待和重試機制

### v1.1.0  
-  新增智能圖片搜尋功能
-  改善文字清理算法
-  支援更多下拉選單映射

### v1.0.0
-  初始版本發布
-  基本 Word 解析功能
-  自動化表單填寫

---

##  貢獻指南

歡迎提交問題報告和功能建議！

### 提交 Issue
請包含以下資訊：
- 使用的 Word 檔案格式
- 完整的錯誤訊息
- 操作系統和瀏覽器版本
- 復現步驟

### 開發環境設置
```bash
git clone <repository-url>
cd automatic-enhanced
npm install
node index.js --word ./docs/測試檔案.docx
```

---

##  致謝

### 教育平台
- **[時空學園 (Time Warrior Academy)](https://sites.google.com/aiv.com.tw/time-warrior-academy/)** - 提供優質的教育平台和學習環境

### 原作者
- **[olyueen](https://github.com/olyueen)** - [automatic](https://github.com/olyueen/automatic) 專案的原始作者，提供了基礎的自動化填寫功能

### 開源專案
感謝以下開源專案：
- [Puppeteer](https://github.com/puppeteer/puppeteer) - 瀏覽器自動化
- [Mammoth.js](https://github.com/mwilliamson/mammoth.js) - Word 文檔解析  
- [Yargs](https://github.com/yargs/yargs) - 命令列參數處理

---

** 開始自動化您的教育內容管理之旅！**

如有任何問題，請隨時開啟 Issue 或聯繫開發團隊。
