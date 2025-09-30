# 自動化任務建立工具

這個專案透過 Node.js 腳本（`index.js`）將 Word 任務檔案轉成後台建立任務／知識卡所需的欄位資料，並搭配 Puppeteer 自動化操作頁面。

## 環境需求
- Node.js 18 或以上版本
- npm（已隨 Node.js 安裝提供）

## 安裝
```bash
npm install
```

## 使用方式
以下指令會讀取 `./docs/第一題.docx`，並同時產生任務與知識卡資料：
```bash
npm run start-task -- --word ./docs/第一題.docx --mode=task+card
```

> 💡 `第一題.docx` 只是作為範例的題目順序命名，請依實際需求將檔名改成「第一題」、「第二題」等對應標題，或使用其他你希望的檔名。

## 執行情境

### 1️⃣ 單檔執行

若只需處理單一 Word 檔案，維持上述指令即可：

```bash
npm run start-task -- --word ./docs/第一題.docx --mode=task+card
```

這樣程式只會處理 `./docs/第一題.docx` 一個檔案。

### 2️⃣ 多檔情境

若需要連續處理多個 Word 檔，有兩種方式可以選擇：

#### 🔹 方法 A：一個一個跑

手動更換檔名後重複執行指令，例如：

```bash
npm run start-task -- --word ./docs/第二題.docx --mode=task+card
```

這種方式最直觀，也最不容易出錯。

#### 🔹 方法 B：用 `--loop` 自動連續處理

若想讓程式自動連續處理，在指令中加入 `--loop`：

```bash
npm run start-task -- --loop --mode=task+card
```

執行後終端機會提示：

```
📄 請輸入下一個 Word 檔路徑（空白直接結束）：
```

依序輸入要處理的檔案路徑：

1. `./docs/第一題.docx`
2. `./docs/第二題.docx`

每個檔案處理完畢後會再詢問下一個檔案路徑，若沒有更多檔案，直接按 Enter 即可結束流程。

### 常用參數
- `--word <路徑>`：指定 Word (`.docx`) 檔案位置。
- `--mode <task|card|task+card>`：決定要輸出任務、知識卡或兩者皆輸出，預設為 `task+card`。
- `--loop`：加上此旗標可在執行完畢後重複流程。
- `--close`：加上此旗標會在流程結束時關閉瀏覽器視窗。
- `--url <網址>`：自動開啟指定的「建立任務」頁面。

若需調整欄位對應，可編輯 `mapping.json`。

## 附註
- `docs/` 目錄中的 Word 範例可作為測試輸入。
- 腳本會在終端機提示必要的輸入，請依指示操作。
