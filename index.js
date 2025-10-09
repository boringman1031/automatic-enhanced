// index.js
import fs from "fs";
import path from "path";
import mammoth from "mammoth";
import puppeteer from "puppeteer";
import yargs from "yargs";
import { hideBin } from "yargs/helpers";
import readline from "readline";

// ---------- CLI ----------
const argv = yargs(hideBin(process.argv))
  .option("mode", { type: "string", default: "task+card", describe: "task | card | task+card" })
  .option("word", { type: "string", alias: "w", describe: "Word 檔路徑（.docx）" })
  .option("loop", { type: "boolean", default: false })
  .option("close", { type: "boolean", default: false })
  .option("url", { type: "string", describe: "（可選）自動前往『建立任務』頁 URL" })
  .option("auto-image", { type: "boolean", default: true, alias: "i", describe: "自動解析並上傳 docx 中的圖片" })
  .option("no-image", { type: "boolean", default: false, describe: "停用自動圖片功能" })
  .help().argv;

console.log("⚙️ argv =", argv);

// ---------- mapping ----------
let mapping = {};
try {
  mapping = JSON.parse(fs.readFileSync("./mapping.json", "utf8"));
  if (!mapping.task) mapping.task = {};
  if (!mapping.card) mapping.card = {};
  console.log("📑 已載入 mapping.json");
} catch (e) {
  console.error("❌ 找不到 mapping.json 或 JSON 損壞");
  process.exit(1);
}

// ---------- 小工具 ----------
function ask(prompt = "") {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((resolve) => rl.question(prompt, (ans) => { rl.close(); resolve(ans); }));
}

function cssEscape(s) {
  return String(s).replace(/["\\]/g, "\\$&");
}

function preview(s, n = 36) {
  const t = (s || "").replace(/\s+/g, " ").trim();
  return t.length > n ? t.slice(0, n) + "..." : t;
}

// 僅做輕量清理：保留 $[[...]]，砍空白/頁碼/系統段
function pruneAfterMarkers(text) {
  const markers = ["【系統】", "知識卡分類", "第一組知識卡", "第二組知識卡", "第三組知識卡"];
  let cut = text.length;
  for (const m of markers) {
    const i = text.indexOf(m);
    if (i !== -1) cut = Math.min(cut, i);
  }
  return text.slice(0, cut);
}

function cleanTextBase(text) {
  let t = String(text);
  t = t.split("\n").filter(l => !/^\d+$/.test(l.trim())).join("\n");
  t = pruneAfterMarkers(t);
  return t.trim();
}

// 只對 description/answerDescription/cardDescription 用的清理
function stripBracketNotesBlock(s) {
  if (!s) return s;
  const lines = String(s).split(/\r?\n/);
  const out = [];
  let firstKeptSeen = false;

  // 若將來發現其它被捲進來的小標題，可擴充這個 RE
  const neighborHeadingsRE = /^(多媒體補充資訊|因材網或外部資訊|卡片圖片|編輯者註解|卡片線索說明)\s*$/;

  for (let line of lines) {
    const trimmed = line.trim();

    if (/^[（(][^）)]+[）)]$/.test(trimmed)) continue;

    if (!firstKeptSeen && /^[（(][^）)]+[）)]/.test(trimmed)) {
      line = line.replace(/^[（(][^）)]+[）)]\s*/, "");
    }

    if (neighborHeadingsRE.test(trimmed)) continue;

    if (line.trim().length && !firstKeptSeen) firstKeptSeen = true;
    out.push(line);
  }
  return out.join("\n").trim();
}

function firstLine(text) {
  return (String(text).split("\n").find(l => l.trim().length) || "").trim();
}

// ---------- 圖片解析 ----------
async function extractCardImages(wordPath) {
  try {
    console.log("📸 開始解析 docx 檔案中的圖片...");
    
    const images = [];
    await mammoth.convertToHtml({
      path: wordPath
    }, {
      convertImage: mammoth.images.imgElement(function(image) {
        return image.read('base64').then(function(imageBuffer) {
          images.push({
            contentType: image.contentType,
            base64Data: imageBuffer,
            altText: image.altText || ''
          });
          
          return {
            src: `data:${image.contentType};base64,${imageBuffer.substring(0, 50)}...`
          };
        });
      })
    });
    
    console.log(`✅ 成功解析 ${images.length} 張圖片`);
    return images;
    
  } catch (error) {
    console.error("❌ 解析圖片時發生錯誤:", error.message);
    return [];
  }
}

// ---------- 圖片檔案處理 ----------
async function saveImageToTemp(imageData, cardIndex) {
  try {
    // 建立暫存資料夾
    const tempDir = path.resolve('./temp_images');
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }
    
    // 根據圖片類型決定副檔名
    const ext = imageData.contentType === 'image/jpeg' ? 'jpg' : 'png';
    const fileName = `card_${cardIndex + 1}.${ext}`;
    const filePath = path.join(tempDir, fileName);
    
    const buffer = Buffer.from(imageData.base64Data, 'base64');
    fs.writeFileSync(filePath, buffer);
    
    console.log(`💾 圖片已儲存至: ${filePath}`);
    return filePath;
    
  } catch (error) {
    console.error(`❌ 儲存圖片失敗 (卡片 ${cardIndex + 1}):`, error.message);
    return null;
  }
}

// ---------- 解析 Word ----------
// 任務欄位（用 mapping.task 的 key 抓段落）
async function parseTaskSections(value) {
  const titles = new Set(Object.keys(mapping.task || {}));
  const resultRaw = {};
  let curKey = null;

  for (const raw of value.split("\n")) {
    const line = raw.trim();
    if (!line) continue;
    if (titles.has(line)) {
      curKey = line;
      resultRaw[curKey] = "";
    } else if (curKey) {
      resultRaw[curKey] += (resultRaw[curKey] ? "\n" : "") + line;
    }
  }

  const taskData = {};
  for (const [title, cfg] of Object.entries(mapping.task)) {
    const name = typeof cfg === "string" ? cfg : cfg.name;
    if (!name || !resultRaw[title]) continue;
    let v = cleanTextBase(resultRaw[title]);

    // 任務的單行欄位
    if (["name", "syllabus", "area", "centuryId", "mainSubjectId", "level"].includes(name)) {
      
      if (name === "level") {
        // 等級欄位：在整個文本中找到類似 "8A"、"01A" 的模式
        const levelMatch = v.match(/\b(\d{1,2}[A-D])\b/);
        if (levelMatch) {
          v = levelMatch[1];
        } else {
          v = firstLine(v);
        }
      } else if (name === "centuryId") {
        // 世紀欄位：提取類似 "21世紀"、"21 世紀" 的值
        const centuryMatch = v.match(/(\d{1,2}\s*世紀|西元前|公元前)/);
        if (centuryMatch) {
          v = centuryMatch[1].replace(/\s+/g, ""); 
        } else {
          v = firstLine(v);
        }
      } else {
        
        v = firstLine(v);
        
        if (name === "mainSubjectId") {
          // 主要學科：提取學科名稱（去除數字編號等）
          v = v.replace(/^\d+\s*/, "").replace(/^其他學科\s*/, "").trim();
        } else if (name === "area") {
          // 地點：直接使用，但清理格式
          v = v.replace(/^\d+\s*/, "").trim();
        }
      }
    }
    if (name.startsWith("missionHintSet")) {
      v = v.split("\n").map(s => s.replace(/（限25字[^）]*）/g, "").trim()).filter(Boolean)[0] || "";
    }

    if (name === "description" || name === "answerDescription") {
      v = stripBracketNotesBlock(v);
    }

    taskData[name] = v.trim();
  }
  return taskData;
}

// 卡片欄位（按 label 循序切卡）
function parseCards(value) {
  const labels = mapping.__cardLabels || ["卡片名稱", "文字內容", "學科", "類別", "課綱"];
  const labelToKey = {
    "卡片名稱": "cardTitle",
    "文字內容": "cardDescription",
    "學科": "cardSubjectId",
    "類別": "cardType",
    "課綱": "syllabus" 
  };

  const lines = value.split("\n").map(s => s.trim()).filter(Boolean);
  const cards = [];
  let cur = null;
  let curField = null;

  for (const line of lines) {
    if (labels.includes(line)) {
      
      if (line === "卡片名稱") {
        if (cur && (cur.cardTitle || cur.cardDescription)) {
          cards.push(cur);
        }
        cur = {};
      }
      curField = labelToKey[line] || null;
      continue;
    }
    if (curField && cur) {
      cur[curField] = (cur[curField] ? cur[curField] + "\n" : "") + line;
    }
  }
  if (cur && (cur.cardTitle || cur.cardDescription)) {
    cards.push(cur);
  }

  // 正規化：名稱取第一行；內容清理；學科/類別正規化
  for (const c of cards) {
    if (c.cardTitle) c.cardTitle = firstLine(c.cardTitle);
    if (c.cardDescription) {
      let cleaned = stripBracketNotesBlock(cleanTextBase(c.cardDescription));
      cleaned = cleaned.split("\n").map(s => s.trim()).filter(Boolean)[0] || ""; 
      c.cardDescription = cleaned;
    }
    // 學科和類別欄位正規化處理
    if (c.cardSubjectId) {
      let subject = firstLine(c.cardSubjectId);

      subject = subject.replace(/^圖片.*$/i, "").trim();
      subject = subject.replace(/（.*）/g, "").trim(); 
      subject = subject.replace(/\s+/g, "");
      c.cardSubjectId = subject;
    }
    if (c.cardType) {
      let type = firstLine(c.cardType);
      type = type.replace(/^圖片.*$/i, "").trim();
      type = type.replace(/（.*）/g, "").trim();
      type = type.replace(/\s+/g, "");
      c.cardType = type;
    }
    if (c.syllabus) {
      c.syllabus = firstLine(c.syllabus);
    }
  }

  return cards.slice(0, 12);
}

async function parseWord(wordPath) {
  const { value } = await mammoth.extractRawText({ path: wordPath });
  const taskData = await parseTaskSections(value);
  const cardDataList = parseCards(value);
  
  if (taskData.syllabus && cardDataList.length > 0) {
    for (const card of cardDataList) {
      if (!card.syllabus) {  
        card.syllabus = taskData.syllabus;
      }
    }
  }
  
  // 解析 docx 中的圖片並與卡片關聯（預設啟用，可用 --no-image 停用）
  const shouldExtractImages = argv['auto-image'] && !argv['no-image'];
  
  if (shouldExtractImages) {
    const cardImages = await extractCardImages(wordPath);
    if (cardImages.length > 0) {
      console.log(`📸 從 docx 檔案中提取到 ${cardImages.length} 張圖片`);
      
      // 將圖片與卡片關聯（按順序對應）
      for (let i = 0; i < Math.min(cardDataList.length, cardImages.length); i++) {
        cardDataList[i].imageData = cardImages[i];
        const tag = `${Math.floor(i/4) + 1}-${(i % 4) + 1}`;
        console.log(`🖼️  卡片 ${tag} (${cardDataList[i].cardTitle || '未命名'}) 已關聯圖片`);
      }
    } else {
      console.log("📷 docx 檔案中未找到圖片，將使用手動上傳模式");
    }
  } else {
    console.log("📷 自動圖片功能已停用（使用 --auto-image 啟用，或移除 --no-image）");
  }
  
  return { taskData, cardDataList };
}

// ---------- DOM 寫入 ----------
async function setByNameNative(page, name, value) {
  const selector = `[name="${cssEscape(name)}"]`;
  
  // 如果是課綱欄位，使用 Puppeteer 的真實打字模擬
  if (name === "syllabus") {
    console.log(`🔍開始處理: ${preview(value)}`);
    
    try {
      // 先找到所有課綱輸入框，選擇最後一個可見的（通常是用戶正在編輯的）
      const targetInput = await page.evaluate(() => {
        const inputs = document.querySelectorAll('[name="syllabus"]');
        for (let i = inputs.length - 1; i >= 0; i--) {
          const input = inputs[i];
          if (input.offsetParent !== null && input.type === 'text') {
            return {
              found: true,
              id: input.id,
              visible: true,
              index: i,
              total: inputs.length
            };
          }
        }
        return { found: false, total: inputs.length };
      });
      
      if (!targetInput.found) {
        console.log(`❌找不到可見的課綱輸入框`);
        return "NF";
      }
      
      console.log(`🎯選擇輸入框 ${targetInput.index + 1}/${targetInput.total}, ID: ${targetInput.id}`);
      
      // 使用特定的 ID 選擇器，手動轉義特殊字符
      const escapedId = targetInput.id.replace(/:/g, '\\:');
      const specificSelector = `#${escapedId}`;
      
      // 確認元素存在並可見
      await page.waitForSelector(specificSelector, { visible: true, timeout: 3000 });
      
      // 點擊輸入框聚焦
      await page.click(specificSelector);
      await new Promise(resolve => setTimeout(resolve, 200));
      
      // 清空現有內容 (Ctrl+A + Delete)
      await page.keyboard.down('Control');
      await page.keyboard.press('KeyA');
      await page.keyboard.up('Control');
      await page.keyboard.press('Delete');
   
      await new Promise(resolve => setTimeout(resolve, 300));
      
      await page.type(specificSelector, value, { delay: 5 }); 
      
      await new Promise(resolve => setTimeout(resolve, 500));
      
      // 點擊輸入框外的區域來失焦，而不是按 Escape
      await page.evaluate((sel) => {
        const input = document.querySelector(sel);
        if (input) {
          input.blur(); 
        }
      }, specificSelector);
      
      await new Promise(resolve => setTimeout(resolve, 300));
      
      const finalValue = await page.evaluate((sel) => {
        const input = document.querySelector(sel);
        return input ? input.value : "NO_INPUT";
      }, specificSelector);
      
      console.log(`🔍 最終結果: ${preview(finalValue)}`);
      
      if (finalValue === value) {
        return "OK";
      } else if (finalValue.includes(value)) {
        console.log(`⚠️包含期望值但有額外內容，嘗試重新設置`);
        
        await page.evaluate((sel, val) => {
          const input = document.querySelector(sel);
          if (input) {
            input.focus();
            input.value = val;
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
            input.blur();
          }
        }, specificSelector, value);
        
        return "OK";
      } else {
        console.log(`⚠️值不匹配，期望: ${preview(value)}, 實際: ${preview(finalValue)}`);
        return "VALUE_MISMATCH";
      }
      
    } catch (error) {
      console.log(`❌失敗: ${error.message}`);
    }
  }
  
  // 標準處理方式
  return page.evaluate((sel, val) => {
    const el = document.querySelector(sel);
    if (!el) return "NF";
    
    const isTA = el.tagName.toLowerCase() === "textarea";
    const proto = isTA ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    const desc = Object.getOwnPropertyDescriptor(proto, "value");
    if (!desc || !desc.set) return "NOSETTER";
    
    el.focus();
    desc.set.call(el, val);
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    el.blur();
    
    return "OK";
  }, selector, value);
}

async function typeIntoInputByLabel(page, labelText, value) {
  return page.evaluate((label, val) => {
    const lab = Array.from(document.querySelectorAll("label,div,span,p,h6,h5"))
      .find(el => (el.textContent || "").trim() === label);
    const row = lab ? lab.closest(".MuiGrid-root, .MuiStack-root, div") : null;
    const input = row ? row.querySelector('input[type="text"], input:not([type]), textarea') : null;
    if (!input) return "NF";

    const isTA = input.tagName.toLowerCase() === "textarea";
    const proto = isTA ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    const desc = Object.getOwnPropertyDescriptor(proto, "value");
    if (!desc || !desc.set) return "NOSETTER";
    input.focus();
    desc.set.call(input, val);
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
    input.blur();
    return "OK";
  }, labelText, value);
}

async function typeIntoRichByLabel(page, labelText, value) {
  return page.evaluate((label, val) => {
    const lab = Array.from(document.querySelectorAll("label,div,span,p,h6,h5"))
      .find(el => (el.textContent || "").trim() === label);
    const row = lab ? lab.closest(".MuiGrid-root, .MuiStack-root, div") : null;
    let box =
      (row && row.querySelector("textarea")) ||
      (row && row.querySelector('[contenteditable="true"]')) ||
      (row && row.querySelector(".ql-editor")) ||
      (row && row.querySelector('[role="textbox"]'));
    if (!box) return "NF";

    if (box.tagName && box.tagName.toLowerCase() === "textarea") {
      const desc = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, "value");
      box.focus();
      desc.set.call(box, val);
      box.dispatchEvent(new Event("input", { bubbles: true }));
      box.dispatchEvent(new Event("change", { bubbles: true }));
      box.blur();
      return "OK";
    }

    box.focus();
    const sel = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(box);
    sel.removeAllRanges();
    sel.addRange(range);
    document.execCommand("delete", false, null);
    document.execCommand("insertText", false, val);
    box.dispatchEvent(new InputEvent("input", { bubbles: true, data: val }));
    box.dispatchEvent(new Event("change", { bubbles: true }));
    return "OK";
  }, labelText, value);
}

// ---------- 下拉選單處理 ---------- 
async function setByDropdown(page, labelText, value) {
  try {
    const result = await page.evaluate(async (label, val) => {
      // 精確的下拉選單ID映射
      const fieldMappings = {
        '主要學科': 'mui-component-select-mainSubjectId',
        '等級': 'mui-component-select-level', 
        '世紀': 'mui-component-select-centuryId',
        '課綱': 'mui-component-select-syllabus',
        '地點': 'mui-component-select-area'
      };
      
      // 直接通過精確的ID查找下拉選單
      if (!fieldMappings[label]) {
        return "UNSUPPORTED_FIELD";
      }
      
      const dropdown = document.getElementById(fieldMappings[label]);
      if (!dropdown) {
        return "ELEMENT_NOT_FOUND:" + fieldMappings[label];
      }
      
      if (dropdown.getAttribute('role') !== 'combobox') {
        return "NOT_COMBOBOX";
      }
      
      console.log('點擊下拉選單:', fieldMappings[label]);
      
      dropdown.focus();
      dropdown.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
      dropdown.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      dropdown.click();
      
      return "CLICKED:" + fieldMappings[label];
    }, labelText, value);

    if (!result.startsWith("CLICKED:")) {
      return result;
    }

    // 等待選項列表出現，使用更簡單的等待條件
    try {
      await page.waitForFunction(
        () => {
          const listboxes = document.querySelectorAll('ul[role="listbox"]');
          return listboxes.length > 0;
        },
        { timeout: 3000 }
      );
    } catch (e) {
      console.log('等待超時，但仍嘗試查找選項列表...');
    }

    await new Promise(resolve => setTimeout(resolve, 500));

    // 處理選項選擇
    const selectResult = await page.evaluate(async (val) => {
      // 獲取所有可能的選項列表，取最後一個（最新打開的）
      const listboxes = document.querySelectorAll('ul[role="listbox"]');
      
      if (listboxes.length === 0) {
        return "NOLISTBOX";
      }
      
      const listbox = listboxes[listboxes.length - 1];

      const options = Array.from(listbox.querySelectorAll('li.MuiMenuItem-root'));
      
      if (options.length === 0) {
        return "NOOPTIONS";
      }
      
      const targetOption = options.find(opt => {
        const text = (opt.textContent || "").trim();
        return text === val;
      });

      if (targetOption) {
        targetOption.click();
        await new Promise(resolve => setTimeout(resolve, 500));
        return "OK";
      } else {
        
        const partialMatch = options.find(opt => {
          const text = (opt.textContent || "").toLowerCase().trim();
          const searchVal = val.toLowerCase().trim();
          return text.includes(searchVal) || searchVal.includes(text);
        });
        
        if (partialMatch) {
          partialMatch.click();
          await new Promise(resolve => setTimeout(resolve, 500));
          return "PARTIAL";
        }
      }
      
      // 返回可用選項供調試
      const availableOptions = options.map(opt => (opt.textContent || "").trim()).slice(0, 8);
      return "NOPTION:" + availableOptions.join(",");
    }, value);

    return selectResult;

  } catch (error) {
    return "TIMEOUT:" + error.message;
  }
}

// 使用 mapping.dropdown 進行值映射
function mapDropdownValue(fieldName, value) {
  if (!mapping.dropdown) return value;
  
  const mappings = {
    'mainSubjectId': mapping.dropdown.subject,
    'cardSubjectId': mapping.dropdown.subject,
    'cardType': mapping.dropdown.type,
    'level': mapping.dropdown.level,
    'centuryId': mapping.dropdown.century,
    'syllabus': mapping.dropdown.syllabus,
    'cardSyllabus': mapping.dropdown.syllabus,
    'area': mapping.dropdown.area
  };
  
  const fieldMapping = mappings[fieldName];
  if (fieldMapping && fieldMapping[value]) {
    return fieldMapping[value];
  }
  
  return value;
}

// ---------- 卡片頁下拉選單處理 ----------
async function setByDropdownForCard(page, labelText, value) {
  try {
    const result = await page.evaluate(async (label, val) => {
      // 卡片頁面的下拉選單可能使用不同的ID模式
      const fieldMappings = {
        '學科': ['mui-component-select-cardSubjectId', 'cardSubjectId', 'subject'],
        '類別': ['mui-component-select-cardType', 'cardType', 'type'],
        '課綱': ['mui-component-select-cardSyllabus', 'cardSyllabus', 'syllabus']
      };
      
      let dropdown = null;
      
      if (fieldMappings[label]) {
        for (const id of fieldMappings[label]) {
          dropdown = document.getElementById(id);
          if (dropdown) {
            console.log('卡片下拉選單找到ID:', id);
            break;
          }
        }
      }
      
      // 如果通過ID找不到，嘗試通過標籤文字查找
      if (!dropdown) {
        console.log('通過文字標籤查找卡片下拉選單:', label);
        
        const labels = Array.from(document.querySelectorAll("*"))
          .filter(el => {
            const text = (el.textContent || "").trim();
            return text === label && (el.tagName === 'LABEL' || el.tagName === 'SPAN' || el.tagName === 'DIV');
          });
          
        for (const lab of labels) {
          let container = lab;
          for (let i = 0; i < 5; i++) {
            container = container.parentElement;
            if (!container) break;
            
            const potentialDropdown = container.querySelector('[role="combobox"]');
            if (potentialDropdown) {
              dropdown = potentialDropdown;
              console.log('找到卡片下拉選單在容器中');
              break;
            }
          }
          if (dropdown) break;
        }
      }
      
      if (!dropdown) {
        return "NODROPDOWN";
      }
      
      if (dropdown.getAttribute('role') !== 'combobox') {
        return "NOT_COMBOBOX";
      }
      
      console.log('點擊卡片下拉選單');
      
      // 使用真實的點擊方式
      dropdown.focus();
      dropdown.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
      dropdown.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      dropdown.click();
      
      return "CLICKED";
    }, labelText, value);

    if (result !== "CLICKED") {
      return result;
    }

    // 等待選項列表出現
    try {
      await page.waitForFunction(
        () => {
          const listboxes = document.querySelectorAll('ul[role="listbox"]');
          return listboxes.length > 0;
        },
        { timeout: 3000 }
      );
    } catch (e) {
      console.log('等待卡片選項列表超時，但仍嘗試查找...');
    }

    await new Promise(resolve => setTimeout(resolve, 500));

    // 處理選項選擇
    const selectResult = await page.evaluate(async (val) => {
      const listboxes = document.querySelectorAll('ul[role="listbox"]');
      
      if (listboxes.length === 0) {
        return "NOLISTBOX";
      }
      
      const listbox = listboxes[listboxes.length - 1];
      const options = Array.from(listbox.querySelectorAll('li.MuiMenuItem-root'));
      
      if (options.length === 0) {
        return "NOOPTIONS";
      }
      
     
      const targetOption = options.find(opt => {
        const text = (opt.textContent || "").trim();
        return text === val;
      });

      if (targetOption) {
        targetOption.click();
        await new Promise(resolve => setTimeout(resolve, 500));
        return "OK";
      } else {
        
        const partialMatch = options.find(opt => {
          const text = (opt.textContent || "").toLowerCase().trim();
          const searchVal = val.toLowerCase().trim();
          return text.includes(searchVal) || searchVal.includes(text);
        });
        
        if (partialMatch) {
          partialMatch.click();
          await new Promise(resolve => setTimeout(resolve, 500));
          return "PARTIAL";
        }
      }
      
      // 返回可用選項供調試
      const availableOptions = options.map(opt => (opt.textContent || "").trim()).slice(0, 8);
      return "NOPTION:" + availableOptions.join(",");
    }, value);

    return selectResult;

  } catch (error) {
    return "TIMEOUT:" + error.message;
  }
}

// ---------- 任務頁填寫 ----------
async function waitForTaskReady(page) {
  try {
    await page.waitForSelector('[name="description"], [name="name"]', { timeout: 5000 });
  } catch {}
}

async function fillTask(page, taskData) {
  console.log("✍️ 任務頁開始填寫...");
  await waitForTaskReady(page);

  // 定義哪些欄位是下拉選單 - 只處理確定存在的
  const dropdownFields = ['mainSubjectId', 'level', 'centuryId'];

  for (const [title, cfg] of Object.entries(mapping.task)) {
    const name = typeof cfg === "string" ? cfg : cfg.name;
    const isRich = !!(typeof cfg === "object" && cfg.rich);
    const isDropdown = dropdownFields.includes(name);
    const labels = (typeof cfg === "object" && cfg.label) ? cfg.label : [title];
    let val = taskData[name];
    
    if (!val) continue;

    // 如果是下拉選單欄位，先進行值映射
    if (isDropdown) {
      const mappedVal = mapDropdownValue(name, val);
      val = mappedVal;
    }

    let done = false;
    
    // 如果是下拉選單，使用下拉選單處理
    if (isDropdown) {
      for (const lb of labels) {
        const r = await setByDropdown(page, lb, val);
        if (r === "OK" || r === "PARTIAL") { 
          done = true; 
          console.log(`✅ ${name} (dropdown): ${preview(val)} ${r === "PARTIAL" ? "(部分匹配)" : ""}`);
          break; 
        } else if (r.startsWith("NOPTION:")) {
          console.log(`⚠️ ${name} 找不到選項 "${val}"，可用選項: ${r.substring(8)}`);
        } else {
          console.log(`❌ ${name} 下拉選單操作失敗: ${r}`);
        }
      }
    }
    
    // 如果下拉選單失敗或不是下拉選單，嘗試原有方法
    if (!done) {
      
      const r1 = await setByNameNative(page, name, val);
      if (r1 === "OK") { done = true; }
     
      if (!done && isRich) {
        for (const lb of labels) {
          const r2 = await typeIntoRichByLabel(page, lb, val);
          if (r2 === "OK") { done = true; break; }
        }
      }
    }
    
    if (!done && !isDropdown) {
      console.log(`⚠️ 找不到欄位: ${name}`);
    } else if (!isDropdown) {
      console.log(`✅ ${name}: ${preview(val)}`);
    }
  }

  console.log("🎉 任務頁完成！");
}

// ---------- 卡片頁填寫（統一處理方式） ----------
async function fillOneCard(page, card, index) {
  console.log(`🎴 開始填寫卡片 ${index + 1}...`);
  
  // 顯示解析出的卡片數據
  console.log(`🔍 卡片數據預覽:`);
  console.log(`   - 卡片名稱: "${card.cardTitle || ''}"`);
  console.log(`   - 學科: "${card.cardSubjectId || ''}"`);
  console.log(`   - 類別: "${card.cardType || ''}"`);

  // 定義卡片的下拉選單欄位
  const cardDropdownFields = ['cardSubjectId', 'cardType'];

  for (const [title, cfg] of Object.entries(mapping.card)) {
    const name = typeof cfg === "string" ? cfg : cfg.name;
    const isRich = !!(typeof cfg === "object" && cfg.rich);
    const isDropdown = cardDropdownFields.includes(name);
    const labels = (typeof cfg === "object" && cfg.label) ? cfg.label : [title];
    let val = card[name];
    
    if (!val) continue;

    // 如果是下拉選單欄位，先進行值映射
    if (isDropdown) {
      const mappedVal = mapDropdownValue(name, val);
      val = mappedVal;
    }

    let done = false;
    
    // 如果是下拉選單，使用下拉選單處理
    if (isDropdown) {
      for (const lb of labels) {
        const r = await setByDropdownForCard(page, lb, val);
        if (r === "OK" || r === "PARTIAL") { 
          done = true; 
          console.log(`✅ ${name} (dropdown): ${preview(val)} ${r === "PARTIAL" ? "(部分匹配)" : ""}`);
          break; 
        } else if (r.startsWith("NOPTION:")) {
          console.log(`⚠️ ${name} 找不到選項 "${val}"，可用選項: ${r.substring(8)}`);
        } else {
          console.log(`❌ ${name} 下拉選單操作失敗: ${r}`);
        }
      }
    }
    
    // 如果下拉選單失敗或不是下拉選單，嘗試原有方法
    if (!done) {
      // 先試 name
      const r1 = await setByNameNative(page, name, val);
      if (r1 === "OK") { done = true; }
      // 富文本備援
      if (!done && isRich) {
        for (const lb of labels) {
          const r2 = await typeIntoRichByLabel(page, lb, val);
          if (r2 === "OK") { done = true; break; }
        }
      }
      // 文字輸入備援
      if (!done) {
        for (const lb of labels) {
          const r3 = await typeIntoInputByLabel(page, lb, val);
          if (r3 === "OK") { done = true; break; }
        }
      }
    }
    
    if (!done && !isDropdown) {
      console.log(`⚠️ 找不到卡片欄位: ${name}`);
    } else if (!isDropdown) {
      console.log(`✅ ${name}: ${preview(val)}`);
    }
  }

  console.log(`🎉 卡片 ${index + 1} 填寫完成！`);
}

// ---------- 上傳圖片視窗：兩欄位自動填 ----------
// ---------- 在對話框中填入搜尋字串（搜尋現有圖片用） ----------
async function typeIntoFirstTextInputInDialog(page, value) {
  return page.evaluate((val) => {
    const dialogs = document.querySelectorAll('[role="dialog"]');
    const dialog = dialogs[dialogs.length - 1];
    if (!dialog) return 'NODIALOG';

    const candidates = Array.from(dialog.querySelectorAll(
      'input[type="text"], input:not([type]), [role="textbox"]'
    )).filter(el => {
      const style = window.getComputedStyle(el);
      return style && style.display !== 'none' && style.visibility !== 'hidden';
    });

    const input = candidates[0];
    if (!input) return 'NF';

    // 優先用原生 setter，確保 React/MUI 能收到事件
    if (input.tagName.toLowerCase() === 'input') {
      const desc = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
      input.focus();
      desc.set.call(input, val);
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
      input.blur();
      return 'OK';
    }

    input.focus();
    const sel = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(input);
    sel.removeAllRanges();
    sel.addRange(range);
    document.execCommand('delete', false, null);
    document.execCommand('insertText', false, val);
    input.dispatchEvent(new InputEvent('input', { bubbles: true, data: val }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
    return 'OK';
  }, value);
}

// 打開《搜尋現有圖片》後，把卡片名稱貼到搜尋框
async function fillSearchExistingImageDialog(page, query) {
  await getTopDialog(page); 
  const r = await typeIntoFirstTextInputInDialog(page, query || '');
  console.log('🔎 搜尋關鍵字:', r, '→', (query || '').slice(0, 24));
}
async function getTopDialog(page) {
  await page.waitForSelector('[role="dialog"]', { timeout: 8000 });
  return page.$$('[role="dialog"]').then(list => list[list.length - 1]);
}

async function typeIntoInputByLabelInDialog(page, labelText, value) {
  return page.evaluate((label, val) => {
    const dialogs = document.querySelectorAll('[role="dialog"]');
    const dialog = dialogs[dialogs.length - 1];
    if (!dialog) return 'NODIALOG';

    const lab = Array.from(dialog.querySelectorAll('label,div,span,p,h6,h5'))
      .find(el => (el.textContent || '').trim() === label);
    const row = lab ? lab.closest('.MuiGrid-root, .MuiStack-root, div') : null;
    const input = row ? row.querySelector('input[type="text"], input:not([type])') : null;
    if (!input) return 'NF';

    const desc = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
    input.focus();
    desc.set.call(input, val);
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
    input.blur();
    return 'OK';
  }, labelText, value);
}

async function typeIntoTextareaByLabelInDialog(page, labelText, value) {
  return page.evaluate((label, val) => {
    const dialogs = document.querySelectorAll('[role="dialog"]');
    const dialog = dialogs[dialogs.length - 1];
    if (!dialog) return 'NODIALOG';

    const lab = Array.from(dialog.querySelectorAll('label,div,span,p,h6,h5'))
      .find(el => (el.textContent || '').trim() === label);
    const row = lab ? lab.closest('.MuiGrid-root, .MuiStack-root, div') : null;
    const box = row && (row.querySelector('textarea') ||
                        row.querySelector('[contenteditable="true"]') ||
                        row.querySelector('.ql-editor') ||
                        row.querySelector('[role="textbox"]'));
    if (!box) return 'NF';

    if (box.tagName && box.tagName.toLowerCase() === 'textarea') {
      const desc = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value');
      box.focus();
      desc.set.call(box, val);
      box.dispatchEvent(new Event('input', { bubbles: true }));
      box.dispatchEvent(new Event('change', { bubbles: true }));
      box.blur();
      return 'OK';
    }

    box.focus();
    const sel = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(box);
    sel.removeAllRanges();
    sel.addRange(range);
    document.execCommand('delete', false, null);
    document.execCommand('insertText', false, val);
    box.dispatchEvent(new InputEvent('input', { bubbles: true, data: val }));
    box.dispatchEvent(new Event('change', { bubbles: true }));
    return 'OK';
  }, labelText, value);
}

async function fillUploadImageDialog(page, title, description, imageData = null) {
  await getTopDialog(page); 
  
  // 先嘗試找到對話框中的所有標籤，用於調試
  const availableLabels = await page.evaluate(() => {
    const dialogs = document.querySelectorAll('[role="dialog"]');
    const dialog = dialogs[dialogs.length - 1];
    if (!dialog) return [];
    
    const labels = Array.from(dialog.querySelectorAll('label,div,span,p,h6,h5'))
      .map(el => (el.textContent || '').trim())
      .filter(text => text.length > 0 && text.length < 50);
    
    return [...new Set(labels)]; 
  });
  
  console.log('🔍 對話框中的可用標籤:', availableLabels);
  
  // 嘗試多種可能的標籤文字
  const nameLabelOptions = ['圖片名稱', '圖片標題', '名稱', '標題', 'Image Name', 'Name'];
  const descLabelOptions = ['圖片描述', '圖片說明', '描述', '說明', 'Image Description', 'Description'];
  
  let r1 = 'NF', r2 = 'NF';
  
  for (const nameLabel of nameLabelOptions) {
    r1 = await typeIntoInputByLabelInDialog(page, nameLabel, title || '');
    if (r1 === 'OK') {
      console.log(`✅ 成功使用標籤: "${nameLabel}"`);
      break;
    }
  }
  
  for (const descLabel of descLabelOptions) {
    r2 = await typeIntoTextareaByLabelInDialog(page, descLabel, description || '');
    if (r2 === 'OK') {
      console.log(`✅ 成功使用標籤: "${descLabel}"`);
      break;
    }
  }
  
  console.log('🖼 圖片名稱:', r1, '圖片描述:', r2);

  // 如果有從 docx 解析出的圖片資料，嘗試自動上傳
  if (imageData) {
    console.log('📸 嘗試自動上傳從 docx 解析的圖片...');
    
    let uploadSuccess = false;
    
    try {
      // 儲存圖片到暫存檔案
      const tempImagePath = await saveImageToTemp(imageData, Date.now());
      if (tempImagePath) {
        // 尋找檔案上傳元素
        const fileInput = await page.$('input[type="file"]');
        if (fileInput) {
          await fileInput.uploadFile(tempImagePath);
          console.log('✅ 圖片已自動上傳');
          uploadSuccess = true;
        } else {
          console.log('⚠️ 未找到檔案上傳元素，請手動選擇圖片');
        }
      }
    } catch (error) {
      console.error('❌ 自動上傳圖片失敗:', error.message);
      console.log('👉 請手動選擇並上傳圖片');
    }
    
    // 只有在成功上傳後才等待
    if (uploadSuccess) {
      try {
        // 短暫等待上傳處理
        await new Promise(resolve => setTimeout(resolve, 2000));
      } catch (waitError) {
        // 等待錯誤不影響主要功能，靜默處理
        console.log('⏳ 等待處理完成...');
      }
    }
  }

  // 若要自動送出，解除註解即可
  // await page.evaluate(() => {
  //   const btn = Array.from(document.querySelectorAll('[role="dialog"] button, [role="dialog"] [role="button"]'))
  //     .find(b => /確定上傳/.test((b.textContent || '')));
  //   btn?.click();
  // });
}

// ---------- 流程 ----------
async function runOnce(page, wordPath) {
  if (argv.url) {
    try { await page.goto(argv.url, { waitUntil: "domcontentloaded", timeout: 60000 }); }
    catch (e) { console.log("⚠️ 自動開啟 URL 失敗，請手動切頁：", e.message); }
  }

  if (!wordPath) { console.error("❌ 需提供 --word"); return; }

  console.log("📂 解析 Word:", wordPath);
  const { taskData, cardDataList } = await parseWord(wordPath);

  // 任務頁
  await ask("👉 請切到『建立任務』頁（需登入）。準備好按 Enter 開始填...");
  await fillTask(await getActivePage(page), taskData);

  if (argv.mode === "task") return;

  // 卡片頁（固定 12 張）
  const cards = cardDataList;
  for (let i = 0; i < cards.length; i++) {
    const tag = `${Math.floor(i/4) + 1}-${(i % 4) + 1}`; // 1-1..3-4

    // 1) 填卡片
    await ask(`👉 請切到『卡片模式』頁（目前要填：${tag}）。準備好後按 Enter 開始...`);
    await fillOneCard(await getActivePage(page), cards[i], i);

    // 2) 【新增】在上傳圖片前，先開『搜尋現有圖片』視窗，貼上卡片名稱到搜尋框
    const keyword = cards[i].cardTitle || '';
    await ask('👉 請在卡片頁按【搜尋現有圖片】打開視窗，準備好後按 Enter，我會貼上搜尋關鍵字...');
    await fillSearchExistingImageDialog(await getActivePage(page), keyword);
    console.log('✅ 已貼上搜尋關鍵字（卡片名稱）到搜尋框');

    // （你可以選擇在這裡暫停，讓你自己按「搜尋」或挑圖）
    // await ask('👉 如需馬上搜尋，請手動按「搜尋」或輸入法 Enter。挑好圖後按 Enter 繼續上傳圖片步驟...');

    // 3) 上傳圖片視窗（增強版：支援自動上傳 docx 中的圖片）
    const hasImageData = cards[i].imageData;
    if (hasImageData) {
      console.log(`🖼️  卡片 ${tag} 包含來自 docx 的圖片，將嘗試自動上傳`);
      await ask('👉 請在卡片頁按【上傳圖片】打開視窗，準備好後按 Enter 繼續（將自動上傳圖片）...');
    } else {
      console.log(`📝 卡片 ${tag} 無圖片資料，需手動上傳`);
      await ask('👉 請在卡片頁按【上傳圖片】打開視窗，準備好後按 Enter 繼續...');
    }
    
    await fillUploadImageDialog(
      await getActivePage(page), 
      cards[i].cardTitle, 
      cards[i].cardDescription,
      cards[i].imageData || null
    );
  }

  console.log("✅ 全部卡片處理完畢！");
}

async function getActivePage(page) {
  const pages = await page.browser().pages();
  const active = pages[pages.length - 1];
  await active.bringToFront();
  return active;
}

// ---------- Windows Chrome（保留） ----------
function getWindowsChromePath() {
  if (process.platform !== "win32") return undefined;
  const cand = [
    path.join(process.env["PROGRAMFILES"] || "", "Google/Chrome/Application/chrome.exe"),
    path.join(process.env["PROGRAMFILES(X86)"] || "", "Google/Chrome/Application/chrome.exe"),
    path.join(process.env["LOCALAPPDATA"] || "", "Google/Chrome/Application/chrome.exe"),
    path.join(process.env["PROGRAMFILES"] || "", "Microsoft/Edge/Application/msedge.exe"),
    path.join(process.env["PROGRAMFILES(X86)"] || "", "Microsoft/Edge/Application/msedge.exe"),
  ].filter(Boolean);
  return cand.find(p => fs.existsSync(p));
}

// ---------- Main ----------
async function main() {
  console.log("🚀 進入 main()");
  const userDataDir = path.resolve("./user_data");
  if (!fs.existsSync(userDataDir)) fs.mkdirSync(userDataDir, { recursive: true });

  const execPath = getWindowsChromePath();

  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: null,
    userDataDir,
    executablePath: execPath,
    args: ["--disable-features=ImprovedKeyboardShortcuts", "--no-sandbox"],
  });

  const [page] = await browser.pages();

  // 啟動時自動前往指定網址
  await page.goto("https://adl.edu.tw/twa-admin/edit/missionEdit", { waitUntil: "domcontentloaded" });

  if (!argv.word) {
    console.error("❌ 請提供 --word");
    if (argv.close) await browser.close();
    process.exit(1);
  }

  // 先跑第一個檔案（必要）
  await runOnce(page, argv.word);

  // 若開啟 --loop，持續詢問下一個檔案路徑
  if (argv.loop) {
    while (true) {
      const next = (await ask("📄 請輸入下一個 Word 檔路徑（直接 Enter 結束）：")).trim();
      if (!next) break;
      await runOnce(await getActivePage(page), next);
    }
  }

  if (argv.close) {
    await browser.close();
    console.log("✅ 關閉瀏覽器，程式結束");
  } else {
    console.log("✅ 流程完成，瀏覽器保持開著（沿用登入）。");
  }
}

main().catch((err) => {
  console.error("❌ 執行錯誤:", err);
});
