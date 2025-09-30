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
  .option("word", { type: "string", describe: "Word 檔路徑（.docx）" })
  .option("loop", { type: "boolean", default: false })
  .option("close", { type: "boolean", default: false })
  .option("url", { type: "string", describe: "（可選）自動前往『建立任務』頁 URL" })
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

    // 整行只有括弧說明 -> 丟掉
    if (/^[（(][^）)]+[）)]$/.test(trimmed)) continue;

    // 第一段若以括弧說明起頭，僅剝掉這段的前導括弧，保留後面正文
    if (!firstKeptSeen && /^[（(][^）)]+[）)]/.test(trimmed)) {
      line = line.replace(/^[（(][^）)]+[）)]\s*/, "");
    }

    // 被誤捲的小標題 -> 丟掉
    if (neighborHeadingsRE.test(trimmed)) continue;

    if (line.trim().length && !firstKeptSeen) firstKeptSeen = true;
    out.push(line);
  }
  return out.join("\n").trim();
}

function firstLine(text) {
  return (String(text).split("\n").find(l => l.trim().length) || "").trim();
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
      v = firstLine(v);
    }
    if (name.startsWith("missionHintSet")) {
      v = v.split("\n").map(s => s.replace(/（限25字[^）]*）/g, "").trim()).filter(Boolean)[0] || "";
    }

    // 只對這兩個欄位剝括弧備註/小標題
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
    "課綱": "cardSyllabus"
  };

  const lines = value.split("\n").map(s => s.trim()).filter(Boolean);
  const cards = [];
  let cur = null;
  let curField = null;

  for (const line of lines) {
    if (labels.includes(line)) {
      // 新卡片開始：遇到下一個「卡片名稱」
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

  // 正規化：名稱取第一行；內容清理；（不再處理學科/類別）
  for (const c of cards) {
    if (c.cardTitle) c.cardTitle = firstLine(c.cardTitle);
    if (c.cardDescription) {
      let cleaned = stripBracketNotesBlock(cleanTextBase(c.cardDescription));
      cleaned = cleaned.split("\n").map(s => s.trim()).filter(Boolean)[0] || ""; // 只取第一個非空行
      c.cardDescription = cleaned;
    }
    // if (c.cardSubjectId) c.cardSubjectId = firstLine(c.cardSubjectId);
    // if (c.cardType)      c.cardType      = firstLine(c.cardType);
  }

  // 固定只取 12 張
  return cards.slice(0, 12);
}

async function parseWord(wordPath) {
  const { value } = await mammoth.extractRawText({ path: wordPath });
  const taskData = await parseTaskSections(value);
  const cardDataList = parseCards(value);
  return { taskData, cardDataList };
}

// ---------- DOM 寫入 ----------
async function setByNameNative(page, name, value) {
  const selector = `[name="${cssEscape(name)}"]`;
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

// ---------- 任務頁填寫 ----------
async function waitForTaskReady(page) {
  try {
    await page.waitForSelector('[name="description"], [name="name"]', { timeout: 5000 });
  } catch {}
}

async function fillTask(page, taskData) {
  console.log("✍️ 任務頁開始填寫...");
  await waitForTaskReady(page);

  for (const [title, cfg] of Object.entries(mapping.task)) {
    const name = typeof cfg === "string" ? cfg : cfg.name;
    const isRich = !!(typeof cfg === "object" && cfg.rich);
    const labels = (typeof cfg === "object" && cfg.label) ? cfg.label : [title];
    const val = taskData[name];
    if (!val) continue;

    let done = false;
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
    console.log(done ? `✅ ${name}: ${preview(val)}` : `⚠️ 找不到欄位: ${name}`);
  }

  console.log("🎉 任務頁完成！");
}

// ---------- 卡片頁填寫（僅標題與文字內容） ----------
async function fillOneCard(page, card, index) {
  const title = card.cardTitle || "";
  const desc  = card.cardDescription || "";

  // 卡片名稱（單行）
  if (title) {
    const r = await typeIntoInputByLabel(page, "卡片名稱", title);
    console.log(r === "OK" ? `✅ cardTitle: ${preview(title)}` : "⚠️ 找不到『卡片名稱』輸入框");
  }

  // 文字內容（多行/textarea/contenteditable 都可）
  if (desc) {
    const r = await typeIntoRichByLabel(page, "文字內容", desc);
    console.log(r === "OK" ? `✅ cardDescription: ${preview(desc)}` : "⚠️ 找不到『文字內容』輸入框");
  }
}

// ---------- 上傳圖片視窗：兩欄位自動填 ----------
// ---------- 在對話框中填入搜尋字串（搜尋現有圖片用） ----------
async function typeIntoFirstTextInputInDialog(page, value) {
  return page.evaluate((val) => {
    const dialogs = document.querySelectorAll('[role="dialog"]');
    const dialog = dialogs[dialogs.length - 1];
    if (!dialog) return 'NODIALOG';

    // 盡量找第一個可見的文字輸入框（含 MUI、一般 input、role=textbox）
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

    // 富文本/role=textbox 的備援
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
  await getTopDialog(page); // 等搜尋視窗出現
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

async function fillUploadImageDialog(page, title, description) {
  await getTopDialog(page); // 等視窗出現
  const r1 = await typeIntoInputByLabelInDialog(page, '圖片名稱', title || '');
  const r2 = await typeIntoTextareaByLabelInDialog(page, '圖片描述', description || '');
  console.log('🖼 圖片名稱:', r1, '圖片描述:', r2);

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

    // 3) 上傳圖片視窗（原本就有的步驟）
    await ask('👉 請在卡片頁按【上傳圖片】打開視窗，準備好後按 Enter 繼續...');
    await fillUploadImageDialog(await getActivePage(page), cards[i].cardTitle, cards[i].cardDescription);
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
