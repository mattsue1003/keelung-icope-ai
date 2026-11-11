// 基隆 ICOPE × AI 工作坊：網頁型簡報 → 可列印講義 PPTX
// node ppt/build.js  （在 keelung-icope-ai 資料夾執行）
const path = require("path");
const pptxgen = require(path.join(__dirname, "node_modules/pptxgenjs"));
const D = require("./slides.json");
const SZ = require("./sizes.json");
const ROOT = path.join(__dirname, "..");
const IMG = (p) => path.join(ROOT, p);

const F = "Microsoft JhengHei";
const C = { bg: "FFFEFA", ink: "1B1A18", muted: "66625B", line: "E4E0D6", soft: "F3F0E8",
  mag: "D7306E", blue: "1773EA", lime: "C9E44A", amber: "FBA705", sky: "47B2E4", code: "F6F3EB" };
const W = 10, H = 5.625, M = 0.5;

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "基隆 ICOPE 培力實務工作坊｜AI 工作流 ICOPE 實戰應用";
pres.author = "鍾孟修 職能治療師";
pres.theme = { headFontFace: F, bodyFontFace: F };

pres.defineSlideMaster({
  title: "PAPER",
  background: { color: C.bg },
  objects: [
    { rect: { x: 9.18, y: 0.28, w: 0.13, h: 0.13, fill: { color: C.mag } } },
    { rect: { x: 9.36, y: 0.28, w: 0.13, h: 0.13, fill: { color: C.lime } } },
    { rect: { x: 9.18, y: 0.46, w: 0.13, h: 0.13, fill: { color: C.blue } } },
    { rect: { x: 9.36, y: 0.46, w: 0.13, h: 0.13, fill: { color: C.amber } } },
    { text: { text: "基隆 ICOPE 培力實務工作坊｜AI 工作流", options: { x: M, y: 5.25, w: 6, h: 0.25, fontFace: F, fontSize: 9, color: C.muted, margin: 0 } } },
  ],
  slideNumber: { x: 9.0, y: 5.22, w: 0.5, h: 0.25, fontFace: F, fontSize: 9, color: C.muted, align: "right" },
});

let secColorIdx = 0;
const accent = (eb) => /實作③|實作④|方法/.test(eb) ? C.amber : /實作⑤|實作⑥|驗算/.test(eb) ? C.mag : /實作①|實作②/.test(eb) ? C.blue : /實作⑦|補充/.test(eb) ? C.lime : C.ink;

function header(s, eyebrow, title, color) {
  s.addShape(pres.shapes.RECTANGLE, { x: M, y: 0.36, w: 0.12, h: 0.12, fill: { color: color || C.mag } });
  s.addText(eyebrow, { x: M + 0.2, y: 0.28, w: 7.5, h: 0.28, fontFace: F, fontSize: 11, bold: true, color: C.muted, margin: 0, isTextBox: true, charSpacing: 1 });
  s.addText(title, { x: M, y: 0.6, w: 8.5, h: 0.6, fontFace: F, fontSize: [...title].length > 26 ? 19 : 24, bold: true, color: C.ink, margin: 0, isTextBox: true, valign: "top" });
}
function imgFit(s, p, x, y, w, h, bg) {
  if (bg !== false) s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: bg || "2A2925" }, rectRadius: 0.06, line: { color: C.line, width: 0.75 } });
  const sz = SZ[p] || [16, 9];
  const r = Math.min((w - 0.08) / sz[0], (h - 0.08) / sz[1]);
  const iw = sz[0] * r, ih = sz[1] * r;
  s.addImage({ path: IMG(p), x: x + (w - iw) / 2, y: y + (h - ih) / 2, w: iw, h: ih, altText: p });
}
function numBadge(s, n, x, y, color) {
  s.addShape(pres.shapes.OVAL, { x, y, w: 0.32, h: 0.32, fill: { color: color || C.ink } });
  s.addText(String(n), { x, y, w: 0.32, h: 0.32, fontFace: F, fontSize: 12, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true });
}
function promptBox(s, lab, text, x, y, w, h, size) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: C.code }, line: { color: C.line, width: 0.75 }, rectRadius: 0.08 });
  s.addText(lab, { x: x + 0.15, y: y + 0.1, w: w - 0.3, h: 0.25, fontFace: F, fontSize: 10, bold: true, color: C.muted, margin: 0, isTextBox: true });
  s.addText(text, { x: x + 0.15, y: y + 0.38, w: w - 0.3, h: h - 0.48, fontFace: F, fontSize: size || 13, color: C.ink, margin: 0, valign: "top", isTextBox: true, paraSpaceAfter: 2 });
}
function chunks(arr) {
  const n = arr.length, k = Math.ceil(n / 3), base = Math.floor(n / k), extra = n % k, out = [];
  let i = 0;
  for (let j = 0; j < k; j++) { const c = base + (j < extra ? 1 : 0); out.push(arr.slice(i, i + c)); i += c; }
  return out;
}
function stepSlides(d, note) {
  const col = accent(d.eyebrow);
  const groups = chunks(d.steps);
  let num = 1;
  groups.forEach((g, gi) => {
    const s = pres.addSlide({ masterName: "PAPER" });
    header(s, d.eyebrow + (groups.length > 1 ? `（${gi + 1}/${groups.length}）` : ""), d.h2, col);
    const n = g.length, gap = 0.3, cw = (W - 2 * M - gap * (n - 1)) / n;
    const ih = n === 3 ? 1.95 : 2.15;
    g.forEach((st, i) => {
      const x = M + i * (cw + gap), y = 1.38;
      imgFit(s, st.img, x, y, cw, ih);
      numBadge(s, num, x, y + ih + 0.14, col === C.lime ? C.ink : col);
      s.addText(st.title, { x: x + 0.42, y: y + ih + 0.08, w: cw - 0.42, h: 0.52, fontFace: F, fontSize: 13, bold: true, color: C.ink, margin: 0, valign: "middle", isTextBox: true, fit: "shrink" });
      s.addText(st.desc, { x, y: y + ih + 0.64, w: cw, h: 5.15 - (y + ih + 0.64), fontFace: F, fontSize: 11.5, color: C.muted, margin: 0, valign: "top", isTextBox: true, paraSpaceAfter: 2 });
      num++;
    });
    s.addNotes(g.map((st, i) => `步驟：${st.title}。截圖說明：${st.cap}`).join("\n"));
  });
  d.prompts.forEach((p) => {
    const s = pres.addSlide({ masterName: "PAPER" });
    header(s, d.eyebrow + "｜提示詞", p.lab.replace(/^提示詞・/, "提示詞：").replace(/^Excel/, "Excel"), col);
    promptBox(s, "複製到 AI 的輸入框，把（ ）裡的內容換成你的", p.text, M, 1.4, W - 2 * M, 2.6, 15);
    if (note) s.addText(note, { x: M, y: 4.2, w: W - 2 * M, h: 0.8, fontFace: F, fontSize: 12, color: C.muted, margin: 0, isTextBox: true, valign: "top" });
    s.addText("網頁版可一鍵複製：mattsue1003.github.io/keelung-icope-ai", { x: M, y: 4.85, w: 6, h: 0.3, fontFace: F, fontSize: 10, color: C.blue, margin: 0, isTextBox: true });
  });
}

const by = (t) => D.find((d) => d.title === t);

// 1 封面
{
  const s = pres.addSlide({ masterName: "PAPER" });
  [[6.6, 0.5, 0.45, C.blue], [7.6, 1.2, 0.22, C.lime], [8.4, 0.75, 0.3, C.mag], [7.1, 2.0, 0.55, C.amber], [8.8, 2.3, 0.18, C.sky], [6.3, 2.7, 0.2, C.lime]].forEach(([x, y, w, c]) =>
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: w, fill: { color: c }, rectRadius: 0.05 }));
  s.addText("基隆 ICOPE 培力實務工作坊", { x: M, y: 0.55, w: 6, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: C.muted, margin: 0, isTextBox: true, charSpacing: 2 });
  s.addText([{ text: "ICOPE ", options: { color: C.ink } }, { text: "× ", options: { color: C.mag } }, { text: "AI", options: { color: C.ink } }], { x: M, y: 1.0, w: 6, h: 0.45, fontFace: "Arial", fontSize: 20, bold: true, margin: 0, isTextBox: true, charSpacing: 8 });
  s.addText("AI 工作流\nICOPE 實戰應用", { x: M, y: 1.5, w: 6, h: 1.6, fontFace: F, fontSize: 40, bold: true, color: C.ink, margin: 0, isTextBox: true, valign: "top" });
  s.addText("評估做完之後的事，交給 AI 打草稿：口述紀錄、讀手冊、做衛教圖、做網頁型簡報、整理指標數據。", { x: M, y: 3.15, w: 5.6, h: 0.65, fontFace: F, fontSize: 13, color: C.muted, margin: 0, isTextBox: true });
  const meta = [["日期", "11/18（三）"], ["時間", "13:30–16:30"], ["地點", "巨匠電腦 基隆分校 9F"], ["講師", "鍾孟修 職能治療師"]];
  s.addShape(pres.shapes.LINE, { x: M, y: 3.98, w: 6.6, h: 0, line: { color: C.line, width: 0.75 } });
  meta.forEach(([k, v], i) => {
    const x = M + i * 1.65;
    s.addText(k, { x, y: 4.1, w: 1.6, h: 0.22, fontFace: F, fontSize: 9, bold: true, color: C.muted, margin: 0, isTextBox: true });
    s.addText(v, { x, y: 4.32, w: 1.6, h: 0.5, fontFace: F, fontSize: 12, bold: true, color: C.ink, margin: 0, isTextBox: true, valign: "top" });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.35, y: 3.05, w: 2.15, h: 2.05, fill: { color: "FFFFFF" }, line: { color: C.line, width: 0.75 }, rectRadius: 0.1 });
  s.addImage({ path: IMG("img/qr-deck.png"), x: 7.75, y: 3.12, w: 1.35, h: 1.35, altText: "網頁型簡報 QR code" });
  s.addText("掃我開啟網頁版簡報", { x: 7.35, y: 4.48, w: 2.15, h: 0.28, fontFace: F, fontSize: 11, bold: true, color: C.ink, align: "center", margin: 0, isTextBox: true });
  s.addText("mattsue1003.github.io/keelung-icope-ai", { x: 7.35, y: 4.74, w: 2.15, h: 0.25, fontFace: "Arial", fontSize: 7.5, color: C.muted, align: "center", margin: 0, isTextBox: true });
}
// 2 講師
{
  const s = pres.addSlide({ masterName: "PAPER" });
  header(s, "講師", "大家好！我是鍾孟修 職能治療師", C.blue);
  s.addImage({ path: IMG("img/speaker.jpg"), x: M, y: 1.45, w: 2.3, h: 2.3, altText: "鍾孟修 職能治療師", rounding: false });
  ["長照復能", "銀髮體適能", "認知及體智", "遊戲化教學"].forEach((t, i) => {
    const x = M + (i % 2) * 1.17, y = 3.95 + Math.floor(i / 2) * 0.42;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.12, h: 0.34, fill: { color: C.soft }, rectRadius: 0.17 });
    s.addText(t, { x, y, w: 1.12, h: 0.34, fontFace: F, fontSize: 9.5, bold: true, color: C.ink, align: "center", valign: "middle", margin: 0, isTextBox: true });
  });
  const list = (title, items, x, w) => {
    s.addText(title, { x, y: 1.4, w, h: 0.3, fontFace: F, fontSize: 11, bold: true, color: C.muted, margin: 0, isTextBox: true });
    s.addText(items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1 } })), { x, y: 1.72, w, h: 3.35, fontFace: F, fontSize: 10.5, color: C.ink, margin: 0, valign: "top", isTextBox: true, paraSpaceAfter: 3 });
  };
  list("協會與機構", ["社團法人台灣生活復能健康促進協會 理事長", "愛迪樂有限公司（105/08）", "愛迪樂居家職能治療所（107/01）", "愛迪樂物理治療所（107/12）", "愛迪樂中山聯合治療所（111/12）", "愛迪樂新北長期照顧服務機構（108/01）", "愛迪樂三重失智症據點（111/01）", "新北市私立愛迪樂東湖國小社區長照機構（112/05）", "銀光復能咖啡館 C 單位巷弄長照站（108/11）", "五股日間照顧中心及身障日間照顧中心（籌備中）"], 3.05, 3.05);
  list("角色與專長", ["職能治療師／執行長", "Homeseen 失智症遊戲包 共同創辦人", "關懷小丑協會 監事", "長庚大學職能治療系系友會 監事", "銀髮體適能教練、皮拉提斯教練 ABC 級", "認知桌遊活動遊戲設計", "MAT 功能性評估國際認證", "簡報遊戲化教學設計"], 6.25, 1.75);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.12, y: 1.45, w: 1.38, h: 2.55, fill: { color: "FFFFFF" }, line: { color: C.line, width: 0.75 }, rectRadius: 0.08 });
  s.addImage({ path: IMG("img/qr-fb.png"), x: 8.2, y: 1.55, w: 1.22, h: 1.22, altText: "粉絲專頁 QR code" });
  s.addText("加入粉絲專頁\n看更多作品", { x: 8.12, y: 2.82, w: 1.38, h: 0.5, fontFace: F, fontSize: 10.5, bold: true, color: C.ink, align: "center", margin: 0, isTextBox: true });
  s.addText("鍾孟修Matt-職能治療師", { x: 8.12, y: 3.36, w: 1.38, h: 0.5, fontFace: F, fontSize: 8, color: C.blue, align: "center", margin: 0, isTextBox: true, hyperlink: { url: "https://www.facebook.com/profile.php?id=100063748491881" } });
}
// 3 地圖
{
  const s = pres.addSlide({ masterName: "PAPER" });
  header(s, "今天的地圖", "一個下午，跑完 ICOPE 的 AI 工作流", C.amber);
  const rows = [["0", "暖身：AI 做出的 ICOPE 工具", "5′", C.mag], ["1", "三個 AI 怎麼選＋守住界線", "15′", C.ink], ["2", "實作① 語音轉文字＋手機即時字幕", "15′", C.blue], ["3", "實作② 讀資料：手冊、影片、筆記本", "20′", C.blue], ["4", "實作③ 衛教圖＋Canva 魔法圖層", "20′", C.amber], ["5", "實作④ 參考設計做網頁型簡報", "20′", C.amber],
    ["6", "實作⑤ 做 PPT：任務模式、筆記本、Canva", "25′", C.mag], ["7", "實作⑥ 看數據：Gemini 與 ChatGPT 對照", "15′", C.mag], ["8", "Napkin：文字變圖表", "5′", C.lime], ["9", "實作⑦ 上架：GitHub Pages", "10′", C.lime], ["10", "實作⑧ AI 代理人：OpenCode", "15′", C.ink], ["11", "提示詞庫＋課後任務", "5′", C.sky]];
  rows.forEach(([n, t, m, c], i) => {
    const col = i < 6 ? 0 : 1, r = i < 6 ? i : i - 6, x = M + col * 4.6, y = 1.4 + r * 0.6;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: y + 0.06, w: 0.36, h: 0.36, fill: { color: c }, rectRadius: 0.05 });
    s.addText(n, { x, y: y + 0.06, w: 0.36, h: 0.36, fontFace: "Arial", fontSize: 11, bold: true, color: c === C.lime || c === C.amber ? C.ink : "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(t, { x: x + 0.5, y, w: 3.3, h: 0.48, fontFace: F, fontSize: 13, bold: true, color: C.ink, margin: 0, valign: "middle", isTextBox: true });
    s.addText(m, { x: x + 3.8, y, w: 0.5, h: 0.48, fontFace: "Arial", fontSize: 11, bold: true, color: C.muted, align: "right", margin: 0, valign: "middle", isTextBox: true });
    s.addShape(pres.shapes.LINE, { x, y: y + 0.54, w: 4.3, h: 0, line: { color: C.line, width: 0.5 } });
  });
  s.addText("── 第 5、6 項之間休息 10 分鐘 ──", { x: M, y: 5.02, w: 8.9, h: 0.25, fontFace: F, fontSize: 10, color: C.muted, align: "center", margin: 0, isTextBox: true });
}
// 4 準備
{
  const s = pres.addSlide({ masterName: "PAPER" });
  header(s, "開始前 5 分鐘", "先把這五件事打勾", C.lime);
  ["電腦 Chrome 登入自己的 Google 帳號（用完記得登出）", "登入 ChatGPT：chatgpt.com（用 Google 帳號）", "打開 Gemini 與 Claude：gemini.google.com、claude.ai", "手機裝 ChatGPT、Gemini App（Claude 可選裝）", "加入 LINE 學員社群（封底 QR code）"].forEach((t, i) => {
    const y = 1.45 + i * 0.6;
    s.addShape(pres.shapes.RECTANGLE, { x: M, y: y + 0.08, w: 0.28, h: 0.28, fill: { color: "FFFFFF" }, line: { color: C.ink, width: 1.25 } });
    s.addText(t, { x: M + 0.45, y, w: 4.5, h: 0.45, fontFace: F, fontSize: 13, color: C.ink, margin: 0, valign: "middle", isTextBox: true });
  });
  [["電腦教室", "教室電腦多半不能安裝軟體。今天的實作都用瀏覽器完成；Typeless 請用自己的手機或筆電試。", C.blue], ["今天的原則", "AI 產出的是草稿。凡是要交到長輩或家屬手上的內容，都由專業人員覆核。", C.mag], ["練習檔", "在網頁版的去識別化頁與實作⑥下載，全部是虛構資料。", C.lime]].forEach(([h, t, c], i) => {
    const y = 1.4 + i * 1.22;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.6, y, w: 3.9, h: 1.1, fill: { color: C.soft }, rectRadius: 0.08 });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.8, y: y + 0.2, w: 0.14, h: 0.14, fill: { color: c } });
    s.addText(h, { x: 6.02, y: y + 0.12, w: 3.3, h: 0.3, fontFace: F, fontSize: 13, bold: true, color: C.ink, margin: 0, isTextBox: true });
    s.addText(t, { x: 5.8, y: y + 0.45, w: 3.5, h: 0.6, fontFace: F, fontSize: 11, color: C.muted, margin: 0, isTextBox: true, valign: "top" });
  });
}
// 5 & 21 作品格
function worksSlide(eyebrow, title, items, color) {
  const s = pres.addSlide({ masterName: "PAPER" });
  header(s, eyebrow, title, color);
  items.forEach(([img, t, url], i) => {
    const x = M + (i % 4) * 2.3, y = 1.38 + Math.floor(i / 4) * 1.95, w = 2.1;
    imgFit(s, img, x, y, w, 1.18, C.soft);
    s.addText(t, { x, y: y + 1.24, w, h: 0.3, fontFace: F, fontSize: 10.5, bold: true, color: C.ink, margin: 0, isTextBox: true, fit: "shrink" });
    s.addText(url, { x, y: y + 1.52, w, h: 0.3, fontFace: "Arial", fontSize: 7.5, color: C.blue, margin: 0, isTextBox: true, valign: "top", hyperlink: { url: "https://" + url } });
  });
}
worksSlide("精選作品", "這些 ICOPE 工具，都是用 AI 做的", [
  ["img/w_six.jpg", "六力反思查詢站", "mattsue1003.github.io/icope-team-six"],
  ["img/w_leaflets.jpg", "ICOPE 衛教單張資源庫", "icope-leaflets-tw-2026.tedandet.chatgpt.site"],
  ["img/w_follow.jpg", "評估完，怎麼讓長輩願意動", "mattsue1003.github.io/follow-order-skill"],
  ["img/w_mi.jpg", "動機式晤談", "mattsue1003.github.io/motivational-interviewing"],
  ["img/w_reablepre.jpg", "復健復能處方箋開立系統", "mattsue1003.github.io/reablepre"],
  ["img/w_bodysore.jpg", "痠痛問題影片學習資料庫", "mattsue1003.github.io/bodysoreproblemlearn"],
  ["img/w_dementia.jpg", "WHO 失智風險指南整理", "mattsue1003.github.io/dementia-risk-reduction-2026"],
  ["img/w_deid.jpg", "長照資料去識別化工作台", "mattsue1003.github.io/ltc-deid-workbench"]], C.mag);
// 6 實習生
{
  const s = pres.addSlide({ masterName: "PAPER" });
  header(s, "觀念", "把 AI 當成「打草稿很快的實習生」", C.ink);
  const hdr = (t) => ({ text: t, options: { bold: true, color: C.muted, fontSize: 10, fill: { color: C.soft } } });
  const opts = { fontFace: F, fontSize: 11, color: C.ink, border: { type: "solid", color: C.line, pt: 0.5 }, valign: "middle", margin: 0.06 };
  s.addText("交給它打草稿", { x: M, y: 1.35, w: 4.3, h: 0.3, fontFace: F, fontSize: 14, bold: true, color: C.blue, margin: 0, isTextBox: true });
  s.addTable([[hdr("擅長"), hdr("ICOPE 現場可以這樣說")], ["整理口述", "把我口述的評估結果，整理成六力表格與追蹤時間"], ["改寫語氣", "把聽力篩檢說明改成阿公阿嬤聽得懂的三句話"], ["摘要長文", "把 ICOPE 手冊『行動能力』整理成 5 點，標頁碼"], ["做素材", "參考喜歡的設計，做衛教圖、團康網頁、網頁型簡報"]], { ...opts, x: M, y: 1.7, w: 4.3, colW: [1.0, 3.3] });
  s.addText("一定要人工把關", { x: 5.2, y: 1.35, w: 4.3, h: 0.3, fontFace: F, fontSize: 14, bold: true, color: C.mag, margin: 0, isTextBox: true });
  s.addTable([[hdr("風險"), hdr("怎麼做")], ["幻覺", "先給資料，規定只能根據文件回答，沒提到就寫「文件未提及」"], ["計算", "要求列出分母與算式，自己挑一項驗算"], ["時效", "切點以國健署最新版手冊為準，請 AI 標出版本"], ["個別化", "只寫一般衛教；轉介與處置由專業人員判斷"]], { ...opts, x: 5.2, y: 1.7, w: 4.3, colW: [1.0, 3.3] });
}
// 6b 三個 AI
{
  const s = pres.addSlide({ masterName: "PAPER" });
  header(s, "今天會用到的三個 AI", "Gemini、ChatGPT、Claude：挑最順手的那個", C.ink);
  [["Gemini", "gemini.google.com", C.blue, ["跟 Google 連在一起：Chrome 側邊欄、YouTube、雲端硬碟", "建立圖像、Canvas 互動網頁", "NotebookLM 做簡報與語音摘要"], "讀 WHO 手冊、衛教圖、網頁型簡報、六力小問答"],
   ["ChatGPT", "chatgpt.com", "10A37F", ["參考圖片做衛教圖，畫完可以局部修改", "上傳 Excel，寫程式幫你算統計", "直接產出 Word、Excel、簡報檔下載"], "參考風格做衛教圖、描述性統計＋Word 報告"],
   ["Claude", "claude.ai", C.amber, ["長文整理、改寫、寫程式都很穩", "做好的網頁直接在旁邊預覽（Artifact）", "一鍵分享連結，或匯出 HTML 上架"], "參考設計做網頁型簡報（這份簡報就是用 Claude 做的）"]].forEach(([n, url, c, pts, today], i) => {
    const x = M + i * 3.07, w = 2.86;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.4, w, h: 3.2, fill: { color: "FFFFFF" }, line: { color: C.line, width: 0.75 }, rectRadius: 0.08 });
    s.addText(n, { x: x + 0.2, y: 1.55, w: w - 0.4, h: 0.45, fontFace: "Arial", fontSize: 20, bold: true, color: c, margin: 0, isTextBox: true });
    s.addText(url, { x: x + 0.2, y: 2.0, w: w - 0.4, h: 0.25, fontFace: "Arial", fontSize: 9, color: C.muted, margin: 0, isTextBox: true });
    s.addText(pts.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < pts.length - 1 } })), { x: x + 0.2, y: 2.35, w: w - 0.4, h: 1.35, fontFace: F, fontSize: 10.5, color: C.ink, margin: 0, valign: "top", isTextBox: true, paraSpaceAfter: 3 });
    s.addText([{ text: "今天做：", options: { bold: true } }, { text: today }], { x: x + 0.2, y: 3.8, w: w - 0.4, h: 0.7, fontFace: F, fontSize: 10, color: C.ink, margin: 0, valign: "top", isTextBox: true });
  });
  s.addText("三家都會出錯。上傳前先去識別化；數字、出處、醫療說法，交出去前自己核對。", { x: M, y: 4.75, w: W - 2 * M, h: 0.3, fontFace: F, fontSize: 11, bold: true, color: C.mag, margin: 0, isTextBox: true });
}
stepSlides(by("三不原則與去識別化"));
stepSlides(by("實作① 語音輸入"));
stepSlides(by("實作① 不用安裝：ChatGPT、Gemini 的麥克風"));
stepSlides(by("語音轉文字設定成繁體中文"));
stepSlides(by("幫重聽長輩：手機即時字幕"));
stepSlides(by("實作② Chrome 問問 Gemini"));
stepSlides(by("實作② 讀影片：YouTube 摘要、Gemini 與 ChatGPT"));
stepSlides(by("實作② 沒有問問 Gemini：直接丟給 Gemini、ChatGPT"));
stepSlides(by("實作② Gemini 筆記本"));
// 10 方法
{
  const d = by("方法：參考別人的 UI 風格");
  const s = pres.addSlide({ masterName: "PAPER" });
  header(s, d.eyebrow, d.h2, C.amber);
  imgFit(s, "img/ref_webtech_hero.jpg", M, 1.4, 4.9, 2.95, C.soft);
  s.addText("參考範例：米洛科技 webtech.com.tw 首頁截圖", { x: M, y: 4.42, w: 4.9, h: 0.3, fontFace: F, fontSize: 10, color: C.muted, margin: 0, isTextBox: true });
  [["找參考", "喜歡的網站、Pinterest、Dribbble、別單位的海報，截 1～2 張最有代表性的畫面。", C.mag], ["說出你看到的風格", "白底、粗黑字、彩色小方塊、圓角卡片、膠囊按鈕、數字統計列。講得越具體越像。", C.blue], ["換成你的內容，劃好界線", "「只參考配色與版面，不要畫截圖裡的人和商標」。參考風格可以，照抄別人的照片、文字、標誌不行。", C.amber]].forEach(([h, t, c], i) => {
    const y = 1.4 + i * 1.08;
    numBadge(s, i + 1, 5.7, y + 0.02, c);
    s.addText(h, { x: 6.12, y, w: 3.4, h: 0.36, fontFace: F, fontSize: 14, bold: true, color: C.ink, margin: 0, valign: "middle", isTextBox: true });
    s.addText(t, { x: 6.12, y: y + 0.38, w: 3.4, h: 0.66, fontFace: F, fontSize: 11, color: C.muted, margin: 0, valign: "top", isTextBox: true });
  });
}
stepSlides(by("實作③ ChatGPT 參考風格做衛教圖"));
stepSlides(by("實作③ Canva 魔法圖層改字改圖"));
stepSlides(by("實作③ 免費版：Gemini 建立圖像"));
stepSlides(by("實作④ 參考設計做網頁型簡報"));
stepSlides(by("實作④ Claude 做網頁型簡報"));
stepSlides(by("實作④ 做互動網頁"));
stepSlides(by("實作⑤ Gemini 做 PPT：要用「任務」"));
stepSlides(by("實作⑤ ChatGPT 工作模式：自己生圖放進簡報"));
stepSlides(by("實作⑤ Gemini 任務＋筆記本做簡報"));
stepSlides(by("實作⑤ NotebookLM 做簡報"));
stepSlides(by("實作⑤ ChatGPT＋Canva 外掛：記得檢查圖層"));
stepSlides(by("實作⑥ 指標數據描述性統計"));
stepSlides(by("實作⑥ ChatGPT 再算一次"));
// 18 驗算
{
  const d = by("驗算：AI 的數字對不對");
  const s = pres.addSlide({ masterName: "PAPER" });
  header(s, d.eyebrow, d.h2, C.mag);
  const red = (t) => ({ text: t, options: { color: C.mag, bold: true, align: "right" } });
  const ok = (t) => ({ text: t, options: { bold: true, align: "right" } });
  const hdr = (t, a) => ({ text: t, options: { bold: true, color: C.muted, fontSize: 10, fill: { color: C.soft }, align: a || "left" } });
  const grn = (t) => ({ text: t, options: { color: "10A37F", bold: true, align: "right" } });
  s.addTable([[hdr("項目"), hdr("Gemini", "right"), hdr("ChatGPT", "right"), hdr("正確答案", "right")],
    ["女／男", red("71／49"), grn("77／43"), ok("77／43")], ["平均年齡", red("76.2"), grn("80.0"), ok("80.0")], ["85 歲以上", red("23"), grn("43"), ok("43")], ["認知異常", red("51"), grn("25"), ok("25")], ["行動異常", red("58"), grn("42"), ok("42")], ["視力異常", red("46"), grn("42"), ok("42")], ["聽力異常", red("39"), grn("36"), ok("36")]],
    { x: M, y: 1.35, w: 4.6, colW: [1.3, 1.05, 1.05, 1.2], fontFace: F, fontSize: 11.5, color: C.ink, border: { type: "solid", color: C.line, pt: 0.5 }, margin: 0.06, rowH: 0.36 });
  s.addText("分母皆為 120 位初評。紅字＝不符、綠字＝正確（2026/11/11 實測）。", { x: M, y: 4.4, w: 4.6, h: 0.4, fontFace: F, fontSize: 10, color: C.muted, margin: 0, isTextBox: true });
  [["57.8%", "後測完成率＝59／102（分母：至少一項異常者）", C.mag], ["7.5→8.9", "SPPB 平均（n＝25，行動異常且完成後測）", C.blue]].forEach(([b, t, c], i) => {
    s.addText(b, { x: 5.4 + i * 2.1, y: 1.35, w: 2, h: 0.55, fontFace: "Arial", fontSize: 26, bold: true, color: c, margin: 0, isTextBox: true });
    s.addText(t, { x: 5.4 + i * 2.1, y: 1.9, w: 2, h: 0.55, fontFace: F, fontSize: 10, color: C.muted, margin: 0, isTextBox: true, valign: "top" });
  });
  promptBox(s, d.prompts[0].lab, d.prompts[0].text, 5.4, 2.6, 4.1, 1.45, 11);
  promptBox(s, d.prompts[1].lab, d.prompts[1].text, 5.4, 4.15, 4.1, 0.85, 11);
}
stepSlides(by("Napkin：文字變圖表"));
stepSlides(by("實作⑦ 上架 GitHub Pages"));
stepSlides(by("實作⑧ OpenCode：選資料夾、選模型"));
stepSlides(by("實作⑧ OpenCode 代理人做 Google 表單"));
stepSlides(by("實作⑧ OpenCode 代理人整理資料夾"));
stepSlides(by("實作⑧ OpenCode 代理人蒐集資料"));
worksSlide("延伸工具", "回去還可以試的工具與範例", [
  ["img/cv_34_order_diff.jpg", "長者遊戲簡報生成器", "mattsue1003.github.io/order-different-slides"],
  ["img/cv_36_guesswho.jpg", "猜猜我是誰（Gemini 分享）", "gemini.google.com/share/9b73f8a57909"],
  ["img/cv_38_nutrition.jpg", "營養標示 10 題挑戰", "gemini.google.com/share/647fe79eaf6c"],
  ["img/cv_19_gemini_spark.jpg", "AI 代理：OpenCode", "opencode.ai/zht"],
  ["img/cv_21_gpt_excel.jpg", "ChatGPT 工作模式", "chatgpt.com"],
  ["img/cv_25_hearing.jpg", "Canva 魔法工作室", "www.canva.com"],
  ["img/cv_18_gemini_tools.jpg", "賓果／桌遊生成器", "mattsue1003.github.io/new-bingo-generator"],
  ["img/cv_28_notebooklm.jpg", "NotebookLM 語音摘要", "notebooklm.google.com"]], C.sky);
// 22 提示詞庫（兩頁）
{
  const d = by("提示詞庫");
  [d.prompts.slice(0, 3), d.prompts.slice(3)].forEach((ps, k) => {
    const s = pres.addSlide({ masterName: "PAPER" });
    header(s, `帶回單位（${k + 1}/2）`, d.h2, C.sky);
    ps.forEach((p, i) => promptBox(s, p.lab, p.text, M, 1.35 + i * 1.25, W - 2 * M, 1.12, 12));
  });
}
// 23 課後任務
{
  const s = pres.addSlide({ masterName: "PAPER" });
  header(s, "課後任務", "兩週內，每項做一次", C.sky);
  ["用語音口述一筆評估紀錄，貼給 AI 整理（用代號）", "用 Chrome 問問 Gemini 讀一份手冊，核對它給的頁碼", "找一個喜歡的設計，請 AI 參考它做一張 A4 衛教圖", "丟一個參考網頁進去，做一份網頁型簡報或團康網頁", "把單位指標去識別化後做描述性統計，自己驗算一項", "上架一個網頁到 GitHub Pages，把網址貼到 LINE 社群"].forEach((t, i) => {
    const y = 1.4 + i * 0.6;
    s.addShape(pres.shapes.RECTANGLE, { x: M, y: y + 0.08, w: 0.28, h: 0.28, fill: { color: "FFFFFF" }, line: { color: C.ink, width: 1.25 } });
    s.addText(t, { x: M + 0.45, y, w: 8.5, h: 0.45, fontFace: F, fontSize: 14, color: C.ink, margin: 0, valign: "middle", isTextBox: true });
  });
}
// 24 結尾
{
  const s = pres.addSlide({ masterName: "PAPER" });
  s.addText("謝謝聆聽", { x: M, y: 0.6, w: 5, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: C.muted, margin: 0, isTextBox: true });
  s.addText("Happiness is not a goal…\nit's a by-product of a life well lived.", { x: M, y: 1.1, w: 4.9, h: 1.5, fontFace: "Arial", fontSize: 22, bold: true, color: C.ink, margin: 0, isTextBox: true, valign: "top" });
  s.addText("快樂不是一個目標，它是踏實人生的副產品。（Eleanor Roosevelt）", { x: M, y: 2.7, w: 4.9, h: 0.4, fontFace: F, fontSize: 12, color: C.muted, margin: 0, isTextBox: true });
  s.addText("粉絲專頁：鍾孟修Matt-職能治療師\n設計｜愛迪樂 鍾孟修 職能治療師　練習資料皆為虛擬資料", { x: M, y: 4.2, w: 5.6, h: 0.6, fontFace: F, fontSize: 10, color: C.muted, margin: 0, isTextBox: true });
  [["img/qr-fb.png", "粉絲專頁 看更多作品"], ["img/qr-line.png", "加入 LINE 學員社群"], ["img/qr-deck.png", "網頁版簡報"]].forEach(([p, t], i) => {
    const x = 5.6 + i * 1.35;
    s.addImage({ path: IMG(p), x, y: 1.4, w: 1.2, h: 1.2, altText: t });
    s.addText(t, { x: x - 0.1, y: 2.65, w: 1.4, h: 0.5, fontFace: F, fontSize: 10.5, bold: true, color: C.ink, align: "center", margin: 0, isTextBox: true });
  });
}

pres.writeFile({ fileName: path.join(ROOT, "基隆ICOPE_AI工作坊_講義.pptx") }).then((f) => console.log("wrote", f));
