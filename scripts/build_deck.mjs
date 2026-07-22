import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const [outputPath, previewDir, assetDir] = process.argv.slice(2);
if (!outputPath || !previewDir || !assetDir) {
  throw new Error("Usage: node build_deck.mjs <output.pptx> <preview-dir> <asset-dir>");
}

const W = 1280;
const H = 720;
const FONT = "Microsoft JhengHei";
const C = {
  ink: "#171717",
  paper: "#F6F2E8",
  white: "#FFFFFF",
  yellow: "#FFCF33",
  coral: "#F05A47",
  teal: "#17877D",
  blue: "#275DAD",
  purple: "#8C5DB3",
  muted: "#6D6A63",
  line: "#D9D4C9",
  dark2: "#242424",
};

const p = Presentation.create({ slideSize: { width: W, height: H } });

function slide(bg = C.paper) {
  const s = p.slides.add();
  s.background.fill = bg;
  return s;
}

function box(s, x, y, w, h, fill, opts = {}) {
  return s.shapes.add({
    geometry: opts.geometry || "rect",
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: { style: "solid", fill: opts.line || "none", width: opts.lineWidth || 0 },
    ...(opts.rotation ? { rotation: opts.rotation } : {}),
  });
}

function text(s, value, x, y, w, h, opts = {}) {
  const shape = s.shapes.add({
    geometry: "textbox",
    name: opts.name,
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = value;
  shape.text.style = {
    fontSize: opts.size || 28,
    typeface: FONT,
    color: opts.color || C.ink,
    bold: Boolean(opts.bold),
    alignment: opts.align || "left",
    ...(opts.lineSpacing ? { lineSpacing: opts.lineSpacing } : {}),
  };
  return shape;
}

function kicker(s, value, color = C.ink) {
  text(s, value, 72, 46, 760, 32, { size: 18, bold: true, color });
}

function title(s, value, color = C.ink, width = 1040) {
  text(s, value, 72, 92, width, 108, { size: 56, bold: true, color, lineSpacing: 0.95 });
}

function page(s, _n, color = C.muted) {
  const actualPage = p.slides.items.length;
  text(s, String(actualPage).padStart(2, "0"), 1160, 44, 52, 28, { size: 18, bold: true, color, align: "right" });
}

function source(s, value, color = C.muted) {
  text(s, value, 72, 675, 1136, 28, { size: 13, color });
}

async function imageOn(s, file, x, y, w, h, fit = "cover", alt = "") {
  const bytes = await fs.readFile(path.join(assetDir, file));
  s.images.add({
    blob: bytes,
    contentType: "image/png",
    alt,
    fit,
    position: { left: x, top: y, width: w, height: h },
  });
}

function bullets(s, items, x, y, w, opts = {}) {
  const lines = items.map((item) => `• ${item}`).join("\n");
  return text(s, lines, x, y, w, opts.h || 300, { size: opts.size || 25, color: opts.color || C.ink, lineSpacing: opts.lineSpacing || 1.25 });
}

// 1
{
  const s = slide(C.ink); page(s, 1, "#BDB9B0"); kicker(s, "交通安全訪視經驗分享", C.white);
  text(s, "訪視可以打分，", 72, 150, 760, 90, { size: 74, bold: true, color: C.white });
  text(s, "安全不能靠打分", 72, 232, 820, 92, { size: 74, bold: true, color: C.coral });
  text(s, "從三次訪視、低效度評鑑，走向人本交通與教育現場的合理角色", 76, 355, 800, 78, { size: 27, color: C.white, lineSpacing: 1.18 });
  text(s, "竹光國中　學務主任 黃隆偉", 76, 465, 600, 40, { size: 24, color: C.white });
  box(s, 76, 610, 650, 12, C.yellow);
  box(s, 1030, 102, 130, 130, C.coral, { geometry: "ellipse" });
}

// Purpose
{
  const s = slide(C.white); page(s, 2); kicker(s, "這場分享不只是成果展示"); title(s, "今天，我想和大家一起釐清四件事");
  const goals = [
    ["01", "經驗", "三次訪視，為什麼會得到三種不同結果？"],
    ["02", "效度", "文件齊全，能代表道路更安全嗎？"],
    ["03", "目的", "交通安全應打造什麼樣的人本環境？"],
    ["04", "邊界", "教育可以做什麼，制度必須承擔什麼？"],
  ];
  const positions = [[72, 245], [654, 245], [72, 425], [654, 425]];
  goals.forEach((goal, i) => {
    const [x, y] = positions[i];
    box(s, x, y, 510, 2, C.line);
    text(s, goal[0], x, y + 24, 66, 56, { size: 38, bold: true, color: i === 3 ? C.teal : C.coral });
    text(s, goal[1], x + 82, y + 25, 390, 38, { size: 28, bold: true });
    text(s, goal[2], x + 82, y + 72, 390, 58, { size: 20, lineSpacing: 1.22 });
  });
  source(s, "希望最後，我們能把「通過訪視」與「真正降低風險」分開思考。", C.ink);
}

// 2
{
  const s = slide(C.yellow); page(s, 3); kicker(s, "先看最重要的證據"); title(s, "同一批資料，三種判決");
  const xs = [72, 426, 780];
  const big = ["第 1", "倒 1", "第 1"];
  const years = ["108 年｜甲等第一名", "111 年｜甲等倒數第一名", "114 年｜優等第一名"];
  xs.forEach((x, i) => {
    box(s, x, 260, 300, 8, i === 1 ? C.coral : C.ink);
    text(s, big[i], x, 282, 300, 100, { size: 76, bold: true, color: i === 1 ? C.coral : C.ink });
    text(s, years[i], x, 400, 300, 38, { size: 23, bold: true, color: i === 1 ? C.coral : C.ink });
  });
  text(s, "資料幾乎一樣。如果真實安全沒有劇烈變動，名次到底代表什麼？", 72, 535, 1010, 64, { size: 29, bold: true });
  source(s, "來源：使用者提供之《交通安全訪視經驗分享115.pptx》第 5–6 頁。");
}


// 114 case 1
{
  const s = slide(C.ink); page(s, 4, "#BDB9B0"); kicker(s, "114 年優等，到底做對了什麼？", C.white);
  text(s, "先把有用的", 72, 145, 820, 90, { size: 76, bold: true, color: C.white });
  text(s, "留下來", 72, 230, 760, 100, { size: 84, bold: true, color: C.coral });
  text(s, "不是教大家做更厚的成果冊，而是拆出能複製、能省時間、能真的降低風險的做法。", 76, 385, 960, 120, { size: 30, color: C.white, lineSpacing: 1.25 });
  box(s, 76, 585, 520, 12, C.yellow);
  source(s, "案例來源：竹光國中 114 年交通安全訪視網站與使用者原始簡報。", "#BDB9B0");
}

// 114 case 2
{
  const s = slide(C.white); page(s, 5); kicker(s, "先看評鑑架構，再決定資料放哪裡"); title(s, "114 訪視看的，是四條工作線");
  const items = [
    ["01", "組織計畫與宣導", "目標、分工、會議與追蹤"],
    ["02", "教學與活動", "課程、活動、能力與回饋"],
    ["03", "安全與輔導", "通學、導護、事故與接送"],
    ["04", "創新與成效", "校本問題與可檢核改變"],
  ];
  const pos = [[72,245],[654,245],[72,430],[654,430]];
  items.forEach((item,i)=>{const[x,y]=pos[i];box(s,x,y,510,2,C.line);text(s,item[0],x,y+22,72,55,{size:38,bold:true,color:i===3?C.teal:C.coral});text(s,item[1],x+86,y+23,405,38,{size:28,bold:true});text(s,item[2],x+86,y+72,405,38,{size:22,color:C.muted});});
  source(s, "來源：竹光國中交通安全訪視網站分類；原簡報第 33 頁。", C.muted);
}

// 114 case 3
{
  const s = slide(); page(s, 6); kicker(s, "效率亮點 01｜資料治理"); title(s, "資料平時進網站，訪視前不用重做", C.ink, 590);
  text(s, "同一份證據，平時就依指標歸位。訪視只是打開網站，不再臨時追照片、重排成果冊。", 72, 300, 520, 150, { size: 29, lineSpacing: 1.3 });
  text(s, "一次蒐集，多次使用。", 72, 500, 520, 52, { size: 34, bold: true, color: C.coral });
  box(s, 650, 125, 530, 430, C.white, { line: C.ink, lineWidth: 2 });
  await imageOn(s, "slide-32.png", 660, 135, 510, 410, "contain", "原簡報：平時即時蒐集並上傳資料到網站");
  source(s, "來源：使用者原簡報第 31–32 頁；竹光國中交通安全訪視網站。", C.muted);
}

// 114 case 4
{
  const s = slide(C.yellow); page(s, 7); kicker(s, "效率亮點 02｜借既有時間做深");
  text(s, "每週", 72, 145, 300, 62, { size: 48, bold: true });
  text(s, "2 堂", 72, 210, 440, 140, { size: 112, bold: true, color: C.coral });
  text(s, "既有導師時間", 72, 360, 520, 64, { size: 46, bold: true });
  text(s, "素材集中在網路，導師直接取用；不另開一門交通安全課。", 655, 225, 480, 150, { size: 38, bold: true, lineSpacing: 1.25 });
  box(s, 655, 430, 350, 10, C.ink);
  source(s, "來源：網站 4-3「學校自評特色與優點」。", C.ink);
}

// 114 case 5
{
  const s = slide(C.ink); page(s, 8, "#BDB9B0"); kicker(s, "效率亮點 03｜公民講堂素材庫", C.white); title(s, "不是每位導師各做一份教材", C.white, 590);
  text(s, "行政端整理主題資料夾，教師在導師時間選用；交通安全只是共享素材庫的一部分。", 72, 310, 520, 150, { size: 28, color: C.white, lineSpacing: 1.3 });
  box(s, 650, 120, 530, 440, C.white);
  await imageOn(s, "slide-28.png", 660, 130, 510, 420, "contain", "原簡報：公民講堂共享資料夾中的交通安全主題");
  source(s, "來源：使用者原簡報第 27–28 頁。", "#BDB9B0");
}

// 114 case 6
{
  const s = slide(C.white); page(s, 9); kicker(s, "效率亮點 04｜不另開課，借學科教風險"); title(s, "現有課程，本來就能承載交通問題");
  const rows = [
    ["數學", "相似形 × 大車視線死角"],
    ["自然", "摩擦力、速率、牛頓運動定律"],
    ["社會", "政府與人民在行政法中的角色"],
    ["英語", "移動情境與交通表達"],
  ];
  rows.forEach((row,i)=>{const y=245+i*90;box(s,72,y,1136,2,C.line);text(s,row[0],72,y+18,160,46,{size:32,bold:true,color:i===0?C.coral:C.ink});text(s,row[1],250,y+18,870,46,{size:28,bold:true});});
  source(s, "來源：網站 2-1-1、2-1-2；課程清單涵蓋七至九年級。", C.muted);
}

// 114 case 7
{
  const s = slide(C.ink); page(s, 10, "#BDB9B0"); kicker(s, "一個可直接帶走的跨域案例", C.white);
  text(s, "相似形", 72, 140, 470, 90, { size: 76, bold: true, color: C.yellow });
  text(s, "不只算比例", 72, 230, 600, 90, { size: 70, bold: true, color: C.white });
  text(s, "用因材網題目模擬大車視線死角，再搭配真實影片，讓抽象幾何變成風險判斷。", 72, 390, 610, 130, { size: 31, color: C.white, lineSpacing: 1.28 });
  text(s, "看得見你，\n你看得見我？", 790, 205, 360, 220, { size: 52, bold: true, color: C.coral, lineSpacing: 1.08 });
  source(s, "來源：網站 2-1-1；使用者原簡報第 23 頁。", "#BDB9B0");
}

// 114 case 8
{
  const s = slide(); page(s, 11); kicker(s, "效率亮點 05｜先調查，再決定教什麼"); title(s, "不是先辦活動，而是先看校本問題");
  const steps = [
    ["01", "資訊課填 Google 表單", "建立全校通學方式"],
    ["02", "量化整理", "看見接送與動線需求"],
    ["03", "SWOT 找策略", "分流、步行、導護配置"],
  ];
  steps.forEach((item,i)=>{const x=72+i*375;box(s,x,285,330,8,i===2?C.teal:C.coral);text(s,item[0],x,310,70,55,{size:38,bold:true,color:i===2?C.teal:C.coral});text(s,item[1],x,378,330,45,{size:28,bold:true});text(s,item[2],x,438,330,66,{size:22,color:C.muted,lineSpacing:1.25});});
  source(s, "來源：網站 3-1-1、1-2-1。", C.muted);
}

// 114 case 9
{
  const s = slide(C.yellow); page(s, 12); kicker(s, "校本策略｜先把動線做清楚"); title(s, "汽車、機車、步行，不搶同一條線");
  const items = [["前門長側","汽車為主"],["前門短側","機車為主"],["後門","機車為主"]];
  items.forEach((item,i)=>{const x=72+i*375;box(s,x,285,330,8,C.ink);text(s,item[0],x,320,330,50,{size:30,bold:true});text(s,item[1],x,395,330,65,{size:42,bold:true,color:i===0?C.coral:C.teal});});
  text(s, "校內再做人車分流，接送區與學生動線分開。", 72, 535,1000,48,{size:30,bold:true});
  source(s, "來源：網站 1-2-1、3-2-1、3-5-1。", C.ink);
}

// 114 case 10
{
  const s = slide(C.ink); page(s, 13, "#BDB9B0"); kicker(s, "導護亮點｜不是站了多少人，而是系統能不能接手", C.white); title(s, "四件事讓導護工作可持續", C.white);
  const labels = ["選拔", "訓練", "配備", "交接"];
  const desc = ["有計畫與名單", "會舉旗、懂風險", "背心、旗子齊備", "班表與違規紀錄"];
  labels.forEach((label,i)=>{const x=72+i*282;box(s,x,290,250,8,i===3?C.yellow:C.coral);text(s,label,x,325,250,55,{size:42,bold:true,color:C.white});text(s,desc[i],x,410,250,70,{size:23,color:"#D5D0C7",lineSpacing:1.25});});
  source(s, "來源：網站 3-3-1、3-3-2。", "#BDB9B0");
}

// 114 case 11
{
  const s = slide(C.white); page(s, 14); kicker(s, "輔導亮點｜資料要回到學生身上"); title(s, "違規資料，不只拿來處罰");
  box(s,72,265,465,285,C.paper);text(s,"立即制止",105,300,390,50,{size:36,bold:true,color:C.coral});text(s,"留下勸導紀錄\n必要時依規定處理",105,385,370,95,{size:25,lineSpacing:1.35});
  box(s,570,265,638,285,C.ink);text(s,"更重要的是",605,300,540,50,{size:34,bold:true,color:C.yellow});text(s,"了解違規原因 → 同理處境 → 教會可調整策略",605,382,540,110,{size:30,bold:true,color:C.white,lineSpacing:1.28});
  source(s, "事故熱點另整理至班親會資料。來源：網站 3-4-1、3-4-2。", C.muted);
}

// 114 case 12
{
  const s = slide(C.ink); page(s, 15, "#BDB9B0"); kicker(s, "114 最值得分享的，不是活動", C.white);
  text(s, "真正的亮點，", 72, 155, 850, 92, { size: 76, bold: true, color: C.white });
  text(s, "是讓工程發生", 72, 245, 920, 105, { size: 82, bold: true, color: C.coral });
  text(s, "學校沒有把危險的最後一哩，再丟回給學生『自己小心』。", 76, 420, 900, 80, { size: 32, color: C.white });
  box(s,76,575,610,12,C.yellow);
  source(s, "來源：網站 4-4「學校創新特色」。", "#BDB9B0");
}

// 114 case 13
{
  const s = slide(); page(s, 16); kicker(s, "工程不會只靠學校完成"); title(s, "一條通學路，需要五方坐在一起");
  const labels=["學校","交通處","教育處","里長","居民"];
  labels.forEach((label,i)=>{const x=72+i*226;box(s,x,285,200,10,i===0?C.coral:i===4?C.teal:C.ink);text(s,label,x,330,200,58,{size:34,bold:true,align:"center"});});
  text(s, "共同會勘｜協商停車｜討論標線｜確認可行方案", 72, 470, 1136, 56, { size: 32, bold: true, align: "center", color: C.coral });
  source(s, "來源：網站 4-4；112.6 至 113.5 的會勘、協調與施工紀錄。", C.muted);
}

// 114 case 14
{
  const s = slide(C.yellow); page(s, 17); kicker(s, "最後一哩，具體改了什麼？");
  const items=[["1 公尺","和平路邊線外推，挪出行走空間"],["行人號誌","提高辨識，降低誤闖紅燈風險"],["113.05.01","開始畫設標線型人行道"]];
  items.forEach((item,i)=>{const x=72+i*375;text(s,item[0],x,200,340,100,{size:i===1?54:64,bold:true,color:i===1?C.teal:C.coral});box(s,x,320,330,8,C.ink);text(s,item[1],x,355,330,100,{size:25,bold:true,lineSpacing:1.3});});
  text(s, "這些是可被現場看見的改變，不是照片數量。", 72, 550, 950, 48, { size: 30, bold: true });
  source(s, "來源：網站 4-4；僅描述工程產出，不宣稱已證明事故下降。", C.ink);
}

// 114 case 15
{
  const s = slide(C.white); page(s, 18); kicker(s, "資料與成果要分開"); title(s, "這才叫成果：風險條件被改變");
  text(s, "有文件", 72, 270, 430, 70, { size: 48, bold: true, color:C.muted });
  bullets(s,["會議紀錄","活動照片","網站頁面"],72,370,420,{size:27,h:180,color:C.muted});
  text(s, "有改變", 650, 270, 430, 70, { size: 48, bold: true, color:C.teal });
  bullets(s,["多出行走空間","號誌更可辨識","跨單位有決議與時程"],650,370,500,{size:27,h:180,color:C.ink});
  source(s, "案例來源：網站 4-4。事故與車速結果仍需另行追蹤。", C.muted);
}

// 114 case 16
{
  const s = slide(); page(s, 19); kicker(s, "教授的回饋：肯定做法，也指出下一步"); title(s, "三個優點，三個還要補強的地方");
  box(s,72,250,505,330,C.white);text(s,"看見的優點",102,282,430,46,{size:32,bold:true,color:C.teal});bullets(s,["計畫執行落實","活動融入交安議題","標誌安排於合適地點"],102,355,420,{size:25,h:180});
  box(s,610,250,598,330,C.ink);text(s,"教授的建議",642,282,500,46,{size:32,bold:true,color:C.yellow});bullets(s,["交通能力融入課程","建立交通能力指標","融入交通四守則與交通素養"],642,355,500,{size:25,color:C.white,h:190});
  source(s, "來源：使用者原簡報第 34 頁「教授的建議」。", C.muted);
}

// 114 case 17
{
  const s = slide(C.yellow); page(s, 20); kicker(s, "訪視時由誰報告？原始資料給了意外答案"); title(s, "三次訪視，三種報告者");
  const items=[["108","生教組長"],["111","校長"],["114","學務主任"]];
  items.forEach((item,i)=>{const x=72+i*375;box(s,x,285,330,8,C.ink);text(s,item[0],x,315,330,55,{size:38,bold:true,color:C.coral});text(s,item[1],x,395,330,65,{size:42,bold:true});});
  text(s, "名次不能只歸因於報告者職稱。", 72, 535, 800, 48, { size: 30, bold:true });
  source(s, "來源：使用者原簡報第 29 頁。", C.ink);
}

// 114 case 18
{
  const s = slide(C.ink); page(s, 21, "#BDB9B0"); kicker(s, "我的建議：不要迷信校長一定要上台", C.white); title(s, "誰最適合報告？", C.white);
  const items=["知道校本問題的人","找得到證據的人","敢說出制度界線的人"];
  items.forEach((item,i)=>{const y=260+i*105;box(s,72,y,1136,2,"#555555");text(s,"0"+(i+1),72,y+18,90,55,{size:40,bold:true,color:C.yellow});text(s,item,190,y+22,900,48,{size:32,bold:true,color:C.white});});
  source(s, "本頁為依三次報告者經驗提出的實務建議，不是評分規定。", "#BDB9B0");
}

// 114 case 19
{
  const s = slide(C.white); page(s, 22); kicker(s, "事半功倍的核心");
  text(s, "效率不是少做，", 72, 155, 850, 90, { size: 76, bold:true });
  text(s, "是每份資料只做一次", 72, 245, 1060, 100, { size: 78, bold:true, color:C.coral });
  text(s, "平時服務教學與安全管理，訪視時只是再次使用。", 76, 420, 920, 62, { size: 32, bold:true });
  box(s,76,570,650,12,C.yellow);
  source(s, "整理自網站即時上傳、公民講堂共用素材與校本資料流程。", C.muted);
}

// 114 case 20
{
  const s = slide(); page(s, 23); kicker(s, "可直接複製的工作流程"); title(s, "五步驟，把訪視變成平常工作副產品");
  const items=[
    ["01","列問題","通學與事故風險"],
    ["02","借時間","導師課與既有學科"],
    ["03","共用素材","一個資料庫"],
    ["04","即時歸檔","依指標放網站"],
    ["05","追工程","責任人與時程"],
  ];
  items.forEach((item,i)=>{const x=72+i*226;box(s,x,270,200,8,i===4?C.teal:C.coral);text(s,item[0],x,300,200,42,{size:30,bold:true,color:i===4?C.teal:C.coral});text(s,item[1],x,370,200,44,{size:30,bold:true});text(s,item[2],x,435,200,62,{size:22,color:C.muted,lineSpacing:1.25});});
  source(s, "案例轉譯：竹光國中 114 訪視網站。", C.muted);
}

// 114 case 21
{
  const s = slide(C.white); page(s, 24); kicker(s, "報告不需要一本書"); title(s, "一張表，只講五件事");
  const items=[["問題","哪裡有風險？"],["做法","誰做了什麼？"],["證據","如何確認做過？"],["結果","什麼條件改變？"],["下一步","誰何時接手？"]];
  items.forEach((item,i)=>{const x=72+i*226;box(s,x,280,200,250,i===3?C.yellow:C.paper);text(s,item[0],x+18,315,164,44,{size:30,bold:true,color:i===3?C.coral:C.ink});text(s,item[1],x+18,395,164,78,{size:22,bold:true,lineSpacing:1.3});});
  source(s, "建議格式：問題 → 做法 → 證據 → 結果 → 下一步。", C.muted);
}

// 114 case 22
{
  const s = slide(C.yellow); page(s, 25); kicker(s, "減量清單"); title(s, "五樣東西，不必為訪視重做");
  const items=["不另開一門交通安全課","不為每次訪視重做網站","不臨時補活動與照片","不重複向處室收同一份資料","不把工程問題包成教育成果"];
  items.forEach((item,i)=>{const y=235+i*72;box(s,72,y,1136,2,"#D7AB21");text(s,"0"+(i+1),72,y+12,70,45,{size:30,bold:true,color:C.coral});text(s,item,175,y+14,940,42,{size:28,bold:true});});
  source(s, "把時間留給真正能降低風險的工作。", C.ink);
}

// 3
{
  const s = slide(); page(s, 26); kicker(s, "訪視的測量邊界"); title(s, "評鑑測到的，可能只是資料呈現能力");
  const labels = ["投入", "流程", "產出", "結果"];
  const desc = ["人力、經費、會議、志工與行政時間", "課程、宣導、網站上傳、資料彙整", "活動場次、文件份數、照片與成果冊", "速度下降？風險降低？通學路線更安全？"];
  const xs = [72, 354, 636, 918];
  xs.forEach((x, i) => {
    box(s, x, 245, 250, 260, i === 3 ? C.yellow : C.white, { line: i === 3 ? C.yellow : C.line, lineWidth: 1.5 });
    text(s, labels[i], x + 20, 270, 210, 42, { size: 30, bold: true });
    text(s, desc[i], x + 20, 332, 210, 125, { size: 21, lineSpacing: 1.25 });
  });
  text(s, "前三者可以證明「做過」；只有結果能回答「有沒有更安全」。", 72, 545, 1030, 48, { size: 28, bold: true });
  source(s, "評估架構參考：UK Government, Magenta Book（process / impact / outcomes）。");
}

// 4
{
  const s = slide(C.white); page(s, 27); kicker(s, "不是反對評鑑，是要求評鑑回答正確問題"); title(s, "四問評鑑：它有資格代表「安全」嗎？");
  const qs = [
    ["01", "構念是什麼？", "評的是交通安全、行政完整度，還是簡報表現？"],
    ["02", "評分穩不穩？", "相同證據換一組人，結果會不會完全不同？"],
    ["03", "能連到結果嗎？", "高分是否對應更低的速度、事故與通學風險？"],
    ["04", "會不會扭曲行為？", "學校是否被迫把時間花在可拍照、可裝訂的工作？"],
  ];
  const pos = [[72,245],[654,245],[72,440],[654,440]];
  qs.forEach((q, i) => {
    const [x,y]=pos[i]; box(s,x,y,510,2,C.line); text(s,q[0],x,y+24,66,56,{size:38,bold:true,color:C.coral});
    text(s,q[1],x+82,y+25,390,38,{size:28,bold:true}); text(s,q[2],x+82,y+72,390,72,{size:20,lineSpacing:1.25});
  });
  source(s, "依三次訪視經驗提出的效度檢核；不宣稱已完成正式心理計量研究。");
}

// 5
{
  const s = slide(C.ink); page(s, 28, "#BDB9B0"); kicker(s, "教育現場已經 overloading", C.white);
  text(s, "把行政負荷當成安全證據，", 72, 120, 720, 80, { size: 54, bold: true, color: C.white });
  text(s, "成本由誰承擔？", 72, 190, 600, 80, { size: 54, bold: true, color: C.coral });
  text(s, "OECD TALIS 2024：平均約 52% 教師把過多行政工作視為工作壓力來源；行政時間增加與福祉下降的關聯尤其明顯。", 72, 310, 610, 165, { size: 25, color: C.white, lineSpacing: 1.28 });
  text(s, "新增任務前\n請先說明：\n要刪掉什麼？", 785, 205, 390, 260, { size: 48, bold: true, color: C.white, lineSpacing: 1.05 });
  box(s, 785, 485, 330, 12, C.yellow);
  source(s, "來源：OECD, Results from TALIS 2024, “The demands of teaching”.", "#BDB9B0");
}

// 6
{
  const s = slide(); page(s, 29); kicker(s, "把注意力拉回真實世界"); title(s, "即使評鑑失真，交通安全仍然真實");
  text(s, "2,858", 72, 270, 500, 120, { size: 105, bold: true, color: C.coral });
  text(s, "臺灣 2025 年道路交通事故\n30 日內死亡人數", 78, 408, 450, 90, { size: 25, lineSpacing: 1.25 });
  text(s, "119萬", 680, 270, 480, 120, { size: 105, bold: true, color: C.teal });
  text(s, "全球每年道路交通死亡；道路傷害仍是 5–29 歲主要死因", 688, 408, 430, 90, { size: 25, lineSpacing: 1.25 });
  source(s, "來源：交通部 168 交通安全入口網（2026-03-25）；WHO Global status report on road safety 2023。");
}

// 7
{
  const s = slide(C.yellow); page(s, 30); kicker(s, "重新定義交通安全教育的目的");
  text(s, "目的不是訓練永不犯錯的用路人。", 72, 130, 1020, 110, { size: 58, bold: true });
  text(s, "而是打造犯錯仍能回家的環境。", 72, 250, 1090, 150, { size: 66, bold: true, color: C.coral, lineSpacing: 0.98 });
  box(s, 72, 480, 130, 10, C.teal);
  text(s, "這就是人本交通，也是 Safe System 的倫理起點。", 72, 515, 910, 60, { size: 30, bold: true });
  source(s, "概念來源：FHWA Zero Deaths and Safe System；WHO Safe System approach。");
}

// 8
{
  const s = slide(C.ink); page(s, 31, "#BDB9B0"); kicker(s, "國際共同語言", C.white); title(s, "Safe System 的六個原則", C.white);
  const data = [
    ["死傷不可接受","目標不是少一點，而是避免死亡與重傷。"], ["人一定會犯錯","系統必須預期錯誤，而不是只責怪個人。"],
    ["人體承受有限","速度與撞擊角度必須符合人的生理極限。"], ["責任共同承擔","設計者、管理者與用路人都有責任。"],
    ["安全要主動","在死亡發生前，先找出高風險位置。"], ["需要多重防護","一道防線失敗，還有下一道保住生命。"],
  ];
  const pos=[[72,240],[472,240],[872,240],[72,425],[472,425],[872,425]];
  data.forEach((d,i)=>{const[x,y]=pos[i];box(s,x,y,340,145,C.dark2);box(s,x,y,8,145,C.yellow);text(s,d[0],x+24,y+22,285,38,{size:27,bold:true,color:C.white});text(s,d[1],x+24,y+70,285,58,{size:18,color:"#D5D0C7",lineSpacing:1.2});});
  source(s, "來源：FHWA, Zero Deaths and Safe System。", "#BDB9B0");
}

// 9
{
  const s = slide(); page(s, 32); kicker(s, "責任必須被看見"); title(s, "教育只是五道防線之一");
+  text(s, "不能獨自背負安全結果", 72, 165, 760, 52, { size: 34, bold: true, color: C.coral });
  const labels=["安全用路人","安全道路","安全速度","安全車輛","事故後照護"];
  const colors=[C.blue,C.teal,C.coral,C.purple,"#55514A"];
  labels.forEach((label,i)=>{const x=72+i*226;box(s,x,285,200,210,colors[i],{geometry:"roundRect"});text(s,label,x+18,365,164,54,{size:25,bold:true,color:C.white,align:"center"});});
  text(s, "任何一層失效，其他層都應降低後果；這才叫系統。", 72, 545, 920, 52, { size: 29, bold: true });
  source(s, "來源：FHWA Safe System elements。");
}

// 10
{
  const s = slide(C.white); page(s, 33); kicker(s, "教育的邊界");
  text(s, "只給資訊，", 72, 130, 520, 76, { size: 58, bold: true });
  text(s, "幾乎不會單獨奏效", 72, 200, 680, 76, { size: 58, bold: true, color: C.coral });
  text(s, "FHWA 的人因教材直接指出：僅提供資訊通常無法改變行為。教育要與工程、執法、制度與可行選擇一起發生。", 72, 315, 560, 180, { size: 25, lineSpacing: 1.28 });
  box(s, 720, 150, 440, 360, C.ink);
  text(s, "不要把\n工程怠惰\n的後果，\n歸咎到教育上。", 770, 205, 340, 250, { size: 42, bold: true, color: C.white, lineSpacing: 1.05 });
  box(s, 770, 480, 210, 10, C.yellow);
  source(s, "來源：FHWA Road Safety Fundamentals, Unit 2；使用者原簡報第 14 頁。");
}

// 11
{
  const s = slide(); page(s, 34); kicker(s, "在地經驗"); title(s, "環境決定人能不能做對", C.ink, 620);
  text(s, "一句「十次車禍九次快」，敵不過讓駕駛自然減速、縮短穿越距離、提升可見性的工程。", 72, 245, 500, 145, { size: 25, lineSpacing: 1.28 });
  text(s, "教育教判斷；\n環境讓正確判斷做得到。", 72, 430, 510, 110, { size: 30, bold: true, color: C.teal, lineSpacing: 1.15 });
  box(s, 644, 115, 540, 430, C.white, { line: C.ink, lineWidth: 2 });
  await imageOn(s, "slide-10.png", 654, 125, 520, 410, "contain", "原簡報中的通學環境改善前後對照");
  box(s, 672, 560, 500, 10, C.yellow);
  source(s, "案例來源：使用者原簡報第 10–11 頁。");
}

// 12
{
  const s = slide(C.white); page(s, 35); kicker(s, "國外不是更會宣導，而是更早改變系統"); title(s, "瑞典與荷蘭把安全放進設計");
  box(s,72,255,530,8,C.ink); text(s,"瑞典｜Vision Zero",72,285,510,52,{size:38,bold:true});
  text(s,"1997 年成為國家政策。倫理前提是死亡與重傷不可接受；道路系統設計者必須承擔防止嚴重後果的責任。",72,355,500,160,{size:24,lineSpacing:1.3});
  box(s,680,255,530,8,C.teal); text(s,"荷蘭｜Sustainable Safety",680,285,510,52,{size:38,bold:true,color:C.teal});
  text(s,"道路功能要清楚、速度與使用者相容、設計可預期、容錯且具限制性；安全是設計原則，不是補救活動。",680,355,500,160,{size:24,lineSpacing:1.3});
  source(s, "來源：Swedish Transport Administration / Vision Zero Academy；SWOV Sustainable Road Safety。");
}

// 13
{
  const s = slide(); page(s, 36); kicker(s, "教育仍然重要，但角色要精準"); title(s, "學校真正該教的三件事");
  const items=[
    ["保護自己","風險感知、視線死角、速度判斷、防衛性用路與不同環境的應變。",C.teal],
    ["看懂系統","辨識道路功能、穿越距離、車速、衝突點，以及設計如何影響行為。",C.yellow],
    ["參與改變","調查通學路線、提出改善方案、與公部門溝通，成為支持人本交通的公民。",C.coral],
  ];
  items.forEach((it,i)=>{const x=72+i*380;box(s,x,255,340,300,C.white);box(s,x,545,340,10,it[2]);text(s,it[0],x+28,290,280,42,{size:32,bold:true});text(s,it[1],x+28,360,280,125,{size:22,lineSpacing:1.3});});
  source(s, "內容整合：Safe System、人本交通與使用者原簡報第 20–25 頁。");
}

// 14
{
  const s = slide(C.ink); page(s, 37, "#BDB9B0"); kicker(s, "不用再新增一套大型課程", C.white); title(s, "一支影片，轉成 10 分鐘的公民課", C.white);
  const steps=[["看見","哪個位置讓人不安？誰最脆弱？"],["命名","問題是速度、視距、動線，還是責任設計？"],["重畫","如果容許人犯錯，道路要改哪一處？"],["發聲","用一張圖、一段理由，提出可回應的建議。"]];
  steps.forEach((st,i)=>{const x=72+i*282;box(s,x,280,250,250,C.dark2);box(s,x,280,250,10,C.yellow);text(s,st[0],x+22,320,205,42,{size:32,bold:true,color:C.white});text(s,st[1],x+22,385,205,90,{size:20,color:"#D5D0C7",lineSpacing:1.25});});
  source(s, "教學轉譯建議：觀察 → 系統語言 → 設計 → 公民表達。", "#BDB9B0");
}

// 15
{
  const s = slide(C.white); page(s, 38); kicker(s, "影片不是答案，是討論的起點"); title(s, "三支可直接使用的影片");
  const videos=[
    ["5′","FHWA｜The Safe System Approach","官方資源：人會犯錯、責任共享、多重防護。","highways.dot.gov/safety/zero-deaths/zero-deaths-resources"],
    ["8′","BicycleDutch｜Systematic Safety","荷蘭街道如何把 Vision Zero 變成系統設計。","youtu.be/5aNtsWvNYKE"],
    ["5′","Not Just Bikes｜Traffic Calming","觀察自我解釋的減速設計，而非只看速限牌。","youtu.be/bAxRYrpbnuA"],
  ];
  videos.forEach((v,i)=>{const y=235+i*130;box(s,72,y,1136,2,C.line);text(s,v[0],72,y+24,100,55,{size:42,bold:true,color:C.coral});text(s,v[1],190,y+23,610,36,{size:27,bold:true});text(s,v[2],190,y+64,610,30,{size:19,color:C.muted});text(s,v[3],845,y+38,335,40,{size:16,bold:true,align:"right"});});
  source(s, "非官方影片用於引發觀察與比較；政策事實請回到 FHWA、WHO、SWOV 等來源。");
}

// 16
{
  const s = slide(); page(s, 39); kicker(s, "不加課，把零散議題變成進程"); title(s, "把既有班會課排成可累積的能力");
  const terms=[["七上","我家在這裡：\n畫出通學路線與風險點"],["七下","危險路口：\n停、看、聽與視距判斷"],["八上","行路禮儀：\n權利義務與弱勢用路人"],["八下","乘車安全：\n速度、碰撞與防護"],["九上","大眾運輸：\n公共空間與移動公平"],["九下","防衛駕駛：\nSafe System 與公民提案"]];
  terms.forEach((t,i)=>{const x=72+i*188;box(s,x,270,164,265,C.white);box(s,x,270,164,8,i%2===0?C.blue:C.teal);text(s,t[0],x+16,300,130,34,{size:24,bold:true});text(s,t[1],x+16,360,130,120,{size:18,lineSpacing:1.25});});
  source(s, "改寫自使用者原簡報第 25–26 頁。");
}

// 17
{
  const s = slide(C.white); page(s, 40); kicker(s, "減量，才有可能做深"); title(s, "停止堆疊，開始聚焦");
  box(s,72,250,500,10,C.coral); text(s,"STOP",72,280,250,52,{size:40,bold:true,color:C.coral});
  bullets(s,["為訪視臨時補照片與成果冊","用口號與道德勸說責怪弱勢者","每個社會問題都新增一堂課","用活動數量代替風險變化"],72,350,500,{size:23,h:230});
  box(s,680,250,500,10,C.teal); text(s,"START",680,280,250,52,{size:40,bold:true,color:C.teal});
  bullets(s,["每年做一次通學路線風險盤點","每屆完成一個學生改善提案","跨公民、數學、科技課共用素材","同一網站即時保存真實工作"],680,350,500,{size:23,h:230});
}

// 18
{
  const s = slide(); page(s, 41); kicker(s, "共同責任不是共同甩鍋"); title(s, "學校負責學習；政府負責讓改變發生");
  box(s,72,250,430,340,C.white); text(s,"學校可以承擔",100,280,360,42,{size:31,bold:true});
  bullets(s,["學生風險感知與系統素養","通學經驗與風險點的回報","跨領域公民提案與公共溝通","校內接送管理與志工協作"],100,350,350,{size:21,h:200});
  box(s,535,250,673,340,C.ink); text(s,"主管機關／教授／交通專業必須承擔",565,280,600,42,{size:31,bold:true,color:C.white});
  bullets(s,["道路設計、速度管理與執法","風險資料、工程優先順序與改善時程","一致、可重現、有申訴機制的評估","刪減行政要求並提供人力、經費、素材","學校提報後得到正式工程回應"],565,350,590,{size:20,color:C.white,h:220});
}

// 19
{
  const s = slide(C.white); page(s, 42); kicker(s, "建議的新評估邏輯"); title(s, "評鑑要從「資料齊全」走向「安全改變」");
  const ev=[
    ["流程證據","是否建立跨局處責任、風險提報與回應機制？",C.muted],
    ["產出證據","完成哪些工程、課程與學生提案？品質如何？",C.yellow],
    ["結果證據","車速、停讓、衝突點、通學感受與系統理解是否改善？",C.teal],
    ["可信機制","評分者校準、多重證據、公開準則、回饋與申訴。",C.coral],
  ];
  ev.forEach((e,i)=>{const x=72+i*282;box(s,x,270,250,280,C.paper);box(s,x,270,250,9,e[2]);text(s,e[0],x+20,305,210,42,{size:28,bold:true});text(s,e[1],x+20,375,210,120,{size:20,lineSpacing:1.3});});
  source(s, "評估設計參考：UK Government Magenta Book；Safe System proactive safety。");
}

// 20
{
  const s = slide(C.ink); page(s, 43, "#BDB9B0"); kicker(s, "請把制度改成支持安全，而不是生產資料", C.white); title(s, "給教授與長官的三個請求", C.white);
  const asks=["少看一本成果冊，多看一條通學路。","新增任務前，先刪除等量行政負擔。","學校提出風險後，請給工程回應、時程與責任人。"];
  asks.forEach((a,i)=>{const y=255+i*120;box(s,72,y,1136,2,"#555555");text(s,`0${i+1}`,72,y+22,100,60,{size:44,bold:true,color:C.yellow});text(s,a,190,y+28,960,52,{size:31,bold:true,color:C.white});});
  source(s, "本頁為演講者依教育現場經驗提出的制度倡議。", "#BDB9B0");
}

// 21
{
  const s = slide(C.yellow); page(s, 44); kicker(s, "最後，不要只問學校做了幾份資料");
  text(s, "請問孩子走出校門時，", 72, 150, 800, 68, { size: 52, bold: true });
  text(s, "這個系統，容許他犯一次錯嗎？", 72, 228, 860, 130, { size: 58, bold: true, color: C.coral, lineSpacing: 0.98 });
  box(s, 72, 445, 630, 12, C.ink);
  box(s, 940, 150, 220, 290, C.white);
  await imageOn(s, "site-qr.png", 960, 170, 180, 180, "contain", "交通安全資料網站 QR Code");
  text(s, "案例資料網站", 955, 380, 190, 32, { size: 18, bold: true, align: "center" });
  source(s, "網站：https://sites.google.com/mail.zgjh.hc.edu.tw/safe/首頁");
}

// 22
{
  const s = slide(C.white); page(s, 45); kicker(s, "帶回去繼續看"); title(s, "延伸資源與影片");
  const res=[
    ["WHO｜Global road safety 2023","who.int/publications/b/68866"],
    ["FHWA｜Zero Deaths and Safe System","highways.dot.gov/safety/zero-deaths"],
    ["SWOV｜Sustainable Road Safety","swov.nl/.../FS Sustainable Safety.pdf"],
    ["OECD｜TALIS 2024","oecd.org/.../the-demands-of-teaching"],
    ["BicycleDutch｜Systematic Safety","youtu.be/5aNtsWvNYKE"],
    ["Not Just Bikes｜Traffic Calming","youtu.be/bAxRYrpbnuA"],
  ];
  const pos=[[72,240],[654,240],[72,360],[654,360],[72,480],[654,480]];
  res.forEach((r,i)=>{const[x,y]=pos[i];box(s,x,y,510,2,C.line);text(s,r[0],x,y+18,500,34,{size:23,bold:true});text(s,r[1],x,y+58,500,28,{size:16,color:C.muted});});
  source(s, "完整來源與使用說明見本專案 SOURCES.md。");
}

// 23
{
  const s = slide(); page(s, 46); kicker(s, "附錄｜案例證據"); title(s, "原始訪視材料保留為案例證據", C.ink, 600);
  text(s, "平時即時上傳、訪視重點、教授建議與評分表都應被保留；但它們證明的是工作歷程，不自動等於道路風險下降。", 72, 255, 500, 180, { size: 26, lineSpacing: 1.3 });
  box(s, 640, 115, 544, 445, C.white, { line: C.ink, lineWidth: 2 });
  await imageOn(s, "slide-35.png", 650, 125, 524, 425, "contain", "原簡報中的訪視評分資料");
  box(s, 670, 578, 480, 10, C.yellow);
  source(s, "案例來源：使用者原簡報第 32–36 頁。");
}

await fs.mkdir(previewDir, { recursive: true });
for (const [index, s] of p.slides.items.entries()) {
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  const png = await p.export({ slide: s, format: "png", scale: 1 });
  await fs.writeFile(path.join(previewDir, `${stem}.png`), new Uint8Array(await png.arrayBuffer()));
  const layout = await s.export({ format: "layout" });
  await fs.writeFile(path.join(previewDir, `${stem}.layout.json`), await layout.text(), "utf8");
}
const montage = await p.export({ format: "webp", montage: true, scale: 1 });
await fs.writeFile(path.join(previewDir, "deck-montage.webp"), new Uint8Array(await montage.arrayBuffer()));
const pptx = await PresentationFile.exportPptx(p);
await pptx.save(outputPath);
console.log(outputPath);
