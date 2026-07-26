import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = path.resolve(import.meta.dirname, "..");
const outputDir = path.join(
  root,
  "outputs",
  "019f98b3-1492-7f20-b3bf-3ff08a7d7916",
);
const previewDir = path.join("/tmp", "traffic-safety-toolkit-previews");
const outputPath = path.join(outputDir, "交通安全組長實戰工具包.xlsx");

const colors = {
  ink: "#171717",
  coral: "#E96552",
  yellow: "#F5D66D",
  teal: "#3C8C84",
  blue: "#355C9A",
  paper: "#F7F3EA",
  white: "#FFFFFF",
  line: "#D9D4CA",
  muted: "#6C6861",
  input: "#FFF3BF",
  paleTeal: "#DCEDEA",
  paleBlue: "#E5ECF7",
};

const workbook = Workbook.create();

function setWidths(sheet, widths) {
  for (const [column, width] of Object.entries(widths)) {
    sheet.getRange(`${column}:${column}`).format.columnWidth = width;
  }
}

function titleBand(sheet, range, text, subtitleRange, subtitle) {
  sheet.getRange(range).merge();
  sheet.getRange(range).values = [[text]];
  sheet.getRange(range).format = {
    fill: colors.ink,
    font: { bold: true, color: colors.white, size: 20 },
    verticalAlignment: "center",
  };
  sheet.getRange(range).format.rowHeight = 36;

  if (subtitleRange && subtitle) {
    sheet.getRange(subtitleRange).merge();
    sheet.getRange(subtitleRange).values = [[subtitle]];
    sheet.getRange(subtitleRange).format = {
      fill: colors.paper,
      font: { color: colors.muted, italic: true, size: 11 },
      wrapText: true,
      verticalAlignment: "center",
    };
    sheet.getRange(subtitleRange).format.rowHeight = 38;
  }
}

function styleHeader(range) {
  range.format = {
    fill: colors.teal,
    font: { bold: true, color: colors.white },
    wrapText: true,
    verticalAlignment: "center",
    horizontalAlignment: "center",
    borders: { preset: "all", style: "thin", color: colors.line },
  };
  range.format.rowHeight = 34;
}

function styleBody(range) {
  range.format = {
    font: { color: colors.ink, size: 10 },
    wrapText: true,
    verticalAlignment: "top",
    borders: {
      insideHorizontal: { style: "thin", color: colors.line },
      insideVertical: { style: "thin", color: colors.line },
      bottom: { style: "thin", color: colors.line },
    },
  };
}

function styleInput(range) {
  range.format = {
    fill: colors.input,
    font: { color: colors.ink },
    wrapText: true,
    verticalAlignment: "top",
    borders: { preset: "all", style: "thin", color: colors.line },
  };
}

function finalizeSheet(sheet, freezeRows = 5) {
  sheet.showGridLines = false;
  if (freezeRows > 0) sheet.freezePanes.freezeRows(freezeRows);
}

// 使用說明
{
  const sheet = workbook.worksheets.add("使用說明");
  titleBand(
    sheet,
    "A1:H2",
    "交通安全組長實戰工具包",
    "A3:H3",
    "目的：讓組長回到學校後，知道第一步要看什麼、問什麼、交給誰，以及如何追蹤。",
  );
  setWidths(sheet, { A: 7, B: 20, C: 21, D: 21, E: 21, F: 21, G: 21, H: 21 });

  sheet.getRange("A5:H5").merge();
  sheet.getRange("A5:H5").values = [["七個問題，就是一套工作流程"]];
  sheet.getRange("A5:H5").format = {
    fill: colors.coral,
    font: { bold: true, color: colors.white, size: 14 },
  };
  sheet.getRange("A6:H12").values = [
    ["01", "哪裡有風險？", "先走現場", "看人車衝突、視距、速度與人行缺口", "", "", "", ""],
    ["02", "怎麼確認？", "用調查與紀錄", "把感覺變成可討論的資料", "", "", "", ""],
    ["03", "現有課程能做什麼？", "借既有課程", "不另開一門課，教可遷移的風險判斷", "", "", "", ""],
    ["04", "學校先改什麼？", "處理校內機制", "分流、導護、輔導與交接", "", "", "", ""],
    ["05", "誰必須一起來？", "提出工程需求", "把學校做不到的問題交給正確單位", "", "", "", ""],
    ["06", "現場改變了什麼？", "追蹤產出與結果", "不把照片與文件直接當成安全成效", "", "", "", ""],
    ["07", "還缺什麼、誰接手？", "留下下一棒", "明確寫出責任人、時程與未完成缺口", "", "", "", ""],
  ];
  for (let row = 6; row <= 12; row += 1) {
    sheet.getRange(`D${row}:H${row}`).merge();
  }
  sheet.getRange("A6:H12").format = {
    wrapText: true,
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: colors.line },
  };
  sheet.getRange("A6:A12").format = {
    fill: colors.yellow,
    font: { bold: true, color: colors.ink, size: 13 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: colors.line },
  };
  sheet.getRange("B6:B12").format.font = { bold: true, color: colors.ink };
  sheet.getRange("C6:C12").format.fill = colors.paleTeal;
  sheet.getRange("A6:H12").format.rowHeight = 36;

  sheet.getRange("A14:H14").merge();
  sheet.getRange("A14:H14").values = [[
    "使用方式：先填「01風險盤點」；需要學生資料時複製「02通學調查」；再用 03～05 把課程、工程與追蹤接起來。黃色儲存格是主要填寫區。",
  ]];
  sheet.getRange("A14:H14").format = {
    fill: colors.input,
    font: { bold: true, color: colors.ink },
    wrapText: true,
    verticalAlignment: "center",
  };
  sheet.getRange("A14:H14").format.rowHeight = 44;
  finalizeSheet(sheet, 5);
}

// 01 風險盤點
{
  const sheet = workbook.worksheets.add("01風險盤點");
  titleBand(
    sheet,
    "A1:I2",
    "01｜30 分鐘通學風險盤點",
    "A3:I3",
    "先記錄可觀察的事實，不急著寫活動。建議在上學或放學尖峰，走一次校門周邊。",
  );
  setWidths(sheet, { A: 13, B: 12, C: 18, D: 18, E: 18, F: 30, G: 25, H: 24, I: 24 });
  const headers = [[
    "日期", "時段", "位置", "主要使用者", "風險類型",
    "看到的事實", "可能造成的後果", "學校可以先做什麼", "需要誰接手",
  ]];
  sheet.getRange("A5:I5").values = headers;
  styleHeader(sheet.getRange("A5:I5"));
  sheet.getRange("A6:I20").values = Array.from({ length: 15 }, () => Array(9).fill(null));
  styleInput(sheet.getRange("A6:I20"));
  sheet.getRange("A6:A20").format.numberFormat = "yyyy-mm-dd";
  sheet.getRange("E6:E20").dataValidation = {
    rule: { type: "list", values: ["人車交錯", "視線遮蔽", "車速", "停等空間", "人行缺口", "號誌標線", "其他"] },
  };
  sheet.getRange("A6:I20").format.rowHeight = 42;
  finalizeSheet(sheet);
}

// 02 通學調查
{
  const sheet = workbook.worksheets.add("02通學調查");
  titleBand(
    sheet,
    "A1:E2",
    "02｜六題通學調查題庫",
    "A3:E3",
    "可直接複製到 Google 表單。調查目的是確認問題，不是蒐集愈多資料愈好。",
  );
  setWidths(sheet, { A: 7, B: 38, C: 18, D: 45, E: 32 });
  sheet.getRange("A5:E5").values = [["題號", "可直接使用的題目", "題型", "建議選項／填答提示", "這題要確認什麼"]];
  styleHeader(sheet.getRange("A5:E5"));
  sheet.getRange("A6:E11").values = [
    [1, "你平常主要用什麼方式到校？", "單選", "步行／自行車／機車接送／汽車接送／公車／其他", "各種通學方式的人數與比例"],
    [2, "你通常從哪一個校門進出？", "單選", "依學校實際校門名稱設定", "不同校門的使用量與分流需求"],
    [3, "如果由家人接送，通常在哪裡上下車？", "單選＋其他", "校門長側／校門短側／後門／鄰近路口／其他", "停等位置是否和步行動線衝突"],
    [4, "上下學路上，哪一個位置讓你最不安心？", "簡答", "請寫地點與方向，例如：和平路往校門方向", "找出學生感受到的風險熱點"],
    [5, "你曾遇過哪些危險情況？", "複選", "車輛太快／視線被擋／人車擠在一起／沒有連續人行空間／號誌不清楚／其他", "辨識風險類型，不只統計事故"],
    [6, "你希望學校或政府優先改善什麼？", "簡答", "請寫出一項最希望先改變的事情", "形成改善優先順序與公民提案素材"],
  ];
  styleBody(sheet.getRange("A6:E11"));
  sheet.getRange("A6:A11").format = {
    fill: colors.yellow,
    font: { bold: true, color: colors.ink, size: 13 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: colors.line },
  };
  sheet.getRange("A6:E11").format.rowHeight = 56;
  sheet.getRange("A13:E13").merge();
  sheet.getRange("A13:E13").values = [[
    "分析時先回答三件事：哪裡最常被提到？哪一種風險最多？哪些問題不是學校單獨能處理？",
  ]];
  sheet.getRange("A13:E13").format = {
    fill: colors.input,
    font: { bold: true, color: colors.ink },
    wrapText: true,
  };
  sheet.getRange("A13:E13").format.rowHeight = 38;
  finalizeSheet(sheet);
}

// 03 課程媒合
{
  const sheet = workbook.worksheets.add("03課程媒合");
  titleBand(
    sheet,
    "A1:F2",
    "03｜把問題放回既有課程",
    "A3:F3",
    "先找現有課程可以承載的能力，不為每個社會問題再新增一堂課。",
  );
  setWidths(sheet, { A: 22, B: 24, C: 24, D: 34, E: 28, F: 30 });
  sheet.getRange("A5:F5").values = [[
    "校本問題", "學生要學會的能力", "可借用課程", "可以怎麼做", "學生產出", "不另加課的方法",
  ]];
  styleHeader(sheet.getRange("A5:F5"));
  sheet.getRange("A6:F13").values = [
    ["大車視線死角", "風險辨識、視距判斷", "數學／自然", "用相似形估算視野；比較車速、反應距離與煞停距離", "一張死角判讀圖", "使用原有相似形、速度與力學單元"],
    ["校門接送混亂", "調查、資料分析、分流思考", "資訊／班會", "用六題表單盤點通學方式與停等位置", "通學方式圖表與分流建議", "利用資訊課表單技能與既有導師時間"],
    ["人行道中斷", "系統理解、公民參與", "社會／科技", "辨識責任單位，畫出改善前後方案", "一頁式工程提案", "融入權利義務、政府職權與設計活動"],
    ["車速過快", "速度風險、人體承受極限", "自然／數學", "比較不同速度下的反應距離與碰撞後果", "速度風險說明圖", "使用原有速度、比例與統計內容"],
    [null, null, null, null, null, null],
    [null, null, null, null, null, null],
    [null, null, null, null, null, null],
    [null, null, null, null, null, null],
  ];
  styleBody(sheet.getRange("A6:F9"));
  styleInput(sheet.getRange("A10:F13"));
  sheet.getRange("A6:F13").format.rowHeight = 62;
  finalizeSheet(sheet);
}

// 04 工程提報
{
  const sheet = workbook.worksheets.add("04工程提報");
  titleBand(
    sheet,
    "A1:H2",
    "04｜一頁式道路風險提報單",
    "A3:H3",
    "學校負責把問題說清楚；工程與制度必須交給有權責的單位接手。",
  );
  setWidths(sheet, { A: 15, B: 18, C: 18, D: 18, E: 18, F: 18, G: 18, H: 18 });
  const labels = [
    ["問題名稱", "B5:H5"],
    ["位置與方向", "B6:H6"],
    ["主要發生時段", "B7:D7"],
    ["影響對象", "F7:H7"],
    ["現場可觀察事實", "B8:H9"],
    ["證據", "B10:H11"],
    ["學校已經做了什麼", "B12:H13"],
    ["為什麼學校無法單獨完成", "B14:H15"],
    ["建議接手單位", "B16:D16"],
    ["希望處理的事情", "F16:H16"],
    ["學校聯絡人", "B17:D17"],
    ["預計追蹤日期", "F17:H17"],
    ["目前狀態", "B18:H18"],
    ["外部回應與下一步", "B19:H21"],
  ];
  for (const [label, inputRange] of labels) {
    const firstCell = inputRange.split(":")[0];
    const row = firstCell.match(/\d+/)[0];
    const labelCell = `A${row}`;
    if (row === "7" || row === "16" || row === "17") {
      const secondLabel = row === "7" ? "E7" : row === "16" ? "E16" : "E17";
      if (label === "主要發生時段" || label === "建議接手單位" || label === "學校聯絡人") {
        sheet.getRange(labelCell).values = [[label]];
      } else {
        sheet.getRange(secondLabel).values = [[label]];
      }
    } else {
      sheet.getRange(labelCell).values = [[label]];
    }
    sheet.getRange(inputRange).merge();
    styleInput(sheet.getRange(inputRange));
  }
  sheet.getRange("A5:A21").format = {
    fill: colors.paleTeal,
    font: { bold: true, color: colors.ink },
    wrapText: true,
    verticalAlignment: "center",
  };
  sheet.getRange("E7:E7").format = {
    fill: colors.paleTeal,
    font: { bold: true, color: colors.ink },
    wrapText: true,
    verticalAlignment: "center",
  };
  sheet.getRange("E16:E17").format = {
    fill: colors.paleTeal,
    font: { bold: true, color: colors.ink },
    wrapText: true,
    verticalAlignment: "center",
  };
  sheet.getRange("F17:H17").format.numberFormat = "yyyy-mm-dd";
  sheet.getRange("B18:H18").dataValidation = {
    rule: { type: "list", values: ["待提報", "已提報", "會勘中", "施工中", "已完成", "暫緩"] },
  };
  sheet.getRange("A5:H21").format.rowHeight = 30;
  sheet.getRange("A8:H15").format.rowHeight = 44;
  sheet.getRange("A19:H21").format.rowHeight = 40;
  finalizeSheet(sheet, 4);
}

// 05 成果追蹤
{
  const sheet = workbook.worksheets.add("05成果追蹤");
  titleBand(
    sheet,
    "A1:J2",
    "05｜問題—做法—證據—結果—下一步追蹤表",
    "A3:J3",
    "文件證明做過；產出說明完成什麼；結果才回答是否更安全。沒有結果資料時，請誠實標示。",
  );
  setWidths(sheet, { A: 8, B: 25, C: 25, D: 25, E: 22, F: 25, G: 25, H: 20, I: 14, J: 14 });
  sheet.getRange("A5:J5").values = [[
    "編號", "問題", "學校做法", "證據", "工作產出",
    "安全結果", "尚未完成", "下一責任人", "狀態", "追蹤日期",
  ]];
  styleHeader(sheet.getRange("A5:J5"));
  const rows = Array.from(
    { length: 20 },
    (_, index) => [index + 1, null, null, null, null, null, null, null, null, null],
  );
  sheet.getRange("A6:J25").values = rows;
  styleInput(sheet.getRange("A6:J25"));
  sheet.getRange("A6:A25").format = {
    fill: colors.paleBlue,
    font: { bold: true, color: colors.blue },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: colors.line },
  };
  sheet.getRange("I6:I25").dataValidation = {
    rule: { type: "list", values: ["待處理", "進行中", "已完成", "暫緩"] },
  };
  sheet.getRange("J6:J25").format.numberFormat = "yyyy-mm-dd";
  sheet.getRange("A6:J25").format.rowHeight = 48;

  sheet.getRange("L1:M1").values = [["追蹤摘要", "數量"]];
  styleHeader(sheet.getRange("L1:M1"));
  sheet.getRange("L2:L5").values = [["全部問題"], ["已完成"], ["進行中"], ["待處理"]];
  sheet.getRange("M2").formulas = [["=COUNTIF(B6:B25,\"?*\")"]];
  sheet.getRange("M3").formulas = [["=COUNTIFS(B6:B25,\"<>\",I6:I25,\"已完成\")"]];
  sheet.getRange("M4").formulas = [["=COUNTIFS(B6:B25,\"<>\",I6:I25,\"進行中\")"]];
  sheet.getRange("M5").formulas = [["=COUNTIFS(B6:B25,\"<>\",I6:I25,\"待處理\")"]];
  sheet.getRange("L2:M5").format = {
    fill: colors.paper,
    font: { color: colors.ink },
    borders: { preset: "all", style: "thin", color: colors.line },
  };
  sheet.getRange("M2:M5").format.font = { bold: true, color: colors.coral, size: 13 };
  setWidths(sheet, { L: 14, M: 10 });
  finalizeSheet(sheet);
}

// 06 竹光案例
{
  const sheet = workbook.worksheets.add("06竹光案例");
  titleBand(
    sheet,
    "A1:H2",
    "06｜竹光國中填寫範例：標線型人行道缺口",
    "A3:H3",
    "這個範例示範如何同時寫出已完成的產出、尚未證明的結果，以及下一棒。",
  );
  setWidths(sheet, { A: 20, B: 28, C: 28, D: 28, E: 28, F: 28, G: 28, H: 28 });
  sheet.getRange("A5:H5").values = [[
    "問題", "可觀察證據", "學校先做", "共同處理", "已有產出", "尚未證明", "剩餘缺口", "下一步",
  ]];
  styleHeader(sheet.getRange("A5:H5"));
  sheet.getRange("A6:H6").values = [[
    "住戶門口標線型人行道中斷，步行動線不連續",
    "現場照片、校內分流紅箭頭、網站 4-4 會勘與施工紀錄",
    "校內人車分流、導護配置、說明通學需求",
    "學校、交通處、教育處、里長與居民共同會勘",
    "和平路邊線外推 1 公尺、改善行人號誌、部分路段畫設標線型人行道",
    "尚無事故、車速或衝突資料，因此不宣稱事故已下降",
    "住戶門口仍有暫不劃設的缺口",
    "確認責任單位、可行方案與下一次追蹤日期",
  ]];
  styleBody(sheet.getRange("A6:H6"));
  sheet.getRange("A6:H6").format.rowHeight = 118;

  sheet.getRange("A8:H8").merge();
  sheet.getRange("A8:H8").values = [[
    "判讀原則：把缺口寫出來，不是承認失敗；是讓下一步有責任人、有時程，也可以被追蹤。",
  ]];
  sheet.getRange("A8:H8").format = {
    fill: colors.yellow,
    font: { bold: true, color: colors.ink, size: 13 },
    wrapText: true,
    verticalAlignment: "center",
  };
  sheet.getRange("A8:H8").format.rowHeight = 44;

  sheet.getRange("A10:H10").merge();
  sheet.getRange("A10:H10").values = [[
    "來源：https://sites.google.com/mail.zgjh.hc.edu.tw/safe/重大創新與成效/4-4",
  ]];
  sheet.getRange("A10:H10").format = {
    font: { color: colors.muted, italic: true },
    wrapText: true,
  };
  finalizeSheet(sheet);
}

await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(previewDir, { recursive: true });

for (const sheetName of [
  "使用說明",
  "01風險盤點",
  "02通學調查",
  "03課程媒合",
  "04工程提報",
  "05成果追蹤",
  "06竹光案例",
]) {
  const preview = await workbook.render({
    sheetName,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  await fs.writeFile(
    path.join(previewDir, `${sheetName}.png`),
    new Uint8Array(await preview.arrayBuffer()),
  );
}

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);

const summary = await workbook.inspect({
  kind: "sheet,table",
  include: "id,name",
  maxChars: 6000,
  tableMaxRows: 4,
  tableMaxCols: 8,
});
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});

console.log(summary.ndjson);
console.log(errors.ndjson);
console.log(outputPath);
console.log(previewDir);
