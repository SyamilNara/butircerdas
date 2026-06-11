import fs from "node:fs/promises";
import fsSync from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, "..");

const xlsxCode = await (await fetch("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js")).text();
const appCode = await fs.readFile(path.join(root, "script.js"), "utf8");

const context = createBrowserLikeContext();
vm.createContext(context);
vm.runInContext(xlsxCode, context);
vm.runInContext(`${appCode}\nthis.__state = state;\nthis.__els = els;`, context);

const files = {
  valid5: path.join(here, "butircerdas-uji-5-soal-10-peserta.xlsx"),
  valid20: path.join(here, "butircerdas-uji-20-soal-30-peserta.xlsx"),
  invalid: path.join(here, "butircerdas-uji-salah-format.xlsx")
};

const checks = [];

testTemplateDownload();
const parsed5 = testWorkbook(files.valid5, 10, 5, "5 soal");
const parsed20 = testWorkbook(files.valid20, 30, 20, "20 soal");
testInvalidWorkbook();
testRenderAndCharts(parsed5.results);
testExports(parsed5.results);
await testGeminiFallback(parsed5.results);
testNoLoginNoDatabase();
testSourceRequirements();

console.log(JSON.stringify({ ok: true, checks }, null, 2));

function testTemplateDownload() {
  const captured = {};
  context.XLSX.writeFile = (workbook, fileName) => {
    captured.workbook = workbook;
    captured.fileName = fileName;
  };
  context.__elements.examName.value = "Uji Template";
  context.__elements.subjectName.value = "Matematika";
  context.__elements.questionCount.value = "5";
  context.__elements.optionSet.value = "ABCDE";

  context.downloadTemplate();
  assert(captured.fileName === "template-butircerdas-5-soal.xlsx", "Template Excel memakai nama file benar");
  assert(captured.workbook.SheetNames.includes("Jawaban"), 'Template punya sheet "Jawaban"');
  assert(captured.workbook.SheetNames.includes("Kunci"), 'Template punya sheet "Kunci"');
  const answerRows = context.XLSX.utils.sheet_to_json(captured.workbook.Sheets.Jawaban, { header: 1, defval: "" });
  const keyRows = context.XLSX.utils.sheet_to_json(captured.workbook.Sheets.Kunci, { header: 1, defval: "" });
  assert(answerRows[0].join("|") === "No|Nama|Kelas|1|2|3|4|5", "Header Jawaban sesuai jumlah soal 5");
  assert(keyRows[0].join("|") === "Nomor|Kunci", "Header Kunci sesuai");
  checks.push({ check: "Template bisa dibuat/download", fileName: captured.fileName });
}

function testWorkbook(filePath, expectedStudents, expectedQuestions, label) {
  const workbook = readWorkbook(filePath);
  const parsed = context.parseWorkbook(workbook);
  const results = context.analyzeData(parsed);

  assert(parsed.students.length === expectedStudents, `Jumlah peserta ${label} terdeteksi`);
  assert(parsed.questionNumbers.length === expectedQuestions, `Jumlah soal ${label} terdeteksi`);
  assert(results.students.length === expectedStudents, `Skor peserta ${label} dihitung`);
  assert(results.items.length === expectedQuestions, `Analisis butir ${label} lengkap`);
  assert(results.students.every((student) => typeof student.correct === "number" && typeof student.wrong === "number" && typeof student.score === "number"), `Skor benar/salah/nilai ${label} valid`);
  assert(results.items.every((item) => item.difficultyCategory && item.discriminationCategory && item.validityCategory), `Kesukaran, daya pembeda, validitas ${label} muncul`);
  assert(Number.isFinite(results.summary.reliability), `Reliabilitas KR-20 ${label} muncul`);
  assert(results.items.every((item) => item.distribution && item.recommendation?.label), `Distraktor dan rekomendasi ${label} muncul`);

  checks.push({
    check: `Upload/parsing dan analisis ${label}`,
    students: parsed.students.length,
    questions: parsed.questionNumbers.length,
    average: results.summary.average,
    reliability: results.summary.reliability,
    firstStudentScore: results.students[0].score
  });
  return { parsed, results };
}

function testInvalidWorkbook() {
  let message = "";
  try {
    context.parseWorkbook(readWorkbook(files.invalid));
  } catch (error) {
    message = error.message;
  }
  assert(message.includes("Nama") && message.includes("tidak ditemukan"), "File salah format memunculkan pesan validasi jelas");
  checks.push({ check: "Validasi file salah format", message });
}

function testRenderAndCharts(results) {
  context.__state.results = results;
  context.renderResults(results);
  context.setResultTabsEnabled(true);
  context.activateTab("charts");

  assert(context.__elements.summaryCards.innerHTML.includes("Reliabilitas"), "Ringkasan menampilkan reliabilitas");
  assert((context.__elements.studentScoreTable.innerHTML.match(/<tr/g) || []).length === results.students.length + 1, "Tabel nilai peserta dirender");
  assert((context.__elements.itemAnalysisTable.innerHTML.match(/<tr/g) || []).length === results.items.length + 1, "Tabel analisis butir dirender");
  assert(Object.keys(context.__state.charts).length === 4, "Empat grafik Chart.js dibuat");
  assert(context.__elements.ruleRecommendation.innerHTML.includes("Rekomendasi Berbasis Rumus"), "Rekomendasi aturan muncul");
  checks.push({ check: "Render tabel, grafik, dan rekomendasi aturan" });
}

function testExports(results) {
  context.__state.results = results;
  const excelCapture = {};
  context.XLSX.writeFile = (workbook, fileName) => {
    excelCapture.workbook = workbook;
    excelCapture.fileName = fileName;
  };
  context.exportExcelReport();
  assert(excelCapture.fileName === "laporan-butircerdas.xlsx", "Export Excel memakai nama benar");
  assert(["Ringkasan", "Nilai Peserta", "Analisis Butir", "Distraktor"].every((sheet) => excelCapture.workbook.SheetNames.includes(sheet)), "Export Excel punya 4 sheet wajib");

  const pdfCapture = { text: [], tables: [], saved: "" };
  context.window.jspdf.jsPDF = class {
    constructor() {
      this.lastAutoTable = { finalY: 44 };
    }
    setFontSize() {}
    text(value) { pdfCapture.text.push(value); }
    autoTable(config) {
      pdfCapture.tables.push(config.head[0].join("|"));
      this.lastAutoTable = { finalY: (this.lastAutoTable?.finalY || 44) + 30 };
    }
    save(fileName) { pdfCapture.saved = fileName; }
  };
  context.exportPdfReport();
  assert(pdfCapture.saved === "laporan-butircerdas.pdf", "Export PDF memakai nama benar");
  assert(pdfCapture.text.some((text) => text.includes("ButirCerdas")), "PDF berisi judul ButirCerdas");
  assert(pdfCapture.tables.some((table) => table.includes("No|Nama|Kelas|Benar|Salah|Nilai")), "PDF berisi tabel nilai peserta");
  assert(pdfCapture.tables.some((table) => table.includes("No|Kunci|Benar|Salah|Kesukaran|Daya Pembeda|Validitas|Distraktor|Rekomendasi")), "PDF berisi tabel analisis butir");
  assert(pdfCapture.tables.some((table) => table.includes("No Soal|Rekomendasi|Alasan")), "PDF berisi tabel rekomendasi");
  checks.push({ check: "Export Excel dan PDF" });
}

async function testGeminiFallback(results) {
  context.__state.results = results;
  context.fetch = async () => ({
    ok: false,
    json: async () => ({ error: "No API key" })
  });
  await context.generateAiRecommendation();
  assert(context.__elements.aiRecommendation.innerHTML === "Rekomendasi AI belum tersedia, tetapi hasil analisis tetap bisa digunakan.", "Fallback Gemini sesuai");
  checks.push({ check: "Gemini fallback tanpa API key" });
}

function testNoLoginNoDatabase() {
  const html = context.__html.toLowerCase();
  assert(!html.includes('type="password"'), "Tidak ada input password");
  assert(!html.includes("register"), "Tidak ada register");
  checks.push({ check: "Tidak ada fitur login/database permanen" });
}

function testSourceRequirements() {
  const html = context.__html;
  const css = context.__css;
  assert(html.includes("cdn.jsdelivr.net/npm/xlsx"), "CDN XLSX ada");
  assert(html.includes("cdn.jsdelivr.net/npm/chart.js"), "CDN Chart.js ada");
  assert(html.includes("cdn.jsdelivr.net/npm/jspdf"), "CDN jsPDF ada");
  assert(html.includes("jspdf-autotable"), "CDN jspdf-autotable ada");
  assert(css.includes("@media (max-width: 640px)") && css.includes("grid-template-columns: 1fr"), "CSS mobile satu kolom ada");
  assert(css.includes("overflow-x: auto"), "Tabel/tab punya scroll horizontal");
  checks.push({ check: "Syarat CDN dan responsif di source" });
}

function readWorkbook(filePath) {
  const bytes = new Uint8Array(Buffer.from(readFileSyncCompat(filePath)));
  return context.XLSX.read(bytes, { type: "array" });
}

function readFileSyncCompat(filePath) {
  return fsSync.readFileSync(filePath);
}

function createBrowserLikeContext() {
  const html = fsSync.readFileSync(path.join(root, "index.html"), "utf8");
  const css = fsSync.readFileSync(path.join(root, "style.css"), "utf8");
  const elements = {};
  const ids = [
    "examName", "subjectName", "questionCount", "optionSet", "downloadTemplateBtn", "excelFile",
    "messageBox", "previewSection", "detectedSummary", "previewTable", "analyzeBtn", "summaryCards",
    "studentScoreTable", "itemAnalysisTable", "ruleRecommendation", "aiRecommendation",
    "aiRecommendationBtn", "exportExcelBtn", "exportPdfBtn", "scoreChart", "difficultyChart",
    "recommendationChart", "reliabilityChart", "templateTab", "uploadTab", "summaryTab",
    "scoresTab", "itemsTab", "chartsTab", "recommendationsTab"
  ];
  ids.forEach((id) => { elements[id] = makeElement(id); });
  elements.examName.value = "Ujian Contoh";
  elements.subjectName.value = "Matematika";
  elements.questionCount.value = "20";
  elements.optionSet.value = "ABCDE";

  const tabs = ["template", "upload", "summary", "scores", "items", "charts", "recommendations"].map((tab, index) => ({
    ...makeElement(`${tab}Button`),
    dataset: { tab },
    disabled: index > 1,
    classList: makeClassList(index === 0 ? ["tab-btn", "active"] : ["tab-btn", ...(index > 1 ? ["result-tab"] : [])])
  }));
  const resultTabs = tabs.filter((button) => ["summary", "scores", "items", "charts", "recommendations"].includes(button.dataset.tab));
  const panels = ["template", "upload", "summary", "scores", "items", "charts", "recommendations"].map((tab) => elements[`${tab}Tab`]);

  const document = {
    addEventListener(event, callback) {
      if (event === "DOMContentLoaded") callback();
    },
    getElementById(id) {
      return elements[id] || makeElement(id);
    },
    querySelectorAll(selector) {
      if (selector === ".tab-btn") return tabs;
      if (selector === ".result-tab") return resultTabs;
      if (selector === ".tab-panel") return panels;
      if (selector === "canvas") return [elements.scoreChart, elements.difficultyChart, elements.recommendationChart, elements.reliabilityChart];
      return [];
    },
    querySelector(selector) {
      if (selector === "#aiRecommendation") return elements.aiRecommendation;
      return null;
    }
  };

  const ctx = {
    console,
    document,
    window: {},
    navigator: {},
    setTimeout,
    clearTimeout,
    __elements: elements,
    __html: html,
    __css: css
  };
  ctx.window = ctx;
  ctx.Chart = class {
    constructor(element, config) {
      this.element = element;
      this.config = config;
      this.destroyed = false;
    }
    destroy() {
      this.destroyed = true;
    }
  };
  ctx.window.jspdf = { jsPDF: class {} };
  return ctx;
}

function makeElement(id) {
  return {
    id,
    value: "",
    innerHTML: "",
    textContent: "",
    disabled: false,
    dataset: {},
    classList: makeClassList(),
    addEventListener() {},
    querySelectorAll() { return []; },
    waitFor() {}
  };
}

function makeClassList(initial = []) {
  const classes = new Set(initial);
  return {
    add(...items) { items.forEach((item) => classes.add(item)); },
    remove(...items) { items.forEach((item) => classes.delete(item)); },
    toggle(item, force) {
      if (force) classes.add(item);
      else classes.delete(item);
    },
    contains(item) { return classes.has(item); },
    toString() { return [...classes].join(" "); }
  };
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}
