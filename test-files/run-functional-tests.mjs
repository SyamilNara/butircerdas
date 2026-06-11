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

await fs.mkdir(here, { recursive: true });

const files = {
  valid10: path.join(here, "butircerdas-uji-10-soal-15-responden.xlsx"),
  valid20: path.join(here, "butircerdas-uji-20-soal-30-responden.xlsx"),
  invalid: path.join(here, "butircerdas-uji-salah-format.xlsx")
};

const checks = [];

testTemplateDownload();
const result10 = testWorkbook(files.valid10, 15, 10, "10 soal");
const result20 = testWorkbook(files.valid20, 30, 20, "20 soal");
testTextInput();
testInvalidWorkbook();
testRender(result10.results);
testExports(result10.results);
testNoDistractor();

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
  context.downloadTemplate();

  assert(captured.fileName === "template-butircerdas-matriks-5-soal.xlsx", "Template matriks memakai nama file benar");
  assert(captured.workbook.SheetNames.includes("Matriks"), 'Template punya sheet "Matriks"');
  const rows = context.XLSX.utils.sheet_to_json(captured.workbook.Sheets.Matriks, { header: 1, defval: "" });
  assert(rows[0].join("|") === "Nama|S1|S2|S3|S4|S5", "Header template matriks sesuai");
  checks.push({ check: "Template matriks bisa dibuat", fileName: captured.fileName });
}

function testWorkbook(filePath, expectedRespondents, expectedItems, label) {
  const workbook = readWorkbook(filePath);
  const parsed = context.parseMatrixWorkbook(workbook);
  const results = context.analyzeMatrix(parsed);

  assert(parsed.respondents.length === expectedRespondents, `Jumlah responden ${label} benar`);
  assert(parsed.itemNumbers.length === expectedItems, `Jumlah soal ${label} benar`);
  assert(results.difficulty.length === expectedItems, `Tabel kesukaran ${label} lengkap`);
  assert(results.discrimination.length === expectedItems, `Tabel daya pembeda ${label} lengkap`);
  assert(results.validity.length === expectedItems, `Tabel validitas ${label} lengkap`);
  assert(results.reliability.items.length === expectedItems, `Tabel reliabilitas ${label} lengkap`);
  assert(Number.isFinite(results.reliability.kr20), `KR-20 ${label} dihitung`);
  assert(results.respondents.every((respondent) => typeof respondent.total === "number"), `Skor responden ${label} dihitung`);
  if (label === "20 soal") {
    assert(results.respondents.some((respondent) => respondent.name === "Responden C"), "Nama kosong otomatis menjadi Responden C");
  }

  context.renderPreview(parsed);
  assert(context.__elements.previewTable.innerHTML.includes(`<th>S${expectedItems}</th>`), `Preview ${label} menampilkan soal terakhir`);

  checks.push({
    check: `Analisis matriks ${label}`,
    respondents: parsed.respondents.length,
    items: parsed.itemNumbers.length,
    kr20: results.reliability.kr20
  });
  return { parsed, results };
}

function testInvalidWorkbook() {
  let message = "";
  try {
    context.parseMatrixWorkbook(readWorkbook(files.invalid));
  } catch (error) {
    message = error.message;
  }
  assert(message.includes("Header soal tidak ditemukan"), "File salah format memunculkan validasi jelas");
  checks.push({ check: "Validasi file salah format", message });
}

function testTextInput() {
  context.fillDataSample("medium");
  assert(context.__elements.dataTextInput.value.includes("Nama,S1,S2"), "Contoh data mengisi textarea");
  context.handleTextData();
  const parsed = context.__state.matrixData;
  assert(parsed.respondents.length === 15, "Input data contoh 10 soal membaca 15 responden");
  assert(parsed.itemNumbers.length === 10, "Input data contoh 10 soal membaca 10 butir");
  assert(context.__elements.previewTable.innerHTML.includes("<th>S10</th>"), "Preview data menampilkan soal terakhir");

  context.fillDataSample("blankNames");
  const blankParsed = context.parseTextMatrix(context.__elements.dataTextInput.value);
  assert(blankParsed.respondents[0].name === "Responden A", "Input data nama kosong menjadi Responden A");
  checks.push({ check: "Input teks dan contoh data berjalan" });
}

function testRender(results) {
  context.__state.results = results;
  context.renderResults(results);
  context.setResultSections(true);

  assert(context.__elements.summaryCards.innerHTML.includes("Reliabilitas KR-20"), "Ringkasan menampilkan reliabilitas");
  assert(context.__elements.difficultyTable.innerHTML.includes("Perhitungan P = B/N"), "Tabel kesukaran punya perhitungan");
  assert(context.__elements.discriminationTable.innerHTML.includes("Perhitungan D"), "Tabel daya pembeda punya perhitungan");
  assert(context.__elements.validityTable.innerHTML.includes("ΣXY"), "Tabel validitas punya komponen hitung");
  assert(context.__elements.reliabilityTable.innerHTML.includes("Rumus KR-20"), "Tabel reliabilitas punya rumus");
  checks.push({ check: "Render semua tabel analisis terpisah" });
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
  assert(["Ringkasan", "Skor Responden", "Kesukaran", "Daya Pembeda", "Validitas", "Reliabilitas"].every((sheet) => excelCapture.workbook.SheetNames.includes(sheet)), "Export Excel punya sheet analisis terpisah");

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
  assert(pdfCapture.tables.some((table) => table.includes("Perhitungan P")), "PDF berisi kesukaran");
  assert(pdfCapture.tables.some((table) => table.includes("Perhitungan D")), "PDF berisi daya pembeda");
  assert(pdfCapture.tables.some((table) => table.includes("ΣXY")), "PDF berisi validitas");
  assert(pdfCapture.tables.some((table) => table.includes("p × q")), "PDF berisi reliabilitas");
  checks.push({ check: "Export Excel/PDF analisis terpisah" });
}

function testNoDistractor() {
  const html = context.__html.toLowerCase();
  const js = appCode.toLowerCase();
  assert(!html.includes("distraktor"), "HTML tidak memuat distraktor");
  assert(!js.includes("distractor"), "JS tidak memuat logika distraktor");
  checks.push({ check: "Distraktor sudah dihapus" });
}

function readWorkbook(filePath) {
  const bytes = new Uint8Array(fsSync.readFileSync(filePath));
  return context.XLSX.read(bytes, { type: "array" });
}

function createBrowserLikeContext() {
  const html = fsSync.readFileSync(path.join(root, "index.html"), "utf8");
  const elements = {};
  const ids = [
    "examName", "subjectName", "questionCount", "validityThreshold", "downloadTemplateBtn",
    "matrixFile", "dataTextInput", "useDataBtn", "messageBox", "previewSection", "detectedSummary", "previewTable",
    "analyzeBtn", "summaryCards", "scoreTable", "difficultyTable", "discriminationTable",
    "validityTable", "reliabilityCards", "reliabilityTable", "groupInfo", "difficultyInsight",
    "discriminationInsight", "validityInsight", "reliabilityInsight", "exportExcelBtn", "exportPdfBtn"
  ];
  ids.forEach((id) => { elements[id] = makeElement(id); });
  elements.examName.value = "Ujian Contoh";
  elements.subjectName.value = "Matematika";
  elements.questionCount.value = "10";
  elements.validityThreshold.value = "";

  const resultNav = ["summarySection", "difficultySection", "discriminationSection", "validitySection", "reliabilitySection"].map((target) => ({
    ...makeElement(`${target}Button`),
    dataset: { target },
    disabled: true
  }));
  const resultSections = resultNav.map((button) => makeElement(button.dataset.target));
  const sampleButtons = ["mini", "medium", "blankNames"].map((sample) => ({
    ...makeElement(`${sample}SampleButton`),
    dataset: { sample }
  }));

  const document = {
    addEventListener(event, callback) {
      if (event === "DOMContentLoaded") callback();
    },
    getElementById(id) {
      return elements[id] || makeElement(id);
    },
    querySelectorAll(selector) {
      if (selector === ".nav-btn") return resultNav;
      if (selector === "[data-target]") return resultNav;
      if (selector === "[data-sample]") return sampleButtons;
      if (selector === ".result-nav") return resultNav;
      if (selector === ".result-section") return resultSections;
      return [];
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
    __html: html
  };
  ctx.window = ctx;
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
    scrollIntoView() {},
    querySelectorAll(selector) {
      if (selector === "thead th") {
        return parseCells(this.innerHTML, "th");
      }
      if (selector === "tbody tr") {
        return parseRows(this.innerHTML);
      }
      if (selector === "td") {
        return parseCells(this.innerHTML, "td");
      }
      return [];
    }
  };
}

function parseRows(html) {
  const rows = [...html.matchAll(/<tr>(.*?)<\/tr>/gs)].map((match) => makeElementFromHtml(match[1]));
  return rows.slice(1);
}

function parseCells(html, tag) {
  return [...html.matchAll(new RegExp(`<${tag}[^>]*>(.*?)<\\/${tag}>`, "gs"))].map((match) => ({
    textContent: match[1].replace(/<[^>]+>/g, "")
  }));
}

function makeElementFromHtml(html) {
  const element = makeElement("row");
  element.innerHTML = html;
  return element;
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
    contains(item) { return classes.has(item); }
  };
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}
