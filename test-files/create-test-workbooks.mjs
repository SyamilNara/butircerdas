import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.dirname(fileURLToPath(import.meta.url));

const matrixRows = [
  ["Ayu", 1, 1, 1, 0, 1, 1, 0, 1, 1, 1],
  ["Bima", 1, 1, 0, 0, 1, 0, 0, 1, 0, 1],
  ["Cici", 1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
  ["Dedi", 0, 1, 0, 0, 1, 0, 0, 1, 0, 0],
  ["Eka", 1, 0, 1, 0, 1, 1, 0, 0, 1, 1],
  ["Fajar", 0, 0, 0, 0, 1, 0, 0, 0, 0, 0],
  ["Gina", 1, 1, 1, 1, 1, 1, 0, 1, 1, 1],
  ["Hana", 1, 0, 1, 0, 0, 1, 0, 1, 0, 1],
  ["Indra", 0, 0, 1, 0, 0, 0, 0, 0, 1, 0],
  ["Jihan", 1, 1, 1, 1, 1, 0, 1, 1, 1, 1],
  ["Kiki", 0, 1, 0, 0, 0, 0, 0, 1, 0, 0],
  ["Lala", 1, 1, 1, 0, 1, 1, 1, 1, 0, 1],
  ["Mira", 1, 0, 0, 0, 1, 0, 0, 0, 0, 1],
  ["Nina", 1, 1, 1, 1, 1, 1, 0, 1, 1, 0],
  ["Oki", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
];

const matrix20Rows = Array.from({ length: 30 }, (_, rowIndex) => {
  const name = rowIndex === 2 ? "" : `Responden ${rowIndex + 1}`;
  const ability = 0.9 - (rowIndex / 29) * 0.65;
  const scores = Array.from({ length: 20 }, (_, itemIndex) => {
    const signal = ((rowIndex * 11 + itemIndex * 17) % 100) / 100;
    const difficulty = (itemIndex % 5) * 0.05;
    return signal < Math.max(0.1, ability - difficulty) ? 1 : 0;
  });
  return [name, ...scores];
});

await fs.mkdir(outputDir, { recursive: true });
await createWorkbook("butircerdas-uji-10-soal-15-responden.xlsx", "Uji Matriks 10 Soal", "Matematika", matrixRows);
await createWorkbook("butircerdas-uji-20-soal-30-responden.xlsx", "Uji Matriks 20 Soal", "Bahasa Indonesia", matrix20Rows);
await createInvalidWorkbook("butircerdas-uji-salah-format.xlsx");

async function createWorkbook(fileName, examName, subject, rows) {
  const workbook = Workbook.create();
  workbook.Props = { Title: examName, Subject: subject, Comments: "ButirCerdas matrix 1/0" };
  const sheet = workbook.worksheets.add("Matriks");
  const questionCount = rows[0].length - 1;
  const headers = ["Nama", ...Array.from({ length: questionCount }, (_, index) => `S${index + 1}`)];
  sheet.getRangeByIndexes(0, 0, rows.length + 1, headers.length).values = [headers, ...rows];
  styleSheet(sheet, headers.length);
  await saveWorkbook(workbook, fileName);
}

async function createInvalidWorkbook(fileName) {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("Matriks");
  sheet.getRange("A1:C3").values = [
    ["Nama", "Soal 1", "Soal 2"],
    ["Salah", "A", "B"],
    ["Format", "C", "D"]
  ];
  styleSheet(sheet, 3);
  await saveWorkbook(workbook, fileName);
}

function styleSheet(sheet, columnCount) {
  sheet.getRangeByIndexes(0, 0, 1, columnCount).format = {
    fill: "#0F766E",
    font: { bold: true, color: "#FFFFFF" }
  };
  sheet.freezePanes.freezeRows(1);
  sheet.getUsedRange().format.borders = { preset: "all", style: "thin", color: "#D9E2DF" };
  sheet.getUsedRange().format.autofitColumns();
}

async function saveWorkbook(workbook, fileName) {
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(path.join(outputDir, fileName));
}
