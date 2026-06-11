import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.dirname(fileURLToPath(import.meta.url));

const keys5 = ["A", "B", "C", "D", "E"];
const students5 = [
  ["Ani", "X A", ["A", "B", "C", "D", "E"]],
  ["Budi", "X A", ["A", "B", "C", "D", "A"]],
  ["Cici", "X A", ["A", "B", "C", "A", "E"]],
  ["Dani", "X A", ["A", "B", "D", "D", "A"]],
  ["Eka", "X A", ["A", "C", "C", "A", "E"]],
  ["Fajar", "X A", ["B", "B", "D", "D", "A"]],
  ["Gita", "X A", ["A", "C", "D", "A", "B"]],
  ["Hani", "X A", ["B", "C", "C", "A", "A"]],
  ["Irfan", "X A", ["C", "A", "D", "B", "E"]],
  ["Joko", "X A", ["B", "C", "D", "A", "A"]]
];

const keys20 = ["A", "B", "C", "D", "E", "A", "B", "C", "D", "E", "A", "B", "C", "D", "E", "A", "B", "C", "D", "E"];
const names = [
  "Adi", "Bella", "Cahya", "Dewa", "Elin", "Farhan", "Gina", "Haris", "Indah", "Jihan",
  "Karin", "Lukman", "Mira", "Naufal", "Oki", "Putri", "Qori", "Raka", "Salsa", "Tegar",
  "Uma", "Vino", "Wulan", "Yoga", "Zahra", "Alya", "Bagas", "Citra", "Dimas", "Elsa"
];
const choices = ["A", "B", "C", "D", "E"];
const students20 = names.map((name, studentIndex) => {
  const ability = 0.92 - (studentIndex / (names.length - 1)) * 0.62;
  const answers = keys20.map((key, questionIndex) => {
    const difficultyPenalty = (questionIndex % 5) * 0.05;
    const signal = ((studentIndex * 13 + questionIndex * 17) % 100) / 100;
    const shouldCorrect = signal < Math.max(0.1, ability - difficultyPenalty);
    return shouldCorrect ? key : choices[(choices.indexOf(key) + studentIndex + questionIndex + 1) % choices.length];
  });
  return [name, `XI ${String.fromCharCode(65 + (studentIndex % 3))}`, answers];
});

await fs.mkdir(outputDir, { recursive: true });
await createValidWorkbook("butircerdas-uji-5-soal-10-peserta.xlsx", "Uji 5 Soal", "Matematika", keys5, students5);
await createValidWorkbook("butircerdas-uji-20-soal-30-peserta.xlsx", "Uji 20 Soal", "Bahasa Indonesia", keys20, students20);
await createInvalidWorkbook("butircerdas-uji-salah-format.xlsx");

async function createValidWorkbook(fileName, examName, subject, keys, students) {
  const workbook = Workbook.create();
  workbook.Props = { Title: examName, Subject: subject, Comments: "ABCDE" };
  const answerSheet = workbook.worksheets.add("Jawaban");
  const keySheet = workbook.worksheets.add("Kunci");

  const questionHeaders = keys.map((_, index) => String(index + 1));
  answerSheet.getRangeByIndexes(0, 0, students.length + 1, keys.length + 3).values = [
    ["No", "Nama", "Kelas", ...questionHeaders],
    ...students.map((student, index) => [index + 1, student[0], student[1], ...student[2]])
  ];
  keySheet.getRangeByIndexes(0, 0, keys.length + 1, 2).values = [
    ["Nomor", "Kunci"],
    ...keys.map((key, index) => [index + 1, key])
  ];
  styleSheet(answerSheet, keys.length + 3);
  styleSheet(keySheet, 2);
  await saveWorkbook(workbook, fileName);
}

async function createInvalidWorkbook(fileName) {
  const workbook = Workbook.create();
  const answerSheet = workbook.worksheets.add("Jawaban");
  const keySheet = workbook.worksheets.add("Kunci");

  answerSheet.getRange("A1:E3").values = [
    ["No", "Peserta", "Kelas", "1", "2"],
    [1, "Salah Header", "X A", "A", "B"],
    [2, "Format Rusak", "X A", "C", "D"]
  ];
  keySheet.getRange("A1:B2").values = [
    ["Nomor", "Kunci"],
    [1, "A"]
  ];
  styleSheet(answerSheet, 5);
  styleSheet(keySheet, 2);
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
