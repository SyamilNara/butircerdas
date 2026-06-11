const state = {
  workbookData: null,
  results: null,
  charts: {}
};

const optionSets = {
  ABCD: ["A", "B", "C", "D"],
  ABCDE: ["A", "B", "C", "D", "E"]
};

const els = {
  examName: document.getElementById("examName"),
  subjectName: document.getElementById("subjectName"),
  questionCount: document.getElementById("questionCount"),
  optionSet: document.getElementById("optionSet"),
  validityThreshold: document.getElementById("validityThreshold"),
  downloadTemplateBtn: document.getElementById("downloadTemplateBtn"),
  excelFile: document.getElementById("excelFile"),
  messageBox: document.getElementById("messageBox"),
  previewSection: document.getElementById("previewSection"),
  detectedSummary: document.getElementById("detectedSummary"),
  previewTable: document.getElementById("previewTable"),
  analyzeBtn: document.getElementById("analyzeBtn"),
  summaryCards: document.getElementById("summaryCards"),
  studentScoreTable: document.getElementById("studentScoreTable"),
  itemAnalysisTable: document.getElementById("itemAnalysisTable"),
  ruleRecommendation: document.getElementById("ruleRecommendation"),
  aiRecommendation: document.getElementById("aiRecommendation"),
  aiRecommendationBtn: document.getElementById("aiRecommendationBtn"),
  exportExcelBtn: document.getElementById("exportExcelBtn"),
  exportPdfBtn: document.getElementById("exportPdfBtn")
};

document.addEventListener("DOMContentLoaded", () => {
  setResultTabsEnabled(false);
  els.downloadTemplateBtn.addEventListener("click", downloadTemplate);
  els.excelFile.addEventListener("change", handleExcelUpload);
  els.analyzeBtn.addEventListener("click", runAnalysis);
  els.exportExcelBtn.addEventListener("click", exportExcelReport);
  els.exportPdfBtn.addEventListener("click", exportPdfReport);
  els.aiRecommendationBtn.addEventListener("click", generateAiRecommendation);
  document.querySelectorAll(".tab-btn").forEach((button) => {
    button.addEventListener("click", () => activateTab(button.dataset.tab));
  });
});

function downloadTemplate() {
  const examName = cleanText(els.examName.value) || "Ujian";
  const subject = cleanText(els.subjectName.value) || "Mata Pelajaran";
  const count = Number(els.questionCount.value);
  const optionSet = els.optionSet.value;

  if (!Number.isInteger(count) || count < 1 || count > 200) {
    showMessage("Jumlah soal harus berupa angka 1 sampai 200.", "error");
    return;
  }

  const headers = ["No", "Nama", "Kelas", ...range(1, count).map(String)];
  const answerRows = [
    headers,
    [1, "Contoh Peserta 1", "X A", ...range(1, count).map(() => "")],
    [2, "Contoh Peserta 2", "X A", ...range(1, count).map(() => "")]
  ];

  const keyRows = [
    ["Nomor", "Kunci"],
    ...range(1, count).map((number) => [number, ""])
  ];

  const workbook = XLSX.utils.book_new();
  workbook.Props = {
    Title: examName,
    Subject: subject,
    Comments: optionSet
  };
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(answerRows), "Jawaban");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(keyRows), "Kunci");
  XLSX.writeFile(workbook, `template-butircerdas-${count}-soal.xlsx`);
  showMessage("Template Excel berhasil dibuat. Isi sheet Jawaban dan Kunci, lalu upload kembali.", "success");
  activateTab("upload");
}

function handleExcelUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (loadEvent) => {
    try {
      const data = new Uint8Array(loadEvent.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const parsed = parseWorkbook(workbook);
      state.workbookData = parsed;
      state.results = null;
      setResultTabsEnabled(false);
      renderPreview(parsed);
      els.previewSection.classList.remove("hidden");
      showMessage("File berhasil dibaca. Periksa preview, lalu klik Analisis Sekarang.", "success");
    } catch (error) {
      state.workbookData = null;
      state.results = null;
      setResultTabsEnabled(false);
      els.previewSection.classList.add("hidden");
      showMessage(error.message, "error");
    }
  };
  reader.readAsArrayBuffer(file);
}

function parseWorkbook(workbook) {
  if (!workbook.SheetNames.includes("Jawaban")) {
    throw new Error('Sheet "Jawaban" tidak ditemukan.');
  }
  if (!workbook.SheetNames.includes("Kunci")) {
    throw new Error('Sheet "Kunci" tidak ditemukan.');
  }

  const answerRows = XLSX.utils.sheet_to_json(workbook.Sheets.Jawaban, { header: 1, defval: "" });
  const keyRows = XLSX.utils.sheet_to_json(workbook.Sheets.Kunci, { header: 1, defval: "" });
  const meta = readMeta(answerRows, workbook);
  const optionList = optionSets[meta.optionSet] || optionSets.ABCDE;
  const headerIndex = findHeaderRow(answerRows, "Nama");
  const answerHeaders = answerRows[headerIndex].map(normalizeHeader);
  const nameIndex = answerHeaders.findIndex((value) => value.toLowerCase() === "nama");
  const classIndex = answerHeaders.findIndex((value) => value.toLowerCase() === "kelas");
  const numberIndex = answerHeaders.findIndex((value) => value.toLowerCase() === "no");
  const questionColumns = answerHeaders
    .map((header, index) => ({ header, index, number: Number(header) }))
    .filter((item) => Number.isInteger(item.number) && item.number > 0)
    .sort((a, b) => a.number - b.number);

  if (nameIndex === -1) throw new Error('Kolom "Nama" tidak ditemukan di sheet Jawaban.');
  if (questionColumns.length === 0) throw new Error("Kolom soal berupa angka tidak ditemukan di sheet Jawaban.");

  const keyHeaderIndex = findHeaderRow(keyRows, "Nomor");
  const keyHeaders = keyRows[keyHeaderIndex].map(normalizeHeader);
  const keyNumberIndex = keyHeaders.findIndex((value) => value.toLowerCase() === "nomor");
  const keyAnswerIndex = keyHeaders.findIndex((value) => value.toLowerCase() === "kunci");
  if (keyNumberIndex === -1 || keyAnswerIndex === -1) {
    throw new Error('Sheet Kunci harus memiliki kolom "Nomor" dan "Kunci".');
  }

  const keys = {};
  const keyNumbers = [];
  const duplicateKeys = [];
  keyRows.slice(keyHeaderIndex + 1).forEach((row) => {
    const number = Number(row[keyNumberIndex]);
    const key = cleanAnswer(row[keyAnswerIndex]);
    if (Number.isInteger(number) && number > 0) {
      if (keyNumbers.includes(number)) duplicateKeys.push(number);
      keyNumbers.push(number);
      keys[number] = key;
    }
  });

  if (duplicateKeys.length) {
    throw new Error(`Nomor soal dobel di sheet Kunci: ${[...new Set(duplicateKeys)].join(", ")}.`);
  }

  const answerQuestionSet = new Set(questionColumns.map((column) => column.number));
  const extraKeys = keyNumbers.filter((number) => !answerQuestionSet.has(number));
  if (extraKeys.length) {
    throw new Error(`Sheet Kunci memiliki nomor soal yang tidak ada di sheet Jawaban: ${extraKeys.join(", ")}.`);
  }

  if (keyNumbers.length !== questionColumns.length) {
    throw new Error(`Jumlah soal di sheet Jawaban adalah ${questionColumns.length}, tetapi jumlah nomor soal di sheet Kunci adalah ${keyNumbers.length}.`);
  }

  const missingKeys = questionColumns.filter((column) => !keys[column.number]);
  if (missingKeys.length) {
    throw new Error(`Kunci jawaban belum lengkap untuk soal: ${missingKeys.map((item) => item.number).join(", ")}.`);
  }

  const invalidKeys = Object.entries(keys).filter(([, key]) => key && !optionList.includes(key));
  if (invalidKeys.length) {
    throw new Error(`Kunci jawaban hanya boleh ${optionList.join("/")}. Periksa nomor: ${invalidKeys.map(([number]) => number).join(", ")}.`);
  }

  const students = answerRows.slice(headerIndex + 1)
    .filter((row) => row.some((cell) => cleanText(cell) !== ""))
    .map((row, rowOffset) => {
      const answers = {};
      questionColumns.forEach((column) => {
        answers[column.number] = cleanAnswer(row[column.index]);
      });
      return {
        no: numberIndex >= 0 ? cleanText(row[numberIndex]) : String(rowOffset + 1),
        name: cleanText(row[nameIndex]),
        className: classIndex >= 0 ? cleanText(row[classIndex]) : "",
        answers,
        sourceRow: headerIndex + rowOffset + 2
      };
    });

  if (!students.length) throw new Error("Data peserta belum ada di sheet Jawaban.");
  const unnamed = students.find((student) => !student.name);
  if (unnamed) throw new Error(`Nama peserta kosong pada baris ${unnamed.sourceRow}.`);

  const invalidAnswers = [];
  students.forEach((student) => {
    questionColumns.forEach((column) => {
      const answer = student.answers[column.number];
      if (answer && !optionList.includes(answer)) {
        invalidAnswers.push(`baris ${student.sourceRow} soal ${column.number}`);
      }
    });
  });
  if (invalidAnswers.length) {
    throw new Error(`Jawaban hanya boleh ${optionList.join("/")}. Periksa ${invalidAnswers.slice(0, 8).join(", ")}${invalidAnswers.length > 8 ? ", dan lainnya" : ""}.`);
  }

  return {
    examName: meta.examName,
    subject: meta.subject,
    optionSet: meta.optionSet,
    optionList,
    questionNumbers: questionColumns.map((column) => column.number),
    keys,
    students
  };
}

function readMeta(rows, workbook) {
  const getValue = (label, fallback) => {
    const row = rows.find((item) => cleanText(item[0]).toLowerCase() === label.toLowerCase());
    return row ? cleanText(row[1]) || fallback : fallback;
  };
  const props = workbook?.Props || {};
  const optionFromProps = cleanText(props.Comments).toUpperCase();
  const optionSet = (getValue("Pilihan Jawaban", optionFromProps || els.optionSet.value).toUpperCase() === "ABCD") ? "ABCD" : "ABCDE";
  return {
    examName: getValue("Nama Ujian", cleanText(props.Title) || cleanText(els.examName.value) || "Ujian"),
    subject: getValue("Mata Pelajaran", cleanText(props.Subject) || cleanText(els.subjectName.value) || "Mata Pelajaran"),
    optionSet
  };
}

function findHeaderRow(rows, requiredHeader) {
  const index = rows.findIndex((row) => row.map(normalizeHeader).includes(requiredHeader));
  if (index === -1) throw new Error(`Baris header dengan kolom "${requiredHeader}" tidak ditemukan.`);
  return index;
}

function renderPreview(data) {
  els.detectedSummary.innerHTML = [
    metricCard("Nama Ujian", data.examName),
    metricCard("Mata Pelajaran", data.subject),
    metricCard("Jumlah Peserta", data.students.length),
    metricCard("Jumlah Soal", data.questionNumbers.length),
    metricCard("Pilihan", data.optionList.join("/"))
  ].join("");

  const headers = ["No", "Nama", "Kelas", ...data.questionNumbers.map(String)];
  const rows = data.students.slice(0, 5).map((student, index) => [
    student.no || index + 1,
    student.name,
    student.className,
    ...data.questionNumbers.map((number) => student.answers[number] || "")
  ]);
  els.previewTable.innerHTML = tableHtml(headers, rows);
}

function runAnalysis() {
  if (!state.workbookData) {
    showMessage("Upload file Excel terlebih dahulu.", "error");
    return;
  }
  const results = analyzeData(state.workbookData);
  state.results = results;
  renderResults(results);
  setResultTabsEnabled(true);
  renderCharts(results);
  activateTab("summary");
  showMessage("Analisis selesai. Hasil sudah ditampilkan.", "success");
}

function analyzeData(data) {
  const questionNumbers = data.questionNumbers;
  const k = questionNumbers.length;
  const scoredStudents = data.students.map((student, index) => {
    const itemScores = {};
    let correct = 0;
    questionNumbers.forEach((number) => {
      const isCorrect = student.answers[number] === data.keys[number];
      itemScores[number] = isCorrect ? 1 : 0;
      correct += itemScores[number];
    });
    return {
      ...student,
      rankSeed: index,
      itemScores,
      correct,
      wrong: k - correct,
      score: round((correct / k) * 100, 2)
    };
  });

  const totalScores = scoredStudents.map((student) => student.correct);
  const totalVariance = populationVariance(totalScores);
  const validityThreshold = parseOptionalThreshold(els.validityThreshold?.value);
  const sorted = [...scoredStudents].sort((a, b) => b.correct - a.correct || a.rankSeed - b.rankSeed);
  const groupSize = Math.max(1, Math.round(scoredStudents.length * 0.27));
  const upper = sorted.slice(0, groupSize);
  const lower = sorted.slice(-groupSize);

  const itemAnalyses = questionNumbers.map((number) => {
    const correctCount = scoredStudents.reduce((sum, student) => sum + student.itemScores[number], 0);
    const wrongCount = scoredStudents.length - correctCount;
    const difficultyValue = correctCount / scoredStudents.length;
    const upperProp = upper.reduce((sum, student) => sum + student.itemScores[number], 0) / upper.length;
    const lowerProp = lower.reduce((sum, student) => sum + student.itemScores[number], 0) / lower.length;
    const discriminationValue = upperProp - lowerProp;
    const itemVector = scoredStudents.map((student) => student.itemScores[number]);
    const correctedTotals = scoredStudents.map((student) => student.correct - student.itemScores[number]);
    const validityValue = pearson(itemVector, correctedTotals);
    const validity = categorizeValidity(validityValue, validityThreshold);
    const distribution = {};
    data.optionList.forEach((option) => {
      distribution[option] = scoredStudents.filter((student) => student.answers[number] === option).length;
    });
    const ineffectiveDistractors = data.optionList.filter((option) => option !== data.keys[number] && distribution[option] === 0);
    const difficultyCategory = categorizeDifficulty(difficultyValue);
    const discriminationCategory = categorizeDiscrimination(discriminationValue);
    const recommendation = makeRecommendation({
      difficultyCategory,
      discriminationCategory,
      validityStatus: validity.status,
      ineffectiveDistractors
    });

    return {
      number,
      key: data.keys[number],
      correctCount,
      wrongCount,
      difficultyValue,
      difficultyCategory,
      discriminationValue,
      discriminationCategory,
      validityValue,
      validityCategory: validity.category,
      validityStatus: validity.status,
      distribution,
      ineffectiveDistractors,
      recommendation
    };
  });

  const pqSum = itemAnalyses.reduce((sum, item) => {
    const p = item.difficultyValue;
    return sum + p * (1 - p);
  }, 0);
  const reliability = k > 1 && totalVariance > 0 ? (k / (k - 1)) * (1 - (pqSum / totalVariance)) : 0;

  const values = scoredStudents.map((student) => student.score);
  return {
    meta: {
      examName: data.examName,
      subject: data.subject,
      optionList: data.optionList,
      questionCount: k,
      studentCount: scoredStudents.length,
      validityThreshold
    },
    students: scoredStudents,
    items: itemAnalyses,
    summary: {
      average: round(average(values), 2),
      highest: round(Math.max(...values), 2),
      lowest: round(Math.min(...values), 2),
      reliability: round(Math.max(-1, Math.min(1, reliability)), 3),
      reliabilityCategory: categorizeReliability(reliability)
    }
  };
}

function renderResults(results) {
  els.summaryCards.innerHTML = [
    metricCard("Jumlah Peserta", results.meta.studentCount),
    metricCard("Jumlah Soal", results.meta.questionCount),
    metricCard("Rata-rata", results.summary.average),
    metricCard("Tertinggi", results.summary.highest),
    metricCard("Terendah", results.summary.lowest),
    metricCard("Reliabilitas", `${results.summary.reliability} (${results.summary.reliabilityCategory})`)
  ].join("");

  els.studentScoreTable.innerHTML = tableHtml(
    ["No", "Nama", "Kelas", "Benar", "Salah", "Nilai"],
    results.students.map((student, index) => [
      student.no || index + 1,
      student.name,
      student.className,
      student.correct,
      student.wrong,
      student.score
    ])
  );

  els.itemAnalysisTable.innerHTML = tableHtml(
    ["No Soal", "Kunci", "Benar", "Salah", "Kesukaran", "Daya Pembeda", "Validitas", "Distraktor", "Rekomendasi"],
    results.items.map((item) => [
      item.number,
      item.key,
      item.correctCount,
      item.wrongCount,
      `${round(item.difficultyValue, 3)} (${item.difficultyCategory})`,
      `${round(item.discriminationValue, 3)} (${item.discriminationCategory})`,
      validityDisplay(item, results.meta.validityThreshold),
      distractorSummary(item, results.meta.optionList),
      item.recommendation.label
    ])
  );

  renderRuleRecommendation(results);
}

function renderRuleRecommendation(results) {
  const counts = countBy(results.items, (item) => item.recommendation.label);
  const flagged = results.items.filter((item) => item.recommendation.label !== "Dipakai");
  els.ruleRecommendation.innerHTML = `
    <h3>Rekomendasi Berbasis Rumus</h3>
    <p>Reliabilitas tes: <strong>${results.summary.reliability}</strong> (${results.summary.reliabilityCategory}).</p>
    <p>Ringkasan rekomendasi: ${Object.entries(counts).map(([label, count]) => `<strong>${count}</strong> ${escapeHtml(label)}`).join(", ")}.</p>
    <ul>
      ${flagged.length
        ? flagged.slice(0, 12).map((item) => `<li>Soal ${item.number}: ${escapeHtml(item.recommendation.reason)}</li>`).join("")
        : "<li>Semua soal berada pada kondisi yang layak dipakai menurut aturan dasar.</li>"}
    </ul>
  `;
  els.aiRecommendation.innerHTML = "Rekomendasi AI opsional akan muncul di sini setelah tombol dibuat.";
}

function renderCharts(results) {
  destroyCharts();
  const scoreBins = {
    "0-49": 0,
    "50-69": 0,
    "70-84": 0,
    "85-100": 0
  };
  results.students.forEach((student) => {
    if (student.score < 50) scoreBins["0-49"] += 1;
    else if (student.score < 70) scoreBins["50-69"] += 1;
    else if (student.score < 85) scoreBins["70-84"] += 1;
    else scoreBins["85-100"] += 1;
  });

  const difficultyCounts = countBy(results.items, (item) => item.difficultyCategory);
  const recommendationCounts = countBy(results.items, (item) => item.recommendation.label);
  const palette = ["#0f766e", "#c76a2d", "#3f6f97", "#8b5d33", "#8a3ffc"];

  state.charts.score = new Chart(document.getElementById("scoreChart"), {
    type: "bar",
    data: { labels: Object.keys(scoreBins), datasets: [{ label: "Peserta", data: Object.values(scoreBins), backgroundColor: palette[0] }] },
    options: chartOptions()
  });
  state.charts.difficulty = new Chart(document.getElementById("difficultyChart"), {
    type: "doughnut",
    data: { labels: Object.keys(difficultyCounts), datasets: [{ data: Object.values(difficultyCounts), backgroundColor: palette }] },
    options: chartOptions()
  });
  state.charts.recommendation = new Chart(document.getElementById("recommendationChart"), {
    type: "bar",
    data: { labels: Object.keys(recommendationCounts), datasets: [{ label: "Soal", data: Object.values(recommendationCounts), backgroundColor: palette[1] }] },
    options: chartOptions()
  });
  state.charts.reliability = new Chart(document.getElementById("reliabilityChart"), {
    type: "bar",
    data: { labels: [results.summary.reliabilityCategory], datasets: [{ label: "KR-20", data: [results.summary.reliability], backgroundColor: palette[2] }] },
    options: { ...chartOptions(), scales: { y: { min: 0, max: 1 } } }
  });
}

async function generateAiRecommendation() {
  if (!state.results) return;
  els.aiRecommendation.innerHTML = "Sedang membuat rekomendasi AI...";
  try {
    const payload = {
      meta: state.results.meta,
      summary: state.results.summary,
      problematicItems: state.results.items
        .filter((item) => item.recommendation.label !== "Dipakai")
        .slice(0, 30)
        .map((item) => ({
          number: item.number,
          difficulty: item.difficultyCategory,
          discrimination: item.discriminationCategory,
          validity: item.validityCategory,
          validityStatus: item.validityStatus,
          distractors: item.ineffectiveDistractors,
          recommendation: item.recommendation.label
        }))
    };
    const response = await fetch("/api/gemini-recommendation", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Rekomendasi AI belum tersedia.");
    els.aiRecommendation.innerHTML = `<h3>Rekomendasi AI</h3><p>${escapeHtml(data.recommendation).replace(/\n/g, "<br>")}</p>`;
  } catch (error) {
    els.aiRecommendation.innerHTML = "Rekomendasi AI belum tersedia, tetapi hasil analisis tetap bisa digunakan.";
  }
}

function exportExcelReport() {
  if (!state.results) return;
  const results = state.results;
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ["ButirCerdas - Ringkasan"],
    ["Nama Ujian", results.meta.examName],
    ["Mata Pelajaran", results.meta.subject],
    ["Jumlah Peserta", results.meta.studentCount],
    ["Jumlah Soal", results.meta.questionCount],
    ["Rata-rata Nilai", results.summary.average],
    ["Nilai Tertinggi", results.summary.highest],
    ["Nilai Terendah", results.summary.lowest],
    ["Reliabilitas KR-20", results.summary.reliability],
    ["Kategori Reliabilitas", results.summary.reliabilityCategory]
  ]), "Ringkasan");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ["No", "Nama", "Kelas", "Benar", "Salah", "Nilai"],
    ...results.students.map((student, index) => [student.no || index + 1, student.name, student.className, student.correct, student.wrong, student.score])
  ]), "Nilai Peserta");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ["No Soal", "Kunci", "Benar", "Salah", "Kesukaran", "Daya Pembeda", "Validitas", "Distraktor", "Rekomendasi", "Alasan"],
    ...results.items.map((item) => [
      item.number,
      item.key,
      item.correctCount,
      item.wrongCount,
      `${round(item.difficultyValue, 3)} (${item.difficultyCategory})`,
      `${round(item.discriminationValue, 3)} (${item.discriminationCategory})`,
      validityDisplay(item, results.meta.validityThreshold),
      distractorSummary(item, results.meta.optionList),
      item.recommendation.label,
      item.recommendation.reason
    ])
  ]), "Analisis Butir");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ["No Soal", "Kunci", ...results.meta.optionList],
    ...results.items.map((item) => [item.number, item.key, ...results.meta.optionList.map((option) => item.distribution[option] || 0)])
  ]), "Distraktor");
  XLSX.writeFile(workbook, "laporan-butircerdas.xlsx");
}

function exportPdfReport() {
  if (!state.results) return;
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "landscape" });
  const results = state.results;
  doc.setFontSize(16);
  doc.text("ButirCerdas - Laporan Analisis Butir Soal", 14, 16);
  doc.setFontSize(10);
  doc.text(`Nama Ujian: ${results.meta.examName}`, 14, 24);
  doc.text(`Mata Pelajaran: ${results.meta.subject}`, 14, 30);
  doc.text(`Peserta: ${results.meta.studentCount} | Soal: ${results.meta.questionCount} | Rata-rata: ${results.summary.average} | KR-20: ${results.summary.reliability} (${results.summary.reliabilityCategory})`, 14, 36);

  doc.autoTable({
    startY: 44,
    head: [["No", "Nama", "Kelas", "Benar", "Salah", "Nilai"]],
    body: results.students.map((student, index) => [student.no || index + 1, student.name, student.className, student.correct, student.wrong, student.score]),
    styles: { fontSize: 8 },
    headStyles: { fillColor: [15, 118, 110] }
  });

  doc.autoTable({
    startY: doc.lastAutoTable.finalY + 10,
    head: [["No", "Kunci", "Benar", "Salah", "Kesukaran", "Daya Pembeda", "Validitas", "Distraktor", "Rekomendasi"]],
    body: results.items.map((item) => [
      item.number,
      item.key,
      item.correctCount,
      item.wrongCount,
      `${round(item.difficultyValue, 3)} (${item.difficultyCategory})`,
      `${round(item.discriminationValue, 3)} (${item.discriminationCategory})`,
      validityDisplay(item, results.meta.validityThreshold),
      distractorSummary(item, results.meta.optionList),
      item.recommendation.label
    ]),
    styles: { fontSize: 7 },
    headStyles: { fillColor: [15, 118, 110] }
  });

  doc.autoTable({
    startY: doc.lastAutoTable.finalY + 10,
    head: [["No Soal", "Rekomendasi", "Alasan"]],
    body: results.items.map((item) => [
      item.number,
      item.recommendation.label,
      item.recommendation.reason
    ]),
    styles: { fontSize: 8, cellWidth: "wrap" },
    columnStyles: {
      0: { cellWidth: 20 },
      1: { cellWidth: 42 },
      2: { cellWidth: 210 }
    },
    headStyles: { fillColor: [199, 106, 45] }
  });
  doc.save("laporan-butircerdas.pdf");
}

function activateTab(tabName) {
  const targetButton = document.querySelector(`.tab-btn[data-tab="${tabName}"]`);
  if (targetButton?.disabled) return;
  const resultTabs = ["summary", "scores", "items", "charts", "recommendations"];
  if (resultTabs.includes(tabName) && !state.results) {
    showMessage("Upload file Excel dan klik Hitung Analisis terlebih dahulu.", "error");
    document.getElementById("templateTab")?.scrollIntoView({ behavior: "smooth", block: "start" });
    return;
  }
  document.querySelectorAll(".tab-btn").forEach((button) => button.classList.toggle("active", button.dataset.tab === tabName));
  const target = document.getElementById(`${tabName}Tab`);
  if (target) {
    target.classList.remove("hidden");
    target.scrollIntoView({ behavior: "smooth", block: "start" });
  }
  if (tabName === "charts" && state.results) {
    renderCharts(state.results);
  }
}

function setResultTabsEnabled(enabled) {
  document.querySelectorAll(".result-tab").forEach((button) => {
    button.disabled = !enabled;
  });
  ["summary", "scores", "items", "charts", "recommendations"].forEach((tabName) => {
    document.getElementById(`${tabName}Tab`)?.classList.toggle("hidden", !enabled);
  });
}

function makeRecommendation({ difficultyCategory, discriminationCategory, validityStatus, ineffectiveDistractors }) {
  const isIdeal =
    difficultyCategory === "Sedang" &&
    ["Baik", "Sangat Baik"].includes(discriminationCategory) &&
    validityStatus === "Valid";
  const difficultyProblem = difficultyCategory === "Mudah" || difficultyCategory === "Sukar";
  const discriminationProblem = ["Buruk", "Jelek"].includes(discriminationCategory);
  const validityProblem = validityStatus === "Tidak Valid";
  const distractorProblem = ineffectiveDistractors.length > 0;
  const badIndicatorCount = [difficultyProblem, discriminationProblem, validityProblem, distractorProblem].filter(Boolean).length;

  const notes = [];
  if (difficultyProblem) notes.push(`tingkat kesukaran ${difficultyCategory.toLowerCase()}`);
  if (discriminationProblem) notes.push("daya pembeda kurang");
  if (validityProblem) notes.push("validitas tidak valid");
  if (distractorProblem) notes.push(`Perbaiki pengecoh ${ineffectiveDistractors.join(", ")} karena tidak efektif`);

  let label = isIdeal ? "Dipakai" : "Revisi";
  if (discriminationProblem || validityProblem) label = "Revisi";
  if (difficultyProblem && discriminationProblem) label = "Revisi Serius";
  if (badIndicatorCount >= 3) label = "Dipertimbangkan Dibuang";

  return {
    label,
    reason: notes.length ? `Perlu diperhatikan karena ${notes.join(", ")}.` : "Indikator utama sudah baik sehingga soal dapat dipakai."
  };
}

function validityDisplay(item, threshold) {
  if (threshold !== null) {
    return `${round(item.validityValue, 3)} (${item.validityStatus}; r tabel ${threshold})`;
  }
  return `${round(item.validityValue, 3)} (${item.validityCategory})`;
}

function distractorSummary(item, optionList) {
  const counts = optionList.map((option) => `${option}: ${item.distribution[option] || 0}`).join(", ");
  const status = item.ineffectiveDistractors.length
    ? `Tidak efektif: ${item.ineffectiveDistractors.join(", ")}`
    : "Efektif";
  return `${counts}. ${status}`;
}

function categorizeDifficulty(value) {
  if (value > 0.7) return "Mudah";
  if (value <= 0.3) return "Sukar";
  return "Sedang";
}

function categorizeDiscrimination(value) {
  if (value < 0) return "Buruk";
  if (value < 0.2) return "Jelek";
  if (value >= 0.7) return "Sangat Baik";
  if (value >= 0.4) return "Baik";
  if (value >= 0.2) return "Cukup";
  return "Jelek";
}

function categorizeValidity(value, threshold) {
  if (threshold !== null) {
    return {
      category: value >= threshold ? "Valid" : "Tidak Valid",
      status: value >= threshold ? "Valid" : "Tidak Valid"
    };
  }
  if (value >= 0.4) return { category: "Tinggi", status: "Valid" };
  if (value >= 0.2) return { category: "Cukup", status: "Cukup" };
  return { category: "Rendah", status: "Tidak Valid" };
}

function categorizeReliability(value) {
  if (value >= 0.8) return "Sangat Tinggi";
  if (value >= 0.6) return "Tinggi";
  if (value >= 0.4) return "Cukup";
  if (value >= 0.2) return "Rendah";
  return "Sangat Rendah";
}

function pearson(x, y) {
  const avgX = average(x);
  const avgY = average(y);
  let numerator = 0;
  let xSum = 0;
  let ySum = 0;
  x.forEach((value, index) => {
    const dx = value - avgX;
    const dy = y[index] - avgY;
    numerator += dx * dy;
    xSum += dx * dx;
    ySum += dy * dy;
  });
  const denominator = Math.sqrt(xSum * ySum);
  return denominator === 0 ? 0 : numerator / denominator;
}

function populationVariance(values) {
  if (values.length < 1) return 0;
  const avg = average(values);
  return values.reduce((sum, value) => sum + ((value - avg) ** 2), 0) / values.length;
}

function parseOptionalThreshold(value) {
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0) return null;
  return Math.min(1, Math.max(0, round(number, 3)));
}

function average(values) {
  return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : 0;
}

function countBy(items, getKey) {
  return items.reduce((acc, item) => {
    const key = getKey(item);
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

function metricCard(label, value) {
  return `<div class="metric-card"><span>${escapeHtml(label)}</span><strong>${escapeHtml(String(value))}</strong></div>`;
}

function tableHtml(headers, rows) {
  return `
    <thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead>
    <tbody>${rows.map((row) => `<tr>${row.map((cell) => `<td>${formatCell(cell)}</td>`).join("")}</tr>`).join("")}</tbody>
  `;
}

function formatCell(value) {
  const text = escapeHtml(String(value ?? ""));
  if (["Dipakai"].includes(value)) return `<span class="badge success">${text}</span>`;
  if (["Revisi", "Cukup", "Sedang"].includes(value)) return `<span class="badge warning">${text}</span>`;
  if (["Revisi Serius", "Dipertimbangkan Dibuang"].includes(value)) return `<span class="badge danger">${text}</span>`;
  return text;
}

function showMessage(message, type = "success") {
  els.messageBox.innerHTML = `<div class="message ${type}">${escapeHtml(message)}</div>`;
}

function destroyCharts() {
  Object.values(state.charts).forEach((chart) => chart.destroy());
  state.charts = {};
}

function chartOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { position: "bottom" }
    }
  };
}

function cleanAnswer(value) {
  return cleanText(value).toUpperCase();
}

function cleanText(value) {
  return String(value ?? "").trim();
}

function normalizeHeader(value) {
  return cleanText(value);
}

function round(value, digits = 2) {
  return Number(Number(value || 0).toFixed(digits));
}

function range(start, end) {
  return Array.from({ length: end - start + 1 }, (_, index) => start + index);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
