const DATA_SAMPLES = {
  mini: `
Nama,S1,S2,S3,S4,S5
Raka,1,1,0,1,1
Salsa,1,0,1,1,0
Naufal,0,1,1,0,1
Tania,1,1,1,1,1
Bagas,0,0,1,0,1
Meisya,1,0,0,1,0
Reno,1,1,1,0,0
Dinda,0,1,0,1,1
  `,
  medium: `
Nama,S1,S2,S3,S4,S5,S6,S7,S8,S9,S10
Raka,1,1,0,1,1,0,1,1,0,1
Salsa,1,0,1,1,0,1,1,0,1,1
Naufal,0,1,1,0,1,1,0,1,1,0
Tania,1,1,1,1,1,1,1,1,1,0
Bagas,0,0,1,0,1,0,1,0,0,1
Meisya,1,0,0,1,0,1,0,1,0,0
Reno,1,1,1,0,0,1,1,1,0,1
Dinda,0,1,0,1,1,0,0,1,1,1
Arman,1,0,1,1,1,0,1,0,1,0
Putri,1,1,1,0,1,1,0,1,0,1
Yusuf,0,1,0,0,1,0,1,1,0,0
Clara,1,1,0,1,0,1,1,0,1,1
Rizky,0,0,0,1,1,0,0,1,0,1
Nabila,1,1,1,1,0,1,0,1,1,1
Farhan,0,0,1,0,0,0,1,0,1,0
  `,
  blankNames: `
Nama,S1,S2,S3,S4,S5,S6
,1,1,0,1,1,0
,1,0,0,1,0,1
,0,1,1,0,1,1
,1,1,1,1,1,1
,0,0,1,0,0,1
  `
};

const state = {
  matrixData: null,
  results: null
};

const els = {
  examName: document.getElementById("examName"),
  subjectName: document.getElementById("subjectName"),
  questionCount: document.getElementById("questionCount"),
  validityThreshold: document.getElementById("validityThreshold"),
  downloadTemplateBtn: document.getElementById("downloadTemplateBtn"),
  matrixFile: document.getElementById("matrixFile"),
  dataTextInput: document.getElementById("dataTextInput"),
  useDataBtn: document.getElementById("useDataBtn"),
  messageBox: document.getElementById("messageBox"),
  previewSection: document.getElementById("previewSection"),
  detectedSummary: document.getElementById("detectedSummary"),
  previewTable: document.getElementById("previewTable"),
  analyzeBtn: document.getElementById("analyzeBtn"),
  summaryCards: document.getElementById("summaryCards"),
  scoreTable: document.getElementById("scoreTable"),
  difficultyTable: document.getElementById("difficultyTable"),
  discriminationTable: document.getElementById("discriminationTable"),
  validityTable: document.getElementById("validityTable"),
  reliabilityCards: document.getElementById("reliabilityCards"),
  reliabilityTable: document.getElementById("reliabilityTable"),
  groupInfo: document.getElementById("groupInfo"),
  difficultyInsight: document.getElementById("difficultyInsight"),
  discriminationInsight: document.getElementById("discriminationInsight"),
  validityInsight: document.getElementById("validityInsight"),
  reliabilityInsight: document.getElementById("reliabilityInsight"),
  exportExcelBtn: document.getElementById("exportExcelBtn"),
  exportPdfBtn: document.getElementById("exportPdfBtn")
};

document.addEventListener("DOMContentLoaded", () => {
  setResultSections(false);
  els.downloadTemplateBtn.addEventListener("click", downloadTemplate);
  els.matrixFile.addEventListener("change", handleMatrixUpload);
  els.useDataBtn.addEventListener("click", handleTextData);
  els.analyzeBtn.addEventListener("click", runAnalysis);
  els.exportExcelBtn.addEventListener("click", exportExcelReport);
  els.exportPdfBtn.addEventListener("click", exportPdfReport);
  document.querySelectorAll("[data-target]").forEach((button) => {
    button.addEventListener("click", () => {
      if (scrollToSection(button.dataset.target)) {
        setActiveNav(button.dataset.target);
      }
    });
  });
  document.querySelectorAll("[data-sample]").forEach((button) => {
    button.addEventListener("click", () => fillDataSample(button.dataset.sample));
  });
  setupSectionObserver();
});

function downloadTemplate() {
  const examName = cleanText(els.examName.value) || "Ujian";
  const subject = cleanText(els.subjectName.value) || "Mata Pelajaran";
  const count = Number(els.questionCount.value);

  if (!Number.isInteger(count) || count < 1 || count > 200) {
    showMessage("Jumlah soal harus berupa angka 1 sampai 200.", "error");
    return;
  }

  const headers = ["Nama", ...range(1, count).map((number) => `S${number}`)];
  const sampleRows = [
    ["Responden A", ...range(1, count).map((number) => (number % 4 === 0 ? 0 : 1))],
    ["Responden B", ...range(1, count).map((number) => (number % 3 === 0 ? 0 : 1))],
    ["", ...range(1, count).map((number) => (number % 5 === 0 ? 0 : 1))]
  ];

  const workbook = XLSX.utils.book_new();
  workbook.Props = {
    Title: examName,
    Subject: subject,
    Comments: "ButirCerdas matrix 1/0"
  };
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([headers, ...sampleRows]), "Matriks");
  XLSX.writeFile(workbook, `template-butircerdas-matriks-${count}-soal.xlsx`);
  showMessage("Template matriks berhasil dibuat. Isi nilai 1 untuk benar dan 0 untuk salah.", "success");
}

function handleMatrixUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (loadEvent) => {
    try {
      const data = new Uint8Array(loadEvent.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const parsed = parseMatrixWorkbook(workbook);
      state.matrixData = parsed;
      state.results = null;
      setResultSections(false);
      renderPreview(parsed);
      els.previewSection.classList.remove("hidden");
      showMessage("File berhasil dibaca. Periksa matriks, lalu klik Hitung Analisis.", "success");
    } catch (error) {
      state.matrixData = null;
      state.results = null;
      setResultSections(false);
      els.previewSection.classList.add("hidden");
      showMessage(error.message, "error");
    }
  };
  reader.readAsArrayBuffer(file);
}

function handleTextData() {
  try {
    const parsed = parseTextMatrix(els.dataTextInput.value);
    state.matrixData = parsed;
    state.results = null;
    setResultSections(false);
    renderPreview(parsed);
    els.previewSection.classList.remove("hidden");
    showMessage("Data berhasil dibaca. Periksa preview, lalu klik Hitung Analisis.", "success");
  } catch (error) {
    state.matrixData = null;
    state.results = null;
    setResultSections(false);
    els.previewSection.classList.add("hidden");
    showMessage(error.message, "error");
  }
}

function fillDataSample(sampleKey) {
  const sample = DATA_SAMPLES[sampleKey] || DATA_SAMPLES.medium;
  els.dataTextInput.value = sample.trim();
  document.querySelectorAll("[data-sample]").forEach((button) => {
    button.classList.toggle("active", button.dataset.sample === sampleKey);
  });
  showMessage("Contoh data sudah dimasukkan. Klik Gunakan Data untuk membaca data.", "success");
}

function parseTextMatrix(text) {
  const content = cleanText(text);
  if (!content) {
    throw new Error("Input data masih kosong. Masukkan data dengan format Nama,S1,S2,S3,...");
  }

  const workbook = XLSX.read(content, { type: "string", raw: true });
  workbook.Props = {
    Title: cleanText(els.examName.value) || "Data",
    Subject: cleanText(els.subjectName.value) || "Mata Pelajaran"
  };
  return parseMatrixWorkbook(workbook);
}

function parseMatrixWorkbook(workbook) {
  const sheetName = workbook.SheetNames.includes("Matriks") ? "Matriks" : workbook.SheetNames[0];
  if (!sheetName) throw new Error("File tidak memiliki sheet.");

  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
  const headerIndex = rows.findIndex((row) => row.some((cell) => /^s\d+$/i.test(cleanText(cell))));
  if (headerIndex === -1) {
    throw new Error('Header soal tidak ditemukan. Gunakan format kolom "Nama, S1, S2, S3, ...".');
  }

  const headers = rows[headerIndex].map(cleanText);
  const nameIndex = headers.findIndex((header) => ["nama", "name", "responden", "respondent"].includes(header.toLowerCase()));
  const itemColumns = headers
    .map((header, index) => ({ header, index, number: parseItemNumber(header) }))
    .filter((item) => item.number !== null)
    .sort((a, b) => a.number - b.number);

  if (!itemColumns.length) {
    throw new Error('Kolom soal tidak ditemukan. Gunakan header S1, S2, S3, dan seterusnya.');
  }

  const dataRows = rows.slice(headerIndex + 1).filter((row) => row.some((cell) => cleanText(cell) !== ""));
  if (!dataRows.length) throw new Error("Data responden belum ada.");

  const respondents = dataRows.map((row, rowOffset) => {
    const name = nameIndex >= 0 ? cleanText(row[nameIndex]) : "";
    const scores = {};
    itemColumns.forEach((item) => {
      const value = normalizeBinary(row[item.index]);
      if (value === null) {
        throw new Error(`Nilai pada baris ${headerIndex + rowOffset + 2}, ${item.header} harus 1 atau 0.`);
      }
      scores[item.number] = value;
    });

    return {
      name: name || defaultRespondentName(rowOffset),
      scores,
      sourceRow: headerIndex + rowOffset + 2
    };
  });

  return {
    examName: cleanText(workbook.Props?.Title) || cleanText(els.examName.value) || "Ujian",
    subject: cleanText(workbook.Props?.Subject) || cleanText(els.subjectName.value) || "Mata Pelajaran",
    itemNumbers: itemColumns.map((item) => item.number),
    respondents
  };
}

function runAnalysis() {
  if (!state.matrixData) {
    showMessage("Masukkan data matriks 1/0 terlebih dahulu.", "error");
    return;
  }
  const results = analyzeMatrix(state.matrixData);
  state.results = results;
  renderResults(results);
  setResultSections(true);
  showMessage("Analisis selesai. Hasil sudah dipisahkan berdasarkan jenis analisis.", "success");
  scrollToSection("summarySection");
}

function analyzeMatrix(data) {
  const items = data.itemNumbers;
  const n = data.respondents.length;
  const k = items.length;
  const validityThreshold = parseOptionalThreshold(els.validityThreshold.value);

  const respondents = data.respondents.map((respondent) => {
    const total = items.reduce((sum, item) => sum + respondent.scores[item], 0);
    return {
      ...respondent,
      total,
      score: round((total / k) * 100, 2)
    };
  });

  const totalScores = respondents.map((respondent) => respondent.total);
  const totalVariance = populationVariance(totalScores);
  const sorted = [...respondents].sort((a, b) => b.total - a.total || a.name.localeCompare(b.name));
  const groupSize = Math.max(1, Math.round(n * 0.27));
  const upperGroup = sorted.slice(0, groupSize);
  const lowerGroup = sorted.slice(-groupSize);

  const difficulty = items.map((item) => {
    const correct = respondents.reduce((sum, respondent) => sum + respondent.scores[item], 0);
    const wrong = n - correct;
    const p = correct / n;
    return {
      item,
      correct,
      wrong,
      n,
      formula: `${correct} / ${n}`,
      p,
      category: categorizeDifficulty(p)
    };
  });

  const discrimination = items.map((item) => {
    const upperCorrect = upperGroup.reduce((sum, respondent) => sum + respondent.scores[item], 0);
    const lowerCorrect = lowerGroup.reduce((sum, respondent) => sum + respondent.scores[item], 0);
    const upperProp = upperCorrect / upperGroup.length;
    const lowerProp = lowerCorrect / lowerGroup.length;
    const d = upperProp - lowerProp;
    return {
      item,
      upperCorrect,
      upperTotal: upperGroup.length,
      lowerCorrect,
      lowerTotal: lowerGroup.length,
      formula: `(${upperCorrect}/${upperGroup.length}) - (${lowerCorrect}/${lowerGroup.length})`,
      d,
      category: categorizeDiscrimination(d)
    };
  });

  const validity = items.map((item) => {
    const x = respondents.map((respondent) => respondent.scores[item]);
    const y = respondents.map((respondent) => respondent.total - respondent.scores[item]);
    const sums = correlationSums(x, y);
    const r = pearsonFromSums(sums);
    const category = categorizeValidity(r, validityThreshold);
    return {
      item,
      sumX: sums.sumX,
      sumY: sums.sumY,
      sumXY: sums.sumXY,
      sumX2: sums.sumX2,
      sumY2: sums.sumY2,
      numerator: (sums.n * sums.sumXY) - (sums.sumX * sums.sumY),
      denominator: Math.sqrt(((sums.n * sums.sumX2) - (sums.sumX ** 2)) * ((sums.n * sums.sumY2) - (sums.sumY ** 2))),
      r,
      category: category.category,
      status: category.status
    };
  });

  const reliabilityItems = difficulty.map((item) => {
    const p = item.p;
    const q = 1 - p;
    return {
      item: item.item,
      p,
      q,
      pq: p * q
    };
  });
  const pqSum = reliabilityItems.reduce((sum, item) => sum + item.pq, 0);
  const kr20 = k > 1 && totalVariance > 0 ? (k / (k - 1)) * (1 - (pqSum / totalVariance)) : 0;

  return {
    meta: {
      examName: data.examName,
      subject: data.subject,
      respondentCount: n,
      itemCount: k,
      validityThreshold,
      groupSize
    },
    respondents,
    upperGroup,
    lowerGroup,
    difficulty,
    discrimination,
    validity,
    reliability: {
      items: reliabilityItems,
      k,
      pqSum,
      totalVariance,
      kr20: round(Math.max(-1, Math.min(1, kr20)), 3),
      category: categorizeReliability(kr20),
      formula: `${k}/(${k}-1) × (1 - ${round(pqSum, 3)} / ${round(totalVariance, 3)})`
    },
    summary: {
      average: round(average(respondents.map((respondent) => respondent.score)), 2),
      highest: round(Math.max(...respondents.map((respondent) => respondent.score)), 2),
      lowest: round(Math.min(...respondents.map((respondent) => respondent.score)), 2)
    }
  };
}

function renderPreview(data) {
  els.detectedSummary.innerHTML = [
    metricCard("Nama Ujian", data.examName),
    metricCard("Mata Pelajaran", data.subject),
    metricCard("Jumlah Responden", data.respondents.length),
    metricCard("Jumlah Soal", data.itemNumbers.length)
  ].join("");

  els.previewTable.innerHTML = tableHtml(
    ["Nama", ...data.itemNumbers.map((item) => `S${item}`)],
    data.respondents.slice(0, 10).map((respondent) => [
      respondent.name,
      ...data.itemNumbers.map((item) => respondent.scores[item])
    ])
  );
}

function renderResults(results) {
  els.summaryCards.innerHTML = [
    metricCard("Jumlah Responden", results.meta.respondentCount),
    metricCard("Jumlah Soal", results.meta.itemCount),
    metricCard("Rata-rata Nilai", results.summary.average),
    metricCard("Nilai Tertinggi", results.summary.highest),
    metricCard("Nilai Terendah", results.summary.lowest),
    metricCard("Reliabilitas KR-20", `${results.reliability.kr20} (${results.reliability.category})`)
  ].join("");

  els.scoreTable.innerHTML = tableHtml(
    ["No", "Nama", "Total Benar", "Total Salah", "Nilai"],
    results.respondents.map((respondent, index) => [
      index + 1,
      respondent.name,
      respondent.total,
      results.meta.itemCount - respondent.total,
      respondent.score
    ])
  );

  els.difficultyTable.innerHTML = tableHtml(
    ["No Soal", "Benar (B)", "Salah", "Jumlah (N)", "Perhitungan P = B/N", "Hasil P", "Kategori"],
    results.difficulty.map((item) => [
      `S${item.item}`,
      item.correct,
      item.wrong,
      item.n,
      item.formula,
      round(item.p, 3),
      item.category
    ])
  );
  els.difficultyInsight.innerHTML = buildDifficultyInsight(results.difficulty);

  els.groupInfo.innerHTML = `Kelompok atas dan bawah masing-masing berjumlah <strong>${results.meta.groupSize}</strong> responden.`;
  els.discriminationTable.innerHTML = tableHtml(
    ["No Soal", "BA", "JA", "BB", "JB", "Perhitungan D", "Hasil D", "Kategori"],
    results.discrimination.map((item) => [
      `S${item.item}`,
      item.upperCorrect,
      item.upperTotal,
      item.lowerCorrect,
      item.lowerTotal,
      item.formula,
      round(item.d, 3),
      item.category
    ])
  );
  els.discriminationInsight.innerHTML = buildDiscriminationInsight(results.discrimination);

  els.validityTable.innerHTML = tableHtml(
    ["No Soal", "ΣX", "ΣY", "ΣXY", "ΣX²", "ΣY²", "Pembilang", "Penyebut", "r hitung", "Keputusan"],
    results.validity.map((item) => [
      `S${item.item}`,
      item.sumX,
      item.sumY,
      item.sumXY,
      item.sumX2,
      item.sumY2,
      round(item.numerator, 3),
      round(item.denominator, 3),
      round(item.r, 3),
      results.meta.validityThreshold !== null
        ? `${item.status} (r tabel ${results.meta.validityThreshold})`
        : `${item.category} / ${item.status}`
    ])
  );
  els.validityInsight.innerHTML = buildValidityInsight(results.validity, results.meta.validityThreshold);

  els.reliabilityCards.innerHTML = [
    metricCard("Jumlah Soal (k)", results.reliability.k),
    metricCard("Σpq", round(results.reliability.pqSum, 3)),
    metricCard("Varians Total", round(results.reliability.totalVariance, 3)),
    metricCard("KR-20", results.reliability.kr20),
    metricCard("Kategori", results.reliability.category)
  ].join("");
  els.reliabilityTable.innerHTML = tableHtml(
    ["No Soal", "p", "q = 1-p", "p × q"],
    results.reliability.items.map((item) => [
      `S${item.item}`,
      round(item.p, 3),
      round(item.q, 3),
      round(item.pq, 3)
    ]).concat([["Rumus KR-20", "", "", results.reliability.formula]])
  );
  els.reliabilityInsight.innerHTML = buildReliabilityInsight(results.reliability);
}

function buildDifficultyInsight(items) {
  const counts = countBy(items.map((item) => item.category));
  return `
    <h3>Hasil kesukaran</h3>
    <p>Total butir: <strong>${items.length}</strong></p>
    <ul>
      <li>Sukar: ${counts.Sukar || 0}</li>
      <li>Sedang: ${counts.Sedang || 0}</li>
      <li>Mudah: ${counts.Mudah || 0}</li>
    </ul>
  `;
}

function buildDiscriminationInsight(items) {
  const counts = countBy(items.map((item) => item.category));
  return `
    <h3>Hasil daya pembeda</h3>
    <p>Butir dengan daya pembeda baik perlu dipertahankan, sedangkan kategori jelek atau buruk perlu diperiksa lagi.</p>
    <ul>
      <li>Sangat Baik: ${counts["Sangat Baik"] || 0}</li>
      <li>Baik: ${counts.Baik || 0}</li>
      <li>Cukup: ${counts.Cukup || 0}</li>
      <li>Jelek/Buruk: ${(counts.Jelek || 0) + (counts.Buruk || 0)}</li>
    </ul>
  `;
}

function buildValidityInsight(items, threshold) {
  const valid = items.filter((item) => item.status === "Valid").length;
  const notValid = items.filter((item) => item.status === "Tidak Valid").length;
  const enough = items.length - valid - notValid;
  const rule = threshold !== null ? `Keputusan memakai r tabel ${threshold}.` : "Keputusan memakai kategori korelasi otomatis.";
  return `
    <h3>Hasil validitas</h3>
    <p>${rule}</p>
    <ul>
      <li>Valid: ${valid}</li>
      <li>Cukup: ${enough}</li>
      <li>Tidak valid: ${notValid}</li>
    </ul>
  `;
}

function buildReliabilityInsight(reliability) {
  return `
    <h3>Hasil reliabilitas</h3>
    <p>Koefisien KR-20 tes adalah <strong>${reliability.kr20}</strong> dengan kategori <strong>${escapeHtml(reliability.category)}</strong>.</p>
    <p>Perhitungan: ${escapeHtml(reliability.formula)}</p>
  `;
}

function exportExcelReport() {
  if (!state.results) return;
  const results = state.results;
  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ["ButirCerdas - Ringkasan"],
    ["Nama Ujian", results.meta.examName],
    ["Mata Pelajaran", results.meta.subject],
    ["Jumlah Responden", results.meta.respondentCount],
    ["Jumlah Soal", results.meta.itemCount],
    ["Rata-rata Nilai", results.summary.average],
    ["Nilai Tertinggi", results.summary.highest],
    ["Nilai Terendah", results.summary.lowest],
    ["Reliabilitas KR-20", results.reliability.kr20],
    ["Kategori Reliabilitas", results.reliability.category]
  ]), "Ringkasan");

  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(tableToRows(els.scoreTable)), "Skor Responden");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(tableToRows(els.difficultyTable)), "Kesukaran");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(tableToRows(els.discriminationTable)), "Daya Pembeda");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(tableToRows(els.validityTable)), "Validitas");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(tableToRows(els.reliabilityTable)), "Reliabilitas");
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
  doc.text(`Responden: ${results.meta.respondentCount} | Soal: ${results.meta.itemCount} | KR-20: ${results.reliability.kr20} (${results.reliability.category})`, 14, 36);

  addPdfTable(doc, "Skor Responden", tableToRows(els.scoreTable), 44);
  addPdfTable(doc, "Tingkat Kesukaran", tableToRows(els.difficultyTable));
  addPdfTable(doc, "Daya Pembeda", tableToRows(els.discriminationTable));
  addPdfTable(doc, "Validitas", tableToRows(els.validityTable));
  addPdfTable(doc, "Reliabilitas", tableToRows(els.reliabilityTable));
  doc.save("laporan-butircerdas.pdf");
}

function addPdfTable(doc, title, rows, startY = null) {
  const y = startY || doc.lastAutoTable.finalY + 12;
  doc.text(title, 14, y);
  doc.autoTable({
    startY: y + 4,
    head: [rows[0]],
    body: rows.slice(1),
    styles: { fontSize: 7 },
    headStyles: { fillColor: [15, 118, 110] }
  });
}

function scrollToSection(id) {
  const resultIds = ["summarySection", "difficultySection", "discriminationSection", "validitySection", "reliabilitySection"];
  if (resultIds.includes(id) && !state.results) {
    showMessage("Masukkan data matriks dan klik Hitung Analisis terlebih dahulu.", "error");
    document.getElementById("inputSection")?.scrollIntoView({ behavior: "smooth", block: "start" });
    setActiveNav("inputSection");
    return false;
  }
  document.getElementById(id)?.scrollIntoView({ behavior: "smooth", block: "start" });
  return true;
}

function setActiveNav(id) {
  document.querySelectorAll(".nav-btn").forEach((button) => {
    button.classList.toggle("active", button.dataset.target === id);
  });
}

function setupSectionObserver() {
  if (!("IntersectionObserver" in window)) {
    setActiveNav("inputSection");
    return;
  }
  const sections = ["inputSection", "summarySection", "difficultySection", "discriminationSection", "validitySection", "reliabilitySection"]
    .map((id) => document.getElementById(id))
    .filter(Boolean);
  const observer = new IntersectionObserver((entries) => {
    const visible = entries
      .filter((entry) => entry.isIntersecting)
      .sort((a, b) => b.intersectionRatio - a.intersectionRatio)[0];
    if (visible) setActiveNav(visible.target.id);
  }, { rootMargin: "-20% 0px -65% 0px", threshold: [0.1, 0.35, 0.6] });
  sections.forEach((section) => observer.observe(section));
  setActiveNav("inputSection");
}

function setResultSections(enabled) {
  document.querySelectorAll(".result-section").forEach((section) => {
    section.classList.toggle("hidden", !enabled);
  });
  document.querySelectorAll(".result-nav").forEach((button) => {
    button.disabled = !enabled;
  });
}

function parseItemNumber(value) {
  const text = cleanText(value);
  const match = text.match(/^s(\d+)$/i);
  if (match) return Number(match[1]);
  return null;
}

function normalizeBinary(value) {
  const text = cleanText(value);
  if (text === "1") return 1;
  if (text === "0") return 0;
  if (value === 1) return 1;
  if (value === 0) return 0;
  return null;
}

function defaultRespondentName(index) {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  if (index < alphabet.length) return `Responden ${alphabet[index]}`;
  return `Responden ${index + 1}`;
}

function categorizeDifficulty(value) {
  if (value <= 0.3) return "Sukar";
  if (value <= 0.7) return "Sedang";
  return "Mudah";
}

function categorizeDiscrimination(value) {
  if (value < 0) return "Buruk";
  if (value < 0.2) return "Jelek";
  if (value < 0.4) return "Cukup";
  if (value < 0.7) return "Baik";
  return "Sangat Baik";
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

function correlationSums(x, y) {
  return x.reduce((acc, value, index) => {
    const yValue = y[index];
    acc.n += 1;
    acc.sumX += value;
    acc.sumY += yValue;
    acc.sumXY += value * yValue;
    acc.sumX2 += value * value;
    acc.sumY2 += yValue * yValue;
    return acc;
  }, { n: 0, sumX: 0, sumY: 0, sumXY: 0, sumX2: 0, sumY2: 0 });
}

function pearsonFromSums(sums) {
  const numerator = (sums.n * sums.sumXY) - (sums.sumX * sums.sumY);
  const denominator = Math.sqrt(((sums.n * sums.sumX2) - (sums.sumX ** 2)) * ((sums.n * sums.sumY2) - (sums.sumY ** 2)));
  return denominator === 0 ? 0 : numerator / denominator;
}

function populationVariance(values) {
  if (!values.length) return 0;
  const avg = average(values);
  return values.reduce((sum, value) => sum + ((value - avg) ** 2), 0) / values.length;
}

function parseOptionalThreshold(value) {
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0) return null;
  return Math.min(1, Math.max(0, round(number, 3)));
}

function countBy(values) {
  return values.reduce((acc, value) => {
    acc[value] = (acc[value] || 0) + 1;
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

function tableToRows(table) {
  const header = [...table.querySelectorAll("thead th")].map((cell) => cell.textContent);
  const body = [...table.querySelectorAll("tbody tr")].map((row) => [...row.querySelectorAll("td")].map((cell) => cell.textContent));
  return [header, ...body];
}

function formatCell(value) {
  const text = escapeHtml(String(value ?? ""));
  if (["Valid", "Tinggi", "Sangat Tinggi", "Baik", "Sangat Baik"].includes(value)) return `<span class="badge success">${text}</span>`;
  if (["Cukup", "Sedang", "Mudah"].includes(value)) return `<span class="badge warning">${text}</span>`;
  if (["Tidak Valid", "Rendah", "Sangat Rendah", "Buruk", "Jelek", "Sukar"].includes(value)) return `<span class="badge danger">${text}</span>`;
  return text;
}

function showMessage(message, type = "success") {
  els.messageBox.innerHTML = `<div class="message ${type}">${escapeHtml(message)}</div>`;
}

function average(values) {
  return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : 0;
}

function round(value, digits = 2) {
  return Number(Number(value || 0).toFixed(digits));
}

function range(start, end) {
  return Array.from({ length: end - start + 1 }, (_, index) => start + index);
}

function cleanText(value) {
  return String(value ?? "").trim();
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
