module.exports = async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  const apiKey = process.env.GEMINI_API_KEY;
  const model = process.env.GEMINI_MODEL || "gemini-2.5-flash";
  if (!apiKey) {
    res.status(503).json({ error: "GEMINI_API_KEY belum diatur di environment variable Vercel." });
    return;
  }

  try {
    const payload = normalizePayload(typeof req.body === "string" ? JSON.parse(req.body || "{}") : req.body || {});
    const prompt = [
      "Buat narasi rekomendasi lengkap dalam bahasa Indonesia untuk hasil analisis butir soal berikut.",
      "Jangan hitung ulang rumus. Gunakan hasil statistik yang diberikan.",
      "Analisis memakai matriks skor 1/0, bukan pilihan A/B/C/D dan bukan analisis distraktor.",
      "Jangan berhenti di tengah kalimat.",
      "Gunakan format berikut:",
      "1. Ringkasan kualitas tes",
      "2. Masalah utama",
      "3. Saran perbaikan",
      "4. Prioritas tindak lanjut",
      "Tulis 180 sampai 260 kata. Gunakan bahasa yang jelas untuk guru. Jangan menyebut distraktor.",
      JSON.stringify(payload, null, 2)
    ].join("\n\n");

    const data = await callGeminiWithRetry({
      apiKey,
      model,
      body: {
        contents: [
          {
            parts: [{ text: prompt }]
          }
        ],
        generationConfig: {
          temperature: 0.25,
          maxOutputTokens: 1800
        }
      }
    });

    const recommendation = data.candidates?.[0]?.content?.parts?.map((part) => part.text).join("\n").trim();
    const finishReason = data.candidates?.[0]?.finishReason;
    if (finishReason === "MAX_TOKENS") {
      res.status(502).json({ error: "Respons Gemini terpotong karena batas output. Coba ulangi rekomendasi AI." });
      return;
    }
    res.status(200).json({ recommendation: recommendation || "Rekomendasi AI belum menghasilkan teks." });
  } catch (error) {
    res.status(500).json({ error: error.message || "Terjadi kesalahan pada serverless function." });
  }
};

async function callGeminiWithRetry({ apiKey, model, body }) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
  const retryableStatuses = new Set([429, 500, 502, 503, 504]);
  let lastError;

  for (let attempt = 0; attempt < 3; attempt += 1) {
    try {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 25000);
      const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        signal: controller.signal,
        body: JSON.stringify(body)
      });
      clearTimeout(timeout);

      const text = await response.text();
      const data = text ? JSON.parse(text) : {};

      if (response.ok) return data;
      const message = data.error?.message || "Gemini API gagal dipanggil.";
      lastError = new Error(message);
      if (!retryableStatuses.has(response.status)) throw lastError;
    } catch (error) {
      lastError = error;
      if (attempt === 2) break;
    }

    await delay(700 * (attempt + 1));
  }

  throw lastError || new Error("Gemini API gagal dipanggil.");
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function normalizePayload(payload) {
  return {
    meta: {
      examName: String(payload.meta?.examName || ""),
      subject: String(payload.meta?.subject || ""),
      questionCount: Number(payload.meta?.questionCount || 0),
      studentCount: Number(payload.meta?.studentCount || 0)
    },
    summary: {
      average: Number(payload.summary?.average || 0),
      highest: Number(payload.summary?.highest || 0),
      lowest: Number(payload.summary?.lowest || 0),
      reliability: Number(payload.summary?.reliability || 0),
      reliabilityCategory: String(payload.summary?.reliabilityCategory || "")
    },
    itemCategoryCounts: {
      difficulty: sanitizeCountObject(payload.itemCategoryCounts?.difficulty),
      discrimination: sanitizeCountObject(payload.itemCategoryCounts?.discrimination),
      validity: sanitizeCountObject(payload.itemCategoryCounts?.validity)
    },
    problematicItems: Array.isArray(payload.problematicItems)
      ? payload.problematicItems.slice(0, 30).map((item) => ({
        number: Number(item.number || 0),
        difficultyIndex: Number(item.difficultyIndex || 0),
        difficulty: String(item.difficulty || ""),
        discriminationIndex: Number(item.discriminationIndex || 0),
        discrimination: String(item.discrimination || ""),
        validityIndex: Number(item.validityIndex || 0),
        validity: String(item.validity || "")
      }))
      : [],
    topItems: Array.isArray(payload.topItems)
      ? payload.topItems.slice(0, 10).map((item) => ({
        number: Number(item.number || 0),
        difficulty: String(item.difficulty || ""),
        discrimination: String(item.discrimination || ""),
        validity: String(item.validity || "")
      }))
      : []
  };
}

function sanitizeCountObject(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  return Object.fromEntries(
    Object.entries(value)
      .slice(0, 12)
      .map(([key, count]) => [String(key), Number(count || 0)])
  );
}
