module.exports = async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  const apiKey = process.env.GEMINI_API_KEY;
  const model = process.env.GEMINI_MODEL || "gemini-2.0-flash";
  if (!apiKey) {
    res.status(503).json({ error: "GEMINI_API_KEY belum diatur di environment variable Vercel." });
    return;
  }

  try {
    const payload = normalizePayload(typeof req.body === "string" ? JSON.parse(req.body || "{}") : req.body || {});
    const prompt = [
      "Buat rekomendasi singkat dalam bahasa Indonesia untuk hasil analisis butir soal berikut.",
      "Jangan hitung ulang rumus. Gunakan hasil statistik yang diberikan.",
      "Tuliskan ringkasan kualitas tes, masalah utama, dan saran perbaikan praktis.",
      JSON.stringify(payload, null, 2)
    ].join("\n\n");

    const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [
          {
            parts: [{ text: prompt }]
          }
        ],
        generationConfig: {
          temperature: 0.4,
          maxOutputTokens: 700
        }
      })
    });

    const data = await response.json();
    if (!response.ok) {
      res.status(response.status).json({ error: data.error?.message || "Gemini API gagal dipanggil." });
      return;
    }

    const recommendation = data.candidates?.[0]?.content?.parts?.map((part) => part.text).join("\n").trim();
    res.status(200).json({ recommendation: recommendation || "Rekomendasi AI belum menghasilkan teks." });
  } catch (error) {
    res.status(500).json({ error: error.message || "Terjadi kesalahan pada serverless function." });
  }
};

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
    problematicItems: Array.isArray(payload.problematicItems)
      ? payload.problematicItems.slice(0, 30).map((item) => ({
        number: Number(item.number || 0),
        difficulty: String(item.difficulty || ""),
        discrimination: String(item.discrimination || ""),
        validity: String(item.validity || ""),
        distractors: Array.isArray(item.distractors) ? item.distractors.map(String).slice(0, 5) : [],
        recommendation: String(item.recommendation || "")
      }))
      : []
  };
}
