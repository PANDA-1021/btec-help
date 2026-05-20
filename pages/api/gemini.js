// pages/api/gemini.js
// Runs server-side on Vercel. API key is read from environment — never sent to browser.

export const config = {
  api: {
    bodyParser: {
      sizeLimit: "20mb",
    },
    // Increase response timeout for large PDFs
    responseLimit: false,
  },
};

const GEMINI_MODELS = [
  "gemini-2.5-flash",
  "gemini-2.0-flash",
  "gemini-1.5-flash",
];

async function callGeminiWithRetry(apiKey, body, retries = 2) {
  for (let attempt = 0; attempt <= retries; attempt++) {
    const modelIdx = Math.min(attempt, GEMINI_MODELS.length - 1);
    const model = GEMINI_MODELS[modelIdx];

    // thinkingConfig is only supported on gemini-2.5-flash; strip it for all other models
    const modelBody = model === "gemini-2.5-flash"
      ? body
      : { ...body, generationConfig: { ...body.generationConfig, thinkingConfig: undefined } };

    const res = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(modelBody),
        signal: AbortSignal.timeout(55000), // 55s — under Vercel's 60s limit
      }
    );

    if (res.ok) {
      const data = await res.json();
      // Handle both standard and thinking-model response shapes
      const candidates = data.candidates || [];
      const parts = candidates[0]?.content?.parts || [];
      // Skip thought parts (type=="thinking"), get the first text part
      const textPart = parts.find((p) => p.text !== undefined && !p.thought);
      const text = textPart?.text || parts[0]?.text || "";
      if (text) return { text, model };
    }

    const errBody = await res.json().catch(() => ({}));
    const errMsg = errBody?.error?.message || `HTTP ${res.status}`;

    // Only hard-stop on auth errors; treat 400 as retryable (model may not support a param)
    if (res.status === 401 || res.status === 403) {
      throw new Error(errMsg);
    }

    if (attempt === retries) throw new Error(`Gemini error after ${retries + 1} attempts: ${errMsg}`);
    // Exponential back-off: 500ms, 1000ms
    await new Promise((r) => setTimeout(r, 500 * (attempt + 1)));
  }
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
  if (!GEMINI_API_KEY) {
    return res.status(500).json({
      error: "Gemini API key not configured. Add GEMINI_API_KEY to your Vercel environment variables.",
    });
  }

  const { prompt, base64Pdf, slideCount, temperature } = req.body;

  if (!prompt) {
    return res.status(400).json({ error: "Missing prompt" });
  }

  const parts = [];
  if (base64Pdf) {
    parts.push({ inline_data: { mime_type: "application/pdf", data: base64Pdf } });
  }
  parts.push({ text: prompt });

  const body = {
    contents: [{ parts }],
    generationConfig: {
      temperature: temperature ?? 0.3,
      maxOutputTokens: 8192,
      // Disable thinking for faster responses (works on 2.5-flash)
      thinkingConfig: { thinkingBudget: 0 },
    },
  };

  try {
    const result = await callGeminiWithRetry(GEMINI_API_KEY, body);
    return res.status(200).json({ text: result.text, model: result.model });
  } catch (err) {
    return res.status(500).json({ error: err.message || "Server error" });
  }
}
