// pages/api/gemini.js
// Runs server-side on Vercel. API key is read from environment — never sent to browser.

export const config = {
  api: {
    bodyParser: {
      sizeLimit: "20mb",
    },
  },
};

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
  if (!GEMINI_API_KEY) {
    return res.status(500).json({
      error:
        "Gemini API key not configured. Add GEMINI_API_KEY to your Vercel environment variables.",
    });
  }

  const { prompt, base64Pdf } = req.body;

  if (!prompt) {
    return res.status(400).json({ error: "Missing prompt" });
  }

  const parts = [];
  if (base64Pdf) {
    parts.push({
      inline_data: { mime_type: "application/pdf", data: base64Pdf },
    });
  }
  parts.push({ text: prompt });

  try {
    const geminiRes = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ contents: [{ parts }] }),
      },
    );

    if (!geminiRes.ok) {
      const err = await geminiRes.json().catch(() => ({}));
      return res.status(geminiRes.status).json({
        error: err?.error?.message || `Gemini error ${geminiRes.status}`,
      });
    }

    const data = await geminiRes.json();
    const text = data.candidates?.[0]?.content?.parts?.[0]?.text || "";
    return res.status(200).json({ text });
  } catch (err) {
    return res.status(500).json({ error: err.message || "Server error" });
  }
}
