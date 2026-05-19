import { useState, useRef } from "react";

const STEPS = {
  IDLE: "idle",
  READING: "reading",
  PARAPHRASING: "paraphrasing",
  BUILDING: "building",
  DONE: "done",
  ERROR: "error",
};

// ── Gemini API ───────────────────────────────────────────────────────────────

async function callGemini(prompt, geminiKey, base64Pdf = null) {
  const parts = [];
  if (base64Pdf) {
    parts.push({ inline_data: { mime_type: "application/pdf", data: base64Pdf } });
  }
  parts.push({ text: prompt });

  const res = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${geminiKey}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ contents: [{ parts }] }),
    }
  );
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err?.error?.message || `Gemini error ${res.status}`);
  }
  const data = await res.json();
  return data.candidates?.[0]?.content?.parts?.[0]?.text || "";
}

// ── Groq API (fallback — text only, no PDF) ──────────────────────────────────

async function callGroq(prompt, groqKey) {
  const res = await fetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${groqKey}`,
    },
    body: JSON.stringify({
      model: "llama3-8b-8192",
      messages: [{ role: "user", content: prompt }],
      temperature: 0.7,
    }),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err?.error?.message || `Groq error ${res.status}`);
  }
  const data = await res.json();
  return data.choices?.[0]?.message?.content || "";
}

// ── Parse JSON from LLM response ─────────────────────────────────────────────

function parseJSON(raw) {
  const clean = raw.replace(/```json|```/g, "").trim();
  return JSON.parse(clean);
}

// ── Extract structure from PDF (Gemini only — it reads PDFs) ─────────────────

const EXTRACT_PROMPT = `You are a document parser. Extract the full content of this BTEC assignment into JSON.
Return ONLY valid JSON, no markdown fences, no explanation.

Format:
{
  "title": "Assignment Title",
  "slides": [
    {
      "heading": "Section heading (keep exactly as written)",
      "bullets": ["Key point or sentence from that section", "Another point"]
    }
  ]
}

Rules:
- Use actual section headings from the document as slide headings (keep them verbatim)
- Break each section into 3-6 concise bullet points capturing the core ideas
- Keep bullets close to the original phrasing (they will be paraphrased next)
- Aim for 6-14 slides total
- If no clear sections, create logical topic groupings`;

async function extractStructure(base64Pdf, keys, addLog) {
  if (keys.gemini) {
    addLog("🔵 Gemini: reading PDF and extracting structure…");
    try {
      const raw = await callGemini(EXTRACT_PROMPT, keys.gemini, base64Pdf);
      return parseJSON(raw);
    } catch (e) {
      addLog(`⚠️  Gemini failed (${e.message}) — falling back to Groq…`);
    }
  }
  // Groq fallback: can't read PDF natively, so ask user's text
  throw new Error("PDF extraction requires a Gemini API key. Groq cannot read PDFs directly.");
}

// ── Paraphrase bullets ────────────────────────────────────────────────────────

function buildParaphrasePrompt(slides) {
  return `You are an academic paraphrasing assistant.
Paraphrase ONLY the bullet points in the JSON below — rewrite them in different words while fully preserving the meaning.
Do NOT change the headings at all.
Return ONLY valid JSON in exactly the same structure, no markdown fences, no explanation.

${JSON.stringify(slides, null, 2)}`;
}

async function paraphraseSlides(slides, keys, addLog) {
  const prompt = buildParaphrasePrompt(slides);

  if (keys.gemini) {
    addLog("🔵 Gemini: paraphrasing bullet points…");
    try {
      const raw = await callGemini(prompt, keys.gemini);
      return parseJSON(raw);
    } catch (e) {
      addLog(`⚠️  Gemini paraphrase failed (${e.message}) — trying Groq…`);
    }
  }

  if (keys.groq) {
    addLog("🟠 Groq: paraphrasing bullet points…");
    try {
      const raw = await callGroq(prompt, keys.groq);
      return parseJSON(raw);
    } catch (e) {
      addLog(`⚠️  Groq paraphrase failed (${e.message})`);
      throw e;
    }
  }

  throw new Error("No API key available for paraphrasing.");
}

// ── PptxGenJS builder ─────────────────────────────────────────────────────────

async function buildPptx(data) {
  await new Promise((resolve, reject) => {
    if (window.PptxGenJS) return resolve();
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundled.js";
    s.onload = resolve;
    s.onerror = reject;
    document.head.appendChild(s);
  });

  const pres = new window.PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.title = data.title || "BTEC Assignment";

  const NAVY = "1B2B4B";
  const ACCENT = "3B82F6";
  const WHITE = "FFFFFF";
  const LIGHT_BG = "F0F4FA";

  // Title slide
  const titleSlide = pres.addSlide();
  titleSlide.background = { color: NAVY };
  titleSlide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 4.2, w: 10, h: 1.425,
    fill: { color: ACCENT }, line: { color: ACCENT },
  });
  titleSlide.addText(data.title || "BTEC Assignment", {
    x: 0.6, y: 1.0, w: 8.8, h: 2.8,
    fontSize: 36, fontFace: "Calibri", bold: true,
    color: WHITE, align: "left", valign: "middle", wrap: true,
  });
  titleSlide.addText("Prepared from BTEC Assignment", {
    x: 0.6, y: 4.35, w: 8.8, h: 0.7,
    fontSize: 14, fontFace: "Calibri", color: WHITE, align: "left", valign: "middle",
  });

  // Content slides
  data.slides.forEach((slide, idx) => {
    const s = pres.addSlide();
    s.background = { color: idx % 2 === 0 ? WHITE : LIGHT_BG };

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 0.12, h: 5.625,
      fill: { color: ACCENT }, line: { color: ACCENT },
    });

    s.addShape(pres.shapes.OVAL, {
      x: 8.9, y: 0.15, w: 0.6, h: 0.6,
      fill: { color: NAVY }, line: { color: NAVY },
    });
    s.addText(String(idx + 1), {
      x: 8.9, y: 0.15, w: 0.6, h: 0.6,
      fontSize: 11, fontFace: "Calibri", bold: true,
      color: WHITE, align: "center", valign: "middle",
    });

    s.addText(slide.heading, {
      x: 0.3, y: 0.1, w: 8.5, h: 0.85,
      fontSize: 22, fontFace: "Calibri", bold: true,
      color: NAVY, align: "left", valign: "middle",
    });

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.3, y: 0.95, w: 8.5, h: 0.025,
      fill: { color: ACCENT }, line: { color: ACCENT },
    });

    const bullets = (slide.bullets || []).slice(0, 7);
    if (bullets.length > 0) {
      s.addText(
        bullets.map((b, i) => ({
          text: b,
          options: {
            bullet: true,
            breakLine: i < bullets.length - 1,
            fontSize: 15, fontFace: "Calibri", color: "2D3748", paraSpaceAfter: 6,
          },
        })),
        { x: 0.3, y: 1.1, w: 9.2, h: 4.2, valign: "top", wrap: true }
      );
    }
  });

  // End slide
  const endSlide = pres.addSlide();
  endSlide.background = { color: NAVY };
  endSlide.addShape(pres.shapes.RECTANGLE, {
    x: 3.5, y: 2.55, w: 3, h: 0.06,
    fill: { color: ACCENT }, line: { color: ACCENT },
  });
  endSlide.addText("Thank You", {
    x: 0, y: 1.5, w: 10, h: 1.5,
    fontSize: 44, fontFace: "Calibri", bold: true,
    color: WHITE, align: "center", valign: "middle",
  });
  endSlide.addText("Generated from BTEC Assignment", {
    x: 0, y: 3.2, w: 10, h: 0.6,
    fontSize: 14, fontFace: "Calibri", color: ACCENT, align: "center",
  });

  const filename = `${(data.title || "BTEC_Assignment").replace(/[^a-zA-Z0-9]/g, "_")}.pptx`;
  await pres.writeFile({ fileName: filename });
  return filename;
}

// ── Input field component ─────────────────────────────────────────────────────

function KeyInput({ label, placeholder, value, onChange, hint, color }) {
  const [show, setShow] = useState(false);
  return (
    <div style={{ marginBottom: 14 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
        <label style={{ color: "#cbd5e1", fontSize: 13, fontWeight: 600 }}>
          <span style={{ color, marginRight: 6 }}>●</span>{label}
        </label>
        <a href={hint.url} target="_blank" rel="noreferrer"
          style={{ color: "#64748b", fontSize: 11, textDecoration: "none" }}>
          Get free key →
        </a>
      </div>
      <div style={{ position: "relative" }}>
        <input
          type={show ? "text" : "password"}
          placeholder={placeholder}
          value={value}
          onChange={e => onChange(e.target.value)}
          style={{
            width: "100%", boxSizing: "border-box",
            background: "rgba(255,255,255,0.05)",
            border: "1px solid rgba(255,255,255,0.12)",
            borderRadius: 8, padding: "10px 40px 10px 12px",
            color: "#f0f4fa", fontSize: 13,
            outline: "none", fontFamily: "monospace",
          }}
        />
        <button onClick={() => setShow(s => !s)} style={{
          position: "absolute", right: 10, top: "50%", transform: "translateY(-50%)",
          background: "none", border: "none", color: "#64748b", cursor: "pointer", fontSize: 14,
        }}>{show ? "🙈" : "👁"}</button>
      </div>
      <div style={{ color: "#475569", fontSize: 11, marginTop: 4 }}>{hint.text}</div>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────

export default function App() {
  const [geminiKey, setGeminiKey] = useState("");
  const [groqKey, setGroqKey] = useState("");
  const [step, setStep] = useState(STEPS.IDLE);
  const [log, setLog] = useState([]);
  const [error, setError] = useState("");
  const [preview, setPreview] = useState(null);
  const [fileName, setFileName] = useState("");
  const fileRef = useRef(null);

  const addLog = (msg) => setLog(l => [...l, msg]);

  const reset = () => {
    setStep(STEPS.IDLE); setLog([]); setError("");
    setPreview(null); setFileName("");
    if (fileRef.current) fileRef.current.value = "";
  };

  const handleFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.type !== "application/pdf") { setError("Please upload a PDF file."); return; }
    if (!geminiKey && !groqKey) { setError("Please enter at least a Gemini API key."); return; }
    if (!geminiKey) { setError("A Gemini API key is required to read the PDF. Groq alone cannot parse PDFs."); return; }

    setError(""); setLog([]); setPreview(null);

    try {
      // Read PDF
      setStep(STEPS.READING);
      addLog("📄 Reading PDF file…");
      const base64 = await new Promise((res, rej) => {
        const r = new FileReader();
        r.onload = () => res(r.result.split(",")[1]);
        r.onerror = () => rej(new Error("Failed to read file"));
        r.readAsDataURL(file);
      });

      // Extract
      const structure = await extractStructure(base64, { gemini: geminiKey, groq: groqKey }, addLog);
      addLog(`✅ Extracted ${structure.slides?.length || 0} sections — "${structure.title}"`);

      // Paraphrase
      setStep(STEPS.PARAPHRASING);
      addLog("✍️  Paraphrasing body content (headings untouched)…");
      const paraphrased = await paraphraseSlides(structure.slides, { gemini: geminiKey, groq: groqKey }, addLog);
      const finalData = { title: structure.title, slides: paraphrased };
      setPreview(finalData);
      addLog("✅ Paraphrasing complete.");

      // Build PPTX
      setStep(STEPS.BUILDING);
      addLog("📊 Building PowerPoint…");
      const fn = await buildPptx(finalData);
      setFileName(fn);
      addLog(`🎉 Done! "${fn}" saved to your downloads.`);
      setStep(STEPS.DONE);
    } catch (err) {
      setError(err.message || "Something went wrong.");
      setStep(STEPS.ERROR);
    }
  };

  const busy = [STEPS.READING, STEPS.PARAPHRASING, STEPS.BUILDING].includes(step);

  const stepLabel = {
    [STEPS.READING]: "Reading PDF with Gemini…",
    [STEPS.PARAPHRASING]: "Paraphrasing content…",
    [STEPS.BUILDING]: "Building PowerPoint…",
  }[step];

  const canSubmit = !!geminiKey.trim();

  return (
    <div style={{
      minHeight: "100vh",
      background: "linear-gradient(135deg, #0f172a 0%, #1e3a5f 50%, #0f172a 100%)",
      display: "flex", alignItems: "center", justifyContent: "center",
      fontFamily: "'Calibri', 'Segoe UI', sans-serif", padding: "2rem",
    }}>
      <style>{`
        @keyframes spin { to { transform: rotate(360deg); } }
        input::placeholder { color: #334155; }
        input:focus { border-color: rgba(59,130,246,0.5) !important; }
        * { transition: border-color 0.2s; }
      `}</style>

      <div style={{
        width: "100%", maxWidth: 660,
        background: "rgba(255,255,255,0.05)",
        backdropFilter: "blur(20px)",
        borderRadius: 20,
        border: "1px solid rgba(255,255,255,0.12)",
        padding: "2.5rem",
        boxShadow: "0 25px 60px rgba(0,0,0,0.5)",
      }}>
        {/* Header */}
        <div style={{ textAlign: "center", marginBottom: "2rem" }}>
          <div style={{
            display: "inline-flex", alignItems: "center", justifyContent: "center",
            width: 64, height: 64, borderRadius: 16,
            background: "linear-gradient(135deg, #3b82f6, #1d4ed8)",
            marginBottom: "1rem",
            boxShadow: "0 8px 24px rgba(59,130,246,0.4)",
          }}>
            <span style={{ fontSize: 30 }}>📑</span>
          </div>
          <h1 style={{ margin: 0, fontSize: 24, fontWeight: 700, color: "#f0f4fa", letterSpacing: "-0.5px" }}>
            BTEC → PowerPoint
          </h1>
          <p style={{ margin: "0.5rem 0 0", color: "#94a3b8", fontSize: 13 }}>
            Powered by Gemini (free) · Groq fallback · No subscriptions
          </p>
        </div>

        {/* API Keys */}
        {!busy && step !== STEPS.DONE && (
          <div style={{
            background: "rgba(0,0,0,0.2)", borderRadius: 12,
            padding: "1.2rem 1.4rem", marginBottom: "1.5rem",
            border: "1px solid rgba(255,255,255,0.07)",
          }}>
            <div style={{ color: "#64748b", fontSize: 11, textTransform: "uppercase", letterSpacing: 1, marginBottom: 14 }}>
              API Keys — stored in browser only, never sent anywhere else
            </div>
            <KeyInput
              label="Gemini API Key (required)"
              placeholder="AIza..."
              value={geminiKey}
              onChange={setGeminiKey}
              color="#4ade80"
              hint={{ text: "Free tier: 15 req/min, 1500/day. Reads PDFs natively.", url: "https://aistudio.google.com/app/apikey" }}
            />
            <KeyInput
              label="Groq API Key (optional fallback)"
              placeholder="gsk_..."
              value={groqKey}
              onChange={setGroqKey}
              color="#fb923c"
              hint={{ text: "Free tier: generous limits. Used as fallback for paraphrasing.", url: "https://console.groq.com/keys" }}
            />
          </div>
        )}

        {/* Upload */}
        {!busy && step !== STEPS.DONE && (
          <label style={{
            display: "block",
            border: `2px dashed ${canSubmit ? "rgba(59,130,246,0.5)" : "rgba(255,255,255,0.1)"}`,
            borderRadius: 14, padding: "2rem 1.5rem", textAlign: "center",
            cursor: canSubmit ? "pointer" : "not-allowed",
            background: canSubmit ? "rgba(59,130,246,0.04)" : "rgba(255,255,255,0.02)",
            marginBottom: "1.5rem",
            opacity: canSubmit ? 1 : 0.5,
          }}>
          <input
            ref={fileRef}
            type="file"
            accept="application/pdf"
            style={{ display: "none" }}
            onChange={handleFile}
            disabled={busy || !canSubmit}
          />
            <div style={{ fontSize: 36, marginBottom: 8 }}>📂</div>
            <div style={{ color: "#cbd5e1", fontSize: 14, fontWeight: 600 }}>
              {canSubmit ? "Click to upload BTEC PDF" : "Enter your Gemini API key first"}
            </div>
            <div style={{ color: "#64748b", fontSize: 12, marginTop: 4 }}>
              PDF only · Headings kept verbatim · Body fully paraphrased
            </div>
          </label>
        )}

        {/* Progress */}
        {busy && (
          <div style={{
            background: "rgba(59,130,246,0.08)", border: "1px solid rgba(59,130,246,0.2)",
            borderRadius: 12, padding: "1.5rem", marginBottom: "1.5rem",
          }}>
            <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
              <div style={{
                width: 20, height: 20, border: "3px solid #3b82f6",
                borderTopColor: "transparent", borderRadius: "50%",
                animation: "spin 0.8s linear infinite",
              }} />
              <span style={{ color: "#93c5fd", fontWeight: 600, fontSize: 14 }}>{stepLabel}</span>
            </div>
            {[
              { key: STEPS.READING, label: "Read & extract PDF structure" },
              { key: STEPS.PARAPHRASING, label: "Paraphrase body content" },
              { key: STEPS.BUILDING, label: "Build PowerPoint file" },
            ].map(({ key, label }) => {
              const order = [STEPS.READING, STEPS.PARAPHRASING, STEPS.BUILDING];
              const cur = order.indexOf(step), idx = order.indexOf(key);
              const done = idx < cur, active = idx === cur;
              return (
                <div key={key} style={{
                  display: "flex", alignItems: "center", gap: 10, padding: "5px 0",
                  opacity: done || active ? 1 : 0.3,
                }}>
                  <div style={{
                    width: 20, height: 20, borderRadius: "50%", flexShrink: 0,
                    background: done ? "#22c55e" : active ? "#3b82f6" : "rgba(255,255,255,0.1)",
                    display: "flex", alignItems: "center", justifyContent: "center",
                    fontSize: 11, color: "#fff", fontWeight: 700,
                  }}>{done ? "✓" : idx + 1}</div>
                  <span style={{ color: done ? "#86efac" : active ? "#93c5fd" : "#64748b", fontSize: 13 }}>{label}</span>
                </div>
              );
            })}
          </div>
        )}

        {/* Log */}
        {log.length > 0 && (
          <div style={{
            background: "rgba(0,0,0,0.35)", borderRadius: 10,
            padding: "1rem 1.2rem", marginBottom: "1.5rem",
            fontFamily: "monospace", fontSize: 12,
            maxHeight: 160, overflowY: "auto",
          }}>
            {log.map((l, i) => (
              <div key={i} style={{ color: "#a5f3fc", lineHeight: 1.8 }}>{l}</div>
            ))}
          </div>
        )}

        {/* Error */}
        {error && (
          <div style={{
            background: "rgba(239,68,68,0.1)", border: "1px solid rgba(239,68,68,0.3)",
            borderRadius: 10, padding: "1rem 1.2rem",
            color: "#fca5a5", fontSize: 13, marginBottom: "1.5rem",
          }}>⚠️ {error}</div>
        )}

        {/* Done */}
        {step === STEPS.DONE && (
          <div style={{
            background: "rgba(34,197,94,0.08)", border: "1px solid rgba(34,197,94,0.3)",
            borderRadius: 14, padding: "1.5rem", textAlign: "center", marginBottom: "1.5rem",
          }}>
            <div style={{ fontSize: 40, marginBottom: 8 }}>🎉</div>
            <div style={{ color: "#86efac", fontWeight: 700, fontSize: 16 }}>PowerPoint Created!</div>
            <div style={{ color: "#4ade80", fontSize: 12, marginTop: 4 }}>{fileName} — check downloads</div>
            {preview && (
              <div style={{ marginTop: 14, textAlign: "left" }}>
                <div style={{ color: "#475569", fontSize: 11, marginBottom: 6, textTransform: "uppercase", letterSpacing: 1 }}>Slides</div>
                {preview.slides.slice(0, 6).map((s, i) => (
                  <div key={i} style={{
                    background: "rgba(255,255,255,0.05)", borderRadius: 7,
                    padding: "5px 10px", marginBottom: 3, color: "#cbd5e1", fontSize: 12,
                  }}>
                    <strong style={{ color: "#93c5fd" }}>#{i + 1}</strong> {s.heading}
                  </div>
                ))}
                {preview.slides.length > 6 && (
                  <div style={{ color: "#475569", fontSize: 11, paddingLeft: 10 }}>
                    + {preview.slides.length - 6} more…
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {(step === STEPS.DONE || step === STEPS.ERROR) && (
          <button onClick={reset} style={{
            width: "100%", padding: "12px", borderRadius: 10,
            background: "linear-gradient(135deg, #3b82f6, #1d4ed8)",
            border: "none", color: "#fff", fontWeight: 700, fontSize: 14,
            cursor: "pointer", boxShadow: "0 4px 16px rgba(59,130,246,0.35)",
          }}>Convert Another PDF</button>
        )}

        {/* How it works */}
        {step === STEPS.IDLE && (
          <div style={{ marginTop: "1.5rem", borderTop: "1px solid rgba(255,255,255,0.06)", paddingTop: "1.2rem" }}>
            <div style={{ color: "#475569", fontSize: 11, textTransform: "uppercase", letterSpacing: 1, marginBottom: 10 }}>How it works</div>
            {[
              ["🔑", "Add keys", "Free Gemini key (required) + optional Groq fallback"],
              ["📄", "Upload PDF", "Drop in any BTEC assignment PDF"],
              ["🔵", "Gemini reads", "Extracts and structures content into slides"],
              ["✍️", "Paraphrase", "All body text rewritten — headings kept verbatim"],
              ["📊", "Download", "A polished .pptx saved to your device instantly"],
            ].map(([icon, title, desc]) => (
              <div key={title} style={{
                display: "flex", alignItems: "flex-start", gap: 10,
                padding: "7px 0", borderBottom: "1px solid rgba(255,255,255,0.04)",
              }}>
                <span style={{ fontSize: 16, flexShrink: 0, marginTop: 1 }}>{icon}</span>
                <div>
                  <span style={{ color: "#e2e8f0", fontWeight: 600, fontSize: 12 }}>{title} — </span>
                  <span style={{ color: "#64748b", fontSize: 12 }}>{desc}</span>
                </div>
              </div>
            ))}
            <div style={{
              marginTop: 14, background: "rgba(34,197,94,0.07)",
              border: "1px solid rgba(34,197,94,0.2)", borderRadius: 8,
              padding: "10px 12px", color: "#86efac", fontSize: 12,
            }}>
              💚 100% free — Gemini gives you 1,500 requests/day at no cost. No card needed.
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
