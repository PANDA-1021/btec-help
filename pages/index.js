import { useState, useRef, useCallback } from "react";
import Head from "next/head";

// ── Step machine ────────────────────────────────────────────────────────────
const S = {
  IDLE: "idle",
  READING: "reading",
  EXTRACTING: "extracting",
  PARAPHRASING: "paraphrasing",
  BUILDING: "building",
  DONE: "done",
  ERROR: "error",
};

const STEP_ORDER = [S.READING, S.EXTRACTING, S.PARAPHRASING, S.BUILDING];
const STEP_LABELS = {
  [S.READING]: "Reading PDF",
  [S.EXTRACTING]: "Extracting structure",
  [S.PARAPHRASING]: "Paraphrasing content",
  [S.BUILDING]: "Building PowerPoint",
};

// ── Themes ──────────────────────────────────────────────────────────────────
const THEMES = {
  navy: { name: "Navy Pro", bg: "1B2B4B", accent: "3B82F6", text: "FFFFFF", light: "F0F4FA", dark: "2D3748" },
  dark: { name: "Dark Slate", bg: "111827", accent: "8B5CF6", text: "F9FAFB", light: "1F2937", dark: "D1D5DB" },
  green: { name: "Emerald", bg: "064E3B", accent: "10B981", text: "ECFDF5", light: "F0FDF4", dark: "1F2937" },
  rose: { name: "Rose Gold", bg: "1C0A0A", accent: "F43F5E", text: "FFF1F2", light: "FFF1F2", dark: "1C1917" },
  slate: { name: "Minimal", bg: "FFFFFF", accent: "475569", text: "0F172A", light: "F8FAFC", dark: "334155" },
};

// ── Gemini API call ─────────────────────────────────────────────────────────
async function callGemini(prompt, base64Pdf = null, opts = {}) {
  const res = await fetch("/api/gemini", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ prompt, base64Pdf, ...opts }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || `Server error ${res.status}`);
  return data.text;
}

// ── JSON parser with robust fallback ────────────────────────────────────────
function parseJSON(raw) {
  const clean = raw
    .replace(/```json\s*/gi, "")
    .replace(/```\s*/g, "")
    .trim();
  try {
    return JSON.parse(clean);
  } catch {
    // Try to grab the first complete {...} block
    const match = clean.match(/\{[\s\S]*\}/);
    if (match) {
      try { return JSON.parse(match[0]); } catch {}
    }
    // Try to grab first [...] block (in case model returned only slides array)
    const arrMatch = clean.match(/\[[\s\S]*\]/);
    if (arrMatch) {
      try { return JSON.parse(arrMatch[0]); } catch {}
    }
    throw new Error("Could not parse Gemini response as JSON. Try a different PDF or reduce slide count.");
  }
}

// ── Prompts ──────────────────────────────────────────────────────────────────
function buildExtractPrompt(targetSlides) {
  return `You are a BTEC document parser. Extract the content of this assignment PDF into structured JSON for a presentation.
Return ONLY valid JSON — no markdown fences, no explanation.

Format exactly:
{
  "title": "Assignment Title",
  "unit": "Unit name/number if found, else empty string",
  "slides": [
    {
      "heading": "Section heading verbatim from document",
      "bullets": ["Key point from this section", "Another point"]
    }
  ]
}

Rules:
- Target exactly ${targetSlides} slides (±2 is acceptable)
- Use ACTUAL headings from the document verbatim
- Each slide: 3–6 concise bullet points capturing core ideas
- Keep bullets close to original phrasing (they will be paraphrased next)
- Do NOT include a title slide (it is added automatically)
- If fewer than ${targetSlides} clear sections exist, split the largest sections`;
}

function buildParaphrasePrompt(slides) {
  return `You are an academic paraphrasing assistant for BTEC students.
Rewrite ONLY the bullet points below in clearly different words while preserving all meaning and technical accuracy.
Do NOT change headings.
Return ONLY valid JSON in the EXACT same structure — no markdown fences, no explanation.

${JSON.stringify({ slides }, null, 2)}`;
}

// ── CDN loader with fallback ─────────────────────────────────────────────────
const PPTX_CDNS = [
  "https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundled.js",
  "https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundled.js",
];

function loadScript(src) {
  return new Promise((resolve, reject) => {
    if (window.PptxGenJS) return resolve();
    const existing = document.querySelector(`script[src="${src}"]`);
    if (existing) {
      existing.addEventListener("load", resolve);
      existing.addEventListener("error", reject);
      return;
    }
    const s = document.createElement("script");
    s.src = src;
    s.async = true;
    s.onload = resolve;
    s.onerror = () => reject(new Error(`CDN failed: ${src}`));
    document.head.appendChild(s);
  });
}

async function loadPptxGen() {
  for (const cdn of PPTX_CDNS) {
    try { await loadScript(cdn); if (window.PptxGenJS) return; } catch {}
  }
  throw new Error("Failed to load pptxgenjs from all CDNs. Check your internet connection.");
}

// ── PPTX builder ────────────────────────────────────────────────────────────
async function buildPptx(data, theme) {
  const T = THEMES[theme] || THEMES.navy;
  await loadPptxGen();

  const pres = new window.PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.title = data.title || "BTEC Assignment";

  const RT = pres.ShapeType?.rect ?? "rect";
  const EL = pres.ShapeType?.ellipse ?? "ellipse";

  // ── Title slide ───────────────────────────────────────────────────────────
  const ts = pres.addSlide();
  ts.background = { color: T.bg };
  ts.addShape(RT, { x: 0, y: 4.2, w: 10, h: 1.425, fill: { color: T.accent }, line: { color: T.accent } });
  ts.addText(data.title || "BTEC Assignment", {
    x: 0.6, y: 0.9, w: 8.8, h: 2.8,
    fontSize: data.title?.length > 60 ? 28 : 36,
    fontFace: "Calibri", bold: true, color: T.text,
    align: "left", valign: "middle", wrap: true,
  });
  if (data.unit) {
    ts.addText(data.unit, {
      x: 0.6, y: 3.7, w: 8.8, h: 0.5,
      fontSize: 13, fontFace: "Calibri", color: T.accent,
      align: "left", valign: "middle",
    });
  }
  ts.addText("Prepared from BTEC Assignment", {
    x: 0.6, y: 4.35, w: 8.8, h: 0.7,
    fontSize: 14, fontFace: "Calibri", color: T.text,
    align: "left", valign: "middle",
  });

  // ── Table of contents slide ───────────────────────────────────────────────
  if (data.slides.length >= 4) {
    const tc = pres.addSlide();
    tc.background = { color: T.light };
    tc.addShape(RT, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: T.accent }, line: { color: T.accent } });
    tc.addText("Contents", {
      x: 0.3, y: 0.1, w: 9, h: 0.8,
      fontSize: 24, fontFace: "Calibri", bold: true, color: T.bg,
      align: "left", valign: "middle",
    });
    tc.addShape(RT, { x: 0.3, y: 0.9, w: 9, h: 0.025, fill: { color: T.accent }, line: { color: T.accent } });

    const cols = data.slides.length > 8 ? 2 : 1;
    const perCol = Math.ceil(data.slides.length / cols);
    data.slides.slice(0, Math.min(data.slides.length, 16)).forEach((sl, i) => {
      const col = Math.floor(i / perCol);
      const row = i % perCol;
      const colW = cols === 2 ? 4.4 : 9;
      const colX = 0.3 + col * 4.7;
      tc.addText(`${i + 1}.  ${sl.heading}`, {
        x: colX, y: 1.05 + row * 0.42, w: colW, h: 0.38,
        fontSize: 11, fontFace: "Calibri", color: T.dark,
        align: "left", valign: "middle",
      });
    });
  }

  // ── Content slides ────────────────────────────────────────────────────────
  data.slides.forEach((slide, idx) => {
    const s = pres.addSlide();
    s.background = { color: idx % 2 === 0 ? "FFFFFF" : T.light };
    s.addShape(RT, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: T.accent }, line: { color: T.accent } });
    s.addShape(EL, { x: 8.9, y: 0.15, w: 0.6, h: 0.6, fill: { color: T.bg }, line: { color: T.bg } });
    s.addText(String(idx + 1), {
      x: 8.9, y: 0.15, w: 0.6, h: 0.6,
      fontSize: 11, fontFace: "Calibri", bold: true, color: "FFFFFF",
      align: "center", valign: "middle",
    });

    const headingLen = (slide.heading || "").length;
    s.addText(slide.heading || "", {
      x: 0.3, y: 0.1, w: 8.5, h: headingLen > 60 ? 1.1 : 0.85,
      fontSize: headingLen > 60 ? 17 : 22,
      fontFace: "Calibri", bold: true, color: T.bg,
      align: "left", valign: "middle", wrap: true,
    });

    const divY = headingLen > 60 ? 1.2 : 0.97;
    s.addShape(RT, { x: 0.3, y: divY, w: 8.5, h: 0.025, fill: { color: T.accent }, line: { color: T.accent } });

    const bullets = (slide.bullets || []).slice(0, 7);
    if (bullets.length > 0) {
      const textY = divY + 0.1;
      const textH = 5.625 - textY - 0.1;
      const fs = bullets.length > 5 ? 13 : 15;
      s.addText(
        bullets.map((b, i) => ({
          text: b,
          options: {
            bullet: true,
            breakLine: i < bullets.length - 1,
            fontSize: fs, fontFace: "Calibri", color: T.dark,
            paraSpaceAfter: bullets.length > 5 ? 4 : 6,
          },
        })),
        { x: 0.3, y: textY, w: 9.2, h: textH, valign: "top", wrap: true }
      );
    }
  });

  // ── End slide ─────────────────────────────────────────────────────────────
  const es = pres.addSlide();
  es.background = { color: T.bg };
  es.addShape(RT, { x: 3.5, y: 2.55, w: 3, h: 0.06, fill: { color: T.accent }, line: { color: T.accent } });
  es.addText("Thank You", {
    x: 0, y: 1.5, w: 10, h: 1.5,
    fontSize: 44, fontFace: "Calibri", bold: true, color: T.text,
    align: "center", valign: "middle",
  });
  es.addText("Generated from BTEC Assignment", {
    x: 0, y: 3.2, w: 10, h: 0.6,
    fontSize: 14, fontFace: "Calibri", color: T.accent, align: "center",
  });

  // ── Save ──────────────────────────────────────────────────────────────────
  const safeName = `${(data.title || "BTEC_Assignment")
    .replace(/[^a-zA-Z0-9 ]/g, " ")
    .replace(/\s+/g, "_")
    .slice(0, 55)}.pptx`;

  const blob = await pres.write({ outputType: "blob" });
  if (!blob || blob.size < 100) throw new Error("PowerPoint file appears empty. Please try again.");

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = safeName;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 15000);
  return safeName;
}

// ── CSS vars ─────────────────────────────────────────────────────────────────
const C = {
  bg: "rgba(255,255,255,0.05)",
  border: "rgba(255,255,255,0.12)",
  blue: "#3b82f6",
  blueDeep: "#1d4ed8",
  text: "#f0f4fa",
  muted: "#94a3b8",
  dim: "#64748b",
  green: "#22c55e",
  greenLight: "#86efac",
  red: "#ef4444",
  redLight: "#fca5a5",
  cyan: "#a5f3fc",
  logBg: "rgba(0,0,0,0.35)",
};

// ── Component ─────────────────────────────────────────────────────────────────
export default function Home() {
  const [step, setStep] = useState(S.IDLE);
  const [log, setLog] = useState([]);
  const [error, setError] = useState("");
  const [preview, setPreview] = useState(null);
  const [fileName, setFileName] = useState("");
  const [dragging, setDragging] = useState(false);
  const [slideCount, setSlideCount] = useState(10);
  const [theme, setTheme] = useState("navy");
  const [progress, setProgress] = useState(0); // 0–100
  const fileRef = useRef(null);
  const logRef = useRef(null);

  const addLog = useCallback((msg) => {
    setLog((l) => [...l, msg]);
    setTimeout(() => { if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight; }, 50);
  }, []);

  const reset = () => {
    setStep(S.IDLE); setLog([]); setError(""); setPreview(null);
    setFileName(""); setProgress(0);
    if (fileRef.current) fileRef.current.value = "";
  };

  const processFile = useCallback(async (file) => {
    if (!file) return;
    if (file.type !== "application/pdf") { setError("Please upload a PDF file."); return; }
    if (file.size > 19 * 1024 * 1024) { setError("PDF is too large (max 19 MB). Try compressing it first."); return; }

    setError(""); setLog([]); setPreview(null); setProgress(0);

    try {
      // Step 1 — Read PDF
      setStep(S.READING); setProgress(5);
      addLog(`📄 Reading "${file.name}" (${(file.size / 1024).toFixed(0)} KB)…`);
      const base64 = await new Promise((res, rej) => {
        const r = new FileReader();
        r.onload = () => res(r.result.split(",")[1]);
        r.onerror = () => rej(new Error("Failed to read file"));
        r.readAsDataURL(file);
      });
      setProgress(15);

      // Pre-load pptxgenjs in parallel while Gemini works
      loadPptxGen().catch(() => {});

      // Step 2 — Extract
      setStep(S.EXTRACTING); setProgress(20);
      addLog(`🔵 Sending to Gemini — extracting ${slideCount} slides…`);
      const rawExtract = await callGemini(buildExtractPrompt(slideCount), base64, { temperature: 0.2 });
      setProgress(50);

      const structure = parseJSON(rawExtract);
      if (!structure.slides || structure.slides.length === 0) {
        throw new Error("Gemini returned no slides. The PDF may be image-only, encrypted, or too short.");
      }
      addLog(`✅ Extracted ${structure.slides.length} sections — "${structure.title}"`);
      if (structure.unit) addLog(`📚 Unit: ${structure.unit}`);

      // Step 3 — Paraphrase
      setStep(S.PARAPHRASING); setProgress(55);
      addLog("✍️  Paraphrasing body content…");
      const rawPara = await callGemini(buildParaphrasePrompt(structure.slides), null, { temperature: 0.7 });
      setProgress(80);

      let paraphrased;
      try {
        const parsed = parseJSON(rawPara);
        // Model may return { slides: [...] } or directly [...]
        paraphrased = Array.isArray(parsed) ? parsed : parsed.slides || parsed;
        if (!Array.isArray(paraphrased)) throw new Error("unexpected shape");
      } catch {
        addLog("⚠️  Paraphrase parse failed — using original bullets.");
        paraphrased = structure.slides;
      }

      const finalData = { title: structure.title, unit: structure.unit || "", slides: paraphrased };
      setPreview(finalData);
      addLog("✅ Paraphrasing complete.");

      // Step 4 — Build PPTX
      setStep(S.BUILDING); setProgress(85);
      addLog("📊 Building PowerPoint…");
      const fn = await buildPptx(finalData, theme);
      setFileName(fn);
      setProgress(100);
      addLog(`🎉 Done! "${fn}" saved to Downloads.`);
      setStep(S.DONE);
    } catch (err) {
      setError(err.message || "Something went wrong. Please try again.");
      setStep(S.ERROR);
    }
  }, [slideCount, theme, addLog]);

  const handleFile = useCallback((e) => processFile(e.target.files?.[0]), [processFile]);

  const handleDrop = useCallback((e) => {
    e.preventDefault(); setDragging(false);
    processFile(e.dataTransfer.files?.[0]);
  }, [processFile]);

  const busy = [S.READING, S.EXTRACTING, S.PARAPHRASING, S.BUILDING].includes(step);

  return (
    <>
      <Head>
        <title>BTEC → PowerPoint</title>
        <meta name="description" content="Turn your BTEC assignment PDF into a paraphrased PowerPoint" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
      </Head>

      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", padding: "2rem" }}>
        <div style={{
          width: "100%", maxWidth: 660,
          background: C.bg, backdropFilter: "blur(20px)",
          borderRadius: 20, border: `1px solid ${C.border}`,
          padding: "2.5rem", boxShadow: "0 25px 60px rgba(0,0,0,0.5)",
        }}>
          {/* Header */}
          <div style={{ textAlign: "center", marginBottom: "2rem" }}>
            <div style={{
              display: "inline-flex", alignItems: "center", justifyContent: "center",
              width: 64, height: 64, borderRadius: 16,
              background: "linear-gradient(135deg, #3b82f6, #1d4ed8)",
              marginBottom: "1rem", boxShadow: "0 8px 24px rgba(59,130,246,0.4)",
            }}>
              <span style={{ fontSize: 30 }}>📑</span>
            </div>
            <h1 style={{ margin: 0, fontSize: 26, fontWeight: 700, color: C.text }}>BTEC → PowerPoint</h1>
            <p style={{ margin: "0.5rem 0 0", color: C.muted, fontSize: 13 }}>
              Upload your assignment PDF · Get a paraphrased presentation instantly
            </p>
          </div>

          {/* Options (idle only) */}
          {step === S.IDLE && (
            <div style={{ display: "flex", gap: 12, marginBottom: "1.5rem", flexWrap: "wrap" }}>
              <div style={{ flex: 1, minWidth: 140 }}>
                <label style={{ display: "block", color: C.muted, fontSize: 11, marginBottom: 5, textTransform: "uppercase", letterSpacing: 0.8 }}>
                  Slides (target)
                </label>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  {[6, 8, 10, 12, 14].map(n => (
                    <button key={n} onClick={() => setSlideCount(n)} style={{
                      padding: "5px 10px", borderRadius: 8, border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600,
                      background: slideCount === n ? C.blue : "rgba(255,255,255,0.08)",
                      color: slideCount === n ? "#fff" : C.muted,
                      transition: "all 0.15s",
                    }}>{n}</button>
                  ))}
                </div>
              </div>

              <div style={{ flex: 1, minWidth: 200 }}>
                <label style={{ display: "block", color: C.muted, fontSize: 11, marginBottom: 5, textTransform: "uppercase", letterSpacing: 0.8 }}>
                  Theme
                </label>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                  {Object.entries(THEMES).map(([k, t]) => (
                    <button key={k} onClick={() => setTheme(k)} style={{
                      padding: "5px 10px", borderRadius: 8, border: theme === k ? `2px solid ${C.blue}` : "2px solid transparent",
                      cursor: "pointer", fontSize: 12, fontWeight: 600,
                      background: `#${t.bg}`, color: `#${t.text}`,
                      boxShadow: theme === k ? `0 0 0 2px ${C.blue}` : "none",
                      transition: "all 0.15s",
                    }}>{t.name}</button>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* Drop zone */}
          {!busy && step !== S.DONE && (
            <label
              onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              onDrop={handleDrop}
              style={{
                display: "block",
                border: `2px dashed ${dragging ? C.blue : "rgba(59,130,246,0.5)"}`,
                borderRadius: 14, padding: "2.5rem 1.5rem",
                textAlign: "center", cursor: "pointer",
                background: dragging ? "rgba(59,130,246,0.15)" : "rgba(59,130,246,0.04)",
                marginBottom: "1.5rem", transition: "background 0.2s, border-color 0.2s",
              }}
              onMouseEnter={e => { e.currentTarget.style.background = "rgba(59,130,246,0.1)"; e.currentTarget.style.borderColor = C.blue; }}
              onMouseLeave={e => { if (!dragging) { e.currentTarget.style.background = "rgba(59,130,246,0.04)"; e.currentTarget.style.borderColor = "rgba(59,130,246,0.5)"; } }}
            >
              <input ref={fileRef} type="file" accept="application/pdf" style={{ display: "none" }} onChange={handleFile} />
              <div style={{ fontSize: 40, marginBottom: 10 }}>{dragging ? "⬇️" : "📂"}</div>
              <div style={{ color: "#cbd5e1", fontSize: 15, fontWeight: 600 }}>
                {dragging ? "Drop to upload" : "Click or drag & drop BTEC PDF"}
              </div>
              <div style={{ color: C.dim, fontSize: 12, marginTop: 4 }}>
                PDF only · Max 19 MB · {slideCount} slides · {THEMES[theme].name} theme
              </div>
            </label>
          )}

          {/* Progress bar */}
          {busy && (
            <div style={{ marginBottom: "1.5rem" }}>
              <div style={{
                height: 6, borderRadius: 3, background: "rgba(255,255,255,0.1)",
                overflow: "hidden", marginBottom: "1rem",
              }}>
                <div style={{
                  height: "100%", width: `${progress}%`,
                  background: `linear-gradient(90deg, ${C.blue}, ${C.blueDeep})`,
                  transition: "width 0.4s ease", borderRadius: 3,
                }} />
              </div>

              <div style={{
                background: "rgba(59,130,246,0.08)", border: "1px solid rgba(59,130,246,0.2)",
                borderRadius: 12, padding: "1.2rem 1.5rem",
              }}>
                <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
                  <div style={{
                    width: 20, height: 20, border: "3px solid #3b82f6", borderTopColor: "transparent",
                    borderRadius: "50%", animation: "spin 0.8s linear infinite", flexShrink: 0,
                  }} />
                  <span style={{ color: "#93c5fd", fontWeight: 600, fontSize: 14 }}>{STEP_LABELS[step]}…</span>
                  <span style={{ marginLeft: "auto", color: C.dim, fontSize: 12 }}>{progress}%</span>
                </div>

                {STEP_ORDER.map((k, i) => {
                  const cur = STEP_ORDER.indexOf(step);
                  const done = i < cur, active = i === cur;
                  return (
                    <div key={k} style={{ display: "flex", alignItems: "center", gap: 10, padding: "4px 0", opacity: done || active ? 1 : 0.3 }}>
                      <div style={{
                        width: 22, height: 22, borderRadius: "50%", flexShrink: 0,
                        background: done ? C.green : active ? C.blue : "rgba(255,255,255,0.1)",
                        display: "flex", alignItems: "center", justifyContent: "center",
                        fontSize: 11, color: "#fff", fontWeight: 700,
                        transition: "background 0.3s",
                      }}>{done ? "✓" : i + 1}</div>
                      <span style={{ color: done ? C.greenLight : active ? "#93c5fd" : C.dim, fontSize: 13 }}>
                        {STEP_LABELS[k]}
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* Log */}
          {log.length > 0 && (
            <div ref={logRef} style={{
              background: C.logBg, borderRadius: 10, padding: "0.9rem 1.1rem",
              marginBottom: "1.5rem", fontFamily: "monospace", fontSize: 12,
              maxHeight: 150, overflowY: "auto",
            }}>
              {log.map((l, i) => (
                <div key={i} style={{ color: C.cyan, lineHeight: 1.85 }}>{l}</div>
              ))}
            </div>
          )}

          {/* Error */}
          {error && (
            <div style={{
              background: "rgba(239,68,68,0.1)", border: "1px solid rgba(239,68,68,0.3)",
              borderRadius: 10, padding: "1rem 1.2rem", color: C.redLight,
              fontSize: 13, marginBottom: "1.5rem",
            }}>
              ⚠️ {error}
            </div>
          )}

          {/* Done */}
          {step === S.DONE && (
            <div style={{
              background: "rgba(34,197,94,0.08)", border: "1px solid rgba(34,197,94,0.3)",
              borderRadius: 14, padding: "1.5rem", marginBottom: "1.5rem",
            }}>
              <div style={{ textAlign: "center", marginBottom: 12 }}>
                <div style={{ fontSize: 36, marginBottom: 6 }}>🎉</div>
                <div style={{ color: C.greenLight, fontWeight: 700, fontSize: 16 }}>PowerPoint Created!</div>
                <div style={{ color: "#4ade80", fontSize: 12, marginTop: 4 }}>{fileName} — check Downloads</div>
              </div>

              {preview && (
                <>
                  <div style={{ color: C.dim, fontSize: 11, textTransform: "uppercase", letterSpacing: 1, marginBottom: 8 }}>
                    {preview.slides.length} slides generated
                  </div>
                  <div style={{ maxHeight: 200, overflowY: "auto" }}>
                    {preview.slides.map((s, i) => (
                      <div key={i} style={{
                        background: "rgba(255,255,255,0.05)", borderRadius: 7,
                        padding: "5px 10px", marginBottom: 3,
                        color: "#cbd5e1", fontSize: 12,
                      }}>
                        <strong style={{ color: "#93c5fd" }}>#{i + 1}</strong>{" "}{s.heading}
                        <span style={{ color: C.dim, marginLeft: 6 }}>({(s.bullets || []).length} bullets)</span>
                      </div>
                    ))}
                  </div>
                </>
              )}
            </div>
          )}

          {/* Action buttons */}
          {(step === S.DONE || step === S.ERROR) && (
            <button onClick={reset} style={{
              width: "100%", padding: "12px", borderRadius: 10,
              background: "linear-gradient(135deg, #3b82f6, #1d4ed8)",
              border: "none", color: "#fff", fontWeight: 700, fontSize: 14,
              cursor: "pointer", boxShadow: "0 4px 16px rgba(59,130,246,0.35)",
              transition: "opacity 0.2s",
            }}
              onMouseEnter={e => e.currentTarget.style.opacity = "0.9"}
              onMouseLeave={e => e.currentTarget.style.opacity = "1"}
            >
              Convert Another PDF
            </button>
          )}

          {/* How it works */}
          {step === S.IDLE && (
            <div style={{ marginTop: "1.5rem", borderTop: "1px solid rgba(255,255,255,0.06)", paddingTop: "1.2rem" }}>
              <div style={{ color: C.dim, fontSize: 11, textTransform: "uppercase", letterSpacing: 1, marginBottom: 10 }}>
                How it works
              </div>
              {[
                ["📄", "Upload", "Drop in any BTEC assignment PDF (max 19 MB)"],
                ["🔵", "Extract", "Gemini reads and structures the content into slides"],
                ["✍️", "Paraphrase", "All body text rewritten — headings kept verbatim"],
                ["📊", "Download", "A polished .pptx file saved to your device instantly"],
              ].map(([icon, title, desc]) => (
                <div key={title} style={{
                  display: "flex", alignItems: "flex-start", gap: 10,
                  padding: "7px 0", borderBottom: "1px solid rgba(255,255,255,0.04)",
                }}>
                  <span style={{ fontSize: 16, flexShrink: 0, marginTop: 1 }}>{icon}</span>
                  <div>
                    <span style={{ color: "#e2e8f0", fontWeight: 600, fontSize: 12 }}>{title} — </span>
                    <span style={{ color: C.dim, fontSize: 12 }}>{desc}</span>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      <style>{`
        @keyframes spin { to { transform: rotate(360deg); } }
        ::-webkit-scrollbar { width: 4px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.15); border-radius: 2px; }
      `}</style>
    </>
  );
}
