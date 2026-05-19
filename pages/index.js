import { useState, useRef } from "react";
import Head from "next/head";

const STEPS = {
  IDLE: "idle",
  READING: "reading",
  PARAPHRASING: "paraphrasing",
  BUILDING: "building",
  DONE: "done",
  ERROR: "error",
};

// ── Call our own Next.js API route (key stays server-side) ───────────────────

async function callGemini(prompt, base64Pdf = null) {
  const res = await fetch("/api/gemini", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ prompt, base64Pdf }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || `Server error ${res.status}`);
  return data.text;
}

// ── JSON parser ──────────────────────────────────────────────────────────────

function parseJSON(raw) {
  // Strip markdown fences if Gemini wraps the response
  const clean = raw
    .replace(/```json\s*/gi, "")
    .replace(/```\s*/g, "")
    .trim();
  try {
    return JSON.parse(clean);
  } catch {
    // Try to extract the first {...} block as fallback
    const match = clean.match(/\{[\s\S]*\}/);
    if (match) return JSON.parse(match[0]);
    throw new Error(
      "Could not parse JSON from Gemini response. Try a different PDF.",
    );
  }
}

// ── Prompts ──────────────────────────────────────────────────────────────────

const EXTRACT_PROMPT = `You are a document parser. Extract the full content of this BTEC assignment into JSON.
Return ONLY valid JSON — no markdown fences, no explanation, no preamble.

Format exactly:
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

function buildParaphrasePrompt(slides) {
  return `You are an academic paraphrasing assistant.
Paraphrase ONLY the bullet points in the JSON below — rewrite them in clearly different words while fully preserving the meaning.
Do NOT change the headings at all.
Return ONLY valid JSON in exactly the same structure — no markdown fences, no explanation, no preamble.

${JSON.stringify(slides, null, 2)}`;
}

// ── PPTX builder (client-side via pptxgenjs CDN) ─────────────────────────────

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
    s.onerror = () =>
      reject(
        new Error(
          "Failed to load pptxgenjs from CDN. Check your internet connection.",
        ),
      );
    document.head.appendChild(s);
  });
}

async function buildPptx(data) {
  // Try primary CDN, fall back to jsDelivr
  try {
    await loadScript(
      "https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundled.js",
    );
  } catch {
    await loadScript(
      "https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundled.js",
    );
  }

  if (!window.PptxGenJS) {
    throw new Error(
      "pptxgenjs library did not load. Please refresh and try again.",
    );
  }

  const pres = new window.PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.title = data.title || "BTEC Assignment";

  const NAVY = "1B2B4B";
  const ACCENT = "3B82F6";
  const WHITE = "FFFFFF";
  const LIGHT = "F0F4FA";

  // ── Title slide ──────────────────────────────────────────────────────────
  const titleSlide = pres.addSlide();
  titleSlide.background = { color: NAVY };
  titleSlide.addShape(pres.ShapeType?.rect ?? "rect", {
    x: 0,
    y: 4.2,
    w: 10,
    h: 1.425,
    fill: { color: ACCENT },
    line: { color: ACCENT },
  });
  titleSlide.addText(data.title || "BTEC Assignment", {
    x: 0.6,
    y: 1.0,
    w: 8.8,
    h: 2.8,
    fontSize: 36,
    fontFace: "Calibri",
    bold: true,
    color: WHITE,
    align: "left",
    valign: "middle",
    wrap: true,
  });
  titleSlide.addText("Prepared from BTEC Assignment", {
    x: 0.6,
    y: 4.35,
    w: 8.8,
    h: 0.7,
    fontSize: 14,
    fontFace: "Calibri",
    color: WHITE,
    align: "left",
    valign: "middle",
  });

  // ── Content slides ───────────────────────────────────────────────────────
  (data.slides || []).forEach((slide, idx) => {
    const s = pres.addSlide();
    s.background = { color: idx % 2 === 0 ? WHITE : LIGHT };

    // Left accent bar
    s.addShape(pres.ShapeType?.rect ?? "rect", {
      x: 0,
      y: 0,
      w: 0.12,
      h: 5.625,
      fill: { color: ACCENT },
      line: { color: ACCENT },
    });
    // Slide number badge
    s.addShape(pres.ShapeType?.ellipse ?? "ellipse", {
      x: 8.9,
      y: 0.15,
      w: 0.6,
      h: 0.6,
      fill: { color: NAVY },
      line: { color: NAVY },
    });
    s.addText(String(idx + 1), {
      x: 8.9,
      y: 0.15,
      w: 0.6,
      h: 0.6,
      fontSize: 11,
      fontFace: "Calibri",
      bold: true,
      color: WHITE,
      align: "center",
      valign: "middle",
    });
    // Heading
    s.addText(slide.heading || "", {
      x: 0.3,
      y: 0.1,
      w: 8.5,
      h: 0.85,
      fontSize: 22,
      fontFace: "Calibri",
      bold: true,
      color: NAVY,
      align: "left",
      valign: "middle",
    });
    // Divider line
    s.addShape(pres.ShapeType?.rect ?? "rect", {
      x: 0.3,
      y: 0.95,
      w: 8.5,
      h: 0.025,
      fill: { color: ACCENT },
      line: { color: ACCENT },
    });

    const bullets = (slide.bullets || []).slice(0, 7);
    if (bullets.length > 0) {
      s.addText(
        bullets.map((b, i) => ({
          text: b,
          options: {
            bullet: true,
            breakLine: i < bullets.length - 1,
            fontSize: 15,
            fontFace: "Calibri",
            color: "2D3748",
            paraSpaceAfter: 6,
          },
        })),
        { x: 0.3, y: 1.1, w: 9.2, h: 4.2, valign: "top", wrap: true },
      );
    }
  });

  // ── End slide ────────────────────────────────────────────────────────────
  const endSlide = pres.addSlide();
  endSlide.background = { color: NAVY };
  endSlide.addShape(pres.ShapeType?.rect ?? "rect", {
    x: 3.5,
    y: 2.55,
    w: 3,
    h: 0.06,
    fill: { color: ACCENT },
    line: { color: ACCENT },
  });
  endSlide.addText("Thank You", {
    x: 0,
    y: 1.5,
    w: 10,
    h: 1.5,
    fontSize: 44,
    fontFace: "Calibri",
    bold: true,
    color: WHITE,
    align: "center",
    valign: "middle",
  });
  endSlide.addText("Generated from BTEC Assignment", {
    x: 0,
    y: 3.2,
    w: 10,
    h: 0.6,
    fontSize: 14,
    fontFace: "Calibri",
    color: ACCENT,
    align: "center",
  });

  // ── Save ─────────────────────────────────────────────────────────────────
  const filename = `${(data.title || "BTEC_Assignment")
    .replace(/[^a-zA-Z0-9]/g, "_")
    .replace(/_+/g, "_")
    .slice(0, 60)}.pptx`;

  // pres.write() returns a Blob when outputType is "blob"
  const blob = await pres.write({ outputType: "blob" });
  if (!blob || blob.size === 0) {
    throw new Error(
      "PowerPoint file was created but appears empty. Please try again.",
    );
  }

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 10000);
  return filename;
}

// ── Main component ────────────────────────────────────────────────────────────

export default function Home() {
  const [step, setStep] = useState(STEPS.IDLE);
  const [log, setLog] = useState([]);
  const [error, setError] = useState("");
  const [preview, setPreview] = useState(null);
  const [fileName, setFileName] = useState("");
  const fileRef = useRef(null);

  const addLog = (msg) => setLog((l) => [...l, msg]);

  const reset = () => {
    setStep(STEPS.IDLE);
    setLog([]);
    setError("");
    setPreview(null);
    setFileName("");
    if (fileRef.current) fileRef.current.value = "";
  };

  const handleFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.type !== "application/pdf") {
      setError("Please upload a PDF file.");
      return;
    }

    setError("");
    setLog([]);
    setPreview(null);

    try {
      // 1. Read PDF as base64
      setStep(STEPS.READING);
      addLog("📄 Reading PDF…");
      const base64 = await new Promise((res, rej) => {
        const r = new FileReader();
        r.onload = () => res(r.result.split(",")[1]);
        r.onerror = () => rej(new Error("Failed to read file"));
        r.readAsDataURL(file);
      });

      // 2. Extract structure with Gemini
      addLog("🔵 Sending to Gemini to extract slide structure…");
      const rawExtract = await callGemini(EXTRACT_PROMPT, base64);
      const structure = parseJSON(rawExtract);
      if (!structure.slides || structure.slides.length === 0) {
        throw new Error(
          "Gemini returned no slides. The PDF might be image-only or too short.",
        );
      }
      addLog(
        `✅ Found ${structure.slides.length} sections — "${structure.title}"`,
      );

      // 3. Paraphrase bullet points
      setStep(STEPS.PARAPHRASING);
      addLog("✍️  Paraphrasing body content (headings kept verbatim)…");
      const rawPara = await callGemini(buildParaphrasePrompt(structure.slides));
      const paraphrased = parseJSON(rawPara);
      const finalData = { title: structure.title, slides: paraphrased };
      setPreview(finalData);
      addLog("✅ Paraphrasing complete.");

      // 4. Build PPTX client-side
      setStep(STEPS.BUILDING);
      addLog("📊 Building PowerPoint file…");
      const fn = await buildPptx(finalData);
      setFileName(fn);
      addLog(`🎉 Done! "${fn}" saved to your Downloads folder.`);
      setStep(STEPS.DONE);
    } catch (err) {
      setError(err.message || "Something went wrong. Please try again.");
      setStep(STEPS.ERROR);
    }
  };

  const busy = [STEPS.READING, STEPS.PARAPHRASING, STEPS.BUILDING].includes(
    step,
  );
  const stepLabel = {
    [STEPS.READING]: "Reading PDF with Gemini…",
    [STEPS.PARAPHRASING]: "Paraphrasing content…",
    [STEPS.BUILDING]: "Building PowerPoint…",
  }[step];

  return (
    <>
      <Head>
        <title>BTEC → PowerPoint</title>
        <meta
          name="description"
          content="Turn your BTEC assignment PDF into a paraphrased PowerPoint"
        />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
      </Head>

      <div
        style={{
          minHeight: "100vh",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          padding: "2rem",
        }}
      >
        <div
          style={{
            width: "100%",
            maxWidth: 640,
            background: "rgba(255,255,255,0.05)",
            backdropFilter: "blur(20px)",
            borderRadius: 20,
            border: "1px solid rgba(255,255,255,0.12)",
            padding: "2.5rem",
            boxShadow: "0 25px 60px rgba(0,0,0,0.5)",
          }}
        >
          {/* Header */}
          <div style={{ textAlign: "center", marginBottom: "2rem" }}>
            <div
              style={{
                display: "inline-flex",
                alignItems: "center",
                justifyContent: "center",
                width: 64,
                height: 64,
                borderRadius: 16,
                background: "linear-gradient(135deg, #3b82f6, #1d4ed8)",
                marginBottom: "1rem",
                boxShadow: "0 8px 24px rgba(59,130,246,0.4)",
              }}
            >
              <span style={{ fontSize: 30 }}>📑</span>
            </div>
            <h1
              style={{
                margin: 0,
                fontSize: 26,
                fontWeight: 700,
                color: "#f0f4fa",
              }}
            >
              BTEC → PowerPoint
            </h1>
            <p style={{ margin: "0.5rem 0 0", color: "#94a3b8", fontSize: 13 }}>
              Upload your assignment PDF · Get a paraphrased presentation
              instantly
            </p>
          </div>

          {/* Upload zone */}
          {!busy && step !== STEPS.DONE && (
            <label
              style={{
                display: "block",
                border: "2px dashed rgba(59,130,246,0.5)",
                borderRadius: 14,
                padding: "2.5rem 1.5rem",
                textAlign: "center",
                cursor: "pointer",
                background: "rgba(59,130,246,0.04)",
                marginBottom: "1.5rem",
                transition: "background 0.2s, border-color 0.2s",
              }}
              onMouseEnter={(e) => {
                e.currentTarget.style.background = "rgba(59,130,246,0.1)";
                e.currentTarget.style.borderColor = "#3b82f6";
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.background = "rgba(59,130,246,0.04)";
                e.currentTarget.style.borderColor = "rgba(59,130,246,0.5)";
              }}
            >
              <input
                ref={fileRef}
                type="file"
                accept="application/pdf"
                style={{ display: "none" }}
                onChange={handleFile}
                disabled={busy}
              />
              <div style={{ fontSize: 40, marginBottom: 10 }}>📂</div>
              <div style={{ color: "#cbd5e1", fontSize: 15, fontWeight: 600 }}>
                Click to upload BTEC PDF
              </div>
              <div style={{ color: "#64748b", fontSize: 12, marginTop: 4 }}>
                PDF only · Headings kept verbatim · Body fully paraphrased
              </div>
            </label>
          )}

          {/* Progress tracker */}
          {busy && (
            <div
              style={{
                background: "rgba(59,130,246,0.08)",
                border: "1px solid rgba(59,130,246,0.2)",
                borderRadius: 12,
                padding: "1.5rem",
                marginBottom: "1.5rem",
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 12,
                  marginBottom: 14,
                }}
              >
                <div
                  style={{
                    width: 20,
                    height: 20,
                    border: "3px solid #3b82f6",
                    borderTopColor: "transparent",
                    borderRadius: "50%",
                    animation: "spin 0.8s linear infinite",
                  }}
                />
                <span
                  style={{ color: "#93c5fd", fontWeight: 600, fontSize: 14 }}
                >
                  {stepLabel}
                </span>
              </div>
              {[
                { key: STEPS.READING, label: "Read & extract PDF structure" },
                { key: STEPS.PARAPHRASING, label: "Paraphrase body content" },
                { key: STEPS.BUILDING, label: "Build PowerPoint file" },
              ].map(({ key, label }) => {
                const order = [
                  STEPS.READING,
                  STEPS.PARAPHRASING,
                  STEPS.BUILDING,
                ];
                const cur = order.indexOf(step),
                  i = order.indexOf(key);
                const done = i < cur,
                  active = i === cur;
                return (
                  <div
                    key={key}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 10,
                      padding: "5px 0",
                      opacity: done || active ? 1 : 0.3,
                    }}
                  >
                    <div
                      style={{
                        width: 20,
                        height: 20,
                        borderRadius: "50%",
                        flexShrink: 0,
                        background: done
                          ? "#22c55e"
                          : active
                            ? "#3b82f6"
                            : "rgba(255,255,255,0.1)",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        fontSize: 11,
                        color: "#fff",
                        fontWeight: 700,
                      }}
                    >
                      {done ? "✓" : i + 1}
                    </div>
                    <span
                      style={{
                        color: done
                          ? "#86efac"
                          : active
                            ? "#93c5fd"
                            : "#64748b",
                        fontSize: 13,
                      }}
                    >
                      {label}
                    </span>
                  </div>
                );
              })}
            </div>
          )}

          {/* Log */}
          {log.length > 0 && (
            <div
              style={{
                background: "rgba(0,0,0,0.35)",
                borderRadius: 10,
                padding: "1rem 1.2rem",
                marginBottom: "1.5rem",
                fontFamily: "monospace",
                fontSize: 12,
                maxHeight: 160,
                overflowY: "auto",
              }}
            >
              {log.map((l, i) => (
                <div key={i} style={{ color: "#a5f3fc", lineHeight: 1.8 }}>
                  {l}
                </div>
              ))}
            </div>
          )}

          {/* Error */}
          {error && (
            <div
              style={{
                background: "rgba(239,68,68,0.1)",
                border: "1px solid rgba(239,68,68,0.3)",
                borderRadius: 10,
                padding: "1rem 1.2rem",
                color: "#fca5a5",
                fontSize: 13,
                marginBottom: "1.5rem",
              }}
            >
              ⚠️ {error}
            </div>
          )}

          {/* Done */}
          {step === STEPS.DONE && (
            <div
              style={{
                background: "rgba(34,197,94,0.08)",
                border: "1px solid rgba(34,197,94,0.3)",
                borderRadius: 14,
                padding: "1.5rem",
                textAlign: "center",
                marginBottom: "1.5rem",
              }}
            >
              <div style={{ fontSize: 40, marginBottom: 8 }}>🎉</div>
              <div style={{ color: "#86efac", fontWeight: 700, fontSize: 16 }}>
                PowerPoint Created!
              </div>
              <div style={{ color: "#4ade80", fontSize: 12, marginTop: 4 }}>
                {fileName} — check your Downloads folder
              </div>
              {preview && (
                <div style={{ marginTop: 14, textAlign: "left" }}>
                  <div
                    style={{
                      color: "#475569",
                      fontSize: 11,
                      marginBottom: 6,
                      textTransform: "uppercase",
                      letterSpacing: 1,
                    }}
                  >
                    Slides generated
                  </div>
                  {preview.slides.slice(0, 6).map((s, i) => (
                    <div
                      key={i}
                      style={{
                        background: "rgba(255,255,255,0.05)",
                        borderRadius: 7,
                        padding: "5px 10px",
                        marginBottom: 3,
                        color: "#cbd5e1",
                        fontSize: 12,
                      }}
                    >
                      <strong style={{ color: "#93c5fd" }}>#{i + 1}</strong>{" "}
                      {s.heading}
                    </div>
                  ))}
                  {preview.slides.length > 6 && (
                    <div
                      style={{
                        color: "#475569",
                        fontSize: 11,
                        paddingLeft: 10,
                      }}
                    >
                      + {preview.slides.length - 6} more…
                    </div>
                  )}
                </div>
              )}
            </div>
          )}

          {(step === STEPS.DONE || step === STEPS.ERROR) && (
            <button
              onClick={reset}
              style={{
                width: "100%",
                padding: "12px",
                borderRadius: 10,
                background: "linear-gradient(135deg, #3b82f6, #1d4ed8)",
                border: "none",
                color: "#fff",
                fontWeight: 700,
                fontSize: 14,
                cursor: "pointer",
                boxShadow: "0 4px 16px rgba(59,130,246,0.35)",
              }}
            >
              Convert Another PDF
            </button>
          )}

          {/* How it works */}
          {step === STEPS.IDLE && (
            <div
              style={{
                marginTop: "1.5rem",
                borderTop: "1px solid rgba(255,255,255,0.06)",
                paddingTop: "1.2rem",
              }}
            >
              <div
                style={{
                  color: "#475569",
                  fontSize: 11,
                  textTransform: "uppercase",
                  letterSpacing: 1,
                  marginBottom: 10,
                }}
              >
                How it works
              </div>
              {[
                ["📄", "Upload", "Drop in any BTEC assignment PDF"],
                [
                  "🔵",
                  "Extract",
                  "Gemini reads and structures the content into slides",
                ],
                [
                  "✍️",
                  "Paraphrase",
                  "All body text rewritten — headings kept verbatim",
                ],
                [
                  "📊",
                  "Download",
                  "A polished .pptx file saved to your device instantly",
                ],
              ].map(([icon, title, desc]) => (
                <div
                  key={title}
                  style={{
                    display: "flex",
                    alignItems: "flex-start",
                    gap: 10,
                    padding: "7px 0",
                    borderBottom: "1px solid rgba(255,255,255,0.04)",
                  }}
                >
                  <span style={{ fontSize: 16, flexShrink: 0, marginTop: 1 }}>
                    {icon}
                  </span>
                  <div>
                    <span
                      style={{
                        color: "#e2e8f0",
                        fontWeight: 600,
                        fontSize: 12,
                      }}
                    >
                      {title} —{" "}
                    </span>
                    <span style={{ color: "#64748b", fontSize: 12 }}>
                      {desc}
                    </span>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      <style>{`
        @keyframes spin {
          to { transform: rotate(360deg); }
        }
      `}</style>
    </>
  );
}
