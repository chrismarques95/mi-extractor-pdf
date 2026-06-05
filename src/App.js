import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

const ACCENT = "#FF3C1D";
const DARK = "#053220";
const CREAM = "#F9F7F0";
const TEAL = "#6C989A";
const GREY = "#A5A5A5";

// ─── tiny helpers ────────────────────────────────────────────────────────────
function fileToBase64(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(",")[1]);
    r.onerror = () => rej(new Error("Error leyendo archivo"));
    r.readAsDataURL(file);
  });
}

function buildPrompt() {
  return `Eres un extractor de datos estructurados a partir de documentos PDF.

Tu tarea es leer TODO el contenido del PDF adjunto e identificar registros de personas.

REGLAS ESTRICTAS:
1. Cada vez que aparezca un nombre y apellido diferente → es una fila diferente.
2. Cada campo asociado a ese nombre/apellido en cualquier hoja o sección del documento → es una columna.
3. Si un mismo campo aparece varias veces para la misma persona con valores diferentes, usa el formato: "NombreCampo_1", "NombreCampo_2", etc.
4. Incluye TODOS los campos que encuentres, sin omitir ninguno.
5. El campo "nombre_completo" siempre debe estar presente y ser la clave de identificación.

Responde ÚNICAMENTE con un objeto JSON válido, sin markdown, sin backticks, sin explicaciones. 
El formato exacto debe ser:

{
  "campos": ["nombre_completo", "campo2", "campo3", ...],
  "registros": [
    { "nombre_completo": "Juan Pérez", "campo2": "valor", "campo3": "valor", ... },
    { "nombre_completo": "Ana García", "campo2": "valor", "campo3": "valor", ... }
  ]
}

Si no encuentras personas/registros, devuelve: { "campos": [], "registros": [] }`;
}

// ─── exportar a xlsx ──────────────────────────────────────────────────────────
function exportToXlsx(allFields, allRecords, fileName) {
  const headers = allFields;
  const rows = allRecords.map((rec) => headers.map((h) => rec[h] ?? ""));
  const wsData = [headers, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Ancho de columnas
  ws["!cols"] = headers.map((h) => ({ wch: Math.max(h.length + 4, 18) }));

  // Estilo encabezados (openpyxl no aplica en SheetJS Community, pero configuramos anyway)
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Datos extraídos");
  XLSX.writeFile(wb, fileName.replace(/\.pdf$/i, "") + "_datos.xlsx");
}

// ─── Chip de estado ───────────────────────────────────────────────────────────
function StatusChip({ status }) {
  const map = {
    idle: { color: GREY, label: "En cola" },
    reading: { color: TEAL, label: "Leyendo…" },
    done: { color: "#2D6A4F", label: "Listo" },
    error: { color: ACCENT, label: "Error" },
  };
  const { color, label } = map[status] || map.idle;
  return (
    <span
      style={{
        fontSize: 11,
        fontWeight: 700,
        letterSpacing: ".06em",
        color,
        background: color + "18",
        border: `1px solid ${color}40`,
        borderRadius: 4,
        padding: "2px 8px",
        textTransform: "uppercase",
      }}
    >
      {label}
    </span>
  );
}

// ─── Barra de progreso animada ────────────────────────────────────────────────
function ProgressBar({ pct }) {
  return (
    <div
      style={{
        height: 3,
        background: GREY + "30",
        borderRadius: 99,
        overflow: "hidden",
        marginTop: 6,
      }}
    >
      <div
        style={{
          height: "100%",
          width: pct + "%",
          background: `linear-gradient(90deg, ${TEAL}, ${ACCENT})`,
          borderRadius: 99,
          transition: "width .4s ease",
        }}
      />
    </div>
  );
}

// ─── Vista previa de tabla ────────────────────────────────────────────────────
function PreviewTable({ fields, records }) {
  if (!records.length) return null;
  return (
    <div style={{ overflowX: "auto", marginTop: 24 }}>
      <p
        style={{
          fontSize: 11,
          color: GREY,
          letterSpacing: ".08em",
          textTransform: "uppercase",
          marginBottom: 8,
        }}
      >
        Vista previa — {records.length} registro{records.length !== 1 ? "s" : ""} ·{" "}
        {fields.length} campo{fields.length !== 1 ? "s" : ""}
      </p>
      <table
        style={{
          borderCollapse: "collapse",
          fontSize: 12,
          fontFamily: "'DM Mono', monospace",
          width: "100%",
        }}
      >
        <thead>
          <tr>
            {fields.slice(0, 8).map((f) => (
              <th
                key={f}
                style={{
                  background: DARK,
                  color: CREAM,
                  padding: "6px 12px",
                  textAlign: "left",
                  fontWeight: 600,
                  whiteSpace: "nowrap",
                  fontSize: 11,
                  letterSpacing: ".05em",
                }}
              >
                {f}
              </th>
            ))}
            {fields.length > 8 && (
              <th
                style={{
                  background: DARK,
                  color: GREY,
                  padding: "6px 12px",
                  fontStyle: "italic",
                  fontSize: 11,
                }}
              >
                +{fields.length - 8} más…
              </th>
            )}
          </tr>
        </thead>
        <tbody>
          {records.slice(0, 6).map((rec, i) => (
            <tr
              key={i}
              style={{ background: i % 2 === 0 ? CREAM : "#E8E6DF" }}
            >
              {fields.slice(0, 8).map((f) => (
                <td
                  key={f}
                  style={{
                    padding: "5px 12px",
                    color: DARK,
                    whiteSpace: "nowrap",
                    maxWidth: 180,
                    overflow: "hidden",
                    textOverflow: "ellipsis",
                  }}
                >
                  {rec[f] ?? ""}
                </td>
              ))}
              {fields.length > 8 && (
                <td style={{ padding: "5px 12px", color: GREY }}>…</td>
              )}
            </tr>
          ))}
        </tbody>
      </table>
      {records.length > 6 && (
        <p style={{ fontSize: 11, color: GREY, marginTop: 6 }}>
          … y {records.length - 6} registro{records.length - 6 !== 1 ? "s" : ""} más en el Excel
        </p>
      )}
    </div>
  );
}

// ─── MAIN ─────────────────────────────────────────────────────────────────────
export default function App() {
  const [files, setFiles] = useState([]); // [{ file, status, result, error }]
  const [processing, setProcessing] = useState(false);
  const [allFields, setAllFields] = useState([]);
  const [allRecords, setAllRecords] = useState([]);
  const [dragOver, setDragOver] = useState(false);
  const inputRef = useRef();

  const addFiles = useCallback((newFiles) => {
    const items = Array.from(newFiles)
      .filter((f) => f.type === "application/pdf")
      .map((file) => ({ file, status: "idle", result: null, error: null }));
    setFiles((prev) => [...prev, ...items]);
    setAllFields([]);
    setAllRecords([]);
  }, []);

  const onDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    addFiles(e.dataTransfer.files);
  };

  const processFiles = async () => {
    if (!files.length || processing) return;
    setProcessing(true);
    setAllFields([]);
    setAllRecords([]);

    const updatedFiles = files.map((f) => ({ ...f, status: "reading", error: null }));
    setFiles(updatedFiles);

    const mergedFieldsSet = new Set();
    const mergedRecords = [];

    for (let i = 0; i < files.length; i++) {
      const item = files[i];
      try {
        const b64 = await fileToBase64(item.file);

        const response = await fetch("https://api.anthropic.com/v1/messages", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            model: "claude-sonnet-4-20250514",
            max_tokens: 4096,
            messages: [
              {
                role: "user",
                content: [
                  {
                    type: "document",
                    source: {
                      type: "base64",
                      media_type: "application/pdf",
                      data: b64,
                    },
                  },
                  { type: "text", text: buildPrompt() },
                ],
              },
            ],
          }),
        });

        if (!response.ok) {
          const err = await response.json().catch(() => ({}));
          throw new Error(err?.error?.message || `HTTP ${response.status}`);
        }

        const data = await response.json();
        const raw = data.content?.find((c) => c.type === "text")?.text || "{}";
        const clean = raw.replace(/```json|```/g, "").trim();
        const parsed = JSON.parse(clean);

        const { campos = [], registros = [] } = parsed;
        campos.forEach((c) => mergedFieldsSet.add(c));
        mergedRecords.push(...registros);

        setFiles((prev) =>
          prev.map((f, idx) =>
            idx === i ? { ...f, status: "done", result: { campos, registros } } : f
          )
        );
      } catch (err) {
        setFiles((prev) =>
          prev.map((f, idx) =>
            idx === i ? { ...f, status: "error", error: err.message } : f
          )
        );
      }

      // Actualizar vista previa progresivamente
      setAllFields(Array.from(mergedFieldsSet));
      setAllRecords([...mergedRecords]);
    }

    setProcessing(false);
  };

  const downloadExcel = () => {
    if (!allRecords.length) return;
    const name =
      files.length === 1
        ? files[0].file.name
        : `extraccion_${files.length}_pdfs`;
    exportToXlsx(allFields, allRecords, name);
  };

  const clearAll = () => {
    setFiles([]);
    setAllFields([]);
    setAllRecords([]);
  };

  const doneCount = files.filter((f) => f.status === "done").length;
  const errorCount = files.filter((f) => f.status === "error").length;
  const progress = files.length ? Math.round((doneCount + errorCount) / files.length * 100) : 0;

  return (
    <div
      style={{
        minHeight: "100vh",
        background: CREAM,
        fontFamily: "'DM Sans', 'Helvetica Neue', sans-serif",
        color: DARK,
        padding: "40px 24px",
        boxSizing: "border-box",
      }}
    >
      {/* Google Fonts */}
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;600;700&family=DM+Mono:wght@400;600&display=swap');
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-thumb { background: ${GREY}60; border-radius: 99px; }
      `}</style>

      <div style={{ maxWidth: 820, margin: "0 auto" }}>
        {/* Header */}
        <div style={{ marginBottom: 40 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 8 }}>
            <div
              style={{
                width: 8,
                height: 32,
                background: ACCENT,
                borderRadius: 2,
              }}
            />
            <h1
              style={{
                margin: 0,
                fontSize: 28,
                fontWeight: 700,
                letterSpacing: "-.02em",
                color: DARK,
              }}
            >
              PDF → Excel
            </h1>
          </div>
          <p style={{ margin: 0, color: GREY, fontSize: 14, paddingLeft: 20 }}>
            Carga uno o más PDFs · Claude extrae los datos · Descarga el Excel consolidado
          </p>
        </div>

        {/* Drop zone */}
        <div
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={onDrop}
          onClick={() => inputRef.current?.click()}
          style={{
            border: `2px dashed ${dragOver ? ACCENT : TEAL + "80"}`,
            borderRadius: 12,
            background: dragOver ? ACCENT + "08" : TEAL + "08",
            padding: "36px 24px",
            textAlign: "center",
            cursor: "pointer",
            transition: "all .2s ease",
            marginBottom: 20,
          }}
        >
          <div style={{ fontSize: 36, marginBottom: 8 }}>📄</div>
          <p style={{ margin: 0, fontWeight: 600, color: DARK }}>
            Arrastra tus PDFs aquí
          </p>
          <p style={{ margin: "4px 0 0", fontSize: 13, color: GREY }}>
            o haz clic para seleccionar archivos
          </p>
          <input
            ref={inputRef}
            type="file"
            accept="application/pdf"
            multiple
            style={{ display: "none" }}
            onChange={(e) => addFiles(e.target.files)}
          />
        </div>

        {/* Lista de archivos */}
        {files.length > 0 && (
          <div
            style={{
              background: "#fff",
              border: `1px solid ${DARK}18`,
              borderRadius: 10,
              overflow: "hidden",
              marginBottom: 20,
            }}
          >
            {files.map((item, i) => (
              <div
                key={i}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 12,
                  padding: "12px 16px",
                  borderBottom: i < files.length - 1 ? `1px solid ${DARK}10` : "none",
                }}
              >
                <span style={{ fontSize: 18 }}>
                  {item.status === "done" ? "✅" : item.status === "error" ? "❌" : "📋"}
                </span>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <p
                    style={{
                      margin: 0,
                      fontSize: 13,
                      fontWeight: 600,
                      color: DARK,
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {item.file.name}
                  </p>
                  {item.error && (
                    <p style={{ margin: "2px 0 0", fontSize: 11, color: ACCENT }}>
                      {item.error}
                    </p>
                  )}
                  {item.result && (
                    <p style={{ margin: "2px 0 0", fontSize: 11, color: TEAL }}>
                      {item.result.registros.length} registro{item.result.registros.length !== 1 ? "s" : ""} ·{" "}
                      {item.result.campos.length} campos
                    </p>
                  )}
                  {item.status === "reading" && <ProgressBar pct={50} />}
                </div>
                <StatusChip status={item.status} />
              </div>
            ))}
          </div>
        )}

        {/* Barra de progreso global */}
        {processing && files.length > 1 && (
          <div style={{ marginBottom: 16 }}>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                fontSize: 12,
                color: GREY,
                marginBottom: 4,
              }}
            >
              <span>Procesando archivos…</span>
              <span>{doneCount + errorCount} / {files.length}</span>
            </div>
            <ProgressBar pct={progress} />
          </div>
        )}

        {/* Botones */}
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          {files.length > 0 && !processing && allRecords.length === 0 && (
            <button
              onClick={processFiles}
              style={{
                background: DARK,
                color: CREAM,
                border: "none",
                borderRadius: 8,
                padding: "12px 28px",
                fontFamily: "inherit",
                fontSize: 14,
                fontWeight: 700,
                cursor: "pointer",
                letterSpacing: ".03em",
                transition: "opacity .2s",
              }}
              onMouseEnter={(e) => (e.target.style.opacity = ".85")}
              onMouseLeave={(e) => (e.target.style.opacity = "1")}
            >
              ▶ Extraer datos
            </button>
          )}

          {allRecords.length > 0 && (
            <button
              onClick={downloadExcel}
              style={{
                background: ACCENT,
                color: "#fff",
                border: "none",
                borderRadius: 8,
                padding: "12px 28px",
                fontFamily: "inherit",
                fontSize: 14,
                fontWeight: 700,
                cursor: "pointer",
                letterSpacing: ".03em",
                transition: "opacity .2s",
              }}
              onMouseEnter={(e) => (e.target.style.opacity = ".85")}
              onMouseLeave={(e) => (e.target.style.opacity = "1")}
            >
              ⬇ Descargar Excel ({allRecords.length} filas)
            </button>
          )}

          {files.length > 0 && !processing && (
            <button
              onClick={clearAll}
              style={{
                background: "transparent",
                color: GREY,
                border: `1px solid ${GREY}60`,
                borderRadius: 8,
                padding: "12px 20px",
                fontFamily: "inherit",
                fontSize: 14,
                cursor: "pointer",
                transition: "all .2s",
              }}
              onMouseEnter={(e) => {
                e.target.style.borderColor = ACCENT;
                e.target.style.color = ACCENT;
              }}
              onMouseLeave={(e) => {
                e.target.style.borderColor = GREY + "60";
                e.target.style.color = GREY;
              }}
            >
              ✕ Limpiar todo
            </button>
          )}
        </div>

        {/* Vista previa */}
        {allRecords.length > 0 && (
          <PreviewTable fields={allFields} records={allRecords} />
        )}

        {/* Footer */}
        <p
          style={{
            marginTop: 48,
            fontSize: 11,
            color: GREY + "99",
            textAlign: "center",
            letterSpacing: ".04em",
          }}
        >
          Los PDFs se procesan vía Claude API · Los datos no se almacenan
        </p>
      </div>
    </div>
  );
}
