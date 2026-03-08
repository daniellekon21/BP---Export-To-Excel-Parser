import { useState, useCallback, useRef, useMemo } from "react";
import { formatTime, dateToStr, monthLabel, MONTH_NAMES } from "./helpers.js";
import { parseCuttingMessages, parseCuttingMessagesNew } from "./parsing/cuttingParser.js";
import { parseBalingMessages } from "./parsing/balingParser.js";
import { downloadWorkbook } from "./excel/downloadWorkbook.js";

function getQuarter(month) {
  return Math.ceil(month / 3);
}

function filterRecords(records, mode, filterYear, filterPeriod) {
  if (mode === "all") return records;
  return records.filter((r) => {
    if (!r.date) return false;
    const { year, month } = r.date;
    if (mode === "year") return filterYear ? year === parseInt(filterYear, 10) : true;
    if (mode === "month") {
      const yearMatch = filterYear ? year === parseInt(filterYear, 10) : true;
      const periodMatch = filterPeriod ? month === parseInt(filterPeriod, 10) : true;
      return yearMatch && periodMatch;
    }
    if (mode === "quarter") {
      const yearMatch = filterYear ? year === parseInt(filterYear, 10) : true;
      const periodMatch = filterPeriod ? getQuarter(month) === parseInt(filterPeriod, 10) : true;
      return yearMatch && periodMatch;
    }
    return true;
  });
}

function parseLogDate(dateStr) {
  if (!dateStr) return null;
  const m = String(dateStr).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  return { day: parseInt(m[1], 10), month: parseInt(m[2], 10), year: parseInt(m[3], 10) };
}

function filterValidationLog(logEntries, mode, filterYear, filterPeriod) {
  if (mode === "all") return logEntries;
  return logEntries.filter((entry) => {
    const d = parseLogDate(entry.date);
    if (!d) return false;
    const { year, month } = d;
    if (mode === "year") return filterYear ? year === parseInt(filterYear, 10) : true;
    if (mode === "month") {
      const yearMatch = filterYear ? year === parseInt(filterYear, 10) : true;
      const periodMatch = filterPeriod ? month === parseInt(filterPeriod, 10) : true;
      return yearMatch && periodMatch;
    }
    if (mode === "quarter") {
      const yearMatch = filterYear ? year === parseInt(filterYear, 10) : true;
      const periodMatch = filterPeriod ? getQuarter(month) === parseInt(filterPeriod, 10) : true;
      return yearMatch && periodMatch;
    }
    return true;
  });
}

function emptyBalingData() {
  return {
    standardRecords: [],
    failedRecords: [],
    scrapRecords: [],
    crcaRecords: [],
    summaryRecords: [],
    validationLog: [],
    allRecords: [],
  };
}

export default function App() {
  const [chatType, setChatType] = useState(null);
  const [cuttingMode, setCuttingMode] = useState("old");
  const [chatText, setChatText] = useState("");
  const [records, setRecords] = useState([]);
  const [summaryRecords, setSummaryRecords] = useState([]);
  const [validationLog, setValidationLog] = useState([]);
  const [balingData, setBalingData] = useState(emptyBalingData());
  const [parsed, setParsed] = useState(false);
  const [fileName, setFileName] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [filterMode, setFilterMode] = useState("all");
  const [filterYear, setFilterYear] = useState("");
  const [filterPeriod, setFilterPeriod] = useState("");
  const [placeholderPolicy, setPlaceholderPolicy] = useState("keep");
  const [exportFileName, setExportFileName] = useState("");
  const fileRef = useRef(null);

  const filterOptions = useMemo(() => {
    if (!records.length) return { months: [], quarters: [], years: [] };
    const monthSet = new Set();
    const quarterSet = new Set();
    const yearSet = new Set();
    for (const r of records) {
      if (!r.date) continue;
      const { year, month } = r.date;
      monthSet.add(`${year}-${String(month).padStart(2, "0")}`);
      quarterSet.add(`${year}-Q${getQuarter(month)}`);
      yearSet.add(String(year));
    }
    return {
      months: [...monthSet].sort(),
      quarters: [...quarterSet].sort(),
      years: [...yearSet].sort(),
    };
  }, [records]);

  const filteredRecords = useMemo(
    () => filterRecords(records, filterMode, filterYear, filterPeriod),
    [records, filterMode, filterYear, filterPeriod]
  );
  const filteredSummaryRecords = useMemo(
    () => filterRecords(summaryRecords, filterMode, filterYear, filterPeriod),
    [summaryRecords, filterMode, filterYear, filterPeriod]
  );
  const filteredValidationLog = useMemo(
    () => filterValidationLog(validationLog, filterMode, filterYear, filterPeriod),
    [validationLog, filterMode, filterYear, filterPeriod]
  );

  const filteredBalingData = useMemo(() => {
    if (chatType !== "baling") return balingData;
    return {
      ...balingData,
      standardRecords: filterRecords(balingData.standardRecords, filterMode, filterYear, filterPeriod),
      failedRecords: filterRecords(balingData.failedRecords, filterMode, filterYear, filterPeriod),
      scrapRecords: filterRecords(balingData.scrapRecords, filterMode, filterYear, filterPeriod),
      crcaRecords: filterRecords(balingData.crcaRecords, filterMode, filterYear, filterPeriod),
      summaryRecords: filterRecords(balingData.summaryRecords, filterMode, filterYear, filterPeriod),
      validationLog: filteredValidationLog,
      allRecords: filterRecords(balingData.allRecords, filterMode, filterYear, filterPeriod),
    };
  }, [chatType, balingData, filterMode, filterYear, filterPeriod, filteredValidationLog]);

  const visibleRecords = useMemo(() => {
    if (chatType === "baling") return filteredBalingData.allRecords;
    if (placeholderPolicy === "keep") return filteredRecords;
    return filteredRecords.filter((r) => !r._syntheticPlaceholder);
  }, [chatType, filteredBalingData, filteredRecords, placeholderPolicy]);

  const allMonthKeys = useMemo(() => {
    const keys = new Set();
    for (const r of records) {
      if (!r?.date) continue;
      keys.add(`${r.date.year}-${String(r.date.month).padStart(2, "0")}`);
    }
    for (const s of summaryRecords) {
      if (!s?.date) continue;
      keys.add(`${s.date.year}-${String(s.date.month).padStart(2, "0")}`);
    }
    return [...keys].sort();
  }, [records, summaryRecords]);

  const exportMonthKeys = useMemo(() => {
    const keys = new Set();
    for (const r of visibleRecords) {
      if (!r?.date) continue;
      keys.add(`${r.date.year}-${String(r.date.month).padStart(2, "0")}`);
    }
    for (const s of filteredSummaryRecords) {
      if (!s?.date) continue;
      keys.add(`${s.date.year}-${String(s.date.month).padStart(2, "0")}`);
    }
    return [...keys].sort();
  }, [visibleRecords, filteredSummaryRecords]);

  const excludedMonthKeys = useMemo(
    () => allMonthKeys.filter((k) => !exportMonthKeys.includes(k)),
    [allMonthKeys, exportMonthKeys]
  );

  const exportScopeWarning = useMemo(() => {
    if (filterMode === "all" || excludedMonthKeys.length === 0) return "";
    const preview = excludedMonthKeys.slice(0, 3).map(monthLabel).join(", ");
    const extra = excludedMonthKeys.length > 3 ? ` +${excludedMonthKeys.length - 3} more` : "";
    return `Export scope excludes ${excludedMonthKeys.length} month(s): ${preview}${extra}.`;
  }, [filterMode, excludedMonthKeys]);

  const handleFile = useCallback((file) => {
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => setChatText(e.target.result);
    reader.readAsText(file);
  }, []);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  }, [handleFile]);

  const handleParse = () => {
    if (!chatText.trim()) return;

    if (chatType === "baling") {
      const parsedBaling = parseBalingMessages(chatText);
      setBalingData(parsedBaling);
      setRecords(parsedBaling.allRecords);
      setSummaryRecords(parsedBaling.summaryRecords);
      setValidationLog(parsedBaling.validationLog);
    } else if (cuttingMode === "new") {
      const parsedCutting = parseCuttingMessagesNew(chatText);
      setBalingData(emptyBalingData());
      setRecords(parsedCutting.records);
      setSummaryRecords(parsedCutting.summaryRecords);
      setValidationLog(parsedCutting.validationLog);
    } else {
      const parsedCutting = parseCuttingMessages(chatText);
      setBalingData(emptyBalingData());
      setRecords(parsedCutting.records);
      setSummaryRecords(parsedCutting.summaryRecords);
      setValidationLog(parsedCutting.validationLog);
    }

    setParsed(true);
  };

  const handleDownload = async () => {
    const defaultFilename = chatType === "baling" ? "BPR_Baling_Data.xlsx" : "BPR_Cutting_Data.xlsx";
    const cleaned = String(exportFileName || "")
      .trim()
      .replace(/[\\/:*?"<>|]/g, "")
      .replace(/\s+/g, " ");
    const filename = cleaned ? (cleaned.toLowerCase().endsWith(".xlsx") ? cleaned : `${cleaned}.xlsx`) : defaultFilename;

    if (exportScopeWarning) {
      const ok = window.confirm(`${exportScopeWarning}\n\nContinue with scoped export?`);
      if (!ok) return;
    }

    try {
      await downloadWorkbook({
        chatType,
        filename,
        cutting: {
          records: visibleRecords,
          summaryRecords: filteredSummaryRecords,
          validationLog: filteredValidationLog,
        },
        baling: filteredBalingData,
      });
    } catch (error) {
      console.error("Excel export failed:", error);
      window.alert("Excel export failed. Please try again.");
    }
  };

  const handleReset = () => {
    setChatType(null);
    setCuttingMode("old");
    setChatText("");
    setRecords([]);
    setSummaryRecords([]);
    setValidationLog([]);
    setBalingData(emptyBalingData());
    setParsed(false);
    setFileName("");
    setFilterMode("all");
    setFilterYear("");
    setFilterPeriod("");
    setPlaceholderPolicy("keep");
    setExportFileName("");
  };

  const accent = "#1B2E5C";
  const accentLight = "#EDF1F7";
  const warmBg = "#F8FAFC";
  const darkText = "#0F172A";
  const mutedText = "#64748B";
  const borderColor = "#E2E8F0";

  return (
    <div style={{ minHeight: "100vh", background: warmBg, fontFamily: "'Inter', 'Segoe UI', 'Helvetica Neue', sans-serif", color: darkText, display: "flex", flexDirection: "column", alignItems: "center", padding: "32px 16px" }}>
      <div style={{ textAlign: "center", marginBottom: 40, maxWidth: 600 }}>
        <div style={{ display: "inline-flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
          <img src="/logo.png" alt="Blue Pyramid" style={{ height: 56, objectFit: "contain" }} />
        </div>
        <p style={{ fontSize: 15, color: mutedText, margin: 0, lineHeight: 1.5 }}>My most humble gift to my love 💙💙💜</p>
      </div>

      {!chatType && (
        <div style={{ display: "flex", gap: 16, flexWrap: "wrap", justifyContent: "center" }}>
          {[
            { id: "baling", icon: "📦", title: "Baling Production", desc: "Bale reports with weights, operators & tyre counts" },
            { id: "cutting", icon: "✂️", title: "Cutting Data", desc: "Hourly cutting machine counts per tyre type" },
          ].map((opt) => (
            <button
              key={opt.id}
              onClick={() => setChatType(opt.id)}
              style={{ width: 260, padding: "28px 24px", background: "white", border: `2px solid ${borderColor}`, borderRadius: 16, cursor: "pointer", textAlign: "left", transition: "all 0.2s", boxShadow: "0 1px 3px rgba(0,0,0,0.04)" }}
              onMouseEnter={(e) => {
                e.currentTarget.style.borderColor = accent;
                e.currentTarget.style.boxShadow = `0 0 0 3px ${accentLight}`;
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.borderColor = borderColor;
                e.currentTarget.style.boxShadow = "0 1px 3px rgba(0,0,0,0.04)";
              }}
            >
              <span style={{ fontSize: 32 }}>{opt.icon}</span>
              <h3 style={{ fontSize: 17, fontWeight: 600, margin: "12px 0 6px" }}>{opt.title}</h3>
              <p style={{ fontSize: 13, color: mutedText, margin: 0, lineHeight: 1.4 }}>{opt.desc}</p>
            </button>
          ))}
        </div>
      )}

      {chatType && !parsed && (
        <div style={{ width: "100%", maxWidth: 560 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 20 }}>
            <button onClick={handleReset} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 14, color: mutedText, padding: "4px 0", display: "flex", alignItems: "center", gap: 4 }}>
              ← Back
            </button>
            <span style={{ color: "#D1D5DB" }}>|</span>
            <span style={{ fontSize: 14, fontWeight: 600 }}>{chatType === "baling" ? "📦 Baling Production" : "✂️ Cutting Data"}</span>
          </div>

          {chatType === "cutting" && (
            <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16, fontSize: 13, color: mutedText }}>
              <span style={{ fontWeight: 500 }}>Format:</span>
              {[
                { value: "old", label: "Old WhatsApp format" },
                { value: "new", label: "New WhatsApp format" },
              ].map((opt) => (
                <label key={opt.value} style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
                  <input type="radio" name="cuttingMode" value={opt.value} checked={cuttingMode === opt.value} onChange={() => setCuttingMode(opt.value)} style={{ accentColor: accent }} />
                  {opt.label}
                </label>
              ))}
            </div>
          )}

          <div
            onDragOver={(e) => {
              e.preventDefault();
              setDragOver(true);
            }}
            onDragLeave={() => setDragOver(false)}
            onDrop={handleDrop}
            onClick={() => fileRef.current?.click()}
            style={{ border: `2px dashed ${dragOver ? accent : "#D1D5DB"}`, borderRadius: 16, padding: "40px 24px", textAlign: "center", cursor: "pointer", background: dragOver ? accentLight : "white", transition: "all 0.2s", marginBottom: 16 }}
          >
            <input ref={fileRef} type="file" accept=".txt,.zip" style={{ display: "none" }} onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])} />
            <div style={{ fontSize: 36, marginBottom: 8 }}>📂</div>
            <p style={{ fontSize: 15, fontWeight: 500, margin: "0 0 4px" }}>{fileName || "Drop your WhatsApp .txt export here"}</p>
            <p style={{ fontSize: 13, color: mutedText, margin: 0 }}>or click to browse</p>
          </div>

          <div style={{ marginBottom: 16 }}>
            <p style={{ fontSize: 12, color: mutedText, textAlign: "center", margin: "0 0 8px", textTransform: "uppercase", letterSpacing: 1 }}>or paste chat text</p>
            <textarea
              value={chatText}
              onChange={(e) => setChatText(e.target.value)}
              placeholder="Paste the exported WhatsApp chat text here..."
              style={{ width: "100%", minHeight: 140, border: "1px solid #E5E7EB", borderRadius: 12, padding: 16, fontSize: 13, fontFamily: "monospace", resize: "vertical", background: "white", boxSizing: "border-box" }}
            />
          </div>

          <button onClick={handleParse} disabled={!chatText.trim()} style={{ width: "100%", padding: "14px 0", background: chatText.trim() ? accent : "#D1D5DB", color: "white", border: "none", borderRadius: 12, fontSize: 15, fontWeight: 600, cursor: chatText.trim() ? "pointer" : "default", transition: "background 0.2s" }}>
            Parse Chat →
          </button>
        </div>
      )}

      {parsed && (
        <div style={{ width: "100%", maxWidth: 900 }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 12 }}>
            <div>
              <h2 style={{ fontSize: 20, fontWeight: 700, margin: 0 }}>{records.length} records parsed</h2>
              <p style={{ fontSize: 13, color: mutedText, margin: "4px 0 0" }}>{chatType === "baling" ? "Baling records across all subtypes" : "Cutting data entries"}</p>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={handleReset} style={{ padding: "10px 20px", background: "white", border: "1px solid #D1D5DB", borderRadius: 10, fontSize: 14, cursor: "pointer" }}>Start Over</button>
              <button onClick={handleDownload} disabled={visibleRecords.length === 0} style={{ padding: "10px 24px", background: visibleRecords.length > 0 ? accent : "#D1D5DB", color: "white", border: "none", borderRadius: 10, fontSize: 14, fontWeight: 600, cursor: visibleRecords.length > 0 ? "pointer" : "default" }}>
                ⬇ Download Excel
              </button>
            </div>
          </div>

          <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", marginBottom: 16, padding: "12px 16px", background: "white", borderRadius: 12, border: `1px solid ${borderColor}` }}>
            <span style={{ fontSize: 13, fontWeight: 600, color: mutedText }}>📅 Export:</span>
            {[{ value: "all", label: "All" }, { value: "month", label: "Month" }, { value: "quarter", label: "Quarter" }, { value: "year", label: "Year" }].map((opt) => (
              <button
                key={opt.value}
                onClick={() => {
                  setFilterMode(opt.value);
                  setFilterYear("");
                  setFilterPeriod("");
                }}
                style={{ padding: "6px 14px", borderRadius: 8, border: filterMode === opt.value ? `2px solid ${accent}` : `1px solid ${borderColor}`, background: filterMode === opt.value ? accentLight : "white", color: filterMode === opt.value ? accent : darkText, fontWeight: filterMode === opt.value ? 600 : 400, fontSize: 13, cursor: "pointer", transition: "all 0.15s", opacity: filterMode !== "all" && filterMode !== opt.value ? 0.4 : 1 }}
              >
                {opt.label}
              </button>
            ))}

            {filterMode !== "all" && (
              <>
                <select value={filterYear} onChange={(e) => setFilterYear(e.target.value)} style={{ padding: "6px 12px", borderRadius: 8, border: `1px solid ${borderColor}`, fontSize: 13, background: "white", cursor: "pointer" }}>
                  <option value="">Year…</option>
                  {filterOptions.years.map((y) => (
                    <option key={y} value={y}>{y}</option>
                  ))}
                </select>

                {filterMode === "month" && (
                  <select value={filterPeriod} onChange={(e) => setFilterPeriod(e.target.value)} style={{ padding: "6px 12px", borderRadius: 8, border: `1px solid ${borderColor}`, fontSize: 13, background: "white", cursor: "pointer" }}>
                    <option value="">Month…</option>
                    {MONTH_NAMES.map((name, i) => (
                      <option key={i + 1} value={i + 1}>{name}</option>
                    ))}
                  </select>
                )}

                {filterMode === "quarter" && (
                  <select value={filterPeriod} onChange={(e) => setFilterPeriod(e.target.value)} style={{ padding: "6px 12px", borderRadius: 8, border: `1px solid ${borderColor}`, fontSize: 13, background: "white", cursor: "pointer" }}>
                    <option value="">Quarter…</option>
                    <option value="1">Q1 (Jan–Mar)</option>
                    <option value="2">Q2 (Apr–Jun)</option>
                    <option value="3">Q3 (Jul–Sep)</option>
                    <option value="4">Q4 (Oct–Dec)</option>
                  </select>
                )}
              </>
            )}

            {filterMode !== "all" && (filterYear || filterPeriod) && (
              <span style={{ fontSize: 12, color: mutedText, marginLeft: 4 }}>{visibleRecords.length} of {records.length} records</span>
            )}

            {chatType === "cutting" && (
              <select value={placeholderPolicy} onChange={(e) => setPlaceholderPolicy(e.target.value)} style={{ marginLeft: "auto", padding: "6px 10px", borderRadius: 8, border: `1px solid ${borderColor}`, fontSize: 12, background: "white", cursor: "pointer" }}>
                <option value="keep">Keep placeholders</option>
                <option value="strict">Strict source only</option>
              </select>
            )}

            <input
              type="text"
              value={exportFileName}
              onChange={(e) => setExportFileName(e.target.value)}
              placeholder={chatType === "baling" ? "BPR_Baling_Data.xlsx" : "BPR_Cutting_Data.xlsx"}
              style={{ marginLeft: chatType === "cutting" ? 0 : "auto", padding: "6px 10px", borderRadius: 8, border: `1px solid ${borderColor}`, fontSize: 12, background: "white", minWidth: 220 }}
              aria-label="Export filename"
            />
          </div>

          {exportScopeWarning && <p style={{ margin: "-8px 0 12px", fontSize: 12, color: "#B45309" }}>⚠ {exportScopeWarning}</p>}

          <div style={{ background: "white", borderRadius: 16, border: "1px solid #E5E7EB", overflow: "auto", boxShadow: "0 1px 3px rgba(0,0,0,0.04)" }}>
            {chatType === "baling" ? (
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ background: "#F9FAFB" }}>
                    {["Date", "Type", "Machine", "Bale", "Status", "Start", "Finish", "Passenger", "4x4", "LC", "MC", "SR", "Agri", "Weight", "Notes"].map((h) => (
                      <th key={h} style={{ padding: "10px 8px", textAlign: "left", borderBottom: "1px solid #E5E7EB", fontWeight: 600, whiteSpace: "nowrap", color: mutedText, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5 }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {visibleRecords.slice(0, 100).map((r, i) => (
                    <tr key={i} style={{ borderBottom: "1px solid #F3F4F6" }}>
                      <td style={{ padding: "8px" }}>{r.date ? dateToStr(r.date) : "—"}</td>
                      <td style={{ padding: "8px" }}>{r.recordType || "—"}</td>
                      <td style={{ padding: "8px" }}>{r.machine || "—"}</td>
                      <td style={{ padding: "8px" }}>{r.baleNumber || r.baleTestCode || "—"}</td>
                      <td style={{ padding: "8px" }}>{r.status || "—"}</td>
                      <td style={{ padding: "8px" }}>{r.startTime ? formatTime(r.startTime) : "—"}</td>
                      <td style={{ padding: "8px" }}>{r.finishTime ? formatTime(r.finishTime) : "—"}</td>
                      <td style={{ padding: "8px" }}>{r.passengerQty ?? "—"}</td>
                      <td style={{ padding: "8px" }}>{r.fourx4Qty ?? "—"}</td>
                      <td style={{ padding: "8px" }}>{r.lcQty ?? "—"}</td>
                      <td style={{ padding: "8px" }}>{r.motorcycleQty ?? "—"}</td>
                      <td style={{ padding: "8px" }}>{r.srQty ?? "—"}</td>
                      <td style={{ padding: "8px" }}>{r.agriQty ?? "—"}</td>
                      <td style={{ padding: "8px", fontWeight: 600 }}>{r.weightKg ?? "—"}</td>
                      <td style={{ padding: "8px" }}>{r.notesFlags || "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : (
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ background: "#F9FAFB" }}>
                    {["Date", "Machine", "Time", "LC", "HC", "Agri", "Tread LC", "Tread HC", "Tread Agri"].map((h) => (
                      <th key={h} style={{ padding: "10px 8px", textAlign: "left", borderBottom: "1px solid #E5E7EB", fontWeight: 600, whiteSpace: "nowrap", color: mutedText, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5 }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {visibleRecords.slice(0, 100).map((r, i) => (
                    <tr key={i} style={{ borderBottom: "1px solid #F3F4F6" }}>
                      <td style={{ padding: "8px" }}>{dateToStr(r.date)}</td>
                      <td style={{ padding: "8px" }}>{r.cmNumber}</td>
                      <td style={{ padding: "8px" }}>{r.startTime && r.finishTime ? `${formatTime(r.startTime)}-${formatTime(r.finishTime)}` : "—"}</td>
                      <td style={{ padding: "8px", fontWeight: 600 }}>{r.light_commercial ?? "—"}</td>
                      <td style={{ padding: "8px", fontWeight: 600 }}>{r.heavy_commercial_t ?? "—"}</td>
                      <td style={{ padding: "8px", fontWeight: 600 }}>{r.agricultural_t ?? "—"}</td>
                      <td style={{ padding: "8px" }}>{r.tread_lc ?? "—"}</td>
                      <td style={{ padding: "8px" }}>{r.tread_hc ?? "—"}</td>
                      <td style={{ padding: "8px" }}>{r.tread_agri ?? "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
            {visibleRecords.length > 100 && <p style={{ textAlign: "center", padding: 12, fontSize: 12, color: mutedText }}>Showing first 100 of {visibleRecords.length} records. All {visibleRecords.length} will be included in the download.</p>}
          </div>

          <p style={{ fontSize: 12, color: mutedText, textAlign: "center", marginTop: 16, lineHeight: 1.5 }}>The Excel file output is separated by workflow: cutting sheets or baling-specific tables.</p>
        </div>
      )}
    </div>
  );
}
