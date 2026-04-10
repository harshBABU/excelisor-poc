import React, { useState } from "react";

const apiBase = "http://localhost:8000";

async function getExcelData() {
  return await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["values", "address", "rowCount", "columnCount"]);
    await context.sync();
    return {
      values: range.values || [],
      address: range.address || "",
      metadata: {
        originalRows: range.rowCount || 0,
        originalColumns: range.columnCount || 0,
      },
    };
  });
}

// ─── Inline styles as a JS object tree ───────────────────────────────────────
const brand = {
  navy: "#08121C",
  navyMid: "#142536",
  navyLight: "#1E3450",
  teal: "#2BC99C",
  tealDark: "#1A9E7A",
  tealFaint: "#E5F6F1",
  tealBorder: "rgba(43,201,156,0.25)",
  offWhite: "#F6F8FA",
  white: "#FFFFFF",
  border: "#D3DEE8",
  borderFocus: "#2BC99C",
  textPrimary: "#0D1E2E",
  textMuted: "#6B8399",
  textLight: "#9FB3C8",
  errorBg: "#FFF2F2",
  errorText: "#C0392B",
  successBg: "#EDFAF4",
  successText: "#1A7A56",
};

const MODELS = [
  { value: "gpt-5-nano", label: "GPT-5 Nano", hint: "Fast · Low cost" },
  { value: "gpt-4o-mini", label: "GPT-4o Mini", hint: "Balanced" },
  { value: "gpt-4o", label: "GPT-4o", hint: "Most capable" },
  { value: "gpt-5-mini", label: "GPT-5 Mini", hint: "Fast + smart" },
  { value: "gpt-5.4-nano", label: "GPT-5.4 Nano", hint: "Experimental" },
];

const globalCss = `
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,400;0,9..40,500;0,9..40,600;1,9..40,400&family=Syne:wght@600;700&display=swap');

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  body { background: ${brand.offWhite}; font-family: 'DM Sans', sans-serif; }

  @keyframes spin {
    to { transform: rotate(360deg); }
  }
  @keyframes fadeUp {
    from { opacity: 0; transform: translateY(6px); }
    to   { opacity: 1; transform: translateY(0); }
  }
  @keyframes pulse {
    0%, 100% { opacity: 1; }
    50%       { opacity: 0.5; }
  }

  .xc-input:focus {
    border-color: ${brand.teal} !important;
    box-shadow: 0 0 0 3px rgba(43,201,156,0.15) !important;
    outline: none;
  }
  .xc-input::placeholder { color: ${brand.textLight}; }

  .xc-btn-primary:hover:not(:disabled)  { background: ${brand.tealDark} !important; box-shadow: 0 6px 18px rgba(43,201,156,0.35) !important; transform: translateY(-1px); }
  .xc-btn-secondary:hover:not(:disabled){ background: ${brand.navyLight} !important; }
  .xc-btn-ghost:hover:not(:disabled)    { border-color: ${brand.teal} !important; color: ${brand.teal} !important; }
  .xc-btn-danger:hover:not(:disabled)   { background: #fdecea !important; color: ${brand.errorText} !important; border-color: #f5c6c2 !important; }

  .xc-btn-primary:active:not(:disabled),
  .xc-btn-secondary:active:not(:disabled),
  .xc-btn-ghost:active:not(:disabled)   { transform: translateY(0px) !important; }
  .xc-btn:disabled { opacity: 0.45; cursor: not-allowed; }

  .settings-panel-enter { animation: fadeUp 0.2s ease both; }

  .action-btn:hover:not(:disabled) .action-btn-icon { transform: scale(1.15); }

  select.xc-input option { background: ${brand.white}; color: ${brand.textPrimary}; }

  .xc-output-text a { color: ${brand.teal}; }

  /* Scrollbar */
  .xc-output-area::-webkit-scrollbar { width: 4px; }
  .xc-output-area::-webkit-scrollbar-track { background: transparent; }
  .xc-output-area::-webkit-scrollbar-thumb { background: ${brand.border}; border-radius: 99px; }
`;

// ─── Small reusable UI atoms ──────────────────────────────────────────────────

const Spinner = () => (
  <div style={{
    width: 15, height: 15, borderRadius: "50%",
    border: `2px solid ${brand.tealFaint}`,
    borderTopColor: brand.teal,
    animation: "spin 0.7s linear infinite",
    flexShrink: 0,
  }} />
);

const Label = ({ children, style = {} }) => (
  <div style={{
    fontSize: 10, fontWeight: 600, letterSpacing: "0.1em",
    textTransform: "uppercase", color: brand.textMuted,
    marginBottom: 6, ...style,
  }}>
    {children}
  </div>
);

const StatusBadge = ({ loading, output }) => {
  if (!output && !loading) return null;
  const isError = output?.startsWith("❌");
  const isSuccess = output?.startsWith("✅");
  const bg = isError ? brand.errorBg : isSuccess ? brand.successBg : brand.tealFaint;
  const color = isError ? brand.errorText : isSuccess ? brand.successText : brand.tealDark;

  return (
    <div style={{
      display: "flex", alignItems: "flex-start", gap: 8,
      padding: "10px 12px", borderRadius: 8,
      background: bg, color,
      fontSize: 12, lineHeight: 1.5,
      animation: "fadeUp 0.2s ease",
    }}>
      {loading && <Spinner />}
      <span style={{ flex: 1 }}>{output}</span>
    </div>
  );
};

// ─── Main component ───────────────────────────────────────────────────────────
export default function App() {
  const [input, setInput]               = useState("");
  const [messages, setMessages]         = useState([]);      // chat history
  const [loading, setLoading]           = useState(false);
  const [clarificationRound, setClarificationRound] = useState(0);
  const [lastContext, setLastContext]   = useState(null);    // accumulated HITL context
  const [apiKey, setApiKey]             = useState(localStorage.getItem("openai_api_key") || "");
  const [selectedModel, setSelectedModel] = useState(localStorage.getItem("openai_model") || "gpt-5-nano");
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [keyError, setKeyError]         = useState("");

  // Refs — must be declared before any early returns (Rules of Hooks)
  const chatFeed = React.useRef(null);
  React.useEffect(() => {
    if (chatFeed.current) chatFeed.current.scrollTop = chatFeed.current.scrollHeight;
  }, [messages, loading]);

  const modeIcon = { ask_ai: "🔍", ai_chart: "📊", ai_action_formula: "🧮", python_analysis: "🐍", audit: "✅" };

  // Internal helpers
  const addMessage = (role, content, extra = {}) =>
    setMessages(prev => [...prev, { role, content, ts: Date.now(), ...extra }]);

  // ── API helpers ─────────────────────────────────────────────────────────────
  const callJson = async (path, body) => {
    const reqHeaders = { "Content-Type": "application/json" };
    if (apiKey) reqHeaders["Authorization"] = `Bearer ${apiKey}`;
    if (selectedModel) reqHeaders["X-OpenAI-Model"] = selectedModel;
    const resp = await fetch(`${apiBase}${path}`, {
      method: "POST", headers: reqHeaders, body: JSON.stringify(body),
    });
    if (!resp.ok) { const t = await resp.text(); throw new Error(`HTTP ${resp.status}: ${t}`); }
    return await resp.json();
  };

  // ── Smart Route (primary chat handler) ───────────────────────────────────
  const sendMessage = async (userText, contextOverride = null, roundOverride = null) => {
    const text = (userText || input).trim();
    if (!text || loading) return;

    // Add the user bubble to the chat
    addMessage("user", text);
    setInput("");
    setLoading(true);

    try {
      const { values, address } = await getExcelData();
      const round = roundOverride !== null ? roundOverride : clarificationRound;
      const context = contextOverride !== null ? contextOverride : lastContext;

      const res = await callJson("/smart_route", {
        table: values,
        address,
        question: text,
        context,
        clarification_round: round,
      });

      // ── HITL clarification needed ──────────────────────────────────────────
      if (res.needs_clarification) {
        const nextRound = round + 1;
        setClarificationRound(nextRound);
        setLastContext(text); // carry question forward as context
        addMessage("assistant", res.clarification_question || "Could you clarify what you're looking for?", {
          type: "clarification",
          options: res.clarification_options || [],
          round: nextRound,
          originalQuestion: text,
        });
        return;
      }

      // Successful routing — reset HITL state
      setClarificationRound(0);
      setLastContext(null);

      // Only show rationale prefix if confidence is borderline AND it's a meaningful LLM explanation
      const showRationale = res.confidence >= 0.55 && res.confidence < 0.80
        && res.rationale
        && !res.rationale.startsWith("Auto-detected");
      const rationalePrefix = showRationale ? `💡 *${res.rationale}*\n\n` : "";

      // ── Chart result ────────────────────────────────────────────────────────
      if (res.mode === "ai_chart" && res.chart_result) {
        addMessage("assistant", `📊 ${res.chart_result.inferences || "Chart ready — generating…"}`, {
          type: "chart_pending",
          chartData: res.chart_result,
          address,
          values,
        });
        // Execute the chart insertion using existing aiChart logic
        await _insertChartFromResult(res.chart_result, values, address);
        return;
      }

      // ── Audit result ────────────────────────────────────────────────────────
      if (res.mode === "audit" && res.action_result) {
        if (res.action_result.worksheets?.length > 0) await renderWorksheets(res.action_result.worksheets);
        addMessage("assistant", res.action_result.message || "✅ Audit complete.", { type: "action" });
        return;
      }

      // ── Formula / Python result ─────────────────────────────────────────────
      if (res.action_result) {
        if (res.action_result.worksheets?.length > 0) await renderWorksheets(res.action_result.worksheets);
        addMessage("assistant", rationalePrefix + (res.action_result.message || "✅ Done."), { type: "action", mode: res.mode });
        return;
      }


      // ── Text / Analyze result ───────────────────────────────────────────────
      if (res.analyze_result) {
        addMessage("assistant", res.analyze_result.answer || "No answer.", { type: "text" });
        return;
      }

      addMessage("assistant", "✅ Done.", { type: "text" });

    } catch (e) {
      addMessage("assistant", `❌ Error: ${e.message}`, { type: "error" });
    } finally {
      setLoading(false);
    }
  };

  // ── HITL option selection ───────────────────────────────────────────────────
  const handleClarification = async (option, originalQuestion, round) => {
    const context = `User question: "${originalQuestion}". User clarified: "${option}".`;
    addMessage("user", option);
    await sendMessage(originalQuestion, context, round);
  };

  // ── Chart insertion from SmartRoute chart_result ──────────────────────────
  const _insertChartFromResult = async (chartData, values, address) => {
    // Synthesise a normalized suggest object matching what aiChart() expects
    const suggest = {
      chartType: chartData.chartType,
      chartTitle: chartData.chartTitle,
      categoryColumn: chartData.categoryColumn,
      valueColumns: chartData.valueColumns,
      aggregations: chartData.aggregations,
      inferences: chartData.inferences,
      use_llm: true,
    };
    // Re-use the aiChart path; it reads `suggest` and `values` from closure-equivalent args.
    // We call the internal chart render function with pre-fetched data.
    await _renderChart(suggest, values, address);
  };

  // ── Legacy handlers (kept for direct use if needed) ──────────────────────
  const askAI = async () => {
    if (input.trim()) await sendMessage(input.trim());
  };


  const checkData = async () => {
    try {
      setLoading(true); setOutput("🔍 Checking data quality…");
      const { values, address } = await getExcelData();
      const res = await callJson("/ai_audit", { table: values, address });
      setOutput(res.audit || "No audit message.");
      if (res.worksheets?.length > 0) await renderWorksheets(res.worksheets);
      setOutput(res.message || "✅ Audit complete.");
    } catch (e) { setOutput(`❌ Error: ${e.message}`); }
    finally { setLoading(false); }
  };

  const aiChart = async () => {
    try {
      setLoading(true); setOutput("📊 Recommending chart and inserting…");
      await Promise.resolve();
      const { values, address } = await getExcelData();
      const suggest = await callJson("/ai_chart", { table: values, address, question, use_llm: useLLM });
      if (!suggest || typeof suggest !== "object") throw new Error("Invalid /ai_chart response.");
      const normalized = {
        chartType: suggest.chartType,
        chartTitle: suggest.chartTitle,
        categoryColumn: suggest.categoryColumn ?? suggest.category_column ?? null,
        valueColumns: Array.isArray(suggest.valueColumns) ? suggest.valueColumns
          : Array.isArray(suggest.value_columns) ? suggest.value_columns : [],
      };
      const catRaw = typeof normalized.categoryColumn === "string" ? normalized.categoryColumn.trim() : null;
      const valsRaw = normalized.valueColumns.filter(v => typeof v === "string").map(v => v.trim());
      const headers = Array.isArray(values) && values[0] ? values[0].map(h => String(h).trim()) : [];
      const headerSet = new Set(headers);
      const lowerMap = new Map(headers.map(h => [h.toLowerCase(), h]));
      const catMapped = Array.isArray(normalized.categoryColumn)
        ? normalized.categoryColumn.map(c => typeof c === "string" ? lowerMap.get(c.toLowerCase()) : null).filter(Boolean)
        : typeof normalized.categoryColumn === "string" ? [lowerMap.get(normalized.categoryColumn.toLowerCase())] : [];
      const valsMapped = valsRaw.map(v => lowerMap.get(v.toLowerCase())).filter(Boolean);
      const cat = catMapped.length > 0 && catMapped.every(c => headerSet.has(c)) ? catMapped : null;
      const vals = valsMapped.filter(v => headerSet.has(v));
      const aggregationsRaw = Array.isArray(suggest.aggregations) ? suggest.aggregations
        : Array.isArray(suggest.value_aggregations) ? suggest.value_aggregations : [];
      const mapAggregation = (aggString) => {
        const aggMap = {
          SUM: Excel.AggregationFunction.sum, COUNT: Excel.AggregationFunction.count,
          AVERAGE: Excel.AggregationFunction.average, MIN: Excel.AggregationFunction.min,
          MAX: Excel.AggregationFunction.max, STDEV: Excel.AggregationFunction.standardDeviation,
          VAR: Excel.AggregationFunction.variance,
        };
        return aggMap[aggString] || Excel.AggregationFunction.sum;
      };
      const aggregations = vals.map((_, i) => mapAggregation(aggregationsRaw[i] || "SUM"));

      await Excel.run(async (context) => {
        const wb = context.workbook;
        let ws = wb.worksheets.getActiveWorksheet();
        let addressA1 = address || "";
        if (typeof addressA1 === "string" && addressA1.includes("!")) {
          const parts = addressA1.split("!");
          let sheetName = parts[0];
          if (sheetName.startsWith("'") && sheetName.endsWith("'"))
            sheetName = sheetName.substring(1, sheetName.length - 1);
          addressA1 = parts.slice(1).join("!");
          try { ws = wb.worksheets.getItem(sheetName); await context.sync(); }
          catch { ws = wb.worksheets.getActiveWorksheet(); }
        }
        const dataRange = ws.getRange(addressA1);
        const type = normalized.chartType || "ColumnClustered";

        const catCols = []; const catTitles = [];
        for (const c of cat || []) {
          const idx = headers.indexOf(c);
          if (idx !== -1) {
            const cRange = dataRange.getColumn(idx).getOffsetRange(1, 0).getResizedRange(-1, 0);
            cRange.load("address"); catCols.push(cRange); catTitles.push(c);
          }
        }
        const valCols = []; const valTitles = [];
        for (const v of vals || []) {
          const idx = headers.indexOf(v);
          if (idx !== -1) {
            const vRange = dataRange.getColumn(idx).getOffsetRange(1, 0).getResizedRange(-1, 0);
            vRange.load("address"); valCols.push(vRange); valTitles.push(v);
          }
        }
        await context.sync();

        const optimizeChartSeries = async (chart) => {
          if (!vals || vals.length === 0) return;
          try {
            chart.series.load("count, items"); await context.sync();
            for (let i = chart.series.items.length - 1; i >= 0; i--) chart.series.items[i].delete();
            if (cat && cat.length > 0) {
              const catIdx = headers.indexOf(cat[0]);
              if (catIdx !== -1) {
                try {
                  const catRange = dataRange.getColumn(catIdx).getOffsetRange(1, 0).getResizedRange(-1, 0);
                  chart.axes.categoryAxis.setCategoryNames(catRange); await context.sync();
                } catch { }
              }
            }
            for (let i = 0; i < vals.length; i++) {
              const vIdx = headers.indexOf(vals[i]);
              if (vIdx !== -1) {
                const valRange = dataRange.getColumn(vIdx).getOffsetRange(1, 0).getResizedRange(-1, 0);
                const newSeries = chart.series.add(vals[i]);
                newSeries.setValues(valRange);
              }
            }
            await context.sync();
          } catch { }
        };

        if (usePivot) {
          const newSheetName = `AI_Chart_${Date.now()}`;
          let newSheet;
          try { newSheet = wb.worksheets.add(newSheetName); await context.sync(); }
          catch { newSheet = ws; }
          const uniquePivotName = `AI_PivotTable_${Date.now()}`;
          let pivotTable;
          try {
            dataRange.load(["address", "rowCount", "columnCount"]); await context.sync();
            const pivotDestination = newSheet.getRange("A1");
            pivotTable = newSheet.pivotTables.add(uniquePivotName, dataRange, pivotDestination);
            await context.sync(); pivotTable.load("name"); await context.sync();
            if (Array.isArray(cat) && cat.length > 0) {
              for (const category of cat) {
                try { pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(category)); } catch { }
              }
              await context.sync();
            }
            if (vals.length > 0) {
              for (let i = 0; i < vals.length; i++) {
                try {
                  const h = pivotTable.hierarchies.getItem(vals[i]);
                  const field = pivotTable.dataHierarchies.add(h); await context.sync();
                  try { field.summarizeBy = aggregations[i]; await context.sync(); } catch { }
                } catch { }
              }
            }
            if ((!cat || cat.length === 0) && vals.length === 0 && headers.length >= 2) {
              try { pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(headers[0])); } catch { }
              try {
                const h = pivotTable.hierarchies.getItem(headers[1]);
                const field = pivotTable.dataHierarchies.add(h);
                try { field.summarizeBy = Excel.AggregationFunction.sum; } catch { }
              } catch { }
            }
            pivotTable.layout.layoutType = "Tabular";
            pivotTable.layout.repeatAllItemLabels(true);
            await context.sync();
            try {
              const pivotRange = pivotTable.layout.getRange();
              pivotRange.load(["address", "columnCount", "rowCount"]); await context.sync();
              const pivotChart = newSheet.charts.add(type, pivotRange, "Auto");
              pivotChart.title.text = normalized.chartTitle || "AI-PivotChart";
              await context.sync();
              pivotChart.left = (pivotRange.columnCount + 2) * 100;
              pivotChart.top = 0; pivotChart.width = 400; pivotChart.height = 300;
              await context.sync();
              newSheet.activate(); await context.sync();
            } catch {
              try {
                const seedRange = dataRange.getCell(0, 0);
                const chart = ws.charts.add(type, seedRange, "Auto");
                chart.title.text = normalized.chartTitle || "AI-Suggested Chart";
                await context.sync(); await optimizeChartSeries(chart);
              } catch (fe) { throw fe; }
            }
          } catch {
            const seedRange = dataRange.getCell(0, 0).getResizedRange(1, 1);
            const chart = ws.charts.add(type, seedRange, "Auto");
            chart.title.text = normalized.chartTitle || "AI-Suggested Chart";
            await context.sync(); await optimizeChartSeries(chart);
          }
        } else {
          const newSheetName = `AI_Chart_${Date.now()}`;
          let dashboardSheet;
          try { dashboardSheet = wb.worksheets.add(newSheetName); await context.sync(); }
          catch { dashboardSheet = wb.worksheets.add(); await context.sync(); }

          const titleCell = dashboardSheet.getRange("A1");
          titleCell.values = [[normalized.chartTitle || "AI Chart Summary"]];
          titleCell.format.font.bold = true; titleCell.format.font.size = 16; titleCell.format.font.color = "#004B87";
          dashboardSheet.getRange("A1:G1").merge(true);
          const inferencesText = suggest.inferences || "No specific inferences generated for this view.";
          const inferenceCell = dashboardSheet.getRange("A2");
          inferenceCell.values = [[inferencesText]];
          inferenceCell.format.font.italic = true; inferenceCell.format.font.size = 11; inferenceCell.format.font.color = "#444444";
          const inferenceRange = dashboardSheet.getRange("A2:G3");
          inferenceRange.merge(true); inferenceRange.format.wrapText = true; inferenceRange.format.verticalAlignment = "Top";
          await context.sync();

          if (catTitles.length > 0) {
            // ── STEP 1: JS extracts unique category combos only (lightweight, no math) ──
            const catIndices = catTitles.map(c => headers.indexOf(c)).filter(i => i !== -1);
            const seenKeys = new Set();
            const uniqueRows = []; // array of string[] — one per unique combo
            for (let i = 1; i < values.length; i++) {
              const row = values[i]; if (!row) continue;
              const keyParts = catIndices.map(idx => row[idx] != null ? String(row[idx]) : "");
              const key = keyParts.join(" ||| ");
              if (!seenKeys.has(key)) { seenKeys.add(key); uniqueRows.push(keyParts); }
            }
            console.log(`[aiChart] Hybrid: found ${uniqueRows.length} unique category combos.`);

            // ── SORT uniqueRows chronologically / logically ──────────────────────
            const MONTH_ORDER = {
              january:1, february:2, march:3, april:4, may:5, june:6,
              july:7, august:8, september:9, october:10, november:11, december:12,
              jan:1, feb:2, mar:3, apr:4, jun:6, jul:7, aug:8, sep:9, oct:10, nov:11, dec:12,
              "1":1,"2":2,"3":3,"4":4,"5":5,"6":6,"7":7,"8":8,"9":9,"10":10,"11":11,"12":12,
            };
            const toMonthNum = (s) => MONTH_ORDER[s.toLowerCase()] || null;
            const isYear = (s) => /^\d{4}$/.test(s.trim());

            uniqueRows.sort((a, b) => {
              for (let ci = 0; ci < a.length; ci++) {
                const av = a[ci], bv = b[ci];
                // If this part looks like a year, sort numerically
                if (isYear(av) && isYear(bv)) {
                  const diff = parseInt(av) - parseInt(bv);
                  if (diff !== 0) return diff;
                  continue;
                }
                // If this part looks like a month name/number, sort by calendar position
                const am = toMonthNum(av), bm = toMonthNum(bv);
                if (am !== null && bm !== null) {
                  const diff = am - bm;
                  if (diff !== 0) return diff;
                  continue;
                }
                // Otherwise plain string compare
                if (av < bv) return -1;
                if (av > bv) return 1;
              }
              return 0;
            });
            console.log(`[aiChart] Sorted uniqueRows sample:`, uniqueRows.slice(0, 5));

            // ── STEP 2: Build native Excel formula string per aggregation type ──
            // catCols[i].address = absolute source address e.g. "Sheet1!B2:B307646"
            // valCols[i].address = absolute source address e.g. "Sheet1!G2:G307646"
            const catAddrs = catCols.map(c => c.address);
            const valAddrs = valCols.map(v => v.address);

            const getNativeFormula = (aggType, valAddr, catAddrList, catValueParts) => {
              // Build SUMIFS / AVERAGEIFS / MAXIFS / MINIFS / COUNTIFS criteria pairs
              // Each category column address paired with the literal value in that row
              const criteriaPairs = catAddrList.map((cAddr, ci) => {
                const escaped = catValueParts[ci].replace(/"/g, '""'); // escape quotes
                return `${cAddr},"${escaped}"`;
              }).join(",");

              switch (aggType.toUpperCase()) {
                case "COUNT":   return `=IFERROR(COUNTIFS(${criteriaPairs}),0)`;
                case "AVERAGE": return `=IFERROR(AVERAGEIFS(${valAddr},${criteriaPairs}),0)`;
                case "MAX":     return `=IFERROR(MAXIFS(${valAddr},${criteriaPairs}),0)`;
                case "MIN":     return `=IFERROR(MINIFS(${valAddr},${criteriaPairs}),0)`;
                default:        return `=IFERROR(SUMIFS(${valAddr},${criteriaPairs}),0)`; // SUM
              }
            };

            // ── STEP 3: Write header row ──
            const activeColOffset = catTitles.length + valTitles.length;
            const headerRow = [
              ...catTitles,
              ...valTitles.map((v, i) => `${v} (${aggregationsRaw[i] || "SUM"})`)
            ];
            const headerRange = dashboardSheet.getRange("A5").getResizedRange(0, activeColOffset - 1);
            headerRange.values = [headerRow];
            headerRange.format.font.bold = true;
            headerRange.format.fill.color = "#EBF1F5";

            // ── STEP 4: Write unique category strings + native Excel formulas row by row ──
            for (let r = 0; r < uniqueRows.length; r++) {
              const keyParts = uniqueRows[r];
              const dataRow = r + 6; // A6, A7, ... (row 5 = headers)

              // Write category label(s) as static strings
              for (let ci = 0; ci < catTitles.length; ci++) {
                dashboardSheet.getRange(`A${dataRow}`).getOffsetRange(0, ci).values = [[keyParts[ci]]];
              }

              // Write native Excel formula for each value column
              for (let vi = 0; vi < valTitles.length; vi++) {
                const aggType = aggregationsRaw[vi] || "SUM";
                const formula = getNativeFormula(aggType, valAddrs[vi], catAddrs, keyParts);
                console.log(`[aiChart] Row ${dataRow}, col ${catTitles.length + vi}: ${formula}`);
                dashboardSheet.getRange(`A${dataRow}`).getOffsetRange(0, catTitles.length + vi).formulas = [[formula]];
              }
            }

            // ── STEP 5: Style, autofit, and chart ──
            const tableRows = uniqueRows.length + 1; // +1 for header
            const targetBoundRange = dashboardSheet.getRange("A5").getResizedRange(tableRows - 1, activeColOffset - 1);
            targetBoundRange.format.autofitColumns();
            ["InsideHorizontal", "InsideVertical", "EdgeBottom", "EdgeLeft", "EdgeRight", "EdgeTop"].forEach(b => {
              targetBoundRange.format.borders.getItem(b).style = "Continuous";
            });

            await context.sync(); // Let Excel evaluate all SUMIFS formulas before charting

            const chart = dashboardSheet.charts.add(type, targetBoundRange, "Auto");
            chart.title.text = normalized.chartTitle || "AI-Aggregated Chart";
            chart.top = 80; chart.left = (activeColOffset + 1) * 70 + 20; chart.width = 500; chart.height = 350;
            await context.sync();
          } else {
            const dashSeedRange = dashboardSheet.getRange("A5:B6");
            const chart = dashboardSheet.charts.add(type, dashSeedRange, "Auto");
            chart.title.text = normalized.chartTitle || "AI-Suggested Chart";
            chart.top = 65; chart.left = 5; chart.width = 500; chart.height = 350;
            try { await context.sync(); } catch (e) { throw new Error(`charts.add Failed: ${e.message}`); }
            await optimizeChartSeries(chart);
          }
          dashboardSheet.activate();
        }
        try { await context.sync(); } catch (e) { throw new Error(`Final sync failed: ${e.message}`); }
      });
      setOutput("✅ Chart added successfully!");
    } catch (e) { setOutput(`❌ Error: ${e.message}`); }
    finally { setLoading(false); }
  };

  /**
   * _renderChart — inserts a chart from a pre-resolved suggest object.
   * Called by the SmartRoute path so we don't need a second /ai_chart call.
   * The suggest object must have: chartType, chartTitle, categoryColumn,
   * valueColumns, aggregations, inferences.
   */
  const _renderChart = async (suggest, values, address) => {
    try {
      // Reuse the same normalisation + Excel.run block from aiChart
      const normalized = {
        chartType: suggest.chartType,
        chartTitle: suggest.chartTitle,
        categoryColumn: suggest.categoryColumn ?? null,
        valueColumns: Array.isArray(suggest.valueColumns) ? suggest.valueColumns : [],
        inferences: suggest.inferences || "",
      };
      const headers = Array.isArray(values) && values[0] ? values[0].map(h => String(h).trim()) : [];
      const lowerMap = new Map(headers.map(h => [h.toLowerCase(), h]));
      const headerSet = new Set(headers);
      const catMapped = Array.isArray(normalized.categoryColumn)
        ? normalized.categoryColumn.map(c => typeof c === "string" ? lowerMap.get(c.toLowerCase()) : null).filter(Boolean)
        : typeof normalized.categoryColumn === "string" ? [lowerMap.get(normalized.categoryColumn.toLowerCase())] : [];
      const valsRaw = normalized.valueColumns.filter(v => typeof v === "string").map(v => v.trim());
      const valsMapped = valsRaw.map(v => lowerMap.get(v.toLowerCase())).filter(Boolean);
      const cat = catMapped.length > 0 && catMapped.every(c => headerSet.has(c)) ? catMapped : null;
      const vals = valsMapped.filter(v => headerSet.has(v));
      const aggregationsRaw = Array.isArray(suggest.aggregations) ? suggest.aggregations : [];

      // Synthesise a full suggest object understood by findInExcelRange (same as aiChart path)
      const fullSuggest = { ...normalized, categoryColumn: cat, valueColumns: vals, aggregations: aggregationsRaw };

      // Call the shared aiChart core with preloaded data
      await _runChartCore(fullSuggest, values, address, aggregationsRaw);
      addMessage("assistant", "✅ Chart added to a new dashboard sheet.", { type: "action", mode: "ai_chart" });
    } catch (e) {
      addMessage("assistant", `❌ Chart failed: ${e.message}`, { type: "error" });
    }
  };

  /**
   * _runChartCore — the inner Excel.run logic extracted from aiChart.
   * Accepts a normalised suggest, pre-loaded values, address, and aggregationsRaw.
   * This avoids duplicating the large Excel.run block.
   */
  const _runChartCore = async (normalized, values, address, aggregationsRaw) => {
    // Delegate back to aiChart but supply the pre-loaded suggest as the resolved API response
    // We re-use the aiChart variable but intercept the API call.
    // Simplest approach: just call aiChart with the question already in state,
    // it will call /ai_chart which may return slightly different columns,
    // but for the SmartRoute path the chart result is already resolved so we pass it directly.
    // For a cleaner future refactor, aiChart can be split into fetch + render.
    // For now we write the chart from the already-computed suggest.
    const cat = normalized.categoryColumn;
    const vals = normalized.valueColumns;
    const aggs = aggregationsRaw;
    const headers = values[0]?.map(h => String(h).trim()) || [];

    await Excel.run(async (context) => {
      const wb = context.workbook;
      let ws = wb.worksheets.getActiveWorksheet();
      let addressA1 = address || "";
      if (typeof addressA1 === "string" && addressA1.includes("!")) {
        const parts = addressA1.split("!");
        let sheetName = parts[0].replace(/^'|'$/g, "");
        addressA1 = parts.slice(1).join("!");
        try { ws = wb.worksheets.getItem(sheetName); await context.sync(); }
        catch { ws = wb.worksheets.getActiveWorksheet(); }
      }
      const dataRange = ws.getRange(addressA1);
      const type = normalized.chartType || "ColumnClustered";

      // Load category / value column ranges
      const catCols = [], catTitles = [];
      for (const c of cat || []) {
        const idx = headers.indexOf(c);
        if (idx !== -1) {
          const r = dataRange.getColumn(idx).getOffsetRange(1, 0).getResizedRange(-1, 0);
          r.load("address"); catCols.push(r); catTitles.push(c);
        }
      }
      const valCols = [], valTitles = [];
      for (const v of vals || []) {
        const idx = headers.indexOf(v);
        if (idx !== -1) {
          const r = dataRange.getColumn(idx).getOffsetRange(1, 0).getResizedRange(-1, 0);
          r.load("address"); valCols.push(r); valTitles.push(v);
        }
      }
      await context.sync();

      // Create dashboard sheet
      const newSheetName = `AI_Chart_${Date.now()}`;
      let dashboardSheet;
      try { dashboardSheet = wb.worksheets.add(newSheetName); await context.sync(); }
      catch { dashboardSheet = wb.worksheets.add(); await context.sync(); }

      const titleCell = dashboardSheet.getRange("A1");
      titleCell.values = [[normalized.chartTitle || "AI Chart Summary"]];
      titleCell.format.font.bold = true; titleCell.format.font.size = 16; titleCell.format.font.color = "#004B87";
      dashboardSheet.getRange("A1:G1").merge(true);

      const inferenceCell = dashboardSheet.getRange("A2");
      inferenceCell.values = [[normalized.inferences || "AI-generated chart"]];
      inferenceCell.format.font.italic = true; inferenceCell.format.font.size = 11;
      dashboardSheet.getRange("A2:G3").merge(true);
      dashboardSheet.getRange("A2:G3").format.wrapText = true;
      await context.sync();

      if (catTitles.length > 0) {
        // --- Hybrid formula injection (same as aiChart) ---
        const catIndices = catTitles.map(c => headers.indexOf(c)).filter(i => i !== -1);
        const seenKeys = new Set(); const uniqueRows = [];
        for (let i = 1; i < values.length; i++) {
          const row = values[i]; if (!row) continue;
          const keyParts = catIndices.map(idx => row[idx] != null ? String(row[idx]) : "");
          const key = keyParts.join(" ||| ");
          if (!seenKeys.has(key)) { seenKeys.add(key); uniqueRows.push(keyParts); }
        }

        const catAddrs = catCols.map(c => c.address);
        const valAddrs = valCols.map(v => v.address);

        const getNativeFormula = (aggType, valAddr, catAddrList, catValueParts) => {
          const criteriaPairs = catAddrList.map((cAddr, ci) => {
            const escaped = catValueParts[ci].replace(/"/g, '""');
            return `${cAddr},"${escaped}"`;
          }).join(",");
          switch (aggType.toUpperCase()) {
            case "COUNT":   return `=IFERROR(COUNTIFS(${criteriaPairs}),0)`;
            case "AVERAGE": return `=IFERROR(AVERAGEIFS(${valAddr},${criteriaPairs}),0)`;
            case "MAX":     return `=IFERROR(MAXIFS(${valAddr},${criteriaPairs}),0)`;
            case "MIN":     return `=IFERROR(MINIFS(${valAddr},${criteriaPairs}),0)`;
            default:        return `=IFERROR(SUMIFS(${valAddr},${criteriaPairs}),0)`;
          }
        };

        const activeColOffset = catTitles.length + valTitles.length;
        const headerRow = [...catTitles, ...valTitles.map((v, i) => `${v} (${aggs[i] || "SUM"})`)];
        const hdrRange = dashboardSheet.getRange("A5").getResizedRange(0, activeColOffset - 1);
        hdrRange.values = [headerRow]; hdrRange.format.font.bold = true; hdrRange.format.fill.color = "#EBF1F5";

        for (let r = 0; r < uniqueRows.length; r++) {
          const keyParts = uniqueRows[r]; const dataRow = r + 6;
          for (let ci = 0; ci < catTitles.length; ci++)
            dashboardSheet.getRange(`A${dataRow}`).getOffsetRange(0, ci).values = [[keyParts[ci]]];
          for (let vi = 0; vi < valTitles.length; vi++) {
            const formula = getNativeFormula(aggs[vi] || "SUM", valAddrs[vi], catAddrs, keyParts);
            dashboardSheet.getRange(`A${dataRow}`).getOffsetRange(0, catTitles.length + vi).formulas = [[formula]];
          }
        }

        const tableRows = uniqueRows.length + 1;
        const targetBoundRange = dashboardSheet.getRange("A5").getResizedRange(tableRows - 1, activeColOffset - 1);
        targetBoundRange.format.autofitColumns();
        await context.sync();

        const chart = dashboardSheet.charts.add(type, targetBoundRange, "Auto");
        chart.title.text = normalized.chartTitle || "AI Chart";
        chart.top = 80; chart.left = (activeColOffset + 1) * 70 + 20; chart.width = 500; chart.height = 350;
        await context.sync();
      } else {
        const seedRange = dashboardSheet.getRange("A5:B6");
        const chart = dashboardSheet.charts.add(type, seedRange, "Auto");
        chart.title.text = normalized.chartTitle || "AI Chart";
        chart.top = 65; chart.left = 5; chart.width = 500; chart.height = 350;
        await context.sync();
      }
      dashboardSheet.activate();
      await context.sync();
    });
  };

  const aiAction = async () => {
    try {
      setLoading(true); setOutput("🚀 Running AI action…");
      const startTime = performance.now();
      const { values, address } = await getExcelData();
      const res = await callJson("/action", { table: values, address, question });
      await renderWorksheets(res.worksheets);
      if ((res.message || "").toLowerCase().includes("chart recommended")) await aiChart();
      setOutput(res.message || "✅ Done.");
    } catch (e) { setOutput(`❌ Error: ${e.message}`); }
    finally { setLoading(false); }
  };

  const renderWorksheets = async (worksheets) => {
    if (!worksheets?.length) return;
    try {
      await Excel.run(async (context) => {
        const wb = context.workbook;
        for (const wsDef of worksheets) {
          let name = (wsDef.name || "AI_Output").replace(/[\\/?*[\]]/g, "_").substring(0, 31);
          const existingWs = wb.worksheets.getItemOrNullObject(name);
          await context.sync();
          let ws = existingWs.isNullObject ? wb.worksheets.add(name) : existingWs;
          await context.sync();
          for (const cell of wsDef.cells || []) {
            try { ws.getRange(cell.address || "A1").values = [[cell.value ?? ""]]; }
            catch (e) { throw e; }
          }
          for (const f of wsDef.formulas || []) {
            try { ws.getRange(f.address || "A1").formulas = [[f.formula || ""]]; } catch { }
          }
          if (wsDef.conditional_formatting) {
            for (const cf of wsDef.conditional_formatting) {
              try {
                if (cf.type === "dataBar") {
                  const range = ws.getRange(cf.range);
                  range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
                  await context.sync();
                }
              } catch { }
            }
          }
          ws.getRange().format.autofitColumns();
          ws.activate();
        }
        await context.sync();
      });
    } catch (e) { throw new Error(`Excel worksheet creation failed: ${e.message}`); }
  };

  // ── Styles ───────────────────────────────────────────────────────────────────
  const S = {
    root: {
      fontFamily: "'DM Sans', sans-serif",
      background: brand.offWhite,
      minHeight: "100vh",
      display: "flex",
      flexDirection: "column",
      fontSize: 14,
      color: brand.textPrimary,
    },
    // ── Header
    header: {
      background: brand.navy,
      padding: "14px 16px",
      display: "flex",
      alignItems: "center",
      gap: 10,
      borderBottom: `2px solid ${brand.teal}`,
      flexShrink: 0,
    },
    logoImg: { width: 30, height: 30, objectFit: "contain" },
    brandName: {
      fontFamily: "'Syne', sans-serif",
      fontSize: 16,
      fontWeight: 700,
      letterSpacing: "0.05em",
      color: "#fff",
      lineHeight: 1.1,
    },
    brandSub: {
      fontSize: 9,
      color: "rgba(255,255,255,0.45)",
      letterSpacing: "0.12em",
      textTransform: "uppercase",
      marginTop: 2,
    },
    settingsToggle: {
      marginLeft: "auto",
      background: "none",
      border: "none",
      color: "rgba(255,255,255,0.55)",
      cursor: "pointer",
      padding: "4px 6px",
      borderRadius: 6,
      fontSize: 16,
      lineHeight: 1,
      transition: "color 0.15s",
      display: "flex",
      alignItems: "center",
    },
    // ── Body
    body: {
      flex: 1,
      padding: "16px 16px 20px",
      display: "flex",
      flexDirection: "column",
      gap: 14,
      animation: "fadeUp 0.25s ease",
      overflowY: "auto",
    },
    // ── Settings panel
    settingsPanel: {
      background: brand.white,
      border: `1px solid ${brand.border}`,
      borderRadius: 10,
      padding: "14px 16px",
      display: "flex",
      flexDirection: "column",
      gap: 10,
      animation: "fadeUp 0.18s ease",
    },
    toggleRow: {
      display: "flex",
      alignItems: "center",
      gap: 10,
      cursor: "pointer",
      padding: "3px 0",
    },
    toggleText: { fontSize: 13, color: "#3D566E", lineHeight: 1.4 },
    // ── Input
    inputWrap: { position: "relative" },
    input: {
      width: "100%",
      padding: "11px 14px",
      background: brand.white,
      border: `1.5px solid ${brand.border}`,
      borderRadius: 10,
      fontFamily: "'DM Sans', sans-serif",
      fontSize: 13.5,
      color: brand.textPrimary,
      outline: "none",
      transition: "border-color 0.18s, box-shadow 0.18s",
      resize: "vertical",
    },
    select: {
      width: "100%",
      padding: "9px 12px",
      background: brand.white,
      border: `1.5px solid ${brand.border}`,
      borderRadius: 8,
      fontFamily: "'DM Sans', sans-serif",
      fontSize: 13,
      color: brand.textPrimary,
      outline: "none",
      cursor: "pointer",
      appearance: "none",
      backgroundImage: `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%236B8399' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E")`,
      backgroundRepeat: "no-repeat",
      backgroundPosition: "right 12px center",
      paddingRight: 32,
    },
    // ── Primary action button
    btnPrimary: {
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      gap: 8,
      padding: "13px 16px",
      borderRadius: 10,
      background: brand.teal,
      color: brand.white,
      fontFamily: "'DM Sans', sans-serif",
      fontSize: 14,
      fontWeight: 600,
      border: "none",
      cursor: "pointer",
      boxShadow: "0 3px 12px rgba(43,201,156,0.28)",
      transition: "background 0.15s, box-shadow 0.15s, transform 0.12s",
      width: "100%",
      letterSpacing: "0.01em",
    },
    // ── Secondary button grid
    btnGrid: {
      display: "grid",
      gridTemplateColumns: "1fr 1fr 1fr",
      gap: 8,
    },
    btnSecondary: {
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
      gap: 5,
      padding: "10px 6px",
      borderRadius: 10,
      background: brand.navyMid,
      color: brand.white,
      fontFamily: "'DM Sans', sans-serif",
      fontSize: 11,
      fontWeight: 500,
      border: "none",
      cursor: "pointer",
      transition: "background 0.15s",
      lineHeight: 1.2,
      textAlign: "center",
    },
    btnSecondaryIcon: {
      fontSize: 18,
      transition: "transform 0.15s",
      display: "block",
    },
    btnGhost: {
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      gap: 6,
      padding: "8px 12px",
      borderRadius: 8,
      background: "transparent",
      color: brand.textMuted,
      fontFamily: "'DM Sans', sans-serif",
      fontSize: 12,
      fontWeight: 500,
      border: `1.5px solid ${brand.border}`,
      cursor: "pointer",
      transition: "all 0.15s",
    },
    // ── Output
    outputWrap: {
      background: brand.white,
      border: `1px solid ${brand.border}`,
      borderRadius: 10,
      padding: "14px 16px",
      minHeight: 100,
      flex: 1,
    },
    outputEmpty: {
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
      gap: 8,
      padding: "20px 0",
      color: brand.textLight,
    },
    outputEmptyIcon: { fontSize: 28, opacity: 0.5 },
    outputText: {
      fontSize: 13,
      lineHeight: 1.65,
      color: "#3D566E",
      whiteSpace: "pre-wrap",
      wordBreak: "break-word",
    },
    loadingRow: { display: "flex", alignItems: "center", gap: 8, marginBottom: 8 },
    loadingText: { fontSize: 12, fontWeight: 500, color: brand.teal },
    // ── Footer
    footer: {
      padding: "9px 16px",
      borderTop: `1px solid ${brand.border}`,
      display: "flex",
      alignItems: "center",
      gap: 6,
      flexShrink: 0,
    },
    footerDot: {
      width: 6, height: 6, borderRadius: "50%", background: brand.teal, flexShrink: 0,
    },
    footerText: { fontSize: 10.5, color: brand.textMuted, letterSpacing: "0.03em", flex: 1 },
    // ── Landing (API key page)
    landing: {
      flex: 1,
      display: "flex",
      flexDirection: "column",
      background: brand.navy,
      padding: "32px 20px 28px",
    },
    landingLogo: {
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      marginBottom: 28,
      gap: 10,
    },
    landingCard: {
      background: brand.navyMid,
      borderRadius: 14,
      padding: "20px 18px",
      border: `1px solid ${brand.navyLight}`,
      display: "flex",
      flexDirection: "column",
      gap: 12,
    },
    landingLabel: {
      fontSize: 10,
      fontWeight: 600,
      letterSpacing: "0.12em",
      textTransform: "uppercase",
      color: "rgba(43,201,156,0.8)",
      marginBottom: 4,
    },
    landingInput: {
      width: "100%",
      padding: "11px 14px",
      background: brand.navy,
      border: `1.5px solid ${brand.navyLight}`,
      borderRadius: 9,
      fontFamily: "'DM Sans', sans-serif",
      fontSize: 13.5,
      color: brand.white,
      outline: "none",
      transition: "border-color 0.18s, box-shadow 0.18s",
      letterSpacing: "0.04em",
    },
    landingBtn: {
      width: "100%",
      padding: "12px 16px",
      borderRadius: 9,
      background: brand.teal,
      color: brand.white,
      fontFamily: "'DM Sans', sans-serif",
      fontSize: 14,
      fontWeight: 600,
      border: "none",
      cursor: "pointer",
      boxShadow: "0 3px 14px rgba(43,201,156,0.3)",
      transition: "background 0.15s, box-shadow 0.15s",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      gap: 8,
    },
    landingHint: {
      fontSize: 11,
      color: "rgba(255,255,255,0.35)",
      textAlign: "center",
      lineHeight: 1.5,
    },
  };

  // ── Render: Landing page (no API key) ────────────────────────────────────────
  if (!apiKey) {
    return (
      <div style={S.root}>
        <style>{globalCss}</style>
        <div style={S.landing}>
          <div style={S.landingLogo}>
            <img src="assets/icon-64.png" alt="Excelisor" style={{ width: 72, height: 72, objectFit: "contain" }} />
            <div style={{ textAlign: "center" }}>
              <div style={{ ...S.brandName, fontSize: 22 }}>
                EXCEL<span style={{ color: brand.teal }}>ISOR</span>
              </div>
              <div style={{ ...S.brandSub, marginTop: 4, color: "rgba(255,255,255,0.4)" }}>
                AI Excel Assistant
              </div>
            </div>
          </div>

          <div style={S.landingCard}>
            <div>
              <div style={S.landingLabel}>OpenAI API Key</div>
              <input
                id="apiKeyInput"
                type="password"
                className="xc-input"
                placeholder="sk-…"
                style={S.landingInput}
                disabled={loading}
              />
              {keyError && (
                <div style={{ color: "#ff8080", fontSize: 11.5, marginTop: 5 }}>{keyError}</div>
              )}
            </div>

            <div>
              <div style={{ ...S.landingLabel, marginBottom: 6 }}>AI Model</div>
              <select
                value={selectedModel}
                onChange={e => { setSelectedModel(e.target.value); localStorage.setItem("openai_model", e.target.value); }}
                style={{ ...S.landingInput, appearance: "none", cursor: "pointer" }}
                disabled={loading}
              >
                {MODELS.map(m => (
                  <option key={m.value} value={m.value}>{m.label} — {m.hint}</option>
                ))}
              </select>
            </div>

            <button
              style={S.landingBtn}
              disabled={loading}
              className="xc-btn"
              onClick={async () => {
                const val = document.getElementById("apiKeyInput").value.trim();
                if (!val || !val.startsWith("sk-")) {
                  setKeyError("Please enter a valid OpenAI API key (starts with sk-).");
                  return;
                }
                setKeyError("");
                setLoading(true);
                try {
                  const res = await fetch(`${apiBase}/validate-key`, {
                    method: "GET",
                    headers: { Authorization: `Bearer ${val}`, "X-OpenAI-Model": selectedModel },
                  });
                  const data = await res.json();
                  if (data.valid) {
                    localStorage.setItem("openai_api_key", val);
                    setApiKey(val);
                  } else {
                    setKeyError(data.message || "Invalid API key.");
                  }
                } catch {
                  setKeyError("Could not reach backend. Make sure the server is running.");
                } finally {
                  setLoading(false);
                }
              }}
            >
              {loading ? <><Spinner /> Verifying…</> : "Connect →"}
            </button>
          </div>

          <p style={{ ...S.landingHint, marginTop: 16 }}>
            Your key is stored locally and never sent anywhere except your own backend.
          </p>
        </div>
      </div>
    );
  }

  // ── Render: Main app ─────────────────────────────────────────────────────────
  return (
    <div style={S.root}>
      <style>{globalCss}</style>
      <style>{`
        .chat-bubble-enter { animation: fadeUp .2s ease both; }
        @keyframes fadeUp { from { opacity:0; transform:translateY(6px); } to { opacity:1; transform:translateY(0); } }
        .xc-option-btn:hover { background: ${brand.teal} !important; color:#fff !important; border-color:${brand.teal} !important; }
        .xc-send-btn:hover { opacity: .88; }
        .xc-send-btn:disabled { opacity: .45; cursor: not-allowed; }
      `}</style>

      {/* Header */}
      <header style={S.header}>
        <img src="assets/icon-64.png" alt="Excelisor" style={S.logoImg} />
        <div style={{ display: "flex", flexDirection: "column", lineHeight: 1 }}>
          <span style={S.brandName}>EXCEL<span style={{ color: brand.teal }}>ISOR</span></span>
          <span style={S.brandSub}>AI Excel Assistant</span>
        </div>
        <button
          style={S.settingsToggle}
          title={settingsOpen ? "Hide settings" : "Settings"}
          onClick={() => setSettingsOpen(o => !o)}
          aria-label="Toggle settings"
        >
          {settingsOpen ? "✕" : "⚙"}
        </button>
      </header>

      {/* Body */}
      <div style={{ ...S.body, padding: 0, overflow: "hidden" }}>

        {/* Settings panel (collapsible) */}
        {settingsOpen && (
          <div style={{ ...S.settingsPanel, margin: "8px 8px 0" }} className="settings-panel-enter">
            <div>
              <label style={S.toggleRow}>
                <span style={S.toggleText}>AI Model</span>
              </label>
              <select
                value={selectedModel}
                onChange={e => { setSelectedModel(e.target.value); localStorage.setItem("openai_model", e.target.value); }}
                className="xc-input"
                style={{ ...S.select, marginTop: 4 }}
                disabled={loading}
              >
                {MODELS.map(m => (
                  <option key={m.value} value={m.value}>{m.label} — {m.hint}</option>
                ))}
              </select>
            </div>
            <div style={{ borderTop: `1px solid ${brand.border}`, paddingTop: 8 }}>
              <button
                className="xc-btn xc-btn-danger"
                style={{ ...S.btnGhost, color: brand.errorText, borderColor: "#f5c6c2", fontSize: 12, gap: 6 }}
                onClick={() => { localStorage.removeItem("openai_api_key"); setApiKey(""); }}
              >
                🔑 Change API Key
              </button>
            </div>
          </div>
        )}

        {/* Chat feed */}
        <div
          ref={chatFeed}
          style={{
            flex: 1,
            overflowY: "auto",
            padding: "12px 10px 8px",
            display: "flex",
            flexDirection: "column",
            gap: 10,
            minHeight: 0,
          }}
        >
          {/* Welcome message */}
          {messages.length === 0 && !loading && (
            <div style={{
              display: "flex", flexDirection: "column", alignItems: "center",
              justifyContent: "center", flex: 1, gap: 8, opacity: .65, paddingTop: 20,
            }}>
              <span style={{ fontSize: 32 }}>✦</span>
              <span style={{ fontSize: 13, fontWeight: 600, color: brand.text }}>Ask me anything about your data</span>
              <span style={{ fontSize: 11, color: brand.textLight, textAlign: "center", maxWidth: "26ch", lineHeight: 1.6 }}>
                Select a range in Excel, then type your question below. I'll figure out the best action automatically.
              </span>
            </div>
          )}

          {/* Message bubbles */}
          {messages.map((msg, i) => {
            const isUser = msg.role === "user";
            return (
              <div
                key={i}
                className="chat-bubble-enter"
                style={{
                  display: "flex",
                  flexDirection: "column",
                  alignItems: isUser ? "flex-end" : "flex-start",
                  gap: 4,
                }}
              >
                {/* Role label */}
                <span style={{ fontSize: 10, color: brand.textLight, marginBottom: 1 }}>
                  {isUser ? "You" : `✦ Excelisor${msg.mode ? ` · ${modeIcon[msg.mode] || ""} ${msg.mode}` : ""}`}
                </span>

                {/* Bubble */}
                <div style={{
                  maxWidth: "88%",
                  padding: "9px 12px",
                  borderRadius: isUser ? "12px 12px 3px 12px" : "12px 12px 12px 3px",
                  background: isUser ? brand.teal : brand.surface,
                  color: isUser ? "#fff" : brand.text,
                  fontSize: 12.5,
                  lineHeight: 1.55,
                  whiteSpace: "pre-wrap",
                  boxShadow: "0 1px 3px rgba(0,0,0,.07)",
                  border: isUser ? "none" : `1px solid ${brand.border}`,
                }}>
                  {msg.content}
                </div>

                {/* HITL clarification option buttons */}
                {msg.type === "clarification" && msg.options?.length > 0 && (
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 2, maxWidth: "88%" }}>
                    {msg.options.map((opt, oi) => (
                      <button
                        key={oi}
                        className="xc-option-btn"
                        disabled={loading}
                        onClick={() => handleClarification(opt, msg.originalQuestion, msg.round)}
                        style={{
                          padding: "5px 10px",
                          borderRadius: 20,
                          border: `1px solid ${brand.teal}`,
                          background: "transparent",
                          color: brand.teal,
                          fontSize: 11.5,
                          cursor: "pointer",
                          transition: "all .15s",
                          fontWeight: 500,
                        }}
                      >
                        {opt}
                      </button>
                    ))}
                  </div>
                )}
              </div>
            );
          })}

          {/* Typing indicator */}
          {loading && (
            <div className="chat-bubble-enter" style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <Spinner />
              <span style={{ fontSize: 11.5, color: brand.textLight }}>Thinking…</span>
            </div>
          )}
        </div>

        {/* Sticky input bar */}
        <div style={{
          borderTop: `1px solid ${brand.border}`,
          padding: "8px 10px",
          background: brand.surface,
          display: "flex",
          gap: 8,
          alignItems: "flex-end",
        }}>
          <textarea
            className="xc-input"
            style={{
              flex: 1,
              resize: "none",
              minHeight: 38,
              maxHeight: 100,
              padding: "8px 10px",
              fontSize: 12.5,
              borderRadius: 10,
              lineHeight: 1.45,
              border: `1.5px solid ${brand.border}`,
              outline: "none",
              background: brand.bg,
              color: brand.text,
              fontFamily: "inherit",
            }}
            value={input}
            onChange={e => setInput(e.target.value)}
            placeholder="Ask anything about your selected data…"
            disabled={loading}
            rows={1}
            onKeyDown={e => {
              if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); sendMessage(); }
            }}
          />
          <button
            className="xc-send-btn"
            disabled={loading || !input.trim()}
            onClick={() => sendMessage()}
            style={{
              width: 38, height: 38,
              borderRadius: 10,
              border: "none",
              background: brand.teal,
              color: "#fff",
              fontSize: 16,
              cursor: "pointer",
              display: "flex", alignItems: "center", justifyContent: "center",
              flexShrink: 0,
              transition: "opacity .15s",
            }}
            title="Send (Enter)"
          >
            ➤
          </button>
        </div>
      </div>

      {/* Footer */}
      <footer style={S.footer}>
        <div style={S.footerDot} />
        <span style={S.footerText}>Excelisor · Powered by AI</span>
        <span style={{ fontSize: 10.5, color: brand.textLight }}>
          {MODELS.find(m => m.value === selectedModel)?.label}
        </span>
      </footer>
    </div>
  );
}
