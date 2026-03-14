/**
 * 健檢報告產生程式 — React 主元件
 * 仿照 Python Tkinter 版的深色 UI 風格
 */
import { useState, useCallback, useRef } from "react";
import { parseExcel } from "../lib/excel-reader";
import { fillReport, makeFilename } from "../lib/docx-generator";
import JSZip from "jszip";

function saveAs(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
import type { PatientRecord } from "../lib/excel-reader";

// ── 色系 ──────────────────────────────────────────────────────────────────────
const C = {
  bgDark:   "#1E2230",
  bgPanel:  "#252B3B",
  bgRowA:   "#2A3147",
  bgRowB:   "#242A3A",
  bgHover:  "#313A52",
  accent:   "#4C9BE8",
  accent2:  "#5BBFA8",
  warn:     "#E8864C",
  success:  "#5BB870",
  textMain: "#E8EAF0",
  textDim:  "#8A94B0",
  textHead: "#FFFFFF",
  border:   "#353D52",
  msBg:     "#3A2A2A",
  msFg:     "#FFB3A0",
};

// ── 欄位定義 ─────────────────────────────────────────────────────────────────
const COLS = [
  { key: "seq",       label: "序",     w: 40  },
  { key: "date",      label: "體檢日",  w: 90  },
  { key: "name",      label: "姓名",    w: 80  },
  { key: "gender",    label: "性別",    w: 48  },
  { key: "age",       label: "年齡",    w: 45  },
  { key: "bmi",       label: "BMI",     w: 60  },
  { key: "bp",        label: "血壓",    w: 90  },
  { key: "glucose",   label: "血糖",    w: 65  },
  { key: "metabolic", label: "代謝症候群", w: 90 },
  { key: "msItems",   label: "風險項目", w: 220 },
];

type SortKey = "date" | "name" | "gender" | "age" | "bmi" | "bp" | "glucose" | "metabolic" | null;

export default function ReportApp() {
  const [excelFile,    setExcelFile]    = useState<File | null>(null);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [records,      setRecords]      = useState<PatientRecord[]>([]);
  const [filtered,     setFiltered]     = useState<PatientRecord[]>([]);
  const [searchKw,     setSearchKw]     = useState("");
  const [selected,     setSelected]     = useState<Set<number>>(new Set()); // indices into filtered
  const [status,       setStatus]       = useState("尚未載入資料");
  const [generating,   setGenerating]   = useState(false);
  const [progress,     setProgress]     = useState({ done: 0, total: 0 });
  const [sortKey,      setSortKey]      = useState<SortKey>(null);
  const [sortAsc,      setSortAsc]      = useState(true);
  const [rangeMode,    setRangeMode]    = useState<"all" | "filtered" | "range">("all");
  const [rangeFrom,    setRangeFrom]    = useState("1");
  const [rangeTo,      setRangeTo]      = useState("10");

  const excelInputRef    = useRef<HTMLInputElement>(null);
  const templateInputRef = useRef<HTMLInputElement>(null);

  // ── 搜尋過濾 ──────────────────────────────────────────────────────────────
  const applyFilter = useCallback((kw: string, recs: PatientRecord[]) => {
    const k = kw.toLowerCase();
    if (!k) return recs;
    return recs.filter(r =>
      r.name.toLowerCase().includes(k) ||
      r.date.toLowerCase().includes(k) ||
      r.msItems.toLowerCase().includes(k) ||
      r.gender.includes(k)
    );
  }, []);

  // ── 匯入 Excel ──────────────────────────────────────────────────────────
  const handleExcelFile = useCallback(async (file: File) => {
    setExcelFile(file);
    setStatus("載入中...");
    try {
      const buf = await file.arrayBuffer();
      const recs = parseExcel(buf);
      setRecords(recs);
      const f = applyFilter(searchKw, recs);
      setFiltered(f);
      setSelected(new Set());
      setStatus(`✅  已載入 ${recs.length} 筆資料  ·  ${file.name}`);
    } catch (e: unknown) {
      setStatus(`❌  載入失敗：${String(e)}`);
    }
  }, [searchKw, applyFilter]);

  // ── 搜尋 ─────────────────────────────────────────────────────────────────
  const handleSearch = useCallback((kw: string) => {
    setSearchKw(kw);
    setFiltered(applyFilter(kw, records));
    setSelected(new Set());
  }, [records, applyFilter]);

  // ── 顯示全部 / 只顯示代謝症候群 ─────────────────────────────────────────
  const showAll = () => {
    setFiltered(records);
    setSearchKw("");
    setSelected(new Set());
    setStatus(`顯示全部 ${records.length} 筆`);
  };
  const showMetabolic = () => {
    const f = records.filter(r => r.metabolic !== "無");
    setFiltered(f);
    setSelected(new Set());
    setStatus(`⚠️  代謝症候群個案：${f.length} 筆`);
  };
  const showHepBC = () => {
    const f = records.filter(r =>
      r.rawData["hbsag"] === "陽性" || r.rawData["hcv"] === "陽性"
    );
    setFiltered(f);
    setSelected(new Set());
    setStatus(`🅱️  B/C肝陽性個案：${f.length} 筆`);
  };

  // ── 排序 ─────────────────────────────────────────────────────────────────
  const handleSort = (key: SortKey) => {
    const newAsc = sortKey === key ? !sortAsc : true;
    setSortKey(key);
    setSortAsc(newAsc);
    if (!key) return;
    const sorted = [...filtered].sort((a, b) => {
      let av = (a as Record<string, unknown>)[key] as string ?? "";
      let bv = (b as Record<string, unknown>)[key] as string ?? "";
      const an = parseFloat(av), bn = parseFloat(bv);
      const cmp = (!isNaN(an) && !isNaN(bn)) ? an - bn : av.localeCompare(bv, "zh-TW");
      return newAsc ? cmp : -cmp;
    });
    setFiltered(sorted);
  };

  // ── 選取行 ────────────────────────────────────────────────────────────────
  const toggleSelect = (idx: number, multi = false) => {
    setSelected(prev => {
      const next = new Set(prev);
      if (!multi) {
        if (next.has(idx) && next.size === 1) { next.clear(); }
        else { next.clear(); next.add(idx); }
      } else {
        next.has(idx) ? next.delete(idx) : next.add(idx);
      }
      return next;
    });
  };
  const selectAll = () => setSelected(new Set(filtered.map((_, i) => i)));
  const clearSel  = () => setSelected(new Set());

  // ── 決定要產生的 records ────────────────────────────────────────────────
  const getTargetRecords = (): PatientRecord[] => {
    if (rangeMode === "all") return records;
    if (rangeMode === "filtered") return filtered;
    // range
    const f = parseInt(rangeFrom);
    const t = parseInt(rangeTo);
    if (isNaN(f) || isNaN(t)) { alert("請輸入正確的起訖筆數"); return []; }
    return records.slice(f - 1, t);
  };

  // ── 產生報告 ─────────────────────────────────────────────────────────────
  const generateReports = async (targetRecords: PatientRecord[]) => {
    if (!templateFile) { alert("請先選擇報告範本 (.docx)"); return; }
    if (targetRecords.length === 0) { alert("沒有可產生的個案"); return; }

    setGenerating(true);
    setProgress({ done: 0, total: targetRecords.length });

    try {
      const templateBuf = await templateFile.arrayBuffer();
      const results: { name: string; blob: Blob }[] = [];
      const errors: string[] = [];

      for (let i = 0; i < targetRecords.length; i++) {
        const rec = targetRecords[i];
        try {
          const blob = await fillReport(templateBuf, rec.rawData);
          results.push({ name: makeFilename(rec.rawData), blob });
        } catch (e) {
          errors.push(`${rec.name}：${String(e)}`);
        }
        setProgress({ done: i + 1, total: targetRecords.length });
      }

      if (results.length === 1) {
        saveAs(results[0].blob, results[0].name);
      } else if (results.length > 1) {
        const zip = new JSZip();
        for (const { name, blob } of results) {
          zip.file(name, blob);
        }
        const zipBlob = await zip.generateAsync({ type: "blob" });
        saveAs(zipBlob, "健康報告.zip");
      }

      if (errors.length > 0) {
        setStatus(`⚠️  完成 ${results.length} 份，${errors.length} 份失敗`);
        alert("部分失敗：\n" + errors.slice(0, 10).join("\n"));
      } else {
        setStatus(`✅  成功產生 ${results.length} 份報告`);
      }
    } catch (e) {
      setStatus(`❌  產生失敗：${String(e)}`);
    } finally {
      setGenerating(false);
    }
  };

  const handleGenerateAll = () => {
    if (selected.size > 0) {
      // 有勾選個案 → 依勾選產生
      const sel = [...selected].map(i => filtered[i]).filter(Boolean);
      generateReports(sel);
    } else {
      // 無勾選 → 依左側產生範圍設定
      generateReports(getTargetRecords());
    }
  };
  const handleGenerateSelected = () => {
    const sel = [...selected].map(i => filtered[i]).filter(Boolean);
    if (sel.length === 0) { alert("請先選取要產生報告的個案（可多選）"); return; }
    generateReports(sel);
  };

  // ── 樣式工具 ─────────────────────────────────────────────────────────────
  const btn = (extra: React.CSSProperties = {}): React.CSSProperties => ({
    background: C.bgPanel, color: C.textMain, border: `1px solid ${C.border}`,
    borderRadius: 6, padding: "6px 14px", cursor: "pointer",
    fontFamily: "Microsoft JhengHei UI, sans-serif", fontSize: 13,
    transition: "background 0.15s",
    ...extra,
  });
  const btnAccent: React.CSSProperties = { ...btn(), background: C.accent, color: C.textHead, fontWeight: "bold" };

  // ── 渲染 ─────────────────────────────────────────────────────────────────
  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh",
                  background: C.bgDark, color: C.textMain,
                  fontFamily: "Microsoft JhengHei UI, sans-serif", fontSize: 13 }}>

      {/* ── 工具列 ────────────────────────────────────────────────────────── */}
      <div style={{ background: C.bgPanel, padding: "8px 16px",
                    display: "flex", alignItems: "center", gap: 8,
                    borderBottom: `1px solid ${C.border}` }}>
        <span style={{ fontSize: 18, marginRight: 4 }}>🏥</span>
        <span style={{ fontWeight: "bold", fontSize: 14, color: C.textHead, marginRight: 16 }}>
          成人健檢報告產生程式
        </span>

        {/* 隱藏的 file input */}
        <input ref={excelInputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
          onChange={e => { const f = e.target.files?.[0]; if (f) handleExcelFile(f); e.target.value = ""; }} />
        <input ref={templateInputRef} type="file" accept=".docx" style={{ display: "none" }}
          onChange={e => { const f = e.target.files?.[0]; if (f) { setTemplateFile(f); setStatus(`範本：${f.name}`); } e.target.value = ""; }} />

        <button style={btn()} onClick={() => excelInputRef.current?.click()}>📂 匯入 Excel</button>
        <button style={btn()} onClick={() => templateInputRef.current?.click()}>📋 選擇範本</button>
        <button style={btn()} onClick={showAll}>👁 全部個案</button>
        <button style={btn({ color: C.warn })} onClick={showMetabolic}>⚠️ 代謝症候群</button>
        <button style={btn({ color: C.accent2 })} onClick={showHepBC}>🦠 B/C肝陽性</button>
        <div style={{ flex: 1 }} />
        <button style={btn()} onClick={handleGenerateSelected} disabled={generating}>
          ✔ 產生選取
        </button>
        <button style={btnAccent} onClick={handleGenerateAll} disabled={generating}>
          {generating ? `⏳ ${progress.done}/${progress.total}` : "▶ 產生報告"}
        </button>
      </div>

      {/* ── 主體 ──────────────────────────────────────────────────────────── */}
      <div style={{ display: "flex", flex: 1, overflow: "hidden", padding: "8px 12px", gap: 8 }}>

        {/* ── 側欄 ────────────────────────────────────────────────────────── */}
        <div style={{ width: 240, background: C.bgPanel, borderRadius: 8,
                      padding: "12px 14px", display: "flex", flexDirection: "column",
                      gap: 8, border: `1px solid ${C.border}`, flexShrink: 0 }}>

          <SideSection title="📄 資料來源" accent={C.accent} border={C.border}>
            <FilePathRow label="Excel 檔案" name={excelFile?.name}
              onClick={() => excelInputRef.current?.click()} accent={C.accent} />
            <FilePathRow label="報告範本 (.docx)" name={templateFile?.name}
              onClick={() => templateInputRef.current?.click()} accent={C.accent} />
          </SideSection>

          <SideSection title="⚙️ 產生範圍" accent={C.accent} border={C.border}>
            {(["all", "filtered", "range"] as const).map(val => (
              <label key={val} style={{ display: "flex", alignItems: "center", gap: 6,
                                        cursor: "pointer", color: C.textMain }}>
                <input type="radio" checked={rangeMode === val}
                  onChange={() => setRangeMode(val)}
                  style={{ accentColor: C.accent }} />
                {{ all: "全部個案", filtered: "目前篩選結果", range: "指定範圍" }[val]}
              </label>
            ))}
            {rangeMode === "range" && (
              <div style={{ display: "flex", alignItems: "center", gap: 4, marginTop: 4 }}>
                <span style={{ color: C.textDim }}>第</span>
                <input value={rangeFrom} onChange={e => setRangeFrom(e.target.value)}
                  style={{ width: 44, background: C.bgRowA, color: C.textMain,
                            border: `1px solid ${C.border}`, borderRadius: 4,
                            padding: "2px 4px", textAlign: "center" }} />
                <span style={{ color: C.textDim }}>至</span>
                <input value={rangeTo} onChange={e => setRangeTo(e.target.value)}
                  style={{ width: 44, background: C.bgRowA, color: C.textMain,
                            border: `1px solid ${C.border}`, borderRadius: 4,
                            padding: "2px 4px", textAlign: "center" }} />
                <span style={{ color: C.textDim }}>筆</span>
              </div>
            )}
          </SideSection>

          <SideSection title="🔵 多選操作" accent={C.accent} border={C.border}>
            <button style={{ ...btn(), width: "100%", textAlign: "left" }}
              onClick={selectAll}>全選目前頁面</button>
            <button style={{ ...btn(), width: "100%", textAlign: "left", marginTop: 4 }}
              onClick={clearSel}>取消選取</button>
            <div style={{ color: C.accent2, fontSize: 12, marginTop: 4 }}>
              已選 {selected.size} 筆
            </div>
          </SideSection>

          <div style={{ flex: 1 }} />
          <div style={{ color: C.textDim, fontSize: 11, borderTop: `1px solid ${C.border}`, paddingTop: 8 }}>
            版本 v1.0 · 全瀏覽器執行
          </div>
        </div>

        {/* ── 右側資料表 ──────────────────────────────────────────────────── */}
        <div style={{ flex: 1, display: "flex", flexDirection: "column",
                      background: C.bgPanel, borderRadius: 8,
                      border: `1px solid ${C.border}`, overflow: "hidden" }}>

          {/* 搜尋列 */}
          <div style={{ display: "flex", alignItems: "center", gap: 8,
                        padding: "8px 12px", borderBottom: `1px solid ${C.border}` }}>
            <span style={{ color: C.textDim }}>🔍</span>
            <input value={searchKw} onChange={e => handleSearch(e.target.value)}
              placeholder="搜尋姓名 / 日期 / 風險..."
              style={{ flex: 1, background: C.bgRowA, color: C.textMain,
                        border: `1px solid ${C.border}`, borderRadius: 6,
                        padding: "6px 10px", fontSize: 13,
                        outline: "none" }} />
            <span style={{ color: C.accent2, fontWeight: "bold", fontSize: 12 }}>
              共 {filtered.length} 筆
            </span>
          </div>

          {/* 進度條 */}
          {generating && (
            <div style={{ height: 4, background: C.border }}>
              <div style={{ height: "100%", background: C.accent,
                            width: `${(progress.done / progress.total) * 100}%`,
                            transition: "width 0.2s" }} />
            </div>
          )}

          {/* 表格 */}
          <div style={{ flex: 1, overflow: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead>
                <tr style={{ background: C.bgPanel, position: "sticky", top: 0, zIndex: 1 }}>
                  <th style={{ width: 32, padding: "8px 4px", color: C.accent,
                                borderBottom: `1px solid ${C.border}` }}>
                    <input type="checkbox"
                      checked={filtered.length > 0 && selected.size === filtered.length}
                      onChange={e => e.target.checked ? selectAll() : clearSel()}
                      style={{ accentColor: C.accent }} />
                  </th>
                  {COLS.map(col => (
                    <th key={col.key}
                      style={{ padding: "8px 6px", color: C.accent, textAlign: "center",
                                borderBottom: `1px solid ${C.border}`, minWidth: col.w,
                                cursor: col.key !== "seq" && col.key !== "msItems" ? "pointer" : "default",
                                userSelect: "none" }}
                      onClick={() => col.key !== "seq" && col.key !== "msItems" && handleSort(col.key as SortKey)}>
                      {col.label}
                      {sortKey === col.key && <span style={{ marginLeft: 4 }}>{sortAsc ? "▲" : "▼"}</span>}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.length === 0 && (
                  <tr>
                    <td colSpan={COLS.length + 1}
                      style={{ textAlign: "center", padding: "40px", color: C.textDim }}>
                      {records.length === 0 ? "請先匯入 Excel 檔案" : "沒有符合條件的個案"}
                    </td>
                  </tr>
                )}
                {filtered.map((rec, i) => {
                  const isSel = selected.has(i);
                  const isMs  = rec.metabolic !== "無";
                  const rowBg = isSel ? C.accent + "33" : isMs ? C.msBg : i % 2 === 0 ? C.bgRowA : C.bgRowB;
                  return (
                    <tr key={`${rec._index}-${i}`}
                      style={{ background: rowBg, cursor: "pointer",
                                color: isMs && !isSel ? C.msFg : C.textMain }}
                      onClick={e => toggleSelect(i, e.ctrlKey || e.metaKey || e.shiftKey)}
                      onDoubleClick={() => generateReports([rec])}>
                      <td style={{ textAlign: "center", padding: "6px 4px" }}>
                        <input type="checkbox" checked={isSel}
                          onChange={() => toggleSelect(i, true)}
                          onClick={e => e.stopPropagation()}
                          style={{ accentColor: C.accent }} />
                      </td>
                      <td style={{ textAlign: "center", padding: "6px 4px", color: C.textDim }}>{i + 1}</td>
                      <td style={{ textAlign: "center", padding: "6px 4px" }}>{rec.date}</td>
                      <td style={{ textAlign: "center", padding: "6px 4px", fontWeight: "bold" }}>{rec.name}</td>
                      <td style={{ textAlign: "center", padding: "6px 4px" }}>{rec.gender}</td>
                      <td style={{ textAlign: "center", padding: "6px 4px" }}>{rec.age}</td>
                      <td style={{ textAlign: "center", padding: "6px 4px",
                                    color: rec.bmiOk ? C.textMain : C.warn }}>{rec.bmi}</td>
                      <td style={{ textAlign: "center", padding: "6px 4px",
                                    color: rec.bpOk ? C.textMain : C.warn }}>{rec.bp}</td>
                      <td style={{ textAlign: "center", padding: "6px 4px",
                                    color: rec.glucOk ? C.textMain : C.warn }}>{rec.glucose}</td>
                      <td style={{ textAlign: "center", padding: "6px 4px",
                                    color: isMs ? C.msFg : C.textMain,
                                    fontWeight: isMs ? "bold" : "normal" }}>{rec.metabolic}</td>
                      <td style={{ padding: "6px 8px", textAlign: "left" }}>{rec.msItems}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* ── 狀態列 ────────────────────────────────────────────────────────── */}
      <div style={{ background: C.bgPanel, borderTop: `1px solid ${C.border}`,
                    padding: "4px 16px", display: "flex", justifyContent: "space-between",
                    alignItems: "center", fontSize: 12 }}>
        <span style={{ color: C.textDim }}>{status}</span>
        <span style={{ color: C.border }}>成人健檢報告產生程式 v1.0</span>
      </div>
    </div>
  );
}

// ── 小工具元件 ────────────────────────────────────────────────────────────────
function SideSection({
  title, accent, border, children
}: { title: string; accent: string; border: string; children: React.ReactNode }) {
  return (
    <div>
      <div style={{ color: accent, fontWeight: "bold", fontSize: 11, marginBottom: 4 }}>{title}</div>
      <div style={{ height: 1, background: border, marginBottom: 8 }} />
      {children}
    </div>
  );
}

function FilePathRow({
  label, name, onClick, accent
}: { label: string; name?: string; onClick: () => void; accent: string }) {
  return (
    <div style={{ marginBottom: 8 }}>
      <div style={{ color: "#8A94B0", fontSize: 11, marginBottom: 3 }}>{label}</div>
      <div style={{ display: "flex", gap: 4 }}>
        <div style={{ flex: 1, background: "#2A3147", border: "1px solid #353D52",
                      borderRadius: 4, padding: "3px 6px", fontSize: 11,
                      color: name ? "#E8EAF0" : "#8A94B0",
                      overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
          {name ?? "（未選擇）"}
        </div>
        <button onClick={onClick}
          style={{ background: "#353D52", color: "#E8EAF0", border: "none",
                    borderRadius: 4, padding: "3px 8px", cursor: "pointer",
                    fontSize: 13 }}>…</button>
      </div>
    </div>
  );
}
