/**
 * Excel 讀取器 — 使用 SheetJS (xlsx) 在瀏覽器端解析 .xlsx
 * 對應 Python 版 load_excel()
 */
import * as XLSX from "xlsx";
import { COL_MAP } from "./report-logic";
import type { RowData } from "./report-logic";

export interface PatientRecord {
  // 顯示用
  _index: number;         // 1-based，對應 Excel 第幾筆
  serialNo: string;       // 序號，如 202602-93001
  date: string;
  name: string;
  gender: string;
  age: string;
  bmi: string;
  bmiOk: boolean;
  bp: string;
  bpOk: boolean;
  glucose: string;
  glucOk: boolean;
  metabolic: string;      // "■ 有" | "無"
  msItems: string;        // "腰圍、血壓" 等
  // 完整 row 資料（供 docx 產生用）
  rawData: RowData;
}

// 序號對照表：病歷號(10位字串) → 序號(如 202602-93001)
export type SerialMapping = Map<string, string>;

/**
 * 自動偵測 Excel 類型
 * - "health"  : 健檢資料（含「健檢號碼」欄）
 * - "mapping" : 序號對照表（含「病歷號」欄）
 * - "unknown" : 無法判斷
 */
export function detectExcelType(buffer: ArrayBuffer): "health" | "mapping" | "unknown" {
  try {
    const wb = XLSX.read(buffer, { type: "array", cellText: false, cellDates: false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: "" }) as string[][];
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const row = (rows[i] as string[]).map(c => String(c).trim());
      if (row.includes("健檢號碼")) return "health";
      if (row.includes("病歷號"))   return "mapping";
    }
  } catch { /* ignore */ }
  return "unknown";
}

/**
 * 解析序號對照表（0308成健-V2.xlsx 格式）
 * 格式：第0列為標題、第1列為欄位標題（含「病歷號」），序號欄為無標題但含 YYYYMM-NNNNN 格式
 */
export function parseMappingExcel(buffer: ArrayBuffer): SerialMapping {
  const wb = XLSX.read(buffer, { type: "array", cellText: true, cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: "" }) as string[][];

  // 找到含「病歷號」的標題列
  let headerRowIdx = -1;
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    if ((rows[i] as string[]).some(cell => String(cell).trim() === "病歷號")) {
      headerRowIdx = i;
      break;
    }
  }
  if (headerRowIdx === -1) return new Map();

  const headers = (rows[headerRowIdx] as string[]).map(h => String(h ?? "").trim());
  const medIdIdx = headers.indexOf("病歷號");
  if (medIdIdx === -1) return new Map();

  // 找序號欄：掃描資料列，看哪欄符合 YYYYMM-NNNNN 格式
  const serialPattern = /^\d{6}-\d{5}$/;
  let serialIdx = -1;
  outer: for (let ri = headerRowIdx + 1; ri < Math.min(headerRowIdx + 6, rows.length); ri++) {
    const row = rows[ri] as string[];
    for (let ci = 0; ci < row.length; ci++) {
      if (serialPattern.test(String(row[ci]).trim())) {
        serialIdx = ci;
        break outer;
      }
    }
  }
  if (serialIdx === -1) return new Map();

  const mapping = new Map<string, string>();
  for (let ri = headerRowIdx + 1; ri < rows.length; ri++) {
    const row = rows[ri] as string[];
    const medIdRaw = String(row[medIdIdx] ?? "").trim().replace(/\D/g, "");
    const serial   = String(row[serialIdx] ?? "").trim();
    if (!medIdRaw || !serialPattern.test(serial)) continue;
    const medId = medIdRaw.padStart(10, "0");
    mapping.set(medId, serial);
  }
  return mapping;
}

/**
 * 解析健檢 Excel（115年社區成人健檢格式）
 * @param buffer      Excel ArrayBuffer
 * @param serialMapping 可選的序號對照表；若提供則以病歷號比對後填入 serialNo
 */
export function parseExcel(buffer: ArrayBuffer, serialMapping?: SerialMapping): PatientRecord[] {
  const wb = XLSX.read(buffer, { type: "array", cellText: true, cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: "" }) as string[][];

  if (rows.length < 2) return [];

  // 動態找到欄位標題列（含「姓名」的那一行）
  let headerRowIdx = 0;
  for (let i = 0; i < Math.min(10, rows.length); i++) {
    if ((rows[i] as string[]).some(cell => String(cell ?? "").trim() === "姓名")) {
      headerRowIdx = i;
      break;
    }
  }
  const headers = (rows[headerRowIdx] as string[]).map(h => String(h ?? "").trim());

  function idx(name: string): number {
    return headers.indexOf(name);
  }

  const nameI    = idx("姓名");
  const dateI    = idx("體檢日");
  const genderI  = idx("性別");
  const ageI     = idx("年齡");
  const bmiI     = idx("身體質量指數");
  const sbpI     = idx("收縮壓");
  const dbpI     = idx("舒張壓");
  const glucI    = idx("血糖");
  const waistI   = idx("腰圍");
  const tgI      = idx("三酸甘油脂");
  const hdlI     = idx("高密度膽固醇");
  const serialI  = idx("序號");
  const healthIdI = idx("健檢號碼");

  const records: PatientRecord[] = [];

  for (let ri = headerRowIdx + 1; ri < rows.length; ri++) {
    const row = rows[ri] as string[];
    const getCellStr = (i: number) =>
      i >= 0 && i < row.length && row[i] !== undefined && row[i] !== null
        ? String(row[i]).trim()
        : "";

    const nameVal = getCellStr(nameI);
    if (!nameVal) continue;

    // 完整 rawData（COL_MAP 轉換）
    const rawData: RowData = {};
    headers.forEach((h, i) => {
      const key = COL_MAP[h] ?? h;
      rawData[key] = getCellStr(i);
    });

    // 血壓
    const sbpStr = getCellStr(sbpI);
    const dbpStr = getCellStr(dbpI);
    let bpStr = sbpStr && dbpStr ? `${sbpStr}/${dbpStr}` : sbpStr;
    const sbpN = parseFloat(sbpStr);
    const dbpN = parseFloat(dbpStr);
    const bpOk = !isNaN(sbpN) && !isNaN(dbpN) ? sbpN < 130 && dbpN < 85 : true;

    // 血糖
    const glucStr = getCellStr(glucI);
    const glucN   = parseFloat(glucStr);
    const glucOk  = !isNaN(glucN) ? glucN < 100 : true;

    // BMI
    const bmiStr = getCellStr(bmiI);
    const bmiN   = parseFloat(bmiStr);
    const bmiOk  = !isNaN(bmiN) ? bmiN >= 18.5 && bmiN < 24 : true;
    const bmiDisplay = !isNaN(bmiN) ? bmiN.toFixed(1) : bmiStr;

    // 代謝症候群
    const gender = getCellStr(genderI);
    const msItems: string[] = [];
    const waistN = parseFloat(getCellStr(waistI));
    if (!isNaN(waistN) && ((gender === "男" && waistN >= 90) || (gender === "女" && waistN >= 80)))
      msItems.push("腰圍");
    if (!bpOk) msItems.push("血壓");
    if (!isNaN(glucN) && glucN >= 100) msItems.push("血糖");
    const tgN = parseFloat(getCellStr(tgI));
    if (!isNaN(tgN) && tgN >= 150) msItems.push("血脂");
    const hdlN = parseFloat(getCellStr(hdlI));
    if (!isNaN(hdlN) && ((gender === "男" && hdlN < 40) || (gender === "女" && hdlN < 50)))
      msItems.push("HDL");
    const ms = msItems.length >= 3;

    // 序號：優先讀 Excel 內的序號欄；若無則從對照表查病歷號
    let serialNo = getCellStr(serialI);
    if (!serialNo && serialMapping) {
      const rawId = getCellStr(healthIdI).replace(/\D/g, "").padStart(10, "0");
      serialNo = serialMapping.get(rawId) ?? "";
    }
    // 同步寫入 rawData 讓 docx 產生時可用
    if (serialNo) rawData["serial_no"] = serialNo;

    records.push({
      _index:   ri,
      serialNo,
      date:     getCellStr(dateI),
      name:     nameVal,
      gender,
      age:      getCellStr(ageI),
      bmi:      bmiDisplay,
      bmiOk,
      bp:       bpStr,
      bpOk,
      glucose:  glucStr,
      glucOk,
      metabolic: ms ? "■ 有" : "無",
      msItems:   msItems.length > 0 ? msItems.join("、") : "—",
      rawData,
    });
  }

  return records;
}
