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

export function parseExcel(buffer: ArrayBuffer): PatientRecord[] {
  const wb = XLSX.read(buffer, { type: "array", cellText: true, cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: "" }) as string[][];

  if (rows.length < 2) return [];

  const headers = (rows[0] as string[]).map(h => String(h ?? "").trim());

  function idx(name: string): number {
    return headers.indexOf(name);
  }

  const nameI   = idx("姓名");
  const dateI   = idx("體檢日");
  const genderI = idx("性別");
  const ageI    = idx("年齡");
  const bmiI    = idx("身體質量指數");
  const sbpI    = idx("收縮壓");
  const dbpI    = idx("舒張壓");
  const glucI   = idx("血糖");
  const waistI  = idx("腰圍");
  const tgI     = idx("三酸甘油脂");
  const hdlI    = idx("高密度膽固醇");

  const records: PatientRecord[] = [];

  for (let ri = 1; ri < rows.length; ri++) {
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
    const bpOk = !isNaN(sbpN) && !isNaN(dbpN) ? sbpN < 130 && dbpN < 80 : true;

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

    records.push({
      _index: ri,
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
