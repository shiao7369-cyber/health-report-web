/**
 * 健檢報告業務邏輯 — 移植自 Python 版 generate_reports.py
 */

export type RowData = Record<string, string>;

// ── 欄位對應 (Excel 標題 → JS 變數名) ─────────────────────────────────────────
export const COL_MAP: Record<string, string> = {
  "體檢日": "exam_date",
  "健檢號碼": "id",
  "姓名": "name",
  "性別": "gender",
  "年齡": "age",
  "生日": "birthday",
  "高血壓": "hypertension",
  "糖尿病": "diabetes",
  "高血脂症": "hyperlipidemia",
  "心臟病": "heart_disease",
  "腎臟病": "kidney_disease",
  "腦中風": "stroke",
  "吸菸": "smoking",
  "喝酒": "drinking",
  "嚼檳榔": "betel_nut",
  "運動": "exercise",
  "憂鬱檢測1": "depression1",
  "憂鬱檢測2": "depression2",
  "身高": "height",
  "體重": "weight",
  "腰圍": "waist",
  "身體質量指數": "bmi",
  "脈搏": "pulse",
  "收縮壓": "sbp",
  "舒張壓": "dbp",
  "蛋白質": "urine_protein",
  "膽固醇": "cholesterol",
  "三酸甘油脂": "triglyceride",
  "ASR(GOT)": "got",
  "ALT(GPT)": "gpt",
  "尿素氮": "bun",
  "肌酐酸": "creatinine",
  "血糖": "glucose",
  "尿酸": "uric_acid",
  "高密度膽固醇": "hdl",
  "低密度膽固醇": "ldl",
  "腎絲球過濾率": "egfr",
  "B型肝炎表面抗原": "hbsag",
  "C型肝炎抗體": "hcv",
  "戒菸": "counsel_quit_smoke",
  "戒酒": "counsel_quit_alcohol",
  "戒檳榔": "counsel_quit_betel",
  "事故傷害預防": "counsel_accident",
  "口腔保健": "counsel_oral",
  "體重控制": "counsel_weight",
  "飲食與營養": "counsel_diet",
  "規律運動": "counsel_exercise_old",
  "維持正常體重": "counsel_maintain_weight",
  "健康飲食": "counsel_healthy_diet",
  "規律運動(含150分鐘/每週)": "counsel_exercise",
  "健康飲食(含我的健康餐盤)": "counsel_healthy_meal",
  "慢性疾病-風險評估": "counsel_chronic",
  "慢性疾病風險值-冠心病": "risk_cad",
  "慢性疾病風險值-糖尿病": "risk_dm",
  "慢性疾病風險值-高血壓": "risk_htn",
  "慢性疾病風險值-腦中風": "risk_stroke",
  "慢性疾病風險值-血管不良事件": "risk_cv",
  "腎功能檢查期別": "kidney_stage",
  "腎病識能": "counsel_kidney",
  "慢性疾病風險評估": "counsel_chronic2",
  "全身理學檢查": "counsel_physical",
  "報告解說": "counsel_report",
  "曾於成健B,C肝檢查": "prev_hep",
};

// ── 安全取值 ─────────────────────────────────────────────────────────────────
export function v(row: RowData, key: string, def = ""): string {
  const val = row[key];
  if (val === undefined || val === null) return def;
  return String(val).trim();
}

function toFloat(s: string): number | null {
  if (!s || s === "-" || s === "") return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

// ── 判斷正常/異常邏輯 ────────────────────────────────────────────────────────

export function classifyBP(sbp: string, dbp: string): "normal" | "elevated" | "stage1" | "stage2" {
  const s = toFloat(sbp);
  const d = toFloat(dbp);
  if (s === null || d === null) return "normal";
  if (s >= 140 || d >= 90) return "stage2";
  if (s >= 130 || d >= 80) return "stage1";
  if (s >= 120 && s <= 129 && d < 80) return "elevated";
  return "normal";
}

export function classifyGlucose(glucose: string): "normal" | "prediab" | "suspect" | "high" {
  const g = toFloat(glucose);
  if (g === null) return "normal";
  if (g >= 200) return "high";
  if (g >= 126) return "suspect";
  if (g >= 100) return "prediab";
  return "normal";
}

export function classifyLipid(
  chol: string, tri: string, hdl: string, ldl: string, gender: string
): "normal" | "mild" | "high" {
  const tc = toFloat(chol);
  const tg = toFloat(tri);
  const hdlV = toFloat(hdl);
  const ldlV = toFloat(ldl);

  if (tc !== null && tc >= 240) return "high";
  if (ldlV !== null && ldlV >= 160) return "high";
  if (tg !== null && tg >= 200) return "high";
  if (tc !== null && tc >= 200) return "mild";
  if (ldlV !== null && ldlV >= 130) return "mild";
  if (tg !== null && tg >= 150) return "mild";
  if (hdlV !== null) {
    if (gender === "男" && hdlV < 40) return "mild";
    if (gender === "女" && hdlV < 50) return "mild";
  }
  return "normal";
}

export function classifyLiver(
  got: string, gpt: string, hbsag = "", hcv = ""
): "normal" | "mild" | "high" | "severe" {
  const g = toFloat(got) ?? 0;
  const gp = toFloat(gpt) ?? 0;
  const mx = Math.max(g, gp);
  const hasHep = hbsag.trim() === "陽性" || hcv.trim() === "陽性";
  if (mx <= 40) return "normal";
  if (mx > 200) return "severe";
  if (mx > 80 || hasHep) return "high";
  return "mild";
}

export function hasProteinuria(urineProtein: string): boolean {
  const val = String(urineProtein).trim();
  return !["−", "-", "", "None", "陰性"].includes(val);
}

export function classifyKidney(
  egfr: string, urineProtein: string
): "normal" | "stage1" | "stage2" | "stage3" | "stage4" {
  const e = toFloat(egfr);
  const protein = hasProteinuria(urineProtein);
  if (e === null) return protein ? "stage1" : "normal";
  if (e < 30) return "stage4";
  if (e < 60) return "stage3";
  if (e < 90) return protein ? "stage2" : "normal";
  return protein ? "stage1" : "normal";
}

export function classifyMetabolic(
  gender: string, waist: string, sbp: string, dbp: string,
  glucose: string, triglyceride: string, hdl: string
): "normal" | "mild" | "moderate" | "severe" {
  const abnormal: string[] = [];

  const w = toFloat(waist);
  if (w !== null && ((gender === "男" && w >= 90) || (gender === "女" && w >= 80)))
    abnormal.push("waist");

  const s = toFloat(sbp) ?? 0;
  const d = toFloat(dbp) ?? 0;
  if (s >= 130 || d >= 85) abnormal.push("bp");

  const g = toFloat(glucose);
  if (g !== null && g >= 100) abnormal.push("glucose");

  const tg = toFloat(triglyceride);
  if (tg !== null && tg >= 150) abnormal.push("tg");

  const hdlV = toFloat(hdl);
  if (hdlV !== null && ((gender === "男" && hdlV < 40) || (gender === "女" && hdlV < 50)))
    abnormal.push("hdl");

  const count = abnormal.length;
  if (count < 3) return "normal";

  const bpSevere = s >= 160 || d >= 100;
  const glucoseSevere = (toFloat(glucose) ?? 0) >= 126;

  if (count === 5 || bpSevere || glucoseSevere) return "severe";
  if (count >= 4) return "moderate";
  return "mild";
}

// ── check_mark ───────────────────────────────────────────────────────────────
function checkMark(value: string): string {
  return ["是", "有", "Yes", "Y"].includes(value.trim()) ? "■" : "□";
}

// ── 健康諮詢文字 ─────────────────────────────────────────────────────────────
export function buildCounselLine1(row: RowData): string {
  const quitSmoke = checkMark(v(row, "counsel_quit_smoke"));
  const quitAlc   = checkMark(v(row, "counsel_quit_alcohol"));
  const quitBet   = checkMark(v(row, "counsel_quit_betel"));

  const bmiF = toFloat(v(row, "bmi"));
  const abnormalBmi = bmiF !== null && (bmiF < 18.5 || bmiF >= 24);
  const maintainWtExcel = checkMark(v(row, "counsel_maintain_weight"));
  const maintainWt = (abnormalBmi || maintainWtExcel === "■") ? "■" : "□";

  const exerciseCounsel = v(row, "counsel_exercise");
  const exerciseHabit   = v(row, "exercise");
  const needExercise = exerciseCounsel === "是" || exerciseHabit.includes("沒有") || exerciseHabit.includes("未達");
  const exercise = needExercise ? "■" : "□";

  return `${quitSmoke}戒菸${quitAlc}節酒         ${quitBet}戒檳榔        ${exercise}規律運動(含150分鐘/每週)        ${maintainWt}維持正常體重`;
}

export function buildCounselLine2(row: RowData): string {
  const healthyMeal = checkMark(v(row, "counsel_healthy_meal"));
  const accident    = checkMark(v(row, "counsel_accident"));
  const oral        = checkMark(v(row, "counsel_oral"));
  return `${healthyMeal}健康飲食(含我的健康餐盤)           ${accident}事故傷害預防                     ${oral}口腔保健`;
}

export function buildCounselKidney(row: RowData): string {
  const kidney = checkMark(v(row, "counsel_kidney"));
  return `${kidney}腎病識能衛教指導(含尿蛋白、eGFR的數據、腎功能期別及其嚴重度、危險因子衛教)`;
}

// ── 檢查結果文字 ─────────────────────────────────────────────────────────────
export function buildBpLine(row: RowData): string {
  const grade = classifyBP(v(row, "sbp"), v(row, "dbp"));
  switch (grade) {
    case "normal":   return "血  壓：■正常□異常：建議□生活型態改善，並定期＿＿個月追蹤□進一步檢查□接受治療";
    case "elevated": return "血  壓：□正常■異常：建議■生活型態改善，並定期＿＿個月追蹤□進一步檢查□接受治療";
    case "stage1":   return "血  壓：□正常■異常：建議■生活型態改善，並定期３個月追蹤□進一步檢查□接受治療";
    case "stage2":   return "血  壓：□正常■異常：建議□生活型態改善，並定期＿＿個月追蹤■進一步檢查■接受治療";
  }
}

export function buildGlucoseLine(row: RowData): string {
  const grade = classifyGlucose(v(row, "glucose"));
  switch (grade) {
    case "normal":  return "飯前血糖：■正常□異常：建議□生活型態改善，並定期＿＿個月追蹤□進一步檢查□接受治療";
    case "prediab": return "飯前血糖：□正常■異常：建議■生活型態改善，並定期3~6個月追蹤□進一步檢查□接受治療";
    case "suspect": return "飯前血糖：□正常■異常：建議□生活型態改善，並定期＿＿個月追蹤■進一步檢查□接受治療";
    case "high":    return "飯前血糖：□正常■異常：建議□生活型態改善，並定期＿＿個月追蹤□進一步檢查■接受治療";
  }
}

export function buildLipidLine(row: RowData): string {
  const grade = classifyLipid(v(row, "cholesterol"), v(row, "triglyceride"), v(row, "hdl"), v(row, "ldl"), v(row, "gender"));
  switch (grade) {
    case "normal": return "血脂肪：■正常□異常：建議□生活型態改善，並定期＿＿個月追蹤□進一步檢查□接受治療";
    case "mild":   return "血脂肪：□正常■異常：建議■生活型態改善，並定期3~6個月追蹤□進一步檢查□接受治療";
    case "high":   return "血脂肪：□正常■異常：建議■生活型態改善，並定期1~3個月追蹤■進一步檢查□接受治療";
  }
}

export function buildKidneyLine(row: RowData): string {
  const egfr         = v(row, "egfr");
  const urineProtein = v(row, "urine_protein");
  const stageRaw     = v(row, "kidney_stage");

  let stageLabel = "";
  const m = stageRaw.match(/第\s*([\w]+)\s*期/);
  if (m) {
    stageLabel = `第${m[1]}期`;
  } else if (stageRaw.includes("正常") || stageRaw.startsWith("0(")) {
    stageLabel = "正常";
  } else if (stageRaw.includes("暫時無法判定")) {
    stageLabel = "待確認";
  } else {
    stageLabel = stageRaw.slice(0, 8);
  }

  const grade = classifyKidney(egfr, urineProtein);
  switch (grade) {
    case "normal": return `腎功能：■正常□異常：期別${stageLabel}建議□生活型態改善，並定期＿個月追蹤□進一步檢查□接受治療`;
    case "stage1": return `腎功能：□正常■異常：期別${stageLabel}建議■生活型態改善，並定期6個月追蹤□進一步檢查□接受治療`;
    case "stage2": return `腎功能：□正常■異常：期別${stageLabel}建議■生活型態改善，並定期3~6個月追蹤□進一步檢查□接受治療`;
    case "stage3": return `腎功能：□正常■異常：期別${stageLabel}建議□生活型態改善，並定期＿個月追蹤■進一步檢查□接受治療`;
    case "stage4": return `腎功能：□正常■異常：期別${stageLabel}建議□生活型態改善，並定期＿個月追蹤■進一步檢查■接受治療`;
  }
}

export function buildLiverLine(row: RowData): string {
  const grade = classifyLiver(v(row, "got"), v(row, "gpt"), v(row, "hbsag"), v(row, "hcv"));
  switch (grade) {
    case "normal": return "肝功能：■正常□異常：建議□生活型態改善，並定期＿＿個月追蹤□進一步檢查□接受治療";
    case "mild":   return "肝功能：□正常■異常：建議■生活型態改善，並定期3~6個月追蹤□進一步檢查□接受治療";
    case "high":   return "肝功能：□正常■異常：建議□生活型態改善，並定期＿＿個月追蹤■進一步檢查□接受治療";
    case "severe": return "肝功能：□正常■異常：建議□生活型態改善，並定期＿＿個月追蹤□進一步檢查■接受治療";
  }
}

export function buildMetabolicLine(row: RowData): string {
  const grade = classifyMetabolic(
    v(row, "gender"), v(row, "waist"), v(row, "sbp"), v(row, "dbp"),
    v(row, "glucose"), v(row, "triglyceride"), v(row, "hdl")
  );
  switch (grade) {
    case "normal":   return "代謝症候群：■沒有□有：建議□生活型態改善，並定期＿＿個月追蹤□進一步檢查□接受治療";
    case "mild":     return "代謝症候群：□沒有■有：建議■生活型態改善，並定期6個月追蹤□進一步檢查□接受治療";
    case "moderate": return "代謝症候群：□沒有■有：建議■生活型態改善，並定期3個月追蹤□進一步檢查□接受治療";
    case "severe":   return "代謝症候群：□沒有■有：建議■生活型態改善，並定期3個月追蹤■進一步檢查□接受治療";
  }
}

export function buildHbsagLine(row: RowData): string {
  const val = v(row, "hbsag");
  if (val === "陰性") {
    return '"△"B型肝炎表面抗原：■陰性        □陽性        □進一步檢查        □接受治療';
  }
  if (val === "陽性") {
    const gotV = toFloat(v(row, "got")) ?? 0;
    const gptV = toFloat(v(row, "gpt")) ?? 0;
    const liverSevere = gotV > 80 || gptV > 80;
    if (liverSevere) {
      return '"△"B型肝炎表面抗原：□陰性        ■陽性        ■進一步檢查        ■接受治療';
    }
    return '"△"B型肝炎表面抗原：□陰性        ■陽性        ■進一步檢查        □接受治療';
  }
  return '"△"B型肝炎表面抗原：□陰性        □陽性        □進一步檢查        □接受治療';
}

export function buildHcvLine(row: RowData): string {
  const val = v(row, "hcv");
  const neg = val === "陰性" ? "■" : "□";
  const pos = val === "陽性" ? "■" : "□";
  return `"△"C型肝炎抗體        ：${neg}陰性        ${pos}陽性        □進一步檢查        □接受治療`;
}

export function buildDepressionLine(row: RowData): string {
  const d1 = v(row, "depression1");
  const d2 = v(row, "depression2");
  const bothNo = d1 === "否" && d2 === "否";
  const n = bothNo ? "■" : "□";
  const y = bothNo ? "□" : "■";
  return `憂鬱檢測：${n}二題皆答「否」${y}二題任一題答「是」，建議轉介至相關單位接受進一步服務`;
}

export function buildRiskLine(row: RowData): { line1: string; line2: string } {
  function riskLabel(raw: string): string {
    const m = raw.match(/(高風險|中風險|低風險)/);
    return m ? m[1] : raw.trim();
  }
  const cad  = riskLabel(v(row, "risk_cad"));
  const dm   = riskLabel(v(row, "risk_dm"));
  const htn  = riskLabel(v(row, "risk_htn"));
  const strk = riskLabel(v(row, "risk_stroke"));
  const cv   = riskLabel(v(row, "risk_cv"));
  return {
    line1: `慢性疾病風險值：冠心病(${cad})、糖尿病(${dm})、高血壓(${htn})、`,
    line2: `腦中風(${strk})、心血管不良事件(${cv})`,
  };
}
