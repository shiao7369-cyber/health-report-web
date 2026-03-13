/**
 * docx 產生器 — 使用 JSZip 在瀏覽器端操作 .docx XML
 * 對應 Python 版 fill_report()
 */
import JSZip from "jszip";
import {
  v, buildBpLine, buildGlucoseLine, buildLipidLine, buildKidneyLine,
  buildLiverLine, buildMetabolicLine, buildCounselLine1, buildCounselLine2,
  buildCounselKidney, buildHbsagLine, buildHcvLine, buildDepressionLine,
  buildRiskLine,
} from "./report-logic";
import type { RowData } from "./report-logic";

const NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W  = (local: string) => `{${NS}}${local}`;

// ── XML 工具 ──────────────────────────────────────────────────────────────────

/** 取節點下所有 w:t 的文字合併 */
function allText(node: Element): string {
  return Array.from(node.getElementsByTagNameNS(NS, "t"))
    .map(t => t.textContent ?? "")
    .join("");
}

/** 清除段落中所有 w:r，在 firstPara 插入新文字 run */
function clearAndFill(paras: Element[], text: string, sz = "20") {
  for (const p of paras) {
    Array.from(p.getElementsByTagNameNS(NS, "r")).forEach(r => r.parentNode!.removeChild(r));
  }
  if (paras.length === 0) return;
  const p = paras[0];
  const doc = p.ownerDocument!;
  const r = doc.createElementNS(NS, "w:r");
  const rpr = doc.createElementNS(NS, "w:rPr");
  const fonts = doc.createElementNS(NS, "w:rFonts");
  fonts.setAttributeNS(NS, "w:ascii",    "標楷體");
  fonts.setAttributeNS(NS, "w:eastAsia", "標楷體");
  fonts.setAttributeNS(NS, "w:hAnsi",    "標楷體");
  const szEl = doc.createElementNS(NS, "w:sz");
  szEl.setAttributeNS(NS, "w:val", sz);
  rpr.appendChild(fonts);
  rpr.appendChild(szEl);
  r.appendChild(rpr);
  const t = doc.createElementNS(NS, "w:t");
  t.textContent = text;
  if (text && (text[0] === " " || text[text.length - 1] === " ")) {
    t.setAttribute("xml:space", "preserve");
  }
  r.appendChild(t);
  p.appendChild(r);
}

/** 找到含 keyword 的儲存格，清空並填入新文字 */
function replaceCellText(tr: Element, keyword: string, newText: string) {
  const tcs = Array.from(tr.getElementsByTagNameNS(NS, "tc"));
  for (const tc of tcs) {
    if (allText(tc).includes(keyword)) {
      const paras = Array.from(tc.getElementsByTagNameNS(NS, "p"));
      clearAndFill(paras, newText);
      return;
    }
  }
}

/** 找到含 keyword 的 w:t 節點，直接替換文字 */
function fillTextNode(tr: Element, keyword: string, newText: string) {
  const tNodes = Array.from(tr.getElementsByTagNameNS(NS, "t"));
  for (const t of tNodes) {
    if ((t.textContent ?? "").includes(keyword)) {
      t.textContent = newText;
      if (newText && (newText[0] === " " || newText[newText.length - 1] === " ")) {
        t.setAttribute("xml:space", "preserve");
      }
      return;
    }
  }
}

/** 取 tr 的儲存格文字陣列 */
function getCellTexts(tr: Element): string[] {
  return Array.from(tr.getElementsByTagNameNS(NS, "tc")).map(allText);
}

// ── 核心填報 ─────────────────────────────────────────────────────────────────

export async function fillReport(
  templateBuffer: ArrayBuffer,
  rowData: RowData
): Promise<Blob> {
  const zip = await JSZip.loadAsync(templateBuffer);
  const xmlStr = await zip.file("word/document.xml")!.async("string");

  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlStr, "text/xml");

  const trs = Array.from(doc.getElementsByTagNameNS(NS, "tr"));

  for (const tr of trs) {
    const cellTexts = getCellTexts(tr);
    const joined = cellTexts.join(" ");

    // ── 姓名 / 生日 ──────────────────────────────────────────────────────────
    if (cellTexts[0]?.includes("姓名")) {
      const tcs = Array.from(tr.getElementsByTagNameNS(NS, "tc"));
      if (tcs.length >= 2) {
        clearAndFill(Array.from(tcs[1].getElementsByTagNameNS(NS, "p")), v(rowData, "name"));
      }
      if (tcs.length >= 4) {
        clearAndFill(Array.from(tcs[3].getElementsByTagNameNS(NS, "p")), v(rowData, "birthday"));
      }
      continue;
    }

    // ── 健康諮詢行 ──────────────────────────────────────────────────────────
    if (joined.includes("□戒菸□節酒")) {
      fillTextNode(tr, "□戒菸□節酒", buildCounselLine1(rowData));
      continue;
    }
    if (joined.includes("□健康飲食(含我的健康餐盤)")) {
      fillTextNode(tr, "□健康飲食(含我的健康餐盤)", buildCounselLine2(rowData));
      continue;
    }
    if (joined.includes("慢性疾病風險評估") && !joined.includes("腎病識能") && !joined.includes("風險值")) {
      const tNodes = Array.from(tr.getElementsByTagNameNS(NS, "t"));
      for (const t of tNodes) {
        if ((t.textContent ?? "").includes("慢性疾病風險評估")) {
          t.textContent = "■慢性疾病風險評估";
        } else if ((t.textContent ?? "").trim() === "□") {
          t.textContent = "";
        }
      }
      continue;
    }
    if (joined.includes("腎病識能衛教指導") || joined.includes("□腎病識能")) {
      fillTextNode(tr, "腎病識能", buildCounselKidney(rowData));
      continue;
    }

    // ── 各項檢查 ────────────────────────────────────────────────────────────
    if (joined.includes("血  壓：")) {
      replaceCellText(tr, "血  壓：", buildBpLine(rowData));
      continue;
    }
    if (joined.includes("飯前血糖：")) {
      replaceCellText(tr, "飯前血糖：", buildGlucoseLine(rowData));
      continue;
    }
    if (joined.includes("血脂肪：")) {
      replaceCellText(tr, "血脂肪：", buildLipidLine(rowData));
      continue;
    }
    if (joined.includes("腎功能：")) {
      replaceCellText(tr, "腎功能：", buildKidneyLine(rowData));
      continue;
    }
    if (joined.includes("肝功能：")) {
      replaceCellText(tr, "肝功能：", buildLiverLine(rowData));
      continue;
    }
    if (joined.includes("代謝症候群：") && !joined.includes("定義")) {
      replaceCellText(tr, "代謝症候群：", buildMetabolicLine(rowData));
      continue;
    }

    // ── 慢性疾病風險值（雙段落）────────────────────────────────────────────
    if (joined.includes("慢性疾病風險值：冠心病(")) {
      const tcs = Array.from(tr.getElementsByTagNameNS(NS, "tc"));
      const tcRisk = tcs.length > 1 ? tcs[1] : tcs[0];
      const paras = Array.from(tcRisk.getElementsByTagNameNS(NS, "p"));
      const { line1, line2 } = buildRiskLine(rowData);

      function makeRun(text: string): Element {
        const rEl = doc.createElementNS(NS, "w:r");
        const rpr = doc.createElementNS(NS, "w:rPr");
        const fnt = doc.createElementNS(NS, "w:rFonts");
        fnt.setAttributeNS(NS, "w:ascii",    "標楷體");
        fnt.setAttributeNS(NS, "w:eastAsia", "標楷體");
        fnt.setAttributeNS(NS, "w:hAnsi",    "標楷體");
        const szEl = doc.createElementNS(NS, "w:sz");
        szEl.setAttributeNS(NS, "w:val", "20");
        rpr.appendChild(fnt);
        rpr.appendChild(szEl);
        rEl.appendChild(rpr);
        const tEl = doc.createElementNS(NS, "w:t");
        tEl.textContent = text;
        if (text && (text[0] === " " || text[text.length - 1] === " ")) {
          tEl.setAttribute("xml:space", "preserve");
        }
        rEl.appendChild(tEl);
        return rEl;
      }

      if (paras.length >= 1) {
        const p0 = paras[0];
        Array.from(p0.getElementsByTagNameNS(NS, "r")).forEach(r => r.parentNode!.removeChild(r));
        p0.appendChild(makeRun(line1));
      }
      if (paras.length >= 2) {
        const p1 = paras[1];
        Array.from(p1.getElementsByTagNameNS(NS, "r")).forEach(r => r.parentNode!.removeChild(r));
        p1.appendChild(makeRun(line2));
      }
      continue;
    }

    // ── B 型肝炎 ─────────────────────────────────────────────────────────────
    if (joined.includes("B型肝炎表面抗原")) {
      fillTextNode(tr, "B型肝炎表面抗原", buildHbsagLine(rowData));
      continue;
    }

    // ── C 型肝炎 ─────────────────────────────────────────────────────────────
    if (joined.includes("C型肝炎抗體")) {
      fillTextNode(tr, "C型肝炎抗體", buildHcvLine(rowData));
      continue;
    }

    // ── 咳嗽症狀（還原範本預設）────────────────────────────────────────────
    if (joined.includes("咳嗽症狀：")) {
      const tNodes = Array.from(tr.getElementsByTagNameNS(NS, "t"));
      for (const t of tNodes) {
        if ((t.textContent ?? "").includes("咳嗽症狀：")) {
          t.textContent = (t.textContent ?? "").replace("☑", "□");
        }
      }
      continue;
    }

    // ── 憂鬱檢測 ────────────────────────────────────────────────────────────
    if (joined.includes("憂鬱檢測：")) {
      fillTextNode(tr, "憂鬱檢測：", buildDepressionLine(rowData));
      continue;
    }
  }

  // 序列化回 XML 並存入 ZIP
  const serializer = new XMLSerializer();
  const newXml = serializer.serializeToString(doc);
  zip.file("word/document.xml", newXml);

  return zip.generateAsync({ type: "blob" });
}

/** 產生檔名：yyyymmdd_姓名_健康報告.docx */
export function makeFilename(rowData: RowData): string {
  const name = v(rowData, "name") || "unnamed";
  const date = (v(rowData, "exam_date") || "").replace(/[/\-]/g, "");
  return `${date}_${name}_健康報告.docx`;
}
