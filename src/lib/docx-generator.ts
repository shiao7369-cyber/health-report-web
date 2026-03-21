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

  // □/■ 在標楷體中字型偏小，需分段並加大字型來補正
  type Seg = { text: string; box: boolean };
  const segs: Seg[] = [];
  for (const ch of text) {
    const box = ch === "□" || ch === "■";
    if (segs.length && segs[segs.length - 1].box === box) {
      segs[segs.length - 1].text += ch;
    } else {
      segs.push({ text: ch, box });
    }
  }

  function makeRun(seg: Seg): Element {
    const rEl = doc.createElementNS(NS, "w:r");
    const rPr = doc.createElementNS(NS, "w:rPr");
    const fnt = doc.createElementNS(NS, "w:rFonts");
    fnt.setAttributeNS(NS, "w:ascii",    "標楷體");
    fnt.setAttributeNS(NS, "w:eastAsia", "標楷體");
    fnt.setAttributeNS(NS, "w:hAnsi",    "標楷體");
    const szEl = doc.createElementNS(NS, "w:sz");
    // 框框字元比中文字視覺偏小，加 12 個半點（+6pt）補正
    szEl.setAttributeNS(NS, "w:val", seg.box ? String(parseInt(sz) + 12) : sz);
    rPr.appendChild(fnt);
    rPr.appendChild(szEl);
    rEl.appendChild(rPr);
    const tEl = doc.createElementNS(NS, "w:t");
    tEl.textContent = seg.text;
    if (seg.text && (seg.text[0] === " " || seg.text[seg.text.length - 1] === " ")) {
      tEl.setAttribute("xml:space", "preserve");
    }
    rEl.appendChild(tEl);
    return rEl;
  }

  for (const seg of segs) {
    p.appendChild(makeRun(seg));
  }
}

/** 找到含 keyword 的儲存格，清空並填入新文字；newText 為 null 時跳過（保留範本原樣） */
function replaceCellText(tr: Element, keyword: string, newText: string | null) {
  if (newText === null) return; // 未測量，保留範本原樣
  const tcs = Array.from(tr.getElementsByTagNameNS(NS, "tc"));
  for (const tc of tcs) {
    if (allText(tc).includes(keyword)) {
      const paras = Array.from(tc.getElementsByTagNameNS(NS, "p"));
      clearAndFill(paras, newText);
      return;
    }
  }
}

/** 找到含 keyword 的 w:t 節點，直接替換文字；newText 為 null 時跳過 */
function fillTextNode(tr: Element, keyword: string, newText: string | null) {
  if (newText === null) return;
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
      replaceCellText(tr, "□戒菸□節酒", buildCounselLine1(rowData));
      continue;
    }
    if (joined.includes("□健康飲食(含我的健康餐盤)")) {
      replaceCellText(tr, "□健康飲食(含我的健康餐盤)", buildCounselLine2(rowData));
      continue;
    }
    if (joined.includes("慢性疾病風險評估") && !joined.includes("腎病識能") && !joined.includes("風險值")) {
      const chronicVal = v(rowData, "counsel_chronic2") || v(rowData, "counsel_chronic");
      const mark = chronicVal === "是" ? "■" : "□";
      replaceCellText(tr, "慢性疾病風險評估", `${mark}慢性疾病風險評估`);
      continue;
    }
    if (joined.includes("代謝症候群定義")) {
      // 保留原文字但統一字型大小為 10pt
      const tcs = Array.from(tr.getElementsByTagNameNS(NS, "tc"));
      for (const tc of tcs) {
        if (allText(tc).includes("代謝症候群定義")) {
          const text = allText(tc);
          const paras = Array.from(tc.getElementsByTagNameNS(NS, "p"));
          clearAndFill(paras, text);
          break;
        }
      }
      continue;
    }
    if (joined.includes("腎病識能衛教指導") || joined.includes("□腎病識能")) {
      replaceCellText(tr, "腎病識能", buildCounselKidney(rowData));
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
      replaceCellText(tr, "B型肝炎表面抗原", buildHbsagLine(rowData));
      continue;
    }

    // ── C 型肝炎 ─────────────────────────────────────────────────────────────
    if (joined.includes("C型肝炎抗體")) {
      replaceCellText(tr, "C型肝炎抗體", buildHcvLine(rowData));
      continue;
    }

    // ── 咳嗽症狀（還原範本預設，走 clearAndFill 套用框框加大）──────────────
    if (joined.includes("咳嗽症狀：")) {
      const tcs = Array.from(tr.getElementsByTagNameNS(NS, "tc"));
      for (const tc of tcs) {
        const tcText = allText(tc);
        if (tcText.includes("咳嗽症狀：")) {
          const newText = tcText.replace(/☑/g, "□");
          const paras = Array.from(tc.getElementsByTagNameNS(NS, "p"));
          clearAndFill(paras, newText);
          break;
        }
      }
      continue;
    }

    // ── 憂鬱檢測 ────────────────────────────────────────────────────────────
    if (joined.includes("憂鬱檢測：")) {
      replaceCellText(tr, "憂鬱檢測：", buildDepressionLine(rowData));
      continue;
    }
  }

  // ── 縮小頁邊距，四邊均為 5mm ─────────────────────────────────────────────────
  const MARGIN = "280"; // 280 twips ≈ 5 mm
  const sectPrList = Array.from(doc.getElementsByTagNameNS(NS, "sectPr"));
  for (const sectPr of sectPrList) {
    const pgMarList = Array.from(sectPr.getElementsByTagNameNS(NS, "pgMar"));
    for (const pgMar of pgMarList) {
      pgMar.setAttributeNS(NS, "w:top",    MARGIN);
      pgMar.setAttributeNS(NS, "w:right",  MARGIN);
      pgMar.setAttributeNS(NS, "w:bottom", MARGIN);
      pgMar.setAttributeNS(NS, "w:left",   MARGIN);
      pgMar.setAttributeNS(NS, "w:header", MARGIN);
      pgMar.setAttributeNS(NS, "w:footer", MARGIN);
    }
  }

  // ── 消除 body 層級的尾端段落高度（Word 強制保留，但會造成底部多餘空白）────────
  const bodyEl = doc.getElementsByTagNameNS(NS, "body")[0];
  if (bodyEl) {
    for (const node of Array.from(bodyEl.childNodes)) {
      const el = node as Element;
      if (el.namespaceURI === NS && el.localName === "p") {
        // 找到或建立 w:pPr
        let pPr = el.getElementsByTagNameNS(NS, "pPr")[0];
        if (!pPr) {
          pPr = doc.createElementNS(NS, "w:pPr");
          el.insertBefore(pPr, el.firstChild);
        }
        // 行距與段落間距設為絕對最小值
        let spacing = pPr.getElementsByTagNameNS(NS, "spacing")[0];
        if (!spacing) {
          spacing = doc.createElementNS(NS, "w:spacing");
          pPr.appendChild(spacing);
        }
        spacing.setAttributeNS(NS, "w:before",   "0");
        spacing.setAttributeNS(NS, "w:after",    "0");
        spacing.setAttributeNS(NS, "w:line",     "1");
        spacing.setAttributeNS(NS, "w:lineRule", "exact");
        // 字型大小設為 1（最小），確保行高趨近於零
        let pRpr = pPr.getElementsByTagNameNS(NS, "rPr")[0];
        if (!pRpr) {
          pRpr = doc.createElementNS(NS, "w:rPr");
          pPr.appendChild(pRpr);
        }
        let pSz = pRpr.getElementsByTagNameNS(NS, "sz")[0];
        if (!pSz) { pSz = doc.createElementNS(NS, "w:sz"); pRpr.appendChild(pSz); }
        pSz.setAttributeNS(NS, "w:val", "1");
      }
    }
  }

  // ── 主表格寬度設為 100% 填滿內文區域 ────────────────────────────────────────
  const tblList = Array.from(doc.getElementsByTagNameNS(NS, "tbl"));
  if (tblList.length > 0) {
    const mainTbl = tblList[0];
    const tblPr = mainTbl.getElementsByTagNameNS(NS, "tblPr")[0];
    if (tblPr) {
      const tblWList = Array.from(tblPr.getElementsByTagNameNS(NS, "tblW"));
      for (const tblW of tblWList) {
        // A5 橫向 - 左右各 280 twips 邊距 → 內文寬 = 11906 - 560 = 11346 twips
        tblW.setAttributeNS(NS, "w:w",    "11346");
        tblW.setAttributeNS(NS, "w:type", "dxa");
      }
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
