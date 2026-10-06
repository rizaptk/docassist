/* OoxmlKit — blok bangunan OOXML premium untuk semua kit (CV, surat, flyer, artikel).
   Murni string builder (testable di Node). Warna/font eksplisit per desain, bukan theme.
   Palet: navy 1F3864, teal 0E5F5B, biru 2E74B5, peach E8B48C, peach-bg FBE9DC, abu F2F2F2, teks 262626. */
import { esc } from "../FlowchartOoxml.js";

export const DXA_PAGE = 9360;
const FONT = "Calibri";

function rpr(o = {}) {
  const f = o.font || FONT; // Cambria untuk serif klasik, Calibri default
  let s = `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}"/>`;
  if (o.bold) s += "<w:b/><w:bCs/>";
  if (o.italic) s += "<w:i/><w:iCs/>";
  if (o.color) s += `<w:color w:val="${o.color}"/>`;
  if (o.size) s += `<w:sz w:val="${o.size * 2}"/><w:szCs w:val="${o.size * 2}"/>`;
  return s;
}

export function R(text, o = {}) {
  return `<w:r><w:rPr>${rpr(o)}</w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}

// Paragraf: o = {align, bold, color, size, after, rule (warna garis bawah), gapBefore}
export function P(text, o = {}) {
  let ppr = "";
  if (o.align) ppr += `<w:jc w:val="${o.align}"/>`;
  ppr += `<w:spacing w:after="${o.after ?? 100}" w:line="276" w:lineRule="auto"/>`;
  if (o.rule) ppr += `<w:pBdr><w:bottom w:val="single" w:sz="9" w:space="1" w:color="${o.rule}"/></w:pBdr>`;
  return `<w:p><w:pPr>${ppr}</w:pPr>${R(text, o)}</w:p>`;
}

export function EMPTY(after = 100) {
  return `<w:p><w:pPr><w:spacing w:after="${after}"/></w:pPr></w:p>`;
}

// Sel tabel: paras = array string XML <w:p>; o = {fill (hex tanpa #), width (dxa)}
export function CELL(paras, o = {}) {
  const w = o.width || 0;
  const shd = o.fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${o.fill}"/>` : "";
  return `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${shd}</w:tcPr>${paras.join("")}</w:tc>`;
}
export function ROW(cells) { return `<w:tr>${cells.join("")}</w:tr>`; }

// Tabel tanpa border: rows = array ROW, widths = array dxa per kolom.
export function TBL(rows, widths) {
  const grid = widths.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
  const nil = ["top", "left", "bottom", "right", "insideH", "insideV"].map((e) => `<w:${e} w:val="nil" w:sz="0" w:space="0" w:color="auto"/>`).join("");
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${nil}</w:tblBorders></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows.join("")}</w:tbl>`;
}

// Bungkus penuh siap setSelectedDataAsync (pola sampel resmi Microsoft).
export function wrapPkg(bodyXml) {
  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512"><pkg:xmlData>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>` +
    `</pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyXml}</w:body></w:document>` +
    `</pkg:xmlData></pkg:part></pkg:package>`;
}

// Bar skill "■■■■■□□□" (teks biasa — aman di semua font, tanpa gambar).
export function skillBar(level = 3, max = 7) {
  const l = Math.max(1, Math.min(max, level || 3));
  return "■".repeat(l) + "□".repeat(max - l);
}
