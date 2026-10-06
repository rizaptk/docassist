/* FlowchartOoxml — flowchart EDITABLE sebagai native DrawingML shapes (wpg group),
   disisip via setSelectedDataAsync(images/ooxml). User bisa: edit teks tiap kotak,
   pindah/resize/restyle, ungroup.
   Struktur meniru sampel resmi Microsoft (ShapeWithText.xml): wsp minimal
   (cNvSpPr→spPr→txbx→bodyPr), grup wpg, wrap wp:inline + mc:AlternateContent.
   Batas jujur (ponytail: upgrade ke stCxn/endCxn refs bila terbukti stabil):
   - panah digambar pada posisi hasil layout → TIDAK menempel otomatis saat kotak dipindah.
   - label edge = kotak teks kecil tanpa garis.
   - layout vertikal (TD) saja; maks 20 node (lebih → fallback gambar).
   Pure string builder — bisa diuji di Node tanpa Word/DOM. */

const EMU = 12700; // 1pt dalam EMU
const NODE_W = 150, NODE_H = 56, RANK_GAP = 66, NODE_GAP = 30, PAD = 24;

const GEOM = { process: "rect", start: "ellipse", end: "ellipse", terminator: "roundRect", decision: "diamond", input: "parallelogram" };
const FILL = { process: "DBEAFE", start: "D1FAE5", end: "FEE2E2", terminator: "E0E7FF", decision: "FEF3C7", input: "F3E8FF" };

export function esc(s) {
  return String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// flat: label berupa objek → gabung values. Anti "[object Object]" di shapes.
const flat = (v) => Array.isArray(v) ? v.map(flat).filter(Boolean).join(", ")
  : (v && typeof v === "object" ? Object.values(v).map(flat).filter(Boolean).join(" — ") : String(v ?? ""));

function normShape(s) { return GEOM[s] ? s : "process"; }

// Rank = longest-path dari root, tahan cycle (relaksasi ≤ N pass).
export function layoutFlow(flow) {
  const raw = flow.nodes || [];
  if (!raw.length) throw new Error("flowchart butuh ≥1 node");
  if (raw.length > 20) throw new Error("maks 20 node untuk mode editable (lebih → gunakan gambar)");
  const nodes = raw.map((n, i) => ({ id: String(n.id ?? "n" + (i + 1)), label: flat(n.label).slice(0, 60), shape: normShape(n.shape) }));
  const ids = new Set(nodes.map((n) => n.id));
  const edges = (flow.edges || []).filter((e) => ids.has(String(e.from)) && ids.has(String(e.to))).slice(0, 30)
    .map((e) => ({ from: String(e.from), to: String(e.to), label: String(e.label ?? "").slice(0, 24) }));
  const rank = {}; nodes.forEach((n) => (rank[n.id] = 0));
  for (let p = 0; p < nodes.length; p++)
    for (const e of edges) rank[e.to] = Math.max(rank[e.to], rank[e.from] + 1);
  const byRank = {};
  nodes.forEach((n) => { n.rank = Math.min(rank[n.id], nodes.length); (byRank[n.rank] = byRank[n.rank] || []).push(n); });
  const ranks = Object.keys(byRank).map(Number).sort((a, b) => a - b);
  const maxW = Math.max(...ranks.map((r) => byRank[r].length * NODE_W + (byRank[r].length - 1) * NODE_GAP));
  const W = maxW + PAD * 2;
  const H = PAD * 2 + ranks.length * NODE_H + (ranks.length - 1) * RANK_GAP;
  ranks.forEach((r, ri) => {
    const row = byRank[r], rowW = row.length * NODE_W + (row.length - 1) * NODE_GAP;
    row.forEach((n, i) => { n.x = PAD + (maxW - rowW) / 2 + i * (NODE_W + NODE_GAP); n.y = PAD + ri * (NODE_H + RANK_GAP); n.w = NODE_W; n.h = NODE_H; });
  });
  const at = {}; nodes.forEach((n) => (at[n.id] = n));
  return { nodes, edges, at, W, H };
}

const pt = (v) => Math.round(v * EMU);

function wspNode(n) {
  const X = pt(n.x), Y = pt(n.y), W = pt(n.w), H = pt(n.h);
  return `<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="${X}" y="${Y}"/><a:ext cx="${W}" cy="${H}"/></a:xfrm>` +
    `<a:prstGeom prst="${GEOM[n.shape]}"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="${FILL[n.shape]}"/></a:solidFill>` +
    `<a:ln w="12700"><a:solidFill><a:srgbClr val="2563EB"/></a:solidFill></a:ln></wps:spPr>` +
    `<wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/><w:color w:val="1E3A8A"/></w:rPr>` +
    `<w:t xml:space="preserve">${esc(n.label)}</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
    `<wps:bodyPr rot="0" vert="horz" wrap="square" anchor="ctr"><a:noAutofit/></wps:bodyPr></wps:wsp>`;
}

// Panah garis lurus ber-arrowhead pada posisi layout (tanpa stCxn/endCxn refs).
function cxnLine(x1, y1, x2, y2) {
  const X = pt(Math.min(x1, x2)), Y = pt(Math.min(y1, y2));
  const W = Math.max(pt(Math.abs(x2 - x1)), 1270), H = Math.max(pt(Math.abs(y2 - y1)), 1270);
  return `<wps:cxnSp><wps:cNvCxnSpPr/><wps:spPr><a:xfrm><a:off x="${X}" y="${Y}"/><a:ext cx="${W}" cy="${H}"/></a:xfrm>` +
    `<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>` +
    `<a:ln w="12700"><a:solidFill><a:srgbClr val="64748B"/></a:solidFill>` +
    `<a:tailEnd type="triangle" w="med" len="med"/></a:ln></wps:spPr></wps:cxnSp>`;
}

function edgeLabelBox(txt, mx, my) {
  const w = Math.min(120, 20 + txt.length * 6), X = pt(mx - w / 2), Y = pt(my - 14);
  return `<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="${X}" y="${Y}"/><a:ext cx="${pt(w)}" cy="${pt(20)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>` +
    `<wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/><w:color w:val="475569"/></w:rPr>` +
    `<w:t xml:space="preserve">${esc(txt)}</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
    `<wps:bodyPr rot="0" vert="horz" wrap="square" anchor="ctr"><a:noAutofit/></wps:bodyPr></wps:wsp>`;
}

export function buildFlowchartOoxml(flow, docId) {
  const { nodes, edges, at, W, H } = layoutFlow(flow);
  const id = docId || (100000 + Math.floor(Math.random() * 899999));
  let kids = "";
  for (const n of nodes) kids += wspNode(n);
  for (const e of edges) {
    const a = at[e.from], b = at[e.to];
    const x1 = a.x + a.w / 2, y1 = a.y + a.h, x2 = b.x + b.w / 2, y2 = b.y;
    kids += cxnLine(x1, y1, x2, y2);
    if (e.label) kids += edgeLabelBox(e.label, (x1 + x2) / 2, (y1 + y2) / 2);
  }
  const GW = pt(W), GH = pt(H);
  const p =
    `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
    `<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr>` +
    `<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${GW}" cy="${GH}"/><wp:docPr id="${id}" name="DocAssist Flowchart"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">` +
    `<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${GW}" cy="${GH}"/><a:chOff x="0" y="0"/><a:chExt cx="${GW}" cy="${GH}"/></a:xfrm></wpg:grpSpPr>` +
    kids + `</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>`;
  // Bungkus paket penuh seperti sampel resmi (setSelectedDataAsync butuh pkg:package).
  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512"><pkg:xmlData>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>` +
    `</pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${p}</w:body></w:document>` +
    `</pkg:xmlData></pkg:part></pkg:package>`;
}

// Turunan mermaid untuk pratinjau di widget (bukan yang disisip).
export function flowToMermaid(flow) {
  const nodes = (flow.nodes || []).map((n, i) => {
    const id = String(n.id ?? "n" + (i + 1)).replace(/[^A-Za-z0-9_]/g, "_");
    const lb = String(n.label ?? "").replace(/"/g, "'");
    const s = normShape(n.shape);
    const open = s === "decision" ? `{` : s === "start" || s === "end" ? `([` : s === "input" ? `[/` : `["`;
    const close = s === "decision" ? `}` : s === "start" || s === "end" ? `])` : s === "input" ? `/]` : `"]`;
    return `${id}${open}"${lb}"${close}`;
  });
  const edges = (flow.edges || []).map((e) => {
    const f = String(e.from).replace(/[^A-Za-z0-9_]/g, "_"), t = String(e.to).replace(/[^A-Za-z0-9_]/g, "_");
    return e.label ? `${f}-->|"${String(e.label).replace(/"/g, "'")}"|${t}` : `${f}-->${t}`;
  });
  return `graph TD\n${[...nodes, ...edges].join("\n")}`;
}
