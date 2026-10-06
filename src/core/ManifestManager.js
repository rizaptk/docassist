/* ManifestManager — DocManifest cache: CustomXmlParts (ikut .docx) + localStorage (instan). Tulis debounced, baca sinkron-cepat. */
const NS = "http://docassist/ai-manifest/v1";
const LS_KEY = "docassist:manifest:v1";
let mem = null, saveT = null;

const blank = (hash = "") => ({ schema: "docassist/manifest/v1", docHash: hash, updatedAt: new Date().toISOString(),
  outline: [], styleProfile: { lang: "id", tone: "formal-netral" }, glossary: [], sectionSummary: {}, styleInventory: [], doNotTouch: [] });

export function cheapHash(meta, winLen) {
  const s = `${meta.paragraf}|${(meta.headings || []).join(";").slice(0, 400)}|${winLen}`;
  let h = 0; for (let i = 0; i < s.length; i++) h = ((h << 5) - h + s.charCodeAt(i)) | 0;
  return (h >>> 0).toString(16);
}

export async function loadManifest() {
  if (mem) return mem;
  try { const raw = localStorage.getItem(LS_KEY); if (raw) { mem = JSON.parse(raw); return mem; } } catch (e) {}
  try {
    const parts = await Office.context.document.customXmlParts.getByNamespaceAsync(NS).then
      ? await new Promise((res) => Office.context.document.customXmlParts.getByNamespaceAsync(NS, (r) => res(r)))
      : null;
    const id = parts && parts.value && parts.value[0];
    if (id) { const xml = await new Promise((res) => id.getXmlAsync((r) => res(r.value))); const m = JSON.parse(xml.match(/<m>([\s\S]*)<\/m>/)?.[1] || "null"); if (m) { mem = m; return m; } }
  } catch (e) {}
  return null;
}

export function saveManifestLocal(m) { mem = m; clearTimeout(saveT); saveT = setTimeout(persist, 600); }
let lastWrite = 0;
async function persist() {
  if (!mem) return;
  mem.updatedAt = new Date().toISOString();
  try { localStorage.setItem(LS_KEY, JSON.stringify(mem)); } catch (e) {}
  if (Date.now() - lastWrite < 30000) return; // throttle tulis dokumen: maks 1x/30 dtk
  lastWrite = Date.now();
  // Tulis ke dokumen di idle agar tidak blokir ketikan user; hapus part lama dulu anti-bengkak.
  const write = () => {
    try {
      Office.context.document.customXmlParts.getByNamespaceAsync(NS, (r) => {
        const parts = (r.status === Office.AsyncResultStatus.Succeeded && r.value) || [];
        let i = 0;
        const next = () => {
          if (i < parts.length) { const p = parts[i++]; try { p.deleteAsync(() => next()); } catch (e) { next(); } }
          else { try { Office.context.document.customXmlParts.addAsync(`<docassist><m>${JSON.stringify(mem)}</m></docassist>`, { namespace: NS }, () => {}); } catch (e) {} }
        };
        next();
      });
    } catch (e) {}
  };
  if ("requestIdleCallback" in window) requestIdleCallback(write, { timeout: 2000 }); else setTimeout(write, 800);
}

export function ensureManifest(hash) { if (!mem) mem = blank(hash); return mem; }
export function isStale(cachedHash, curHash) { return cachedHash !== curHash; }
