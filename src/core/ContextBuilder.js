/* ContextBuilder — layered context hemat token. L0+L1 selalu; L2 dari cache; L3 hanya on-demand. */
import { bacaSeleksi, bacaJendela, bacaMeta } from "./WordOps.js";

const L1_MAX = 4000;
let lastWin = { seleksi: "", konteks: "", at: 0 };
let pending = null;

function debounceIdle(fn, ms = 400) {
  return (...a) => { clearTimeout(pending); pending = setTimeout(() => {
    if ("requestIdleCallback" in window) requestIdleCallback(() => fn(...a), { timeout: 800 });
    else fn(...a);
  }, ms); };
}

// Dipanggil saat seleksi berubah — hanya update ringan, TIDAK panggil AI/Word berat tiap ketik.
export const onSelectionHint = debounceIdle(async (cb) => {
  const t = await bacaSeleksi();
  if (t != null && t !== lastWin.seleksi) { lastWin = { seleksi: t, konteks: t.slice(0, 500), at: Date.now() }; cb && cb(t); }
});

export async function buildContext(manifest) {
  const [meta, win] = await Promise.all([bacaMeta(), bacaJendela(L1_MAX)]);
  const l0 = { paragraf: meta.paragraf, nHead: (meta.headings || []).length, kursorAda: !!(win.seleksi || "").length };
  const l2 = manifest ? {
    outline: manifest.outline || [], glossary: manifest.glossary || [],
    tone: manifest.styleProfile?.tone, lang: manifest.styleProfile?.lang || "id",
  } : null;
  return { l0, l1: win, l2 };
}

export function budgetInfo(ctx) {
  const s = JSON.stringify(ctx).length;
  return { chars: s, estTokens: Math.round(s / 4) };
}
