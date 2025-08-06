// utils.js - Normalisierung, Parser & Vergleichslogik (optimiert)

function toNumber(val) {
  if (val === null || val === undefined) return null;
  if (typeof val === 'number') return val;
  const s = String(val).replace(/\s/g, '').replace(',', '.').trim();
  const m = s.match(/-?\d+(\.\d+)?/);
  return m ? parseFloat(m[0]) : null;
}

function normalizeWeightToKg(value) {
  if (value === null || value === undefined || value === '') return null;
  const s = String(value).toLowerCase().replace(',', '.').replace(/\s+/g, '');
  const num = toNumber(s);
  if (num === null) return null;
  if (s.includes('mg')) return num / 1e6;
  if (s.includes('g') && !s.includes('kg')) return num / 1000.0;
  if (s.includes('t')) return num * 1000.0;
  return num; // assume kg
}

function normalizeLenToMm(val) {
  if (val === null || val === undefined || val === '') return null;
  const s = String(val).toLowerCase().replace(',', '.').replace(/\s+/g, '');
  const num = toNumber(s);
  if (num === null) return null;
  if (s.includes('cm')) return num * 10.0;
  if (s.includes('m')) return num * 1000.0;
  return num; // default mm
}

// Parse dimension text to L,B,H (mm)
function parseDimensionsToLBH(text) {
  if (!text) return { L: null, B: null, H: null, raw: '' };
  let raw = String(text).trim();
  let s = raw.toLowerCase()
    .replace(/[，、]/g, ',')
    .replace(/[×xX*]/g, 'x')
    .replace(/\s+/g, '');
  const nums = s.match(/-?\d+(?:\.\d+)?/g) || [];
  const L = nums[0] ? normalizeLenToMm(nums[0]) : null;
  const B = nums[1] ? normalizeLenToMm(nums[1]) : null;
  const H = nums[2] ? normalizeLenToMm(nums[2]) : null;
  return { L, B, H, raw };
}

function normPartNo(s) {
  if (!s && s !== 0) return '';
  return String(s).toUpperCase().replace(/[\s-]/g, '');
}

function mapMaterialClassificationToExcel(text) {
  if (!text) return null;
  const t = String(text).toLowerCase();
  if (t.includes('nicht') && t.includes('schweiss') && t.includes('guss') && t.includes('klebe') && t.includes('schmiede')) {
    return 'OHNE/N/N/N/N';
  }
  if (t.includes('nicht') && (t.includes('schweiss') || t.includes('schweiß'))) {
    return 'OHNE/N/N/N/N';
  }
  return null;
}

// Default tolerance is 0 (strict). server.js can override by passing pct.
function withinToleranceKG(excelKg, webKg, pct=0, eps=1e-6) {
  if (excelKg == null || webKg == null) return false;
  const tol = Math.abs(excelKg) * (pct / 100.0);
  return Math.abs(webKg - excelKg) <= Math.max(tol, eps);
}

function a2vUrl(a2v) {
  if (!a2v) return null;
  return `https://www.mymobase.com/de/p/${String(a2v).trim()}`;
}

module.exports = {
  toNumber,
  normalizeWeightToKg,
  normalizeLenToMm,
  parseDimensionsToLBH,
  normPartNo,
  mapMaterialClassificationToExcel,
  withinToleranceKG,
  a2vUrl
};