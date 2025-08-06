/* server.js - Backend mit Excel-Verarbeitung & Web-Scraping */

const express = require('express');
const { chromium } = require('playwright');
const cors = require('cors');
const helmet = require('helmet');
const path = require('path');
const multer = require('multer');
const ExcelJS = require('exceljs');
const { execSync } = require('child_process');

const {
  toNumber,
  normalizeWeightToKg,
  parseDimensionsToLBH,
  normPartNo,
  mapMaterialClassificationToExcel,
  withinToleranceKG,
  a2vUrl
} = require('./utils');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(express.json());
app.use(express.static(__dirname));

// ---- Scraper ----
class SiemensProductScraper {
  constructor() {
    this.baseUrl = "https://www.mymobase.com/de/p/";
    this.browser = null;
  }
  async initBrowser() {
    if (!this.browser) {
      try {
        this.browser = await chromium.launch({
          headless: true,
          args: ['--no-sandbox','--disable-setuid-sandbox','--disable-dev-shm-usage']
        });
      } catch (error) {
        // try installing browsers on the fly
        try {
          execSync('npx playwright install --with-deps chromium', { stdio: 'inherit' });
          this.browser = await chromium.launch({ headless: true, args: ['--no-sandbox'] });
        } catch (e) {
          throw new Error('Chromium konnte nicht gestartet werden. Bitte führen Sie "npm run install-browsers" aus.');
        }
      }
    }
    return this.browser;
  }
  async closeBrowser() { if (this.browser) { await this.browser.close(); this.browser = null; } }

  async scrapeProduct(a2v) {
    const url = `${this.baseUrl}${a2v}`;
    const result = {
      URL: url,
      A2V: a2v,
      'Weitere Artikelnummer': 'Nicht gefunden',
      Produkttitel: 'Nicht gefunden',
      Gewicht: 'Nicht gefunden',
      Abmessung: 'Nicht gefunden',
      Werkstoff: 'Nicht gefunden',
      Materialklassifizierung: 'Nicht gefunden',
      Status: 'Init'
    };
    try {
      const browser = await this.initBrowser();
      const page = await browser.newPage();
      await page.goto(url, { waitUntil: 'networkidle', timeout: 45000 });

      // Titel
      try {
        const title = await page.title();
        if (title && !title.includes('404')) result.Produkttitel = title.replace(' | MoBase','').trim();
      } catch {}

      // Tabellen/Key-Value auslesen
      const kvPairs = await page.evaluate(() => {
        function add(map, k, v) {
          if (!k || !v) return;
          const key = k.trim().toLowerCase();
          const val = v.trim();
          if (!map[key]) map[key] = val;
        }
        const data = {};
        document.querySelectorAll('table').forEach(t => {
          t.querySelectorAll('tr').forEach(tr => {
            const tds = tr.querySelectorAll('td,th');
            if (tds.length >= 2) add(data, tds[0].textContent, tds[1].textContent);
          });
        });
        document.querySelectorAll('dl').forEach(dl => {
          const dts = dl.querySelectorAll('dt'); const dds = dl.querySelectorAll('dd');
          for (let i=0;i<Math.min(dts.length,dds.length);i++) add(data, dts[i].textContent, dds[i].textContent);
        });
        // generic spans/divs with colon
        document.querySelectorAll('div,span,li').forEach(el => {
          const txt = (el.textContent||'').trim();
          const idx = txt.indexOf(':');
          if (idx > 0 && idx < 80) {
            const key = txt.slice(0, idx);
            const val = txt.slice(idx+1);
            if (val && /\d|\w/.test(val)) add(data, key, val);
          }
        });
        return data;
      });

      function pick(keyContains) {
        const keys = Object.keys(kvPairs||{});
        for (const k of keys) {
          const low = k.toLowerCase();
          let ok = true;
          for (const needle of keyContains) {
            if (!low.includes(needle)) { ok = false; break; }
          }
          if (ok) return kvPairs[k];
        }
        return null;
      }

      const weitere = pick(['weitere','artikelnummer']) || pick(['additional','material','number']) || pick(['part','number']);
      if (weitere) result['Weitere Artikelnummer'] = weitere;

      const abm = pick(['abmess']) || pick(['dimension']);
      if (abm) result.Abmessung = abm;

      const gew = pick(['gewicht']) || pick(['weight']);
      if (gew) result.Gewicht = gew;

      const werk = pick(['werkstoff']) || pick(['material']) && !pick(['material','klass']);
      if (werk) result.Werkstoff = werk;

      const klass = pick(['material','klass']) || pick(['material','class']);
      if (klass) result.Materialklassifizierung = klass;

      result.Status = 'Erfolgreich';
      await page.close();
    } catch (e) {
      result.Status = `Fehler: ${e.message}`;
    }
    return result;
  }
}
const scraper = new SiemensProductScraper();

// ---- Helpers for comparison & row building ----
const COLS = { Z:'Z', E:'E', C:'C', S:'S', U:'U', V:'V', W:'W', P:'P', N:'N' };
const HEADER_ROW = 3;
const FIRST_DATA_ROW = 4;

function cell(worksheet, col, row) { return worksheet.getCell(`${col}${row}`); }

function statusCellFill(status) {
  if (status === 'GREEN') return { type:'pattern', pattern:'solid', fgColor:{ argb:'FFD5F4E6' } }; // green-ish
  if (status === 'RED')   return { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFDEAEA' } }; // red-ish
  return { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFF3CD' } }; // orange-ish
}

function compareText(excel, web) {
  if (!excel && !web) return { status:'ORANGE', comment:'Beide fehlen' };
  if (!excel) return { status:'ORANGE', comment:'Excel fehlt' };
  if (!web)   return { status:'ORANGE', comment:'Web fehlt' };
  const a = String(excel).trim().toLowerCase().replace(/\s+/g,' ');
  const b = String(web).trim().toLowerCase().replace(/\s+/g,' ');
  return a === b ? { status:'GREEN', comment:'identisch' } : { status:'RED', comment:'abweichend' };
}

function comparePartNo(excel, web) {
  if (!excel && !web) return { status:'ORANGE', comment:'Beide fehlen' };
  if (!excel) return { status:'ORANGE', comment:'Excel fehlt' };
  if (!web)   return { status:'ORANGE', comment:'Web fehlt' };
  return normPartNo(excel) === normPartNo(web)
    ? { status:'GREEN', comment:'identisch (normalisiert)' }
    : { status:'RED', comment:`abweichend: Excel ${excel} vs. Web ${web}` };
}

function compareWeight(excelVal, webVal) {
  const exKg = normalizeWeightToKg(excelVal);
  const wbKg = normalizeWeightToKg(webVal);
  if (exKg == null && wbKg == null) return { status:'ORANGE', comment:'Beide fehlen' };
  if (exKg == null) return { status:'ORANGE', comment:'Excel fehlt' };
  if (wbKg == null) return { status:'ORANGE', comment:'Web fehlt/unklar' };
  const ok = withinToleranceKG(exKg, wbKg, 5);
  const diffPct = ((wbKg - exKg) / Math.max(1e-9, Math.abs(exKg))) * 100;
  return ok
    ? { status:'GREEN', comment:`Δ ${diffPct.toFixed(1)}%` }
    : { status:'RED', comment:`Excel ${exKg.toFixed(3)} kg vs. Web ${wbKg.toFixed(3)} kg (${diffPct.toFixed(1)}%)` };
}

function compareDimensions(excelU, excelV, excelW, webDimText) {
  const L = toNumber(excelU); const B = toNumber(excelV); const H = toNumber(excelW);
  const fromWeb = parseDimensionsToLBH(webDimText);
  const allExcelPresent = L!=null && B!=null && H!=null;
  const anyExcel = L!=null || B!=null || H!=null;
  if (!anyExcel && !fromWeb.L && !fromWeb.B && !fromWeb.H) return { status:'ORANGE', comment:'Beide fehlen' };
  if (!anyExcel) return { status:'ORANGE', comment:'Excel fehlt' };
  if (!fromWeb.L && !fromWeb.B && !fromWeb.H) return { status:'ORANGE', comment:'Web fehlt/unklar' };
  const eq = (a,b)=> (a!=null && b!=null && Math.abs(a-b) < 1e-6);
  const match = eq(L, fromWeb.L) && eq(B, fromWeb.B) && eq(H, fromWeb.H);
  return match
    ? { status:'GREEN', comment:'L×B×H identisch (mm)' }
    : { status:'RED', comment:`Excel ${L||''}×${B||''}×${H||''} mm vs. Web ${fromWeb.L||''}×${fromWeb.B||''}×${fromWeb.H||''} mm` };
}

function compareMaterialClass(excelN, webText) {
  const mapped = mapMaterialClassificationToExcel(webText);
  if (!excelN && !mapped) return { status:'ORANGE', comment:'Beide fehlen' };
  if (!excelN) return { status:'ORANGE', comment:'Excel fehlt' };
  if (!mapped) return { status:'ORANGE', comment:'Web nicht interpretierbar' };
  return String(excelN).trim().toUpperCase() === mapped
    ? { status:'GREEN', comment:'identisch' }
    : { status:'RED', comment:`Excel ${excelN} vs. Web ${mapped}` };
}

// ---- Routes ----
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

app.get('/api/health', (req, res) => {
  res.json({
    status: 'OK',
    service: 'DB Produktvergleich API',
    browser: scraper.browser ? 'Bereit' : 'Nicht initialisiert',
    timestamp: new Date().toISOString()
  });
});

// File upload
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

/**
 * POST /api/process-excel
 * multipart/form-data: file=<excel>
 * Returns processed Excel with two extra rows per product.
 */
app.post('/api/process-excel', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Bitte Excel-Datei hochladen (file).' });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);

    // Process every worksheet; only first may have DB layout but we keep 1:1
    for (const ws of wb.worksheets) {
      // Determine last row with data by scanning column C or Z
      let lastRow = ws.lastRow?.number || 0;
      // ensure header row exists
      if (lastRow < HEADER_ROW) continue;

      // Collect products start..end first to avoid index shift
      const products = [];
      for (let r = FIRST_DATA_ROW; r <= lastRow; r++) {
        const anyValue = ['A','B','C','Z'].some(c => ws.getCell(`${c}${r}`).value && String(ws.getCell(`${c}${r}`).value).trim() !== '');
        if (anyValue) products.push(r);
      }

      // Iterate from bottom to top to insert rows
      for (let i = products.length - 1; i >= 0; i--) {
        const r = products[i];

        // Read Excel row fields
        const A2V = ws.getCell(`${COLS.Z}${r}`).value;
        const manufNoExcel = ws.getCell(`${COLS.E}${r}`).value;
        const titleExcel = ws.getCell(`${COLS.C}${r}`).value;
        const weightExcel = ws.getCell(`${COLS.S}${r}`).value;
        const lenExcel = ws.getCell(`${COLS.U}${r}`).value;
        const widExcel = ws.getCell(`${COLS.V}${r}`).value;
        const heiExcel = ws.getCell(`${COLS.W}${r}`).value;
        const werkstoffExcel = ws.getCell(`${COLS.P}${r}`).value;
        const noteExcel = ws.getCell(`${COLS.N}${r}`).value;

        // Scrape if we have A2V
        let webData = {
          URL: a2vUrl(A2V),
          A2V,
          'Weitere Artikelnummer': null,
          Produkttitel: null,
          Gewicht: null,
          Abmessung: null,
          Werkstoff: null,
          Materialklassifizierung: null,
          Status: 'Nicht versucht'
        };
        if (A2V && String(A2V).toUpperCase().startsWith('A2V')) {
          webData = await scraper.scrapeProduct(String(A2V).trim());
        }

        // ---- Insert two rows after r ----
        const insertAt = r + 1;
        ws.spliceRows(insertAt, 0, [null]); // Web-Daten
        ws.spliceRows(insertAt + 1, 0, [null]); // Vergleich
        const webRow = insertAt;
        const cmpRow = insertAt + 1;

        // Label cells (optional, we leave other cells empty)
        // Fill Web-Daten row ONLY in allowed columns
        ws.getCell(`${COLS.Z}${webRow}`).value = A2V || '';
        ws.getCell(`${COLS.E}${webRow}`).value = webData['Weitere Artikelnummer'] || '';
        ws.getCell(`${COLS.C}${webRow}`).value = webData.Produkttitel || '';
        ws.getCell(`${COLS.S}${webRow}`).value = webData.Gewicht || '';
        // Dimensions -> split to U/V/W
        const dimParsed = parseDimensionsToLBH(webData.Abmessung);
        if (dimParsed.L != null) ws.getCell(`${COLS.U}${webRow}`).value = dimParsed.L;
        if (dimParsed.B != null) ws.getCell(`${COLS.V}${webRow}`).value = dimParsed.B;
        if (dimParsed.H != null) ws.getCell(`${COLS.W}${webRow}`).value = dimParsed.H;
        ws.getCell(`${COLS.P}${webRow}`).value = webData.Werkstoff || '';
        ws.getCell(`${COLS.N}${webRow}`).value = mapMaterialClassificationToExcel(webData.Materialklassifizierung) || '';

        // Comparison row ONLY in allowed columns (set status + short comment, plus background color)
        // Z (A2V) - direct text compare
        const cmpZ = compareText(A2V || '', webData.A2V || A2V || '');
        ws.getCell(`${COLS.Z}${cmpRow}`).value = cmpZ.comment;
        ws.getCell(`${COLS.Z}${cmpRow}`).fill = statusCellFill(cmpZ.status);

        // E (Herstellerartikelnummer) vs "Weitere Artikelnummer"
        const cmpE = comparePartNo(manufNoExcel || '', webData['Weitere Artikelnummer'] || '');
        ws.getCell(`${COLS.E}${cmpRow}`).value = cmpE.comment;
        ws.getCell(`${COLS.E}${cmpRow}`).fill = statusCellFill(cmpE.status);

        // C (Material-Kurztext) vs Produkttitel
        const cmpC = compareText(titleExcel || '', webData.Produkttitel || '');
        ws.getCell(`${COLS.C}${cmpRow}`).value = cmpC.comment;
        ws.getCell(`${COLS.C}${cmpRow}`).fill = statusCellFill(cmpC.status);

        // S (Gewicht) ±5%
        const cmpS = compareWeight(weightExcel, webData.Gewicht);
        ws.getCell(`${COLS.S}${cmpRow}`).value = cmpS.comment;
        ws.getCell(`${COLS.S}${cmpRow}`).fill = statusCellFill(cmpS.status);

        // U/V/W (L/B/H) vs Abmessungen
        const cmpDim = compareDimensions(lenExcel, widExcel, heiExcel, webData.Abmessung);
        ws.getCell(`${COLS.U}${cmpRow}`).value = cmpDim.comment;
        ws.getCell(`${COLS.U}${cmpRow}`).fill = statusCellFill(cmpDim.status);
        // leave V/W empty per requirement (only allowed columns may be written; U/V/W are allowed, but we put comment in U)

        // P (Werkstoff)
        const cmpP = compareText(werkstoffExcel || '', webData.Werkstoff || '');
        ws.getCell(`${COLS.P}${cmpRow}`).value = cmpP.comment;
        ws.getCell(`${COLS.P}${cmpRow}`).fill = statusCellFill(cmpP.status);

        // N (Materialklassifizierung mapping)
        const cmpN = compareMaterialClass(noteExcel || '', webData.Materialklassifizierung || '');
        ws.getCell(`${COLS.N}${cmpRow}`).value = cmpN.comment;
        ws.getCell(`${COLS.N}${cmpRow}`).fill = statusCellFill(cmpN.status);
      }
    }

    // Write to buffer and respond
    const out = await wb.xlsx.writeBuffer();
    const fileName = 'DB_Produktvergleich_verarbeitet.xlsx';
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition',`attachment; filename="${fileName}"`);
    res.send(Buffer.from(out));

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// graceful shutdown
process.on('SIGINT', async () => { await scraper.closeBrowser(); process.exit(0); });
process.on('SIGTERM', async () => { await scraper.closeBrowser(); process.exit(0); });

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
  console.log(`Health: http://localhost:${PORT}/api/health`);
});