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
      console.log('🚀 Starte Browser...');
      try {
        this.browser = await chromium.launch({
          headless: true,
          args: [
            '--no-sandbox', 
            '--disable-setuid-sandbox',
            '--disable-dev-shm-usage',
            '--disable-web-security',
            '--disable-features=VizDisplayCompositor'
          ]
        });
      } catch (error) {
        console.error('❌ Browser-Start fehlgeschlagen:', error.message);
        
        // Versuche Browser zu installieren und erneut zu starten
        console.log('🔄 Versuche Browser-Installation...');
        try {
          execSync('npx playwright install chromium', { stdio: 'inherit' });
          
          this.browser = await chromium.launch({
            headless: true,
            args: [
              '--no-sandbox', 
              '--disable-setuid-sandbox',
              '--disable-dev-shm-usage',
              '--disable-web-security',
              '--disable-features=VizDisplayCompositor'
            ]
          });
          console.log('✅ Browser erfolgreich installiert und gestartet');
        } catch (installError) {
          console.error('❌ Browser-Installation fehlgeschlagen:', installError.message);
          throw new Error('Browser konnte nicht installiert werden. Bitte führen Sie "npm run install-browsers" manuell aus.');
        }
      }
    }
    return this.browser;
  }

  async closeBrowser() { 
    if (this.browser) { 
      await this.browser.close(); 
      this.browser = null; 
    } 
  }

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
      
      // Set user agent to appear more like a real browser
      await page.setExtraHTTPHeaders({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
      });

      console.log(`🔍 Lade Seite: ${url}`);
      
      const response = await page.goto(url, { 
        waitUntil: 'networkidle', 
        timeout: 30000 
      });

      if (!response) {
        result.Status = "Keine Antwort vom Server";
        return result;
      }

      if (response.status() === 404) {
        result.Status = "Produkt nicht gefunden (404)";
        return result;
      } else if (response.status() !== 200) {
        result.Status = `HTTP-Fehler: ${response.status()}`;
        return result;
      }

      // Wait for page to load
      await page.waitForTimeout(2000);

      // Extract page title
      try {
        const title = await page.title();
        if (title && !title.includes('404') && !title.includes('Not Found')) {
          result.Produkttitel = title.replace(" | MoBase", "").replace(" | Siemens Mobility", "").trim();
        }
      } catch (e) {
        console.log('⚠️ Titel nicht gefunden:', e.message);
      }

      // PRIMARY METHOD: Table-based extraction
      console.log('🔄 Starte robuste Table-basierte Extraktion...');
      await this.extractTechnicalData(page, result);

      // SECONDARY METHOD: Extract data from JavaScript initialData object
      if (result.Werkstoff === "Nicht gefunden" || result.Materialklassifizierung === "Nicht gefunden" || result['Statistische Warennummer'] === "Nicht gefunden") {
        console.log('🔄 Ergänze fehlende Felder mit JavaScript initialData...');
        await this.extractFromInitialData(page, result);
      } else {
        console.log('✅ Table-Extraktion vollständig - Skip initialData');
      }

      // TERTIARY METHOD: HTML fallback extraction
      if (result.Werkstoff === "Nicht gefunden" || result.Materialklassifizierung === "Nicht gefunden") {
        console.log('🔄 Verwende erweiterte HTML-Extraktion...');
        await this.extractFromHTML(page, result);
      }

      // Extract product details from various selectors
      await this.extractProductDetails(page, result);

      // Check if we got meaningful data
      const hasData = result.Werkstoff !== "Nicht gefunden" || 
                    result.Materialklassifizierung !== "Nicht gefunden" ||
                    result.Gewicht !== "Nicht gefunden" ||
                    result.Abmessung !== "Nicht gefunden";
      
      if (hasData) {
        result.Status = "Erfolgreich";
        console.log('✅ Scraping erfolgreich - Daten gefunden');
      } else {
        result.Status = "Teilweise erfolgreich - Wenig Daten gefunden";
        console.log('⚠️ Scraping unvollständig - Wenig Daten extrahiert');
        
        // Try to get at least the page title if nothing else worked
        try {
          const pageTitle = await page.title();
          if (pageTitle && !pageTitle.includes('404')) {
            result.Produkttitel = pageTitle.replace(" | MoBase", "").trim();
          }
        } catch (titleError) {
          console.log('⚠️ Auch Titel-Extraktion fehlgeschlagen');
        }
      }
      
      await page.close();
      
    } catch (e) {
      console.error('❌ Scraping Fehler:', e);
      result.Status = `Fehler: ${e.message}`;
    }
    
    return result;
  }

  async extractTechnicalData(page, result) {
    try {
      // Look for tables with technical data
      const tables = await page.$$('table');
      
      for (const table of tables) {
        const rows = await table.$$('tr');
        
        for (const row of rows) {
          const cells = await row.$$('td, th');
          
          if (cells.length >= 2) {
            const keyElement = cells[0];
            const valueElement = cells[1];
            
            const key = (await keyElement.innerText()).trim().toLowerCase();
            const value = (await valueElement.innerText()).trim();
            
            if (key && value) {
              this.mapTechnicalField(key, value, result);
            }
          }
        }
      }

      // Look for product specifications in div elements
      const specDivs = await page.$$('[class*="spec"], [class*="detail"], [class*="info"]');
      for (const div of specDivs) {
        try {
          const text = await div.innerText();
          this.parseSpecificationText(text, result);
        } catch (e) {
          // Continue if this div fails
        }
      }

    } catch (error) {
      console.log('⚠️ Technische Daten Extraktion Fehler:', error.message);
    }
  }

  async extractProductDetails(page, result) {
    // Common selectors for product information
    const selectors = {
      title: ['h1', '.product-title', '.title', '[data-testid="product-title"]'],
      description: ['.description', '.product-description', '.details'],
      weight: ['[data-testid="weight"]', '.weight', '[class*="weight"]'],
      dimensions: ['[data-testid="dimensions"]', '.dimensions', '[class*="dimension"]'],
      material: ['.material', '[data-testid="material"]', '[class*="material"]'],
      availability: ['.availability', '.stock', '[data-testid="availability"]']
    };

    for (const [field, selectorList] of Object.entries(selectors)) {
      for (const selector of selectorList) {
        try {
          const element = await page.$(selector);
          if (element) {
            const text = await element.innerText();
            if (text && text.trim()) {
              this.mapProductField(field, text.trim(), result);
              break; // Found content for this field
            }
          }
        } catch (e) {
          // Continue to next selector
        }
      }
    }
  }

  mapTechnicalField(key, value, result) {
    console.log(`🔍 Table-Field Mapping: "${key}" = "${value}"`);
    
    // EXAKTE REIHENFOLGE WIE IN DEINEM PYTHON CODE - Spezifische Felder ZUERST!
    if (key.includes('abmessung') || key.includes('größe') || key.includes('dimension')) {
      result.Abmessung = this.interpretDimensions(value);
      console.log(`✅ Abmessung aus Tabelle zugeordnet: ${value}`);
    } else if (key.includes('gewicht') && !key.includes('einheit')) {
      result.Gewicht = value;
      console.log(`✅ Gewicht aus Tabelle zugeordnet: ${value}`);
    } else if (key.includes('werkstoff') && !key.includes('klassifizierung')) {
      // NUR exakte "werkstoff" Übereinstimmung, NICHT bei "materialklassifizierung"
      result.Werkstoff = value;
      console.log(`✅ Werkstoff aus Tabelle zugeordnet: ${value}`);
    } else if (key.includes('weitere artikelnummer') || key.includes('additional article number') || key.includes('part number')) {
      result["Weitere Artikelnummer"] = value;
      console.log(`✅ Weitere Artikelnummer aus Tabelle zugeordnet: ${value}`);
    } else if (key.includes('materialklassifizierung') || key.includes('material classification')) {
      result.Materialklassifizierung = value;
      console.log(`✅ Materialklassifizierung aus Tabelle zugeordnet: ${value}`);
      if (value.toLowerCase().includes('nicht schweiss')) {
        result["Materialklassifizierung Bewertung"] = "OHNE/N/N/N/N";
      }
    } else if (key.includes('statistische warennummer') || key.includes('statistical') || key.includes('import')) {
      result["Statistische Warennummer"] = value;
      console.log(`✅ Statistische Warennummer aus Tabelle zugeordnet: ${value}`);
    } else if (key.includes('ursprungsland') || key.includes('origin')) {
      result.Ursprungsland = value;
      console.log(`✅ Ursprungsland aus Tabelle zugeordnet: ${value}`);
    } else if (key.includes('verfügbar') || key.includes('stock') || key.includes('lager')) {
      result.Verfügbarkeit = value;
    } else {
      console.log(`❓ Unbekannter Table-Schlüssel: "${key}" = "${value}"`);
    }
  }

  mapProductField(field, value, result) {
    switch (field) {
      case 'title':
        if (!result.Produkttitel || result.Produkttitel === "Nicht gefunden") {
          result.Produkttitel = value;
        }
        break;
      case 'description':
        if (!result.Produktbeschreibung || result.Produktbeschreibung === "Nicht gefunden") {
          result.Produktbeschreibung = value;
        }
        break;
      case 'weight':
        result.Gewicht = value;
        break;
      case 'dimensions':
        result.Abmessung = this.interpretDimensions(value);
        break;
      case 'material':
        result.Werkstoff = value;
        break;
      case 'availability':
        result.Verfügbarkeit = value;
        break;
    }
  }

  parseSpecificationText(text, result) {
    const lines = text.split('\n');
    
    for (const line of lines) {
      const colonIndex = line.indexOf(':');
      if (colonIndex > 0) {
        const key = line.substring(0, colonIndex).trim().toLowerCase();
        const value = line.substring(colonIndex + 1).trim();
        
        if (key && value) {
          this.mapTechnicalField(key, value, result);
        }
      }
    }
  }

  interpretDimensions(text) {
    if (!text) return "Nicht gefunden";
    
    const cleanText = text.replace(/\s+/g, '').toLowerCase();
    
    // Check for diameter x height pattern
    if (cleanText.includes('⌀') || cleanText.includes('ø')) {
      const match = cleanText.match(/[⌀ø]?(\d+)[x×](\d+)/);
      if (match) {
        return `Durchmesser×Höhe: ${match[1]}×${match[2]} mm`;
      }
    }
    
    // Check for L x B x H pattern
    const lbhMatch = cleanText.match(/(\d+)[x×](\d+)[x×](\d+)/);
    if (lbhMatch) {
      return `L×B×H: ${lbhMatch[1]}×${lbhMatch[2]}×${lbhMatch[3]} mm`;
    }
    
    // Check for L x B pattern
    const lbMatch = cleanText.match(/(\d+)[x×](\d+)/);
    if (lbMatch) {
      return `L×B: ${lbMatch[1]}×${lbMatch[2]} mm`;
    }
    
    return text;
  }

  async extractFromInitialData(page, result) {
    try {
      console.log('🔍 Extrahiere Daten aus window.initialData...');
      
      // Extract data from window.initialData JavaScript object
      const productData = await page.evaluate(() => {
        try {
          const initialData = window.initialData;
          if (!initialData || !initialData['product/dataProduct']) {
            return null;
          }
          
          const productInfo = initialData['product/dataProduct'].data.product;
          
          // Extract basic product info
          const extractedData = {
            name: productInfo.name || '',
            description: productInfo.description || '',
            code: productInfo.code || '',
            url: productInfo.url || '',
            technicalSpecs: []
          };
          
          // Extract technical specifications from multiple possible locations
          if (productInfo.localizations && productInfo.localizations.technicalSpecifications) {
            extractedData.technicalSpecs = productInfo.localizations.technicalSpecifications;
          }
          
          // Also extract direct product properties as backup
          extractedData.directProperties = {
            weight: productInfo.weight || '',
            dimensions: productInfo.dimensions || '',
            basicMaterial: productInfo.basicMaterial || '',
            materialClassification: productInfo.materialClassification || '',
            importCodeNumber: productInfo.importCodeNumber || '',
            additionalMaterialNumbers: productInfo.additionalMaterialNumbers || ''
          };
          
          return extractedData;
        } catch (e) {
          console.log('JavaScript extraction error:', e);
          return null;
        }
      });

      if (productData) {
        console.log('✅ Produktdaten aus initialData gefunden');
        
        // Map basic product information
        if (productData.name) {
          result.Produkttitel = productData.name;
        }
        
        if (productData.description) {
          result.Produktbeschreibung = productData.description;
        }
        
        if (productData.url) {
          result.Produktlink = `https://www.mymobase.com${productData.url}`;
        }
        
        // Map technical specifications with improved key matching
        if (productData.technicalSpecs && productData.technicalSpecs.length > 0) {
          // Debug: Show all available keys
          console.log('📋 Alle verfügbaren technische Spezifikationen:');
          productData.technicalSpecs.forEach(spec => {
            console.log(`   "${spec.key}" = "${spec.value}"`);
          });
          productData.technicalSpecs.forEach(spec => {
            const key = spec.key.toLowerCase().trim();
            const value = spec.value;
            
            console.log(`🔍 Mapping spec: "${key}" = "${value}"`);
            
            // KRITISCH: NUR fehlende Felder ergänzen, nicht überschreiben!
            if (key.includes('materialklassifizierung') || key.includes('material classification')) {
              if (!result.Materialklassifizierung || result.Materialklassifizierung === "Nicht gefunden") {
                result.Materialklassifizierung = value;
                console.log(`✅ InitialData Materialklassifizierung ergänzt: ${value}`);
              }
            } else if (key.includes('statistische warennummer') || key.includes('statistical') || key.includes('import')) {
              if (!result['Statistische Warennummer'] || result['Statistische Warennummer'] === "Nicht gefunden") {
                result['Statistische Warennummer'] = value;
                console.log(`✅ InitialData Statistische Warennummer ergänzt: ${value}`);
              }
            } else if (key.includes('weitere artikelnummer') || key.includes('additional material')) {
              if (!result['Weitere Artikelnummer'] || result['Weitere Artikelnummer'] === "Nicht gefunden") {
                result['Weitere Artikelnummer'] = value;
                console.log(`✅ InitialData Weitere Artikelnummer ergänzt: ${value}`);
              }
            } else if (key.includes('abmessungen') || key.includes('dimension')) {
              if (!result.Abmessung || result.Abmessung === "Nicht gefunden") {
                result.Abmessung = value;
                console.log(`✅ InitialData Abmessung ergänzt: ${value}`);
              }
            } else if (key.includes('gewicht') || key.includes('weight')) {
              if (!result.Gewicht || result.Gewicht === "Nicht gefunden") {
                result.Gewicht = value;
                console.log(`✅ InitialData Gewicht ergänzt: ${value}`);
              }
            } else if (key.includes('werkstoff') && !key.includes('klassifizierung')) {
              // NUR ergänzen wenn Werkstoff fehlt
              if (!result.Werkstoff || result.Werkstoff === "Nicht gefunden") {
                result.Werkstoff = value;
                console.log(`✅ InitialData Werkstoff ergänzt: ${value}`);
              }
            } else {
              console.log(`🔄 InitialData Skip: "${key}" = "${value}"`);
            }
          });
        }
        
        // Fallback: Use direct properties if technical specs didn't provide everything
        if (productData.directProperties) {
          if (result.Gewicht === "Nicht gefunden" && productData.directProperties.weight) {
            result.Gewicht = productData.directProperties.weight.toString();
            console.log(`🔄 Fallback Gewicht: ${result.Gewicht}`);
          }
          if (result.Abmessung === "Nicht gefunden" && productData.directProperties.dimensions) {
            result.Abmessung = productData.directProperties.dimensions;
            console.log(`🔄 Fallback Abmessung: ${result.Abmessung}`);
          }
          if (result.Werkstoff === "Nicht gefunden" && productData.directProperties.basicMaterial) {
            result.Werkstoff = productData.directProperties.basicMaterial;
            console.log(`🔄 Fallback Werkstoff: ${result.Werkstoff}`);
          }
          if (result.Materialklassifizierung === "Nicht gefunden" && productData.directProperties.materialClassification) {
            result.Materialklassifizierung = productData.directProperties.materialClassification;
            console.log(`🔄 Fallback Materialklassifizierung: ${result.Materialklassifizierung}`);
          }
          if (result['Statistische Warennummer'] === "Nicht gefunden" && productData.directProperties.importCodeNumber) {
            result['Statistische Warennummer'] = productData.directProperties.importCodeNumber;
            console.log(`🔄 Fallback Statistische Warennummer: ${result['Statistische Warennummer']}`);
          }
          if (result['Weitere Artikelnummer'] === "Nicht gefunden" && productData.directProperties.additionalMaterialNumbers) {
            result['Weitere Artikelnummer'] = productData.directProperties.additionalMaterialNumbers;
            console.log(`🔄 Fallback Weitere Artikelnummer: ${result['Weitere Artikelnummer']}`);
          }
        }
        
        console.log('📊 Extrahierte Daten:', {
          titel: result.Produkttitel,
          weitere_artikelnummer: result['Weitere Artikelnummer'],
          abmessung: result.Abmessung,
          gewicht: result.Gewicht,
          werkstoff: result.Werkstoff,
          materialklassifizierung: result.Materialklassifizierung,
          statistische_warennummer: result['Statistische Warennummer']
        });
        
      } else {
        console.log('⚠️ Keine initialData gefunden, verwende Fallback-Methode');
      }
      
    } catch (error) {
      console.log('⚠️ Fehler bei initialData Extraktion:', error.message);
    }
  }

  async extractFromHTML(page, result) {
    try {
      console.log('🔍 Extrahiere Daten direkt aus HTML DOM...');
      
      // Extract technical specifications from HTML tables/divs
      const htmlData = await page.evaluate(() => {
        const data = {};
        
        // Look for various table structures
        const tables = document.querySelectorAll('table');
        tables.forEach(table => {
          const rows = table.querySelectorAll('tr');
          rows.forEach(row => {
            const cells = row.querySelectorAll('td, th');
            if (cells.length >= 2) {
              const key = cells[0].textContent.trim().toLowerCase();
              const value = cells[1].textContent.trim();
              
              if (key && value && value !== '-' && value !== '') {
                data[key] = value;
              }
            }
          });
        });
        
        // Look for definition lists
        const dls = document.querySelectorAll('dl');
        dls.forEach(dl => {
          const dts = dl.querySelectorAll('dt');
          const dds = dl.querySelectorAll('dd');
          
          for (let i = 0; i < Math.min(dts.length, dds.length); i++) {
            const key = dts[i].textContent.trim().toLowerCase();
            const value = dds[i].textContent.trim();
            
            if (key && value && value !== '-' && value !== '') {
              data[key] = value;
            }
          }
        });
        
        // Look for specific classes or data attributes
        const specElements = document.querySelectorAll('[class*="spec"], [class*="detail"], [data-spec]');
        specElements.forEach(element => {
          const text = element.textContent;
          if (text.includes(':')) {
            const parts = text.split(':');
            if (parts.length >= 2) {
              const key = parts[0].trim().toLowerCase();
              const value = parts.slice(1).join(':').trim();
              if (key && value && value !== '-' && value !== '') {
                data[key] = value;
              }
            }
          }
        });
        
        return data;
      });
      
      console.log('📊 HTML-Extraktion gefunden:', Object.keys(htmlData));
      
      // Map HTML data to result fields
      Object.entries(htmlData).forEach(([key, value]) => {
        if (key.includes('weitere artikelnummer') && !result['Weitere Artikelnummer']) {
          result['Weitere Artikelnummer'] = value;
        } else if (key.includes('abmess') && !result.Abmessung) {
          result.Abmessung = value;
        } else if (key.includes('gewicht') && !result.Gewicht) {
          result.Gewicht = value;
        } else if (key.includes('werkstoff') && !result.Werkstoff) {
          result.Werkstoff = value;
        } else if (key.includes('materialklassifizierung') && !result.Materialklassifizierung) {
          result.Materialklassifizierung = value;
        } else if (key.includes('statistische warennummer') && !result['Statistische Warennummer']) {
          result['Statistische Warennummer'] = value;
        }
      });
      
    } catch (error) {
      console.log('⚠️ Fehler bei HTML-Extraktion:', error.message);
    }
  }

  interpretMaterialClassification(classification) {
    if (!classification) return "Nicht bewertet";
    
    const lower = classification.toLowerCase();
    
    if (lower.includes('nicht schweiss') && lower.includes('guss') && lower.includes('klebe') && lower.includes('schmiede')) {
      return "OHNE/N/N/N/N - Material ist nicht schweißbar, gießbar, klebbar oder schmiedbar";
    }
    
    if (lower.includes('nicht schweiss')) {
      return "Nicht schweißbar - Material kann nicht geschweißt werden";
    }
    
    return classification; // Return original if no specific interpretation found
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
