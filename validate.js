#!/usr/bin/env node

/**
 * Spanish Language Conversion Validator (Strict)
 * Validates that the Spanish page contains distinctive Spanish words.
 * Clears browser cache and performs Ctrl+F5 hard refresh after navigation.
 * Outputs an Excel report in the "reports/" folder with a timestamp.
 *
 * Usage:
 *   node validate.js --file urls.csv [--debug]
 *   node validate.js https://url1.com https://url2.com [--debug]
 */

const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// ============================================================
//  CONFIGURATION
// ============================================================
const SAMPLE_COUNT = 8;
const REPORTS_DIR = path.join(__dirname, 'reports');

// DISTINCTIVE SPANISH WORDS (unlikely to appear in English text)
const SPANISH_WORDS = [
  'protegemos', 'vehículo', 'propiedad', 'negocios', 'inversiones',
  'reclamos', 'factura', 'código', 'cotización', 'explorar',
  'financieros', 'jubilación', 'solicitar', 'póliza', 'temporal',
  'personas', 'empresas', 'mascotas', 'eventos', 'empleados',
  'bienes', 'sueños', 'necesidades', 'ahorro', 'hogar',
  'viaje', 'salud', 'seguro', 'cobertura', 'tarifa',
  'analiza', 'ingresar', 'comenzar', 'buscar', 'protección',
  'disposición', 'beneficios', 'miembros',
  'peyton manning', 'mucho más', 'hace falta', 'iniciar sesión',
  'gustaría hacer', 'paquete', 'código postal', 'comenzar la cotización',
  'centro de información', 'recursos de seguro', 'pequeñas empresas',
  'deportes de motor', 'planificación de emergencias',
  'agricultura y agroindustria', 'centro de recursos cibernéticos',
  'encontrar un profesional financiero',
  'preguntas frecuentes sobre inversiones', 'finanzas a nivel nacional',
  'ahora desde nationwide', 'el blog advisor advocate', 'agencia forward'
];

if (!fs.existsSync(REPORTS_DIR)) {
  fs.mkdirSync(REPORTS_DIR, { recursive: true });
}

// ============================================================
//  HELP / USAGE
// ============================================================
function showHelp() {
  console.log(`
Spanish Language Conversion Validator

Usage:
  node validate.js --file <filename> [--debug]
  node validate.js <url1> <url2> ... [--debug]

Options:
  --file <filename>   Read URLs from CSV or plain text (one per line)
  --debug             Run browser in non‑headless mode (visible)
  --help              Show this help

Output: Excel report saved in "reports/spanish_conversion_report_<timestamp>.xlsx"
`);
  process.exit(0);
}

// ============================================================
//  INPUT PARSING
// ============================================================
function parseArguments() {
  const args = process.argv.slice(2);
  if (args.includes('--help')) showHelp();

  const urls = [];
  let inputFile = null;
  let debug = false;

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg === '--file') {
      inputFile = args[++i];
      if (!inputFile) {
        console.error('❌ Missing filename after --file');
        process.exit(1);
      }
    } else if (arg === '--debug') {
      debug = true;
    } else if (!arg.startsWith('--')) {
      urls.push(arg);
    }
  }

  let urlList = [];
  if (inputFile) {
    urlList = readUrlsFromFile(inputFile);
  } else if (urls.length > 0) {
    urlList = urls;
  } else {
    console.warn('⚠️  No input provided. Using default fallback URLs.');
    urlList = [
      'https://www.example.com/page1',
      'https://www.example.com/page2',
      'https://www.example.com/page3'
    ];
  }

  return { urls: urlList, debug };
}

// ============================================================
//  URL READERS
// ============================================================
function readUrlsFromFile(filePath) {
  if (!fs.existsSync(filePath)) {
    console.error(`❌ File not found: ${filePath}`);
    process.exit(1);
  }

  const ext = path.extname(filePath).toLowerCase();
  let lines;
  try {
    const content = fs.readFileSync(filePath, 'utf-8');
    lines = content.split('\n').filter(line => line.trim().length > 0);
  } catch (err) {
    console.error(`❌ Error reading file: ${err.message}`);
    process.exit(1);
  }

  if (ext === '.csv') {
    return parseCsvUrls(lines);
  } else {
    return lines
      .map(line => line.replace(/^#.*$/, '').trim())
      .filter(u => u.length > 0);
  }
}

function parseCsvUrls(lines) {
  function parseRow(row) {
    const result = [];
    let current = '';
    let inQuotes = false;
    for (let i = 0; i < row.length; i++) {
      const ch = row[i];
      if (ch === '"') {
        inQuotes = !inQuotes;
      } else if (ch === ',' && !inQuotes) {
        result.push(current.trim());
        current = '';
      } else {
        current += ch;
      }
    }
    result.push(current.trim());
    return result;
  }

  const rows = lines.map(parseRow);
  if (rows.length === 0) return [];

  const header = rows[0].map(h => h.toLowerCase());
  const urlColIndex = header.findIndex(h =>
    ['url', 'link', 'website', 'site', 'address'].includes(h)
  );

  const startRow = (urlColIndex !== -1) ? 1 : 0;
  const colIndex = (urlColIndex !== -1) ? urlColIndex : 0;

  const urls = [];
  for (let i = startRow; i < rows.length; i++) {
    const row = rows[i];
    if (row.length > colIndex) {
      let url = row[colIndex].replace(/^"|"$/g, '').trim();
      if (url) urls.push(url);
    }
  }
  return urls;
}

// ============================================================
//  ROBUST "En Español" LINK FINDER
// ============================================================
async function findEnEspanolLink(page) {
  const footerSelectors = [
    'footer', '.footer', '#footer', 'div[class*="footer"]',
    'div[class*="Footer"]', '.site-footer', '.footer-container',
    '[role="contentinfo"]', '.nav-footer', '.nw-footer', '.global-footer'
  ];

  for (const sel of footerSelectors) {
    try {
      const footer = await page.$(sel);
      if (footer) {
        const linkHandle = await page.evaluateHandle((footerEl) => {
          const links = footerEl.querySelectorAll('a');
          for (const a of links) {
            const text = a.textContent.trim().toLowerCase();
            if (text.includes('en español') || text === 'español' || text.includes('espanol')) {
              return a;
            }
          }
          return null;
        }, footer);
        const isNull = await page.evaluate(el => el === null, linkHandle);
        if (!isNull) {
          return linkHandle;
        }
      }
    } catch (_) {}
  }

  console.warn('⚠️  No footer found or link not in footer; scanning whole page.');
  const linkHandle = await page.evaluateHandle(() => {
    const links = document.querySelectorAll('a');
    for (const a of links) {
      const text = a.textContent.trim().toLowerCase();
      if (text.includes('en español') || text === 'español' || text.includes('espanol')) {
        return a;
      }
    }
    return null;
  });

  const isNull = await page.evaluate(el => el === null, linkHandle);
  if (isNull) return null;
  return linkHandle;
}

// ============================================================
//  CLICKABLE CHECK
// ============================================================
async function isClickable(page, elementHandle) {
  if (!elementHandle) return { clickable: false, reason: 'Element handle is null' };
  try {
    const result = await page.evaluate(el => {
      const style = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      const visible = style.display !== 'none' &&
                      style.visibility !== 'hidden' &&
                      el.offsetParent !== null &&
                      rect.width > 0 && rect.height > 0;
      const enabled = !el.disabled;
      const pointerEvents = style.pointerEvents !== 'none';
      return { visible, enabled, pointerEvents, rect: { width: rect.width, height: rect.height } };
    }, elementHandle);
    if (!result.visible) return { clickable: false, reason: 'Element not visible' };
    if (!result.enabled) return { clickable: false, reason: 'Element disabled' };
    if (!result.pointerEvents) return { clickable: false, reason: 'pointer-events: none' };
    if (result.rect.width === 0 || result.rect.height === 0)
      return { clickable: false, reason: 'Zero size' };
    return { clickable: true, reason: '' };
  } catch (e) {
    return { clickable: false, reason: `Error: ${e.message}` };
  }
}

// ============================================================
//  URL PATTERN VALIDATION
// ============================================================
function isValidSpanishUrl(url) {
  return /^https:\/\/espanol\..*/.test(url) || /\/espanol/.test(url);
}

// ============================================================
//  CLEAR BROWSER CACHE (using CDP)
// ============================================================
async function clearBrowserCache(page) {
  try {
    const client = await page.target().createCDPSession();
    await client.send('Network.clearBrowserCache');
    console.log('   ✅ Browser cache cleared.');
  } catch (err) {
    console.warn(`   ⚠️  Could not clear cache: ${err.message}`);
  }
}

// ============================================================
//  EXTRACT MAIN CONTENT
// ============================================================
async function extractMainContent(page, count = SAMPLE_COUNT, debug = false) {
  console.log('   Waiting for app-root...');
  try {
    await page.waitForSelector('app-root', { timeout: 15000 });
    console.log('   ✅ app-root found.');
  } catch (_) {
    console.warn('   ⚠️  app-root not found, falling back to body.');
  }

  console.log('   Waiting for content to load...');
  try {
    await page.waitForFunction(
      () => {
        const root = document.querySelector('app-root');
        if (!root) return false;
        return root.textContent.trim().length > 100;
      },
      { timeout: 20000 }
    );
    console.log('   ✅ Content loaded.');
  } catch (_) {
    console.warn('   ⚠️  Content not loaded after 20s, proceeding anyway.');
  }

  await page.evaluate(async () => {
    await new Promise((resolve) => {
      let totalHeight = 0;
      const distance = 100;
      const timer = setInterval(() => {
        const scrollHeight = document.body.scrollHeight;
        window.scrollBy(0, distance);
        totalHeight += distance;
        if (totalHeight >= scrollHeight) {
          clearInterval(timer);
          resolve();
        }
      }, 100);
    });
  });
  await new Promise(r => setTimeout(r, 3000));

  const texts = await page.evaluate((c, isDebug) => {
    function getClassString(el) {
      if (!el) return '';
      if (typeof el.className === 'string') return el.className;
      if (el.className && typeof el.className.baseVal === 'string') return el.className.baseVal;
      return String(el.className || '');
    }

    function getCleanText(el) {
      const clone = el.cloneNode(true);
      const remove = clone.querySelectorAll('style, script, noscript, link, meta, svg, button, input, select, textarea');
      remove.forEach(el => el.remove());
      return clone.textContent.trim();
    }

    function isUIElement(el) {
      const tag = el.tagName.toLowerCase();
      if (['header', 'footer', 'nav', 'aside'].includes(tag)) return true;
      const id = (el.id || '').toLowerCase();
      const cls = getClassString(el).toLowerCase();
      const uiPatterns = ['nav', 'menu', 'sidebar', 'breadcrumb', 'toolbar', 'topbar', 'header', 'footer'];
      for (const p of uiPatterns) {
        if (id.includes(p) || cls.includes(p)) return true;
      }
      let parent = el.parentElement;
      while (parent) {
        const pid = (parent.id || '').toLowerCase();
        const pcls = getClassString(parent).toLowerCase();
        for (const p of uiPatterns) {
          if (pid.includes(p) || pcls.includes(p)) return true;
        }
        parent = parent.parentElement;
      }
      return false;
    }

    const container = document.querySelector('app-root') || document.body;
    const all = container.querySelectorAll('*');
    const candidates = [];
    for (const el of all) {
      if (isUIElement(el)) continue;
      const text = getCleanText(el);
      if (text.length > 50) {
        candidates.push(text);
      }
    }

    candidates.sort((a, b) => b.length - a.length);

    const samples = [];
    const seen = new Set();
    for (const text of candidates) {
      if (samples.length >= c) break;
      if (!seen.has(text)) {
        seen.add(text);
        samples.push(text);
      }
    }

    if (samples.length < c) {
      const bodyText = getCleanText(container);
      const chunks = bodyText.split(/\n\s*\n|\.\s+|\?\s+/).filter(s => s.length > 40);
      for (const chunk of chunks) {
        if (samples.length >= c) break;
        if (!seen.has(chunk)) {
          seen.add(chunk);
          samples.push(chunk);
        }
      }
    }

    if (isDebug && samples.length > 0) {
      console.log('   Debug: first sample (first 150 chars):', samples[0].substring(0, 150));
    }
    return samples.slice(0, c);
  }, count, debug);

  return texts;
}

// ============================================================
//  CHECK FOR SPANISH WORDS
// ============================================================
function hasSpanishContent(samples) {
  if (!samples || samples.length === 0) {
    return { found: false, details: 'No samples extracted' };
  }

  const foundWords = [];
  for (const sample of samples) {
    const lower = sample.toLowerCase();
    for (const word of SPANISH_WORDS) {
      if (lower.includes(word)) {
        if (!foundWords.includes(word)) {
          foundWords.push(word);
        }
      }
    }
  }

  const hasSpanish = foundWords.length > 0;

  return {
    found: hasSpanish,
    foundWords: foundWords.slice(0, 5),
    totalFound: foundWords.length,
    samplesAnalyzed: samples.length
  };
}

// ============================================================
//  CLICK + NAVIGATION HANDLER
// ============================================================
async function clickAndWaitForLanguagePage(page, linkHandle, debug = false) {
  const newPagePromise = new Promise((resolve) => {
    const timeout = setTimeout(() => { resolve(null); }, 2000);
    const browser = page.browser();
    const onTarget = async (target) => {
      if (target.type() === 'page') {
        const newPage = await target.page();
        if (newPage && newPage !== page) {
          clearTimeout(timeout);
          browser.off('targetcreated', onTarget);
          resolve(newPage);
        }
      }
    };
    browser.on('targetcreated', onTarget);
    setTimeout(() => {
      browser.off('targetcreated', onTarget);
      resolve(null);
    }, 3000);
  });

  console.log('   Clicking link...');
  await linkHandle.click();

  const navigationPromise = page.waitForNavigation({
    waitUntil: 'networkidle2',
    timeout: 10000
  }).catch(() => null);

  const initialUrl = page.url();
  const urlChangePromise = page.waitForFunction(
    (url) => window.location.href !== url,
    { timeout: 10000 },
    initialUrl
  ).catch(() => null);

  const result = await Promise.race([
    navigationPromise.then(() => ({ type: 'navigation', page: page })),
    urlChangePromise.then(() => ({ type: 'urlchange', page: page })),
    newPagePromise.then((newPage) => newPage ? ({ type: 'newpage', page: newPage }) : null)
  ]);

  if (result) {
    console.log(`   Navigation detected via ${result.type}`);
    return result.page;
  }

  console.log('   No navigation detected, checking content...');
  const samples = await extractMainContent(page, 5);
  const result2 = hasSpanishContent(samples);
  if (result2.found) {
    console.log('   Content contains Spanish words (no navigation).');
    return page;
  }

  console.log('   No successful navigation or Spanish content detected.');
  return null;
}

// ============================================================
//  MAIN VALIDATION FUNCTION
// ============================================================
async function validateSpanishConversion(url, debug = false) {
  const result = {
    url,
    elementFound: 'Not Found',
    linkValid: 'N/A',
    linkHref: '',
    spanishTranslate: 'Fail',
    status: 'FAIL',
    pageError: 'N/A',
    details: ''
  };

  const browser = await puppeteer.launch({
    headless: !debug,
    slowMo: debug ? 50 : 0,
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });
  let page;
  try {
    page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 800 });

    // ---- CLEAR BROWSER CACHE BEFORE NAVIGATION ----
    console.log('   Clearing browser cache...');
    await clearBrowserCache(page);

    console.log(`   Navigating to ${url} ...`);
    let response;
    try {
      // Also add cache-busting query param or use ignoreCache flag
      response = await page.goto(url, {
        waitUntil: 'networkidle2',
        timeout: 30000
      });
      if (!response) {
        throw new Error('No response received');
      }
      if (!response.ok()) {
        result.pageError = `HTTP ${response.status()} ${response.statusText()}`;
      }
    } catch (err) {
      result.pageError = err.message;
      result.details = `Page error: ${err.message}`;
      result.status = 'FAIL';
      return result;
    }
    console.log('   Page loaded.');

    console.log('   Searching for "En Español" link...');
    const linkHandle = await findEnEspanolLink(page);
    if (!linkHandle) {
      console.log('   ❌ Link NOT found.');
      result.details = 'En Español link missing, hidden, disabled, or has no href';
      result.status = 'FAIL';
      return result;
    }
    result.elementFound = 'Found';
    const href = await page.evaluate(el => el.href || el.getAttribute('href') || '', linkHandle);
    result.linkHref = href;
    console.log(`   ✅ Link found: ${href}`);

    console.log('   Checking if link is clickable...');
    const clickInfo = await isClickable(page, linkHandle);
    if (!clickInfo.clickable) {
      console.log(`   ❌ Link not clickable: ${clickInfo.reason}`);
      try {
        await page.evaluate(el => el.click(), linkHandle);
        result.linkValid = 'Force Clicked';
      } catch (e) {
        result.linkValid = 'Invalid';
        result.details = 'En Español link missing, hidden, disabled, or has no href';
        result.status = 'FAIL';
        return result;
      }
    } else {
      console.log('   ✅ Link is clickable.');
      result.linkValid = 'Valid';
    }

    const targetPage = await clickAndWaitForLanguagePage(page, linkHandle, debug);
    if (!targetPage) {
      result.linkValid = 'Invalid';
      result.details = 'Click did not lead to navigation or content change';
      result.status = 'FAIL';
      return result;
    }

    // ---- HARD REFRESH USING Ctrl+F5 ----
    console.log('   Performing hard refresh (Ctrl+F5)...');
    try {
      await targetPage.keyboard.down('Control');
      await targetPage.keyboard.press('F5');
      await targetPage.keyboard.up('Control');
      await targetPage.waitForNavigation({ waitUntil: 'networkidle2', timeout: 10000 });
      await new Promise(r => setTimeout(r, 2000));
      console.log('   ✅ Hard refresh completed.');
    } catch (refreshError) {
      console.warn(`   ⚠️  Hard refresh failed: ${refreshError.message}, continuing anyway.`);
    }

    const currentPage = targetPage;
    const newUrl = currentPage.url();
    console.log(`   Current URL: ${newUrl}`);

    if (!isValidSpanishUrl(newUrl)) {
      console.log(`   ❌ URL pattern mismatch.`);
      result.linkValid = 'Invalid';
      result.details = 'Spanish translation not detected (URL pattern mismatch)';
      result.status = 'FAIL';
      return result;
    } else {
      console.log('   ✅ URL pattern matches.');
    }

    console.log('   Extracting main content...');
    const samples = await extractMainContent(currentPage, SAMPLE_COUNT, debug);
    console.log(`   Extracted ${samples.length} samples.`);
    if (samples.length === 0) {
      result.details = 'Spanish translation not detected';
      result.status = 'FAIL';
      return result;
    }

    const spanishCheck = hasSpanishContent(samples);
    console.log(`   Found ${spanishCheck.totalFound} Spanish word(s): ${spanishCheck.foundWords.join(', ')}`);

    if (spanishCheck.found) {
      console.log('   ✅ PASS – Spanish content found.');
      result.spanishTranslate = 'Pass';
      result.details = 'Spanish translation successfully validated.';
      result.status = 'PASS';
    } else {
      console.log('   ❌ FAIL – No Spanish content found.');
      result.spanishTranslate = 'Fail';
      result.details = 'Spanish translation not detected';
      result.status = 'FAIL';
    }

  } catch (error) {
    console.log(`   ❌ Unexpected error: ${error.message}`);
    result.details = `Unexpected error: ${error.message}`;
    result.status = 'FAIL';
    if (!result.pageError || result.pageError === 'N/A') {
      result.pageError = error.message;
    }
  } finally {
    await browser.close();
  }

  return result;
}

// ============================================================
//  RUN VALIDATION (saves Excel report with timestamp)
// ============================================================
async function runValidation(urls, debug = false) {
  const results = [];
  for (const url of urls) {
    console.log(`\n🔍 Processing ${url} ...`);
    const res = await validateSpanishConversion(url, debug);
    results.push(res);
    await new Promise(r => setTimeout(r, 1500));
  }

  const data = results.map(r => ({
    'URL': r.url,
    'En Español Element': r.elementFound,
    'En Español Link': r.linkHref || r.linkValid,
    'Spanish Translate': r.spanishTranslate,
    'Status': r.status,
    'Page Error': r.pageError,
    'Validation Details': r.details
  }));

  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Validation');

  const now = new Date();
  const ts = now.getFullYear() + '-' +
             String(now.getMonth() + 1).padStart(2, '0') + '-' +
             String(now.getDate()).padStart(2, '0') + '_' +
             String(now.getHours()).padStart(2, '0') + '-' +
             String(now.getMinutes()).padStart(2, '0') + '-' +
             String(now.getSeconds()).padStart(2, '0');
  const excelPath = path.join(REPORTS_DIR, `spanish_conversion_report_${ts}.xlsx`);
  XLSX.writeFile(wb, excelPath);
  console.log(`\n✅ Excel report written to ${excelPath}`);
}

// ============================================================
//  MAIN
// ============================================================
async function main() {
  const { urls, debug } = parseArguments();
  console.log(`📋 Validating ${urls.length} URL(s):`);
  urls.forEach((u, i) => console.log(`  ${i+1}. ${u}`));
  if (debug) console.log('🐞 Debug mode: browser will be visible\n');
  await runValidation(urls, debug);
}

main().catch(err => {
  console.error('❌ Fatal error:', err);
  process.exit(1);
});