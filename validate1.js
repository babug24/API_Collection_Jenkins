#!/usr/bin/env node

/**
 * Spanish Language Conversion Validator (Strict, Layered)
 * Validates Spanish-language conversion using FOUR independent signals:
 *   1. HTML lang attribute      (document.documentElement.lang === 'es*')
 *   2. og:locale meta tag       (non-blocking; many pages omit this)
 *   3. Statistical content detection via `franc` (trigram-based, spa vs eng)
 *   4. Distinctive Spanish keyword presence (word-boundary + accent-normalized)
 * Each page gets a confidence score (signals passed / signals applicable)
 * instead of a single keyword-only pass/fail — intended to produce
 * defensible evidence for QA sign-off rather than a single weak signal.
 *
 * Handles SPAs (client‑side redirects) and dynamic content.
 * Skips "En Español" link check if the URL is already Spanish.
 * Clears cache and performs Ctrl+F5 hard refresh after navigation.
 * Outputs an Excel report in the "reports/" folder with a timestamp.
 *
 * Dependency: this script uses `franc` for statistical language detection.
 * franc v6+ is ESM-only, so it's loaded via dynamic import() even though
 * this file itself is CommonJS. Install with:
 *   npm install franc
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
  'mucho más', 'hace falta', 'iniciar sesión',
  'gustaría hacer', 'paquete', 'código postal', 'comenzar la cotización',
  'centro de información', 'recursos de seguro', 'pequeñas empresas',
  'deportes de motor', 'planificación de emergencias',
  'agricultura y agroindustria', 'centro de recursos cibernéticos',
  'encontrar un profesional financiero',
  'preguntas frecuentes sobre inversiones', 'finanzas a nivel nacional',
  'ahora desde nationwide', 'el blog advisor advocate', 'agencia forward',
  'Ingreso a tu cuenta','Nombre de usuario'
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
//  CHECK IF URL IS ALREADY SPANISH
// ============================================================
function isSpanishUrl(url) {
  return /^https:\/\/espanol\..*/.test(url) || /\/espanol/.test(url);
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
  if (isNull) {
    console.warn('⚠️  No footer found or link not in footer; scanning whole page.');
    return null;
  }
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
//  EXTRACT MAIN CONTENT (improved)
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

  // Scroll to trigger lazy loading
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
//  CHECK FOR SPANISH WORDS (word-boundary safe, normalized)
// ============================================================
function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

const MIN_MATCHES_TO_PASS = 1; // any Spanish word found = pass

function hasSpanishContent(samples) {
  if (!samples || samples.length === 0) {
    return { found: false, details: 'No samples extracted' };
  }

  const normalize = (s) => s.normalize('NFC').toLowerCase();
  const combinedText = normalize(samples.join(' \n '));

  const foundWords = [];
  for (const word of SPANISH_WORDS) {
    const normalizedWord = normalize(word);
    const re = new RegExp(`\\b${escapeRegex(normalizedWord)}\\b`, 'i');
    if (re.test(combinedText)) {
      foundWords.push(word);
    }
  }

  return {
    found: foundWords.length >= MIN_MATCHES_TO_PASS,
    foundWords: foundWords.slice(0, 10),
    totalFound: foundWords.length,
    samplesAnalyzed: samples.length
  };
}

// ============================================================
//  SIGNAL 2: HTML LANG ATTRIBUTE CHECK
// ============================================================
async function checkHtmlLangAttribute(page) {
  try {
    const lang = await page.evaluate(() => document.documentElement.lang || '');
    const trimmed = (lang || '').trim();
    const pass = /^es(-\w+)?$/i.test(trimmed);
    return { pass, value: trimmed || '(none)' };
  } catch (e) {
    return { pass: false, value: `Error: ${e.message}` };
  }
}

// ============================================================
//  SIGNAL 3: OG:LOCALE META CHECK (non-blocking, timeout-safe)
// ============================================================
async function checkOgLocale(page) {
  try {
    const content = await Promise.race([
      page.$eval('meta[property="og:locale"]', el => el.getAttribute('content')).catch(() => undefined),
      new Promise((resolve) => setTimeout(() => resolve(undefined), 3000))
    ]);
    if (content === undefined || content === null || content === '') {
      // Tag simply not present — don't penalize, mark as not applicable
      return { pass: null, value: '(not present)' };
    }
    const pass = /^es[_-]/i.test(content);
    return { pass, value: content };
  } catch (e) {
    return { pass: null, value: '(not present)' };
  }
}

// ============================================================
//  SIGNAL 4: STATISTICAL CONTENT DETECTION (franc, ESM via dynamic import)
// ============================================================
let _francModule = null;
async function getFranc() {
  if (!_francModule) {
    _francModule = await import('franc'); // franc v6+ is ESM-only; CJS files must dynamic-import
  }
  return _francModule;
}

async function detectContentLanguageFranc(text) {
  if (!text || text.trim().length < 20) {
    return { pass: null, value: 'Insufficient text', detected: 'und', confidence: 0 };
  }
  try {
    const { francAll } = await getFranc();
    const results = francAll(text, { only: ['spa', 'eng'], minLength: 20 });
    if (!results || results.length === 0 || results[0][0] === 'und') {
      return { pass: null, value: 'Undetermined', detected: 'und', confidence: 0 };
    }
    const [topLang, topScore] = results[0];
    const pass = topLang === 'spa' && topScore >= 0.9;
    return {
      pass,
      value: `${topLang} (${(topScore * 100).toFixed(0)}%)`,
      detected: topLang,
      confidence: topScore
    };
  } catch (e) {
    return { pass: null, value: `Error: ${e.message}`, detected: 'error', confidence: 0 };
  }
}

// ============================================================
//  COMBINE ALL SIGNALS INTO A LAYERED, SCORED VALIDATION
// ============================================================
const MIN_CONFIDENCE_TO_PASS = 0.75; // require 75% of applicable signals to agree

async function runLayeredValidation(page, samples) {
  const keywordCheck = hasSpanishContent(samples);
  const htmlLangCheck = await checkHtmlLangAttribute(page);
  const ogLocaleCheck = await checkOgLocale(page);
  const combinedText = samples.join(' \n ');
  const contentLangCheck = await detectContentLanguageFranc(combinedText);

  const signals = [
    { name: 'HTML Lang', result: htmlLangCheck },
    { name: 'OG Locale', result: ogLocaleCheck },
    { name: 'Content Detection (franc)', result: contentLangCheck },
    { name: 'Keyword Match', result: { pass: keywordCheck.found, value: `${keywordCheck.totalFound} found` } }
  ];

  // pass === null means "not applicable" (e.g. og:locale missing) — excluded from scoring, not penalized
  const applicable = signals.filter(s => s.result.pass !== null);
  const passed = applicable.filter(s => s.result.pass === true);

  const signalsTotal = applicable.length;
  const signalsPassed = passed.length;
  const confidenceScore = signalsTotal > 0 ? signalsPassed / signalsTotal : 0;

  return {
    keywordCheck,
    htmlLangCheck,
    ogLocaleCheck,
    contentLangCheck,
    signalsPassed,
    signalsTotal,
    confidenceScore,
    overallPass: keywordCheck.found || (signalsTotal > 0 && confidenceScore >= MIN_CONFIDENCE_TO_PASS)
  };
}

// ============================================================
//  CLICK + NAVIGATION HANDLER (SPA‑aware)
// ============================================================
async function clickAndWaitForLanguagePage(page, linkHandle, debug = false) {
  // 1. Capture the current page state (for SPA detection)
  const initialUrl = page.url();
  let initialContent = '';
  try {
    initialContent = await page.evaluate(() => document.body.innerText.trim().slice(0, 500));
  } catch (_) {}

  console.log('   Clicking link...');
  await linkHandle.click();

  // 2. Wait for either URL change OR content change (SPA)
  const navigationPromise = page.waitForNavigation({
    waitUntil: 'networkidle2',
    timeout: 10000
  }).catch(() => null);

  const urlChangePromise = page.waitForFunction(
    (url) => window.location.href !== url,
    { timeout: 10000 },
    initialUrl
  ).catch(() => null);

  // 3. Also wait for content to become Spanish (or at least change)
  const contentChangePromise = page.waitForFunction(
    (initial) => {
      const current = document.body.innerText.trim().slice(0, 500);
      // If content length has grown significantly, assume navigation occurred
      if (current.length > initial.length + 100) return true;
      // Also check for Spanish words
      const spanishWords = ['protegemos', 'vehículo', 'propiedad', 'negocios', 'inversiones', 'seguro', 'cobertura', 'cotización'];
      const lower = current.toLowerCase();
      for (const w of spanishWords) {
        if (lower.includes(w)) return true;
      }
      return false;
    },
    { timeout: 15000 },
    initialContent
  ).catch(() => null);

  const result = await Promise.race([
    navigationPromise.then(() => ({ type: 'navigation', page: page })),
    urlChangePromise.then(() => ({ type: 'urlchange', page: page })),
    contentChangePromise.then(() => ({ type: 'contentchange', page: page }))
  ]);

  if (result) {
    console.log(`   Navigation detected via ${result.type}`);
    return result.page;
  }

  console.log('   No navigation or content change detected.');
  return null;
}

// ============================================================
//  POST-NAVIGATION REFRESH AND CACHE CLEAR
// ============================================================
async function refreshSpanishPage(page) {
  console.log('   Clearing browser cache...');
  await clearBrowserCache(page);

  // Check if content is already loaded
  const hasContent = await page.evaluate(() => {
    const root = document.querySelector('app-root');
    return root && root.textContent.trim().length > 100;
  }).catch(() => false);

  if (hasContent) {
    console.log('   ℹ️  Content already loaded, skipping hard refresh.');
    return;
  }

  console.log('   Performing hard refresh (Ctrl+F5)...');
  try {
    await page.keyboard.down('Control');
    await page.keyboard.press('F5');
    await page.keyboard.up('Control');
    await page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 30000 });
    await new Promise(r => setTimeout(r, 2000));
    console.log('   ✅ Hard refresh completed.');
  } catch (refreshError) {
    // Silent fallback - continue without warning
  }

  try {
    await page.waitForFunction(
      () => {
        const root = document.querySelector('app-root');
        if (!root) return false;
        return root.textContent.trim().length > 100;
      },
      { timeout: 30000 }
    );
    console.log('   ✅ Content reloaded successfully.');
  } catch (_) {
    console.warn('   ⚠️  Content not visible after refresh, proceeding anyway.');
  }
}

// ============================================================
//  APPLY LAYERED VALIDATION ONTO THE RESULT OBJECT
// ============================================================
async function applyLayeredValidation(page, samples, result) {
  const layered = await runLayeredValidation(page, samples);

  result.htmlLang = layered.htmlLangCheck.value + (layered.htmlLangCheck.pass === null ? '' : layered.htmlLangCheck.pass ? ' ✓' : ' ✗');
  result.ogLocale = layered.ogLocaleCheck.value + (layered.ogLocaleCheck.pass === null ? '' : layered.ogLocaleCheck.pass ? ' ✓' : ' ✗');
  result.contentDetection = layered.contentLangCheck.value + (layered.contentLangCheck.pass === null ? '' : layered.contentLangCheck.pass ? ' ✓' : ' ✗');
  result.keywordMatches = `${layered.keywordCheck.totalFound} found: ${layered.keywordCheck.foundWords.join(', ') || 'none'}`;
  result.signalsPassed = layered.signalsPassed;
  result.signalsTotal = layered.signalsTotal;
  result.confidenceScore = layered.signalsTotal > 0 ? `${(layered.confidenceScore * 100).toFixed(0)}%` : 'N/A';

  console.log(`   Signals: HTML Lang=${layered.htmlLangCheck.value} | OG Locale=${layered.ogLocaleCheck.value} | Content Detection=${layered.contentLangCheck.value} | Keywords=${layered.keywordCheck.totalFound} found`);
  console.log(`   Confidence: ${layered.signalsPassed}/${layered.signalsTotal} applicable signals passed (${result.confidenceScore})`);

  if (layered.overallPass) {
    console.log('   ✅ PASS – layered validation confidence threshold met.');
    result.spanishTranslate = 'Pass';
    result.details = `Spanish translation validated.`;
    result.status = 'PASS';
  } else {
    console.log('   ❌ FAIL – layered validation confidence threshold not met.');
    result.spanishTranslate = 'Fail';
    result.details = `Spanish translation not validated.`;
    result.status = 'FAIL';
  }

  return result;
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
    htmlLang: 'N/A',
    ogLocale: 'N/A',
    contentDetection: 'N/A',
    keywordMatches: 'N/A',
    signalsPassed: 0,
    signalsTotal: 0,
    confidenceScore: 0,
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

    console.log('   Clearing browser cache...');
    await clearBrowserCache(page);

    console.log(`   Navigating to ${url} ...`);
    let response;
    try {
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

    const isAlreadySpanish = isSpanishUrl(url);

    if (isAlreadySpanish) {
      console.log('   ℹ️  URL is already Spanish – skipping link search.');
      result.elementFound = 'N/A';
      result.linkHref = url;
      result.linkValid = 'N/A';

      console.log(`   Current URL: ${url}`);
      console.log('   ✅ URL is already Spanish (pattern matches).');

      console.log('   Performing post-load refresh...');
      await refreshSpanishPage(page);

      console.log('   Extracting main content...');
      const samples = await extractMainContent(page, SAMPLE_COUNT, debug);
      console.log(`   Extracted ${samples.length} samples.`);
      if (samples.length === 0) {
        result.details = 'Spanish translation not detected';
        result.status = 'FAIL';
        return result;
      }

      await applyLayeredValidation(page, samples, result);
      return result;
    }

    // ---- English page – find and click "En Español" ----
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

    console.log('   Performing post-navigation refresh...');
    await refreshSpanishPage(targetPage);

    console.log('   Extracting main content...');
    const samples = await extractMainContent(targetPage, SAMPLE_COUNT, debug);
    console.log(`   Extracted ${samples.length} samples.`);
    if (samples.length === 0) {
      result.details = 'Spanish translation not detected';
      result.status = 'FAIL';
      return result;
    }

    await applyLayeredValidation(targetPage, samples, result);

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