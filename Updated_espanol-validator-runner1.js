const { chromium } = require('playwright');
const { mergeConfig, dismissCookieBanner, findEnEspanolLink, validateSpanishTranslation, takeScreenshot } = require('./lib/espanol-validator-core');
const path = require('path');
const fs = require('fs');
const { parse } = require('csv-parse/sync');
const franc = require('franc');

const DEFAULT_URL_CSV = path.join(__dirname, 'url.csv');
const config = mergeConfig({});

// ==================== Config Parsing ====================

function parseCommandLineArgs() {
  const args = {
    csvFile: null,
    batchSize: 10,        // default batch size for concurrent URLs
    limit: null,           // max number of URLs to process (null = all)
    verbose: process.argv.includes('--verbose') || process.argv.includes('-v')
  };

  for (let i = 2; i < process.argv.length; i++) {
    if ((process.argv[i] === '--file' || process.argv[i] === '-f') && i + 1 < process.argv.length) {
      args.csvFile = process.argv[i + 1];
      i++;
    } else if ((process.argv[i] === '--batch-size' || process.argv[i] === '-b') && i + 1 < process.argv.length) {
      const size = parseInt(process.argv[i + 1], 10);
      if (!isNaN(size) && size > 0) args.batchSize = size;
      i++;
    } else if ((process.argv[i] === '--limit' || process.argv[i] === '-l') && i + 1 < process.argv.length) {
      const limit = parseInt(process.argv[i + 1], 10);
      if (!isNaN(limit) && limit > 0) args.limit = limit;
      i++;
    }
  }
  return args;
}

// ==================== Helpers ====================

function readUrlsFromCsv(filePath) {
  const text = fs.readFileSync(filePath, 'utf-8');
  const records = parse(text, {
    columns: true,
    skip_empty_lines: true,
    trim: true
  });
  const urlColumn = Object.keys(records[0] || {}).find(k => k.toLowerCase() === 'url');
  if (!urlColumn) {
    console.warn('CSV does not contain a "url" column; using first column values.');
    const rows = parse(text, { skip_empty_lines: true, trim: true });
    return rows.map(row => row[0]).filter(Boolean);
  }
  return records.map(r => r[urlColumn]).filter(Boolean);
}

async function getStartUrls(args) {
  let urls = [];
  
  // If --file flag provided, use that CSV
  if (args.csvFile) {
    const absolutePath = path.isAbsolute(args.csvFile) ? args.csvFile : path.join(process.cwd(), args.csvFile);
    if (fs.existsSync(absolutePath)) {
      urls = readUrlsFromCsv(absolutePath);
    } else {
      console.error(`CSV file not found: ${absolutePath}`);
      return [];
    }
  } else {
    // Legacy: check for direct CSV filename or URL in argv[2]
    const arg = process.argv[2];
    if (!arg && fs.existsSync(DEFAULT_URL_CSV)) {
      urls = readUrlsFromCsv(DEFAULT_URL_CSV);
    } else if (arg && arg.toLowerCase().endsWith('.csv')) {
      const csvPath = path.isAbsolute(arg) ? arg : path.join(process.cwd(), arg);
      if (fs.existsSync(csvPath)) urls = readUrlsFromCsv(csvPath);
    } else if (arg && !arg.startsWith('-')) {
      urls = [arg];
    } else if (fs.existsSync(DEFAULT_URL_CSV)) {
      urls = readUrlsFromCsv(DEFAULT_URL_CSV);
    } else {
      urls = ['https://www.nationwide.com/'];
    }
  }

  // Apply limit if specified
  if (args.limit && urls.length > args.limit) {
    urls = urls.slice(0, args.limit);
  }

  return urls;
}

// Normalize text: lowercase and remove diacritics for keyword matching
function normalizeText(text) {
  return text.normalize('NFKD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
}

// Check for hreflang links with 'es' anywhere (including es-ES, es-MX, etc.)
async function hasHreflangSpanish(page) {
  const links = await page.locator('link[rel="alternate"][hreflang]').all();
  for (const link of links) {
    const hreflang = await link.getAttribute('hreflang').catch(() => null);
    if (hreflang && hreflang.toLowerCase().startsWith('es')) return true;
  }
  return false;
}

// Check for og:locale meta tag with 'es' anywhere
async function hasOgLocaleSpanish(page) {
  const meta = await page.locator('meta[property="og:locale"]').first();
  if (!meta) return false;
  const content = await meta.getAttribute('content').catch(() => null);
  return content && content.toLowerCase().startsWith('es');
}

// Enhanced Spanish detection with logging of the winning signal
async function isAlreadySpanishPage(page, url) {
  const signals = { url: false, htmlLang: false, hreflang: false, ogLocale: false, franc: false, keyword: false };
  let detectionMethod = 'none';

  // 1. Direct URL
  if (isDirectSpanishUrl(url)) {
    signals.url = true;
    detectionMethod = 'url';
    return { isSpanish: true, method: detectionMethod };
  }

  // 2. HTML lang attribute
  try {
    const htmlLang = await page.locator('html').getAttribute('lang').catch(() => '');
    if (typeof htmlLang === 'string' && htmlLang.toLowerCase().startsWith('es')) {
      signals.htmlLang = true;
      detectionMethod = 'html-lang';
      return { isSpanish: true, method: detectionMethod };
    }
  } catch (_) {}

  // 3. hreflang links
  if (await hasHreflangSpanish(page)) {
    signals.hreflang = true;
    detectionMethod = 'hreflang';
    return { isSpanish: true, method: detectionMethod };
  }

  // 4. og:locale meta
  if (await hasOgLocaleSpanish(page)) {
    signals.ogLocale = true;
    detectionMethod = 'og-locale';
    return { isSpanish: true, method: detectionMethod };
  }

  // 5. Body analysis (franc + keyword heuristics)
  let bodyText = '';
  try {
    // Try to extract main content only (if <main> exists, else fallback to body)
    const main = await page.locator('main, article').first();
    if (await main.count() > 0) {
      bodyText = await main.innerText().catch(() => '');
    } else {
      bodyText = await page.locator('body').innerText().catch(() => '');
    }
  } catch (_) { bodyText = ''; }

  // Strip excess whitespace and get sample
  const cleanText = bodyText.replace(/\s+/g, ' ').trim();
  const sample = cleanText.slice(0, 2000); // larger sample for franc

  // Check franc if we have enough text (at least 100 chars)
  let francDetected = false;
  if (sample.length > 100) {
    try {
      const detected = franc(sample, { minLength: 100 });
      if (detected === 'spa') {
        francDetected = true;
        signals.franc = true;
        detectionMethod = 'franc';
        return { isSpanish: true, method: detectionMethod };
      }
    } catch (_) {}
  }

  // Keyword heuristics: only use if franc didn't fire, and require at least 2 different keywords or a minimum count
  // We'll also require that the total occurrences of Spanish keywords is >= 3 to avoid single false positive.
  const normalizedText = normalizeText(cleanText);
  const spanishSignals = ['seguro', 'servicios', 'contacto', 'privacidad', 'términos', 'cobertura', 'hogar', 'auto', 'vida', 'español'];
  // Also include common accented forms: español, servicios, etc. Already normalized.
  let keywordCount = 0;
  const foundKeywords = new Set();
  for (const word of spanishSignals) {
    // Count occurrences (case-insensitive, already normalized)
    const matches = (normalizedText.match(new RegExp(word, 'g')) || []).length;
    if (matches > 0) {
      foundKeywords.add(word);
      keywordCount += matches;
    }
  }

  // Only count if we have at least 2 different keywords AND total occurrences >= 3
  if (foundKeywords.size >= 2 && keywordCount >= 3) {
    signals.keyword = true;
    detectionMethod = 'keyword';
    return { isSpanish: true, method: detectionMethod };
  }

  // If we got here, not Spanish
  return { isSpanish: false, method: 'none' };
}

// Old function replaced; we keep the old name for compatibility but now returns object
// We'll adapt usage to unpack.

function isDirectSpanishUrl(url) {
  try {
    const u = new URL(url);
    const host = u.hostname.toLowerCase();
    const pathname = u.pathname.toLowerCase();
    const search = u.search.toLowerCase();
    if (host.includes('espanol')) return true;
    if (pathname.includes('/es/') || pathname.includes('/espanol/') || /(^|[?&])lang=es($|[=&?])/.test(search)) return true;
  } catch (_) { /* ignore */ }
  return false;
}

async function isApplicationUnavailablePage(page, url) {
  let statusCode = 0;
  try {
    const response = await page.waitForResponse(
      resp => resp.url() === url && resp.status() >= 200,
      { timeout: 5000 }
    );
    statusCode = response.status();
  } catch (_) { /* no response captured */ }
  if (statusCode >= 400) return true;
  try {
    const title = await page.title().catch(() => '');
    const bodyText = await page.locator('body').innerText().catch(() => '');
    const combined = `${title}\n${bodyText}`.toLowerCase();
    return combined.includes('application unavailable') ||
           combined.includes('temporarily unavailable') ||
           combined.includes('this page is currently unavailable');
  } catch (_) { return false; }
}

async function findEnEspanolLinkWithRetry(page, maxRetries = 3) {
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      const linkInfo = await findEnEspanolLink(page);
      if (linkInfo.exists && linkInfo.visible && linkInfo.enabled && linkInfo.href) {
        return linkInfo;
      }
      if (linkInfo.exists) {
        await page.waitForTimeout(500);
        continue;
      }
      await page.waitForTimeout(300);
    } catch (_) {
      await page.waitForTimeout(500);
    }
  }
  return await findEnEspanolLink(page);
}

// ==================== Core Validation ====================

async function validateEnEspanolForUrl(page, context, url) {
  const result = {
    url,
    enEspanolExists: false,
    enEspanolLink: null,
    enEspanolVisible: false,
    enEspanolEnabled: false,
    spanishUrl: null,
    spanishTranslate: 'No',
    detectedLanguage: 'unknown',
    status: 'SKIPPED',
    error: null,
    pageError: null,  // capture errors during page execution when link exists
    evidence: [],
    screenshotPath: null,
    detectionMethod: null  // store which signal triggered Spanish detection
  };

  try {
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: config.navigationTimeout });

    if (await isApplicationUnavailablePage(page, url)) {
      result.status = 'SKIPPED';
      result.error = 'Application Unavailable page; validation skipped because the page is not available for testing.';
      result.screenshotPath = await takeScreenshot(page, 'skip', config);
      return result;
    }

    await dismissCookieBanner(page, config);

    const linkInfo = await findEnEspanolLinkWithRetry(page, 3);
    result.enEspanolExists = linkInfo.exists;
    result.enEspanolVisible = linkInfo.visible;
    result.enEspanolEnabled = linkInfo.enabled;
    result.enEspanolLink = linkInfo.href;

    // Use enhanced Spanish detection
    const spanishCheck = await isAlreadySpanishPage(page, url);
    const alreadySpanish = spanishCheck.isSpanish;
    result.detectionMethod = spanishCheck.method;

    if (!linkInfo.exists || !linkInfo.visible || !linkInfo.enabled || !linkInfo.href) {
      if (alreadySpanish) {
        result.spanishUrl = page.url();
        const translation = await validateSpanishTranslation(page, Object.assign({}, config, { acceptLanguage: 'es' }));
        result.spanishTranslate = translation.spanishTranslate;
        result.detectedLanguage = translation.detectedLanguage;
        result.evidence = translation.evidence;
        result.status = translation.spanishTranslate === 'Yes' ? 'PASS' : 'FAIL';
        result.error = translation.message;
        if (result.status === 'FAIL') result.screenshotPath = await takeScreenshot(page, 'fail', config);
        return result;
      }
      result.status = 'SKIPPED';
      result.error = 'En Español link missing, hidden, disabled, or has no href';
      result.screenshotPath = await takeScreenshot(page, 'skip', config);
      return result;
    }

    // --- Click the link ---
    let newPage = null;
    const waitForPage = context.waitForEvent('page', { timeout: 10000 })
      .then(p => { newPage = p; })
      .catch(() => {});

    let clickError = null;
    try {
      await linkInfo.locator.click({ timeout: 10000 });
    } catch (err) {
      clickError = err;
      result.pageError = `Click failed: ${err.message}`;
    }

    if (clickError) {
      result.status = 'FAIL';
      result.error = `Click failed: ${clickError.message}`;
      result.screenshotPath = await takeScreenshot(page, 'fail', config);
      return result;
    }

    await waitForPage;

    let spanishPage;
    try {
      if (newPage) {
        spanishPage = newPage;
        await spanishPage.waitForLoadState('domcontentloaded', { timeout: config.navigationTimeout });
      } else {
        await page.waitForLoadState('domcontentloaded', { timeout: config.navigationTimeout });
        spanishPage = page;
      }
    } catch (pageLoadErr) {
      result.pageError = `Page load error: ${pageLoadErr.message}`;
      result.status = 'FAIL';
      result.error = result.pageError;
      result.screenshotPath = await takeScreenshot(page, 'fail', config);
      return result;
    }

    result.spanishUrl = await spanishPage.url();

    // Validate Spanish translation on the resulting page – THIS IS THE ONLY DECISIVE CHECK
    const translation = await validateSpanishTranslation(spanishPage, Object.assign({}, config, { acceptLanguage: 'es' }));
    result.spanishTranslate = translation.spanishTranslate;
    result.detectedLanguage = translation.detectedLanguage;
    result.evidence = translation.evidence;
    result.status = translation.spanishTranslate === 'Yes' ? 'PASS' : 'FAIL';
    result.error = translation.message;
    if (result.status === 'FAIL') {
      result.screenshotPath = await takeScreenshot(spanishPage, 'fail', config);
    }

    // --- Clean up ---
    if (spanishPage !== page) {
      await spanishPage.close().catch(() => {});
    } else {
      await page.goBack({ timeout: 10000 }).catch(() => {});
    }

  } catch (err) {
    result.status = 'FAIL';
    result.error = `Unexpected error: ${err.message}`;
    result.screenshotPath = await takeScreenshot(page, 'fail', config);
  }
  return result;
}

// ==================== Main Execution ====================

(async () => {
  const args = parseCommandLineArgs();
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    ignoreHTTPSErrors: true,
    viewport: { width: 1280, height: 720 }
  });
  const results = [];
  const startTime = Date.now();

  try {
    const startUrls = await getStartUrls(args);
    if (!startUrls.length) {
      console.error('No URLs found in url.csv and no URL argument provided.');
      return;
    }

    console.log(`\n=== En Español Validator ===`);
    console.log(`Total URLs to process: ${startUrls.length}`);
    if (args.limit) console.log(`Limit: ${args.limit}`);
    console.log(`Batch size: ${args.batchSize}`);
    console.log(`Starting validation...\n`);

    // Process URLs in batches
    for (let batchStart = 0; batchStart < startUrls.length; batchStart += args.batchSize) {
      const batchEnd = Math.min(batchStart + args.batchSize, startUrls.length);
      const batchUrls = startUrls.slice(batchStart, batchEnd);
      const batchNum = Math.floor(batchStart / args.batchSize) + 1;
      const totalBatches = Math.ceil(startUrls.length / args.batchSize);

      if (args.verbose) {
        console.log(`\n[Batch ${batchNum}/${totalBatches}] Processing ${batchUrls.length} URLs...`);
      }

      // Process URLs in parallel within batch
      const batchPromises = batchUrls.map(async (url) => {
        let page = null;
        try {
          page = await context.newPage().catch(err => {
            console.error(`Failed to create page for ${url}: ${err.message}`);
            return null;
          });

          if (!page) {
            const result = {
              url,
              enEspanolExists: false,
              enEspanolLink: null,
              enEspanolVisible: false,
              enEspanolEnabled: false,
              spanishUrl: null,
              spanishTranslate: 'No',
              detectedLanguage: 'unknown',
              status: 'FAILED',
              error: 'Failed to create browser page',
              pageError: null,
              evidence: [],
              screenshotPath: null,
              detectionMethod: 'none'
            };
            return result;
          }

          const res = await validateEnEspanolForUrl(page, context, url);
          const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
          const processed = results.length + 1;
          const remaining = startUrls.length - processed;
          const rate = processed / (elapsed / 60); // URLs per minute
          const etaMinutes = remaining > 0 ? (remaining / rate).toFixed(1) : '0';

          console.log(`[${processed}/${startUrls.length}] ${url} → ${res.status} (${res.spanishTranslate}) [ETA: ${etaMinutes}m]`);
          return res;
        } catch (err) {
          console.error(`Error processing ${url}: ${err.message}`);
          const result = {
            url,
            enEspanolExists: false,
            enEspanolLink: null,
            enEspanolVisible: false,
            enEspanolEnabled: false,
            spanishUrl: null,
            spanishTranslate: 'No',
            detectedLanguage: 'unknown',
            status: 'FAILED',
            error: `Unexpected error: ${err.message}`,
            pageError: null,
            evidence: [],
            screenshotPath: null,
            detectionMethod: 'none'
          };
          return result;
        } finally {
          if (page) {
            await page.close().catch(() => {});
          }
        }
      });

      const batchResults = await Promise.all(batchPromises);
      results.push(...batchResults);

      // Add small delay between batches to avoid overwhelming the system
      if (batchEnd < startUrls.length) {
        await new Promise(resolve => setTimeout(resolve, 500));
      }
    }
  } catch (e) {
    console.error('Critical error:', e.message);
  } finally {
    await browser.close();
  }

  // ==================== Reports ====================

  const summary = {
    total: results.length,
    pass: results.filter(r => r.status === 'PASS').length,
    fail: results.filter(r => r.status === 'FAIL').length,
    skip: results.filter(r => r.status === 'SKIPPED').length,
    unavailable: results.filter(r => r.error && /application unavailable|temporarily unavailable|currently unavailable/i.test(r.error)).length
  };

  const reportDir = config.reportDir;
  if (!fs.existsSync(reportDir)) fs.mkdirSync(reportDir, { recursive: true });

  // JSON report
  const jsonPath = path.join(reportDir, `espanol_validation_${Date.now()}.json`);
  fs.writeFileSync(jsonPath, JSON.stringify({ summary, results }, null, 2), 'utf-8');
  console.log('JSON report saved to', jsonPath);

  // HTML report
  try {
    const htmlPath = path.join(reportDir, `espanol_validation_${Date.now()}.html`);
    const rows = results.map(r => ({
      url: r.url,
      enElement: r.enEspanolExists ? 'Yes' : 'No',
      enLink: r.spanishUrl || r.enEspanolLink || '',
      spanishTranslate: r.spanishTranslate === 'Yes' ? 'Yes' : 'No',
      status: r.status || '',
      detectionMethod: r.detectionMethod || 'none',
      pageError: r.pageError || '',
      details: r.error && r.error.length > 0 ? r.error : (r.evidence ? r.evidence.join('; ') : '')
    }));

    let html = `<!doctype html><html><head><meta charset="utf-8"><title>En Español Validation (Strict)</title>
      <style>table{border-collapse:collapse;width:100%}th,td{border:1px solid #ccc;padding:6px;text-align:left}th{background:#f0f0f0}</style>
      </head><body><h2>En Español Validation – Strict Mode</h2>
      <p><strong>Summary:</strong> ${JSON.stringify(summary)}</p>
      <p><em>Detection Method shows which signal (url, html-lang, hreflang, og-locale, franc, keyword) triggered Spanish detection.</em></p>
      <table><thead><tr><th>URL</th><th>En Español Element</th><th>En Español Link</th><th>Spanish Translate</th><th>Status</th><th>Detection Method</th><th>Page Error</th><th>Validation Details</th></tr></thead><tbody>`;
    for (const r of rows) {
      html += `<tr><td>${escapeHtml(r.url)}</td><td>${escapeHtml(r.enElement)}</td><td>${escapeHtml(r.enLink)}</td><td>${escapeHtml(r.spanishTranslate)}</td><td>${escapeHtml(r.status)}</td><td>${escapeHtml(r.detectionMethod)}</td><td>${escapeHtml(r.pageError)}</td><td>${escapeHtml(r.details)}</td></tr>`;
    }
    html += `</tbody></table></body></html>`;
    fs.writeFileSync(htmlPath, html, 'utf-8');
    console.log('HTML report saved to', htmlPath);
  } catch (e) {
    console.error('Failed to write HTML report', e.message);
  }

  // Excel report
  try {
    const Excel = require('exceljs');
    const workbook = new Excel.Workbook();
    const sheet = workbook.addWorksheet('EnEspanol');
    sheet.columns = [
      { header: 'URL', key: 'url', width: 60 },
      { header: 'En Español Element', key: 'enElement', width: 15 },
      { header: 'En Español Link', key: 'enLink', width: 60 },
      { header: 'Spanish Translate', key: 'spanishTranslate', width: 15 },
      { header: 'Status', key: 'status', width: 10 },
      { header: 'Detection Method', key: 'detectionMethod', width: 20 },
      { header: 'Page Error', key: 'pageError', width: 50 },
      { header: 'Validation Details', key: 'details', width: 80 }
    ];
    for (const r of results) {
      sheet.addRow({
        url: r.url,
        enElement: r.enEspanolExists ? 'Yes' : 'No',
        enLink: r.spanishUrl || r.enEspanolLink || '',
        spanishTranslate: r.spanishTranslate === 'Yes' ? 'Yes' : 'No',
        status: r.status || '',
        detectionMethod: r.detectionMethod || 'none',
        pageError: r.pageError || '',
        details: r.error && r.error.length > 0 ? r.error : (r.evidence ? r.evidence.join('; ') : '')
      });
    }
    const excelPath = path.join(reportDir, `espanol_validation_${Date.now()}.xlsx`);
    await workbook.xlsx.writeFile(excelPath);
    console.log('Excel report saved to', excelPath);
  } catch (e) {
    console.error('Failed to write Excel report', e.message);
  }

  // Print final summary with timing
  const totalSeconds = (Date.now() - startTime) / 1000;
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = Math.floor(totalSeconds % 60);
  const urlsPerMinute = (results.length / (totalSeconds / 60)).toFixed(1);

  console.log('\n=== Validation Complete ===');
  console.log(`Summary: ${JSON.stringify(summary)}`);
  console.log(`Time elapsed: ${minutes}m ${seconds}s`);
  console.log(`Average rate: ${urlsPerMinute} URLs/minute`);
  console.log(`Total URLs processed: ${results.length}`);
})();

function escapeHtml(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}