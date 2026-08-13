const { chromium } = require('playwright');
const { mergeConfig, dismissCookieBanner, findEnEspanolLink, validateSpanishTranslation, takeScreenshot } = require('./lib/espanol-validator-core');
const path = require('path');
const fs = require('fs');
const { parse } = require('csv-parse/sync');

// franc v6+ is ESM-only – lazy load
let _francPromise = null;
async function getFranc() {
  if (!_francPromise) {
    _francPromise = import('franc').then(mod => mod.franc || mod.default);
  }
  return _francPromise;
}

const DEFAULT_URL_CSV = path.join(__dirname, 'url.csv');
const config = mergeConfig({});

// ==================== Helpers ====================

function readUrlsFromCsv(filePath) {
  const text = fs.readFileSync(filePath, 'utf-8');
  const records = parse(text, { columns: true, skip_empty_lines: true, trim: true });
  const urlColumn = Object.keys(records[0] || {}).find(k => k.toLowerCase() === 'url');
  if (!urlColumn) {
    const rows = parse(text, { skip_empty_lines: true, trim: true });
    return rows.map(row => row[0]).filter(Boolean);
  }
  return records.map(r => r[urlColumn]).filter(Boolean);
}

async function getStartUrls() {
  const arg = process.argv[2];
  if (!arg && fs.existsSync(DEFAULT_URL_CSV)) return readUrlsFromCsv(DEFAULT_URL_CSV);
  if (arg && arg.toLowerCase().endsWith('.csv')) {
    const csvPath = path.isAbsolute(arg) ? arg : path.join(process.cwd(), arg);
    if (fs.existsSync(csvPath)) return readUrlsFromCsv(csvPath);
  }
  if (arg) return [arg];
  if (fs.existsSync(DEFAULT_URL_CSV)) return readUrlsFromCsv(DEFAULT_URL_CSV);
  return ['https://www.nationwide.com/'];
}

function normalizeText(text) {
  return text.normalize('NFKD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
}

async function hasHreflangSpanish(page) {
  const links = await page.locator('link[rel="alternate"][hreflang]').all();
  for (const link of links) {
    const hreflang = await link.getAttribute('hreflang').catch(() => null);
    if (hreflang && hreflang.toLowerCase().startsWith('es')) return true;
  }
  return false;
}

async function hasOgLocaleSpanish(page) {
  const meta = page.locator('meta[property="og:locale"]').first();
  const count = await meta.count().catch(() => 0);
  if (count === 0) return false;
  const content = await meta.getAttribute('content', { timeout: 1000 }).catch(() => null);
  return Boolean(content) && content.toLowerCase().startsWith('es');
}

async function isAlreadySpanishPage(page, url) {
  if (isDirectSpanishUrl(url)) return { isSpanish: true, method: 'url' };
  try {
    const htmlLang = await page.locator('html').getAttribute('lang').catch(() => '');
    if (htmlLang && htmlLang.toLowerCase().startsWith('es')) return { isSpanish: true, method: 'html-lang' };
  } catch (_) {}
  if (await hasHreflangSpanish(page)) return { isSpanish: true, method: 'hreflang' };
  if (await hasOgLocaleSpanish(page)) return { isSpanish: true, method: 'og-locale' };

  let bodyText = '';
  try {
    const main = await page.locator('main, article').first();
    if (await main.count() > 0) bodyText = await main.innerText().catch(() => '');
    else bodyText = await page.locator('body').innerText().catch(() => '');
  } catch (_) { bodyText = ''; }
  const cleanText = bodyText.replace(/\s+/g, ' ').trim();
  const sample = cleanText.slice(0, 2000);

  if (sample.length > 100) {
    try {
      const francFn = await getFranc();
      const detected = francFn(sample, { minLength: 100 });
      if (detected === 'spa') return { isSpanish: true, method: 'franc' };
    } catch (_) {}
  }

  const normalizedText = normalizeText(cleanText);
  const spanishSignals = ['seguro', 'servicios', 'contacto', 'privacidad', 'terminos', 'cobertura', 'vida', 'espanol'].map(normalizeText);
  let keywordCount = 0;
  const foundKeywords = new Set();
  for (const word of spanishSignals) {
    const matches = (normalizedText.match(new RegExp(`\\b${word}\\b`, 'g')) || []).length;
    if (matches > 0) { foundKeywords.add(word); keywordCount += matches; }
  }
  if (foundKeywords.size >= 2 && keywordCount >= 3) return { isSpanish: true, method: 'keyword' };

  return { isSpanish: false, method: 'none' };
}

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
    const response = await page.waitForResponse(resp => resp.url() === url && resp.status() >= 200, { timeout: 5000 });
    statusCode = response.status();
  } catch (_) {}
  if (statusCode >= 400) return true;
  try {
    const title = await page.title().catch(() => '');
    const bodyText = await page.locator('body').innerText().catch(() => '');
    const combined = `${title}\n${bodyText}`.toLowerCase();
    return combined.includes('application unavailable') || combined.includes('temporarily unavailable') || combined.includes('this page is currently unavailable') || combined.includes('# application unavailable');
  } catch (_) { return false; }
}

async function findEnEspanolLinkWithRetry(page, maxRetries = 3) {
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      const linkInfo = await findEnEspanolLink(page);
      if (linkInfo.exists && linkInfo.visible && linkInfo.enabled && linkInfo.href) return linkInfo;
      if (linkInfo.exists) await page.waitForTimeout(500);
      else await page.waitForTimeout(300);
    } catch (_) { await page.waitForTimeout(500); }
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
    errors: [], // Collect all errors encountered
    pageError: null, // Capture page/console errors
    evidence: [],
    screenshotPath: null,
    detectionMethod: null
  };

  try {
    // Use domcontentloaded – networkidle times out
    try {
      // Set up error listeners to capture page errors
      page.on('console', msg => {
        if (msg.type() === 'error') {
          result.pageError = msg.text();
        }
      });
      page.on('error', err => {
        if (!result.pageError) result.pageError = err.message;
      });

      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: config.navigationTimeout });
    } catch (err) {
      result.status = 'FAIL';
      result.error = `Navigation error: ${err.message}`;
      result.errors.push(result.error);
      result.screenshotPath = await takeScreenshot(page, 'fail', config);
      return result;
    }

    if (await isApplicationUnavailablePage(page, url)) {
      result.status = 'SKIPPED';
      result.error = 'Application Unavailable page; validation skipped.';
      result.errors.push(result.error);
      result.screenshotPath = await takeScreenshot(page, 'skip', config);
      return result;
    }

    // Check if page error occurred - skip validation if error exists
    if (result.pageError) {
      result.status = 'FAIL';
      result.error = `Page error detected: ${result.pageError}`;
      result.errors.push(result.error);
      return result;
    }

    try {
      await dismissCookieBanner(page, config);
    } catch (err) {
      result.errors.push(`Cookie banner error: ${err.message}`);
    }

    let linkInfo;
    try {
      linkInfo = await findEnEspanolLinkWithRetry(page, 3);
    } catch (err) {
      result.status = 'FAIL';
      result.error = `Error finding En Español link: ${err.message}`;
      result.errors.push(result.error);
      result.screenshotPath = await takeScreenshot(page, 'fail', config);
      return result;
    }
    
    result.enEspanolExists = linkInfo.exists;
    result.enEspanolVisible = linkInfo.visible;
    result.enEspanolEnabled = linkInfo.enabled;
    result.enEspanolLink = linkInfo.href;

    let spanishCheck;
    try {
      spanishCheck = await isAlreadySpanishPage(page, url);
    } catch (err) {
      result.errors.push(`Error detecting Spanish page: ${err.message}`);
      spanishCheck = { isSpanish: false, method: 'error' };
    }
    const alreadySpanish = spanishCheck.isSpanish;
    result.detectionMethod = spanishCheck.method;

    if (!linkInfo.exists || !linkInfo.visible || !linkInfo.enabled || !linkInfo.href) {
      if (alreadySpanish) {
        try {
          result.spanishUrl = page.url();
          const translation = await validateSpanishTranslation(page, Object.assign({}, config, { acceptLanguage: 'es' }));
          result.spanishTranslate = translation.spanishTranslate;
          result.detectedLanguage = translation.detectedLanguage;
          result.evidence = translation.evidence;
          result.status = translation.spanishTranslate === 'Yes' ? 'PASS' : 'FAIL';
          result.error = translation.message;
          if (result.status === 'FAIL') result.screenshotPath = await takeScreenshot(page, 'fail', config);
          return result;
        } catch (err) {
          result.status = 'FAIL';
          result.error = `Spanish translation validation error: ${err.message}`;
          result.errors.push(result.error);
          result.screenshotPath = await takeScreenshot(page, 'fail', config);
          return result;
        }
      }
      result.status = 'SKIPPED';
      result.error = 'En Español link missing, hidden, disabled, or has no href';
      result.errors.push(result.error);
      result.screenshotPath = await takeScreenshot(page, 'skip', config);
      return result;
    }

    // --- Click the link ---
    let newPage = null;
    let pageOpened = false;
    
    const waitForNewPage = context.waitForEvent('page', { timeout: 3000 })
      .then(p => { 
        newPage = p;
        pageOpened = true;
      })
      .catch(() => { 
        // No new page opened - likely same-page navigation
        pageOpened = false;
      });

    let clickError = null;
    try {
      await linkInfo.locator.click({ timeout: 30000 });
    } catch (err) {
      clickError = err;
    }

    if (clickError) {
      result.status = 'FAIL';
      result.error = `Click failed: ${clickError.message}`;
      result.errors.push(result.error);
      result.screenshotPath = await takeScreenshot(page, 'fail', config);
      return result;
    }

    // Wait briefly to see if new page opens
    await waitForNewPage;
    await page.waitForTimeout(1000); // Give page time to start loading

    let spanishPage;
    try {
      if (newPage) {
        // New page opened in new tab
        spanishPage = newPage;
        // Set up error listeners for Spanish page
        spanishPage.on('console', msg => {
          if (msg.type() === 'error' && !result.pageError) {
            result.pageError = msg.text();
          }
        });
        spanishPage.on('error', err => {
          if (!result.pageError) result.pageError = err.message;
        });
        await spanishPage.waitForLoadState('domcontentloaded', { timeout: config.navigationTimeout });
      } else {
        // Same-page navigation
        await page.waitForLoadState('domcontentloaded', { timeout: config.navigationTimeout });
        spanishPage = page;
      }
    } catch (err) {
      result.status = 'FAIL';
      result.error = `Page load error: ${err.message}`;
      result.errors.push(result.error);
      result.screenshotPath = await takeScreenshot(page, 'fail', config);
      return result;
    }

    try {
      result.spanishUrl = await spanishPage.url();
    } catch (err) {
      result.errors.push(`Error getting Spanish page URL: ${err.message}`);
    }

    // Check if page error occurred on Spanish page - skip validation if error exists
    if (result.pageError) {
      result.status = 'FAIL';
      result.error = `Spanish page error detected: ${result.pageError}`;
      result.errors.push(result.error);
      if (spanishPage !== page) await spanishPage.close().catch(() => {});
      return result;
    }

    // =========================================================
    // VALIDATE LINK ACTUALLY LED TO SPANISH CONTENT
    // =========================================================
    // Check if we got redirected to login/oauth/error instead of Spanish page
    const redirectUrl = result.spanishUrl || '';
    const isUnexpectedRedirect = redirectUrl.includes('authorization.oauth') 
      || redirectUrl.includes('login') 
      || redirectUrl.includes('signin')
      || redirectUrl.includes('error')
      || redirectUrl.includes('404');
    
    if (isUnexpectedRedirect) {
      result.status = 'FAIL';
      result.error = `Clicking "En Español" redirected to ${redirectUrl.split('?')[0]} instead of Spanish content`;
      result.errors.push(result.error);
      result.screenshotPath = await takeScreenshot(spanishPage, 'fail', config);
      if (spanishPage !== page) await spanishPage.close().catch(() => {});
      return result;
    }
    // =========================================================

    // =========================================================
    // WAIT FOR TRANSLATION TO RENDER
    // =========================================================
    // Step 1: Wait for the footer indicator "Privacidad" (optional - don't fail if not found)
    let privacidadFound = false;
    try {
      await spanishPage.locator('body').getByText('Privacidad', { exact: false }).first().waitFor({ timeout: 10000 });
      privacidadFound = true;
      console.log(`✅ Footer loaded (found "Privacidad") for ${url}`);
    } catch (e) {
      // Don't fail - just log that we didn't find it
      result.errors.push(`"Privacidad" footer not found (${e.message.split('\n')[0]})`);
      console.warn(`⚠️ "Privacidad" not found for ${url}, continuing with wait`);
    }

    // Step 2: Always wait 2-3 seconds for AJAX/dynamic content to fully load
    await spanishPage.waitForTimeout(2500);
    // =========================================================

    // Validate Spanish translation – now the content should be fully translated
    let translation;
    try {
      translation = await validateSpanishTranslation(spanishPage, Object.assign({}, config, { acceptLanguage: 'es' }));
    } catch (err) {
      result.status = 'FAIL';
      result.error = `Translation validation error: ${err.message}`;
      result.errors.push(result.error);
      result.screenshotPath = await takeScreenshot(spanishPage, 'fail', config);
      if (spanishPage !== page) await spanishPage.close().catch(() => {});
      return result;
    }
    
    result.spanishTranslate = translation.spanishTranslate;
    result.detectedLanguage = translation.detectedLanguage;
    result.evidence = translation.evidence;
    result.status = translation.spanishTranslate === 'Yes' ? 'PASS' : 'FAIL';
    result.error = translation.message;
    if (result.status === 'FAIL') {
      result.screenshotPath = await takeScreenshot(spanishPage, 'fail', config);
    }

    // Clean up
    if (spanishPage !== page) {
      await spanishPage.close().catch(() => {});
    } else {
      await page.goBack({ timeout: 10000 }).catch(() => {});
    }

  } catch (err) {
    result.status = 'FAIL';
    result.error = `Unexpected error: ${err.message}`;
    result.errors.push(result.error);
    result.screenshotPath = await takeScreenshot(page, 'fail', config);
  }
  return result;
}

// ==================== Main Execution ====================

(async () => {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    ignoreHTTPSErrors: true,
    viewport: { width: 1280, height: 720 }
  });
  const results = [];

  try {
    const startUrls = await getStartUrls();
    if (!startUrls.length) {
      console.error('No URLs found.');
      return;
    }

    for (const url of startUrls) {
      const page = await context.newPage();
      try {
        const res = await validateEnEspanolForUrl(page, context, url);
        results.push(res);
        console.log(`${url} → ${res.status} (${res.spanishTranslate}) [detection: ${res.detectionMethod || 'none'}]`);
      } finally {
        await page.close().catch(() => {});
      }
    }
  } catch (e) {
    console.error(e.message);
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
      details: r.errors && r.errors.length > 0 ? r.errors.join('; ') : (r.error && r.error.length > 0 ? r.error : (r.evidence ? r.evidence.join('; ') : ''))
    }));

    let html = `<!doctype html><html><head><meta charset="utf-8"><title>En Español Validation (Strict)</title>
      <style>table{border-collapse:collapse;width:100%}th,td{border:1px solid #ccc;padding:6px;text-align:left}</style>
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
      { header: 'Page Error', key: 'pageError', width: 60 },
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
        details: r.errors && r.errors.length > 0 ? r.errors.join('; ') : (r.error && r.error.length > 0 ? r.error : (r.evidence ? r.evidence.join('; ') : ''))
      });
    }
    const excelPath = path.join(reportDir, `espanol_validation_${Date.now()}.xlsx`);
    await workbook.xlsx.writeFile(excelPath);
    console.log('Excel report saved to', excelPath);
  } catch (e) {
    console.error('Failed to write Excel report', e.message);
  }

  console.log('Summary:', summary);
})();

function escapeHtml(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}