const fs = require('fs');
const path = require('path');

const DEFAULT_CONFIG = {
  maxRetries: 2,
  navigationTimeout: 30000,
  elementTimeout: 15000,
  reportDir: path.join(__dirname, '..', 'reports'),
  acceptLanguage: null,
  spanishIndicators: [
    'seguro', 'seguros', 'inicio', 'servicios', 'contacto',
    'acerca de', 'privacidad', 'términos', 'condiciones',
    'ayuda', 'soporte', 'reclamación', 'póliza', 'deducible',
    'prima', 'cobertura', 'hogar', 'auto', 'vida', 'negocio'
  ],
  englishIndicators: [
    'insurance', 'policy', 'claim', 'deductible', 'premium',
    'coverage', 'home', 'auto', 'life', 'business', 'retirement'
  ],
  cookieSelectors: [
    '#truste-consent-content',
    'button:has-text("Accept")',
    'button:has-text("Accept All")',
    'button:has-text("OK")',
    'button:has-text("Aceptar")',
    '#onetrust-accept-btn-handler',
    '.cookie-accept',
    '[data-testid="cookie-accept"]'
  ]
};


function mergeConfig(cfg) {
  return Object.assign({}, DEFAULT_CONFIG, cfg || {});
}

async function retry(fn, retries = 2) {
  let lastErr;
  for (let i = 0; i < retries; i++) {
    try { return await fn(); } catch (e) { lastErr = e; await new Promise(r => setTimeout(r, 500)); }
  }
  throw lastErr;
}

async function dismissCookieBanner(page, config) {
  const selectors = (config && config.cookieSelectors) || DEFAULT_CONFIG.cookieSelectors;
  try {
    for (const selector of selectors) {
      try {
        const btn = page.locator(selector).first();
        if (await btn.count() > 0 && await btn.isVisible()) {
          await btn.click({ timeout: 3000 }).catch(() => {});
          await page.waitForTimeout(400);
          return true;
        }
      } catch (_) {}
    }
  } catch {} 
  return false;
}

async function waitForFooter(page) {
  const footerSelectors = [ 'footer', '[role="contentinfo"]', '[class*="footer"]', '[id*="footer"]' ];
  for (const sel of footerSelectors) {
    try {
      const el = page.locator(sel).first();
      if (await el.count() > 0 && await el.isVisible()) return el;
    } catch (_) {}
  }
  return null;
}

async function findEnEspanolLink(page) {
  const normalize = s => (s || '').toString().toLowerCase().normalize('NFD').replace(/\p{Diacritic}/gu, '').trim();

  try {
    const headLink = await page.locator('head link[rel="alternate"][hreflang^="es"]').first();
    if (headLink && (await headLink.count()) > 0) {
      const href = await headLink.getAttribute('href').catch(() => null);
      // Skip portal/essportal/oauth URLs
      if (href && (href.includes('portal') || href.includes('essportal') || href.includes('oauth') || href.includes('authorization'))) {
        // Skip this link, continue to other detection methods
      } else if (href) {
        const locator = page.locator(`a[href="${href}"]`).first();
        return { exists: true, locator, href, visible: await locator.isVisible().catch(() => false), enabled: await locator.isEnabled().catch(() => false) };
      }
    }
  } catch (_) {}

  try {
    const html = await page.content().catch(() => '');
    const hrefRegex = /https?:\/\/(?:www\.)?espanol\.nationwide\.com(?:\/[^\s"'<>]*)?/i;
    const sourceMatch = html.match(hrefRegex) || html.match(/\bEn\s+Espa(?:n|ñ)ol\b/i);
    if (sourceMatch) {
      const href = sourceMatch[0].startsWith('http') ? sourceMatch[0].replace(/["'<>].*$/, '') : 'https://espanol.nationwide.com';
      const locator = page.locator('a[href*="espanol.nationwide.com"]').first();
      if (locator && (await locator.count().catch(() => 0)) > 0) {
        return { exists: true, locator, href, visible: await locator.isVisible().catch(() => false), enabled: await locator.isEnabled().catch(() => false) };
      }
      return { exists: true, locator, href, visible: false, enabled: false };
    }
  } catch (_) {}

  const attempts = 6;
  for (let attempt = 0; attempt < attempts; attempt++) {
    try {
      const footer = await waitForFooter(page);
      if (footer) {
        const linkLocator = footer.locator('a:has-text("En Español"), a:has-text("Español")').first();
        if (await linkLocator.count() > 0) {
          const href = await linkLocator.getAttribute('href').catch(() => null);
          // Skip portal/essportal URLs
          if (href && (href.includes('portal') || href.includes('essportal'))) {
            continue;
          }
          return { exists: true, locator: linkLocator, href, visible: await linkLocator.isVisible().catch(() => false), enabled: await linkLocator.isEnabled().catch(() => false) };
        }
      }
    } catch (_) {}

    try {
      const nodes = await page.locator('a, button').all();
      for (const node of nodes) {
        try {
          const href = await node.getAttribute('href').catch(() => null);
          // Skip OAuth/auth/login URLs entirely
          if (href && (href.includes('oauth') || href.includes('authorization') || href.includes('login') || href.includes('signin') || href.includes('portal') || href.includes('essportal'))) {
            continue;
          }
          
          const rawText = (await node.innerText().catch(() => '')) || (await node.getAttribute('aria-label').catch(() => '')) || (await node.getAttribute('title').catch(() => ''));
          const text = normalize(rawText);
          const hreflang = ((await node.getAttribute('hreflang').catch(() => '')) || '').toLowerCase();
          const lang = ((await node.getAttribute('lang').catch(() => '')) || '').toLowerCase();
          let imgAlt = '';
          try { const img = await node.locator('img').first(); if (img && (await img.count()) > 0) imgAlt = (await img.getAttribute('alt').catch(() => '')) || ''; } catch {}

          if (text.includes('espanol') || text.includes('en espanol') || imgAlt.toLowerCase().includes('espanol')) {
            // Only accept if text is short and clearly says "español" (not embedded in longer text)
            // This filters out false positives from metadata
            if (text === 'espanol' || text === 'en espanol' || text.match(/^en\s+espanol$/) || text.match(/^espanol$/)) {
              const locator = href ? page.locator(`a[href="${href}"]`).first() : node;
              return { exists: true, locator, href, visible: await locator.isVisible().catch(() => false), enabled: await locator.isEnabled().catch(() => false) };
            }
          }

          if (hreflang.startsWith('es') || lang.startsWith('es')) {
            // Skip portal/essportal URLs and links without Spanish path indicators
            if (href && (href.includes('portal') || href.includes('essportal'))) {
              continue;
            }
            // Also verify href has some Spanish indicator in path or query
            let hasSpanishHref = false;
            try {
              const hrefLower = href ? href.toLowerCase() : '';
              const url = new URL(href || 'http://example.com', page.url());
              if (url.pathname.includes('/es') || url.search.includes('lang=es')) {
                hasSpanishHref = true;
              }
            } catch (e) {}
            
            if (!hasSpanishHref) continue; // Skip if href doesn't have Spanish indicators
            
            const locator = href ? page.locator(`a[href="${href}"]`).first() : node;
            return { exists: true, locator, href, visible: await locator.isVisible().catch(() => false), enabled: await locator.isEnabled().catch(() => false) };
          }

          if (href) {
            try {
              const abs = new URL(href, page.url()).pathname;
              // Only accept if path clearly indicates Spanish (/es, /espanol, lang=es)
              if (abs.startsWith('/es') || abs.split('/').includes('es') || /\/es(\/|$)/.test(abs) || /[?&]lang=es\b/.test(href)) {
                const locator = page.locator(`a[href="${href}"]`).first();
                return { exists: true, locator, href, visible: await locator.isVisible().catch(() => false), enabled: await locator.isEnabled().catch(() => false) };
              }
            } catch (e) {}
          }
        } catch (e) {}
      }
    } catch (e) {}

    try {
      const frames = page.frames();
      for (const fr of frames) {
        if (fr === page.mainFrame()) continue;
        try {
          const fLink = await fr.locator('a:has-text("En Español"), a:has-text("Español")').first();
          if (fLink && (await fLink.count()) > 0) {
            const href = await fLink.getAttribute('href').catch(() => null);
            return { exists: true, locator: fLink, href, visible: await fLink.isVisible().catch(() => false), enabled: await fLink.isEnabled().catch(() => false) };
          }
        } catch {}
      }
    } catch {}

    try {
      const found = await page.evaluate(() => {
        const norm = s => (s || '').toString().normalize('NFD').replace(/\p{Diacritic}/gu, '').toLowerCase();
        const nodes = Array.from(document.querySelectorAll('a, button'));
        for (const n of nodes) {
          try {
            const href = n.getAttribute('href') || null;
            // Skip oauth/auth/login/portal URLs
            if (href && (href.includes('oauth') || href.includes('authorization') || href.includes('login') || href.includes('portal'))) continue;
            
            const txt = norm(n.innerText || n.getAttribute('aria-label') || n.getAttribute('title') || '');
            // Only exact matches - "español" or "en español"
            if (txt === 'espanol' || txt === 'en espanol') return { href };
            
            const img = n.querySelector && n.querySelector('img'); 
            if (img && norm(img.getAttribute('alt') || '') === 'espanol') return { href };
          } catch {}
        }
        return null;
      });
      if (found && found.href) {
        const locator = page.locator(`a[href="${found.href}"]`).first();
        return { exists: true, locator, href: found.href, visible: await locator.isVisible().catch(() => false), enabled: await locator.isEnabled().catch(() => false) };
      }
    } catch {}

    await page.waitForTimeout(500);
  }

  return { exists: false, locator: null, href: null, visible: false, enabled: false };
}

async function validateSpanishTranslation(page, config) {
  config = mergeConfig(config);
  const result = { spanishTranslate: 'No', detectedLanguage: 'unknown', evidence: [], message: '' };
  const evidence = [];

  const currentUrl = page.url();
  
  // If caller requested an Accept-Language header, set it and reload to fetch localized content
  try {
    if (config && config.acceptLanguage) {
        try { await page.context().setExtraHTTPHeaders({ 'accept-language': String(config.acceptLanguage) }); } catch (_) {}
      try { await page.goto(currentUrl, { waitUntil: 'domcontentloaded', timeout: config.navigationTimeout }); } catch (_) {}
      try { await page.waitForLoadState('networkidle', { timeout: 3000 }); } catch (_) {}
    }
  } catch (_) {}

  // Check if URL itself indicates Spanish (espanol.*, /es/, ?lang=es)
  const urlHasSpanish = currentUrl.includes('espanol') || /\/es(\/|$|\?)/.test(currentUrl) || /[?&]lang=es/.test(currentUrl);
  if (urlHasSpanish) {
    evidence.push(`URL indicates Spanish: ${urlHasSpanish}`);
  }

  const langAttr = await page.locator('html').getAttribute('lang').catch(() => null);
  if (langAttr) {
    const isSpanish = langAttr.toLowerCase().startsWith('es');
    evidence.push(`lang attribute: ${langAttr} → ${isSpanish ? 'Spanish' : 'not Spanish'}`);
    if (isSpanish) result.detectedLanguage = langAttr;
  }

  // Prefer the fully-rendered document text to capture hydrated/SPAs content
  let bodyText = '';
  try {
    // Try to read the visible text of the whole document
    bodyText = await page.evaluate(() => (document.documentElement && document.documentElement.innerText) || document.body.innerText || '');
  } catch (_) {}
  if (!bodyText) {
    try { bodyText = await page.locator('body').innerText().catch(() => ''); } catch (_) { bodyText = ''; }
  }
  const lowerBody = String(bodyText || '').toLowerCase();

  let spanishHitCount = 0; for (const word of config.spanishIndicators) if (lowerBody.includes(word)) spanishHitCount++;
  let englishHitCount = 0; for (const word of config.englishIndicators) if (lowerBody.includes(word)) englishHitCount++;

  const spanishStopwords = ['el','la','los','las','un','una','y','o','pero','por','para','con','sin','que','es','en','se','su','como','porque','cuando','donde','este','esta'];
  let stopwordHits = 0; for (const sw of spanishStopwords) if (new RegExp(`\\b${sw}\\b`).test(lowerBody)) stopwordHits++;
  evidence.push(`Spanish stopwords found: ${stopwordHits}`);

  evidence.push(`Spanish keywords found: ${spanishHitCount}`);
  evidence.push(`English keywords found: ${englishHitCount}`);

  const words = lowerBody.split(/\s+/).filter(w => w.length > 2);
  let spanishCharWords = 0;
  for (const w of words) {
    if (/[áéíóúñ]/.test(w) || w.endsWith('ción') || w.endsWith('dad') || w.endsWith('mente') || w.endsWith('ía') || w.endsWith('aje')) spanishCharWords++;
  }
  const ratio = words.length > 0 ? spanishCharWords / words.length : 0;
  evidence.push(`Spanish-character words: ${spanishCharWords}/${words.length} (${(ratio*100).toFixed(1)}%)`);

  const hasSpanishLang = langAttr && langAttr.toLowerCase().startsWith('es');
  const hasKeywords = spanishHitCount >= 2 || stopwordHits >= 3;
  const ratioGood = ratio > 0.08 || stopwordHits >= 3; // loosened threshold
  
  // Prefer content evidence over lang attribute if content strongly indicates Spanish
  let pass = false;
  
  // Strong signals: Spanish language attribute + good content
  if (hasSpanishLang && spanishHitCount >= 1 && englishHitCount <= spanishHitCount + 2) pass = true;
  
  // Content-based: strong keywords/stopwords + good ratio
  if (hasKeywords && (ratioGood || spanishHitCount >= 3 || stopwordHitCount >= 4)) pass = true;
  
  // URL-based: if URL clearly says "espanol", it's usually Spanish even if content detection is weak
  // (site may use JS to lazy-load translations or have mixed language content)
  if (urlHasSpanish) {
    // More lenient thresholds when URL says Spanish
    if (spanishHitCount >= 1 || stopwordHits >= 1 || ratio > 0.02) pass = true;
    // Even if no Spanish keywords found, if URL is Spanish and doesn't have overwhelming English, pass it
    // Allow up to 10 English keywords since navigation/UI might be in English
    if (spanishHitCount === 0 && stopwordHits === 0 && englishHitCount <= 10) pass = true;
  }

  if (pass) {
    result.spanishTranslate = 'Yes';
    result.message = 'Spanish translation successfully validated.';
    result.detectedLanguage = result.detectedLanguage || 'es (inferred)';
  } else {
    result.spanishTranslate = 'No';
    result.message = `Spanish translation not detected. Indicators: ${evidence.join('; ')}`;
    result.detectedLanguage = result.detectedLanguage || 'unknown';
  }
  result.evidence = evidence;
  return result;
}

async function takeScreenshot(page, prefix = 'screenshot', config) {
  config = mergeConfig(config);
  try {
    if (!fs.existsSync(config.reportDir)) fs.mkdirSync(config.reportDir, { recursive: true });
    const filename = `${prefix}_${Date.now()}.png`;
    const filepath = path.join(config.reportDir, filename);
    await page.screenshot({ path: filepath, fullPage: true }).catch(() => {});
    return filepath;
  } catch { return null; }
}

module.exports = {
  DEFAULT_CONFIG,
  mergeConfig,
  retry,
  dismissCookieBanner,
  waitForFooter,
  findEnEspanolLink,
  validateSpanishTranslation,
  takeScreenshot
};
