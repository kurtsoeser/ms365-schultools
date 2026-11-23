/**
 * Landing-Screenshots je Persona (helles Teal) via Playwright.
 * Daten: scripts/landing-screens-data.mjs
 *
 * Voraussetzung: npm run dev
 * Usage: node scripts/capture-landing-screens.mjs
 */
import { chromium } from 'playwright';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import fs from 'node:fs';
import { allLandingShots } from './landing-screens-data.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');
const outDir = path.join(root, 'landing', 'assets', 'screens');
const base = process.env.MS365_SHOT_BASE || 'http://localhost:5173';

const VIEWPORT = { width: 1440, height: 900 };
const VIEWPORT_CLIP = { x: 0, y: 0, width: 1440, height: 900 };

const SHOT_ACCOUNT = {
  username: 'mueller@schule.demo',
  name: 'Maria Müller',
  localAccountId: 'landing-shot-demo',
  homeAccountId: 'landing-shot-demo',
  tenantId: 'demo-tenant'
};

async function clearOverlays(page) {
  await page.evaluate(() => {
    try {
      document.documentElement.removeAttribute('data-ms365-license');
      document
        .querySelectorAll(
          [
            '.ms365-license-gate',
            '.ms365-onboarding-overlay',
            '.ms365-pin-gate',
            '.ms365-auth-required',
            '.auth-required-overlay',
            '[data-ms365-login-modal]',
            '.toast',
            '.fr-toast',
            '.sa-toast',
            '.akt-toast',
            '.ms365-toast',
            '.ms365-it-first-run-chip',
            '.ms365-lehrer-portal-setup-banner',
            '#ms365LehrerPortalSetupBanner',
            '.dash-setup-banner',
            '[class*="spotlight"]',
            '[class*="product-tour"]',
            '[class*="tour-overlay"]'
          ].join(', ')
        )
        .forEach((el) => el.remove());
      document.querySelectorAll('[role="status"], .fr-alert--info, .akt-alert--info, .sa-alert--info').forEach((el) => {
        const t = String(el.textContent || '');
        if (/Demo-Ansicht|Demo SJ|Paket geladen|Konto wird gelesen|Stammdaten übernommen/i.test(t)) {
          el.remove();
        }
      });
      // Login-Modale ohne feste Klasse: Überschrift treffen
      document.querySelectorAll('h1, h2, h3, p, strong').forEach((el) => {
        const t = String(el.textContent || '');
        if (/Anmeldung erforderlich|Mit Microsoft anmelden/i.test(t)) {
          const box = el.closest('dialog, .modal, .overlay, [role="dialog"], section, div');
          if (box && box !== document.body && box !== document.documentElement) {
            if (box.querySelector('main, .it-workspace-shell, #dashAppOnlyStack, .app-tool-chrome')) return;
            const root = box.parentElement && box.parentElement.children.length < 4 ? box.parentElement : box;
            if (
              root &&
              root !== document.body &&
              root !== document.documentElement &&
              !root.querySelector('main, .it-workspace-shell, #dashAppOnlyStack')
            ) {
              root.remove();
            }
          }
        }
      });
      document.body.style.overflow = '';
    } catch (_) {}
  });
}

async function enableTheme(page) {
  await page.evaluate(() => {
    try {
      localStorage.setItem('ms365-theme-v1', 'light');
      localStorage.setItem('ms365-brand-v1', 'teal');
      localStorage.setItem('ms365-onboarding-welcome-v1', 'seen');
      localStorage.setItem('ms365-dashboard-setup-dismissed-v1', '1');
      document.documentElement.setAttribute('data-theme', 'light');
      document.documentElement.setAttribute('data-brand', 'teal');
      document.documentElement.style.colorScheme = 'light';
    } catch (_) {}
  });
  await page
    .evaluate(() => {
      const btn = document.querySelector(
        '[data-ms365-setup-dismiss], #dashSetupDismiss, .dash-setup-banner button'
      );
      if (btn) btn.click();
    })
    .catch(() => {});
}

async function fakeLoggedInTeacher(page) {
  await page.evaluate((acc) => {
    const info = {
      username: acc.username,
      name: acc.name,
      localAccountId: acc.localAccountId,
      homeAccountId: acc.homeAccountId,
      tenantId: acc.tenantId,
      idTokenClaims: {
        preferred_username: acc.username,
        name: acc.name,
        tid: acc.tenantId
      }
    };
    window.ms365AuthIsLoggedIn = () => true;
    window.ms365AuthGetAccountInfo = () => info;
    window.ms365AuthGetAccount = () => ({
      username: acc.username,
      name: acc.name,
      localAccountId: acc.localAccountId,
      homeAccountId: acc.homeAccountId,
      tenantId: acc.tenantId
    });
    window.dispatchEvent(new CustomEvent('ms365-auth-state-changed'));
    window.dispatchEvent(new CustomEvent('ms365-auth-widget-ready'));
  }, SHOT_ACCOUNT);
}

async function settle(page, ms = 700) {
  await page.waitForTimeout(ms);
  await clearOverlays(page).catch(() => {});
  const landing = await page
    .evaluate(() => new URLSearchParams(location.search || '').get('landingShot') === '1')
    .catch(() => false);
  if (!landing) await fakeLoggedInTeacher(page).catch(() => {});
  await page.waitForTimeout(200);
  await clearOverlays(page).catch(() => {});
}

async function clickFirst(page, selectors) {
  for (const sel of selectors) {
    const loc = page.locator(sel).first();
    if ((await loc.count()) === 0) continue;
    if (!(await loc.isVisible().catch(() => false))) continue;
    await loc.click({ timeout: 4000 }).catch(() => {});
    return true;
  }
  return false;
}

async function prepareBasic(page) {
  await enableTheme(page);
  await settle(page, 500);
  await page.evaluate(() => window.scrollTo(0, 0));
}

async function prepareDashLehrer(page) {
  await prepareBasic(page);
  await page
    .waitForSelector('#dashAppOnlyStack .dash-portal-row, #dashAppOnlyStack .dash-portal-card, .dash-portal-home', {
      timeout: 8000
    })
    .catch(() => {});
  await settle(page, 700);
  await page.evaluate(() => window.scrollTo(0, 0));
}

async function prepareDashIt(page) {
  await prepareBasic(page);
  await page
    .waitForFunction(
      () => {
        const text = document.body ? document.body.innerText || '' : '';
        return text.length > 400 && /Stammdaten|Mitgliedschaften|Schuljahresstart/.test(text);
      },
      { timeout: 12000 }
    )
    .catch(() => {});
  await settle(page, 500);
  await page.evaluate(() => window.scrollTo(0, 0)).catch(() => {});
}

async function prepareStammdaten(page) {
  await prepareBasic(page);
  await page.evaluate(() => {
    const set = (sel, val) => {
      const el = document.querySelector(sel);
      if (!el) return;
      el.value = val;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
    };
    set('#schoolName, [name="schoolName"]', 'MS Musterstadt');
    set('#schoolDomain, [name="schoolDomain"], #emailDomain', 'schule.demo');
    set('#schoolSkz, [name="skz"], #schulkennzahl', '123456');
  });
  await settle(page, 400);
}

async function prepareOrganigramm(page) {
  await prepareBasic(page);
  await clickFirst(page, [
    'a[href="#organigramm"]',
    'button:has-text("Organigramm")',
    '[data-tab="organigramm"]',
    '#tab-organigramm'
  ]);
  await settle(page, 800);
  await clickFirst(page, [
    'button:has-text("Aus Stammdaten")',
    'button:has-text("Neu aufbauen")',
    'button:has-text("Organigramm erzeugen")',
    'button:has-text("Demo")'
  ]);
  await settle(page, 1400);
  await clickFirst(page, ['button:has-text("Schließen")', 'button:has-text("OK")']);
  await settle(page, 500);
  await page.evaluate(() => {
    const el =
      document.querySelector('#organigramm, [data-organigramm], .org-canvas, .schul-organigramm') ||
      document.querySelector('main');
    if (el) el.scrollIntoView({ block: 'start' });
  });
  await settle(page, 400);
}

async function prepareGruppen(page) {
  await prepareBasic(page);
  await clickFirst(page, ['#ssBtnDemo', 'button:has-text("Demo")']);
  await settle(page, 800);
  await clickFirst(page, ['button:has-text("Schließen")']);
  await settle(page, 300);
}

async function prepareKursteams(page) {
  await prepareBasic(page);
  await clickFirst(page, ['button:has-text("Schließen")']);
  await settle(page, 400);
}

async function prepareLandingReady(page) {
  await prepareBasic(page);
  await page
    .waitForFunction(
      () => {
        const el = document.querySelector('#frApp, #saApp, #aktApp, #lfrApp');
        if (!el) return (document.body.innerText || '').length > 700;
        return (el.innerText || '').trim().length > 40;
      },
      { timeout: 8000 }
    )
    .catch(() => {});
  await settle(page, 300);
}

const prepares = {
  basic: prepareBasic,
  'dash-lehrer': prepareDashLehrer,
  'dash-it': prepareDashIt,
  stammdaten: prepareStammdaten,
  organigramm: prepareOrganigramm,
  gruppen: prepareGruppen,
  kursteams: prepareKursteams,
  'demo-click': async (page) => {
    await prepareBasic(page);
    await clickFirst(page, [
      'button:has-text("Demo laden")',
      'button:has-text("Demo-Daten")',
      'button:has-text("Demo")',
      '#pwBtnDemo',
      '#ebLoadDemo'
    ]);
    await settle(page, 900);
  },
  'fr-liste': prepareLandingReady,
  'fr-antrag': prepareLandingReady,
  'fr-schueler': prepareLandingReady,
  'sa-list': prepareLandingReady,
  'akt-list': prepareLandingReady,
  eb: prepareLandingReady,
  srdp: prepareLandingReady,
  krank: async (page) => {
    await prepareBasic(page);
    await clickFirst(page, ['button:has-text("Demo laden")', 'button:has-text("Demo")']);
    await settle(page, 800);
  },
  ausgaben: async (page) => {
    await prepareBasic(page);
    await clickFirst(page, ['button:has-text("Demo laden")', 'button:has-text("Demo")']);
    await settle(page, 800);
  }
};

/** Hero + Legacy-Dateien, die die Landing noch referenziert. */
const EXTRA = [
  {
    file: 'dashboard.png',
    captureUrl: '/index.html?landingShot=1&persona=lehrer',
    prepare: 'dash-lehrer'
  },
  {
    file: 'werkzeuge.png',
    captureUrl: '/index.html?landingShot=1&persona=it',
    prepare: 'dash-it'
  },
  {
    file: 'hilfe.png',
    captureUrl: '/hilfe.html',
    prepare: 'basic'
  }
];

async function main() {
  fs.mkdirSync(outDir, { recursive: true });

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    viewport: VIEWPORT,
    deviceScaleFactor: 2
  });

  await context.addInitScript(() => {
    try {
      localStorage.setItem('ms365-theme-v1', 'light');
      localStorage.setItem('ms365-brand-v1', 'teal');
      localStorage.setItem('ms365-onboarding-welcome-v1', 'seen');
      localStorage.setItem('ms365-dashboard-setup-dismissed-v1', '1');
      localStorage.setItem('ms365-lehrer-welcome-v1', 'seen');
      const tourDone = JSON.stringify({
        v: 1,
        phase: 'skipped',
        step: 0,
        updatedAt: Date.now()
      });
      localStorage.setItem('ms365-it-first-run-v1', tourDone);
      localStorage.setItem('ms365-lehrer-first-run-v1', tourDone);
      localStorage.setItem('ms365-it-first-run-chip-dismiss-v1', '1');
      localStorage.setItem('ms365-lehrer-first-run-chip-dismiss-v1', '1');
      const shotPersona = new URLSearchParams(location.search || '').get('persona');
      const shell =
        shotPersona === 'it' ? 'full' : shotPersona === 'schueler' ? 'schueler' : 'portal';
      localStorage.setItem('ms365-dashboard-it-preview-v1', shell);
      localStorage.setItem('ms365-dashboard-it-preview-v1:demo-tenant', shell);
      if (shotPersona === 'it') sessionStorage.setItem('ms365-it-workspace-v1', '1');
      else sessionStorage.removeItem('ms365-it-workspace-v1');
      sessionStorage.setItem('ms365-access-granted-v1', '1');
      localStorage.setItem(
        'ms365-license-me-v1',
        JSON.stringify({
          ok: true,
          status: 'active',
          cachedAt: Date.now(),
          tenantId: 'demo-tenant'
        })
      );
      window.confirm = () => false;
      window.alert = () => {};
    } catch (_) {}
  });

  const shots = [
    ...EXTRA,
    ...allLandingShots().map((s) => ({
      file: s.file,
      captureUrl: s.captureUrl,
      prepare: s.prepare || 'basic'
    }))
  ];

  // Dedup by file (EXTRA first wins for dashboard aliases)
  const seen = new Set();
  const unique = [];
  for (const s of shots) {
    if (seen.has(s.file)) continue;
    seen.add(s.file);
    unique.push(s);
  }

  const only = String(process.env.MS365_SHOT_ONLY || '')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);

  async function captureOne(shot) {
    const page = await context.newPage();
    page.setDefaultTimeout(12000);
    page.setDefaultNavigationTimeout(25000);
    const url = base + shot.captureUrl;
    let watchdog;
    const kill = new Promise((_, reject) => {
      watchdog = setTimeout(() => reject(new Error('shot-watchdog')), 40000);
    });
    try {
      return await Promise.race([
        kill,
        (async () => {
          console.error('  goto', shot.file);
          await page.goto(url, { waitUntil: 'commit', timeout: 20000 });
          console.error('  committed', shot.file);
          await page.waitForTimeout(1000);
          await clearOverlays(page).catch(() => {});
          const prep = prepares[shot.prepare] || prepareBasic;
          await Promise.race([
            prep(page),
            new Promise((_, reject) => setTimeout(() => reject(new Error('prep-timeout')), 20000))
          ]);
          console.error('  prepped', shot.file);
          await clearOverlays(page).catch(() => {});
          await page.waitForTimeout(250);
          const outPath = path.join(outDir, shot.file);
          await page.screenshot({
            path: outPath,
            type: 'png',
            clip: VIEWPORT_CLIP
          });
          const st = fs.statSync(outPath);
          console.log('  saved', shot.file, Math.round(st.size / 1024) + 'KB');
          return st.size;
        })()
      ]);
    } finally {
      clearTimeout(watchdog);
      await Promise.race([
        page.close().catch(() => {}),
        new Promise((resolve) => setTimeout(resolve, 4000))
      ]);
    }
  }

  let savedCount = 0;
  for (const shot of unique) {
    if (only.length && !only.includes(shot.file)) continue;
    console.log('→', shot.file, base + shot.captureUrl);
    try {
      let size = await captureOne(shot);
      if (size < 80 * 1024) {
        console.log('  retry (zu leer)', shot.file);
        size = await captureOne(shot);
      }
      savedCount += 1;
    } catch (err) {
      console.error('  fail', shot.file, err && err.message ? err.message : err);
      try {
        await captureOne(shot);
        savedCount += 1;
      } catch (err2) {
        console.error('  fail2', shot.file, err2 && err2.message ? err2.message : err2);
      }
    }
  }

  await browser.close();
  console.log('done ·', savedCount, 'files');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
