/**
 * Capture light-mode landing screenshots via Playwright.
 * Usage: node scripts/capture-landing-screens.mjs
 */
import { chromium } from 'playwright';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import fs from 'node:fs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');
const outDir = path.join(root, 'landing', 'assets', 'screens');
const base = process.env.MS365_SHOT_BASE || 'http://127.0.0.1:5173';

const VIEWPORT = { width: 1440, height: 900 };

const shots = [
  {
    file: 'dashboard.png',
    url: '/index.html',
    prepare: async (page) => {
      await enableDemo(page);
      await page.evaluate(() => window.scrollTo(0, 0));
      await page.waitForTimeout(350);
    }
  },
  {
    file: 'werkzeuge.png',
    url: '/index.html',
    prepare: async (page) => {
      await enableDemo(page);
      await page.evaluate(() => {
        const el = document.getElementById('dashboard-tools');
        if (el) {
          const top = el.getBoundingClientRect().top + window.scrollY - 12;
          window.scrollTo(0, Math.max(0, top));
        }
      });
      await page.waitForTimeout(350);
    }
  },
  {
    file: 'einrichtung.png',
    url: '/einrichtung.html',
    prepare: async (page) => {
      await enableDemo(page);
      await page.evaluate(() => window.scrollTo(0, 0));
      await page.waitForTimeout(300);
    }
  },
  {
    file: 'kursteams.png',
    url: '/tools/kursteams.html',
    prepare: async (page) => {
      await enableDemo(page);
      await page.evaluate(() => window.scrollTo(0, 0));
      await page.waitForTimeout(300);
    }
  },
  {
    file: 'gruppen.png',
    url: '/tools/schulstruktur-sync.html',
    prepare: async (page) => {
      await enableDemo(page);
      await page.evaluate(() => window.scrollTo(0, 0));
      await page.waitForTimeout(300);
    }
  },
  {
    file: 'schularbeiten.png',
    url: '/tools/schularbeiten-planer.html',
    prepare: async (page) => {
      await enableDemo(page);
      await page.evaluate(() => window.scrollTo(0, 0));
      await page.waitForTimeout(400);
    }
  },
  {
    file: 'projektwochen.png',
    url: '/tools/projektwochen.html',
    prepare: async (page) => {
      await enableDemo(page);
      await page.evaluate(() => window.scrollTo(0, 0));
      await page.waitForTimeout(400);
    }
  },
  {
    file: 'aufraeumen.png',
    url: '/tools/cleanup-playbook.html',
    prepare: async (page) => {
      await enableDemo(page);
      await page.evaluate(() => window.scrollTo(0, 0));
      await page.waitForTimeout(300);
    }
  },
  {
    file: 'hilfe.png',
    url: '/hilfe.html',
    prepare: async (page) => {
      await page.evaluate(() => window.scrollTo(0, 0));
      await page.waitForTimeout(300);
    }
  }
];

/** Alle Screens: Viewport-Ausschnitt (kein Full-Page), helles Teal-Branding. */
const VIEWPORT_CLIP = { x: 0, y: 0, width: 1440, height: 900 };

async function clearOverlays(page) {
  await page.evaluate(() => {
    try {
      document.documentElement.removeAttribute('data-ms365-license');
      document.querySelectorAll('.ms365-license-gate, .ms365-onboarding-overlay').forEach((el) => el.remove());
      document.querySelectorAll('body > *').forEach((el) => {
        el.style.pointerEvents = '';
        el.style.userSelect = '';
      });
    } catch (_) {}
  });
}

async function enableDemo(page) {
  await clearOverlays(page);
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

    const demo = window.ms365DemoMode;
    if (demo && typeof demo.activate === 'function') {
      try {
        demo.activate();
      } catch (_) {}
    } else {
      try {
        localStorage.setItem('ms365-demo-mode-v1', '1');
      } catch (_) {}
    }
  });
  await clearOverlays(page);
  // Dismiss setup banner if present
  await page.evaluate(() => {
    const btn = document.querySelector('[data-ms365-setup-dismiss], #dashSetupDismiss, .dash-setup-banner button');
    if (btn) btn.click();
  }).catch(() => {});
}

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
    } catch (_) {}
  });

  for (const shot of shots) {
    const page = await context.newPage();
    const url = base + shot.url;
    console.log('→', shot.file, url);
    await page.goto(url, { waitUntil: 'networkidle', timeout: 60000 }).catch(async () => {
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
    });
    await page.waitForTimeout(700);
    await clearOverlays(page);

    const killer = setInterval(() => {
      clearOverlays(page).catch(() => {});
    }, 300);

    try {
      if (shot.prepare) await shot.prepare(page);
      await clearOverlays(page);
      await page.waitForTimeout(300);

      const outPath = path.join(outDir, shot.file);
      await page.screenshot({
        path: outPath,
        type: 'png',
        clip: shot.clip || VIEWPORT_CLIP
      });
      const st = fs.statSync(outPath);
      console.log('  saved', shot.file, Math.round(st.size / 1024) + 'KB');
    } finally {
      clearInterval(killer);
      await page.close();
    }
  }

  await browser.close();
  console.log('done');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
