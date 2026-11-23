import { chromium } from 'playwright';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');
const outDir = path.join(root, 'design-lab', 'shots');
fs.mkdirSync(outDir, { recursive: true });
const base = process.env.MS365_VERIFY_URL || 'http://127.0.0.1:5173';

async function clearGates(page) {
    await page.evaluate(() => {
        document.querySelectorAll(
            '.ms365-license-gate, .pin-gate, [data-ms365-overlay], .ms365-tour, .product-tour, dialog[open], .ms365-modal-backdrop, #ms365PinGate, #licenseGate'
        ).forEach((el) => {
            try {
                el.remove();
            } catch {
                /* ignore */
            }
        });
        document.body.style.pointerEvents = 'auto';
        document.documentElement.style.pointerEvents = 'auto';
        const blockers = Array.from(document.querySelectorAll('body *')).filter((el) => {
            const s = getComputedStyle(el);
            return (
                (s.position === 'fixed' || s.position === 'absolute') &&
                s.pointerEvents !== 'none' &&
                parseInt(s.zIndex || '0', 10) >= 40 &&
                el.getBoundingClientRect().width > window.innerWidth * 0.5
            );
        });
        blockers.forEach((el) => {
            try {
                el.style.display = 'none';
                el.style.pointerEvents = 'none';
            } catch {
                /* ignore */
            }
        });
    });
}

async function shot(page, name) {
    await page.screenshot({ path: path.join(outDir, name), fullPage: false });
}

async function main() {
    const browser = await chromium.launch({ headless: true });
    const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });

    // Catalog fix check
    await page.goto(base + '/index.html', { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForTimeout(1000);
    await clearGates(page);
    await page.evaluate(() => {
        document.documentElement.setAttribute('data-dash-it-shell', '');
        document.documentElement.setAttribute('data-theme', 'light');
        const btn = Array.from(document.querySelectorAll('button, a, [role="tab"]')).find((b) =>
            /Alle Werkzeuge/i.test(b.textContent || '')
        );
        if (btn) btn.click();
        document.dispatchEvent(new CustomEvent('ms365-dash-catalog-ready'));
    });
    await page.waitForTimeout(700);
    await clearGates(page);
    const catalog = await page.evaluate(() => {
        const catalogEl = document.getElementById('dash-catalog');
        const nav = catalogEl && catalogEl.querySelector('.dash-tasks-split__nav');
        const detail = catalogEl && catalogEl.querySelector('.dash-tasks-split__detail');
        if (!catalogEl || !nav || !detail) return { ok: false, reason: 'missing' };
        const nr = nav.getBoundingClientRect();
        const dr = detail.getBoundingClientRect();
        const gap = Math.round(dr.left - nr.right);
        return {
            ok: gap >= 0 && gap <= 28,
            gap,
            cols: getComputedStyle(catalogEl).gridTemplateColumns
        };
    });
    await shot(page, 'catalog-split-fix.png');

    // Stammdaten rail
    await page.goto(base + '/tenant.html', { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForTimeout(1000);
    await clearGates(page);
    await page.evaluate(() => {
        document.documentElement.setAttribute('data-theme', 'light');
        document.body.classList.add('app-shell-chrome', 'ts-admin-page');
    });
    const railLight = await page.evaluate(() => {
        const rail = document.querySelector('.ts-admin-rail');
        const tabs = document.getElementById('tenantMainTabs');
        const shell = document.querySelector('.ts-admin-shell');
        if (!rail || !tabs || !shell) return { ok: false };
        const rr = rail.getBoundingClientRect();
        const sr = shell.getBoundingClientRect();
        const cols = getComputedStyle(shell).gridTemplateColumns;
        return {
            ok: rr.width >= 200 && rr.width <= 300 && sr.left < 40,
            railWidth: Math.round(rr.width),
            cols,
            tabCount: tabs.querySelectorAll('.tab-btn').length
        };
    });
    await clearGates(page);
    await page.evaluate(() => {
        const btn = document.getElementById('tabMainKlassen');
        if (btn) btn.click();
        if (typeof window.ms365TenantTabActivate === 'function') {
            window.ms365TenantTabActivate('tabMainKlassen');
        }
    });
    await page.waitForTimeout(400);
    await clearGates(page);
    await shot(page, 'stammdaten-klassen-light.png');

    await page.evaluate(() => document.documentElement.setAttribute('data-theme', 'dark'));
    await page.waitForTimeout(200);
    await clearGates(page);
    await shot(page, 'stammdaten-klassen-dark.png');

    await page.evaluate(() => {
        const btn = document.getElementById('tabMainStammdaten');
        if (btn) btn.click();
        if (typeof window.ms365TenantTabActivate === 'function') {
            window.ms365TenantTabActivate('tabMainStammdaten');
        }
    });
    await page.waitForTimeout(300);
    await clearGates(page);
    await shot(page, 'stammdaten-profil-dark.png');

    // Auth menu brand box gone
    await page.goto(base + '/index.html', { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForTimeout(800);
    await clearGates(page);
    await page.evaluate(() => {
        const t = document.getElementById('ms365AuthBadge');
        if (t) t.click();
    });
    await page.waitForTimeout(200);
    const auth = await page.evaluate(() => {
        const note = document.querySelector('.ms365-auth-menu__brand-note');
        const text = document.getElementById('ms365AuthDropdown')?.innerText || '';
        return {
            hasBrandNote: !!(note && note.offsetParent !== null && note.textContent.trim()),
            hasMarkenText: /Markenfarbe/.test(text)
        };
    });
    await shot(page, 'auth-menu-no-brand.png');

    console.log(JSON.stringify({ catalog, railLight, auth }, null, 2));
    await browser.close();
    if (!catalog.ok || !railLight.ok || auth.hasMarkenText) process.exit(2);
}

main().catch((e) => {
    console.error(e);
    process.exit(1);
});
