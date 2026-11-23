/**
 * Verify Alle Werkzeuge catalog split: content starts next to rail (no mid gap).
 */
import { chromium } from 'playwright';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');
const outDir = path.join(root, 'design-lab', 'shots');
const url = process.env.MS365_VERIFY_URL || 'http://127.0.0.1:5173/index.html';

async function clearOverlays(page) {
    await page.evaluate(() => {
        document.querySelectorAll(
            '.ms365-license-gate, .pin-gate, [data-ms365-overlay], .ms365-tour, .product-tour'
        ).forEach((el) => {
            try {
                el.remove();
            } catch {
                /* ignore */
            }
        });
        document.documentElement.setAttribute('data-dash-it-shell', '');
        document.body.classList.add('dash-catalog-split-active');
        const sections = document.querySelectorAll('#dashCatalogSection, #dashboard-tools, #dash-catalog');
        sections.forEach((el) => {
            el.hidden = false;
            el.style.display = '';
            el.style.visibility = 'visible';
        });
        // Switch to Alle Werkzeuge if toggle exists
        const btn =
            document.querySelector('[data-dash-main="catalog"], [data-dash-view="catalog"], #dashMainCatalog') ||
            Array.from(document.querySelectorAll('button, a')).find((b) =>
                /Alle Werkzeuge/i.test(b.textContent || '')
            );
        if (btn) btn.click();
    });
}

async function main() {
    const browser = await chromium.launch({ headless: true });
    const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForTimeout(1200);
    await clearOverlays(page);
    await page.waitForTimeout(800);

    // Ensure split mounted
    await page.evaluate(async () => {
        try {
            const mod = await import('/src/shared/dashboard-catalog-split.js');
            if (mod && typeof mod.mountDashboardCatalogSplit === 'function') {
                mod.mountDashboardCatalogSplit();
            }
        } catch {
            /* ignore */
        }
        document.dispatchEvent(new CustomEvent('ms365-dash-catalog-ready'));
    });
    await page.waitForTimeout(500);
    await clearOverlays(page);

    const metrics = await page.evaluate(() => {
        const catalog = document.getElementById('dash-catalog');
        const nav = catalog && catalog.querySelector('.dash-tasks-split__nav');
        const detail = catalog && catalog.querySelector('.dash-tasks-split__detail');
        const personal = catalog && catalog.querySelector('.dash-catalog-split__personal');
        if (!catalog || !nav || !detail) {
            return {
                ok: false,
                reason: 'missing nodes',
                hasCatalog: !!catalog,
                hasNav: !!nav,
                hasDetail: !!detail,
                classes: catalog ? catalog.className : ''
            };
        }
        const cr = catalog.getBoundingClientRect();
        const nr = nav.getBoundingClientRect();
        const dr = detail.getBoundingClientRect();
        const pr = personal ? personal.getBoundingClientRect() : null;
        const cs = getComputedStyle(catalog);
        const gap = Math.round(dr.left - nr.right);
        const ok = gap >= 0 && gap <= 24 && Math.abs(dr.left - (nr.right + gap)) < 2;
        return {
            ok,
            gap,
            navWidth: Math.round(nr.width),
            navRight: Math.round(nr.right),
            detailLeft: Math.round(dr.left),
            detailWidth: Math.round(dr.width),
            catalogWidth: Math.round(cr.width),
            personalWidth: pr ? Math.round(pr.width) : null,
            personalLeft: pr ? Math.round(pr.left) : null,
            gridTemplateColumns: cs.gridTemplateColumns,
            classes: catalog.className
        };
    });

    await page.screenshot({
        path: path.join(outDir, 'catalog-split-fix.png'),
        fullPage: false
    });

    console.log(JSON.stringify(metrics, null, 2));
    await browser.close();
    if (!metrics.ok) process.exit(2);
}

main().catch((err) => {
    console.error(err);
    process.exit(1);
});
