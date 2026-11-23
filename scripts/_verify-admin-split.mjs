import { chromium } from 'playwright';

const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
const errors = [];
page.on('pageerror', (err) => errors.push(String(err)));

await page.addInitScript(() => {
  try {
    localStorage.removeItem('ms365-admin-active-tab-v1');
  } catch (e) {}
});

await page.goto('http://localhost:5173/admin.html#overview', { waitUntil: 'networkidle' });
await page.evaluate(() => {
  const kill = () => document.getElementById('ms365AdminOperatorBlock')?.remove();
  kill();
  new MutationObserver(kill).observe(document.body, { childList: true, subtree: true });
});
await page.waitForTimeout(400);

const overview = await page.evaluate(() => {
  const nav = [...document.querySelectorAll('[data-admin-tab]')].map((b) => b.textContent.replace(/\s+/g, ' ').trim());
  const title = document.getElementById('adminAppTitle')?.textContent || '';
  const ov = document.getElementById('adminPanelOverview');
  const lic = document.getElementById('adminPanelLicenses');
  return {
    nav,
    title,
    hash: location.hash,
    overviewHidden: ov?.hidden ?? null,
    licensesHidden: lic?.hidden ?? null,
    tiles: document.querySelectorAll('#adminLicOverview .admin-lic-tile').length,
    contacts: !!document.getElementById('adminLicContactsTitle'),
    tilesInSchools: !!document.querySelector('#adminPanelLicenses #adminLicOverview'),
    selected: document.querySelector('[data-admin-tab][aria-selected="true"]')?.getAttribute('data-admin-tab')
  };
});

await page.evaluate(() => document.querySelector('[data-admin-tab="licenses"]')?.click());
await page.waitForTimeout(250);
const schools = await page.evaluate(() => ({
  title: document.getElementById('adminAppTitle')?.textContent || '',
  hash: location.hash.split('/')[0],
  selected: document.querySelector('[data-admin-tab][aria-selected="true"]')?.getAttribute('data-admin-tab'),
  overviewHidden: document.getElementById('adminPanelOverview')?.hidden ?? null,
  licensesHidden: document.getElementById('adminPanelLicenses')?.hidden ?? null,
  tilesInSchools: !!document.querySelector('#adminPanelLicenses #adminLicOverview')
}));

await page.evaluate(() => document.querySelector('[data-admin-tab="overview"]')?.click());
await page.waitForTimeout(200);
await page.evaluate(() => {
  const tiles = [...document.querySelectorAll('#adminLicOverview .admin-lic-tile')];
  const due = tiles.find((b) => /Läuft in 30 Tagen ab/.test(b.textContent || ''));
  due?.click();
});
await page.waitForTimeout(250);
const afterTile = await page.evaluate(() => ({
  title: document.getElementById('adminAppTitle')?.textContent || '',
  hash: location.hash.split('/')[0],
  selected: document.querySelector('[data-admin-tab][aria-selected="true"]')?.getAttribute('data-admin-tab'),
  licensesHidden: document.getElementById('adminPanelLicenses')?.hidden ?? null
}));

console.log(JSON.stringify({ overview, schools, afterTile, errors }, null, 2));
await browser.close();
