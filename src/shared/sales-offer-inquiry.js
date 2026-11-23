/**
 * Angebots-Empfehlung und mailto-Anfrage (Landing + Tests).
 * Baut auf sales-bundles-catalog und license-feature-catalog auf.
 */
import {
    SALES_ADDONS,
    SALES_BUNDLES,
    matchSalesBundle,
    salesBundleById
} from './sales-bundles-catalog.js';
import {
    LICENSE_PACKAGES,
    packageById,
    packageLabel,
    sumPackageListPrices
} from './license-feature-catalog.js';

export const SALES_INQUIRY_MAIL = 'kontakt@ms365.schule';

/**
 * @param {string[]|null|undefined} packageIds
 * @returns {string[]}
 */
export function normalizePackageIds(packageIds) {
    const seen = new Set();
    const out = [];
    (Array.isArray(packageIds) ? packageIds : []).forEach((raw) => {
        const id = String(raw || '')
            .trim()
            .toLowerCase();
        if (!id || seen.has(id)) return;
        seen.add(id);
        out.push(id);
    });
    return out;
}

/**
 * Abhängigkeiten ergänzen (intranet für schulapps/ablauf/…).
 * @param {string[]|null|undefined} packageIds
 * @returns {string[]}
 */
export function expandPackageDeps(packageIds) {
    const set = new Set(normalizePackageIds(packageIds));
    let changed = true;
    while (changed) {
        changed = false;
        for (const id of [...set]) {
            const pkg = packageById(id);
            const req = (pkg && pkg.requires) || [];
            for (const r of req) {
                if (!set.has(r)) {
                    set.add(r);
                    changed = true;
                }
            }
        }
    }
    return LICENSE_PACKAGES.map((p) => p.id).filter((id) => set.has(id));
}

/**
 * @param {string[]} selected
 * @param {SalesBundle} bundle
 */
function diffToBundle(selected, bundle) {
    const sel = new Set(selected);
    const need = bundle.packageIds || [];
    const missing = need.filter((id) => !sel.has(id));
    const extra = selected.filter((id) => need.indexOf(id) < 0);
    return { missing, extra };
}

/**
 * @typedef {{
 *   mode: 'bundle'|'custom',
 *   bundle: import('./sales-bundles-catalog.js').SalesBundle|null,
 *   nearest: import('./sales-bundles-catalog.js').SalesBundle|null,
 *   listSum: number,
 *   offerEuro: number|null,
 *   tip: string,
 *   missing: string[],
 *   extra: string[]
 * }} SalesOfferRecommendation
 */

/**
 * Entscheidungsgrundlage: Bundle-Treffer oder individuelles Paket.
 * @param {string[]|null|undefined} packageIds
 * @returns {SalesOfferRecommendation}
 */
export function recommendSalesOffer(packageIds) {
    const selected = expandPackageDeps(packageIds);
    const listSum = sumPackageListPrices(selected);
    const empty = {
        mode: 'custom',
        bundle: null,
        nearest: null,
        listSum,
        offerEuro: listSum || null,
        tip: 'Wählen Sie Module oder eines der drei Bundles – wir melden uns mit einer konkreten Empfehlung.',
        missing: [],
        extra: []
    };
    if (!selected.length) return empty;

    const matched = matchSalesBundle(selected);
    if (matched.bundle && matched.exact) {
        const save = Math.max(0, listSum - (matched.bundle.priceEuro || 0));
        return {
            mode: 'bundle',
            bundle: matched.bundle,
            nearest: matched.bundle,
            listSum,
            offerEuro: matched.bundle.priceEuro,
            tip: save
                ? 'Gegenüber Einzelmodulen ca. ' + save + ' € günstiger.'
                : 'Schulpauschale statt Einzelmodule.',
            missing: [],
            extra: []
        };
    }

    let nearest = null;
    let bestScore = Infinity;
    let bestSize = -1;
    for (const bundle of SALES_BUNDLES) {
        const { missing, extra } = diffToBundle(selected, bundle);
        const score = missing.length * 2 + extra.length;
        const size = (bundle.packageIds || []).length;
        if (score < bestScore || (score === bestScore && size > bestSize)) {
            bestScore = score;
            bestSize = size;
            nearest = bundle;
        }
    }

    if (!nearest) {
        return {
            mode: 'custom',
            bundle: null,
            nearest: null,
            listSum,
            offerEuro: listSum,
            tip: 'Individuelle Zusammenstellung – wir prüfen mit Ihnen die passende Schulpauschale.',
            missing: [],
            extra: []
        };
    }

    const { missing, extra } = diffToBundle(selected, nearest);

    if (missing.length > 0 && missing.length <= 3) {
        return {
            mode: 'custom',
            bundle: null,
            nearest,
            listSum,
            offerEuro: listSum,
            tip:
                'Für viele Schulen lohnt „' +
                nearest.label +
                '“ (' +
                nearest.priceLabel +
                '). Es fehlt noch: ' +
                missing.map((id) => packageLabel(id) || id).join(', ') +
                '.',
            missing,
            extra
        };
    }

    if (missing.length === 0 && extra.length > 0) {
        const extraSum = sumPackageListPrices(extra);
        return {
            mode: 'custom',
            bundle: null,
            nearest,
            listSum,
            offerEuro: listSum,
            tip:
                'Kern passt zu „' +
                nearest.label +
                '“ (' +
                nearest.priceLabel +
                '). Zusätzlich gewählt: ' +
                extra.map((id) => packageLabel(id) || id).join(', ') +
                (extraSum ? ' (ca. +' + extraSum + ' € Listenpreis).' : '.'),
            missing,
            extra
        };
    }

    return {
        mode: 'custom',
        bundle: null,
        nearest,
        listSum,
        offerEuro: listSum,
        tip:
            'Individuelle Zusammenstellung (Listenpreis-Orientierung ' +
            listSum +
            ' €). Nächstes Bundle: „' +
            nearest.label +
            '“ – oft günstiger als Modul für Modul.',
        missing,
        extra
    };
}

/**
 * @param {{
 *   to?: string,
 *   name?: string,
 *   school?: string,
 *   email?: string,
 *   message?: string,
 *   packageIds?: string[],
 *   addonIds?: string[],
 *   recommendation?: SalesOfferRecommendation|null
 * }} input
 * @returns {{ href: string, subject: string, body: string }}
 */
export function buildSalesInquiryMailto(input) {
    const to = String((input && input.to) || SALES_INQUIRY_MAIL).trim();
    const name = String((input && input.name) || '').trim();
    const school = String((input && input.school) || '').trim();
    const email = String((input && input.email) || '').trim();
    const message = String((input && input.message) || '').trim();
    const packageIds = expandPackageDeps((input && input.packageIds) || []);
    const addonIds = normalizePackageIds((input && input.addonIds) || []);
    const rec =
        (input && input.recommendation) || recommendSalesOffer(packageIds);

    const choiceLabel =
        rec.mode === 'bundle' && rec.bundle
            ? 'Bundle: ' + rec.bundle.label + ' (' + rec.bundle.priceLabel + ')'
            : 'Individuell (Listenpreis-Orientierung ' +
              (rec.listSum || 0) +
              ' € / Schuljahr)';

    const moduleLines = packageIds.length
        ? packageIds
              .map((id) => {
                  const p = packageById(id);
                  const price =
                      p && typeof p.priceEuro === 'number'
                          ? p.priceEuro + ' €'
                          : '–';
                  return '- ' + (packageLabel(id) || id) + ' (' + price + ')';
              })
              .join('\n')
        : '- (noch keine Module gewählt)';

    const addonLines = addonIds.length
        ? addonIds
              .map((id) => {
                  const a = SALES_ADDONS.find((x) => x.id === id);
                  return a
                      ? '- ' + a.label + ' (' + a.priceLabel + ')'
                      : '- ' + id;
              })
              .join('\n')
        : '- keine';

    const subject =
        'Angebotsanfrage MS365-Schul-Tools – ' +
        (rec.mode === 'bundle' && rec.bundle ? rec.bundle.label : 'Individuell');

    const body = [
        'Guten Tag,',
        '',
        'ich interessiere mich für MS365-Schul-Tools.',
        '',
        'Schule: ' + (school || '–'),
        'Name: ' + (name || '–'),
        'E-Mail: ' + (email || '–'),
        '',
        'Auswahl: ' + choiceLabel,
        '',
        'Module:',
        moduleLines,
        '',
        'Add-ons / Einrichtung:',
        addonLines,
        '',
        'Empfehlung auf der Homepage:',
        rec.tip || '–',
        '',
        'Nachricht:',
        message || '–',
        '',
        'Mit freundlichen Grüßen',
        name || ''
    ].join('\n');

    const href =
        'mailto:' +
        encodeURIComponent(to).replace(/%40/g, '@') +
        '?subject=' +
        encodeURIComponent(subject) +
        '&body=' +
        encodeURIComponent(body);

    return { href, subject, body, to };
}

export function salesAddonById(id) {
    const want = String(id || '')
        .trim()
        .toLowerCase();
    return SALES_ADDONS.find((a) => a.id === want) || null;
}

export function salesBundlePackageSet(bundleId) {
    const b = salesBundleById(bundleId);
    return b ? (b.packageIds || []).slice() : [];
}
