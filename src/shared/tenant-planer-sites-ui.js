/**
 * Stammdaten: Primär-Site / Planer & Apps Sites + zentrale Postfächer.
 */
import {
    SCHOOL_TOOL_SITE_KINDS,
    normalizeSchoolToolSitesSetup,
    patchSchoolToolSitesSetup,
    resolveAllSchoolToolSiteUrls,
    resolveSchoolToolSiteUrl,
    suggestPrimarySiteUrl
} from './school-tool-sites.js';
import {
    readIntranetSiteUrl,
    readSchoolIntranetSiteUrl
} from './stammdaten-intranet-listen-ui.js';
import { activateStammdatenSetupPane } from './tenant-stammdaten-setup-nav.js';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.Read.All'
];

const TOOL_ICONS = {
    schularbeiten: 'bi-journal-check',
    projektwochen: 'bi-calendar2-week',
    freistellung: 'bi-person-walking',
    lehrerFreistellung: 'bi-briefcase',
    aktivitaeten: 'bi-bus-front',
    wahlfach: 'bi-ui-checks-grid'
};

function el(id) {
    return document.getElementById(id);
}

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function setStatus(id, text, tone) {
    const status = el(id);
    if (!status) return;
    status.textContent = text || '';
    status.dataset.tone = tone || '';
}

function toast(m, kind) {
    if (typeof window.ms365ShowToast === 'function') {
        window.ms365ShowToast(m, { kind: kind || 'info', title: 'Planer & Apps' });
        return;
    }
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
}

function shortSiteLabel(url) {
    const raw = String(url || '').trim();
    if (!raw) return 'Noch keine Site';
    try {
        const u = new URL(raw);
        const m = u.pathname.match(/\/sites\/([^/]+)/i);
        if (m && m[1]) return '/sites/' + decodeURIComponent(m[1]);
        return u.hostname || raw;
    } catch {
        const m = raw.match(/\/sites\/([^/?#]+)/i);
        if (m && m[1]) return '/sites/' + m[1];
        return raw.length > 42 ? raw.slice(0, 40) + '…' : raw;
    }
}

function sameSite(a, b) {
    const x = String(a || '')
        .trim()
        .replace(/\/+$/, '')
        .toLowerCase();
    const y = String(b || '')
        .trim()
        .replace(/\/+$/, '')
        .toLowerCase();
    return !!(x && y && x === y);
}

function readFormPatch() {
    return {
        planerDefaultSiteUrl: String((el('tenantPlanerDefaultSiteUrl') && el('tenantPlanerDefaultSiteUrl').value) || '').trim(),
        planerSiteUrls: {
            schularbeiten: String((el('tenantPlanerSiteSa') && el('tenantPlanerSiteSa').value) || '').trim(),
            projektwochen: String((el('tenantPlanerSitePw') && el('tenantPlanerSitePw').value) || '').trim(),
            freistellung: String((el('tenantPlanerSiteFr') && el('tenantPlanerSiteFr').value) || '').trim(),
            aktivitaeten: String((el('tenantPlanerSiteAkt') && el('tenantPlanerSiteAkt').value) || '').trim(),
            wahlfach: String((el('tenantPlanerSiteWz') && el('tenantPlanerSiteWz').value) || '').trim()
        },
        schoolMailboxes: {
            elternbriefZentral: String((el('tenantMailboxElternbrief') && el('tenantMailboxElternbrief').value) || '')
                .trim()
                .toLowerCase(),
            elternbriefName: String(
                (el('tenantMailboxElternbriefName') && el('tenantMailboxElternbriefName').value) || 'Elternbrief Schule'
            ).trim(),
            schulkalender: String((el('tenantMailboxSchulkalender') && el('tenantMailboxSchulkalender').value) || '')
                .trim()
                .toLowerCase()
        }
    };
}

function fillFormFromSetup() {
    let setup = {};
    try {
        setup =
            window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function'
                ? window.ms365AppDataV2.getSetup() || {}
                : {};
    } catch {
        setup = {};
    }
    const n = normalizeSchoolToolSitesSetup(setup);
    const def = n.planerDefaultSiteUrl || suggestPrimarySiteUrl(setup);
    const defEl = el('tenantPlanerDefaultSiteUrl');
    if (defEl && !String(defEl.value || '').trim()) defEl.value = def;
    else if (defEl && n.planerDefaultSiteUrl) defEl.value = n.planerDefaultSiteUrl;

    const map = {
        tenantPlanerSiteSa: n.planerSiteUrls.schularbeiten,
        tenantPlanerSitePw: n.planerSiteUrls.projektwochen,
        tenantPlanerSiteFr: n.planerSiteUrls.freistellung,
        tenantPlanerSiteAkt: n.planerSiteUrls.aktivitaeten,
        tenantPlanerSiteWz: n.planerSiteUrls.wahlfach,
        tenantMailboxElternbrief: n.schoolMailboxes.elternbriefZentral,
        tenantMailboxElternbriefName: n.schoolMailboxes.elternbriefName,
        tenantMailboxSchulkalender: n.schoolMailboxes.schulkalender
    };
    Object.keys(map).forEach(function (id) {
        const input = el(id);
        if (input) input.value = map[id] || '';
    });

    refreshStatusUi();
}

function refreshStatusUi() {
    const list = el('tenantPrimarySiteMapHint');
    const meta = el('tenantPlanerAppsStatusMeta');
    const setup = (function () {
        try {
            return window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function'
                ? window.ms365AppDataV2.getSetup() || {}
                : {};
        } catch {
            return {};
        }
    })();
    const normalized = normalizeSchoolToolSitesSetup(setup);
    const primary =
        String((el('tenantPlanerDefaultSiteUrl') && el('tenantPlanerDefaultSiteUrl').value) || '').trim() ||
        normalized.planerDefaultSiteUrl ||
        suggestPrimarySiteUrl(setup);
    const urls = resolveAllSchoolToolSiteUrls(setup);

    let linked = 0;
    let overrides = 0;
    const rows = SCHOOL_TOOL_SITE_KINDS.map(function (k) {
        const url = urls[k.id] || resolveSchoolToolSiteUrl(k.id, setup) || '';
        const has = !!url;
        if (has) linked += 1;
        const isPrimary = has && sameSite(url, primary);
        if (has && !isPrimary) overrides += 1;
        const tag = !has ? 'offen' : isPrimary ? 'Primär-Site' : 'Eigene Site';
        const tagClass = !has ? '' : isPrimary ? ' is-primary' : ' is-override';
        const iconClass = !has ? 'is-empty' : 'is-ok';
        const icon = TOOL_ICONS[k.id] || 'bi-box';
        return (
            '<li class="ts-planer-apps-status__item">' +
            '<span class="ts-planer-apps-status__icon ' +
            iconClass +
            '" aria-hidden="true"><i class="bi ' +
            icon +
            '"></i></span>' +
            '<span class="ts-planer-apps-status__body">' +
            '<span class="ts-planer-apps-status__name">' +
            escapeHtml(k.label) +
            '</span>' +
            '<span class="ts-planer-apps-status__site" title="' +
            escapeHtml(url || '') +
            '">' +
            escapeHtml(shortSiteLabel(url)) +
            '</span>' +
            '</span>' +
            '<span class="ts-planer-apps-status__tag' +
            tagClass +
            '">' +
            escapeHtml(tag) +
            '</span>' +
            '</li>'
        );
    });

    if (list) list.innerHTML = rows.join('');
    if (meta) {
        meta.textContent =
            linked +
            ' von ' +
            SCHOOL_TOOL_SITE_KINDS.length +
            ' verbunden' +
            (overrides ? ' · ' + overrides + ' mit eigener Site' : '');
    }

    const eb = String((el('tenantMailboxElternbrief') && el('tenantMailboxElternbrief').value) || '').trim();
    const sk = String((el('tenantMailboxSchulkalender') && el('tenantMailboxSchulkalender').value) || '').trim();
    const ebBadge = el('tenantMailboxElternbriefBadge');
    const skBadge = el('tenantMailboxSchulkalenderBadge');
    if (ebBadge) {
        ebBadge.hidden = !eb;
        ebBadge.textContent = eb ? 'gesetzt' : '';
    }
    if (skBadge) {
        skBadge.hidden = !sk;
        skBadge.textContent = sk ? 'gesetzt' : '';
    }

    // Overrides-Details öffnen, wenn bereits Werte gesetzt
    const details = el('tenantPlanerSiteOverrides');
    if (details && !details.dataset.autoOpened) {
        const hasOverride = Object.keys(normalized.planerSiteUrls).some(function (key) {
            return !!normalized.planerSiteUrls[key];
        });
        if (hasOverride) details.open = true;
        details.dataset.autoOpened = '1';
    }
}

function persistForm(opts) {
    const patch = readFormPatch();
    if (opts && opts.applyLocalEmpty) patch.applyLocalEmpty = true;
    patchSchoolToolSitesSetup(patch);
    refreshStatusUi();
    try {
        window.dispatchEvent(new CustomEvent('ms365-school-data-written', { detail: { source: 'planer-sites' } }));
    } catch {
        /* ignore */
    }
}

async function verifyDefaultSite() {
    const input = el('tenantPlanerDefaultSiteUrl');
    const webUrl = String((input && input.value) || '').trim();
    setStatus('tenantPlanerDefaultSiteStatus', '', '');
    if (!webUrl) {
        setStatus('tenantPlanerDefaultSiteStatus', 'Bitte zuerst eine SharePoint-Webadresse eintragen.', 'warn');
        return;
    }
    const G = window.ms365SpoGraph;
    if (!G || typeof G.getGraphToken !== 'function' || typeof G.resolveSiteFromWebUrl !== 'function') {
        setStatus('tenantPlanerDefaultSiteStatus', 'SharePoint-Hilfen nicht geladen – Seite neu laden.', 'error');
        return;
    }
    setStatus('tenantPlanerDefaultSiteStatus', 'Prüfe Website …', 'pending');
    try {
        const token = await G.getGraphToken(SCOPES);
        const site = await G.resolveSiteFromWebUrl(token, webUrl);
        const name = site && site.displayName ? String(site.displayName).trim() : '';
        const id = site && site.id ? String(site.id) : '';
        if (!id) throw new Error('Site konnte nicht aufgelöst werden.');
        if (input) input.value = webUrl;
        persistForm();
        setStatus(
            'tenantPlanerDefaultSiteStatus',
            (name ? 'Website gefunden: „' + name + '“' : 'Website erkannt.') +
                (site.webUrl ? ' · ' + site.webUrl : ''),
            'ok'
        );
    } catch (e) {
        setStatus(
            'tenantPlanerDefaultSiteStatus',
            'Nicht erreichbar oder keine Berechtigung: ' + (e && e.message ? e.message : String(e)),
            'error'
        );
    }
}

function bindAutosave() {
    const ids = [
        'tenantPlanerDefaultSiteUrl',
        'tenantPlanerSiteSa',
        'tenantPlanerSitePw',
        'tenantPlanerSiteFr',
        'tenantPlanerSiteAkt',
        'tenantPlanerSiteWz',
        'tenantMailboxElternbrief',
        'tenantMailboxElternbriefName',
        'tenantMailboxSchulkalender'
    ];
    ids.forEach(function (id) {
        const input = el(id);
        if (!input || input.dataset.planerSitesBound === '1') return;
        input.dataset.planerSitesBound = '1';
        let timer = null;
        const run = function () {
            persistForm();
        };
        input.addEventListener('input', function () {
            if (timer) clearTimeout(timer);
            timer = setTimeout(run, 450);
        });
        input.addEventListener('change', run);
    });
}

function bindActions() {
    const verify = el('tenantPlanerDefaultSiteVerify');
    if (verify && verify.dataset.planerSitesBound !== '1') {
        verify.dataset.planerSitesBound = '1';
        verify.addEventListener('click', function () {
            void verifyDefaultSite();
        });
    }

    const fromIntranet = el('tenantPlanerDefaultFromIntranet');
    if (fromIntranet && fromIntranet.dataset.planerSitesBound !== '1') {
        fromIntranet.dataset.planerSitesBound = '1';
        fromIntranet.addEventListener('click', function () {
            const u = readIntranetSiteUrl();
            const input = el('tenantPlanerDefaultSiteUrl');
            if (!u) {
                toast('Keine Intranet-Listen-URL gesetzt.', 'warn');
                return;
            }
            if (input) input.value = u;
            persistForm();
            toast('Primär-Site von Intranet-Listen übernommen.');
        });
    }

    const fromHub = el('tenantPlanerDefaultFromHub');
    if (fromHub && fromHub.dataset.planerSitesBound !== '1') {
        fromHub.dataset.planerSitesBound = '1';
        fromHub.addEventListener('click', function () {
            const u = readSchoolIntranetSiteUrl();
            const input = el('tenantPlanerDefaultSiteUrl');
            if (!u) {
                toast('Keine Schul-Intranet-URL gesetzt.', 'warn');
                return;
            }
            if (input) input.value = u;
            persistForm();
            toast('Primär-Site vom Schul-Intranet übernommen.');
        });
    }

    const applyEmpty = el('tenantPlanerApplyEmptyTools');
    if (applyEmpty && applyEmpty.dataset.planerSitesBound !== '1') {
        applyEmpty.dataset.planerSitesBound = '1';
        applyEmpty.addEventListener('click', function () {
            persistForm({ applyLocalEmpty: true });
            toast('Leere Tool-Sites mit Primär-Site befüllt.');
        });
    }

    document.querySelectorAll('[data-stammdaten-goto-pane]').forEach(function (btn) {
        if (btn.dataset.planerGotoBound === '1') return;
        btn.dataset.planerGotoBound = '1';
        btn.addEventListener('click', function () {
            const pane = String(btn.getAttribute('data-stammdaten-goto-pane') || '').trim();
            if (pane) activateStammdatenSetupPane(pane, { scroll: true });
        });
    });
}

export function mountTenantPlanerSitesUi() {
    const panel = el('tenantPlanerSitesPanel');
    if (!panel) return;
    fillFormFromSetup();
    bindAutosave();
    bindActions();
}

export default { mountTenantPlanerSitesUi };
