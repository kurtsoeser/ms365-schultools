/**
 * First-Run-Tour für Lehrkräfte: Willkommen → Portal-Spotlight (Apps, Bereiche, Feedback).
 */
import { getDashboardPersonas } from './dashboard-persona-session.js';
import {
    isLehrerWelcomeTourDone,
    LEHRER_FIRST_RUN_CHIP_DISMISS_KEY,
    LEHRER_FIRST_RUN_KEY,
    LEHRER_WELCOME_KEY,
    lehrerChipLabelFromPhase,
    markLehrerWelcomeTourDone,
    patchLehrerFirstRunState,
    readLehrerFirstRunState,
    shouldShowLehrerFirstRunChip,
    shouldShowLehrerTourMenuButton,
    shouldStartLehrerWelcomeTour
} from './product-tour-logic.js';
import { openWelcomeTour, startSpotlightTour } from './product-tour.js';
import {
    gatherLehrerPortalReadinessInput,
    evaluateLehrerPortalReadiness,
    isLehrerPortalReady
} from './lehrer-portal-readiness.js';

const FEEDBACK_HINT =
    '<p><strong>Feedback ist erwünscht.</strong> In den Pilotwochen hilft uns jeder Hinweis – unten in der Fußzeile unter <em>Feedback</em> oder jederzeit über denselben Dialog.</p>';

function openFeedbackPrefill(title) {
    try {
        if (window.ms365Feedback && typeof window.ms365Feedback.open === 'function') {
            window.ms365Feedback.open({ title: title || 'Lehrer First-Run / Pilot' });
            return;
        }
    } catch {
        /* ignore */
    }
    try {
        document.querySelector('.site-feedback-link')?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    } catch {
        /* ignore */
    }
}

function waitForSelector(selector, timeoutMs) {
    return new Promise((resolve) => {
        const existing = document.querySelector(selector);
        if (existing) {
            resolve(existing);
            return;
        }
        const started = Date.now();
        const timer = window.setInterval(() => {
            const el = document.querySelector(selector);
            if (el || Date.now() - started > (timeoutMs || 2500)) {
                window.clearInterval(timer);
                resolve(el || null);
            }
        }, 120);
    });
}

function readCachedPersonas() {
    try {
        const api = window.ms365DashboardAudience;
        if (api && typeof api.getPersonas === 'function') return api.getPersonas();
    } catch {
        /* ignore */
    }
    return null;
}

function readLayoutMode() {
    try {
        const api = window.ms365DashboardAudience;
        if (api && typeof api.getLayoutMode === 'function') return String(api.getLayoutMode() || '');
    } catch {
        /* ignore */
    }
    try {
        if (document.documentElement.hasAttribute('data-dash-app-only')) return 'apps-lehrer';
    } catch {
        /* ignore */
    }
    return '';
}

function isPortalPreviewLayout(mode) {
    const m = String(mode || '');
    return (
        m === 'preview-lehrer' ||
        m === 'apps-lehrer' ||
        document.documentElement.hasAttribute('data-dash-app-only')
    );
}

function isLoggedInAccount() {
    try {
        if (typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            return !!(info && (info.oid || info.upn || info.username));
        }
    } catch {
        /* ignore */
    }
    try {
        if (typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn()) return true;
    } catch {
        /* ignore */
    }
    const p = readCachedPersonas();
    return !!(p && p.loggedIn);
}

/**
 * @param {number} [timeoutMs]
 * @returns {Promise<{ loggedIn: boolean, isLehrer: boolean, isIt: boolean }>}
 */
async function resolveLehrerPersona(timeoutMs) {
    const budget = Math.max(500, Number(timeoutMs) || 8000);
    const started = Date.now();

    const fromCache = () => {
        const p = readCachedPersonas();
        if (p && p.loggedIn) {
            return { loggedIn: true, isLehrer: !!p.isLehrer, isIt: !!p.isIt };
        }
        return null;
    };

    const hit = fromCache();
    if (hit && hit.isLehrer) return hit;

    try {
        const personas = await getDashboardPersonas({ force: false });
        if (personas && personas.loggedIn) {
            return {
                loggedIn: true,
                isLehrer: !!personas.isLehrer,
                isIt: !!personas.isIt
            };
        }
    } catch {
        /* ignore */
    }

    while (Date.now() - started < budget) {
        const again = fromCache();
        if (again && again.isLehrer) return again;
        await new Promise((r) => window.setTimeout(r, 300));
        try {
            const personas = await getDashboardPersonas({ force: false });
            if (personas && personas.loggedIn) {
                return {
                    loggedIn: true,
                    isLehrer: !!personas.isLehrer,
                    isIt: !!personas.isIt
                };
            }
        } catch {
            /* ignore */
        }
    }

    const late = fromCache();
    if (late) return late;
    return { loggedIn: isLoggedInAccount(), isLehrer: false, isIt: false };
}

function ensurePortalHome() {
    try {
        const homeBtn = document.querySelector('[data-portal-nav="home"]');
        if (homeBtn && homeBtn.getAttribute('aria-selected') !== 'true' && typeof homeBtn.click === 'function') {
            homeBtn.click();
        }
    } catch {
        /* ignore */
    }
}

function clickPortalNav(navId) {
    try {
        const btn = document.querySelector('[data-portal-nav="' + navId + '"]');
        if (btn && typeof btn.click === 'function') btn.click();
    } catch {
        /* ignore */
    }
}

function welcomePanels() {
    return [
        {
            id: 'role',
            title: 'Sie sind im Lehrer-Portal',
            bodyHtml:
                '<p>Hier finden Sie die <strong>Alltags-Apps</strong> für den Unterricht und den Schulbetrieb – Freistellungen, Schularbeiten, Aktivitäten, Termine und mehr.</p>' +
                '<ul><li>Kein IT-Backend: Einrichtung und Stammdaten macht die Schul-IT</li>' +
                '<li>Ihre Apps hängen an der SharePoint-Welt Ihrer Schule</li>' +
                '<li>Bereiche wie Unterricht, Schulalltag oder Eltern gliedern den Alltag</li></ul>'
        },
        {
            id: 'apps',
            title: 'Ausprobieren ist erwünscht',
            bodyHtml:
                '<p>Auf dem Dashboard sehen Sie zuerst <strong>Meine Apps</strong> – die wichtigsten Einstiege. Über die Tabs links öffnen Sie weitere Bereiche und Tools.</p>' +
                '<p class="muted">Fehlt eine App? Oft muss die Schul-IT noch die Tool-Site oder die Lizenz freischalten – kurz nachfragen oder Feedback senden.</p>'
        },
        {
            id: 'feedback',
            title: 'Pilotwochen: Feedback bitte',
            primaryLabel: 'Tour auf dem Portal starten',
            bodyHtml:
                FEEDBACK_HINT +
                '<p>Unklar, hakt oder fehlt etwas? Kurz melden – wir priorisieren danach die nächsten Verbesserungen.</p>'
        }
    ];
}

function portalSpotlightSteps() {
    return [
        {
            id: 'hero',
            selector: '.dash-portal-hero, #dashQuickApps',
            title: 'Ihr Lehrer-Dashboard',
            bodyHtml:
                '<p>Begrüßung und Schulkarte – der Einstieg ins Portal. Von hier aus starten Sie in den Schulalltag, nicht ins IT-Backend.</p>',
            placement: 'bottom',
            before: async () => {
                ensurePortalHome();
                await waitForSelector('.dash-portal-hero, #dashQuickApps', 2500);
            }
        },
        {
            id: 'nav',
            selector: '.dash-portal-shell__nav, [data-portal-nav]',
            title: 'Bereiche für den Alltag',
            bodyHtml:
                '<p>Links liegen die Bereiche. Ein Klick öffnet die Übersicht und klappt die Werkzeuge auf – Unterricht, Schulalltag, Betrieb, Eltern und Kollegium, je nachdem, was an Ihrer Schule freigeschaltet ist.</p>',
            placement: 'right',
            padding: 6,
            before: async () => {
                ensurePortalHome();
                await waitForSelector('.dash-portal-shell__nav', 2000);
            }
        },
        {
            id: 'apps',
            selector: '.dash-portal-section[aria-label="Meine Apps"], .dash-portal-section, .dash-quick-apps__grid',
            title: 'Meine Apps',
            bodyHtml:
                '<p>Hier liegen Ihre wichtigsten Tools als Kacheln. Tippen Sie eine App an und schauen Sie sich in Ruhe um – Anträge, Kalender, Genehmigungen, Formulare.</p>',
            placement: 'top',
            before: async () => {
                ensurePortalHome();
                await waitForSelector('.dash-portal-section, .dash-quick-apps__grid', 2000);
            }
        },
        {
            id: 'section',
            selector: '.dash-portal-section, .dash-portal-shell__main',
            title: 'Mehr Apps in den Bereichen',
            bodyHtml:
                '<p>Ein Bereich (z. B. Unterricht) zeigt alle zugehörigen Apps. So finden Sie Freistellungen, Schularbeiten oder sRDP, ohne den ganzen Katalog zu durchsuchen.</p>',
            placement: 'top',
            before: async () => {
                const firstSection = document.querySelector(
                    '.dash-portal-shell__nav [data-portal-nav]:not([data-portal-nav="home"]):not([data-portal-nav="it-bereich"])'
                );
                const id = firstSection ? firstSection.getAttribute('data-portal-nav') : '';
                if (id) clickPortalNav(id);
                else ensurePortalHome();
                await waitForSelector('.dash-portal-section, .dash-portal-shell__main', 2000);
            }
        },
        {
            id: 'links',
            selector: '.dash-portal-links, .dash-portal-links__title',
            title: 'Schnellzugriff',
            bodyHtml:
                '<p>Outlook, Teams, Kalender und schulinterne Links (WebUntis, Intranet …) – einmal setzen, dann immer griffbereit. Lehrkräfte können die Links über <em>Links bearbeiten</em> anpassen.</p>',
            placement: 'bottom',
            before: async () => {
                ensurePortalHome();
                await waitForSelector('.dash-portal-links', 2000);
            }
        },
        {
            id: 'feedback',
            selector: '.site-feedback-link, #siteCreditKurtrocks',
            title: 'Feedback immer griffbereit',
            primaryLabel: 'Feedback öffnen',
            bodyHtml:
                FEEDBACK_HINT +
                '<p>Als Nächstes: eine App öffnen und den Alltag ausprobieren. Die Tour können Sie im Kontomenü jederzeit erneut starten.</p>',
            placement: 'top',
            before: async () => {
                ensurePortalHome();
                await waitForSelector('.site-feedback-link, #siteCreditKurtrocks', 3000);
            },
            onPrimary: async () => {
                openFeedbackPrefill('Lehrer-Portal / First-Run');
                return 'finish';
            }
        }
    ];
}

function toast(msg) {
    try {
        if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
    } catch {
        /* ignore */
    }
}

function isDashboardPage() {
    try {
        const path = String(location.pathname || '');
        if (/tenant\.html$/i.test(path) || /\/tools\//i.test(path) || /admin\.html$/i.test(path)) {
            return false;
        }
        if (/index\.html$/i.test(path)) return true;
        const last = path.split('/').filter(Boolean).pop() || '';
        return !/\.html$/i.test(last);
    } catch {
        return false;
    }
}

function dashboardHref(withTour) {
    try {
        const p = String(location.pathname || '');
        const base = /\/tools\//i.test(p) ? '../index.html' : 'index.html';
        return withTour ? base + '?tour=lehrer' : base;
    } catch {
        return withTour ? 'index.html?tour=lehrer' : 'index.html';
    }
}

function closeOpenTours() {
    try {
        document.getElementById('ms365OnboardingOverlay')?.remove();
        document.getElementById('ms365ProductTourRoot')?.remove();
        document.documentElement.classList.remove('ms365-product-tour-active');
    } catch {
        /* ignore */
    }
}

function isChipDismissed() {
    try {
        return sessionStorage.getItem(LEHRER_FIRST_RUN_CHIP_DISMISS_KEY) === '1';
    } catch {
        return false;
    }
}

function setChipDismissed(on) {
    try {
        if (on) sessionStorage.setItem(LEHRER_FIRST_RUN_CHIP_DISMISS_KEY, '1');
        else sessionStorage.removeItem(LEHRER_FIRST_RUN_CHIP_DISMISS_KEY);
    } catch {
        /* ignore */
    }
}

function finishFirstRun(reason) {
    markLehrerWelcomeTourDone(localStorage, reason === 'skipped' ? 'skipped' : 'seen');
    patchLehrerFirstRunState(localStorage, {
        phase: reason === 'skipped' ? 'skipped' : 'done',
        step: 0
    });
    if (reason !== 'skipped') setChipDismissed(false);
    refreshLehrerFirstRunChrome();
}

/**
 * First-Run-Status zurücksetzen (zum erneuten Testen).
 */
export function resetLehrerFirstRun() {
    try {
        localStorage.removeItem(LEHRER_WELCOME_KEY);
        localStorage.removeItem(LEHRER_FIRST_RUN_KEY);
    } catch {
        /* ignore */
    }
    setChipDismissed(false);
    refreshLehrerFirstRunChrome();
}

/**
 * Komplette Tour von vorn (Dashboard). Auf anderen Seiten → Redirect mit ?tour=lehrer.
 * @param {{ silent?: boolean }} [opts]
 */
export async function restartLehrerFirstRun(opts) {
    closeOpenTours();
    resetLehrerFirstRun();
    if (!isDashboardPage()) {
        window.location.href = dashboardHref(true);
        return true;
    }
    if (!opts || !opts.silent) toast('Lehrer-Tour startet von vorn …');
    return startLehrerDashboardFirstRun({ force: true });
}

/**
 * Offene Tour fortsetzen (Chip / Resume).
 */
export async function resumeLehrerFirstRun() {
    closeOpenTours();
    return restartLehrerFirstRun();
}

function ensureContinueChip() {
    let chip = document.getElementById('ms365LehrerFirstRunChip');
    if (chip) return chip;
    chip = document.createElement('div');
    chip.id = 'ms365LehrerFirstRunChip';
    chip.className = 'ms365-it-first-run-chip ms365-lehrer-first-run-chip';
    chip.hidden = true;
    chip.innerHTML =
        '<button type="button" class="ms365-it-first-run-chip__go" id="ms365LehrerFirstRunChipGo">' +
        '<span class="ms365-it-first-run-chip__kicker">Lehrer-Portal</span>' +
        '<span class="ms365-it-first-run-chip__label" id="ms365LehrerFirstRunChipLabel">Lehrer-Tour fortsetzen</span>' +
        '</button>' +
        '<button type="button" class="ms365-it-first-run-chip__dismiss" id="ms365LehrerFirstRunChipDismiss" title="Für diese Sitzung ausblenden" aria-label="Ausblenden">×</button>';
    document.body.appendChild(chip);

    chip.querySelector('#ms365LehrerFirstRunChipGo')?.addEventListener('click', () => {
        resumeLehrerFirstRun();
    });
    chip.querySelector('#ms365LehrerFirstRunChipDismiss')?.addEventListener('click', () => {
        setChipDismissed(true);
        chip.hidden = true;
    });
    return chip;
}

function closeAuthMenu() {
    try {
        document.getElementById('ms365AuthMenu')?.classList.remove('is-open');
        const drop = document.getElementById('ms365AuthDropdown');
        if (drop) drop.hidden = true;
        const badge = document.getElementById('ms365AuthBadge');
        if (badge) badge.setAttribute('aria-expanded', 'false');
    } catch {
        /* ignore */
    }
}

/**
 * Chip + Menüpunkt an Persona anpassen.
 * @param {{ isLehrer?: boolean, isIt?: boolean, loggedIn?: boolean }} [persona]
 */
export async function refreshLehrerFirstRunChrome(persona) {
    let isLehrer = persona && typeof persona.isLehrer === 'boolean' ? persona.isLehrer : null;
    let isIt = persona && typeof persona.isIt === 'boolean' ? persona.isIt : null;
    let loggedIn = persona && typeof persona.loggedIn === 'boolean' ? persona.loggedIn : null;
    if (isLehrer == null || loggedIn == null || isIt == null) {
        try {
            const p = await resolveLehrerPersona(2500);
            isLehrer = !!p.isLehrer;
            isIt = !!p.isIt;
            loggedIn = !!p.loggedIn || isLoggedInAccount();
        } catch {
            isLehrer = false;
            isIt = false;
            loggedIn = isLoggedInAccount();
        }
    }

    const state = readLehrerFirstRunState(localStorage);
    const showChip = shouldShowLehrerFirstRunChip({
        isLehrer,
        isIt,
        loggedIn,
        phase: state.phase,
        chipDismissed: isChipDismissed()
    });

    const chip = ensureContinueChip();
    const labelEl = chip.querySelector('#ms365LehrerFirstRunChipLabel');
    if (labelEl) labelEl.textContent = lehrerChipLabelFromPhase(state.phase);
    const tourOpen = !!(
        document.getElementById('ms365OnboardingOverlay') ||
        document.getElementById('ms365ProductTourRoot') ||
        document.documentElement.classList.contains('ms365-product-tour-active')
    );
    chip.hidden = !showChip || tourOpen;

    const portalPreview = isPortalPreviewLayout(readLayoutMode());
    const tourBtn = document.getElementById('ms365AuthLehrerTourBtn');
    if (tourBtn) {
        tourBtn.hidden = !shouldShowLehrerTourMenuButton({
            loggedIn,
            isLehrer,
            isIt,
            portalPreview
        });
        if (!tourBtn.dataset.lehrerTourBound) {
            tourBtn.dataset.lehrerTourBound = '1';
            tourBtn.addEventListener('click', () => {
                closeAuthMenu();
                restartLehrerFirstRun();
            });
        }
    }
}

/**
 * @param {{ force?: boolean }} [opts]
 */
export async function startLehrerDashboardFirstRun(opts) {
    const force = !!(opts && opts.force);
    const persona = await resolveLehrerPersona(force ? 2000 : 8000);
    const isLehrer = !!persona.isLehrer;
    const isIt = !!persona.isIt;
    const loggedIn = !!persona.loggedIn || isLoggedInAccount();
    const state = readLehrerFirstRunState(localStorage);
    const welcomeDone = isLehrerWelcomeTourDone(localStorage);
    const portalReady = isLehrerPortalReady(
        evaluateLehrerPortalReadiness(gatherLehrerPortalReadinessInput())
    );

    if (
        !shouldStartLehrerWelcomeTour({
            force,
            isLehrer,
            isIt,
            loggedIn,
            welcomeDone,
            phase: state.phase,
            portalReady
        })
    ) {
        return false;
    }

    // Tour braucht die Portal-Shell (Apps-only). IT im Backend → kurz auf Portal wechseln.
    try {
        if (!isPortalPreviewLayout(readLayoutMode())) {
            const api = window.ms365DashboardAudience;
            if (api && typeof api.setLayoutMode === 'function') {
                api.setLayoutMode('preview-lehrer');
            }
        }
    } catch {
        /* ignore */
    }

    // Portal muss gerendert sein (Apps-only Shell)
    await waitForSelector('#dashQuickApps, .dash-portal-shell', force ? 2500 : 5000);

    patchLehrerFirstRunState(localStorage, { phase: 'welcome', step: 0 });

    openWelcomeTour(welcomePanels(), {
        title: 'Willkommen im Lehrer-Portal',
        lead: 'Kurze Orientierung: Dashboard, Apps und Feedback-Kanal für die Pilotwochen.',
        onComplete(reason) {
            if (reason === 'skipped') {
                finishFirstRun('skipped');
                return;
            }
            patchLehrerFirstRunState(localStorage, { phase: 'dashboard', step: 0 });
            window.setTimeout(() => {
                startSpotlightTour(portalSpotlightSteps(), {
                    onComplete(spotReason) {
                        finishFirstRun(spotReason === 'skipped' ? 'skipped' : 'done');
                    }
                });
            }, 180);
        }
    });
    return true;
}

function urlForceTour() {
    try {
        const q = new URLSearchParams(window.location.search || '');
        const v = String(q.get('tour') || '')
            .trim()
            .toLowerCase();
        return v === 'lehrer' || v === 'teacher' || v === 'portal';
    } catch {
        return false;
    }
}

function wireChromeListeners() {
    if (wireChromeListeners.bound) return;
    wireChromeListeners.bound = true;
    const refresh = () => {
        refreshLehrerFirstRunChrome();
    };
    window.addEventListener('ms365-dashboard-persona-ready', refresh);
    window.addEventListener('ms365-auth-state-changed', refresh);
    window.addEventListener('ms365-it-shell-changed', refresh);
    window.addEventListener('ms365-dash-app-only-changed', refresh);
    window.setTimeout(refresh, 800);
    window.setTimeout(refresh, 2000);
}

/**
 * @param {{ force?: boolean }} [opts]
 */
export function mountLehrerFirstRunTour(opts) {
    const force = !!(opts && opts.force) || urlForceTour();
    wireChromeListeners();

    const run = () => {
        refreshLehrerFirstRunChrome();
        if (force) {
            try {
                localStorage.removeItem(LEHRER_WELCOME_KEY);
                localStorage.removeItem(LEHRER_FIRST_RUN_KEY);
                setChipDismissed(false);
            } catch {
                /* ignore */
            }
        }
        startLehrerDashboardFirstRun({ force });
        window.setTimeout(() => refreshLehrerFirstRunChrome(), 400);
    };
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => window.setTimeout(run, 700));
    } else {
        window.setTimeout(run, 700);
    }
}

window.ms365LehrerFirstRun = {
    start: (o) => startLehrerDashboardFirstRun(o || { force: true }),
    restart: (o) => restartLehrerFirstRun(o),
    resume: () => resumeLehrerFirstRun(),
    reset: () => resetLehrerFirstRun(),
    refreshChrome: () => refreshLehrerFirstRunChrome(),
    mount: mountLehrerFirstRunTour,
    openFeedback: openFeedbackPrefill
};
