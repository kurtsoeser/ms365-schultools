/**
 * First-Run-Tour für Lehrkräfte und den Schulalltag: Willkommen → Kacheln, Links, Apps, Bereiche.
 */
import { readItShellChoice } from './dashboard-layout-mode.js';
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
    '<p><strong>Feedback ist erwünscht.</strong> In den Pilotwochen hilft uns jeder Hinweis – über das Fragezeichen unten rechts, Eintrag <em>Feedback</em>.</p>';

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

function openFooterHelp() {
    try {
        if (window.ms365FooterHelp && typeof window.ms365FooterHelp.open === 'function') {
            window.ms365FooterHelp.open();
        }
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
            title: 'Sie sind im Schulportal',
            bodyHtml:
                '<p>Hier liegt der Schulalltag: Termine, Anträge, Schularbeiten, Aktivitäten und die übrigen Werkzeuge, die die Schul-IT für Sie freigeschaltet hat.</p>' +
                '<ul><li>Einrichtung und Stammdaten macht die Schul-IT</li>' +
                '<li>Die Startseite zeigt Begrüßung, Kacheln <strong>Für dich</strong>, Schnellzugriff und Meine Apps</li>' +
                '<li>Links gliedern die Bereiche: Unterricht, Anträge, Schulalltag, Betrieb, Eltern, Kollegium, Intranet</li>' +
                '<li>Schüler:innen nutzen dieselbe Startseite mit ihren eigenen Apps und Kacheln</li></ul>'
        },
        {
            id: 'apps',
            title: 'Kacheln, Apps, Bereiche',
            bodyHtml:
                '<p><strong>Für dich</strong> sind Ihre Kacheln: was heute ansteht und was offen ist. Mit <em>Anordnen</em> ziehen Sie sie, ändern die Größe, nehmen welche dazu oder weg. <em>Fertig</em> merkt sich das für Ihr Konto auf diesem Gerät.</p>' +
                '<p><strong>Meine Apps</strong> ist die kurze Liste darunter. Die übrigen Werkzeuge sitzen in den Bereichen links – ein Klick auf den Bereich klappt sie auf.</p>' +
                '<p class="muted">Fehlt eine App oder eine Kachel? Oft hat die Schul-IT das Werkzeug noch nicht freigegeben – kurz nachfragen oder Feedback senden.</p>'
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
            selector: '.dash-portal-hero:not(.dash-portal-hero--section), #dashQuickApps',
            title: 'Ihre Startseite',
            bodyHtml:
                '<p>Begrüßung und Schulkarte. Darunter kommt der Alltag: Kacheln, Links und die kurzen App-Einstiege. Die Einrichtung bleibt bei der Schul-IT.</p>',
            placement: 'bottom',
            before: async () => {
                ensurePortalHome();
                await waitForSelector('.dash-portal-hero, #dashQuickApps', 2500);
            }
        },
        {
            id: 'tiles',
            selector: '#dashPortalBriefing:not([hidden]), .dash-portal-section[aria-label="Meine Apps"]',
            title: 'Kacheln „Für dich“',
            bodyHtml:
                '<p>Hier sehen Sie, was ansteht: Termine von heute und der nächsten Tage, offene Aktivitäten, Freistellungen, Schularbeiten, Aufgaben, Elternbriefe, eigene Technikmeldungen. Ein Klick öffnet das Werkzeug.</p>' +
                '<p><strong>Anordnen</strong> schaltet das Raster frei: ziehen, an der Ecke Breite und Höhe ändern, über <em>Kachel</em> eine weitere dazu (auch dieselbe App als Liste oder als Zahl), mit × entfernen. <em>Fertig</em> speichert die Anordnung für Ihr Konto auf diesem Gerät, <em>Zurücksetzen</em> holt die Standardauswahl zurück.</p>' +
                '<p class="muted">Ist der Bereich noch leer, hat die Schul-IT die Werkzeuge noch nicht freigegeben oder es liegen noch keine Einträge vor.</p>',
            placement: 'top',
            before: async () => {
                ensurePortalHome();
                await waitForSelector(
                    '#dashPortalBriefing:not([hidden]), .dash-portal-section[aria-label="Meine Apps"]',
                    4000
                );
            }
        },
        {
            id: 'links',
            selector: '.dash-portal-links, .dash-portal-links__title',
            title: 'Schnellzugriff',
            bodyHtml:
                '<p>Outlook, Teams, Kalender und schulinterne Links (WebUntis, Sokrates, Intranet). Über <em>Links bearbeiten</em> passen Sie die Leiste an. Die Kacheln darüber bleiben Ihr persönlicher Überblick, die Links der Sprung nach draußen.</p>',
            placement: 'bottom',
            before: async () => {
                ensurePortalHome();
                await waitForSelector('.dash-portal-links', 2000);
            }
        },
        {
            id: 'nav',
            selector: '.dash-portal-shell__nav, [data-portal-nav]',
            title: 'Bereiche links',
            bodyHtml:
                '<p><strong>Dashboard</strong> ist diese Startseite. Die übrigen Einträge sind Bereiche: Unterricht, Anträge und Formulare, Schulalltag, Betrieb &amp; Technik, Eltern &amp; Ablauf, Kollegium, Intranet.</p>' +
                '<p>Ein Klick öffnet den Bereich und klappt die Werkzeuge darunter auf. Ein zweiter Klick klappt sie wieder zu. Sichtbar ist nur, was die Schule freigegeben hat. Das Intranet öffnet die Schulwebsite.</p>',
            placement: 'right',
            padding: 6,
            before: async () => {
                ensurePortalHome();
                await waitForSelector('.dash-portal-shell__nav', 2000);
            }
        },
        {
            id: 'apps',
            selector: '.dash-portal-section[aria-label="Meine Apps"]',
            title: 'Meine Apps',
            bodyHtml:
                '<p>Die kurze Liste auf der Startseite: was die Schule angepinnt hat, Ihre Favoriten, sonst die wichtigsten Einstiege. Ein Klick öffnet das Werkzeug.</p>' +
                '<p>Alles Weitere liegt in den Bereichen links – die Startseite zeigt absichtlich nur einen Ausschnitt.</p>',
            placement: 'top',
            before: async () => {
                ensurePortalHome();
                await waitForSelector('.dash-portal-section[aria-label="Meine Apps"]', 2000);
            }
        },
        {
            id: 'section',
            selector: '.dash-portal-section[data-portal-section], .dash-portal-shell__main',
            title: 'Ein Bereich, die passenden Werkzeuge',
            bodyHtml:
                '<p>Unterricht bündelt Schularbeiten und Unterrichtsteams. Anträge und Formulare: Freistellungen, sRDP, Wahlfach. Schulalltag: Termine, Aktivitäten, Projektwochen, Aufgaben. Eltern &amp; Ablauf: Briefe, Elternsprechtag, Einsammler, Krankmeldung. Betrieb: Technik, Räume, Ausgaben.</p>' +
                '<p>Jede Zeile öffnet genau dieses Werkzeug.</p>',
            placement: 'top',
            before: async () => {
                const preferred =
                    document.querySelector('[data-portal-nav="unterricht"]') ||
                    document.querySelector(
                        '.dash-portal-shell__nav [data-portal-nav]:not([data-portal-nav="home"]):not([data-portal-nav="it-bereich"])'
                    );
                const id = preferred ? preferred.getAttribute('data-portal-nav') : '';
                if (id) clickPortalNav(id);
                else ensurePortalHome();
                await waitForSelector('.dash-portal-section[data-portal-section], .dash-portal-shell__main', 2000);
            }
        },
        {
            id: 'feedback',
            selector: '.site-feedback-link, #siteCreditKurtrocks',
            title: 'Feedback immer griffbereit',
            primaryLabel: 'Feedback öffnen',
            bodyHtml:
                FEEDBACK_HINT +
                '<p>Als Nächstes: eine Kachel oder eine App öffnen und den Alltag ausprobieren. Die Tour starten Sie im Kontomenü jederzeit neu.</p>',
            placement: 'top',
            before: async () => {
                ensurePortalHome();
                await waitForSelector('.site-feedback-link, #appFooterHelpBtn', 3000);
                openFooterHelp();
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

    if (!force && readItShellChoice() === 'full') return false;

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

    // Portal-Tour darf die gespeicherte IT-Arbeitsfläche nicht auf „Portal“ umschreiben.
    // setLayoutMode('preview-lehrer') hat das bisher dauerhaft gespeichert – danach landete
    // jeder Wechsel auf dem Lehrer-Dashboard.

    patchLehrerFirstRunState(localStorage, { phase: 'welcome', step: 0 });

    openWelcomeTour(welcomePanels(), {
        title: 'Willkommen im Schulportal',
        lead: 'Startseite, Kacheln „Für dich“, Bereiche links und der Feedback-Kanal für die Pilotwochen.',
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
