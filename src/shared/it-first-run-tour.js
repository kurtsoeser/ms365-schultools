/**
 * First-Run-Tour für Schul-IT: Willkommen → Portal, IT-Leiste, Stammdaten, Tool-Site.
 */
import { getDashboardPersonas } from './dashboard-persona-session.js';
import { readItShellChoice } from './dashboard-layout-mode.js';
import {
    chipLabelFromPhase,
    isFirstSaveTipDone,
    isWelcomeTourDone,
    IT_FIRST_RUN_CHIP_DISMISS_KEY,
    markFirstSaveTipDone,
    markWelcomeTourDone,
    patchItFirstRunState,
    readItFirstRunState,
    resumeTargetFromPhase,
    shouldShowFirstSaveTip,
    shouldShowItFirstRunChip,
    shouldStartItStammdatenTour,
    shouldStartItWelcomeTour
} from './product-tour-logic.js';
import { openWelcomeTour, startSpotlightTour } from './product-tour.js';
import { activateStammdatenSetupPane } from './tenant-stammdaten-setup-nav.js';

const FEEDBACK_HINT =
    '<p><strong>Feedback ist erwünscht.</strong> In den Pilotwochen hilft uns jeder Hinweis – über das Fragezeichen unten rechts, Eintrag <em>Feedback</em>.</p>';

function openFeedbackPrefill(title) {
    try {
        if (window.ms365Feedback && typeof window.ms365Feedback.open === 'function') {
            window.ms365Feedback.open({ title: title || 'IT First-Run / Pilot' });
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

function activateStammdatenPane(pane) {
    try {
        activateStammdatenSetupPane(pane, { persist: false });
    } catch {
        /* ignore */
    }
}

function activateTenantMainTab(tabId) {
    try {
        if (typeof window.ms365TenantTabActivate === 'function') {
            window.ms365TenantTabActivate(tabId);
            return;
        }
    } catch {
        /* ignore */
    }
    try {
        window.dispatchEvent(
            new CustomEvent('ms365-register-tab-request', { detail: { tabBtnId: tabId } })
        );
        return;
    } catch {
        /* ignore */
    }
    try {
        const btn = document.getElementById(tabId);
        if (btn && typeof btn.click === 'function') btn.click();
    } catch {
        /* ignore */
    }
}

function isItPortalLayout() {
    try {
        const api = window.ms365DashboardAudience;
        if (api && typeof api.getLayoutMode === 'function') {
            const mode = String(api.getLayoutMode() || '');
            return (
                mode === 'preview-lehrer' ||
                mode === 'preview-schueler' ||
                mode === 'apps-lehrer' ||
                mode === 'apps-schueler'
            );
        }
    } catch {
        /* ignore */
    }
    try {
        return document.documentElement.hasAttribute('data-dash-app-only');
    } catch {
        return false;
    }
}

/** Gemerkte Ansicht, falls die Tour nur für die Backend-Schritte auf IT schaltet. */
let itTourLayoutRestore = '';

/** IT-Schritte brauchen das Backend. Die gespeicherte Portal-Wahl bleibt danach erhalten. */
function ensureItBackend() {
    try {
        const api = window.ms365DashboardAudience;
        if (!api || typeof api.setLayoutMode !== 'function') return;
        const current = typeof api.getLayoutMode === 'function' ? String(api.getLayoutMode() || '') : '';
        // Gespeicherte IT-Wahl nie auf Portal zurückschreiben (Tour-Restore).
        if (readItShellChoice() === 'full') {
            itTourLayoutRestore = '';
            if (current !== 'full') api.setLayoutMode('full');
            return;
        }
        if (current === 'full') return;
        if (!itTourLayoutRestore) itTourLayoutRestore = current || 'preview-lehrer';
        api.setLayoutMode('full');
    } catch {
        /* ignore */
    }
}

function restoreItTourLayout() {
    const mode = itTourLayoutRestore;
    if (!mode) return;
    itTourLayoutRestore = '';
    try {
        const api = window.ms365DashboardAudience;
        if (api && typeof api.setLayoutMode === 'function') api.setLayoutMode(mode);
    } catch {
        /* ignore */
    }
}

/**
 * @param {string} taskId
 */
async function selectItArea(taskId) {
    ensureItBackend();
    await waitForSelector('#dashTasksSplit, #' + taskId, 2500);
    try {
        const api = window.ms365DashTasksSplit;
        if (api && typeof api.select === 'function') api.select(taskId, { pulse: false });
    } catch {
        /* ignore */
    }
    await waitForSelector('#' + taskId, 1500);
}

/** @param {'tasks'|'catalog'|'playbooks'} section */
function setDashboardMainSection(section) {
    try {
        const api = window.ms365DashMainSection;
        if (api && typeof api.set === 'function') {
            api.set(section);
            return;
        }
    } catch {
        /* ignore */
    }
    try {
        document.documentElement.setAttribute('data-dash-main-section', section);
        const btn = document.querySelector('[data-dash-main-section="' + section + '"]');
        if (btn && typeof btn.click === 'function') btn.click();
    } catch {
        /* ignore */
    }
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
 * @returns {Promise<{ loggedIn: boolean, isIt: boolean }>}
 */
async function resolveItPersona(timeoutMs) {
    const budget = Math.max(500, Number(timeoutMs) || 8000);
    const started = Date.now();

    const fromCache = () => {
        const p = readCachedPersonas();
        if (p && p.loggedIn) return { loggedIn: true, isIt: !!p.isIt };
        return null;
    };

    const hit = fromCache();
    if (hit && hit.isIt) return hit;

    try {
        const personas = await getDashboardPersonas({ force: false });
        if (personas && personas.loggedIn) {
            return { loggedIn: true, isIt: !!personas.isIt };
        }
    } catch {
        /* ignore */
    }

    while (Date.now() - started < budget) {
        const again = fromCache();
        if (again && again.isIt) return again;
        await new Promise((r) => window.setTimeout(r, 300));
        try {
            const personas = await getDashboardPersonas({ force: false });
            if (personas && personas.loggedIn) {
                return { loggedIn: true, isIt: !!personas.isIt };
            }
        } catch {
            /* ignore */
        }
    }

    const late = fromCache();
    if (late) return late;
    return { loggedIn: isLoggedInAccount(), isIt: false };
}

function welcomePanels() {
    return [
        {
            id: 'role',
            title: 'Zwei Ansichten: Portal und IT',
            bodyHtml:
                '<p>Als Schul-IT starten Sie im <strong>Portal</strong> – derselben Startseite, die Lehrkräfte und Schüler:innen sehen. Oben rechts schalten Sie auf <strong>IT</strong>, das Einrichtungs-Backend.</p>' +
                '<ul><li><strong>Portal:</strong> Begrüßung, Kacheln „Für dich“, Schnellzugriff und Meine Apps</li>' +
                '<li><strong>IT:</strong> linke Leiste mit Stammdaten, MS365 Mitgliedschaft, Klassen, Schul-Apps und den übrigen Bereichen</li>' +
                '<li>Die Schüler-Vorschau liegt im Kontomenü</li></ul>'
        },
        {
            id: 'stammdaten',
            title: 'Sinnvoller Start im IT-Bereich',
            bodyHtml:
                '<p>Empfohlene Reihenfolge für neue Schulen:</p>' +
                '<ul><li><strong>Stammdaten → Schulprofil</strong> – Name, Domain, Logo, danach E-Mail-Muster</li>' +
                '<li><strong>Lehrer, Klassen, Fächer</strong> – in derselben Leiste, Import (z. B. WebUntis) oder manuell</li>' +
                '<li><strong>Schulprofil → Planer &amp; Apps</strong> – SharePoint-Website für die Tool-Listen</li>' +
                '<li><strong>Blaupause Erste Einrichtung</strong> – unter Stammdaten bei „Mehr …“</li>' +
                '<li><strong>Schul-Apps → Portal vorbereiten</strong> – damit im Portal Apps und Kacheln erscheinen</li></ul>' +
                '<p class="muted">Die Tour orientiert. Abhaken tun Sie in der Blaupause.</p>'
        },
        {
            id: 'feedback',
            title: 'Pilotwochen: viel Feedback bitte',
            primaryLabel: 'Tour auf dem Dashboard starten',
            bodyHtml:
                FEEDBACK_HINT +
                '<p>Hakt etwas, ist unklar oder fehlt ein Schritt? Kurz melden – wir priorisieren danach die nächsten Verbesserungen.</p>'
        }
    ];
}

function dashboardSpotlightSteps() {
    const steps = [];

    if (isItPortalLayout()) {
        steps.push(
            {
                id: 'portal-home',
                selector: '.dash-portal-hero, #dashQuickApps',
                title: 'Das Portal, das alle sehen',
                bodyHtml:
                    '<p>Das ist die Startseite für Lehrkräfte und Schüler:innen – und Ihr Standard, bis Sie auf <strong>IT</strong> schalten.</p>' +
                    '<p>Unter der Begrüßung liegen die Kacheln <strong>Für dich</strong> (Termine, offene Anträge, Aufgaben), danach Schnellzugriff und <strong>Meine Apps</strong>. Die Kacheln kann jede Person selbst anordnen.</p>',
                placement: 'bottom',
                before: async () => {
                    restoreItTourLayout();
                    await waitForSelector('.dash-portal-hero, #dashPortalBriefing, #dashQuickApps', 2500);
                }
            },
            {
                id: 'portal-tiles',
                selector: '#dashPortalBriefing:not([hidden]), .dash-portal-section[aria-label="Meine Apps"]',
                title: 'Kacheln „Für dich“',
                bodyHtml:
                    '<p>Sobald ein Werkzeug freigeschaltet ist und Daten hat, erscheint es hier als Kachel: heute, nächste Termine, Aktivitäten, Freistellungen, Schularbeiten, Aufgaben, Elternbriefe, Technik.</p>' +
                    '<p><strong>Anordnen</strong> verschiebt, vergrößert, ergänzt oder entfernt Kacheln. <em>Fertig</em> speichert das für dieses Konto auf diesem Gerät. Welche Kacheln es überhaupt gibt, entscheiden Sie unter <strong>Schul-Apps → Portal vorbereiten</strong>.</p>',
                placement: 'top',
                before: async () => {
                    restoreItTourLayout();
                    await waitForSelector(
                        '#dashPortalBriefing:not([hidden]), .dash-portal-section[aria-label="Meine Apps"]',
                        4000
                    );
                }
            }
        );
    }

    steps.push(
        {
            id: 'switch',
            selector: '#ms365HeaderShellSwitch, #ms365AuthWidget',
            title: 'Umschalter Portal / IT',
            bodyHtml:
                '<p>Oben rechts: <strong>Portal</strong> ist die Alltags-Startseite, <strong>IT</strong> die Einrichtung. Die Schüler-Vorschau bleibt im Kontomenü.</p>' +
                '<p>Als Nächstes wechseln wir ins IT-Backend – dort liegen Stammdaten, Bereiche und Portal vorbereiten.</p>',
            placement: 'bottom',
            before: async () => {
                await waitForSelector('#ms365HeaderShellSwitch, #ms365AuthWidget', 2500);
            }
        },
        {
            id: 'rail',
            selector: '#itWorkspaceNav, #dashTasksSplit .dash-tasks-split__nav',
            title: 'IT-Bereiche in der Leiste',
            bodyHtml:
                '<p>Links liegen die Bereiche, jeder klappt seine Werkzeuge auf: <strong>Stammdaten</strong>, MS365 Mitgliedschaft, Klassen &amp; Unterricht, Personen, Schuljahresstart, Intranet, <strong>Schul-Apps</strong>, Aufräumen.</p>' +
                '<p>Playbooks und Katalog sind keine eigenen Reiter mehr – die Checklisten hängen im passenden Bereich. Die Suche darüber findet Werkzeuge über alle Bereiche.</p>',
            placement: 'right',
            padding: 6,
            before: async () => {
                setDashboardMainSection('tasks');
                ensureItBackend();
                await waitForSelector('#itWorkspaceNav, #dashTasksSplit .dash-tasks-split__nav', 2500);
            }
        },
        {
            id: 'import',
            selector: '#dashTaskImportVerknuepfen',
            title: 'Stammdaten und Blaupause',
            bodyHtml:
                '<p>Der erste Bereich heißt ohne Schuldaten <strong>Erste Einrichtung</strong>, mit Daten <strong>Stammdaten</strong>. In der Leiste: Schulprofil, Lehrer, Klassen, Fächer – Import und Blaupause unter <em>Mehr …</em>.</p>' +
                '<p>Die Blaupause ist die Checkliste: Profil, Tool-Site, Import, Verknüpfen, Sicherung. Die Tour hier orientiert nur.</p>',
            placement: 'left',
            padding: 6,
            before: async () => {
                setDashboardMainSection('tasks');
                await selectItArea('dashTaskImportVerknuepfen');
            }
        },
        {
            id: 'schulapps',
            selector: '#dashTaskSchulApps, [data-task-id="dashTaskSchulApps"]',
            title: 'Schul-Apps und Portal vorbereiten',
            bodyHtml:
                '<p><strong>Portal vorbereiten</strong> gibt frei, was Lehrkräfte und Schüler:innen auf der Startseite sehen: Apps in den Bereichen und die Kacheln „Für dich“.</p>' +
                '<p>Dieselben Schul-Apps (Schularbeiten, Freistellungen, Elternsprechtag …) richten Sie hier ein. Schuljahresstart, Intranet und Aufräumen sind eigene Bereiche in derselben Leiste.</p>',
            placement: 'left',
            before: async () => {
                setDashboardMainSection('tasks');
                await selectItArea('dashTaskSchulApps');
            }
        },
        {
            id: 'feedback',
            selector: '.site-feedback-link, #siteCreditKurtrocks',
            title: 'Feedback immer griffbereit',
            primaryLabel: 'Zur Blaupause (Playbook)',
            bodyHtml:
                FEEDBACK_HINT +
                '<p>Als Nächstes die Blaupause öffnen – oder im Bereich Stammdaten das Schulprofil ansehen und danach die Checkliste abhaken. Portal vorbereiten folgt, sobald Profil und Tool-Site stehen.</p>',
            placement: 'top',
            before: async () => {
                await waitForSelector('.site-feedback-link, #appFooterHelpBtn', 3000);
                openFooterHelp();
            },
            onPrimary: async () => {
                markWelcomeTourDone(localStorage, 'seen');
                patchItFirstRunState(localStorage, { phase: 'done', step: 0 });
                window.location.href = 'tools/playbook-daten-import-verknuepfen.html';
                return 'finish';
            }
        }
    );

    return steps;
}

function stammdatenSpotlightSteps() {
    return [
        {
            id: 'profile',
            selector: '#tenantSchoolProfileCard, [data-stammdaten-pane-panel="profile"]',
            title: '1 · Schulprofil',
            bodyHtml:
                '<p>Schulname, Domain und Logo bilden die Basis. Ohne Domain bleiben Import-Mails und Vorschläge oft leer.</p>',
            placement: 'right',
            before: async () => {
                activateTenantMainTab('tabMainStammdaten');
                activateStammdatenPane('profile');
                await waitForSelector('#tenantSchoolProfileCard', 2000);
            }
        },
        {
            id: 'email',
            selector: '#tenantSchoolEmailPatternCard, [data-stammdaten-pane-panel="email"]',
            title: '2 · E-Mail-Muster',
            bodyHtml:
                '<p>Muster für fehlende Adressen bei Importen – einmal setzen, spart später viel Nacharbeit.</p>',
            placement: 'right',
            before: async () => {
                activateStammdatenPane('email');
                await waitForSelector('#tenantSchoolEmailPatternCard', 1500);
            }
        },
        {
            id: 'teachers',
            selector: '#tabMainLehrer, #panelMainLehrer',
            title: '3 · Lehrkräfte eintragen',
            bodyHtml:
                '<p>In der Leiste unter <strong>Personen → Lehrer</strong>. Aus Microsoft&nbsp;365 übernehmen, WebUntis/CSV importieren oder manuell pflegen. Die Apps brauchen diese Liste für Rollen und Zielgruppen.</p>',
            placement: 'bottom',
            before: async () => {
                activateTenantMainTab('tabMainLehrer');
                await waitForSelector('#tabMainLehrer, #panelMainLehrer', 1500);
            }
        },
        {
            id: 'classes',
            selector: '#tabMainKlassen, #panelMainKlassen',
            title: '4 · Klassen',
            bodyHtml:
                '<p>Unter <strong>Unterricht → Klassen</strong>, mit Klassenvorstand. Daneben liegen Fächer, ARGE und Lehrer je Klasse – Grundlage für Gruppen, Planer und Kommunikation.</p>',
            placement: 'bottom',
            before: async () => {
                activateTenantMainTab('tabMainKlassen');
                await waitForSelector('#tabMainKlassen, #panelMainKlassen', 1500);
            }
        },
        {
            id: 'tool-site',
            selector: '#tenantPlanerSitesPanel, #tenantPlanerDefaultSiteUrl, [data-stammdaten-pane="spo-planer"]',
            title: '5 · Planer & Apps',
            bodyHtml:
                '<p>Unter <strong>Schulprofil → SharePoint &amp; Dienste → Planer &amp; Apps</strong> die Website im Schul-Tenant eintragen. Dort liegen die Listen der Schul-Apps.</p>' +
                '<p>Ohne diese Tool-Site bleiben Freistellungen, Schularbeiten, Aktivitäten und die Kacheln im Portal leer.</p>',
            placement: 'right',
            before: async () => {
                activateTenantMainTab('tabMainStammdaten');
                activateStammdatenPane('spo-planer');
                await waitForSelector('#tenantPlanerSitesPanel, #tenantPlanerDefaultSiteUrl', 2000);
            }
        },
        {
            id: 'feedback-end',
            selector: '.site-feedback-link, #siteCreditKurtrocks',
            title: 'Als Nächstes: Blaupause, dann Portal',
            primaryLabel: 'Blaupause öffnen',
            bodyHtml:
                '<p>Die Checkliste <strong>Erste Einrichtung</strong> (Stammdaten → Mehr …) führt Import, Verknüpfen und Sicherung. Danach <strong>Schul-Apps → Portal vorbereiten</strong>, damit die Startseite der Lehrkräfte Apps und Kacheln zeigt.</p>' +
                FEEDBACK_HINT,
            placement: 'top',
            before: async () => {
                activateTenantMainTab('tabMainStammdaten');
                activateStammdatenPane('profile');
                await waitForSelector('.site-feedback-link, #appFooterHelpBtn', 2500);
                openFooterHelp();
            },
            onPrimary: async () => {
                window.location.href = 'tools/playbook-daten-import-verknuepfen.html';
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

function isTenantPage() {
    try {
        return /(?:^|\/)tenant\.html$/i.test(String(location.pathname || ''));
    } catch {
        return false;
    }
}

function dashboardHref(withTour) {
    try {
        const p = String(location.pathname || '');
        const base = /\/tools\//i.test(p) ? '../index.html' : 'index.html';
        return withTour ? base + '?tour=it' : base;
    } catch {
        return withTour ? 'index.html?tour=it' : 'index.html';
    }
}

function tenantHref() {
    try {
        const p = String(location.pathname || '');
        const base = /\/tools\//i.test(p) ? '../tenant.html' : 'tenant.html';
        return base + '#stammdaten';
    } catch {
        return 'tenant.html#stammdaten';
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
        return sessionStorage.getItem(IT_FIRST_RUN_CHIP_DISMISS_KEY) === '1';
    } catch {
        return false;
    }
}

function setChipDismissed(on) {
    try {
        if (on) sessionStorage.setItem(IT_FIRST_RUN_CHIP_DISMISS_KEY, '1');
        else sessionStorage.removeItem(IT_FIRST_RUN_CHIP_DISMISS_KEY);
    } catch {
        /* ignore */
    }
}

function finishFirstRun(reason) {
    markWelcomeTourDone(localStorage, reason === 'skipped' ? 'skipped' : 'seen');
    patchItFirstRunState(localStorage, {
        phase: reason === 'skipped' ? 'skipped' : 'done',
        step: 0
    });
    if (reason !== 'skipped') setChipDismissed(false);
    refreshItFirstRunChrome();
}

/**
 * First-Run-Status zurücksetzen (zum erneuten Testen).
 */
export function resetItFirstRun() {
    try {
        localStorage.removeItem('ms365-onboarding-welcome-v1');
        localStorage.removeItem('ms365-it-first-run-v1');
        localStorage.removeItem('ms365-it-first-save-tip-v1');
    } catch {
        /* ignore */
    }
    setChipDismissed(false);
    refreshItFirstRunChrome();
}

function ensureBackupPanelOpen() {
    try {
        const wrap = document.getElementById('ms365BackupHeader');
        const panel = document.getElementById('ms365BackupPanel');
        const headerBtn = document.getElementById('ms365BackupHeaderBtn');
        const open = wrap && wrap.classList.contains('is-open') && panel && !panel.hidden;
        if (!open && headerBtn && typeof headerBtn.click === 'function') headerBtn.click();
    } catch {
        /* ignore */
    }
}

function itLibraryTipStep(primaryLabel) {
    return {
        id: 'it-library',
        selector:
            '#ms365BackupHeader .ms365-backup-panel__setup, .ms365-backup-panel__setup, a[data-ms365-spo-sync="setup"], #ms365BackupHeader',
        title: 'Bitte IT-Sicherungsbibliothek anlegen',
        primaryLabel: primaryLabel || 'Verstanden',
        bodyHtml:
            '<p>Die Daten liegen zuerst <strong>nur in diesem Browser</strong>. Legen Sie bitte die <strong>IT-Sicherungsbibliothek</strong> in SharePoint an – dann können Sie manuell sichern und laden (mit Versionierung).</p>' +
            '<p>Im Panel: <em>IT-Sicherungsbibliothek einrichten</em>. Die Tool-Listen der Apps sind eine andere Website: Schulprofil → Planer &amp; Apps.</p>' +
            FEEDBACK_HINT,
        placement: 'bottom',
        before: async () => {
            ensureBackupPanelOpen();
            if (isTenantPage()) activateStammdatenPane('spo-it');
            await waitForSelector(
                '#ms365BackupHeader .ms365-backup-panel__setup, .ms365-backup-panel__setup, #ms365BackupHeader',
                2000
            );
        }
    };
}

/** Tip nach „Speichern“ in den Stammdaten. */
function firstSaveTipSteps() {
    return [
        {
            id: 'local-save',
            selector: '#tenantSaveBar, #tenantSettingsSave, [data-tenant-action="save"]',
            title: 'Gespeichert – lokal in diesem Browser',
            bodyHtml:
                '<p><strong>Speichern</strong> schreibt die Stammdaten in <em>diesen</em> Browser (Lokal-First). Andere PCs oder ein geleerter Cache sehen die Daten so noch nicht.</p>' +
                '<p>Deshalb gleich die Sicherung mitdenken – nicht nur auf den einen Rechner verlassen.</p>',
            placement: 'top',
            padding: 6
        },
        {
            id: 'tool-site',
            selector: '#tenantPlanerSitesPanel, #tenantPlanerDefaultSiteUrl, [data-stammdaten-pane="spo-planer"]',
            title: 'Schul-SharePoint für die Tools',
            bodyHtml:
                '<p>Damit Lehrkräfte Schulaktivitäten, Freistellungen und die Kacheln auf der Startseite nutzen können, brauchen die Apps eine <strong>SharePoint-Website in Ihrem Schul-Tenant</strong> – dort liegen die Listen.</p>' +
                '<p>Unter <strong>Schulprofil → SharePoint &amp; Dienste → Planer &amp; Apps</strong> die Primär-Site eintragen und prüfen. Diese URL wird automatisch dem Betreiber gemeldet (Feld „Tool-Site-URL“).</p>',
            placement: 'right',
            before: async () => {
                activateTenantMainTab('tabMainStammdaten');
                activateStammdatenPane('spo-planer');
                await waitForSelector('#tenantPlanerSitesPanel, #tenantPlanerDefaultSiteUrl', 2000);
            }
        },
        {
            id: 'browser-backup',
            selector: '#ms365BackupHeader, #tenantSaveBar .ts-register-save-drop, .ts-register-save-drop--header',
            title: 'Browser-Backup (JSON)',
            bodyHtml:
                '<p>Über das <strong>Disketten-Symbol</strong> oben oder „Browser-Backup“ können Sie alle lokalen App-Daten als JSON exportieren und auf einem anderen Gerät wieder importieren.</p>',
            placement: 'bottom',
            before: async () => {
                ensureBackupPanelOpen();
                await waitForSelector('#ms365BackupHeader', 1500);
            }
        },
        itLibraryTipStep('Verstanden')
    ];
}

/** Tip beim ersten Klick auf das Floppy-/Datensicherungs-Symbol im Header. */
function backupHeaderTipSteps() {
    return [
        {
            id: 'backup-open',
            selector: '#ms365BackupHeader, #ms365BackupHeaderBtn',
            title: 'Datensicherung',
            bodyHtml:
                '<p>Hier sichern Sie Ihre Schuldaten: <strong>Browser-Backup</strong> (JSON) und – empfohlen – die <strong>IT-Sicherungsbibliothek</strong> in SharePoint.</p>' +
                '<p>Ohne SharePoint-Sicherung bleiben die Daten an diesen Browser gebunden.</p>',
            placement: 'bottom',
            before: async () => {
                ensureBackupPanelOpen();
                await waitForSelector('#ms365BackupHeader', 1500);
            }
        },
        itLibraryTipStep('Sicherungsbibliothek merken')
    ];
}

let firstSaveTipInFlight = false;

/**
 * Kurz-Tour nach dem ersten manuellen Speichern.
 * @param {{ force?: boolean, reason?: string }} [opts]
 */
export function startFirstSaveTipTour(opts) {
    const force = !!(opts && opts.force);
    const reason = opts && opts.reason != null ? opts.reason : 'manual-save';
    if (
        !shouldShowFirstSaveTip({
            force,
            tipDone: isFirstSaveTipDone(localStorage),
            reason,
            isTenantPage: isTenantPage()
        })
    ) {
        return false;
    }
    if (firstSaveTipInFlight && !force) return false;
    if (
        document.getElementById('ms365OnboardingOverlay') ||
        document.getElementById('ms365ProductTourRoot')
    ) {
        return false;
    }

    firstSaveTipInFlight = true;
    markFirstSaveTipDone(localStorage, 'seen');

    const steps = reason === 'backup-header' ? backupHeaderTipSteps() : firstSaveTipSteps();

    window.setTimeout(() => {
        startSpotlightTour(steps, {
            skipLabel: 'Später',
            onComplete(tipReason) {
                firstSaveTipInFlight = false;
                if (tipReason === 'skipped') markFirstSaveTipDone(localStorage, 'skipped');
                else markFirstSaveTipDone(localStorage, 'seen');
            }
        });
    }, 420);
    return true;
}

function wireFirstSaveTipListener() {
    if (wireFirstSaveTipListener.bound) return;
    wireFirstSaveTipListener.bound = true;
    window.addEventListener('ms365-tenant-settings-changed', (ev) => {
        try {
            const reason = ev && ev.detail ? String(ev.detail.reason || '') : '';
            if (reason !== 'manual-save') return;
            startFirstSaveTipTour({ reason: 'manual-save' });
        } catch {
            /* ignore */
        }
    });
    // Floppy oben = Datensicherung → gleicher Erst-Hinweis inkl. Sicherungsbibliothek
    window.addEventListener('ms365-backup-header-opened', () => {
        try {
            startFirstSaveTipTour({ reason: 'backup-header' });
        } catch {
            /* ignore */
        }
    });
}

/**
 * Komplette Tour von vorn (Dashboard). Auf anderen Seiten → Redirect mit ?tour=it.
 * @param {{ silent?: boolean }} [opts]
 */
export async function restartItFirstRun(opts) {
    closeOpenTours();
    resetItFirstRun();
    if (!isDashboardPage()) {
        window.location.href = dashboardHref(true);
        return true;
    }
    if (!opts || !opts.silent) toast('IT-Tour startet von vorn …');
    return startItDashboardFirstRun({ force: true });
}

/**
 * Offene Einrichtung fortsetzen (Chip / Resume).
 */
export async function resumeItFirstRun() {
    closeOpenTours();
    const state = readItFirstRunState(localStorage);
    const target = resumeTargetFromPhase(state.phase);

    if (target === 'stammdaten') {
        patchItFirstRunState(localStorage, { phase: 'stammdaten', step: 0 });
        if (!isTenantPage()) {
            window.location.href = tenantHref();
            return true;
        }
        setChipDismissed(false);
        toast('Stammdaten-Einrichtung fortsetzen …');
        return startItStammdatenFirstRun({ force: true });
    }

    return restartItFirstRun();
}

function ensureContinueChip() {
    let chip = document.getElementById('ms365ItFirstRunChip');
    if (chip) return chip;
    chip = document.createElement('div');
    chip.id = 'ms365ItFirstRunChip';
    chip.className = 'ms365-it-first-run-chip';
    chip.hidden = true;
    chip.innerHTML =
        '<button type="button" class="ms365-it-first-run-chip__go" id="ms365ItFirstRunChipGo">' +
        '<span class="ms365-it-first-run-chip__kicker">Schul-IT Einrichtung</span>' +
        '<span class="ms365-it-first-run-chip__label" id="ms365ItFirstRunChipLabel">IT-Tour fortsetzen</span>' +
        '</button>' +
        '<button type="button" class="ms365-it-first-run-chip__dismiss" id="ms365ItFirstRunChipDismiss" title="Für diese Sitzung ausblenden" aria-label="Ausblenden">×</button>';
    document.body.appendChild(chip);

    chip.querySelector('#ms365ItFirstRunChipGo')?.addEventListener('click', () => {
        resumeItFirstRun();
    });
    chip.querySelector('#ms365ItFirstRunChipDismiss')?.addEventListener('click', () => {
        setChipDismissed(true);
        chip.hidden = true;
    });
    return chip;
}

/**
 * Chip + Menüpunkt an Persona/Phase anpassen.
 * @param {{ isIt?: boolean, loggedIn?: boolean }} [persona]
 */
export async function refreshItFirstRunChrome(persona) {
    let isIt = persona && typeof persona.isIt === 'boolean' ? persona.isIt : null;
    let loggedIn = persona && typeof persona.loggedIn === 'boolean' ? persona.loggedIn : null;
    if (isIt == null || loggedIn == null) {
        try {
            const p = await resolveItPersona(2500);
            isIt = !!p.isIt;
            loggedIn = !!p.loggedIn || isLoggedInAccount();
        } catch {
            isIt = false;
            loggedIn = isLoggedInAccount();
        }
    }

    const state = readItFirstRunState(localStorage);
    const showChip = shouldShowItFirstRunChip({
        isIt,
        loggedIn,
        phase: state.phase,
        chipDismissed: isChipDismissed()
    });

    const chip = ensureContinueChip();
    const labelEl = chip.querySelector('#ms365ItFirstRunChipLabel');
    if (labelEl) labelEl.textContent = chipLabelFromPhase(state.phase);
    const tourOpen = !!(
        document.getElementById('ms365OnboardingOverlay') ||
        document.getElementById('ms365ProductTourRoot') ||
        document.documentElement.classList.contains('ms365-product-tour-active')
    );
    chip.hidden = !showChip || tourOpen;

    const tourBtn = document.getElementById('ms365AuthItTourBtn');
    if (tourBtn) {
        tourBtn.hidden = !(loggedIn && isIt);
        if (!tourBtn.dataset.itTourBound) {
            tourBtn.dataset.itTourBound = '1';
            tourBtn.addEventListener('click', () => {
                try {
                    document.getElementById('ms365AuthMenu')?.classList.remove('is-open');
                    const drop = document.getElementById('ms365AuthDropdown');
                    if (drop) drop.hidden = true;
                    const badge = document.getElementById('ms365AuthBadge');
                    if (badge) badge.setAttribute('aria-expanded', 'false');
                } catch {
                    /* ignore */
                }
                restartItFirstRun();
            });
        }
    }
}

/**
 * @param {{ force?: boolean }} [opts]
 */
export async function startItDashboardFirstRun(opts) {
    const force = !!(opts && opts.force);
    const persona = await resolveItPersona(force ? 2000 : 8000);
    const isIt = !!persona.isIt;
    const loggedIn = !!persona.loggedIn || isLoggedInAccount();
    const state = readItFirstRunState(localStorage);
    const welcomeDone = isWelcomeTourDone(localStorage);

    if (
        !shouldStartItWelcomeTour({
            force,
            isIt,
            loggedIn,
            welcomeDone,
            phase: state.phase
        })
    ) {
        return false;
    }

    patchItFirstRunState(localStorage, { phase: 'welcome', step: 0 });

    openWelcomeTour(welcomePanels(), {
        title: 'Willkommen, Schul-IT',
        lead: 'Portal und IT-Bereich, Stammdaten-Pfad, Kacheln der Lehrkräfte und der Feedback-Kanal.',
        onComplete(reason) {
            if (reason === 'skipped') {
                finishFirstRun('skipped');
                return;
            }
            patchItFirstRunState(localStorage, { phase: 'dashboard', step: 0 });
            window.setTimeout(() => {
                startSpotlightTour(dashboardSpotlightSteps(), {
                    onComplete(spotReason) {
                        restoreItTourLayout();
                        if (readItFirstRunState(localStorage).phase === 'stammdaten') return;
                        finishFirstRun(spotReason === 'skipped' ? 'skipped' : 'done');
                    }
                });
            }, 180);
        }
    });
    return true;
}

/**
 * @param {{ force?: boolean }} [opts]
 */
export async function startItStammdatenFirstRun(opts) {
    const force = !!(opts && opts.force);
    const persona = await resolveItPersona(force ? 1500 : 6000);
    const isIt = !!persona.isIt;
    const loggedIn = !!persona.loggedIn || isLoggedInAccount();
    const state = readItFirstRunState(localStorage);

    if (
        !shouldStartItStammdatenTour({
            force,
            isIt,
            loggedIn,
            phase: state.phase
        })
    ) {
        return false;
    }

    window.setTimeout(() => {
        startSpotlightTour(stammdatenSpotlightSteps(), {
            skipLabel: 'Einrichtung später',
            onComplete(reason) {
                finishFirstRun(reason === 'skipped' ? 'skipped' : 'done');
            }
        });
    }, 350);
    return true;
}

function urlForceTour() {
    try {
        const q = new URLSearchParams(window.location.search || '');
        const v = String(q.get('tour') || '').trim().toLowerCase();
        return v === 'it' || v === '1' || v === 'first-run';
    } catch {
        return false;
    }
}

function urlForceSaveTip() {
    try {
        const q = new URLSearchParams(window.location.search || '');
        const v = String(q.get('tour') || '').trim().toLowerCase();
        return v === 'save' || v === 'first-save' || v === 'backup';
    } catch {
        return false;
    }
}

function wireChromeListeners() {
    if (wireChromeListeners.bound) return;
    wireChromeListeners.bound = true;
    const refresh = () => {
        refreshItFirstRunChrome();
    };
    window.addEventListener('ms365-dashboard-persona-ready', refresh);
    window.addEventListener('ms365-auth-state-changed', refresh);
    // Menü wird oft erst nach Auth-Widget gebaut
    window.setTimeout(refresh, 800);
    window.setTimeout(refresh, 2000);
}

/**
 * @param {{ force?: boolean, page?: 'dashboard'|'stammdaten' }} [opts]
 */
export function mountItFirstRunTour(opts) {
    const page = (opts && opts.page) || 'dashboard';
    const force = !!(opts && opts.force) || urlForceTour();
    wireChromeListeners();
    // Speichern unten (Stammdaten) + Floppy oben (Datensicherung) auf Dashboard/Tenant
    wireFirstSaveTipListener();
    const forceSaveTip = urlForceSaveTip();

    const run = () => {
        refreshItFirstRunChrome();
        if (force) {
            try {
                localStorage.removeItem('ms365-onboarding-welcome-v1');
                localStorage.removeItem('ms365-it-first-save-tip-v1');
                if (page === 'stammdaten') {
                    patchItFirstRunState(localStorage, { phase: 'stammdaten', step: 0 });
                } else {
                    localStorage.removeItem('ms365-it-first-run-v1');
                }
                setChipDismissed(false);
            } catch {
                /* ignore */
            }
        }
        if (forceSaveTip && (page === 'stammdaten' || isTenantPage())) {
            try {
                localStorage.removeItem('ms365-it-first-save-tip-v1');
            } catch {
                /* ignore */
            }
            startFirstSaveTipTour({ force: true, reason: 'manual-save' });
            window.setTimeout(() => refreshItFirstRunChrome(), 400);
            return;
        }
        if (page === 'stammdaten') startItStammdatenFirstRun({ force });
        else startItDashboardFirstRun({ force });
        window.setTimeout(() => refreshItFirstRunChrome(), 400);
    };
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => window.setTimeout(run, 600));
    } else {
        window.setTimeout(run, 600);
    }
}

window.ms365ItFirstRun = {
    start: (o) => startItDashboardFirstRun(o || { force: true }),
    startStammdaten: (o) => startItStammdatenFirstRun(o || { force: true }),
    startFirstSaveTip: (o) => startFirstSaveTipTour(o || { force: true, reason: 'manual-save' }),
    restart: (o) => restartItFirstRun(o),
    resume: () => resumeItFirstRun(),
    reset: () => resetItFirstRun(),
    refreshChrome: () => refreshItFirstRunChrome(),
    mount: mountItFirstRunTour,
    openFeedback: openFeedbackPrefill
};
