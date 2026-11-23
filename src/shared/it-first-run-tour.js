/**
 * First-Run-Tour für Schul-IT: Willkommen → Dashboard-Spotlight → Stammdaten-Fortsetzung.
 */
import { getDashboardPersonas } from './dashboard-persona-session.js';
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
    '<p><strong>Feedback ist erwünscht.</strong> In den Pilotwochen hilft uns jeder Hinweis – unten in der Fußzeile unter <em>Feedback</em> oder jederzeit über denselben Dialog.</p>';

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
            title: 'Sie sind als Schul-IT unterwegs',
            bodyHtml:
                '<p>Als Tenant-Administrator sehen Sie das <strong>IT-Backend</strong>: Stammdaten, Playbooks, Gruppenhygiene und die Einrichtung der Schul-Apps.</p>' +
                '<ul><li>Admin-Zustimmung (Consent) in Microsoft&nbsp;365 freigeben</li>' +
                '<li>Schuldaten einmal zentral pflegen – die Apps greifen darauf zu</li>' +
                '<li>Lehrkräfte und Schüler:innen sehen später ihr Portal, nicht dieses Backend</li></ul>'
        },
        {
            id: 'stammdaten',
            title: 'Sinnvoller Start: Stammdaten + Tool-Site',
            bodyHtml:
                '<p>Empfohlene Reihenfolge für neue Schulen:</p>' +
                '<ul><li><strong>Schulprofil</strong> – Name, Domain, Logo</li>' +
                '<li><strong>Lehrkräfte &amp; Klassen</strong> – Import (z. B. WebUntis) oder manuell</li>' +
                '<li><strong>Primär-Site (SharePoint)</strong> – Website im Schul-Tenant für die Tool-Listen</li>' +
                '<li>Danach die Blaupause <strong>Erste Einrichtung</strong> (Playbook) abhaken</li></ul>' +
                '<p class="muted">Die Tour orientiert nur – die Checkliste im Playbook ist die Arbeitsgrundlage.</p>'
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
    return [
        {
            id: 'greeting',
            selector: '#dashGreetingBand',
            title: 'Ihr IT-Überblick',
            bodyHtml:
                '<p>Hier sehen Sie Begrüßung und Kurzstatus. Als Schul-IT ist das der Einstieg ins Backend – nicht die Lehrer-Portal-Ansicht.</p>',
            placement: 'bottom',
            before: async () => {
                setDashboardMainSection('tasks');
            }
        },
        {
            id: 'import',
            selector: '#dashTaskImportVerknuepfen',
            title: 'Blaupause: Erste Einrichtung',
            bodyHtml:
                '<p>Das Playbook <strong>Erste Einrichtung</strong> ist Ihre Checkliste – nicht nur „Daten importieren“, sondern der rote Faden: Profil, Tool-Site, Import, Verknüpfen, Sicherung.</p>' +
                '<p>Die Tour hier orientiert nur; im Playbook arbeiten Sie die Schritte mit Haken ab.</p>',
            placement: 'right',
            padding: 6,
            before: async () => {
                setDashboardMainSection('tasks');
                await waitForSelector('#dashTaskImportVerknuepfen', 1500);
            }
        },
        {
            id: 'playbooks',
            selector: '#dashTaskSchuljahr, [data-task-id="dashTaskSchuljahr"]',
            title: 'Weitere Playbooks in den Bereichen',
            bodyHtml:
                '<p>Schuljahresstart, Intranet, Freistellungen &amp; Co. sitzen in der linken Navigation – jeweils in der passenden Gruppe. Zuerst Fundament und Import-Checkliste.</p>',
            placement: 'right',
            before: async () => {
                setDashboardMainSection('tasks');
                try {
                    if (window.ms365DashTasksSplit && typeof window.ms365DashTasksSplit.select === 'function') {
                        window.ms365DashTasksSplit.select('dashTaskSchuljahr', { pulse: false });
                    }
                } catch {
                    /* ignore */
                }
                await waitForSelector('#dashTaskSchuljahr, [data-task-id="dashTaskSchuljahr"]', 1500);
            }
        },
        {
            id: 'feedback',
            selector: '.site-feedback-link, #siteCreditKurtrocks',
            title: 'Feedback immer griffbereit',
            primaryLabel: 'Zur Blaupause (Playbook)',
            bodyHtml:
                FEEDBACK_HINT +
                '<p>Als Nächstes: Playbook öffnen – oder kurz Stammdaten-Schulprofil ansehen und dann die Checkliste.</p>',
            placement: 'top',
            before: async () => {
                await waitForSelector('.site-feedback-link', 3000);
            },
            onPrimary: async () => {
                markWelcomeTourDone(localStorage, 'seen');
                patchItFirstRunState(localStorage, { phase: 'done', step: 0 });
                window.location.href = 'tools/playbook-daten-import-verknuepfen.html';
                return 'finish';
            }
        }
    ];
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
                '<p>Aus Microsoft&nbsp;365 übernehmen, WebUntis/CSV importieren oder manuell pflegen. Die Apps brauchen diese Liste für Rollen und Zielgruppen.</p>',
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
                '<p>Klassenliste mit Klassenvorstand – Grundlage für Gruppen, Planer und Kommunikation.</p>',
            placement: 'bottom',
            before: async () => {
                activateTenantMainTab('tabMainKlassen');
                await waitForSelector('#tabMainKlassen, #panelMainKlassen', 1500);
            }
        },
        {
            id: 'feedback-end',
            selector: '.site-feedback-link, #siteCreditKurtrocks',
            title: 'Als Nächstes: Blaupause',
            primaryLabel: 'Blaupause öffnen',
            bodyHtml:
                '<p>Schulprofil können Sie jederzeit nachziehen. Die Checkliste <strong>Erste Einrichtung</strong> führt Import, Verknüpfen und Sicherung – dort arbeiten Sie mit Haken weiter.</p>' +
                FEEDBACK_HINT,
            placement: 'top',
            before: async () => {
                activateTenantMainTab('tabMainStammdaten');
                activateStammdatenPane('profile');
                await waitForSelector('.site-feedback-link', 2500);
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
            '<p>Im Panel: <em>IT-Sicherungsbibliothek einrichten</em> – oder unter Stammdaten → SharePoint &amp; Dienste.</p>' +
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
                '<p>Damit Lehrkräfte Schulaktivitäten, Freistellungen &amp; Co. nutzen können, brauchen die Apps eine <strong>SharePoint-Website in Ihrem Schul-Tenant</strong> – dort liegen die Listen.</p>' +
                '<p>Unter <strong>Stammdaten → Planer &amp; Apps</strong> die Primär-Site eintragen und prüfen. Diese URL wird automatisch dem Betreiber gemeldet (Feld „Tool-Site-URL“), damit das System weiß, wo Ihre Schule arbeitet.</p>',
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
        lead: 'Die nächsten Minuten: Orientierung, Stammdaten-Pfad und Feedback-Kanal für die Pilotwochen.',
        onComplete(reason) {
            if (reason === 'skipped') {
                finishFirstRun('skipped');
                return;
            }
            patchItFirstRunState(localStorage, { phase: 'dashboard', step: 0 });
            window.setTimeout(() => {
                startSpotlightTour(dashboardSpotlightSteps(), {
                    onComplete(spotReason) {
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
