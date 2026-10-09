/**
 * Nach Login: wenn IT-Bibliothek fehlt und Auto-Link/Discovery scheitern –
 * einmal pro Tab-Session Site-URL abfragen oder auf Einrichtung verweisen.
 */
import { IT_LIBRARY_TITLE, collectItLibraryLinkHints } from './stammdaten-sharepoint-sync-logic.js';
import {
    setupPageHref,
    writeItLibraryFormDraft,
    readItLibraryFormDraft,
    loadItMeta
} from './stammdaten-sharepoint-sync-api.js';
import { tryAutoLinkItLibrary } from './stammdaten-sharepoint-auto-link.js';

const SESSION_PROMPT_KEY_PREFIX = 'ms365-it-library-prompt-done-v1';

function toast(m, opts) {
    if (typeof window.ms365ShowToast === 'function') {
        window.ms365ShowToast(m, opts || {});
        return;
    }
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
}

function getTenantId() {
    try {
        if (typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            if (info && info.tenantId) return String(info.tenantId).trim();
        }
    } catch {
        /* ignore */
    }
    return '';
}

function sessionPromptKey() {
    const tid = getTenantId();
    return SESSION_PROMPT_KEY_PREFIX + (tid ? ':' + tid : '');
}

export function wasItLibraryPromptDoneThisSession() {
    try {
        return sessionStorage.getItem(sessionPromptKey()) === '1';
    } catch {
        return false;
    }
}

export function markItLibraryPromptDoneThisSession() {
    try {
        sessionStorage.setItem(sessionPromptKey(), '1');
    } catch {
        /* ignore */
    }
}

function readSetup() {
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
            return window.ms365AppDataV2.getSetup() || {};
        }
    } catch {
        /* ignore */
    }
    return {};
}

function defaultSiteSuggestion() {
    const hints = collectItLibraryLinkHints({
        itMeta: loadItMeta(),
        setup: readSetup(),
        formDraft: readItLibraryFormDraft()
    });
    return hints.siteUrl || '';
}

function promptAsync(message, defaultValue, options) {
    if (typeof window.ms365AppDialogPrompt === 'function') {
        return window.ms365AppDialogPrompt(message, defaultValue, options || {});
    }
    const v = window.prompt(String(message || ''), defaultValue != null ? String(defaultValue) : '');
    return Promise.resolve(v == null ? null : String(v));
}

function confirmAsync(message, options) {
    if (typeof window.ms365AppDialogConfirm === 'function') {
        return window.ms365AppDialogConfirm(message, options || {});
    }
    return Promise.resolve(window.confirm(String(message || '')));
}

function looksLikeSiteUrl(value) {
    const s = String(value || '').trim();
    if (!s) return false;
    if (!/^https?:\/\//i.test(s)) return false;
    return /sharepoint\.com|\.sharepoint\./i.test(s) || /\/sites\//i.test(s);
}

/**
 * @param {{ skipped?: string, error?: string }} [linkResult]
 * @returns {Promise<{ linked?: boolean, skipped?: string, meta?: object, via?: string }>}
 */
export async function promptForMissingItLibrary(linkResult) {
    if (wasItLibraryPromptDoneThisSession()) {
        return { linked: false, skipped: 'prompt-already-done' };
    }
    markItLibraryPromptDoneThisSession();

    const skipped = linkResult && linkResult.skipped ? String(linkResult.skipped) : '';
    if (skipped === 'no-token') {
        toast(
            'IT-Sicherungsbibliothek: SharePoint-Rechte fehlen (Sites.ReadWrite.All). Mit IT-Konto anmelden oder Zustimmung prüfen.',
            { kind: 'warning', title: 'IT-Sicherung' }
        );
        return { linked: false, skipped: 'no-token' };
    }

    const title = IT_LIBRARY_TITLE;
    const want = await confirmAsync(
        'Auf diesem Browser ist noch keine IT-Sicherungsbibliothek verknüpft.\n\n' +
            'Wenn an der Schule bereits die Bibliothek „' +
            title +
            '“ existiert, können Sie jetzt die SharePoint-Site-URL angeben – ' +
            'die App verknüpft sie und gleicht die Sicherung ab.\n\n' +
            'Site-URL jetzt eingeben?',
        {
            title: 'IT-Sicherungsbibliothek',
            okText: 'Site-URL eingeben',
            cancelText: 'Später',
            kind: 'info'
        }
    );

    if (!want) {
        toast(
            'Hinweis: IT-Sicherungsbibliothek unter Stammdaten-Übergabe einrichten oder Site-URL später angeben.',
            { kind: 'info', title: 'IT-Sicherung', durationMs: 8000 }
        );
        return { linked: false, skipped: 'prompt-declined' };
    }

    const suggested = defaultSiteSuggestion();
    const entered = await promptAsync(
        'SharePoint-Site-URL der Schule (z. B. https://schule.sharepoint.com/sites/IT oder Intranet-Site), ' +
            'auf der die Bibliothek „' +
            title +
            '“ liegt:',
        suggested,
        {
            title: 'IT-Sicherungsbibliothek verknüpfen',
            inputLabel: 'Site-URL',
            okText: 'Verknüpfen',
            cancelText: 'Abbrechen',
            kind: 'info'
        }
    );

    if (entered == null) {
        return { linked: false, skipped: 'prompt-cancelled' };
    }

    const siteUrl = String(entered || '').trim().replace(/\/$/, '');
    if (!looksLikeSiteUrl(siteUrl)) {
        toast('Bitte eine gültige SharePoint-Site-URL eingeben (https://…/sites/…).', {
            kind: 'warning',
            title: 'IT-Sicherung'
        });
        const goSetup = await confirmAsync(
            'Die Eingabe sieht nicht nach einer SharePoint-Site-URL aus.\n\nZur Einrichtungsseite wechseln?',
            {
                title: 'IT-Sicherungsbibliothek',
                okText: 'Einrichten öffnen',
                cancelText: 'Schließen',
                kind: 'warning'
            }
        );
        if (goSetup) {
            window.location.href = setupPageHref();
            return { linked: false, skipped: 'goto-setup' };
        }
        return { linked: false, skipped: 'invalid-url' };
    }

    try {
        writeItLibraryFormDraft({ siteUrl: siteUrl, libraryTitle: title });
    } catch {
        /* ignore */
    }

    const result = await tryAutoLinkItLibrary({
        siteUrl: siteUrl,
        listTitle: title,
        allowDiscover: true
    });

    if (result && result.linked) {
        return Object.assign({}, result, { via: result.via || 'prompt' });
    }

    const goSetup = await confirmAsync(
        'Auf dieser Site wurde die Bibliothek „' +
            title +
            '“ nicht gefunden (oder der Zugriff fehlt).\n\n' +
            'Zur Einrichtungsseite wechseln und die IT-Sicherungsbibliothek anlegen bzw. prüfen?',
        {
            title: 'Bibliothek nicht gefunden',
            okText: 'Einrichten öffnen',
            cancelText: 'Schließen',
            kind: 'warning'
        }
    );
    if (goSetup) {
        window.location.href = setupPageHref();
        return { linked: false, skipped: 'goto-setup', error: result && result.error };
    }
    return {
        linked: false,
        skipped: (result && result.skipped) || 'library-not-found',
        error: result && result.error
    };
}

/**
 * Dialog nur auf Übersichts-/IT-Seiten – nicht in jedem Fachwerkzeug stören.
 * Discovery/Auto-Link läuft trotzdem überall still.
 */
export function shouldShowItLibraryPromptOnThisPage(pathname) {
    const p = String(
        pathname != null
            ? pathname
            : (typeof window !== 'undefined' && window.location && window.location.pathname) || ''
    ).replace(/\\/g, '/');
    if (/\/tools\//i.test(p)) return false;
    if (/stammdaten-uebergabe\.html/i.test(p)) return false;
    if (/stammdaten-backup-abgleich\.html/i.test(p)) return false;
    if (/\/index\.html$/i.test(p)) return true;
    if (p === '/' || /^\/[^/]+\/?$/i.test(p)) return true;
    if (/\/tenant\.html$/i.test(p)) return true;
    if (/dashboard-werkzeug-zugriff\.html$/i.test(p)) return true;
    return false;
}

/**
 * Ob nach fehlgeschlagenem Auto-Link ein Prompt sinnvoll ist.
 * @param {{ skipped?: string }|null|undefined} linkResult
 * @param {string} [pathname]
 */
export function shouldPromptForMissingItLibrary(linkResult, pathname) {
    if (!linkResult || linkResult.linked) return false;
    if (!shouldShowItLibraryPromptOnThisPage(pathname)) return false;
    if (wasItLibraryPromptDoneThisSession()) return false;
    const s = String(linkResult.skipped || '');
    // no-token: stiller Toast nur auf IT-Seiten; kein Site-URL-Dialog ohne Graph-Recht
    if (s === 'no-token') return true;
    return (
        s === 'no-hints' ||
        s === 'library-not-found' ||
        s === 'site-unresolved' ||
        s === 'site-missing'
    );
}

export default {
    promptForMissingItLibrary,
    shouldPromptForMissingItLibrary,
    shouldShowItLibraryPromptOnThisPage,
    wasItLibraryPromptDoneThisSession,
    markItLibraryPromptDoneThisSession
};
