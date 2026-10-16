/**
 * Stammdaten Lokal-First (Umbau): kein SharePoint-/IT-Bibliothek-Sync,
 * kein Intranet-Listen-Export – nur lokaler v2-Speicher + manueller JSON-Export.
 *
 * Flag auf false setzen, wenn Phase 4 (manuelle IT-Bib) wieder freigeschaltet wird.
 * Siehe docs/stammdaten-local-first.md
 */
export const STAMMDATEN_LOCAL_FIRST = true;

export const LOCAL_FIRST_BANNER =
    'Lokal-First-Umbau: Speichern nur in diesem Browser. SharePoint-IT-Bibliothek und Intranet-Listen-Export sind deaktiviert. Sicherung: manuelles Browser-Backup (JSON).';

export function isStammdatenLocalFirst() {
    return STAMMDATEN_LOCAL_FIRST === true;
}

/** Netzwerk-Sync IT-Bibliothek (Auto + manuell Upload/Load). */
export function isSpoNetworkSyncAllowed() {
    return !isStammdatenLocalFirst();
}

/** Intranet-Stammdaten-Listen nach SharePoint schreiben. */
export function isIntranetExportAllowed() {
    return !isStammdatenLocalFirst();
}

export function localFirstStatusMessage() {
    return LOCAL_FIRST_BANNER;
}

export function localFirstShortLabel() {
    return 'Lokal-First: nur Browser';
}

/**
 * UI-Elemente für SPO-Sync / IT-Bib deaktivieren und beschriften.
 * @param {ParentNode|Document|null|undefined} [root]
 */
export function applyLocalFirstSpoUi(root) {
    if (!isStammdatenLocalFirst()) return;
    const doc = root && root.querySelectorAll ? root : typeof document !== 'undefined' ? document : null;
    if (!doc) return;
    const nodes = doc.querySelectorAll(
        '[data-ms365-spo-sync="upload"], [data-ms365-spo-sync="load"], [data-ms365-spo-sync="setup"]'
    );
    nodes.forEach(function (el) {
        if (!el || el.dataset.localFirstDisabled === '1') return;
        el.dataset.localFirstDisabled = '1';
        el.setAttribute('aria-disabled', 'true');
        el.title = LOCAL_FIRST_BANNER;
        if (el.tagName === 'BUTTON' || el.tagName === 'INPUT') {
            el.disabled = true;
        } else if (el.tagName === 'A') {
            el.addEventListener(
                'click',
                function (ev) {
                    ev.preventDefault();
                    ev.stopPropagation();
                    if (typeof window.ms365ShowToast === 'function') {
                        window.ms365ShowToast(LOCAL_FIRST_BANNER, { kind: 'info', title: 'Lokal-First' });
                    } else if (typeof window.ms365ToastOrAlert === 'function') {
                        window.ms365ToastOrAlert(LOCAL_FIRST_BANNER);
                    }
                },
                true
            );
            el.style.pointerEvents = 'none';
            el.style.opacity = '0.55';
        }
    });
}

/**
 * Intranet-Sync-Buttons deaktivieren.
 * @param {ParentNode|Document|null|undefined} [root]
 */
export function applyLocalFirstIntranetUi(root) {
    if (!isStammdatenLocalFirst()) return;
    const doc = root && root.querySelectorAll ? root : typeof document !== 'undefined' ? document : null;
    if (!doc) return;
    doc.querySelectorAll('[data-tenant-intranet-sync]').forEach(function (btn) {
        if (!btn || btn.dataset.localFirstDisabled === '1') return;
        btn.dataset.localFirstDisabled = '1';
        btn.disabled = true;
        btn.title = LOCAL_FIRST_BANNER;
        btn.setAttribute('aria-disabled', 'true');
    });
}

if (typeof window !== 'undefined') {
    window.ms365StammdatenLocalFirst = {
        STAMMDATEN_LOCAL_FIRST,
        LOCAL_FIRST_BANNER,
        isStammdatenLocalFirst,
        isSpoNetworkSyncAllowed,
        isIntranetExportAllowed,
        localFirstStatusMessage,
        localFirstShortLabel,
        applyLocalFirstSpoUi,
        applyLocalFirstIntranetUi
    };
}

export default {
    STAMMDATEN_LOCAL_FIRST,
    LOCAL_FIRST_BANNER,
    isStammdatenLocalFirst,
    isSpoNetworkSyncAllowed,
    isIntranetExportAllowed,
    localFirstStatusMessage,
    localFirstShortLabel,
    applyLocalFirstSpoUi,
    applyLocalFirstIntranetUi
};
