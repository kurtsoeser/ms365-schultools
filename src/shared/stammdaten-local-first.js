/**
 * Stammdaten Lokal-First (Umbau): kanonisch lokal im Browser.
 *
 * Phase 4: manuelle IT-Sicherungsbibliothek (Einrichten + Sichern + Laden)
 * ist wieder erlaubt. Auto-Sync/Abgleich und Intranet-Listen-Export bleiben aus.
 *
 * Siehe docs/stammdaten-local-first.md
 */
export const STAMMDATEN_LOCAL_FIRST = true;

/** Manuelles Einrichten/Sichern/Laden in die IT-Bibliothek (ohne Auto-Sync). */
export const STAMMDATEN_IT_MANUAL_BACKUP = true;

export const LOCAL_FIRST_BANNER = STAMMDATEN_IT_MANUAL_BACKUP
    ? 'Lokal-First: Speichern in diesem Browser. IT-Sicherungsbibliothek: manuell sichern und laden (mit Versionierung). Kein Auto-Abgleich. Intranet-Listen-Export pausiert.'
    : 'Lokal-First-Umbau: Speichern nur in diesem Browser. SharePoint-IT-Bibliothek und Intranet-Listen-Export sind deaktiviert. Sicherung: manuelles Browser-Backup (JSON).';

export function isStammdatenLocalFirst() {
    return STAMMDATEN_LOCAL_FIRST === true;
}

/** Manuelles Upload/Load/Setup der IT-Sicherungsbibliothek. */
export function isSpoManualBackupAllowed() {
    if (STAMMDATEN_IT_MANUAL_BACKUP === true) return true;
    return !isStammdatenLocalFirst();
}

/** Auto-Push/Pull und Konflikt-/Abgleich-Logik. */
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
    return STAMMDATEN_IT_MANUAL_BACKUP
        ? 'Lokal-First · IT manuell'
        : 'Lokal-First: nur Browser';
}

/**
 * UI für manuelle SPO-Aktionen nur sperren, wenn Phase-4-Manual aus ist.
 * @param {ParentNode|Document|null|undefined} [root]
 */
export function applyLocalFirstSpoUi(root) {
    if (isSpoManualBackupAllowed()) return;
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
 * Abgleich-/Auto-Sync-UI während Lokal-First ausblenden.
 * @param {ParentNode|Document|null|undefined} [root]
 */
export function applyLocalFirstCompareUi(root) {
    if (!isStammdatenLocalFirst()) return;
    const doc = root && root.querySelectorAll ? root : typeof document !== 'undefined' ? document : null;
    if (!doc) return;
    doc.querySelectorAll('[data-local-first-hide="spo-compare"]').forEach(function (el) {
        el.hidden = true;
        el.setAttribute('aria-hidden', 'true');
    });
    // Fallback: Abgleich-Blöcke ohne data-Attribut
    doc.querySelectorAll('.ms365-backup-panel__block').forEach(function (block) {
        const label = block.querySelector('.ms365-backup-panel__label');
        const t = label ? String(label.textContent || '') : '';
        if (/Abgleich/i.test(t)) {
            block.hidden = true;
            block.setAttribute('aria-hidden', 'true');
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
        STAMMDATEN_IT_MANUAL_BACKUP,
        LOCAL_FIRST_BANNER,
        isStammdatenLocalFirst,
        isSpoManualBackupAllowed,
        isSpoNetworkSyncAllowed,
        isIntranetExportAllowed,
        localFirstStatusMessage,
        localFirstShortLabel,
        applyLocalFirstSpoUi,
        applyLocalFirstCompareUi,
        applyLocalFirstIntranetUi
    };
}

export default {
    STAMMDATEN_LOCAL_FIRST,
    STAMMDATEN_IT_MANUAL_BACKUP,
    LOCAL_FIRST_BANNER,
    isStammdatenLocalFirst,
    isSpoManualBackupAllowed,
    isSpoNetworkSyncAllowed,
    isIntranetExportAllowed,
    localFirstStatusMessage,
    localFirstShortLabel,
    applyLocalFirstSpoUi,
    applyLocalFirstCompareUi,
    applyLocalFirstIntranetUi
};
