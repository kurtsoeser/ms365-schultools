/**
 * Abgleich lokales vs. SharePoint-Backup und Nutzerentscheidung (Import / behalten / hochladen).
 */
import {
    compareBackupPayloads,
    formatBackupCompareDe,
    isLikelyFreshLocalBackup
} from './stammdaten-sharepoint-sync-logic.js';
import { promptBackupCompareChoice } from './stammdaten-sharepoint-backup-compare-ui.js';

function confirmAsync(msg, opts) {
    const o = opts || {};
    if (typeof window.ms365AppDialogConfirm === 'function') {
        return window.ms365AppDialogConfirm(msg, {
            title: o.title,
            okText: o.confirmLabel,
            cancelText: o.cancelLabel,
            kind: o.kind
        });
    }
    return Promise.resolve(window.confirm(msg));
}

function canUseCompareUi() {
    return typeof document !== 'undefined' && document.body;
}

export function getLocalBackupSnapshot() {
    try {
        const bb = window.ms365BrowserBackup;
        if (bb && typeof bb.buildBackup === 'function') return bb.buildBackup();
    } catch {
        /* ignore */
    }
    return null;
}

/**
 * @param {object} remotePayload SharePoint-Backup-JSON
 * @param {{ remoteLastModified?: string, auto?: boolean }} [ctx]
 * @returns {Promise<'import'|'keep-local'|'push-local'|'skip'>}
 */
export async function resolveSharePointImportChoice(remotePayload, ctx) {
    const options = ctx || {};
    const local = getLocalBackupSnapshot();
    const cmp = compareBackupPayloads(local, remotePayload);
    if (cmp.identical) {
        return 'skip';
    }

    if (options.auto && isLikelyFreshLocalBackup(local)) {
        return 'import';
    }

    const title = options.auto ? 'Sicherung nach Anmeldung' : 'Von SharePoint einlesen';

    if (canUseCompareUi()) {
        try {
            const choice = await promptBackupCompareChoice({
                compare: cmp,
                remotePayload: remotePayload,
                remoteLastModified: options.remoteLastModified,
                title: title
            });
            if (choice === 'import' || choice === 'keep-local' || choice === 'push-local') {
                return choice;
            }
        } catch {
            /* Fallback Textdialog */
        }
    }

    const body =
        formatBackupCompareDe(cmp, { remoteLastModified: options.remoteLastModified }) +
        '\n\n' +
        (cmp.newerSide === 'remote'
            ? 'SharePoint-Stand übernehmen? (Abbrechen = lokal behalten)'
            : cmp.newerSide === 'local'
              ? 'Lokal ist neuer. Trotzdem von SharePoint übernehmen? (Abbrechen = lokal behalten)'
              : 'Von SharePoint übernehmen? (Abbrechen = lokal behalten)');

    const ok = await confirmAsync(body, {
        title: title,
        confirmLabel: 'Von SharePoint übernehmen',
        cancelLabel: 'Lokal behalten',
        kind: cmp.newerSide === 'local' ? 'warning' : 'info'
    });
    if (ok) return 'import';

    if (cmp.newerSide === 'local') {
        const push = await confirmAsync(
            'Lokalen Stand jetzt nach SharePoint hochladen?\n\nDamit wird die Sicherungsdatei auf SharePoint aktualisiert.',
            {
                title: 'Nach SharePoint sichern',
                confirmLabel: 'Hochladen',
                cancelLabel: 'Später',
                kind: 'info'
            }
        );
        if (push) return 'push-local';
    }

    return 'keep-local';
}

export default {
    getLocalBackupSnapshot,
    resolveSharePointImportChoice
};
