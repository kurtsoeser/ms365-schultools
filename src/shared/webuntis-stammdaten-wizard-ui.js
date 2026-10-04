/**
 * Legacy-Einstieg: leitet auf die Import-Vollseite um (kein Modal mehr).
 */
import { resolveReturnUrl } from './webuntis-stammdaten-import-handoff.js';

function importPageHref(from) {
    const base = 'tools/webuntis-stammdaten-import.html';
    const f = from || 'tenant';
    return base + '?from=' + encodeURIComponent(f);
}

/**
 * @param {object} [options]
 * @param {'tenant'|'einrichtung'} [options.from]
 * @deprecated Nutzen Sie direkt tools/webuntis-stammdaten-import.html
 */
export function openWebUntisStammdatenWizard(options) {
    const from =
        options && options.from
            ? options.from
            : document.getElementById('tenantSettingsForm')
              ? 'tenant'
              : document.getElementById('swBtnWebuntisStammdatenWizard') || document.getElementById('setupWizardForm')
                ? 'einrichtung'
                : 'tenant';
    try {
        sessionStorage.setItem('ms365WebuntisImportReturn', resolveReturnUrl(from));
    } catch {
        /* ignore */
    }
    window.location.href = importPageHref(from);
    return Promise.resolve({ cancelled: true });
}
