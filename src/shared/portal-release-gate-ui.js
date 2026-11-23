/**
 * IT-UI: Portal freigeben / zurücknehmen (b6).
 * Ort: Portal vorbereiten (dashboard-werkzeug-zugriff.html).
 */
import {
    evaluateItPortalReleaseGate,
    gatherLehrerPortalReadinessInput,
    isPortalReleased,
    patchPortalReleased,
    readPortalReleasedAt
} from './portal-release-gate.js';

function currentSetup() {
    try {
        return window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function'
            ? window.ms365AppDataV2.getSetup() || {}
            : {};
    } catch {
        return {};
    }
}

/**
 * @param {HTMLElement} body
 */
function render(body) {
    const input = gatherLehrerPortalReadinessInput();
    const gate = evaluateItPortalReleaseGate(input);
    const released = isPortalReleased(currentSetup());
    const at = readPortalReleasedAt(currentSetup());

    const status = released
        ? '<p class="dash-pg-section__note"><strong>Portal freigegeben</strong>' +
          (at ? ' · ' + new Date(at).toLocaleString('de-AT') : '') +
          '</p>'
        : '<p class="dash-pg-section__note"><strong>Portal nicht freigegeben</strong> – Lehrkräfte sehen sonst den Hinweis „Portal noch nicht freigegeben“.</p>';

    const miss =
        gate.missingLabels && gate.missingLabels.length
            ? '<ul class="ts-portal-release__list">' +
              gate.missingLabels.map((l) => '<li>' + String(l) + '</li>').join('') +
              '</ul>'
            : '';

    body.innerHTML =
        status +
        '<p class="dash-pg-section__lead" style="margin:0 0 0.75rem">' +
        gate.message +
        '</p>' +
        miss +
        '<div class="ts-portal-release__actions" style="display:flex;gap:0.5rem;flex-wrap:wrap;margin-top:0.5rem">' +
        '<button type="button" class="btn" id="dashPortalReleaseBtn"' +
        (gate.canRelease && !released ? '' : ' disabled') +
        '><i class="bi bi-door-open" aria-hidden="true"></i> Portal freigeben</button>' +
        '<button type="button" class="btn alt" id="dashPortalUnreleaseBtn"' +
        (released ? '' : ' disabled') +
        '>Freigabe zurücknehmen</button>' +
        '</div>';

    const btn = body.querySelector('#dashPortalReleaseBtn');
    const un = body.querySelector('#dashPortalUnreleaseBtn');
    if (btn) {
        btn.onclick = function () {
            try {
                patchPortalReleased(true);
                render(body);
            } catch (e) {
                window.alert((e && e.message) || String(e));
            }
        };
    }
    if (un) {
        un.onclick = function () {
            try {
                patchPortalReleased(false);
                render(body);
            } catch (e) {
                window.alert((e && e.message) || String(e));
            }
        };
    }
}

/**
 * @param {string} [mountSelector]
 */
export function mountPortalReleaseGateUi(mountSelector) {
    if (typeof document === 'undefined') return;
    const sel = mountSelector || '[data-ms365-dashboard-portal-release-mount]';
    const root = document.querySelector(sel);
    if (!root) return;

    root.innerHTML =
        '<section class="dash-pg-section dash-portal-release" aria-labelledby="dashPortalReleaseTitle">' +
        '<div class="dash-pg-section__head">' +
        '<div>' +
        '<h3 id="dashPortalReleaseTitle"><i class="bi bi-door-open" aria-hidden="true"></i> Lehrer-Portal freigeben</h3>' +
        '<p class="dash-pg-section__lead">' +
        'Erst freigeben, wenn Tool-Site und Klassen/Lehrer stehen. Danach verschwindet der Hinweis bei Lehrkräften.' +
        '</p>' +
        '</div>' +
        '</div>' +
        '<div data-portal-release-body></div>' +
        '</section>';

    const body = root.querySelector('[data-portal-release-body]');
    if (body) render(/** @type {HTMLElement} */ (body));
}

/** @deprecated Alias – früher auf tenant.html */
export function bootPortalReleaseGateUi() {
    mountPortalReleaseGateUi();
}

if (typeof window !== 'undefined') {
    window.ms365BootPortalReleaseGateUi = bootPortalReleaseGateUi;
    window.ms365MountPortalReleaseGateUi = mountPortalReleaseGateUi;
}
