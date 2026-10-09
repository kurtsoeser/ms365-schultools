/**
 * UI für Zugriffsdiagnose (Planer).
 */
import {
    runFreistellungAccessDiagnostics,
    formatDiagnosticsReport,
    isFreistellungAccessDebugEnabled
} from './freistellung-planer-access-debug.js';

function esc(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

const STATUS_LABEL = {
    ok: 'OK',
    warn: 'Hinweis',
    fail: 'Fehler',
    skip: '–',
    info: 'Info'
};

export function shouldShowFreistellungAccessDebugEntry(state) {
    if (isFreistellungAccessDebugEnabled()) return true;
    return !!(state && state.planerAccessDenied && state.accountEmail);
}

export function renderFreistellungAccessDebugEntry(state) {
    if (!shouldShowFreistellungAccessDebugEntry(state)) return '';
    const open = isFreistellungAccessDebugEnabled();
    return `
    <div class="fr-access-debug" id="frAccessDebugWrap" data-fr-debug-open="${open ? '1' : '0'}">
      <p class="fr-access-debug__lead">
        <button type="button" class="btn btn-sm" id="frAccessDebugRun" aria-expanded="${open ? 'true' : 'false'}">
          <i class="bi bi-clipboard2-pulse"></i>Zugriff Schritt für Schritt prüfen
        </button>
        <span class="muted" style="font-size:0.85rem;margin-left:6px;">Für IT/KV – zeigt nur dieses Konto &amp; technische Checks.</span>
      </p>
      <div class="fr-access-debug__panel" id="frAccessDebugPanel" ${open ? '' : 'hidden'}>
        <div class="fr-access-debug__toolbar">
          <button type="button" class="btn btn-sm" id="frAccessDebugCopy" disabled>Bericht kopieren</button>
          <span class="muted" id="frAccessDebugStatus" aria-live="polite"></span>
        </div>
        <ol class="fr-access-debug__steps" id="frAccessDebugSteps"></ol>
        <p class="fr-access-debug__summary muted" id="frAccessDebugSummary"></p>
      </div>
    </div>`;
}

function renderSteps(steps) {
    return (steps || [])
        .map(function (st) {
            const lab = STATUS_LABEL[st.status] || st.status;
            return (
                '<li class="fr-access-debug__step fr-access-debug__step--' +
                esc(st.status) +
                '">' +
                '<span class="fr-access-debug__badge">' +
                esc(lab) +
                '</span> ' +
                '<strong>' +
                esc(st.title) +
                '</strong>' +
                (st.detail ? '<div class="fr-access-debug__detail">' + esc(st.detail) + '</div>' : '') +
                '</li>'
            );
        })
        .join('');
}

let lastReport = null;

/**
 * @param {HTMLElement} root
 * @param {object} state
 */
export function wireFreistellungAccessDebug(root, state) {
    if (!root) return;
    const wrap = root.querySelector('#frAccessDebugWrap');
    if (!wrap) return;

    const btnRun = root.querySelector('#frAccessDebugRun');
    const panel = root.querySelector('#frAccessDebugPanel');
    const stepsEl = root.querySelector('#frAccessDebugSteps');
    const summaryEl = root.querySelector('#frAccessDebugSummary');
    const statusEl = root.querySelector('#frAccessDebugStatus');
    const btnCopy = root.querySelector('#frAccessDebugCopy');

    async function run() {
        if (panel) panel.hidden = false;
        if (btnRun) btnRun.setAttribute('aria-expanded', 'true');
        if (statusEl) statusEl.textContent = 'Prüfe …';
        if (stepsEl) stepsEl.innerHTML = '<li class="muted">Bitte warten …</li>';
        if (summaryEl) summaryEl.textContent = '';
        if (btnCopy) btnCopy.disabled = true;
        try {
            const report = await Promise.race([
                runFreistellungAccessDiagnostics(state, { refreshRemote: true }),
                new Promise(function (_, reject) {
                    setTimeout(function () {
                        reject(new Error('Diagnose-Timeout (90 s) – Seite neu laden und erneut versuchen.'));
                    }, 90000);
                })
            ]);
            lastReport = report;
            if (stepsEl) stepsEl.innerHTML = renderSteps(report.steps);
            if (summaryEl) summaryEl.textContent = report.summary || '';
            if (statusEl) statusEl.textContent = 'Fertig – ' + new Date().toLocaleTimeString('de-AT');
            if (btnCopy) btnCopy.disabled = false;
        } catch (e) {
            if (stepsEl) {
                stepsEl.innerHTML =
                    '<li class="fr-access-debug__step fr-access-debug__step--fail"><strong>Diagnose abgebrochen</strong><div class="fr-access-debug__detail">' +
                    esc(e && e.message ? e.message : String(e)) +
                    '</div></li>';
            }
            if (statusEl) statusEl.textContent = 'Fehler';
        }
    }

    if (btnRun) {
        btnRun.addEventListener('click', function () {
            run();
        });
    }
    if (btnCopy) {
        btnCopy.addEventListener('click', function () {
            if (!lastReport) return;
            const text = formatDiagnosticsReport(lastReport);
            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(text).then(
                    function () {
                        if (statusEl) statusEl.textContent = 'Bericht in Zwischenablage.';
                    },
                    function () {
                        window.prompt('Bericht kopieren:', text);
                    }
                );
            } else {
                window.prompt('Bericht kopieren:', text);
            }
        });
    }

    if (wrap.getAttribute('data-fr-debug-open') === '1') {
        run();
    }
}
