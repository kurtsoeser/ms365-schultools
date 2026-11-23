/**
 * Einheitliche IT-Setup-Assistenten-Schale (Listen / Rechte / Paket / Flow).
 * Referenz-Optik: Schulaktivitäten-Listen (Website → Rollen → Paket).
 *
 * Domäne bleibt tool-spezifisch – nur Stepper, Phase-Hint, Footer, Summary, Protokoll.
 *
 * @see docs/it-setup-assistent.md
 */

/**
 * @param {number|string} n
 * @param {number} count
 */
export function clampItSetupStep(n, count) {
    const max = Math.max(1, Number(count) || 1);
    const v = parseInt(String(n), 10);
    if (!Number.isFinite(v) || v < 1) return 1;
    if (v > max) return max;
    return v;
}

/**
 * @param {number} step
 * @param {number} count
 * @param {string[]} labels
 */
export function formatItSetupPhaseHint(step, count, labels) {
    const s = clampItSetupStep(step, count);
    const label = Array.isArray(labels) && labels[s - 1] ? String(labels[s - 1]) : '';
    return label
        ? 'Schritt ' + s + ' von ' + count + ': ' + label
        : 'Schritt ' + s + ' von ' + count;
}

/**
 * HTML für Status-Chips (Phase + optional Rechte-Hinweis).
 * @param {{ phaseId?: string, phaseText?: string, rightsText?: string }} opts
 */
export function renderItSetupStatusChipsHtml(opts) {
    const o = opts || {};
    const phaseId = o.phaseId ? ' id="' + escapeAttr(o.phaseId) + '"' : '';
    const phase = o.phaseText != null ? String(o.phaseText) : 'Schritt …';
    const rights =
        o.rightsText != null
            ? String(o.rightsText)
            : 'Schreibrechte + AllSites.FullControl für Rechte';
    return (
        '<div class="tm-hero__actions it-setup-status" style="margin:0 0 14px;" aria-label="Status">' +
        '<span class="tm-perm"><i class="bi bi-signpost-2"></i><span' +
        phaseId +
        '>' +
        escapeHtml(phase) +
        '</span></span>' +
        '<span class="tm-perm"><i class="bi bi-shield-check"></i>' +
        escapeHtml(rights) +
        '</span></div>'
    );
}

/**
 * @param {{
 *   id: string,
 *   stepAttr: string,
 *   steps: Array<{ n: number, label: string, icon?: string, panelId?: string, value?: string }>,
 *   cols?: number,
 *   ariaLabel?: string
 * }} opts
 */
export function renderItSetupStepperHtml(opts) {
    const o = opts || {};
    const cols = Math.max(2, Number(o.cols) || (o.steps && o.steps.length) || 4);
    const attr = String(o.stepAttr || 'data-it-setup-step');
    const cards = (o.steps || [])
        .map((st, idx) => {
            const n = st.n != null ? st.n : idx + 1;
            const active = idx === 0 ? ' is-active' : '';
            const ariaSel = idx === 0 ? 'true' : 'false';
            const panel = st.panelId ? ' aria-controls="' + escapeAttr(st.panelId) + '"' : '';
            const icon = st.icon || 'bi-circle';
            const value = st.value
                ? '<span class="tm-glance__value" data-it-setup-glance-value>' +
                  escapeHtml(st.value) +
                  '</span>'
                : '';
            return (
                '<button type="button" role="tab" class="tm-glance__card' +
                active +
                '" ' +
                attr +
                '="' +
                n +
                '" aria-selected="' +
                ariaSel +
                '"' +
                panel +
                '>' +
                '<i class="bi ' +
                escapeAttr(icon) +
                '" aria-hidden="true"></i>' +
                '<span class="tm-glance__label">' +
                escapeHtml(String(n) + ' · ' + String(st.label || '')) +
                '</span>' +
                value +
                '</button>'
            );
        })
        .join('');
    return (
        '<section class="tm-glance tm-glance--compact it-setup-glance it-setup-glance--nav it-setup-glance--cols-' +
        cols +
        '" id="' +
        escapeAttr(o.id || 'itSetupGlance') +
        '" role="tablist" aria-label="' +
        escapeAttr(o.ariaLabel || 'Setup-Schritte') +
        '">' +
        cards +
        '</section>'
    );
}

/**
 * @param {{ id?: string, open?: boolean, logId?: string }} opts
 */
export function renderItSetupProtocolHtml(opts) {
    const o = opts || {};
    const open = o.open ? ' open' : '';
    const id = o.id ? ' id="' + escapeAttr(o.id) + '"' : '';
    const logId = o.logId || 'itSetupLog';
    return (
        '<details class="tm-tech it-setup-protocol"' +
        id +
        open +
        '>' +
        '<summary>Protokoll</summary>' +
        '<div id="' +
        escapeAttr(logId) +
        '" class="tm-log" style="margin-top:10px;" aria-live="polite"></div>' +
        '</details>'
    );
}

/**
 * @param {Array<{ label: string, value: string }|string>} items
 */
export function renderItSetupSummaryListHtml(items) {
    const lis = (Array.isArray(items) ? items : [])
        .map((it) => {
            if (typeof it === 'string') return '<li>' + escapeHtml(it) + '</li>';
            return (
                '<li><strong>' +
                escapeHtml(it.label || '') +
                ':</strong> ' +
                escapeHtml(it.value || '') +
                '</li>'
            );
        })
        .join('');
    return '<ul class="it-setup-summary-list">' + lis + '</ul>';
}

/**
 * Wizard-Controller: Stepper + Footer Zurück/Weiter + localStorage.
 *
 * @param {{
 *   stepCount: number,
 *   labels: string[],
 *   storageKey?: string,
 *   glanceSelector: string,
 *   stepAttr: string,
 *   panelIdPrefix: string,
 *   phaseHintId?: string,
 *   backId?: string,
 *   nextId?: string,
 *   validateStep?: (step: number) => string|null|undefined,
 *   onShowStep?: (step: number) => void,
 *   nextLabelForStep?: (step: number) => string|null
 * }} opts
 */
export function createItSetupWizardController(opts) {
    const o = opts || {};
    const count = Math.max(1, Number(o.stepCount) || 1);
    const labels = Array.isArray(o.labels) ? o.labels : [];
    const storageKey = o.storageKey || '';
    const stepAttr = o.stepAttr || 'data-it-setup-step';
    const panelPrefix = o.panelIdPrefix || 'itSetupStep';

    function loadStep() {
        if (!storageKey) return 1;
        try {
            return clampItSetupStep(localStorage.getItem(storageKey), count);
        } catch {
            return 1;
        }
    }

    function saveStep(step) {
        if (!storageKey) return;
        try {
            localStorage.setItem(storageKey, String(clampItSetupStep(step, count)));
        } catch {
            /* ignore */
        }
    }

    function $(id) {
        return typeof document !== 'undefined' ? document.getElementById(id) : null;
    }

    function showStep(n) {
        const step = clampItSetupStep(n, count);
        saveStep(step);
        for (let i = 1; i <= count; i++) {
            const panel = $(panelPrefix + i);
            if (panel) panel.hidden = i !== step;
        }
        const glanceRoot =
            typeof document !== 'undefined' ? document.querySelector(o.glanceSelector) : null;
        if (glanceRoot) {
            glanceRoot.querySelectorAll('[' + stepAttr + ']').forEach((btn) => {
                const s = parseInt(btn.getAttribute(stepAttr) || '0', 10);
                const active = s === step;
                btn.classList.toggle('is-active', active);
                btn.setAttribute('aria-selected', active ? 'true' : 'false');
            });
        }
        if (o.phaseHintId) {
            const hint = $(o.phaseHintId);
            if (hint) hint.textContent = formatItSetupPhaseHint(step, count, labels);
        }
        const back = o.backId ? $(o.backId) : null;
        const next = o.nextId ? $(o.nextId) : null;
        if (back) back.disabled = step <= 1;
        if (next) {
            next.hidden = step >= count;
            if (step < count) {
                const custom =
                    typeof o.nextLabelForStep === 'function' ? o.nextLabelForStep(step) : null;
                const label =
                    custom ||
                    (step === count - 1 ? 'Zur Übersicht' : 'Weiter');
                next.innerHTML = escapeHtml(label) + ' <i class="bi bi-arrow-right"></i>';
            }
        }
        if (typeof o.onShowStep === 'function') o.onShowStep(step);
        return step;
    }

    function tryAdvance(fromStep) {
        if (typeof o.validateStep !== 'function') return true;
        const err = o.validateStep(fromStep);
        if (err) {
            if (typeof window !== 'undefined' && typeof window.ms365ToastOrAlert === 'function') {
                window.ms365ToastOrAlert(String(err));
            } else if (typeof window !== 'undefined') {
                window.alert(String(err));
            }
            return false;
        }
        return true;
    }

    function wire() {
        if (typeof document === 'undefined') return;
        const glanceRoot = document.querySelector(o.glanceSelector);
        if (glanceRoot) {
            glanceRoot.querySelectorAll('[' + stepAttr + ']').forEach((btn) => {
                btn.addEventListener('click', () => {
                    const target = parseInt(btn.getAttribute(stepAttr) || '1', 10);
                    const cur = loadStep();
                    if (target > cur) {
                        for (let s = cur; s < target; s++) {
                            if (!tryAdvance(s)) {
                                showStep(s);
                                return;
                            }
                        }
                    }
                    showStep(target);
                });
            });
        }
        const back = o.backId ? $(o.backId) : null;
        const next = o.nextId ? $(o.nextId) : null;
        if (back) {
            back.addEventListener('click', () => showStep(loadStep() - 1));
        }
        if (next) {
            next.addEventListener('click', () => {
                const cur = loadStep();
                if (!tryAdvance(cur)) return;
                showStep(cur + 1);
            });
        }
        showStep(loadStep());
    }

    return {
        loadStep,
        saveStep,
        showStep,
        tryAdvance,
        wire,
        clampStep: (n) => clampItSetupStep(n, count)
    };
}

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function escapeAttr(s) {
    return escapeHtml(s).replace(/'/g, '&#39;');
}
