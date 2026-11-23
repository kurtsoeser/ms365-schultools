/**
 * Gemeinsame Playbook-Persistenz (Cleanup-Pattern).
 * Keys: ms365-<id>-playbook-v1
 *
 * Wave 0: Hybrid Checkbox + optional Ready-Checks (playbook-ready.js).
 * Deep-Links: playbook-deeplink.js
 */
import { notifyAppLocalDataChanged } from './app-local-data-notify.js';
import { formatReadyHint, normalizeReadyResult, readyForStep } from './playbook-ready.js';

export function loadPlaybookState(storageKey, storage) {
    const store = storage || (typeof localStorage !== 'undefined' ? localStorage : null);
    if (!store) return {};
    try {
        return JSON.parse(store.getItem(storageKey) || '{}') || {};
    } catch {
        return {};
    }
}

export function savePlaybookState(storageKey, state) {
    try {
        localStorage.setItem(storageKey, JSON.stringify(state && typeof state === 'object' ? state : {}));
    } catch {
        /* ignore */
    }
    notifyAppLocalDataChanged('playbook', { key: storageKey });
}

/**
 * Bindet Checkboxen [data-cp-check] + Fortschritt #cpProgress + Reset #cpReset.
 * Optional: evaluateGates / evaluateStepUnlock (Unlock) und evaluateReadyChecks (Hybrid-Hinweis).
 *
 * DOM-Hinweise:
 * - [data-cp-ready-hint] → Ready-Check-Text (bevorzugt)
 * - [data-cp-gate-hint] → Unlock-Warnung; Fallback für Ready, wenn kein ready-hint
 *
 * @param {{
 *   storageKey: string,
 *   stepIds: string[],
 *   gateMode?: 'warn'|'block',
 *   evaluateGates?: () => Record<string, { ok: boolean, message: string }>,
 *   evaluateStepUnlock?: (stepId: string, index: number, state: Record<string, boolean>, gates: Record<string, { ok: boolean, message: string }>) => { unlocked: boolean, message?: string },
 *   evaluateReadyChecks?: () => Record<string, { status?: string, ok?: boolean, message?: string }>
 * }} opts
 */
export function wirePlaybookPage(opts) {
    const key = opts.storageKey;
    const ids = Array.isArray(opts.stepIds) ? opts.stepIds : [];
    const gateMode = opts.gateMode === 'block' ? 'block' : 'warn';
    const evaluateGates = typeof opts.evaluateGates === 'function' ? opts.evaluateGates : null;
    const evaluateStepUnlock =
        typeof opts.evaluateStepUnlock === 'function' ? opts.evaluateStepUnlock : null;
    const evaluateReadyChecks =
        typeof opts.evaluateReadyChecks === 'function' ? opts.evaluateReadyChecks : null;

    function stepArticle(id) {
        return (
            document.querySelector('[data-cp-step="' + id + '"]') ||
            (function () {
                const chk = document.querySelector('[data-cp-check="' + id + '"]');
                return chk ? chk.closest('.pb-step, .cp-step') : null;
            })()
        );
    }

    function setHint(el, text) {
        if (!el) return;
        if (text) {
            el.hidden = false;
            el.textContent = text;
        } else {
            el.hidden = true;
            el.textContent = '';
        }
    }

    function applyReadyChecks() {
        if (!evaluateReadyChecks) return;
        let map = {};
        try {
            map = evaluateReadyChecks() || {};
        } catch {
            map = {};
        }
        ids.forEach(function (id) {
            const article = stepArticle(id);
            if (!article) return;
            const result = readyForStep(map, id);
            ['ok', 'warn', 'missing', 'unknown'].forEach(function (s) {
                article.classList.remove('cp-step--ready-' + s);
                article.classList.remove('pb-step--ready-' + s);
            });
            article.classList.add('cp-step--ready-' + result.status);
            article.classList.add('pb-step--ready-' + result.status);
            article.setAttribute('data-cp-ready', result.status);

            const readyHint = article.querySelector('[data-cp-ready-hint]');
            const gateHint = article.querySelector('[data-cp-gate-hint]');
            const text = formatReadyHint(result);
            const show = !!(text && result.status !== 'unknown');
            if (readyHint) {
                setHint(readyHint, show ? text : '');
                if (show) readyHint.setAttribute('data-cp-ready-status', result.status);
                else readyHint.removeAttribute('data-cp-ready-status');
            } else if (gateHint && !evaluateGates) {
                setHint(gateHint, show ? text : '');
            }
        });
    }

    function applyGates(state) {
        if (!evaluateGates) {
            applyReadyChecks();
            return;
        }
        const gates = evaluateGates() || {};
        ids.forEach(function (id, index) {
            const article = stepArticle(id);
            const chk = document.querySelector('[data-cp-check="' + id + '"]');
            const gateHint = article ? article.querySelector('[data-cp-gate-hint]') : null;
            let unlock = { unlocked: true, message: '' };
            if (evaluateStepUnlock) {
                unlock = evaluateStepUnlock(id, index, state, gates) || unlock;
            } else {
                const g = gates[id];
                if (g && !g.ok) unlock = { unlocked: false, message: g.message };
            }
            const gated = !unlock.unlocked;
            const block = gateMode === 'block' && gated;
            if (article) {
                article.classList.toggle('cp-step--locked', block);
                article.classList.toggle('pb-step--locked', block);
                article.classList.toggle('cp-step--gate-warn', gateMode === 'warn' && gated);
                article.classList.toggle('pb-step--gate-warn', gateMode === 'warn' && gated);
                article.setAttribute('aria-disabled', block ? 'true' : 'false');
            }
            if (chk) {
                chk.disabled = block;
                if (block && chk.checked) {
                    chk.checked = false;
                    state[id] = false;
                    savePlaybookState(key, state);
                }
            }
            if (gateHint) {
                if (gated && unlock.message) {
                    setHint(
                        gateHint,
                        gateMode === 'warn' ? 'Hinweis: ' + unlock.message : unlock.message
                    );
                } else {
                    setHint(gateHint, '');
                }
            }
        });
        applyReadyChecks();
    }

    function refresh() {
        const state = loadPlaybookState(key);
        let done = 0;
        ids.forEach(function (id) {
            const el = document.querySelector('[data-cp-check="' + id + '"]');
            if (!el) return;
            el.checked = !!state[id];
            if (el.checked) done++;
        });
        const p = document.getElementById('cpProgress');
        if (p) p.textContent = done + ' / ' + ids.length + ' erledigt';
        applyGates(state);
    }

    document.querySelectorAll('[data-cp-check]').forEach(function (el) {
        el.addEventListener('change', function () {
            if (el.disabled) {
                el.checked = false;
                return;
            }
            const state = loadPlaybookState(key);
            state[el.getAttribute('data-cp-check')] = !!el.checked;
            savePlaybookState(key, state);
            refresh();
        });
    });

    const reset = document.getElementById('cpReset');
    if (reset) {
        reset.addEventListener('click', function () {
            const ok =
                typeof window.ms365AppDialogConfirm === 'function'
                    ? null
                    : window.confirm('Alle Haken zurücksetzen?');
            function doReset() {
                savePlaybookState(key, {});
                refresh();
            }
            if (typeof window.ms365AppDialogConfirm === 'function') {
                window.ms365AppDialogConfirm('Alle Haken zurücksetzen?', { title: 'Zurücksetzen', danger: true }).then(
                    function (yes) {
                        if (yes) doReset();
                    }
                );
            } else if (ok) doReset();
        });
    }
    refresh();

    window.addEventListener('ms365-app-local-data-changed', function () {
        refresh();
    });
}

if (typeof window !== 'undefined') {
    window.ms365Playbook = {
        loadPlaybookState: loadPlaybookState,
        savePlaybookState: savePlaybookState,
        wirePlaybookPage: wirePlaybookPage,
        normalizeReadyResult: normalizeReadyResult
    };
}
