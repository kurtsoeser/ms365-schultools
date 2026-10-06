/**
 * Gemeinsame Playbook-Persistenz (Cleanup-Pattern).
 * Keys: ms365-<id>-playbook-v1
 */
import { notifyAppLocalDataChanged } from './app-local-data-notify.js';

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
 * @param {{ storageKey: string, stepIds: string[], gateMode?: 'warn'|'block', evaluateGates?: () => Record<string, { ok: boolean, message: string }>, evaluateStepUnlock?: (stepId: string, index: number, state: Record<string, boolean>, gates: Record<string, { ok: boolean, message: string }>) => { unlocked: boolean, message?: string } }} opts
 */
export function wirePlaybookPage(opts) {
    const key = opts.storageKey;
    const ids = Array.isArray(opts.stepIds) ? opts.stepIds : [];
    const gateMode = opts.gateMode === 'block' ? 'block' : 'warn';
    const evaluateGates = typeof opts.evaluateGates === 'function' ? opts.evaluateGates : null;
    const evaluateStepUnlock =
        typeof opts.evaluateStepUnlock === 'function' ? opts.evaluateStepUnlock : null;

    function stepArticle(id) {
        return (
            document.querySelector('[data-cp-step="' + id + '"]') ||
            (function () {
                const chk = document.querySelector('[data-cp-check="' + id + '"]');
                return chk ? chk.closest('.cp-step') : null;
            })()
        );
    }

    function applyGates(state) {
        if (!evaluateGates) return;
        const gates = evaluateGates() || {};
        ids.forEach(function (id, index) {
            const article = stepArticle(id);
            const chk = document.querySelector('[data-cp-check="' + id + '"]');
            const hint = article ? article.querySelector('[data-cp-gate-hint]') : null;
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
                article.classList.toggle('cp-step--gate-warn', gateMode === 'warn' && gated);
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
            if (hint) {
                if (gated && unlock.message) {
                    hint.hidden = false;
                    hint.textContent =
                        gateMode === 'warn' ? 'Hinweis: ' + unlock.message : unlock.message;
                } else {
                    hint.hidden = true;
                    hint.textContent = '';
                }
            }
        });
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
        wirePlaybookPage: wirePlaybookPage
    };
}
