/**
 * Gemeinsame Playbook-Persistenz (Cleanup-Pattern).
 * Keys: ms365-<id>-playbook-v1
 */
export function loadPlaybookState(storageKey) {
    try {
        return JSON.parse(localStorage.getItem(storageKey) || '{}') || {};
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
}

/**
 * Bindet Checkboxen [data-cp-check] + Fortschritt #cpProgress + Reset #cpReset.
 * @param {{ storageKey: string, stepIds: string[] }} opts
 */
export function wirePlaybookPage(opts) {
    const key = opts.storageKey;
    const ids = Array.isArray(opts.stepIds) ? opts.stepIds : [];

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
    }

    document.querySelectorAll('[data-cp-check]').forEach(function (el) {
        el.addEventListener('change', function () {
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
}

if (typeof window !== 'undefined') {
    window.ms365Playbook = {
        loadPlaybookState: loadPlaybookState,
        savePlaybookState: savePlaybookState,
        wirePlaybookPage: wirePlaybookPage
    };
}
