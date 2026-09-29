/**
 * Einheitliches Bulk-Fortschritts-UI (Analyse 03 UX-2).
 *
 * Markup (einmal im Tool):
 * <div id="…" data-ms365-bulk-progress hidden>
 *   <div class="ms365-bulk-progress__bar"><div data-fill></div></div>
 *   <p data-text></p>
 *   <button type="button" data-cancel>Abbrechen</button>
 *   <ul data-errors></ul>
 * </div>
 */
export function createBulkProgress(rootOrId) {
    const root =
        typeof rootOrId === 'string' ? document.getElementById(rootOrId) : rootOrId;
    if (!root) {
        return {
            show: function () {},
            hide: function () {},
            set: function () {},
            addError: function () {},
            clearErrors: function () {},
            isCancelled: function () {
                return false;
            },
            resetCancel: function () {},
            exportErrorsCsv: function () {
                return '';
            }
        };
    }

    const fill = root.querySelector('[data-fill]');
    const text = root.querySelector('[data-text]');
    const errList = root.querySelector('[data-errors]');
    const cancelBtn = root.querySelector('[data-cancel]');
    let cancelled = false;
    let total = 0;
    let done = 0;

    if (cancelBtn) {
        cancelBtn.addEventListener('click', function () {
            cancelled = true;
            if (text) text.textContent = (text.textContent || '') + ' – Abbruch …';
        });
    }

    function setBar() {
        if (!fill) return;
        const pct = total > 0 ? Math.max(0, Math.min(100, Math.round((done / total) * 100))) : 0;
        fill.style.width = pct + '%';
    }

    return {
        show: function (message, t) {
            cancelled = false;
            total = typeof t === 'number' ? t : 0;
            done = 0;
            root.hidden = false;
            if (text) text.textContent = message || '';
            if (errList) errList.replaceChildren();
            setBar();
        },
        hide: function () {
            root.hidden = true;
        },
        set: function (current, t, message) {
            if (typeof t === 'number') total = t;
            done = current;
            if (message && text) text.textContent = message;
            else if (text && total) text.textContent = current + ' / ' + total;
            setBar();
        },
        addError: function (line) {
            if (!errList) return;
            const li = document.createElement('li');
            li.textContent = String(line || '');
            errList.appendChild(li);
        },
        clearErrors: function () {
            if (errList) errList.replaceChildren();
        },
        isCancelled: function () {
            return cancelled;
        },
        resetCancel: function () {
            cancelled = false;
        },
        exportErrorsCsv: function () {
            if (!errList) return '';
            const lines = ['Fehler'];
            errList.querySelectorAll('li').forEach(function (li) {
                const s = String(li.textContent || '').replace(/"/g, '""');
                lines.push('"' + s + '"');
            });
            return '\uFEFF' + lines.join('\n');
        }
    };
}

if (typeof window !== 'undefined') {
    window.ms365BulkProgress = { createBulkProgress: createBulkProgress };
}
