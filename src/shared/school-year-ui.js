/**
 * Globales Schuljahr: Select + „Neues Schuljahr“ (User-Menü, einmal pro Jahr relevant).
 */

function currentSchoolYearLabel() {
    try {
        if (window.ms365SchoolYear && typeof window.ms365SchoolYear.currentSchoolYearLabel === 'function') {
            return window.ms365SchoolYear.currentSchoolYearLabel();
        }
    } catch {
        /* ignore */
    }
    const y = new Date().getFullYear();
    return String(y) + '/' + String(y + 1).slice(2);
}

export function setCurrentSchoolYearInV2(nextLabel, opts) {
    try {
        if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.setCurrentYear !== 'function') return false;
        window.ms365AppDataV2.setCurrentYear(String(nextLabel || '').trim(), opts || {});
        return true;
    } catch {
        return false;
    }
}

export function renderSchoolYearSelect(selectEl) {
    if (!selectEl) return;
    try {
        if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.getContainer !== 'function') {
            selectEl.replaceChildren();
            const o = document.createElement('option');
            o.value = currentSchoolYearLabel();
            o.textContent = currentSchoolYearLabel();
            selectEl.appendChild(o);
            selectEl.value = o.value;
            return;
        }
        const c = window.ms365AppDataV2.getContainer();
        const cur = c && c.years ? String(c.years.current || '') : '';
        const years = typeof window.ms365AppDataV2.listYears === 'function' ? window.ms365AppDataV2.listYears() : [];
        const list = years.length ? years : cur ? [cur] : [currentSchoolYearLabel()];
        selectEl.replaceChildren();
        list.forEach(function (y) {
            const opt = document.createElement('option');
            opt.value = String(y);
            opt.textContent = String(y);
            selectEl.appendChild(opt);
        });
        selectEl.value = cur && list.includes(cur) ? cur : list[0];
    } catch {
        /* ignore */
    }
}

function dispatchSchoolYearChanged(year, reason) {
    try {
        window.dispatchEvent(
            new CustomEvent('ms365-school-year-changed', {
                detail: { year: String(year || '').trim(), reason: reason || 'switch' }
            })
        );
    } catch {
        /* ignore */
    }
    try {
        if (typeof window.ms365RefreshContextBar === 'function') window.ms365RefreshContextBar();
    } catch {
        /* ignore */
    }
}

async function dlgPrompt(msg, def, opts) {
    if (typeof window.ms365AppDialogPrompt === 'function') {
        return window.ms365AppDialogPrompt(msg, def, opts);
    }
    return window.prompt(msg, def);
}

async function dlgConfirm(msg, opts) {
    if (typeof window.ms365AppDialogConfirm === 'function') {
        return window.ms365AppDialogConfirm(msg, opts);
    }
    return window.confirm(msg);
}

function suggestNextYearLabel(cur) {
    const m = String(cur || '').match(/^(\d{4})\s*\/\s*(\d{2}|\d{4})/);
    if (!m) return '';
    const y = parseInt(m[1], 10);
    if (!isFinite(y)) return '';
    return String(y + 1) + '/' + String(y + 2).slice(2);
}

export function bindSchoolYearControls() {
    const selectEl = document.getElementById('schoolYearSelect');
    const addBtn = document.getElementById('schoolYearAddBtn');

    if (selectEl && selectEl.dataset.schoolYearBound !== '1') {
        selectEl.dataset.schoolYearBound = '1';
        selectEl.addEventListener('change', function () {
            const y = String(selectEl.value || '').trim();
            if (!y) return;
            setCurrentSchoolYearInV2(y);
            dispatchSchoolYearChanged(y, 'switch');
        });
    }

    if (addBtn && addBtn.dataset.schoolYearBound !== '1') {
        addBtn.dataset.schoolYearBound = '1';
        addBtn.addEventListener('click', function () {
            void (async function () {
                const cur = selectEl ? String(selectEl.value || '').trim() : '';
                const suggest = suggestNextYearLabel(cur);
                const next = await dlgPrompt('Neues Schuljahr (z. B. 2027/28)', suggest || currentSchoolYearLabel(), {
                    title: 'Schuljahr',
                    inputLabel: 'Bezeichnung'
                });
                if (next == null || !String(next).trim()) return;
                const copy = await dlgConfirm('Schüler & Klassen aus dem aktuellen Schuljahr übernehmen?', {
                    title: 'Schuljahr',
                    okText: 'Ja, übernehmen',
                    cancelText: 'Nein'
                });
                setCurrentSchoolYearInV2(String(next).trim(), copy && cur ? { copyFrom: cur } : {});
                renderSchoolYearSelect(selectEl);
                dispatchSchoolYearChanged(String(next).trim(), 'added');
            })();
        });
    }

    if (selectEl) renderSchoolYearSelect(selectEl);
}

if (typeof window !== 'undefined') {
    window.ms365SchoolYearUi = {
        renderSchoolYearSelect,
        setCurrentSchoolYearInV2,
        bindSchoolYearControls
    };
}
