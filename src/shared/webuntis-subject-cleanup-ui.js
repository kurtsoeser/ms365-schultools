/**
 * Fächer bereinigen (WebUntis-Import) – UI wie Kursteams-Schritt „Fächer ausschließen“.
 * Erwartet window.ms365KursteamSubjectFilterLogic (kursteam-subject-filter-logic.js).
 */

function ks() {
    return window.ms365KursteamSubjectFilterLogic;
}

/**
 * @param {object} opts
 * @param {() => { code: string, selected?: boolean }[]} opts.getSubjectRows
 * @param {(excludedCodes: string[]) => void} opts.onExcludedChange
 */
export function mountWebuntisSubjectCleanup(opts) {
    const getRows = opts && opts.getSubjectRows;
    const onExcluded = opts && opts.onExcludedChange;
    const list = document.getElementById('wuSubjectFilterList');
    const summary = document.getElementById('wuSubjectFilterSummary');
    const search = document.getElementById('wuSubjectFilterSearch');
    const host = document.getElementById('wuStammdatenSubjectCleanup');
    if (!list || !getRows || !onExcluded) return { refresh: function () {} };

    const KS = ks();
    if (!KS) return { refresh: function () {} };

    let excluded = new Set();

    function availableCodes() {
        return KS.uniqSortedSubjectTokens(
            (getRows() || []).map(function (r) {
                return r.code;
            })
        );
    }

    function emitExcluded() {
        onExcluded(Array.from(excluded));
    }

    function setExcluded(tokens) {
        excluded = new Set(KS.uniqSortedSubjectTokens(tokens));
        emitExcluded();
    }

    function updateSummary(available, familyCount) {
        if (summary) summary.textContent = KS.subjectFilterSummaryText(available.length, excluded.size, familyCount);
    }

    function applySearch(query) {
        const q = KS.normalizeSubjectToken(query);
        list.querySelectorAll('[data-subject-family], [data-subject]').forEach(function (node) {
            if (node.hasAttribute('data-subject-family')) {
                const base = String(node.getAttribute('data-subject-family') || '');
                const variants = String(node.getAttribute('data-variants') || '');
                const hit = !q || base.includes(q) || variants.includes(q);
                node.style.display = hit ? '' : 'none';
                return;
            }
            const subj = String(node.getAttribute('data-subject') || '');
            node.style.display = !q || subj.includes(q) ? '' : 'none';
        });
    }

    function setFamilyCheckboxState(cb, variants) {
        const n = variants.filter(function (v) {
            return excluded.has(v);
        }).length;
        cb.checked = n === variants.length && variants.length > 0;
        cb.indeterminate = n > 0 && n < variants.length;
    }

    function syncExcludedFromSelection() {
        excluded = new Set();
        (getRows() || []).forEach(function (r) {
            const code = KS.normalizeSubjectToken(r && r.code);
            if (code && !r.selected) excluded.add(code);
        });
    }

    function refresh() {
        syncExcludedFromSelection();
        const available = availableCodes();
        const groups = KS.groupSubjectsByBase(available);
        const familyCount = groups.filter(function (g) {
            return g.isFamily;
        }).length;

        list.replaceChildren();

        groups.forEach(function (group) {
            if (!group.isFamily) {
                const subj = group.variants[0];
                const label = document.createElement('label');
                label.className = 'subject-filter-item';
                label.setAttribute('data-subject', subj);
                const cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.checked = excluded.has(subj);
                cb.title = 'Fach vom Import ausschließen';
                cb.addEventListener('change', function () {
                    if (cb.checked) excluded.add(subj);
                    else excluded.delete(subj);
                    emitExcluded();
                    updateSummary(available, familyCount);
                });
                const text = document.createElement('code');
                text.textContent = subj;
                label.append(cb, text);
                list.appendChild(label);
                return;
            }

            const wrap = document.createElement('div');
            wrap.className = 'subject-filter-family';
            wrap.setAttribute('data-subject-family', group.base);
            wrap.setAttribute('data-variants', group.variants.join(','));

            const head = document.createElement('label');
            head.className = 'subject-filter-item subject-filter-family-head';
            const familyCb = document.createElement('input');
            familyCb.type = 'checkbox';
            familyCb.title = 'Gesamte Familie ausschließen';
            setFamilyCheckboxState(familyCb, group.variants);
            familyCb.addEventListener('change', function () {
                if (familyCb.checked) group.variants.forEach(function (v) {
                    excluded.add(v);
                });
                else
                    group.variants.forEach(function (v) {
                        excluded.delete(v);
                    });
                emitExcluded();
                refresh();
            });
            const headCode = document.createElement('code');
            headCode.textContent = group.base;
            const headMeta = document.createElement('span');
            headMeta.className = 'subject-filter-family-meta';
            headMeta.textContent = ' Familie (' + group.variants.length + ')';
            head.append(familyCb, headCode, headMeta);
            wrap.appendChild(head);

            const variants = document.createElement('div');
            variants.className = 'subject-filter-family-variants';
            group.variants.forEach(function (subj) {
                const label = document.createElement('label');
                label.className = 'subject-filter-item subject-filter-item--variant';
                label.setAttribute('data-subject', subj);
                const cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.checked = excluded.has(subj);
                cb.addEventListener('change', function () {
                    if (cb.checked) excluded.add(subj);
                    else excluded.delete(subj);
                    emitExcluded();
                    updateSummary(available, familyCount);
                    setFamilyCheckboxState(familyCb, group.variants);
                });
                const text = document.createElement('code');
                text.textContent = subj;
                label.append(cb, text);
                variants.appendChild(label);
            });
            wrap.appendChild(variants);
            list.appendChild(wrap);
        });

        updateSummary(available, familyCount);
        if (search && search.value) applySearch(search.value);
        if (host) host.hidden = available.length === 0;
    }

    if (!mountWebuntisSubjectCleanup._wired) {
        mountWebuntisSubjectCleanup._wired = true;
        if (search) search.addEventListener('input', function () {
            applySearch(search.value);
        });
        const btnNone = document.getElementById('wuSubjectFilterExcludeNone');
        if (btnNone)
            btnNone.addEventListener('click', function () {
                setExcluded([]);
                refresh();
            });
        const btnAll = document.getElementById('wuSubjectFilterExcludeAll');
        if (btnAll)
            btnAll.addEventListener('click', function () {
                setExcluded(availableCodes());
                refresh();
            });
        const btnDefault = document.getElementById('wuSubjectFilterResetDefault');
        if (btnDefault)
            btnDefault.addEventListener('click', function () {
                setExcluded(['ADM', 'DIR', 'KUST', 'BFK', 'BIB', 'AUFSICHT', 'ORD', 'KV']);
                refresh();
            });
        const btnNumbered = document.getElementById('wuSubjectFilterExcludeNumbered');
        if (btnNumbered)
            btnNumbered.addEventListener('click', function () {
                const numbered = availableCodes().filter(function (s) {
                    return !!KS.splitSubjectBaseAndSuffix(s).suffix;
                });
                const next = new Set(excluded);
                numbered.forEach(function (s) {
                    next.add(s);
                });
                setExcluded(Array.from(next));
                refresh();
            });
        const btnKeepBase = document.getElementById('wuSubjectFilterKeepFamilyBase');
        if (btnKeepBase)
            btnKeepBase.addEventListener('click', function () {
                const available = availableCodes();
                const groups = KS.groupSubjectsByBase(available);
                const next = new Set(excluded);
                groups.forEach(function (g) {
                    if (!g.isFamily) return;
                    g.variants.forEach(function (v) {
                        if (v !== g.base) next.add(v);
                    });
                });
                setExcluded(Array.from(next));
                refresh();
            });
    }

    return {
        refresh: refresh,
        setExcluded: setExcluded,
        getExcluded: function () {
            return Array.from(excluded);
        }
    };
}
