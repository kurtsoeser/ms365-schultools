const KS = window.ms365KursteamSubjectFilterLogic;
window.ms365AssertModules({ KS }, 'kursteam-subject-filter-ui.js');

/**
 * Fachfilter-Liste, Suche, Schnellbuttons (Schritt Bereinigung).
 * Nummerierte Fächer (OMAI1/2/3) werden als Familien gruppiert.
 * @param {object} ns window.ms365Kursteam
 */
function mount(ns) {
    ns.parseExcludeSubjectsFromInput = function parseExcludeSubjectsFromInput() {
        const el = document.getElementById('excludeSubjects');
        if (!el) return [];
        return KS.parseExcludeSubjectsFromString(el.value);
    };

    ns.setExcludeSubjectsInput = function setExcludeSubjectsInput(tokens) {
        const el = document.getElementById('excludeSubjects');
        if (!el) return;
        el.value = KS.uniqSortedSubjectTokens(tokens).join(',');
    };

    function collectAvailableSubjectsFromRawData() {
        return KS.collectSubjectsFromRows(ns.rawData);
    }

    function updateSubjectFilterSummary(available, excluded, familyCount) {
        const el = document.getElementById('subjectFilterSummary');
        if (!el) return;
        el.textContent = KS.subjectFilterSummaryText(available.length, excluded.length, familyCount);
    }

    function applySearchToSubjectList(query) {
        const q = KS.normalizeSubjectToken(query);
        const list = document.getElementById('subjectFilterList');
        if (!list) return;
        Array.from(list.querySelectorAll('[data-subject-family], [data-subject]')).forEach((node) => {
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

    function setFamilyCheckboxState(cb, variants, excluded) {
        const n = variants.filter((v) => excluded.has(v)).length;
        cb.checked = n === variants.length && variants.length > 0;
        cb.indeterminate = n > 0 && n < variants.length;
    }

    function wireSubjectFilterEventsOnce() {
        if (wireSubjectFilterEventsOnce._wired) return;
        wireSubjectFilterEventsOnce._wired = true;

        const search = document.getElementById('subjectFilterSearch');
        if (search) {
            search.addEventListener('input', () => applySearchToSubjectList(search.value));
        }

        const btnNone = document.getElementById('subjectFilterExcludeNone');
        if (btnNone) {
            btnNone.addEventListener('click', () => {
                ns.setExcludeSubjectsInput([]);
                ns.refreshSubjectFilterUI();
            });
        }

        const btnAll = document.getElementById('subjectFilterExcludeAll');
        if (btnAll) {
            btnAll.addEventListener('click', () => {
                ns.setExcludeSubjectsInput(collectAvailableSubjectsFromRawData());
                ns.refreshSubjectFilterUI();
            });
        }

        const btnDefault = document.getElementById('subjectFilterResetDefault');
        if (btnDefault) {
            btnDefault.addEventListener('click', () => {
                ns.setExcludeSubjectsInput(['ORD', 'DIR', 'KV']);
                ns.refreshSubjectFilterUI();
            });
        }

        const btnNumbered = document.getElementById('subjectFilterExcludeNumbered');
        if (btnNumbered) {
            btnNumbered.addEventListener('click', () => {
                const available = collectAvailableSubjectsFromRawData();
                const numbered = available.filter((s) => !!KS.splitSubjectBaseAndSuffix(s).suffix);
                const current = new Set(ns.parseExcludeSubjectsFromInput());
                numbered.forEach((s) => current.add(s));
                ns.setExcludeSubjectsInput(Array.from(current));
                ns.refreshSubjectFilterUI();
            });
        }

        const input = document.getElementById('excludeSubjects');
        if (input) {
            input.addEventListener('input', () => ns.refreshSubjectFilterUI());
        }
    }

    ns.refreshSubjectFilterUI = function refreshSubjectFilterUI() {
        const list = document.getElementById('subjectFilterList');
        const search = document.getElementById('subjectFilterSearch');
        if (!list) return;

        wireSubjectFilterEventsOnce();

        const available = collectAvailableSubjectsFromRawData();
        const excluded = new Set(ns.parseExcludeSubjectsFromInput());
        const groups = KS.groupSubjectsByBase(available);
        const familyCount = groups.filter((g) => g.isFamily).length;

        list.replaceChildren();

        groups.forEach((group) => {
            if (!group.isFamily) {
                const subj = group.variants[0];
                const label = document.createElement('label');
                label.className = 'subject-filter-item';
                label.setAttribute('data-subject', subj);

                const cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.checked = excluded.has(subj);
                cb.addEventListener('change', () => {
                    const current = new Set(ns.parseExcludeSubjectsFromInput());
                    if (cb.checked) current.add(subj);
                    else current.delete(subj);
                    ns.setExcludeSubjectsInput(Array.from(current));
                    updateSubjectFilterSummary(available, Array.from(current), familyCount);
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
            familyCb.title = 'Gesamte Fach-Familie ausschließen';
            setFamilyCheckboxState(familyCb, group.variants, excluded);
            familyCb.addEventListener('change', () => {
                const current = new Set(ns.parseExcludeSubjectsFromInput());
                if (familyCb.checked) group.variants.forEach((v) => current.add(v));
                else group.variants.forEach((v) => current.delete(v));
                ns.setExcludeSubjectsInput(Array.from(current));
                ns.refreshSubjectFilterUI();
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

            group.variants.forEach((subj) => {
                const label = document.createElement('label');
                label.className = 'subject-filter-item subject-filter-item--variant';
                label.setAttribute('data-subject', subj);

                const cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.checked = excluded.has(subj);
                cb.addEventListener('change', () => {
                    const current = new Set(ns.parseExcludeSubjectsFromInput());
                    if (cb.checked) current.add(subj);
                    else current.delete(subj);
                    ns.setExcludeSubjectsInput(Array.from(current));
                    ns.refreshSubjectFilterUI();
                });

                const text = document.createElement('code');
                text.textContent = subj;

                label.append(cb, text);
                variants.appendChild(label);
            });

            wrap.appendChild(variants);
            list.appendChild(wrap);
        });

        updateSubjectFilterSummary(available, Array.from(excluded), familyCount);
        if (search) applySearchToSubjectList(search.value);
    };
}

window.ms365KursteamSubjectFilterUI = {
    mount
};
