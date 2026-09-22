/**
 * Komfort-Modus: ein oder mehrere Kursteams zeilenweise vorbereiten und anlegen.
 * Einstieg: Dashboard „Einzelne Unterrichtsteams hinzufügen“ oder ?mode=single
 */
(function () {
    'use strict';

    const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

    let rowSeq = 0;

    function loadTenantLists() {
        let settings = null;
        try {
            if (typeof window.ms365TenantSettingsLoad === 'function') {
                settings = window.ms365TenantSettingsLoad();
            }
        } catch {
            settings = null;
        }
        const classes = Array.isArray(settings && settings.classes) ? settings.classes : [];
        const subjects = Array.isArray(settings && settings.subjects) ? settings.subjects : [];
        const teachers = Array.isArray(settings && settings.teachers) ? settings.teachers : [];
        return { classes, subjects, teachers };
    }

    function fillDatalist(listId, values) {
        const list = document.getElementById(listId);
        if (!list) return;
        list.replaceChildren();
        (values || []).forEach((v) => {
            if (!v) return;
            const opt = document.createElement('option');
            opt.value = v;
            list.appendChild(opt);
        });
    }

    ns.refreshSingleTeamDatalists = function refreshSingleTeamDatalists() {
        const { classes, subjects, teachers } = loadTenantLists();
        fillDatalist(
            'singleTeamKlasseList',
            classes
                .map((c) => String((c && (c.code || c.name)) || '').trim())
                .filter(Boolean)
        );
        fillDatalist(
            'singleTeamFachList',
            subjects.map((s) => String((s && s.code) || '').trim()).filter(Boolean)
        );
        fillDatalist(
            'singleTeamLehrerList',
            teachers.map((t) => String((t && t.code) || '').trim()).filter(Boolean)
        );
    };

    ns.resolveSingleTeamOwnerEmail = function resolveSingleTeamOwnerEmail(lehrerCode) {
        const code = String(lehrerCode || '')
            .trim()
            .toUpperCase();
        if (!code) return '';
        if (typeof ns.syncTeacherEmailsFromTenant === 'function') {
            ns.syncTeacherEmailsFromTenant([code], { quiet: true });
        }
        const map = ns.teacherEmailMapping || {};
        if (map[code]) return String(map[code]).trim();
        if (typeof ns.resolveTeacherMatchFromTenant === 'function') {
            const m = ns.resolveTeacherMatchFromTenant(code);
            if (m && m.email) return String(m.email).trim();
        }
        return '';
    };

    function fieldVal(rowEl, field) {
        const inp = rowEl && rowEl.querySelector('[data-st-field="' + field + '"]');
        return inp ? String(inp.value || '').trim() : '';
    }

    function readSingleTeamRowsFromDom() {
        const wrap = document.getElementById('singleTeamRows');
        if (!wrap) return [];
        return Array.from(wrap.querySelectorAll('.single-team-row')).map((rowEl) => ({
            rowEl,
            klasse: fieldVal(rowEl, 'klasse'),
            fach: fieldVal(rowEl, 'fach'),
            lehrer: fieldVal(rowEl, 'lehrer'),
            gruppe: fieldVal(rowEl, 'gruppe'),
            owner: fieldVal(rowEl, 'owner')
        }));
    }

    function buildOptionsForTeamBuild(mapping) {
        const KTB = window.ms365KursteamTeamBuild;
        const KT = window.ms365KursteamTeamNames;
        const KS = window.ms365KursteamSubjectFilterLogic;
        if (!KTB || typeof KTB.buildTeamEntriesFromRows !== 'function') return null;

        const yearPrefix =
            (document.getElementById('yearPrefix') && document.getElementById('yearPrefix').value) ||
            (typeof ns.calcYearPrefix === 'function' ? ns.calcYearPrefix() : 'SJ26');
        const emailDomain =
            typeof window.ms365GetTeacherEmailDomainSuffix === 'function'
                ? window.ms365GetTeacherEmailDomainSuffix()
                : '@';
        const separator =
            document.getElementById('teamSeparator') && document.getElementById('teamSeparator').value
                ? document.getElementById('teamSeparator').value
                : ' | ';
        const pattern =
            typeof ns.getPatternFromBuilder === 'function' ? ns.getPatternFromBuilder() : null;
        const stripEl = document.getElementById('stripSubjectTrailingDigits');
        const combineModeEl = document.getElementById('classCombineMode');
        const classCombineMode =
            combineModeEl && (combineModeEl.value === 'smart' || combineModeEl.value === 'letters')
                ? combineModeEl.value
                : 'concat';

        return {
            yearPrefix,
            emailDomain,
            separator,
            pattern:
                pattern ||
                (KT && typeof KT.defaultTeamNamePattern === 'function' ? KT.defaultTeamNamePattern() : null),
            combineClassNames: ns.combineClassNames,
            isCombinedClassCell: ns.isCombinedClassCell,
            buildGruppenmailBase: ns.buildGruppenmailBase,
            formatKlasseSegmentForGruppenmail: ns.formatKlasseSegmentForGruppenmail,
            sanitizeGruppeForMail: ns.sanitizeGruppeForMail,
            INVALID_CHARS_REPLACE: ns.INVALID_CHARS_REPLACE,
            INVALID_CHARS_TEST: ns.INVALID_CHARS_TEST,
            teacherEmailMapping: mapping || {},
            stripSubjectTrailingDigits: stripEl ? !!stripEl.checked : false,
            classCombineMode,
            normalizeNumberedSubjectFields:
                KS && typeof KS.normalizeNumberedSubjectFields === 'function'
                    ? KS.normalizeNumberedSubjectFields
                    : null
        };
    }

    /** @returns {object|null} */
    ns.buildTeamEntryFromSingleRow = function buildTeamEntryFromSingleRow(row) {
        if (!row || !row.klasse || !row.fach || !row.lehrer) return null;
        const KTB = window.ms365KursteamTeamBuild;
        if (!KTB) return null;

        if (typeof ns.syncTeacherEmailsFromTenant === 'function') {
            ns.syncTeacherEmailsFromTenant([row.lehrer.toUpperCase()], { quiet: true });
        }
        const mapping = Object.assign({}, ns.teacherEmailMapping || {});
        if (row.owner && row.owner.includes('@')) {
            mapping[row.lehrer.toUpperCase()] = row.owner.toLowerCase();
        }
        const opts = buildOptionsForTeamBuild(mapping);
        if (!opts) return null;

        const teams = KTB.buildTeamEntriesFromRows(
            [{ klasse: row.klasse, fach: row.fach, lehrer: row.lehrer, gruppe: row.gruppe || '' }],
            opts
        );
        const team = teams && teams[0] ? teams[0] : null;
        if (team && row.owner && row.owner.includes('@')) {
            team.besitzer = row.owner.toLowerCase();
            team.mappingUsed = true;
            team.isValid = !!(team.teamName && team.gruppenmail && team.besitzer);
            team.error = team.isValid ? null : 'Unvollständige Daten';
        }
        return team;
    };

    /** Rückwärtskompatibel für Tests: erste Zeile bzw. Legacy-Felder. */
    ns.buildSingleTeamPreviewEntry = function buildSingleTeamPreviewEntry() {
        const rows = readSingleTeamRowsFromDom().filter((r) => r.klasse || r.fach || r.lehrer);
        if (rows.length) return ns.buildTeamEntryFromSingleRow(rows[0]);
        return null;
    };

    ns.buildAllSingleTeamEntries = function buildAllSingleTeamEntries() {
        const rows = readSingleTeamRowsFromDom();
        const out = [];
        rows.forEach((r) => {
            if (!r.klasse && !r.fach && !r.lehrer && !r.gruppe && !r.owner) return;
            const team = ns.buildTeamEntryFromSingleRow(r);
            out.push({ row: r, team });
        });
        return out;
    };

    ns.updateSingleTeamOwnerForRow = function updateSingleTeamOwnerForRow(rowEl) {
        if (!rowEl) return;
        const lehrerInp = rowEl.querySelector('[data-st-field="lehrer"]');
        const ownerInp = rowEl.querySelector('[data-st-field="owner"]');
        if (!lehrerInp || !ownerInp) return;
        const lehrer = String(lehrerInp.value || '').trim();
        if (!lehrer) return;
        const email = ns.resolveSingleTeamOwnerEmail(lehrer);
        if (email && !String(ownerInp.value || '').trim()) {
            ownerInp.value = email;
        }
        ns.updateSingleTeamPreview();
    };

    ns.updateSingleTeamPreview = function updateSingleTeamPreview() {
        const box = document.getElementById('singleTeamPreview');
        if (!box) return;
        const built = ns.buildAllSingleTeamEntries();
        const ready = built.filter((x) => x.team && x.team.isValid);
        const incomplete = built.filter((x) => x.row.klasse || x.row.fach || x.row.lehrer).filter((x) => !x.team || !x.team.isValid);

        if (!ready.length && !incomplete.length) {
            box.innerHTML =
                '<p class="muted" style="margin:0;">Vorschau erscheint nach Klasse, Fach und Lehrkraft.</p>';
            return;
        }

        let html = '<div style="display:grid;gap:10px;">';
        ready.forEach((x, i) => {
            const t = x.team;
            html +=
                '<div style="padding-bottom:8px;border-bottom:1px solid var(--border);">' +
                '<div style="font-weight:600;margin-bottom:4px;">Team ' +
                (i + 1) +
                '</div>' +
                '<div><strong>Name:</strong> ' +
                ns.escapeHtml(t.teamName) +
                '</div>' +
                '<div><strong>Gruppenmail:</strong> <code>' +
                ns.escapeHtml(t.gruppenmail) +
                '</code></div>' +
                '<div><strong>Besitzer:</strong> ' +
                ns.escapeHtml(t.besitzer) +
                '</div>' +
                '</div>';
        });
        if (incomplete.length) {
            html +=
                '<div style="color:var(--danger,#c0392b);font-size:0.9em;">' +
                incomplete.length +
                ' Zeile(n) noch unvollständig (Klasse, Fach, Lehrkraft, Besitzer-E-Mail).</div>';
        }
        html += '</div>';
        box.innerHTML = html;
    };

    function wireRowInputs(rowEl) {
        rowEl.querySelectorAll('[data-st-field]').forEach((inp) => {
            inp.addEventListener('input', () => {
                if (inp.getAttribute('data-st-field') === 'lehrer') {
                    ns.updateSingleTeamOwnerForRow(rowEl);
                } else {
                    ns.updateSingleTeamPreview();
                }
            });
            inp.addEventListener('change', () => {
                if (inp.getAttribute('data-st-field') === 'lehrer') {
                    ns.updateSingleTeamOwnerForRow(rowEl);
                } else {
                    ns.updateSingleTeamPreview();
                }
            });
        });
        const rm = rowEl.querySelector('[data-st-remove]');
        if (rm) {
            rm.addEventListener('click', () => {
                const wrap = document.getElementById('singleTeamRows');
                const count = wrap ? wrap.querySelectorAll('.single-team-row').length : 0;
                if (count <= 1) {
                    rowEl.querySelectorAll('[data-st-field]').forEach((inp) => {
                        inp.value = '';
                    });
                    ns.updateSingleTeamPreview();
                    return;
                }
                rowEl.remove();
                ns.updateSingleTeamPreview();
            });
        }
    }

    function attrEsc(text) {
        if (typeof ns.attrEscape === 'function') return ns.attrEscape(text);
        return String(text ?? '')
            .replace(/&/g, '&amp;')
            .replace(/"/g, '&quot;')
            .replace(/</g, '&lt;');
    }

    ns.createSingleTeamRowElement = function createSingleTeamRowElement(prefill) {
        const id = 'st-row-' + ++rowSeq;
        const p = prefill && typeof prefill === 'object' ? prefill : {};
        const row = document.createElement('div');
        row.className = 'single-team-row';
        row.id = id;
        row.innerHTML =
            '<div class="filter-group">' +
            '<label>Klasse</label>' +
            '<input type="text" data-st-field="klasse" list="singleTeamKlasseList" placeholder="z. B. 4AK" autocomplete="off" value="' +
            attrEsc(p.klasse || '') +
            '">' +
            '</div>' +
            '<div class="filter-group">' +
            '<label>Fach</label>' +
            '<input type="text" data-st-field="fach" list="singleTeamFachList" placeholder="z. B. WINF" autocomplete="off" value="' +
            attrEsc(p.fach || '') +
            '">' +
            '</div>' +
            '<div class="filter-group">' +
            '<label>Lehrkraft</label>' +
            '<input type="text" data-st-field="lehrer" list="singleTeamLehrerList" placeholder="z. B. MEI" autocomplete="off" value="' +
            attrEsc(p.lehrer || '') +
            '">' +
            '</div>' +
            '<div class="filter-group">' +
            '<label>Gruppe</label>' +
            '<input type="text" data-st-field="gruppe" placeholder="opt." autocomplete="off" value="' +
            attrEsc(p.gruppe || '') +
            '">' +
            '</div>' +
            '<div class="filter-group">' +
            '<label>Besitzer-E-Mail</label>' +
            '<input type="email" data-st-field="owner" placeholder="aus Stammdaten" autocomplete="off" value="' +
            attrEsc(p.owner || '') +
            '">' +
            '</div>' +
            '<button type="button" class="btn btn-small single-team-row__remove" data-st-remove title="Zeile entfernen" aria-label="Zeile entfernen"><i class="bi bi-trash"></i></button>';
        wireRowInputs(row);
        return row;
    };

    ns.addSingleTeamRow = function addSingleTeamRow(prefill, opts) {
        const wrap = document.getElementById('singleTeamRows');
        if (!wrap) return null;
        const row = ns.createSingleTeamRowElement(prefill);
        wrap.appendChild(row);
        const o = opts && typeof opts === 'object' ? opts : {};
        if (!o.quiet) {
            const klasseInp = row.querySelector('[data-st-field="klasse"]');
            if (klasseInp) klasseInp.focus();
        }
        if (prefill && prefill.lehrer) ns.updateSingleTeamOwnerForRow(row);
        else ns.updateSingleTeamPreview();
        return row;
    };

    ns.resetSingleTeamRows = function resetSingleTeamRows() {
        const wrap = document.getElementById('singleTeamRows');
        if (!wrap) return;
        wrap.replaceChildren();
        ns.addSingleTeamRow({}, { quiet: true });
    };

    ns.mountNameConfigForMode = function mountNameConfigForMode(mode) {
        const block = document.getElementById('kursteamNameConfigBlock');
        const hostSingle = document.getElementById('kursteamNameConfigHostSingle');
        const hostStep5 = document.getElementById('kursteamNameConfigHostStep5');
        if (!block) return;
        const target = mode === 'single' ? hostSingle : hostStep5;
        if (target && block.parentElement !== target) {
            target.appendChild(block);
        }
        const genBtn = document.getElementById('btnGenerateTeamNames');
        if (genBtn) {
            genBtn.style.display = mode === 'single' ? 'none' : '';
        }
        if (typeof ns.renderTeamNameBuilder === 'function') {
            try {
                ns.renderTeamNameBuilder();
            } catch {
                /* ignore */
            }
        }
    };

    ns.applySingleTeamChrome = function applySingleTeamChrome(active) {
        const on = !!active;
        const setDisp = (id, show) => {
            const el = document.getElementById(id);
            if (el) el.style.display = show ? '' : 'none';
        };
        setDisp('kursteamWizardSteps', !on);
        setDisp('kursteamBulkStep0Header', !on);
        setDisp('kursteamBulkIntro', !on);
        setDisp('kursteamBulkChecklist', !on);
        setDisp('kursteamSingleStep0Header', on);
        setDisp('kursteamSingleBanner', on);
        setDisp('kursteamSingleCreateNav', on && ns.currentStep === 7);

        ns.mountNameConfigForMode(on ? 'single' : 'bulk');

        const title = document.getElementById('kursteamCreateStepTitle');
        const sub = document.getElementById('kursteamCreateStepSub');
        const n = Array.isArray(ns.teamsData) ? ns.teamsData.filter((t) => t && t.isValid).length : 0;
        if (title) {
            title.textContent = on
                ? n > 1
                    ? 'Unterrichtsteams anlegen'
                    : 'Unterrichtsteam anlegen'
                : 'Schritt 6 – Kursteams anlegen';
        }
        if (sub) {
            sub.textContent = on
                ? n > 1
                    ? n + ' Teams online oder per CMD anlegen.'
                    : 'Dieses Team online oder per CMD anlegen.'
                : 'Kursteams online über das Azure-Backend anlegen oder alternativ die CMD-Datei per Doppelklick starten.';
        }

        try {
            document.body.classList.toggle('kursteam-mode-single', on);
        } catch {
            /* ignore */
        }
    };

    ns.showSingleTeamForm = function showSingleTeamForm(show) {
        const form = document.getElementById('kursteamSingleForm');
        const cards = document.getElementById('kursteamStep0Cards');
        if (form) form.style.display = show ? 'block' : 'none';
        if (cards) cards.style.display = show ? 'none' : '';
        ns.applySingleTeamChrome(!!show);
    };

    ns.startKursteamSingleTeam = function startKursteamSingleTeam() {
        ns.kursteamEntryMode = 'single';
        if (!ns.teamsGenerated) {
            ns.rawData = [];
            ns.filteredData = [];
            if (typeof ns.invalidateTeams === 'function') ns.invalidateTeams();
            ns.resetSingleTeamRows();
        } else {
            const wrap = document.getElementById('singleTeamRows');
            if (wrap && !wrap.querySelector('.single-team-row')) ns.resetSingleTeamRows();
        }
        ns.refreshSingleTeamDatalists();
        ns.showSingleTeamForm(true);
        if (typeof ns.goToStep === 'function') ns.goToStep(0);
        ns.applySingleTeamChrome(true);
        // yearPrefix ggf. noch leer → aus Datum
        const yp = document.getElementById('yearPrefix');
        if (yp && !String(yp.value || '').trim() && typeof ns.calcYearPrefix === 'function') {
            yp.value = ns.calcYearPrefix();
        }
        const first = document.querySelector('#singleTeamRows [data-st-field="klasse"]');
        if (first) first.focus();
        ns.updateSingleTeamPreview();
        if (typeof ns.showToast === 'function') {
            ns.showToast('Namensschema oben einstellen, Zeilen ausfüllen – mit + weitere Teams.');
        }
    };

    ns.cancelKursteamSingleTeam = function cancelKursteamSingleTeam() {
        ns.kursteamEntryMode = 'unset';
        ns.showSingleTeamForm(false);
        ns.applySingleTeamChrome(false);
        if (typeof ns.goToStep === 'function') ns.goToStep(0);
    };

    /**
     * Baut alle gültigen Teams aus den Zeilen und springt zum Anlege-Schritt.
     * @returns {boolean}
     */
    ns.commitSingleTeamAndGoCreate = function commitSingleTeamAndGoCreate() {
        const built = ns.buildAllSingleTeamEntries();
        const valid = [];
        const dataRows = [];
        built.forEach((x) => {
            if (!x.team || !x.team.isValid) return;
            valid.push(x.team);
            dataRows.push({
                id: Date.now() + Math.random(),
                klasse: x.row.klasse,
                fach: x.row.fach,
                lehrer: x.row.lehrer,
                gruppe: x.row.gruppe || '',
                original: { singleTeam: true }
            });
        });

        if (!valid.length) {
            if (typeof ns.showToast === 'function') {
                ns.showToast('Bitte mindestens eine vollständige Zeile (Klasse, Fach, Lehrkraft, Besitzer).');
            }
            return false;
        }

        ns.rawData = dataRows;
        ns.filteredData = dataRows.slice();
        ns.teamsData = valid;
        ns.teamsGenerated = true;
        ns.kursteamEntryMode = 'single';

        if (typeof ns.resolveDuplicateGruppenmails === 'function') {
            ns.resolveDuplicateGruppenmails(ns.teamsData);
        }
        if (typeof ns.displayTeamsData === 'function') ns.displayTeamsData();
        if (typeof ns.markAutoSaveDirty === 'function') ns.markAutoSaveDirty();

        if (typeof ns.goToStep === 'function') ns.goToStep(7);
        if (typeof ns.applySingleTeamChrome === 'function') ns.applySingleTeamChrome(true);
        if (typeof ns.showToast === 'function') {
            ns.showToast(
                valid.length === 1
                    ? '1 Team vorbereitet – jetzt online oder per CMD anlegen.'
                    : valid.length + ' Teams vorbereitet – jetzt online oder per CMD anlegen.'
            );
        }
        return true;
    };

    ns.bootSingleTeamModeFromQuery = function bootSingleTeamModeFromQuery() {
        try {
            const q = new URLSearchParams(window.location.search || '');
            const mode = String(q.get('mode') || '').toLowerCase();
            if (mode !== 'single' && mode !== 'einzeln') return false;
            ns.startKursteamSingleTeam();
            return true;
        } catch {
            return false;
        }
    };

    function wireSingleTeamForm() {
        /* Rows werden dynamisch verdrahtet; + Button via onclick. */
        const wrap = document.getElementById('singleTeamRows');
        if (!wrap) return;
        if (!wrap.querySelector('.single-team-row')) {
            try {
                ns.resetSingleTeamRows();
            } catch {
                /* DOM unvollständig (z. B. Unit-Test) */
            }
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', wireSingleTeamForm);
    } else {
        wireSingleTeamForm();
    }

    window.startKursteamSingleTeam = function () {
        return ns.startKursteamSingleTeam();
    };
    window.cancelKursteamSingleTeam = function () {
        return ns.cancelKursteamSingleTeam();
    };
    window.commitSingleTeamAndGoCreate = function () {
        return ns.commitSingleTeamAndGoCreate();
    };
    window.addSingleTeamRow = function () {
        return ns.addSingleTeamRow();
    };
})();
