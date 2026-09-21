/**
 * Komfort-Modus: genau ein Kursteam vorbereiten und zur Anlage springen.
 * Einstieg: Dashboard „Ein Team nachziehen“ oder ?mode=single
 */
(function () {
    'use strict';

    const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

    function val(id) {
        const el = document.getElementById(id);
        return el ? String(el.value || '').trim() : '';
    }

    function setVal(id, v) {
        const el = document.getElementById(id);
        if (el) el.value = v == null ? '' : String(v);
    }

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

    ns.updateSingleTeamOwnerHint = function updateSingleTeamOwnerHint() {
        const lehrer = val('singleTeamLehrer');
        const hint = document.getElementById('singleTeamOwnerHint');
        const ownerEl = document.getElementById('singleTeamOwner');
        if (!lehrer) {
            if (hint) hint.textContent = 'Lehrkraft-Kürzel eingeben – E-Mail kommt aus den Stammdaten.';
            return;
        }
        const email = ns.resolveSingleTeamOwnerEmail(lehrer);
        if (email) {
            if (ownerEl && !String(ownerEl.value || '').trim()) ownerEl.value = email;
            if (hint) hint.textContent = 'Aus Stammdaten: ' + email;
        } else if (hint) {
            hint.textContent = 'Keine E-Mail in den Stammdaten – bitte Besitzer manuell eintragen.';
        }
        ns.updateSingleTeamPreview();
    };

    ns.buildSingleTeamPreviewEntry = function buildSingleTeamPreviewEntry() {
        const klasse = val('singleTeamKlasse');
        const fach = val('singleTeamFach');
        const lehrer = val('singleTeamLehrer');
        const gruppe = val('singleTeamGruppe');
        const ownerOverride = val('singleTeamOwner');

        if (!klasse || !fach || !lehrer) return null;

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

        if (typeof ns.syncTeacherEmailsFromTenant === 'function') {
            ns.syncTeacherEmailsFromTenant([lehrer.toUpperCase()], { quiet: true });
        }
        const mapping = Object.assign({}, ns.teacherEmailMapping || {});
        if (ownerOverride && ownerOverride.includes('@')) {
            mapping[lehrer.toUpperCase()] = ownerOverride.toLowerCase();
        }

        const teams = KTB.buildTeamEntriesFromRows(
            [{ klasse, fach, lehrer, gruppe }],
            {
                yearPrefix,
                emailDomain,
                separator,
                pattern: pattern || (KT && typeof KT.defaultTeamNamePattern === 'function' ? KT.defaultTeamNamePattern() : null),
                combineClassNames: ns.combineClassNames,
                isCombinedClassCell: ns.isCombinedClassCell,
                buildGruppenmailBase: ns.buildGruppenmailBase,
                formatKlasseSegmentForGruppenmail: ns.formatKlasseSegmentForGruppenmail,
                sanitizeGruppeForMail: ns.sanitizeGruppeForMail,
                INVALID_CHARS_REPLACE: ns.INVALID_CHARS_REPLACE,
                INVALID_CHARS_TEST: ns.INVALID_CHARS_TEST,
                teacherEmailMapping: mapping,
                stripSubjectTrailingDigits: stripEl ? !!stripEl.checked : false,
                normalizeNumberedSubjectFields:
                    KS && typeof KS.normalizeNumberedSubjectFields === 'function'
                        ? KS.normalizeNumberedSubjectFields
                        : null
            }
        );

        return teams && teams[0] ? teams[0] : null;
    };

    ns.updateSingleTeamPreview = function updateSingleTeamPreview() {
        const box = document.getElementById('singleTeamPreview');
        if (!box) return;
        const team = ns.buildSingleTeamPreviewEntry();
        if (!team) {
            box.innerHTML =
                '<p class="muted" style="margin:0;">Vorschau erscheint nach Klasse, Fach und Lehrkraft.</p>';
            return;
        }
        box.innerHTML =
            '<div style="display:grid;gap:6px;">' +
            '<div><strong>Team-Name:</strong> ' +
            ns.escapeHtml(team.teamName) +
            '</div>' +
            '<div><strong>Gruppenmail:</strong> <code>' +
            ns.escapeHtml(team.gruppenmail) +
            '</code></div>' +
            '<div><strong>Besitzer:</strong> ' +
            ns.escapeHtml(team.besitzer) +
            '</div>' +
            (team.isValid
                ? ''
                : '<div style="color:var(--danger,#c0392b);">' +
                  ns.escapeHtml(team.error || 'Unvollständig') +
                  '</div>') +
            '</div>';
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

        const title = document.getElementById('kursteamCreateStepTitle');
        const sub = document.getElementById('kursteamCreateStepSub');
        if (title) {
            title.textContent = on ? 'Unterrichtsteam anlegen' : 'Schritt 6 – Kursteams anlegen';
        }
        if (sub) {
            sub.textContent = on
                ? 'Genau dieses eine Team online oder per CMD anlegen.'
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
        // Nicht invalidateTeams hier erzwingen wenn wir nur zurück vom Anlege-Schritt kommen
        if (!ns.teamsGenerated) {
            ns.rawData = [];
            ns.filteredData = [];
            if (typeof ns.invalidateTeams === 'function') ns.invalidateTeams();
        }
        ns.refreshSingleTeamDatalists();
        ns.showSingleTeamForm(true);
        if (typeof ns.goToStep === 'function') ns.goToStep(0);
        ns.applySingleTeamChrome(true);
        const k = document.getElementById('singleTeamKlasse');
        if (k) k.focus();
        ns.updateSingleTeamOwnerHint();
        ns.updateSingleTeamPreview();
        if (typeof ns.showToast === 'function') {
            ns.showToast('Einzelnes Unterrichtsteam – Felder ausfüllen, dann anlegen.');
        }
    };

    ns.cancelKursteamSingleTeam = function cancelKursteamSingleTeam() {
        ns.kursteamEntryMode = 'unset';
        ns.showSingleTeamForm(false);
        ns.applySingleTeamChrome(false);
        if (typeof ns.goToStep === 'function') ns.goToStep(0);
    };

    /**
     * Baut genau ein gültiges Team und springt zum Anlege-Schritt.
     * @returns {boolean}
     */
    ns.commitSingleTeamAndGoCreate = function commitSingleTeamAndGoCreate() {
        const team = ns.buildSingleTeamPreviewEntry();
        if (!team) {
            if (typeof ns.showToast === 'function') {
                ns.showToast('Bitte Klasse, Fach und Lehrkraft ausfüllen.');
            }
            return false;
        }
        const ownerOverride = val('singleTeamOwner');
        if (ownerOverride && ownerOverride.includes('@')) {
            team.besitzer = ownerOverride.toLowerCase();
            team.mappingUsed = true;
            team.isValid = !!(team.teamName && team.gruppenmail && team.besitzer);
            team.error = team.isValid ? null : 'Unvollständige Daten';
        }
        if (!team.isValid) {
            if (typeof ns.showToast === 'function') {
                ns.showToast(team.error || 'Team noch unvollständig (Besitzer-E-Mail prüfen).');
            }
            return false;
        }

        const klasse = val('singleTeamKlasse');
        const fach = val('singleTeamFach');
        const lehrer = val('singleTeamLehrer');
        const gruppe = val('singleTeamGruppe');
        const id = Date.now();
        const row = { id, klasse, fach, lehrer, gruppe, original: { singleTeam: true } };
        ns.rawData = [row];
        ns.filteredData = [row];
        ns.teamsData = [team];
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
            ns.showToast('1 Team vorbereitet – jetzt online oder per CMD anlegen.');
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
        ['singleTeamKlasse', 'singleTeamFach', 'singleTeamLehrer', 'singleTeamGruppe', 'singleTeamOwner'].forEach(
            (id) => {
                const el = document.getElementById(id);
                if (!el) return;
                el.addEventListener('input', () => {
                    if (id === 'singleTeamLehrer') ns.updateSingleTeamOwnerHint();
                    else ns.updateSingleTeamPreview();
                });
                el.addEventListener('change', () => {
                    if (id === 'singleTeamLehrer') ns.updateSingleTeamOwnerHint();
                    else ns.updateSingleTeamPreview();
                });
            }
        );
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
})();
