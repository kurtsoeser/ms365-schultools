/**
 * UI-Bootstrap: Wizard + Berechtigungen für sharepoint-liste-stammdaten.html
 */
import { applyStammdatenPackagePermissions } from './stammdaten-liste-permissions.js';
import {
    initStammdatenPermissionsUi,
    readPermissionsFromPickers,
    persistPickersToStorage
} from './stammdaten-permissions-ui.js';
import {
    SPS_WIZARD_STEP_COUNT,
    buildWizardSummary,
    clampWizardStep,
    loadWizardStep,
    saveWizardStep,
    validateWizardStep,
    wizardPhaseHint
} from './sharepoint-liste-stammdaten-wizard.js';
import { resolveIntranetListTitle } from '../../shared/intranet-list-title-logic.js';
import { readGrantRowsFromDom } from './stammdaten-perm-matrix-ui.js';

function $(id) {
    return document.getElementById(id);
}

function log(msg) {
    const el = $('spsLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + msg;
    el.scrollTop = el.scrollHeight;
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function readRunOptsFromDom() {
    const alwaysNew = $('spsAlwaysNew') && $('spsAlwaysNew').checked;
    return {
        syncMode: !alwaysNew,
        removeOrphans: !($('spsRemoveOrphans') && !$('spsRemoveOrphans').checked)
    };
}

export function collectListOptsFromForm() {
    return {
        schueler: $('spsWantSchueler') && $('spsWantSchueler').checked,
        faecher: $('spsWantFaecher') && $('spsWantFaecher').checked,
        fachgruppen: $('spsWantFachgruppen') && $('spsWantFachgruppen').checked,
        arges: $('spsWantArge') && $('spsWantArge').checked,
        klassen: $('spsWantKlassen') && $('spsWantKlassen').checked,
        schuelerTitle: String($('spsSchuelerName') && $('spsSchuelerName').value || '').trim(),
        faecherTitle: String($('spsFaecherName') && $('spsFaecherName').value || '').trim(),
        fachgruppenTitle: String($('spsFachgruppenName') && $('spsFachgruppenName').value || '').trim(),
        argesTitle: String($('spsArgeName') && $('spsArgeName').value || '').trim(),
        klassenTitle: String($('spsKlassenName') && $('spsKlassenName').value || '').trim()
    };
}

function wizardContextFromDom() {
    const lists = collectListOptsFromForm();
    const ro = readRunOptsFromDom();
    return {
        siteUrl: String($('spsSiteUrl') && $('spsSiteUrl').value || '').trim(),
        syncMode: ro.syncMode,
        removeOrphans: ro.removeOrphans,
        lists: {
            schueler: lists.schueler,
            faecher: lists.faecher,
            fachgruppen: lists.fachgruppen,
            arges: lists.arges,
            klassen: lists.klassen
        },
        listTitles: {
            schueler: lists.schuelerTitle,
            faecher: lists.faecherTitle,
            fachgruppen: lists.fachgruppenTitle,
            arges: lists.argesTitle,
            klassen: lists.klassenTitle
        },
        skipPerms: $('spsSkipPerms') && $('spsSkipPerms').checked
    };
}

function refreshSummaryPane() {
    const host = $('spsSummary');
    if (!host) return;
    const sum = buildWizardSummary(wizardContextFromDom());
    host.innerHTML =
        '<ul class="sps-summary-list">' +
        '<li><strong>Website:</strong> ' +
        escapeHtml(sum.site) +
        '</li>' +
        '<li><strong>Modus:</strong> ' +
        escapeHtml(sum.mode) +
        ' · ' +
        escapeHtml(sum.orphans) +
        '</li>' +
        '<li><strong>Listen:</strong> ' +
        escapeHtml(sum.listsText) +
        '</li>' +
        '<li><strong>Berechtigungen:</strong> ' +
        escapeHtml(sum.perms) +
        '</li>' +
        '</ul>';
}

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function showWizardStep(n) {
    const step = clampWizardStep(n);
    saveWizardStep(step);
    for (let i = 1; i <= SPS_WIZARD_STEP_COUNT; i++) {
        const panel = $('spsWizardStep' + i);
        if (!panel) continue;
        const on = i === step;
        panel.hidden = !on;
        panel.setAttribute('aria-hidden', on ? 'false' : 'true');
    }
    document.querySelectorAll('#spsWizardGlance [data-sps-wizard-step]').forEach(function (btn) {
        const sn = parseInt(btn.getAttribute('data-sps-wizard-step'), 10);
        const on = sn === step;
        btn.classList.toggle('is-active', on);
        btn.setAttribute('aria-selected', on ? 'true' : 'false');
        btn.setAttribute('tabindex', on ? '0' : '-1');
    });
    const hint = $('spsPhaseHint');
    if (hint) hint.textContent = wizardPhaseHint(step);
    const back = $('spsWizardBack');
    const next = $('spsWizardNext');
    if (back) back.disabled = step <= 1;
    if (next) {
        const isLast = step >= SPS_WIZARD_STEP_COUNT;
        next.textContent = isLast ? 'Fertig' : 'Weiter';
        next.hidden = isLast;
    }
    const runFooter = $('spsWizardRunFooter');
    if (runFooter) runFooter.hidden = step !== SPS_WIZARD_STEP_COUNT;
    if (step === SPS_WIZARD_STEP_COUNT) refreshSummaryPane();
    try {
        const u = new URL(window.location.href);
        u.searchParams.set('step', String(step));
        window.history.replaceState({}, '', u);
    } catch {
        /* ignore */
    }
}

function tryAdvance(fromStep) {
    const err = validateWizardStep(fromStep, wizardContextFromDom());
    if (err) {
        toast(err);
        return false;
    }
    return true;
}

function wireWizard() {
    document.querySelectorAll('[data-sps-wizard-step]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const target = parseInt(btn.getAttribute('data-sps-wizard-step'), 10);
            const cur = loadWizardStep();
            if (target > cur) {
                for (let s = cur; s < target; s++) {
                    if (!tryAdvance(s)) {
                        showWizardStep(s);
                        return;
                    }
                }
            }
            showWizardStep(target);
        });
    });
    const back = $('spsWizardBack');
    const next = $('spsWizardNext');
    if (back) {
        back.addEventListener('click', function () {
            showWizardStep(loadWizardStep() - 1);
        });
    }
    if (next) {
        next.addEventListener('click', function () {
            const cur = loadWizardStep();
            if (!tryAdvance(cur)) return;
            if (cur >= SPS_WIZARD_STEP_COUNT) {
                toast('Unten „Abgleichen / anlegen“ starten oder Schritt 1–4 anpassen.');
                return;
            }
            showWizardStep(cur + 1);
        });
    }
    let start = loadWizardStep();
    try {
        const q = parseInt(new URLSearchParams(window.location.search).get('step') || '', 10);
        if (q >= 1 && q <= SPS_WIZARD_STEP_COUNT) start = q;
    } catch {
        /* ignore */
    }
    showWizardStep(start);
}

function prefillFromSetup() {
    try {
        const api = window.ms365AppDataV2;
        const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
        if (!setup) return;
        const url = setup.intranetSiteUrl ? String(setup.intranetSiteUrl).trim() : '';
        if (url && $('spsSiteUrl') && !$('spsSiteUrl').value) $('spsSiteUrl').value = url;
        const t =
            setup.intranetListTitles && typeof setup.intranetListTitles === 'object'
                ? setup.intranetListTitles
                : {};
        const map = [
            ['spsSchuelerName', 'schueler'],
            ['spsFaecherName', 'faecher'],
            ['spsFachgruppenName', 'fachgruppen'],
            ['spsArgeName', 'arges'],
            ['spsKlassenName', 'klassen']
        ];
        map.forEach(function (pair) {
            const el = $(pair[0]);
            if (!el || String(el.value || '').trim()) return;
            const title = resolveIntranetListTitle(pair[1], t[pair[1]]);
            if (title) el.value = title;
        });
    } catch {
        /* ignore */
    }
}

window.ms365SpoStammdatenApplyPermissions = async function (webUrl, listOpts, logFn) {
    const write = typeof logFn === 'function' ? logFn : log;
    const grantRows = readGrantRowsFromDom();
    const perms = readPermissionsFromPickers(null, 'spsSkipPerms');
    persistPickersToStorage(null, 'spsSkipPerms');
    return await applyStammdatenPackagePermissions(
        webUrl,
        { ...perms, grantRows },
        write,
        listOpts || collectListOptsFromForm()
    );
};

function syncListNameFieldsDisabled() {
    document.querySelectorAll('[data-sps-list-toggle]').forEach(function (cb) {
        const id = cb.getAttribute('data-sps-list-toggle');
        const nameEl = id ? $(id) : null;
        if (!nameEl) return;
        const on = cb.checked;
        nameEl.disabled = !on;
        const row = cb.closest('.sps-list-row');
        if (row) row.classList.toggle('sps-list-row--off', !on);
    });
    const wantKlassen = $('spsWantKlassen');
    const klassenPersonen = $('spsKlassenPersonen');
    if (klassenPersonen && wantKlassen) {
        klassenPersonen.disabled = !wantKlassen.checked;
        if (!wantKlassen.checked) klassenPersonen.checked = false;
    }
}

document.addEventListener('DOMContentLoaded', function () {
    initStammdatenPermissionsUi();
    prefillFromSetup();
    syncListNameFieldsDisabled();
    document.querySelectorAll('[data-sps-list-toggle], #spsWantKlassen').forEach(function (el) {
        el.addEventListener('change', syncListNameFieldsDisabled);
    });
    wireWizard();

    const permsBtn = $('spsmBtnPerms');
    if (permsBtn) {
        permsBtn.addEventListener('click', function () {
            const webUrl = String($('spsSiteUrl') && $('spsSiteUrl').value || '').trim();
            if (!webUrl) {
                toast('Website-URL fehlt (Schritt 1).');
                showWizardStep(1);
                return;
            }
            if (
                !window.confirm(
                    'Berechtigungen auf die gewählten Stammdaten-Listen anwenden?\n\nVererbung wird gebrochen; breite Site-Gruppen entfernt; Entra-Gruppen erhalten die Rollen gemäß Profil.'
                )
            ) {
                return;
            }
            if ($('spsLog')) $('spsLog').textContent = '';
            window
                .ms365SpoStammdatenApplyPermissions(webUrl, collectListOptsFromForm(), log)
                .then(function () {
                    toast('Berechtigungen angewendet.');
                })
                .catch(function (e) {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                });
        });
    }
});
