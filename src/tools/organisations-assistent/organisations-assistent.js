import { getEl } from '../../shared/utils/dom.js';
import { dlgConfirm } from '../../shared/utils/dialog.js';
import {
    PLAYBOOK_REQUIRED_IDS,
    appendRunLogEntry,
    applyInferredPlaybookDone,
    applyYearActivated,
    buildSchoolYearRunPreview,
    inferPlaybookFromContainer,
    isSchoolYearLabel,
    markPlaybookStep,
    nextSchoolYearLabel,
    normalizePlaybook,
    normalizeRunLog,
    openItemsFromContainer,
    playbookProgress,
    playbookStepDefs,
    summarizeRunPreview,
    yearAlreadyExists
} from './organisations-assistent-logic.js';
import { initCohortPanel, refreshCohortPanel } from './organisations-assistent-cohorts.js';

function normStr(v) {
    return String(v ?? '').trim();
}

function appData() {
    return window.ms365AppDataV2;
}

function getContainer() {
    const api = appData();
    if (!api || typeof api.getContainer !== 'function') return null;
    return api.getContainer();
}

function getRows() {
    const c = getContainer();
    const rows = c && c.structure && Array.isArray(c.structure.rows) ? c.structure.rows : [];
    return rows;
}

function getSettings() {
    const c = getContainer();
    const s =
        c && c.structure && c.structure.settings && typeof c.structure.settings === 'object'
            ? c.structure.settings
            : {};
    return s;
}

function saveStructurePatch(patch) {
    const api = appData();
    if (!api || typeof api.getContainer !== 'function' || typeof api.setContainer !== 'function') {
        throw new Error('Lokale Daten (app-data-v2) nicht verfügbar.');
    }
    const c = api.getContainer();
    if (!c.structure) c.structure = { rows: [], memberships: {}, settings: {} };
    if (patch.rows) c.structure.rows = patch.rows;
    if (patch.memberships) c.structure.memberships = patch.memberships;
    if (patch.settings) c.structure.settings = Object.assign({}, c.structure.settings || {}, patch.settings);
    api.setContainer(c);
}

function getOrganisationAssist(settings) {
    const st =
        settings && settings.organisationAssist && typeof settings.organisationAssist === 'object'
            ? settings.organisationAssist
            : {};
    const cohortPlans = Array.isArray(st.cohortPlans) ? st.cohortPlans : [];
    const currentYear = currentYearFromData();
    const inferred = inferPlaybookFromContainer(getContainer());
    const playbook = applyInferredPlaybookDone(
        normalizePlaybook(st.playbook, currentYear),
        inferred,
        currentYear
    );
    return {
        cohortPlans: cohortPlans,
        playbook: playbook,
        inferred: inferred,
        runLog: normalizeRunLog(st.runLog)
    };
}

function setOrganisationAssist(settings, partial) {
    const cur = getOrganisationAssist(settings);
    const next = {
        cohortPlans: cur.cohortPlans,
        playbook: cur.playbook,
        runLog: cur.runLog
    };
    if (partial && partial.playbook) next.playbook = normalizePlaybook(partial.playbook, currentYearFromData());
    if (partial && Object.prototype.hasOwnProperty.call(partial, 'cohortPlans')) {
        next.cohortPlans = Array.isArray(partial.cohortPlans) ? partial.cohortPlans : [];
    }
    if (partial && partial.runLog) next.runLog = normalizeRunLog(partial.runLog);
    return Object.assign({}, settings || {}, {
        organisationAssist: next
    });
}

function currentYearFromData() {
    const c = getContainer();
    return c && c.years ? normStr(c.years.current) : '';
}

function listYearsFromData() {
    const api = appData();
    if (api && typeof api.listYears === 'function') return api.listYears();
    const c = getContainer();
    const by = c && c.years && c.years.byLabel && typeof c.years.byLabel === 'object' ? c.years.byLabel : {};
    return Object.keys(by).sort();
}

function currentYearBucket() {
    const c = getContainer();
    const y = currentYearFromData();
    const bucket = y && c && c.years && c.years.byLabel ? c.years.byLabel[y] : null;
    const classes = bucket && Array.isArray(bucket.classes) ? bucket.classes : [];
    const students = bucket && Array.isArray(bucket.students) ? bucket.students : [];
    return { year: y, classes: classes.length, students: students.length };
}

function saveRunLog(runLog) {
    const settings = getSettings();
    const nextSettings = setOrganisationAssist(settings, { runLog: normalizeRunLog(runLog) });
    saveStructurePatch({ settings: nextSettings });
}

function statusLabel(status) {
    if (status === 'done') return 'erledigt';
    if (status === 'ready') return 'bereit';
    if (status === 'blocked') return 'noch nicht verknüpft';
    return 'später';
}

function renderRunPreview() {
    const list = getEl('oaRunActions');
    const meta = getEl('oaRunMeta');
    const logEl = getEl('oaRunLog');
    if (!list) return;
    const oa = getOrganisationAssist(getSettings());
    const preview = buildSchoolYearRunPreview(getContainer(), oa.playbook);
    const sum = summarizeRunPreview(preview);
    if (meta) {
        meta.textContent =
            (preview.currentYear ? 'Schuljahr ' + preview.currentYear + '. ' : '') +
            sum.done +
            ' erledigt, ' +
            sum.ready +
            ' bereit, Rest später. Nichts wird hier automatisch nach Microsoft 365 geschrieben.';
    }
    list.replaceChildren();
    preview.actions.forEach(function (a, idx) {
        const li = document.createElement('li');
        li.setAttribute('data-run-status', a.status);
        const strong = document.createElement('strong');
        strong.textContent = String(idx + 1) + '. ' + a.title + ' (' + statusLabel(a.status) + ')';
        li.appendChild(strong);
        if (a.detail) {
            const p = document.createElement('span');
            p.className = 'muted';
            p.style.display = 'block';
            p.style.marginTop = '4px';
            p.textContent = a.detail;
            li.appendChild(p);
        }
        list.appendChild(li);
    });
    if (logEl) {
        logEl.replaceChildren();
        const entries = oa.runLog || [];
        if (!entries.length) {
            const p = document.createElement('p');
            p.className = 'muted';
            p.style.margin = '0';
            p.textContent = 'Noch kein Protokoll. „Vorschau merken“ speichert den Stand lokal.';
            logEl.appendChild(p);
        } else {
            entries
                .slice()
                .reverse()
                .forEach(function (e) {
                    const p = document.createElement('p');
                    p.style.margin = '0 0 8px';
                    const when = e.at ? String(e.at).replace('T', ' ').slice(0, 16) : '';
                    p.textContent = (when ? when + ' — ' : '') + e.summary;
                    logEl.appendChild(p);
                });
        }
    }
}

function snapshotRunPreview() {
    const oa = getOrganisationAssist(getSettings());
    const preview = buildSchoolYearRunPreview(getContainer(), oa.playbook);
    const sum = summarizeRunPreview(preview);
    const summary =
        'Vorschau ' +
        (preview.currentYear || 'ohne Jahr') +
        ': ' +
        sum.done +
        '/' +
        sum.total +
        ' erledigt, ' +
        sum.ready +
        ' bereit.';
    saveRunLog(appendRunLogEntry(oa.runLog, { mode: 'preview', summary: summary }));
    renderRunPreview();
    flash('Vorschau im Protokoll gespeichert (lokal, ohne Microsoft 365).', true);
}

function notifyYearChanged() {
    try {
        window.dispatchEvent(new CustomEvent('ms365-tenant-settings-changed', { detail: { source: 'organisations-assistent' } }));
    } catch {
        /* ignore */
    }
    if (typeof window.ms365RefreshContextBar === 'function') window.ms365RefreshContextBar();
}

function flash(msg, ok) {
    const el = getEl('oaStatus');
    if (!el) return;
    el.textContent = msg;
    el.style.color = ok ? '#146c43' : '#842029';
    el.style.fontWeight = '700';
}

function renderYearCard() {
    const info = currentYearBucket();
    const yearEl = getEl('oaCurrentYear');
    const metaEl = getEl('oaYearMeta');
    if (yearEl) yearEl.textContent = info.year || 'Noch kein Schuljahr gesetzt';
    if (metaEl) {
        metaEl.textContent = info.year
            ? info.classes + ' Klasse(n), ' + info.students + ' Schüler:innen in den Stammdaten dieses Jahres.'
            : 'Legen Sie ein Schuljahr an, damit Klassen- und Schülerlisten zugeordnet werden.';
    }

    const years = listYearsFromData();
    const sel = getEl('oaExistingYear');
    const inp = getEl('oaNewYear');
    if (sel) {
        const prev = normStr(sel.value);
        sel.innerHTML = '';
        const opt0 = document.createElement('option');
        opt0.value = '';
        opt0.textContent = years.length ? '— wählen —' : 'Keine gespeicherten Jahre';
        sel.appendChild(opt0);
        years.forEach(function (y) {
            const o = document.createElement('option');
            o.value = y;
            o.textContent = y === info.year ? y + ' (aktuell)' : y;
            sel.appendChild(o);
        });
        if (prev && years.indexOf(prev) >= 0) sel.value = prev;
        else if (info.year && years.indexOf(info.year) >= 0) sel.value = info.year;
    }
    if (inp && !normStr(inp.value)) {
        inp.value = nextSchoolYearLabel(info.year);
        inp.placeholder = nextSchoolYearLabel(info.year);
    }
    updateCopyHint();
}

function targetYearFromForm() {
    const typed = normStr(getEl('oaNewYear') && getEl('oaNewYear').value);
    if (typed) return typed;
    return normStr(getEl('oaExistingYear') && getEl('oaExistingYear').value);
}

function updateCopyHint() {
    const box = getEl('oaCopyLists');
    if (!box) return;
    const tgt = targetYearFromForm();
    const exists = yearAlreadyExists(listYearsFromData(), tgt);
    box.disabled = exists;
    const wrap = box.closest('label');
    if (wrap) {
        wrap.style.opacity = exists ? '0.55' : '';
        wrap.title = exists
            ? 'Dieses Schuljahr existiert bereits – es wird nur umgeschaltet, Listen bleiben wie gespeichert.'
            : '';
    }
}

function renderPlaybook() {
    const host = getEl('oaPlaybookList');
    if (!host) return;
    const oa = getOrganisationAssist(getSettings());
    const pb = oa.playbook;
    const inferred = oa.inferred || inferPlaybookFromContainer(getContainer());
    const prog = playbookProgress(pb, PLAYBOOK_REQUIRED_IDS);
    const lab = getEl('oaProgressLabel');
    const pct = getEl('oaProgressPct');
    const fill = getEl('oaProgressFill');
    if (lab) lab.textContent = prog.done + ' von ' + prog.total + ' erledigt';
    if (pct) pct.textContent = String(prog.pct) + ' %';
    if (fill) fill.style.width = String(prog.pct) + '%';

    playbookStepDefs().forEach(function (step, idx) {
        const cb = getEl('oaStep-' + step.id);
        if (cb) cb.checked = !!pb.done[step.id];
        const title = getEl('oaStepTitle-' + step.id);
        if (title) {
            const tag = title.querySelector('.oa-later');
            title.textContent = String(idx + 1) + '. ' + step.title + (tag ? ' ' : '');
            if (step.optional) {
                const span = document.createElement('span');
                span.className = 'oa-later';
                span.textContent = 'optional';
                title.appendChild(span);
            }
        }
        const blurb = getEl('oaStepBlurb-' + step.id);
        if (blurb) blurb.textContent = step.blurb;
        const a = host.querySelector('[data-oa-step-href="' + step.id + '"]');
        if (a) {
            a.href = step.href;
            if (step.hrefLabel && a.childNodes.length <= 1) a.textContent = step.hrefLabel;
        }
        const statusEl = getEl('oaStepStatus-' + step.id);
        if (statusEl && inferred.hints && inferred.hints[step.id]) {
            statusEl.textContent = inferred.hints[step.id];
            statusEl.hidden = false;
        } else if (statusEl) {
            statusEl.hidden = true;
        }
    });

    const openEl = getEl('oaOpenList');
    const openBox = getEl('oaOpenBox');
    if (openEl) {
        const items = openItemsFromContainer(getContainer(), pb);
        openEl.replaceChildren();
        items.forEach(function (it) {
            const li = document.createElement('li');
            const a = document.createElement('a');
            a.href = it.href;
            a.textContent = it.text;
            li.appendChild(a);
            openEl.appendChild(li);
        });
        if (openBox) openBox.hidden = !items.length;
    }

    // Auto-erkannte Schritte dauerhaft speichern (wie beim Jahr-Schritt).
    const stored = normalizePlaybook(
        (getSettings().organisationAssist && getSettings().organisationAssist.playbook) || {},
        currentYearFromData()
    );
    const enriched = applyInferredPlaybookDone(stored, inferred, currentYearFromData());
    if (
        enriched.done.students !== stored.done.students ||
        enriched.done.subjects !== stored.done.subjects
    ) {
        savePlaybook(enriched);
    }
}

function playbookStepPanels() {
    return Array.from(document.querySelectorAll('#oaPlaybookList details[data-oa-step]'));
}

function onPlaybookStepOpened(id) {
    if (id === 'names' && typeof window.ms365ClassTeamsRolloverRefresh === 'function') {
        window.ms365ClassTeamsRolloverRefresh();
    }
    if (id === 'graduates') refreshCohortPanel();
}

function openPlaybookStep(id, scroll) {
    const want = String(id || '');
    playbookStepPanels().forEach(function (d) {
        d.open = d.getAttribute('data-oa-step') === want;
    });
    const el = document.querySelector('#oaPlaybookList details[data-oa-step="' + want + '"]');
    if (el && el.open) onPlaybookStepOpened(want);
    if (scroll && el && typeof el.scrollIntoView === 'function') {
        el.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
}

function firstIncompleteStepId() {
    const oa = getOrganisationAssist(getSettings());
    const defs = playbookStepDefs();
    for (let i = 0; i < defs.length; i++) {
        if (defs[i].optional) continue;
        if (!oa.playbook.done[defs[i].id]) return defs[i].id;
    }
    return defs[0] ? defs[0].id : '';
}

function renderExpertMeta() {
    const el = getEl('oaExpertMeta');
    if (!el) return;
    const n = getRows().length;
    const api = appData();
    let teamsN = 0;
    if (api && typeof api.getContainer === 'function' && typeof api.normalizeCoreClassTeams === 'function') {
        teamsN = api.normalizeCoreClassTeams(api.getContainer().core.classTeams || []).length;
    }
    if (!n && !teamsN) {
        el.textContent =
            'Keine SOLL-Strukturzeilen gespeichert. Der Alltags-Schuljahreswechsel braucht diesen Baum nicht.';
        return;
    }
    el.textContent =
        String(n) +
        ' Zeile(n) in der lokalen SOLL-Struktur' +
        (teamsN ? ', ' + String(teamsN) + ' Klassengruppe(n)' : '') +
        '. Den Baum nur in der Gruppenverwaltung ändern – hier nicht duplizieren.';
}

function refreshAll() {
    renderYearCard();
    renderPlaybook();
    renderRunPreview();
    renderExpertMeta();
    refreshCohortPanel();
}

async function activateYear() {
    const api = appData();
    if (!api || typeof api.setCurrentYear !== 'function') {
        throw new Error('Lokale Daten (app-data-v2) nicht verfügbar.');
    }
    const tgt = targetYearFromForm();
    if (!tgt) throw new Error('Bitte ein Schuljahr eintragen (z. B. 2026/27).');
    if (!isSchoolYearLabel(tgt)) {
        throw new Error('Schuljahr bitte als 2026/27 (oder 2026/2027) schreiben.');
    }
    const cur = currentYearFromData();
    const exists = yearAlreadyExists(listYearsFromData(), tgt);
    const copy = !exists && !!(getEl('oaCopyLists') && getEl('oaCopyLists').checked);
    const msg = exists
        ? 'Schuljahr ' + tgt + ' ist bereits gespeichert. Es wird nur zum aktuellen Jahr.'
        : copy && cur
          ? 'Neues Schuljahr ' + tgt + ' anlegen und Schüler- sowie Klassenliste aus ' + cur + ' kopieren?'
          : 'Neues Schuljahr ' + tgt + ' anlegen? Schüler- und Klassenliste bleiben leer, bis Sie sie in den Stammdaten pflegen.';
    const ok = await dlgConfirm(msg, {
        title: 'Schuljahr aktivieren',
        okText: exists ? 'Umschalten' : 'Anlegen'
    });
    if (!ok) return;
    api.setCurrentYear(tgt, copy && cur ? { copyFrom: cur } : {});
    const settings = getSettings();
    const oa = getOrganisationAssist(settings);
    const nextSettings = setOrganisationAssist(settings, { playbook: applyYearActivated(oa.playbook, tgt) });
    saveStructurePatch({ settings: nextSettings });
    notifyYearChanged();
    refreshAll();
    if (typeof window.ms365ClassTeamsRolloverRefresh === 'function') {
        window.ms365ClassTeamsRolloverRefresh();
    }
    openPlaybookStep('names', true);
    flash(
        exists
            ? 'Aktuelles Schuljahr ist jetzt ' + tgt + '.'
            : 'Schuljahr ' + tgt + ' aktiviert' + (copy ? ' (Listen kopiert).' : '.') + ' Ergänzen Sie neue erste Klassen in den Stammdaten.',
        true
    );
}

function bind() {
    getEl('oaExistingYear')?.addEventListener('change', function () {
        const v = normStr(getEl('oaExistingYear').value);
        const inp = getEl('oaNewYear');
        if (inp && v) inp.value = v;
        updateCopyHint();
    });
    getEl('oaNewYear')?.addEventListener('input', updateCopyHint);
    getEl('oaActivateYear')?.addEventListener('click', function () {
        activateYear().catch(function (e) {
            flash(e.message || String(e), false);
        });
    });
    getEl('oaRunPreview')?.addEventListener('click', function () {
        renderRunPreview();
        flash('Vorschau aktualisiert.', true);
    });
    getEl('oaRunSaveLog')?.addEventListener('click', function () {
        try {
            snapshotRunPreview();
        } catch (e) {
            flash(e.message || String(e), false);
        }
    });

    getEl('oaPlaybookList')?.addEventListener('change', function (ev) {
        const t = ev.target;
        if (!t || t.dataset.stepId == null) return;
        try {
            const oa = getOrganisationAssist(getSettings());
            savePlaybook(markPlaybookStep(oa.playbook, t.dataset.stepId, !!t.checked));
            renderPlaybook();
        } catch (e) {
            flash(e.message || String(e), false);
        }
    });

    playbookStepPanels().forEach(function (d) {
        d.addEventListener('toggle', function () {
            if (!d.open) return;
            const id = d.getAttribute('data-oa-step') || '';
            playbookStepPanels().forEach(function (other) {
                if (other !== d) other.open = false;
            });
            onPlaybookStepOpened(id);
        });
    });
}

function hashToStepId(hash) {
    const h = String(hash || '').toLowerCase();
    if (h === '#jahr' || h === '#year') return 'year';
    if (h === '#namen' || h === '#names' || h === '#classes') return 'names';
    if (h === '#abschluss' || h === '#kohorte' || h === '#coh' || h === '#graduates') return 'graduates';
    if (h === '#students' || h === '#schueler') return 'students';
    if (h === '#kursteams') return 'kursteams';
    if (h === '#subjects' || h === '#faecher') return 'subjects';
    if (h === '#expert' || h === '#struktur') return 'expert';
    return '';
}

function openHashDetails() {
    const fromHash = hashToStepId(location.hash);
    openPlaybookStep(fromHash || firstIncompleteStepId(), !!fromHash);
}

bind();
initCohortPanel();
refreshAll();
openHashDetails();
window.addEventListener('hashchange', openHashDetails);
window.addEventListener('ms365-tenant-settings-changed', function () {
    renderYearCard();
    renderPlaybook();
    renderRunPreview();
    renderExpertMeta();
    refreshCohortPanel();
    if (typeof window.ms365ClassTeamsRolloverRefresh === 'function') {
        window.ms365ClassTeamsRolloverRefresh();
    }
});
window.addEventListener('ms365-class-teams-rollover-saved', function () {
    try {
        const oa = getOrganisationAssist(getSettings());
        savePlaybook(markPlaybookStep(oa.playbook, 'names', true));
        renderPlaybook();
        flash('Anzeigenamen gespeichert. Checklisten-Schritt ist markiert.', true);
    } catch (e) {
        flash(e.message || String(e), false);
    }
});
window.addEventListener('ms365-cohort-linked', function () {
    try {
        const oa = getOrganisationAssist(getSettings());
        savePlaybook(markPlaybookStep(oa.playbook, 'graduates', true));
        renderPlaybook();
        flash('Abschlussjahrgang verknüpft. Checklisten-Schritt ist markiert.', true);
    } catch (e) {
        flash(e.message || String(e), false);
    }
});
