/**
 * Diplomarbeiten / VWA UI – Setup (Option C), Jahrgänge/Kanäle, Legacy-Teams.
 * Transport: shared/graph-client.js (R22 / b11) für Graph-JSON; Token weiter über Unified Groups.
 */
import { graphJson as sharedGraphJson } from '../../shared/graph-client.js';
import { pickEntraUser, pickEntraUsers } from '../../shared/entra-user-picker.js';
import { escapeHtml } from '../../shared/utils/strings.js';
import {
    linkCohortTeamEnsured,
    provisionArbeitChannel,
    provisionCohortTeam,
    searchTeamsGroups,
    syncArbeitChannel
} from './diplomarbeiten-graph.js';
import {
    buildDiplomPlan,
    buildCohortTeamDisplayName,
    buildCohortTeamMailNickname,
    buildPrivateChannelDisplayName,
    cohortListStatus,
    diplomNameChoices,
    diplomSetupChoices,
    filterDiplomGroups,
    labelsForArbeitTyp,
    maturajahrgangChoices
} from './diplomarbeiten-logic.js';
import { pullDiplomSettingsFromSchool, pushDiplomSettingsToSchool } from './diplomarbeiten-school.js';
import { loadDiplomSettings, saveDiplomSettings } from './diplomarbeiten-settings.js';
import {
    ARBEIT_STATUS_LABEL,
    getArbeit,
    getCohort,
    loadState,
    patchArbeit,
    upsertArbeit,
    upsertCohort
} from './diplomarbeiten-store.js';

/**
 * @type {{
 *   mentor: {id:string,displayName:string,mail:string}|null,
 *   students: Array<{id:string,displayName:string,mail:string}>,
 *   teamHits: object[],
 *   teamQuery: string,
 *   teamSearched: boolean,
 *   editId: string,
 *   topic: string,
 *   channelName: string,
 *   flash: {type:string,text:string}|null
 * }}
 */
const draftPick = {
    mentor: null,
    students: [],
    teamHits: [],
    teamQuery: '',
    teamSearched: false,
    editId: '',
    topic: '',
    channelName: '',
    flash: null
};

/** @type {{ arbeitTyp: string, modus: string }|null} */
let setupDraft = null;
/** @type {string} */
let schoolNote = '';

function setupSelection() {
    const s = settings();
    if (!setupDraft) {
        setupDraft = { arbeitTyp: s.arbeitTyp, modus: s.modus };
    }
    return setupDraft;
}

function setFlash(type, text) {
    draftPick.flash = text ? { type: type || 'ok', text: String(text) } : null;
    paintFlash();
}

function bindFlashClose(host) {
    const btn = host && host.querySelector('#daFlashClose');
    if (!btn) return;
    btn.addEventListener('click', function () {
        draftPick.flash = null;
        render();
    });
}

function paintFlash() {
    const host = $('daApp');
    if (!host) return;
    const existing = host.querySelector('.da-flash');
    if (!draftPick.flash) {
        if (existing) existing.remove();
        return;
    }
    const html = renderFlash();
    if (existing) existing.outerHTML = html;
    else host.insertAdjacentHTML('afterbegin', html);
    bindFlashClose(host);
}

function clearArbeitDraft() {
    draftPick.mentor = null;
    draftPick.students = [];
    draftPick.editId = '';
    draftPick.topic = '';
    draftPick.channelName = '';
}

function suggestChannelName(topic, students) {
    const L = labels();
    const student = (students || [])
        .map(function (s) {
            return s.displayName || s.mail || '';
        })
        .filter(Boolean)
        .join(' / ');
    return buildPrivateChannelDisplayName({
        topic: topic || '',
        student: student,
        arbeitTyp: settings().arbeitTyp
    });
}

function $(id) {
    return document.getElementById(id);
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function settings() {
    return loadDiplomSettings();
}

function labels() {
    return labelsForArbeitTyp(settings().arbeitTyp);
}

/** @type {string} */
let actionLog = '';

function resetActionLog() {
    actionLog = '';
    const el = $('daLog');
    if (el) el.textContent = '';
}

function appendActionLog(msg) {
    const line = String(msg || '');
    if (!line) return;
    actionLog += (actionLog ? '\n' : '') + line;
    const el = $('daLog');
    if (el) {
        el.textContent = actionLog;
        const details = el.closest('details');
        if (details) details.open = true;
    }
}

function protocolDetails() {
    return (
        '<details class="tm-tech da-log"' +
        (actionLog ? ' open' : '') +
        '><summary>Protokoll</summary><pre id="daLog" class="tm-log" style="white-space:pre-wrap">' +
        escapeHtml(actionLog) +
        '</pre></details>'
    );
}

/* ——— Legacy (eigene Teams) ——— */

function readLegacyForm() {
    return {
        year: ($('daYear') && $('daYear').value) || '',
        topic: ($('daTopic') && $('daTopic').value) || '',
        student: ($('daStudent') && $('daStudent').value) || '',
        mentor: ($('daMentor') && $('daMentor').value) || '',
        asTeam: !($('daGroupOnly') && $('daGroupOnly').checked),
        arbeitTyp: settings().arbeitTyp
    };
}

function refreshLegacyPreview() {
    const plan = buildDiplomPlan(readLegacyForm());
    const el = $('daPreview');
    if (!el) return;
    if (!plan.ok) {
        el.innerHTML = '<p class="muted">' + escapeHtml(plan.issues.join(' · ')) + '</p>';
        return;
    }
    el.innerHTML =
        '<p><strong>Anzeigename:</strong> ' +
        escapeHtml(plan.displayName) +
        '</p><p><strong>Alias:</strong> <code>' +
        escapeHtml(plan.mailNickname) +
        '</code></p><p class="muted">' +
        escapeHtml(plan.description) +
        '</p>';
}

async function createLegacyDiplom() {
    const G = window.ms365GraphUnifiedGroups;
    if (!G) throw new Error('Graph-Modul nicht geladen.');
    const plan = buildDiplomPlan(readLegacyForm());
    if (!plan.ok) throw new Error(plan.issues.join(', '));
    if (!window.confirm('Gruppe anlegen?\n\n' + plan.displayName + '\n' + plan.mailNickname)) return;

    const token = await G.getGraphToken([
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.ReadWrite.All',
        'https://graph.microsoft.com/Team.Create'
    ]);
    log('Lege Gruppe an …');
    const group = await G.createUnifiedGroup(token, plan.displayName, plan.mailNickname, plan.description);
    log('Gruppe: ' + (group.id || ''));
    if (plan.asTeam) {
        log('Provisioniere Team …');
        await G.provisionTeamForGroup(token, group.id);
        log('Team bereit.');
    }
    setFlash('ok', labels().singular + ' angelegt. In Teams heißt es „' + plan.displayName + '“.');
    await loadDiplomList();
    render();
}

/** @type {object[]|null} */
let legacyGroups = null;

const DIPLOM_GROUP_FILTERS = [
    "startswith(mailNickname,'dipl-')",
    "startswith(mailNickname,'vwa-')",
    "startswith(displayName,'Diplomarbeit')",
    "startswith(displayName,'Vorwissenschaftliche')",
    "startswith(displayName,'VWA')"
];
async function readDiplomGroupFilter(token, filter) {
    const select = 'id,displayName,mailNickname,mail,description';
    let next =
        "/groups?$filter=" +
        filter +
        '&$select=' +
        select +
        '&$top=100';
    const rows = [];
    let page = 0;
    while (next && page < 5) {
        page++;
        let data;
        try {
            data = await sharedGraphJson('GET', next, token, undefined, { ConsistencyLevel: 'eventual' });
        } catch (e) {
            if (page > 1) throw e;
            data = await sharedGraphJson('GET', next, token, undefined);
        }
        if (Array.isArray(data.value)) rows.push.apply(rows, data.value);
        next = data['@odata.nextLink'] || null;
    }
    return rows;
}

function paintLegacyList() {
    const body = $('daBody');
    const status = $('daListStatus');
    if (!body) return;
    if (!legacyGroups) {
        body.innerHTML = '<tr><td colspan="3" class="muted">Noch nicht geladen.</td></tr>';
        if (status) status.textContent = '';
        return;
    }
    const q = ($('daFilter') && $('daFilter').value) || '';
    const rows = filterDiplomGroups(legacyGroups, q);
    if (!rows.length) {
        body.innerHTML = '<tr><td colspan="3" class="muted">Keine Treffer.</td></tr>';
        if (status) status.textContent = legacyGroups.length + ' geladen, kein Treffer für die Suche';
        return;
    }
    body.innerHTML = rows
        .map(function (g) {
            return (
                '<tr><td>' +
                escapeHtml(g.displayName || '') +
                '</td><td><code>' +
                escapeHtml(g.mailNickname || '') +
                '</code></td><td>' +
                escapeHtml(g.mail || '') +
                '</td></tr>'
            );
        })
        .join('');
    if (status) status.textContent = rows.length + ' Treffer';
}

async function loadDiplomList() {
    const G = window.ms365GraphUnifiedGroups;
    if (!G) throw new Error('Graph-Modul nicht geladen.');
    const status = $('daListStatus');
    if (status) status.textContent = 'Lade Diplom- und VWA-Teams …';
    const token = await G.getGraphToken([
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.Read.All',
        'https://graph.microsoft.com/Directory.Read.All'
    ]);
    const byId = new Map();
    let failed = 0;
    let lastError = null;
    for (let i = 0; i < DIPLOM_GROUP_FILTERS.length; i++) {
        try {
            const rows = await readDiplomGroupFilter(token, DIPLOM_GROUP_FILTERS[i]);
            rows.forEach(function (g) {
                if (g && g.id) byId.set(String(g.id), g);
            });
        } catch (e) {
            failed++;
            lastError = e;
        }
    }
    if (!byId.size && failed === DIPLOM_GROUP_FILTERS.length) {
        throw lastError || new Error('Teams konnten nicht geladen werden.');
    }
    legacyGroups = Array.from(byId.values());
    paintLegacyList();
}

/* ——— Shell / Views ——— */

function ensureShell() {
    const host = $('daApp');
    if (!host) return null;
    return host;
}

function render() {
    const host = ensureShell();
    if (!host) return;
    const s = settings();
    const L = labels();

    if (!s.setupDone) {
        host.innerHTML = renderSetup(L);
        bindSetup(host);
        return;
    }

    let view = host.getAttribute('data-view') || (s.modus === 'legacy' ? 'legacy' : 'jahrgaenge');
    if (s.modus !== 'legacy' && view === 'legacy') view = 'jahrgaenge';
    if (s.modus === 'legacy' && view !== 'setup') view = 'legacy';
    // „Arbeiten“-Tab ohne Kontext → Jahrgangsliste
    if (view === 'arbeiten' && !host.getAttribute('data-cohort')) view = 'jahrgaenge';
    const tabView = view === 'cohort' ? 'jahrgaenge' : view;
    host.innerHTML =
        renderTabs(L, s, tabView) +
        '<div class="da-view">' +
        (view === 'setup'
            ? renderSetup(L)
            : view === 'legacy'
              ? renderLegacy(L)
              : view === 'cohort'
                ? renderCohortDetail(L)
                : renderJahrgaenge(L)) +
        '</div>';
    bindTabs(host);
    bindFlashClose(host);
    if (view === 'setup') bindSetup(host);
    else if (view === 'legacy') bindLegacy(host);
    else if (view === 'cohort') bindCohortDetail(host);
    else bindJahrgaenge(host);
}

function renderTabs(L, s, view) {
    const tabs = [
        { id: 'jahrgaenge', label: 'Jahrgänge', show: s.modus === 'channels' },
        { id: 'legacy', label: 'Legacy-Teams', show: s.modus === 'legacy' },
        { id: 'setup', label: 'Setup', show: true }
    ];
    return (
        '<div class="da-tabs" role="tablist">' +
        tabs
            .filter((t) => t.show)
            .map(function (t) {
                return (
                    '<button type="button" class="fb-btn' +
                    (view === t.id ? '' : ' fb-btn--ghost') +
                    '" data-da-view="' +
                    t.id +
                    '">' +
                    escapeHtml(t.label) +
                    '</button>'
                );
            })
            .join(' ') +
        '<span class="muted" style="margin-left:8px">' +
        escapeHtml(L.toolTitle) +
        ' · ' +
        (s.modus === 'channels' ? 'Ein Team pro Jahrgang' : 'Ein Team pro Arbeit') +
        '</span></div>'
    );
}

function renderSetup(L) {
    const pick = setupSelection();
    const names = diplomNameChoices()
        .map(function (n) {
            const on = pick.arbeitTyp === n.arbeitTyp;
            return (
                '<button type="button" class="fb-btn' +
                (on ? '' : ' fb-btn--ghost') +
                '" data-da-typ="' +
                n.arbeitTyp +
                '" aria-pressed="' +
                (on ? 'true' : 'false') +
                '">' +
                escapeHtml(n.label) +
                '</button>'
            );
        })
        .join('');
    const cards = diplomSetupChoices(pick.arbeitTyp)
        .map(function (c) {
            const on = pick.modus === c.modus;
            return (
                '<button type="button" class="da-choice' +
                (on ? ' da-choice--on' : '') +
                '" data-da-modus="' +
                c.modus +
                '" aria-pressed="' +
                (on ? 'true' : 'false') +
                '">' +
                '<strong>' +
                escapeHtml(c.title) +
                '</strong>' +
                '<span class="da-choice__text">' +
                escapeHtml(c.text) +
                '</span>' +
                '<span class="da-choice__example">' +
                escapeHtml(c.example) +
                '</span></button>'
            );
        })
        .join('');
    const note = schoolNote
        ? '<p class="da-school-note" role="status">' + escapeHtml(schoolNote) + '</p>'
        : '';
    return (
        '<section class="tm-panel">' +
        '<div class="tm-panel__head"><i class="bi bi-mortarboard" aria-hidden="true"></i><div>' +
        '<h3>Wie sollen die Arbeiten in Teams liegen?</h3>' +
        '<p>Die Wahl gilt für die Schule. ' +
        escapeHtml(L.singular) +
        ' ist nur die Bezeichnung, nicht ein zweites Werkzeug.</p></div></div>' +
        '<p class="da-name-label">So heißen die Arbeiten an der Schule</p>' +
        '<div class="da-name-row">' +
        names +
        '</div>' +
        '<div class="da-choices">' +
        cards +
        '</div>' +
        note +
        '<div class="tm-actions">' +
        '<button type="button" class="fb-btn" id="daSetupSave"><i class="bi bi-check2"></i>Für die Schule festlegen</button>' +
        '</div></section>'
    );
}

function bindSetup(host) {
    host.querySelectorAll('[data-da-typ]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const pick = setupSelection();
            pick.arbeitTyp = btn.getAttribute('data-da-typ') || 'diplom';
            schoolNote = '';
            render();
        });
    });
    host.querySelectorAll('[data-da-modus]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const pick = setupSelection();
            pick.modus = btn.getAttribute('data-da-modus') === 'legacy' ? 'legacy' : 'channels';
            schoolNote = '';
            render();
        });
    });
    const btn = host.querySelector('#daSetupSave');
    if (!btn) return;
    btn.addEventListener('click', function () {
        const pick = setupSelection();
        btn.disabled = true;
        saveDiplomSettings({
            arbeitTyp: pick.arbeitTyp,
            modus: pick.modus,
            setupDone: true
        });
        const h = ensureShell();
        if (h) {
            h.setAttribute('data-view', pick.modus === 'legacy' ? 'legacy' : 'jahrgaenge');
        }
        updateChrome();
        void pushDiplomSettingsToSchool().then(function (result) {
            if (result && result.ok) {
                schoolNote = 'Für die Schule gespeichert. Ein anderer Rechner übernimmt die Wahl beim Öffnen.';
                toast('Für die Schule festgelegt.');
            } else if (result && result.reason === 'no-library') {
                schoolNote =
                    'In diesem Browser gespeichert. Die IT-Sicherung der Schule ist noch nicht verbunden, deshalb gilt die Wahl hier, bis die Sicherung steht.';
                toast('In diesem Browser gespeichert.');
            } else {
                schoolNote =
                    'In diesem Browser gespeichert. Die Schul-Sicherung hat die Wahl nicht übernommen.';
                toast('In diesem Browser gespeichert.');
            }
            setupDraft = null;
            render();
        });
    });
}

function clearTeamSearch() {
    draftPick.teamHits = [];
    draftPick.teamQuery = '';
    draftPick.teamSearched = false;
}

function cohortTeamName(cohort) {
    if (cohort && cohort.displayName) return cohort.displayName;
    const s = settings();
    return buildCohortTeamDisplayName({
        maturajahrgang: cohort ? cohort.maturajahrgang : '',
        arbeitTyp: s.arbeitTyp
    });
}

function renderTeamSetup(cohort) {
    const name = cohortTeamName(cohort);
    const hits = draftPick.teamHits || [];
    const list = hits.length
        ? '<ul class="da-team-list">' +
          hits
              .map(function (g) {
                  const meta = g.hasTeam ? 'Schon ein Team' : 'Wird zum Team';
                  const mail = g.mail || g.mailNickname || '';
                  return (
                      '<li><button type="button" class="da-team-hit" data-link-team="' +
                      escapeHtml(g.id) +
                      '" data-link-name="' +
                      escapeHtml(g.displayName) +
                      '" data-link-nick="' +
                      escapeHtml(g.mailNickname) +
                      '"><span class="da-team-hit__name">' +
                      escapeHtml(g.displayName || 'Ohne Namen') +
                      '</span><span class="da-team-hit__meta">' +
                      escapeHtml(mail ? mail + ' · ' + meta : meta) +
                      '</span></button></li>'
                  );
              })
              .join('') +
          '</ul>'
        : draftPick.teamSearched
          ? '<p class="muted">Kein Team mit diesem Namen.</p>'
          : '';
    return (
        '<div class="da-team-setup">' +
        '<p class="da-team-lead">Noch kein Team für diesen Jahrgang.</p>' +
        '<p class="da-team-name-label">In Teams steht</p>' +
        '<p class="da-team-name">' +
        escapeHtml(name) +
        '</p>' +
        '<button type="button" class="fb-btn" id="daCohortCreate"><i class="bi bi-plus-lg"></i>Team anlegen</button>' +
        '<div class="da-team-search">' +
        '<label for="daTeamSearch">Bestehendes Team suchen</label>' +
        '<div class="tm-actions da-inline">' +
        '<input id="daTeamSearch" value="' +
        escapeHtml(draftPick.teamQuery || '') +
        '" placeholder="Name in Teams">' +
        '<button type="button" class="fb-btn fb-btn--ghost" id="daTeamSearchBtn">Suchen</button></div>' +
        list +
        '</div>' +
        '<details class="da-advanced"><summary>Erweitert</summary>' +
        '<div class="tm-field"><label for="daTeamGuid">Gruppen-ID</label>' +
        '<div class="tm-actions da-inline">' +
        '<input id="daTeamGuid" placeholder="xxxxxxxx-xxxx-…" spellcheck="false">' +
        '<button type="button" class="fb-btn fb-btn--ghost" id="daTeamGuidLink">Verbinden</button></div></div></details></div>'
    );
}

function openCohort(id) {
    const h = ensureShell();
    if (!h) return;
    clearArbeitDraft();
    clearTeamSearch();
    draftPick.flash = null;
    h.setAttribute('data-view', 'cohort');
    h.setAttribute('data-cohort', id || '');
    render();
}

function renderJahrgangSteps(L) {
    return (
        '<ol class="da-steps">' +
        '<li><span class="da-steps__n">1</span><div><strong>Jahrgang</strong>' +
        '<p>Maturajahr oben auswählen.</p></div></li>' +
        '<li><span class="da-steps__n">2</span><div><strong>Team</strong>' +
        '<p>Ein Team für diesen Jahrgang.</p></div></li>' +
        '<li><span class="da-steps__n">3</span><div><strong>Kanäle</strong>' +
        '<p>Jede ' +
        escapeHtml(L.singular) +
        ' bekommt darin einen privaten Kanal.</p></div></li>' +
        '</ol>'
    );
}

function renderJahrgaenge(L) {
    const state = loadState();
    const choices = maturajahrgangChoices(
        new Date(),
        state.cohorts.map(function (c) {
            return c.maturajahrgang;
        })
    );
    const rows = state.cohorts
        .map(function (c) {
            const status = cohortListStatus(c, state.arbeiten);
            const meta = c.displayName
                ? '<span class="da-cohort-row__meta">' + escapeHtml(c.displayName) + '</span>'
                : '';
            return (
                '<button type="button" class="da-cohort-row" data-cohort-open="' +
                escapeHtml(c.id) +
                '"><span><span class="da-cohort-row__name">' +
                escapeHtml(c.maturajahrgang) +
                '</span>' +
                meta +
                '</span><span class="da-cohort-row__status">' +
                escapeHtml(status.label) +
                '</span></button>'
            );
        })
        .join('');

    const now = new Date();
    const start = now.getMonth() >= 8 ? now.getFullYear() : now.getFullYear() - 1;
    const currentValue = start + '/' + String(start + 1).slice(-2);
    const selectedValue = (
        choices.find(function (c) {
            return c.value === currentValue;
        }) ||
        choices.find(function (c) {
            return c.value > currentValue;
        }) ||
        choices[choices.length - 1] || { value: '' }
    ).value;
    const picker = choices.length
        ? '<div class="da-year-row"><div class="tm-field"><label for="daCohortJahr">Maturajahrgang</label>' +
          '<select id="daCohortJahr">' +
          choices
              .map(function (c) {
                  return (
                      '<option value="' +
                      escapeHtml(c.value) +
                      '"' +
                      (c.value === selectedValue ? ' selected' : '') +
                      '>' +
                      escapeHtml(c.label) +
                      '</option>'
                  );
              })
              .join('') +
          '</select></div>' +
          '<button type="button" class="fb-btn" id="daCohortAdd"><i class="bi bi-plus-lg"></i>Jahrgang anlegen</button></div>'
        : '<p class="muted">Die angebotenen Maturajahre sind schon angelegt.</p>';

    return (
        '<section class="tm-panel">' +
        '<div class="tm-panel__head"><i class="bi bi-calendar3" aria-hidden="true"></i><div>' +
        '<h3>Maturajahrgänge</h3>' +
        '<p>Jahrgang, dann Team, dann Kanäle.</p></div></div>' +
        picker +
        (rows ? '<div class="da-cohorts">' + rows + '</div>' : renderJahrgangSteps(L)) +
        '</section>'
    );
}

function bindJahrgaenge(host) {
    const add = host.querySelector('#daCohortAdd');
    if (add) {
        add.addEventListener('click', function () {
            const jahr = String((host.querySelector('#daCohortJahr') || {}).value || '').trim();
            const s = settings();
            const displayName = buildCohortTeamDisplayName({
                maturajahrgang: jahr,
                arbeitTyp: s.arbeitTyp
            });
            const mailNickname = buildCohortTeamMailNickname({
                maturajahrgang: jahr,
                arbeitTyp: s.arbeitTyp
            });
            const res = upsertCohort({ maturajahrgang: jahr, displayName, mailNickname });
            if (!res.ok) {
                toast(res.errors.join(' '));
                return;
            }
            toast('Jahrgang ' + jahr + ' steht in der Liste.');
            render();
        });
    }
    host.querySelectorAll('[data-cohort-open]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            openCohort(btn.getAttribute('data-cohort-open'));
        });
    });
}

function renderPickedPeople() {
    const mentor = draftPick.mentor
        ? '<div class="da-chip">' +
          '<span class="da-chip__name">' +
          escapeHtml(draftPick.mentor.displayName || '–') +
          '</span>' +
          '<span class="da-chip__mail">' +
          escapeHtml(draftPick.mentor.mail || '') +
          '</span>' +
          '<button type="button" class="da-chip__x" id="daClearMentor" aria-label="Entfernen">×</button></div>'
        : '<span class="muted">noch niemand gewählt</span>';
    const students =
        draftPick.students.length > 0
            ? '<div class="da-chip-row">' +
              draftPick.students
                  .map(function (u, i) {
                      return (
                          '<div class="da-chip">' +
                          '<span class="da-chip__name">' +
                          escapeHtml(u.displayName || '–') +
                          '</span>' +
                          '<span class="da-chip__mail">' +
                          escapeHtml(u.mail || '') +
                          '</span>' +
                          '<button type="button" class="da-chip__x" data-rm-student="' +
                          i +
                          '" aria-label="Entfernen">×</button></div>'
                      );
                  })
                  .join('') +
              '</div>'
            : '<span class="muted">noch niemand gewählt</span>';
    return { mentor: mentor, students: students };
}

function renderFlash() {
    if (!draftPick.flash) return '';
    const cls =
        draftPick.flash.type === 'err'
            ? 'da-flash da-flash--err'
            : draftPick.flash.type === 'warn'
              ? 'da-flash da-flash--warn'
              : 'da-flash da-flash--ok';
    return (
        '<div class="' +
        cls +
        '" role="status">' +
        escapeHtml(draftPick.flash.text) +
        ' <button type="button" class="da-flash__close" id="daFlashClose" aria-label="Schließen">×</button></div>'
    );
}

function renderArbeitList(cohort, count, rows) {
    if (!cohort.teamId && !count) return '';
    const countText = count === 1 ? '1 Arbeit' : count ? count + ' Arbeiten' : 'Noch keine Arbeit.';
    return (
        '<section class="tm-panel da-panel" id="daArbeitList">' +
        '<div class="tm-panel__head"><i class="bi bi-list-ul" aria-hidden="true"></i><div>' +
        '<h3>Arbeiten in ' +
        escapeHtml(cohort.maturajahrgang) +
        '</h3>' +
        '<p>' +
        countText +
        '</p>' +
        (count
            ? '<p class="da-done-note">Abschließen setzt nur die Markierung in dieser Übersicht. Der Kanal in Teams bleibt. Archivieren passiert in Teams.</p>'
            : '') +
        '</div></div>' +
        (count
            ? '<div class="da-table-wrap"><table class="dash-hygiene-table da-table">' +
              '<thead><tr><th>Thema / Kanal</th><th>Betreuung</th><th>Schüler</th><th>Status</th><th></th></tr></thead>' +
              '<tbody>' +
              rows +
              '</tbody></table></div>'
            : '') +
        '</section>'
    );
}

function renderArbeitForm(L, cohort, editing, picked, channelPlaceholder) {
    if (!cohort.teamId) return '';
    const existing = editing ? getArbeit(draftPick.editId) : null;
    const buttonLabel =
        existing && existing.channelId ? 'Kanal in Teams aktualisieren' : 'Kanal in Teams anlegen';
    return (
        '<section class="tm-panel da-panel" id="daArbeitForm">' +
        '<div class="tm-panel__head"><i class="bi bi-journal-plus" aria-hidden="true"></i><div>' +
        '<h3>' +
        (editing ? escapeHtml(L.singular) + ' bearbeiten' : 'Neue ' + escapeHtml(L.singular)) +
        '</h3>' +
        '<p>Thema, Kanalname, Betreuung und Schüler:innen.</p></div></div>' +
        '<div class="da-form-grid">' +
        '<div class="tm-field da-field"><label for="daArbeitTopic">Thema / Titel</label><input id="daArbeitTopic" value="' +
        escapeHtml(draftPick.topic) +
        '" placeholder="z. B. KI im Unterricht"></div>' +
        '<div class="tm-field da-field"><label for="daKanalName">Name des privaten Kanals</label><input id="daKanalName" value="' +
        escapeHtml(draftPick.channelName) +
        '" placeholder="' +
        escapeHtml(channelPlaceholder || 'DA – Name / Thema') +
        '" maxlength="50">' +
        '<p class="muted da-field-hint">So heißt der private Kanal in Teams (max. 50 Zeichen). Leer = Vorschlag aus Thema und Personen.</p></div>' +
        '<div class="tm-field da-field"><label>Betreuungslehrer:in</label>' +
        '<div class="da-pick-row"><button type="button" class="fb-btn fb-btn--ghost" id="daPickMentor"><i class="bi bi-search"></i>Betreuung suchen</button></div>' +
        '<div class="da-picked">' +
        picked.mentor +
        '</div></div>' +
        '<div class="tm-field da-field"><label>Schüler:innen</label>' +
        '<div class="da-pick-row"><button type="button" class="fb-btn fb-btn--ghost" id="daPickStudents"><i class="bi bi-search"></i>Schüler:innen suchen</button></div>' +
        '<div class="da-picked">' +
        picked.students +
        '</div></div>' +
        '</div>' +
        '<div class="tm-actions da-form-actions">' +
        '<button type="button" class="fb-btn" id="daArbeitAddChannel"><i class="bi bi-plus-lg"></i>' +
        buttonLabel +
        '</button>' +
        (editing ? '<button type="button" class="fb-btn fb-btn--ghost" id="daArbeitCancelEdit">Abbrechen</button>' : '') +
        '</div></section>'
    );
}

function renderCohortDetail(L) {
    const host = ensureShell();
    const cohortId = (host && host.getAttribute('data-cohort')) || '';
    const cohort = getCohort(cohortId);
    if (!cohort) {
        return (
            '<p class="muted">Jahrgang nicht gefunden.</p>' +
            '<button type="button" class="fb-btn fb-btn--ghost" data-da-view="jahrgaenge">Zurück</button>'
        );
    }
    const state = loadState();
    const arbeiten = state.arbeiten.filter(function (a) {
        return a.cohortId === cohort.id;
    });
    const picked = renderPickedPeople();

    const editing = !!draftPick.editId;
    const channelPlaceholder = suggestChannelName(draftPick.topic, draftPick.students);
    const rows =
        arbeiten
            .map(function (a) {
                return (
                    '<tr' +
                    (draftPick.editId === a.id ? ' class="da-row-active"' : '') +
                    '><td>' +
                    escapeHtml(a.topic || a.student || '–') +
                    '<div class="muted" style="font-size:0.8em">Kanal: ' +
                    escapeHtml(a.channelName || '–') +
                    '</div></td><td>' +
                    escapeHtml(a.mentorName || a.mentorEmail || '–') +
                    '</td><td>' +
                    escapeHtml(
                        (a.studentNames && a.studentNames.filter(Boolean).length
                            ? a.studentNames.filter(Boolean).join(', ')
                            : (a.studentEmails || []).length + ' Person(en)')
                    ) +
                    '</td><td>' +
                    escapeHtml(ARBEIT_STATUS_LABEL[a.status] || a.status) +
                    '</td><td class="da-row-actions">' +
                    '<button type="button" class="fb-btn fb-btn--ghost" data-arbeit-edit="' +
                    escapeHtml(a.id) +
                    '"><i class="bi bi-pencil"></i>Bearbeiten</button> ' +
                    (a.channelId
                        ? (a.channelWebUrl
                              ? '<a class="fb-btn fb-btn--ghost" href="' +
                                escapeHtml(a.channelWebUrl) +
                                '" target="_blank" rel="noopener">Kanal öffnen</a> '
                              : '') +
                          '<button type="button" class="fb-btn" data-arbeit-sync="' +
                          escapeHtml(a.id) +
                          '"><i class="bi bi-cloud-upload"></i>In Teams aktualisieren</button> '
                        : '<button type="button" class="fb-btn" data-arbeit-channel="' +
                          escapeHtml(a.id) +
                          '">Kanal anlegen</button> ') +
                    '<button type="button" class="fb-btn fb-btn--ghost" data-arbeit-done="' +
                    escapeHtml(a.id) +
                    '">Abschließen</button></td></tr>'
                );
            })
            .join('') ||
        '<tr><td colspan="5" class="muted">Noch keine Arbeiten in diesem Jahrgang.</td></tr>';

    return (
        '<div class="da-cohort">' +
        '<p class="da-back"><button type="button" class="fb-btn fb-btn--ghost" data-da-view="jahrgaenge"><i class="bi bi-arrow-left"></i>Alle Jahrgänge</button></p>' +
        renderFlash() +
        '<section class="tm-panel da-panel">' +
        '<div class="tm-panel__head"><i class="bi bi-mortarboard" aria-hidden="true"></i><div>' +
        '<h3>' +
        escapeHtml(L.plural) +
        ' ' +
        escapeHtml(cohort.maturajahrgang) +
        '</h3>' +
        '<p>' +
        (cohort.teamId
            ? 'Team, dann die Kanäle der Arbeiten.'
            : 'Zuerst das Team für diesen Jahrgang.') +
        '</p></div></div>' +
        (cohort.teamId
            ? '<p class="ta-msg ta-msg--ok da-team-banner">Team verknüpft: <strong>' +
              escapeHtml(cohort.displayName || cohort.teamId) +
              '</strong>' +
              (cohort.teamWebUrl
                  ? ' · <a href="' +
                    escapeHtml(cohort.teamWebUrl) +
                    '" target="_blank" rel="noopener">In Teams öffnen</a>'
                  : '') +
              '</p>'
            : renderTeamSetup(cohort)) +
        protocolDetails() +
        '</section>' +
        renderArbeitList(cohort, arbeiten.length, rows) +
        renderArbeitForm(L, cohort, editing, picked, channelPlaceholder) +
        '</div>'
    );
}

function log(msg) {
    appendActionLog(msg);
}

function appendLog(_host, msg) {
    appendActionLog(msg);
}

function bindCohortDetail(host) {
    host.querySelectorAll('[data-da-view]').forEach(function (btn) {
        /* also bound globally via bindTabs; keep for back button inside view */
    });

    const create = host.querySelector('#daCohortCreate');
    if (create) {
        create.addEventListener('click', function () {
            const id = (ensureShell() && ensureShell().getAttribute('data-cohort')) || '';
            resetActionLog();
            create.disabled = true;
            appendLog(host, 'Team wird angelegt …');
            provisionCohortTeam(id, {
                onLog: function (m) {
                    appendLog(host, m);
                }
            })
                .then(function () {
                    const name = (getCohort(id) || {}).displayName || '';
                    setFlash(
                        'ok',
                        name
                            ? 'Team angelegt. In Teams heißt es „' + name + '“.'
                            : 'Team angelegt.'
                    );
                    render();
                })
                .catch(function (e) {
                    setFlash('err', e && e.message ? e.message : String(e));
                    render();
                });
        });
    }

    const searchBtn = host.querySelector('#daTeamSearchBtn');
    if (searchBtn) {
        searchBtn.addEventListener('click', function () {
            void (async function () {
                const q = String((host.querySelector('#daTeamSearch') || {}).value || '');
                draftPick.teamQuery = q;
                if (q.trim().length < 2) {
                    draftPick.teamHits = [];
                    draftPick.teamSearched = false;
                    toast('Mindestens zwei Zeichen.');
                    render();
                    return;
                }
                try {
                    draftPick.teamHits = await searchTeamsGroups(q);
                    draftPick.teamSearched = true;
                    if (!draftPick.teamHits.length) toast('Kein Team mit diesem Namen.');
                } catch (e) {
                    toast(e && e.message ? e.message : String(e));
                    draftPick.teamHits = [];
                    draftPick.teamSearched = true;
                }
                render();
            })();
        });
        const searchInput = host.querySelector('#daTeamSearch');
        if (searchInput) {
            searchInput.addEventListener('keydown', function (ev) {
                if (ev.key !== 'Enter') return;
                ev.preventDefault();
                searchBtn.click();
            });
        }
    }

    host.querySelectorAll('[data-link-team]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const id = (ensureShell() && ensureShell().getAttribute('data-cohort')) || '';
            const gid = btn.getAttribute('data-link-team');
            resetActionLog();
            linkCohortTeamEnsured(id, gid, {
                displayName: btn.getAttribute('data-link-name') || '',
                mailNickname: btn.getAttribute('data-link-nick') || '',
                onLog: function (m) {
                    appendLog(host, m);
                }
            })
                .then(function () {
                    clearTeamSearch();
                    setFlash('ok', 'Team verknüpft.');
                    render();
                })
                .catch(function (e) {
                    setFlash('err', e && e.message ? e.message : String(e));
                    render();
                });
        });
    });

    const guidLink = host.querySelector('#daTeamGuidLink');
    if (guidLink) {
        guidLink.addEventListener('click', function () {
            const id = (ensureShell() && ensureShell().getAttribute('data-cohort')) || '';
            const gid = String((host.querySelector('#daTeamGuid') || {}).value || '').trim();
            if (!gid) {
                toast('Bitte eine Gruppen-ID eintragen.');
                return;
            }
            resetActionLog();
            linkCohortTeamEnsured(id, gid, {
                onLog: function (m) {
                    appendLog(host, m);
                }
            })
                .then(function () {
                    setFlash('ok', 'Team verknüpft.');
                    render();
                })
                .catch(function (e) {
                    setFlash('err', e && e.message ? e.message : String(e));
                    render();
                });
        });
    }

    const topicEl = host.querySelector('#daArbeitTopic');
    const kanalEl = host.querySelector('#daKanalName');
    if (topicEl) {
        topicEl.addEventListener('input', function () {
            draftPick.topic = topicEl.value;
            if (kanalEl && !String(draftPick.channelName || '').trim()) {
                kanalEl.placeholder = suggestChannelName(draftPick.topic, draftPick.students);
            }
        });
    }
    if (kanalEl) {
        kanalEl.addEventListener('input', function () {
            draftPick.channelName = kanalEl.value;
        });
    }

    const pickMentor = host.querySelector('#daPickMentor');
    if (pickMentor) {
        pickMentor.addEventListener('click', function () {
            draftPick.topic = topicEl ? topicEl.value : draftPick.topic;
            draftPick.channelName = kanalEl ? kanalEl.value : draftPick.channelName;
            void pickEntraUser({
                title: 'Betreuungslehrer:in wählen',
                hint: 'Lehrkraft in Microsoft 365 suchen.'
            }).then(function (u) {
                if (!u) return;
                draftPick.mentor = u;
                render();
            });
        });
    }
    const clearMentor = host.querySelector('#daClearMentor');
    if (clearMentor) {
        clearMentor.addEventListener('click', function () {
            draftPick.mentor = null;
            render();
        });
    }
    const pickStudents = host.querySelector('#daPickStudents');
    if (pickStudents) {
        pickStudents.addEventListener('click', function () {
            draftPick.topic = topicEl ? topicEl.value : draftPick.topic;
            draftPick.channelName = kanalEl ? kanalEl.value : draftPick.channelName;
            void pickEntraUsers({
                title: 'Schüler:innen wählen',
                hint: 'Mehrere Personen aus Microsoft 365 anhaken und übernehmen.'
            }).then(function (list) {
                if (!list || !list.length) return;
                const map = new Map(
                    draftPick.students.map(function (s) {
                        return [s.id || s.mail, s];
                    })
                );
                list.forEach(function (u) {
                    map.set(u.id || u.mail, u);
                });
                draftPick.students = Array.from(map.values());
                if (!String(draftPick.channelName || '').trim()) {
                    draftPick.channelName = suggestChannelName(draftPick.topic, draftPick.students);
                }
                render();
            });
        });
    }
    host.querySelectorAll('[data-rm-student]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const i = Number(btn.getAttribute('data-rm-student'));
            draftPick.students.splice(i, 1);
            render();
        });
    });

    async function saveArbeit(andTeams) {
        const cohortId = (ensureShell() && ensureShell().getAttribute('data-cohort')) || '';
        const topic = String((host.querySelector('#daArbeitTopic') || {}).value || '').trim();
        let channelName = String((host.querySelector('#daKanalName') || {}).value || '').trim();
        if (!draftPick.mentor || !draftPick.mentor.mail) {
            setFlash('err', 'Bitte Betreuungslehrer:in aus Microsoft 365 wählen.');
            render();
            return;
        }
        if (!draftPick.students.length) {
            setFlash('err', 'Bitte mindestens eine Schüler:in aus Microsoft 365 wählen.');
            render();
            return;
        }
        const studentLabel = draftPick.students
            .map(function (s) {
                return s.displayName || s.mail;
            })
            .join(' / ');
        if (!channelName) {
            channelName = suggestChannelName(topic || studentLabel, draftPick.students);
        }
        if (!channelName) {
            setFlash('err', 'Bitte einen Kanalnamen eingeben.');
            render();
            return;
        }
        const existing = draftPick.editId ? getArbeit(draftPick.editId) : null;
        const res = upsertArbeit({
            id: draftPick.editId || undefined,
            cohortId: cohortId,
            topic: topic || studentLabel,
            student: studentLabel,
            mentorName: draftPick.mentor.displayName || '',
            mentorEmail: draftPick.mentor.mail || '',
            studentEmails: draftPick.students.map(function (s) {
                return s.mail;
            }),
            studentNames: draftPick.students.map(function (s) {
                return s.displayName || '';
            }),
            channelName: channelName,
            channelId: existing ? existing.channelId : '',
            channelWebUrl: existing ? existing.channelWebUrl : '',
            status: existing ? existing.status : 'geplant'
        });
        if (!res.ok) {
            setFlash('err', res.errors.join(' '));
            render();
            return;
        }

        if (!andTeams) {
            clearArbeitDraft();
            setFlash('ok', 'Arbeit gespeichert.');
            render();
            return;
        }

        resetActionLog();
        try {
            if (res.item.channelId) {
                await syncArbeitChannel(res.item.id, {
                    onLog: function (m) {
                        appendLog(host, m);
                    }
                });
                clearArbeitDraft();
                setFlash('ok', 'Gespeichert und in Teams aktualisiert (Kanalname + Mitglieder).');
            } else {
                await provisionArbeitChannel(res.item.id, {
                    onLog: function (m) {
                        appendLog(host, m);
                    }
                });
                clearArbeitDraft();
                setFlash('ok', 'Arbeit gespeichert und privater Kanal in Teams angelegt.');
            }
        } catch (e) {
            setFlash(
                'err',
                'Lokal gespeichert, Teams-Update fehlgeschlagen: ' +
                    (e && e.message ? e.message : String(e))
            );
        }
        render();
    }

    const addCh = host.querySelector('#daArbeitAddChannel');
    if (addCh) {
        addCh.addEventListener('click', function () {
            addCh.disabled = true;
            void saveArbeit(true);
        });
    }
    const cancelEdit = host.querySelector('#daArbeitCancelEdit');
    if (cancelEdit) {
        cancelEdit.addEventListener('click', function () {
            clearArbeitDraft();
            setFlash(null, '');
            render();
        });
    }

    host.querySelectorAll('[data-arbeit-edit]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const id = btn.getAttribute('data-arbeit-edit');
            const a = getArbeit(id);
            if (!a) return;
            draftPick.editId = a.id;
            draftPick.topic = a.topic || '';
            draftPick.channelName = a.channelName || '';
            draftPick.mentor = a.mentorEmail
                ? { id: '', displayName: a.mentorName || '', mail: a.mentorEmail }
                : null;
            draftPick.students = (a.studentEmails || []).map(function (mail, i) {
                return {
                    id: mail,
                    displayName: (a.studentNames && a.studentNames[i]) || mail,
                    mail: mail
                };
            });
            draftPick.flash = null;
            render();
            const form = document.getElementById('daArbeitForm');
            if (form) form.scrollIntoView({ behavior: 'smooth', block: 'start' });
        });
    });

    host.querySelectorAll('[data-arbeit-channel]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const id = btn.getAttribute('data-arbeit-channel');
            resetActionLog();
            provisionArbeitChannel(id, {
                onLog: function (m) {
                    appendLog(host, m);
                }
            })
                .then(function () {
                    setFlash('ok', 'Privater Kanal angelegt.');
                    render();
                })
                .catch(function (e) {
                    setFlash('err', e && e.message ? e.message : String(e));
                    render();
                });
        });
    });

    host.querySelectorAll('[data-arbeit-sync]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const id = btn.getAttribute('data-arbeit-sync');
            resetActionLog();
            syncArbeitChannel(id, {
                onLog: function (m) {
                    appendLog(host, m);
                }
            })
                .then(function () {
                    setFlash('ok', 'Kanal in Teams aktualisiert.');
                    render();
                })
                .catch(function (e) {
                    setFlash('err', e && e.message ? e.message : String(e));
                    render();
                });
        });
    });

    host.querySelectorAll('[data-arbeit-done]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const id = btn.getAttribute('data-arbeit-done');
            if (!getArbeit(id)) return;
            patchArbeit(id, { status: 'abgeschlossen' });
            setFlash('ok', 'In der Übersicht als abgeschlossen markiert. Der Kanal in Teams bleibt.');
            render();
        });
    });
}

function renderLegacy(L) {
    return (
        renderFlash() +
        '<section class="tm-panel">' +
        '<div class="tm-panel__head"><i class="bi bi-plus-circle" aria-hidden="true"></i><div>' +
        '<h3>Legacy: eigene ' +
        escapeHtml(L.singular) +
        '-Gruppe</h3>' +
        '<p>Wie bisher: eine Gruppe/Team pro Arbeit.</p></div></div>' +
        '<div class="tm-field"><label for="daYear">Jahr</label><input type="text" id="daYear" inputmode="numeric"></div>' +
        '<div class="tm-field"><label for="daTopic">Thema</label><input type="text" id="daTopic" placeholder="z. B. KI im Unterricht"></div>' +
        '<div class="tm-field"><label for="daStudent">Schüler:in (optional)</label><input type="text" id="daStudent"></div>' +
        '<div class="tm-field"><label for="daMentor">Betreuung (optional)</label><input type="text" id="daMentor"></div>' +
        '<label class="checkbox-label" style="display:flex;gap:8px;margin:8px 0"><input type="checkbox" id="daGroupOnly"><span>Nur Gruppe (kein Team)</span></label>' +
        '<div id="daPreview" class="tm-log" style="min-height:60px;margin:10px 0"></div>' +
        '<div class="tm-actions"><button type="button" class="fb-btn" id="daBtnCreate"><i class="bi bi-plus-lg"></i>Anlegen</button></div>' +
        protocolDetails() +
        '</section>' +
        '<section class="tm-panel">' +
        '<div class="tm-panel__head"><i class="bi bi-search" aria-hidden="true"></i><div>' +
        '<h3>Bestehende Gruppen</h3>' +
        '<p>Die Suche filtert die geladene Liste. Es werden nicht alle Gruppen der Schule gelesen.</p></div></div>' +
        '<div class="tm-field"><label for="daFilter">Suche</label><input type="search" id="daFilter"></div>' +
        '<div class="tm-actions"><button type="button" class="fb-btn fb-btn--ghost" id="daBtnList"><i class="bi bi-arrow-clockwise"></i>Diplom- und VWA-Teams laden</button></div>' +
        '<p id="daListStatus" class="muted"></p>' +
        '<div style="overflow:auto"><table class="dash-hygiene-table" style="width:100%;font-size:0.92em">' +
        '<thead><tr><th>Name</th><th>Alias</th><th>Mail</th></tr></thead>' +
        '<tbody id="daBody"><tr><td colspan="3" class="muted">Noch nicht geladen.</td></tr></tbody></table></div>' +
        '</section>'
    );
}

function bindLegacy() {
    ['daYear', 'daTopic', 'daStudent', 'daMentor', 'daGroupOnly'].forEach(function (id) {
        const el = $(id);
        if (el) el.addEventListener('input', refreshLegacyPreview);
        if (el) el.addEventListener('change', refreshLegacyPreview);
    });
    refreshLegacyPreview();
    const create = $('daBtnCreate');
    if (create) {
        create.addEventListener('click', function () {
            createLegacyDiplom().catch(function (e) {
                setFlash('err', e && e.message ? e.message : String(e));
            });
        });
    }
    const list = $('daBtnList');
    if (list) {
        list.addEventListener('click', function () {
            loadDiplomList().catch(function (e) {
                toast(e && e.message ? e.message : String(e));
            });
        });
    }
    const filter = $('daFilter');
    if (filter) {
        filter.addEventListener('input', function () {
            paintLegacyList();
        });
    }
    paintLegacyList();
}

function bindTabs(host) {
    host.querySelectorAll('[data-da-view]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const next = btn.getAttribute('data-da-view') || 'jahrgaenge';
            host.setAttribute('data-view', next);
            if (next !== 'cohort') host.removeAttribute('data-cohort');
            clearTeamSearch();
            render();
        });
    });
}

function updateChrome() {
    const L = labels();
    const name = document.querySelector('.header-tool-indicator__name');
    if (name) {
        name.innerHTML = '<i class="bi bi-mortarboard"></i>' + escapeHtml(L.toolTitle);
    }
    const title = document.querySelector('title');
    if (title) title.textContent = 'MS365-Schul-Tools – ' + L.toolTitle;
}

document.addEventListener('DOMContentLoaded', function () {
    updateChrome();
    if ($('daApp')) {
        render();
        void pullDiplomSettingsFromSchool().then(function (result) {
            if (result && result.applied) {
                setupDraft = null;
                schoolNote = 'Wahl der Schule übernommen.';
                updateChrome();
                render();
            }
        });
        return;
    }
    /* Fallback: alte HTML-Struktur ohne #daApp */
    bindLegacy();
});
