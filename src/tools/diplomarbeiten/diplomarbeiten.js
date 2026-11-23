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
    filterDiplomGroups,
    labelsForArbeitTyp
} from './diplomarbeiten-logic.js';
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
    editId: '',
    topic: '',
    channelName: '',
    flash: null
};

function setFlash(type, text) {
    draftPick.flash = text ? { type: type || 'ok', text: String(text) } : null;
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

function log(msg) {
    const el = $('daLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + String(msg || '');
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
    toast(labels().singular + '-Gruppe angelegt.');
    await loadDiplomList();
}

async function loadDiplomList() {
    const G = window.ms365GraphUnifiedGroups;
    if (!G) throw new Error('Graph-Modul nicht geladen.');
    const status = $('daListStatus');
    if (status) status.textContent = 'Lade Gruppen …';
    const token = await G.getGraphToken([
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.Read.All',
        'https://graph.microsoft.com/Directory.Read.All'
    ]);
    let next = '/groups?$select=id,displayName,mailNickname,mail,description&$top=999';
    const all = [];
    let page = 0;
    while (next && page < 30) {
        page++;
        const data = await sharedGraphJson('GET', next, token, undefined);
        if (Array.isArray(data.value)) all.push.apply(all, data.value);
        next = data['@odata.nextLink'] || null;
        if (status) status.textContent = 'Gruppen: ' + all.length;
    }
    const q = ($('daFilter') && $('daFilter').value) || '';
    const rows = filterDiplomGroups(all, q);
    const body = $('daBody');
    if (!body) return;
    if (!rows.length) {
        body.innerHTML = '<tr><td colspan="3" class="muted">Keine Treffer.</td></tr>';
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
        host.innerHTML = renderSetup(L, s);
        bindSetup(host);
        return;
    }

    let view = host.getAttribute('data-view') || (s.modus === 'legacy' ? 'legacy' : 'jahrgaenge');
    // „Arbeiten“-Tab ohne Kontext → Jahrgangsliste
    if (view === 'arbeiten' && !host.getAttribute('data-cohort')) view = 'jahrgaenge';
    const tabView = view === 'cohort' ? 'jahrgaenge' : view;
    host.innerHTML =
        renderTabs(L, s, tabView) +
        '<div class="da-view">' +
        (view === 'setup'
            ? renderSetup(L, s)
            : view === 'legacy'
              ? renderLegacy(L)
              : view === 'cohort'
                ? renderCohortDetail(L)
                : renderJahrgaenge(L)) +
        '</div>';
    bindTabs(host);
    if (view === 'setup') bindSetup(host);
    else if (view === 'legacy') bindLegacy(host);
    else if (view === 'cohort') bindCohortDetail(host);
    else bindJahrgaenge(host);
}

function renderTabs(L, s, view) {
    const tabs = [
        { id: 'jahrgaenge', label: 'Jahrgänge', show: s.modus === 'channels' },
        { id: 'legacy', label: 'Legacy-Teams', show: true },
        { id: 'setup', label: 'Setup', show: true }
    ];
    return (
        '<div class="da-tabs" role="tablist">' +
        tabs
            .filter((t) => t.show)
            .map(function (t) {
                return (
                    '<button type="button" class="btn' +
                    (view === t.id ? ' btn-success' : '') +
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
        (s.modus === 'channels' ? 'Dach-Team + Kanäle' : 'Legacy') +
        '</span></div>'
    );
}

function renderSetup(L, s) {
    return (
        '<section class="tm-panel">' +
        '<div class="tm-panel__head"><i class="bi bi-sliders" aria-hidden="true"></i><div>' +
        '<h3>Setup – ' +
        escapeHtml(L.toolTitle) +
        '</h3>' +
        '<p>Wizard-Umschalter: Standard (Dach-Team pro Maturajahrgang) oder Legacy (eigene Teams).</p></div></div>' +
        '<div class="tm-field"><label for="daArbeitTyp">Abschlussarbeit-Typ (Schulform)</label>' +
        '<select id="daArbeitTyp">' +
        '<option value="diplom"' +
        (s.arbeitTyp === 'diplom' ? ' selected' : '') +
        '>BHS – Diplomarbeit</option>' +
        '<option value="vwa"' +
        (s.arbeitTyp === 'vwa' ? ' selected' : '') +
        '>AHS – VWA / ABA</option>' +
        '</select></div>' +
        '<div class="tm-field"><label for="daModus">Betriebsmodus</label>' +
        '<select id="daModus">' +
        '<option value="channels"' +
        (s.modus === 'channels' ? ' selected' : '') +
        '>Standard: Dach-Team + private Kanäle (pro Maturajahrgang)</option>' +
        '<option value="legacy"' +
        (s.modus === 'legacy' ? ' selected' : '') +
        '>Legacy: eigene Teams/Gruppen pro Arbeit</option>' +
        '</select></div>' +
        '<div class="tm-actions">' +
        '<button type="button" class="btn btn-success" id="daSetupSave"><i class="bi bi-check2"></i>Speichern &amp; starten</button>' +
        '</div></section>'
    );
}

function bindSetup(host) {
    const btn = host.querySelector('#daSetupSave') || $('daSetupSave');
    if (!btn) return;
    btn.addEventListener('click', function () {
        const typEl = host.querySelector('#daArbeitTyp') || $('daArbeitTyp');
        const modusEl = host.querySelector('#daModus') || $('daModus');
        saveDiplomSettings({
            arbeitTyp: typEl ? typEl.value : 'diplom',
            modus: modusEl ? modusEl.value : 'channels',
            setupDone: true
        });
        const h = ensureShell();
        if (h) {
            h.setAttribute(
                'data-view',
                (modusEl && modusEl.value === 'legacy' ? 'legacy' : 'jahrgaenge')
            );
        }
        updateChrome();
        toast('Setup gespeichert.');
        render();
    });
}

function openCohort(id) {
    const h = ensureShell();
    if (!h) return;
    clearArbeitDraft();
    draftPick.teamHits = [];
    draftPick.flash = null;
    h.setAttribute('data-view', 'cohort');
    h.setAttribute('data-cohort', id || '');
    render();
}

function renderJahrgaenge(L) {
    const state = loadState();
    const rows =
        state.cohorts
            .map(function (c) {
                const nArbeiten = state.arbeiten.filter(function (a) {
                    return a.cohortId === c.id;
                }).length;
                return (
                    '<tr><td><strong>' +
                    escapeHtml(c.maturajahrgang) +
                    '</strong></td><td>' +
                    escapeHtml(c.displayName || '–') +
                    '</td><td>' +
                    (c.teamId
                        ? '<span class="muted">Team verknüpft</span>'
                        : '<span class="muted">kein Team</span>') +
                    '</td><td>' +
                    nArbeiten +
                    '</td><td>' +
                    '<button type="button" class="btn btn-sm btn-success" data-cohort-open="' +
                    escapeHtml(c.id) +
                    '"><i class="bi bi-box-arrow-in-right"></i>Öffnen</button>' +
                    '</td></tr>'
                );
            })
            .join('') || '<tr><td colspan="5" class="muted">Noch kein Jahrgang.</td></tr>';

    return (
        '<section class="tm-panel">' +
        '<div class="tm-panel__head"><i class="bi bi-calendar3" aria-hidden="true"></i><div>' +
        '<h3>Maturajahrgänge</h3>' +
        '<p>Jahrgang öffnen → Dach-Team (M365-Gruppe mit Teams) verbinden → Arbeiten &amp; private Kanäle.</p></div></div>' +
        '<div class="tm-field"><label for="daCohortJahr">Neuer Maturajahrgang</label>' +
        '<input id="daCohortJahr" placeholder="2026/27"></div>' +
        '<div class="tm-actions"><button type="button" class="btn btn-success" id="daCohortAdd"><i class="bi bi-plus-lg"></i>Jahrgang anlegen</button></div>' +
        '<div style="overflow:auto;margin-top:12px"><table class="dash-hygiene-table" style="width:100%;font-size:0.92em">' +
        '<thead><tr><th>Jahrgang</th><th>Team-Name</th><th>Team</th><th>Arbeiten</th><th></th></tr></thead>' +
        '<tbody>' +
        rows +
        '</tbody></table></div>' +
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
            toast('Jahrgang angelegt – jetzt öffnen und Team verbinden.');
            openCohort(res.item.id);
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

function renderCohortDetail(L) {
    const host = ensureShell();
    const cohortId = (host && host.getAttribute('data-cohort')) || '';
    const cohort = getCohort(cohortId);
    if (!cohort) {
        return (
            '<p class="muted">Jahrgang nicht gefunden.</p>' +
            '<button type="button" class="btn" data-da-view="jahrgaenge">Zurück</button>'
        );
    }
    const state = loadState();
    const arbeiten = state.arbeiten.filter(function (a) {
        return a.cohortId === cohort.id;
    });
    const picked = renderPickedPeople();
    const teamHits = (draftPick.teamHits || [])
        .map(function (g) {
            return (
                '<button type="button" class="btn btn-sm" data-link-team="' +
                escapeHtml(g.id) +
                '" data-link-name="' +
                escapeHtml(g.displayName) +
                '" data-link-nick="' +
                escapeHtml(g.mailNickname) +
                '">' +
                escapeHtml(g.displayName) +
                (g.hasTeam
                    ? ' <span class="muted">(Team)</span>'
                    : ' <span class="muted">(Gruppe → Team wird ergänzt)</span>') +
                '</button>'
            );
        })
        .join(' ');

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
                    '<button type="button" class="btn btn-sm" data-arbeit-edit="' +
                    escapeHtml(a.id) +
                    '"><i class="bi bi-pencil"></i>Bearbeiten</button> ' +
                    (a.channelId
                        ? (a.channelWebUrl
                              ? '<a class="btn btn-sm" href="' +
                                escapeHtml(a.channelWebUrl) +
                                '" target="_blank" rel="noopener">Kanal öffnen</a> '
                              : '') +
                          '<button type="button" class="btn btn-sm btn-success" data-arbeit-sync="' +
                          escapeHtml(a.id) +
                          '"><i class="bi bi-cloud-upload"></i>In Teams aktualisieren</button> '
                        : '<button type="button" class="btn btn-sm btn-success" data-arbeit-channel="' +
                          escapeHtml(a.id) +
                          '">Kanal anlegen</button> ') +
                    '<button type="button" class="btn btn-sm" data-arbeit-done="' +
                    escapeHtml(a.id) +
                    '">Abschließen</button></td></tr>'
                );
            })
            .join('') ||
        '<tr><td colspan="5" class="muted">Noch keine Arbeiten in diesem Jahrgang.</td></tr>';

    return (
        '<div class="da-cohort">' +
        '<p class="da-back"><button type="button" class="btn" data-da-view="jahrgaenge"><i class="bi bi-arrow-left"></i>Alle Jahrgänge</button></p>' +
        renderFlash() +
        '<section class="tm-panel da-panel">' +
        '<div class="tm-panel__head"><i class="bi bi-mortarboard" aria-hidden="true"></i><div>' +
        '<h3>' +
        escapeHtml(L.plural) +
        ' ' +
        escapeHtml(cohort.maturajahrgang) +
        '</h3>' +
        '<p>Dach-Team = Microsoft-365-Gruppe <strong>mit Teams-Funktionalität</strong>.</p></div></div>' +
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
            : '<div class="ta-detail da-team-setup">' +
              '<p><strong>Noch kein Dach-Team.</strong> Neues Team anlegen oder bestehendes suchen und verbinden.</p>' +
              '<div class="tm-actions">' +
              '<button type="button" class="btn btn-success" id="daCohortCreate"><i class="bi bi-plus-lg"></i>Neues Team anlegen</button>' +
              '</div>' +
              '<div class="tm-field da-field"><label for="daTeamSearch">Bestehendes Team / M365-Gruppe suchen</label>' +
              '<div class="tm-actions da-inline">' +
              '<input id="daTeamSearch" placeholder="z. B. Diplomarbeiten 2026">' +
              '<button type="button" class="btn" id="daTeamSearchBtn">Suchen</button></div>' +
              (teamHits ? '<div class="tm-actions da-hit-row">' + teamHits + '</div>' : '') +
              '</div>' +
              '<div class="tm-field da-field"><label for="daTeamGuid">Oder Gruppen-/Team-ID (GUID)</label>' +
              '<div class="tm-actions da-inline">' +
              '<input id="daTeamGuid" placeholder="xxxxxxxx-xxxx-…" spellcheck="false">' +
              '<button type="button" class="btn" id="daTeamGuidLink">Verbinden</button></div></div>' +
              '</div>') +
        '<details class="tm-tech da-log"><summary>Protokoll</summary><pre id="daLog" class="tm-log" style="white-space:pre-wrap"></pre></details>' +
        '</section>' +
        '<section class="tm-panel da-panel" id="daArbeitForm">' +
        '<div class="tm-panel__head"><i class="bi bi-journal-plus" aria-hidden="true"></i><div>' +
        '<h3>' +
        (editing ? escapeHtml(L.singular) + ' bearbeiten' : 'Neue ' + escapeHtml(L.singular)) +
        '</h3>' +
        '<p>Thema, <strong>Kanalname</strong>, Betreuung und Schüler:innen – danach speichern und optional in Teams schreiben.</p></div></div>' +
        (!cohort.teamId
            ? '<p class="muted">Zuerst Dach-Team verbinden, dann Arbeiten anlegen.</p>'
            : '') +
        '<div class="da-form-grid">' +
        '<div class="tm-field da-field"><label for="daArbeitTopic">Thema / Titel</label><input id="daArbeitTopic" value="' +
        escapeHtml(draftPick.topic) +
        '" ' +
        (cohort.teamId ? '' : 'disabled ') +
        'placeholder="z. B. KI im Unterricht"></div>' +
        '<div class="tm-field da-field"><label for="daKanalName">Name des privaten Kanals</label><input id="daKanalName" value="' +
        escapeHtml(draftPick.channelName) +
        '" ' +
        (cohort.teamId ? '' : 'disabled ') +
        'placeholder="' +
        escapeHtml(channelPlaceholder || 'DA – Name / Thema') +
        '" maxlength="50">' +
        '<p class="muted da-field-hint">So heißt der private Kanal in Teams (max. 50 Zeichen). Leer = Vorschlag aus Thema/Personen.</p></div>' +
        '<div class="tm-field da-field"><label>Betreuungslehrer:in</label>' +
        '<div class="da-pick-row"><button type="button" class="btn" id="daPickMentor" ' +
        (cohort.teamId ? '' : 'disabled ') +
        '><i class="bi bi-search"></i>Betreuung suchen</button></div>' +
        '<div class="da-picked">' +
        picked.mentor +
        '</div></div>' +
        '<div class="tm-field da-field"><label>Schüler:innen</label>' +
        '<div class="da-pick-row"><button type="button" class="btn" id="daPickStudents" ' +
        (cohort.teamId ? '' : 'disabled ') +
        '><i class="bi bi-search"></i>Schüler:innen suchen</button></div>' +
        '<div class="da-picked">' +
        picked.students +
        '</div></div>' +
        '</div>' +
        '<div class="tm-actions da-form-actions">' +
        '<button type="button" class="btn btn-success" id="daArbeitAdd" ' +
        (cohort.teamId ? '' : 'disabled ') +
        '><i class="bi bi-check2"></i>' +
        (editing ? 'Änderungen speichern' : 'Arbeit speichern') +
        '</button>' +
        '<button type="button" class="btn" id="daArbeitAddChannel" ' +
        (cohort.teamId ? '' : 'disabled ') +
        '><i class="bi bi-plus-lg"></i>' +
        (editing && getArbeit(draftPick.editId) && getArbeit(draftPick.editId).channelId
            ? 'Speichern &amp; in Teams aktualisieren'
            : 'Speichern &amp; privaten Kanal anlegen') +
        '</button>' +
        (editing
            ? '<button type="button" class="btn" id="daArbeitCancelEdit">Abbrechen</button>'
            : '') +
        '</div></section>' +
        '<section class="tm-panel da-panel">' +
        '<div class="tm-panel__head"><i class="bi bi-list-ul" aria-hidden="true"></i><div>' +
        '<h3>Arbeiten in ' +
        escapeHtml(cohort.maturajahrgang) +
        '</h3>' +
        '<p>' +
        arbeiten.length +
        ' Einträge – bearbeiten, Kanal anlegen/aktualisieren</p></div></div>' +
        '<div class="da-table-wrap"><table class="dash-hygiene-table da-table">' +
        '<thead><tr><th>Thema / Kanal</th><th>Betreuung</th><th>Schüler</th><th>Status</th><th></th></tr></thead>' +
        '<tbody>' +
        rows +
        '</tbody></table></div></section></div>'
    );
}

function appendLog(host, msg) {
    const logEl = host.querySelector('#daLog');
    if (logEl) logEl.textContent += (logEl.textContent ? '\n' : '') + msg;
}

function bindCohortDetail(host) {
    host.querySelectorAll('[data-da-view]').forEach(function (btn) {
        /* also bound globally via bindTabs; keep for back button inside view */
    });

    const create = host.querySelector('#daCohortCreate');
    if (create) {
        create.addEventListener('click', function () {
            const id = (ensureShell() && ensureShell().getAttribute('data-cohort')) || '';
            const logEl = host.querySelector('#daLog');
            if (logEl) logEl.textContent = '';
            provisionCohortTeam(id, {
                onLog: function (m) {
                    appendLog(host, m);
                }
            })
                .then(function () {
                    toast('Dach-Team (M365-Gruppe + Teams) angelegt.');
                    render();
                })
                .catch(function (e) {
                    toast(e && e.message ? e.message : String(e));
                });
        });
    }

    const searchBtn = host.querySelector('#daTeamSearchBtn');
    if (searchBtn) {
        searchBtn.addEventListener('click', function () {
            void (async function () {
                const q = String((host.querySelector('#daTeamSearch') || {}).value || '');
                try {
                    draftPick.teamHits = await searchTeamsGroups(q);
                    if (!draftPick.teamHits.length) toast('Keine Treffer.');
                    else toast(draftPick.teamHits.length + ' Treffer – bitte auswählen.');
                } catch (e) {
                    toast(e && e.message ? e.message : String(e));
                    draftPick.teamHits = [];
                }
                render();
            })();
        });
    }

    host.querySelectorAll('[data-link-team]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const id = (ensureShell() && ensureShell().getAttribute('data-cohort')) || '';
            const gid = btn.getAttribute('data-link-team');
            const logEl = host.querySelector('#daLog');
            if (logEl) logEl.textContent = '';
            linkCohortTeamEnsured(id, gid, {
                displayName: btn.getAttribute('data-link-name') || '',
                mailNickname: btn.getAttribute('data-link-nick') || '',
                onLog: function (m) {
                    appendLog(host, m);
                }
            })
                .then(function () {
                    draftPick.teamHits = [];
                    toast('Team verknüpft (inkl. Teams-Funktionalität).');
                    render();
                })
                .catch(function (e) {
                    toast(e && e.message ? e.message : String(e));
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
            const logEl = host.querySelector('#daLog');
            if (logEl) logEl.textContent = '';
            linkCohortTeamEnsured(id, gid, {
                onLog: function (m) {
                    appendLog(host, m);
                }
            })
                .then(function () {
                    toast('Team verknüpft.');
                    render();
                })
                .catch(function (e) {
                    toast(e && e.message ? e.message : String(e));
                });
        });
    }

    const flashClose = host.querySelector('#daFlashClose');
    if (flashClose) {
        flashClose.addEventListener('click', function () {
            draftPick.flash = null;
            render();
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

        const logEl = host.querySelector('#daLog');
        if (logEl) logEl.textContent = '';
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

    const add = host.querySelector('#daArbeitAdd');
    if (add) {
        add.addEventListener('click', function () {
            void saveArbeit(false);
        });
    }
    const addCh = host.querySelector('#daArbeitAddChannel');
    if (addCh) {
        addCh.addEventListener('click', function () {
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
            const logEl = host.querySelector('#daLog');
            if (logEl) logEl.textContent = '';
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
            const logEl = host.querySelector('#daLog');
            if (logEl) logEl.textContent = '';
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
            setFlash('ok', 'Als abgeschlossen markiert.');
            render();
        });
    });
}

function renderLegacy(L) {
    return (
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
        '<div class="tm-actions"><button type="button" class="btn btn-success" id="daBtnCreate"><i class="bi bi-plus-lg"></i>Anlegen</button></div>' +
        '<details class="tm-tech" style="margin-top:12px"><summary>Protokoll</summary><pre id="daLog" class="tm-log" style="white-space:pre-wrap"></pre></details>' +
        '</section>' +
        '<section class="tm-panel">' +
        '<div class="tm-panel__head"><i class="bi bi-search" aria-hidden="true"></i><div>' +
        '<h3>Bestehende Gruppen</h3>' +
        '<p>Filter nach <code>dipl-</code> / <code>vwa-</code> bzw. Anzeigenamen.</p></div></div>' +
        '<div class="tm-field"><label for="daFilter">Suche</label><input type="search" id="daFilter"></div>' +
        '<div class="tm-actions"><button type="button" class="btn" id="daBtnList"><i class="bi bi-arrow-clockwise"></i>Einlesen</button></div>' +
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
                toast(e && e.message ? e.message : String(e));
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
            /* list already loaded – re-filter would need cache; ignore for MVP */
        });
    }
}

function bindTabs(host) {
    host.querySelectorAll('[data-da-view]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const next = btn.getAttribute('data-da-view') || 'jahrgaenge';
            host.setAttribute('data-view', next);
            if (next !== 'cohort') host.removeAttribute('data-cohort');
            draftPick.teamHits = [];
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
        return;
    }
    /* Fallback: alte HTML-Struktur ohne #daApp */
    bindLegacy();
});
