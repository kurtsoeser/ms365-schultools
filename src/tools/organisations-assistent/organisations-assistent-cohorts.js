import { getEl } from '../../shared/utils/dom.js';
import { dlgConfirm } from '../../shared/utils/dialog.js';
import {
    buildCohortRows,
    cohortDisplayName,
    cohortMailNickname
} from './organisations-assistent-logic.js';

function gug() {
    const G = window.ms365GraphUnifiedGroups;
    if (!G) throw new Error('graph-unified-groups.js muss vor diesem Skript geladen werden.');
    return G;
}

function live() {
    const L = window.ms365SlgLiveDetails;
    if (!L) throw new Error('slg-live-details.js muss vor diesem Skript geladen werden.');
    return L;
}

function gd() {
    const G = window.ms365GroupDetail;
    if (!G) throw new Error('group-detail.js muss vor diesem Skript geladen werden.');
    return G;
}

function dataV2() {
    return window.ms365AppDataV2 || null;
}

function normStr(v) {
    return String(v ?? '').trim();
}

function normEmail(v) {
    return normStr(v).toLowerCase();
}

function toast(msg) {
    const el = document.getElementById('toast');
    if (el) {
        el.textContent = msg;
        el.classList.add('show');
        clearTimeout(toast._t);
        toast._t = setTimeout(function () {
            el.classList.remove('show');
        }, 3800);
        return;
    }
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
}

function isDirektionRole(roleRaw) {
    const r = normStr(roleRaw).toLowerCase();
    return !!r && (r.indexOf('direktion') !== -1 || r.indexOf('direktor') !== -1);
}

function readDirektion() {
    const out = [];
    const seen = new Set();
    function add(email) {
        const em = normEmail(email);
        if (!em || em.indexOf('@') === -1 || seen.has(em)) return;
        seen.add(em);
        out.push(em);
    }
    const settings = typeof window.ms365TenantSettingsLoad === 'function' ? window.ms365TenantSettingsLoad() : null;
    const adminTs = settings && Array.isArray(settings.admin) ? settings.admin : [];
    adminTs.forEach(function (row) {
        if (!isDirektionRole(row && row.role)) return;
        add(row && row.email);
    });
    if (!out.length) {
        const api = dataV2();
        const c = api && typeof api.getContainer === 'function' ? api.getContainer() : null;
        const admin = c && c.core && Array.isArray(c.core.admin) ? c.core.admin : [];
        admin.forEach(function (row) {
            if (!isDirektionRole(row && row.role)) return;
            add(row && row.email);
        });
    }
    return out;
}

function currentClasses() {
    const api = dataV2();
    if (!api || typeof api.getContainer !== 'function') return { year: '', classes: [] };
    const c = api.getContainer();
    const year = c && c.years ? normStr(c.years.current) : '';
    const bucket = year && c.years && c.years.byLabel ? c.years.byLabel[year] : null;
    const classes = bucket && Array.isArray(bucket.classes) ? bucket.classes : [];
    return { year: year, classes: classes };
}

function catalogLinks() {
    const api = dataV2();
    const su = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
    return su && Array.isArray(su.catalogLinks) ? su.catalogLinks : [];
}

function legacyPlans() {
    const api = dataV2();
    if (!api || typeof api.getContainer !== 'function') return [];
    const c = api.getContainer();
    const st =
        c && c.structure && c.structure.settings && c.structure.settings.organisationAssist
            ? c.structure.settings.organisationAssist
            : {};
    return Array.isArray(st.cohortPlans) ? st.cohortPlans : [];
}

/** @type {string} */
let activeYear = '';
/** @type {string[]} */
let direktion = [];
let mounted = false;

function rows() {
    const ctx = currentClasses();
    return buildCohortRows(ctx.classes, catalogLinks(), ctx.year, legacyPlans());
}

function getActiveRow() {
    const y = normStr(activeYear);
    const list = rows();
    for (let i = 0; i < list.length; i++) {
        if (list[i].year === y) return list[i];
    }
    return null;
}

function getActiveGroupId() {
    const api = dataV2();
    if (!api || typeof api.getCatalogLink !== 'function') return null;
    const link = api.getCatalogLink('cohort', activeYear);
    const id = link && link.graphGroupId ? String(link.graphGroupId).trim() : '';
    return id || null;
}

function phaseLabel(phase) {
    if (phase === 'gerade-abgeschlossen') return 'gerade abgeschlossen';
    if (phase === 'aktuell') return 'aktueller Abschlussjahrgang';
    if (phase === 'vergangen') return 'früher';
    if (phase === 'kommend') return 'später';
    return '';
}

function renderLeftList() {
    const host = getEl('oaCohortListItems');
    const summary = getEl('oaCohortSummary');
    const empty = getEl('oaCohortEmptyHint');
    const wrap = getEl('oaCohortDetailWrap');
    if (!host) return;
    host.replaceChildren();
    const list = rows();
    let matchedN = 0;
    list.forEach(function (row) {
        if (row.graphGroupId) matchedN += 1;
    });
    if (summary) {
        summary.textContent =
            String(list.length) +
            ' Abschlussjahr' +
            (list.length === 1 ? '' : 'e') +
            ' · ' +
            String(matchedN) +
            ' mit Microsoft-365-Gruppe';
    }
    const hasRows = list.length > 0;
    if (empty) empty.style.display = hasRows ? 'none' : '';
    if (wrap) wrap.style.display = hasRows ? '' : 'none';

    if (!list.length) {
        const li = document.createElement('li');
        const p = document.createElement('p');
        p.className = 'muted';
        p.style.margin = '10px 12px';
        p.textContent = 'Keine Abschlussjahre in der Klassenliste des aktuellen Schuljahrs.';
        li.appendChild(p);
        host.appendChild(li);
        return;
    }

    list.forEach(function (row) {
        const li = document.createElement('li');
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.setAttribute('data-oa-cohort-year', row.year);
        if (row.year === activeYear) btn.setAttribute('aria-current', 'true');
        const main = document.createElement('span');
        main.className = 'slg-side-main';
        const t = document.createElement('span');
        t.className = 'slg-side-title';
        t.textContent = row.displayName;
        const meta = document.createElement('span');
        meta.className = 'muted slg-side-meta';
        const bits = [];
        const pl = phaseLabel(row.phase);
        if (pl) bits.push(pl);
        if (row.classCount) bits.push(String(row.classCount) + ' Klasse(n)');
        bits.push(row.graphGroupId ? 'Gruppe verknüpft' : 'noch keine Gruppe');
        meta.textContent = bits.join(' · ');
        main.appendChild(t);
        main.appendChild(meta);
        btn.appendChild(main);
        li.appendChild(btn);
        host.appendChild(li);
    });
}

function renderOwnerPreview() {
    const el = document.getElementById('slgOwnerPreview');
    if (!el) return;
    el.replaceChildren();
    if (!direktion.length) {
        const p = document.createElement('p');
        p.style.margin = '0';
        p.style.color = '#6c757d';
        p.textContent = 'Keine Direktion in den Stammdaten gefunden.';
        el.appendChild(p);
        return;
    }
    direktion.forEach(function (em) {
        const d = document.createElement('div');
        d.textContent = em;
        d.style.padding = '4px 0';
        d.style.borderBottom = '1px solid #eef1f4';
        el.appendChild(d);
    });
}

function renderMemberPreview() {
    const el = document.getElementById('slgMemberPreview');
    if (!el) return;
    el.replaceChildren();
    const row = getActiveRow();
    const p = document.createElement('p');
    p.style.margin = '0';
    p.style.color = '#6c757d';
    if (row && row.classLabels.length) {
        p.textContent =
            'Klassen dieses Abschlussjahrs: ' +
            row.classLabels.join(', ') +
            '. Mitglieder nach dem Match in Microsoft 365 pflegen (kein automatischer Listen-Sync).';
    } else {
        p.textContent =
            'Keine Klassen mit diesem Abschlussjahr in der aktuellen Liste. Die Gruppe kann trotzdem angelegt oder verknüpft werden.';
    }
    el.appendChild(p);
}

function applyCreateDefaults() {
    const row = getActiveRow();
    const y = row ? row.year : activeYear;
    const dn = document.getElementById('slgNewDisplayName');
    const nn = document.getElementById('slgNewMailNick');
    const desc = document.getElementById('slgNewDescription');
    if (dn) dn.value = row && row.displayName ? row.displayName : cohortDisplayName(y);
    if (nn) nn.value = row && row.mailNickname ? row.mailNickname : cohortMailNickname(y);
    if (desc) {
        desc.value = cohortDisplayName(y) + ' (MS365-Schulverwaltung)';
    }
    const search = document.getElementById('slgGroupSearch');
    if (search && !normStr(search.value)) {
        search.value = (row && row.mailNickname) || cohortMailNickname(y) || '';
    }
}

function refreshMatchUi() {
    const gid = getActiveGroupId();
    const row = getActiveRow();
    const title = document.getElementById('slgDetailTitle');
    if (title) title.textContent = row ? row.displayName : 'Abschlussjahrgang';
    live().resetCaches();
    live().setMatchedMode(!!gid);
    live().fillForm(gid ? { id: gid } : null);
    renderOwnerPreview();
    renderMemberPreview();
}

function ensureActiveYear() {
    const list = rows();
    if (!list.length) {
        activeYear = '';
        return;
    }
    const has = list.some(function (r) {
        return r.year === activeYear;
    });
    if (has) return;
    const prefer =
        list.find(function (r) {
            return r.phase === 'gerade-abgeschlossen';
        }) ||
        list.find(function (r) {
            return r.phase === 'aktuell';
        }) ||
        list[0];
    activeYear = prefer.year;
}

function setActiveYear(year) {
    activeYear = String(year || '').trim();
    const search = document.getElementById('slgGroupSearch');
    if (search) search.value = '';
    gd().clearSearchResults();
    renderLeftList();
    applyCreateDefaults();
    gd().setTab('general');
    refreshMatchUi();
    if (getActiveGroupId()) live().loadGroup({ silent: true });
}

function persistMatch(g, mode) {
    const api = dataV2();
    const row = getActiveRow();
    if (api && typeof api.upsertCatalogLink === 'function') {
        api.upsertCatalogLink({
            kind: 'cohort',
            code: activeYear,
            graphGroupId: g && g.id ? String(g.id) : '',
            displayName: (g && g.displayName) || (row && row.displayName) || cohortDisplayName(activeYear),
            mailNickname: (g && g.mailNickname) || (row && row.mailNickname) || cohortMailNickname(activeYear),
            mode: mode
        });
    }
    renderLeftList();
    try {
        window.dispatchEvent(
            new CustomEvent('ms365-cohort-linked', { detail: { year: activeYear, mode: mode || '' } })
        );
    } catch {
        /* ignore */
    }
}

function persistUnmatch() {
    const api = dataV2();
    if (api && typeof api.clearCatalogLinkGroup === 'function') {
        api.clearCatalogLinkGroup('cohort', activeYear);
    } else {
        persistMatch({ id: '', displayName: '', mailNickname: '' }, '');
    }
    renderLeftList();
}

function mountDetail() {
    gd().mount('#oaCohortDetailHost', {
        title: 'Abschlussjahrgang',
        searchPlaceholder: 'z. B. maturajg2026 oder Abschlussjahrgang',
        unmatchedCreateHint:
            'Legt eine Microsoft-365-Gruppe für alle Klassen dieses Abschlussjahrs an. Das ist nicht die Klassengruppe – der Klassen-Alias bleibt unberührt.',
        membersUnmatchedHint:
            'Klassen dieses Jahres stehen in den Stammdaten. Mitglieder nach dem Match in Graph pflegen.',
        membersUnmatchedTitle: 'Klassen dieses Abschlussjahrs',
        membersMatchedHint:
            'Live aus Microsoft Graph. Es gibt keinen automatischen Abgleich mit der Klassenliste.',
        emptyHintHtml:
            'Keine Abschlussjahre gefunden. Tragen Sie in den <a href="../tenant.html">Stammdaten</a> bei den Klassen ein Abschlussjahr ein.',
        features: {
            syncMembers: false,
            emptyHint: true
        },
        ids: { emptyHint: 'oaCohortEmptyHint', wrap: 'oaCohortDetailWrap' },
        live: {
            toast: toast,
            dlgConfirm: dlgConfirm,
            getGroupId: getActiveGroupId,
            ensureDirektionOwners: function (token, gid) {
                if (!direktion.length) throw new Error('Keine Direktion-Adressen in den Stammdaten.');
                return gug().ensureOwners(token, gid, direktion);
            },
            onUnmatched: function () {
                renderOwnerPreview();
                renderMemberPreview();
                renderLeftList();
            },
            onAfterLoad: function () {
                renderLeftList();
            }
        },
        match: {
            persistMatch: persistMatch,
            persistUnmatch: persistUnmatch,
            canSearch: function () {
                return activeYear
                    ? { ok: true }
                    : { ok: false, message: 'Bitte zuerst ein Abschlussjahr wählen.' };
            },
            canCreate: function () {
                return activeYear
                    ? { ok: true }
                    : { ok: false, message: 'Bitte zuerst ein Abschlussjahr wählen.' };
            },
            ensureOwners: function (token, gid) {
                return gug().ensureOwners(token, gid, direktion || []);
            }
        },
        onTabUnmatched: function (tab) {
            if (tab === 'owners') renderOwnerPreview();
            if (tab === 'members') renderMemberPreview();
        }
    });
}

export function refreshCohortPanel() {
    direktion = readDirektion();
    ensureActiveYear();
    renderLeftList();
    if (!mounted) return;
    applyCreateDefaults();
    refreshMatchUi();
}

export function initCohortPanel() {
    if (!getEl('oaCohortDetailHost') || !window.ms365GroupDetail) return;
    direktion = readDirektion();
    mountDetail();
    mounted = true;
    ensureActiveYear();
    const listHost = getEl('oaCohortListItems');
    if (listHost && !listHost.dataset.bound) {
        listHost.dataset.bound = '1';
        listHost.addEventListener('click', function (ev) {
            const t = ev.target;
            const item = t && t.closest ? t.closest('button[data-oa-cohort-year]') : null;
            if (!item) return;
            setActiveYear(item.getAttribute('data-oa-cohort-year') || '');
        });
    }
    const reload = getEl('oaCohortReload');
    if (reload && !reload.dataset.bound) {
        reload.dataset.bound = '1';
        reload.addEventListener('click', function () {
            refreshCohortPanel();
            toast('Abschlussjahre neu eingelesen.');
        });
    }
    refreshCohortPanel();
    if (getActiveGroupId()) live().loadGroup({ silent: true });
}
