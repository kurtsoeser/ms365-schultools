/**
 * Microsoft-365-Gruppe je Verwaltungs-Zielgruppe (Stammdaten-Board).
 */
import {
    BUILTIN_AUDIENCE_SCHULLEITUNG,
    BUILTIN_AUDIENCE_VERWALTUNG,
    collectEmailsForAudienceGroup,
    isBuiltinAudienceGroupId
} from './administration-audience-groups.js';

function normStr(v) {
    return String(v ?? '').trim();
}

function gug() {
    const G = window.ms365GraphUnifiedGroups;
    if (!G) throw new Error('graph-unified-groups.js fehlt');
    return G;
}

function getSetup() {
    const api = window.ms365AppDataV2;
    return api && typeof api.getSetup === 'function' ? api.getSetup() || {} : {};
}

function patchSetup(patch) {
    const api = window.ms365AppDataV2;
    if (api && typeof api.patchSetup === 'function') api.patchSetup(patch);
}

/**
 * @param {string} groupId
 */
export function getAudienceGroupGraphId(groupId) {
    const id = normStr(groupId);
    const matched = getSetup().matched || {};
    if (id === BUILTIN_AUDIENCE_SCHULLEITUNG) return normStr(matched.schulleitungGroupId);
    if (id === BUILTIN_AUDIENCE_VERWALTUNG) return normStr(matched.verwaltungGroupId);
    const map = matched.verwaltungAudienceGroupIds && typeof matched.verwaltungAudienceGroupIds === 'object'
        ? matched.verwaltungAudienceGroupIds
        : {};
    return normStr(map[id]);
}

/**
 * @param {string} groupId
 * @param {{ id?: string, displayName?: string, mailNickname?: string, mode?: string }|null} group
 */
export function patchAudienceGroupGraphId(groupId, group) {
    const gid = normStr(group && group.id);
    const id = normStr(groupId);
    const api = window.ms365AppDataV2;
    if (id === BUILTIN_AUDIENCE_SCHULLEITUNG) {
        patchSetup({ matched: { schulleitungGroupId: gid || null } });
        if (gid && api && typeof api.upsertCatalogLink === 'function') {
            api.upsertCatalogLink({
                kind: 'sammelgruppe',
                code: 'schulleitung',
                graphGroupId: gid,
                displayName: normStr(group && group.displayName),
                mailNickname: normStr(group && group.mailNickname),
                mode: (group && group.mode) || 'matched',
                syncStatus: ''
            });
        }
        return;
    }
    if (id === BUILTIN_AUDIENCE_VERWALTUNG) {
        patchSetup({ matched: { verwaltungGroupId: gid || null } });
        if (gid && api && typeof api.upsertCatalogLink === 'function') {
            api.upsertCatalogLink({
                kind: 'sammelgruppe',
                code: 'verwaltung',
                graphGroupId: gid,
                displayName: normStr(group && group.displayName),
                mailNickname: normStr(group && group.mailNickname),
                mode: (group && group.mode) || 'matched',
                syncStatus: ''
            });
        }
        return;
    }
    const matched = getSetup().matched || {};
    const map = Object.assign({}, matched.verwaltungAudienceGroupIds || {});
    if (gid) map[id] = gid;
    else delete map[id];
    patchSetup({ matched: { verwaltungAudienceGroupIds: map } });
}

/**
 * @param {{ id: string, label: string, mailNick?: string, newDisplayName?: string, builtin?: boolean }} group
 * @param {string} schoolName
 */
export function audienceGroupNamingDefaults(group, schoolName) {
    const setup = getSetup();
    const label = normStr(group && group.label) || 'Verwaltungsgruppe';
    const school = normStr(schoolName);
    let nick = '';
    let displayName = '';
    if (group && group.id === BUILTIN_AUDIENCE_SCHULLEITUNG) {
        const d = setup.schulleitungDraft || {};
        nick = normStr(d.slNewMailNick) || 'schulleitung';
        displayName = normStr(d.slNewDisplayName) || (school ? 'Schulleitung ' + school : 'Schulleitung');
    } else if (group && group.id === BUILTIN_AUDIENCE_VERWALTUNG) {
        const d = setup.verwaltungDraft || {};
        nick = normStr(d.vwNewMailNick) || 'verwaltung';
        displayName = normStr(d.vwNewDisplayName) || (school ? 'Verwaltung ' + school : 'Schulverwaltung');
    } else {
        nick = normStr(group && group.mailNick) || slugMailNickFromLabel(label);
        displayName = normStr(group && group.newDisplayName) || label;
    }
    nick = gug().sanitizeUnifiedGroupMailNickname(nick || 'verwaltung');
    return {
        mailNick: nick,
        displayName: displayName,
        description: label + ' (MS365-Schul-Tools / Stammdaten-Zielgruppe)'
    };
}

function slugMailNickFromLabel(label) {
    return normStr(label)
        .toLowerCase()
        .replace(/ä/g, 'ae')
        .replace(/ö/g, 'oe')
        .replace(/ü/g, 'ue')
        .replace(/ß/g, 'ss')
        .replace(/[^a-z0-9]+/g, '-')
        .replace(/^-+|-+$/g, '')
        .slice(0, 48) || 'verwaltung-gruppe';
}

function schoolDomain() {
    return typeof window.ms365GetSchoolDomainNoAt === 'function' ? normStr(window.ms365GetSchoolDomainNoAt()) : '';
}

function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function paintMatchStatus(el, payload, expectedNick) {
    if (!el) return;
    el.style.background = '';
    el.style.color = '';
    if (!payload) {
        el.textContent = '—';
        el.title = 'Noch nicht geprüft';
        return;
    }
    if (payload.loading) {
        el.textContent = 'Prüfe …';
        return;
    }
    if (payload.found && payload.group) {
        const g = payload.group;
        const gid = normStr(g.id);
        const shown = normStr(g.displayName) || normStr(g.mailNickname) || expectedNick || 'gematcht';
        el.innerHTML =
            '<span style="color:#0d8050;font-weight:700;">✓</span> <code style="font-size:0.88em;">' +
            escapeHtml(shown) +
            '</code>';
        el.title = (g.displayName || '') + (gid ? '\nID: ' + gid : '');
        el.style.background = 'color-mix(in srgb, #0d8050 8%, transparent)';
        return;
    }
    if (payload.notFound) {
        el.innerHTML = '<span style="color:#856404;font-weight:700;">✗</span> nicht gefunden';
        el.title = 'Keine passende Gruppe – Anlegen oder Suchen';
        return;
    }
    if (payload.error) {
        el.textContent = 'Fehler';
        el.title = String(payload.error);
    }
}

async function resolveGroup(expectedNick, expectedDn, storedId) {
    const token = await gug().getGraphToken();
    if (storedId && typeof gug().fetchGroup === 'function') {
        try {
            const g = await gug().fetchGroup(token, storedId);
            if (g && g.id) return { found: true, group: g };
        } catch {
            /* stored invalid */
        }
    }
    const domain = schoolDomain().replace(/^@+/, '');
    const mail = domain && expectedNick ? expectedNick + '@' + domain : '';
    const queries = [mail, expectedNick, expectedDn].filter(Boolean);
    let hits = [];
    for (let i = 0; i < queries.length; i++) {
        hits = await gug().searchUnifiedGroups(token, queries[i]);
        if (hits && hits.length) break;
    }
    const nickLc = normStr(expectedNick).toLowerCase();
    const match =
        Array.isArray(hits) &&
        hits.find(function (g) {
            const mn = normStr(g && g.mailNickname).toLowerCase();
            if (mn && mn === nickLc) return true;
            const m = normStr(g && g.mail).toLowerCase();
            return mail && m === mail.toLowerCase();
        });
    if (match) return { found: true, group: match };
    return { found: false, notFound: true, hits: hits || [] };
}

/**
 * @param {{
 *   group: object,
 *   memberships: object[],
 *   schoolName: string,
 *   setSummary?: (msg: string, kind?: string) => void,
 *   onGroupMetaChange?: (groupId: string, patch: { mailNick?: string, newDisplayName?: string }) => void,
 *   onMatchChanged?: () => void,
 * }} ctx
 */
function actBtn(iconClass, label, title, variant) {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'ts-audience-board__act-btn ts-audience-board__act-btn--' + (variant || 'neutral');
    const t = title || label;
    b.title = t;
    b.setAttribute('aria-label', t);
    b.innerHTML =
        '<i class="bi ' + iconClass + '" aria-hidden="true"></i><span class="ts-audience-board__act-label">' + label + '</span>';
    return b;
}

function actLink(iconClass, label, title, href) {
    const a = document.createElement('a');
    a.className = 'ts-audience-board__act-btn ts-audience-board__act-btn--neutral';
    a.href = href;
    const t = title || label;
    a.title = t;
    a.setAttribute('aria-label', t);
    a.innerHTML =
        '<i class="bi ' + iconClass + '" aria-hidden="true"></i><span class="ts-audience-board__act-label">' + label + '</span>';
    return a;
}

function labeledField(labelText, inputEl) {
    const wrap = document.createElement('label');
    wrap.className = 'ts-audience-board__m365-field';
    const lab = document.createElement('span');
    lab.className = 'ts-audience-board__m365-label';
    lab.textContent = labelText;
    wrap.append(lab, inputEl);
    return wrap;
}

/**
 * @returns {{ inline: HTMLElement, aux: HTMLElement }}
 */
export function buildAudienceGroupM365Panel(ctx) {
    const group = ctx.group;
    const inline = document.createElement('div');
    inline.className = 'ts-audience-board__m365-inline';
    const aux = document.createElement('div');
    aux.className = 'ts-audience-board__m365-aux';

    const naming = audienceGroupNamingDefaults(group, ctx.schoolName);
    const statusEl = document.createElement('span');
    statusEl.className = 'ts-audience-board__m365-status muted';
    const storedId = getAudienceGroupGraphId(group.id);
    if (storedId) {
        paintMatchStatus(statusEl, {
            found: true,
            group: { id: storedId, displayName: naming.displayName, mailNickname: naming.mailNick }
        });
    } else {
        paintMatchStatus(statusEl, null);
    }

    const inpDn = document.createElement('input');
    inpDn.type = 'text';
    inpDn.className = 'ts-audience-board__m365-inp';
    inpDn.placeholder = 'z. B. Schulleitung …';
    inpDn.value = naming.displayName;
    inpDn.autocomplete = 'off';
    inpDn.title = 'Anzeigename der Microsoft-365-Gruppe';
    const inpNick = document.createElement('input');
    inpNick.type = 'text';
    inpNick.className = 'ts-audience-board__m365-inp';
    inpNick.placeholder = 'z. B. schulleitung';
    inpNick.value = naming.mailNick;
    inpNick.autocomplete = 'off';
    inpNick.spellcheck = false;
    inpNick.title = 'Mail-Alias (Nickname)';
    function refreshMailPreview() {
        const dom = schoolDomain().replace(/^@+/, '');
        const nick = gug().sanitizeUnifiedGroupMailNickname(inpNick.value || '');
        inpNick.title = dom && nick ? 'Alias · ' + nick + '@' + dom : 'Mail-Alias (Nickname)';
    }
    inpNick.addEventListener('input', refreshMailPreview);
    refreshMailPreview();

    if (!isBuiltinAudienceGroupId(group.id)) {
        const persistMeta = function () {
            if (typeof ctx.onGroupMetaChange === 'function') {
                ctx.onGroupMetaChange(group.id, {
                    mailNick: inpNick.value,
                    newDisplayName: inpDn.value
                });
            }
        };
        inpDn.addEventListener('change', persistMeta);
        inpNick.addEventListener('change', persistMeta);
    } else {
        inpDn.title = 'Für Standardgruppen: Werkzeug Verwaltung oder Setup-Draft';
        inpNick.title = inpDn.title;
    }

    const actions = document.createElement('div');
    actions.className = 'ts-audience-board__m365-btns';
    actions.setAttribute('role', 'group');
    actions.setAttribute('aria-label', 'Microsoft-365-Gruppe');

    const btnVerify = actBtn('bi-microsoft', 'Prüfen', 'Microsoft-365-Gruppe prüfen und zuordnen', 'brand');
    const btnCreate = actBtn('bi-plus-lg', 'Anlegen', 'Neue Microsoft-365-Gruppe erstellen', 'success');
    const btnUnmatch = actBtn('bi-link-45deg', 'Lösen', 'Zuordnung zur Microsoft-365-Gruppe aufheben', 'danger');
    btnUnmatch.hidden = !storedId;
    const btnOpen = actBtn('bi-box-arrow-up-right', 'Entra', 'Gruppe in Microsoft Entra öffnen', 'neutral');
    btnOpen.hidden = !storedId;
    const btnSync = actBtn('bi-arrow-repeat', 'Sync', 'Mitglieder dieser Zielgruppe in die M365-Gruppe übernehmen', 'sync');
    btnSync.hidden = !storedId;
    const btnSearchToggle = actBtn('bi-search', 'Suchen', 'Bestehende Gruppe suchen und zuordnen', 'neutral');
    const linkTool = actLink(
        'bi-building-gear',
        'Werkzeug',
        'Werkzeug Verwaltung – vollständiger Abgleich',
        'tools/verwaltung.html?kind=' +
            (group.id.indexOf('va-') === 0 ? 'audience:' + encodeURIComponent(group.id) : encodeURIComponent(group.id))
    );

    const searchWrap = document.createElement('div');
    searchWrap.className = 'ts-audience-board__m365-search';
    searchWrap.hidden = true;
    const searchIn = document.createElement('input');
    searchIn.type = 'search';
    searchIn.placeholder = 'Gruppe suchen …';
    searchIn.className = 'ts-audience-board__input';
    const btnSearch = document.createElement('button');
    btnSearch.type = 'button';
    btnSearch.className = 'btn btn-sm';
    btnSearch.innerHTML = '<i class="bi bi-search"></i>';
    searchWrap.append(searchIn, btnSearch);
    const searchResults = document.createElement('div');
    searchResults.className = 'ts-audience-board__m365-results';
    searchResults.hidden = true;

    function notify(msg, kind) {
        if (typeof ctx.setSummary === 'function') ctx.setSummary(msg, kind);
    }

    function afterMatch(g) {
        btnUnmatch.hidden = !g;
        btnOpen.hidden = !g;
        btnSync.hidden = !g;
        if (typeof ctx.onMatchChanged === 'function') ctx.onMatchChanged();
    }

    btnVerify.addEventListener('click', async function () {
        const nick = gug().sanitizeUnifiedGroupMailNickname(inpNick.value || naming.mailNick);
        const dn = normStr(inpDn.value) || naming.displayName;
        paintMatchStatus(statusEl, { loading: true }, nick);
        try {
            const res = await resolveGroup(nick, dn, getAudienceGroupGraphId(group.id));
            if (res.found && res.group) {
                patchAudienceGroupGraphId(group.id, res.group);
                paintMatchStatus(statusEl, { found: true, group: res.group }, nick);
                notify('Microsoft-365-Gruppe gefunden: ' + (res.group.displayName || nick), 'ok');
                afterMatch(res.group.id);
            } else {
                patchAudienceGroupGraphId(group.id, null);
                paintMatchStatus(statusEl, { notFound: true }, nick);
                notify('Keine passende Gruppe gefunden.', 'warn');
                afterMatch(null);
            }
        } catch (e) {
            paintMatchStatus(statusEl, { error: e.message || e }, nick);
            notify('Prüfen: ' + (e.message || e), 'warn');
        }
    });

    btnCreate.addEventListener('click', async function () {
        const nick = gug().sanitizeUnifiedGroupMailNickname(inpNick.value || naming.mailNick);
        const dn = normStr(inpDn.value) || naming.displayName;
        paintMatchStatus(statusEl, { loading: true }, nick);
        try {
            const token = await gug().getGraphToken();
            const created = await gug().createUnifiedGroup(token, dn, nick, naming.description);
            patchAudienceGroupGraphId(group.id, created);
            paintMatchStatus(statusEl, { found: true, group: created }, nick);
            notify('Gruppe angelegt: ' + dn, 'ok');
            afterMatch(created && created.id);
        } catch (e) {
            paintMatchStatus(statusEl, { error: e.message || e }, nick);
            notify('Anlegen: ' + (e.message || e), 'warn');
        }
    });

    btnUnmatch.addEventListener('click', function () {
        patchAudienceGroupGraphId(group.id, null);
        paintMatchStatus(statusEl, null);
        notify('Zuordnung zur Microsoft-365-Gruppe gelöst.', 'ok');
        afterMatch(null);
    });

    btnOpen.addEventListener('click', function () {
        const gid = getAudienceGroupGraphId(group.id);
        if (!gid) return;
        const url = 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Overview/groupId/' + gid;
        window.open(url, '_blank', 'noopener,noreferrer');
    });

    btnSync.addEventListener('click', async function () {
        const gid = getAudienceGroupGraphId(group.id);
        if (!gid) {
            notify('Zuerst Microsoft-365-Gruppe prüfen oder anlegen.', 'warn');
            return;
        }
        const emails = collectEmailsForAudienceGroup(ctx.memberships || [], group.id);
        if (!emails.length) {
            notify('Keine E-Mails in dieser Zielgruppe.', 'warn');
            return;
        }
        btnSync.disabled = true;
        try {
            const token = await gug().getGraphToken();
            const r = await gug().syncEmailsToGroup(token, gid, emails, group.label, function () {});
            notify(
                'Sync: ' + (r.ok || 0) + ' hinzugefügt/aktualisiert' + (r.fail ? ', ' + r.fail + ' Fehler' : '') + '.',
                r.fail ? 'warn' : 'ok'
            );
        } catch (e) {
            notify('Sync: ' + (e.message || e), 'warn');
        } finally {
            btnSync.disabled = false;
        }
    });

    btnSearchToggle.addEventListener('click', function () {
        searchWrap.hidden = !searchWrap.hidden;
        if (!searchWrap.hidden) searchIn.focus();
    });

    btnSearch.addEventListener('click', async function () {
        const q = normStr(searchIn.value);
        if (!q) return;
        searchResults.hidden = false;
        searchResults.textContent = 'Suche …';
        try {
            const token = await gug().getGraphToken();
            const hits = await gug().searchUnifiedGroups(token, q);
            searchResults.replaceChildren();
            if (!hits || !hits.length) {
                searchResults.textContent = 'Keine Treffer.';
                return;
            }
            hits.slice(0, 8).forEach(function (g) {
                const row = document.createElement('div');
                row.className = 'ts-audience-board__m365-hit';
                row.innerHTML =
                    '<span><strong>' +
                    escapeHtml(g.displayName || '') +
                    '</strong> <code>' +
                    escapeHtml(g.mailNickname || '') +
                    '</code></span>';
                const b = document.createElement('button');
                b.type = 'button';
                b.className = 'btn btn-sm btn-primary';
                b.textContent = 'Zuordnen';
                b.addEventListener('click', function () {
                    patchAudienceGroupGraphId(group.id, g);
                    paintMatchStatus(statusEl, { found: true, group: g }, inpNick.value);
                    notify('Gruppe zugeordnet.', 'ok');
                    searchResults.hidden = true;
                    afterMatch(g.id);
                });
                row.appendChild(b);
                searchResults.appendChild(row);
            });
        } catch (e) {
            searchResults.textContent = String(e.message || e);
        }
    });

    actions.append(btnVerify, btnCreate, btnUnmatch, btnOpen, btnSync, btnSearchToggle, linkTool);
    inline.append(labeledField('Anzeigename', inpDn), labeledField('Alias', inpNick), statusEl, actions);
    aux.append(searchWrap, searchResults);
    return { inline, aux };
}
