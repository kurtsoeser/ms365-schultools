/**
 * Anzeige der verknüpften Sammelgruppe in der geführten Einrichtung
 * (Name, Alias, Mail, Besitzer, Mitglieder, …). Rein, ohne Graph/DOM.
 */

import { escapeHtml, compareDe } from './utils/strings.js';

export const SW_MATCH_MEMBER_PREVIEW = 50;

export function personSummaryLabel(p) {
    if (!p || typeof p !== 'object') return '';
    const dn = p.displayName ? String(p.displayName).trim() : '';
    const upn = p.userPrincipalName || p.mail ? String(p.userPrincipalName || p.mail).trim() : '';
    if (dn && upn && dn !== upn) return dn + ' (' + upn + ')';
    return dn || upn || (p.id ? String(p.id) : '');
}

export function visibilityLabel(v) {
    const s = String(v || '').trim();
    if (s === 'Public') return 'Öffentlich';
    if (s === 'Private') return 'Privat';
    return s || '';
}

export function formatGroupDateTime(iso) {
    const s = String(iso || '').trim();
    if (!s) return '';
    const d = new Date(s);
    if (isNaN(d.getTime())) return s;
    try {
        return d.toLocaleString('de-AT', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
    } catch {
        return s;
    }
}

export function entraGroupOverviewUrl(groupId) {
    const id = String(groupId || '').trim();
    if (!id) return '';
    return (
        'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Overview/groupId/' +
        encodeURIComponent(id)
    );
}

export function teamsGroupConversationsUrl(groupId) {
    const id = String(groupId || '').trim();
    if (!id) return '';
    return (
        'https://teams.microsoft.com/l/team/' +
        encodeURIComponent(id) +
        '/conversations?groupId=' +
        encodeURIComponent(id)
    );
}

function dash(v) {
    const s = String(v || '').trim();
    return s ? escapeHtml(s) : '–';
}

function dlRow(label, valueHtml) {
    return (
        '<dt>' +
        escapeHtml(label) +
        '</dt><dd>' +
        (valueHtml == null || valueHtml === '' ? '–' : valueHtml) +
        '</dd>'
    );
}

function peopleListHtml(people) {
    const list = Array.isArray(people) ? people.slice() : [];
    if (!list.length) {
        return '<p class="sw-match-people__empty">Keine Einträge.</p>';
    }
    list.sort(function (a, b) {
        return compareDe(personSummaryLabel(a), personSummaryLabel(b));
    });
    let html = '<ul class="sw-match-people__list">';
    for (let i = 0; i < list.length; i++) {
        html += '<li>' + escapeHtml(personSummaryLabel(list[i]) || '–') + '</li>';
    }
    html += '</ul>';
    return html;
}

/**
 * @param {object} model
 * @param {string} [model.groupId]
 * @param {object|null} [model.group]
 * @param {object[]} [model.owners]
 * @param {object[]} [model.members]
 * @param {number} [model.memberCount]
 * @param {boolean} [model.membersTruncated]
 * @param {string} [model.artLabel]
 * @param {boolean} [model.hasTeam]
 * @param {'idle'|'loading'|'error'|'needsLogin'|'partial'|'ready'} [model.status]
 * @param {string} [model.error]
 */
export function buildLinkedGroupSummaryHtml(model) {
    const m = model && typeof model === 'object' ? model : {};
    const g = m.group && typeof m.group === 'object' ? m.group : {};
    const gid = String(m.groupId || g.id || '').trim();
    if (!gid) {
        return '<span class="sw-match-empty">Noch keine Gruppe gewählt.</span>';
    }

    const name = String(g.displayName || '').trim();
    const alias = String(g.mailNickname || '').trim();
    const mail = String(g.mail || '').trim();
    const desc = String(g.description || '').trim();
    const vis = visibilityLabel(g.visibility);
    const art = String(m.artLabel || '').trim();
    const created = formatGroupDateTime(g.createdDateTime);
    const expires = formatGroupDateTime(g.expirationDateTime);
    const renewed = formatGroupDateTime(g.renewedDateTime);
    const owners = Array.isArray(m.owners) ? m.owners : [];
    const members = Array.isArray(m.members) ? m.members : [];
    const count =
        typeof m.memberCount === 'number' && m.memberCount >= 0 ? m.memberCount : members.length;
    const truncated = !!m.membersTruncated || (count > members.length && members.length > 0);
    const status = String(m.status || '').trim() || (name ? 'partial' : 'idle');
    const entra = entraGroupOverviewUrl(gid);
    const teamsUrl = m.hasTeam ? teamsGroupConversationsUrl(gid) : '';

    let html = '<div class="sw-match-card">';
    html += '<div class="sw-match-card__head">';
    html += '<div class="sw-match-card__titles">';
    html += '<div class="sw-match-card__kicker">Verknüpfte Gruppe</div>';
    html += '<div class="sw-match-card__title">' + dash(name || 'Gruppe') + '</div>';
    if (art) html += '<div class="sw-match-card__art">' + escapeHtml(art) + '</div>';
    html += '</div>';
    html += '<div class="sw-match-card__actions">';
    if (entra) {
        html +=
            '<a class="btn" href="' +
            escapeHtml(entra) +
            '" target="_blank" rel="noopener noreferrer"><i class="bi bi-box-arrow-up-right" aria-hidden="true"></i>Entra öffnen</a>';
    }
    if (teamsUrl) {
        html +=
            '<a class="btn" href="' +
            escapeHtml(teamsUrl) +
            '" target="_blank" rel="noopener noreferrer"><i class="bi bi-microsoft-teams" aria-hidden="true"></i>Team öffnen</a>';
    }
    const refreshLabel = status === 'needsLogin' ? 'Details aus Microsoft 365 laden' : 'Aktualisieren';
    const refreshIcon = status === 'needsLogin' ? 'bi-cloud-download' : 'bi-arrow-clockwise';
    html +=
        '<button type="button" class="btn" data-sw-match-refresh="1"><i class="bi ' +
        refreshIcon +
        '" aria-hidden="true"></i>' +
        escapeHtml(refreshLabel) +
        '</button>';
    html += '</div></div>';

    if (status === 'loading') {
        html += '<p class="sw-match-card__status">Live-Daten werden aus Microsoft 365 geladen …</p>';
    } else if (status === 'needsLogin') {
        html +=
            '<p class="sw-match-card__status">Melden Sie sich an, um Besitzer, Mitglieder und weitere Felder aus Microsoft 365 zu laden.</p>';
    } else if (status === 'error') {
        html +=
            '<p class="sw-match-card__status sw-match-card__status--err">' +
            escapeHtml(m.error || 'Details konnten nicht geladen werden.') +
            '</p>';
    }

    html += '<dl class="sw-match-dl">';
    html += dlRow('Anzeigename', dash(name));
    html += dlRow('Alias', alias ? '<code>' + escapeHtml(alias) + '</code>' : '–');
    html += dlRow('Gruppen-ID', '<code>' + escapeHtml(gid) + '</code>');
    html += dlRow(
        'Mail',
        mail ? '<a href="mailto:' + escapeHtml(mail) + '">' + escapeHtml(mail) + '</a>' : '–'
    );
    html += dlRow('Beschreibung', dash(desc));
    html += dlRow('Sichtbarkeit', dash(vis));
    html += dlRow('Team', m.hasTeam ? 'Vorhanden' : status === 'ready' ? 'Kein Team' : '–');
    html += dlRow('Erstellt', dash(created));
    if (expires) html += dlRow('Ablauf', dash(expires));
    if (renewed) html += dlRow('Zuletzt verlängert', dash(renewed));
    html += '</dl>';

    const peopleReady = status === 'ready' || owners.length > 0 || members.length > 0 || count > 0;
    if (peopleReady || status === 'ready') {
        html += '<div class="sw-match-people">';
        html += '<div class="sw-match-people__col">';
        html += '<h3 class="sw-match-people__h">Besitzer (' + String(owners.length) + ')</h3>';
        html += peopleListHtml(owners);
        html += '</div>';
        html += '<div class="sw-match-people__col">';
        const shown = members.length;
        const headCount = count >= 0 ? count : shown;
        html += '<h3 class="sw-match-people__h">Mitglieder (' + String(headCount) + ')</h3>';
        html += peopleListHtml(members);
        if (truncated && shown > 0) {
            html +=
                '<p class="sw-match-people__more">Angezeigt: ' +
                String(shown) +
                ' von ' +
                String(headCount) +
                '. Die vollständige Liste finden Sie in Entra.</p>';
        }
        html += '</div></div>';
    }

    html += '</div>';
    return html;
}
