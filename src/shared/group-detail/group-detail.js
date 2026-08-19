(function () {
    'use strict';

    var ATTR_MOUNTED = 'data-gd-mounted';
    var ATTR_WIRED = 'data-gd-wired';
    var liveWired = false;
    var ENTRA_GROUP =
        'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Members/groupId/';

    /** @type {null|{ toast?: Function, live?: object, match?: object, onTabUnmatched?: Function }} */
    var session = null;
    var activeTab = 'general';

    function escapeHtml(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function liveApi() {
        const L = window.ms365SlgLiveDetails;
        if (!L) throw new Error('slg-live-details.js muss vor group-detail.js geladen werden.');
        return L;
    }

    function gug() {
        const G = window.ms365GraphUnifiedGroups;
        if (!G) throw new Error('graph-unified-groups.js muss vor diesem Skript geladen werden.');
        return G;
    }

    function toast(msg) {
        if (session && session.live && session.live.toast) session.live.toast(msg);
    }

    function featureBag(opts) {
        const f = (opts && opts.features) || {};
        return {
            aliasEditable: !!f.aliasEditable,
            createdAt: f.createdAt !== false,
            teams: f.teams !== false,
            renewExpiration: f.renewExpiration !== false,
            smtpSlot: !!f.smtpSlot,
            syncMembers: f.syncMembers !== false,
            membershipReview: f.membershipReview === true,
            emptyHint: !!f.emptyHint,
            header: f.header !== false,
            matchUi: f.matchUi !== false,
            openEntra: f.openEntra !== false,
            deleteGroup: !!f.deleteGroup,
            teamArchive: !!f.teamArchive,
            ensureDirektion: f.ensureDirektion !== false,
            visibilityUnsupported: !!f.visibilityUnsupported
        };
    }

    function idBag(opts) {
        const ids = (opts && opts.ids) || {};
        return {
            emptyHint: ids.emptyHint || 'gdEmptyHint',
            wrap: ids.wrap || 'gdDetailWrap',
            headActions: ids.headActions || '',
            afterWrap: ids.afterWrap || '',
            ownerExtra: ids.ownerExtra || ''
        };
    }

    function defaults(opts) {
        const o = opts && typeof opts === 'object' ? opts : {};
        const f = featureBag(o);
        const aliasInUpdate = f.aliasEditable;
        return {
            title: o.title || 'Gruppe',
            subtitle: o.subtitle || 'Gruppe matchen oder anlegen',
            emptyHintHtml:
                o.emptyHintHtml ||
                'Keine Einträge in dieser Liste. Bitte unter <a href="../tenant.html">Stammdaten</a> pflegen.',
            searchPlaceholder: o.searchPlaceholder || 'Name, Mail oder Alias …',
            unmatchedCreateHint:
                o.unmatchedCreateHint ||
                'Legt eine Microsoft 365‑Gruppe (Unified) an. Optional auch als Team bereitstellen.',
            matchedHintHtml:
                o.matchedHintHtml ||
                (aliasInUpdate
                    ? 'Gematchte Microsoft 365‑Gruppe. Hervorgehobene Felder speichert <strong>Update</strong> (Anzeigename, Alias, Beschreibung, Sichtbarkeit). Graue Felder sind nur Anzeige.'
                    : f.matchUi
                      ? 'Gematchte Microsoft 365‑Gruppe. Hervorgehobene Felder speichert <strong>Update</strong> (Anzeigename, Beschreibung, Sichtbarkeit). Graue Felder sind nur Anzeige.'
                      : 'Microsoft 365‑Gruppe im Tenant. Hervorgehobene Felder speichert <strong>Update</strong> (Anzeigename, Beschreibung, Sichtbarkeit). Graue Felder sind nur Anzeige.'),
            updateHintHtml:
                o.updateHintHtml ||
                (aliasInUpdate
                    ? 'Update speichert <strong>Anzeigename</strong>, <strong>Alias</strong>, <strong>Beschreibung</strong> und <strong>Sichtbarkeit</strong>. Zum Wechseln der Gruppe: „Match lösen“.'
                    : f.matchUi
                      ? 'Update speichert <strong>Anzeigename</strong>, <strong>Beschreibung</strong> und <strong>Sichtbarkeit</strong>. Zum Wechseln der Gruppe: „Match lösen“.'
                      : 'Update speichert <strong>Anzeigename</strong>, <strong>Beschreibung</strong> und <strong>Sichtbarkeit</strong> (Unified) sowie optional den <strong>Teams‑Archiv‑Status</strong>. E‑Mail/Alias bleiben unverändert.'),
            ownersUnmatchedHint:
                o.ownersUnmatchedHint ||
                'Besitzer werden aus der Verwaltungsliste (Rolle „Direktion“) abgeleitet und beim Anlegen gesetzt. Nach dem Match erscheinen hier die Live‑Besitzer aus Microsoft Graph.',
            membersUnmatchedHint:
                o.membersUnmatchedHint ||
                'Nach dem Match können Mitglieder live in Microsoft Graph gepflegt werden.',
            membersUnmatchedTitle: o.membersUnmatchedTitle || 'Vorschau',
            membersMatchedHint:
                o.membersMatchedHint ||
                (f.syncMembers
                    ? 'Live aus Microsoft Graph. „Mitglieder synchronisieren“ gleicht die Gruppe mit der Schul‑Liste ab: fehlende Adressen werden hinzugefügt, Mitglieder die nicht in der Liste stehen werden entfernt.'
                    : 'Live aus Microsoft Graph. Mitglieder hier suchen, hinzufügen oder entfernen.')
        };
    }

    function buildMarkup(opts) {
        const t = defaults(opts);
        const f = featureBag(opts);
        const ids = idBag(opts);
        const aliasClass = f.aliasEditable ? 'field-editable' : 'field-readonly';
        const aliasRo = f.aliasEditable ? '' : ' readonly';
        const headId = ids.headActions
            ? ' id="' + escapeHtml(ids.headActions) + '"'
            : '';
        let html = '';
        if (f.header) {
            html +=
                '<div class="panel-head" style="display:flex; justify-content:space-between; align-items:flex-start; gap:12px; flex-wrap:wrap;">' +
                '<div>' +
                '<h2 id="slgDetailTitle" style="margin:0;">' +
                escapeHtml(t.title) +
                '</h2>' +
                '<div class="muted" id="slgDetailSubtitle" style="margin-top:6px;">' +
                escapeHtml(t.subtitle) +
                '</div>' +
                '</div>' +
                '<div class="detail-actions" style="margin-top:0;"' +
                headId +
                '>';
            if (f.openEntra) {
                html +=
                    '<button type="button" class="btn" id="slgBtnOpenEntra" disabled><i class="bi bi-box-arrow-up-right"></i>Entra öffnen</button>';
            }
            if (f.matchUi) {
                html +=
                    '<button type="button" class="btn btn-danger" id="slgBtnUnmatch" disabled><i class="bi bi-x-circle"></i>Match lösen</button>';
            }
            html += '</div></div>';
        }
        html += f.header ? '<div class="panel-body">' : '<div class="gd-embed">';
        if (f.emptyHint) {
            html +=
                '<div id="' +
                escapeHtml(ids.emptyHint) +
                '" class="hint" style="display:none;">' +
                t.emptyHintHtml +
                '</div>';
        }
        html +=
            '<div id="' +
            escapeHtml(ids.wrap) +
            '">' +
            '<div class="detail-tabs-sticky" aria-label="Detail-Tabs">' +
            '<div class="detail-tabs" role="tablist" id="slgDetailTabs" aria-label="Tabs Details">' +
            '<button type="button" class="detail-tab-btn" id="slgTabBtnGeneral" role="tab" data-slg-tab="general" aria-selected="true" aria-controls="slgTabGeneral">Allgemein</button>' +
            '<button type="button" class="detail-tab-btn" id="slgTabBtnOwners" role="tab" data-slg-tab="owners" aria-selected="false" aria-controls="slgTabOwners">Besitzer</button>' +
            '<button type="button" class="detail-tab-btn" id="slgTabBtnMembers" role="tab" data-slg-tab="members" aria-selected="false" aria-controls="slgTabMembers">Mitglieder</button>' +
            '</div>' +
            '</div>' +
            '<div id="slgTabGeneral" class="tab-panel active" role="tabpanel" data-slg-tab-content="general" aria-labelledby="slgTabBtnGeneral">';
        if (f.matchUi) {
            html +=
                '<div id="slgUnmatchedPanel">' +
                '<div class="field" style="margin-top:0;">' +
                '<label for="slgGroupSearch">Bestehende Gruppe suchen (Name, Mail, Alias, Beschreibung)</label>' +
                '<div class="slg-search-wrap">' +
                '<input type="text" id="slgGroupSearch" autocomplete="off" placeholder="' +
                escapeHtml(t.searchPlaceholder) +
                '">' +
                '<button type="button" class="btn" id="slgBtnSearch"><i class="bi bi-search"></i>Suchen</button>' +
                '</div>' +
                '<div id="slgGroupSearchResults" class="slg-search-results" style="display:none;"></div>' +
                '</div>' +
                '<div style="margin-top:18px; padding-top:16px; border-top:1px solid var(--border);">' +
                '<div class="slg-section-title">Neue Gruppe anlegen</div>' +
                '<p class="muted" style="margin:0 0 10px;">' +
                escapeHtml(t.unmatchedCreateHint) +
                '</p>' +
                '<div class="form-grid" style="margin-top:0;">' +
                '<div class="field" style="grid-column:1 / -1;">' +
                '<label for="slgNewDisplayName">Anzeigename</label>' +
                '<input type="text" id="slgNewDisplayName" maxlength="200" autocomplete="off">' +
                '</div>' +
                '<div class="field">' +
                '<label for="slgNewMailNick">Alias / Mail‑Nickname</label>' +
                '<input type="text" id="slgNewMailNick" maxlength="60" autocomplete="off" spellcheck="false" inputmode="latin">' +
                '</div>' +
                '<div class="field">' +
                '<label for="slgNewDescription">Beschreibung</label>' +
                '<input type="text" id="slgNewDescription" maxlength="512" autocomplete="off">' +
                '</div>' +
                '</div>' +
                '<label class="checkbox-label" for="slgNewCreateTeam" style="margin-top:12px; display:inline-flex;">' +
                '<input type="checkbox" id="slgNewCreateTeam">' +
                '<span>Auch als <strong>Team</strong> anlegen (optional)</span>' +
                '</label>' +
                '<div class="detail-actions">' +
                '<button type="button" class="btn" id="slgBtnCreate"><i class="bi bi-plus-circle"></i>Anlegen &amp; matchen</button>' +
                '</div>' +
                '</div>' +
                '</div>';
        }
        html +=
            '<div id="slgMatchedPanel"' +
            (f.matchUi ? ' style="display:none;"' : '') +
            '>' +
            '<div class="gd-group-photo" id="slgGroupPhotoWrap">' +
            '<div class="gd-group-photo__avatar" id="slgGroupPhotoAvatar" aria-hidden="true">' +
            '<img id="slgGroupPhotoImg" alt="" hidden>' +
            '<span id="slgGroupPhotoInitials" class="gd-group-photo__initials">–</span>' +
            '</div>' +
            '<div class="gd-group-photo__body">' +
            '<div class="gd-group-photo__title">Gruppenbild</div>' +
            '<div class="detail-actions gd-group-photo__actions">' +
            '<label class="btn" for="slgGroupPhotoFile"><i class="bi bi-image"></i>Bild hochladen</label>' +
            '<input type="file" id="slgGroupPhotoFile" accept="image/jpeg,image/png,image/webp" hidden>' +
            '<button type="button" class="btn btn-ghost" id="slgBtnRemoveGroupPhoto" hidden><i class="bi bi-trash"></i>Entfernen</button>' +
            '</div>' +
            '<p class="muted gd-group-photo__hint">JPEG, PNG oder WebP, max. 4&nbsp;MB. Wird in Microsoft&nbsp;365 angezeigt; bei Gruppen mit Team wird das Bild zusätzlich per Graph an Teams mitgesetzt (Verzögerung/SharePoint-Sync möglich).</p>' +
            '</div>' +
            '</div>' +
            '<div class="hint" style="margin-bottom:12px;">' +
            t.matchedHintHtml +
            '</div>' +
            '<div class="form-grid" style="margin-top:0;">' +
            '<div class="field field-editable">' +
            '<label for="slgLiveName">Anzeigename</label>' +
            '<input type="text" id="slgLiveName" maxlength="200" autocomplete="off">' +
            '</div>' +
            '<div class="field ' +
            aliasClass +
            '">' +
            '<label for="slgLiveAlias">Alias</label>' +
            '<input type="text" id="slgLiveAlias" maxlength="60" autocomplete="off" spellcheck="false"' +
            aliasRo +
            '>' +
            '</div>' +
            '<div class="field field-editable" style="grid-column:1 / -1;">' +
            '<label for="slgLiveDescription">Beschreibung</label>' +
            '<textarea id="slgLiveDescription" maxlength="1024" placeholder="Beschreibung der Microsoft‑365‑Gruppe"></textarea>' +
            '</div>' +
            '<div class="field field-readonly">' +
            '<label for="slgLiveMail">E‑Mail</label>' +
            '<input type="text" id="slgLiveMail" readonly>';
        if (f.smtpSlot) {
            html +=
                '<details class="jg-smtp-drop" id="jgSmtpDrop">' +
                '<summary>SMTP auf Schul‑Domain …</summary>' +
                '<div class="jg-smtp-drop__body">' +
                '<div id="jgSmtpHint">–</div>' +
                '<div class="detail-actions" style="margin-top:10px;">' +
                '<button type="button" class="btn" id="jgBtnSmtpThis"><i class="bi bi-envelope-at"></i>Skript erzeugen</button>' +
                '</div>' +
                '<textarea id="jgSmtpScript" readonly spellcheck="false" aria-label="Exchange-Skript für SMTP"></textarea>' +
                '</div>' +
                '</details>';
        }
        html +=
            '</div>' +
            '<div class="field field-readonly">' +
            '<label for="slgLiveId">Gruppen‑ID</label>' +
            '<input type="text" id="slgLiveId" readonly style="font-family:Consolas, \'Segoe UI\', monospace;">' +
            '</div>' +
            '<div class="field field-readonly">' +
            '<label for="slgLiveArt">Art</label>' +
            '<input type="text" id="slgLiveArt" readonly>' +
            '</div>' +
            '<div class="field field-editable">' +
            '<label for="slgLiveVisibility">Sichtbarkeit</label>' +
            '<select id="slgLiveVisibility">' +
            (f.visibilityUnsupported ? '<option value="">(nicht unterstützt)</option>' : '') +
            '<option value="Private">Privat</option>' +
            '<option value="Public">Öffentlich</option>' +
            '</select>' +
            '</div>';
        if (f.createdAt) {
            html +=
                '<div class="field field-readonly">' +
                '<label for="slgLiveCreated">Erstellt am</label>' +
                '<input type="text" id="slgLiveCreated" readonly placeholder="–">' +
                '</div>';
        }
        html +=
            '<div class="field field-readonly">' +
            '<label for="slgLiveExpires">Ablaufdatum</label>' +
            '<input type="text" id="slgLiveExpires" readonly placeholder="kein Ablaufdatum">';
        if (f.renewExpiration) {
            html +=
                '<details class="jg-smtp-drop" id="jgExpiresDrop">' +
                '<summary>Ablauf verlängern …</summary>' +
                '<div class="jg-smtp-drop__body">' +
                '<div id="slgLiveExpiresHint" class="muted">Microsoft setzt kein frei wählbares Datum. Es kommt aus der Gruppen-Lebenszyklusrichtlinie des Tenants.</div>' +
                '<div class="detail-actions" style="margin-top:10px;">' +
                '<button type="button" class="btn" id="slgBtnRenewExpires"><i class="bi bi-arrow-repeat"></i>Ablauf verlängern</button>' +
                '</div>' +
                '</div>' +
                '</details>';
        }
        html += '</div>';
        if (f.teams) {
            html +=
                '<div class="field field-readonly" style="grid-column:1 / -1;">' +
                '<label>Microsoft Teams</label>' +
                '<div class="jg-team-quiet">' +
                '<span id="slgLiveTeam" class="jg-team-status">–</span>' +
                '<button type="button" class="jg-quiet-btn" id="slgBtnProvisionTeam" hidden>Team anlegen</button>' +
                '<a id="slgLiveTeamLink" class="jg-quiet-link" href="#" target="_blank" rel="noopener noreferrer" hidden>In Teams öffnen</a>' +
                '</div>' +
                '</div>';
        }
        html += '</div>';
        if (f.teamArchive) {
            html +=
                '<div id="slgTeamArchiveWrap" class="gd-archive-card" style="display:none;">' +
                '<div class="gd-archive-card__head">' +
                '<span class="gd-archive-card__icon" aria-hidden="true"><i class="bi bi-archive"></i></span>' +
                '<div>' +
                '<div class="gd-archive-card__title">Teams‑Archiv</div>' +
                '<p class="gd-archive-card__lead">Nur für Microsoft 365‑Gruppen mit gebundenem Team (Allgemeines Team, Kursteam usw.). Reine Gruppen ohne Team können hier nicht archiviert werden.</p>' +
                '</div>' +
                '</div>' +
                '<div class="gd-archive-card__body">' +
                '<div class="field">' +
                '<label for="slgArchiveState">Status</label>' +
                '<select id="slgArchiveState">' +
                '<option value="active">Aktiv (nicht archiviert)</option>' +
                '<option value="archived">Archiviert</option>' +
                '</select>' +
                '</div>' +
                '<p id="slgArchiveHint" class="gd-archive-card__hint" style="display:none;"></p>' +
                '<label class="gd-archive-option" for="slgArchiveSpoReadonly">' +
                '<input type="checkbox" id="slgArchiveSpoReadonly">' +
                '<span>' +
                '<strong>SharePoint‑Website schreibgeschützt</strong>' +
                '<small>Beim Archivieren dürfen Mitglieder Dateien auf der Teamwebsite nicht mehr ändern. Nur mit delegierter (persönlicher) Anmeldung.</small>' +
                '</span>' +
                '</label>' +
                '</div>' +
                '</div>';
        }
        html +=
            '<div class="detail-actions gd-detail-actions">' +
            '<div class="gd-detail-actions__primary">' +
            '<button type="button" class="btn btn-success" id="slgBtnUpdateGroup"><i class="bi bi-save"></i>Update</button>' +
            '<button type="button" class="btn btn-ghost" id="slgBtnRefreshGroup"><i class="bi bi-arrow-clockwise"></i>Neu laden</button>';
        if (!f.header && f.openEntra) {
            html +=
                '<button type="button" class="btn btn-ghost" id="slgBtnOpenEntra" disabled><i class="bi bi-box-arrow-up-right"></i>Entra öffnen</button>';
        }
        html += '</div>';
        if (f.deleteGroup) {
            html +=
                '<div class="gd-detail-actions__danger">' +
                '<button type="button" class="btn btn-ghost gd-btn-danger" id="slgBtnDeleteGroup"><i class="bi bi-trash"></i>Löschen</button>' +
                '</div>';
        }
        html +=
            '</div>' +
            '<p class="gd-detail-hint">' +
            t.updateHintHtml +
            '</p>' +
            '</div>' +
            '</div>' +
            '<div id="slgTabOwners" class="tab-panel" role="tabpanel" data-slg-tab-content="owners" aria-labelledby="slgTabBtnOwners">';
        if (f.matchUi) {
            html +=
                '<div id="slgOwnersUnmatched">' +
                '<p class="muted" style="margin:0;">' +
                escapeHtml(t.ownersUnmatchedHint) +
                '</p>' +
                '<div class="slg-section-title" style="margin-top:14px;">Geplante Besitzer (Direktion)</div>' +
                '<div id="slgOwnerPreview" class="slg-owner-member-box"></div>' +
                '</div>';
        }
        html +=
            '<div id="slgOwnersMatched"' +
            (f.matchUi ? ' style="display:none;"' : '') +
            '>' +
            '<div id="slgOwnerSingleWrap">' +
            '<h3 style="margin:0 0 10px;color:#32325d;font-size:1.05em;">Besitzer</h3>' +
            '<div class="hint" style="margin-bottom:10px;">Live aus Microsoft Graph. Der letzte Besitzer kann nicht entfernt werden.</div>' +
            '<div style="display:grid;grid-template-columns:1fr auto;gap:10px;align-items:end;">' +
            '<div class="field" style="margin:0;">' +
            '<label for="slgOwnerSearch">Benutzer suchen</label>' +
            '<input type="text" id="slgOwnerSearch" placeholder="Name oder UPN/E‑Mail" autocomplete="off">' +
            '</div>' +
            '<button type="button" class="btn" id="slgOwnerSearchBtn" style="margin:0;"><i class="bi bi-search"></i>Suchen</button>' +
            '</div>' +
            '<div class="field" style="margin-top:10px;">' +
            '<label for="slgOwnerSearchResults">Treffer</label>' +
            '<div id="slgOwnerSearchResults" class="slg-user-checklist" aria-label="Treffer"></div>' +
            '</div>' +
            '<div class="detail-actions" style="margin-top:10px;">' +
            '<button type="button" class="btn btn-success" id="slgOwnerAddBtn"><i class="bi bi-person-plus"></i>Besitzer hinzufügen</button>' +
            '<button type="button" class="btn" id="slgOwnersReloadBtn"><i class="bi bi-arrow-clockwise"></i>Besitzer neu laden</button>';
        if (f.ensureDirektion) {
            html +=
                '<button type="button" class="btn" id="slgOwnersEnsureDirektionBtn"><i class="bi bi-shield-check"></i>Direktion setzen</button>';
        }
        html +=
            '</div>' +
            '<div id="slgOwnersList" class="slg-owner-member-box" style="max-height:320px;" aria-label="Besitzer-Liste"></div>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div id="slgTabMembers" class="tab-panel" role="tabpanel" data-slg-tab-content="members" aria-labelledby="slgTabBtnMembers">';
        if (f.matchUi) {
            html +=
                '<div id="slgMembersUnmatched">' +
                '<p class="muted" style="margin:0;">' +
                escapeHtml(t.membersUnmatchedHint) +
                '</p>' +
                '<div class="slg-section-title" style="margin-top:14px;">' +
                escapeHtml(t.membersUnmatchedTitle) +
                '</div>' +
                '<div id="slgMemberPreview" class="slg-owner-member-box"></div>' +
                '</div>';
        }
        html +=
            '<div id="slgMembersMatched"' +
            (f.matchUi ? ' style="display:none;"' : '') +
            '>' +
            '<h3 style="margin:0 0 10px;color:#32325d;font-size:1.05em;">Mitglieder</h3>' +
            '<div class="hint" style="margin-bottom:10px;">' +
            escapeHtml(t.membersMatchedHint) +
            '</div>';
        if (f.syncMembers) {
            html +=
                '<div class="detail-actions" style="margin-top:0;">' +
                '<button type="button" class="btn btn-success" id="slgBtnSync"><i class="bi bi-person-plus-fill"></i>Mitglieder synchronisieren</button>';
            if (f.membershipReview) {
                html +=
                    '<button type="button" class="btn" id="slgBtnMembershipReview"><i class="bi bi-intersect"></i>Mitglieder vergleichen</button>';
            }
            html += '</div>';
        }
        html +=
            '<div style="display:grid;grid-template-columns:1fr auto;gap:10px;align-items:end;margin-top:14px;">' +
            '<div class="field" style="margin:0;">' +
            '<label for="slgMemberSearch">Benutzer suchen</label>' +
            '<input type="text" id="slgMemberSearch" placeholder="Name oder UPN/E‑Mail" autocomplete="off">' +
            '</div>' +
            '<button type="button" class="btn" id="slgMemberSearchBtn" style="margin:0;"><i class="bi bi-search"></i>Suchen</button>' +
            '</div>' +
            '<div class="field" style="margin-top:10px;">' +
            '<label for="slgMemberSearchResults">Treffer</label>' +
            '<div id="slgMemberSearchResults" class="slg-user-checklist" aria-label="Treffer"></div>' +
            '</div>' +
            '<div class="detail-actions" style="margin-top:10px;">' +
            '<button type="button" class="btn btn-success" id="slgMemberAddBtn"><i class="bi bi-person-plus"></i>Mitglied hinzufügen</button>' +
            '<button type="button" class="btn" id="slgMembersReloadBtn"><i class="bi bi-arrow-clockwise"></i>Mitglieder neu laden</button>' +
            '</div>' +
            '<div id="slgMembersList" class="slg-owner-member-box" style="max-height:360px;" aria-label="Mitglieder-Liste"></div>';
        if (f.syncMembers) {
            html +=
                '<div class="slg-section-title" style="margin-top:16px;">Protokoll (Listen‑Sync)</div>' +
                '<div id="slgSyncLog" class="slg-sync-log"></div>';
        }
        html += '</div></div></div></div>';
        return html;
    }

    function resolveHost(target) {
        if (!target) return document.getElementById('groupDetailHost');
        if (typeof target === 'string') return document.querySelector(target);
        return target;
    }

    function normStr(v) {
        return String(v ?? '').trim();
    }

    function gate(fn) {
        if (typeof fn !== 'function') return { ok: true };
        const r = fn();
        if (r === false) return { ok: false, message: '' };
        if (r && r.ok === false) return r;
        return { ok: true };
    }

    function getGroupId() {
        if (session && session.live && typeof session.live.getGroupId === 'function') {
            return session.live.getGroupId();
        }
        return null;
    }

    function setTab(tab) {
        activeTab = tab === 'owners' || tab === 'members' ? tab : 'general';
        document.querySelectorAll('#slgDetailTabs .detail-tab-btn[data-slg-tab]').forEach(function (b) {
            const on = b.getAttribute('data-slg-tab') === activeTab;
            b.setAttribute('aria-selected', on ? 'true' : 'false');
        });
        document.querySelectorAll('[data-slg-tab-content]').forEach(function (p) {
            p.classList.toggle('active', p.getAttribute('data-slg-tab-content') === activeTab);
        });
        const gid = getGroupId();
        if (!gid) {
            if (session && typeof session.onTabUnmatched === 'function') session.onTabUnmatched(activeTab);
            return;
        }
        liveApi().onTab(activeTab, gid);
    }

    function clearSearchResults() {
        const host = document.getElementById('slgGroupSearchResults');
        if (!host) return;
        host.replaceChildren();
        host.style.display = 'none';
    }

    function renderGroupSearchResults(list) {
        const host = document.getElementById('slgGroupSearchResults');
        if (!host) return;
        host.replaceChildren();
        if (!list || !list.length) {
            host.style.display = 'block';
            const p = document.createElement('div');
            p.className = 'muted';
            p.textContent = 'Keine passenden Microsoft 365‑Gruppen (Unified) gefunden.';
            host.appendChild(p);
            return;
        }
        host.style.display = 'block';
        const box = document.createElement('div');
        box.style.border = '1px solid #ced4da';
        box.style.borderRadius = '12px';
        box.style.background = '#fff';
        box.style.overflow = 'hidden';
        list.forEach(function (g, idx) {
            const row = document.createElement('div');
            row.className = 'slg-search-result-row';
            if (idx === 0) row.style.borderTop = '0';
            const dn = normStr(g && g.displayName) || '(ohne Namen)';
            const mail = normStr(g && g.mail) || '–';
            const nick = normStr(g && g.mailNickname) || '–';
            const gid = normStr(g && g.id);
            if (window.ms365GroupPhotoThumb && typeof window.ms365GroupPhotoThumb.createThumb === 'function') {
                row.appendChild(
                    window.ms365GroupPhotoThumb.createThumb({
                        groupId: gid,
                        displayName: dn,
                        size: 'search'
                    })
                );
            }
            const left = document.createElement('div');
            left.className = 'slg-search-result-row__main';
            left.innerHTML =
                '<div class="slg-search-result-row__title">' +
                escapeHtml(dn) +
                '</div>' +
                '<div class="muted slg-search-result-row__meta">Mail‑Nickname: <code>' +
                escapeHtml(nick) +
                '</code> · SMTP: ' +
                escapeHtml(mail) +
                '</div>' +
                '<div class="muted slg-search-result-row__meta">Gruppen‑ID: <code>' +
                escapeHtml(gid) +
                '</code></div>';
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn';
            btn.textContent = 'Matchen';
            btn.addEventListener('click', function () {
                applyMatch(g, 'matched');
            });
            row.appendChild(left);
            row.appendChild(btn);
            box.appendChild(row);
        });
        host.appendChild(box);
        if (window.ms365GroupPhotoThumb && typeof window.ms365GroupPhotoThumb.hydrate === 'function') {
            window.ms365GroupPhotoThumb.hydrate(box);
        }
    }

    function applyMatch(g, mode) {
        if (!g || !g.id) return;
        const m = session && session.match;
        if (m && typeof m.persistMatch === 'function') m.persistMatch(g, mode);
        liveApi().fillForm(g);
        liveApi().setMatchedMode(true);
        liveApi().loadGroup({ silent: true });
        if (m && typeof m.afterMatch === 'function') m.afterMatch(g, mode);
        toast(mode === 'created' ? 'Gruppe angelegt und gematcht.' : 'Gruppe gematcht.');
    }

    async function runSearchGroups() {
        const m = session && session.match;
        const g1 = gate(m && m.canSearch);
        if (!g1.ok) {
            if (g1.message) toast(g1.message);
            return;
        }
        const inp = document.getElementById('slgGroupSearch');
        const q = inp && inp.value ? inp.value.trim() : '';
        if (!q) {
            toast('Bitte einen Suchbegriff eingeben.');
            return;
        }
        try {
            const token = await gug().getGraphToken();
            const list = await gug().searchUnifiedGroups(token, q);
            renderGroupSearchResults(list);
            if (!list.length) toast('Keine passenden Gruppen gefunden.');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function runCreateAndMatch() {
        const m = session && session.match;
        const g1 = gate(m && m.canCreate);
        if (!g1.ok) {
            if (g1.message) toast(g1.message);
            return;
        }
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const dd = document.getElementById('slgNewDescription');
        const ct = document.getElementById('slgNewCreateTeam');
        const displayName = dn ? dn.value : '';
        const mailNick = nn ? nn.value : '';
        const desc = dd ? dd.value : '';
        if (!normStr(displayName) || !normStr(mailNick)) {
            toast('Bitte Anzeigename und Alias/Mail‑Nickname ausfüllen.');
            return;
        }
        try {
            const token = await gug().getGraphToken();
            const g = await gug().createUnifiedGroup(token, displayName, mailNick, desc);
            if (m && typeof m.ensureOwners === 'function') await m.ensureOwners(token, g.id);
            if (m && typeof m.afterCreate === 'function') await m.afterCreate(token, g);
            if (ct && ct.checked) {
                toast('Gruppe angelegt – Team wird bereitgestellt …');
                await gug().provisionTeamForGroup(token, g.id);
            }
            applyMatch(g, 'created');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    function runUnmatch() {
        if (!getGroupId()) return;
        const m = session && session.match;
        if (m && typeof m.persistUnmatch === 'function') m.persistUnmatch();
        liveApi().loadGroup({ silent: true });
        if (m && typeof m.afterMatch === 'function') m.afterMatch(null, 'unmatch');
        toast('Match gelöst.');
    }

    function openEntraForMatched() {
        const gid = getGroupId();
        if (!gid) return;
        window.open(ENTRA_GROUP + encodeURIComponent(gid), '_blank', 'noopener');
    }

    function onClick(id, fn) {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', fn);
    }

    function wireUi(host) {
        if (!host || host.getAttribute(ATTR_WIRED) === '1') return;
        host.setAttribute(ATTR_WIRED, '1');
        document.querySelectorAll('#slgDetailTabs .detail-tab-btn[data-slg-tab]').forEach(function (b) {
            b.addEventListener('click', function () {
                setTab(b.getAttribute('data-slg-tab') || 'general');
            });
        });
        onClick('slgBtnSearch', function () {
            runSearchGroups();
        });
        onClick('slgBtnCreate', function () {
            runCreateAndMatch();
        });
        onClick('slgBtnUnmatch', runUnmatch);
        onClick('slgBtnOpenEntra', openEntraForMatched);
        onClick('slgBtnDeleteGroup', function () {
            if (session && session.live && typeof session.live.onDelete === 'function') {
                session.live.onDelete();
            }
        });
        const groupSearch = document.getElementById('slgGroupSearch');
        if (groupSearch) {
            groupSearch.addEventListener('keydown', function (ev) {
                if (ev.key === 'Enter') {
                    ev.preventDefault();
                    runSearchGroups();
                }
            });
        }
    }

    function bindLive(opts) {
        const liveCtx = (opts && opts.live) || {};
        liveApi().bind({
            toast: liveCtx.toast,
            dlgConfirm: liveCtx.dlgConfirm,
            getGroupId: liveCtx.getGroupId,
            getGraphToken: liveCtx.getGraphToken,
            confirmUpdate: liveCtx.confirmUpdate,
            alwaysMatched: featureBag(opts).matchUi === false,
            getActiveTab: function () {
                return activeTab;
            },
            ensureDirektionOwners: liveCtx.ensureDirektionOwners,
            onUnmatched: liveCtx.onUnmatched,
            onAfterLoad: liveCtx.onAfterLoad,
            onAfterUpdate: liveCtx.onAfterUpdate,
            onMembersCount: liveCtx.onMembersCount
        });
        if (!liveWired) {
            liveApi().wire();
            liveWired = true;
        }
    }

    function attachAfterWrap(host, ids) {
        if (!ids.afterWrap) return;
        const extra = document.getElementById(ids.afterWrap);
        const body = (host && host.querySelector('.panel-body')) || host;
        if (!extra || !body || extra.parentNode === body) return;
        body.appendChild(extra);
    }

    function attachOwnerExtra(ids) {
        if (!ids.ownerExtra) return;
        const extra = document.getElementById(ids.ownerExtra);
        const host = document.getElementById('slgOwnersMatched');
        if (!extra || !host) return;
        if (extra.parentNode === host && host.firstElementChild === extra) return;
        host.insertBefore(extra, host.firstChild);
    }

    function mount(target, opts) {
        const host = resolveHost(target);
        if (!host) {
            throw new Error('group-detail: Host #groupDetailHost nicht gefunden.');
        }
        const already = host.getAttribute(ATTR_MOUNTED) === '1' && host.querySelector('#slgLiveName');
        if (!already) {
            host.innerHTML = buildMarkup(opts);
            host.setAttribute(ATTR_MOUNTED, '1');
        }
        session = {
            live: (opts && opts.live) || {},
            match: (opts && opts.match) || {},
            onTabUnmatched: opts && opts.onTabUnmatched
        };
        attachAfterWrap(host, idBag(opts));
        attachOwnerExtra(idBag(opts));
        if (opts && opts.live) bindLive(opts);
        wireUi(host);
        return host;
    }

    window.ms365GroupDetail = {
        mount: mount,
        buildMarkup: buildMarkup,
        setTab: setTab,
        getActiveTab: function () {
            return activeTab;
        },
        clearSearchResults: clearSearchResults
    };
})();
