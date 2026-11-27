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
            visibilityUnsupported: !!f.visibilityUnsupported,
            createTargetChoice: !!f.createTargetChoice
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
                    ? 'Verknüpfte Microsoft 365‑Gruppe. <strong>Änderungen speichern</strong> sichert Anzeigename, Alias, Beschreibung und Sichtbarkeit.'
                    : f.matchUi
                      ? 'Verknüpfte Microsoft 365‑Gruppe. <strong>Änderungen speichern</strong> sichert Anzeigename, Beschreibung und Sichtbarkeit.'
                      : 'Microsoft 365‑Gruppe im Tenant. <strong>Änderungen speichern</strong> sichert Anzeigename, Beschreibung und Sichtbarkeit.'),
            updateHintHtml:
                o.updateHintHtml ||
                (aliasInUpdate
                    ? '<strong>Änderungen speichern</strong> sichert <strong>Anzeigename</strong>, <strong>Alias</strong>, <strong>Beschreibung</strong> und <strong>Sichtbarkeit</strong>. Zum Wechseln: im Overflow‑Menü „M365-Verknüpfung lösen".'
                    : f.matchUi
                      ? '<strong>Änderungen speichern</strong> sichert <strong>Anzeigename</strong>, <strong>Beschreibung</strong> und <strong>Sichtbarkeit</strong>. Zum Wechseln: im Overflow‑Menü „M365-Verknüpfung lösen".'
                      : '<strong>Änderungen speichern</strong> sichert <strong>Anzeigename</strong>, <strong>Beschreibung</strong> und <strong>Sichtbarkeit</strong> (Unified) sowie optional den <strong>Teams‑Archiv‑Status</strong>. E‑Mail/Alias bleiben unverändert.'),
            ownersUnmatchedHint:
                o.ownersUnmatchedHint ||
                'Besitzer werden aus der Verwaltungsliste (Rolle „Direktion") abgeleitet und beim Anlegen gesetzt. Nach dem Match erscheinen hier die Live‑Besitzer aus Microsoft Graph.',
            membersUnmatchedHint:
                o.membersUnmatchedHint ||
                'Nach dem Match können Mitglieder live in Microsoft Graph gepflegt werden.',
            membersUnmatchedTitle: o.membersUnmatchedTitle || 'Vorschau',
            membersMatchedHint:
                o.membersMatchedHint ||
                (f.syncMembers
                    ? 'Live aus Microsoft Graph. Mitglieder hier suchen, hinzufügen oder entfernen. Der Microsoft‑365‑Abgleich vergleicht auf Wunsch die Gruppe mit den Stammdaten.'
                    : 'Live aus Microsoft Graph. Mitglieder hier suchen, hinzufügen oder entfernen.')
        };
    }

    /* ───── buildMarkup ───── */

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

        /* ── Object header ── */
        if (f.header) {
            html += '<div class="gd-object-header">';

            html +=
                '<div class="gd-object-header__avatar">' +
                '<div class="gd-group-photo" id="slgGroupPhotoWrap">' +
                '<div class="gd-group-photo__avatar" id="slgGroupPhotoAvatar" aria-hidden="true">' +
                '<img id="slgGroupPhotoImg" alt="" hidden>' +
                '<span id="slgGroupPhotoInitials" class="gd-group-photo__initials">\u2013</span>' +
                '</div>' +
                '<div class="gd-group-photo__body"></div>' +
                '</div>' +
                '<div class="gd-photo-actions">' +
                '<label class="btn" for="slgGroupPhotoFile">Bild ändern</label>' +
                '<input type="file" id="slgGroupPhotoFile" accept="image/jpeg,image/png,image/webp" hidden>' +
                '<details class="gd-overflow" id="slgPhotoOverflow">' +
                '<summary aria-label="Foto-Optionen">\u00B7\u00B7\u00B7</summary>' +
                '<div class="gd-overflow__menu" role="menu">' +
                '<button type="button" id="slgBtnRemoveGroupPhoto" class="gd-overflow__danger" role="menuitem" hidden>Entfernen</button>' +
                '</div></details>' +
                '</div>' +
                '</div>';

            html +=
                '<div class="gd-object-header__main">' +
                '<h2 id="slgDetailTitle">' + escapeHtml(t.title) + '</h2>' +
                '<div class="gd-object-header__kind-line" id="slgDetailStatusRow" hidden>' +
                '<span id="slgDetailSubtitle">' + escapeHtml(t.subtitle) + '</span>' +
                '<span class="gd-sync-badge gd-sync-badge--ok" id="slgDetailSyncBadge"><i class="bi bi-check-circle-fill" aria-hidden="true"></i>Verknüpft</span>' +
                '</div>' +
                '<div class="gd-object-header__counts" id="slgDetailCounts" hidden></div>' +
                '</div>';

            html += '<div class="gd-object-header__actions"' + headId + '>';
            if (f.openEntra) {
                html +=
                    '<button type="button" class="btn" id="slgBtnOpenEntra" disabled><i class="bi bi-box-arrow-up-right"></i>In Entra öffnen</button>';
            }
            if (f.matchUi || f.deleteGroup) {
                html +=
                    '<details class="gd-overflow">' +
                    '<summary aria-label="Weitere Aktionen">\u00B7\u00B7\u00B7</summary>' +
                    '<div class="gd-overflow__menu" role="menu">';
                if (f.matchUi) {
                    html +=
                        '<button type="button" id="slgBtnUnmatch" role="menuitem" disabled>M365-Verknüpfung lösen</button>';
                }
                if (f.deleteGroup) {
                    html +=
                        '<button type="button" id="slgBtnDeleteGroup" class="gd-overflow__danger" role="menuitem">Löschen</button>';
                }
                html += '</div></details>';
            }
            html += '</div></div>';
        }

        /* ── Panel body / embed ── */
        html += f.header ? '<div class="panel-body">' : '<div class="gd-embed">';
        if (f.emptyHint) {
            html +=
                '<div id="' + escapeHtml(ids.emptyHint) + '" class="hint" style="display:none;">' +
                t.emptyHintHtml +
                '</div>';
        }
        html +=
            '<div id="' + escapeHtml(ids.wrap) + '">';

        /* ── Tabs ── */
        html +=
            '<div class="detail-tabs-sticky" aria-label="Detail-Tabs">' +
            '<div class="detail-tabs" role="tablist" id="slgDetailTabs" aria-label="Tabs Details">' +
            '<button type="button" class="detail-tab-btn" id="slgTabBtnGeneral" role="tab" data-slg-tab="general" aria-selected="true" aria-controls="slgTabGeneral" tabindex="0">Allgemein</button>' +
            '<button type="button" class="detail-tab-btn" id="slgTabBtnOwners" role="tab" data-slg-tab="owners" aria-selected="false" aria-controls="slgTabOwners" tabindex="-1">Besitzer <span class="gd-tab-count" id="slgTabOwnersCount" hidden></span></button>' +
            '<button type="button" class="detail-tab-btn" id="slgTabBtnMembers" role="tab" data-slg-tab="members" aria-selected="false" aria-controls="slgTabMembers" tabindex="-1">Mitglieder <span class="gd-tab-count" id="slgTabMembersCount" hidden></span></button>' +
            '</div></div>';

        /* ── General tab ── */
        html +=
            '<div id="slgTabGeneral" class="tab-panel active" role="tabpanel" data-slg-tab-content="general" aria-labelledby="slgTabBtnGeneral">';

        /* Unmatched panel (match search + create) */
        if (f.matchUi) {
            html +=
                '<div id="slgUnmatchedPanel">' +
                '<div class="field" style="margin-top:0;">' +
                '<label for="slgGroupSearch">Bestehende Gruppe suchen (Name, Mail, Alias, Beschreibung)</label>' +
                '<div class="slg-search-wrap">' +
                '<input type="text" id="slgGroupSearch" autocomplete="off" placeholder="' +
                escapeHtml(t.searchPlaceholder) +
                '">' +
                '<button type="button" class="btn alt" id="slgBtnSearch"><i class="bi bi-search"></i>Suchen</button>' +
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
                '<div class="field" style="grid-column:1 / -1;">' +
                '<label for="slgNewMailNick">Unique ID / Alias</label>' +
                '<div class="gd-mailnick-row">' +
                '<input type="text" id="slgNewMailNick" maxlength="60" autocomplete="off" spellcheck="false" inputmode="latin" placeholder="z. B. arge-fvv">' +
                (opts && opts.match && typeof opts.match.suggestMailNickname === 'function'
                    ? '<button type="button" class="btn btn-ghost" id="slgBtnMailNickStandard" title="Vorschlag aus Präfix + Kürzel einsetzen">Standard</button>'
                    : '') +
                '</div>' +
                '<small class="muted" id="slgNewMailNickHint" style="display:block;margin-top:6px;line-height:1.4;">' +
                (opts && opts.match && typeof opts.match.suggestMailNickname === 'function'
                    ? 'Frei editierbar. Leer lassen oder „Standard": Vorschlag aus Präfix + Kürzel.'
                    : 'Alias der Microsoft‑365‑Gruppe (mailNickname).') +
                '</small>' +
                '</div>' +
                '<div class="field" style="grid-column:1 / -1;">' +
                '<label for="slgNewDescription">Beschreibung</label>' +
                '<input type="text" id="slgNewDescription" maxlength="512" autocomplete="off">' +
                '</div>' +
                '</div>' +
                (f.createTargetChoice
                    ? '<fieldset class="gd-create-target" style="margin-top:12px;border:1px solid var(--border);border-radius:12px;padding:10px 12px;">' +
                      '<legend style="padding:0 6px;font-size:0.85em;font-weight:700;">Ziel beim Anlegen</legend>' +
                      '<label class="checkbox-label" for="slgNewCreateTargetMail" style="display:flex;align-items:flex-start;gap:8px;margin:0 0 8px;">' +
                      '<input type="radio" name="slgNewCreateTarget" id="slgNewCreateTargetMail" value="mail" checked>' +
                      '<span><strong>Nur E-Mail-Verteiler</strong> – Microsoft 365‑Gruppe ohne Team (mailfähig)</span>' +
                      '</label>' +
                      '<label class="checkbox-label" for="slgNewCreateTargetTeam" style="display:flex;align-items:flex-start;gap:8px;margin:0;">' +
                      '<input type="radio" name="slgNewCreateTarget" id="slgNewCreateTargetTeam" value="team">' +
                      '<span><strong>Microsoft 365‑Gruppe inkl. Team</strong> – mit Teams‑Arbeitsbereich</span>' +
                      '</label>' +
                      '</fieldset>'
                    : '<label class="checkbox-label" for="slgNewCreateTeam" style="margin-top:12px; display:inline-flex;">' +
                      '<input type="checkbox" id="slgNewCreateTeam">' +
                      '<span>Auch als <strong>Team</strong> anlegen (optional)</span>' +
                      '</label>') +
                '<div class="detail-actions">' +
                '<button type="button" class="btn btn-success" id="slgBtnCreate"><i class="bi bi-plus-circle"></i>Anlegen &amp; matchen</button>' +
                '</div>' +
                '</div>' +
                '</div>';
        }

        /* Matched panel */
        html +=
            '<div id="slgMatchedPanel"' +
            (f.matchUi ? ' style="display:none;"' : '') +
            '>';

        if (!f.header) {
            html +=
                '<div class="gd-group-photo" id="slgGroupPhotoWrap">' +
                '<div class="gd-group-photo__avatar" id="slgGroupPhotoAvatar" aria-hidden="true">' +
                '<img id="slgGroupPhotoImg" alt="" hidden>' +
                '<span id="slgGroupPhotoInitials" class="gd-group-photo__initials">\u2013</span>' +
                '</div>' +
                '<div class="gd-group-photo__body">' +
                '<div class="gd-group-photo__title">Gruppenbild</div>' +
                '<div class="detail-actions gd-group-photo__actions">' +
                '<label class="btn" for="slgGroupPhotoFile"><i class="bi bi-image"></i>Bild hochladen</label>' +
                '<input type="file" id="slgGroupPhotoFile" accept="image/jpeg,image/png,image/webp" hidden>' +
                '<button type="button" class="btn btn-ghost" id="slgBtnRemoveGroupPhoto" hidden><i class="bi bi-trash"></i>Entfernen</button>' +
                '</div>' +
                '<p class="muted gd-group-photo__hint">JPEG, PNG oder WebP, max. 4&nbsp;MB.</p>' +
                '</div>' +
                '</div>';
        }

        html +=
            '<div class="gd-save-feedback" id="slgSaveFeedback" role="status" aria-live="polite">' +
            '<i class="bi bi-check-circle-fill" aria-hidden="true"></i>Gespeichert</div>';

        /* Editable fields */
        html +=
            '<div class="form-grid" style="margin-top:0;">' +
            '<div class="field field-editable">' +
            '<label for="slgLiveName">Anzeigename</label>' +
            '<input type="text" id="slgLiveName" maxlength="200" autocomplete="off">' +
            '</div>' +
            '<div class="field ' + aliasClass + '">' +
            '<label for="slgLiveAlias">Alias</label>' +
            '<input type="text" id="slgLiveAlias" maxlength="60" autocomplete="off" spellcheck="false"' +
            aliasRo +
            '>' +
            '</div>' +
            '<div class="field field-editable" style="grid-column:1 / -1;">' +
            '<label for="slgLiveDescription">Beschreibung</label>' +
            '<textarea id="slgLiveDescription" maxlength="1024" placeholder="Beschreibung der Microsoft\u200965\u2011Gruppe"></textarea>' +
            '</div>' +
            '<div class="field field-editable">' +
            '<label for="slgLiveVisibility">Sichtbarkeit</label>' +
            '<select id="slgLiveVisibility">' +
            (f.visibilityUnsupported ? '<option value="">(nicht unterstützt)</option>' : '') +
            '<option value="Private">Privat</option>' +
            '<option value="Public">Öffentlich</option>' +
            '</select>' +
            '</div>' +
            '</div>';

        /* Read-only metadata (gd-meta blocks) */
        html +=
            '<div class="gd-section" style="margin-top:18px;">' +
            '<div class="gd-section__title">Microsoft 365</div>' +
            '<div class="gd-meta-grid">';

        html +=
            '<div class="gd-meta">' +
            '<div class="gd-meta__label">E\u2011Mail</div>' +
            '<div class="gd-meta__row"><span class="gd-meta__value" id="slgLiveMail">\u2013</span></div>';
        if (f.smtpSlot) {
            html +=
                '<details class="jg-smtp-drop" id="jgSmtpDrop">' +
                '<summary>SMTP auf Schul\u2011Domain …</summary>' +
                '<div class="jg-smtp-drop__body">' +
                '<div id="jgSmtpHint">\u2013</div>' +
                '<div class="detail-actions" style="margin-top:10px;">' +
                '<button type="button" class="btn" id="jgBtnSmtpThis"><i class="bi bi-envelope-at"></i>Skript erzeugen</button>' +
                '</div>' +
                '<textarea id="jgSmtpScript" readonly spellcheck="false" aria-label="Exchange-Skript für SMTP"></textarea>' +
                '</div>' +
                '</details>';
        }
        html += '</div>';

        html +=
            '<div class="gd-meta">' +
            '<div class="gd-meta__label">Gruppen\u2011ID</div>' +
            '<div class="gd-meta__row"><span class="gd-meta__value gd-meta__value--mono" id="slgLiveId">\u2013</span></div>' +
            '</div>';

        html +=
            '<div class="gd-meta">' +
            '<div class="gd-meta__label">Art</div>' +
            '<div class="gd-meta__row"><span class="gd-meta__value" id="slgLiveArt">\u2013</span></div>' +
            '</div>';

        if (f.createdAt) {
            html +=
                '<div class="gd-meta">' +
                '<div class="gd-meta__label">Erstellt am</div>' +
                '<div class="gd-meta__row"><span class="gd-meta__value" id="slgLiveCreated">\u2013</span></div>' +
                '</div>';
        }

        html +=
            '<div class="gd-meta">' +
            '<div class="gd-meta__label">Ablaufdatum</div>' +
            '<div class="gd-meta__row"><span class="gd-meta__value" id="slgLiveExpires">\u2013</span></div>';
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
                '<div class="gd-meta" style="grid-column:1 / -1;">' +
                '<div class="gd-meta__label">Microsoft Teams</div>' +
                '<div class="gd-meta__row">' +
                '<div class="jg-team-quiet">' +
                '<span id="slgLiveTeam" class="jg-team-status">\u2013</span>' +
                '<button type="button" class="jg-quiet-btn" id="slgBtnProvisionTeam" hidden>Team anlegen</button>' +
                '<a id="slgLiveTeamLink" class="jg-quiet-link" href="#" target="_blank" rel="noopener noreferrer" hidden>In Teams öffnen</a>' +
                '<span id="slgLiveTeamHint" class="muted" hidden></span>' +
                '</div></div></div>';
        }

        html += '</div></div>';

        /* Team archive */
        if (f.teamArchive) {
            html +=
                '<div id="slgTeamArchiveWrap" class="gd-archive-card" style="display:none;">' +
                '<div class="gd-archive-card__head">' +
                '<span class="gd-archive-card__icon" aria-hidden="true"><i class="bi bi-archive"></i></span>' +
                '<div>' +
                '<div class="gd-archive-card__title">Teams\u2011Archiv</div>' +
                '<p class="gd-archive-card__lead">Nur für Microsoft 365\u2011Gruppen mit gebundenem Team (Allgemeines Team, Kursteam usw.). Reine Gruppen ohne Team können hier nicht archiviert werden.</p>' +
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
                '<strong>SharePoint\u2011Website schreibgeschützt</strong>' +
                '<small>Beim Archivieren dürfen Mitglieder Dateien auf der Teamwebsite nicht mehr ändern. Nur mit delegierter (persönlicher) Anmeldung.</small>' +
                '</span>' +
                '</label>' +
                '</div>' +
                '</div>';
        }

        /* Detail actions (save, refresh, etc.) */
        html +=
            '<div class="detail-actions gd-detail-actions">' +
            '<div class="gd-detail-actions__primary">' +
            '<button type="button" class="btn btn-success" id="slgBtnUpdateGroup" disabled><i class="bi bi-save"></i>Änderungen speichern</button>' +
            '<button type="button" class="btn btn-ghost" id="slgBtnRefreshGroup"><i class="bi bi-arrow-clockwise"></i>Neu laden</button>';
        if (!f.header && f.openEntra) {
            html +=
                '<button type="button" class="btn btn-ghost" id="slgBtnOpenEntra" disabled><i class="bi bi-box-arrow-up-right"></i>In Entra öffnen</button>';
        }
        html += '</div>';
        if (!f.header && f.deleteGroup) {
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
            '</div>';

        /* ── Owners tab ── */
        html +=
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
            '<div class="gd-people-head">' +
            '<h3 class="gd-people-head__title">BESITZER</h3>' +
            '<span class="gd-people-head__count" id="slgOwnersHeadCount">\u2013</span>' +
            '</div>' +
            '<p class="muted gd-member-target" id="slgOwnerTarget" hidden></p>' +
            '<div class="gd-people-toolbar">' +
            '<div class="field">' +
            '<label class="visually-hidden" for="slgOwnerSearch">Besitzer suchen oder hinzufügen</label>' +
            '<input type="text" id="slgOwnerSearch" placeholder="Besitzer suchen oder hinzufügen …" autocomplete="off">' +
            '</div>' +
            '<button type="button" class="btn btn-success" id="slgOwnerAddBtn"><i class="bi bi-plus-lg"></i>Hinzufügen</button>' +
            '</div>' +
            '<button type="button" class="btn btn-ghost visually-hidden" id="slgOwnerSearchBtn" tabindex="-1">Suchen</button>' +
            '<div id="slgOwnerSearchResults" class="slg-user-checklist gd-people-results" aria-label="Treffer"></div>' +
            '<div class="detail-actions" style="margin:0 0 10px; gap:8px;">' +
            '<button type="button" class="btn btn-ghost" id="slgOwnersReloadBtn"><i class="bi bi-arrow-clockwise"></i>Aktualisieren</button>';
        if (f.ensureDirektion) {
            html +=
                '<button type="button" class="btn btn-ghost" id="slgOwnersEnsureDirektionBtn"><i class="bi bi-shield-check"></i>Direktion als Besitzer hinzufügen</button>';
        }
        html +=
            '</div>' +
            '<div id="slgOwnersList" class="slg-owner-member-box" data-child="gd-person-row" aria-label="Besitzer-Liste"></div>' +
            '</div>' +
            '</div>' +
            '</div>';

        /* ── Members tab ── */
        html +=
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
            '<div class="gd-people-head">' +
            '<h3 class="gd-people-head__title">MITGLIEDER</h3>' +
            '<span class="gd-people-head__count" id="slgMembersHeadCount">\u2013</span>' +
            '</div>' +
            '<p class="muted gd-member-target" id="slgMemberTarget" hidden></p>' +
            '<div class="gd-people-toolbar">' +
            '<div class="field">' +
            '<label class="visually-hidden" for="slgMemberSearch">Mitglieder suchen oder hinzufügen</label>' +
            '<input type="text" id="slgMemberSearch" placeholder="Mitglieder suchen oder hinzufügen …" autocomplete="off">' +
            '</div>' +
            '<button type="button" class="btn btn-success" id="slgMemberAddBtn"><i class="bi bi-plus-lg"></i>Hinzufügen</button>' +
            '</div>' +
            '<button type="button" class="btn btn-ghost visually-hidden" id="slgMemberSearchBtn" tabindex="-1">Suchen</button>' +
            '<div id="slgMemberSearchResults" class="slg-user-checklist gd-people-results" aria-label="Treffer"></div>' +
            '<div class="detail-actions" style="margin:0 0 10px;">' +
            '<button type="button" class="btn btn-ghost" id="slgMembersReloadBtn"><i class="bi bi-arrow-clockwise"></i>Aktualisieren</button>' +
            '</div>' +
            '<div id="slgMembersList" class="slg-owner-member-box" data-child="gd-person-row" aria-label="Mitglieder-Liste"></div>';

        /* Sync card */
        if (f.syncMembers) {
            html +=
                '<div class="gd-sync-card" id="slgSyncTrustHost">' +
                '<div class="gd-sync-card__title">Microsoft\u201065\u2011Abgleich</div>' +
                '<p class="gd-sync-card__status" id="slgSyncStatusLine">Mitglieder mit Stammdaten abgleichen.</p>' +
                '<div class="gd-sync-card__actions">' +
                '<button type="button" class="btn" id="slgBtnSync"><i class="bi bi-search"></i>Änderungen prüfen</button>';
            if (f.membershipReview) {
                html +=
                    '<button type="button" class="btn btn-ghost" id="slgBtnMembershipReview"><i class="bi bi-intersect"></i>Vergleichen</button>';
            }
            html +=
                '<details class="gd-overflow" id="slgSyncOverflow">' +
                '<summary aria-label="Weitere Sync-Aktionen">Weitere Aktionen</summary>' +
                '<div class="gd-overflow__menu" role="menu">' +
                '<button type="button" id="slgBtnShowSyncLog" role="menuitem">Synchronisationsprotokoll</button>' +
                '</div></details>' +
                '</div>' +
                '<div class="gd-sync-preview" id="slgSyncPreview" hidden>' +
                '<div class="gd-sync-preview__title">Änderungsvorschau</div>' +
                '<ul class="gd-sync-preview__lines" id="slgSyncPreviewLines"></ul>' +
                '<div class="detail-actions" style="margin:0;gap:8px;">' +
                '<button type="button" class="btn btn-ghost" id="slgBtnSyncCancel">Abbrechen</button>' +
                '<button type="button" class="btn btn-success" id="slgBtnSyncApply">Änderungen übernehmen</button>' +
                '</div></div>' +
                '<details class="jg-settings" id="slgSyncSettings" style="margin-top:10px;">' +
                '<summary>Synchronisationseinstellungen</summary>' +
                '<div style="margin-top:8px;">' +
                '<input type="checkbox" id="slgSyncDryRun" checked hidden>' +
                '<label class="checkbox-label" style="margin:0;display:flex;gap:8px;align-items:flex-start;">' +
                '<input type="checkbox" id="slgSyncAllowLeave">' +
                '<span>Personen aus Microsoft 365 entfernen, wenn sie nicht mehr in den Stammdaten enthalten sind.<br><small class="muted">Dies kann bestehende Gruppenmitgliedschaften entfernen.</small></span>' +
                '</label></div></details>' +
                '<div id="slgSyncBulkProgress" class="ms365-bulk-progress" data-ms365-bulk-progress hidden style="margin-top:10px;">' +
                '<div class="ms365-bulk-progress__bar"><div data-fill></div></div>' +
                '<p data-text></p>' +
                '<button type="button" class="btn btn-sm" data-cancel>Abbrechen</button>' +
                '<ul data-errors></ul>' +
                '</div>' +
                '<div class="gd-sync-log-wrap" id="slgSyncLogWrap" hidden>' +
                '<div class="gd-sync-card__title" style="margin-top:12px;">Synchronisationsprotokoll</div>' +
                '<div id="slgSyncLog" class="slg-sync-log"></div>' +
                '</div></div>';
        }

        html += '</div></div></div></div>';
        return html;
    }

    /* ───── Helpers ───── */

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

    function syncMembershipTargetLabels() {
        const titleEl = document.getElementById('slgDetailTitle');
        const name = titleEl ? String(titleEl.textContent || '').trim() : '';
        const text = name ? 'Für diese Gruppe: ' + name : '';
        ['slgOwnerTarget', 'slgMemberTarget'].forEach(function (id) {
            const el = document.getElementById(id);
            if (el) el.textContent = text;
        });
    }

    function setTab(tab) {
        activeTab = tab === 'owners' || tab === 'members' ? tab : 'general';
        document.querySelectorAll('#slgDetailTabs .detail-tab-btn[data-slg-tab]').forEach(function (b) {
            const on = b.getAttribute('data-slg-tab') === activeTab;
            b.setAttribute('aria-selected', on ? 'true' : 'false');
            b.setAttribute('tabindex', on ? '0' : '-1');
        });
        document.querySelectorAll('[data-slg-tab-content]').forEach(function (p) {
            p.classList.toggle('active', p.getAttribute('data-slg-tab-content') === activeTab);
        });
        syncMembershipTargetLabels();
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
            p.textContent = 'Keine passenden Microsoft 365\u2011Gruppen (Unified) gefunden.';
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
            const mail = normStr(g && g.mail) || '\u2013';
            const nick = normStr(g && g.mailNickname) || '\u2013';
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
                '<div class="muted slg-search-result-row__meta">Mail\u2011Nickname: <code>' +
                escapeHtml(nick) +
                '</code> \u00B7 SMTP: ' +
                escapeHtml(mail) +
                '</div>' +
                '<div class="muted slg-search-result-row__meta">Gruppen\u2011ID: <code>' +
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
        const targetTeam = document.getElementById('slgNewCreateTargetTeam');
        const displayName = dn ? dn.value : '';
        let mailNick = nn ? nn.value : '';
        if (!normStr(mailNick) && m && typeof m.suggestMailNickname === 'function') {
            mailNick = String(m.suggestMailNickname() || '').trim();
            if (nn && mailNick) nn.value = mailNick;
        }
        const desc = dd ? dd.value : '';
        const wantTeam = !!(targetTeam && targetTeam.checked) || !!(ct && ct.checked);
        if (!normStr(displayName) || !normStr(mailNick)) {
            toast('Bitte Anzeigename und Unique ID / Alias ausfüllen.');
            return;
        }
        try {
            const token = await gug().getGraphToken();
            const g = await gug().createUnifiedGroup(token, displayName, mailNick, desc);
            if (m && typeof m.ensureOwners === 'function') await m.ensureOwners(token, g.id);
            if (m && typeof m.afterCreate === 'function') await m.afterCreate(token, g);
            if (wantTeam) {
                toast('Gruppe angelegt – Team wird bereitgestellt …');
                await gug().provisionTeamForGroup(token, g.id);
            }
            applyMatch(g, 'created');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    function applySuggestedMailNickname() {
        const m = session && session.match;
        const nn = document.getElementById('slgNewMailNick');
        if (!nn) return;
        if (!m || typeof m.suggestMailNickname !== 'function') {
            toast('Kein Standard-Alias hinterlegt.');
            return;
        }
        const v = String(m.suggestMailNickname() || '').trim();
        nn.value = v;
        nn.dispatchEvent(new Event('input', { bubbles: true }));
        if (v) toast('Standard-Alias eingesetzt: ' + v);
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

    /* ───── Sync helpers (exported) ───── */

    /**
     * Master-Listen-KPI: ruhig bei Sync, Soll/Ist nur bei Abweichung.
     * @param {HTMLElement|null} wrap
     * @param {{ listN: number, groupN: number|null|undefined, linked: boolean, memberWord?: string }} opts
     * @returns {{ known: boolean, same: boolean, delta: number }}
     */
    function paintMasterMemberSyncCounts(wrap, opts) {
        if (!wrap) return { known: false, same: false, delta: 0 };
        var listN = opts && typeof opts.listN === 'number' && opts.listN >= 0 ? opts.listN : 0;
        var groupN = opts ? opts.groupN : null;
        var linked = !!(opts && opts.linked);
        var memberWord = (opts && opts.memberWord) || 'Mitglieder';
        var known = linked && typeof groupN === 'number' && groupN >= 0;
        wrap.classList.remove('is-match', 'is-mismatch', 'slg-side-count--quiet');
        wrap.replaceChildren();

        if (known && listN === groupN) {
            wrap.classList.add('is-match', 'slg-side-count--quiet');
            var kpi = document.createElement('span');
            kpi.className = 'slg-side-kpi';
            kpi.innerHTML =
                '<span class="n">' +
                String(listN) +
                '</span><span class="l">' +
                escapeHtml(memberWord) +
                '</span>';
            var ok = document.createElement('span');
            ok.className = 'slg-side-sync-ok';
            ok.innerHTML = '<i class="bi bi-check-circle-fill" aria-hidden="true"></i>Synchron';
            wrap.appendChild(kpi);
            wrap.appendChild(ok);
            wrap.title = listN + ' ' + memberWord + ' · synchron';
            wrap.setAttribute('aria-label', listN + ' ' + memberWord + ', synchron');
            return { known: true, same: true, delta: 0 };
        }

        wrap.classList.add('slg-side-count--dual');
        var left = document.createElement('span');
        left.className = 'slg-side-kpi';
        left.innerHTML = '<span class="n">' + String(listN) + '</span><span class="l">erwartet</span>';
        var sep = document.createElement('span');
        sep.className = 'slg-side-kpi-sep';
        sep.setAttribute('aria-hidden', 'true');
        sep.textContent = '·';
        var right = document.createElement('span');
        right.className = 'slg-side-kpi';
        var rightN = known ? String(groupN) : linked ? '…' : '–';
        right.innerHTML = '<span class="n">' + rightN + '</span><span class="l">in M365</span>';
        wrap.appendChild(left);
        wrap.appendChild(sep);
        wrap.appendChild(right);

        if (known) {
            var delta = Math.abs(listN - groupN);
            wrap.classList.add('is-mismatch');
            var deltaEl = document.createElement('span');
            deltaEl.className = 'slg-side-delta';
            deltaEl.innerHTML =
                '<i class="bi bi-exclamation-triangle-fill" aria-hidden="true"></i>' +
                String(delta) +
                (delta === 1 ? ' Abweichung' : ' Abweichungen');
            wrap.appendChild(deltaEl);
            wrap.title = listN + ' erwartet · ' + groupN + ' in Microsoft 365';
            wrap.setAttribute(
                'aria-label',
                listN +
                    ' erwartet, ' +
                    groupN +
                    ' in Microsoft 365, ' +
                    delta +
                    (delta === 1 ? ' Abweichung' : ' Abweichungen')
            );
            return { known: true, same: false, delta: delta };
        }

        wrap.title = linked
            ? listN + ' erwartet · Mitgliederzahl wird aus Microsoft 365 geladen.'
            : listN + ' erwartet · noch nicht mit Microsoft 365 verknüpft.';
        wrap.setAttribute(
            'aria-label',
            listN + ' erwartet, Microsoft 365 ' + (linked ? 'wird geladen' : 'unbekannt')
        );
        return { known: false, same: false, delta: 0 };
    }

    function fillSyncPreviewLines(joinCount, leaveCount, stayCount) {
        var host = document.getElementById('slgSyncPreviewLines');
        if (!host) return;
        host.replaceChildren();
        function add(text) {
            var li = document.createElement('li');
            li.textContent = text;
            host.appendChild(li);
        }
        var joinN = typeof joinCount === 'number' ? joinCount : 0;
        var leaveN = typeof leaveCount === 'number' ? leaveCount : 0;
        var stayN = typeof stayCount === 'number' ? stayCount : 0;
        add(joinN === 1 ? '1 Person wird hinzugefügt' : String(joinN) + ' Personen werden hinzugefügt');
        add(leaveN === 1 ? '1 Person wird entfernt' : String(leaveN) + ' Personen werden entfernt');
        add(stayN === 1 ? '1 bleibt unverändert' : String(stayN) + ' bleiben unverändert');
    }

    function setSyncPreviewVisible(show) {
        var el = document.getElementById('slgSyncPreview');
        if (el) el.hidden = !show;
    }

    function setSyncLogVisible(show) {
        var el = document.getElementById('slgSyncLogWrap');
        if (el) el.hidden = !show;
    }

    /* ───── wireUi ───── */

    function wireUi(host) {
        if (!host || host.getAttribute(ATTR_WIRED) === '1') return;
        host.setAttribute(ATTR_WIRED, '1');

        document.querySelectorAll('#slgDetailTabs .detail-tab-btn[data-slg-tab]').forEach(function (b) {
            b.addEventListener('click', function () {
                setTab(b.getAttribute('data-slg-tab') || 'general');
            });
        });

        var tablist = document.getElementById('slgDetailTabs');
        if (tablist) {
            tablist.addEventListener('keydown', function (ev) {
                var tabs = Array.from(tablist.querySelectorAll('.detail-tab-btn[data-slg-tab]'));
                var idx = tabs.indexOf(document.activeElement);
                if (idx < 0) return;
                var next = -1;
                if (ev.key === 'ArrowRight') next = (idx + 1) % tabs.length;
                else if (ev.key === 'ArrowLeft') next = (idx - 1 + tabs.length) % tabs.length;
                else if (ev.key === 'Home') next = 0;
                else if (ev.key === 'End') next = tabs.length - 1;
                if (next >= 0 && next !== idx) {
                    ev.preventDefault();
                    tabs[next].focus();
                    setTab(tabs[next].getAttribute('data-slg-tab') || 'general');
                }
            });
        }

        onClick('slgBtnSearch', function () {
            runSearchGroups();
        });
        onClick('slgBtnCreate', function () {
            runCreateAndMatch();
        });
        onClick('slgBtnMailNickStandard', function () {
            applySuggestedMailNickname();
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
        syncMembershipTarget: syncMembershipTargetLabels,
        clearSearchResults: clearSearchResults,
        paintMasterMemberSyncCounts: paintMasterMemberSyncCounts,
        fillSyncPreviewLines: fillSyncPreviewLines,
        setSyncPreviewVisible: setSyncPreviewVisible,
        setSyncLogVisible: setSyncLogVisible
    };
})();
