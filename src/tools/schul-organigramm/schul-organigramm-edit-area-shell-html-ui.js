/**
 * Area editor HTML dialog construction.
 */
import {
    sammelgruppeKeyForArea,
    isKlassenchatClassNode,
    isArgeAreaNode,
    isClassAreaNode,
    canDeleteSchulOrganigrammArea,
    isAreaNode,
    normalizeSchulOrganigramm,
    normalizeSchulOrgGroupKind,
    suggestSchulOrgMailNickname,
    normalizeSchulOrgM365,
    normalizeSchulOrgM365Binding,
    emptySchulOrgM365Bindings,
    listSchulOrgClassFunctions,
    listSchulOrgArgeFunctions,
    listArgeCatalogGroupRefs,
    audienceIdForArea,
    jahrgangKeyForArea,
    functionAggregateDefForArea,
    isFunctionAggregateArea,
    isSgaBoardArea
} from './schul-organigramm-logic.js';
import { findClass, findArge } from './schul-organigramm-edit-find.js';
import { resolveClassGraphGroupId } from '../../shared/membership-hygiene.js';
import { getAudienceGroupGraphId } from '../../shared/tenant-admin-audience-m365.js';

/**
 * @param {Record<string, any>} ctx
 * @param {string|object} areaIdOrNode
 * @returns {{ overlay: HTMLElement, area: object, settings: object, sammelKey: string, isKlassenchatArea: boolean }|null}
 */
export function buildAreaEditorOverlay(ctx, areaIdOrNode) {
    const loadSettings = ctx.loadSettings;
    const toast = ctx.toast;
    const closeOverlay = ctx.closeOverlay;
    const escapeHtml = ctx.escapeHtml;
    const normStr = ctx.normStr;
    const editDialogHtml = ctx.editDialogHtml;
    const editFootHtml = ctx.editFootHtml;
    const areaM365ObjectsHtml = ctx.areaM365ObjectsHtml;

        closeOverlay();
        const settings = loadSettings();
        const layout = normalizeSchulOrganigramm(settings.schulOrganigramm);
        const areaId =
            typeof areaIdOrNode === 'string'
                ? normStr(areaIdOrNode)
                : normStr(areaIdOrNode && areaIdOrNode.id);
        let area = layout.nodes.find(function (n) {
            return n.id === areaId && isAreaNode(n);
        });
        if (!area) {
            toast('Bereich nicht gefunden.', 'warning');
        return null;
        }
        // Explizites groupMode=none / gelöste Verknüpfung nicht wieder überschreiben
        const hadStoredM365 = !!(area.m365 && typeof area.m365 === 'object');
        const isKlassenchatArea = isKlassenchatClassNode(area);
        const defaultGroupKind = isKlassenchatArea ? 'groupChat' : 'unified';
        area = Object.assign({}, area, {
            m365: area.m365
                ? Object.assign(
                      {
                          groupMode: 'none',
                          groupKind: defaultGroupKind,
                          groupId: null,
                          mailboxMode: 'none',
                          mailboxUpn: null
                      },
                      area.m365
                  )
                : {
                      groupMode: 'none',
                      groupKind: defaultGroupKind,
                      groupId: null,
                      mailboxMode: 'none',
                      mailboxUpn: null
                  }
        });
        area.m365.groupKind = normalizeSchulOrgGroupKind(
            area.m365.groupKind || defaultGroupKind
        );
        area.m365 = normalizeSchulOrgM365(area.m365) || area.m365;
        if (isKlassenchatArea) {
            const chatIdMeta = normStr(area.meta && area.meta.chatId);
            if (!area.m365.bindings) area.m365.bindings = emptySchulOrgM365Bindings();
            if (chatIdMeta || area.m365.bindings.groupChat.mode === 'none') {
                area.m365 = normalizeSchulOrgM365({
                    bindings: Object.assign({}, area.m365.bindings, {
                        groupChat: normalizeSchulOrgM365Binding(
                            {
                                mode: chatIdMeta ? 'match' : 'none',
                                id: chatIdMeta || null,
                                label: normStr(area.meta && area.meta.chatTopic) || null,
                                webUrl: normStr(area.meta && area.meta.chatWebUrl) || null
                            },
                            'groupChat'
                        )
                    })
                });
            }
        }
        const sammelKey = sammelgruppeKeyForArea(area);
        // Nur Erstvorschlag, wenn noch nie M365 am Bereich gespeichert wurde
        if (!hadStoredM365) {
            try {
                const container =
                    typeof window !== 'undefined' &&
                    window.ms365AppDataV2 &&
                    typeof window.ms365AppDataV2.getContainer === 'function'
                        ? window.ms365AppDataV2.getContainer()
                        : null;
                const setup = container && container.setup ? container.setup : {};
                const matched = setup.matched && typeof setup.matched === 'object' ? setup.matched : {};
                if (sammelKey === 'lehrer' || sammelKey === 'schueler') {
                    const gid =
                        sammelKey === 'lehrer'
                            ? String(matched.lehrerGroupId || '').trim()
                            : String(matched.schuelerGroupId || '').trim();
                    if (gid) {
                        area.m365.groupMode = 'match';
                        area.m365.groupKind = 'unified';
                        area.m365.groupId = gid;
                        area.m365.groupLabel =
                            sammelKey === 'lehrer' ? 'Lehrer-Sammelgruppe' : 'Schüler-Sammelgruppe';
                    }
                } else if (isClassAreaNode(area) && area.refKey) {
                    // Geerbte Klassengruppe anzeigen, damit „Gruppe lösen“ sie wirklich abschalten kann
                    const cls = findClass(settings, area.refKey);
                    const classTeams =
                        container && container.core && Array.isArray(container.core.classTeams)
                            ? container.core.classTeams
                            : [];
                    const matchMap =
                        setup.classGroupMatchByKey && typeof setup.classGroupMatchByKey === 'object'
                            ? setup.classGroupMatchByKey
                            : {};
                    const gid = cls ? resolveClassGraphGroupId(cls, classTeams, matchMap) : '';
                    if (gid) {
                        area.m365.groupMode = 'match';
                        area.m365.groupKind = 'unified';
                        area.m365.groupId = gid;
                        area.m365.groupLabel = 'Klassengruppe (Stammdaten)';
                    }
                } else if (isArgeAreaNode(area) && area.refKey) {
                    const refs = listArgeCatalogGroupRefs(setup);
                    const code = normStr(area.refKey).toUpperCase();
                    const hit = refs.find(function (r) {
                        return r.code === code;
                    });
                    if (hit && hit.groupId) {
                        area.m365.groupMode = 'match';
                        area.m365.groupKind = 'unified';
                        area.m365.groupId = hit.groupId;
                        area.m365.groupLabel = hit.label || 'ARGE-Gruppe';
                    }
                } else {
                    const aidHeal = audienceIdForArea(area);
                    if (aidHeal) {
                        const gid = getAudienceGroupGraphId(aidHeal);
                        if (gid) {
                            area.m365.groupMode = 'match';
                            area.m365.groupKind = 'unified';
                            area.m365.groupId = gid;
                            area.m365.groupLabel =
                                normStr(area.label) || 'Verwaltungsgruppe';
                        }
                    }
                }
            } catch {
                /* ignore */
            }
        }

        const overlay = document.createElement('div');
        overlay.id = 'schulOrgEditOverlay';
        overlay.className = 'modal-overlay open schul-org-edit-overlay';
        const audId = audienceIdForArea(area);
        const jgKey = jahrgangKeyForArea(area);
        const isArgeArea = isArgeAreaNode(area);
        const isClassArea = isClassAreaNode(area);
        const fnAggDef = functionAggregateDefForArea(area);
        const isFnAggBoard = isFunctionAggregateArea(area);
        const isSgaBoard = isSgaBoardArea(area);
        const classFnRoleOptions = isClassArea
            ? listSchulOrgClassFunctions()
                  .map(function (d) {
                      return (
                          '<option value="fn:' +
                          escapeHtml(d.id) +
                          '">' +
                          escapeHtml(d.label) +
                          '</option>'
                      );
                  })
                  .join('')
            : '';
        const argeFnRoleOptions = isArgeArea
            ? listSchulOrgArgeFunctions()
                  .map(function (d) {
                      return (
                          '<option value="fn:' +
                          escapeHtml(d.id) +
                          '">' +
                          escapeHtml(d.label) +
                          '</option>'
                      );
                  })
                  .join('')
            : '';
        // Lehrkräfte: Personen hier pflegen → teachers[] (Lokal-First)
        const peopleLocked = !!(
            (sammelKey && sammelKey !== 'lehrer') ||
            isSgaBoard ||
            isFnAggBoard
        );
        const peopleSub = sammelKey
            ? sammelKey === 'lehrer'
                ? 'Bereich = M365-Sammelgruppe · Personen hier pflegen die Lehrerliste in den Stammdaten'
                : sammelKey === 'klassenchats'
                  ? 'Bereich = Klassenlehrer-Chats · Jahrgänge und Klassen darunter · Bezeichnung frei wählbar'
                  : 'Bereich = M365-Sammelgruppe · Klassen in den Jahrgangs-Bereichen'
            : isSgaBoard
              ? 'SGA: Schülervertretung, Lehrer, Eltern · optional M365-Gruppe/Verteiler/Postfach'
              : isFnAggBoard
                ? 'Alle ' +
                  ((fnAggDef && fnAggDef.labelPlural) || 'Funktionsträger') +
                  ' aus den Gruppen · dort zuweisen · hier Microsoft 365 (Chat/Gruppe/Postfach)'
                : isArgeArea
                  ? 'Bereich = M365-Gruppe · Mitglieder der Gruppe · Besitzer = ARGE-Leitung (auch unter ARGE-Leiter)'
                  : isKlassenchatArea
                    ? 'Bereich = Teams-Gruppenchat (KlassenChat) · Lehrkräfte aus Belegung/KV · Anlage im Werkzeug KlassenChats'
                    : isClassArea
                      ? 'Bereich = M365-Klassengruppe · Schüler:innen, Klassenvorstand & Klassenfunktionen'
                      : jgKey
                        ? 'Klassen dieses Jahrgangs · optional Gruppe/Verteiler/Postfach zuordnen oder anlegen'
                        : audId
                          ? 'Personen · M365-Objekt matchen oder neu anlegen'
                          : 'Jeder Bereich entspricht einem M365-Objekt (Gruppe, Sicherheitsgruppe, Verteiler, Gruppenchat oder Postfach)';
        const stammdatenBtn =
            isArgeArea && area.refKey
                ? '<button type="button" class="btn btn-sm" data-area-open-stammdaten title="Kürzel, Fächer, Leitung manuell">' +
                  '<i class="bi bi-journal-bookmark" aria-hidden="true"></i> ARGE-Stammdaten</button>'
                : isClassArea && area.refKey
                  ? '<button type="button" class="btn btn-sm" data-area-open-stammdaten title="Kürzel, Abschlussjahr, KV">' +
                    '<i class="bi bi-mortarboard" aria-hidden="true"></i> Klassen-Stammdaten</button>'
                  : isSgaBoard
                    ? '<button type="button" class="btn btn-sm" data-area-open-stammdaten title="Schüler, Lehrer, Eltern im SGA">' +
                      '<i class="bi bi-people" aria-hidden="true"></i> SGA-Personen</button>'
                    : '';
        overlay.innerHTML = editDialogHtml({
            wide: true,
            icon: isArgeArea
                ? 'bi-journal-bookmark'
                : isClassArea
                  ? 'bi-mortarboard'
                  : isSgaBoard
                    ? 'bi-building'
                    : 'bi-pencil-square',
            title: isSgaBoard ? 'SGA bearbeiten' : 'Bereich bearbeiten',
            sub: peopleSub,
            body:
                '<div class="field"><label for="schulOrgAreaLabel">Bezeichnung</label>' +
                '<input type="text" id="schulOrgAreaLabel" value="' +
                escapeHtml(area.label) +
                '"' +
                (isSgaBoard ? ' readonly' : '') +
                '></div>' +
                (stammdatenBtn
                    ? '<div class="schul-org-edit-people-actions" style="margin-bottom:0.5rem">' +
                      stammdatenBtn +
                      '</div>'
                    : '') +
                '<div class="schul-org-area-tabs" data-area-tabs>' +
                '<div class="schul-org-area-tabs__list" role="tablist" aria-label="Bereich bearbeiten">' +
                '<button type="button" class="schul-org-area-tabs__btn is-active" role="tab" id="schulOrgAreaTabPeople" aria-selected="true" aria-controls="schulOrgAreaPanelPeople" data-area-tab="people">' +
                '<i class="bi bi-people" aria-hidden="true"></i> Personen</button>' +
                '<button type="button" class="schul-org-area-tabs__btn" role="tab" id="schulOrgAreaTabM365" aria-selected="false" aria-controls="schulOrgAreaPanelM365" data-area-tab="m365" tabindex="-1">' +
                '<i class="bi bi-microsoft" aria-hidden="true"></i> Microsoft 365-Objekte</button>' +
                '</div>' +
                '<div class="schul-org-area-tabs__panel is-active" role="tabpanel" id="schulOrgAreaPanelPeople" aria-labelledby="schulOrgAreaTabPeople" data-area-panel="people">' +
                '<div class="schul-org-area-people">' +
                '<div class="schul-org-area-people__head">' +
                '<strong>Personen in diesem Bereich</strong>' +
                (peopleLocked
                    ? '<span class="muted">' +
                      (isSgaBoard
                          ? 'in den Unterbereichen Schülervertretung / Lehrer / Eltern'
                          : 'aus Stammdaten / verknüpfter Gruppe') +
                      '</span>'
                    : '<div class="schul-org-edit-people-actions">' +
                      (isArgeArea || isClassArea
                          ? '<span class="muted">Stammdaten + Gruppe</span>' +
                            '<label class="schul-org-area-people__role">' +
                            '<span>Funktion</span>' +
                            '<select data-area-person-role aria-label="Funktion / Rolle">' +
                            (isClassArea
                                ? '<option value="student">Schüler:in</option>' +
                                  '<option value="kv">Klassenvorstand</option>' +
                                  classFnRoleOptions
                                : '<option value="argeMember">Mitglied</option>' +
                                  '<option value="argeHead">ARGE-Leitung</option>' +
                                  argeFnRoleOptions) +
                            '</select></label>'
                          : '') +
                      '<button type="button" class="btn btn-sm btn-success" data-area-add-person-manual>' +
                      '<i class="bi bi-person-plus" aria-hidden="true"></i> Person</button>' +
                      '<button type="button" class="btn btn-sm btn-primary" data-area-add-person>' +
                      '<i class="bi bi-microsoft" aria-hidden="true"></i> Aus Microsoft 365</button>' +
                      '</div>') +
                '</div>' +
                (peopleLocked
                    ? ''
                    : '<div class="schul-org-area-people__manual" data-area-manual-form hidden>' +
                      '<div class="schul-org-edit-grid">' +
                      (isArgeArea || isClassArea
                          ? '<div class="field"><label for="schulOrgAreaPersonRole">Funktion</label>' +
                            '<select id="schulOrgAreaPersonRole" data-area-person-role-manual>' +
                            (isClassArea
                                ? '<option value="student">Schüler:in</option>' +
                                  '<option value="kv">Klassenvorstand</option>' +
                                  classFnRoleOptions
                                : '<option value="argeMember">Mitglied</option>' +
                                  '<option value="argeHead">ARGE-Leitung</option>' +
                                  argeFnRoleOptions) +
                            '</select></div>'
                          : '') +
                      '<div class="field"><label for="schulOrgAreaPersonName">Name</label>' +
                      '<input type="text" id="schulOrgAreaPersonName" maxlength="120" autocomplete="name" placeholder="Vor- und Nachname"></div>' +
                      '<div class="field"><label for="schulOrgAreaPersonEmail">E-Mail <span class="muted">(optional)</span></label>' +
                      '<input type="email" id="schulOrgAreaPersonEmail" maxlength="160" autocomplete="email" placeholder="name@schule.at"></div>' +
                      '</div>' +
                      '<p class="muted schul-org-edit-hint">' +
                      (isClassArea
                          ? 'Funktion wählen (Schüler:in, Klassenvorstand oder z. B. Klassensprecher), dann Person zuweisen. Ohne E-Mail wird eine lokale Kennung erzeugt.'
                          : isArgeArea
                            ? 'Funktion wählen (ARGE-Leitung oder Mitglied), dann Person zuweisen. Ohne E-Mail wird eine lokale Kennung erzeugt.'
                            : sammelKey === 'lehrer'
                              ? 'Wird in die Lehrerliste (Stammdaten) geschrieben; Kürzel wird automatisch vergeben. Ohne E-Mail wird eine lokale Kennung erzeugt.'
                              : 'Ohne E-Mail wird eine lokale Kennung erzeugt.') +
                      '</p>' +
                      '<button type="button" class="btn btn-sm btn-success" data-area-manual-submit>' +
                      '<i class="bi bi-check-lg" aria-hidden="true"></i> Hinzufügen</button>' +
                      '</div>') +
                '<ul class="schul-org-area-people__list" data-area-people-list></ul>' +
                (peopleLocked
                    ? ''
                    : '<div class="schul-org-inline-eup-host" data-schul-org-inline-eup hidden></div>') +
                '</div></div>' +
                '<div class="schul-org-area-tabs__panel" role="tabpanel" id="schulOrgAreaPanelM365" aria-labelledby="schulOrgAreaTabM365" data-area-panel="m365" hidden>' +
                '<fieldset class="schul-org-area-m365">' +
                '<legend class="visually-hidden">Microsoft 365-Objekte</legend>' +
                '<p class="muted schul-org-area-m365__hint">' +
                'Mehrere Objekttypen parallel möglich (z. B. Teams-Gruppenchat + Sicherheitsgruppe + Postfach). ' +
                'Links anhaken, was genutzt werden soll – rechts: bestehendes zuordnen oder neu anlegen.' +
                '</p>' +
                '<div class="schul-org-m365-objects schul-org-m365-objects--stack" data-m365-objects>' +
                areaM365ObjectsHtml(area, { isKlassenchatArea: isKlassenchatArea }) +
                '</div>' +
                // Versteckte Legacy-Felder für bestehende Create/Match-Handler
                '<div class="schul-org-area-m365__legacy" hidden aria-hidden="true">' +
                '<select id="schulOrgAreaGroupKind">' +
                '<option value="unified">Microsoft 365-Gruppe</option>' +
                '<option value="security">Sicherheitsgruppe</option>' +
                '<option value="mailSecurity">E-Mail-Sicherheitsgruppe</option>' +
                '<option value="distribution">Verteilerliste</option>' +
                '<option value="groupChat">Teams-Gruppenchat</option>' +
                '</select>' +
                '<select id="schulOrgAreaGroupMode">' +
                '<option value="none">Keine Gruppe</option>' +
                '<option value="match">Bestehende zuordnen</option>' +
                '<option value="create">Neu in Microsoft 365 anlegen</option>' +
                '</select>' +
                '<select id="schulOrgAreaMailboxMode">' +
                '<option value="none">Keins</option>' +
                '<option value="match">Bestehendes zuordnen</option>' +
                '<option value="create">Neu in Exchange anlegen</option>' +
                '</select>' +
                '<button type="button" data-area-pick-group></button>' +
                '<button type="button" data-area-clear-group></button>' +
                '<button type="button" data-area-pick-mail></button>' +
                '<button type="button" data-area-clear-mail></button>' +
                '<p data-area-group-sum></p><p data-area-chat-sum></p><p data-area-mail-sum></p>' +
                '<span data-area-kind-icon></span>' +
                '<div data-area-group-match-row></div><div data-area-mail-match-row></div>' +
                '</div>' +
                '<div class="schul-org-inline-eup-host" data-schul-org-inline-egp hidden></div>' +
                '<div class="schul-org-area-m365__create" data-area-group-create-row hidden>' +
                '<p class="muted"><strong data-area-create-kind-label>Gruppe</strong> neu anlegen</p>' +
                '<p class="muted schul-org-area-m365__group-people" data-area-group-people></p>' +
                '<div class="field"><label for="schulOrgAreaCreateName">Anzeigename</label>' +
                '<input type="text" id="schulOrgAreaCreateName" value="' +
                escapeHtml(area.label) +
                '"></div>' +
                '<div class="field"><label for="schulOrgAreaCreateNick">Mail-Nickname / Alias</label>' +
                '<input type="text" id="schulOrgAreaCreateNick" value="' +
                escapeHtml(suggestSchulOrgMailNickname(area.label)) +
                '" autocomplete="off"></div>' +
                '<div class="field"><label for="schulOrgAreaGroupOwner">Besitzer</label>' +
                '<div class="schul-org-edit-people-actions schul-org-area-owner-row">' +
                '<select id="schulOrgAreaGroupOwner" data-area-group-owner>' +
                '<option value="">— aus Bereich wählen —</option>' +
                '</select>' +
                '<button type="button" class="btn btn-sm" data-area-pick-owner title="Beliebige Person aus Microsoft 365">' +
                '<i class="bi bi-microsoft" aria-hidden="true"></i> Aus Microsoft 365</button>' +
                '</div>' +
                '<div class="schul-org-inline-eup-host" data-schul-org-inline-eup-owner hidden></div>' +
                '<p class="muted schul-org-edit-hint">Personen aus diesem Bereich, Rollen und Unterbereichen werden flach als Mitglieder hinzugefügt. Besitzer muss nicht Mitglied sein.</p></div>' +
                '<button type="button" class="btn btn-sm btn-primary" data-area-create-group>' +
                '<i class="bi bi-cloud-plus" aria-hidden="true"></i> Jetzt in Microsoft 365 anlegen</button>' +
                '<p class="muted schul-org-area-m365__create-hint" data-area-create-hint></p>' +
                '</div>' +
                '<div class="schul-org-inline-eup-host" data-schul-org-inline-eup-mail hidden></div>' +
                '<div class="schul-org-area-m365__create" data-area-mail-create-row hidden>' +
                '<p class="muted schul-org-area-m365__mail-hint">Freigegebene Postfächer legt Exchange Online an (nicht Graph). Es entsteht eine doppelklickbare <strong>.cmd</strong>-Datei mit geführtem Assistenten. Personen aus Bereich, Rollen und Unterbereichen erhalten FullAccess und SendAs.</p>' +
                '<p class="muted schul-org-area-m365__mail-delegates" data-area-mail-delegates></p>' +
                '<div class="field"><label for="schulOrgAreaMailCreateName">Anzeigename</label>' +
                '<input type="text" id="schulOrgAreaMailCreateName" value="' +
                escapeHtml(area.label) +
                '" autocomplete="off"></div>' +
                '<div class="field"><label for="schulOrgAreaMailCreateAlias">Alias</label>' +
                '<input type="text" id="schulOrgAreaMailCreateAlias" value="' +
                escapeHtml(suggestSchulOrgMailNickname(area.label)) +
                '" autocomplete="off"></div>' +
                '<div class="field"><label for="schulOrgAreaMailCreateSmtp">Primäre E-Mail</label>' +
                '<input type="email" id="schulOrgAreaMailCreateSmtp" value="" autocomplete="off" placeholder="alias@schule.at"></div>' +
                '<div class="schul-org-edit-people-actions">' +
                '<button type="button" class="btn btn-sm btn-primary" data-area-create-mail>' +
                '<i class="bi bi-terminal" aria-hidden="true"></i> Assistent erzeugen &amp; vormerken</button>' +
                '<button type="button" class="btn btn-sm" data-area-mail-script-copy hidden>' +
                '<i class="bi bi-clipboard" aria-hidden="true"></i> PowerShell kopieren</button>' +
                '<button type="button" class="btn btn-sm btn-success" data-area-mail-script-dl hidden>' +
                '<i class="bi bi-download" aria-hidden="true"></i> .cmd herunterladen</button>' +
                '</div>' +
                '<textarea class="schul-org-area-m365__script" data-area-mail-script hidden readonly rows="10" aria-label="Exchange-PowerShell-Assistent"></textarea>' +
                '</div>' +
                '</fieldset></div></div>',
            foot: canDeleteSchulOrganigrammArea(area)
                ? editFootHtml('data-schul-org-delete-area', 'Bereich löschen', 'data-schul-org-save')
                : '<span class="schul-org-edit-foot-spacer"></span>' +
                  '<button type="button" class="btn alt" data-schul-org-close>Abbrechen</button>' +
                  '<button type="button" class="btn btn-success" data-schul-org-save>Speichern</button>'
        });

    return {
        overlay: overlay,
        area: area,
        settings: settings,
        sammelKey: sammelKey,
        isKlassenchatArea: isKlassenchatArea
    };
}
