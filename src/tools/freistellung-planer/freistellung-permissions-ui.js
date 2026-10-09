/**
 * Entra-Gruppen-Picker für Freistellungen-Setup.
 */
import {
    wireEntraGroupPickerFields,
    readGroupPickerField,
    fillGroupPickerField
} from '../../shared/entra-group-picker.js';
import {
    normalizePermissionsConfig,
    normalizePlannerExtraEntraGroups,
    savePermissionsConfig,
    loadPermissionsConfig,
    loadEffectivePermissionsConfig,
    PERMS_STORAGE_KEY,
    plannerGrantRowsForUi
} from './freistellung-planer-permissions.js';
import { publishPlannerPermissionsToSite } from './freistellung-planer-remote-config.js';
import { loadExtraKategorien } from './freistellung-planer-kategorien.js';
import {
    patchFreistellungKategorieColumn,
    patchFreistellungKlasseColumn
} from './freistellung-planer-graph.js';
import { classCatalogFromSchoolStammdaten } from './freistellung-planer-remote-config.js';
import { wireFreistellungKategorienAdmin } from './freistellung-kategorien-ui.js';
import { wireFreistellungKlassenAdmin } from './freistellung-klassen-ui.js';
import { pickEntraUser } from '../../shared/entra-user-picker.js';
import {
    normalizePlannerUsers,
    mergePlannerUser,
    direktionUsersFromTenantStammdaten
} from './freistellung-planer-direktion-users.js';
import { kvUsersFromSchoolStammdaten } from './freistellung-planer-list-permissions.js';
import {
    FR_STAMMDATEN_GROUP_ROLES,
    stripStammdatenGroupFieldsFromPatch,
    filterEditableGroupFields
} from '../../shared/planner-stammdaten-audience-ui.js';
import { loadSchoolAudienceGroups } from '../../shared/school-audience-groups.js';
import { normalizeAllowedJahrgang, normalizeJahrgangGroups } from './freistellung-planer-jahrgang-scope.js';
import {
    migrateLegacyFreistellungPerms,
    mergeDuplicateGrantRows
} from './freistellung-planner-grant-matrix.js';
import {
    initFreistellungPlannerGrantMatrixUi,
    readGrantRowsFromDom,
    appendGrantRow
} from './freistellung-planner-grant-matrix-ui.js';

export const SETUP_GROUP_FIELDS = [
    {
        role: 'groupDirektion',
        labelInputId: 'frPermDirektion',
        idInputId: 'frPermDirektionId',
        pickBtnId: 'frPermDirektionPick',
        clearBtnId: 'frPermDirektionClear',
        dialogTitle: 'Direktion / Verwaltung'
    },
    {
        role: 'groupKv',
        labelInputId: 'frPermKv',
        idInputId: 'frPermKvId',
        pickBtnId: 'frPermKvPick',
        clearBtnId: 'frPermKvClear',
        dialogTitle: 'Klassenvorstände'
    },
    {
        role: 'groupSchueler',
        labelInputId: 'frPermSchueler',
        idInputId: 'frPermSchuelerId',
        pickBtnId: 'frPermSchuelerPick',
        clearBtnId: 'frPermSchuelerClear',
        dialogTitle: 'Schüler'
    }
];

/** Direktion, KV und Schüler – alle im Setup wählbar. */
export const FR_EDITABLE_GROUP_FIELDS = filterEditableGroupFields(SETUP_GROUP_FIELDS, FR_STAMMDATEN_GROUP_ROLES);

/** @type {Record<string, Array<{ id: string, displayName: string, mail: string }>>} */
const extraUsersDraft = {
    direktionUsers: [],
    kvUsers: [],
    schuelerUsers: []
};

/** @type {Array<{ jahrgang: string, groupId: string, groupLabel: string }>} */
let jahrgangGroupsDraft = [];

const EXTRA_ENTRA_GROUP_SPECS = [
    {
        configKey: 'direktionGroups',
        listId: 'frPermDirektionGroups',
        addBtnId: 'frPermDirektionGroupAdd',
        primaryIdEl: 'frPermDirektionId',
        pickTitle: 'Weitere Entra-Gruppe für die Direktion-Ansicht'
    },
    {
        configKey: 'kvGroups',
        listId: 'frPermKvGroups',
        addBtnId: 'frPermKvGroupAdd',
        primaryIdEl: 'frPermKvId',
        pickTitle: 'Weitere Entra-Gruppe für die KV-Ansicht'
    },
    {
        configKey: 'schuelerGroups',
        listId: 'frPermSchuelerGroups',
        addBtnId: 'frPermSchuelerGroupAdd',
        primaryIdEl: 'frPermSchuelerId',
        pickTitle: 'Weitere Entra-Gruppe für die Schüler-Ansicht'
    }
];

/** @type {Record<'direktionGroups'|'kvGroups'|'schuelerGroups', Array<{ groupId: string, groupLabel: string }>>} */
const extraEntraGroupsDraft = {
    direktionGroups: [],
    kvGroups: [],
    schuelerGroups: []
};

function readAllowedJahrgangInput() {
    const el = document.getElementById('frAllowedJahrgang');
    return normalizeAllowedJahrgang(el ? el.value : '');
}

function renderJahrgangGroupsList() {
    const ul = document.getElementById('frJahrgangGroupsList');
    if (!ul) return;
    jahrgangGroupsDraft = normalizeJahrgangGroups(jahrgangGroupsDraft);
    if (!jahrgangGroupsDraft.length) {
        ul.innerHTML =
            '<li class="fr-setup-user-row fr-setup-user-row--empty muted">Noch keine Jahrgangs-Gruppen.</li>';
        return;
    }
    ul.innerHTML = jahrgangGroupsDraft
        .map(
            (g, idx) =>
                `<li class="fr-setup-user-row">` +
                `<span class="fr-setup-user-row__label">Jg ${escapeHtml(g.jahrgang)}: ${escapeHtml(
                    g.groupLabel || g.groupId
                )}</span>` +
                `<button type="button" class="fr-setup-user-row__rm btn btn-sm alt" data-fr-jg-rm="${idx}" title="Entfernen" aria-label="Entfernen"><i class="bi bi-x"></i></button>` +
                `</li>`
        )
        .join('');
    ul.querySelectorAll('[data-fr-jg-rm]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const i = Number(btn.getAttribute('data-fr-jg-rm'));
            const list = normalizeJahrgangGroups(jahrgangGroupsDraft);
            list.splice(i, 1);
            jahrgangGroupsDraft = list;
            renderJahrgangGroupsList();
            persistPickersToStorage(SETUP_GROUP_FIELDS);
        });
    });
}

function renderExtraEntraGroupsList(spec) {
    const ul = document.getElementById(spec.listId);
    if (!ul) return;
    const key = spec.configKey;
    extraEntraGroupsDraft[key] = normalizePlannerExtraEntraGroups(extraEntraGroupsDraft[key]);
    const list = extraEntraGroupsDraft[key];
    if (!list.length) {
        ul.innerHTML =
            '<li class="fr-setup-user-row fr-setup-user-row--empty muted">Noch keine weiteren Gruppen.</li>';
        return;
    }
    ul.innerHTML = list
        .map(
            (g, idx) =>
                `<li class="fr-setup-user-row">` +
                `<span class="fr-setup-user-row__label">${escapeHtml(g.groupLabel || g.groupId)}</span>` +
                `<button type="button" class="fr-setup-user-row__rm btn btn-sm alt" data-fr-eg-rm="${key}:${idx}" title="Entfernen" aria-label="Entfernen"><i class="bi bi-x"></i></button>` +
                `</li>`
        )
        .join('');
    ul.querySelectorAll('[data-fr-eg-rm]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const raw = String(btn.getAttribute('data-fr-eg-rm') || '');
            const sep = raw.indexOf(':');
            const cfgKey = raw.slice(0, sep);
            const i = Number(raw.slice(sep + 1));
            if (!cfgKey || !extraEntraGroupsDraft[cfgKey]) return;
            const next = normalizePlannerExtraEntraGroups(extraEntraGroupsDraft[cfgKey]);
            next.splice(i, 1);
            extraEntraGroupsDraft[cfgKey] = next;
            renderExtraEntraGroupsList(spec);
            persistPickersToStorage(SETUP_GROUP_FIELDS);
        });
    });
}

/**
 * @param {() => void} [onChange]
 */
function wireExtraEntraGroupsUi(onChange) {
    EXTRA_ENTRA_GROUP_SPECS.forEach((spec) => {
        renderExtraEntraGroupsList(spec);
        const addBtn = document.getElementById(spec.addBtnId);
        if (addBtn && !addBtn.dataset.frEgWired) {
            addBtn.dataset.frEgWired = '1';
            addBtn.addEventListener('click', () => {
                pickEntraGroup({ title: spec.pickTitle })
                    .then((sel) => {
                        if (!sel || !sel.id) return;
                        const primaryId = String(
                            (document.getElementById(spec.primaryIdEl) || {}).value || ''
                        ).trim().toLowerCase();
                        if (primaryId && String(sel.id).toLowerCase() === primaryId) {
                            if (typeof window.ms365ToastOrAlert === 'function') {
                                window.ms365ToastOrAlert('Diese Gruppe ist bereits als Hauptgruppe gewählt.');
                            }
                            return;
                        }
                        const key = spec.configKey;
                        extraEntraGroupsDraft[key] = normalizePlannerExtraEntraGroups([
                            ...extraEntraGroupsDraft[key],
                            {
                                groupId: sel.id,
                                groupLabel: sel.label || sel.displayName || sel.id
                            }
                        ]);
                        renderExtraEntraGroupsList(spec);
                        if (typeof onChange === 'function') onChange();
                    })
                    .catch((e) => {
                        const msg = e && e.message ? e.message : String(e);
                        if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
                    });
            });
        }
    });
    const stammBtn = document.getElementById('frPermSchuelerStammdaten');
    if (stammBtn && !stammBtn.dataset.frSgWired) {
        stammBtn.dataset.frSgWired = '1';
        stammBtn.addEventListener('click', () => {
            const aud = loadSchoolAudienceGroups();
            const id = String(aud.groupSchuelerId || '').trim();
            if (!id) {
                if (typeof window.ms365ToastOrAlert === 'function') {
                    window.ms365ToastOrAlert(
                        'Keine Schüler-Sammelgruppe in den Stammdaten – bitte unter Stammdaten / MS 365-Gruppenverwaltung verknüpfen.'
                    );
                }
                return;
            }
            fillGroupPickerField(SETUP_GROUP_FIELDS.find((f) => f.role === 'groupSchueler'), {
                id,
                label: aud.groupSchuelerName || id
            });
            if (typeof onChange === 'function') onChange();
            if (typeof window.ms365ToastOrAlert === 'function') {
                window.ms365ToastOrAlert('Schüler-Sammelgruppe aus Stammdaten übernommen.');
            }
        });
    }
}

/**
 * @param {() => void} [onChange]
 */
function wireJahrgangScopeUi(onChange) {
    const allowedEl = document.getElementById('frAllowedJahrgang');
    if (allowedEl && !allowedEl.dataset.frJgWired) {
        allowedEl.dataset.frJgWired = '1';
        const notify = () => {
            if (typeof onChange === 'function') onChange();
        };
        allowedEl.addEventListener('change', notify);
        allowedEl.addEventListener('input', notify);
    }
    const addBtn = document.getElementById('frJahrgangGroupAdd');
    if (addBtn && !addBtn.dataset.frJgWired) {
        addBtn.dataset.frJgWired = '1';
        addBtn.addEventListener('click', () => {
            const jgRaw = window.prompt('Schulstufe / Jahrgang (z. B. 5 oder 10):', '');
            const jahrgang = String(jgRaw || '').trim();
            if (!jahrgang) return;
            pickEntraGroup({ title: 'Jahrgangs-Entra-Gruppe' })
                .then((sel) => {
                    if (!sel || !sel.id) return;
                    jahrgangGroupsDraft = normalizeJahrgangGroups([
                        ...jahrgangGroupsDraft,
                        {
                            jahrgang,
                            groupId: sel.id,
                            groupLabel: sel.label || sel.displayName || sel.id
                        }
                    ]);
                    renderJahrgangGroupsList();
                    if (typeof onChange === 'function') onChange();
                })
                .catch((e) => {
                    const msg = e && e.message ? e.message : String(e);
                    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
                });
        });
    }
}

const EXTRA_USER_UI = [
    {
        configKey: 'direktionUsers',
        listId: 'frPermDirektionUsers',
        addBtnId: 'frPermDirektionUserAdd',
        stammdatenBtnId: 'frPermDirektionStammdaten',
        pickTitle: 'Person für Direktion / Verwaltung',
        pickHint: 'Sekretariat, Administration oder Schulleitung.',
        emptyHint: 'Noch keine Einzelpersonen – z. B. aus Stammdaten übernehmen.'
    },
    {
        configKey: 'kvUsers',
        listId: 'frPermKvUsers',
        addBtnId: 'frPermKvUserAdd',
        stammdatenBtnId: 'frPermKvStammdaten',
        pickTitle: 'Person als Klassenvorstand',
        pickHint: 'Zusätzlich zur KV-Entra-Gruppe – z. B. Vertretung oder einzelner KV ohne Gruppe.',
        emptyHint: 'Noch keine Einzelpersonen – werden aus Stammdaten (Klassen) vorgefüllt, sobald KV-E-Mails hinterlegt sind.'
    },
    {
        configKey: 'schuelerUsers',
        listId: 'frPermSchuelerUsers',
        addBtnId: 'frPermSchuelerUserAdd',
        pickTitle: 'Schüler/in einzeln',
        pickHint: 'Zusätzlich zur Schüler-Entra-Gruppe – z. B. Einzelfall ohne Gruppenmitgliedschaft.',
        emptyHint: 'Noch keine Einzelpersonen.'
    }
];

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function formatUserLine(u) {
    const name = String(u.displayName || '').trim();
    const mail = String(u.mail || '').trim();
    if (name && mail && name.toLowerCase() !== mail.toLowerCase()) {
        return escapeHtml(name) + ' <span class="fr-setup-user-row__mail">(' + escapeHtml(mail) + ')</span>';
    }
    return escapeHtml(name || mail);
}

function renderExtraUsersList(spec) {
    const ul = document.getElementById(spec.listId);
    if (!ul) return;
    const users = normalizePlannerUsers(extraUsersDraft[spec.configKey]);
    extraUsersDraft[spec.configKey] = users;
    if (!users.length) {
        ul.innerHTML =
            '<li class="fr-setup-user-row fr-setup-user-row--empty muted">' + escapeHtml(spec.emptyHint) + '</li>';
        return;
    }
    ul.innerHTML = users
        .map(
            (u, idx) =>
                `<li class="fr-setup-user-row">` +
                `<span class="fr-setup-user-row__label">${formatUserLine(u)}</span>` +
                `<button type="button" class="fr-setup-user-row__rm btn btn-sm alt" data-fr-user-rm="${idx}" title="Entfernen" aria-label="Entfernen"><i class="bi bi-x"></i></button>` +
                `</li>`
        )
        .join('');
    ul.querySelectorAll('[data-fr-user-rm]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const i = Number(btn.getAttribute('data-fr-user-rm'));
            const list = normalizePlannerUsers(extraUsersDraft[spec.configKey]);
            list.splice(i, 1);
            extraUsersDraft[spec.configKey] = list;
            renderExtraUsersList(spec);
            persistPickersToStorage(SETUP_GROUP_FIELDS);
        });
    });
}

function renderAllExtraUsersLists() {
    EXTRA_USER_UI.forEach((spec) => renderExtraUsersList(spec));
}

function setExtraUsersDraft(configKey, users) {
    extraUsersDraft[configKey] = normalizePlannerUsers(users);
    const spec = EXTRA_USER_UI.find((s) => s.configKey === configKey);
    if (spec) renderExtraUsersList(spec);
}

function wireExtraUsersUi(onChange) {
    EXTRA_USER_UI.forEach((spec) => {
        const addBtn = document.getElementById(spec.addBtnId);
        if (addBtn && !addBtn.dataset.frUserWired) {
            addBtn.dataset.frUserWired = '1';
            addBtn.addEventListener('click', () => {
                pickEntraUser({ title: spec.pickTitle, hint: spec.pickHint })
                    .then((sel) => {
                        if (!sel) return;
                        const mail = String(sel.mail || sel.userPrincipalName || '').trim().toLowerCase();
                        if (!mail) {
                            if (typeof window.ms365ToastOrAlert === 'function') {
                                window.ms365ToastOrAlert(
                                    'Keine E-Mail/UPN für diese Person – bitte anderen Datensatz wählen.'
                                );
                            }
                            return;
                        }
                        extraUsersDraft[spec.configKey] = mergePlannerUser(extraUsersDraft[spec.configKey], {
                            id: sel.id,
                            displayName: sel.displayName,
                            mail
                        });
                        renderExtraUsersList(spec);
                        if (typeof onChange === 'function') onChange();
                    })
                    .catch((e) => {
                        const msg = e && e.message ? e.message : String(e);
                        if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
                    });
            });
        }
        if (spec.stammdatenBtnId) {
            const stBtn = document.getElementById(spec.stammdatenBtnId);
            if (stBtn && !stBtn.dataset.frUserWired) {
                stBtn.dataset.frUserWired = '1';
                stBtn.addEventListener('click', () => {
                    const fromStamm =
                        spec.configKey === 'kvUsers'
                            ? kvUsersFromSchoolStammdaten()
                            : direktionUsersFromTenantStammdaten();
                    if (!fromStamm.length) {
                        if (typeof window.ms365ToastOrAlert === 'function') {
                            window.ms365ToastOrAlert(
                                spec.configKey === 'kvUsers'
                                    ? 'Keine Klassenvorstand-E-Mails in den Stammdaten – bitte unter Stammdaten → Klassen pflegen.'
                                    : 'Keine Verwaltungs-Personen in den Stammdaten – bitte unter Stammdaten → Verwaltung pflegen.'
                            );
                        }
                        return;
                    }
                    fromStamm.forEach((u) => {
                        extraUsersDraft[spec.configKey] = mergePlannerUser(extraUsersDraft[spec.configKey], u);
                    });
                    renderExtraUsersList(spec);
                    if (typeof onChange === 'function') onChange();
                    if (typeof window.ms365ToastOrAlert === 'function') {
                        window.ms365ToastOrAlert(fromStamm.length + ' Person(en) aus Stammdaten ergänzt.');
                    }
                });
            }
        }
    });
}

/**
 * @param {typeof SETUP_GROUP_FIELDS} fieldDefs
 */
export function readPermissionsFromPickers(_fieldDefs) {
    return {
        plannerGrantRows: readGrantRowsFromDom('frPermMatrixBody'),
        allowedJahrgang: readAllowedJahrgangInput(),
        jahrgangGroups: normalizeJahrgangGroups(jahrgangGroupsDraft)
    };
}

/**
 * @param {object} cfg
 * @param {typeof SETUP_GROUP_FIELDS} fieldDefs
 */
export function fillPermissionsPickers(cfg, fieldDefs) {
    const c = normalizePermissionsConfig(cfg);
    fieldDefs.forEach((f) => {
        fillGroupPickerField(f, {
            id: c[f.role + 'Id'],
            label: c[f.role]
        });
    });
}

/**
 * @param {typeof SETUP_GROUP_FIELDS} fieldDefs
 */
export function persistPickersToStorage(fieldDefs) {
    const patch = stripStammdatenGroupFieldsFromPatch(
        readPermissionsFromPickers(fieldDefs),
        FR_STAMMDATEN_GROUP_ROLES
    );
    const skip = document.getElementById('frSkipPerms');
    if (skip) patch.skipPerms = !!skip.checked;
    return savePermissionsConfig(patch);
}

/**
 * @param {typeof SETUP_GROUP_FIELDS} fieldDefs
 * @param {() => void} [onChange]
 */
export function wirePermissionGroupPickers(fieldDefs, onChange) {
    wireEntraGroupPickerFields({
        fields: filterEditableGroupFields(fieldDefs, FR_STAMMDATEN_GROUP_ROLES).map((f) => ({
            labelInputId: f.labelInputId,
            idInputId: f.idInputId,
            pickBtnId: f.pickBtnId,
            clearBtnId: f.clearBtnId,
            dialogTitle: f.dialogTitle
        })),
        onChange
    });
}

export function saveFreistellungPlannerPermissionsFromForm() {
    return persistPickersToStorage(SETUP_GROUP_FIELDS);
}

export async function saveAndPublishFreistellungPlannerGroups() {
    saveFreistellungPlannerPermissionsFromForm();
    let siteUrl = '';
    let listId = '';
    try {
        const setup = JSON.parse(localStorage.getItem('ms365-freistellung-setup-v1') || '{}');
        siteUrl = String(setup.siteUrl || '').trim();
        listId = String(setup.listId || '').trim();
    } catch {
        /* ignore */
    }
    if (!siteUrl) {
        return { ok: false, reason: 'no-site', local: true };
    }
    try {
        const sharePoint = await publishPlannerPermissionsToSite(
            siteUrl,
            loadEffectivePermissionsConfig(),
            listId || undefined
        );
        if (listId && sharePoint && sharePoint.ok) {
            try {
                const G = window.ms365SpoGraph;
                if (G) {
                    const tok = await G.getGraphToken([
                        'https://graph.microsoft.com/User.Read',
                        'https://graph.microsoft.com/Sites.ReadWrite.All'
                    ]);
                    const site = await G.resolveSiteFromWebUrl(tok, siteUrl);
                    await patchFreistellungKategorieColumn(site.id, listId, loadExtraKategorien());
                    const klasseCodes = classCatalogFromSchoolStammdaten()
                        .map((r) => r.code)
                        .filter(Boolean);
                    if (klasseCodes.length) {
                        await patchFreistellungKlasseColumn(site.id, listId, klasseCodes, {
                            webUrl: siteUrl
                        });
                    }
                }
            } catch {
                /* Kategorien-Spalte optional */
            }
        }
        if (!sharePoint || !sharePoint.ok) {
            return {
                ok: false,
                reason: (sharePoint && sharePoint.reason) || 'unknown',
                local: true,
                sharePoint
            };
        }
        return { ok: true, local: true, sharePoint };
    } catch (e) {
        return {
            ok: false,
            reason: 'error',
            message: e && e.message ? e.message : String(e),
            local: true
        };
    }
}

export function formatFreistellungPlannerPublishToast(pub) {
    if (!pub || !pub.local) return 'Gespeichert.';
    if (pub.ok && pub.sharePoint) {
        const sp = pub.sharePoint;
        const where = [];
        if (sp.listDescription) where.push('Listen-Beschreibung');
        if (sp.path) where.push('JSON ' + sp.path);
        return (
            'Planer-Gruppen auf SharePoint veröffentlicht' +
            (where.length ? ' (' + where.join(', ') + ')' : '') +
            '. Einstellungen zusätzlich in diesem Browser.'
        );
    }
    if (pub.reason === 'no-site') {
        return 'Gruppen in diesem Browser gespeichert. SharePoint: Site-URL im Setup fehlt – bitte eintragen und erneut speichern.';
    }
    if (pub.reason === 'no-groups') {
        return 'Gruppen in diesem Browser gespeichert. SharePoint: mindestens eine Entra-Gruppe (KV oder Schüler) wählen.';
    }
    if (pub.reason === 'error') {
        return (
            'Gruppen in diesem Browser gespeichert. SharePoint-Fehler: ' +
            (pub.message || 'unbekannt') +
            ' – ggf. anmelden und „Liste anlegen / prüfen“ ausführen.'
        );
    }
    return 'Gruppen in diesem Browser gespeichert.';
}

let setupPermissionsBooted = false;

export function initFreistellungSetupPermissions() {
    if (setupPermissionsBooted) return;
    try {
        initFreistellungSetupPermissionsCore();
        setupPermissionsBooted = true;
    } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        console.error('[freistellung-setup-perms]', e);
        if (typeof window.ms365ToastOrAlert === 'function') {
            window.ms365ToastOrAlert('Berechtigungs-UI konnte nicht starten: ' + msg);
        }
    }
}

function buildInitialGrantRows(cfg) {
    let rows = plannerGrantRowsForUi(cfg);
    if (!rows.length) rows = migrateLegacyFreistellungPerms(cfg);

    let explicitlySaved = { direktionUsers: false, kvUsers: false, schuelerUsers: false };
    try {
        const raw = JSON.parse(localStorage.getItem(PERMS_STORAGE_KEY) || '{}');
        explicitlySaved.direktionUsers = Array.isArray(raw.direktionUsers);
        explicitlySaved.kvUsers = Array.isArray(raw.kvUsers);
        explicitlySaved.schuelerUsers = Array.isArray(raw.schuelerUsers);
    } catch {
        /* ignore */
    }

    if (!explicitlySaved.direktionUsers) {
        direktionUsersFromTenantStammdaten().forEach(function (u) {
            rows.push({
                principalType: 'user',
                groupId: '',
                groupLabel: '',
                mail: u.mail,
                displayName: u.displayName || u.mail,
                roles: { direktion: true, kv: false, schueler: false }
            });
        });
    }
    if (!explicitlySaved.kvUsers) {
        kvUsersFromSchoolStammdaten().forEach(function (u) {
            rows.push({
                principalType: 'user',
                groupId: '',
                groupLabel: '',
                mail: u.mail,
                displayName: u.displayName || u.mail,
                roles: { direktion: false, kv: true, schueler: false }
            });
        });
    }
    if (!String(cfg.groupSchuelerId || '').trim()) {
        const aud = loadSchoolAudienceGroups();
        if (aud.groupSchuelerId) {
            rows.push({
                principalType: 'group',
                groupId: aud.groupSchuelerId,
                groupLabel: aud.groupSchuelerName || aud.groupSchuelerId,
                mail: '',
                displayName: '',
                roles: { direktion: false, kv: false, schueler: true }
            });
        }
    }
    return mergeDuplicateGrantRows(rows);
}

function wireStammdatenMatrixHelpers(persist) {
    const dirBtn = document.getElementById('frPermDirektionStammdaten');
    if (dirBtn && !dirBtn.dataset.frMatrixWired) {
        dirBtn.dataset.frMatrixWired = '1';
        dirBtn.addEventListener('click', function () {
            const fromStamm = direktionUsersFromTenantStammdaten();
            if (!fromStamm.length) {
                if (typeof window.ms365ToastOrAlert === 'function') {
                    window.ms365ToastOrAlert(
                        'Keine Verwaltungs-Personen in den Stammdaten – bitte unter Stammdaten → Verwaltung pflegen.'
                    );
                }
                return;
            }
            const tbody = document.getElementById('frPermMatrixBody');
            if (!tbody) return;
            fromStamm.forEach(function (u) {
                appendGrantRow(tbody, {
                    principalType: 'user',
                    mail: u.mail,
                    displayName: u.displayName || u.mail,
                    roles: { direktion: true, kv: false, schueler: false }
                });
            });
            persist();
            if (typeof window.ms365ToastOrAlert === 'function') {
                window.ms365ToastOrAlert(fromStamm.length + ' Person(en) aus Stammdaten ergänzt.');
            }
        });
    }
    const kvBtn = document.getElementById('frPermKvStammdaten');
    if (kvBtn && !kvBtn.dataset.frMatrixWired) {
        kvBtn.dataset.frMatrixWired = '1';
        kvBtn.addEventListener('click', function () {
            const fromStamm = kvUsersFromSchoolStammdaten();
            if (!fromStamm.length) {
                if (typeof window.ms365ToastOrAlert === 'function') {
                    window.ms365ToastOrAlert(
                        'Keine Klassenvorstand-E-Mails in den Stammdaten – bitte unter Stammdaten → Klassen pflegen.'
                    );
                }
                return;
            }
            const tbody = document.getElementById('frPermMatrixBody');
            if (!tbody) return;
            fromStamm.forEach(function (u) {
                appendGrantRow(tbody, {
                    principalType: 'user',
                    mail: u.mail,
                    displayName: u.displayName || u.mail,
                    roles: { direktion: false, kv: true, schueler: false }
                });
            });
            persist();
            if (typeof window.ms365ToastOrAlert === 'function') {
                window.ms365ToastOrAlert(fromStamm.length + ' KV-Person(en) aus Stammdaten ergänzt.');
            }
        });
    }
    const schBtn = document.getElementById('frPermSchuelerStammdaten');
    if (schBtn && !schBtn.dataset.frMatrixWired) {
        schBtn.dataset.frMatrixWired = '1';
        schBtn.addEventListener('click', function () {
            const aud = loadSchoolAudienceGroups();
            if (!aud.groupSchuelerId) {
                if (typeof window.ms365ToastOrAlert === 'function') {
                    window.ms365ToastOrAlert('Keine Schüler-Sammelgruppe in den Stammdaten konfiguriert.');
                }
                return;
            }
            const tbody = document.getElementById('frPermMatrixBody');
            if (!tbody) return;
            appendGrantRow(tbody, {
                principalType: 'group',
                groupId: aud.groupSchuelerId,
                groupLabel: aud.groupSchuelerName || aud.groupSchuelerId,
                roles: { direktion: false, kv: false, schueler: true }
            });
            persist();
            if (typeof window.ms365ToastOrAlert === 'function') {
                window.ms365ToastOrAlert('Schüler-Sammelgruppe aus Stammdaten ergänzt.');
            }
        });
    }
}

export function initFreistellungSetupPermissionsCore() {
    const cfg = loadPermissionsConfig();
    const initialRows = buildInitialGrantRows(cfg);

    jahrgangGroupsDraft = normalizeJahrgangGroups(cfg.jahrgangGroups);
    const allowedEl = document.getElementById('frAllowedJahrgang');
    if (allowedEl) {
        allowedEl.value = (cfg.allowedJahrgang || []).join(', ');
    }
    renderJahrgangGroupsList();

    const skip = document.getElementById('frSkipPerms');
    if (skip) skip.checked = !!cfg.skipPerms;
    const persist = () => persistPickersToStorage(SETUP_GROUP_FIELDS);
    initFreistellungPlannerGrantMatrixUi('frPermMatrixBody', initialRows, persist);
    wireStammdatenMatrixHelpers(persist);
    wireJahrgangScopeUi(persist);
    persist();
    const btn = document.getElementById('frBtnSavePerms');
    if (btn) {
        btn.addEventListener('click', async () => {
            const pub = await saveAndPublishFreistellungPlannerGroups();
            const msg = formatFreistellungPlannerPublishToast(pub);
            if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
        });
    }

    const setupSiteCtx = () => {
        let siteUrl = '';
        let listId = '';
        try {
            const setup = JSON.parse(localStorage.getItem('ms365-freistellung-setup-v1') || '{}');
            siteUrl = String(setup.siteUrl || '').trim();
            listId = String(setup.listId || '').trim();
        } catch {
            /* ignore */
        }
        return { siteUrl, listId };
    };
    try {
        if (window.ms365StammdatenCanonical && typeof window.ms365StammdatenCanonical.reconcileStammdatenStorage === 'function') {
            window.ms365StammdatenCanonical.reconcileStammdatenStorage();
        }
    } catch {
        /* ignore */
    }
    wireFreistellungKlassenAdmin(
        {
            mountId: 'frSetupKlassenList',
            syncBtnId: 'frSetupKlassenSyncSp',
            reloadBtnId: 'frSetupKlassenReload',
            cleanBtnId: 'frSetupKlassenClean'
        },
        setupSiteCtx
    );
    wireFreistellungKategorienAdmin(
        {
            listId: 'frSetupKatExtraList',
            addId: 'frSetupKatExtraAdd',
            newInputId: 'frSetupKatExtraNew',
            syncListBtnId: 'frSetupKatSyncSp'
        },
        setupSiteCtx
    );
}
