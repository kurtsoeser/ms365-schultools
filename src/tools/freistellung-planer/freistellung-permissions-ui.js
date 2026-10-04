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
    savePermissionsConfig,
    loadPermissionsConfig,
    PERMS_STORAGE_KEY
} from './freistellung-planer-permissions.js';
import { publishPlannerPermissionsToSite } from './freistellung-planer-remote-config.js';
import { loadExtraKategorien } from './freistellung-planer-kategorien.js';
import { patchFreistellungKategorieColumn } from './freistellung-planer-graph.js';
import { wireFreistellungKategorienAdmin } from './freistellung-kategorien-ui.js';
import { pickEntraUser } from '../../shared/entra-user-picker.js';
import {
    normalizePlannerUsers,
    mergePlannerUser,
    direktionUsersFromTenantStammdaten
} from './freistellung-planer-direktion-users.js';

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

/** @type {Record<string, Array<{ id: string, displayName: string, mail: string }>>} */
const extraUsersDraft = {
    direktionUsers: [],
    kvUsers: [],
    schuelerUsers: []
};

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
        pickTitle: 'Person als Klassenvorstand',
        pickHint: 'Zusätzlich zur KV-Entra-Gruppe – z. B. Vertretung oder einzelner KV ohne Gruppe.',
        emptyHint: 'Noch keine Einzelpersonen.'
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
        if (addBtn) {
            addBtn.addEventListener('click', () => {
                pickEntraUser({ title: spec.pickTitle, hint: spec.pickHint })
                    .then((sel) => {
                        if (!sel || !sel.mail) return;
                        extraUsersDraft[spec.configKey] = mergePlannerUser(extraUsersDraft[spec.configKey], {
                            id: sel.id,
                            displayName: sel.displayName,
                            mail: sel.mail
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
            if (stBtn) {
                stBtn.addEventListener('click', () => {
                    const fromStamm = direktionUsersFromTenantStammdaten();
                    if (!fromStamm.length) {
                        if (typeof window.ms365ToastOrAlert === 'function') {
                            window.ms365ToastOrAlert(
                                'Keine Verwaltungs-Personen in den Stammdaten – bitte unter Stammdaten → Verwaltung pflegen.'
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
export function readPermissionsFromPickers(fieldDefs) {
    const out = {};
    fieldDefs.forEach((f) => {
        const r = readGroupPickerField(f);
        out[f.role] = r.label;
        out[f.role + 'Id'] = r.id;
    });
    out.direktionUsers = normalizePlannerUsers(extraUsersDraft.direktionUsers);
    out.kvUsers = normalizePlannerUsers(extraUsersDraft.kvUsers);
    out.schuelerUsers = normalizePlannerUsers(extraUsersDraft.schuelerUsers);
    return out;
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
    const patch = readPermissionsFromPickers(fieldDefs);
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
        fields: fieldDefs.map((f) => ({
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
            loadPermissionsConfig(),
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

export function initFreistellungSetupPermissions() {
    const cfg = loadPermissionsConfig();
    fillPermissionsPickers(cfg, SETUP_GROUP_FIELDS);

    let explicitlySaved = { direktionUsers: false, kvUsers: false, schuelerUsers: false };
    try {
        const raw = JSON.parse(localStorage.getItem(PERMS_STORAGE_KEY) || '{}');
        explicitlySaved.direktionUsers = Array.isArray(raw.direktionUsers);
        explicitlySaved.kvUsers = Array.isArray(raw.kvUsers);
        explicitlySaved.schuelerUsers = Array.isArray(raw.schuelerUsers);
    } catch {
        /* ignore */
    }

    let dirUsers = normalizePlannerUsers(cfg.direktionUsers);
    if (!dirUsers.length && !explicitlySaved.direktionUsers) {
        dirUsers = direktionUsersFromTenantStammdaten();
    }
    setExtraUsersDraft('direktionUsers', dirUsers);
    setExtraUsersDraft('kvUsers', cfg.kvUsers);
    setExtraUsersDraft('schuelerUsers', cfg.schuelerUsers);

    const skip = document.getElementById('frSkipPerms');
    if (skip) skip.checked = !!cfg.skipPerms;
    wirePermissionGroupPickers(SETUP_GROUP_FIELDS, () => {
        persistPickersToStorage(SETUP_GROUP_FIELDS);
    });
    wireExtraUsersUi(() => {
        persistPickersToStorage(SETUP_GROUP_FIELDS);
    });
    const btn = document.getElementById('frBtnSavePerms');
    if (btn) {
        btn.addEventListener('click', async () => {
            const pub = await saveAndPublishFreistellungPlannerGroups();
            const msg = formatFreistellungPlannerPublishToast(pub);
            if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
        });
    }

    wireFreistellungKategorienAdmin(
        {
            listId: 'frSetupKatExtraList',
            addId: 'frSetupKatExtraAdd',
            newInputId: 'frSetupKatExtraNew',
            syncListBtnId: 'frSetupKatSyncSp'
        },
        () => {
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
        }
    );
}
