/**
 * Entra-Gruppen (mehrere) + Einzelpersonen für sRDP-Listenrechte
 * (Vollzugriff / Lesende Rechte / Kandidat:innen).
 */
import { pickEntraGroup } from '../../shared/entra-group-picker.js';
import { createPlannerExtraUsersController } from '../../shared/planner-extra-users-ui.js';
import { normalizePlannerUsers } from '../freistellung-planer/freistellung-planer-direktion-users.js';
import { loadSchoolAudienceGroups } from '../../shared/school-audience-groups.js';
import {
    normalizeEntraGroupEntries,
    mergeLegacyEntraGroup,
    formatEntraGroupsSummary
} from '../../shared/entra-group-normalize.js';

export const SRS_PERMS_STORAGE_KEY = 'ms365-srdp-list-perms-v1';

/** @typedef {{ groupId: string, groupLabel: string }} SrdpEntraGroup */

const GROUP_SPECS = [
    {
        configKey: 'adminGroups',
        listId: 'srsAdminGroups',
        addBtnId: 'srsAdminGroupAdd',
        pickTitle: 'Gruppe für Vollzugriff',
        pickHint: 'Microsoft-365-Gruppe mit Vollzugriff auf die Anmeldeliste.',
        emptyHint: 'Noch keine Gruppe – Gruppe hinzufügen oder Person wählen.',
        ariaLabel: 'Vollzugriff-Gruppen'
    },
    {
        configKey: 'lehrerGroups',
        listId: 'srsLehrerGroups',
        addBtnId: 'srsLehrerGroupAdd',
        pickTitle: 'Gruppe für Lesende Rechte',
        pickHint: 'Gruppe, die alle Anmeldungen sehen darf (z. B. Lehrkräfte, Sekretariat).',
        emptyHint: 'Noch keine Gruppe.',
        ariaLabel: 'Lesende-Rechte-Gruppen'
    },
    {
        configKey: 'kandidatenGroups',
        listId: 'srsKandidatenGroups',
        addBtnId: 'srsKandidatenGroupAdd',
        pickTitle: 'Gruppe für Kandidat:innen',
        pickHint: 'Klassen-/Jahrgangsgruppe mit Mitwirken auf der eigenen Anmeldung.',
        emptyHint: 'Noch keine Gruppe.',
        ariaLabel: 'Kandidat:innen-Gruppen'
    }
];

const EXTRA_USER_SPECS = [
    {
        configKey: 'adminUsers',
        listId: 'srsAdminUsers',
        addBtnId: 'srsAdminUserAdd',
        stammdatenBtnId: 'srsAdminStammdaten',
        pickTitle: 'Person mit Vollzugriff',
        pickHint: 'Vollzugriff auf die sRDP-Liste (zusätzlich zu den Gruppen).',
        emptyHint: 'Noch keine Einzelpersonen – Gruppe wählen oder Person hinzufügen.'
    },
    {
        configKey: 'lehrerUsers',
        listId: 'srsLehrerUsers',
        addBtnId: 'srsLehrerUserAdd',
        pickTitle: 'Person mit Lesenden Rechten',
        pickHint: 'Sieht alle Anmeldungen (z. B. Sekretariat, einzelne Lehrkräfte).',
        emptyHint: 'Noch keine Einzelpersonen.'
    },
    {
        configKey: 'kandidatenUsers',
        listId: 'srsKandidatenUsers',
        addBtnId: 'srsKandidatenUserAdd',
        pickTitle: 'Kandidat:in einzeln',
        pickHint: 'Mitwirken auf der sRDP-Liste (zusätzlich zu den Jahrgangs-Gruppen).',
        emptyHint: 'Noch keine Einzelpersonen.'
    }
];

const extraCtrl = createPlannerExtraUsersController(EXTRA_USER_SPECS);

/** @type {Record<'adminGroups'|'lehrerGroups'|'kandidatenGroups', SrdpEntraGroup[]>} */
const groupsDraft = {
    adminGroups: [],
    lehrerGroups: [],
    kandidatenGroups: []
};

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * @param {unknown} raw
 */
export function normalizeSrdpPermissionsConfig(raw) {
    const r = raw && typeof raw === 'object' ? raw : {};
    const adminGroups = mergeLegacyEntraGroup(r.adminGroups, r.groupAdminId, r.groupAdmin);
    const lehrerGroups = mergeLegacyEntraGroup(r.lehrerGroups, r.groupLehrerId, r.groupLehrer);
    const kandidatenGroups = mergeLegacyEntraGroup(
        r.kandidatenGroups,
        r.groupKandidatenId,
        r.groupKandidaten
    );
    return {
        adminGroups,
        lehrerGroups,
        kandidatenGroups,
        groupAdmin: adminGroups[0] ? adminGroups[0].groupLabel : '',
        groupAdminId: adminGroups[0] ? adminGroups[0].groupId : '',
        groupLehrer: lehrerGroups[0] ? lehrerGroups[0].groupLabel : '',
        groupLehrerId: lehrerGroups[0] ? lehrerGroups[0].groupId : '',
        groupKandidaten: kandidatenGroups[0] ? kandidatenGroups[0].groupLabel : '',
        groupKandidatenId: kandidatenGroups[0] ? kandidatenGroups[0].groupId : '',
        adminUsers: normalizePlannerUsers(r.adminUsers),
        lehrerUsers: normalizePlannerUsers(r.lehrerUsers),
        kandidatenUsers: normalizePlannerUsers(r.kandidatenUsers),
        skipPerms: !!r.skipPerms
    };
}

export function loadSrdpPermissionsConfig() {
    try {
        const raw = JSON.parse(localStorage.getItem(SRS_PERMS_STORAGE_KEY) || '{}');
        return normalizeSrdpPermissionsConfig(raw);
    } catch {
        return normalizeSrdpPermissionsConfig({});
    }
}

/**
 * Stammdaten-Vorschläge (Lehrkräfte-Gruppe), falls noch nichts gesetzt.
 */
export function loadEffectiveSrdpPermissionsConfig() {
    const saved = loadSrdpPermissionsConfig();
    const aud = loadSchoolAudienceGroups();
    const next = { ...saved };
    if (!next.lehrerGroups.length && aud.groupLehrerId) {
        next.lehrerGroups = mergeLegacyEntraGroup(
            [],
            aud.groupLehrerId,
            aud.groupLehrerName || aud.groupLehrerId
        );
        next.groupLehrerId = next.lehrerGroups[0] ? next.lehrerGroups[0].groupId : '';
        next.groupLehrer = next.lehrerGroups[0] ? next.lehrerGroups[0].groupLabel : '';
    }
    return normalizeSrdpPermissionsConfig(next);
}

/**
 * @param {Partial<ReturnType<typeof normalizeSrdpPermissionsConfig>>} patch
 */
export function saveSrdpPermissionsConfig(patch) {
    const next = normalizeSrdpPermissionsConfig({
        ...loadSrdpPermissionsConfig(),
        ...(patch || {})
    });
    try {
        localStorage.setItem(SRS_PERMS_STORAGE_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    return next;
}

/**
 * @param {'admin'|'lehrer'|'kandidaten'} which
 */
function htmlPermCard(which) {
    const gSpec =
        which === 'admin' ? GROUP_SPECS[0] : which === 'lehrer' ? GROUP_SPECS[1] : GROUP_SPECS[2];
    const uSpec =
        which === 'admin'
            ? EXTRA_USER_SPECS[0]
            : which === 'lehrer'
              ? EXTRA_USER_SPECS[1]
              : EXTRA_USER_SPECS[2];

    if (which === 'admin') {
        return `<article class="fr-setup-perm-card">
          <div class="fr-setup-perm-card__head">
            <span class="fr-setup-perm-card__icon" aria-hidden="true"><i class="bi bi-building"></i></span>
            <div>
              <h4>Vollzugriff</h4>
              <p>Volle Rechte auf die Anmeldeliste – z. B. Direktion, IT, Verwaltung.</p>
            </div>
          </div>
          <p class="fr-setup-perm-sub">Microsoft 365-Gruppen (mehrere möglich)</p>
          <ul id="${gSpec.listId}" class="fr-setup-user-list" aria-label="${escapeHtml(gSpec.ariaLabel)}"></ul>
          <div class="fr-setup-egp fr-setup-egp--tight">
            <button type="button" class="btn btn-sm" id="${gSpec.addBtnId}"><i class="bi bi-people"></i>Gruppe</button>
          </div>
          <p class="fr-setup-perm-sub">Einzelpersonen (optional)</p>
          <ul id="${uSpec.listId}" class="fr-setup-user-list" aria-label="Vollzugriff Einzelpersonen"></ul>
          <div class="fr-setup-egp fr-setup-egp--tight">
            <button type="button" class="btn btn-sm" id="${uSpec.addBtnId}"><i class="bi bi-person-plus"></i>Person</button>
            <button type="button" class="btn btn-sm alt" id="${uSpec.stammdatenBtnId}" title="Verwaltung aus Stammdaten"><i class="bi bi-journal-text"></i>Aus Stammdaten</button>
          </div>
        </article>`;
    }
    if (which === 'lehrer') {
        return `<article class="fr-setup-perm-card">
          <div class="fr-setup-perm-card__head">
            <span class="fr-setup-perm-card__icon" aria-hidden="true"><i class="bi bi-eye"></i></span>
            <div>
              <h4>Lesende Rechte</h4>
              <p>Alle Anmeldungen sehen – z. B. Lehrkräfte, Sekretariat. Kandidat:innen sehen nur die eigene Zeile.</p>
            </div>
          </div>
          <p class="fr-setup-perm-sub">Microsoft 365-Gruppen (mehrere möglich)</p>
          <ul id="${gSpec.listId}" class="fr-setup-user-list" aria-label="${escapeHtml(gSpec.ariaLabel)}"></ul>
          <div class="fr-setup-egp fr-setup-egp--tight">
            <button type="button" class="btn btn-sm" id="${gSpec.addBtnId}"><i class="bi bi-people"></i>Gruppe</button>
          </div>
          <p class="fr-setup-perm-sub">Einzelpersonen (optional)</p>
          <ul id="${uSpec.listId}" class="fr-setup-user-list" aria-label="Lesende Rechte Einzelpersonen"></ul>
          <div class="fr-setup-egp fr-setup-egp--tight">
            <button type="button" class="btn btn-sm" id="${uSpec.addBtnId}"><i class="bi bi-person-plus"></i>Person</button>
          </div>
        </article>`;
    }
    return `<article class="fr-setup-perm-card">
      <div class="fr-setup-perm-card__head">
        <span class="fr-setup-perm-card__icon" aria-hidden="true"><i class="bi bi-mortarboard"></i></span>
        <div>
          <h4>Kandidat:innen</h4>
          <p>Mitwirken – nur die eigene Anmeldung sehen und pflegen.</p>
        </div>
      </div>
      <p class="fr-setup-perm-sub">Microsoft 365-Gruppen (mehrere möglich)</p>
      <ul id="${gSpec.listId}" class="fr-setup-user-list" aria-label="${escapeHtml(gSpec.ariaLabel)}"></ul>
      <div class="fr-setup-egp fr-setup-egp--tight">
        <button type="button" class="btn btn-sm" id="${gSpec.addBtnId}"><i class="bi bi-people"></i>Gruppe</button>
      </div>
      <p class="fr-setup-perm-sub">Einzelpersonen (optional)</p>
      <ul id="${uSpec.listId}" class="fr-setup-user-list" aria-label="Kandidat:innen Einzelpersonen"></ul>
      <div class="fr-setup-egp fr-setup-egp--tight">
        <button type="button" class="btn btn-sm" id="${uSpec.addBtnId}"><i class="bi bi-person-plus"></i>Person</button>
      </div>
    </article>`;
}

/**
 * @param {ReturnType<typeof normalizeSrdpPermissionsConfig>} [_p]
 */
export function htmlSrdpSetupEntraPermGrid(_p) {
    return `<div class="fr-setup-perm-grid" role="group" aria-label="Gruppen und Rechte">
      ${htmlPermCard('admin')}
      ${htmlPermCard('lehrer')}
      ${htmlPermCard('kandidaten')}
    </div>`;
}

function renderGroupsList(spec) {
    const ul = document.getElementById(spec.listId);
    if (!ul) return;
    const key = /** @type {'adminGroups'|'lehrerGroups'|'kandidatenGroups'} */ (spec.configKey);
    groupsDraft[key] = normalizeEntraGroupEntries(groupsDraft[key]);
    const list = groupsDraft[key];
    if (!list.length) {
        ul.innerHTML =
            '<li class="fr-setup-user-row fr-setup-user-row--empty muted">' +
            escapeHtml(spec.emptyHint) +
            '</li>';
        return;
    }
    ul.innerHTML = list
        .map(
            (g, idx) =>
                `<li class="fr-setup-user-row">` +
                `<span class="fr-setup-user-row__label">${escapeHtml(g.groupLabel || g.groupId)}</span>` +
                `<button type="button" class="fr-setup-user-row__rm btn btn-sm alt" data-srs-eg-rm="${key}:${idx}" title="Entfernen" aria-label="Entfernen"><i class="bi bi-x"></i></button>` +
                `</li>`
        )
        .join('');
    ul.querySelectorAll('[data-srs-eg-rm]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const raw = String(btn.getAttribute('data-srs-eg-rm') || '');
            const sep = raw.indexOf(':');
            const cfgKey = /** @type {'adminGroups'|'lehrerGroups'|'kandidatenGroups'|''} */ (
                raw.slice(0, sep)
            );
            const i = Number(raw.slice(sep + 1));
            if (!cfgKey || !groupsDraft[cfgKey]) return;
            const next = normalizeEntraGroupEntries(groupsDraft[cfgKey]);
            next.splice(i, 1);
            groupsDraft[cfgKey] = next;
            renderGroupsList(spec);
            persistSrdpPickersToStorage();
        });
    });
}

export function readSrdpPermissionsFromPickers() {
    const skipEl = document.getElementById('srsSkipPerms');
    const out = {
        skipPerms: !!(skipEl && skipEl.checked),
        adminGroups: normalizeEntraGroupEntries(groupsDraft.adminGroups),
        lehrerGroups: normalizeEntraGroupEntries(groupsDraft.lehrerGroups),
        kandidatenGroups: normalizeEntraGroupEntries(groupsDraft.kandidatenGroups)
    };
    Object.assign(out, extraCtrl.readPatch());
    return normalizeSrdpPermissionsConfig(out);
}

export function fillSrdpPermissionsPickers(cfg) {
    const c = normalizeSrdpPermissionsConfig(cfg || loadEffectiveSrdpPermissionsConfig());
    groupsDraft.adminGroups = c.adminGroups.slice();
    groupsDraft.lehrerGroups = c.lehrerGroups.slice();
    groupsDraft.kandidatenGroups = c.kandidatenGroups.slice();
    GROUP_SPECS.forEach((spec) => renderGroupsList(spec));
    let explicit = {};
    try {
        const raw = JSON.parse(localStorage.getItem(SRS_PERMS_STORAGE_KEY) || '{}');
        explicit = {
            adminUsers: Array.isArray(raw.adminUsers),
            lehrerUsers: Array.isArray(raw.lehrerUsers),
            kandidatenUsers: Array.isArray(raw.kandidatenUsers)
        };
    } catch {
        explicit = {};
    }
    extraCtrl.loadFromConfig(c, explicit);
}

/**
 * @param {() => void} [onChange]
 */
export function wireSrdpPermissionPickers(onChange) {
    GROUP_SPECS.forEach((spec) => {
        renderGroupsList(spec);
        const addBtn = document.getElementById(spec.addBtnId);
        if (addBtn && !addBtn.dataset.srsEgWired) {
            addBtn.dataset.srsEgWired = '1';
            addBtn.addEventListener('click', () => {
                pickEntraGroup({
                    title: spec.pickTitle,
                    hint: spec.pickHint
                })
                    .then((sel) => {
                        if (!sel || !sel.id) return;
                        const key = /** @type {'adminGroups'|'lehrerGroups'|'kandidatenGroups'} */ (
                            spec.configKey
                        );
                        groupsDraft[key] = normalizeEntraGroupEntries([
                            ...groupsDraft[key],
                            {
                                groupId: sel.id,
                                groupLabel: sel.label || sel.displayName || sel.id
                            }
                        ]);
                        renderGroupsList(spec);
                        if (typeof onChange === 'function') onChange();
                    })
                    .catch((e) => {
                        const msg = e && e.message ? e.message : String(e);
                        if (typeof window.ms365ToastOrAlert === 'function') {
                            window.ms365ToastOrAlert(msg);
                        }
                    });
            });
        }
    });
    extraCtrl.wire(() => {
        if (typeof onChange === 'function') onChange();
    });
}

export function persistSrdpPickersToStorage(skipPerms) {
    const patch = readSrdpPermissionsFromPickers();
    if (skipPerms != null) patch.skipPerms = !!skipPerms;
    return saveSrdpPermissionsConfig(patch);
}

/**
 * Host einmal mounten, Pickers verdrahten, Werte füllen.
 * @param {HTMLElement|null} host
 * @param {() => void} [onChange]
 */
export function mountSrdpEntraPermHost(host, onChange) {
    if (!host) return;
    // v2: Multi-Gruppen-Listen statt Einzel-Picker
    if (host.dataset.srsPermMounted !== '2') {
        host.innerHTML = htmlSrdpSetupEntraPermGrid(loadEffectiveSrdpPermissionsConfig());
        host.dataset.srsPermMounted = '2';
        wireSrdpPermissionPickers(onChange);
    }
    fillSrdpPermissionsPickers(loadEffectiveSrdpPermissionsConfig());
}

/**
 * Kurztext für Zusammenfassung.
 * @param {{ label?: string, id?: string, groups?: unknown[], users?: unknown[] }} slot
 */
export function formatSrdpAudienceSummary(slot) {
    const groups = normalizeEntraGroupEntries(
        (slot && slot.groups) ||
            (slot && slot.id
                ? [{ groupId: slot.id, groupLabel: slot.label || '' }]
                : [])
    );
    const groupLabel = formatEntraGroupsSummary(groups) || String((slot && slot.label) || '').trim();
    const users = normalizePlannerUsers(slot && slot.users);
    if (groupLabel && users.length) return groupLabel + ' + ' + users.length + ' Person(en)';
    if (groupLabel) return groupLabel;
    if (users.length) return users.length + ' Person(en)';
    return '— nicht gesetzt —';
}
