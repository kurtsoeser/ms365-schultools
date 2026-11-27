/**
 * Analyse 02 Phase B – ESM + shared/utils/strings.
 */
import { normStr, normCode, normEmail } from '../../shared/utils/strings.js';
import {
    afgNeedsDisplayNameCleanup,
    buildAfgBulkCreateConfirmMessage,
    buildAfgBulkMatchExistingConfirmMessage,
    buildAfgBulkRenameConfirmMessage,
    buildAfgMailNickname,
    collectAfgCreateOwners,
    displayNameForAfgCreate,
    formatAfgBulkCreatePreviewLines,
    pickAfgGroupByNickname,
    stripAfgDisplayNamePrefix
} from './arge-fachgruppen-bulk-logic.js';

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

/** @type {'subject' | 'arge'} */
let activeKind = 'subject';
let activeCode = '';
let listFilter = '';
/** @type {'create' | 'edit'} */
let catalogModalMode = 'create';
/** Original-Kürzel beim Bearbeiten (für Umbenennung / Match-Migration). */
let catalogEditOriginalCode = '';
/** Nutzer hat Unique ID beim Anlegen manuell geändert – Standard nicht überschreiben. */
let createNickDirty = false;
/** @type {{ code: string, name: string, subjects?: string[] }[]} */
let catalog = { subject: [], arge: [] };
/** @type {string[]} */
let direktion = [];
/** @type {{ subject: Set<string>, arge: Set<string> }} */
let selectedKeysByKind = { subject: new Set(), arge: new Set() };
/** Mehrfachauswahl: Checkboxen nur im Selection Mode. */
let selectionMode = false;

function selectedKeys() {
    if (!selectedKeysByKind[activeKind]) selectedKeysByKind[activeKind] = new Set();
    return selectedKeysByKind[activeKind];
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
    } else if (typeof window.ms365ToastOrAlert === 'function') {
        window.ms365ToastOrAlert(msg);
    } else {
        window.alert(msg);
    }
}

function dlgConfirm(message, options) {
    if (typeof window.ms365AppDialogConfirm === 'function') {
        return window.ms365AppDialogConfirm(message, options || {});
    }
    return Promise.resolve(window.confirm(message));
}

function getCatalogLink(kind, code) {
    const api = dataV2();
    if (api && typeof api.getCatalogLink === 'function') return api.getCatalogLink(kind, code);
    return null;
}

function upsertCatalogLink(entry) {
    const api = dataV2();
    if (api && typeof api.upsertCatalogLink === 'function') return api.upsertCatalogLink(entry);
    return null;
}

function rowsForKind(kind) {
    return kind === 'arge' ? catalog.arge : catalog.subject;
}

function getActiveRow() {
    const list = rowsForKind(activeKind);
    const code = normCode(activeCode);
    for (let i = 0; i < list.length; i++) {
        if (normCode(list[i].code) === code) return list[i];
    }
    return null;
}

function getActiveGroupId() {
    const link = getCatalogLink(activeKind, activeCode);
    const id = link && link.graphGroupId ? String(link.graphGroupId).trim() : '';
    return id || null;
}

function mailPrefix(kind) {
    const api = dataV2();
    const su = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
    const def = kind === 'arge' ? 'arge-' : 'fach';
    const stored = kind === 'arge' ? su && su.argeGroupMailPrefix : su && su.subjectGroupMailPrefix;
    // null/undefined → Default; '' bleibt leer (Alias = nur Kürzel).
    const raw = stored == null ? def : String(stored);
    if (api && typeof api.mailNicknamePrefixSanitize === 'function') {
        return api.mailNicknamePrefixSanitize(raw, 24);
    }
    return raw
        .trim()
        .toLowerCase()
        .replace(/[@()[\]\\";:<>,\s]/g, '')
        .slice(0, 24);
}

function deriveNick(kind, code) {
    return buildAfgMailNickname(mailPrefix(kind), code);
}

/** Einmalig alten Default „ag“ (→ agfvv) auf „arge-“ umstellen. */
function maybeUpgradeLegacyAgPrefix() {
    try {
        const api = dataV2();
        if (!api || typeof api.getSetup !== 'function' || typeof api.patchSetup !== 'function') return;
        const su = api.getSetup();
        const cur =
            typeof api.mailNicknamePrefixSanitize === 'function'
                ? api.mailNicknamePrefixSanitize(su && su.argeGroupMailPrefix, 24)
                : String((su && su.argeGroupMailPrefix) || '')
                      .trim()
                      .toLowerCase();
        if (cur !== 'ag') return;
        const key = 'ms365-arge-mail-prefix-ag-upgraded-v1';
        if (localStorage.getItem(key) === '1') return;
        api.patchSetup({ argeGroupMailPrefix: 'arge-' });
        localStorage.setItem(key, '1');
        toast('Mail‑Alias‑Präfix für ARGE von „ag“ auf „arge-“ umgestellt (einstellbar unter dem Katalog).');
    } catch (_) {
        /* ignore */
    }
}

function syncMailPrefixUi() {
    const inp = document.getElementById('afgMailPrefix');
    const preview = document.getElementById('afgMailPrefixPreview');
    if (inp && document.activeElement !== inp) {
        inp.value = mailPrefix(activeKind);
        inp.placeholder = 'leer = nur Kürzel';
    }
    const sampleCode = activeCode || (activeKind === 'arge' ? 'FVV' : 'M');
    const nick = deriveNick(activeKind, sampleCode);
    if (preview) preview.textContent = nick || '–';
    updateCatalogNickHint();
}

function persistMailPrefixFromDom() {
    const inp = document.getElementById('afgMailPrefix');
    if (!inp) return;
    const api = dataV2();
    if (!api || typeof api.patchSetup !== 'function') return;
    let v = String(inp.value || '').trim();
    if (typeof api.mailNicknamePrefixSanitize === 'function') {
        v = api.mailNicknamePrefixSanitize(v, 24);
    } else {
        v = v
            .toLowerCase()
            .replace(/[@()[\]\\";:<>,\s]/g, '')
            .slice(0, 24);
    }
    inp.value = v;
    if (activeKind === 'arge') {
        api.patchSetup({ argeGroupMailPrefix: v });
    } else {
        api.patchSetup({ subjectGroupMailPrefix: v });
    }
    applyCreateDefaults();
    syncMailPrefixUi();
}

function updateCatalogNickHint() {
    const hint = document.getElementById('afgNewNickHint');
    if (!hint) return;
    const modal = document.getElementById('afgCatalogModal');
    const open = modal && modal.classList.contains('open');
    if (!open) {
        hint.textContent = '';
        return;
    }
    const codeEl = document.getElementById('afgNewCode');
    const code = normCode(codeEl && codeEl.value);
    const pre = mailPrefix(activeKind);
    if (!code) {
        hint.textContent = pre
            ? 'Beim Anlegen der M365‑Gruppe: Präfix „' + pre + '“ + Kürzel.'
            : 'Beim Anlegen der M365‑Gruppe: Alias = nur das Kürzel (kein Präfix).';
        return;
    }
    hint.textContent = 'Vorschau Unique ID / Alias: ' + deriveNick(activeKind, code);
}

function isDirektionRole(roleRaw) {
    const r = normStr(roleRaw).toLowerCase();
    return !!r && (r.indexOf('direktion') !== -1 || r.indexOf('direktor') !== -1);
}

function readLists() {
    const settings = typeof window.ms365TenantSettingsLoad === 'function' ? window.ms365TenantSettingsLoad() : null;
    catalog.subject = Array.isArray(settings && settings.subjects) ? settings.subjects.slice() : [];
    catalog.arge = Array.isArray(settings && settings.arges) ? settings.arges.slice() : [];
    const out = [];
    const seen = new Set();
    const admin = settings && Array.isArray(settings.admin) ? settings.admin : [];
    admin.forEach(function (row) {
        if (!isDirektionRole(row && row.role)) return;
        const em = normEmail(row && row.email);
        if (!em || em.indexOf('@') === -1 || seen.has(em)) return;
        seen.add(em);
        out.push(em);
    });
    direktion = out;
}

function dispatchTenantSettingsChanged(saved, reason) {
    try {
        window.dispatchEvent(
            new CustomEvent('ms365-tenant-settings-changed', {
                detail: { settings: saved, reason: reason || 'afg-catalog' }
            })
        );
    } catch (_) {
        /* ignore */
    }
}

function setCatalogModalError(msg) {
    const el = document.getElementById('afgCatalogModalError');
    if (!el) return;
    const text = normStr(msg);
    el.textContent = text;
    el.style.display = text ? '' : 'none';
}

function fillArgeSubjectPicker(selectedCodes) {
    const host = document.getElementById('afgNewSubjectsList');
    const empty = document.getElementById('afgNewSubjectsEmpty');
    if (!host) return;
    host.replaceChildren();
    const selected = new Set(
        (Array.isArray(selectedCodes) ? selectedCodes : []).map(function (c) {
            return normCode(c);
        })
    );
    const subjects = catalog.subject.slice().sort(function (a, b) {
        return String(a.code || '').localeCompare(String(b.code || ''), 'de', { sensitivity: 'base' });
    });
    if (empty) empty.style.display = subjects.length ? 'none' : '';
    subjects.forEach(function (row) {
        const id = 'afgSubj_' + normCode(row.code);
        const lab = document.createElement('label');
        lab.setAttribute('for', id);
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.id = id;
        cb.value = normCode(row.code);
        if (selected.has(normCode(row.code))) cb.checked = true;
        const span = document.createElement('span');
        span.textContent = (row.name || row.code) + (row.name && row.code ? ' (' + row.code + ')' : '');
        lab.appendChild(cb);
        lab.appendChild(span);
        host.appendChild(lab);
    });
}

function selectedArgeSubjects() {
    const host = document.getElementById('afgNewSubjectsList');
    if (!host) return [];
    const out = [];
    const seen = new Set();
    host.querySelectorAll('input[type="checkbox"]:checked').forEach(function (cb) {
        const code = normCode(cb.value);
        if (!code || seen.has(code)) return;
        seen.add(code);
        out.push(code);
    });
    return out;
}

function updateCatalogActionButtons() {
    const has = !!getActiveRow();
    const editBtn = document.getElementById('afgBtnEditCatalog');
    const delBtn = document.getElementById('afgBtnDeleteCatalog');
    if (editBtn) editBtn.disabled = !has;
    if (delBtn) delBtn.disabled = !has;
}

function openCatalogModal(mode) {
    const modal = document.getElementById('afgCatalogModal');
    if (!modal) return;
    catalogModalMode = mode === 'edit' ? 'edit' : 'create';
    const isArge = activeKind === 'arge';
    const title = document.getElementById('afgCatalogModalTitle');
    const hint = document.getElementById('afgCatalogModalHint');
    const wrap = document.getElementById('afgNewSubjectsWrap');
    const codeEl = document.getElementById('afgNewCode');
    const nameEl = document.getElementById('afgNewName');
    const saveBtn = document.getElementById('afgCatalogModalSave');
    const row = catalogModalMode === 'edit' ? getActiveRow() : null;

    if (catalogModalMode === 'edit' && !row) {
        toast('Bitte zuerst einen Katalogeintrag wählen.');
        return;
    }

    catalogEditOriginalCode = catalogModalMode === 'edit' ? normCode(row.code) : '';

    if (title) {
        if (catalogModalMode === 'edit') {
            title.textContent = isArge ? 'ARGE bearbeiten' : 'Fachgruppe bearbeiten';
        } else {
            title.textContent = isArge ? 'Neue ARGE anlegen' : 'Neue Fachgruppe anlegen';
        }
    }
    if (hint) {
        if (catalogModalMode === 'edit') {
            hint.textContent = isArge
                ? 'Änderungen werden in die Stammdaten geschrieben. Bei neuem Kürzel wird ein vorhandenes Match mit umgehängt.'
                : 'Änderungen werden in die Stammdaten geschrieben. Bei neuem Kürzel wird ein vorhandenes Match mit umgehängt.';
        } else {
            hint.textContent = isArge
                ? 'Die ARGE wird in die Stammdaten geschrieben und erscheint danach im Katalog.'
                : 'Das Fach wird in die Stammdaten geschrieben und erscheint danach im Katalog.';
        }
    }
    if (wrap) wrap.hidden = !isArge;
    if (isArge) {
        fillArgeSubjectPicker(catalogModalMode === 'edit' && row ? row.subjects || [] : []);
    }
    if (codeEl) codeEl.value = catalogModalMode === 'edit' && row ? row.code || '' : '';
    if (nameEl) nameEl.value = catalogModalMode === 'edit' && row ? row.name || '' : '';
    if (saveBtn) {
        saveBtn.innerHTML =
            catalogModalMode === 'edit'
                ? '<i class="bi bi-check-lg"></i>Speichern'
                : '<i class="bi bi-check-lg"></i>Anlegen';
    }
    setCatalogModalError('');
    modal.classList.add('open');
    modal.setAttribute('aria-hidden', 'false');
    updateCatalogNickHint();
    if (codeEl) {
        setTimeout(function () {
            codeEl.focus();
            try {
                codeEl.select();
            } catch (_) {
                /* ignore */
            }
        }, 30);
    }
}

function closeCatalogModal() {
    const modal = document.getElementById('afgCatalogModal');
    if (!modal) return;
    modal.classList.remove('open');
    modal.setAttribute('aria-hidden', 'true');
    catalogModalMode = 'create';
    catalogEditOriginalCode = '';
    setCatalogModalError('');
    updateCatalogNickHint();
}

function codeExistsInKind(kind, code, exceptCode) {
    const list = rowsForKind(kind);
    const key = normCode(code);
    const skip = normCode(exceptCode);
    return list.some(function (r) {
        const c = normCode(r.code);
        if (skip && c === skip) return false;
        return c === key;
    });
}

function remapSubjectCodeInArges(arges, fromCode, toCode) {
    const from = normCode(fromCode);
    const to = normCode(toCode);
    if (!from) return Array.isArray(arges) ? arges.slice() : [];
    return (Array.isArray(arges) ? arges : []).map(function (a) {
        const subjects = Array.isArray(a.subjects) ? a.subjects : [];
        const nextSubj = [];
        const seen = new Set();
        subjects.forEach(function (s) {
            let c = normCode(s);
            if (c === from) c = to;
            if (!c || seen.has(c)) return;
            seen.add(c);
            nextSubj.push(c);
        });
        return Object.assign({}, a, { subjects: nextSubj });
    });
}

function stripSubjectFromArges(arges, subjectCode) {
    const drop = normCode(subjectCode);
    return (Array.isArray(arges) ? arges : []).map(function (a) {
        const subjects = (Array.isArray(a.subjects) ? a.subjects : []).filter(function (s) {
            return normCode(s) !== drop;
        });
        return Object.assign({}, a, { subjects: subjects });
    });
}

function persistCatalogCreate(kind, entry) {
    if (typeof window.ms365TenantSettingsLoad !== 'function' || typeof window.ms365TenantSettingsSave !== 'function') {
        throw new Error('Stammdaten nicht verfügbar (tenant-settings-core.js).');
    }
    const current = window.ms365TenantSettingsLoad() || {};
    const next = Object.assign({}, current);
    if (kind === 'arge') {
        const list = Array.isArray(current.arges) ? current.arges.slice() : [];
        list.push({
            code: entry.code,
            name: entry.name,
            subjects: Array.isArray(entry.subjects) ? entry.subjects.slice() : []
        });
        next.arges = list;
    } else {
        const list = Array.isArray(current.subjects) ? current.subjects.slice() : [];
        list.push({ code: entry.code, name: entry.name });
        next.subjects = list;
    }
    const saved = window.ms365TenantSettingsSave(next);
    dispatchTenantSettingsChanged(saved, 'afg-catalog-create');
    return saved;
}

function persistCatalogUpdate(kind, originalCode, entry) {
    if (typeof window.ms365TenantSettingsLoad !== 'function' || typeof window.ms365TenantSettingsSave !== 'function') {
        throw new Error('Stammdaten nicht verfügbar (tenant-settings-core.js).');
    }
    const from = normCode(originalCode);
    const to = normCode(entry.code);
    const current = window.ms365TenantSettingsLoad() || {};
    const next = Object.assign({}, current);

    if (kind === 'arge') {
        const list = Array.isArray(current.arges) ? current.arges.slice() : [];
        const idx = list.findIndex(function (r) {
            return normCode(r.code) === from;
        });
        if (idx < 0) throw new Error('ARGE nicht mehr gefunden.');
        list[idx] = {
            code: to,
            name: entry.name,
            subjects: Array.isArray(entry.subjects) ? entry.subjects.slice() : []
        };
        next.arges = list;
    } else {
        const list = Array.isArray(current.subjects) ? current.subjects.slice() : [];
        const idx = list.findIndex(function (r) {
            return normCode(r.code) === from;
        });
        if (idx < 0) throw new Error('Fach nicht mehr gefunden.');
        list[idx] = { code: to, name: entry.name };
        next.subjects = list;
        if (from !== to) {
            next.arges = remapSubjectCodeInArges(current.arges, from, to);
        }
    }

    const saved = window.ms365TenantSettingsSave(next);
    if (from !== to) {
        const api = dataV2();
        if (api && typeof api.renameCatalogLink === 'function') {
            try {
                api.renameCatalogLink(kind, from, to);
            } catch (e) {
                throw new Error((e && e.message) || String(e));
            }
        }
    }
    dispatchTenantSettingsChanged(saved, 'afg-catalog-update');
    return saved;
}

function persistCatalogDelete(kind, code) {
    if (typeof window.ms365TenantSettingsLoad !== 'function' || typeof window.ms365TenantSettingsSave !== 'function') {
        throw new Error('Stammdaten nicht verfügbar (tenant-settings-core.js).');
    }
    const key = normCode(code);
    const current = window.ms365TenantSettingsLoad() || {};
    const next = Object.assign({}, current);
    if (kind === 'arge') {
        next.arges = (Array.isArray(current.arges) ? current.arges : []).filter(function (r) {
            return normCode(r.code) !== key;
        });
    } else {
        next.subjects = (Array.isArray(current.subjects) ? current.subjects : []).filter(function (r) {
            return normCode(r.code) !== key;
        });
        next.arges = stripSubjectFromArges(current.arges, key);
    }
    const saved = window.ms365TenantSettingsSave(next);
    const api = dataV2();
    if (api && typeof api.removeCatalogLink === 'function') {
        api.removeCatalogLink(kind, key);
    } else if (api && typeof api.clearCatalogLinkGroup === 'function') {
        api.clearCatalogLinkGroup(kind, key);
    }
    dispatchTenantSettingsChanged(saved, 'afg-catalog-delete');
    return saved;
}

async function offerCreateM365Group(kind, code) {
    const label = kind === 'arge' ? 'ARGE' : 'Fachgruppe';
    const ok = await dlgConfirm(
        label +
            ' „' +
            code +
            '“ ist im Katalog. Jetzt auch eine Microsoft‑365‑Gruppe anlegen und matchen?',
        {
            title: 'M365‑Gruppe anlegen?',
            okText: 'Ja, anlegen',
            cancelText: 'Später'
        }
    );
    if (!ok) return;
    const btn = document.getElementById('slgBtnCreate');
    if (btn) {
        btn.click();
        return;
    }
    toast('Bitte rechts unter „Neue Gruppe anlegen“ auf „Anlegen & matchen“ klicken.');
}

async function submitCatalogModal() {
    const kind = activeKind;
    const codeEl = document.getElementById('afgNewCode');
    const nameEl = document.getElementById('afgNewName');
    const code = normCode(codeEl && codeEl.value);
    const name = normStr(nameEl && nameEl.value);
    const editing = catalogModalMode === 'edit';
    const original = catalogEditOriginalCode;

    if (!code) {
        setCatalogModalError('Bitte ein Kürzel eingeben.');
        if (codeEl) codeEl.focus();
        return;
    }
    if (!name) {
        setCatalogModalError('Bitte einen Namen eingeben.');
        if (nameEl) nameEl.focus();
        return;
    }
    if (codeExistsInKind(kind, code, editing ? original : '')) {
        setCatalogModalError(
            (kind === 'arge' ? 'ARGE' : 'Fach') + ' mit Kürzel „' + code + '“ gibt es bereits.'
        );
        if (codeEl) codeEl.focus();
        return;
    }

    const entry =
        kind === 'arge'
            ? { code: code, name: name, subjects: selectedArgeSubjects() }
            : { code: code, name: name };

    try {
        if (editing) persistCatalogUpdate(kind, original, entry);
        else persistCatalogCreate(kind, entry);
    } catch (e) {
        setCatalogModalError((e && e.message) || String(e));
        return;
    }

    closeCatalogModal();
    readLists();
    setActiveCode(code);
    toast(
        (kind === 'arge' ? 'ARGE' : 'Fachgruppe') +
            ' „' +
            code +
            '“ ' +
            (editing ? 'gespeichert.' : 'angelegt.')
    );
    if (!editing) await offerCreateM365Group(kind, code);
}

async function deleteActiveCatalogEntry() {
    const row = getActiveRow();
    if (!row) {
        toast('Bitte zuerst einen Katalogeintrag wählen.');
        return;
    }
    const kind = activeKind;
    const code = normCode(row.code);
    const label = kind === 'arge' ? 'ARGE' : 'Fachgruppe';
    const link = getCatalogLink(kind, code);
    const matched = !!(link && link.graphGroupId);
    let msg =
        label +
        ' „' +
        (row.name || code) +
        '“ (' +
        code +
        ') aus den Stammdaten entfernen?';
    if (matched) {
        msg +=
            '\n\nDie Verknüpfung zur Microsoft‑365‑Gruppe wird gelöst. Die Gruppe selbst bleibt in Entra erhalten.';
    }
    if (kind === 'subject') {
        msg += '\n\nFalls ARGEs dieses Fach referenzieren, wird die Zuordnung dort entfernt.';
    }
    const ok = await dlgConfirm(msg, {
        title: label + ' löschen?',
        okText: 'Löschen',
        cancelText: 'Abbrechen',
        danger: true
    });
    if (!ok) return;
    try {
        persistCatalogDelete(kind, code);
    } catch (e) {
        toast((e && e.message) || String(e));
        return;
    }
    selectedKeys().delete(code);
    readLists();
    ensureActiveCode();
    renderLeftList();
    applyCreateDefaults();
    refreshMatchUi();
    updateCatalogActionButtons();
    toast(label + ' „' + code + '“ gelöscht.');
}

function renderOwnerPreview() {
    const el = document.getElementById('slgOwnerPreview');
    if (!el) return;
    el.replaceChildren();
    if (!direktion.length) {
        const p = document.createElement('p');
        p.style.margin = '0';
        p.style.color = '#6c757d';
        p.textContent = 'Keine Direktion‑Besitzer in den Stammdaten gefunden.';
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
    if (activeKind === 'arge' && row && Array.isArray(row.subjects) && row.subjects.length) {
        p.textContent = 'Zugeordnete Fächer: ' + row.subjects.join(', ') + '. Mitglieder nach dem Match in Graph pflegen.';
    } else {
        p.textContent = 'Keine Mitgliederliste in den Stammdaten. Nach dem Match Personen über Suche hinzufügen.';
    }
    el.appendChild(p);
}

function applyCreateDefaults(forceNick) {
    const row = getActiveRow();
    const code = row ? row.code : activeCode;
    const name = row && row.name ? row.name : code;
    const dn = document.getElementById('slgNewDisplayName');
    const nn = document.getElementById('slgNewMailNick');
    const desc = document.getElementById('slgNewDescription');
    const hint = document.getElementById('slgNewMailNickHint');
    const suggested = deriveNick(activeKind, code);
    const titleName = name || code || '';
    if (dn) dn.value = titleName;
    if (nn && (forceNick || !createNickDirty)) {
        nn.value = suggested;
        createNickDirty = false;
    }
    if (hint) {
        hint.textContent =
            'Frei editierbar. Standard' +
            (mailPrefix(activeKind) ? ' aus Präfix + Kürzel' : ' = Kürzel (kein Präfix)') +
            ': ' +
            (suggested || '–') +
            '. Leer oder „Standard“ setzt diesen Vorschlag beim Anlegen.';
    }
    if (desc) {
        desc.value =
            titleName +
            (activeKind === 'arge' ? ' (MS365-Schul-Tools / ARGE)' : ' (MS365-Schul-Tools / Fachgruppe)');
    }
    const search = document.getElementById('slgGroupSearch');
    if (search && !normStr(search.value)) search.value = name || code || '';
}

function refreshMatchUi() {
    const gid = getActiveGroupId();
    const row = getActiveRow();
    const title = document.getElementById('slgDetailTitle');
    if (title) {
        const prefix = activeKind === 'arge' ? 'ARGE' : 'Fach';
        title.textContent = row ? prefix + ' ' + (row.name || row.code) : prefix;
    }
    live().resetCaches();
    live().setMatchedMode(!!gid);
    live().fillForm(gid ? { id: gid } : null);
    renderOwnerPreview();
    renderMemberPreview();
}

function revealCatalogRow(host, row) {
    if (!host || !row || typeof row.getBoundingClientRect !== 'function') return;
    const hostRect = host.getBoundingClientRect();
    const rowRect = row.getBoundingClientRect();
    if (rowRect.top < hostRect.top) host.scrollTop -= hostRect.top - rowRect.top;
    else if (rowRect.bottom > hostRect.bottom) host.scrollTop += rowRect.bottom - hostRect.bottom;
}

function renderLeftList(opts) {
    const host = document.getElementById('afgListItems');
    const summary = document.getElementById('afgListSummary');
    const empty = document.getElementById('afgEmptyHint');
    const wrap = document.getElementById('afgDetailWrap');
    if (!host) return;
    const resetScroll = !!(opts && opts.resetScroll);
    const revealActive = !!(opts && opts.revealActive);
    const keepScroll = resetScroll ? 0 : host.scrollTop;
    function finishListScroll() {
        host.scrollTop = keepScroll;
        if (!revealActive) return;
        revealCatalogRow(host, host.querySelector('button[aria-current="true"]'));
    }
    host.replaceChildren();
    const q = listFilter.toLowerCase();
    const all = rowsForKind(activeKind);
    const list = all.filter(function (row) {
        if (!q) return true;
        const hay = (row.code + ' ' + (row.name || '')).toLowerCase();
        return hay.indexOf(q) !== -1;
    });
    let matchedN = 0;
    all.forEach(function (row) {
        const link = getCatalogLink(activeKind, row.code);
        if (link && link.graphGroupId) matchedN++;
    });
    const masterTitle = document.getElementById('afgMasterTitle');
    if (masterTitle) {
        masterTitle.textContent = activeKind === 'arge' ? 'ARGE' : 'Fachgruppen';
    }
    if (summary) {
        summary.textContent =
            String(all.length) +
            (activeKind === 'arge' ? ' Gruppen' : ' Gruppen') +
            ' · ' +
            String(matchedN) +
            ' verknüpft' +
            (q ? ' · Treffer: ' + String(list.length) : '');
    }
    const hasRows = all.length > 0;
    if (empty) empty.style.display = hasRows ? 'none' : '';
    if (wrap) wrap.style.display = hasRows ? '' : 'none';

    if (!list.length) {
        const li = document.createElement('li');
        const p = document.createElement('p');
        p.className = 'muted';
        p.style.margin = '10px 12px';
        p.textContent = hasRows ? 'Keine Treffer im Filter.' : 'Liste ist leer.';
        li.appendChild(p);
        host.appendChild(li);
        updateBulkCount();
        updateCatalogActionButtons();
        finishListScroll();
        return;
    }

    list.forEach(function (row) {
        const link = getCatalogLink(activeKind, row.code);
        const gid = link && link.graphGroupId ? String(link.graphGroupId) : '';
        const key = normCode(row.code);
        const li = document.createElement('li');
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.setAttribute('data-afg-code', row.code);
        if (normCode(row.code) === normCode(activeCode)) btn.setAttribute('aria-current', 'true');
        const main = document.createElement('span');
        main.className = 'slg-side-main';
        const t = document.createElement('span');
        t.className = 'slg-side-title';
        t.textContent = (row.name || row.code) + (row.name && row.code ? ' (' + row.code + ')' : '');
        const meta = document.createElement('span');
        meta.className = 'muted slg-side-meta';
        const badge = document.createElement('span');
        badge.className = 'jg-match-badge gd-sync-badge ' + (gid ? 'is-ok gd-sync-badge--ok' : 'is-warn gd-sync-badge--warn');
        const ico = document.createElement('i');
        ico.className = gid ? 'bi bi-check-circle-fill' : 'bi bi-exclamation-circle';
        ico.setAttribute('aria-hidden', 'true');
        badge.appendChild(ico);
        badge.appendChild(document.createTextNode(gid ? 'Verknüpft' : 'Nicht verknüpft'));
        meta.appendChild(badge);
        btn.classList.add(gid ? 'is-matched' : 'is-unmatched');
        if (window.ms365GroupPhotoThumb && typeof window.ms365GroupPhotoThumb.createThumb === 'function') {
            btn.insertBefore(
                window.ms365GroupPhotoThumb.createThumb({
                    groupId: gid,
                    displayName: (row.name || row.code || '').trim(),
                    size: 'list'
                }),
                btn.firstChild
            );
        }
        main.appendChild(t);
        main.appendChild(meta);
        btn.appendChild(main);
        if (selectionMode) {
            const pick = document.createElement('label');
            pick.className = 'afg-pick';
            pick.title = 'Für Mehrfachauswahl markieren';
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.setAttribute('data-afg-pick', key);
            cb.checked = selectedKeys().has(key);
            cb.disabled = false;
            cb.addEventListener('click', function (ev) {
                ev.stopPropagation();
            });
            cb.addEventListener('change', function () {
                if (cb.checked) selectedKeys().add(key);
                else selectedKeys().delete(key);
                updateBulkCount();
            });
            pick.appendChild(cb);
            li.appendChild(pick);
        }
        li.appendChild(btn);
        host.appendChild(li);
    });
    host.classList.toggle('tree--selecting', selectionMode);
    if (window.ms365GroupPhotoThumb && typeof window.ms365GroupPhotoThumb.hydrate === 'function') {
        window.ms365GroupPhotoThumb.hydrate(host);
    }
    updateBulkCount();
    updateSelectionModeUi();
    updateCatalogActionButtons();
    finishListScroll();
}

function updateSelectionModeUi() {
    const btn = document.getElementById('afgBtnSelectionMode');
    if (btn) {
        btn.setAttribute('aria-pressed', selectionMode ? 'true' : 'false');
        btn.innerHTML = selectionMode
            ? '<i class="bi bi-x-lg"></i>Auswahl beenden'
            : '<i class="bi bi-check2-square"></i>Auswählen';
    }
    const host = document.getElementById('afgListItems');
    if (host) host.classList.toggle('tree--selecting', selectionMode);
}

function setSelectionMode(on) {
    selectionMode = !!on;
    if (!selectionMode) {
        selectedKeysByKind[activeKind] = new Set();
        showBulkOwnerPanel(false);
    }
    renderLeftList();
}

function ensureActiveCode() {
    const list = rowsForKind(activeKind);
    if (!list.length) {
        activeCode = '';
        return;
    }
    const has = list.some(function (r) {
        return normCode(r.code) === normCode(activeCode);
    });
    if (!has) activeCode = list[0].code;
}

function setActiveKind(kind) {
    activeKind = kind === 'arge' ? 'arge' : 'subject';
    createNickDirty = false;
    document.querySelectorAll('[data-afg-kind]').forEach(function (b) {
        b.setAttribute('aria-pressed', b.getAttribute('data-afg-kind') === activeKind ? 'true' : 'false');
    });
    ensureActiveCode();
    showBulkOwnerPanel(false);
    setBulkStatus('', false);
    const search = document.getElementById('slgGroupSearch');
    if (search) search.value = '';
    gd().clearSearchResults();
    renderLeftList({ resetScroll: true, revealActive: true });
    applyCreateDefaults(true);
    syncMailPrefixUi();
    gd().setTab('general');
    refreshMatchUi();
    if (getActiveGroupId()) live().loadGroup({ silent: true });
}

function setActiveCode(code) {
    const next = normCode(code);
    if (next && next === normCode(activeCode)) return;
    const apply = function () {
        activeCode = next;
        createNickDirty = false;
        const search = document.getElementById('slgGroupSearch');
        if (search) search.value = '';
        gd().clearSearchResults();
        renderLeftList({ revealActive: true });
        applyCreateDefaults(true);
        syncMailPrefixUi();
        gd().setTab('general');
        refreshMatchUi();
        if (getActiveGroupId()) live().loadGroup({ silent: true });
    };
    if (live() && typeof live().confirmIfDirty === 'function') {
        const label = (function () {
            const rows = rowsForKind(activeKind);
            for (let i = 0; i < rows.length; i++) {
                if (normCode(rows[i].code) === normCode(activeCode)) {
                    return rows[i].name || rows[i].code || activeCode;
                }
            }
            return activeCode || 'diese Gruppe';
        })();
        Promise.resolve(live().confirmIfDirty(label)).then(function (ok) {
            if (ok) apply();
            else renderLeftList({ revealActive: true });
        });
        return;
    }
    apply();
}

function persistMatchFor(kind, code, g, mode) {
    const k = kind === 'arge' ? 'arge' : 'subject';
    const c = normCode(code);
    const list = rowsForKind(k);
    let row = null;
    for (let i = 0; i < list.length; i++) {
        if (normCode(list[i].code) === c) {
            row = list[i];
            break;
        }
    }
    upsertCatalogLink({
        kind: k,
        code: c,
        graphGroupId: g && g.id ? String(g.id) : '',
        displayName: (g && g.displayName) || (row && row.name) || '',
        mailNickname: (g && g.mailNickname) || '',
        mode: mode
    });
}

function persistMatch(g, mode) {
    persistMatchFor(activeKind, activeCode, g, mode);
    renderLeftList();
}

function getAfgBulkCreateTeam() {
    const el = document.getElementById('afgBulkCreateTeam');
    return !!(el && el.checked);
}

function ownersForCatalogRow(row) {
    return collectAfgCreateOwners(row, direktion);
}

function displayNameForCreate(row, kind) {
    void kind;
    return displayNameForAfgCreate(row);
}

function descriptionForCreate(row, kind) {
    const title = displayNameForAfgCreate(row);
    return title + (kind === 'arge' ? ' (MS365-Schul-Tools / ARGE)' : ' (MS365-Schul-Tools / Fachgruppe)');
}

function persistUnmatchFor(kind, code) {
    const api = dataV2();
    const k = kind === 'arge' ? 'arge' : 'subject';
    const c = normCode(code);
    if (api && typeof api.clearCatalogLinkGroup === 'function') {
        api.clearCatalogLinkGroup(k, c);
    } else {
        upsertCatalogLink({
            kind: k,
            code: c,
            graphGroupId: '',
            displayName: '',
            mailNickname: '',
            mode: ''
        });
    }
    selectedKeysByKind[k] && selectedKeysByKind[k].delete(c);
}

function persistUnmatch() {
    persistUnmatchFor(activeKind, activeCode);
    renderLeftList();
}

function sleep(ms) {
    return new Promise(function (r) {
        setTimeout(r, ms);
    });
}

function setBulkStatus(text, show) {
    const el = document.getElementById('afgBulkStatus');
    if (!el) return;
    if (show === false || !text) {
        el.hidden = true;
        el.textContent = '';
        return;
    }
    el.hidden = false;
    el.textContent = text;
}

function pruneSelection() {
    const keep = new Set();
    const list = rowsForKind(activeKind);
    const keys = selectedKeys();
    keys.forEach(function (key) {
        for (let i = 0; i < list.length; i++) {
            if (normCode(list[i].code) === key) {
                keep.add(key);
                break;
            }
        }
    });
    selectedKeysByKind[activeKind] = keep;
}

function collectSelectedMatched() {
    pruneSelection();
    const out = [];
    const list = rowsForKind(activeKind);
    selectedKeys().forEach(function (key) {
        let row = null;
        for (let i = 0; i < list.length; i++) {
            if (normCode(list[i].code) === key) {
                row = list[i];
                break;
            }
        }
        if (!row) return;
        const link = getCatalogLink(activeKind, row.code);
        const id = link && link.graphGroupId ? String(link.graphGroupId).trim() : '';
        if (!id) return;
        out.push({
            key: key,
            row: row,
            link: link,
            id: id,
            name: normStr(row.name) || normStr(row.code) || id
        });
    });
    return out;
}

function collectSelectedUnmatched() {
    pruneSelection();
    const out = [];
    const list = rowsForKind(activeKind);
    selectedKeys().forEach(function (key) {
        let row = null;
        for (let i = 0; i < list.length; i++) {
            if (normCode(list[i].code) === key) {
                row = list[i];
                break;
            }
        }
        if (!row) return;
        const link = getCatalogLink(activeKind, row.code);
        if (link && link.graphGroupId) return;
        const nick = deriveNick(activeKind, row.code);
        out.push({
            key: key,
            row: row,
            code: normCode(row.code),
            nick: nick,
            name: normStr(row.name) || normStr(row.code) || key
        });
    });
    return out;
}

function updateBulkCount() {
    pruneSelection();
    const n = selectedKeys().size;
    const el = document.getElementById('afgBulkCount');
    const bar = document.getElementById('afgSelectionBar');
    if (el) {
        el.textContent = n === 1 ? '1 ausgewählt' : String(n) + ' ausgewählt';
    }
    if (bar) bar.classList.toggle('is-active', n > 0);
    if (n === 0) showBulkOwnerPanel(false);
}

function visibleMatchedRows() {
    const q = listFilter.toLowerCase();
    return rowsForKind(activeKind).filter(function (row) {
        const link = getCatalogLink(activeKind, row.code);
        if (!link || !link.graphGroupId) return false;
        if (!q) return true;
        const hay = (row.code + ' ' + (row.name || '')).toLowerCase();
        return hay.indexOf(q) !== -1;
    });
}

function kindLabel(n) {
    if (activeKind === 'arge') return n === 1 ? 'ARGE' : 'ARGEs';
    return n === 1 ? 'Fachgruppe' : 'Fachgruppen';
}

function visibleUnmatchedRows() {
    const q = listFilter.toLowerCase();
    return rowsForKind(activeKind).filter(function (row) {
        const link = getCatalogLink(activeKind, row.code);
        if (link && link.graphGroupId) return false;
        if (!q) return true;
        const hay = (row.code + ' ' + (row.name || '')).toLowerCase();
        return hay.indexOf(q) !== -1;
    });
}

function selectVisibleMatched() {
    selectionMode = true;
    visibleMatchedRows().forEach(function (row) {
        selectedKeys().add(normCode(row.code));
    });
    renderLeftList();
    const n = collectSelectedMatched().length;
    toast(
        n
            ? String(n) + ' verknüpfte ' + kindLabel(n) + ' angekreuzt.'
            : 'Keine verknüpften Gruppen in der aktuellen Liste.'
    );
}

function selectVisibleUnmatched() {
    selectionMode = true;
    const rows = visibleUnmatchedRows();
    rows.forEach(function (row) {
        selectedKeys().add(normCode(row.code));
    });
    renderLeftList();
    toast(
        rows.length
            ? String(rows.length) + ' nicht verknüpfte ' + kindLabel(rows.length) + ' angekreuzt.'
            : 'Keine nicht verknüpften Einträge in der aktuellen Liste.'
    );
}

function clearSelection() {
    selectedKeysByKind[activeKind] = new Set();
    showBulkOwnerPanel(false);
    renderLeftList();
}

function openAliasPrefixSettings() {
    const ov = document.getElementById('afgMasterOverflow');
    if (ov) ov.open = false;
    const panel = document.getElementById('afgMailPrefixPanel');
    if (!panel) return;
    panel.hidden = false;
    panel.open = true;
    const inp = document.getElementById('afgMailPrefix');
    if (inp) {
        try {
            inp.focus();
        } catch {
            /* ignore */
        }
    }
}

function showBulkOwnerPanel(show) {
    const panel = document.getElementById('afgBulkOwnerPanel');
    if (!panel) return;
    panel.hidden = !show;
    if (show) {
        const inp = document.getElementById('afgBulkOwnerSearch');
        if (inp) inp.focus();
    }
}

function fillBulkOwnerSelect(users) {
    const sel = document.getElementById('afgBulkOwnerResults');
    if (!sel) return;
    sel.replaceChildren();
    if (!users || !users.length) {
        const opt = document.createElement('option');
        opt.value = '';
        opt.textContent = '(keine Treffer)';
        sel.appendChild(opt);
        return;
    }
    users.forEach(function (u) {
        const opt = document.createElement('option');
        opt.value = u.id || '';
        opt.textContent = gug().personLabel(u) || (u.id ? String(u.id) : '');
        sel.appendChild(opt);
    });
}

async function runBulkOwnerSearch() {
    const inp = document.getElementById('afgBulkOwnerSearch');
    const q = inp ? String(inp.value || '').trim() : '';
    if (!q) {
        toast('Bitte einen Namen oder eine E‑Mail eingeben.');
        return;
    }
    const btn = document.getElementById('afgBulkOwnerSearchBtn');
    if (btn) btn.disabled = true;
    try {
        const token = await gug().getGraphToken();
        const users = await gug().searchUsers(token, q);
        fillBulkOwnerSelect(users);
        toast('Suche: ' + users.length + ' Treffer.');
    } catch (e) {
        toast('Suche: ' + (e.message || e));
    } finally {
        if (btn) btn.disabled = false;
    }
}

async function runBulkSetOwner() {
    const items = collectSelectedMatched();
    if (!items.length) {
        toast('Bitte zuerst gematchte ' + kindLabel(2) + ' ankreuzen.');
        return;
    }
    const sel = document.getElementById('afgBulkOwnerResults');
    const userId = sel && sel.value ? String(sel.value).trim() : '';
    if (!userId) {
        toast('Bitte zuerst einen Benutzer suchen und auswählen.');
        showBulkOwnerPanel(true);
        return;
    }
    const label = sel.options[sel.selectedIndex]
        ? String(sel.options[sel.selectedIndex].textContent || '').trim()
        : userId;
    if (
        !(await dlgConfirm(
            '„' +
                label +
                '“ als Besitzer zu ' +
                String(items.length) +
                ' Gruppe(n) hinzufügen?\n\nBestehende Besitzer bleiben erhalten.',
            { title: 'Besitzer setzen', okText: 'Hinzufügen' }
        ))
    ) {
        return;
    }
    const applyBtn = document.getElementById('afgBulkOwnerApply');
    if (applyBtn) applyBtn.disabled = true;
    let ok = 0;
    let skip = 0;
    let fail = 0;
    const lines = [];
    setBulkStatus('Besitzer wird gesetzt …');
    try {
        const token = await gug().getGraphToken();
        for (let i = 0; i < items.length; i++) {
            const it = items[i];
            try {
                await gug().addOwnerWithMemberFallback(token, it.id, userId);
                ok++;
                lines.push('OK  ' + it.name);
            } catch (e) {
                if (gug().isDuplicateMemberError(e)) {
                    skip++;
                    lines.push('schon Besitzer  ' + it.name);
                } else {
                    fail++;
                    lines.push('Fehler  ' + it.name + ': ' + (e.message || e));
                }
            }
            if ((i + 1) % 6 === 0) await sleep(120);
        }
        setBulkStatus(lines.join('\n'));
        toast('Besitzer: neu ' + ok + ', bereits vorhanden ' + skip + ', Fehler ' + fail + '.');
        if (getActiveGroupId()) {
            try {
                live().invalidateMembership();
                if (gd().getActiveTab() === 'owners') await live().loadOwners();
            } catch {
                /* ignore */
            }
        }
    } catch (e) {
        setBulkStatus('Abbruch: ' + (e.message || e));
        toast('Besitzer setzen: ' + (e.message || e));
    } finally {
        if (applyBtn) applyBtn.disabled = false;
    }
}

async function runBulkCreateAndMatch() {
    const items = collectSelectedUnmatched();
    if (!items.length) {
        toast('Bitte zuerst ungematchte ' + kindLabel(2) + ' ankreuzen (rote Badges).');
        return;
    }
    const createTeam = getAfgBulkCreateTeam();
    const preview = formatAfgBulkCreatePreviewLines(items, 12);
    if (
        !(await dlgConfirm(
            buildAfgBulkCreateConfirmMessage({
                kindLabel: kindLabel(items.length),
                count: items.length,
                previewLines: preview,
                createTeam: createTeam
            }),
            { title: 'Gruppen anlegen & matchen', okText: 'Anlegen' }
        ))
    ) {
        return;
    }
    const btn = document.getElementById('afgBtnBulkCreate');
    if (btn) btn.disabled = true;
    let ok = 0;
    let fail = 0;
    const lines = [];
    setBulkStatus('Gruppen werden angelegt …');
    try {
        const token = await gug().getGraphToken();
        for (let i = 0; i < items.length; i++) {
            const it = items[i];
            const nick = it.nick || deriveNick(activeKind, it.code);
            if (!nick) {
                fail++;
                lines.push('Fehler  ' + it.code + ': kein Alias ableitbar (Präfix/Kürzel prüfen).');
                continue;
            }
            const displayName = displayNameForCreate(it.row, activeKind);
            const desc = descriptionForCreate(it.row, activeKind);
            try {
                const g = await gug().createUnifiedGroup(token, displayName, nick, desc);
                const owners = ownersForCatalogRow(it.row);
                if (owners.length) {
                    try {
                        await gug().ensureOwners(token, g.id, owners);
                    } catch {
                        /* optional */
                    }
                }
                if (createTeam && typeof gug().provisionTeamForGroup === 'function') {
                    try {
                        await gug().provisionTeamForGroup(token, g.id);
                    } catch (teamErr) {
                        const msg = String((teamErr && teamErr.message) || teamErr || '');
                        if (/\b409\b/.test(msg) || /Conflict|already exists|already provisioned/i.test(msg)) {
                            /* Team bereits vorhanden */
                        } else {
                            lines.push('Team fehlgeschlagen  ' + it.code + ': ' + msg);
                        }
                    }
                }
                persistMatchFor(activeKind, it.code, g, 'created');
                selectedKeys().delete(it.key);
                ok++;
                lines.push(
                    'OK  ' + it.code + '  →  ' + nick + (createTeam ? ' (+ Team)' : '')
                );
                if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
                    window.ms365ActionLog.append({
                        tool: 'afg',
                        action: 'createAndMatch',
                        target: nick,
                        summary: displayName,
                        result: 'ok'
                    });
                }
            } catch (e) {
                fail++;
                lines.push('Fehler  ' + it.code + ': ' + (e.message || e));
            }
            if ((i + 1) % 4 === 0) await sleep(200);
        }
        renderLeftList();
        if (items.some(function (it) {
            return normCode(it.code) === normCode(activeCode);
        })) {
            refreshMatchUi();
            if (getActiveGroupId()) {
                try {
                    live().invalidateMembership();
                } catch {
                    /* ignore */
                }
                try {
                    live().loadGroup({ silent: true });
                } catch {
                    /* ignore */
                }
            }
        }
        setBulkStatus(lines.join('\n'));
        toast('Anlegen: ' + ok + ' OK, ' + fail + ' Fehler.');
    } catch (e) {
        renderLeftList();
        setBulkStatus('Abbruch: ' + (e.message || e));
        toast('Anlegen: ' + (e.message || e));
    } finally {
        if (btn) btn.disabled = false;
    }
}

async function runBulkMatchExisting() {
    const items = collectSelectedUnmatched();
    if (!items.length) {
        toast('Bitte zuerst ungematchte ' + kindLabel(2) + ' ankreuzen (rote Badges).');
        return;
    }
    const preview = formatAfgBulkCreatePreviewLines(items, 12);
    if (
        !(await dlgConfirm(
            buildAfgBulkMatchExistingConfirmMessage({
                kindLabel: kindLabel(items.length),
                count: items.length,
                previewLines: preview
            }),
            { title: 'Bestehende matchen', okText: 'Abgleichen' }
        ))
    ) {
        return;
    }
    const btn = document.getElementById('afgBtnBulkMatchExisting');
    if (btn) btn.disabled = true;
    let ok = 0;
    let miss = 0;
    let fail = 0;
    const lines = [];
    setBulkStatus('Bestehende Gruppen werden gesucht …');
    try {
        const token = await gug().getGraphToken();
        for (let i = 0; i < items.length; i++) {
            const it = items[i];
            const nick = it.nick || deriveNick(activeKind, it.code);
            if (!nick) {
                fail++;
                lines.push('Fehler  ' + it.code + ': kein Alias ableitbar (Präfix/Kürzel prüfen).');
                continue;
            }
            try {
                const list = await gug().searchUnifiedGroups(token, nick);
                const g = pickAfgGroupByNickname(list, nick);
                if (!g || !g.id) {
                    miss++;
                    lines.push('kein Treffer  ' + it.code + '  →  ' + nick);
                    continue;
                }
                persistMatchFor(activeKind, it.code, g, 'matched');
                selectedKeys().delete(it.key);
                ok++;
                lines.push(
                    'OK  ' +
                        it.code +
                        '  →  ' +
                        (normStr(g.displayName) || nick) +
                        ' (' +
                        nick +
                        ')'
                );
                if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
                    window.ms365ActionLog.append({
                        tool: 'afg',
                        action: 'matchExisting',
                        target: nick,
                        summary: normStr(g.displayName) || nick,
                        result: 'ok'
                    });
                }
            } catch (e) {
                fail++;
                lines.push('Fehler  ' + it.code + ': ' + (e.message || e));
            }
            if ((i + 1) % 4 === 0) await sleep(200);
        }
        renderLeftList();
        if (
            items.some(function (it) {
                return normCode(it.code) === normCode(activeCode);
            })
        ) {
            refreshMatchUi();
            if (getActiveGroupId()) {
                try {
                    live().invalidateMembership();
                } catch {
                    /* ignore */
                }
                try {
                    live().loadGroup({ silent: true });
                } catch {
                    /* ignore */
                }
            }
        }
        setBulkStatus(lines.join('\n'));
        toast('Abgleich: ' + ok + ' gematcht, ' + miss + ' ohne Treffer, ' + fail + ' Fehler.');
    } catch (e) {
        renderLeftList();
        setBulkStatus('Abbruch: ' + (e.message || e));
        toast('Abgleich: ' + (e.message || e));
    } finally {
        if (btn) btn.disabled = false;
    }
}

async function runBulkStripDisplayPrefix() {
    const items = collectSelectedMatched();
    if (!items.length) {
        toast('Bitte zuerst gematchte ' + kindLabel(2) + ' ankreuzen.');
        return;
    }
    const btn = document.getElementById('afgBtnBulkRename');
    if (btn) btn.disabled = true;
    let plan = [];
    let skip = 0;
    let fail = 0;
    const lines = [];
    setBulkStatus('Anzeigenamen werden geprüft …');
    try {
        const token = await gug().getGraphToken();
        for (let i = 0; i < items.length; i++) {
            const it = items[i];
            try {
                const g = await gug().fetchGroup(token, it.id);
                const rawDn = normStr(g && g.displayName);
                const rawDesc = normStr(g && g.description);
                const nextDn = stripAfgDisplayNamePrefix(rawDn);
                const nextDesc = afgNeedsDisplayNameCleanup(rawDesc)
                    ? stripAfgDisplayNamePrefix(rawDesc)
                    : rawDesc;
                if (!afgNeedsDisplayNameCleanup(rawDn) && nextDesc === rawDesc) {
                    skip++;
                    lines.push('bereits ok  ' + (rawDn || it.name));
                    continue;
                }
                plan.push({
                    key: it.key,
                    id: it.id,
                    from: rawDn,
                    to: nextDn || rawDn,
                    description: nextDesc !== rawDesc ? nextDesc : undefined,
                    code: it.key
                });
            } catch (e) {
                fail++;
                lines.push('Fehler  ' + it.name + ': ' + (e.message || e));
            }
            if ((i + 1) % 6 === 0) await sleep(120);
        }
        if (!plan.length) {
            setBulkStatus(lines.join('\n') || 'Keine Namen mit Präfix „Arbeitsgruppe “ / „Fach “ gefunden.');
            toast(
                fail
                    ? 'Namen prüfen: nichts zu ändern, ' + fail + ' Fehler.'
                    : 'Keine Anzeigenamen mit „Arbeitsgruppe “ / „Fach “ – nichts zu tun.'
            );
            return;
        }
        const preview =
            plan
                .slice(0, 12)
                .map(function (p) {
                    return p.from + '  →  ' + p.to;
                })
                .join('\n') + (plan.length > 12 ? '\n…' : '');
        if (
            !(await dlgConfirm(
                buildAfgBulkRenameConfirmMessage({
                    count: plan.length,
                    previewLines: preview
                }),
                { title: 'Namen bereinigen', okText: 'Umbenennen' }
            ))
        ) {
            setBulkStatus(lines.concat(['Abbruch vor dem Umbenennen.']).join('\n'));
            return;
        }
        setBulkStatus('Anzeigenamen werden aktualisiert …');
        let ok = 0;
        for (let i = 0; i < plan.length; i++) {
            const p = plan[i];
            try {
                if (typeof gug().patchGroupDisplayName !== 'function') {
                    throw new Error('patchGroupDisplayName fehlt.');
                }
                await gug().patchGroupDisplayName(token, p.id, p.to, p.description);
                const link = getCatalogLink(activeKind, p.code);
                if (link) {
                    persistMatchFor(activeKind, p.code, {
                        id: p.id,
                        displayName: p.to,
                        mailNickname: link.mailNickname || ''
                    }, link.mode || 'matched');
                }
                ok++;
                lines.push('OK  ' + p.from + '  →  ' + p.to);
                if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
                    window.ms365ActionLog.append({
                        tool: 'afg',
                        action: 'stripDisplayPrefix',
                        target: p.id,
                        summary: p.from + ' → ' + p.to,
                        result: 'ok'
                    });
                }
            } catch (e) {
                fail++;
                lines.push('Fehler  ' + p.from + ': ' + (e.message || e));
            }
            if ((i + 1) % 4 === 0) await sleep(200);
        }
        renderLeftList();
        if (plan.some(function (p) {
            return normCode(p.code) === normCode(activeCode);
        })) {
            refreshMatchUi();
            if (getActiveGroupId()) {
                try {
                    live().loadGroup({ silent: true });
                } catch {
                    /* ignore */
                }
            }
        }
        setBulkStatus(lines.join('\n'));
        toast('Namen: ' + ok + ' geändert, ' + skip + ' bereits ok, ' + fail + ' Fehler.');
    } catch (e) {
        renderLeftList();
        setBulkStatus('Abbruch: ' + (e.message || e));
        toast('Namen bereinigen: ' + (e.message || e));
    } finally {
        if (btn) btn.disabled = false;
    }
}

async function runBulkDelete() {
    const items = collectSelectedMatched();
    if (!items.length) {
        toast('Bitte zuerst gematchte ' + kindLabel(2) + ' ankreuzen.');
        return;
    }
    const preview =
        items
            .slice(0, 12)
            .map(function (it) {
                return it.name;
            })
            .join('\n') + (items.length > 12 ? '\n…' : '');
    if (
        !(await dlgConfirm(
            String(items.length) +
                ' Microsoft‑365‑Gruppe(n) wirklich löschen?\n\n' +
                preview +
                '\n\nDie Gruppen verschwinden in Entra/Teams. Das lokale Match wird gelöst. Das lässt sich nicht rückgängig machen.',
            { title: 'Gruppen löschen', okText: 'Löschen', danger: true }
        ))
    ) {
        return;
    }
    const delBtn = document.getElementById('afgBtnBulkDelete');
    if (delBtn) delBtn.disabled = true;
    let ok = 0;
    let fail = 0;
    const lines = [];
    const deletedKeys = [];
    setBulkStatus('Gruppen werden gelöscht …');
    try {
        const token = await gug().getGraphToken();
        for (let i = 0; i < items.length; i++) {
            const it = items[i];
            try {
                if (typeof gug().deleteUnifiedGroup !== 'function') {
                    throw new Error('deleteUnifiedGroup fehlt.');
                }
                await gug().deleteUnifiedGroup(token, it.id);
                persistUnmatchFor(activeKind, it.key);
                deletedKeys.push(it.key);
                ok++;
                lines.push('gelöscht  ' + it.name);
            } catch (e) {
                const msg = String((e && e.message) || e || '');
                if (/\b404\b/.test(msg) || /Request_ResourceNotFound/i.test(msg)) {
                    persistUnmatchFor(activeKind, it.key);
                    deletedKeys.push(it.key);
                    ok++;
                    lines.push('bereits weg  ' + it.name);
                } else {
                    fail++;
                    lines.push('Fehler  ' + it.name + ': ' + msg);
                }
            }
            if ((i + 1) % 4 === 0) await sleep(200);
        }
        renderLeftList();
        if (deletedKeys.indexOf(normCode(activeCode)) >= 0) {
            refreshMatchUi();
        }
        setBulkStatus(lines.join('\n'));
        toast('Löschen: ' + ok + ' erledigt, ' + fail + ' Fehler.');
    } catch (e) {
        renderLeftList();
        setBulkStatus('Abbruch: ' + (e.message || e));
        toast('Löschen: ' + (e.message || e));
    } finally {
        if (delBtn) delBtn.disabled = false;
    }
}

function onClick(id, fn) {
    const el = document.getElementById(id);
    if (el) el.addEventListener('click', fn);
}

function mountDetail() {
    gd().mount('#groupDetailHost', {
        title: 'Fachgruppe',
        searchPlaceholder: 'z. B. ARGE Sprachen oder fach-d',
        unmatchedCreateHint:
            'Legt eine Microsoft 365‑Gruppe (Unified) an und verknüpft sie mit diesem Katalogeintrag. Unique ID / Alias können Sie frei setzen – sonst gilt Präfix + Kürzel bzw. nur das Kürzel, wenn das Präfix leer ist. Optional auch als Team bereitstellen.',
        membersUnmatchedHint:
            'In den Stammdaten gibt es keine Mitgliederliste für Fach/ARGE. Nach der Verknüpfung können Sie Mitglieder live in Graph pflegen.',
        membersUnmatchedTitle: 'Hinweis aus den Stammdaten',
        membersMatchedHint:
            'Live aus Microsoft Graph. Es gibt keinen automatischen Listen‑Sync – Mitglieder hier suchen, hinzufügen oder entfernen.',
        emptyHintHtml:
            'Keine Einträge in dieser Liste. Legen Sie einen Eintrag über <strong>Neue Gruppe</strong> an oder pflegen Sie Fächer bzw. ARGE unter <a href="../tenant.html">Stammdaten</a>.',
        features: {
            syncMembers: false,
            emptyHint: true
        },
        ids: { emptyHint: 'afgEmptyHint', wrap: 'afgDetailWrap' },
        live: {
            toast: toast,
            dlgConfirm: dlgConfirm,
            getGroupId: getActiveGroupId,
            ensureDirektionOwners: function (token, gid) {
                if (!direktion.length) throw new Error('Keine Direktion‑Adressen in den Stammdaten.');
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
            suggestMailNickname: function () {
                const row = getActiveRow();
                const code = row ? row.code : activeCode;
                return deriveNick(activeKind, code);
            },
            canSearch: function () {
                return activeCode
                    ? { ok: true }
                    : { ok: false, message: 'Bitte zuerst einen Katalogeintrag wählen.' };
            },
            canCreate: function () {
                return activeCode
                    ? { ok: true }
                    : { ok: false, message: 'Bitte zuerst einen Katalogeintrag wählen.' };
            },
            ensureOwners: function (token, gid) {
                return gug().ensureOwners(token, gid, ownersForCatalogRow(getActiveRow()));
            }
        },
        onTabUnmatched: function (tab) {
            if (tab === 'owners') renderOwnerPreview();
            if (tab === 'members') renderMemberPreview();
        }
    });
}

function wire() {
    document.querySelectorAll('[data-afg-kind]').forEach(function (b) {
        b.addEventListener('click', function () {
            setActiveKind(b.getAttribute('data-afg-kind') === 'arge' ? 'arge' : 'subject');
            if (getActiveGroupId()) live().loadGroup({ silent: true });
        });
    });
    const listHost = document.getElementById('afgListItems');
    if (listHost) {
        listHost.addEventListener('click', function (ev) {
            const t = ev.target;
            const item = t && t.closest ? t.closest('button[data-afg-code]') : null;
            if (!item) return;
            setActiveCode(item.getAttribute('data-afg-code') || '');
        });
    }
    const filter = document.getElementById('afgListFilter');
    if (filter) {
        filter.addEventListener('input', function () {
            listFilter = String(filter.value || '').trim();
            renderLeftList();
        });
    }
    onClick('slgBtnReloadLists', function () {
        readLists();
        ensureActiveCode();
        renderLeftList();
        applyCreateDefaults();
        refreshMatchUi();
        toast('Listen neu eingelesen.');
    });
    onClick('afgBtnAddCatalog', function () {
        openCatalogModal('create');
    });
    onClick('afgBtnEditCatalog', function () {
        openCatalogModal('edit');
    });
    onClick('afgBtnDeleteCatalog', function () {
        void deleteActiveCatalogEntry();
    });
    onClick('afgCatalogModalCancel', function () {
        closeCatalogModal();
    });
    onClick('afgCatalogModalSave', function () {
        void submitCatalogModal();
    });
    const modal = document.getElementById('afgCatalogModal');
    if (modal) {
        modal.addEventListener('click', function (ev) {
            if (ev.target === modal) closeCatalogModal();
        });
    }
    const codeEl = document.getElementById('afgNewCode');
    const nameEl = document.getElementById('afgNewName');
    function onModalKeydown(ev) {
        if (ev.key !== 'Enter' || ev.shiftKey) return;
        const modalOpen = modal && modal.classList.contains('open');
        if (!modalOpen) return;
        ev.preventDefault();
        void submitCatalogModal();
    }
    if (codeEl) codeEl.addEventListener('keydown', onModalKeydown);
    if (nameEl) nameEl.addEventListener('keydown', onModalKeydown);
    if (codeEl) {
        codeEl.addEventListener('input', updateCatalogNickHint);
    }
    const prefixInp = document.getElementById('afgMailPrefix');
    if (prefixInp) {
        prefixInp.addEventListener('change', persistMailPrefixFromDom);
        prefixInp.addEventListener('blur', persistMailPrefixFromDom);
        prefixInp.addEventListener('input', function () {
            const preview = document.getElementById('afgMailPrefixPreview');
            const api = dataV2();
            let pre = String(prefixInp.value || '').trim();
            if (api && typeof api.mailNicknamePrefixSanitize === 'function') {
                pre = api.mailNicknamePrefixSanitize(pre, 24);
            } else {
                pre = pre
                    .toLowerCase()
                    .replace(/[@()[\]\\";:<>,\s]/g, '')
                    .slice(0, 24);
            }
            const sampleCode = activeCode || (activeKind === 'arge' ? 'FVV' : 'M');
            const nick = buildAfgMailNickname(pre, sampleCode);
            if (preview) preview.textContent = nick || '–';
        });
    }
    const nickInp = document.getElementById('slgNewMailNick');
    if (nickInp) {
        nickInp.addEventListener('input', function () {
            createNickDirty = true;
        });
    }
    const nickStdBtn = document.getElementById('slgBtnMailNickStandard');
    if (nickStdBtn) {
        nickStdBtn.addEventListener('click', function () {
            createNickDirty = false;
        });
    }
    document.addEventListener('keydown', function (ev) {
        if (ev.key !== 'Escape') return;
        if (modal && modal.classList.contains('open')) closeCatalogModal();
    });
    onClick('afgBtnSelectionMode', function () {
        setSelectionMode(!selectionMode);
    });
    onClick('afgBtnSelectUnmatched', selectVisibleUnmatched);
    onClick('afgBtnSelectMatched', selectVisibleMatched);
    onClick('afgBtnSelectNone', clearSelection);
    onClick('afgBtnAliasSettings', openAliasPrefixSettings);
    onClick('afgBtnBulkCreate', function () {
        runBulkCreateAndMatch().catch(function () {});
    });
    onClick('afgBtnBulkMatchExisting', function () {
        runBulkMatchExisting().catch(function () {});
    });
    onClick('afgBtnBulkRename', function () {
        runBulkStripDisplayPrefix().catch(function () {});
    });
    onClick('afgBtnBulkOwner', function () {
        if (!collectSelectedMatched().length) {
            toast('Bitte zuerst gematchte ' + kindLabel(2) + ' ankreuzen.');
            return;
        }
        showBulkOwnerPanel(true);
    });
    onClick('afgBtnBulkDelete', function () {
        runBulkDelete().catch(function () {});
    });
    onClick('afgBulkOwnerSearchBtn', function () {
        runBulkOwnerSearch().catch(function () {});
    });
    onClick('afgBulkOwnerApply', function () {
        runBulkSetOwner().catch(function () {});
    });
    const bulkOwnerSearch = document.getElementById('afgBulkOwnerSearch');
    if (bulkOwnerSearch) {
        bulkOwnerSearch.addEventListener('keydown', function (ev) {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                runBulkOwnerSearch().catch(function () {});
            }
        });
    }
}

function init() {
    maybeUpgradeLegacyAgPrefix();
    mountDetail();
    readLists();
    ensureActiveCode();
    wire();
    setActiveKind('subject');
    syncMailPrefixUi();
}

// Im Vite-Build läuft dieses Modul vor den klassischen defer-Skripten der Seite
// (group-detail.js, slg-live-details.js, app-data-v2.js). Erst ab DOMContentLoaded
// sind sie garantiert ausgeführt – 'interactive' reicht dafür nicht.
if (document.readyState === 'complete') {
    init();
} else {
    document.addEventListener('DOMContentLoaded', init);
}
