(function () {
    'use strict';

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
    /** @type {{ code: string, name: string, subjects?: string[] }[]} */
    let catalog = { subject: [], arge: [] };
    /** @type {string[]} */
    let direktion = [];
    /** @type {{ subject: Set<string>, arge: Set<string> }} */
    let selectedKeysByKind = { subject: new Set(), arge: new Set() };

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

    function normStr(v) {
        return String(v ?? '').trim();
    }
    function normCode(v) {
        return normStr(v).toUpperCase();
    }
    function normEmail(v) {
        return normStr(v).toLowerCase();
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
        const raw = kind === 'arge' ? (su && su.argeGroupMailPrefix) || 'ag' : (su && su.subjectGroupMailPrefix) || 'fach';
        if (api && typeof api.mailNicknamePrefixSanitize === 'function') {
            return api.mailNicknamePrefixSanitize(raw, 24) || (kind === 'arge' ? 'ag' : 'fach');
        }
        return kind === 'arge' ? 'ag' : 'fach';
    }

    function deriveNick(kind, code) {
        const pre = mailPrefix(kind);
        const tail = gug().sanitizeMailNickname(String(code || 'x')).slice(0, 40);
        return gug().sanitizeUnifiedGroupMailNickname(String(pre + tail).toLowerCase()).slice(0, 60);
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
                    ? 'Änderungen werden in die Schul‑Einstellungen geschrieben. Bei neuem Kürzel wird ein vorhandenes Match mit umgehängt.'
                    : 'Änderungen werden in die Schul‑Einstellungen geschrieben. Bei neuem Kürzel wird ein vorhandenes Match mit umgehängt.';
            } else {
                hint.textContent = isArge
                    ? 'Die ARGE wird in die Schul‑Einstellungen geschrieben und erscheint danach im Katalog.'
                    : 'Das Fach wird in die Schul‑Einstellungen geschrieben und erscheint danach im Katalog.';
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
            throw new Error('Schul‑Einstellungen nicht verfügbar (tenant-settings-core.js).');
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
            throw new Error('Schul‑Einstellungen nicht verfügbar (tenant-settings-core.js).');
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
            throw new Error('Schul‑Einstellungen nicht verfügbar (tenant-settings-core.js).');
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
            ') aus den Schul‑Einstellungen entfernen?';
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
            p.textContent = 'Keine Direktion‑Besitzer in den Schul‑Einstellungen gefunden.';
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

    function applyCreateDefaults() {
        const row = getActiveRow();
        const code = row ? row.code : activeCode;
        const name = row && row.name ? row.name : code;
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const desc = document.getElementById('slgNewDescription');
        const label = activeKind === 'arge' ? 'Arbeitsgruppe ' : 'Fach ';
        if (dn) dn.value = label + (name || code || '');
        if (nn) nn.value = deriveNick(activeKind, code);
        if (desc) {
            desc.value =
                label +
                (name || code || '') +
                (activeKind === 'arge' ? ' (MS365-Schulverwaltung / ARGE)' : ' (MS365-Schulverwaltung / Fachgruppe)');
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

    function renderLeftList() {
        const host = document.getElementById('afgListItems');
        const summary = document.getElementById('afgListSummary');
        const empty = document.getElementById('afgEmptyHint');
        const wrap = document.getElementById('afgDetailWrap');
        if (!host) return;
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
        if (summary) {
            summary.textContent =
                String(all.length) +
                (activeKind === 'arge' ? ' ARGE' : ' Fächer') +
                ' · ' +
                String(matchedN) +
                ' gematcht' +
                (q ? ' · Filter: ' + String(list.length) : '');
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
            badge.className = 'jg-match-badge ' + (gid ? 'is-ok' : 'is-warn');
            const ico = document.createElement('i');
            ico.className = gid ? 'bi bi-check-circle-fill' : 'bi bi-exclamation-circle-fill';
            ico.setAttribute('aria-hidden', 'true');
            badge.appendChild(ico);
            badge.appendChild(document.createTextNode(gid ? 'Gematcht' : 'Kein Match'));
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
            const pick = document.createElement('label');
            pick.className = 'afg-pick';
            pick.title = gid ? 'Für Sammelaktion auswählen' : 'Nur gematchte Gruppen können ausgewählt werden';
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.setAttribute('data-afg-pick', key);
            cb.checked = gid ? selectedKeys().has(key) : false;
            cb.disabled = !gid;
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
            li.appendChild(btn);
            host.appendChild(li);
        });
        if (window.ms365GroupPhotoThumb && typeof window.ms365GroupPhotoThumb.hydrate === 'function') {
            window.ms365GroupPhotoThumb.hydrate(host);
        }
        updateBulkCount();
        updateCatalogActionButtons();
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
        document.querySelectorAll('[data-afg-kind]').forEach(function (b) {
            b.setAttribute('aria-pressed', b.getAttribute('data-afg-kind') === activeKind ? 'true' : 'false');
        });
        ensureActiveCode();
        showBulkOwnerPanel(false);
        setBulkStatus('', false);
        const search = document.getElementById('slgGroupSearch');
        if (search) search.value = '';
        gd().clearSearchResults();
        renderLeftList();
        applyCreateDefaults();
        gd().setTab('general');
        refreshMatchUi();
    }

    function setActiveCode(code) {
        activeCode = normCode(code);
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
        const row = getActiveRow();
        upsertCatalogLink({
            kind: activeKind,
            code: activeCode,
            graphGroupId: g && g.id ? String(g.id) : '',
            displayName: (g && g.displayName) || (row && row.name) || '',
            mailNickname: (g && g.mailNickname) || '',
            mode: mode
        });
        renderLeftList();
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
            let row = null;
            for (let i = 0; i < list.length; i++) {
                if (normCode(list[i].code) === key) {
                    row = list[i];
                    break;
                }
            }
            if (!row) return;
            const link = getCatalogLink(activeKind, row.code);
            if (link && link.graphGroupId) keep.add(key);
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

    function updateBulkCount() {
        pruneSelection();
        const n = selectedKeys().size;
        const el = document.getElementById('afgBulkCount');
        if (el) {
            const label = n === 1 ? '1 Gruppe ausgewählt' : String(n) + ' Gruppen ausgewählt';
            el.innerHTML = '<i class="bi bi-check2-square" aria-hidden="true"></i>' + label;
            el.classList.toggle('is-active', n > 0);
        }
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

    function selectVisibleMatched() {
        visibleMatchedRows().forEach(function (row) {
            selectedKeys().add(normCode(row.code));
        });
        renderLeftList();
        const n = collectSelectedMatched().length;
        toast(
            n
                ? String(n) + ' gematchte ' + kindLabel(n) + ' angekreuzt.'
                : 'Keine gematchten Gruppen in der aktuellen Liste.'
        );
    }

    function clearSelection() {
        selectedKeysByKind[activeKind] = new Set();
        renderLeftList();
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
                'Legt eine Microsoft 365‑Gruppe (Unified) an und verknüpft sie mit diesem Katalogeintrag. Optional auch als Team bereitstellen.',
            membersUnmatchedHint:
                'In den Stammdaten gibt es keine Mitgliederliste für Fach/ARGE. Nach dem Match können Sie Mitglieder live in Graph pflegen.',
            membersUnmatchedTitle: 'Hinweis aus den Schul‑Einstellungen',
            membersMatchedHint:
                'Live aus Microsoft Graph. Es gibt keinen automatischen Listen‑Sync – Mitglieder hier suchen, hinzufügen oder entfernen.',
            emptyHintHtml:
                'Keine Einträge in dieser Liste. Legen Sie einen Eintrag über <strong>Neu</strong> an oder pflegen Sie Fächer bzw. ARGE unter <a href="../tenant.html">Schul‑Einstellungen</a>.',
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
                    if (!direktion.length) throw new Error('Keine Direktion‑Adressen in den Schul‑Einstellungen.');
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
                    return gug().ensureOwners(token, gid, direktion || []);
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
        document.addEventListener('keydown', function (ev) {
            if (ev.key !== 'Escape') return;
            if (modal && modal.classList.contains('open')) closeCatalogModal();
        });
        onClick('afgBtnSelectMatched', selectVisibleMatched);
        onClick('afgBtnSelectNone', clearSelection);
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
        mountDetail();
        readLists();
        ensureActiveCode();
        wire();
        setActiveKind('subject');
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
