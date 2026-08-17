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
            p.textContent = 'Keine Direktion‑Owner in den Schul‑Einstellungen gefunden.';
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
            return;
        }

        list.forEach(function (row) {
            const link = getCatalogLink(activeKind, row.code);
            const gid = link && link.graphGroupId ? String(link.graphGroupId) : '';
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
            meta.textContent = gid ? 'Gematcht: ' + gid : 'Noch kein Match';
            main.appendChild(t);
            main.appendChild(meta);
            btn.appendChild(main);
            li.appendChild(btn);
            host.appendChild(li);
        });
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

    function persistUnmatch() {
        const api = dataV2();
        if (api && typeof api.clearCatalogLinkGroup === 'function') {
            api.clearCatalogLinkGroup(activeKind, activeCode);
        } else {
            persistMatch({ id: '', displayName: '', mailNickname: '' }, '');
        }
        renderLeftList();
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
