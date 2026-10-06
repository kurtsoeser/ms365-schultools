/**
 * Entra-ID-Gruppen suchen und per Dialog auswählen (Microsoft 365 / Sicherheitsgruppen).
 */
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const GROUP_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Group.Read.All'
];

function graphApi() {
    const g = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!g || typeof g.getGraphToken !== 'function' || typeof g.graphJson !== 'function') {
        throw new Error('Graph-Hilfen nicht geladen (spo-graph-shared).');
    }
    return g;
}

/**
 * @param {string} token
 * @param {string} pathOrUrl
 * @param {Record<string, string>} [extraHeaders]
 */
async function graphGet(token, pathOrUrl, extraHeaders) {
    const ug =
        typeof window !== 'undefined' && window.ms365GraphUnifiedGroups
            ? window.ms365GraphUnifiedGroups
            : null;
    if (ug && typeof ug.graphJson === 'function') {
        return await ug.graphJson('GET', pathOrUrl, token, undefined, extraHeaders || undefined);
    }
    const G = graphApi();
    let url = pathOrUrl;
    if (url.indexOf('http') !== 0) {
        url = 'https://graph.microsoft.com/v1.0' + (url.indexOf('/') === 0 ? url : '/' + url);
    }
    const headers = { Authorization: 'Bearer ' + token };
    if (extraHeaders && typeof extraHeaders === 'object') {
        Object.keys(extraHeaders).forEach((k) => {
            headers[k] = extraHeaders[k];
        });
    }
    const res = await fetch(url, { method: 'GET', headers });
    const text = await res.text();
    let data = null;
    if (text) {
        try {
            data = JSON.parse(text);
        } catch {
            data = text;
        }
    }
    if (!res.ok) {
        const msg =
            typeof data === 'object' && data && data.error
                ? data.error.message || JSON.stringify(data.error)
                : text || String(res.status);
        throw new Error(msg);
    }
    return data || {};
}

function escapeSearchPhrase(raw) {
    return String(raw || '')
        .replace(/"/g, '\\"')
        .replace(/\r?\n/g, ' ')
        .trim();
}

function dedupeGroups(list) {
    const seen = new Map();
    (list || []).forEach((g) => {
        if (g && g.id && !seen.has(g.id)) seen.set(g.id, g);
    });
    return [...seen.values()];
}

function odataEscape(s) {
    return String(s || '').replace(/'/g, "''");
}

/**
 * @param {string} queryRaw
 * @param {string} [token]
 * @returns {Promise<Array<{ id: string, displayName: string, mail: string, mailNickname: string, groupTypes: string[], securityEnabled: boolean }>>}
 */
export async function searchEntraGroups(queryRaw, token) {
    const q = String(queryRaw || '').trim();
    if (!q) return [];
    const tok = token || (await graphApi().getGraphToken(GROUP_SCOPES));
    const select = 'id,displayName,mail,mailNickname,groupTypes,securityEnabled,mailEnabled';
    const G = graphApi();

    if (GUID_RE.test(q)) {
        try {
            const g = await G.graphJson(
                'GET',
                '/groups/' + encodeURIComponent(q) + '?$select=' + encodeURIComponent(select),
                tok,
                undefined,
                'v1.0'
            );
            return g && g.id ? [g] : [];
        } catch {
            return [];
        }
    }

    const eventual = { ConsistencyLevel: 'eventual' };

    try {
        const phrase = escapeSearchPhrase(q);
        const aqs =
            '(displayName:' +
            phrase +
            ' OR mail:' +
            phrase +
            ' OR mailNickname:' +
            phrase +
            ')';
        const path =
            '/groups?$search=' +
            encodeURIComponent('"' + aqs + '"') +
            '&$select=' +
            encodeURIComponent(select) +
            '&$count=true&$top=25';
        const data = await graphGet(tok, path, eventual);
        if (Array.isArray(data.value) && data.value.length) return data.value;
    } catch {
        /* $search nicht verfügbar – Fallback */
    }

    const queryVariants = [q];
    const upper = q.toUpperCase();
    const title =
        q.length >= 1 ? q.charAt(0).toUpperCase() + q.slice(1).toLowerCase() : q;
    if (upper !== q) queryVariants.push(upper);
    if (title !== q && title !== upper) queryVariants.push(title);

    /** @type {object[]} */
    let merged = [];
    for (let vi = 0; vi < queryVariants.length; vi++) {
        const variant = queryVariants[vi];
        const esc = odataEscape(variant);
        try {
            const containsFilter =
                "contains(displayName,'" +
                esc +
                "') or contains(mailNickname,'" +
                esc +
                "') or contains(mail,'" +
                esc +
                "')";
            const pathContains =
                '/groups?$count=true&$filter=' +
                encodeURIComponent(containsFilter) +
                '&$select=' +
                encodeURIComponent(select) +
                '&$top=25';
            const dataContains = await graphGet(tok, pathContains, eventual);
            merged = merged.concat((dataContains && dataContains.value) || []);
        } catch {
            /* contains erfordert Advanced Query – nächster Versuch */
        }
        try {
            const prefixFilter =
                "startswith(displayName,'" +
                esc +
                "') or startswith(mailNickname,'" +
                esc +
                "') or startswith(mail,'" +
                esc +
                "')";
            const pathPrefix =
                '/groups?$filter=' +
                encodeURIComponent(prefixFilter) +
                '&$select=' +
                encodeURIComponent(select) +
                '&$top=25';
            const dataPrefix = await graphGet(tok, pathPrefix, undefined);
            merged = merged.concat((dataPrefix && dataPrefix.value) || []);
        } catch {
            /* ignore */
        }
    }
    return dedupeGroups(merged);
}

/**
 * @param {object} g
 */
export function formatGroupOptionLabel(g) {
    const name = String((g && g.displayName) || '').trim() || '(ohne Name)';
    const mail = String((g && g.mail) || (g && g.mailNickname) || '').trim();
    const types = (g && g.groupTypes) || [];
    const tags = [];
    if (g && g.securityEnabled) tags.push('Sicherheit');
    if (types.indexOf('Unified') >= 0) tags.push('M365');
    if (g && g.mailEnabled && types.indexOf('Unified') < 0) tags.push('Mail');
    const tag = tags.length ? ' · ' + tags.join(', ') : '';
    return mail ? name + ' (' + mail + ')' + tag : name + tag;
}

/**
 * @param {{ title?: string, hint?: string }} [opts]
 * @returns {Promise<{ id: string, displayName: string, label: string }|null>}
 */
export function pickEntraGroup(opts) {
    const title = (opts && opts.title) || 'Microsoft 365-Gruppe wählen';
    const hint =
        (opts && opts.hint) ||
        'Name, Mail, Alias oder Objekt-ID. Es werden Entra-Gruppen und Sicherheitsgruppen angezeigt.';

    return new Promise((resolve) => {
        const overlay = document.createElement('div');
        overlay.className = 'modal-overlay ms365-entra-group-picker';
        overlay.setAttribute('role', 'dialog');
        overlay.setAttribute('aria-modal', 'true');
        overlay.innerHTML =
            '<div class="modal-box" style="max-width:520px;width:min(96vw,520px);">' +
            '<h3 class="modal-box__title" style="margin:0 0 8px;">' +
            escapeHtml(title) +
            '</h3>' +
            '<p class="muted" style="margin:0 0 12px;font-size:0.9rem;">' +
            escapeHtml(hint) +
            '</p>' +
            '<div class="tm-field"><label>Suche</label>' +
            '<input type="search" class="ms365-egp-search" autocomplete="off" placeholder="z. B. Lehrer, SG-Schüler …"></div>' +
            '<p class="ms365-egp-status muted" style="min-height:1.2em;margin:8px 0;font-size:0.88em;"></p>' +
            '<ul class="ms365-egp-results" style="list-style:none;margin:0;padding:0;max-height:280px;overflow:auto;border:1px solid var(--border,#ddd);border-radius:8px;"></ul>' +
            '<div class="modal-box__actions" style="margin-top:14px;display:flex;gap:8px;justify-content:flex-end;">' +
            '<button type="button" class="btn alt ms365-egp-cancel">Abbrechen</button>' +
            '</div></div>';

        document.body.appendChild(overlay);
        overlay.classList.add('open');

        const searchEl = overlay.querySelector('.ms365-egp-search');
        const statusEl = overlay.querySelector('.ms365-egp-status');
        const listEl = overlay.querySelector('.ms365-egp-results');

        let debounce = 0;
        let closed = false;

        function close(result) {
            if (closed) return;
            closed = true;
            overlay.classList.remove('open');
            overlay.remove();
            resolve(result || null);
        }

        function renderResults(groups) {
            listEl.replaceChildren();
            if (!groups.length) {
                const li = document.createElement('li');
                li.className = 'muted';
                li.style.padding = '12px';
                li.textContent = 'Keine Treffer – Suchbegriff anpassen.';
                listEl.appendChild(li);
                return;
            }
            groups.forEach((g) => {
                const li = document.createElement('li');
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'ms365-egp-item';
                btn.style.cssText =
                    'display:block;width:100%;text-align:left;padding:10px 12px;border:0;border-bottom:1px solid var(--border,#eee);background:transparent;cursor:pointer;font:inherit;';
                btn.innerHTML =
                    '<strong>' +
                    escapeHtml(String(g.displayName || '')) +
                    '</strong><br><span class="muted" style="font-size:0.85em;">' +
                    escapeHtml(formatGroupOptionLabel(g)) +
                    '</span>';
                btn.addEventListener('click', () => {
                    close({
                        id: String(g.id),
                        displayName: String(g.displayName || ''),
                        label: formatGroupOptionLabel(g)
                    });
                });
                btn.addEventListener('mouseenter', () => {
                    btn.style.background = 'var(--surface-2,#f4f6f8)';
                });
                btn.addEventListener('mouseleave', () => {
                    btn.style.background = 'transparent';
                });
                li.appendChild(btn);
                listEl.appendChild(li);
            });
        }

        async function runSearch(q) {
            statusEl.textContent = 'Suche …';
            try {
                const list = await searchEntraGroups(q);
                statusEl.textContent = list.length ? list.length + ' Treffer' : 'Keine Treffer';
                renderResults(list);
            } catch (e) {
                statusEl.textContent = e && e.message ? e.message : String(e);
                renderResults([]);
            }
        }

        overlay.querySelector('.ms365-egp-cancel').addEventListener('click', () => close(null));
        overlay.addEventListener('click', (ev) => {
            if (ev.target === overlay) close(null);
        });

        searchEl.addEventListener('input', () => {
            const q = String(searchEl.value || '').trim();
            clearTimeout(debounce);
            if (q.length < 2) {
                statusEl.textContent = 'Mindestens 2 Zeichen eingeben.';
                listEl.replaceChildren();
                return;
            }
            debounce = setTimeout(() => runSearch(q), 280);
        });
        searchEl.addEventListener('keydown', (ev) => {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                const q = String(searchEl.value || '').trim();
                if (q.length >= 2) runSearch(q);
            }
        });

        setTimeout(() => searchEl.focus(), 50);
    });
}

/**
 * Mehrere Entra-Gruppen per Suche auswählen (Checkboxen).
 * @param {{
 *   title?: string,
 *   hint?: string,
 *   initialQueries?: string[],
 *   filterFn?: (g: object) => boolean,
 *   searchFn?: (query: string, token?: string) => Promise<object[]>
 * }} [opts]
 * @returns {Promise<object[]>}
 */
export function pickEntraGroupsMulti(opts) {
    const title = (opts && opts.title) || 'Gruppen aus Microsoft 365 übernehmen';
    const hint =
        (opts && opts.hint) ||
        'Gruppen suchen und auswählen. Standard-Suche: arge, fg, ag – Sie können den Suchbegriff anpassen.';
    const initialQueries = Array.isArray(opts && opts.initialQueries) ? opts.initialQueries : ['arge', 'fg', 'ag'];
    const filterFn = opts && typeof opts.filterFn === 'function' ? opts.filterFn : null;
    const searchFn =
        opts && typeof opts.searchFn === 'function'
            ? opts.searchFn
            : (q, tok) => searchEntraGroups(q, tok);

    return new Promise((resolve) => {
        const overlay = document.createElement('div');
        overlay.className = 'modal-overlay ms365-entra-group-picker ms365-entra-group-picker--multi';
        overlay.setAttribute('role', 'dialog');
        overlay.setAttribute('aria-modal', 'true');
        overlay.innerHTML =
            '<div class="modal-box" style="max-width:560px;width:min(96vw,560px);">' +
            '<h3 class="modal-box__title" style="margin:0 0 8px;">' +
            escapeHtml(title) +
            '</h3>' +
            '<p class="muted" style="margin:0 0 12px;font-size:0.9rem;">' +
            escapeHtml(hint) +
            '</p>' +
            '<div class="tm-field"><label>Suche</label>' +
            '<input type="search" class="ms365-egp-search" autocomplete="off" placeholder="z. B. arge, fg-deutsch, ag-sport …"></div>' +
            '<label class="cp-check" style="display:flex;align-items:center;gap:8px;margin:10px 0 6px;font-size:0.9em;">' +
            '<input type="checkbox" class="ms365-egp-prefix-only" checked> Nur Gruppen mit Alias/Name arge-, fg- oder ag-</label>' +
            '<p class="ms365-egp-status muted" style="min-height:1.2em;margin:8px 0;font-size:0.88em;"></p>' +
            '<ul class="ms365-egp-results" style="list-style:none;margin:0;padding:0;max-height:300px;overflow:auto;border:1px solid var(--border,#ddd);border-radius:8px;"></ul>' +
            '<div class="modal-box__actions" style="margin-top:14px;display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap;">' +
            '<button type="button" class="btn alt ms365-egp-cancel">Abbrechen</button>' +
            '<button type="button" class="btn btn-success ms365-egp-apply" disabled>Übernehmen</button>' +
            '</div></div>';

        document.body.appendChild(overlay);
        overlay.classList.add('open');

        const searchEl = overlay.querySelector('.ms365-egp-search');
        const prefixOnlyEl = overlay.querySelector('.ms365-egp-prefix-only');
        const statusEl = overlay.querySelector('.ms365-egp-status');
        const listEl = overlay.querySelector('.ms365-egp-results');
        const applyBtn = overlay.querySelector('.ms365-egp-apply');

        const selected = new Map();
        let debounce = 0;
        let closed = false;
        let lastGroups = [];

        function close(result) {
            if (closed) return;
            closed = true;
            overlay.classList.remove('open');
            overlay.remove();
            resolve(Array.isArray(result) ? result : []);
        }

        function defaultArgePrefixFilter(g) {
            const nick = String((g && g.mailNickname) || '').toLowerCase();
            const dn = String((g && g.displayName) || '').toLowerCase();
            if (/^(arge|fg|ag)[-._]/.test(nick)) return true;
            if (/\b(arge|fg|ag)[-._]/.test(nick)) return true;
            if (/^(arge|fachgruppe|fg|ag)\b/.test(dn)) return true;
            return false;
        }

        function passesFilter(g) {
            if (filterFn && !filterFn(g)) return false;
            if (prefixOnlyEl && prefixOnlyEl.checked && !defaultArgePrefixFilter(g)) return false;
            return true;
        }

        function updateApplyState() {
            if (applyBtn) applyBtn.disabled = selected.size === 0;
        }

        function renderResults(groups) {
            lastGroups = groups;
            listEl.replaceChildren();
            const visible = groups.filter(passesFilter);
            if (!visible.length) {
                const li = document.createElement('li');
                li.className = 'muted';
                li.style.padding = '12px';
                li.textContent = groups.length
                    ? 'Keine Treffer mit aktuellem Filter – Filter abwählen oder Suchbegriff ändern.'
                    : 'Keine Treffer – Suchbegriff anpassen.';
                listEl.appendChild(li);
                return;
            }
            visible.forEach((g) => {
                const id = String(g.id || '');
                const li = document.createElement('li');
                li.style.borderBottom = '1px solid var(--border,#eee)';
                const row = document.createElement('label');
                row.style.cssText =
                    'display:flex;gap:10px;align-items:flex-start;padding:10px 12px;cursor:pointer;font:inherit;';
                const cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.checked = selected.has(id);
                cb.addEventListener('change', () => {
                    if (cb.checked) selected.set(id, g);
                    else selected.delete(id);
                    updateApplyState();
                });
                const body = document.createElement('span');
                body.innerHTML =
                    '<strong>' +
                    escapeHtml(String(g.displayName || '')) +
                    '</strong><br><span class="muted" style="font-size:0.85em;">' +
                    escapeHtml(formatGroupOptionLabel(g)) +
                    '</span>';
                row.appendChild(cb);
                row.appendChild(body);
                li.appendChild(row);
                listEl.appendChild(li);
            });
        }

        async function runSearch(q) {
            statusEl.textContent = 'Suche …';
            try {
                const list = await searchFn(q);
                statusEl.textContent = list.length ? list.length + ' Treffer' : 'Keine Treffer';
                renderResults(list);
            } catch (e) {
                statusEl.textContent = e && e.message ? e.message : String(e);
                renderResults([]);
            }
        }

        async function runInitialQueries() {
            statusEl.textContent = 'Lade Vorschläge (arge, fg, ag) …';
            try {
                const tok = await graphApi().getGraphToken(GROUP_SCOPES);
                const seen = new Map();
                for (let i = 0; i < initialQueries.length; i++) {
                    const q = String(initialQueries[i] || '').trim();
                    if (!q) continue;
                    const list = await searchFn(q, tok);
                    (list || []).forEach((g) => {
                        if (g && g.id && !seen.has(g.id)) seen.set(g.id, g);
                    });
                }
                const merged = [];
                seen.forEach((g) => merged.push(g));
                statusEl.textContent = merged.length ? merged.length + ' Vorschläge' : 'Keine Vorschläge – bitte suchen';
                renderResults(merged);
            } catch (e) {
                statusEl.textContent = e && e.message ? e.message : String(e);
                renderResults([]);
            }
        }

        overlay.querySelector('.ms365-egp-cancel').addEventListener('click', () => close([]));
        overlay.addEventListener('click', (ev) => {
            if (ev.target === overlay) close([]);
        });
        applyBtn.addEventListener('click', () => {
            const out = [];
            selected.forEach((g) => out.push(g));
            close(out);
        });
        if (prefixOnlyEl) {
            prefixOnlyEl.addEventListener('change', () => renderResults(lastGroups));
        }

        searchEl.addEventListener('input', () => {
            const q = String(searchEl.value || '').trim();
            clearTimeout(debounce);
            if (q.length < 2) {
                if (!q.length) {
                    runInitialQueries();
                    return;
                }
                statusEl.textContent = 'Mindestens 2 Zeichen eingeben.';
                listEl.replaceChildren();
                return;
            }
            debounce = setTimeout(() => runSearch(q), 280);
        });
        searchEl.addEventListener('keydown', (ev) => {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                const q = String(searchEl.value || '').trim();
                if (q.length >= 2) runSearch(q);
            }
        });

        runInitialQueries();
        setTimeout(() => searchEl.focus(), 50);
    });
}

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * @param {{ fields: Array<{ labelInputId: string, idInputId: string, pickBtnId: string, clearBtnId?: string, dialogTitle?: string }>, onChange?: () => void }} spec
 */
export function wireEntraGroupPickerFields(spec) {
    const fields = (spec && spec.fields) || [];
    const onChange = spec && spec.onChange;

    fields.forEach((f) => {
        const pickBtn = document.getElementById(f.pickBtnId);
        const clearBtn = f.clearBtnId ? document.getElementById(f.clearBtnId) : null;
        const labelInput = document.getElementById(f.labelInputId);
        const idInput = document.getElementById(f.idInputId);
        if (!pickBtn || !labelInput || !idInput) return;
        if (pickBtn.dataset.frEgpWired === '1') return;
        pickBtn.dataset.frEgpWired = '1';

        pickBtn.addEventListener('click', () => {
            pickEntraGroup({ title: f.dialogTitle || 'Gruppe wählen' })
                .then((sel) => {
                    if (!sel) return;
                    idInput.value = sel.id;
                    labelInput.value = sel.label || sel.displayName;
                    if (typeof onChange === 'function') onChange();
                })
                .catch((e) => {
                    const msg = e && e.message ? e.message : String(e);
                    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
                    else window.alert(msg);
                });
        });

        if (clearBtn && clearBtn.dataset.frEgpWired !== '1') {
            clearBtn.dataset.frEgpWired = '1';
            clearBtn.addEventListener('click', () => {
                idInput.value = '';
                labelInput.value = '';
                if (typeof onChange === 'function') onChange();
            });
        }
    });
}

/**
 * @param {{ labelInputId: string, idInputId: string }} field
 */
export function readGroupPickerField(field) {
    const labelInput = document.getElementById(field.labelInputId);
    const idInput = document.getElementById(field.idInputId);
    return {
        id: String((idInput && idInput.value) || '').trim(),
        label: String((labelInput && labelInput.value) || '').trim()
    };
}

/**
 * @param {{ labelInputId: string, idInputId: string }} field
 * @param {{ id?: string, label?: string }} value
 */
export function fillGroupPickerField(field, value) {
    const labelInput = document.getElementById(field.labelInputId);
    const idInput = document.getElementById(field.idInputId);
    if (idInput) idInput.value = String((value && value.id) || '').trim();
    if (labelInput) labelInput.value = String((value && value.label) || '').trim();
}
