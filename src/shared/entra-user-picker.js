/**
 * Entra-Benutzer suchen und per Dialog auswählen.
 */
import { getGraphPickerApi } from './graph-picker-backend.js';

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const USER_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All'
];

const USER_SELECT = 'id,displayName,mail,userPrincipalName,givenName,surname';

function graphApi() {
    return getGraphPickerApi();
}

function odataEscape(s) {
    return String(s || '').replace(/'/g, "''");
}

function escapeSearchPhrase(raw) {
    return (
        '"' +
        String(raw || '')
            .replace(/"/g, '')
            .replace(/\r?\n/g, ' ')
            .trim() +
        '"'
    );
}

/**
 * GET mit optionalen Extra-Headern (z. B. ConsistencyLevel für $search/$count).
 * spo-graph-shared akzeptiert keine Header als 5. Argument – daher ug zuerst, sonst fetch.
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
    if (!extraHeaders || typeof extraHeaders !== 'object' || !Object.keys(extraHeaders).length) {
        const G = graphApi();
        return await G.graphJson('GET', pathOrUrl, token, undefined, 'v1.0');
    }
    let url = pathOrUrl;
    if (url.indexOf('http') !== 0) {
        url = 'https://graph.microsoft.com/v1.0' + (url.indexOf('/') === 0 ? url : '/' + url);
    }
    const headers = { Authorization: 'Bearer ' + token };
    Object.keys(extraHeaders).forEach(function (k) {
        headers[k] = extraHeaders[k];
    });
    const res = await fetch(url, { method: 'GET', headers: headers });
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

function dedupeUsers(list) {
    const seen = new Map();
    (list || []).forEach(function (u) {
        if (u && u.id && !seen.has(u.id)) seen.set(u.id, u);
    });
    return Array.from(seen.values());
}

/**
 * @param {object} u
 */
export function formatUserOptionLabel(u) {
    const name = String((u && u.displayName) || '').trim() || '(ohne Name)';
    const mail = String((u && u.mail) || (u && u.userPrincipalName) || '').trim();
    return mail ? name + ' (' + mail + ')' : name;
}

/**
 * Personen in Entra/M365 suchen.
 * Nutzt $search (Nachname/Teiltreffer), dann contains, dann startswith – analog zur Gruppensuche.
 *
 * @param {string} queryRaw
 * @param {string} [token]
 */
export async function searchEntraUsers(queryRaw, token) {
    const q = String(queryRaw || '').trim();
    if (!q) return [];
    const tok = token || (await graphApi().getGraphToken(USER_SCOPES));
    const G = graphApi();
    const select = USER_SELECT;
    const eventual = { ConsistencyLevel: 'eventual' };

    if (GUID_RE.test(q)) {
        try {
            const u = await G.graphJson(
                'GET',
                '/users/' + encodeURIComponent(q) + '?$select=' + encodeURIComponent(select),
                tok,
                undefined,
                'v1.0'
            );
            return u && u.id ? [u] : [];
        } catch {
            return [];
        }
    }

    if (q.indexOf('@') !== -1) {
        const esc = odataEscape(q);
        const filter = "mail eq '" + esc + "' or userPrincipalName eq '" + esc + "'";
        try {
            const data = await graphGet(
                tok,
                '/users?$filter=' +
                    encodeURIComponent(filter) +
                    '&$select=' +
                    encodeURIComponent(select) +
                    '&$top=25'
            );
            const list = Array.isArray(data.value) ? data.value : [];
            if (list.length) return list;
        } catch {
            /* weiter mit Volltext */
        }
    }

    try {
        const phrase = escapeSearchPhrase(q);
        const aqs =
            '(displayName:' +
            phrase +
            ' OR givenName:' +
            phrase +
            ' OR surname:' +
            phrase +
            ' OR mail:' +
            phrase +
            ' OR userPrincipalName:' +
            phrase +
            ')';
        const path =
            '/users?$search=' +
            encodeURIComponent('"' + aqs + '"') +
            '&$select=' +
            encodeURIComponent(select) +
            '&$count=true&$top=25';
        const data = await graphGet(tok, path, eventual);
        const list = Array.isArray(data.value) ? data.value : [];
        if (list.length) return list;
    } catch {
        /* $search nicht verfügbar – Fallback */
    }

    const queryVariants = [q];
    const upper = q.toUpperCase();
    const title = q.length >= 1 ? q.charAt(0).toUpperCase() + q.slice(1).toLowerCase() : q;
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
                "') or contains(givenName,'" +
                esc +
                "') or contains(surname,'" +
                esc +
                "') or contains(mail,'" +
                esc +
                "') or contains(userPrincipalName,'" +
                esc +
                "')";
            const pathContains =
                '/users?$count=true&$filter=' +
                encodeURIComponent(containsFilter) +
                '&$select=' +
                encodeURIComponent(select) +
                '&$top=25';
            const dataContains = await graphGet(tok, pathContains, eventual);
            merged = merged.concat((dataContains && dataContains.value) || []);
        } catch {
            /* contains braucht Advanced Query */
        }
        try {
            const prefixFilter =
                "startswith(displayName,'" +
                esc +
                "') or startswith(givenName,'" +
                esc +
                "') or startswith(surname,'" +
                esc +
                "') or startswith(mail,'" +
                esc +
                "') or startswith(userPrincipalName,'" +
                esc +
                "')";
            const pathPrefix =
                '/users?$filter=' +
                encodeURIComponent(prefixFilter) +
                '&$select=' +
                encodeURIComponent(select) +
                '&$top=25';
            const dataPrefix = await graphGet(tok, pathPrefix);
            merged = merged.concat((dataPrefix && dataPrefix.value) || []);
        } catch {
            /* ignore */
        }
        if (merged.length) break;
    }
    return dedupeUsers(merged);
}

/**
 * @param {{ title?: string, hint?: string, initialQuery?: string }} [opts]
 * @returns {Promise<{ id: string, displayName: string, mail: string, label: string }|null>}
 */
export function pickEntraUser(opts) {
    const title = (opts && opts.title) || 'Person aus Microsoft Entra wählen';
    const hint =
        (opts && opts.hint) || 'Name oder E-Mail-Adresse (mindestens 2 Zeichen).';
    const initialQuery = String((opts && opts.initialQuery) || '').trim();

    return new Promise((resolve) => {
        const overlay = document.createElement('div');
        overlay.className = 'modal-overlay ms365-entra-user-picker';
        overlay.setAttribute('role', 'dialog');
        overlay.setAttribute('aria-modal', 'true');
        overlay.innerHTML =
            '<div class="modal-box ms365-egp-modal">' +
            '<div class="ms365-egp-modal__head">' +
            '<span class="ms365-egp-modal__icon" aria-hidden="true"><i class="bi bi-person-bounding-box"></i></span>' +
            '<div class="ms365-egp-modal__titles">' +
            '<h3 class="modal-box__title ms365-egp-modal__title">' +
            escapeHtml(title) +
            '</h3>' +
            '<p class="muted ms365-egp-modal__hint">' +
            escapeHtml(hint) +
            '</p></div></div>' +
            '<div class="field ms365-egp-search-wrap">' +
            '<label for="ms365-eup-search-input">Suche in Microsoft Entra</label>' +
            '<div class="ms365-egp-search-row">' +
            '<i class="bi bi-search ms365-egp-search-icon" aria-hidden="true"></i>' +
            '<input type="search" id="ms365-eup-search-input" class="ms365-eup-search ms365-egp-search" autocomplete="off" placeholder="Name oder E-Mail …">' +
            '</div></div>' +
            '<p class="ms365-eup-status ms365-egp-status muted" role="status" aria-live="polite"></p>' +
            '<ul class="ms365-eup-results ms365-egp-results" aria-label="Suchergebnisse"></ul>' +
            '<div class="modal-box__actions ms365-egp-modal__actions">' +
            '<button type="button" class="btn alt ms365-eup-cancel">Abbrechen</button>' +
            '</div></div>';

        document.body.appendChild(overlay);
        overlay.classList.add('open');

        const searchEl = overlay.querySelector('.ms365-eup-search');
        const statusEl = overlay.querySelector('.ms365-eup-status');
        const listEl = overlay.querySelector('.ms365-eup-results');

        let debounce = 0;
        let closed = false;

        function close(result) {
            if (closed) return;
            closed = true;
            overlay.classList.remove('open');
            overlay.remove();
            resolve(result || null);
        }

        function renderResults(users) {
            listEl.replaceChildren();
            if (!users.length) {
                const li = document.createElement('li');
                li.className = 'ms365-egp-empty muted';
                li.textContent = 'Keine Treffer – Suchbegriff anpassen.';
                listEl.appendChild(li);
                return;
            }
            users.forEach((u) => {
                const li = document.createElement('li');
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'ms365-egp-item ms365-eup-item';
                const mail = String(u.mail || u.userPrincipalName || '').trim();
                btn.innerHTML =
                    '<span class="ms365-egp-item__name">' +
                    escapeHtml(String(u.displayName || '')) +
                    '</span>' +
                    '<span class="ms365-egp-item__meta muted">' +
                    escapeHtml(mail || '–') +
                    '</span>';
                btn.addEventListener('click', () => {
                    close({
                        id: String(u.id),
                        displayName: String(u.displayName || ''),
                        mail,
                        label: formatUserOptionLabel(u)
                    });
                });
                li.appendChild(btn);
                listEl.appendChild(li);
            });
        }

        async function runSearch(q) {
            statusEl.textContent = 'Suche …';
            try {
                const list = await searchEntraUsers(q);
                statusEl.textContent = list.length ? list.length + ' Treffer' : 'Keine Treffer';
                renderResults(list);
            } catch (e) {
                statusEl.textContent = e && e.message ? e.message : String(e);
                renderResults([]);
            }
        }

        overlay.querySelector('.ms365-eup-cancel').addEventListener('click', () => close(null));
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

        setTimeout(() => {
            searchEl.focus();
            if (initialQuery.length >= 2) {
                searchEl.value = initialQuery;
                runSearch(initialQuery);
            } else if (initialQuery) {
                searchEl.value = initialQuery;
                statusEl.textContent = 'Mindestens 2 Zeichen eingeben.';
            }
        }, 50);
    });
}

/**
 * Mehrere Entra-Benutzer suchen und auswählen.
 * @param {{ title?: string, hint?: string, initialQuery?: string, maxCount?: number }} [opts]
 * @returns {Promise<Array<{ id: string, displayName: string, mail: string, label: string }>>}
 */
export function pickEntraUsers(opts) {
    const title = (opts && opts.title) || 'Personen aus Microsoft Entra wählen';
    const hint =
        (opts && opts.hint) ||
        'Namen suchen, Treffer anhaken und mit „Übernehmen“ bestätigen.';
    const initialQuery = String((opts && opts.initialQuery) || '').trim();
    const maxCount =
        typeof (opts && opts.maxCount) === 'number' && opts.maxCount > 0 ? Math.floor(opts.maxCount) : 0;

    return new Promise((resolve) => {
        const overlay = document.createElement('div');
        overlay.className = 'modal-overlay ms365-entra-user-picker ms365-entra-user-picker--multi';
        overlay.setAttribute('role', 'dialog');
        overlay.setAttribute('aria-modal', 'true');
        overlay.innerHTML =
            '<div class="modal-box ms365-egp-modal">' +
            '<div class="ms365-egp-modal__head">' +
            '<span class="ms365-egp-modal__icon" aria-hidden="true"><i class="bi bi-people"></i></span>' +
            '<div class="ms365-egp-modal__titles">' +
            '<h3 class="modal-box__title ms365-egp-modal__title">' +
            escapeHtml(title) +
            '</h3>' +
            '<p class="muted ms365-egp-modal__hint">' +
            escapeHtml(hint) +
            (maxCount ? ' Maximal ' + maxCount + ' Person(en).' : '') +
            '</p></div></div>' +
            '<div class="field ms365-egp-search-wrap">' +
            '<label for="ms365-eup-multi-search">Suche in Microsoft Entra</label>' +
            '<div class="ms365-egp-search-row">' +
            '<i class="bi bi-search ms365-egp-search-icon" aria-hidden="true"></i>' +
            '<input type="search" id="ms365-eup-multi-search" class="ms365-eup-search ms365-egp-search" autocomplete="off" placeholder="Name oder E-Mail …">' +
            '</div></div>' +
            '<p class="ms365-eup-status ms365-egp-status muted" role="status" aria-live="polite"></p>' +
            '<ul class="ms365-eup-results ms365-egp-results" aria-label="Suchergebnisse"></ul>' +
            '<p class="ms365-eup-selected muted" style="margin:8px 0 0;font-size:0.88em;" aria-live="polite"></p>' +
            '<div class="modal-box__actions ms365-egp-modal__actions">' +
            '<button type="button" class="btn alt ms365-eup-cancel">Abbrechen</button>' +
            '<button type="button" class="btn btn-success ms365-eup-apply" disabled>Übernehmen</button>' +
            '</div></div>';

        document.body.appendChild(overlay);
        overlay.classList.add('open');

        const searchEl = overlay.querySelector('.ms365-eup-search');
        const statusEl = overlay.querySelector('.ms365-eup-status');
        const listEl = overlay.querySelector('.ms365-eup-results');
        const selectedEl = overlay.querySelector('.ms365-eup-selected');
        const applyBtn = overlay.querySelector('.ms365-eup-apply');

        /** @type {Map<string, { id: string, displayName: string, mail: string, label: string }>} */
        const selected = new Map();
        let debounce = 0;
        let closed = false;

        function close(result) {
            if (closed) return;
            closed = true;
            overlay.classList.remove('open');
            overlay.remove();
            resolve(Array.isArray(result) ? result : []);
        }

        function toPick(u) {
            const mail = String((u && (u.mail || u.userPrincipalName)) || '').trim();
            return {
                id: String((u && u.id) || ''),
                displayName: String((u && u.displayName) || ''),
                mail: mail,
                label: formatUserOptionLabel(u)
            };
        }

        function updateSelectedUi() {
            if (applyBtn) applyBtn.disabled = selected.size === 0;
            if (selectedEl) {
                selectedEl.textContent = selected.size
                    ? selected.size + ' ausgewählt: ' + Array.from(selected.values())
                          .map(function (p) {
                              return p.displayName || p.mail;
                          })
                          .join(', ')
                    : 'Noch niemand ausgewählt.';
            }
        }

        function renderResults(users) {
            listEl.replaceChildren();
            if (!users.length) {
                const li = document.createElement('li');
                li.className = 'ms365-egp-empty muted';
                li.textContent = 'Keine Treffer – Suchbegriff anpassen.';
                listEl.appendChild(li);
                return;
            }
            users.forEach((u) => {
                const pick = toPick(u);
                if (!pick.id) return;
                const li = document.createElement('li');
                const row = document.createElement('label');
                row.className = 'ms365-egp-item ms365-eup-item ms365-eup-item--check';
                row.style.cssText =
                    'display:flex;gap:10px;align-items:flex-start;cursor:pointer;width:100%;box-sizing:border-box;';
                const cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.checked = selected.has(pick.id);
                cb.addEventListener('change', () => {
                    if (cb.checked) {
                        if (maxCount && selected.size >= maxCount && !selected.has(pick.id)) {
                            cb.checked = false;
                            statusEl.textContent = 'Maximal ' + maxCount + ' Person(en) erlaubt.';
                            return;
                        }
                        selected.set(pick.id, pick);
                    } else {
                        selected.delete(pick.id);
                    }
                    updateSelectedUi();
                });
                const body = document.createElement('span');
                body.innerHTML =
                    '<span class="ms365-egp-item__name">' +
                    escapeHtml(pick.displayName) +
                    '</span>' +
                    '<span class="ms365-egp-item__meta muted">' +
                    escapeHtml(pick.mail || '–') +
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
                const list = await searchEntraUsers(q);
                statusEl.textContent = list.length ? list.length + ' Treffer' : 'Keine Treffer';
                renderResults(list);
            } catch (e) {
                statusEl.textContent = e && e.message ? e.message : String(e);
                renderResults([]);
            }
        }

        overlay.querySelector('.ms365-eup-cancel').addEventListener('click', () => close([]));
        applyBtn.addEventListener('click', () => close(Array.from(selected.values())));
        overlay.addEventListener('click', (ev) => {
            if (ev.target === overlay) close([]);
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

        updateSelectedUi();
        setTimeout(() => {
            searchEl.focus();
            if (initialQuery.length >= 2) {
                searchEl.value = initialQuery;
                runSearch(initialQuery);
            } else if (initialQuery) {
                searchEl.value = initialQuery;
                statusEl.textContent = 'Mindestens 2 Zeichen eingeben.';
            }
        }, 50);
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
 * Entra-Personensuche direkt in einen Host einbetten (kein zweites Modal).
 * Für Dialoge, die bereits als Overlay offen sind.
 *
 * @param {HTMLElement} host
 * @param {{
 *   multi?: boolean,
 *   hint?: string,
 *   initialQuery?: string,
 *   maxCount?: number,
 *   excludeEmails?: string[]|Set<string>,
 *   onPick?: (user: { id: string, displayName: string, mail: string, label: string }) => void,
 *   onApply?: (users: Array<{ id: string, displayName: string, mail: string, label: string }>) => void,
 *   onCancel?: () => void
 * }} [opts]
 * @returns {{ destroy: () => void, focus: () => void }}
 */
export function mountInlineEntraUserPicker(host, opts) {
    const options = opts && typeof opts === 'object' ? opts : {};
    const multi = !!options.multi;
    const hint =
        options.hint ||
        (multi
            ? 'Name suchen, Treffer anhaken und übernehmen – ohne Extra-Fenster.'
            : 'Name oder E-Mail suchen und Person übernehmen – ohne Extra-Fenster.');
    const initialQuery = String(options.initialQuery || '').trim();
    const maxCount =
        typeof options.maxCount === 'number' && options.maxCount > 0
            ? Math.floor(options.maxCount)
            : 0;
    /** @type {Set<string>} */
    const exclude = new Set();
    const exclIn = options.excludeEmails;
    if (exclIn) {
        (exclIn instanceof Set ? Array.from(exclIn) : Array.isArray(exclIn) ? exclIn : []).forEach(
            function (e) {
                const m = String(e || '')
                    .trim()
                    .toLowerCase();
                if (m.indexOf('@') !== -1) exclude.add(m);
            }
        );
    }

    if (!host || !(host instanceof HTMLElement)) {
        return {
            destroy: function () {},
            focus: function () {}
        };
    }

    const uid = 'inline-eup-' + String(Date.now() % 1e9);
    host.hidden = false;
    host.classList.add('ms365-inline-eup');
    host.innerHTML =
        '<div class="ms365-inline-eup__inner">' +
        '<p class="ms365-inline-eup__hint muted">' +
        escapeHtml(hint) +
        '</p>' +
        '<div class="field ms365-egp-search-wrap">' +
        '<label for="' +
        uid +
        '-search">Suche in Microsoft 365</label>' +
        '<div class="ms365-egp-search-row">' +
        '<i class="bi bi-search ms365-egp-search-icon" aria-hidden="true"></i>' +
        '<input type="search" id="' +
        uid +
        '-search" class="ms365-eup-search ms365-egp-search" autocomplete="off" placeholder="Name oder E-Mail …">' +
        '</div></div>' +
        '<p class="ms365-eup-status ms365-egp-status muted" role="status" aria-live="polite"></p>' +
        '<ul class="ms365-eup-results ms365-egp-results ms365-inline-eup__results" aria-label="Suchergebnisse"></ul>' +
        (multi
            ? '<p class="ms365-eup-selected muted ms365-inline-eup__selected" aria-live="polite"></p>'
            : '') +
        '<div class="ms365-inline-eup__actions">' +
        '<button type="button" class="btn btn-sm alt ms365-inline-eup-cancel">Schließen</button>' +
        (multi
            ? '<button type="button" class="btn btn-sm btn-success ms365-inline-eup-apply" disabled>Übernehmen</button>'
            : '') +
        '</div></div>';

    const searchEl = /** @type {HTMLInputElement|null} */ (host.querySelector('.ms365-eup-search'));
    const statusEl = host.querySelector('.ms365-eup-status');
    const listEl = host.querySelector('.ms365-eup-results');
    const selectedEl = host.querySelector('.ms365-eup-selected');
    const applyBtn = /** @type {HTMLButtonElement|null} */ (
        host.querySelector('.ms365-inline-eup-apply')
    );
    const cancelBtn = host.querySelector('.ms365-inline-eup-cancel');

    /** @type {Map<string, { id: string, displayName: string, mail: string, label: string }>} */
    const selected = new Map();
    let debounce = 0;
    let destroyed = false;

    function toPick(u) {
        const mail = String((u && (u.mail || u.userPrincipalName)) || '').trim();
        return {
            id: String((u && u.id) || ''),
            displayName: String((u && u.displayName) || ''),
            mail: mail,
            label: formatUserOptionLabel(u)
        };
    }

    /**
     * @param {'cancel'|'done'} [reason]
     */
    function destroy(reason) {
        if (destroyed) return;
        destroyed = true;
        clearTimeout(debounce);
        host.replaceChildren();
        host.hidden = true;
        host.classList.remove('ms365-inline-eup');
        if (reason === 'cancel' && typeof options.onCancel === 'function') options.onCancel();
    }

    function updateSelectedUi() {
        if (applyBtn) applyBtn.disabled = selected.size === 0;
        if (selectedEl) {
            selectedEl.textContent = selected.size
                ? selected.size +
                  ' ausgewählt: ' +
                  Array.from(selected.values())
                      .map(function (p) {
                          return p.displayName || p.mail;
                      })
                      .join(', ')
                : 'Noch niemand ausgewählt.';
        }
    }

    function renderResults(users) {
        if (!listEl || destroyed) return;
        listEl.replaceChildren();
        const filtered = (users || []).filter(function (u) {
            const mail = String((u && (u.mail || u.userPrincipalName)) || '')
                .trim()
                .toLowerCase();
            return !mail || !exclude.has(mail);
        });
        if (!filtered.length) {
            const li = document.createElement('li');
            li.className = 'ms365-egp-empty muted';
            li.textContent = 'Keine Treffer – Suchbegriff anpassen.';
            listEl.appendChild(li);
            return;
        }
        filtered.forEach(function (u) {
            const pick = toPick(u);
            if (!pick.id) return;
            const li = document.createElement('li');
            if (multi) {
                const row = document.createElement('label');
                row.className = 'ms365-egp-item ms365-eup-item ms365-eup-item--check';
                const cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.checked = selected.has(pick.id);
                cb.addEventListener('change', function () {
                    if (cb.checked) {
                        if (maxCount && selected.size >= maxCount && !selected.has(pick.id)) {
                            cb.checked = false;
                            if (statusEl) {
                                statusEl.textContent = 'Maximal ' + maxCount + ' Person(en) erlaubt.';
                            }
                            return;
                        }
                        selected.set(pick.id, pick);
                    } else {
                        selected.delete(pick.id);
                    }
                    updateSelectedUi();
                });
                const body = document.createElement('span');
                body.innerHTML =
                    '<span class="ms365-egp-item__name">' +
                    escapeHtml(pick.displayName) +
                    '</span>' +
                    '<span class="ms365-egp-item__meta muted">' +
                    escapeHtml(pick.mail || '–') +
                    '</span>';
                row.appendChild(cb);
                row.appendChild(body);
                li.appendChild(row);
            } else {
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'ms365-egp-item ms365-eup-item';
                btn.innerHTML =
                    '<span class="ms365-egp-item__name">' +
                    escapeHtml(pick.displayName) +
                    '</span>' +
                    '<span class="ms365-egp-item__meta muted">' +
                    escapeHtml(pick.mail || '–') +
                    '</span>';
                btn.addEventListener('click', function () {
                    if (typeof options.onPick === 'function') options.onPick(pick);
                    destroy('done');
                });
                li.appendChild(btn);
            }
            listEl.appendChild(li);
        });
    }

    async function runSearch(q) {
        if (!statusEl || destroyed) return;
        statusEl.textContent = 'Suche …';
        try {
            const list = await searchEntraUsers(q);
            if (destroyed) return;
            statusEl.textContent = list.length ? list.length + ' Treffer' : 'Keine Treffer';
            renderResults(list);
        } catch (e) {
            if (destroyed) return;
            statusEl.textContent = e && e.message ? e.message : String(e);
            renderResults([]);
        }
    }

    if (cancelBtn) {
        cancelBtn.addEventListener('click', function () {
            destroy('cancel');
        });
    }
    if (applyBtn) {
        applyBtn.addEventListener('click', function () {
            const list = Array.from(selected.values());
            if (typeof options.onApply === 'function') options.onApply(list);
            destroy('done');
        });
    }
    if (searchEl) {
        searchEl.addEventListener('input', function () {
            const q = String(searchEl.value || '').trim();
            clearTimeout(debounce);
            if (q.length < 2) {
                if (statusEl) statusEl.textContent = 'Mindestens 2 Zeichen eingeben.';
                if (listEl) listEl.replaceChildren();
                return;
            }
            debounce = window.setTimeout(function () {
                void runSearch(q);
            }, 280);
        });
    }

    if (multi) updateSelectedUi();
    setTimeout(function () {
        if (destroyed || !searchEl) return;
        searchEl.focus();
        if (initialQuery.length >= 2) {
            searchEl.value = initialQuery;
            void runSearch(initialQuery);
        } else if (initialQuery) {
            searchEl.value = initialQuery;
            if (statusEl) statusEl.textContent = 'Mindestens 2 Zeichen eingeben.';
        } else if (statusEl) {
            statusEl.textContent = 'Mindestens 2 Zeichen eingeben.';
        }
    }, 30);

    return {
        destroy: destroy,
        focus: function () {
            if (searchEl) searchEl.focus();
        }
    };
}
