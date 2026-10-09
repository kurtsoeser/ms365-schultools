/**
 * Entra-Benutzer suchen und per Dialog auswählen.
 */
import { getGraphPickerApi } from './graph-picker-backend.js';

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const USER_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All'
];

function graphApi() {
    return getGraphPickerApi();
}

function odataEscape(s) {
    return String(s || '').replace(/'/g, "''");
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
 * @param {string} queryRaw
 * @param {string} [token]
 */
export async function searchEntraUsers(queryRaw, token) {
    const q = String(queryRaw || '').trim();
    if (!q) return [];
    const tok = token || (await graphApi().getGraphToken(USER_SCOPES));
    const G = graphApi();
    const select = 'id,displayName,mail,userPrincipalName';

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

    const esc = odataEscape(q);
    const filter =
        "startswith(displayName,'" +
        esc +
        "') or startswith(mail,'" +
        esc +
        "') or startswith(userPrincipalName,'" +
        esc +
        "')";
    const path =
        '/users?$filter=' +
        encodeURIComponent(filter) +
        '&$select=' +
        encodeURIComponent(select) +
        '&$top=25';
    const data = await G.graphJson('GET', path, tok, undefined, 'v1.0');
    return Array.isArray(data.value) ? data.value : [];
}

/**
 * @param {{ title?: string, hint?: string }} [opts]
 * @returns {Promise<{ id: string, displayName: string, mail: string, label: string }|null>}
 */
export function pickEntraUser(opts) {
    const title = (opts && opts.title) || 'Person aus Microsoft Entra wählen';
    const hint =
        (opts && opts.hint) || 'Name oder E-Mail-Adresse (mindestens 2 Zeichen).';

    return new Promise((resolve) => {
        const overlay = document.createElement('div');
        overlay.className = 'modal-overlay ms365-entra-user-picker';
        overlay.setAttribute('role', 'dialog');
        overlay.setAttribute('aria-modal', 'true');
        overlay.innerHTML =
            '<div class="modal-box ms365-egp-modal">' +
            '<div class="ms365-egp-modal__head">' +
            '<span class="ms365-egp-modal__icon" aria-hidden="true"><i class="bi bi-envelope-at"></i></span>' +
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
