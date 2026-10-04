/**
 * Entra-Benutzer suchen und per Dialog auswählen.
 */
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const USER_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All'
];

function graphApi() {
    const g = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!g || typeof g.getGraphToken !== 'function' || typeof g.graphJson !== 'function') {
        throw new Error('Graph-Hilfen nicht geladen (spo-graph-shared).');
    }
    return g;
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
            '<div class="modal-box" style="max-width:520px;width:min(96vw,520px);">' +
            '<h3 class="modal-box__title" style="margin:0 0 8px;">' +
            escapeHtml(title) +
            '</h3>' +
            '<p class="muted" style="margin:0 0 12px;font-size:0.9rem;">' +
            escapeHtml(hint) +
            '</p>' +
            '<div class="tm-field"><label>Suche</label>' +
            '<input type="search" class="ms365-eup-search" autocomplete="off" placeholder="z. B. Sekretariat, name@schule.at …"></div>' +
            '<p class="ms365-eup-status muted" style="min-height:1.2em;margin:8px 0;font-size:0.88em;"></p>' +
            '<ul class="ms365-eup-results" style="list-style:none;margin:0;padding:0;max-height:280px;overflow:auto;border:1px solid var(--border,#ddd);border-radius:8px;"></ul>' +
            '<div class="modal-box__actions" style="margin-top:14px;display:flex;gap:8px;justify-content:flex-end;">' +
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
                li.className = 'muted';
                li.style.padding = '12px';
                li.textContent = 'Keine Treffer – Suchbegriff anpassen.';
                listEl.appendChild(li);
                return;
            }
            users.forEach((u) => {
                const li = document.createElement('li');
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'ms365-eup-item';
                btn.style.cssText =
                    'display:block;width:100%;text-align:left;padding:10px 12px;border:0;border-bottom:1px solid var(--border,#eee);background:transparent;cursor:pointer;font:inherit;';
                const mail = String(u.mail || u.userPrincipalName || '').trim();
                btn.innerHTML =
                    '<strong>' +
                    escapeHtml(String(u.displayName || '')) +
                    '</strong><br><span class="muted" style="font-size:0.85em;">' +
                    escapeHtml(formatUserOptionLabel(u)) +
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
