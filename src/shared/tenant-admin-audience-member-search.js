/**
 * Microsoft-365-Benutzersuche zum Hinzufügen von Personen zu Verwaltungs-Zielgruppen.
 */

function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function graphApi() {
    const G = window.ms365GraphUnifiedGroups;
    if (!G || typeof G.searchUsers !== 'function' || typeof G.getGraphToken !== 'function') return null;
    return G;
}

/**
 * @param {HTMLElement} host
 * @param {{ onSelect: (user: { email: string, displayName: string }) => void, setSummary?: (msg: string, kind?: string) => void }} opts
 */
export function mountAudienceMemberM365Search(host, opts) {
    if (!host || typeof opts.onSelect !== 'function') return;

    const wrap = document.createElement('div');
    wrap.className = 'ts-audience-board__m365-user-search';

    const searchIn = document.createElement('input');
    searchIn.type = 'search';
    searchIn.className = 'ts-audience-board__input ts-audience-board__m365-user-search-in';
    searchIn.placeholder = 'Microsoft 365: Name oder E-Mail suchen …';
    searchIn.autocomplete = 'off';
    searchIn.setAttribute('aria-label', 'Person in Microsoft 365 suchen');

    const btnGo = document.createElement('button');
    btnGo.type = 'button';
    btnGo.className = 'btn btn-sm';
    btnGo.title = 'Suchen';
    btnGo.innerHTML = '<i class="bi bi-search" aria-hidden="true"></i>';

    const hits = document.createElement('ul');
    hits.className = 'ts-audience-board__m365-user-hits';
    hits.hidden = true;

    let timer = null;
    let inFlight = 0;

    function notify(msg, kind) {
        if (typeof opts.setSummary === 'function') opts.setSummary(msg, kind);
    }

    function pickUser(u) {
        const mail = String(u.mail || u.userPrincipalName || '')
            .trim()
            .toLowerCase();
        const name = String(u.displayName || '').trim();
        if (!mail || mail.indexOf('@') === -1) return;
        hits.hidden = true;
        hits.replaceChildren();
        opts.onSelect({ email: mail, displayName: name });
        searchIn.value = name ? name + ' · ' + mail : mail;
    }

    function renderHits(users) {
        hits.replaceChildren();
        if (!users.length) {
            const li = document.createElement('li');
            li.className = 'muted';
            li.textContent = 'Keine Treffer';
            hits.appendChild(li);
            hits.hidden = false;
            return;
        }
        users.slice(0, 12).forEach(function (u) {
            const mail = String(u.mail || u.userPrincipalName || '').trim();
            const name = String(u.displayName || mail).trim();
            const li = document.createElement('li');
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'ts-audience-board__m365-user-hit';
            btn.innerHTML = '<strong>' + escapeHtml(name) + '</strong> <span class="muted">' + escapeHtml(mail) + '</span>';
            btn.addEventListener('click', function () {
                pickUser(u);
            });
            li.appendChild(btn);
            hits.appendChild(li);
        });
        hits.hidden = false;
    }

    async function runSearch() {
        const q = String(searchIn.value || '').trim();
        if (q.length < 2) {
            hits.hidden = true;
            hits.replaceChildren();
            return;
        }
        const G = graphApi();
        if (!G) {
            notify('Microsoft-365-Suche: Graph-Modul nicht geladen.', 'warn');
            return;
        }
        const flight = ++inFlight;
        btnGo.disabled = true;
        try {
            const token = await G.getGraphToken();
            const users = await G.searchUsers(token, q);
            if (flight !== inFlight) return;
            renderHits(users || []);
        } catch (e) {
            if (flight === inFlight) {
                notify('Suche: ' + (e && e.message ? e.message : String(e)), 'warn');
                hits.hidden = true;
            }
        } finally {
            if (flight === inFlight) btnGo.disabled = false;
        }
    }

    searchIn.addEventListener('input', function () {
        clearTimeout(timer);
        timer = setTimeout(runSearch, 320);
    });
    searchIn.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') {
            e.preventDefault();
            clearTimeout(timer);
            runSearch();
        } else if (e.key === 'Escape') {
            hits.hidden = true;
        }
    });
    btnGo.addEventListener('click', function () {
        clearTimeout(timer);
        runSearch();
    });

    document.addEventListener(
        'click',
        function (ev) {
            if (!wrap.contains(ev.target)) hits.hidden = true;
        },
        true
    );

    wrap.append(searchIn, btnGo, hits);
    host.appendChild(wrap);
}
