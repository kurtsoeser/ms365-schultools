/**
 * Entra-Abgleich / Anlegen / Entfernen – gleiche Aktionsleiste wie Verwaltungs-Rollentabelle.
 */

function normEmail(emailRaw) {
    return String(emailRaw ?? '').trim().toLowerCase();
}

function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

export function getDirectoryMatchByEmail(emailRaw) {
    const em = normEmail(emailRaw);
    if (!em) return null;
    const api = window.ms365AppDataV2;
    const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
    const map = setup && setup.directoryMatchByEmail ? setup.directoryMatchByEmail : {};
    return map[em] || null;
}

/**
 * @param {HTMLElement} el
 * @param {string} emailRaw
 */
export function paintDirectoryMatchCell(el, emailRaw) {
    if (!el) return;
    el.className = 'ts-audience-board__member-cell ts-audience-board__member-ms';
    const em = normEmail(emailRaw);
    const m = em && em.indexOf('@') !== -1 ? getDirectoryMatchByEmail(em) : null;
    el.style.background = '';
    if (!em || em.indexOf('@') === -1) {
        el.style.color = 'var(--muted)';
        el.textContent = '–';
        el.title = 'E-Mail nötig für Abgleich mit Microsoft Entra';
        return;
    }
    if (m && m.graphUserId) {
        const gid = String(m.graphUserId);
        const short = gid.length > 14 ? gid.slice(0, 12) + '…' : gid;
        el.innerHTML =
            '<span class="ts-audience-board__ms-ok">✓</span> <code>' + escapeHtml(short) + '</code>';
        el.title =
            (m.displayName ? m.displayName : '') +
            (m.userPrincipalName ? '\n' + m.userPrincipalName : '') +
            '\nObject-ID: ' + gid;
        el.style.background = 'color-mix(in srgb, #0d8050 8%, transparent)';
        return;
    }
    if (m && m.notFound) {
        el.innerHTML =
            '<span class="ts-audience-board__ms-warn">✗</span> <span class="muted">nicht gefunden</span>';
        el.title = 'Kein Benutzer mit mail oder UPN gleich dieser E-Mail';
        return;
    }
    el.style.color = 'var(--muted)';
    el.textContent = '–';
    el.title = 'Noch nicht geprüft – „Prüfen“ in Aktion';
}

/**
 * @param {HTMLElement} parent
 * @param {{
 *   email: string,
 *   personName?: string,
 *   onVerify?: (email: string) => Promise<void>,
 *   onCreateUser?: (email: string, name: string) => Promise<void>,
 *   onRemove?: () => void,
 *   removeTitle?: string,
 * }} opts
 */
export function appendAdminDirectoryActionButtons(parent, opts) {
    const em = normEmail(opts.email);
    const name = String(opts.personName ?? '').trim();
    const wrap = document.createElement('div');
    wrap.className = 'ts-audience-board__member-actions';

    const dir = em && em.indexOf('@') !== -1 ? getDirectoryMatchByEmail(em) : null;

    if (em && em.indexOf('@') !== -1 && typeof opts.onVerify === 'function') {
        const btnCheck = document.createElement('button');
        btnCheck.type = 'button';
        btnCheck.className = 'ts-audience-board__dir-btn ts-audience-board__dir-btn--brand';
        btnCheck.title = 'Diese E-Mail in Microsoft Entra prüfen';
        btnCheck.setAttribute('aria-label', btnCheck.title);
        btnCheck.innerHTML = '<i class="bi bi-microsoft" aria-hidden="true"></i>';
        if (dir && dir.graphUserId) btnCheck.disabled = true;
        else {
            btnCheck.addEventListener('click', async function () {
                btnCheck.disabled = true;
                try {
                    await opts.onVerify(em);
                } finally {
                    btnCheck.disabled = false;
                }
            });
        }
        wrap.appendChild(btnCheck);
    }

    if (em && em.indexOf('@') !== -1 && name && typeof opts.onCreateUser === 'function') {
        const btnCreate = document.createElement('button');
        btnCreate.type = 'button';
        btnCreate.className = 'ts-audience-board__dir-btn ts-audience-board__dir-btn--create';
        btnCreate.title = 'Neuen Benutzer in Microsoft Entra ID anlegen';
        btnCreate.setAttribute('aria-label', btnCreate.title);
        btnCreate.innerHTML = '<i class="bi bi-person-plus" aria-hidden="true"></i>';
        if (dir && dir.graphUserId) btnCreate.disabled = true;
        else {
            btnCreate.addEventListener('click', async function () {
                btnCreate.disabled = true;
                try {
                    await opts.onCreateUser(em, name);
                } finally {
                    btnCreate.disabled = false;
                }
            });
        }
        wrap.appendChild(btnCreate);
    }

    if (typeof opts.onRemove === 'function') {
        const btnDel = document.createElement('button');
        btnDel.type = 'button';
        btnDel.className = 'ts-audience-board__dir-btn ts-audience-board__dir-btn--danger';
        btnDel.textContent = '✕';
        btnDel.title = opts.removeTitle || 'Entfernen';
        btnDel.setAttribute('aria-label', btnDel.title);
        btnDel.addEventListener('click', function () {
            opts.onRemove();
        });
        wrap.appendChild(btnDel);
    }

    parent.appendChild(wrap);
    return wrap;
}
