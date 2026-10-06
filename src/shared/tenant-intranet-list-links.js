/**
 * Zuletzt abgeglichene SharePoint-Intranet-Listen (Schnell-Link „Liste öffnen“).
 */
const STORAGE_KEY = 'ms365-intranet-list-links-v1';

export function readIntranetListLinks() {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        const parsed = raw ? JSON.parse(raw) : {};
        return parsed && typeof parsed === 'object' ? parsed : {};
    } catch {
        return {};
    }
}

/**
 * @param {string} kind
 * @param {{ url: string, title?: string, count?: number }} entry
 */
export function saveIntranetListLink(kind, entry) {
    const k = String(kind || '').trim();
    const url = entry && entry.url ? String(entry.url).trim() : '';
    if (!k || !url) return;
    const map = readIntranetListLinks();
    map[k] = {
        url: url,
        title: entry.title ? String(entry.title).trim() : '',
        count: entry.count != null ? Number(entry.count) : null,
        at: new Date().toISOString()
    };
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(map));
    } catch {
        /* ignore */
    }
}

export function refreshIntranetListOpenLinks() {
    const stored = readIntranetListLinks();
    document.querySelectorAll('[data-intranet-list-open]').forEach(function (a) {
        const kind = String(a.getAttribute('data-intranet-list-open') || '').trim();
        const row = stored[kind];
        if (row && row.url) {
            a.href = row.url;
            a.hidden = false;
            const label = row.title || kind;
            a.title = 'SharePoint-Liste „' + label + '“ öffnen';
            if (row.at) {
                a.setAttribute('data-intranet-list-synced-at', row.at);
            }
        } else {
            a.hidden = true;
            a.removeAttribute('href');
        }
    });
}

export function wrapIntranetSyncButtons() {
    document.querySelectorAll('[data-tenant-intranet-sync]').forEach(function (btn) {
        const parent = btn.parentElement;
        if (parent && parent.classList && parent.classList.contains('ts-intranet-sync-wrap')) {
            return;
        }
        const kind = String(btn.getAttribute('data-tenant-intranet-sync') || '').trim();
        const wrap = document.createElement('span');
        wrap.className = 'ts-intranet-sync-wrap';
        const link = document.createElement('a');
        link.className = 'ts-intranet-list-open';
        link.setAttribute('data-intranet-list-open', kind);
        link.target = '_blank';
        link.rel = 'noopener';
        link.hidden = true;
        link.innerHTML = '<i class="bi bi-box-arrow-up-right" aria-hidden="true"></i> Liste öffnen';
        if (parent) {
            parent.insertBefore(wrap, btn);
            wrap.appendChild(btn);
            wrap.appendChild(link);
        }
    });
    refreshIntranetListOpenLinks();
}

export function showIntranetSyncToast(listTitle, count, hasListLink) {
    const title = String(listTitle || 'Liste').trim();
    const parts = ['„' + title + '“ wurde auf SharePoint abgeglichen.'];
    if (count != null && !Number.isNaN(Number(count))) {
        parts.push(String(count) + ' Zeilen.');
    }
    if (hasListLink) {
        parts.push('Neben dem Button: „Liste öffnen“.');
    }
    const msg = parts.join(' ');
    if (typeof window.ms365ShowToast === 'function') {
        window.ms365ShowToast(msg, { kind: 'success', title: 'Intranet-Liste', durationMs: 9000 });
        return;
    }
    if (typeof window.ms365ToastOrAlert === 'function') {
        window.ms365ToastOrAlert(msg);
    } else {
        window.alert(msg);
    }
}
