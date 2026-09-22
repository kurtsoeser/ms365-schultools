/**
 * Release-Notes: lokal + veröffentlicht (public/release-notes.json).
 * Felder: id, at, title, bodyHtml (bevorzugt), body (Plaintext-Fallback),
 * kind: feature|fix|other, source: manual|github|local,
 * images: [{ src, alt }], gitSha (optional, für Auto-Dedup).
 */

export const RELEASE_NOTES_KEY = 'ms365-schooltool-release-notes-v1';
export const RELEASE_NOTES_LAST_SEEN_AT_KEY = 'ms365-schooltool-release-notes-last-seen-at-v1';

const ALLOWED_TAGS = new Set([
    'P',
    'BR',
    'STRONG',
    'B',
    'EM',
    'I',
    'U',
    'UL',
    'OL',
    'LI',
    'A',
    'DIV',
    'SPAN'
]);

function safeJsonParse(raw) {
    try {
        return JSON.parse(String(raw));
    } catch {
        return null;
    }
}

function nowIso() {
    try {
        return new Date().toISOString();
    } catch {
        return '';
    }
}

function normalizeIsoDate(raw) {
    const s = String(raw == null ? '' : raw).trim();
    if (!s) return '';
    const d = new Date(s);
    if (Number.isNaN(d.getTime())) return '';
    return d.toISOString();
}

function escapeHtml(s) {
    return String(s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/** Plaintext → einfaches HTML (Absätze). */
export function plainTextToHtml(text) {
    const t = String(text || '').trim();
    if (!t) return '';
    return t
        .split(/\n{2,}/)
        .map((block) => '<p>' + escapeHtml(block).replace(/\n/g, '<br>') + '</p>')
        .join('');
}

/**
 * Erlaubt nur einfache Formatierung (kein Script/Style).
 * @param {string} dirty
 */
export function sanitizeReleaseHtml(dirty) {
    const raw = String(dirty || '').trim();
    if (!raw) return '';
    if (typeof document === 'undefined') {
        return raw
            .replace(/<script[\s\S]*?>[\s\S]*?<\/script>/gi, '')
            .replace(/on\w+\s*=\s*(['"]).*?\1/gi, '')
            .replace(/<\/?(?:img|iframe|object|embed|link|meta|style)[^>]*>/gi, '');
    }
    const tpl = document.createElement('template');
    tpl.innerHTML = raw;
    const walk = (node) => {
        const children = Array.from(node.childNodes);
        for (const child of children) {
            if (child.nodeType === 3) continue;
            if (child.nodeType !== 1) {
                child.remove();
                continue;
            }
            const el = /** @type {HTMLElement} */ (child);
            const tag = el.tagName;
            if (!ALLOWED_TAGS.has(tag)) {
                const parent = el.parentNode;
                while (el.firstChild) parent.insertBefore(el.firstChild, el);
                el.remove();
                continue;
            }
            [...el.attributes].forEach((attr) => {
                const name = attr.name.toLowerCase();
                if (tag === 'A' && name === 'href') {
                    const href = String(attr.value || '').trim();
                    if (!/^(https?:|mailto:|#)/i.test(href)) el.removeAttribute(attr.name);
                    else {
                        el.setAttribute('rel', 'noopener noreferrer');
                        if (/^https?:/i.test(href)) el.setAttribute('target', '_blank');
                    }
                    return;
                }
                el.removeAttribute(attr.name);
            });
            walk(el);
        }
    };
    walk(tpl.content);
    return tpl.innerHTML.trim();
}

function normalizeImage(raw, idx) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const src = String(o.src || o.url || '').trim();
    if (!src) return null;
    if (!/^(data:image\/|https?:\/\/|\.?\.?\/|release-notes\/)/i.test(src)) return null;
    return {
        src: src,
        alt: String(o.alt || o.caption || 'Screenshot ' + (idx + 1)).trim()
    };
}

function normalizeKind(raw) {
    const k = String(raw || '')
        .trim()
        .toLowerCase();
    if (k === 'feature' || k === 'feat') return 'feature';
    if (k === 'fix' || k === 'bugfix' || k === 'bug') return 'fix';
    return 'other';
}

function normalizeSource(raw) {
    const s = String(raw || '')
        .trim()
        .toLowerCase();
    if (s === 'github' || s === 'auto') return 'github';
    if (s === 'local') return 'local';
    return 'manual';
}

export function normalizeNote(raw, idx) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const id = String(o.id || '').trim() || String('n_' + idx + '_' + Math.random().toString(16).slice(2));
    const at = normalizeIsoDate(o.at) || nowIso();
    const title = String(o.title || '').trim();
    let bodyHtml = String(o.bodyHtml || '').trim();
    const body = String(o.body || '').trim();
    if (!bodyHtml && body) bodyHtml = plainTextToHtml(body);
    bodyHtml = sanitizeReleaseHtml(bodyHtml);
    const images = (Array.isArray(o.images) ? o.images : [])
        .map((img, i) => normalizeImage(img, i))
        .filter(Boolean)
        .slice(0, 6);
    const note = {
        id: id,
        at: at,
        title: title,
        bodyHtml: bodyHtml,
        body: body,
        kind: normalizeKind(o.kind),
        source: normalizeSource(o.source),
        images: images
    };
    const gitSha = String(o.gitSha || '').trim();
    if (gitSha) note.gitSha = gitSha;
    return note;
}

export function loadReleaseNotes(storage = localStorage) {
    try {
        const raw = storage.getItem(RELEASE_NOTES_KEY);
        if (!raw) return [];
        const parsed = safeJsonParse(raw);
        if (!parsed) return [];
        const arr = Array.isArray(parsed) ? parsed : Array.isArray(parsed.notes) ? parsed.notes : [];
        return arr.map((n, i) => normalizeNote(n, i)).sort((a, b) => (a.at < b.at ? 1 : -1));
    } catch {
        return [];
    }
}

export function saveReleaseNotes(notes, storage = localStorage) {
    const arr = Array.isArray(notes) ? notes : [];
    const normalized = arr.map((n, i) => normalizeNote(n, i));
    try {
        storage.setItem(RELEASE_NOTES_KEY, JSON.stringify(normalized));
    } catch {
        // ignore (Quota)
    }
    return normalized;
}

export function appendReleaseNote(note, storage = localStorage) {
    const existing = loadReleaseNotes(storage);
    const next = existing.slice();
    next.push(normalizeNote(note, next.length));
    return saveReleaseNotes(next, storage);
}

export function updateReleaseNote(id, patch, storage = localStorage) {
    const want = String(id || '').trim();
    const existing = loadReleaseNotes(storage);
    const next = existing.map((n, i) => {
        if (n.id !== want) return n;
        return normalizeNote(Object.assign({}, n, patch, { id: n.id }), i);
    });
    return saveReleaseNotes(next, storage);
}

export function deleteReleaseNote(id, storage = localStorage) {
    const want = String(id || '').trim();
    return saveReleaseNotes(
        loadReleaseNotes(storage).filter((n) => n.id !== want),
        storage
    );
}

export function getLastSeenAt(storage = localStorage) {
    try {
        const raw = storage.getItem(RELEASE_NOTES_LAST_SEEN_AT_KEY);
        if (!raw) return '';
        return normalizeIsoDate(raw) || '';
    } catch {
        return '';
    }
}

export function setLastSeenAt(atIso, storage = localStorage) {
    const value = normalizeIsoDate(atIso);
    try {
        if (value) storage.setItem(RELEASE_NOTES_LAST_SEEN_AT_KEY, value);
        else storage.removeItem(RELEASE_NOTES_LAST_SEEN_AT_KEY);
    } catch {
        // ignore
    }
    return value;
}

export function getNewReleaseNotes({ notes, lastSeenAtIso }) {
    const lastSeen = normalizeIsoDate(lastSeenAtIso);
    const lastTime = lastSeen ? new Date(lastSeen).getTime() : 0;
    return (Array.isArray(notes) ? notes : []).filter((n) => {
        const t = new Date(n.at).getTime();
        return !Number.isNaN(t) && t > lastTime;
    });
}

/**
 * Veröffentlichtes JSON (GitHub Pages) laden und mit lokalen Notizen mergen.
 * Lokal gewinnt bei gleicher id.
 */
export async function loadMergedReleaseNotes(storage = localStorage, fetchImpl = fetch) {
    const local = loadReleaseNotes(storage);
    let published = [];
    try {
        const url = new URL('release-notes.json', typeof window !== 'undefined' ? window.location.href : 'http://local/');
        const res = await fetchImpl(url.href, { cache: 'no-cache' });
        if (res.ok) {
            const data = await res.json();
            const arr = Array.isArray(data) ? data : Array.isArray(data && data.notes) ? data.notes : [];
            published = arr.map((n, i) => normalizeNote(n, i));
        }
    } catch {
        published = [];
    }

    const byId = new Map();
    published.forEach((n) => byId.set(n.id, n));
    local.forEach((n) => byId.set(n.id, n));
    return Array.from(byId.values()).sort((a, b) => (a.at < b.at ? 1 : -1));
}

/** Export-Format für public/release-notes.json */
export function toPublishedJson(notes) {
    return JSON.stringify(
        (Array.isArray(notes) ? notes : []).map((n, i) => normalizeNote(n, i)),
        null,
        2
    );
}
