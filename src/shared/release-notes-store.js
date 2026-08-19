/**
 * Lokale Release-Notes (ohne Server).
 * - Admin kann Einträge hinzufügen/ändern (localStorage).
 * - User sehen beim nächsten Öffnen der App neue Einträge (per last-seen timestamp).
 */

export const RELEASE_NOTES_KEY = 'ms365-schooltool-release-notes-v1';
export const RELEASE_NOTES_LAST_SEEN_AT_KEY = 'ms365-schooltool-release-notes-last-seen-at-v1';

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

function normalizeNote(raw, idx) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const id = String(o.id || '').trim() || String('n_' + idx + '_' + Math.random().toString(16).slice(2));
    const at = normalizeIsoDate(o.at) || nowIso();
    const title = String(o.title || '').trim();
    const body = String(o.body || '').trim();
    return { id: id, at: at, title: title, body: body };
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
        // ignore
    }
    return normalized;
}

export function appendReleaseNote(note, storage = localStorage) {
    const existing = loadReleaseNotes(storage);
    const next = existing.slice();
    next.push(normalizeNote(note, next.length));
    return saveReleaseNotes(next, storage);
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

