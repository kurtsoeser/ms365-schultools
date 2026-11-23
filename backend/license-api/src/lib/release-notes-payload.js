'use strict';

const MAX_NOTES = 40;
const MAX_IMAGES = 6;

function normalizeAudience(raw) {
    const a = String(raw || '')
        .trim()
        .toLowerCase();
    if (a === 'it' || a === 'schul-it' || a === 'schulit') return 'it';
    if (a === 'enduser' || a === 'endnutzer' || a === 'lehrer' || a === 'lehrkraft') return 'enduser';
    return 'all';
}

function normalizeKind(raw) {
    const k = String(raw || '')
        .trim()
        .toLowerCase();
    if (k === 'feature' || k === 'feat') return 'feature';
    if (k === 'fix' || k === 'bugfix' || k === 'bug') return 'fix';
    return 'other';
}

function sanitizeHtml(html) {
    return String(html || '')
        .replace(/<script[\s\S]*?>[\s\S]*?<\/script>/gi, '')
        .replace(/\son\w+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, '')
        .replace(/javascript:/gi, '');
}

function normalizeImage(raw, idx) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const src = String(o.src || '').trim();
    if (!src) return null;
    if (!/^(data:image\/(?:png|jpeg|jpg|webp|gif);base64,|https?:\/\/|release-notes\/)/i.test(src)) return null;
    if (src.length > 900_000) {
        const err = new Error('Ein Screenshot ist zu groß. Bitte ein kleineres Bild verwenden.');
        err.status = 413;
        throw err;
    }
    return {
        src: src,
        alt: String(o.alt || 'Screenshot ' + (idx + 1)).slice(0, 180)
    };
}

/**
 * @param {unknown} rawNotes
 * @returns {Array<object>}
 */
function normalizeReleaseNotes(rawNotes) {
    const list = Array.isArray(rawNotes) ? rawNotes : [];
    if (list.length > MAX_NOTES) {
        const err = new Error('Maximal ' + MAX_NOTES + ' Neuigkeiten.');
        err.status = 400;
        throw err;
    }
    return list
        .map((raw, idx) => {
            const o = raw && typeof raw === 'object' ? raw : {};
            const title = String(o.title || '').trim().slice(0, 180);
            if (!title) return null;
            const images = [];
            const rawImages = Array.isArray(o.images) ? o.images : [];
            for (let i = 0; i < rawImages.length && images.length < MAX_IMAGES; i++) {
                const img = normalizeImage(rawImages[i], i);
                if (img) images.push(img);
            }
            return {
                id: String(o.id || 'rn_' + idx).trim().slice(0, 80),
                at: String(o.at || new Date().toISOString()).slice(0, 40),
                title: title,
                bodyHtml: sanitizeHtml(o.bodyHtml).slice(0, 20000),
                body: String(o.body || '').slice(0, 20000),
                kind: normalizeKind(o.kind),
                source: 'manual',
                audience: normalizeAudience(o.audience),
                images: images
            };
        })
        .filter(Boolean);
}

/**
 * @param {Array<object>} notes
 * @param {'public'|'full'} view
 */
function notesForView(notes, view) {
    const list = Array.isArray(notes) ? notes : [];
    if (view === 'full') return list;
    return list.filter((n) => !n || !n.audience || n.audience === 'all');
}

module.exports = { normalizeReleaseNotes, notesForView, normalizeAudience };
