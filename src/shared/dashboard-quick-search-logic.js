/**
 * Ranking & Matching für die IT-Schnellsuche (Admin-Center-Stil).
 * Kein Volltext über Status-Badges – nur Titel, ID, Keywords, Kurztext.
 */

/** @typedef {'werkzeug'|'bereich'|'playbook'|'einstellung'} QuickSearchKind */

/**
 * @typedef {{
 *   id: string,
 *   title: string,
 *   blurb?: string,
 *   keywords?: string,
 *   href?: string,
 *   kind: QuickSearchKind,
 *   group?: string,
 *   icon?: string,
 *   taskId?: string,
 *   toolId?: string
 * }} QuickSearchEntry
 */

/**
 * @typedef {{
 *   entry: QuickSearchEntry,
 *   score: number,
 *   match: 'exact'|'prefix'|'token'|'keyword'|'blurb'
 * }} QuickSearchHit
 */

/** @param {string} s */
export function normalizeSearchText(s) {
    return String(s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/ä/g, 'ae')
        .replace(/ö/g, 'oe')
        .replace(/ü/g, 'ue')
        .replace(/ß/g, 'ss')
        .replace(/[^a-z0-9]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

/** @param {string} s */
export function tokenizeSearchText(s) {
    const n = normalizeSearchText(s);
    return n ? n.split(' ').filter(Boolean) : [];
}

/**
 * Kurze Codes / Aliase: stark gewichten, wenn Query exakt so lautet.
 * @type {Record<string, string[]>}
 */
export const QUICK_SEARCH_ALIASES = {
    srdp: ['srdp-anmeldung', 'srdp-mail-setup', 'sharepoint-liste-srdp'],
    slg: ['slg-schueler', 'slg-lehrer'],
    kv: ['klassenvorstaende'],
    matura: ['srdp-anmeldung', 'srdp-mail-setup', 'sharepoint-liste-srdp'],
    webuntis: ['webuntis-stammdaten-import', 'stammdaten-quelle-waehlen'],
    freistellung: ['freistellung-planer', 'lehrer-freistellung-planer', 'playbook-freistellungen'],
    dashboard: ['einstellung-portal-vorbereiten', 'einstellung-portal'],
    portal: ['einstellung-portal-vorbereiten', 'einstellung-portal'],
    'portal vorbereiten': ['einstellung-portal-vorbereiten'],
    portalvorbereiten: ['einstellung-portal-vorbereiten'],
    werkzeugzugriff: ['einstellung-portal-vorbereiten'],
    'werkzeug zugriff': ['einstellung-portal-vorbereiten']
};

/**
 * @param {QuickSearchEntry} entry
 * @param {string} rawQuery
 * @returns {QuickSearchHit | null}
 */
export function scoreQuickSearchEntry(entry, rawQuery) {
    const q = normalizeSearchText(rawQuery);
    if (!q || !entry) return null;

    const title = normalizeSearchText(entry.title);
    const id = normalizeSearchText(String(entry.id || entry.toolId || '').replace(/-/g, ' '));
    const idCompact = normalizeSearchText(String(entry.id || entry.toolId || '').replace(/-/g, ''));
    const keywords = normalizeSearchText(entry.keywords || '');
    const blurb = normalizeSearchText(entry.blurb || '');
    const qCompact = q.replace(/\s+/g, '');

    let score = 0;
    /** @type {QuickSearchHit['match']} */
    let match = 'keyword';

    if (title === q) {
        score = 100;
        match = 'exact';
    } else if (title.startsWith(q + ' ') || title.startsWith(q)) {
        score = 88;
        match = 'prefix';
    } else if (tokenizeSearchText(entry.title).some(function (t) {
        return t.startsWith(q);
    })) {
        score = 78;
        match = 'token';
    }

    if (id === q || idCompact === qCompact) {
        score = Math.max(score, 92);
        match = score >= 92 ? 'exact' : match;
    } else if (id.startsWith(q) || idCompact.startsWith(qCompact)) {
        score = Math.max(score, 82);
        if (score === 82) match = 'prefix';
    }

    const kwTokens = tokenizeSearchText(entry.keywords || '');
    if (kwTokens.indexOf(q) !== -1 || keywords.split(' ').indexOf(q) !== -1) {
        score = Math.max(score, 70);
        if (score === 70) match = 'keyword';
    } else if (keywords && (' ' + keywords + ' ').indexOf(' ' + q + ' ') !== -1) {
        score = Math.max(score, 66);
        if (score === 66) match = 'keyword';
    } else if (keywords.indexOf(q) !== -1) {
        // Substring in Keywords nur bei Query-Länge >= 3, sonst zu laut
        if (q.length >= 3) {
            score = Math.max(score, 48);
            if (score === 48) match = 'keyword';
        }
    }

    if (blurb && q.length >= 4 && blurb.indexOf(q) !== -1) {
        score = Math.max(score, 28);
        if (score === 28) match = 'blurb';
    }

    const aliasIds = QUICK_SEARCH_ALIASES[q] || QUICK_SEARCH_ALIASES[qCompact] || null;
    if (aliasIds && aliasIds.indexOf(String(entry.toolId || entry.id)) !== -1) {
        score = Math.max(score, 95);
        match = 'exact';
    }

    // Playbooks leicht bevorzugen, wenn Titel/Alias trifft
    if (entry.kind === 'playbook' && score >= 70) score += 2;
    // Bereiche nur zeigen, wenn sie wirklich matchen – nicht als Noise
    if (entry.kind === 'bereich' && score < 70) score -= 8;

    if (score < 40) return null;
    return { entry: entry, score: score, match: match };
}

/**
 * @param {QuickSearchEntry[]} entries
 * @param {string} rawQuery
 * @param {{ limit?: number }} [opts]
 * @returns {QuickSearchHit[]}
 */
export function rankQuickSearch(entries, rawQuery, opts) {
    const limit = opts && opts.limit != null ? opts.limit : 12;
    const hits = [];
    (entries || []).forEach(function (entry) {
        const hit = scoreQuickSearchEntry(entry, rawQuery);
        if (hit) hits.push(hit);
    });
    hits.sort(function (a, b) {
        if (b.score !== a.score) return b.score - a.score;
        return String(a.entry.title).localeCompare(String(b.entry.title), 'de');
    });
    return hits.slice(0, Math.max(1, limit));
}

/**
 * Treffer nach Kategorie gruppieren (Admin-Center-Reihenfolge).
 * @param {QuickSearchHit[]} hits
 * @returns {{ kind: QuickSearchKind, label: string, hits: QuickSearchHit[] }[]}
 */
export function groupQuickSearchHits(hits) {
    /** @type {Record<QuickSearchKind, string>} */
    const labels = {
        werkzeug: 'Werkzeuge',
        bereich: 'Bereiche',
        playbook: 'Playbooks',
        einstellung: 'Einstellungen'
    };
    /** @type {QuickSearchKind[]} */
    const order = ['werkzeug', 'playbook', 'bereich', 'einstellung'];
    return order
        .map(function (kind) {
            return {
                kind: kind,
                label: labels[kind],
                hits: (hits || []).filter(function (h) {
                    return h.entry.kind === kind;
                })
            };
        })
        .filter(function (g) {
            return g.hits.length > 0;
        });
}

/**
 * Katalog-Karte → Such-Eintrag (ohne Status-Text aus dem DOM).
 * @param {Element} card
 * @returns {QuickSearchEntry | null}
 */
export function entryFromCatalogCard(card) {
    if (!card) return null;
    const toolId = String(card.getAttribute('data-tool-id') || '').trim();
    if (!toolId) return null;
    const titleEl = card.querySelector('h2');
    let title = titleEl
        ? String(titleEl.textContent || '')
              .replace(/\s+/g, ' ')
              .trim()
        : '';
    if (!title) title = String(card.getAttribute('aria-label') || toolId).trim();
    const p = card.querySelector('p');
    const blurb = p
        ? String(p.textContent || '')
              .replace(/\s+/g, ' ')
              .trim()
        : '';
    const link = card.querySelector('a.btn[href], a[href]');
    const href = link ? String(link.getAttribute('href') || '').trim() : '';
    const cluster = String(card.getAttribute('data-cluster') || '').trim();
    const keywords = String(card.getAttribute('data-search') || '').trim();
    const iconEl = titleEl && titleEl.querySelector('i.bi');
    const isPlaybook = /^playbook-/.test(toolId) || /playbook/i.test(title);
    return {
        id: toolId,
        toolId: toolId,
        title: title,
        blurb: blurb,
        keywords: keywords,
        href: href,
        kind: isPlaybook ? 'playbook' : 'werkzeug',
        group: cluster,
        icon: iconEl && iconEl.className ? String(iconEl.className) : 'bi bi-box-arrow-up-right'
    };
}

/**
 * Ob eine Katalog-Karte zur Query passt (Seite filtern – gleiche Logik wie Ranking-Schwelle).
 * @param {Element} card
 * @param {string} needle
 */
export function catalogCardMatchesQuery(card, needle) {
    if (!needle) return true;
    const entry = entryFromCatalogCard(card);
    if (!entry) return false;
    return !!scoreQuickSearchEntry(entry, needle);
}
