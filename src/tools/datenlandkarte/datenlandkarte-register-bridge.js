/**
 * Datenlandkarte ↔ Stammdaten (Phase 5): Deep-Links und Sync-Ampeln.
 */

/** @type {'tool' | 'register'} */
let hrefContext = 'tool';

/**
 * @param {'tool' | 'register'} ctx
 */
export function setDatenlandkarteHrefContext(ctx) {
    hrefContext = ctx === 'register' ? 'register' : 'tool';
}

export function getDatenlandkarteHrefContext() {
    return hrefContext;
}

function registerHash(hash) {
    const h = String(hash || '').trim();
    const withHash = h.startsWith('#') ? h : h ? `#${h}` : '#stammdaten';
    return hrefContext === 'register' ? withHash : `../tenant.html${withHash}`;
}

/**
 * @param {string} [href]
 */
export function resolveDatenlandkarteHref(href) {
    const s = String(href || '').trim();
    if (!s) return '';
    if (s.startsWith('../tenant.html')) {
        const tail = s.slice('../tenant.html'.length);
        return registerHash(tail || 'stammdaten');
    }
    if (hrefContext === 'register' && s.startsWith('../')) {
        return s.replace(/^\.\.\//, '');
    }
    return s;
}

/** @type {Record<string, { stamm: string, spo: string }>} */
export const SYNC_LINK_METRICS = {
    'spo-sp-klassen-stamm': { stamm: 'classes', spo: 'spoListKlassen' },
    'spo-sp-teachers-stamm': { stamm: 'subjects', spo: 'spoListFaecher' },
    'spo-sp-schueler-stamm': { stamm: 'students', spo: 'spoListSchuelerinnen' }
};

const BLOCK_PRIMARY_HASH = {
    'stamm-hub': 'stammdaten',
    'stamm-classes': 'klassen',
    'stamm-teachers': 'lehrer',
    'stamm-subjects': 'faecher',
    'stamm-arges': 'arge',
    'year-students': 'schueler',
    'year-guardians': 'schueler',
    'import-webuntis': 'schueler',
    'spo-sp-klassen': 'stammdaten',
    'spo-sp-faecher': 'stammdaten',
    'spo-sp-schueler': 'stammdaten',
    'spo-lehrer': 'stammdaten'
};

const BLOCK_PRIMARY_TOOL = {
    'm365-catalog': '../tenant.html#stammdaten'
};

/**
 * @param {{ id: string, href?: string, layer?: string }} block
 */
export function primaryHrefForBlock(block) {
    if (!block) return '';
    if (BLOCK_PRIMARY_HASH[block.id]) return registerHash(BLOCK_PRIMARY_HASH[block.id]);
    if (BLOCK_PRIMARY_TOOL[block.id]) return resolveDatenlandkarteHref(BLOCK_PRIMARY_TOOL[block.id]);
    if (block.layer === 'sharepoint' && String(block.id || '').startsWith('spo-')) {
        return registerHash('stammdaten');
    }
    if (block.layer === 'stamm' || block.layer === 'schuljahr') {
        return resolveDatenlandkarteHref(block.href) || registerHash('stammdaten');
    }
    return resolveDatenlandkarteHref(block.href);
}

/**
 * @param {{ id: string, href?: string, layer?: string }} block
 * @returns {Array<{ label: string, href: string }>}
 */
export function actionsForBlock(block) {
    if (!block) return [];
    const out = [];
    const primary = primaryHrefForBlock(block);
    if (primary && (block.layer === 'stamm' || block.layer === 'schuljahr' || block.id === 'import-webuntis')) {
        out.push({
            label: hrefContext === 'register' ? 'Stammdaten-Tab öffnen' : 'Stammdaten öffnen',
            href: primary
        });
    } else if (primary && block.layer === 'sharepoint') {
        out.push({ label: 'Intranet synchronisieren', href: primary });
        const toolHref = resolveDatenlandkarteHref(block.href);
        if (toolHref && toolHref !== primary) {
            out.push({ label: 'Listen-Tool', href: toolHref });
        }
    } else if (block.href) {
        out.push({ label: 'Tool öffnen', href: resolveDatenlandkarteHref(block.href) });
    }
    if (block.id === 'import-webuntis') {
        out.push({
            label: 'WebUntis-Import',
            href: resolveDatenlandkarteHref('../tools/webuntis-stammdaten-import.html?from=tenant')
        });
    }
    if (block.id === 'year-unterricht' && block.href) {
        out.push({ label: 'Kursteams', href: resolveDatenlandkarteHref(block.href) });
    }
    return out;
}

function metricNumber(metrics, key) {
    const m = metrics && metrics[key];
    if (!m) return null;
    const v = m.value;
    if (typeof v === 'number' && isFinite(v)) return v;
    if (typeof v === 'string' && /^\d+$/.test(v.trim())) return Number(v.trim());
    return null;
}

/**
 * @param {string} linkId
 * @param {Record<string, { value: number|string, hint?: string }>} metrics
 * @returns {'ok'|'warn'|'unknown'|'na'}
 */
export function linkSyncHealth(linkId, metrics) {
    const pair = SYNC_LINK_METRICS[linkId];
    if (!pair) return 'na';
    const local = metricNumber(metrics, pair.stamm);
    const remote = metricNumber(metrics, pair.spo);
    if (local === null || remote === null) return 'unknown';
    if (local === 0 && remote === 0) return 'ok';
    if (local === remote) return 'ok';
    if (local > 0 && remote === 0) return 'warn';
    if (Math.abs(local - remote) <= 2) return 'ok';
    return 'warn';
}

export function linkHealthLabel(health) {
    if (health === 'ok') return { icon: 'bi-check-circle-fill', title: 'Zähler stimmen überein (lokal vs. SharePoint)' };
    if (health === 'warn') return { icon: 'bi-exclamation-triangle-fill', title: 'Abweichung – Intranet ggf. aktualisieren' };
    if (health === 'unknown') return { icon: 'bi-question-circle', title: 'SharePoint-Zähler noch nicht geladen' };
    return { icon: '', title: '' };
}
