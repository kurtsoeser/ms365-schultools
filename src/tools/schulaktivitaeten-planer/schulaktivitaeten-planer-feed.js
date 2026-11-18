/**
 * ICS-Feed auf der SharePoint-Site veröffentlichen (abonnierbar für Outlook u. a.).
 */
import { buildIcs, toWebcalUrl } from './schulaktivitaeten-planer-export.js';

export const FEED_REL_PATH = 'MS365-Schultools/kalender/schulaktivitaeten.ics';
export const FEED_META_KEY = 'ms365-akt-ics-feed-v1';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

function G() {
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
}

function encodeDriveRootPath(relativePath) {
    const parts = String(relativePath || '')
        .replace(/^\/+/, '')
        .split('/')
        .filter(Boolean);
    if (!parts.length) return 'root:';
    return 'root:/' + parts.map((p) => encodeURIComponent(p)).join('/') + ':';
}

/**
 * @returns {{ httpsUrl: string, webcalUrl: string, path: string, publishedAt: string, count: number }|null}
 */
export function loadFeedMeta() {
    try {
        const raw = JSON.parse(localStorage.getItem(FEED_META_KEY) || 'null');
        if (!raw || typeof raw !== 'object') return null;
        return {
            httpsUrl: String(raw.httpsUrl || ''),
            webcalUrl: String(raw.webcalUrl || toWebcalUrl(raw.httpsUrl || '')),
            path: String(raw.path || FEED_REL_PATH),
            publishedAt: String(raw.publishedAt || ''),
            count: Number(raw.count) || 0
        };
    } catch {
        return null;
    }
}

/**
 * @param {object} meta
 */
export function saveFeedMeta(meta) {
    try {
        localStorage.setItem(FEED_META_KEY, JSON.stringify(meta || {}));
    } catch {
        /* ignore */
    }
    if (window.ms365BrowserBackup && typeof window.ms365BrowserBackup.notifyLocalDataChanged === 'function') {
        window.ms365BrowserBackup.notifyLocalDataChanged('akt-feed');
    }
}

async function resolveSiteDrive(tok, siteWebUrl) {
    const site = await G().resolveSiteFromWebUrl(tok, siteWebUrl);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    const drive = await G().graphJson(
        'GET',
        G().graphPathSite(siteId) + '/drive?$select=id,name,webUrl',
        tok,
        undefined,
        'v1.0'
    );
    const driveId = drive && drive.id ? String(drive.id) : '';
    if (!driveId) throw new Error('Dokumentbibliothek der Site nicht gefunden.');
    return { siteId, driveId, driveWebUrl: drive.webUrl || '' };
}

/**
 * Genehmigte Aktivitäten als ICS-Datei auf die Site schreiben.
 * @param {object} opts
 * @param {string} opts.siteUrl
 * @param {object[]} opts.items
 * @param {object} opts.stammdaten
 * @param {string} [opts.calendarName]
 */
export async function publishAktivitaetenIcsFeed(opts) {
    const o = opts || {};
    const siteUrl = String(o.siteUrl || '')
        .trim()
        .replace(/\/$/, '');
    if (!siteUrl) throw new Error('Bitte zuerst die SharePoint-Site laden.');

    const ics = buildIcs(o.items || [], o.stammdaten || {}, o.calendarName || 'Schulaktivitäten', {
        onlyApproved: true
    });
    const count = (String(ics).match(/BEGIN:VEVENT/g) || []).length;

    const tok = await G().getGraphToken(SCOPES);
    const { driveId } = await resolveSiteDrive(tok, siteUrl);
    const enc = encodeDriveRootPath(FEED_REL_PATH);
    const putUrl =
        G().graphBase('v1.0') + '/drives/' + encodeURIComponent(driveId) + '/' + enc + '/content';

    const putRes = await fetch(putUrl, {
        method: 'PUT',
        headers: {
            Authorization: 'Bearer ' + tok,
            'Content-Type': 'text/calendar; charset=utf-8'
        },
        body: ics
    });
    if (!putRes.ok) {
        const text = await putRes.text();
        throw new Error('ICS-Feed konnte nicht gespeichert werden: ' + (text || putRes.status));
    }
    let item = null;
    try {
        item = await putRes.json();
    } catch {
        item = null;
    }

    // Metadaten nachlesen (webUrl)
    if (!item || !item.webUrl) {
        item = await G().graphJson(
            'GET',
            G().graphBase('v1.0') +
                '/drives/' +
                encodeURIComponent(driveId) +
                '/' +
                enc +
                '?$select=id,name,webUrl',
            tok,
            undefined,
            'v1.0'
        );
    }

    const fileWebUrl = String((item && item.webUrl) || '').trim();
    if (!fileWebUrl) throw new Error('Feed-Datei gespeichert, aber keine webUrl erhalten.');

    // Roh-Download-URL erleichtert Abos (Outlook/Apple holen periodisch den Dateiinhalt).
    const httpsUrl = preferDownloadUrl(fileWebUrl);

    // Organisationsweiter Freigabelink (Anmeldung nötig) – Fallback/Weitergabe.
    let orgShareUrl = '';
    try {
        const itemId = item && item.id ? String(item.id) : '';
        if (itemId) {
            const link = await G().graphJson(
                'POST',
                G().graphBase('v1.0') +
                    '/drives/' +
                    encodeURIComponent(driveId) +
                    '/items/' +
                    encodeURIComponent(itemId) +
                    '/createLink',
                tok,
                { type: 'view', scope: 'organization' },
                'v1.0'
            );
            orgShareUrl = String((link && link.link && link.link.webUrl) || '').trim();
        }
    } catch {
        /* organisation link optional */
    }

    const meta = {
        httpsUrl,
        fileWebUrl,
        webcalUrl: toWebcalUrl(httpsUrl),
        shareUrl: orgShareUrl || httpsUrl,
        path: FEED_REL_PATH,
        publishedAt: new Date().toISOString(),
        count,
        siteUrl
    };
    saveFeedMeta(meta);
    return meta;
}

/**
 * SharePoint-Datei-URL um download=1 ergänzen (ICS-Clients brauchen den Rohinhalt).
 * @param {string} url
 */
export function preferDownloadUrl(url) {
    const u = String(url || '').trim();
    if (!u) return '';
    if (/[?&]download=/i.test(u)) return u;
    if (!/\.sharepoint\.com/i.test(u)) return u;
    return u + (u.includes('?') ? '&' : '?') + 'download=1';
}
