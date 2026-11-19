/**
 * ICS-Feed auf der SharePoint-Site veröffentlichen (abonnierbar für Outlook u. a.).
 *
 * Outlook „Internetkalender“ braucht Roh-ICS ohne interaktiven Login.
 * SharePoint-Pfad-URLs + webcal:// scheitern deshalb oft (Login-HTML statt ICS).
 * Strategie: anonyme Freigabe wenn Tenant erlaubt; sonst Org-Download-URL + klare Hinweise.
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
 * @param {string} siteWebUrl
 * @param {string} uniqueId
 */
export function buildDownloadAspxUrl(siteWebUrl, uniqueId) {
    const base = String(siteWebUrl || '')
        .trim()
        .replace(/\/$/, '');
    const id = String(uniqueId || '').trim();
    if (!base || !id) return '';
    return base + '/_layouts/15/download.aspx?UniqueId=' + encodeURIComponent(id);
}

/**
 * SharePoint-Datei-URL um download=1 ergänzen (Browser-Rohdownload).
 * @param {string} url
 */
export function preferDownloadUrl(url) {
    const u = String(url || '').trim();
    if (!u) return '';
    if (/[?&]download=/i.test(u)) return u;
    if (!/\.sharepoint\.com/i.test(u)) return u;
    return u + (u.includes('?') ? '&' : '?') + 'download=1';
}

/**
 * Sharing-Link für Kalender-Clients möglichst als Download ansprechen.
 * @param {string} shareWebUrl
 */
export function preferShareDownloadUrl(shareWebUrl) {
    const u = String(shareWebUrl || '').trim();
    if (!u) return '';
    // :u: / :x: / :b: Freigabelinks – download=1 liefert oft die Datei statt der Viewer-Seite
    if (/\/:([uxb]):\//i.test(u) || /sharepoint\.com/i.test(u)) {
        return preferDownloadUrl(u);
    }
    return u;
}

/**
 * Beste HTTPS-URL für Outlook „Kalender aus dem Internet“.
 * @param {{
 *   anonymousUrl?: string,
 *   downloadAspxUrl?: string,
 *   httpsUrl?: string,
 *   shareUrl?: string
 * }} meta
 */
export function resolveOutlookSubscribeHttpsUrl(meta) {
    const m = meta || {};
    return (
        String(m.anonymousUrl || '').trim() ||
        String(m.downloadAspxUrl || '').trim() ||
        String(m.httpsUrl || '').trim() ||
        String(m.shareUrl || '').trim() ||
        ''
    );
}

/**
 * @returns {object|null}
 */
export function loadFeedMeta() {
    try {
        const raw = JSON.parse(localStorage.getItem(FEED_META_KEY) || 'null');
        if (!raw || typeof raw !== 'object') return null;
        const httpsUrl = String(raw.httpsUrl || '');
        const anonymousUrl = String(raw.anonymousUrl || '');
        const downloadAspxUrl = String(raw.downloadAspxUrl || '');
        const shareUrl = String(raw.shareUrl || '');
        const outlookHttps = resolveOutlookSubscribeHttpsUrl({
            anonymousUrl,
            downloadAspxUrl,
            httpsUrl,
            shareUrl
        });
        return {
            httpsUrl,
            anonymousUrl,
            downloadAspxUrl,
            fileWebUrl: String(raw.fileWebUrl || ''),
            shareUrl: shareUrl || httpsUrl,
            webcalUrl: String(raw.webcalUrl || toWebcalUrl(anonymousUrl || outlookHttps || httpsUrl)),
            outlookHttpsUrl: outlookHttps,
            anonymousShare: !!raw.anonymousShare || !!anonymousUrl,
            subscribeHint: String(raw.subscribeHint || ''),
            path: String(raw.path || FEED_REL_PATH),
            publishedAt: String(raw.publishedAt || ''),
            count: Number(raw.count) || 0,
            siteUrl: String(raw.siteUrl || '')
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
    return { siteId, driveId, driveWebUrl: drive.webUrl || '', siteWebUrl: site.webUrl || siteWebUrl };
}

/**
 * @param {string} tok
 * @param {string} driveId
 * @param {string} itemId
 * @param {'anonymous'|'organization'} scope
 */
async function tryCreateViewLink(tok, driveId, itemId, scope) {
    try {
        const link = await G().graphJson(
            'POST',
            G().graphBase('v1.0') +
                '/drives/' +
                encodeURIComponent(driveId) +
                '/items/' +
                encodeURIComponent(itemId) +
                '/createLink',
            tok,
            { type: 'view', scope },
            'v1.0'
        );
        return String((link && link.link && link.link.webUrl) || '').trim();
    } catch {
        return '';
    }
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
    const { driveId, siteWebUrl } = await resolveSiteDrive(tok, siteUrl);
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

    // Metadaten inkl. UniqueId für download.aspx
    item = await G().graphJson(
        'GET',
        G().graphBase('v1.0') +
            '/drives/' +
            encodeURIComponent(driveId) +
            '/' +
            enc +
            '?$select=id,name,webUrl,sharepointIds',
        tok,
        undefined,
        'v1.0'
    );

    const fileWebUrl = String((item && item.webUrl) || '').trim();
    if (!fileWebUrl) throw new Error('Feed-Datei gespeichert, aber keine webUrl erhalten.');

    const itemId = item && item.id ? String(item.id) : '';
    const uniqueId = String(
        (item && item.sharepointIds && item.sharepointIds.listItemUniqueId) || ''
    ).trim();
    const downloadAspxUrl = buildDownloadAspxUrl(siteWebUrl || siteUrl, uniqueId);
    const pathHttps = preferDownloadUrl(fileWebUrl);

    let anonymousUrl = '';
    let orgShareUrl = '';
    if (itemId) {
        anonymousUrl = preferShareDownloadUrl(await tryCreateViewLink(tok, driveId, itemId, 'anonymous'));
        orgShareUrl = preferShareDownloadUrl(
            await tryCreateViewLink(tok, driveId, itemId, 'organization')
        );
    }

    const anonymousShare = !!anonymousUrl;
    const outlookHttps = resolveOutlookSubscribeHttpsUrl({
        anonymousUrl,
        downloadAspxUrl,
        httpsUrl: pathHttps,
        shareUrl: orgShareUrl
    });

    let subscribeHint = '';
    if (anonymousShare) {
        subscribeHint =
            'Öffentliche Freigabe aktiv. In Outlook: Kalender hinzufügen → Aus dem Internet → HTTPS-Link einfügen (nicht webcal). Apple: webcal kopieren.';
    } else {
        subscribeHint =
            'Kein anonymer Freigabelink (Tenant blockiert oft „Jeder“). Outlook-Internetkalender scheitert dann am Login. Alternativen: ICS herunterladen und importieren, oder genehmigte Aktivität „In meinen Kalender“. IT kann anonyme Freigaben für die Site erlauben und den Feed neu veröffentlichen.';
    }

    const meta = {
        httpsUrl: pathHttps,
        fileWebUrl,
        downloadAspxUrl,
        anonymousUrl,
        anonymousShare,
        webcalUrl: toWebcalUrl(anonymousUrl || outlookHttps),
        outlookHttpsUrl: outlookHttps,
        shareUrl: orgShareUrl || anonymousUrl || pathHttps,
        subscribeHint,
        path: FEED_REL_PATH,
        publishedAt: new Date().toISOString(),
        count,
        siteUrl
    };
    saveFeedMeta(meta);
    return meta;
}
