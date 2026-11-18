/**
 * Admin-Client gegen Gateway (Function Key). Nicht in öffentlicher Config.
 */

const SSP_AZURE_HOST_RE = /func-ms365-ssp-dev\.azurewebsites\.net/i;

/**
 * Lokal (Vite): über Same-Origin-Proxy `/__ssp-gw` → Azure, sonst Browser-CORS.
 * Produktion (app.ms365.schule): direkte Azure-URL.
 */
function baseUrl(settings) {
    const configured = String(settings.apiBaseUrl || '').replace(/\/$/, '');
    if (typeof window === 'undefined') return configured;
    const host = window.location.hostname;
    const isLocal = host === 'localhost' || host === '127.0.0.1';
    if (isLocal && (!configured || SSP_AZURE_HOST_RE.test(configured))) {
        return `${window.location.origin}/__ssp-gw/api`;
    }
    return configured;
}

function headers(settings, json = true) {
    const h = {};
    if (json) h['Content-Type'] = 'application/json';
    const key = String(settings.functionKey || '').trim();
    // Gateway prüft x-admin-key (STAMMDATEN_ADMIN_KEY); x-functions-key als Alias
    if (key) {
        h['x-admin-key'] = key;
        h['x-functions-key'] = key;
    }
    return h;
}

/**
 * Freundliche Admin-Meldungen – kein Roh-Graph, Cache bleibt lokal.
 * @param {Response} res
 * @param {object} data
 * @param {string} label
 */
export function softGatewayError(res, data, label = 'Gateway') {
    const detail = String(data?.error || data?.message || '').trim();
    const status = Number(res?.status || 0);
    if (status === 401) {
        return (
            detail ||
            'Anmeldung am Gateway fehlgeschlagen (Admin-Key). Key prüfen unter Einrichtung – lokale Demo-Daten bleiben erhalten.'
        );
    }
    if (status === 403) {
        return (
            detail ||
            'Kein Zugriff auf Gateway/SharePoint (403). Berechtigungen oder Site-URL prüfen – bestehende Liste bleibt im Cache.'
        );
    }
    if (status === 404) {
        return `${label}: Route nicht gefunden. API-Base-URL muss auf …/api enden (nicht …/api/health).`;
    }
    if (status === 409) {
        return detail || 'Statuswechsel nicht erlaubt (geprüfte Datensätze werden nicht überschrieben).';
    }
    if (status >= 500 || status === 0) {
        return (
            detail ||
            `${label} vorübergehend nicht erreichbar. Bitte später erneut versuchen – lokale Daten bleiben erhalten.`
        );
    }
    return detail || `${label}: HTTP ${status || '?'}`;
}

async function parse(res, label = 'Gateway') {
    const data = await res.json().catch(() => ({}));
    if (!res.ok) {
        throw new Error(softGatewayError(res, data, label));
    }
    return data;
}

export async function remoteEnsureList(settings) {
    const url = `${baseUrl(settings)}/ssp-ensure-list`;
    if (!baseUrl(settings)) {
        throw new Error('API-Base-URL fehlt (z. B. https://func-ms365-ssp-dev.azurewebsites.net/api).');
    }
    let res;
    try {
        res = await fetch(url, {
            method: 'POST',
            headers: headers(settings),
            body: JSON.stringify({
                siteUrl: settings.spSiteUrl || '',
                listName: settings.spListName || 'SchuelerAnmeldungen'
            })
        });
    } catch (e) {
        const origin = typeof window !== 'undefined' ? window.location.origin : '?';
        throw new Error(
            `Netzwerk/CORS zu ${url} (Seite: ${origin}): ${e?.message || e}. Lokal Vite neu starten (Proxy /__ssp-gw). Auf app.ms365.schule Origin in Azure CORS.`
        );
    }
    return parse(res, 'Liste anlegen');
}

export async function remoteEnroll(settings, { title, payload } = {}) {
    const res = await fetch(`${baseUrl(settings)}/ssp-enroll`, {
        method: 'POST',
        headers: headers(settings),
        body: JSON.stringify({ title, payload })
    });
    return parse(res);
}

export async function remoteResetPin(settings, id) {
    const res = await fetch(`${baseUrl(settings)}/ssp-reset-pin`, {
        method: 'POST',
        headers: headers(settings),
        body: JSON.stringify({ id })
    });
    return parse(res);
}

export async function remoteList(settings) {
    const res = await fetch(`${baseUrl(settings)}/ssp-list`, {
        method: 'GET',
        headers: headers(settings, false)
    });
    return parse(res);
}

/**
 * @param {object} settings
 * @param {string[]} ids
 * @param {string} status
 * @param {{ actorEmail?: string, actorName?: string }} [actor]
 */
export async function remoteSetStatus(settings, ids, status, actor = {}) {
    const res = await fetch(`${baseUrl(settings)}/ssp-set-status`, {
        method: 'POST',
        headers: headers(settings),
        body: JSON.stringify({
            ids,
            status,
            actorEmail: actor.actorEmail || '',
            actorName: actor.actorName || ''
        })
    });
    return parse(res);
}

export async function remoteHealth(settings) {
    const base = baseUrl(settings);
    if (!base) {
        throw new Error('API-Base-URL fehlt (z. B. https://func-ms365-ssp-dev.azurewebsites.net/api).');
    }
    let res;
    try {
        res = await fetch(`${base}/health`, { method: 'GET' });
    } catch (e) {
        throw new Error(`Netzwerk/CORS zu ${base}/health: ${e?.message || e}`);
    }
    return parse(res, 'Gateway prüfen');
}
