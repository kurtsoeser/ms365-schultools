import { safeJsonParse } from '../../shared/utils/json.js';
import { getEl } from '../../shared/utils/dom.js';
import { compareDe } from '../../shared/utils/strings.js';
import {
    applyPlaceKindsToRows,
    buildPlaceEmailKindMap,
    classifyDirectorySharedMailboxCandidate,
    countMailboxKinds,
    filterRowsByMailboxKind,
    inferMailboxKind,
    mailboxKindLabel,
    mergeReportWithDirectory,
    normalizeMailboxKind,
    sharedMailboxesFromUsageJson
} from './postfaecher-load.js';

const CACHE_KEY = 'ms365-postfaecher-cache-v1';
const GRAPH_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All',
    'https://graph.microsoft.com/Reports.Read.All'
];
const GRAPH_SCOPES_HEURISTIC = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All'
];
const GRAPH_SCOPES_PLACES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Place.Read.All'
];
const USER_SELECT =
    'id,displayName,mail,userPrincipalName,mailNickname,accountEnabled,userType,givenName,surname,jobTitle,department,officeLocation,businessPhones,mobilePhone,assignedLicenses';

function psSingleQuote(value) {
    return "'" + String(value == null ? '' : value).replace(/'/g, "''") + "'";
}

function psHereString(value) {
    const safe = String(value == null ? '' : value).replace(/^\s*'@\s*$/gm, "' @");
    return "@'\n" + safe + "\n'@";
}

function setProgress(on, text) {
    const el = getEl('mbProgress');
    if (!el) return;
    el.style.display = on ? '' : 'none';
    if (text) el.textContent = String(text);
}

async function getGraphToken(scopes) {
    if (typeof window.ms365AuthAcquireToken === 'function') {
        return await window.ms365AuthAcquireToken(scopes);
    }
    throw new Error('Bitte oben rechts anmelden (MSAL-Widget nicht verfügbar).');
}

function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
}

async function graphRequest(method, pathOrUrl, token, body, extraHeaders) {
    const url = pathOrUrl.indexOf('http') === 0 ? pathOrUrl : 'https://graph.microsoft.com/v1.0' + pathOrUrl;
    let attempt = 0;
    while (true) {
        const headers = { Authorization: 'Bearer ' + token };
        if (extraHeaders && typeof extraHeaders === 'object') Object.assign(headers, extraHeaders);
        let payload = undefined;
        if (body !== undefined) {
            headers['Content-Type'] = 'application/json';
            payload = JSON.stringify(body);
        }
        const res = await fetch(url, { method, headers, body: payload });
        if (res.status === 429 && attempt < 8) {
            const ra = parseInt(res.headers.get('Retry-After') || '5', 10);
            await sleep((isNaN(ra) ? 5 : ra) * 1000);
            attempt++;
            continue;
        }
        return res;
    }
}

async function graphJson(method, pathOrUrl, token, body, extraHeaders) {
    const res = await graphRequest(method, pathOrUrl, token, body, extraHeaders);
    const text = await res.text();
    let data = null;
    if (text) {
        try {
            data = JSON.parse(text);
        } catch {
            data = text;
        }
    }
    if (!res.ok) {
        const msg = typeof data === 'object' && data && data.error ? JSON.stringify(data.error) : text || String(res.status);
        throw new Error(method + ' ' + pathOrUrl + ': ' + msg);
    }
    return data || {};
}

async function fetchAllPages(token, initialPath, onProgress, extraHeaders) {
    const out = [];
    let next = initialPath;
    let page = 0;
    while (next) {
        page++;
        const data = await graphJson('GET', next, token, undefined, extraHeaders);
        const vals = data.value;
        if (Array.isArray(vals)) for (let i = 0; i < vals.length; i++) out.push(vals[i]);
        next = data['@odata.nextLink'] || null;
        if (typeof onProgress === 'function') onProgress({ page, loaded: out.length, hasMore: !!next });
    }
    return out;
}

function sortMailboxRows(rows) {
    const out = Array.isArray(rows) ? rows.slice() : [];
    out.sort((a, b) => {
        const conf = (x) => (x && x.highConfidence ? 0 : 1);
        const c = conf(a) - conf(b);
        if (c !== 0) return c;
        return compareDe(a.name, b.name);
    });
    return out;
}

async function tryFetchUsers(token, filterExpr, onProgress, label) {
    const filter = encodeURIComponent(filterExpr);
    const initial =
        '/users?$count=true&$filter=' + filter + '&$select=' + encodeURIComponent(USER_SELECT) + '&$top=999';
    return fetchAllPages(token, initial, (p) => {
        if (typeof onProgress === 'function') {
            onProgress({
                phase: 'directory',
                message: `${label}: Seite ${p.page} – ${p.loaded} …` + (p.hasMore ? '' : '')
            });
        }
    }, { ConsistencyLevel: 'eventual' });
}

/**
 * Verzeichnis-basiert (funktioniert im Browser ohne Reports-Redirect).
 * Mehrere Filter-Passes, weil Shared MBs je Tenant unterschiedlich aussehen.
 */
async function loadSharedMailboxesViaDirectory(token, onProgress) {
    const buckets = [];
    const errors = [];

    // Pass 1: klassisch – deaktiviert + Mail (deckt die meisten Shared MBs ab)
    try {
        buckets.push(
            ...(await tryFetchUsers(token, 'accountEnabled eq false and mail ne null', onProgress, 'Deaktivierte Konten'))
        );
    } catch (e) {
        errors.push('Pass1: ' + (e?.message || e));
    }

    // Pass 2: ohne Lizenz + Mail (auch aktivierte Shared MBs / Resource-ähnlich)
    try {
        buckets.push(
            ...(await tryFetchUsers(
                token,
                "mail ne null and userType eq 'Member' and assignedLicenses/$count eq 0",
                onProgress,
                'Konten ohne Lizenz'
            ))
        );
    } catch (e) {
        errors.push('Pass2: ' + (e?.message || e));
    }

    // Pass 3: einfache Abfrage ohne $count (manche Tenants blocken Advanced Queries)
    if (!buckets.length) {
        try {
            if (typeof onProgress === 'function') {
                onProgress({ phase: 'directory', message: 'Fallback: deaktivierte Konten (einfache Abfrage) …' });
            }
            const simplePath =
                '/users?$filter=' +
                encodeURIComponent('accountEnabled eq false') +
                '&$select=' +
                encodeURIComponent(USER_SELECT) +
                '&$top=999';
            buckets.push(...(await fetchAllPages(token, simplePath, undefined)));
        } catch (e) {
            errors.push('Pass3: ' + (e?.message || e));
        }
    }

    if (!buckets.length && errors.length) {
        throw new Error(errors.join(' | '));
    }

    const byId = new Map();
    buckets.forEach((u) => {
        const id = String(u?.id || '');
        if (id) byId.set(id, u);
    });

    const rows = [];
    byId.forEach((u) => {
        const row = classifyDirectorySharedMailboxCandidate(u);
        if (row) rows.push(row);
    });

    // Places (Räume) – optional, braucht Place.Read.All
    let placeMap = new Map();
    try {
        if (typeof onProgress === 'function') {
            onProgress({ phase: 'places', message: 'Places (Räume) werden gelesen …' });
        }
        let placeToken = token;
        try {
            placeToken = await getGraphToken(GRAPH_SCOPES_PLACES);
        } catch {
            // gleicher Token versuchen
        }
        const rooms = await fetchAllPages(
            placeToken,
            '/places/microsoft.graph.room?$select=id,displayName,emailAddress&$top=999',
            (p) => {
                if (typeof onProgress === 'function') {
                    onProgress({
                        phase: 'places',
                        message: `Räume Seite ${p.page} – ${p.loaded} …`
                    });
                }
            }
        );
        placeMap = buildPlaceEmailKindMap(
            (rooms || []).map((r) => ({ ...r, kind: 'room', '@odata.type': '#microsoft.graph.room' }))
        );
    } catch {
        // Places optional – Heuristik bleibt
    }

    const typed = applyPlaceKindsToRows(rows, placeMap);

    if (typeof onProgress === 'function') {
        const c = countMailboxKinds(typed);
        onProgress({
            phase: 'directory',
            message: `Verzeichnis: ${typed.length} Kandidaten (${c.shared} freigegeben, ${c.room} Räume, ${c.equipment} Geräte) aus ${byId.size} Konten.`
        });
    }
    return { rows: sortMailboxRows(typed), scanned: byId.size, errors };
}

/**
 * JSON-Report ohne CSV-Redirect (CORS-sicherer als text/csv).
 * Schlägt oft fehl (Rechte/RecipientType) – dann Verzeichnis nutzen.
 */
async function loadSharedMailboxesViaReportJson(token, onProgress) {
    if (typeof onProgress === 'function') {
        onProgress({ phase: 'report', message: 'Mailbox-Usage-Report (JSON) …' });
    }
    const path =
        "https://graph.microsoft.com/beta/reports/getMailboxUsageDetail(period='D180')?$format=application/json";
    const all = await fetchAllPages(token, path, (p) => {
        if (typeof onProgress === 'function') {
            onProgress({
                phase: 'report',
                message: `Report JSON Seite ${p.page} – ${p.loaded} …`
            });
        }
    });
    const parsed = sharedMailboxesFromUsageJson({ value: all });
    if (!parsed.ok) {
        throw new Error('Report-JSON ohne Recipient Type (Shared).');
    }
    let users = [];
    try {
        users = await tryFetchUsers(token, 'accountEnabled eq false and mail ne null', onProgress, 'Anreichern');
    } catch {
        // ignore
    }
    return sortMailboxRows(mergeReportWithDirectory(parsed.rows, users));
}

function loadCache() {
    try {
        const raw = localStorage.getItem(CACHE_KEY);
        const obj = raw ? safeJsonParse(raw) : null;
        const rows = obj && Array.isArray(obj.rows) ? obj.rows : [];
        return { rows, loadedAt: obj && obj.loadedAt ? String(obj.loadedAt) : '' };
    } catch {
        return { rows: [], loadedAt: '' };
    }
}

function saveCache(rows) {
    const out = Array.isArray(rows) ? rows : [];
    try {
        localStorage.setItem(CACHE_KEY, JSON.stringify({ rows: out, loadedAt: new Date().toISOString() }));
    } catch {
        // ignore
    }
}

/**
 * Primär: Verzeichnis (/users) – funktioniert im Browser zuverlässig.
 * Optional: Report JSON zur Verfeinerung (ohne CSV-CORS-Redirect).
 * Goldstandard bleibt Exchange-Export/Import.
 */
async function loadSharedMailboxesLive(onProgress) {
    let warnParts = [];
    let token = await getGraphToken(GRAPH_SCOPES_HEURISTIC);

    if (typeof onProgress === 'function') {
        onProgress({ phase: 'directory', message: 'Verzeichnis wird gelesen …' });
    }
    const dir = await loadSharedMailboxesViaDirectory(token, onProgress);
    let rows = dir.rows || [];
    let source = 'directory';

    // Optional: Report JSON nur wenn Verzeichnis nichts lieferte (vermeidet unnötigen Reports-Consent)
    if (!rows.length) {
        try {
            token = await getGraphToken(GRAPH_SCOPES);
            if (typeof onProgress === 'function') {
                onProgress({ phase: 'report', message: 'Verzeichnis leer – versuche Report JSON …' });
            }
            const reportRows = await loadSharedMailboxesViaReportJson(token, onProgress);
            if (reportRows.length) {
                rows = sortMailboxRows(reportRows);
                source = 'report';
            }
        } catch (reportErr) {
            warnParts.push(
                'Report fehlgeschlagen (' +
                    String(reportErr?.message || reportErr) +
                    '). Bitte Export/Import nutzen.'
            );
        }
    }

    if (dir.errors && dir.errors.length) {
        warnParts.push('Teilfehler Verzeichnis: ' + dir.errors.join(' | '));
    }

    saveCache(rows);
    return {
        rows,
        source,
        warn: warnParts.join(' ')
    };
}

function bind() {
    const cache = loadCache();
    /** @type {{id:string,name:string,mail:string,upn:string,alias:string,highConfidence?:boolean,kind?:string}[]} */
    let rows = (cache.rows || []).map((r) => ({
        ...r,
        kind: normalizeMailboxKind(r.kind) || inferMailboxKind(r)
    }));
    let selectedId = '';
    let newMode = false;

    function filteredRows() {
        const kindFilter = String(getEl('mbKindFilter')?.value || 'shared');
        const byKind = filterRowsByMailboxKind(rows, kindFilter);
        const q = String(getEl('mbSearch')?.value || '').trim().toLowerCase();
        if (!q) return byKind;
        return byKind.filter((r) => {
            const hay = (r.name + ' ' + r.mail + ' ' + r.upn + ' ' + r.alias + ' ' + (r.kind || '')).toLowerCase();
            return hay.includes(q);
        });
    }

    function kindPill(kind) {
        const k = normalizeMailboxKind(kind) || 'shared';
        const el = document.createElement('div');
        el.className = 'pill';
        el.textContent = mailboxKindLabel(k);
        if (k === 'room') {
            el.style.borderColor = 'rgba(13,110,253,0.25)';
            el.style.background = 'rgba(13,110,253,0.1)';
            el.style.color = '#0b5ed7';
        } else if (k === 'equipment') {
            el.style.borderColor = 'rgba(111,66,193,0.25)';
            el.style.background = 'rgba(111,66,193,0.1)';
            el.style.color = '#6f42c1';
        } else {
            el.style.borderColor = 'rgba(25,135,84,0.25)';
            el.style.background = 'rgba(25,135,84,0.1)';
            el.style.color = '#198754';
        }
        return el;
    }

    function renderList() {
        const tree = getEl('mbTree');
        if (!tree) return;
        tree.replaceChildren();
        const visible = filteredRows();
        if (!visible.length) {
            const li = document.createElement('li');
            li.style.padding = '10px 12px';
            li.style.color = '#6c757d';
            const kindFilter = String(getEl('mbKindFilter')?.value || 'shared');
            if (!rows.length) {
                li.textContent = 'Noch keine Daten – „Tenant einlesen“ oder Export/Import.';
            } else if (kindFilter !== 'all') {
                li.textContent = 'Keine Treffer in diesem Typ – Filter auf „Alle“ stellen oder Suche leeren.';
            } else {
                li.textContent = 'Keine Treffer.';
            }
            tree.appendChild(li);
            return;
        }
        visible.forEach((r) => {
            const li = document.createElement('li');
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.dataset.mbSelect = r.id;
            btn.setAttribute('aria-current', selectedId === r.id ? 'true' : 'false');

            const name = document.createElement('div');
            name.style.minWidth = '0';
            name.style.flex = '1';
            name.style.fontWeight = '900';
            name.style.color = '#32325d';
            name.style.overflow = 'hidden';
            name.style.textOverflow = 'ellipsis';
            name.style.whiteSpace = 'nowrap';
            name.textContent = r.name || '(ohne Anzeigename)';

            const meta = document.createElement('div');
            meta.className = 'pill';
            meta.textContent = r.mail || r.alias || '–';

            btn.appendChild(kindPill(r.kind));
            if (r.highConfidence === false) {
                const warn = document.createElement('div');
                warn.className = 'pill';
                warn.style.borderColor = 'rgba(133,100,4,0.22)';
                warn.style.background = 'rgba(255,193,7,0.12)';
                warn.style.color = '#856404';
                warn.textContent = 'unsicher';
                btn.appendChild(warn);
            }

            btn.appendChild(name);
            btn.appendChild(meta);
            li.appendChild(btn);
            tree.appendChild(li);
        });
    }

    function renderDetail() {
        const hint = getEl('mbHint');
        const detail = getEl('mbDetail');
        const newDetail = getEl('mbNewDetail');
        const cur = rows.find((x) => x.id === selectedId) || null;
        if (hint) hint.style.display = cur || newMode ? 'none' : '';
        if (detail) detail.style.display = cur && !newMode ? '' : 'none';
        if (newDetail) newDetail.style.display = newMode ? '' : 'none';
        if (!cur) return;
        const n = getEl('mbName');
        const m = getEl('mbMail');
        const u = getEl('mbUpn');
        const i = getEl('mbId');
        const k = getEl('mbKind');
        if (n) n.value = cur.name || '';
        if (m) m.value = cur.mail || '';
        if (u) u.value = cur.upn || '';
        if (i) i.value = cur.id || '';
        if (k) k.value = mailboxKindLabel(cur.kind || inferMailboxKind(cur));
    }

    function downloadText(filename, text) {
        const blob = new Blob([String(text || '')], { type: 'text/plain;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 500);
    }

    function parseDelegateLines(raw) {
        const lines = String(raw || '').split(/\r\n|\n|\r/);
        const seen = new Set();
        const out = [];
        lines.forEach((l) => {
            const t = String(l || '').trim();
            if (!t || t.startsWith('#')) return;
            const key = t.toLowerCase();
            if (seen.has(key)) return;
            seen.add(key);
            out.push(t);
        });
        return out;
    }

    function selectedNewRights() {
        const rights = [];
        if (getEl('mbNewRightFullAccess')?.checked) rights.push('FullAccess');
        if (getEl('mbNewRightSendAs')?.checked) rights.push('SendAs');
        if (getEl('mbNewRightSendOnBehalf')?.checked) rights.push('SendOnBehalf');
        return rights;
    }

    function buildPermissionScript(mailboxUpnOrMail, mode, rights, delegates) {
        const mb = String(mailboxUpnOrMail || '').trim();
        const m = String(mode || 'read').trim();
        const r = String(rights || 'FullAccess').trim();
        const dels = Array.isArray(delegates) ? delegates : [];

        const lines = [];
        lines.push('# Freigegebene Postfächer: Berechtigungen (Exchange Online)');
        lines.push('# Voraussetzungen: Install-Module ExchangeOnlineManagement -Scope CurrentUser');
        lines.push('# Anmeldung: Connect-ExchangeOnline (Admin-Konto)');
        lines.push('');
        lines.push('$ErrorActionPreference = "Stop"');
        lines.push('');
        lines.push('if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) {');
        lines.push('  Write-Host "ExchangeOnlineManagement fehlt. Installiere: Install-Module ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Yellow');
        lines.push('}');
        lines.push('Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue');
        lines.push('');
        lines.push('try {');
        lines.push('  Connect-ExchangeOnline -ShowBanner:$false | Out-Null');
        lines.push('} catch {');
        lines.push('  Write-Host ("Connect-ExchangeOnline fehlgeschlagen: " + $_.Exception.Message) -ForegroundColor Red');
        lines.push('  throw');
        lines.push('}');
        lines.push('');
        lines.push(`$Mailbox = "${mb.replace(/"/g, '""')}"`);
        lines.push('Write-Host ("Postfach: " + $Mailbox) -ForegroundColor Cyan');
        lines.push('');
        lines.push('Write-Host "== Report ==" -ForegroundColor Cyan');
        lines.push('Write-Host "FullAccess:" -ForegroundColor DarkCyan');
        lines.push('Get-MailboxPermission -Identity $Mailbox | Where-Object { $_.IsInherited -eq $false -and $_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "S-1-5-*" } | Select-Object User,AccessRights,Deny,IsInherited | Format-Table -AutoSize');
        lines.push('');
        lines.push('Write-Host "SendAs:" -ForegroundColor DarkCyan');
        lines.push('Get-RecipientPermission -Identity $Mailbox | Where-Object { $_.Trustee -and $_.AccessRights -contains "SendAs" } | Select-Object Trustee,AccessRights,IsInherited | Format-Table -AutoSize');
        lines.push('');
        lines.push('Write-Host "SendOnBehalf:" -ForegroundColor DarkCyan');
        lines.push('Get-Mailbox -Identity $Mailbox | Select-Object -ExpandProperty GrantSendOnBehalfTo | ForEach-Object { $_.Name }');
        lines.push('');

        if (m === 'set') {
            lines.push('Write-Host "== Änderungen ==" -ForegroundColor Cyan');
            if (!dels.length) {
                lines.push('Write-Host "Keine Delegierten angegeben – nur Report ausgegeben." -ForegroundColor Yellow');
            } else {
                lines.push('$Delegates = @(' + dels.map((x) => `"${String(x).replace(/"/g, '""')}"`).join(', ') + ')');
                lines.push('');
                if (r === 'FullAccess') {
                    lines.push('foreach ($d in $Delegates) {');
                    lines.push('  if ($d.StartsWith("-")) {');
                    lines.push('    $u = $d.Substring(1)');
                    lines.push('    Write-Host ("Remove FullAccess: " + $u) -ForegroundColor Yellow');
                    lines.push('    Remove-MailboxPermission -Identity $Mailbox -User $u -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Continue | Out-Null');
                    lines.push('  } else {');
                    lines.push('    Write-Host ("Add FullAccess: " + $d) -ForegroundColor Green');
                    lines.push('    Add-MailboxPermission -Identity $Mailbox -User $d -AccessRights FullAccess -InheritanceType All -AutoMapping:$false -ErrorAction Continue | Out-Null');
                    lines.push('  }');
                    lines.push('}');
                } else if (r === 'SendAs') {
                    lines.push('foreach ($d in $Delegates) {');
                    lines.push('  if ($d.StartsWith("-")) {');
                    lines.push('    $u = $d.Substring(1)');
                    lines.push('    Write-Host ("Remove SendAs: " + $u) -ForegroundColor Yellow');
                    lines.push('    Remove-RecipientPermission -Identity $Mailbox -Trustee $u -AccessRights SendAs -Confirm:$false -ErrorAction Continue | Out-Null');
                    lines.push('  } else {');
                    lines.push('    Write-Host ("Add SendAs: " + $d) -ForegroundColor Green');
                    lines.push('    Add-RecipientPermission -Identity $Mailbox -Trustee $d -AccessRights SendAs -Confirm:$false -ErrorAction Continue | Out-Null');
                    lines.push('  }');
                    lines.push('}');
                } else if (r === 'SendOnBehalf') {
                    lines.push('# SendOnBehalf ist eine Liste am Mailbox-Objekt. Hier: Setze auf bestehend + Add/Remove.');
                    lines.push('$existing = @(Get-Mailbox -Identity $Mailbox).GrantSendOnBehalfTo');
                    lines.push('$set = New-Object System.Collections.Generic.List[object]');
                    lines.push('foreach ($x in $existing) { $set.Add($x) }');
                    lines.push('foreach ($d in $Delegates) {');
                    lines.push('  if ($d.StartsWith("-")) {');
                    lines.push('    $u = $d.Substring(1)');
                    lines.push('    Write-Host ("Remove SendOnBehalf: " + $u) -ForegroundColor Yellow');
                    lines.push('    $set = New-Object System.Collections.Generic.List[object] ($set | Where-Object { $_.Name -ne $u -and $_.PrimarySmtpAddress -ne $u })');
                    lines.push('  } else {');
                    lines.push('    Write-Host ("Add SendOnBehalf: " + $d) -ForegroundColor Green');
                    lines.push('    try { $rec = Get-Recipient -Identity $d -ErrorAction Stop; $set.Add($rec) } catch { Write-Host ("  nicht gefunden: " + $d) -ForegroundColor Red }');
                    lines.push('  }');
                    lines.push('}');
                    lines.push('Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo $set -ErrorAction Continue');
                }
                lines.push('');
                lines.push('Write-Host "== Report nach Änderungen ==" -ForegroundColor Cyan');
                lines.push('Get-MailboxPermission -Identity $Mailbox | Where-Object { $_.IsInherited -eq $false -and $_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "S-1-5-*" } | Select-Object User,AccessRights,Deny,IsInherited | Format-Table -AutoSize');
            }
            lines.push('');
        }

        lines.push('Disconnect-ExchangeOnline -Confirm:$false | Out-Null');
        lines.push('Write-Host "Fertig." -ForegroundColor Green');
        return lines.join('\n');
    }

    function buildCreateMailboxScript(options) {
        const displayName = String(options?.displayName || '').trim();
        const alias = String(options?.alias || '').trim();
        const primarySmtp = String(options?.primarySmtp || '').trim();
        const description = String(options?.description || '').trim();
        const delegates = Array.isArray(options?.delegates) ? options.delegates : [];
        const rights = Array.isArray(options?.rights) ? options.rights : [];
        const autoMapping = options?.autoMapping === true;
        const hideFromAddressLists = options?.hideFromAddressLists === true;

        const lines = [];
        lines.push('# Freigegebenes Postfach erstellen (Exchange Online)');
        lines.push('# Voraussetzungen: Install-Module ExchangeOnlineManagement -Scope CurrentUser');
        lines.push('# Anmeldung: Connect-ExchangeOnline (Admin-Konto)');
        lines.push('');
        lines.push('$ErrorActionPreference = "Stop"');
        lines.push('');
        lines.push('if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) {');
        lines.push('  Write-Host "ExchangeOnlineManagement fehlt. Installiere: Install-Module ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Yellow');
        lines.push('}');
        lines.push('Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue');
        lines.push('');
        lines.push('try {');
        lines.push('  Connect-ExchangeOnline -ShowBanner:$false | Out-Null');
        lines.push('} catch {');
        lines.push('  Write-Host ("Connect-ExchangeOnline fehlgeschlagen: " + $_.Exception.Message) -ForegroundColor Red');
        lines.push('  throw');
        lines.push('}');
        lines.push('');
        lines.push('$DisplayName = ' + psSingleQuote(displayName));
        lines.push('$Alias = ' + psSingleQuote(alias));
        lines.push('$PrimarySmtpAddress = ' + psSingleQuote(primarySmtp));
        lines.push('$Description = ' + psHereString(description));
        lines.push('$HideFromAddressLists = $' + (hideFromAddressLists ? 'true' : 'false'));
        lines.push('$Delegates = @(' + delegates.map(psSingleQuote).join(', ') + ')');
        lines.push('');
        lines.push('if ([string]::IsNullOrWhiteSpace($DisplayName)) { throw "Anzeigename fehlt." }');
        lines.push('if ([string]::IsNullOrWhiteSpace($Alias)) { throw "Alias fehlt." }');
        lines.push('');
        lines.push('$newMailboxParams = @{');
        lines.push('  Shared = $true');
        lines.push('  Name = $DisplayName');
        lines.push('  DisplayName = $DisplayName');
        lines.push('  Alias = $Alias');
        lines.push('}');
        lines.push('if (-not [string]::IsNullOrWhiteSpace($PrimarySmtpAddress)) {');
        lines.push('  $newMailboxParams.PrimarySmtpAddress = $PrimarySmtpAddress');
        lines.push('}');
        lines.push('');
        lines.push('Write-Host ("Erstelle freigegebenes Postfach: " + $DisplayName) -ForegroundColor Cyan');
        lines.push('New-Mailbox @newMailboxParams | Out-Null');
        lines.push('$MailboxIdentity = if ([string]::IsNullOrWhiteSpace($PrimarySmtpAddress)) { $Alias } else { $PrimarySmtpAddress }');
        lines.push('');
        lines.push('Write-Host "Warte auf Verfügbarkeit in Exchange Online ..." -ForegroundColor DarkCyan');
        lines.push('$mailbox = $null');
        lines.push('for ($i = 1; $i -le 24; $i++) {');
        lines.push('  try {');
        lines.push('    $mailbox = Get-Mailbox -Identity $MailboxIdentity -ErrorAction Stop');
        lines.push('    break');
        lines.push('  } catch {');
        lines.push('    Start-Sleep -Seconds 5');
        lines.push('  }');
        lines.push('}');
        lines.push('if (-not $mailbox) { throw "Postfach wurde erstellt, ist aber noch nicht abrufbar. Script später erneut für Rechte ausführen." }');
        lines.push('');
        lines.push('if (-not [string]::IsNullOrWhiteSpace($Description)) {');
        lines.push('  Set-Mailbox -Identity $MailboxIdentity -Notes $Description');
        lines.push('}');
        lines.push('Set-Mailbox -Identity $MailboxIdentity -HiddenFromAddressListsEnabled:$HideFromAddressLists');
        lines.push('');
        if (!rights.length) {
            lines.push('Write-Host "Keine Rechte ausgewählt – es werden keine Delegierten berechtigt." -ForegroundColor Yellow');
        } else {
            lines.push('if ($Delegates.Count -eq 0) {');
            lines.push('  Write-Host "Keine Delegierten angegeben – Rechte-Schritte werden übersprungen." -ForegroundColor Yellow');
            lines.push('} else {');
            if (rights.includes('FullAccess')) {
                lines.push('  Write-Host "Setze FullAccess ..." -ForegroundColor Cyan');
                lines.push('  foreach ($d in $Delegates) {');
                lines.push('    Add-MailboxPermission -Identity $MailboxIdentity -User $d -AccessRights FullAccess -InheritanceType All -AutoMapping:$' + (autoMapping ? 'true' : 'false') + ' -ErrorAction Continue | Out-Null');
                lines.push('  }');
            }
            if (rights.includes('SendAs')) {
                lines.push('  Write-Host "Setze SendAs ..." -ForegroundColor Cyan');
                lines.push('  foreach ($d in $Delegates) {');
                lines.push('    Add-RecipientPermission -Identity $MailboxIdentity -Trustee $d -AccessRights SendAs -Confirm:$false -ErrorAction Continue | Out-Null');
                lines.push('  }');
            }
            if (rights.includes('SendOnBehalf')) {
                lines.push('  Write-Host "Setze SendOnBehalf ..." -ForegroundColor Cyan');
                lines.push('  Set-Mailbox -Identity $MailboxIdentity -GrantSendOnBehalfTo @{Add=$Delegates} -ErrorAction Continue');
            }
            lines.push('}');
        }
        lines.push('');
        lines.push('Write-Host "== Ergebnis ==" -ForegroundColor Cyan');
        lines.push('Get-Mailbox -Identity $MailboxIdentity | Select-Object DisplayName,Alias,PrimarySmtpAddress,RecipientTypeDetails,HiddenFromAddressListsEnabled | Format-List');
        lines.push('Write-Host "FullAccess:" -ForegroundColor DarkCyan');
        lines.push('Get-MailboxPermission -Identity $MailboxIdentity | Where-Object { $_.IsInherited -eq $false -and $_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "S-1-5-*" } | Select-Object User,AccessRights,Deny,IsInherited | Format-Table -AutoSize');
        lines.push('Write-Host "SendAs:" -ForegroundColor DarkCyan');
        lines.push('Get-RecipientPermission -Identity $MailboxIdentity | Where-Object { $_.Trustee -and $_.AccessRights -contains "SendAs" } | Select-Object Trustee,AccessRights,IsInherited | Format-Table -AutoSize');
        lines.push('Write-Host "SendOnBehalf:" -ForegroundColor DarkCyan');
        lines.push('Get-Mailbox -Identity $MailboxIdentity | Select-Object -ExpandProperty GrantSendOnBehalfTo | ForEach-Object { $_.Name }');
        lines.push('');
        lines.push('Disconnect-ExchangeOnline -Confirm:$false | Out-Null');
        lines.push('Write-Host "Fertig." -ForegroundColor Green');
        return lines.join('\n');
    }

    function updatePsScript() {
        const modeSel = getEl('mbPsMode');
        const rightsSel = getEl('mbPsRights');
        const delTa = getEl('mbPsDelegates');
        const out = getEl('mbPsScript');
        if (!out) return;
        const cur = rows.find((x) => x.id === selectedId) || null;
        const mb = cur ? (cur.upn || cur.mail || '') : '';
        const mode = modeSel ? String(modeSel.value || 'read') : 'read';
        const rights = rightsSel ? String(rightsSel.value || 'FullAccess') : 'FullAccess';
        const delegates = delTa ? parseDelegateLines(delTa.value) : [];
        out.value = mb ? buildPermissionScript(mb, mode, rights, delegates) : '';
    }

    function updateNewPsScript() {
        const out = getEl('mbNewPsScript');
        if (!out) return;
        const delegates = parseDelegateLines(getEl('mbNewDelegates')?.value || '');
        out.value = buildCreateMailboxScript({
            displayName: getEl('mbNewDisplayName')?.value || '',
            alias: getEl('mbNewAlias')?.value || '',
            primarySmtp: getEl('mbNewPrimarySmtp')?.value || '',
            description: getEl('mbNewDescription')?.value || '',
            delegates,
            rights: selectedNewRights(),
            autoMapping: String(getEl('mbNewAutoMapping')?.value || 'false') === 'true',
            hideFromAddressLists: String(getEl('mbNewHiddenFromAddressLists')?.value || 'false') === 'true'
        });
    }

    function rerender() {
        renderList();
        renderDetail();
        updatePsScript();
        updateNewPsScript();
    }

    getEl('mbTree')?.addEventListener('click', (ev) => {
        const t = ev.target;
        const btn = t && t.closest ? t.closest('button[data-mb-select]') : null;
        if (!btn) return;
        const id = btn.getAttribute('data-mb-select');
        if (!id) return;
        selectedId = String(id);
        newMode = false;
        rerender();
    });

    getEl('mbBtnNew')?.addEventListener('click', () => {
        selectedId = '';
        newMode = true;
        rerender();
        getEl('mbNewDisplayName')?.focus();
    });

    getEl('mbNewCancel')?.addEventListener('click', () => {
        newMode = false;
        rerender();
    });

    getEl('mbSearch')?.addEventListener('input', () => rerender());
    getEl('mbKindFilter')?.addEventListener('change', () => rerender());

    getEl('mbCopyUpn')?.addEventListener('click', async () => {
        const cur = rows.find((x) => x.id === selectedId) || null;
        if (!cur) return;
        const text = cur.upn || cur.mail || '';
        if (!text) return;
        try {
            await navigator.clipboard.writeText(text);
        } catch {
            // ignore
        }
    });

    ['mbPsMode', 'mbPsRights', 'mbPsDelegates'].forEach((id) => {
        const el = getEl(id);
        if (!el) return;
        const evt = id === 'mbPsMode' || id === 'mbPsRights' ? 'change' : 'input';
        el.addEventListener(evt, () => updatePsScript());
    });

    [
        'mbNewDisplayName',
        'mbNewAlias',
        'mbNewPrimarySmtp',
        'mbNewDescription',
        'mbNewDelegates',
        'mbNewRightFullAccess',
        'mbNewRightSendAs',
        'mbNewRightSendOnBehalf',
        'mbNewAutoMapping',
        'mbNewHiddenFromAddressLists'
    ].forEach((id) => {
        const el = getEl(id);
        if (!el) return;
        const evt = el.tagName === 'SELECT' || el.type === 'checkbox' ? 'change' : 'input';
        el.addEventListener(evt, () => updateNewPsScript());
    });

    getEl('mbPsCopy')?.addEventListener('click', async () => {
        const ta = getEl('mbPsScript');
        const text = ta ? String(ta.value || '') : '';
        if (!text) return;
        try {
            await navigator.clipboard.writeText(text);
        } catch {
            // ignore
        }
    });

    getEl('mbPsDownload')?.addEventListener('click', () => {
        const ta = getEl('mbPsScript');
        const text = ta ? String(ta.value || '') : '';
        if (!text) return;
        const cur = rows.find((x) => x.id === selectedId) || null;
        const name = cur ? String(cur.alias || cur.mail || 'postfach').replace(/[^a-zA-Z0-9-_]/g, '') : 'postfach';
        downloadText(`postfach-permissions-${name}.ps1`, text);
    });

    getEl('mbNewPsCopy')?.addEventListener('click', async () => {
        const ta = getEl('mbNewPsScript');
        const text = ta ? String(ta.value || '') : '';
        if (!text) return;
        try {
            await navigator.clipboard.writeText(text);
        } catch {
            // ignore
        }
    });

    getEl('mbNewPsDownload')?.addEventListener('click', () => {
        const ta = getEl('mbNewPsScript');
        const text = ta ? String(ta.value || '') : '';
        if (!text) return;
        const rawName = getEl('mbNewAlias')?.value || getEl('mbNewPrimarySmtp')?.value || 'postfach';
        const name = String(rawName).replace(/[^a-zA-Z0-9-_]/g, '') || 'postfach';
        downloadText(`postfach-neu-${name}.ps1`, text);
    });

    async function loadNowGraph() {
        const btn = getEl('mbBtnLoadGraph');
        if (btn) btn.disabled = true;
        try {
            setProgress(true, 'Starte – lese freigegebene Postfächer aus dem Verzeichnis …');
            const result = await loadSharedMailboxesLive((p) => {
                if (p && p.message) setProgress(true, p.message);
                else if (p && p.page != null) {
                    setProgress(true, `Seite ${p.page} – bisher ${p.loaded} …` + (p.hasMore ? '' : ' (fertig)'));
                }
            });
            rows = result.rows || [];
            const counts = countMailboxKinds(rows);
            const srcLabel =
                result.source === 'report+directory'
                    ? 'Report+Verzeichnis'
                    : result.source === 'directory'
                      ? 'Verzeichnis'
                      : String(result.source || 'Graph');
            let msg = `Fertig (${srcLabel}): ${rows.length} Postfächer – ${counts.shared} freigegeben, ${counts.room} Räume, ${counts.equipment} Geräte.`;
            if (result.warn) msg += ' Hinweis: ' + result.warn;
            if (!rows.length) {
                msg +=
                    ' Keine Treffer im Verzeichnis. Bitte „Export“ (Exchange Online PowerShell) → „Import“ nutzen – das liest SharedMailbox zuverlässig.';
            } else {
                const unsure = rows.filter((r) => r.highConfidence === false).length;
                if (unsure) {
                    msg += ` Davon ${unsure} als „unsicher“ markiert.`;
                }
                msg += ' Filter oben: standardmäßig nur Freigegebene.';
            }
            setProgress(true, msg);
            if (!result.warn && rows.length) setTimeout(() => setProgress(false, ''), 2800);
            selectedId = '';
            rerender();
        } catch (e) {
            setProgress(
                true,
                'Fehler: ' +
                    (e?.message || String(e)) +
                    ' – Alternative: Export (PowerShell) → Import JSON.'
            );
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    function buildExoExportScript() {
        const lines = [];
        lines.push('# Export: Shared Mailboxes (Exchange Online) -> JSON');
        lines.push('# Voraussetzungen: Install-Module ExchangeOnlineManagement -Scope CurrentUser');
        lines.push('$ErrorActionPreference = "Stop"');
        lines.push('');
        lines.push('Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue');
        lines.push('try { Connect-ExchangeOnline -ShowBanner:$false | Out-Null } catch { throw }');
        lines.push('');
        lines.push('$outFile = "shared-mailboxes-export.json"');
        lines.push('$items = @()');
        lines.push(
            "Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox,RoomMailbox,EquipmentMailbox | ForEach-Object {"
        );
        lines.push('  $kind = switch ($_.RecipientTypeDetails) {');
        lines.push("    'RoomMailbox' { 'room' }");
        lines.push("    'EquipmentMailbox' { 'equipment' }");
        lines.push("    Default { 'shared' }");
        lines.push('  }');
        lines.push('  $items += [pscustomobject]@{');
        lines.push('    id = $_.ExternalDirectoryObjectId');
        lines.push('    name = $_.DisplayName');
        lines.push('    mail = $_.PrimarySmtpAddress');
        lines.push('    upn = $_.UserPrincipalName');
        lines.push('    alias = $_.Alias');
        lines.push('    kind = $kind');
        lines.push('    recipientTypeDetails = [string]$_.RecipientTypeDetails');
        lines.push('  }');
        lines.push('}');
        lines.push('$items | ConvertTo-Json -Depth 5 | Out-File -FilePath $outFile -Encoding UTF8');
        lines.push('Disconnect-ExchangeOnline -Confirm:$false | Out-Null');
        lines.push(
            'Write-Host ("OK: " + $items.Count + " Postfächer (Shared/Room/Equipment) -> " + $outFile) -ForegroundColor Green'
        );
        return lines.join('\n');
    }

    function showExoExport() {
        const wrap = getEl('mbExoExportWrap');
        const ta = getEl('mbExoExportScript');
        if (wrap) wrap.style.display = '';
        if (ta) ta.value = buildExoExportScript();
    }

    getEl('mbBtnLoadExo')?.addEventListener('click', showExoExport);
    getEl('mbExoHide')?.addEventListener('click', () => {
        const wrap = getEl('mbExoExportWrap');
        if (wrap) wrap.style.display = 'none';
    });
    getEl('mbExoCopy')?.addEventListener('click', async () => {
        const ta = getEl('mbExoExportScript');
        const text = ta ? String(ta.value || '') : '';
        if (!text) return;
        try {
            await navigator.clipboard.writeText(text);
        } catch {
            // ignore
        }
    });
    getEl('mbExoDownload')?.addEventListener('click', () => {
        const ta = getEl('mbExoExportScript');
        const text = ta ? String(ta.value || '') : '';
        if (!text) return;
        downloadText('shared-mailboxes-export.ps1', text);
    });

    getEl('mbBtnLoadGraph')?.addEventListener('click', loadNowGraph);

    getEl('mbImportJson')?.addEventListener('change', async (e) => {
        const f = e.target.files && e.target.files[0];
        if (!f) return;
        try {
            const text = await f.text();
            const obj = safeJsonParse(text);
            const arr = Array.isArray(obj) ? obj : (obj && Array.isArray(obj.rows) ? obj.rows : null);
            if (!arr) throw new Error('Ungültiges JSON. Erwartet: Array oder { rows: [...] }.');
            rows = arr
                .map((r) => {
                    const name = String(r.name || r.displayName || '');
                    const mail = String(r.mail || '');
                    const upn = String(r.upn || r.userPrincipalName || '');
                    const alias = String(r.alias || r.mailNickname || '');
                    const kind =
                        normalizeMailboxKind(r.kind) ||
                        normalizeMailboxKind(r.recipientTypeDetails) ||
                        inferMailboxKind({ name, alias, mail, upn });
                    return {
                        id: String(r.id || ''),
                        name,
                        mail,
                        upn,
                        alias,
                        kind,
                        highConfidence: true,
                        source: 'import'
                    };
                })
                .filter((x) => x.id);
            rows.sort((a, b) => compareDe(a.name, b.name));
            saveCache(rows);
            selectedId = '';
            const c = countMailboxKinds(rows);
            rerender();
            setProgress(
                true,
                `Import OK: ${rows.length} – ${c.shared} freigegeben, ${c.room} Räume, ${c.equipment} Geräte.`
            );
            setTimeout(() => setProgress(false, ''), 2000);
        } catch (err) {
            setProgress(true, 'Import fehlgeschlagen: ' + (err?.message || String(err)));
        } finally {
            e.target.value = '';
        }
    });

    // Optional: prominenter Button im Header neben Login
    try {
        const slot = typeof window.ms365AuthGetActionSlot === 'function' ? window.ms365AuthGetActionSlot() : null;
        if (slot && !document.getElementById('mbHeaderLoadBtn')) {
            const hb = document.createElement('button');
            hb.type = 'button';
            hb.className = 'btn btn-success';
            hb.id = 'mbHeaderLoadBtn';
            hb.style.margin = '0';
            hb.style.padding = '10px 12px';
            hb.style.borderRadius = '10px';
            hb.innerHTML = '<i class="bi bi-box-seam"></i>Exchange‑Export';
            hb.addEventListener('click', showExoExport);
            slot.appendChild(hb);
        }
    } catch {
        // ignore
    }

    rerender();
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bind);
else bind();
