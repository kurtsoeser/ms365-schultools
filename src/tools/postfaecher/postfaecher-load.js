/**
 * Robustes Einlesen freigegebener Postfächer (ohne DOM).
 * Primär: Verzeichnis (Graph /users) – im Browser zuverlässig.
 * Optional: Mailbox-Usage-Report als JSON (ohne CSV-Redirect/CORS).
 */

export function normalizeRecipientType(value) {
    return String(value || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '');
}

export function isSharedRecipientType(value) {
    const t = normalizeRecipientType(value);
    return t === 'shared' || t === 'sharedmailbox' || t.indexOf('shared') === 0;
}

export function parseCsvLine(line) {
    const out = [];
    let cur = '';
    let inQuotes = false;
    const s = String(line || '');
    for (let i = 0; i < s.length; i++) {
        const ch = s[i];
        if (inQuotes) {
            if (ch === '"') {
                if (s[i + 1] === '"') {
                    cur += '"';
                    i++;
                } else {
                    inQuotes = false;
                }
            } else {
                cur += ch;
            }
        } else if (ch === '"') {
            inQuotes = true;
        } else if (ch === ',') {
            out.push(cur);
            cur = '';
        } else {
            cur += ch;
        }
    }
    out.push(cur);
    return out;
}

export function parseMailboxUsageCsv(csvText) {
    const raw = String(csvText || '').replace(/^\uFEFF/, '');
    const lines = raw.split(/\r\n|\n|\r/).filter((l) => String(l || '').trim() !== '');
    if (lines.length < 2) {
        return { headers: [], rows: [], hasRecipientType: false };
    }
    const headers = parseCsvLine(lines[0]).map((h) => String(h || '').trim());
    const normHeaders = headers.map((h) => h.toLowerCase().replace(/\s+/g, ''));
    const idx = (names) => {
        for (let i = 0; i < names.length; i++) {
            const n = names[i].toLowerCase().replace(/\s+/g, '');
            const at = normHeaders.indexOf(n);
            if (at !== -1) return at;
        }
        return -1;
    };
    const iUpn = idx(['User Principal Name', 'userPrincipalName']);
    const iName = idx(['Display Name', 'displayName']);
    const iDeleted = idx(['Is Deleted', 'isDeleted']);
    const iRecipient = idx(['Recipient Type', 'RecipientType', 'recipientType']);
    const rows = [];
    for (let li = 1; li < lines.length; li++) {
        const cols = parseCsvLine(lines[li]);
        const upn = iUpn >= 0 ? String(cols[iUpn] || '').trim() : '';
        if (!upn) continue;
        const deletedRaw = iDeleted >= 0 ? String(cols[iDeleted] || '').trim().toLowerCase() : 'false';
        const isDeleted = deletedRaw === 'true' || deletedRaw === '1' || deletedRaw === 'yes';
        rows.push({
            upn,
            name: iName >= 0 ? String(cols[iName] || '').trim() : '',
            isDeleted,
            recipientType: iRecipient >= 0 ? String(cols[iRecipient] || '').trim() : ''
        });
    }
    return { headers, rows, hasRecipientType: iRecipient >= 0 };
}

export function sharedMailboxesFromUsageReport(csvText) {
    const parsed = parseMailboxUsageCsv(csvText);
    if (!parsed.hasRecipientType) {
        return { ok: false, reason: 'missing-recipient-type', rows: [] };
    }
    const rows = parsed.rows
        .filter((r) => !r.isDeleted && isSharedRecipientType(r.recipientType))
        .map((r) => ({
            id: '',
            name: r.name || r.upn,
            mail: '',
            upn: r.upn,
            alias: '',
            highConfidence: true,
            kind: 'shared',
            source: 'report'
        }));
    return { ok: true, reason: '', rows };
}

/** JSON-Variante des Usage-Reports (vermeidet CSV-302 → reports.office.com / CORS). */
export function sharedMailboxesFromUsageJson(payload) {
    const list = Array.isArray(payload)
        ? payload
        : payload && Array.isArray(payload.value)
          ? payload.value
          : null;
    if (!list) return { ok: false, reason: 'invalid-json', rows: [] };

    let sawRecipient = false;
    const rows = [];
    for (let i = 0; i < list.length; i++) {
        const item = list[i] || {};
        const recipientType =
            item.recipientType || item.RecipientType || item['Recipient Type'] || '';
        if (recipientType) sawRecipient = true;
        const deleted = item.isDeleted === true || String(item.isDeleted || '').toLowerCase() === 'true';
        if (deleted) continue;
        if (!isSharedRecipientType(recipientType)) continue;
        const upn = String(item.userPrincipalName || item['User Principal Name'] || '').trim();
        if (!upn) continue;
        rows.push({
            id: '',
            name: String(item.displayName || item['Display Name'] || upn).trim(),
            mail: '',
            upn,
            alias: '',
            highConfidence: true,
            kind: 'shared',
            source: 'report'
        });
    }
    if (!sawRecipient) return { ok: false, reason: 'missing-recipient-type', rows: [] };
    return { ok: true, reason: '', rows };
}

export function scoreSharedMailboxPersonSignals(u) {
    const given = String(u?.givenName || '').trim();
    const sur = String(u?.surname || '').trim();
    const job = String(u?.jobTitle || '').trim();
    const dept = String(u?.department || '').trim();
    const off = String(u?.officeLocation || '').trim();
    const phones = Array.isArray(u?.businessPhones) ? u.businessPhones.filter(Boolean) : [];
    const mobile = String(u?.mobilePhone || '').trim();
    return [given, sur, job, dept, off, mobile, phones.length ? 'phones' : ''].filter(Boolean).length;
}

/**
 * Klassifiziert Verzeichnis-Benutzer als mögliche Shared Mailboxes / Ressourcen.
 * Shared MBs sind oft deaktiviert; manche Tenants haben sie aber aktiviert.
 */
export function classifyDirectorySharedMailboxCandidate(u) {
    if (!u || !u.id) return null;
    const mail = String(u.mail || '').trim();
    const upn = String(u.userPrincipalName || '').trim();
    if (!mail && !upn) return null;

    const userType = String(u.userType || 'Member').trim().toLowerCase();
    if (userType === 'guest') return null;

    const licenses = Array.isArray(u.assignedLicenses) ? u.assignedLicenses.length : 0;
    const enabled = u.accountEnabled === true;
    const personSignals = scoreSharedMailboxPersonSignals(u);

    const looksClassic = enabled === false && personSignals === 0;
    const looksEnabledShared = enabled === true && personSignals === 0 && licenses === 0;
    const looksDisabledWithMail = enabled === false && !!(mail || upn);

    if (!looksClassic && !looksEnabledShared && !looksDisabledWithMail) return null;

    const highConfidence = looksClassic && licenses === 0;
    const name = String(u.displayName || '');
    const alias = String(u.mailNickname || '');
    const kind = inferMailboxKind({ name, alias, mail: mail || upn });
    return {
        id: String(u.id || ''),
        name,
        mail: mail || upn,
        upn: upn || mail,
        alias,
        highConfidence,
        kind,
        source: 'directory',
        accountEnabled: enabled,
        licenseCount: licenses
    };
}

/** shared | room | equipment – aus Exchange RecipientTypeDetails oder Freitext. */
export function normalizeMailboxKind(value) {
    const t = String(value || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '');
    if (!t) return '';
    if (t === 'shared' || t === 'sharedmailbox') return 'shared';
    if (t === 'room' || t === 'roommailbox') return 'room';
    if (t === 'equipment' || t === 'equipmentmailbox' || t === 'resource' || t === 'ressource') {
        return 'equipment';
    }
    return '';
}

/**
 * Heuristik / Places: Raum- und Geräte-Postfächer von Shared trennen.
 * placeKind aus Graph Places (room) hat Vorrang.
 */
export function inferMailboxKind(opts) {
    const placeKind = normalizeMailboxKind(opts?.placeKind || opts?.kind || '');
    if (placeKind === 'room' || placeKind === 'equipment') return placeKind;

    const hay = [opts?.name, opts?.alias, opts?.mail, opts?.upn]
        .map((x) => String(x || '').toLowerCase())
        .join(' ');

    // Räume (DE/EN, Schulkontext)
    if (
        /\b(raum|room|räume|besprech|konferenz|meeting\s*room|klassenraum|horraum|hörsaal|turnsaal|gymnastik|aula|bibliothekssaal|seminarraum|musiksaal)\b/.test(
            hay
        ) ||
        /\br(aum)?[\s._-]?\d{1,4}\b/.test(hay) ||
        /\broom[\s._-]?\d{1,4}\b/.test(hay)
    ) {
        return 'room';
    }

    // Geräte / sonstige Ressourcen
    if (
        /\b(gerät|geraet|equipment|ressource|resource|beamer|projektor|projector|laptop(wagen)?|notebookwagen|kamera|camcorder|drucker|plotter|tablet(wagen)?|medienwagen|technikwagen)\b/.test(
            hay
        )
    ) {
        return 'equipment';
    }

    return 'shared';
}

export function mailboxKindLabel(kind) {
    const k = normalizeMailboxKind(kind) || String(kind || 'shared');
    if (k === 'room') return 'Raum';
    if (k === 'equipment') return 'Gerät';
    if (k === 'shared') return 'Freigegeben';
    return 'Unbekannt';
}

/** Places-E-Mails → kind; überschreibt Heuristik. */
export function applyPlaceKindsToRows(rows, placeByEmail) {
    const map = placeByEmail instanceof Map ? placeByEmail : new Map();
    return (rows || []).map((r) => {
        const keys = [r.mail, r.upn]
            .map((x) => String(x || '').trim().toLowerCase())
            .filter(Boolean);
        let placeKind = '';
        for (let i = 0; i < keys.length; i++) {
            if (map.has(keys[i])) {
                placeKind = map.get(keys[i]);
                break;
            }
        }
        const kind = inferMailboxKind({
            name: r.name,
            alias: r.alias,
            mail: r.mail,
            upn: r.upn,
            placeKind: placeKind || r.kind
        });
        return { ...r, kind };
    });
}

export function buildPlaceEmailKindMap(places) {
    const map = new Map();
    (places || []).forEach((p) => {
        const email = String(p?.emailAddress || p?.mail || '').trim().toLowerCase();
        if (!email) return;
        const odata = String(p?.['@odata.type'] || p?.odataType || '').toLowerCase();
        let kind = 'room';
        if (odata.indexOf('equipment') !== -1 || odata.indexOf('space') !== -1) kind = 'equipment';
        if (odata.indexOf('room') !== -1) kind = 'room';
        // Explizit übergebenes kind
        const forced = normalizeMailboxKind(p?.kind);
        if (forced) kind = forced;
        map.set(email, kind);
    });
    return map;
}

export function filterRowsByMailboxKind(rows, kindFilter) {
    const f = String(kindFilter || 'shared').trim().toLowerCase();
    const list = Array.isArray(rows) ? rows : [];
    if (f === 'all' || f === '*') return list.slice();
    if (f === 'resources') {
        return list.filter((r) => {
            const k = normalizeMailboxKind(r.kind) || inferMailboxKind(r);
            return k === 'room' || k === 'equipment';
        });
    }
    return list.filter((r) => {
        const k = normalizeMailboxKind(r.kind) || inferMailboxKind(r);
        return k === f;
    });
}

export function countMailboxKinds(rows) {
    const counts = { shared: 0, room: 0, equipment: 0, other: 0 };
    (rows || []).forEach((r) => {
        const k = normalizeMailboxKind(r.kind) || inferMailboxKind(r);
        if (k === 'shared' || k === 'room' || k === 'equipment') counts[k]++;
        else counts.other++;
    });
    return counts;
}

export function mapDirectoryUserToMailboxRow(u, source) {
    const classified = classifyDirectorySharedMailboxCandidate(u);
    if (classified) {
        return { ...classified, source: source || classified.source || 'heuristic' };
    }
    const personSignals = scoreSharedMailboxPersonSignals(u);
    const licenses = Array.isArray(u?.assignedLicenses) ? u.assignedLicenses.length : 0;
    return {
        id: String(u?.id || ''),
        name: String(u?.displayName || ''),
        mail: String(u?.mail || ''),
        upn: String(u?.userPrincipalName || ''),
        alias: String(u?.mailNickname || ''),
        highConfidence: personSignals === 0 && licenses === 0 && u?.accountEnabled === false,
        source: source || 'heuristic'
    };
}

export function mergeMailboxRowsByKey(rows) {
    const map = new Map();
    (rows || []).forEach((r) => {
        if (!r) return;
        const key = String(r.id || r.upn || r.mail || '')
            .trim()
            .toLowerCase();
        if (!key) return;
        const prev = map.get(key);
        if (!prev) {
            map.set(key, r);
            return;
        }
        const score = (x) =>
            (x.source === 'report' ? 4 : 0) +
            (x.highConfidence ? 2 : 0) +
            (x.id && String(x.id).indexOf('@') === -1 ? 1 : 0);
        if (score(r) >= score(prev)) {
            map.set(key, { ...prev, ...r, id: r.id || prev.id });
        } else {
            map.set(key, { ...r, ...prev, id: prev.id || r.id });
        }
    });
    return Array.from(map.values());
}

export function mergeReportWithDirectory(reportRows, directoryUsers) {
    const byUpn = new Map();
    const byMail = new Map();
    (directoryUsers || []).forEach((u) => {
        const upn = String(u.userPrincipalName || '').trim().toLowerCase();
        const mail = String(u.mail || '').trim().toLowerCase();
        if (upn) byUpn.set(upn, u);
        if (mail) byMail.set(mail, u);
    });
    return (reportRows || []).map((r) => {
        const key = String(r.upn || '').trim().toLowerCase();
        const u = byUpn.get(key) || byMail.get(key) || null;
        if (!u) {
            return {
                ...r,
                id: r.id || r.upn,
                mail: r.mail || r.upn,
                alias: r.alias || String(r.upn || '').split('@')[0] || '',
                kind: normalizeMailboxKind(r.kind) || 'shared'
            };
        }
        return {
            id: String(u.id || r.upn),
            name: String(u.displayName || r.name || ''),
            mail: String(u.mail || r.mail || r.upn || ''),
            upn: String(u.userPrincipalName || r.upn || ''),
            alias: String(u.mailNickname || r.alias || ''),
            highConfidence: true,
            kind: normalizeMailboxKind(r.kind) || 'shared',
            source: 'report'
        };
    });
}
