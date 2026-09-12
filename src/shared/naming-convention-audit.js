/**
 * Namenskonvention-Audit: DisplayName / UPN prüfen.
 * Kein DOM, kein Graph.
 */

function normStr(v) {
    return String(v == null ? '' : v).trim();
}

function localPart(upnOrMail) {
    const s = normStr(upnOrMail).toLowerCase();
    const at = s.indexOf('@');
    return at === -1 ? s : s.slice(0, at);
}

function isNumericLocal(local) {
    return /^\d{3,}$/.test(String(local || ''));
}

/**
 * @param {object} user Graph-User
 * @param {{ displayOrder?: 'given-sur'|'sur-given', allowNumericUpn?: boolean, domain?: string }} [rules]
 */
export function analyzeUserNaming(user, rules) {
    const r = rules || {};
    const displayOrder = r.displayOrder === 'sur-given' ? 'sur-given' : 'given-sur';
    const allowNumericUpn = r.allowNumericUpn !== false;
    const u = user || {};
    const dn = normStr(u.displayName);
    const given = normStr(u.givenName);
    const sur = normStr(u.surname);
    const upn = normStr(u.userPrincipalName);
    const mail = normStr(u.mail);
    const local = localPart(upn || mail);
    const synced = u.onPremisesSyncEnabled === true;

    const issues = [];
    let expectedDisplay = '';
    if (given && sur) {
        expectedDisplay = displayOrder === 'sur-given' ? sur + ' ' + given : given + ' ' + sur;
        if (dn && dn.toLowerCase() !== expectedDisplay.toLowerCase()) {
            // auch "Sur, Given" und reine Zahl
            const alt = displayOrder === 'sur-given' ? given + ' ' + sur : sur + ' ' + given;
            if (dn.toLowerCase() !== alt.toLowerCase()) {
                if (isNumericLocal(dn) || /^\d+$/.test(dn)) {
                    issues.push({ code: 'display_numeric', message: 'Anzeigename ist eine Nummer statt Name.' });
                } else {
                    issues.push({
                        code: 'display_order',
                        message: 'Anzeigename weicht von Konvention „' + expectedDisplay + '“ ab.'
                    });
                }
            }
        }
    } else if (dn && isNumericLocal(dn)) {
        issues.push({ code: 'display_numeric', message: 'Anzeigename ist eine Nummer statt Name.' });
    }

    if (local) {
        if (isNumericLocal(local) && !allowNumericUpn) {
            issues.push({ code: 'upn_numeric', message: 'UPN/Local-Part ist numerisch.' });
        } else if (given && sur && !isNumericLocal(local)) {
            const expectGivenSur = (given + '.' + sur).toLowerCase().replace(/\s+/g, '');
            const expectSurGiven = (sur + '.' + given).toLowerCase().replace(/\s+/g, '');
            const loc = local.replace(/\s+/g, '');
            if (loc !== expectGivenSur && loc !== expectSurGiven && !loc.includes(sur.toLowerCase())) {
                issues.push({
                    code: 'upn_pattern',
                    message: 'UPN-Local-Part entspricht nicht vorname.nachname.'
                });
            }
        }
    }

    let severity = 'ok';
    let action = 'none';
    if (issues.length) {
        if (synced) {
            severity = 'ad_export';
            action = 'export';
        } else {
            severity = 'cloud_fix';
            action = 'patch';
        }
    }

    const patch = {};
    if (action === 'patch' && given && sur) {
        patch.displayName = expectedDisplay || (displayOrder === 'sur-given' ? sur + ' ' + given : given + ' ' + sur);
        if (given) patch.givenName = given;
        if (sur) patch.surname = sur;
    }

    return {
        id: String(u.id || ''),
        displayName: dn,
        givenName: given,
        surname: sur,
        userPrincipalName: upn,
        mail,
        onPremisesSyncEnabled: synced,
        issues,
        severity,
        action,
        expectedDisplay: expectedDisplay || '',
        patch
    };
}

/**
 * @param {object[]} users
 * @param {object} [rules]
 */
export function analyzeUsersNaming(users, rules) {
    const list = Array.isArray(users) ? users : [];
    const rows = list.map((u) => analyzeUserNaming(u, rules));
    const summary = {
        total: rows.length,
        ok: 0,
        cloud_fix: 0,
        ad_export: 0
    };
    rows.forEach(function (r) {
        if (r.severity === 'ok') summary.ok++;
        else if (r.severity === 'cloud_fix') summary.cloud_fix++;
        else if (r.severity === 'ad_export') summary.ad_export++;
    });
    return { rows, summary };
}

/**
 * CSV für AD-IT (nur sync-Benutzer mit Problemen).
 * @param {ReturnType<typeof analyzeUserNaming>[]} rows
 */
export function buildAdExportCsv(rows) {
    const lines = [
        'displayName;givenName;surname;userPrincipalName;mail;expectedDisplay;issues;onPremisesSyncEnabled'
    ];
    (Array.isArray(rows) ? rows : [])
        .filter((r) => r && r.severity === 'ad_export')
        .forEach(function (r) {
            const issues = (r.issues || []).map((i) => i.message).join(' | ');
            lines.push(
                [
                    r.displayName,
                    r.givenName,
                    r.surname,
                    r.userPrincipalName,
                    r.mail,
                    r.expectedDisplay,
                    issues,
                    'true'
                ]
                    .map((c) => '"' + String(c || '').replace(/"/g, '""') + '"')
                    .join(';')
            );
        });
    return lines.join('\r\n');
}

export default { analyzeUserNaming, analyzeUsersNaming, buildAdExportCsv };
