/**
 * Erlaubte Besitzer-E-Mail-Domains für Kursteams (Backend / Graph).
 * Mehrere verifizierte Domains im selben Mandanten (z. B. Login @kurtsoeser.at, Lehrer @kurtrocks.com).
 */

export function emailDomainFromAddress(addr) {
    const m = String(addr || '')
        .trim()
        .match(/@([^@]+)$/i);
    return m ? m[1].toLowerCase() : '';
}

function normalizeDomainToken(v) {
    return String(v ?? '')
        .trim()
        .replace(/^@+/, '')
        .toLowerCase();
}

function addDomain(set, value) {
    const d = normalizeDomainToken(value);
    if (d && d.includes('.')) set.add(d);
}

function collectEmailsFromList(rows, field) {
    const out = [];
    if (!Array.isArray(rows)) return out;
    rows.forEach((row) => {
        if (!row || typeof row !== 'object') return;
        const em = String(row[field] || row.email || '').trim();
        if (em) out.push(em);
    });
    return out;
}

/**
 * @param {string[] | string | undefined} raw
 */
export function parseVerifiedEmailDomainsField(raw) {
    if (!raw) return [];
    if (Array.isArray(raw)) {
        return raw.map((d) => normalizeDomainToken(d)).filter((d) => d.includes('.'));
    }
    return String(raw)
        .split(/[,;\s]+/)
        .map((d) => normalizeDomainToken(d))
        .filter((d) => d.includes('.'));
}

/**
 * Domains, die als Besitzer-UPN für Kursteams im aktuellen Mandanten gelten.
 * @param {{
 *   loginUpn?: string,
 *   schoolDomainNoAt?: string,
 *   tenantSettings?: object | null,
 *   extraEmails?: string[]
 * }} ctx
 * @returns {Set<string>}
 */
export function collectAllowedKursteamOwnerDomains(ctx) {
    const allowed = new Set();
    const tenant = ctx && ctx.tenantSettings;

    addDomain(allowed, emailDomainFromAddress(ctx && ctx.loginUpn));
    addDomain(allowed, ctx && ctx.schoolDomainNoAt);
    if (tenant && typeof tenant === 'object') {
        addDomain(allowed, tenant.domain);
        parseVerifiedEmailDomainsField(tenant.verifiedEmailDomains).forEach((d) => allowed.add(d));

        collectEmailsFromList(tenant.teachers, 'email').forEach((em) =>
            addDomain(allowed, emailDomainFromAddress(em))
        );
        collectEmailsFromList(tenant.classes, 'headEmail').forEach((em) =>
            addDomain(allowed, emailDomainFromAddress(em))
        );
        collectEmailsFromList(tenant.administration, 'email').forEach((em) =>
            addDomain(allowed, emailDomainFromAddress(em))
        );
        collectEmailsFromList(tenant.admin, 'email').forEach((em) =>
            addDomain(allowed, emailDomainFromAddress(em))
        );
        collectEmailsFromList(tenant.arges, 'headEmail').forEach((em) =>
            addDomain(allowed, emailDomainFromAddress(em))
        );
        collectEmailsFromList(tenant.sga, 'email').forEach((em) =>
            addDomain(allowed, emailDomainFromAddress(em))
        );
        collectEmailsFromList(tenant.studentCouncil, 'email').forEach((em) =>
            addDomain(allowed, emailDomainFromAddress(em))
        );
    }

    const extra = (ctx && ctx.extraEmails) || [];
    extra.forEach((em) => addDomain(allowed, emailDomainFromAddress(em)));

    return allowed;
}

/**
 * @param {string[]} ownerAddresses
 * @param {Set<string> | string[]} allowedDomains
 * @returns {string[]}
 */
export function findForeignOwnerDomains(ownerAddresses, allowedDomains) {
    const allowed =
        allowedDomains instanceof Set
            ? allowedDomains
            : new Set(
                  (Array.isArray(allowedDomains) ? allowedDomains : [])
                      .map((d) => normalizeDomainToken(d))
                      .filter(Boolean)
              );
    const foreign = new Set();
    (ownerAddresses || []).forEach((addr) => {
        const d = emailDomainFromAddress(addr);
        if (!d) return;
        if (!allowed.has(d)) foreign.add(d);
    });
    return [...foreign].sort();
}

/**
 * @param {{
 *   tenantId?: string,
 *   loginUpn?: string,
 *   teams?: Array<{ besitzer?: string }>,
 *   schoolDomainNoAt?: string,
 *   tenantSettings?: object | null,
 *   extraEmails?: string[]
 * }} input
 * @returns {{ ok: boolean, message?: string, allowedDomains?: string[], foreignDomains?: string[] }}
 */
export function validateKursteamBackendTenantContext(input) {
    const tenantId = input && input.tenantId;
    const loginUpn = String((input && input.loginUpn) || '').trim();
    const teams = (input && input.teams) || [];

    if (!tenantId) {
        return {
            ok: false,
            message:
                'Kein Mandant erkannt. Bitte unten links bei Microsoft mit dem Schul-Konto anmelden.'
        };
    }

    if (!loginUpn) {
        return {
            ok: false,
            message: 'Bitte bei Microsoft anmelden, bevor Kursteams online angelegt werden.'
        };
    }

    const ownerAddresses = teams.map((t) => String(t.besitzer || '').trim()).filter(Boolean);
    const ownerDomains = [...new Set(ownerAddresses.map(emailDomainFromAddress).filter(Boolean))];

    if (!ownerDomains.length) {
        return { ok: true, allowedDomains: [], foreignDomains: [] };
    }

    const allowed = collectAllowedKursteamOwnerDomains(input);
    const foreign = findForeignOwnerDomains(ownerAddresses, allowed);

    if (!foreign.length) {
        return {
            ok: true,
            allowedDomains: [...allowed].sort(),
            foreignDomains: []
        };
    }

    const allowedList = [...allowed].sort();
    return {
        ok: false,
        allowedDomains: allowedList,
        foreignDomains: foreign,
        message:
            'Die Besitzer nutzen Domain(s) (' +
            foreign.join(', ') +
            '), die nicht zu den erlaubten Mandanten-Domains passen (' +
            (allowedList.length ? allowedList.join(', ') : 'keine konfiguriert') +
            '). In den Stammdaten Schul-Domain und Lehrer-E-Mails pflegen, optional weitere Domains ergänzen, oder Besitzer in Schritt 5 anpassen. Angemeldet: ' +
            loginUpn +
            '.'
    };
}
