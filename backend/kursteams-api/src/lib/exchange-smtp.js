'use strict';

const crypto = require('crypto');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const { getConfig } = require('./config');
const { sleep, graphJson } = require('./graph-client');

const EXO_SCOPE = 'https://outlook.office365.com/.default';

/** @type {Map<string, import('@azure/msal-node').ConfidentialClientApplication>} */
const exoApps = new Map();

function normalizeDomain(domain) {
    return String(domain || '')
        .trim()
        .replace(/^@+/, '')
        .toLowerCase();
}

function domainFromEmail(email) {
    const m = String(email || '').trim().match(/@([^@]+)$/i);
    return m ? normalizeDomain(m[1]) : '';
}

function parseClientCertificate() {
    const privateKeyEnv = process.env.KURSTEAMS_EXO_CERT_PRIVATE_KEY;
    const thumbprintEnv = process.env.KURSTEAMS_EXO_CERT_THUMBPRINT;
    if (privateKeyEnv && thumbprintEnv) {
        return {
            thumbprint: String(thumbprintEnv).replace(/\s/g, '').toUpperCase(),
            privateKey: String(privateKeyEnv).replace(/\\n/g, '\n')
        };
    }

    const pem = process.env.KURSTEAMS_EXO_CERT_PEM;
    if (!pem) return null;

    const normalized = String(pem).replace(/\\n/g, '\n');
    const certMatch = normalized.match(/-----BEGIN CERTIFICATE-----[\s\S]+?-----END CERTIFICATE-----/);
    const keyMatch =
        normalized.match(/-----BEGIN (?:RSA )?PRIVATE KEY-----[\s\S]+?-----END (?:RSA )?PRIVATE KEY-----/) ||
        normalized.match(/-----BEGIN ENCRYPTED PRIVATE KEY-----[\s\S]+?-----END ENCRYPTED PRIVATE KEY-----/);

    if (!certMatch || !keyMatch) return null;

    let thumbprint = '';
    try {
        const cert = new crypto.X509Certificate(certMatch[0]);
        thumbprint = cert.fingerprint.replace(/:/g, '').toUpperCase();
    } catch {
        return null;
    }

    return { thumbprint, privateKey: keyMatch[0] };
}

function isExoConfigured() {
    return !!parseClientCertificate();
}

function getExoMsalApp(tenantId) {
    const cfg = getConfig();
    const cert = parseClientCertificate();
    if (!cert) {
        throw new Error('Exchange-Zertifikat nicht konfiguriert (KURSTEAMS_EXO_CERT_*).');
    }
    const tid = String(tenantId).trim();
    if (exoApps.has(tid)) return exoApps.get(tid);

    const app = new ConfidentialClientApplication({
        auth: {
            clientId: cfg.clientId,
            authority: 'https://login.microsoftonline.com/' + tid,
            clientCertificate: cert
        }
    });
    exoApps.set(tid, app);
    return app;
}

async function getExoToken(tenantId) {
    const msal = getExoMsalApp(tenantId);
    const result = await msal.acquireTokenByClientCredential({ scopes: [EXO_SCOPE] });
    if (!result || !result.accessToken) {
        throw new Error(
            'Exchange: Kein Access Token (Exchange.ManageAsApp + Zertifikat + Admin Consent im Mandanten?)'
        );
    }
    return result.accessToken;
}

/**
 * @param {string} graphToken
 * @returns {Promise<{ routingDomain: string, defaultDomain: string }>}
 */
async function getTenantMailDomains(graphToken) {
    const org = await graphJson('GET', '/organization?$select=verifiedDomains', graphToken, undefined);
    const list = (org.value && org.value[0] && org.value[0].verifiedDomains) || [];
    let routingDomain = '';
    let defaultDomain = '';
    for (const d of list) {
        const name = String((d && d.name) || '').trim();
        if (!name) continue;
        if (d.isDefault) defaultDomain = name;
        if (/\.onmicrosoft\.com$/i.test(name)) routingDomain = name;
    }
    if (!routingDomain) routingDomain = defaultDomain;
    return { routingDomain, defaultDomain };
}

async function invokeSetUnifiedGroupPrimarySmtp(routingDomain, exoToken, groupId, primarySmtp, anchorUpn) {
    const uri =
        'https://outlook.office365.com/adminapi/beta/' +
        encodeURIComponent(routingDomain) +
        '/InvokeCommand';

    const body = JSON.stringify({
        CmdletInput: {
            CmdletName: 'Set-UnifiedGroup',
            Parameters: {
                Identity: groupId,
                PrimarySmtpAddress: primarySmtp
            }
        }
    });

    const res = await fetch(uri, {
        method: 'POST',
        headers: {
            Authorization: 'Bearer ' + exoToken,
            'Content-Type': 'application/json; charset=utf-8',
            'X-AnchorMailbox': 'UPN:' + anchorUpn,
            'X-Prefer': 'odata.maxpagesize=1'
        },
        body
    });

    const text = await res.text();
    if (!res.ok) {
        throw new Error('Set-UnifiedGroup: ' + res.status + ' ' + text);
    }

    if (text && /error|exception|failed/i.test(text) && !/"@odata.context"/i.test(text)) {
        throw new Error('Set-UnifiedGroup: ' + text.slice(0, 500));
    }
}

/**
 * Setzt die primäre SMTP-Adresse auf nickname@Schuldomain (wie Set-UnifiedGroup in PowerShell).
 * Graph legt Gruppen zuerst unter der Tenant-Standarddomain an (oft *.onmicrosoft.com).
 *
 * @param {{
 *   tenantId: string,
 *   groupId: string,
 *   mailNickname: string,
 *   mailDomain: string,
 *   ownerUpn: string,
 *   graphToken: string,
 *   log: (msg: string) => void
 * }} options
 * @returns {Promise<{ applied: boolean, smtp?: string, reason?: string }>}
 */
async function applySchoolDomainSmtp(options) {
    const mailDomain = normalizeDomain(options.mailDomain);
    const mailNickname = String(options.mailNickname || '').trim();
    if (!mailDomain || !mailNickname) {
        return { applied: false, reason: 'no-domain' };
    }

    const wantedSmtp = mailNickname + '@' + mailDomain;
    const group = await graphJson(
        'GET',
        '/groups/' + options.groupId + '?$select=mail,mailNickname',
        options.graphToken,
        undefined
    );
    const currentMail = String(group.mail || '').trim().toLowerCase();
    if (currentMail === wantedSmtp.toLowerCase()) {
        options.log('E-Mail bereits ' + wantedSmtp + '.');
        return { applied: true, smtp: wantedSmtp };
    }

    if (!isExoConfigured()) {
        options.log(
            'Hinweis: Primäre Adresse ist ' +
                (group.mail || mailNickname + '@…') +
                ' – Ziel wäre ' +
                wantedSmtp +
                '. Exchange-Zertifikat im Backend fehlt (KURSTEAMS_EXO_CERT_*).'
        );
        return { applied: false, reason: 'exo-not-configured', smtp: group.mail || '' };
    }

    const { routingDomain } = await getTenantMailDomains(options.graphToken);
    if (!routingDomain) {
        throw new Error('Tenant-Routing-Domain (.onmicrosoft.com) nicht ermittelbar.');
    }

    const anchorUpn = String(options.ownerUpn || '').trim() || 'admin@' + routingDomain;
    const exoToken = await getExoToken(options.tenantId);

    for (let attempt = 0; attempt < 6; attempt++) {
        try {
            options.log('Exchange: PrimarySmtpAddress = ' + wantedSmtp + ' …');
            await invokeSetUnifiedGroupPrimarySmtp(
                routingDomain,
                exoToken,
                options.groupId,
                wantedSmtp,
                anchorUpn
            );
            await sleep(3000);
            const updated = await graphJson(
                'GET',
                '/groups/' + options.groupId + '?$select=mail',
                options.graphToken,
                undefined
            );
            const newMail = String(updated.mail || '').trim();
            if (newMail.toLowerCase() === wantedSmtp.toLowerCase()) {
                options.log('Exchange: OK → ' + newMail);
                return { applied: true, smtp: newMail };
            }
            if (newMail && domainFromEmail(newMail) === mailDomain) {
                options.log('Exchange: OK → ' + newMail);
                return { applied: true, smtp: newMail };
            }
            throw new Error('Adresse nach Set-UnifiedGroup noch ' + (newMail || 'leer'));
        } catch (e) {
            if (attempt < 5) {
                const wait = 15;
                options.log(
                    'Exchange: Warte auf Postfach (' +
                        (attempt + 1) +
                        '/6), erneut in ' +
                        wait +
                        ' s …'
                );
                await sleep(wait * 1000);
                continue;
            }
            throw e;
        }
    }

    return { applied: false, reason: 'timeout' };
}

module.exports = {
    normalizeDomain,
    domainFromEmail,
    isExoConfigured,
    applySchoolDomainSmtp,
    getTenantMailDomains
};
