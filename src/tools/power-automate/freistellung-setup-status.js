/**
 * Technik-Konto für Power Automate (Flow-Besitzer & Connections).
 * @param {Record<string, unknown>} [cfg]
 */
export function effectiveFreistellungFlowAccount(cfg) {
    const c = cfg || {};
    const explicit = String(c.flowServiceAccount || '').trim().toLowerCase();
    if (explicit) return explicit;
    return String(c.emailMailbox || '').trim().toLowerCase();
}

/**
 * Fortschritts-Anzeige Freistellungen-Setup (rein aus gespeicherten Werten).
 * @param {Record<string, unknown>} cfg
 * @param {{ flowImported?: boolean, onboardingDone?: number, onboardingTotal?: number }} [opts]
 */
export function computeSetupGlance(cfg, opts = {}) {
    const c = cfg || {};
    const siteUrl = String(c.siteUrl || '').trim();
    const listId = String(c.listId || '').trim();
    const flowAccount = effectiveFreistellungFlowAccount(c);
    const emailsOk =
        !!String(c.emailDirektion || '').trim() &&
        !!String(c.emailSonder || '').trim() &&
        !!String(c.emailMailbox || '').trim() &&
        !!flowAccount;

    const onboardingTotal = Math.max(0, Number(opts.onboardingTotal) || 0);
    const onboardingDone = Math.max(0, Number(opts.onboardingDone) || 0);
    const prepOk = onboardingTotal > 0 ? onboardingDone >= onboardingTotal : null;

    return {
        emailsOk,
        listOk: !!(siteUrl && listId),
        flowOk: !!opts.flowImported,
        prepOk,
        onboardingDone,
        onboardingTotal
    };
}
