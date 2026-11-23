/**
 * b6 – IT-Go-Live-Gate für das Lehrer-Portal.
 * Erst wenn IT „Portal freigeben“ setzt, gilt portalReleased = true.
 */

import {
    evaluateLehrerPortalReadiness,
    gatherLehrerPortalReadinessInput,
    isLehrerPortalReady,
    missingLabels
} from './lehrer-portal-readiness.js';

/**
 * @param {unknown} setup
 * @returns {string}
 */
export function readPortalReleasedAt(setup) {
    const s = setup && typeof setup === 'object' ? setup : {};
    const direct = String(s.portalReleasedAt || '').trim();
    if (direct) return direct;
    const profile =
        s.schoolProfile && typeof s.schoolProfile === 'object' ? s.schoolProfile : {};
    return String(profile.portalReleasedAt || '').trim();
}

/**
 * @param {unknown} setup
 * @returns {boolean}
 */
export function isPortalReleased(setup) {
    return !!readPortalReleasedAt(setup);
}

/**
 * Lehrer-Ansicht: zusätzlich portalReleased nötig, sobald Lizenz erlaubt ist.
 * @param {Parameters<typeof evaluateLehrerPortalReadiness>[0] & { portalReleased?: boolean }} [input]
 */
export function evaluateLehrerPortalWithReleaseGate(input) {
    const o = input || {};
    const base = evaluateLehrerPortalReadiness(o);
    if (base.status !== 'ready') return base;
    if (o.allowed !== true) return base;
    if (o.portalReleased === true) return base;
    return {
        status: 'setup_incomplete',
        missing: ['portalRelease'],
        title: 'Portal noch nicht freigegeben',
        message:
            'Die IT hat das Lehrer-Portal noch nicht freigegeben. Bitte warten Sie, bis Domain, Klassen und Tool-Site stehen und die IT unter „Portal vorbereiten“ freigegeben hat.'
    };
}

/**
 * @returns {ReturnType<typeof evaluateLehrerPortalWithReleaseGate>}
 */
export function readLehrerPortalWithReleaseGate() {
    const input = gatherLehrerPortalReadinessInput();
    let portalReleased = false;
    try {
        if (
            typeof window !== 'undefined' &&
            window.ms365AppDataV2 &&
            typeof window.ms365AppDataV2.getSetup === 'function'
        ) {
            portalReleased = isPortalReleased(window.ms365AppDataV2.getSetup());
        }
    } catch {
        portalReleased = false;
    }
    return evaluateLehrerPortalWithReleaseGate({ ...input, portalReleased });
}

/**
 * @param {boolean} released
 * @returns {string}
 */
export function patchPortalReleased(released) {
    if (
        typeof window === 'undefined' ||
        !window.ms365AppDataV2 ||
        typeof window.ms365AppDataV2.patchSetup !== 'function'
    ) {
        throw new Error('App-Daten nicht geladen.');
    }
    const at = released ? new Date().toISOString() : '';
    window.ms365AppDataV2.patchSetup({ portalReleasedAt: at });
    return at;
}

/**
 * IT-Checkliste vor Freigabe.
 * @param {Parameters<typeof evaluateLehrerPortalReadiness>[0]} [input]
 */
export function evaluateItPortalReleaseGate(input) {
    const prereq = evaluateLehrerPortalReadiness(input);
    const canRelease = isLehrerPortalReady(prereq);
    return {
        canRelease,
        prerequisites: prereq,
        missingLabels: missingLabels(prereq.missing),
        title: canRelease ? 'Bereit zur Freigabe' : 'Noch nicht freigabefähig',
        message: canRelease
            ? 'Tool-Site und Stammdaten sind da. Unter „Portal vorbereiten“ können Sie das Lehrer-Portal freigeben.'
            : 'Zuerst: ' +
              (missingLabels(prereq.missing).join(', ') || 'fehlende Voraussetzungen') +
              '.'
    };
}

export { gatherLehrerPortalReadinessInput };
