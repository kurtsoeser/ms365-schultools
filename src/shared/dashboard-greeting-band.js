/**
 * Dashboard: personalisiertes Begrüßungsband statt Dekorations-Hero.
 */

/** @param {number} [hour] */
export function greetingForHour(hour) {
    const h = Number.isFinite(hour) ? hour : new Date().getHours();
    if (h < 11) return 'Guten Morgen';
    if (h < 18) return 'Guten Tag';
    return 'Guten Abend';
}

/** @param {string} displayName */
export function extractFirstName(displayName) {
    const s = String(displayName || '').trim();
    if (!s || s === 'Konto' || s === '–') return '';
    return s.split(/\s+/)[0] || '';
}

/**
 * @param {string[]} tones data-tone von Fortschrittszeilen
 * @param {{ mismatch?: number, emptyList?: number, unmatched?: number } | null} hygieneCounts
 */
export function countOpenDeviations(tones, hygieneCounts) {
    let n = 0;
    (tones || []).forEach(function (t) {
        const tone = String(t || '').trim();
        if (!tone) return;
        if (tone === 'warn' || tone === 'mismatch' || tone === 'unmatched') n += 1;
    });
    const c = hygieneCounts || {};
    if ((c.mismatch || 0) > 0 || (c.emptyList || 0) > 0) n += 1;
    else if ((c.unmatched || 0) > 0) n += 1;
    return n;
}

/** @param {number} n */
export function formatDeviationPhrase(n) {
    const v = Math.max(0, Number(n) || 0);
    if (!v) return '';
    if (v === 1) return '1 offene Abweichung';
    return v + ' offene Abweichungen';
}

/**
 * @param {{ greeting: string, firstName?: string, deviations?: number, schuljahrSteps?: string, hasSchoolData?: boolean }} snap
 */
export function formatGreetingBandText(snap) {
    const s = snap && typeof snap === 'object' ? snap : {};
    const lead = [s.greeting || greetingForHour(), s.firstName ? ', ' + s.firstName : ''].join('');
    const parts = [lead];

    const devPhrase = formatDeviationPhrase(s.deviations);
    if (devPhrase) {
        parts.push(devPhrase);
    } else if (s.hasSchoolData) {
        parts.push('Keine offenen Abweichungen');
    } else {
        parts.push('Schuldaten einrichten – Playbook oder Schulregister');
    }

    const steps = String(s.schuljahrSteps || '').trim();
    if (steps) parts.push('Schuljahresstart: ' + steps);

    return parts.join(' · ');
}

/**
 * HTML für Status-Banner (Spec 2.6 – Schwerpunkte in &lt;strong&gt;).
 * @param {{ greeting: string, firstName?: string, deviations?: number, schuljahrSteps?: string, hasSchoolData?: boolean }} snap
 */
export function formatGreetingBandHtml(snap) {
    const s = snap && typeof snap === 'object' ? snap : {};
    const lead = [s.greeting || greetingForHour(), s.firstName ? ', ' + s.firstName : ''].join('');
    const chunks = [lead];

    const devPhrase = formatDeviationPhrase(s.deviations);
    if (devPhrase) {
        chunks.push('<strong>' + devPhrase + '</strong>');
    } else if (s.hasSchoolData) {
        chunks.push('Keine offenen Abweichungen');
    } else {
        chunks.push('Schuldaten einrichten – Playbook oder Schulregister');
    }

    const steps = String(s.schuljahrSteps || '').trim();
    if (steps) {
        chunks.push('Schuljahresstart: <strong>' + steps + '</strong>');
    }

    return chunks.join(' · ');
}

function progressToneFromEl(id) {
    const el = document.getElementById(id);
    if (!el) return '';
    const text = String(el.textContent || '').trim();
    if (!text) return '';
    return String(el.getAttribute('data-tone') || '').trim();
}

function readHygieneCounts() {
    try {
        const h = window.ms365MembershipHygiene;
        if (h && typeof h.loadHygieneScanCache === 'function') {
            const cache = h.loadHygieneScanCache();
            return cache && cache.counts ? cache.counts : null;
        }
    } catch {
        /* ignore */
    }
    return null;
}

function readSchuljahrStepsLabel() {
    try {
        const sj = window.ms365SchuljahresstartPlaybook;
        const pb = window.ms365Playbook;
        if (sj && pb && typeof pb.loadPlaybookState === 'function' && typeof sj.schuljahresstartPlaybookProgress === 'function') {
            const st = pb.loadPlaybookState(sj.SCHULJAHRSTART_PLAYBOOK_STORAGE_KEY);
            const prog = sj.schuljahresstartPlaybookProgress(st);
            if (prog && prog.total) return prog.done + '/' + prog.total + ' Schritte';
        }
    } catch {
        /* ignore */
    }
    const el = document.getElementById('dashProgSchuljahr');
    if (!el) return '';
    const text = String(el.textContent || '').trim();
    const m = text.match(/(\d+)\s*\/\s*(\d+)/);
    if (m) return m[1] + '/' + m[2] + ' Schritte';
    return '';
}

function readDisplayName() {
    const menu = document.getElementById('ms365AuthMenuName');
    if (menu && String(menu.textContent || '').trim()) return String(menu.textContent).trim();
    const badge = document.getElementById('ms365AuthBadgeText');
    if (badge && String(badge.textContent || '').trim()) return String(badge.textContent).trim();
    return '';
}

function hasMeaningfulSchoolData() {
    try {
        if (typeof window.ms365TenantSettingsLoad === 'function') {
            const s = window.ms365TenantSettingsLoad();
            const domain = String((s && s.domain) || '').trim();
            const schoolName = String((s && s.schoolName) || '').trim();
            const subjects = ((s && s.subjects) || []).length;
            const teachers = ((s && s.teachers) || []).length;
            const students = ((s && s.students) || []).length;
            const classes = ((s && s.classes) || []).length;
            return !!(domain || schoolName || subjects || teachers || students || classes);
        }
    } catch {
        /* ignore */
    }
    return false;
}

export function buildGreetingBandSnapshot() {
    const tones = [
        progressToneFromEl('dashProgGruppen'),
        progressToneFromEl('dashProgUnterricht'),
        progressToneFromEl('dashProgSchuljahr')
    ];
    return {
        greeting: greetingForHour(),
        firstName: extractFirstName(readDisplayName()),
        deviations: countOpenDeviations(tones, readHygieneCounts()),
        schuljahrSteps: readSchuljahrStepsLabel(),
        hasSchoolData: hasMeaningfulSchoolData()
    };
}

export function refreshDashboardGreetingBand() {
    const el = document.getElementById('dashGreetingBandText');
    if (!el) return;
    const snap = buildGreetingBandSnapshot();
    el.innerHTML = formatGreetingBandHtml(snap);
}

export function mountDashboardGreetingBand() {
    const band = document.getElementById('dashGreetingBand');
    if (!band || band.dataset.bound === '1') return;
    band.dataset.bound = '1';

    function refresh() {
        refreshDashboardGreetingBand();
    }

    refresh();
    window.addEventListener('ms365-auth-state-changed', refresh);
    window.addEventListener('ms365-auth-widget-ready', refresh);
    window.addEventListener('ms365-local-data-changed', refresh);
    window.addEventListener('ms365-tenant-settings-changed', refresh);
    window.addEventListener('ms365-dash-main-section-changed', refresh);

    window.__ms365DashGreetingRefresh = refresh;
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', mountDashboardGreetingBand);
    } else {
        mountDashboardGreetingBand();
    }
}
